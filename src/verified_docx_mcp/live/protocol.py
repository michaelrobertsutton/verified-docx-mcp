"""Wire protocol for the WSS ops channel (issue #106 WP-2:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

Typed messages exchanged between the bridge (``live/bridge.py``, running in
the MCP process) and the Word task pane (``addin/taskpane.js``) over
``wss://127.0.0.1:<ops_port>/ops``. Every message is a JSON object with a
``type`` field; ``to_json``/``from_json`` on each dataclass below is the only
place that knows the wire shape (camelCase keys, matching the pane's own
JS naming) versus the Python-side snake_case field names. ``from_json``
raises ``ProtocolError`` (never a bare ``KeyError``/``TypeError``) on a
malformed message, so the bridge can turn a bad message into a clean
disconnect instead of a stack trace.

Message directions
-------------------
Pane -> server (unsolicited):
  ``hello``      -- sent once, immediately after the WSS connection opens.
  ``heartbeat``  -- sent every 5s thereafter (``live/session.py``'s
                    ``SessionRegistry`` evicts a session after 3 missed
                    heartbeats -- see that module).

Server -> pane (request/reply, correlated by ``request_id``):
  ``request``    -- one of the ops below; the pane must reply with exactly
                    one ``reply`` carrying the same ``request_id``.

Pane -> server (solicited):
  ``reply``      -- ``{type: "reply", request_id, ok, result?, error?}``.
                    ``error`` is ``{code, message}`` when ``ok`` is false.

Ops (the ``op`` field of a ``request``, and the payload each one carries)
--------------------------------------------------------------------------
Each op's docstring below names the Word JavaScript API calls
``addin/taskpane.js``'s dispatcher makes to satisfy it — this is the
contract the fake pane (``tests/unit/fake_pane.py``) mimics and the real
pane must honor.

  ping            -- no payload. Pane replies ``{}`` once it has round-
                     tripped a no-op through ``Word.run`` (proves the
                     JS engine, not just the socket, is alive).

  describe        -- no payload. Word calls: ``context.document.body.load
                     ("text")``, ``context.document.load("changeTrackingMode,
                     saved")``, ``Office.context.document.url``. Reply
                     result: ``DescribeResult`` (documentUrl, bodySha256,
                     changeTrackingMode, saved).

  search          -- ``SearchPayload`` (find, matchCase, matchWholeWord).
                     Word calls: ``body.search(find, {matchCase,
                     matchWholeWord})``, ``.load(["text"])`` on each
                     returned range, plus enough of the surrounding
                     paragraph text (loaded separately) to slice
                     contextBefore/contextAfter. Reply result:
                     ``{"matches": [SearchMatch, ...]}``.

  replace         -- ``ReplacePayload`` (find, expected_matches, replace,
                     track_changes). Word calls: ``body.search(find, ...)``;
                     if the match count != expected_matches, no edit is
                     made and the pane replies ok=false (LIVE_OP_FAILED,
                     not a protocol error) with the actual count in
                     ``error.message``. Otherwise, for each match range:
                     optionally ``document.changeTrackingMode =
                     Word.ChangeTrackingMode.trackAll`` for the duration,
                     then ``range.insertText(payload.replace,
                     Word.InsertLocation.replace)``, then re-loads the
                     range's text to confirm. Reply result:
                     ``ReplaceResult`` (applied, match_count, matches:
                     list of ``ReplaceMatchResult`` {before, after}, pre,
                     post -- pre/post are ``body.text`` SHA-256 read
                     before and after the whole op).

  format          -- ``FormatPayload`` (find, expected_matches, bold,
                     italic, underline, track_changes). Same
                     search-and-count-gate as replace; on a match, Word
                     calls: ``range.font.bold``/``.italic``/``.underline``
                     assignment for whichever of the three are not
                     ``None``, under the same optional
                     ``changeTrackingMode`` toggle. Reply result: same
                     shape as ``ReplaceResult`` (before/after are the
                     matched text, unchanged by a format-only op, so the
                     evidence is the pre/post body hash plus
                     match_count).

  comments_list   -- no payload. Word calls: ``body.getComments()``,
                     ``comment.load(["id","content","authorName",
                     "creationDate","resolved"])``, ``comment.getRange()
                     .load("text")`` for anchorText, and
                     ``comment.getReplies()`` loaded the same way (minus
                     anchorText -- a reply has no anchor of its own).
                     Reply result: ``{"comments": [CommentInfo, ...]}``.

  comment_add     -- ``CommentAddPayload`` (find, expected_matches, text).
                     Same search-and-count-gate. Word calls:
                     ``range.insertComment(payload.text)``, then reloads
                     the new ``Comment.id``. Reply result:
                     ``{"comment_id": <str>, "pre": <sha>, "post": <sha>}``
                     -- pre/post are included even though body text is
                     unchanged by a comment insert, for the same staleness
                     check every mutating op supports.

  comment_reply   -- ``CommentReplyPayload`` (comment_id, text). Word
                     calls: ``comment.reply(payload.text)`` on the
                     ``Comment`` looked up by the pane's own in-memory
                     ``Comment.id`` map (never the OOXML durableId -- see
                     the WP-1 result recorded in docs/live-mode.md: the
                     two ids are unrelated). Reply result:
                     ``{"reply_id": <str>}``.

  comment_resolve -- ``CommentResolvePayload`` (comment_id, resolved).
                     Word calls: ``comment.resolved = payload.resolved``.
                     Reply result: ``{"resolved": <bool>}``.

  save            -- no payload. Word calls: ``context.document.save()``.
                     Reply result: ``{"saved": true}``.

Staleness (``expected_body_sha256``)
-------------------------------------
``search``/``replace``/``format``/``comment_add`` payloads carry an
optional ``expected_body_sha256``. The pane does not act on it -- it is
read by ``live/session.py``'s ``LiveSession.request`` after the reply
comes back: if the caller supplied one and the reply's ``result["pre"]``
(the pane's body hash *before* it ran the op) differs, the caller gets
``LiveStale`` instead of trusting a result computed against text the
caller no longer expects.
"""

from __future__ import annotations

import dataclasses
import json
from typing import Any


class ProtocolError(ValueError):
    """A message did not match its expected wire shape."""


VALID_OPS: frozenset[str] = frozenset(
    {
        "ping",
        "describe",
        "search",
        "replace",
        "format",
        "comments_list",
        "comment_add",
        "comment_reply",
        "comment_resolve",
        "save",
    }
)


# Pane-side OpError.code values for a search-and-count-gate refusal on
# `replace`/`format`/`comment_add` (issue #106 WP-3/WP-4: both add a
# write_mode="live" path and both need to tell "nothing matched" apart
# from "the wrong number matched" the same way the file-mode ZERO_MATCH/
# MATCH_COUNT_MISMATCH split already does -- see errors.py's ErrorCode
# members of the same name, which server.py's live tool wrappers map
# these two onto). Lowercase, snake_case, and named identically on both
# WPs' branches so the two additions merge without conflict.
OP_ERROR_ZERO_MATCH = "zero_match"
OP_ERROR_MATCH_COUNT_MISMATCH = "match_count_mismatch"


def _require(obj: dict[str, Any], key: str, expected_type: type | tuple[type, ...]) -> Any:
    if key not in obj:
        raise ProtocolError(f"missing required field {key!r}")
    value = obj[key]
    if not isinstance(value, expected_type):
        raise ProtocolError(f"field {key!r} must be {expected_type}, got {type(value).__name__}")
    return value


def _optional(obj: dict[str, Any], key: str, expected_type: type | tuple[type, ...], default: Any = None) -> Any:
    if key not in obj or obj[key] is None:
        return default
    value = obj[key]
    if not isinstance(value, expected_type):
        raise ProtocolError(f"field {key!r} must be {expected_type}, got {type(value).__name__}")
    return value


def _as_dict(raw: str | bytes | dict[str, Any]) -> dict[str, Any]:
    if isinstance(raw, dict):
        return raw
    try:
        obj = json.loads(raw)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise ProtocolError(f"not valid JSON: {exc}") from exc
    if not isinstance(obj, dict):
        raise ProtocolError(f"top-level message must be a JSON object, got {type(obj).__name__}")
    return obj


# ---------------------------------------------------------------------------
# Pane -> server (unsolicited)
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True)
class HelloMessage:
    """First message a pane sends after the WSS connection opens.

    ``document_name`` is NOT sent on the wire -- ``live/session.py``
    derives it from ``document_url`` (the basename) when it keys the
    session, per the plan's "keyed by document file name from
    documentUrl".
    """

    document_url: str
    host: str
    platform: str
    requirement_sets: dict[str, Any]
    body_sha256: str

    def to_json(self) -> dict[str, Any]:
        return {
            "type": "hello",
            "documentUrl": self.document_url,
            "host": self.host,
            "platform": self.platform,
            "requirementSets": self.requirement_sets,
            "bodySha256": self.body_sha256,
        }

    @classmethod
    def from_json(cls, raw: str | bytes | dict[str, Any]) -> HelloMessage:
        obj = _as_dict(raw)
        msg_type = _require(obj, "type", str)
        if msg_type != "hello":
            raise ProtocolError(f"expected type 'hello', got {msg_type!r}")
        return cls(
            document_url=_require(obj, "documentUrl", str),
            host=_require(obj, "host", str),
            platform=_require(obj, "platform", str),
            requirement_sets=_require(obj, "requirementSets", dict),
            body_sha256=_require(obj, "bodySha256", str),
        )


@dataclasses.dataclass(frozen=True)
class HeartbeatMessage:
    """Sent every 5s by a connected pane."""

    document_url: str
    body_sha256: str

    def to_json(self) -> dict[str, Any]:
        return {"type": "heartbeat", "documentUrl": self.document_url, "bodySha256": self.body_sha256}

    @classmethod
    def from_json(cls, raw: str | bytes | dict[str, Any]) -> HeartbeatMessage:
        obj = _as_dict(raw)
        msg_type = _require(obj, "type", str)
        if msg_type != "heartbeat":
            raise ProtocolError(f"expected type 'heartbeat', got {msg_type!r}")
        return cls(
            document_url=_require(obj, "documentUrl", str),
            body_sha256=_require(obj, "bodySha256", str),
        )


# ---------------------------------------------------------------------------
# Server -> pane (request) / pane -> server (reply)
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True)
class OpRequest:
    request_id: str
    op: str
    payload: dict[str, Any] = dataclasses.field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.op not in VALID_OPS:
            raise ProtocolError(f"unknown op {self.op!r}; must be one of {sorted(VALID_OPS)}")

    def to_json(self) -> dict[str, Any]:
        return {"type": "request", "request_id": self.request_id, "op": self.op, "payload": self.payload}

    @classmethod
    def from_json(cls, raw: str | bytes | dict[str, Any]) -> OpRequest:
        obj = _as_dict(raw)
        msg_type = _require(obj, "type", str)
        if msg_type != "request":
            raise ProtocolError(f"expected type 'request', got {msg_type!r}")
        return cls(
            request_id=_require(obj, "request_id", str),
            op=_require(obj, "op", str),
            payload=_optional(obj, "payload", dict, default={}),
        )


@dataclasses.dataclass(frozen=True)
class OpError:
    code: str
    message: str

    def to_json(self) -> dict[str, Any]:
        return {"code": self.code, "message": self.message}

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> OpError:
        return cls(code=_require(obj, "code", str), message=_require(obj, "message", str))


@dataclasses.dataclass(frozen=True)
class OpReply:
    request_id: str
    ok: bool
    result: dict[str, Any] | None = None
    error: OpError | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {"type": "reply", "request_id": self.request_id, "ok": self.ok}
        if self.result is not None:
            out["result"] = self.result
        if self.error is not None:
            out["error"] = self.error.to_json()
        return out

    @classmethod
    def from_json(cls, raw: str | bytes | dict[str, Any]) -> OpReply:
        obj = _as_dict(raw)
        msg_type = _require(obj, "type", str)
        if msg_type != "reply":
            raise ProtocolError(f"expected type 'reply', got {msg_type!r}")
        ok = _require(obj, "ok", bool)
        result = _optional(obj, "result", dict, default=None)
        error_obj = _optional(obj, "error", dict, default=None)
        error = OpError.from_json(error_obj) if error_obj is not None else None
        if not ok and error is None:
            raise ProtocolError("reply with ok=false must carry an 'error' object")
        return cls(request_id=_require(obj, "request_id", str), ok=ok, result=result, error=error)


def peek_type(raw: str | bytes | dict[str, Any]) -> str:
    """Return a message's ``type`` field without fully validating it --
    used by the bridge's receive loop to route a message to the right
    ``from_json`` before it commits to a shape."""
    obj = _as_dict(raw)
    return _require(obj, "type", str)


# ---------------------------------------------------------------------------
# Per-op payloads and results (server-side construction / pane-side
# contract). These are documented shapes, not wire envelopes on their own
# -- a payload dataclass's ``to_json()`` becomes an ``OpRequest.payload``;
# a result dataclass's ``from_json()`` parses an ``OpReply.result``.
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True)
class SearchPayload:
    find: str
    match_case: bool = False
    match_whole_word: bool = False
    expected_body_sha256: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "matchCase": self.match_case,
            "matchWholeWord": self.match_whole_word,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> SearchPayload:
        find = _require(obj, "find", str)
        if not find:
            raise ProtocolError("'find' must not be empty")
        return cls(
            find=find,
            match_case=_optional(obj, "matchCase", bool, default=False),
            match_whole_word=_optional(obj, "matchWholeWord", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class SearchMatch:
    index: int
    text: str
    context_before: str
    context_after: str

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> SearchMatch:
        return cls(
            index=_require(obj, "index", int),
            text=_require(obj, "text", str),
            context_before=_optional(obj, "contextBefore", str, default=""),
            context_after=_optional(obj, "contextAfter", str, default=""),
        )


@dataclasses.dataclass(frozen=True)
class ReplacePayload:
    find: str
    expected_matches: int
    replace: str
    track_changes: bool = False
    expected_body_sha256: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "expected_matches": self.expected_matches,
            "replace": self.replace,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> ReplacePayload:
        find = _require(obj, "find", str)
        if not find:
            raise ProtocolError("'find' must not be empty")
        expected_matches = _require(obj, "expected_matches", int)
        if expected_matches < 1:
            raise ProtocolError("'expected_matches' must be >= 1")
        return cls(
            find=find,
            expected_matches=expected_matches,
            replace=_require(obj, "replace", str),
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class FormatPayload:
    find: str
    expected_matches: int
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    track_changes: bool = False
    expected_body_sha256: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "expected_matches": self.expected_matches,
            "bold": self.bold,
            "italic": self.italic,
            "underline": self.underline,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> FormatPayload:
        find = _require(obj, "find", str)
        if not find:
            raise ProtocolError("'find' must not be empty")
        expected_matches = _require(obj, "expected_matches", int)
        if expected_matches < 1:
            raise ProtocolError("'expected_matches' must be >= 1")
        bold = _optional(obj, "bold", bool, default=None)
        italic = _optional(obj, "italic", bool, default=None)
        underline = _optional(obj, "underline", bool, default=None)
        if bold is None and italic is None and underline is None:
            raise ProtocolError("at least one of bold/italic/underline must be set")
        return cls(
            find=find,
            expected_matches=expected_matches,
            bold=bold,
            italic=italic,
            underline=underline,
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class ReplaceMatchResult:
    before: str
    after: str

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> ReplaceMatchResult:
        return cls(before=_require(obj, "before", str), after=_require(obj, "after", str))


@dataclasses.dataclass(frozen=True)
class CommentReplyInfo:
    id: str
    content: str
    author_name: str
    creation_date: str

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CommentReplyInfo:
        return cls(
            id=_require(obj, "id", str),
            content=_require(obj, "content", str),
            author_name=_require(obj, "authorName", str),
            creation_date=_require(obj, "creationDate", str),
        )


@dataclasses.dataclass(frozen=True)
class CommentInfo:
    id: str
    content: str
    author_name: str
    creation_date: str
    resolved: bool
    anchor_text: str
    replies: list[CommentReplyInfo] = dataclasses.field(default_factory=list)

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CommentInfo:
        replies_raw = _optional(obj, "replies", list, default=[])
        return cls(
            id=_require(obj, "id", str),
            content=_require(obj, "content", str),
            author_name=_require(obj, "authorName", str),
            creation_date=_require(obj, "creationDate", str),
            resolved=_require(obj, "resolved", bool),
            anchor_text=_optional(obj, "anchorText", str, default=""),
            replies=[CommentReplyInfo.from_json(r) for r in replies_raw],
        )


@dataclasses.dataclass(frozen=True)
class CommentAddPayload:
    find: str
    expected_matches: int
    text: str
    expected_body_sha256: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "expected_matches": self.expected_matches,
            "text": self.text,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CommentAddPayload:
        find = _require(obj, "find", str)
        if not find:
            raise ProtocolError("'find' must not be empty")
        expected_matches = _require(obj, "expected_matches", int)
        if expected_matches < 1:
            raise ProtocolError("'expected_matches' must be >= 1")
        text = _require(obj, "text", str)
        if not text:
            raise ProtocolError("'text' must not be empty")
        return cls(
            find=find,
            expected_matches=expected_matches,
            text=text,
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class CommentReplyPayload:
    comment_id: str
    text: str

    def to_json(self) -> dict[str, Any]:
        return {"comment_id": self.comment_id, "text": self.text}

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CommentReplyPayload:
        comment_id = _require(obj, "comment_id", str)
        if not comment_id:
            raise ProtocolError("'comment_id' must not be empty")
        text = _require(obj, "text", str)
        if not text:
            raise ProtocolError("'text' must not be empty")
        return cls(comment_id=comment_id, text=text)


@dataclasses.dataclass(frozen=True)
class CommentResolvePayload:
    comment_id: str
    resolved: bool

    def to_json(self) -> dict[str, Any]:
        return {"comment_id": self.comment_id, "resolved": self.resolved}

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CommentResolvePayload:
        comment_id = _require(obj, "comment_id", str)
        if not comment_id:
            raise ProtocolError("'comment_id' must not be empty")
        return cls(comment_id=comment_id, resolved=_require(obj, "resolved", bool))


@dataclasses.dataclass(frozen=True)
class DescribeResult:
    document_url: str
    body_sha256: str
    change_tracking_mode: str
    saved: bool

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> DescribeResult:
        return cls(
            document_url=_require(obj, "documentUrl", str),
            body_sha256=_require(obj, "bodySha256", str),
            change_tracking_mode=_require(obj, "changeTrackingMode", str),
            saved=_require(obj, "saved", bool),
        )
