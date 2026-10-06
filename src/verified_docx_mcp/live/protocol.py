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

  body_ooxml      -- no payload (issue #33; capability ``body_ooxml``).
                     Word calls: ``body.getOoxml()``, ``body.load("text")``.
                     Reply result: ``{ooxml, bodySha256, documentUrl,
                     strippedParts}`` -- ``ooxml`` is the Flat OPC
                     ``pkg:package`` string with every ``pkg:binaryData``
                     payload emptied (``strippedParts`` names them);
                     ``bodySha256`` is computed exactly as ``describe``'s.
                     Refuses with ``OP_ERROR_TOO_LARGE`` when the reply
                     would exceed the pane's size limit.

  search        -- ``SearchPayload`` (find, matchCase, matchWholeWord).
                     Word calls: ``body.search(find, {matchCase,
                     matchWholeWord})``, ``.load(["text"])`` on each
                     returned range, plus enough of the surrounding
                     paragraph text (loaded separately) to slice
                     contextBefore/contextAfter. Reply result:
                     ``{"matches": [SearchMatch, ...]}``.

  replace         -- ``ReplacePayload`` (find, expected_matches, replace,
                     track_changes, rowAnchor). rowAnchor (issue #22 B2,
                     optional -- REQUIRES the pane report the "row_scope"
                     capability in its own hello, checked by
                     ``live/write_mode.py``'s ``require_capability``
                     BEFORE this op is ever sent, so an old pane can never
                     silently ignore it and run an unscoped op instead):
                     the pane first resolves rowAnchor's own table row
                     (``range.parentTableCellOrNullObject.parentTableCellOrNullObject``
                     null -> LIVE_OP_FAILED "rowAnchor is not in a table
                     cell"), then scopes `find`'s own
                     ``body.search``/count-gate/edit to that row's own
                     ``Range`` (``row.getRange()``) instead of the whole
                     body. Word calls (unscoped case): ``body.search(find,
                     ...)``; if the match count != expected_matches, no
                     edit is made and the pane replies ok=false
                     (LIVE_OP_FAILED, not a protocol error) with the
                     actual count in ``error.message``. Otherwise, for
                     each match range:
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
                     italic, underline, strike, color, track_changes,
                     rowAnchor). rowAnchor is the same row-scoping/
                     capability-gated mechanism ``replace`` documents
                     above. Same search-and-count-gate as replace; on a match,
                     Word calls: ``range.font.bold``/``.italic``/
                     ``.underline``/``.strikeThrough`` assignment for
                     whichever of the four are not ``None``, and
                     ``range.font.color = "#RRGGBB"`` when ``color`` is
                     given, under the same optional ``changeTrackingMode``
                     toggle. After ``context.sync()``, the pane RE-LOADS
                     ``range.font.color``/``.strikeThrough`` per match and
                     includes the read-back values in the reply
                     (``colorAfter``/``strikeAfter`` on each match) --
                     issue #22: a caller checks these against what it
                     asked for, rather than trusting an echo of the
                     request, so a write that silently didn't take (a
                     protected range, a stale object reference) is
                     detectable. Reply result: same shape as
                     ``ReplaceResult`` (before/after are the matched text,
                     unchanged by a format-only op, so the evidence is
                     the pre/post body hash plus match_count) plus the
                     two read-back fields above on each match.

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

  cell_get        -- ``CellGetPayload`` (table_index, row_index, cell_index;
                     all 1-based, table_index numbered exactly like
                     ``list_tables``' ``table_id``). Issue #27; REQUIRES
                     the pane's ``"cell_edit"`` capability (checked by
                     ``live/write_mode.py``'s ``require_capability``
                     before the op is sent). Word calls:
                     ``body.tables`` (each ``.load("nestingLevel")``),
                     ``table.rows`` -> ``row.cells`` -> the addressed
                     ``TableCell``, ``cell.body.load("text")``. A document
                     containing a NESTED table is refused (LIVE_OP_FAILED):
                     ``body.tables`` cannot be mapped onto ``table_id``
                     numbering once tables nest. Reply result:
                     ``{"text": <cell body text>}``.

  cell_set        -- ``CellSetPayload`` (table_index, row_index,
                     cell_index, paragraphs, expected_before_text,
                     track_changes). Compare-and-set: the pane re-reads the
                     addressed cell's ``body.text`` and refuses
                     (LIVE_OP_FAILED) unless it equals
                     ``expected_before_text``, so a co-author's edit
                     between ``cell_get`` and this op is never silently
                     overwritten. ``paragraphs`` is a list of paragraphs,
                     each a list of runs ``{text, bold, italic, link,
                     hard_break}``. Word calls: ``cell.body.clear()``, then
                     per paragraph ``paragraph.insertText(run.text,
                     "End")`` with ``range.font.bold``/``.italic`` and
                     ``range.hyperlink`` set per run, under the same
                     optional ``changeTrackingMode`` toggle as ``replace``.
                     Reply result: ``{applied, before, after, pre, post}``
                     -- ``before``/``after`` are the cell's own
                     ``body.text`` read before/after; ``pre``/``post`` are
                     the whole-body SHA-256 as for ``replace``.
                     Issue #75 (pane capability ``"cell_multiline"``):
                     ``cell_get`` also returns ``paragraphs`` (one string
                     per paragraph); paragraphs 2+ are written via
                     ``paragraphs.getLast()`` and a hard break via
                     ``insertBreak(line, "After")`` (Word rejects "End").
                     Paragraphs 2+ are written by inserting the empty
                     paragraphs, syncing, and writing into the loaded
                     items (fallback: chained ``insertParagraph("After")``).
                     After writing, the pane re-reads the paragraphs; on a
                     mismatch it retries the other strategy, then rewrites
                     the old paragraphs as plain text (untracked writes
                     only; never OOXML) and refuses with
                     ``OP_ERROR_CELL_WRITE_ROLLED_BACK`` or, if the restore
                     did not take (or the write was tracked),
                     ``OP_ERROR_CELL_WRITE_NOT_ROLLED_BACK``.

  table_get       -- ``TableGetPayload`` (table_index, 1-based, numbered like
                     ``list_tables``' ``table_id``). Issue #34; REQUIRES the
                     pane's ``"table_edit"`` capability. Word calls:
                     ``body.tables`` (``nestingLevel``), the addressed
                     ``Table`` (``style``, ``headerRowCount``,
                     ``getRange().getOoxml()`` scanned for ``w:gridSpan`` /
                     ``w:vMerge``), then per cell ``body.text``,
                     ``shadingColor``, ``horizontalAlignment``,
                     ``verticalAlignment``, ``columnWidth`` and
                     ``body.font`` (bold/color/size). Alignments are
                     reported in the SERVER's vocabulary (left/center/right/
                     both, top/center/bottom), never Word's enum names.
                     A document containing a NESTED table is refused.
                     Reply result: ``{style, headerRowCount, merged,
                     rows: [[{text, fill, align, valign, width, bold,
                     color, size}]]}`` (fill/color are "#RRGGBB" or "";
                     bold/size are null when mixed).

  table_insert    -- ``TableInsertPayload`` (anchor, rows, style,
                     style_builtin, style_from_table_index, header_rows,
                     column_widths_pt, font_size_pt, track_changes,
                     expectedBodySha256). Issue #34; REQUIRES ``"table_edit"``.
                     The pane compares ``expectedBodySha256`` with its own
                     body hash BEFORE any write and refuses with pane code
                     ``stale`` on a mismatch (the session's post-hoc
                     ``result.pre`` check cannot undo a write that already
                     landed). The style is resolved and validated BEFORE
                     the table is inserted (pane code ``style_not_found``),
                     so there is never an insert-then-delete rollback. The
                     anchor is a unique top-level paragraph (``paragraph_text``
                     + ``position`` before/after; 0 matches ``zero_match``,
                     >1 ``match_count_mismatch``), ``after_table_index``, or
                     absent (append to the body end). Word calls:
                     ``paragraph.insertTable`` / ``table.insertTable`` /
                     ``body.insertTable`` (WordApi 1.3), ``table.style`` /
                     ``styleBuiltIn`` / ``headerRowCount`` / ``font.size``,
                     per cell ``TableCell.shadingColor`` / ``horizontalAlignment``
                     / ``verticalAlignment`` / ``columnWidth`` plus the same
                     run writes as ``cell_set``; the new table's own index
                     is found by ``getRange().compareLocationWith`` (a
                     ``ClientResult``: ``.value`` is read after sync). Reply
                     result: ``{applied, table_index, pre, post}``.

  paragraph_insert -- (anchor, expected_matches, position, paragraphs
                     [{text, style, color}], track_changes,
                     expectedBodySha256). Issue #77; REQUIRES the pane's
                     ``"paragraph_insert"`` capability. The pane refuses
                     BEFORE any write with: ``stale`` (body hash moved),
                     ``STYLE_NOT_FOUND`` / ``UNSUPPORTED_STYLE_TYPE`` (a
                     style the document lacks, or a non-paragraph style),
                     ``ZERO_MATCH`` / ``MATCH_COUNT_MISMATCH`` (the anchor
                     must match exactly ``expected_matches``),
                     ``STRUCTURAL_BOUNDARY`` (anchor in a table cell) and
                     ``OUTSIDE_LOCKED_SECTION`` (``before`` under section
                     locks). Word calls: ``Paragraph.insertParagraph``
                     (the first from the anchor paragraph, each later one
                     chained "After" the previous so every write lands in a
                     synced proxy), ``Paragraph.style`` / ``styleBuiltIn``
                     (``styleBuiltIn = Normal`` when no style is given, so a
                     heading anchor's style is not inherited) and
                     ``font.color``. Every inserted paragraph's text, style
                     and color, and the whole paragraph list, are read back.
                     A mismatch is pane code ``VERIFICATION_FAILED``: an
                     untracked write is deleted again (the message says
                     whether that restored the list); a tracked write is
                     left for the Review pane. Reply result: ``{applied,
                     position, anchor_index, inserted: [{text, style,
                     style_builtin, color, requested_style: {kind, value}}],
                     before_count, after_count, pre, post}``.

  cells_set       -- ``CellsSetPayload`` (cells, track_changes,
                     expectedBodySha256). Issue #34; REQUIRES ``"table_edit"``.
                     ``cells`` is a list of ``{table_index, row_index,
                     cell_index, paragraphs, expected_before_text}``. The pane
                     checks the body hash, reads every addressed cell, then
                     READS THEM ALL AGAIN in the batch right before writing;
                     any mismatch refuses with nothing written. The writes
                     are one Word batch, which is NOT a transaction: this is
                     best-effort under co-authoring. Reply result:
                     ``{applied, before: [..], after: [..], pre, post}``.

  save            -- no payload. Word calls: ``context.document.save()``.
                     Reply result: ``{"saved": true}``.

Staleness (``expected_body_sha256``)
-------------------------------------
``search``/``replace``/``format``/``comment_add``/``cell_set``/
``table_insert``/``cells_set``/``paragraph_insert`` payloads
carry an optional ``expected_body_sha256``. The pane does not act on it -- it is
read by ``live/session.py``'s ``LiveSession.request`` after the reply
comes back: if the caller supplied one and the reply's ``result["pre"]``
(the pane's body hash *before* it ran the op) differs, the caller gets
``LiveStale`` instead of trusting a result computed against text the
caller no longer expects.
"""

from __future__ import annotations

import dataclasses
import json
import math
import re
from typing import Any


class ProtocolError(ValueError):
    """A message did not match its expected wire shape."""


VALID_OPS: frozenset[str] = frozenset(
    {
        "ping",
        "describe",
        "body_ooxml",
        "search",
        "replace",
        "format",
        "comments_list",
        "comment_add",
        "comment_reply",
        "comment_resolve",
        "cell_get",
        "cell_set",
        "table_get",
        "table_insert",
        "cells_set",
        "paragraph_delete",
        "paragraph_insert",
        "shapes_list",
        "textboxes_list",
        "textboxes_read",
        "scope_describe",
        "revisions_list",
        "revisions_accept",
        "revisions_reject",
        "save",
        # issue #66: tag a document so Word opens the Live pane with it
        "autoopen_get",
        "autoopen_set",
        # issue #47: heading sections (key, level, hash) for section locks
        "sections_list",
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
# Issue #34: the pane refused BEFORE writing because its own body hash no
# longer equals the caller's expectedBodySha256 (maps to LIVE_STALE), and a
# table style that does not exist (maps to STYLE_NOT_FOUND).
# Issue #75: cell_set's pane-side self-check failed; the cell was restored /
# could not be restored (both map to VERIFICATION_FAILED server-side).
OP_ERROR_CELL_WRITE_ROLLED_BACK = "cell_write_rolled_back"
OP_ERROR_CELL_WRITE_NOT_ROLLED_BACK = "cell_write_not_rolled_back"
OP_ERROR_STALE = "stale"
OP_ERROR_STYLE_NOT_FOUND = "style_not_found"

# `body_ooxml` refusal (issue #33): the document is too large to ship as one
# websocket frame even with binaries stripped. Same lowercase convention.
OP_ERROR_TOO_LARGE = "too_large"


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

    ``capabilities`` (issue #22 B2, optional -- defaults to ``[]``):
    op-level feature names THIS pane build implements beyond the
    baseline op set, e.g. ``"row_scope"`` (``rowAnchor`` on
    ``replace``/``format``). Distinct from ``requirement_sets``
    (WordApi version support the HOST reports) -- a pane can run on a
    WordApi version new enough for a feature's underlying Office JS
    calls while still running an OLDER BUILD of ``taskpane.js`` that
    never implements the wire-level payload field for it. An
    already-connected pane from before a capability existed sends no
    ``capabilities`` field at all (a real gap this defends against: an
    old pane silently ignoring an unknown payload key and running an
    UNSCOPED op instead of refusing) -- ``from_json`` defaults it to
    ``[]`` rather than requiring it, so that pane still connects; a
    caller that needs a specific capability checks for its presence
    before sending an op that depends on it
    (``live/write_mode.py``'s ``require_capability``), rather than the
    hello parse itself refusing.
    """

    document_url: str
    host: str
    platform: str
    requirement_sets: dict[str, Any]
    body_sha256: str
    capabilities: frozenset[str] = frozenset()
    # issue #39: per-load pane id (optional; an older pane sends none), so a
    # response / live_status can say WHICH pane instance answered.
    instance_id: str = ""

    def to_json(self) -> dict[str, Any]:
        return {
            "type": "hello",
            "documentUrl": self.document_url,
            "host": self.host,
            "platform": self.platform,
            "requirementSets": self.requirement_sets,
            "bodySha256": self.body_sha256,
            "capabilities": sorted(self.capabilities),
            "instanceId": self.instance_id,
        }

    @classmethod
    def from_json(cls, raw: str | bytes | dict[str, Any]) -> HelloMessage:
        obj = _as_dict(raw)
        msg_type = _require(obj, "type", str)
        if msg_type != "hello":
            raise ProtocolError(f"expected type 'hello', got {msg_type!r}")
        capabilities_raw = _optional(obj, "capabilities", list, default=[])
        return cls(
            document_url=_require(obj, "documentUrl", str),
            host=_require(obj, "host", str),
            platform=_require(obj, "platform", str),
            requirement_sets=_require(obj, "requirementSets", dict),
            body_sha256=_require(obj, "bodySha256", str),
            capabilities=frozenset(capabilities_raw),
            instance_id=_optional(obj, "instanceId", str, default=""),
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
    row_anchor: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "expected_matches": self.expected_matches,
            "replace": self.replace,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        if self.row_anchor is not None:
            out["rowAnchor"] = self.row_anchor
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
            row_anchor=_optional(obj, "rowAnchor", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class FormatPayload:
    find: str
    expected_matches: int
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: str | None = None
    track_changes: bool = False
    expected_body_sha256: str | None = None
    row_anchor: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "find": self.find,
            "expected_matches": self.expected_matches,
            "bold": self.bold,
            "italic": self.italic,
            "underline": self.underline,
            "strike": self.strike,
            "color": self.color,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        if self.row_anchor is not None:
            out["rowAnchor"] = self.row_anchor
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
        strike = _optional(obj, "strike", bool, default=None)
        color = _optional(obj, "color", str, default=None)
        if bold is None and italic is None and underline is None and strike is None and color is None:
            raise ProtocolError("at least one of bold/italic/underline/strike/color must be set")
        return cls(
            find=find,
            expected_matches=expected_matches,
            bold=bold,
            italic=italic,
            underline=underline,
            strike=strike,
            color=color,
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
            row_anchor=_optional(obj, "rowAnchor", str, default=None),
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


def _require_index(obj: dict[str, Any], key: str) -> int:
    value = _require(obj, key, int)
    if value < 1:
        raise ProtocolError(f"{key!r} must be >= 1 (1-based)")
    return value


@dataclasses.dataclass(frozen=True)
class CellGetPayload:
    table_index: int
    row_index: int
    cell_index: int

    def to_json(self) -> dict[str, Any]:
        return {
            "table_index": self.table_index,
            "row_index": self.row_index,
            "cell_index": self.cell_index,
        }

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CellGetPayload:
        return cls(
            table_index=_require_index(obj, "table_index"),
            row_index=_require_index(obj, "row_index"),
            cell_index=_require_index(obj, "cell_index"),
        )


@dataclasses.dataclass(frozen=True)
class CellSetPayload:
    table_index: int
    row_index: int
    cell_index: int
    paragraphs: list[list[dict[str, Any]]]
    expected_before_text: str
    track_changes: bool = False
    expected_body_sha256: str | None = None

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "table_index": self.table_index,
            "row_index": self.row_index,
            "cell_index": self.cell_index,
            "paragraphs": self.paragraphs,
            "expected_before_text": self.expected_before_text,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CellSetPayload:
        paragraphs = _require(obj, "paragraphs", list)
        for paragraph in paragraphs:
            if not isinstance(paragraph, list) or not all(isinstance(run, dict) for run in paragraph):
                raise ProtocolError("'paragraphs' must be a list of lists of run objects")
        return cls(
            table_index=_require_index(obj, "table_index"),
            row_index=_require_index(obj, "row_index"),
            cell_index=_require_index(obj, "cell_index"),
            paragraphs=paragraphs,
            expected_before_text=_require(obj, "expected_before_text", str),
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


# ---------------------------------------------------------------------------
# Live tables (issue #34). Validation happens in __post_init__ so a payload
# built by the server is checked BEFORE anything is sent, not only when a
# pane parses it.
# ---------------------------------------------------------------------------

_HEX6_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_ALIGNS = frozenset({"left", "center", "right", "both"})
_VALIGNS = frozenset({"top", "center", "bottom"})
_POSITIONS = frozenset({"before", "after"})
_CELL_KEYS = frozenset({"paragraphs", "fill", "color", "bold", "align", "valign"})


def _is_finite_positive(value: Any) -> bool:
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(value)
        and value > 0
    )


def _validate_paragraphs(paragraphs: Any, where: str) -> None:
    if not isinstance(paragraphs, list):
        raise ProtocolError(f"{where}: 'paragraphs' must be a list")
    for paragraph in paragraphs:
        if not isinstance(paragraph, list) or not all(isinstance(run, dict) for run in paragraph):
            raise ProtocolError(f"{where}: 'paragraphs' must be a list of lists of run objects")
        for run in paragraph:
            if not isinstance(run.get("text", ""), str):
                raise ProtocolError(f"{where}: a run's 'text' must be a string")
            for flag in ("bold", "italic", "hard_break"):
                if run.get(flag) is not None and not isinstance(run[flag], bool):
                    raise ProtocolError(f"{where}: a run's {flag!r} must be a bool")
            if run.get("link") is not None and not isinstance(run["link"], str):
                raise ProtocolError(f"{where}: a run's 'link' must be a string")


@dataclasses.dataclass(frozen=True)
class TableGetPayload:
    table_index: int

    def to_json(self) -> dict[str, Any]:
        return {"table_index": self.table_index}

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> TableGetPayload:
        return cls(table_index=_require_index(obj, "table_index"))


@dataclasses.dataclass(frozen=True)
class TableInsertPayload:
    rows: list[list[dict[str, Any]]]
    anchor: dict[str, Any] | None = None
    style: str | None = None
    style_builtin: str | None = None
    style_from_table_index: int | None = None
    header_rows: int = 0
    column_widths_pt: list[float] | None = None
    font_size_pt: float | None = None
    track_changes: bool = False
    expected_body_sha256: str | None = None

    def __post_init__(self) -> None:
        if not self.rows or any(not isinstance(r, list) or not r for r in self.rows):
            raise ProtocolError("'rows' must be a non-empty list of non-empty cell lists")
        width = len(self.rows[0])
        for r, row in enumerate(self.rows):
            if len(row) != width:
                raise ProtocolError(f"'rows' must be rectangular: row {r} has {len(row)} cell(s), row 0 has {width}")
            for c, cell in enumerate(row):
                self._validate_cell(cell, f"rows[{r}][{c}]")
        self._validate_anchor()
        sources = [s for s in (self.style, self.style_builtin, self.style_from_table_index) if s is not None]
        if len(sources) > 1:
            raise ProtocolError("at most one of 'style', 'style_builtin', 'style_from_table_index'")
        if self.style_from_table_index is not None and (
            isinstance(self.style_from_table_index, bool)
            or not isinstance(self.style_from_table_index, int)
            or self.style_from_table_index < 1
        ):
            raise ProtocolError("'style_from_table_index' must be an int >= 1")
        if (
            isinstance(self.header_rows, bool)
            or not isinstance(self.header_rows, int)
            or not 0 <= self.header_rows < len(self.rows)
        ):
            raise ProtocolError(f"'header_rows' must be an int in [0, {len(self.rows) - 1}]")
        if self.column_widths_pt is not None and (
            len(self.column_widths_pt) != width or not all(_is_finite_positive(w) for w in self.column_widths_pt)
        ):
            raise ProtocolError(f"'column_widths_pt' must be {width} positive finite number(s)")
        if self.font_size_pt is not None and not _is_finite_positive(self.font_size_pt):
            raise ProtocolError("'font_size_pt' must be a positive finite number")

    @staticmethod
    def _validate_cell(cell: Any, where: str) -> None:
        if not isinstance(cell, dict):
            raise ProtocolError(f"{where} must be an object")
        unknown = set(cell) - _CELL_KEYS
        if unknown:
            raise ProtocolError(f"{where} has unknown key(s) {sorted(unknown)}")
        _validate_paragraphs(cell.get("paragraphs", []), where)
        for key in ("fill", "color"):
            value = cell.get(key)
            if value is not None and not (isinstance(value, str) and _HEX6_RE.match(value)):
                raise ProtocolError(f"{where}: {key!r} must be a 6-digit hex color")
        if cell.get("bold") is not None and not isinstance(cell["bold"], bool):
            raise ProtocolError(f"{where}: 'bold' must be a bool")
        if cell.get("align") is not None and cell["align"] not in _ALIGNS:
            raise ProtocolError(f"{where}: 'align' must be one of {sorted(_ALIGNS)}")
        if cell.get("valign") is not None and cell["valign"] not in _VALIGNS:
            raise ProtocolError(f"{where}: 'valign' must be one of {sorted(_VALIGNS)}")

    def _validate_anchor(self) -> None:
        anchor = self.anchor
        if anchor is None:
            return
        if not isinstance(anchor, dict):
            raise ProtocolError("'anchor' must be an object")
        has_para = "paragraph_text" in anchor
        has_table = "after_table_index" in anchor
        if has_para == has_table:
            raise ProtocolError("'anchor' needs exactly one of 'paragraph_text' or 'after_table_index'")
        unknown = set(anchor) - {"paragraph_text", "position", "after_table_index"}
        if unknown:
            raise ProtocolError(f"'anchor' has unknown key(s) {sorted(unknown)}")
        if has_para:
            if not isinstance(anchor["paragraph_text"], str) or not anchor["paragraph_text"].strip():
                raise ProtocolError("'anchor.paragraph_text' must be a non-empty string")
            if anchor.get("position") not in _POSITIONS:
                raise ProtocolError(f"'anchor.position' must be one of {sorted(_POSITIONS)}")
        else:
            _require_index(anchor, "after_table_index")
            if "position" in anchor:
                raise ProtocolError("'anchor.position' is only valid with 'paragraph_text'")

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "rows": self.rows,
            "anchor": self.anchor,
            "style": self.style,
            "style_builtin": self.style_builtin,
            "style_from_table_index": self.style_from_table_index,
            "header_rows": self.header_rows,
            "column_widths_pt": self.column_widths_pt,
            "font_size_pt": self.font_size_pt,
            "track_changes": self.track_changes,
        }
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> TableInsertPayload:
        return cls(
            rows=_require(obj, "rows", list),
            anchor=_optional(obj, "anchor", dict, default=None),
            style=_optional(obj, "style", str, default=None),
            style_builtin=_optional(obj, "style_builtin", str, default=None),
            style_from_table_index=_optional(obj, "style_from_table_index", int, default=None),
            header_rows=_optional(obj, "header_rows", int, default=0),
            column_widths_pt=_optional(obj, "column_widths_pt", list, default=None),
            font_size_pt=_optional(obj, "font_size_pt", (int, float), default=None),
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


@dataclasses.dataclass(frozen=True)
class CellsSetPayload:
    cells: list[dict[str, Any]]
    track_changes: bool = False
    expected_body_sha256: str | None = None

    def __post_init__(self) -> None:
        if not self.cells:
            raise ProtocolError("'cells' must not be empty")
        seen: set[tuple[int, int, int]] = set()
        for i, cell in enumerate(self.cells):
            if not isinstance(cell, dict):
                raise ProtocolError(f"cells[{i}] must be an object")
            address = (
                _require_index(cell, "table_index"),
                _require_index(cell, "row_index"),
                _require_index(cell, "cell_index"),
            )
            if address in seen:
                raise ProtocolError(f"cells[{i}] repeats address {address}")
            seen.add(address)
            _require(cell, "expected_before_text", str)
            _validate_paragraphs(_require(cell, "paragraphs", list), f"cells[{i}]")

    def to_json(self) -> dict[str, Any]:
        out: dict[str, Any] = {"cells": self.cells, "track_changes": self.track_changes}
        if self.expected_body_sha256 is not None:
            out["expectedBodySha256"] = self.expected_body_sha256
        return out

    @classmethod
    def from_json(cls, obj: dict[str, Any]) -> CellsSetPayload:
        return cls(
            cells=_require(obj, "cells", list),
            track_changes=_optional(obj, "track_changes", bool, default=False),
            expected_body_sha256=_optional(obj, "expectedBodySha256", str, default=None),
        )


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
