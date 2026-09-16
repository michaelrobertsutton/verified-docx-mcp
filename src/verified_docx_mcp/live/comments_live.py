"""Live-mode comment tools (issue #106 WP-4:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

``write_mode="live"`` for ``add_anchored_comment``/``reply_to_comment``/
``resolve_comment``, and ``source="live"`` for ``list_open_items`` --
going through the WSS bridge (``live/bridge.py``, ``live/session.py``,
``live/protocol.py``, all WP-2) instead of the OOXML ``.docx`` parts
``comments.py``/``tracked_changes.py`` (file mode) read and write
directly. ``server.py``'s tool wrappers call the functions below when
they resolve to live mode; file mode is untouched, byte for byte.

Correlation (the bridge between the two id spaces)
----------------------------------------------------
WP-1's spike result, recorded in ``docs/live-mode.md``: Office.js's
``Comment.id`` is UNRELATED to the OOXML ``durableId``/``w:id`` a
document's own ``word/commentsIds.xml``/``word/comments.xml`` carry --
the two numbering schemes are independent, and a live comment's id (a
``live:<Comment.id>`` handle here) is only ever meaningful within the one
live session that created it; it does not survive a Word restart. So a
caller holding a durableId (from an earlier file-mode call, or from a
lead pasting one in) needs a way to reach the SAME comment through the
live path, and ``list_open_items(source="live")`` needs a way to tell the
caller which durableId a given live comment probably corresponds to.
Both directions go through ``correlate_comments`` below: for each live
comment, the best-matching file-mode comment (read from the same
document's on-disk snapshot, via ``tracked_changes._parse_comments`` --
the same function file-mode ``list_open_items`` itself uses) by
normalized anchor text + normalized content, tie-broken by author and
creation date when both sides have them. This is advisory, never
authoritative -- a caller that already has a ``live:<id>`` handle should
always prefer it; correlation only exists to make a durableId/`w:id`
usable when that is all a caller has.

Normalization
-------------
- Anchor text: collapse all whitespace runs (including the pane's own
  paragraph-boundary whitespace) to single spaces and strip the ends --
  ``" ".join(text.split())``.
- Content: multi-paragraph content differs in shape between the two
  sides for the exact reason WP-1 flagged -- the pane's own
  ``Comment.content``/``getReplies()`` join a multi-paragraph comment's
  paragraphs with ``\\r``, while ``tracked_changes.execute_list_open_items``
  (and ``comments._comment_record``) join the SAME paragraphs with no
  separator at all. Normalizing means dropping ``\\r``/``\\n`` outright on
  both sides (never replacing with a space -- the file side already has
  no space between "comment." and "Second", so introducing one would
  break the very case this normalization exists to fix), not a general
  whitespace collapse.

Confidence
----------
- ``exact``   -- normalized content matches AND normalized anchor text
  matches AND (author matches, when both sides report one) AND (creation
  dates are within 2 minutes of each other, when both sides report one).
- ``content-only`` -- normalized content matches but the anchor text (or
  author/date, when present) does not corroborate it.
- ``none``    -- no file-mode comment's normalized content matches this
  live comment's at all.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import Any

from . import bridge as live_bridge
from .protocol import OP_ERROR_MATCH_COUNT_MISMATCH, OP_ERROR_ZERO_MATCH
from .session import (
    LiveDisconnected,
    LiveOpFailed,
    LiveSession,
    LiveStale,
    LiveUnavailable,
)

VALID_WRITE_MODES = ("auto", "file", "live")
VALID_SOURCES = ("auto", "file", "live")

_CORRELATION_DATE_TOLERANCE = timedelta(minutes=2)
_LIVE_HANDLE_PREFIX = "live:"


# ---------------------------------------------------------------------------
# Write-mode / source resolution and session lookup.
#
# The write-mode resolver and session lookup live in live/write_mode.py (WP-3);
# the thin wrappers below keep the comment tools on that single rule.


def _document_name_and_path(path: str) -> tuple[Any, str]:
    from .. import paths as _paths

    resolved = _paths.resolve_allowed_docx_path(path, must_exist=True)
    return resolved, resolved.name


class _NoSessions:
    """Stand-in registry when the bridge is not running: every lookup misses."""

    def get(self, _document_name: str):
        return None


def _registry():
    # Never start_in_background() here: an ordinary file-mode call must not
    # bind the bridge's real ports as a side effect (same rule as
    # live/write_mode.py). Only live_status, or an already-connected pane,
    # brings the bridge up.
    return live_bridge.current_registry() or _NoSessions()


def _make_error(code, message: str, diagnostics: dict[str, Any]):
    from ..errors import _make_error as _make

    return _make(code, message, diagnostics)


def _resolve_write_mode(path: str, requested: str) -> str:
    """Delegates to the shared resolver (live/write_mode.py, WP-3) so the
    comment tools and the text tools apply one identical "auto" rule."""
    from . import write_mode as _wm

    return _wm.resolve_write_mode(path, requested)


# list_open_items' own read rule (live whenever a pane session exists; reads never refuse).
def _resolve_source(path: str, requested: str) -> str:
    from ..errors import ErrorCode

    if requested not in VALID_SOURCES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"source must be one of {VALID_SOURCES}, got {requested!r}",
            {"source": requested},
        )
    if requested in ("file", "live"):
        if requested == "live":
            _resolved, document_name = _document_name_and_path(path)
            if _registry().get(document_name) is None:
                raise _make_error(
                    ErrorCode.LIVE_UNAVAILABLE,
                    f"source='live' requested but no connected pane session for document {document_name!r}",
                    {"document_name": document_name},
                )
        return requested

    # "auto": a read never refuses, so this is simpler than write_mode's
    # own "auto" -- live whenever a pane session is connected for this
    # document (a lock-owner check would only ever narrow a READ, which
    # core/document-backend-protocol.md never asks of a read tool).
    _resolved, document_name = _document_name_and_path(path)
    return "live" if _registry().get(document_name) is not None else "file"


def _session_for(path: str) -> LiveSession:
    """Delegates to live.write_mode.live_session_for (WP-3)."""
    from . import write_mode as _wm

    return _wm.live_session_for(path)


def _raise_from_live_error(exc: BaseException) -> None:
    """Map a live/session.py exception onto the matching errors.py
    ErrorCode and raise it (never returns) -- server.py's live tool
    wrappers let this propagate as the VerifyError it already is."""
    from ..errors import ErrorCode

    if isinstance(exc, LiveUnavailable):
        raise _make_error(ErrorCode.LIVE_UNAVAILABLE, str(exc), {}) from exc
    if isinstance(exc, LiveDisconnected):
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc), {}) from exc
    if isinstance(exc, LiveStale):
        raise _make_error(
            ErrorCode.LIVE_STALE, str(exc), {"expected": exc.expected, "actual": exc.actual}
        ) from exc
    if isinstance(exc, LiveOpFailed):
        code = exc.code
        if code == OP_ERROR_ZERO_MATCH:
            raise _make_error(ErrorCode.ZERO_MATCH, exc.message, {"pane_error_code": code}) from exc
        if code == OP_ERROR_MATCH_COUNT_MISMATCH:
            raise _make_error(ErrorCode.MATCH_COUNT_MISMATCH, exc.message, {"pane_error_code": code}) from exc
        raise _make_error(ErrorCode.LIVE_OP_FAILED, exc.message, {"pane_error_code": code}) from exc
    raise exc


def _request(session: LiveSession, op: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    try:
        return session.request_threadsafe(op, payload or {})
    except (LiveUnavailable, LiveDisconnected, LiveOpFailed, LiveStale) as exc:
        _raise_from_live_error(exc)
        raise  # unreachable; _raise_from_live_error always raises


# ---------------------------------------------------------------------------
# Correlation.
# ---------------------------------------------------------------------------


def _normalize_anchor(text: str | None) -> str:
    return " ".join((text or "").split())


def _normalize_content(text: str | None) -> str:
    # Drop paragraph separators outright (never replace with a space) --
    # see this module's own docstring: file mode's own paragraph join has
    # no separator at all, so a dropped "\r"/"\n" is what makes the two
    # sides comparable, not a collapsed one.
    return (text or "").replace("\r", "").replace("\n", "")


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    text = value.strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt


def _dates_corroborate(live_date: str | None, file_date: str | None) -> bool:
    live_dt, file_dt = _parse_date(live_date), _parse_date(file_date)
    if live_dt is None or file_dt is None:
        return True  # unknown on either side never disqualifies a match
    return abs(live_dt - file_dt) <= _CORRELATION_DATE_TOLERANCE


def _authors_corroborate(live_author: str | None, file_author: str | None) -> bool:
    if not live_author or not file_author:
        return True
    return live_author == file_author


def correlate_comments(
    live_comments_raw: list[dict[str, Any]], file_comments: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """For each raw live comment (the wire shape ``comments_list`` returns
    -- ``id``, ``content``, ``authorName``, ``creationDate``, ``anchorText``)
    find the best-matching entry in *file_comments* (``tracked_changes.
    _parse_comments``'s own shape -- ``comment_id`` (durableId, or a raw
    ``w:id`` fallback), ``w_id``, ``content``, ``quoted_text``, ``author``,
    ``created_time``).

    Returns one entry per live comment: ``{live_comment_id: "live:<id>",
    comment_id, w_id, confidence}`` -- ``comment_id``/``w_id`` are ``None``
    when ``confidence`` is ``"none"``. Advisory only: never used to gate
    anything, only to let a caller holding a durableId/w:id reach the live
    comment it probably corresponds to (see ``_resolve_comment_handle``).
    """
    out: list[dict[str, Any]] = []
    for live in live_comments_raw:
        live_anchor = _normalize_anchor(live.get("anchorText"))
        live_content = _normalize_content(live.get("content"))
        live_author = live.get("authorName")
        live_date = live.get("creationDate")

        best: dict[str, Any] | None = None
        best_confidence = "none"
        if live_content:
            for f in file_comments:
                if _normalize_content(f.get("content")) != live_content:
                    continue
                anchor_ok = bool(live_anchor) and live_anchor == _normalize_anchor(f.get("quoted_text"))
                author_ok = _authors_corroborate(live_author, f.get("author"))
                date_ok = _dates_corroborate(live_date, f.get("created_time"))
                if anchor_ok and author_ok and date_ok:
                    best, best_confidence = f, "exact"
                    break  # an exact match is the best this can do
                if best is None:
                    best, best_confidence = f, "content-only"

        out.append(
            {
                "live_comment_id": f"{_LIVE_HANDLE_PREFIX}{live.get('id')}",
                "comment_id": best.get("comment_id") if best else None,
                "w_id": best.get("w_id") if best else None,
                "confidence": best_confidence,
            }
        )
    return out


def _resolve_comment_handle(comment_id: str, correlation: list[dict[str, Any]]) -> tuple[str, str]:
    """*comment_id* is either a ``live:<id>`` handle (used as-is) or a
    durableId/``w:id`` resolved through *correlation* (exact or
    content-only confidence only -- "none" entries carry no id to match
    against). Returns (live_handle_without_prefix, resolved_via) where
    resolved_via is ``"live-handle"`` | ``"durableId-correlation"`` |
    ``"w_id-correlation"``. Raises INVALID_INPUT, naming both id spaces,
    when nothing matches."""
    from ..errors import ErrorCode

    if comment_id.startswith(_LIVE_HANDLE_PREFIX):
        return comment_id[len(_LIVE_HANDLE_PREFIX) :], "live-handle"

    for entry in correlation:
        if entry["confidence"] not in ("exact", "content-only"):
            continue
        if entry["comment_id"] is not None and entry["comment_id"] == comment_id:
            return entry["live_comment_id"][len(_LIVE_HANDLE_PREFIX) :], "durableId-correlation"
        if entry["w_id"] is not None and entry["w_id"] == comment_id:
            return entry["live_comment_id"][len(_LIVE_HANDLE_PREFIX) :], "w_id-correlation"

    available_durable = sorted({e["comment_id"] for e in correlation if e["comment_id"]})
    available_live = sorted(e["live_comment_id"] for e in correlation)
    raise _make_error(
        ErrorCode.INVALID_INPUT,
        f"comment_id {comment_id!r} does not match a live:<id> handle or any correlated durableId/w:id "
        "for the connected pane session; call list_open_items(source='live') for the current live ids "
        "and their correlation.",
        {
            "comment_id": comment_id,
            "available_live_handles": available_live,
            "available_durable_ids": available_durable,
        },
    )


# ---------------------------------------------------------------------------
# Snapshot reads (file-mode comments/suggestions for the SAME document, used
# both to build the correlation table and to keep list_open_items(source=
# "live")'s pending_suggestions shape identical to file mode's).
# ---------------------------------------------------------------------------


def _file_comments_and_suggestions(path: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    from .. import mutations, projection, tracked_changes
    from .. import server as _server

    resolved, _name = _document_name_and_path(path)
    local_path, is_temp = _server._read_local_copy(resolved)
    try:
        proj = projection.project_part(local_path)
        document_root, _raw = mutations._load_document(local_path)
        suggestions = tracked_changes._find_pending_suggestions(document_root)
        file_comments = tracked_changes._parse_comments(local_path, proj)
        return file_comments, suggestions
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


def _live_comment_record(raw: dict[str, Any]) -> dict[str, Any]:
    anchor_text = raw.get("anchorText") or ""
    replies = [
        {
            "comment_id": f"{_LIVE_HANDLE_PREFIX}{r.get('id')}",
            "content": r.get("content", ""),
            "author": r.get("authorName"),
            "created_time": r.get("creationDate"),
        }
        for r in (raw.get("replies") or [])
    ]
    return {
        "comment_id": f"{_LIVE_HANDLE_PREFIX}{raw.get('id')}",
        "w_id": None,
        "content": raw.get("content", ""),
        "resolved": bool(raw.get("resolved", False)),
        "reply_count": len(replies),
        "replies": replies,
        "quoted_text": anchor_text,
        "anchor_text": anchor_text,
        "author": raw.get("authorName"),
        "created_time": raw.get("creationDate"),
        "modified_time": "",
        "scope": "document",
    }


def _live_state(path: str, session: LiveSession) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """(raw_live_comments, correlation) for every comment the pane
    currently reports (resolved included) -- callers filter resolved
    themselves per their own needs (list_open_items excludes it;
    handle resolution does not care)."""
    result = _request(session, "comments_list")
    raw_comments = result.get("comments") or []
    file_comments, _suggestions = _file_comments_and_suggestions(path)
    correlation = correlate_comments(raw_comments, file_comments)
    return raw_comments, correlation


# ---------------------------------------------------------------------------
# list_open_items(source="live")
# ---------------------------------------------------------------------------


def execute_list_open_items_live(path: str) -> dict[str, Any]:
    resolved, _name = _document_name_and_path(path)
    session = _session_for(path)
    raw_comments, correlation = _live_state(path, session)
    _file_comments, suggestions = _file_comments_and_suggestions(path)

    # list_open_items filters resolved the same way file mode does (an
    # "open item" is, by definition, not a resolved one) -- but the pane
    # sees resolved comments too, so this filter has to happen here
    # rather than being naturally absent from what the pane reports.
    open_raw = [c for c in raw_comments if not c.get("resolved", False)]
    open_ids = {f"{_LIVE_HANDLE_PREFIX}{c.get('id')}" for c in open_raw}

    return {
        "path": str(resolved),
        "source": "live",
        "comments": [_live_comment_record(c) for c in open_raw],
        "pending_suggestions": suggestions,
        "correlation": [entry for entry in correlation if entry["live_comment_id"] in open_ids],
    }


def execute_list_open_items(path: str, source: str = "auto") -> dict[str, Any]:
    from .. import tracked_changes

    mode = _resolve_source(path, source)
    if mode == "live":
        return execute_list_open_items_live(path)
    result = tracked_changes.execute_list_open_items(path)
    result["source"] = "file"
    return result


# ---------------------------------------------------------------------------
# add_anchored_comment(write_mode="live")
# ---------------------------------------------------------------------------


def execute_add_anchored_comment_live(
    path: str, quote: str, text: str, expected_matches: int
) -> dict[str, Any]:
    from .. import audit

    resolved, document_name = _document_name_and_path(path)
    session = _session_for(path)

    result = _request(
        session, "comment_add", {"find": quote, "expected_matches": expected_matches, "text": text}
    )
    live_id = result["comment_id"]
    comment_handle = f"{_LIVE_HANDLE_PREFIX}{live_id}"
    pre, post = result.get("pre"), result.get("post")

    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": expected_matches,
        "rung": "live",
        "before": quote,
        "after": quote,  # a comment never edits document text, live or otherwise
        "revision_before": f"live:sha256:{pre}" if pre else None,
        "revision_after": f"live:sha256:{post}" if post else None,
        "audit_logged": False,
        "comment_id": comment_handle,
        "comment_ids": [comment_handle],
        "orphaned_comment_ids": [],
        "write_mode": "live",
        "verified_via": "word-addin",
        "document_name": document_name,
        "author": "word-signed-in-user",
    }
    logged, _reason = audit.append_audit(path=str(resolved), tool="add_anchored_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# reply_to_comment(write_mode="live") / resolve_comment(write_mode="live")
# ---------------------------------------------------------------------------


def execute_reply_to_comment_live(path: str, comment_id: str, text: str) -> dict[str, Any]:
    from .. import audit

    resolved, document_name = _document_name_and_path(path)
    session = _session_for(path)

    raw_comments, correlation = _live_state(path, session)
    live_handle, resolved_via = _resolve_comment_handle(comment_id, correlation)
    parent_raw = next((c for c in raw_comments if c.get("id") == live_handle), None)
    parent_anchor = (parent_raw or {}).get("anchorText", "")

    _request(session, "comment_reply", {"comment_id": live_handle, "text": text})

    # Verify by re-listing rather than trusting the op's own ok=true --
    # same discipline as file mode's own post-write re-read.
    after_raw = _request(session, "comments_list").get("comments") or []
    parent_after = next((c for c in after_raw if c.get("id") == live_handle), None)
    reply_found = bool(parent_after) and any(
        r.get("content") == text for r in (parent_after.get("replies") or [])
    )
    if not reply_found:
        from ..errors import ErrorCode, _make_error

        raise _make_error(
            ErrorCode.LIVE_OP_FAILED,
            f"reply to live comment {comment_id!r} was not found after re-listing the pane's comments",
            {"comment_id": comment_id, "live_handle": f"{_LIVE_HANDLE_PREFIX}{live_handle}"},
        )

    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": 1,
        "rung": "reply",
        "before": parent_anchor,
        "after": parent_anchor,  # a reply never edits document text
        "revision_before": None,
        "revision_after": None,
        "audit_logged": False,
        "comment_id": f"{_LIVE_HANDLE_PREFIX}{live_handle}",
        "parent_comment_id": comment_id,
        "comment_id_resolved_via": resolved_via,
        "write_mode": "live",
        "verified_via": "word-addin",
        "document_name": document_name,
        "author": "word-signed-in-user",
    }
    logged, _reason = audit.append_audit(path=str(resolved), tool="reply_to_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


def execute_resolve_comment_live(path: str, comment_id: str) -> dict[str, Any]:
    from .. import audit
    from ..errors import ErrorCode, _make_error

    resolved, document_name = _document_name_and_path(path)
    session = _session_for(path)

    _raw_comments, correlation = _live_state(path, session)
    live_handle, resolved_via = _resolve_comment_handle(comment_id, correlation)

    _request(session, "comment_resolve", {"comment_id": live_handle, "resolved": True})

    # Independent post-op re-read -- the pane's own reply is not trusted
    # on its own, same discipline as file mode's own resolve_comment.
    after_raw = _request(session, "comments_list").get("comments") or []
    parent_after = next((c for c in after_raw if c.get("id") == live_handle), None)
    if parent_after is None or not parent_after.get("resolved", False):
        raise _make_error(
            ErrorCode.COMMENT_STILL_OPEN,
            f"comment_id {comment_id!r} is still open after the live resolve attempt",
            {"comment_id": comment_id, "live_handle": f"{_LIVE_HANDLE_PREFIX}{live_handle}"},
        )

    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": 1,
        "rung": "resolve",
        "before": "open",
        "after": "resolved",
        "revision_before": None,
        "revision_after": None,
        "audit_logged": False,
        "comment_id": f"{_LIVE_HANDLE_PREFIX}{live_handle}",
        "comment_id_resolved_via": resolved_via,
        "write_mode": "live",
        "verified_via": "word-addin",
        "document_name": document_name,
        "author": "word-signed-in-user",
    }
    logged, _reason = audit.append_audit(path=str(resolved), tool="resolve_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
