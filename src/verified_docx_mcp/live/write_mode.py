"""Shared ``write_mode`` resolver + live evidence builder (issue #106 WP-3:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

This module is imported by two parallel work packages against the same
WP-2 base (main ``4c93d42``): WP-3 (``replace_text``/``format_text``/
``live_save`` -- this PR) and WP-4 (the live comment tools, a separate
PR). It exists so both sides share exactly one ``auto`` selection rule
and one evidence shape instead of each reimplementing (and potentially
diverging on) the same logic. Per the plan, its public surface stays
small:

  ``resolve_write_mode(path, requested)`` -- decide ``"file"`` vs
      ``"live"`` for one call.
  ``live_session_for(path)`` -- the connected ``LiveSession`` for a
      document, or a typed ``LIVE_UNAVAILABLE`` error.
  ``live_evidence(...)`` -- build + audit-log the live-mode evidence
      envelope (the same eight keys as file mode, plus the live-only
      keys the plan names).

Plus two small helpers every live write op needs and that would
otherwise be duplicated in both PRs' own tool modules:

  ``check_not_stale(revision_before, current_body_sha256)`` -- the
      ``LIVE_STALE`` pre-flight check.
  ``classify_op_failed(exc)`` -- map a pane's ``ok=false`` reply to the
      specific ``ErrorCode`` a caller should raise (``ZERO_MATCH`` /
      ``MATCH_COUNT_MISMATCH`` vs a generic ``LIVE_OP_FAILED``).

Deliberately NOT here: never calls ``live/bridge.py``'s
``start_in_background()``. Every lookup below goes through
``live_bridge.current_registry()`` instead, which returns ``None``
(never binds a socket) when no bridge is running. This matters because
``resolve_write_mode`` runs on every ``replace_text``/``format_text``
call, including the overwhelming common case where the caller never
touched anything live-related (``write_mode="auto"``, no pane ever
connected) -- that path must be a cheap, side-effect-free lock check,
never an attempt to bind the bridge's real ports (53135/53136 in
production; see this repo's own hard rule against a test or an
unrelated call path ever touching those two specific ports). Starting
the bridge lazily stays ``live_status``'s job (and, from here on, any
live op's, once a session already exists) -- ``resolve_write_mode``
only ever asks "is one already running, and if so, is a pane already
connected for this document", never "start one so I can ask".
"""

from __future__ import annotations

import re
from typing import Any

from .. import audit, paths
from ..errors import ErrorCode, _make_error
from . import bridge as live_bridge
from .session import LiveOpFailed, LiveSession

_VALID_WRITE_MODES = ("auto", "file", "live")
_LIVE_REVISION_PREFIX = "live:sha256:"


def _document_name(path: str) -> str:
    """The file name a connected pane would report for *path* -- the same
    basename ``live/session.py``'s ``document_name_from_url`` derives
    from the pane's own ``hello.documentUrl``, and the same key
    ``SessionRegistry`` uses. Requires the file to exist (a live session
    can only ever exist for a file Word already has open, which by
    definition already exists on disk)."""
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    return resolved.name


def _desktop_word_owner_present(path: str) -> bool:
    """True if ``lock_status``'s owner-file check reports a DESKTOP WORD
    owner file for *path* (not LibreOffice, and not merely "some sync
    artifact present") -- the first half of the ``auto`` rule.

    Deferred import: ``server.py`` imports this module at process start
    (for ``replace_text``/``format_text``/``live_save``'s ``write_mode``
    dispatch), so a module-level ``from .. import server`` here would be
    a circular import -- the same deferred-import pattern ``mutations.py``
    already uses for ``tracked_changes.py``, and ``server.py``'s own
    ``_raise_tool_error`` uses for ``fastmcp.exceptions``.

    ``quiesce_interval=0.0``: ``execute_lock_status`` also samples the
    file twice ~1.5s apart (by default) to report sync-quiesce state,
    which this check does not need (only ``owner_file`` matters here) --
    passing 0 skips that wait rather than taxing every ``auto``-mode call
    with a needless 1.5s pause.
    """
    from .. import server as server_module

    status = server_module.execute_lock_status(path, quiesce_interval=0.0)
    owner = status["owner_file"]
    return bool(owner.get("present")) and owner.get("format") == "word"


def resolve_write_mode(path: str, requested: str) -> str:
    """Resolve *requested* (``"auto"``/``"file"``/``"live"``) to the
    concrete mode a ``replace_text``/``format_text``/comment-tool call
    should use for *path*.

    - ``"file"`` -- always ``"file"``. No bridge/registry lookup at all,
      so a caller passing this explicitly gets today's file-mode path
      byte-for-byte, with zero added cost or side effect.
    - ``"live"`` -- always ``"live"`` if a pane session is connected for
      *path*'s document name; otherwise raises ``LIVE_UNAVAILABLE``.
    - ``"auto"`` (default) -- ``"live"`` only when BOTH hold: a pane
      session is connected for *path*'s document name, AND
      ``lock_status`` reports a desktop Word owner file for *path*.
      Otherwise ``"file"``. Matching is by file name exactly as a
      connected pane reports it (``live/session.py``'s
      ``document_name_from_url``), never by full path -- a synced folder
      can differ in absolute path between machines while naming the same
      file the pane opened.

    Raises ``INVALID_INPUT`` for any other *requested* value.
    """
    if requested not in _VALID_WRITE_MODES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"write_mode must be one of {_VALID_WRITE_MODES!r}; got {requested!r}",
            {"write_mode": requested},
        )
    if requested == "file":
        return "file"

    document_name = _document_name(path)
    registry = live_bridge.current_registry()
    session = registry.get(document_name) if registry is not None else None

    if requested == "live":
        if session is None:
            raise _make_error(
                ErrorCode.LIVE_UNAVAILABLE,
                f"write_mode='live' requested but no connected pane session for document "
                f"{document_name!r}. Open the Live pane in Word (docs/live-mode.md) and retry, "
                "or call live_status to confirm the bridge is running.",
                {"document_name": document_name, "path": path},
            )
        return "live"

    # auto
    if session is None:
        return "file"
    if not _desktop_word_owner_present(path):
        return "file"
    return "live"


def live_session_for(path: str) -> LiveSession:
    """The connected ``LiveSession`` for *path*'s document name.

    Raises ``LIVE_UNAVAILABLE`` if no bridge is running, or none is
    connected for that document -- the same condition
    ``resolve_write_mode(path, "live")`` checks; a caller that already
    resolved ``mode == "live"`` calls this next to get the actual session
    object to send ops to.
    """
    document_name = _document_name(path)
    registry = live_bridge.current_registry()
    session = registry.get(document_name) if registry is not None else None
    if session is None:
        raise _make_error(
            ErrorCode.LIVE_UNAVAILABLE,
            f"no connected pane session for document {document_name!r}. Open the Live pane in "
            "Word (docs/live-mode.md) and retry, or call live_status to confirm the bridge is "
            "running.",
            {"document_name": document_name, "path": path},
        )
    return session


def check_not_stale(revision_before: str | None, current_body_sha256: str) -> None:
    """``LIVE_STALE`` pre-flight check, run before an op is sent.

    A no-op unless *revision_before* is given AND starts with
    ``"live:sha256:"`` (a file-mode revision token, or ``None``, is left
    alone here -- a caller mixing modes across calls is not this
    function's concern). Mirrors ``REVISION_CONFLICT`` on the file path,
    but is its own code: a live document has no OOXML revision token to
    compare, only the pane's own body-hash-before-the-op.
    """
    if not revision_before or not revision_before.startswith(_LIVE_REVISION_PREFIX):
        return
    expected = revision_before[len(_LIVE_REVISION_PREFIX) :]
    if expected != current_body_sha256:
        raise _make_error(
            ErrorCode.LIVE_STALE,
            f"revision_before {revision_before!r} no longer matches the live document's current "
            f"body hash ({_LIVE_REVISION_PREFIX}{current_body_sha256}) -- someone (the lead, "
            "typing) changed the document since that revision was read. Re-read and retry.",
            {"revision_before": revision_before, "current_body_sha256": current_body_sha256},
        )


_FOUND_COUNT_RE = re.compile(r"found (\d+)")


def classify_op_failed(exc: LiveOpFailed) -> ErrorCode:
    """Map a pane's ``ok=false`` reply to ``ZERO_MATCH``,
    ``MATCH_COUNT_MISMATCH``, or a generic ``LIVE_OP_FAILED``.

    Prefers the pane's own ``error.code`` when it already names one of
    the two specific cases (``"zero_match"`` / ``"match_count_mismatch"``,
    case-insensitive) -- a protocol extension this function is written to
    honor immediately if a future pane version sends it. Today's fake
    pane (``tests/unit/fake_pane.py``), and per ``live/protocol.py``'s own
    documented op contract, both still report a fixed ``"LIVE_OP_FAILED"``
    wire code for every ``expected_matches`` gate failure, with the
    actual count folded into the message text (``"expected N match(es)
    for 'x', found M"``) -- changing that wire-level code was ruled out
    here because an existing WP-2 test
    (``test_live_session.py``'s ``test_expected_matches_mismatch_raises_
    live_op_failed_with_pane_message``) pins today's fake pane to
    exactly that code/message shape. So, as a fallback, this parses the
    actual count straight out of the pane's own documented message
    format: 0 found -> ``ZERO_MATCH``, any other mismatched count ->
    ``MATCH_COUNT_MISMATCH``. A message that matches neither shape (an
    op failure unrelated to a match count -- not expected from
    ``replace``/``format``/``comment_add`` today, but not impossible from
    a future op) falls back to the generic ``LIVE_OP_FAILED``.
    """
    pane_code = (exc.code or "").strip().lower()
    if pane_code == "zero_match":
        return ErrorCode.ZERO_MATCH
    if pane_code == "match_count_mismatch":
        return ErrorCode.MATCH_COUNT_MISMATCH

    match = _FOUND_COUNT_RE.search(exc.message or "")
    if match:
        return ErrorCode.ZERO_MATCH if match.group(1) == "0" else ErrorCode.MATCH_COUNT_MISMATCH
    return ErrorCode.LIVE_OP_FAILED


def live_evidence(
    *,
    applied: bool,
    match_count: int,
    rung: Any,
    before: str,
    after: str,
    pre_body_sha256: str,
    post_body_sha256: str,
    document_name: str,
    tool: str,
    path: str,
) -> dict[str, Any]:
    """Build + audit-log the live-mode evidence envelope for one mutating
    live op, and return it.

    Carries the same eight keys file-mode evidence always carries
    (``applied``, ``match_count``, ``rung``, ``before``, ``after``,
    ``revision_before``, ``revision_after``, ``audit_logged``), plus the
    live-only keys the plan names: ``write_mode: "live"``,
    ``verified_via: "word-addin"``, ``document_name``,
    ``track_changes_author: "word-signed-in-user"`` (the pane cannot set
    ``w:author`` -- track-changes authorship in live mode is always
    whoever is signed into that copy of Word), and
    ``orphaned_comment_ids: []`` (live mode never removes a comment
    anchor out from under a write the way a forced file-mode write can --
    see ``COMMENT_ANCHORS_IN_RANGE``/``force`` in
    ``core/document-backend-protocol.md`` §4 -- so this is always empty,
    present for shape parity with the file-mode evidence a caller may
    already expect this key on).

    Deliberately absent: ``conflict_copy_detected`` and its siblings --
    live mode never produces a sync-conflict copy the way a file-mode
    write can, because Word owns the file for the whole session; there is
    nothing for a conflict-copy sweep to find.

    ``revision_before``/``revision_after`` are ``"live:sha256:<hex>"`` of
    the pane's own pre/post ``body.text`` hashes (there is no OOXML
    revision token to read until ``live_save`` writes the file back to
    disk -- see that tool for the bridge to file-mode's own tokens).

    ``tool``/``path`` go straight to ``audit.append_audit`` (the same
    audit log every file-mode mutating tool writes to) so a live edit is
    not invisible to the audit trail just because it did not go through
    ``mutations.atomic_replace_docx_parts``.
    """
    evidence: dict[str, Any] = {
        "applied": applied,
        "match_count": match_count,
        "rung": rung,
        "before": before,
        "after": after,
        "revision_before": f"{_LIVE_REVISION_PREFIX}{pre_body_sha256}",
        "revision_after": f"{_LIVE_REVISION_PREFIX}{post_body_sha256}",
        "audit_logged": False,
        "write_mode": "live",
        "verified_via": "word-addin",
        "document_name": document_name,
        "track_changes_author": "word-signed-in-user",
        "orphaned_comment_ids": [],
    }
    logged, _ = audit.append_audit(path=path, tool=tool, evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
