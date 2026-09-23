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
  ``raise_if_live_session_active(resolved)`` (issue #154) -- the
      file-mode-write counterpart of the three above: raises
      ``LIVE_SESSION_ACTIVE`` if a pane session is connected for
      *resolved*'s document name. ``mutations._guard_before_write`` calls
      this before every file-mode write (and ``atomic_replace_docx_parts``
      calls it again immediately before the write itself lands, to
      narrow the race between the two).

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
connected) -- that path must be a cheap, side-effect-free registry
lookup, never an attempt to bind the bridge's real ports (53135/53136 in
production; see this repo's own hard rule against a test or an
unrelated call path ever touching those two specific ports). Starting
the bridge lazily stays ``live_status``'s job (and, from here on, any
live op's, once a session already exists) -- ``resolve_write_mode``
only ever asks "is one already running, and if so, is a pane already
connected for this document", never "start one so I can ask".

Issue #154 fix: ``auto`` used to also require ``lock_status`` to report
a desktop Word owner file next to a connected pane session -- a signal
Word for Mac never writes into a SharePoint/OneDrive-synced folder, so
``auto`` could never reach live on that platform/storage combination.
That condition is gone; ``auto`` now matches ``comments_live.py``'s own
read-side ``_resolve_source`` rule exactly: a connected pane session is
sufficient on its own. See ``mutations._guard_before_write`` for the
matching fix on the file-mode side (a connected pane now REFUSES a
file-mode write with ``LIVE_SESSION_ACTIVE`` instead of relying on that
same dead owner-file signal to protect it).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

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


def _session_url_local_path(document_url: str) -> Path | None:
    """*document_url* as a local filesystem path, or ``None`` if it isn't
    one (a SharePoint/OneDrive web URL, e.g. ``https://...`` -- the
    actual shape of issue #154's incident). Only a bare path or an
    explicit ``file://`` URL resolves; anything else is left to
    basename-only matching (see ``_check_session_identity`` below), since
    a web URL carries no local path this process could compare against.
    """
    parsed = urlparse(document_url)
    if parsed.scheme in ("", "file"):
        raw = unquote(parsed.path) if parsed.scheme == "file" else document_url
        if raw:
            return Path(raw)
    return None


def _check_session_identity(path: str, session: LiveSession) -> None:
    """Issue #154 WP-1b: ``SessionRegistry`` keys purely on basename
    (deliberate -- a synced folder's absolute path differs per machine),
    so ``auto``/``live`` routing a *different* document that happens to
    share a file name into this session's pane would silently mutate the
    wrong file. Match counts on the pane's side cannot catch this; only
    comparing *path* against the session's own ``document_url`` can.

    Refuses (``LIVE_SESSION_MISMATCH``) only when ``document_url``
    resolves to a LOCAL path (see ``_session_url_local_path``) that
    differs from *path*'s own resolved real path. When ``document_url``
    is a web URL (SharePoint/OneDrive -- issue #154's actual incident
    shape), there is no local path to compare, so this falls back to the
    basename match the registry lookup already performed and does not
    fail closed -- doing so would disable live mode on exactly the
    platform/storage combination this issue is about.
    """
    session_local = _session_url_local_path(session.document_url)
    if session_local is None:
        return
    target_resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    try:
        session_real = session_local.resolve()
    except OSError:
        return
    if session_real != target_resolved.resolve():
        raise _make_error(
            ErrorCode.LIVE_SESSION_MISMATCH,
            f"a connected pane session shares this document's file name ({session.document_name!r}) "
            f"but its own document_url resolves to a different file ({session_real}) than the one "
            f"requested ({target_resolved}). Refusing to route to a session that may be editing a "
            "different document than the one this call names.",
            {"path": str(target_resolved), "session_document_url": session.document_url},
        )


def resolve_write_mode(path: str, requested: str) -> str:
    """Resolve *requested* (``"auto"``/``"file"``/``"live"``) to the
    concrete mode a ``replace_text``/``format_text``/comment-tool call
    should use for *path*.

    - ``"file"`` -- always ``"file"``. No bridge/registry lookup at all,
      so a caller passing this explicitly gets today's file-mode path
      byte-for-byte, with zero added cost or side effect.
    - ``"live"`` -- always ``"live"`` if a pane session is connected for
      *path*'s document name; otherwise raises ``LIVE_UNAVAILABLE``.
    - ``"auto"`` (default, issue #154) -- ``"live"`` whenever a pane
      session is connected for *path*'s document name; otherwise
      ``"file"``. Matching is by file name exactly as a connected pane
      reports it (``live/session.py``'s ``document_name_from_url``),
      never by full path -- a synced folder can differ in absolute path
      between machines while naming the same file the pane opened. (A
      desktop Word owner file used to be required too; that signal never
      appears on Word for Mac against a SharePoint/OneDrive-synced path,
      so ``auto`` could never route live there -- see this module's own
      header comment.)

    Either branch that resolves to ``"live"`` also runs
    ``_check_session_identity`` (issue #154 WP-1b) to refuse
    (``LIVE_SESSION_MISMATCH``) a session whose own ``document_url``
    names a different local file than *path*.

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
        _check_session_identity(path, session)
        return "live"

    # auto (issue #154: session alone is sufficient -- no owner-file check)
    if session is None:
        return "file"
    _check_session_identity(path, session)
    return "live"


def raise_if_live_session_active(resolved: Path) -> None:
    """Issue #154: refuse a FILE-MODE write with ``LIVE_SESSION_ACTIVE``
    when a pane session is connected for *resolved*'s document name (by
    basename, same key ``SessionRegistry`` uses -- ``resolved`` is
    already a resolved, existing path, so no ``_document_name`` call is
    needed here).

    Called from two places, on purpose, to narrow (not eliminate) the
    window between the guard and the write actually landing:
    ``mutations._guard_before_write`` (well before the write: hazard
    scan, revision check, markdown render all still have to happen) and
    ``mutations.atomic_replace_docx_parts`` itself (immediately before
    the write -- a pane can connect during the guard's own
    sync-quiesce wait, which is up to ~10s). Never starts the bridge
    lazily (same rule as ``resolve_write_mode`` -- an ordinary file-mode
    write must not bind ports 53135/53136).

    The remedy in the message is "save and CLOSE THE DOCUMENT in Word",
    never "close the pane": closing only the task pane drops its socket
    but leaves Word holding the document open with an unsaved in-memory
    buffer, so a file write immediately after would still race Word's
    own next autosave -- the exact hazard this function exists to stop.
    """
    document_name = resolved.name
    registry = live_bridge.current_registry()
    session = registry.get(document_name) if registry is not None else None
    if session is None:
        return
    raise _make_error(
        ErrorCode.LIVE_SESSION_ACTIVE,
        f"a Live pane session is connected for {document_name!r} (document_url="
        f"{session.document_url!r}). Writing to disk now would race Word's own autosave of its "
        "open in-memory copy and the file-mode write would likely be silently reverted (issue "
        "#154). Use write_mode=\"live\" on this call if it takes one, or save and CLOSE THE "
        "DOCUMENT in Word -- closing only the Live pane is not enough, Word keeps its buffer "
        "open until the document itself closes -- then retry.",
        {"document_name": document_name, "session_document_url": session.document_url},
    )


def live_session_for(path: str) -> LiveSession:
    """The connected ``LiveSession`` for *path*'s document name.

    Raises ``LIVE_UNAVAILABLE`` if no bridge is running, or none is
    connected for that document -- the same condition
    ``resolve_write_mode(path, "live")`` checks; a caller that already
    resolved ``mode == "live"`` calls this next to get the actual session
    object to send ops to. Also runs ``_check_session_identity`` (issue
    #154 WP-1b), same as ``resolve_write_mode``'s own ``"live"`` branch.
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
    _check_session_identity(path, session)
    return session


def require_capability(session: LiveSession, capability: str, *, feature_description: str) -> None:
    """Issue #22 B2: raise ``LIVE_CAPABILITY_MISSING`` unless *session*'s
    own ``hello.capabilities`` reported *capability*.

    Called BEFORE an op that depends on a wire-level payload field a pane
    build might not implement (e.g. ``rowAnchor``) is ever sent. An
    already-connected pane that predates the capability sends no
    ``capabilities`` field at all (``HelloMessage.from_json`` defaults it
    to an empty ``frozenset`` -- see that class's own docstring), so this
    always fails closed for it: a stale pane can never silently ignore
    the new payload key and run an unscoped op instead of refusing, which
    is the exact failure mode this function exists to close off on the
    wire, not just in this server's own Python.
    """
    if capability in session.hello.capabilities:
        return
    raise _make_error(
        ErrorCode.LIVE_CAPABILITY_MISSING,
        f"the connected pane for {session.document_name!r} did not report the {capability!r} "
        f"capability, required for {feature_description}. Reload the Live pane in Word (it may be "
        "running a build from before this feature existed) and retry.",
        {"document_name": session.document_name, "capability": capability, "feature": feature_description},
    )


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
