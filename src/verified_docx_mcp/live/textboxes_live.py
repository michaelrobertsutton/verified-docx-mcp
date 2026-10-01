"""Live text-box scope for ``replace_text`` / ``format_text`` and the
``list_textboxes`` read (issue #35).

Floating callout boxes are Word shapes, so the body-only live ``replace``/
``format`` ops never reach them (a box-only ``find`` came back ZERO_MATCH and
callers fell back to unverified AppleScript). A non-body ``scope`` sends the
same ops to the pane with two extra payload fields, and this module does the
server half of making that write verified:

1. ``textbox_list`` first, recording each target's ``text_sha256``.
2. The write carries ``expect`` ({shape_id: text_sha256}); the pane re-reads
   every target and refuses (``stale_target`` -> ``LIVE_STALE``) BEFORE
   mutating if any differs. The session layer's own staleness check compares
   the reply's body hash after the pane has already written, and a body hash
   does not cover text boxes at all, so it cannot protect these writes.
3. Verification never trusts the write's own reply alone: every touched shape
   is compared with the exact expected text, then a second, independent
   ``textbox_list`` must agree.

Targets are Word ``Shape.id`` values, never positional keys: file mode numbers
``w:txbxContent`` in XML order (``projection.py``) and nothing proves Word's
shape order matches it.

Failure classes: a refusal before the write (capability, stale target, count
gate, incomplete coverage) changed nothing and is safe to retry. A disconnect,
timeout, malformed reply or failed verification AFTER the write was sent may
have been applied; those raise with that stated and write an ``applied:
"unknown"`` audit entry, so an uncertain outcome is never invisible.

Known limit, not solved here: Office.js batches are not document locks, so an
edit by a co-author between the pane's compare-and-set and its write inside
one ``Word.run`` can still be overwritten.
"""

from __future__ import annotations

import hashlib
from typing import Any

from .. import audit
from ..errors import ErrorCode, _make_error
from . import bridge as live_bridge
from . import write_mode
from .protocol import (
    OP_ERROR_CAPABILITY_MISSING,
    OP_ERROR_STALE_TARGET,
    SCOPE_ALL,
    SCOPE_BODY,
    SCOPE_SHAPE_PREFIX,
    SCOPE_TEXTBOXES,
    ProtocolError,
    TextboxInfo,
    TextboxListResult,
    valid_scope,
)
from .session import LiveDisconnected, LiveOpFailed, LiveSession, LiveStale

TEXTBOX_CAPABILITY = "textbox_scope"

_MAY_HAVE_APPLIED = (
    " The write may have been applied in Word -- call list_textboxes to check the current text "
    "before retrying."
)


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# scope validation
# ---------------------------------------------------------------------------


def validate_scope(scope: str, within_row_containing: str | None = None) -> str:
    """*scope* if it is one of body/textboxes/all/shape:<id>; ``INVALID_INPUT``
    otherwise, or if combined with ``within_row_containing`` (row anchors are
    body tables)."""
    if not isinstance(scope, str) or not valid_scope(scope):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f'scope must be "body", "textboxes", "all" or "shape:<id>"; got {scope!r}',
            {"scope": scope},
        )
    if scope != SCOPE_BODY and within_row_containing:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "within_row_containing anchors on a body table row and cannot be combined with a "
            f"text-box scope (got scope={scope!r})",
            {"scope": scope},
        )
    return scope


def refuse_in_file_mode(scope: str) -> None:
    """Text-box writes are live-only for now (file-mode ``replace_text`` only
    walks the main body)."""
    if scope != SCOPE_BODY:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"scope={scope!r} is a live-only scope: file-mode replace_text/format_text only edit the "
            'main body. Connect the Live pane (write_mode="live"), or use scope="body".',
            {"scope": scope},
        )


# ---------------------------------------------------------------------------
# pane calls
# ---------------------------------------------------------------------------


def _map_op_failed(exc: LiveOpFailed) -> ErrorCode:
    code = (exc.code or "").strip().lower()
    if code == OP_ERROR_STALE_TARGET:
        return ErrorCode.LIVE_STALE
    if code == OP_ERROR_CAPABILITY_MISSING:
        return ErrorCode.LIVE_CAPABILITY_MISSING
    return write_mode.classify_op_failed(exc)


def _list(session: LiveSession) -> TextboxListResult:
    try:
        raw = session.request_threadsafe("textbox_list", {})
    except LiveOpFailed as exc:
        raise _make_error(_map_op_failed(exc), exc.message, {"pane_code": exc.code}) from exc
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc
    try:
        listing = TextboxListResult.from_json(raw)
    except ProtocolError as exc:
        raise _make_error(
            ErrorCode.LIVE_OP_FAILED,
            f"the pane's textbox_list reply is malformed: {exc}",
            {"stage": "textbox_list"},
        ) from exc
    ids = [t.shape_id for t in listing.textboxes]
    if len(set(ids)) != len(ids):
        raise _make_error(
            ErrorCode.LIVE_OP_FAILED,
            "the pane's textbox_list reply repeats a shape id",
            {"stage": "textbox_list", "shape_ids": ids},
        )
    for t in listing.textboxes:
        if _sha(t.text) != t.text_sha256:
            raise _make_error(
                ErrorCode.LIVE_OP_FAILED,
                f"the pane's textbox_list reply is inconsistent: shape {t.shape_id}'s text does not "
                "match its text_sha256",
                {"stage": "textbox_list", "shape_id": t.shape_id},
            )
    return listing


def _describe(session: LiveSession) -> str:
    try:
        return session.request_threadsafe("describe")["bodySha256"]
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc


def _info(t: TextboxInfo) -> dict[str, Any]:
    return {
        "shape_id": t.shape_id,
        "type": t.type,
        "group_path": t.group_path,
        "text": t.text,
        "text_sha256": t.text_sha256,
        "paragraph_count": t.paragraph_count,
    }


def list_textboxes_live(path: str) -> dict[str, Any]:
    """The ``list_textboxes`` tool: every text-bearing shape in the connected
    document body, with the Word-session-scoped ``shape_id`` a write targets."""
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, TEXTBOX_CAPABILITY, feature_description="list_textboxes")
    body_hash = _describe(session)
    listing = _list(session)
    notes = [
        (
            "shape_id is Word's Shape.id: valid only for this Word session and NOT a file-mode "
            "textbox-<n> key (file mode numbers text boxes in XML order; Word's order may differ)."
        ),
        "Only the document body is covered; text boxes anchored in headers/footers are not listed.",
    ]
    registry = live_bridge.current_registry()
    collisions = getattr(registry, "collisions", None)
    warnings: list[str] = []
    if collisions is not None and any(
        c.get("document_name") == session.document_name for c in collisions()
    ):
        warnings.append("basename_collision")
    result: dict[str, Any] = {
        "source": "live",
        "document_name": session.document_name,
        "revision": f"live:sha256:{body_hash}",
        "textboxes": [_info(t) for t in listing.textboxes],
        "incomplete": listing.incomplete,
        "skipped": listing.skipped,
        "notes": notes,
    }
    if warnings:
        result["warnings"] = warnings
    return result


# ---------------------------------------------------------------------------
# scoped writes
# ---------------------------------------------------------------------------


def _select_targets(listing: TextboxListResult, scope: str) -> list[TextboxInfo]:
    if scope.startswith(SCOPE_SHAPE_PREFIX):
        wanted = scope[len(SCOPE_SHAPE_PREFIX) :]
        picked = [t for t in listing.textboxes if t.shape_id == wanted]
        if not picked:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"no readable text shape with id {wanted!r}; call list_textboxes for the current ids",
                {"scope": scope, "available_shape_ids": [t.shape_id for t in listing.textboxes]},
            )
        return picked
    if listing.incomplete:
        raise _make_error(
            ErrorCode.LIVE_OP_FAILED,
            f"text-shape coverage is incomplete, so scope={scope!r} would silently skip some shapes. "
            'Target one box with scope="shape:<id>" instead (see list_textboxes).',
            {"incomplete": listing.incomplete},
        )
    if scope == SCOPE_TEXTBOXES and not listing.textboxes:
        raise _make_error(
            ErrorCode.ZERO_MATCH,
            "the live document has no text boxes (list_textboxes found none)",
            {"scope": scope},
        )
    return list(listing.textboxes)


def _audit_unknown(path: str, tool: str, scope: str, target_ids: list[str], reason: str) -> None:
    audit.append_audit(
        path=path,
        tool=tool,
        evidence={
            "applied": "unknown",
            "write_mode": "live",
            "scope": scope,
            "textbox_ids": target_ids,
            "reason": reason,
        },
    )


class _Run:
    """Shared state for one scoped write, plus the single failure path used
    after the write op has been sent."""

    def __init__(self, path: str, tool: str, scope: str, targets: list[TextboxInfo]) -> None:
        self.path = path
        self.tool = tool
        self.scope = scope
        self.targets = targets
        self.ids = [t.shape_id for t in targets]

    def fail(self, code: ErrorCode, message: str, diagnostics: dict[str, Any] | None = None) -> Any:
        _audit_unknown(self.path, self.tool, self.scope, self.ids, message)
        raise _make_error(code, message + _MAY_HAVE_APPLIED, diagnostics or {})


def _send_write(
    session: LiveSession,
    op: str,
    payload: dict[str, Any],
    run: _Run,
    *,
    body_scope: bool,
    pre_hash: str,
) -> dict[str, Any]:
    try:
        # The session layer's staleness check compares the REPLY's body hash,
        # so it is only meaningful when the body is itself a target.
        return session.request_threadsafe(
            op, payload, expected_body_sha256=pre_hash if body_scope else None
        )
    except LiveOpFailed as exc:
        # Pane refusals happen before any mutation (count gate, stale target,
        # incomplete coverage): nothing changed, safe to retry.
        raise _make_error(_map_op_failed(exc), exc.message, {"pane_code": exc.code}) from exc
    except LiveStale as exc:
        run.fail(ErrorCode.LIVE_STALE, str(exc), {"expected": exc.expected, "actual": exc.actual})
    except LiveDisconnected as exc:
        run.fail(ErrorCode.LIVE_DISCONNECTED, str(exc))
    raise AssertionError("unreachable")  # pragma: no cover


def _reply_shapes(result: dict[str, Any], run: _Run) -> dict[str, dict[str, Any]]:
    shapes = result.get("shapes")
    matches = result.get("matches")
    if not isinstance(shapes, list) or not isinstance(matches, list):
        run.fail(ErrorCode.VERIFICATION_FAILED, "the pane's reply is missing `shapes`/`matches`")
    by_id: dict[str, dict[str, Any]] = {}
    for s in shapes:
        if not isinstance(s, dict) or not all(
            isinstance(s.get(k), str) for k in ("shape_id", "pre_text", "post_text")
        ):
            run.fail(
                ErrorCode.VERIFICATION_FAILED, "the pane's reply has a malformed `shapes` entry"
            )
        by_id[s["shape_id"]] = s
    if set(by_id) != set(run.ids):
        run.fail(
            ErrorCode.VERIFICATION_FAILED,
            f"the pane reported shapes {sorted(by_id)} but the write targeted {sorted(run.ids)}",
        )
    return by_id


def _check_counts(
    result: dict[str, Any], by_id: dict[str, dict[str, Any]], expected_matches: int, run: _Run
) -> int:
    body_count = result.get("body_match_count", 0)
    if not isinstance(body_count, int) or isinstance(body_count, bool) or body_count < 0:
        run.fail(
            ErrorCode.VERIFICATION_FAILED,
            "the pane's `body_match_count` is not a non-negative integer",
        )
    shape_total = sum(int(s.get("match_count", 0)) for s in by_id.values())
    reported = result.get("match_count")
    n_matches = len(result.get("matches") or [])
    if not (shape_total + body_count == reported == n_matches == expected_matches):
        run.fail(
            ErrorCode.VERIFICATION_FAILED,
            f"match counts disagree: expected {expected_matches}, pane reported match_count={reported}, "
            f"{n_matches} match record(s), {shape_total} in text boxes + {body_count} in the body",
            {"expected": expected_matches, "match_count": reported, "records": n_matches},
        )
    return body_count


def _check_listed_text(
    by_id: dict[str, dict[str, Any]], targets: list[TextboxInfo], run: _Run
) -> None:
    for t in targets:
        if _sha(by_id[t.shape_id]["pre_text"]) != t.text_sha256:
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"the pane operated on text for shape {t.shape_id} that differs from what was listed",
                {"shape_id": t.shape_id},
            )


def _second_read(
    session: LiveSession,
    run: _Run,
    listing_before: TextboxListResult,
    expected_text: dict[str, str],
    *,
    relaxed_contains: str | None = None,
    relaxed_ids: set[str] | None = None,
) -> TextboxListResult:
    """Independent re-read. Every target must match *expected_text* exactly.
    Under track changes (tracked deletions can stay in the text) the targets
    in *relaxed_ids* -- only those that actually had matches -- instead must
    differ from the original and contain *relaxed_contains*. Every
    non-targeted shape must be unchanged (checked for shape scope)."""
    try:
        after = _list(session)
    except Exception as exc:  # noqa: BLE001 -- includes VerifyError from _list
        run.fail(
            ErrorCode.VERIFICATION_FAILED,
            f"the independent re-read of the text boxes failed: {exc}",
        )
    now = {t.shape_id: t for t in after.textboxes}
    before = {t.shape_id: t for t in listing_before.textboxes}
    for shape_id, want in expected_text.items():
        got = now.get(shape_id)
        if got is None:
            run.fail(ErrorCode.VERIFICATION_FAILED, f"shape {shape_id} is missing on re-read")
        if relaxed_contains is None or shape_id not in (relaxed_ids or set()):
            ok = got.text == want
        else:
            ok = got.text != before[shape_id].text and (
                relaxed_contains == "" or relaxed_contains in got.text
            )
        if not ok:
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"independent re-read of shape {shape_id} does not show the expected text",
                {
                    "shape_id": shape_id,
                    "expected_sha256": _sha(want),
                    "actual_sha256": got.text_sha256,
                },
            )
    for shape_id, t in before.items():
        if shape_id in expected_text:
            continue
        got = now.get(shape_id)
        if (
            got is not None
            and got.text_sha256 != t.text_sha256
            and run.scope.startswith(SCOPE_SHAPE_PREFIX)
        ):
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"shape {shape_id} was not targeted but its text changed during the write",
                {"shape_id": shape_id},
            )
    return after


def _prepare(
    path: str, find: str, scope: str, revision_before: str | None
) -> tuple[LiveSession, str, TextboxListResult, list[TextboxInfo]]:
    if not find:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty")
    session = write_mode.live_session_for(path)
    write_mode.require_capability(
        session, TEXTBOX_CAPABILITY, feature_description=f"scope={scope!r}"
    )
    pre_hash = _describe(session)
    write_mode.check_not_stale(revision_before, pre_hash)
    listing = _list(session)
    targets = _select_targets(listing, scope)
    return session, pre_hash, listing, targets


def execute_replace_text_scoped(
    path: str,
    find: str,
    replace: str,
    expected_matches: int,
    *,
    scope: str,
    revision_before: str | None = None,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Live ``replace_text`` over text boxes (``scope`` != ``"body"``)."""
    if find == replace:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "find and replace are identical; replacing would change nothing and cannot be verified",
        )
    session, pre_hash, listing, targets = _prepare(path, find, scope, revision_before)
    run = _Run(path, "replace_text", scope, targets)
    payload = {
        "find": find,
        "expected_matches": expected_matches,
        "replace": replace,
        "track_changes": track_changes,
        "scope": scope,
        "expect": {t.shape_id: t.text_sha256 for t in targets},
    }
    body_scope = scope == SCOPE_ALL
    result = _send_write(session, "replace", payload, run, body_scope=body_scope, pre_hash=pre_hash)

    if not result.get("applied"):
        run.fail(ErrorCode.VERIFICATION_FAILED, "the pane did not report applied=true")
    by_id = _reply_shapes(result, run)
    body_count = _check_counts(result, by_id, expected_matches, run)
    _check_listed_text(by_id, targets, run)

    post_hash = result.get("post")
    if body_count > 0 and post_hash == pre_hash:
        run.fail(
            ErrorCode.VERIFICATION_FAILED,
            "body matches were replaced but the body hash did not change",
        )

    matches = result.get("matches") or []
    expected_text: dict[str, str] = {}
    for t in targets:
        s = by_id[t.shape_id]
        want = t.text.replace(find, replace)
        expected_text[t.shape_id] = want
        n = int(s.get("match_count", 0))
        if track_changes:
            # Tracked deletions can stay in the text, so exact equality is not
            # achievable; weaker (and reported as such in the evidence).
            ok = (n == 0 and s["post_text"] == t.text) or (
                n > 0 and s["post_text"] != t.text and (replace == "" or replace in s["post_text"])
            )
        else:
            ok = s["post_text"] == want
        if not ok:
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"the pane's post-write text for shape {t.shape_id} is not what replacing "
                f"{find!r} with {replace!r} should produce",
                {
                    "shape_id": t.shape_id,
                    "expected_sha256": _sha(want),
                    "actual_sha256": _sha(s["post_text"]),
                },
            )
    for m in matches:
        after = m.get("after")
        if not isinstance(after, str) or (
            after != replace if not track_changes else replace not in after
        ):
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"a matched range's read-back text {after!r} does not equal the requested replacement {replace!r}",
                {"matches": matches},
            )

    touched = {shape_id for shape_id, s in by_id.items() if int(s.get("match_count", 0)) > 0}
    after_listing = _second_read(
        session,
        run,
        listing,
        expected_text,
        relaxed_contains=replace if track_changes else None,
        relaxed_ids=touched,
    )
    if body_count > 0:
        confirm = _describe(session)
        if confirm == pre_hash:
            run.fail(
                ErrorCode.VERIFICATION_FAILED, "an independent describe shows the body unchanged"
            )

    return write_mode.live_evidence(
        applied=True,
        match_count=expected_matches,
        rung=3,
        before="\n".join(m.get("before", "") for m in matches),
        after="\n".join(m.get("after", "") for m in matches),
        pre_body_sha256=pre_hash,
        post_body_sha256=post_hash or pre_hash,
        document_name=session.document_name,
        tool="replace_text",
        path=path,
        extra=_extra(scope, listing, after_listing, run.ids, track_changes, "exact"),
    )


def _extra(
    scope: str,
    before: TextboxListResult,
    after: TextboxListResult,
    ids: list[str],
    track_changes: bool,
    kind: str,
) -> dict[str, Any]:
    return {
        "scope": scope,
        "textbox_ids": ids,
        "textbox_text_sha256": {
            "before": {t.shape_id: t.text_sha256 for t in before.textboxes if t.shape_id in ids},
            "after": {t.shape_id: t.text_sha256 for t in after.textboxes if t.shape_id in ids},
        },
        "verification": "relaxed_track_changes" if track_changes and kind == "exact" else kind,
        "second_read": True,
    }


def _underline_on(value: Any) -> bool:
    return str(value or "none").strip().lower() not in ("none", "")


def execute_format_text_scoped(
    path: str,
    find: str,
    style: dict[str, bool | str],
    expected_matches: int,
    *,
    scope: str,
    revision_before: str | None = None,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Live ``format_text`` over text boxes (``scope`` != ``"body"``). *style*
    is already validated by the caller (``text_edit._validate_style``)."""
    session, pre_hash, listing, targets = _prepare(path, find, scope, revision_before)
    run = _Run(path, "format_text", scope, targets)
    payload = {
        "find": find,
        "expected_matches": expected_matches,
        "bold": style.get("bold"),
        "italic": style.get("italic"),
        "underline": style.get("underline"),
        "strike": style.get("strike"),
        "color": style.get("color"),
        "track_changes": track_changes,
        "scope": scope,
        "expect": {t.shape_id: t.text_sha256 for t in targets},
    }
    result = _send_write(
        session, "format", payload, run, body_scope=scope == SCOPE_ALL, pre_hash=pre_hash
    )

    if not result.get("applied"):
        run.fail(ErrorCode.VERIFICATION_FAILED, "the pane did not report applied=true")
    by_id = _reply_shapes(result, run)
    _check_counts(result, by_id, expected_matches, run)
    _check_listed_text(by_id, targets, run)

    # Formatting never changes text.
    for t in targets:
        if by_id[t.shape_id]["post_text"] != t.text:
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"format_text changed the text of shape {t.shape_id}; it must never change content",
                {"shape_id": t.shape_id},
            )

    matches = result.get("matches") or []
    checks: list[tuple[str, str, Any]] = []
    if style.get("bold") is not None:
        checks.append(("bold", "boldAfter", bool(style["bold"])))
    if style.get("italic") is not None:
        checks.append(("italic", "italicAfter", bool(style["italic"])))
    if style.get("underline") is not None:
        checks.append(("underline", "underlineAfter", bool(style["underline"])))
    if style.get("strike") is not None:
        checks.append(("strike", "strikeAfter", bool(style["strike"])))
    for m in matches:
        if m.get("after") != m.get("before"):
            run.fail(
                ErrorCode.VERIFICATION_FAILED,
                f"a matched range's text changed ({m.get('before')!r} -> {m.get('after')!r}); "
                "format_text must never change content",
                {"matches": matches},
            )
        for name, key, want in checks:
            if key not in m:
                run.fail(
                    ErrorCode.VERIFICATION_FAILED, f"the pane did not read back {name} ({key})"
                )
            got = _underline_on(m[key]) if name == "underline" else bool(m[key])
            if got != want:
                run.fail(
                    ErrorCode.VERIFICATION_FAILED,
                    f"the pane's read-back {name} ({m[key]!r}) does not equal the requested {want!r}",
                    {"matches": matches},
                )
        if style.get("color") is not None:
            got_color = (m.get("colorAfter") or "").upper().lstrip("#")
            if got_color != str(style["color"]).upper().lstrip("#"):
                run.fail(
                    ErrorCode.VERIFICATION_FAILED,
                    f"the pane's read-back font.color ({m.get('colorAfter')!r}) does not equal the "
                    f"requested color ({style['color']!r})",
                    {"matches": matches},
                )

    expected_text = {t.shape_id: t.text for t in targets}
    after_listing = _second_read(session, run, listing, expected_text)

    return write_mode.live_evidence(
        applied=True,
        match_count=expected_matches,
        rung=2,
        before="\n".join(m.get("before", "") for m in matches),
        after="\n".join(m.get("after", "") for m in matches),
        pre_body_sha256=pre_hash,
        post_body_sha256=result.get("post") or pre_hash,
        document_name=session.document_name,
        tool="format_text",
        path=path,
        extra=_extra(
            scope,
            listing,
            after_listing,
            run.ids,
            track_changes,
            "text_unchanged_formatting_read_back",
        ),
    )
