"""add_anchored_comment_batch (issue #113): N anchored comments in one call.

The server-side loop the model used to run by hand, one ``add_anchored_comment`` per
finding. Same contract as ``replace_text_batch``: per-item ``{index, status, ...}``, a
``batch_status``, the first failure stops the batch and NOTHING IS EVER RETRIED.

Live mode: one pane ``comment_add`` per comment, chained on the body hash. The pane
comments only the FIRST match, so every item needs ``expected_matches == 1``. Each item's
status is ``applied`` / ``refused_before_write`` / ``unknown`` / ``not_attempted``; earlier
comments stay when a later one fails.

File mode: all comments are staged on ONE in-memory document and saved once, guarded and
verified. A save is all-or-nothing, so a failure mid-batch persists nothing: earlier items
report ``not_persisted``, the failing one its code, the rest ``not_attempted``.

``section_key`` scopes an item to one heading section. File mode scopes ``locate`` to the
section's text, so the quote may repeat elsewhere. Live mode cannot scope the pane's
search, so there it is a check (against a NON-observing read of the live body, which never
re-arms a stale shared-bridge client): the quote must occur in the section and nowhere
else (``AMBIGUOUS_SPAN`` otherwise).

One audit record per call holds every item's evidence, written before the lean envelope
compacts the response.
"""

from __future__ import annotations

import hashlib
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from . import audit, comments, mutations, paths, projection, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, VerifyError, _make_error
from .live import comments_live, reads_live
from .live import write_mode as live_write_mode
from .live.session import LiveDisconnected, LiveOpFailed, LiveStale, LiveStaleAfterWrite
from .locate import locate
from .text_edit import _batch_refused_before_write

MAX_COMMENTS = 50
_MESSAGE_CHARS = 240
_ITEM_KEYS = {"quote", "content", "expected_matches", "section_key"}


def validate_comments(comments_arg: Any) -> list[dict[str, Any]]:
    """1..50 items, validated up front (nothing is sent or staged before this passes)."""
    if not isinstance(comments_arg, list) or not 1 <= len(comments_arg) <= MAX_COMMENTS:
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"comments must be a list of 1..{MAX_COMMENTS} comment objects"
        )
    out: list[dict[str, Any]] = []
    for i, item in enumerate(comments_arg):
        if not isinstance(item, dict):
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}] must be an object")
        unknown = set(item) - _ITEM_KEYS
        if unknown:
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}]: unknown key(s) {sorted(unknown)}")
        quote, content = item.get("quote"), item.get("content")
        expected, section_key = item.get("expected_matches", 1), item.get("section_key")
        if not isinstance(quote, str) or not quote:
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}].quote must be a non-empty string")
        if not isinstance(content, str) or not content:
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}].content must be a non-empty string")
        if isinstance(expected, bool) or not isinstance(expected, int) or expected < 1:
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}].expected_matches must be a positive integer")
        if section_key is not None and (not isinstance(section_key, str) or not section_key):
            raise _make_error(ErrorCode.INVALID_INPUT, f"comments[{i}].section_key must be a non-empty string")
        out.append({"quote": quote, "content": content, "expected_matches": expected, "section_key": section_key})
    return out


# --- section scoping (shared by both modes) ---------------------------------------------


def _heading_sections(path: Path, *, ignore_trailing_empty: bool = False) -> dict[str, dict[str, Any]]:
    found = projection.find_sections_impl(path, ignore_trailing_empty=ignore_trailing_empty)
    return {s["section_key"]: s for s in found if s["kind"] == "heading"}


def _section_entry(sections: dict[str, dict[str, Any]], section_key: str) -> dict[str, Any]:
    entry = sections.get(section_key)
    if entry is None:
        raise _make_error(
            ErrorCode.SECTION_NOT_FOUND,
            f"No heading section with section_key {section_key!r} was found (call find_sections).",
            {"section_key": section_key},
        )
    return entry


def section_span_filter(proj: projection.Projection, entry: dict[str, Any]) -> Callable[[int, int], bool]:
    """A ``locate`` span_filter accepting spans inside *entry*'s paragraph range."""
    refs = [m.para_ref for m in proj.paragraphs]
    try:
        lo_idx, hi_idx = refs.index(entry["start_para_ref"]), refs.index(entry["end_para_ref"])
    except ValueError:
        return lambda _s, _e: False
    in_range = set(refs[lo_idx : hi_idx + 1])
    spans = [(s, e) for s, e, para_ref, _run in proj.offset_map if para_ref in in_range]
    if not spans:
        return lambda _s, _e: False
    lo, hi = min(s for s, _ in spans), max(e for _, e in spans)
    return lambda s, e: lo <= s and e <= hi


# --- reporting ---------------------------------------------------------------------------


def _status_counts(results: list[dict[str, Any]]) -> int:
    return sum(1 for r in results if r["status"] == "applied")


def _public_item(entry: dict[str, Any], *, verbose: bool) -> dict[str, Any]:
    row: dict[str, Any] = {"index": entry["index"], "status": entry["status"]}
    for key in ("comment_id", "comment_ids", "code"):
        if entry.get(key) is not None:
            row[key] = entry[key]
    if entry.get("message"):
        row["message"] = entry["message"][:_MESSAGE_CHARS]
    if verbose and entry["status"] == "applied":
        row["evidence"] = entry.get("evidence", {})
    return row


def _audit_item(entry: dict[str, Any], item: dict[str, Any], *, excerpts: bool) -> dict[str, Any]:
    """Full per-item evidence for the audit record. Document text (the quote) and the
    comment text are kept only while excerpts are enabled: nested values are not covered by
    the audit's top-level redaction."""
    row = {k: v for k, v in entry.items() if k not in ("evidence", "message")}
    row.update(
        {
            "quote_chars": len(item["quote"]),
            "quote_sha256": hashlib.sha256(item["quote"].encode("utf-8")).hexdigest(),
            "content_chars": len(item["content"]),
            "content_sha256": hashlib.sha256(item["content"].encode("utf-8")).hexdigest(),
            "expected_matches": item["expected_matches"],
            "section_key": item["section_key"],
        }
    )
    row.update(entry.get("evidence", {}))
    if excerpts:
        row["quote"], row["content"] = item["quote"], item["content"]
    else:
        row.pop("before", None)
    if entry.get("message"):
        row["message_chars"] = len(entry["message"])
    return row


def _finish(
    path: str,
    *,
    mode: str,
    items: list[dict[str, Any]],
    results: list[dict[str, Any]],
    batch_status: str,
    revision_before_actual: str | None,
    revision_before_given: str | None,
    revision_after: str | None,
    verbose: bool,
    audit_extra: dict[str, Any],
    public_extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    applied_count = _status_counts(results)
    excerpts = audit._audit_excerpts_enabled(True)
    evidence: dict[str, Any] = {
        "applied": batch_status == "complete",
        "batch_status": batch_status,
        "applied_count": applied_count,
        "item_count": len(items),
        "revision_before": revision_before_actual,
        "revision_after": revision_after,
        "write_mode": mode,
        "audit_logged": False,
        "results": [_audit_item(r, items[r["index"]], excerpts=excerpts) for r in results],
        **audit_extra,
    }
    ref = audit.new_evidence_ref()
    logged, _ = audit.append_audit(path=path, tool="add_anchored_comment_batch", evidence=evidence, evidence_ref=ref)

    report: dict[str, Any] = {
        "applied": evidence["applied"],
        "batch_status": batch_status,
        "applied_count": applied_count,
        "item_count": len(items),
        "revision_before": revision_before_given or revision_before_actual,
        "revision_after": revision_after,
        "write_mode": mode,
        "results": [_public_item(r, verbose=verbose) for r in results],
    }
    report.update(public_extra or {})
    if logged:
        report["evidence_ref"] = ref
    else:
        report["audit_logged"] = False
    return report


def _unattempted(results: list[dict[str, Any]], total: int, *, start: int) -> None:
    for i in range(start, total):
        results.append({"index": i, "status": "not_attempted"})


# --- live mode ---------------------------------------------------------------------------


def _live_section_checker(session: Any, items: list[dict[str, Any]], pre_hash: str) -> Callable[[dict[str, Any]], None] | None:
    """A checker for items with a section_key, built from ONE non-observing read of the
    live body (``session`` is ``live_session_for``'s status view), or None if no item
    names a section."""
    if not any(item["section_key"] for item in items):
        return None
    live_write_mode.require_capability(
        session, reads_live.BODY_OOXML_CAPABILITY, feature_description="section_key checks on live comments"
    )
    reply = reads_live._fetch_body_ooxml(session)
    if reply["bodySha256"] != pre_hash:
        raise _make_error(
            ErrorCode.LIVE_STALE,
            "The live document changed while the batch was being prepared; nothing was written. Retry.",
        )
    with tempfile.TemporaryDirectory(prefix="verified-docx-comment-batch-") as tmp:
        docx = Path(tmp) / "live.docx"
        reads_live.flat_opc_to_docx(reply["ooxml"], docx)
        sections = _heading_sections(docx, ignore_trailing_empty=True)
        proj = projection.project_part(docx)

    def check(item: dict[str, Any]) -> None:
        entry = _section_entry(sections, item["section_key"])
        in_section = section_span_filter(proj, entry)
        # Raises ZERO_MATCH (not in the section) / MATCH_COUNT_MISMATCH (twice in it).
        locate(item["quote"], proj, 1, span_filter=in_section)
        try:
            anywhere = locate(item["quote"], proj, 1)
        except VerifyError as exc:
            if exc.envelope.error_code != ErrorCode.MATCH_COUNT_MISMATCH:
                raise
            anywhere = None
        if anywhere is None or not in_section(*anywhere.spans[0]):
            raise _make_error(
                ErrorCode.AMBIGUOUS_SPAN,
                f"quote also occurs outside section {item['section_key']!r}; live mode comments the "
                "pane's first match only, so it could land in the wrong place. Lengthen the quote so "
                "it is unique. Nothing was written.",
                {"section_key": item["section_key"]},
            )

    return check


def _live_batch(
    path: str, items: list[dict[str, Any]], revision_before: str | None, verbose: bool
) -> dict[str, Any]:
    for i, item in enumerate(items):
        if item["expected_matches"] != 1:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"comments[{i}].expected_matches must be 1 in live mode: the pane comments only the "
                "first match. Use a longer, unique quote.",
            )
    resolved, document_name = comments_live._document_name_and_path(path)
    session = live_write_mode.live_session_for(path)  # a non-observing view on a shared bridge
    try:
        pre_hash = session.request_threadsafe("describe")["bodySha256"]
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc
    revision_pre = f"live:sha256:{pre_hash}"
    audit_extra = {"verified_via": "word-addin", "document_name": document_name, "author": "word-signed-in-user"}

    results: list[dict[str, Any]] = []
    try:
        live_write_mode.check_not_stale(revision_before, pre_hash)
        check_section = _live_section_checker(session, items, pre_hash)
    except VerifyError as exc:
        if exc.envelope.error_code != ErrorCode.LIVE_STALE:
            raise
        results.append(
            {"index": 0, "status": "refused_before_write", "code": exc.envelope.error_code.value,
             "message": exc.envelope.message}
        )
        _unattempted(results, len(items), start=1)
        return _finish(
            str(resolved), mode="live", items=items, results=results, batch_status="none_applied",
            revision_before_actual=revision_pre, revision_before_given=revision_before,
            revision_after=revision_pre, verbose=verbose, audit_extra=audit_extra,
        )

    current = pre_hash
    unknown = False
    for i, item in enumerate(items):
        if unknown or (results and results[-1]["status"] != "applied"):
            results.append({"index": i, "status": "not_attempted"})
            continue
        try:
            if item["section_key"] and check_section is not None:
                check_section(item)
        except VerifyError as exc:
            results.append(
                {"index": i, "status": "refused_before_write", "code": exc.envelope.error_code.value,
                 "message": exc.envelope.message}
            )
            continue
        try:
            reply = _send_comment_add(session, item, current)
        except VerifyError as exc:
            code = exc.envelope.error_code.value
            if _batch_refused_before_write(exc):
                results.append({"index": i, "status": "refused_before_write", "code": code, "message": exc.envelope.message})
            else:
                unknown = True
                results.append({"index": i, "status": "unknown", "code": code, "message": exc.envelope.message})
            continue
        except Exception as exc:  # noqa: BLE001 - after dispatch any other failure leaves the outcome unknown
            unknown = True
            results.append({"index": i, "status": "unknown", "code": ErrorCode.LIVE_OP_FAILED.value, "message": str(exc)})
            continue
        current = reply.get("post") or current
        handle = f"{comments_live._LIVE_HANDLE_PREFIX}{reply['comment_id']}"
        results.append(
            {
                "index": i, "status": "applied", "comment_id": handle,
                "evidence": {"match_count": 1, "revision_after": f"live:sha256:{current}"},
            }
        )

    applied = _status_counts(results)
    if unknown:
        batch_status = "unknown"
    elif applied == len(items):
        batch_status = "complete"
    elif applied == 0:
        batch_status = "none_applied"
    else:
        batch_status = "partial"
    return _finish(
        str(resolved), mode="live", items=items, results=results, batch_status=batch_status,
        revision_before_actual=revision_pre, revision_before_given=revision_before,
        revision_after=None if unknown else f"live:sha256:{current}", verbose=verbose, audit_extra=audit_extra,
    )


def _send_comment_add(session: Any, item: dict[str, Any], pre_hash: str) -> dict[str, Any]:
    payload = {
        "find": item["quote"], "expected_matches": 1, "text": item["content"], "expectedBodySha256": pre_hash,
    }
    try:
        return session.request_threadsafe("comment_add", payload, expected_body_sha256=pre_hash)
    except LiveStaleAfterWrite as exc:
        raise _make_error(
            ErrorCode.LIVE_STALE_AFTER_WRITE,
            "the pane added the comment, but the document had moved since it was read: the comment "
            "ALREADY LANDED. Do not retry; re-read the document.",
        ) from exc
    except LiveStale as exc:
        raise _make_error(
            ErrorCode.LIVE_STALE, "The live document changed since it was read; nothing was written."
        ) from exc
    except LiveOpFailed as exc:
        raise _make_error(live_write_mode.classify_op_failed(exc), exc.message, {"pane_code": exc.code}) from exc
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc


# --- file mode ---------------------------------------------------------------------------


def _file_batch(
    path: str,
    items: list[dict[str, Any]],
    revision_before: str | None,
    allow_concurrent_editor: bool,
    verbose: bool,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(
        resolved, revision_before, allow_concurrent_editor=allow_concurrent_editor, live_capable=True
    )
    revision_pre = pre_revision["token"]
    document_root, raw_xml = mutations._load_document(resolved)
    author = resolve_author_name()
    date = tracked_changes.current_revision_date()
    parts = comments._CommentParts(resolved)
    next_id = comments._next_comment_id(parts.comments_root if not parts.comments_created else None)
    taken_hex = comments._existing_hex_ids(parts.ext_root, parts.ids_root)
    sections: dict[str, dict[str, Any]] | None = None

    results: list[dict[str, Any]] = []
    expected_anchors: dict[str, str] = {}
    failure: VerifyError | None = None
    for i, item in enumerate(items):
        try:
            span_filter_for = None
            if item["section_key"]:
                if sections is None:
                    sections = _heading_sections(resolved)
                entry = _section_entry(sections, item["section_key"])
                span_filter_for = lambda proj, entry=entry: section_span_filter(proj, entry)
            staged = comments.stage_anchored_comment(
                document_root, parts, quote=item["quote"], text=item["content"],
                expected_matches=item["expected_matches"], author=author, date=date,
                taken_hex=taken_hex, next_comment_id=next_id, span_filter_for=span_filter_for,
            )
        except VerifyError as exc:
            failure = exc
            results.append(
                {"index": i, "status": "refused_before_write", "code": exc.envelope.error_code.value,
                 "message": exc.envelope.message}
            )
            _unattempted(results, len(items), start=i + 1)
            break
        next_id = staged.next_comment_id
        expected_anchors.update(staged.anchored_text)
        ids = staged.durable_ids
        entry_result: dict[str, Any] = {
            "index": i, "status": "applied",
            "evidence": {"match_count": staged.locate_result.match_count, "before": staged.before_excerpt,
                         "rung": staged.locate_result.rung},
        }
        if len(ids) == 1:
            entry_result["comment_id"] = ids[0]
        else:
            entry_result["comment_ids"] = ids
        results.append(entry_result)

    def _all_not_persisted() -> None:
        for r in results:
            if r["status"] == "applied":
                r["status"] = "not_persisted"
                r.pop("comment_id", None)
                r.pop("comment_ids", None)

    audit_extra: dict[str, Any] = {}
    if failure is not None:
        _all_not_persisted()
        return _finish(
            str(resolved), mode="file", items=items, results=results, batch_status="none_applied",
            revision_before_actual=revision_pre, revision_before_given=revision_before,
            revision_after=revision_pre, verbose=verbose, audit_extra=audit_extra,
        )

    # One guarded, verified save for the whole batch.
    save_error: VerifyError | None = None
    conflict_sweep: dict[str, Any] | None = None
    try:
        comments._ensure_person(parts.people_root, author=author)
        comments.assert_ids_bracket_text(projection.project_document_root(document_root), expected_anchors)

        def _post_verify(written_path: Path) -> None:
            comments.assert_ids_bracket_text(projection.project_part(written_path), expected_anchors)

        overrides = parts.build_overrides(document_root=document_root, raw_xml=raw_xml)
        conflict_sweep = mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=_post_verify)
    except VerifyError as exc:
        save_error = exc
    except ValueError as exc:  # the pre-save anchor check
        save_error = _make_error(ErrorCode.VERIFICATION_FAILED, str(exc))

    if save_error is not None:
        landed = bool(save_error.envelope.diagnostics.get("rollback_skipped"))
        for r in results:
            if landed:
                r["status"] = "unknown"
            else:
                r["status"] = "not_persisted"
                r.pop("comment_id", None)
                r.pop("comment_ids", None)
        # Every item carries the cause, so a caller reading one row sees it.
        results[0]["code"] = save_error.envelope.error_code.value
        results[0]["message"] = save_error.envelope.message
        return _finish(
            str(resolved), mode="file", items=items, results=results,
            batch_status="unknown" if landed else "none_applied",
            revision_before_actual=revision_pre, revision_before_given=revision_before,
            revision_after=None if landed else revision_pre, verbose=verbose, audit_extra=audit_extra,
        )

    assert conflict_sweep is not None
    sweep_evidence: dict[str, Any] = {}
    mutations._merge_conflict_sweep(sweep_evidence, conflict_sweep)
    public_extra = {k: v for k, v in sweep_evidence.items() if k in ("conflict_copy_detected", "conflict_copies", "sibling_files_changed")}
    return _finish(
        str(resolved), mode="file", items=items, results=results, batch_status="complete",
        revision_before_actual=revision_pre, revision_before_given=revision_before,
        revision_after=conflict_sweep["revision_after"], verbose=verbose,
        audit_extra=sweep_evidence, public_extra=public_extra,
    )


def execute_add_anchored_comment_batch(
    path: str,
    comments_arg: list[dict[str, Any]],
    *,
    write_mode: str = "auto",
    revision_before: str | None = None,
    allow_concurrent_editor: bool = False,
    verbose: bool = False,
) -> dict[str, Any]:
    """See the module docstring. Argument errors and failures before the first write (a
    missing pane, a stale guard in file mode) raise; everything after is reported
    per item."""
    items = validate_comments(comments_arg)
    mode = comments_live._resolve_write_mode(path, write_mode)
    if mode == "live":
        return _live_batch(path, items, revision_before, verbose)
    return _file_batch(path, items, revision_before, allow_concurrent_editor, verbose)
