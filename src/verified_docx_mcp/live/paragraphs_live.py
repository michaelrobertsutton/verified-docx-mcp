"""Whole-paragraph deletion through Word, guarded and independently verified."""
from typing import Any

from .. import audit
from ..errors import ErrorCode, VerifyError, _make_error
from . import comments_live
from . import write_mode as wm


def delete_paragraph(path: str, anchor_text: str, revision_before: str,
                     track_changes: bool = False) -> dict[str, Any]:
    if not anchor_text or not revision_before.startswith("live:sha256:"):
        raise _make_error(ErrorCode.INVALID_INPUT, "A whole-paragraph anchor and live body revision are required")
    session = wm.live_session_for(path)
    wm.require_capability(session, "delete_paragraph", feature_description="whole-paragraph deletion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    wm.check_not_stale(revision_before, pre)
    try:
        result = comments_live._request(session, "paragraph_delete", {
            "anchor_text": anchor_text, "expectedBodySha256": pre, "track_changes": track_changes})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="delete_paragraph", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    evidence = wm.live_evidence(
        shape_result=result,
        applied=True, match_count=1, rung=1, before=anchor_text, after="",
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="delete_paragraph", path=path)
    evidence.update(result)
    return evidence


def clear_numbering(path: str, table_id: int | None = None, paragraph_text: str | None = None,
                    revision_before: str | None = None, write_mode: str = "live") -> dict[str, Any]:
    """Live-only ``clear_numbering`` (issue #79): detach list numbering from every list-item
    paragraph of one table, or from the single paragraph whose whole text equals *paragraph_text*."""
    if (table_id is None) == (paragraph_text is None):
        raise _make_error(ErrorCode.INVALID_INPUT, "pass exactly one of table_id or paragraph_text")
    if table_id is not None and (isinstance(table_id, bool) or not isinstance(table_id, int) or table_id < 1):
        raise _make_error(ErrorCode.INVALID_INPUT, f"table_id must be a positive integer, got {table_id!r}")
    if paragraph_text is not None and not paragraph_text.strip():
        raise _make_error(ErrorCode.INVALID_INPUT, "paragraph_text must be a non-empty whole-paragraph text")
    wm.require_live_only(path, write_mode, tool="clear_numbering")
    session = wm.live_session_for(path)
    wm.require_capability(session, "paragraphs_unlist", feature_description="clear_numbering")
    pre = comments_live._request(session, "describe")["bodySha256"]
    wm.check_not_stale(revision_before, pre)
    scope: dict[str, Any] = {"table_index": table_id} if table_id is not None else {"paragraph_text": paragraph_text}
    try:
        result = comments_live._request(session, "paragraphs_unlist", {**scope, "expectedBodySha256": pre})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="clear_numbering", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    problems = []
    if not result.get("applied") or not isinstance(result.get("unlisted"), int):
        problems.append("the pane did not report applied=true with an unlisted count")
    if result.get("remaining_list_items"):
        problems.append(f"{result['remaining_list_items']} list item(s) remain in scope")
    if problems:
        wm.audit_live_failure(
            tool="clear_numbering", path=path, document_name=session.document_name,
            pre_body_sha256=pre, post_body_sha256=result.get("post"), reason="; ".join(problems))
        raise _make_error(ErrorCode.VERIFICATION_FAILED,
                          "live numbering removal did not verify: " + "; ".join(problems), {"result": result})
    evidence = wm.live_evidence(
        shape_result=result, applied=True, match_count=result["unlisted"], rung=1,
        before=f"{result['unlisted']} list item(s)", after="0 list item(s)",
        pre_body_sha256=pre, post_body_sha256=result.get("post", pre),
        document_name=session.document_name, tool="clear_numbering", path=path)
    evidence.update({k: result[k] for k in ("unlisted", "remaining_list_items") if k in result})
    evidence["scope"] = {"table_id": table_id} if table_id is not None else {"paragraph_text": paragraph_text}
    return evidence


def delete_table(path: str, table_id: int, expected_text: str, revision_before: str | None = None,
                 track_changes: bool = False, write_mode: str = "live") -> dict[str, Any]:
    """Live-only ``delete_table`` (issue #79): delete one top-level table after the pane has
    compared its cell text with *expected_text* (cells tab-joined, rows newline-joined,
    whitespace-insensitive)."""
    if isinstance(table_id, bool) or not isinstance(table_id, int) or table_id < 1:
        raise _make_error(ErrorCode.INVALID_INPUT, f"table_id must be a positive integer, got {table_id!r}")
    if not isinstance(expected_text, str) or not expected_text.strip():
        raise _make_error(ErrorCode.INVALID_INPUT,
                          "expected_text is required (the table's current text as get_table reports it)")
    wm.require_live_only(path, write_mode, tool="delete_table")
    session = wm.live_session_for(path)
    wm.require_capability(session, "table_delete", feature_description="delete_table")
    pre = comments_live._request(session, "describe")["bodySha256"]
    wm.check_not_stale(revision_before, pre)
    try:
        result = comments_live._request(session, "table_delete", {
            "table_index": table_id, "expected_text": expected_text,
            "track_changes": track_changes, "expectedBodySha256": pre})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="delete_table", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    before_count, after_count = result.get("tables_before"), result.get("tables_after")
    removed_ok = (isinstance(before_count, int) and isinstance(after_count, int)
                  and (after_count == before_count - 1 or (track_changes and after_count == before_count)))
    if not result.get("applied") or not removed_ok:
        wm.audit_live_failure(
            tool="delete_table", path=path, document_name=session.document_name,
            pre_body_sha256=pre, post_body_sha256=result.get("post"),
            reason="the pane did not report the table count dropping by one")
        raise _make_error(ErrorCode.VERIFICATION_FAILED,
                          "live table deletion did not verify: the pane did not report applied=true with the "
                          "table count dropping by exactly one. Nothing was rolled back; inspect the document.",
                          {"result": result})
    evidence = wm.live_evidence(
        shape_result=result, applied=True, match_count=1, rung=1,
        before=result.get("deleted_text", expected_text), after="",
        pre_body_sha256=pre, post_body_sha256=result.get("post", pre),
        document_name=session.document_name, tool="delete_table", path=path)
    evidence.update({k: result[k] for k in ("table_index", "tables_before", "tables_after", "tracked") if k in result})
    evidence["track_changes"] = track_changes
    return evidence


def _spec_problems(specs: list[dict[str, Any]], inserted: list[dict[str, Any]]) -> list[str]:
    """Independent check of the pane's read-back against what was asked."""
    problems: list[str] = []
    if len(inserted) != len(specs):
        return [f"pane reported {len(inserted)} inserted paragraph(s), expected {len(specs)}"]
    for i, (spec, got) in enumerate(zip(specs, inserted, strict=True), start=1):
        if got.get("text") != spec["text"]:
            problems.append(f"paragraph {i} text {got.get('text')!r} != {spec['text']!r}")
        requested = got.get("requested_style") or {}
        if requested.get("kind") == "name":
            if got.get("style") != spec.get("style"):
                problems.append(f"paragraph {i} style {got.get('style')!r} != {spec.get('style')!r}")
        else:
            want = (spec.get("style") or "Normal").replace(" ", "").lower()
            if (got.get("style_builtin") or "").lower() != want:
                problems.append(f"paragraph {i} built-in style {got.get('style_builtin')!r} != {want!r}")
        if spec.get("color") and str(got.get("color") or "").lstrip("#").upper() != spec["color"]:
            problems.append(f"paragraph {i} color {got.get('color')!r} != #{spec['color']}")
    return problems


def insert_paragraphs(path: str, anchor: str, position: str, paragraphs: list[dict[str, Any]],
                      expected_matches: int, revision_before: str | None = None,
                      track_changes: bool = False) -> dict[str, Any]:
    """Live ``insert_paragraphs`` (issue #77): ``paragraph_insert`` over the ops channel.

    *paragraphs* are the already-validated, normalized specs (see
    ``paragraph_insert.validate_insert_request``). The pane re-reads every
    inserted paragraph; this function checks that read-back again on the
    server side before reporting success.
    """
    session = wm.live_session_for(path)
    wm.require_capability(session, "paragraph_insert", feature_description="paragraph insertion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    wm.check_not_stale(revision_before, pre)
    wire = [{key: spec[key] for key in ("text", "style", "color") if spec.get(key) is not None}
            for spec in paragraphs]
    try:
        result = comments_live._request(session, "paragraph_insert", {
            "anchor": anchor, "expected_matches": expected_matches, "position": position,
            "paragraphs": wire, "track_changes": track_changes, "expectedBodySha256": pre})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            wm.audit_live_failure(
                tool="insert_paragraphs", path=path, document_name=session.document_name,
                pre_body_sha256=pre, post_body_sha256=None, reason=exc.envelope.message)
        raise
    inserted = result.get("inserted") or []
    problems = _spec_problems(paragraphs, inserted)
    if not result.get("applied") or result.get("post") == pre:
        problems.append("the pane did not report a changed body")
    if problems:
        wm.audit_live_failure(
            tool="insert_paragraphs", path=path, document_name=session.document_name,
            pre_body_sha256=pre, post_body_sha256=result.get("post"),
            reason="; ".join(problems), detail={"inserted": inserted})
        raise _make_error(
            ErrorCode.VERIFICATION_FAILED,
            "live paragraph insertion did not verify: " + "; ".join(problems) +
            ". Nothing was rolled back by the server; Word owns the document -- inspect it "
            "(and Word's Review pane if tracked) and fix by hand.",
            {"problems": problems, "inserted": inserted})
    evidence = wm.live_evidence(
        shape_result=result, applied=True, match_count=1, rung=1, before=anchor,
        after="\n".join(item["text"] for item in inserted),
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="insert_paragraphs", path=path)
    evidence.update({key: result[key] for key in ("position", "anchor_index", "before_count", "after_count")
                     if key in result})
    evidence["inserted"] = inserted
    evidence["track_changes"] = track_changes
    return evidence
