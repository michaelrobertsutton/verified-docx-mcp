"""Whole-paragraph deletion through Word, guarded and independently verified."""
from typing import Any

from .. import audit
from ..errors import ErrorCode, VerifyError, _make_error
from . import comments_live, write_mode


def delete_paragraph(path: str, anchor_text: str, revision_before: str,
                     track_changes: bool = False) -> dict[str, Any]:
    if not anchor_text or not revision_before.startswith("live:sha256:"):
        raise _make_error(ErrorCode.INVALID_INPUT, "A whole-paragraph anchor and live body revision are required")
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "delete_paragraph", feature_description="whole-paragraph deletion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    write_mode.check_not_stale(revision_before, pre)
    try:
        result = comments_live._request(session, "paragraph_delete", {
            "anchor_text": anchor_text, "expectedBodySha256": pre, "track_changes": track_changes})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="delete_paragraph", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    evidence = write_mode.live_evidence(
        shape_result=result,
        applied=True, match_count=1, rung=1, before=anchor_text, after="",
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="delete_paragraph", path=path)
    evidence.update(result)
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
    session = write_mode.live_session_for(path)
    write_mode.require_capability(session, "paragraph_insert", feature_description="paragraph insertion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    write_mode.check_not_stale(revision_before, pre)
    wire = [{key: spec[key] for key in ("text", "style", "color") if spec.get(key) is not None}
            for spec in paragraphs]
    try:
        result = comments_live._request(session, "paragraph_insert", {
            "anchor": anchor, "expected_matches": expected_matches, "position": position,
            "paragraphs": wire, "track_changes": track_changes, "expectedBodySha256": pre})
    except VerifyError as exc:
        if exc.envelope.diagnostics.get("pane_error_code") == "stale":
            raise _make_error(ErrorCode.LIVE_STALE, exc.envelope.message,
                              exc.envelope.diagnostics) from exc
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            write_mode.audit_live_failure(
                tool="insert_paragraphs", path=path, document_name=session.document_name,
                pre_body_sha256=pre, post_body_sha256=None, reason=exc.envelope.message)
        raise
    inserted = result.get("inserted") or []
    problems = _spec_problems(paragraphs, inserted)
    if not result.get("applied") or result.get("post") == pre:
        problems.append("the pane did not report a changed body")
    if problems:
        write_mode.audit_live_failure(
            tool="insert_paragraphs", path=path, document_name=session.document_name,
            pre_body_sha256=pre, post_body_sha256=result.get("post"),
            reason="; ".join(problems), detail={"inserted": inserted})
        raise _make_error(
            ErrorCode.VERIFICATION_FAILED,
            "live paragraph insertion did not verify: " + "; ".join(problems) +
            ". Nothing was rolled back by the server; Word owns the document -- inspect it "
            "(and Word's Review pane if tracked) and fix by hand.",
            {"problems": problems, "inserted": inserted})
    evidence = write_mode.live_evidence(
        shape_result=result, applied=True, match_count=1, rung=1, before=anchor,
        after="\n".join(item["text"] for item in inserted),
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="insert_paragraphs", path=path)
    evidence.update({key: result[key] for key in ("position", "anchor_index", "before_count", "after_count")
                     if key in result})
    evidence["inserted"] = inserted
    evidence["track_changes"] = track_changes
    return evidence
