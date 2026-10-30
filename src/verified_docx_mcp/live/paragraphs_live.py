"""Whole-paragraph deletion through Word, guarded and independently verified."""
from typing import Any

from .. import audit
from ..errors import ErrorCode, VerifyError, _make_error
from . import comments_live
from . import write_mode as wm


def _resolve_anchored_paragraph(session, pre: str, anchor_text: str, expected_length: int | None,
                                find_sha256: str | None) -> str:
    """Issue #113: an anchored ``anchor_text`` ("<prefix>…<suffix>") resolved on
    the live body (a non-observing read) to the whole text of one paragraph,
    as Word's paragraph.text reports it."""
    from .. import locate, projection
    from .reads_live import live_projection

    proj = live_projection(session, expected_body_sha256=pre)
    span = locate.locate_anchored(anchor_text, proj, expected_length=expected_length,
                                  find_sha256=find_sha256).spans[0]
    bounds = locate.paragraph_spans(proj)
    texts = projection.paragraph_texts(proj)
    for i, meta in enumerate(proj.paragraphs):
        if bounds.get(meta.para_ref) == span:
            return texts[i]
    raise _make_error(ErrorCode.INVALID_INPUT,
                      "the anchored anchor_text must cover one whole paragraph, from its first to its last character")


def delete_paragraph(path: str, anchor_text: str, revision_before: str,
                     track_changes: bool = False, expected_length: int | None = None,
                     find_sha256: str | None = None) -> dict[str, Any]:
    from .. import locate

    if not anchor_text or not revision_before.startswith("live:sha256:"):
        raise _make_error(ErrorCode.INVALID_INPUT, "A whole-paragraph anchor and live body revision are required")
    anchored = locate.split_anchored(anchor_text) is not None
    if anchored:
        locate.check_anchored_args(expected_length, find_sha256)
    elif expected_length is not None or find_sha256 is not None:
        raise _make_error(ErrorCode.INVALID_INPUT,
                          "expected_length/find_sha256 apply only to an anchored anchor_text ('<prefix>…<suffix>')")
    session = wm.live_session_for(path)
    wm.require_capability(session, "delete_paragraph", feature_description="whole-paragraph deletion")
    pre = comments_live._request(session, "describe")["bodySha256"]
    wm.check_not_stale(revision_before, pre)
    if anchored:
        anchor_text = _resolve_anchored_paragraph(session, pre, anchor_text, expected_length, find_sha256)
    try:
        result = comments_live._request(session, "paragraph_delete", {
            "anchor_text": anchor_text, "expectedBodySha256": pre, "track_changes": track_changes})
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.VERIFICATION_FAILED:
            audit.append_audit(path=path, tool="delete_paragraph", evidence={
                "applied": None, "write_mode": "live", "error": exc.envelope.to_dict()})
        raise
    # A pane predating compact_evidence_v1 replies with every paragraph before and
    # after. Those go to the audit log (live_evidence's `extra` is merged before the
    # audit write); the caller only ever gets the compact form below.
    audited = {k: result[k] for k in ("before_paragraphs", "after_paragraphs") if k in result}
    evidence = wm.live_evidence(
        shape_result=result,
        applied=True, match_count=1, rung=1, before=anchor_text, after="",
        pre_body_sha256=pre, post_body_sha256=result["post"],
        document_name=session.document_name, tool="delete_paragraph", path=path, extra=audited)
    for key in audited:
        evidence.pop(key, None)
    evidence.update(_compact_delete_result(result, anchor_text))
    return evidence


_EXCERPT_CHARS = 200


def _clip(text: str | None) -> str | None:
    if text is None:
        return None
    return text if len(text) <= _EXCERPT_CHARS else text[:_EXCERPT_CHARS] + "…"


def _compact_delete_result(result: dict[str, Any], anchor_text: str) -> dict[str, Any]:
    """Counts and the two neighbouring paragraphs, never the whole section.

    A current pane (``compact_evidence_v1``) already sends ``neighbours``; for an older
    pane that sent every paragraph, derive the same thing from its lists.
    """
    before = result.get("before_paragraphs")
    after = result.get("after_paragraphs")
    if isinstance(before, list) and isinstance(after, list):
        index = before.index(anchor_text) if anchor_text in before else 0
        neighbours = {
            "previous": before[index - 1] if index > 0 else None,
            "next": before[index + 1] if index + 1 < len(before) else None,
        }
        counts = (len(before), len(after))
    else:
        neighbours = result.get("neighbours") or {"previous": None, "next": None}
        counts = (result.get("before_count"), result.get("after_count"))
    return {
        "paragraphs_before": counts[0],
        "paragraphs_after": counts[1],
        "before_excerpt": _clip(anchor_text),
        "neighbours": {"previous": _clip(neighbours.get("previous")), "next": _clip(neighbours.get("next"))},
    }


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
