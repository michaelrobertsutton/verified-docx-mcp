# New for issue #28 WP-06. No GoogleDocs-MCP module maps onto this one
# wholesale: that server's replace_text/format_text (mutations.py,
# formatting.py) compile Docs API batchUpdate requests (deleteContentRange +
# insertText, or updateTextStyle) against a server-assigned index space --
# there is no run-splitting to do, because the Docs API itself owns run
# boundaries. This backend edits the OOXML w:r/w:t tree directly, so THIS
# module's own job -- splitting a run that a match's boundary falls in the
# middle of, cloning its w:rPr verbatim onto the surviving pieces, and
# inheriting the first run's w:rPr onto a freshly-inserted replacement run --
# has no Google-side analogue to lift. What IS lifted: locate.py's
# normalization ladder (from verify.py, see that module's own header), and
# the guard/atomic-write/audit SHAPE of mutations.py's markdown-mutation
# pipeline (guard -> compile -> write -> post-read -> assemble evidence ->
# audit), applied here to a per-match edit instead of a whole-body/section
# rewrite.
"""Guarded pipeline for the two text-matching mutating tools: replace_text
(delete + reinsert, rung-agnostic to the caller) and format_text (style-only,
no content change).

Both share:
  1. locate.locate() over projection.project_document_root(document_root) --
     the normalization ladder, STRUCTURAL_BOUNDARY refusal, and
     crosses_comment_range/crosses_revision warnings (this WP only WARNS;
     WP-07 adds the actual TRACKED_CHANGES_PRESENT refusal on top of this
     same detection -- see tracked_changes.py).
  2. Run splitting (_atoms_for_span + the per-tool apply function below):
     a run whose text-bearing child (almost always w:t; a w:tab/w:br/
     w:noBreakHyphen/w:softHyphen/w:sym node is exactly one character wide,
     so a match boundary can only ever land AT its edges, never inside it)
     straddles a match boundary splits into up to three pieces -- unmatched
     prefix (the original element, shortened in place), the matched middle,
     and unmatched suffix (a NEW sibling element, its w:rPr CLONED VERBATIM
     from the original) -- never touching a run entirely outside the match.
  3. The same guard/atomic-write/audit machinery mutations.py already built
     (lock/sync/revision guard, atomic_replace_docx_parts with its .jsbak
     rollback, append_audit) -- reused directly, not reimplemented.
  4. The eight standard evidence keys, plus runs_before/runs_after (those
     exact names -- the lifted audit redaction set, audit._AUDIT_REDACTED_KEYS,
     already covers them) and, when present, a non-fatal `warnings` list.

Scope limit, documented rather than silently mishandled: a run whose parent
holds it via anything other than direct list membership one level up (the
common case for every fixture in this repo) is handled generically via
element identity (parent_elem.insert/.remove), so nesting inside a
hyperlink/w:ins/w:smartTag/inline w:sdt works the same as a bare paragraph
child. What is NOT specially handled: two matches (expected_matches > 1)
whose spans are directly adjacent with no unmatched character between them,
sharing one boundary run atom -- mutating that shared atom for the first
match could invalidate the second match's own element reference. Untested
and undocumented as supported; every fixture and acceptance test in this WP
uses either a single match or matches separated by ordinary unmatched text.
"""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection
from .errors import ErrorCode, _make_error
from .locate import LocateResult, locate
from .projection import RunEvent, W_NS

_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Style allowlist (format_text) -- our own rPr model carries strike in
# addition to Google's bold/italic/underline (projection._run_properties
# already reads/writes all four), so this server's format_text supports one
# more field than the Google original.
# ---------------------------------------------------------------------------

_STYLE_ALLOWLIST = ("bold", "italic", "underline", "strike")
_BOOL_TOGGLE_TAGS = {"bold": "b", "italic": "i", "strike": "strike"}

_EXCERPT_RADIUS = 200


def _validate_style(style: Any) -> dict[str, bool]:
    if not isinstance(style, dict):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "style must be an object mapping bold/italic/underline/strike to true/false",
            {"style": repr(style)},
        )
    if not style:
        raise _make_error(ErrorCode.INVALID_INPUT, "style must not be empty")
    unknown = sorted(set(style) - set(_STYLE_ALLOWLIST))
    if unknown:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"unknown style key(s): {unknown}; allowed: {list(_STYLE_ALLOWLIST)}",
            {"unknown_keys": unknown},
        )
    non_bool = {k: repr(v) for k, v in style.items() if type(v) is not bool}
    if non_bool:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"style values must be true/false booleans, got: {non_bool}",
            {"invalid_values": non_bool},
        )
    return style


# ---------------------------------------------------------------------------
# Run atoms: RunEvents overlapping a span, in document order.
# ---------------------------------------------------------------------------


def _atoms_for_span(proj: projection.Projection, start: int, end: int) -> list[tuple[int, int, RunEvent]]:
    run_events = [e for e in proj.events if isinstance(e, RunEvent)]
    atoms = [
        (s, e, event)
        for event, (s, e, _pr, _rr) in zip(run_events, proj.offset_map)
        if s < end and e > start
    ]
    atoms.sort(key=lambda a: a[0])
    return atoms


def _excerpt(text: str, start: int, end: int, radius: int = _EXCERPT_RADIUS) -> str:
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    return text[lo:hi]


def _run_dict(event: RunEvent, clip_start: int, clip_end: int, node_start: int) -> dict[str, Any]:
    return {
        "text": event.text[clip_start - node_start : clip_end - node_start],
        "start": clip_start,
        "end": clip_end,
        "bold": bool(event.rpr.get("bold", False)),
        "italic": bool(event.rpr.get("italic", False)),
        "underline": bool(event.rpr.get("underline")),
        "strike": bool(event.rpr.get("strike", False)),
    }


def _collect_style_runs(proj: projection.Projection, spans: list[tuple[int, int]]) -> list[list[dict[str, Any]]]:
    """Per span, the run(s) overlapping it, clipped to the span -- the
    docx analogue of GoogleDocs-MCP verify.py's _collect_style_runs. A
    `find` straddling two runs (e.g. half already bold) surfaces both, so a
    caller sees the prior run boundaries directly rather than one flattened
    style (same rationale as the Google original)."""
    result: list[list[dict[str, Any]]] = []
    for start, end in spans:
        span_runs = [_run_dict(event, max(s, start), min(e, end), s) for s, e, event in _atoms_for_span(proj, start, end)]
        result.append(span_runs)
    return result


def _style_matches(runs_before: list[list[dict[str, Any]]], style: dict[str, bool]) -> bool:
    return all(
        run.get(field) == value
        for span_runs in runs_before
        for run in span_runs
        for field, value in style.items()
    )


# ---------------------------------------------------------------------------
# Element construction / style toggling
# ---------------------------------------------------------------------------


def _build_run(text_value: str, rpr_elem: Any | None) -> Any:
    r = ET.Element(_w("r"))
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = ET.SubElement(r, _w("t"))
    t.text = text_value
    if text_value == "" or text_value != text_value.strip():
        t.set(f"{{{_XML_NS}}}space", "preserve")
    return r


def _set_node_text(event: RunEvent, new_text: str) -> None:
    """Mutate the live text-bearing element's own text to *new_text*.

    Only a <w:t> node's .text is ever actually reassigned: every other
    kind this projects (w:tab/w:br/w:noBreakHyphen/w:softHyphen/w:sym) is
    exactly one character wide, so a caller's own before/after slicing can
    only ever produce "" (route to removal instead, never here) or that
    node's own full original text (a no-op) for one of those -- there is
    no partial-character case to handle.
    """
    if event.text_elem is None:
        return
    if projection._ln(event.text_elem) == "t":
        event.text_elem.text = new_text
        if new_text == "" or new_text != new_text.strip():
            event.text_elem.set(f"{{{_XML_NS}}}space", "preserve")


def _toggle_style(r_elem: Any, style: dict[str, bool]) -> None:
    rpr = None
    for child in r_elem:
        if projection._ln(child) == "rPr":
            rpr = child
            break
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        r_elem.insert(0, rpr)

    for field, value in style.items():
        if field == "underline":
            existing = None
            for child in rpr:
                if projection._ln(child) == "u":
                    existing = child
                    break
            if value:
                if existing is None:
                    ET.SubElement(rpr, _w("u"), {_w("val"): "single"})
                else:
                    existing.set(_w("val"), "single")
            elif existing is not None:
                rpr.remove(existing)
            continue

        tag = _BOOL_TOGGLE_TAGS[field]
        existing = None
        for child in rpr:
            if projection._ln(child) == tag:
                existing = child
                break
        if value:
            if existing is None:
                ET.SubElement(rpr, _w(tag))
            else:
                for key in list(existing.attrib):
                    del existing.attrib[key]
        elif existing is not None:
            rpr.remove(existing)


# ---------------------------------------------------------------------------
# replace_text: run splitting + single-replacement-run insertion
# ---------------------------------------------------------------------------


def _apply_replace_span(start: int, end: int, replace_value: str, atoms: list[tuple[int, int, RunEvent]]) -> None:
    """Delete [start, end)'s text from the live tree and insert exactly one
    new <w:r> carrying *replace_value*, its w:rPr cloned verbatim from the
    FIRST overlapping run's ORIGINAL (pre-mutation) w:rPr (issue #28 plan
    WP-06: "replacement text inherits the first run's w:rPr")."""
    if not atoms:
        raise _make_error(ErrorCode.INVALID_INPUT, "match span has no overlapping run content", {"start": start, "end": end})

    first_s, first_e, first_event = atoms[0]
    last_s, last_e, last_event = atoms[-1]

    first_rpr = None
    if first_event.r_elem is not None:
        for child in first_event.r_elem:
            if projection._ln(child) == "rPr":
                first_rpr = child
                break
    cloned_rpr_for_replacement = copy.deepcopy(first_rpr) if first_rpr is not None else None

    # Middle atoms are always fully inside [start, end) -- see this
    # module's docstring / locate.py's structural-boundary guarantee (every
    # atom here shares one paragraph) -- remove outright.
    for _s, _e, event in atoms[1:-1]:
        if event.parent_elem is not None and event.r_elem is not None:
            _remove_if_present(event.parent_elem, event.r_elem)

    if last_event is not first_event:
        local_end = min(last_e, end) - last_s
        after_text = last_event.text[local_end:]
        if after_text:
            _set_node_text(last_event, after_text)
        elif last_event.parent_elem is not None and last_event.r_elem is not None:
            _remove_if_present(last_event.parent_elem, last_event.r_elem)

    local_start = max(start, first_s) - first_s
    before_text = first_event.text[:local_start]
    if first_event is last_event:
        local_end = min(end, first_e) - first_s
        after_text = first_event.text[local_end:]
    else:
        after_text = ""

    new_run = _build_run(replace_value, cloned_rpr_for_replacement)
    parent = first_event.parent_elem
    if parent is None or first_event.r_elem is None:
        raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to edit", {"start": start, "end": end})

    if before_text:
        _set_node_text(first_event, before_text)
        idx = list(parent).index(first_event.r_elem) + 1
        parent.insert(idx, new_run)
        if after_text:
            after_run = _build_run(after_text, copy.deepcopy(first_rpr) if first_rpr is not None else None)
            parent.insert(idx + 1, after_run)
    else:
        idx = list(parent).index(first_event.r_elem)
        parent.insert(idx, new_run)
        if after_text:
            after_run = _build_run(after_text, copy.deepcopy(first_rpr) if first_rpr is not None else None)
            parent.insert(idx + 1, after_run)
        _remove_if_present(parent, first_event.r_elem)


def _remove_if_present(parent: Any, elem: Any) -> None:
    for child in list(parent):
        if child is elem:
            parent.remove(elem)
            return


# ---------------------------------------------------------------------------
# format_text: run splitting + rPr toggling (no content change)
# ---------------------------------------------------------------------------


def _apply_format_span(start: int, end: int, style: dict[str, bool], atoms: list[tuple[int, int, RunEvent]]) -> None:
    for s, e, event in atoms:
        local_start = max(start, s) - s
        local_end = min(end, e) - s
        full_text = event.text
        before_text = full_text[:local_start]
        after_text = full_text[local_end:]

        if not before_text and not after_text:
            # Entirely inside the span -- toggle in place, no split.
            if event.r_elem is not None:
                _toggle_style(event.r_elem, style)
            continue

        if event.r_elem is None or event.parent_elem is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to edit", {"start": start, "end": end})

        original_rpr = None
        for child in event.r_elem:
            if projection._ln(child) == "rPr":
                original_rpr = child
                break
        middle_text = full_text[local_start:local_end]
        middle_run = _build_run(middle_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
        _toggle_style(middle_run, style)

        parent = event.parent_elem
        idx = list(parent).index(event.r_elem)
        if before_text and after_text:
            _set_node_text(event, before_text)
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx + 1, middle_run)
            parent.insert(idx + 2, after_run)
        elif before_text:
            _set_node_text(event, before_text)
            parent.insert(idx + 1, middle_run)
        else:  # after_text only
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx, middle_run)
            parent.insert(idx + 1, after_run)
            _remove_if_present(parent, event.r_elem)


# ---------------------------------------------------------------------------
# Shared write core
# ---------------------------------------------------------------------------


def _serialize_and_write(resolved: Path, document_root: Any, raw_xml: bytes, *, post_verify) -> None:
    document_decls = mutations._capture_source_namespaces(raw_xml)
    new_xml_bytes = mutations._serialize_xml(document_root, document_decls)
    overrides = {projection.DEFAULT_PART: new_xml_bytes}
    mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=post_verify)


def _evidence(
    *,
    applied: bool,
    match_count: int,
    rung: str,
    before: str,
    after: str,
    revision_before: str,
    revision_after: str,
    audit_logged: bool,
    runs_before: list[list[dict[str, Any]]],
    runs_after: list[list[dict[str, Any]]],
    warnings: list[str],
) -> dict[str, Any]:
    evidence: dict[str, Any] = {
        "applied": applied,
        "match_count": match_count,
        "rung": rung,
        "before": before,
        "after": after,
        "revision_before": revision_before,
        "revision_after": revision_after,
        "audit_logged": audit_logged,
        "runs_before": runs_before,
        "runs_after": runs_after,
    }
    if warnings:
        evidence["warnings"] = warnings
    return evidence


# ---------------------------------------------------------------------------
# Tool 1: replace_text
# ---------------------------------------------------------------------------


def execute_replace_text(
    path: str,
    find: str,
    replace: str,
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = locate(find, proj, expected_matches)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    new_run_t_elems: list[Any] = []
    for start, end in locate_result.spans:
        atoms = _atoms_for_span(proj, start, end)
        _apply_replace_span(start, end, replace, atoms)

    # Re-walk the SAME live tree (not a re-parse) to find where the
    # inserted replacement text landed, for the after excerpt/runs_after,
    # and to build the "intended after" text the post-write re-read is
    # checked against below.
    mutated_proj = projection.project_document_root(document_root)
    after_excerpt = _excerpt(mutated_proj.text, before_first[0], before_first[0] + len(replace))
    runs_after = _collect_style_runs(mutated_proj, [(before_first[0], before_first[0] + len(replace))])
    intended_after_text = mutated_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")

    _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence = _evidence(
        applied=True,
        match_count=locate_result.match_count,
        rung=locate_result.rung,
        before=before_excerpt,
        after=after_excerpt,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        runs_before=runs_before,
        runs_after=runs_after,
        warnings=locate_result.warnings,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_text", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool 2: format_text
# ---------------------------------------------------------------------------


def execute_format_text(
    path: str,
    find: str,
    style: dict[str, bool],
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
) -> dict[str, Any]:
    style = _validate_style(style)
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = locate(find, proj, expected_matches)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    # No-op check: an idempotent re-run must not create a new revision (the
    # same rationale as GoogleDocs-MCP's format_text -- see formatting.py).
    if _style_matches(runs_before, style):
        after_excerpt = before_excerpt
        evidence = _evidence(
            applied=True,
            match_count=locate_result.match_count,
            rung=locate_result.rung,
            before=before_excerpt,
            after=after_excerpt,
            revision_before=pre_revision["token"],
            revision_after=pre_revision["token"],
            audit_logged=False,
            runs_before=runs_before,
            runs_after=runs_before,
            warnings=locate_result.warnings,
        )
        logged, _ = audit.append_audit(path=str(resolved), tool="format_text", evidence=evidence)
        evidence["audit_logged"] = logged
        return evidence

    for start, end in locate_result.spans:
        atoms = _atoms_for_span(proj, start, end)
        _apply_format_span(start, end, style, atoms)

    # format_text never changes character counts, so the spans are stable
    # across the mutation -- no need to re-locate to find them again.
    mutated_proj = projection.project_document_root(document_root)
    after_excerpt = _excerpt(mutated_proj.text, before_first[0], before_first[1])
    runs_after = _collect_style_runs(mutated_proj, locate_result.spans)
    intended_after_text = mutated_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")
        post_proj = projection.project_part(written_path)
        post_runs = _collect_style_runs(post_proj, locate_result.spans)
        if not _style_matches(post_runs, style):
            raise ValueError(f"re-read style does not match the requested style: {post_runs}")

    _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence = _evidence(
        applied=True,
        match_count=locate_result.match_count,
        rung=locate_result.rung,
        before=before_excerpt,
        after=after_excerpt,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        runs_before=runs_before,
        runs_after=runs_after,
        warnings=locate_result.warnings,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="format_text", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
