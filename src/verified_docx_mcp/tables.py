# New for issue #28 WP-14. GoogleDocs-MCP's tables.py (server.py:826-1006,
# a batchUpdate compiler against the Docs API's table JSON model) is a
# partial analogue: its merged-cell detection/refusal shape
# (_has_merged_cells -> INVALID_INPUT on replace_table_row) is mirrored
# here (see _has_merged_cells/_has_nested_table below and
# execute_replace_table_row's refusal), but the underlying mechanics are
# new -- there is no Docs-API JSON to walk; this module reads/writes real
# ``w:tbl``/``w:tr``/``w:tc``/``w:tcPr`` OOXML directly, and Google's server
# has no analogue at all of replace_cell_markdown (a cell-scoped write that
# leaves w:tcPr untouched) -- Google offers only whole-row/whole-table
# writes, never a single cell, because the Docs API has no operation that
# rewrites one cell's content without also touching its style range. The
# guard/atomic-write/audit SHAPE (guard -> compile -> write -> post-read ->
# evidence -> audit) reuses mutations.py's existing pipeline directly
# (mutations._guard_before_write, mutations._write_and_verify,
# mutations.atomic_replace_docx_parts) rather than reimplementing it.
"""Table tools: ``list_tables``/``get_table`` (read-only) and
``replace_table_row``/``replace_cell_markdown``/``insert_table``
(mutating) -- issue #28 WP-14.

Table/row/cell addressing, throughout this module: ``table_id`` is the
same 1-based, document-order id ``read_document(format="runs")``'s
``table_start``/``table_end`` structural records and ``container_chain``
already use (projection.py's ``_Ids.next_table_id``) -- a nested ``w:tbl``
inside a cell gets its own id, assigned at the point it is encountered in
a full document-order walk, not a "sub-id" of its host table. ``row_index``
and ``cell_index`` are 1-based, matching ``container_chain``'s own
``{"table_id", "row", "cell"}`` convention (projection.py's
``_walk_table``) -- so a caller can go directly from ``read_document``'s
own output to this module's arguments with no re-indexing.

``get_table`` reports ``w:gridSpan``/``w:vMerge`` per cell, per issue #28
plan WP-14's own text. ``replace_table_row`` refuses (``MERGED_OR_NESTED_TABLE``)
the moment the TARGET TABLE (not just the target row) contains ANY merged
cell (``w:gridSpan`` != 1, or a ``w:vMerge``) or a nested ``w:tbl`` --
mirroring GoogleDocs-MCP's own ``_has_merged_cells``/nested-table refusal
in ``execute_replace_table_row`` (whole-table scope, not row scope, "as
Google does" per the plan text). ``replace_cell_markdown`` is the escape
hatch for exactly that case: a cell-scoped write that never touches
``w:tcPr`` (so a merged/nested table's own band-and-border formatting,
carried entirely in ``w:tcPr``, survives byte-identical) -- the only path
this server offers for Appendix-A style band-and-border tables. Nested
multi-level bulleted content inside a cell (issue #28 WP-16a) works via
the SAME ``markdown_to_ooxml.render_blocks`` entry point the body-level
markdown tools already use -- that function carries no body-vs-cell
assumption (see its own module docstring), so ``replace_table_row``'s and
``replace_cell_markdown``'s cell content is built exactly the way
``append_markdown`` builds body content, multi-level ``w:numPr`` included.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, markdown_to_ooxml, mutations, paths, projection, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .projection import DEFAULT_PART, W_NS, TableBoundaryEvent

_CELL_CONTENT_TAGS = frozenset({"p", "tbl", "sdt"})


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Table/row/cell element lookup (operates on a live, already-parsed
# document_root -- project_document_root walks that SAME tree without
# re-parsing, so TableBoundaryEvent.element is a live, mutable Element).
# ---------------------------------------------------------------------------


def _find_table_element(document_root: Any, table_id: int) -> Any:
    proj = projection.project_document_root(document_root)
    for event in proj.events:
        if isinstance(event, TableBoundaryEvent) and event.kind == "table_start" and event.table_id == table_id:
            return event.element
    raise _make_error(
        ErrorCode.TABLE_NOT_FOUND,
        f"No table with table_id {table_id} (call list_tables to enumerate the current ones).",
        {"table_id": table_id},
    )


def _find_row(tbl: Any, row_index: int, *, table_id: int) -> Any:
    rows = [r for r in tbl if projection._ln(r) == "tr"]
    if row_index < 1 or row_index > len(rows):
        raise _make_error(
            ErrorCode.TABLE_ROW_NOT_FOUND,
            f"row_index {row_index} is out of range for table_id {table_id} ({len(rows)} row(s)).",
            {"table_id": table_id, "row_index": row_index, "row_count": len(rows)},
        )
    return rows[row_index - 1]


def _find_cell(tr: Any, cell_index: int, *, table_id: int, row_index: int) -> Any:
    cells = [c for c in tr if projection._ln(c) == "tc"]
    if cell_index < 1 or cell_index > len(cells):
        raise _make_error(
            ErrorCode.TABLE_CELL_NOT_FOUND,
            f"cell_index {cell_index} is out of range for table_id {table_id} row {row_index} ({len(cells)} cell(s)).",
            {"table_id": table_id, "row_index": row_index, "cell_index": cell_index, "cell_count": len(cells)},
        )
    return cells[cell_index - 1]


def _row_cells(tr: Any) -> list[Any]:
    return [c for c in tr if projection._ln(c) == "tc"]


def _cell_tcpr(tc: Any) -> Any | None:
    for child in tc:
        if projection._ln(child) == "tcPr":
            return child
    return None


def _cell_grid_span(tc: Any) -> int:
    tcpr = _cell_tcpr(tc)
    if tcpr is None:
        return 1
    for child in tcpr:
        if projection._ln(child) == "gridSpan":
            val = projection._attr(child, "val")
            if val is not None and val.isdigit():
                return int(val)
    return 1


def _cell_v_merge(tc: Any) -> str:
    """"none" | "restart" | "continue" -- matches projection.py's own
    _table_to_markdown convention exactly: a bare <w:vMerge/> (no val, or
    any val other than "restart") is a continuation cell."""
    tcpr = _cell_tcpr(tc)
    if tcpr is None:
        return "none"
    for child in tcpr:
        if projection._ln(child) == "vMerge":
            val = projection._attr(child, "val")
            if val is not None and val.strip().lower() == "restart":
                return "restart"
            return "continue"
    return "none"


def _has_merged_cells(tbl: Any) -> bool:
    """True if any DIRECT row/cell of *tbl* itself (not a nested table's
    own cells) carries a gridSpan != 1 or a vMerge."""
    for tr in tbl:
        if projection._ln(tr) != "tr":
            continue
        for tc in tr:
            if projection._ln(tc) != "tc":
                continue
            if _cell_grid_span(tc) != 1 or _cell_v_merge(tc) != "none":
                return True
    return False


def _has_nested_table(tbl: Any) -> bool:
    for tr in tbl:
        if projection._ln(tr) != "tr":
            continue
        for tc in tr:
            if projection._ln(tc) != "tc":
                continue
            for child in tc:
                if projection._ln(child) == "tbl":
                    return True
    return False


def _table_stats(tbl: Any) -> tuple[int, int, bool, bool]:
    """(row_count, col_count, has_merged_cells, has_nested_table).
    col_count prefers tblGrid's own gridCol count; falls back to the
    widest row's own gridSpan-summed cell count when tblGrid is absent
    (defensive -- every fixture and every table this module itself builds
    always has one)."""
    rows = [r for r in tbl if projection._ln(r) == "tr"]
    grid = next((c for c in tbl if projection._ln(c) == "tblGrid"), None)
    if grid is not None:
        col_count = sum(1 for c in grid if projection._ln(c) == "gridCol")
    else:
        col_count = 0
        for tr in rows:
            span_sum = sum(_cell_grid_span(tc) for tc in tr if projection._ln(tc) == "tc")
            col_count = max(col_count, span_sum)
    return len(rows), col_count, _has_merged_cells(tbl), _has_nested_table(tbl)


# ---------------------------------------------------------------------------
# Read: list_tables / get_table
# ---------------------------------------------------------------------------


def list_tables_impl(docx_path: Path, part_name: str = DEFAULT_PART) -> list[dict[str, Any]]:
    proj = projection.project_part(docx_path, part_name)
    tables: list[dict[str, Any]] = []
    parent_stack: list[int] = []
    for event in proj.events:
        if not isinstance(event, TableBoundaryEvent):
            continue
        if event.kind == "table_start":
            row_count, col_count, has_merged, has_nested = _table_stats(event.element)
            tables.append(
                {
                    "table_id": event.table_id,
                    "row_count": row_count,
                    "col_count": col_count,
                    "has_merged_cells": has_merged,
                    "has_nested_table": has_nested,
                    "nested_in_table_id": parent_stack[-1] if parent_stack else None,
                }
            )
            parent_stack.append(event.table_id)
        else:
            parent_stack.pop()
    return tables


def execute_list_tables(path: str, part: str = DEFAULT_PART) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    from . import server as _server

    local_path, is_temp = _server._read_local_copy(resolved)
    try:
        return {"path": str(resolved), "part": part, "tables": list_tables_impl(local_path, part)}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


def get_table_impl(docx_path: Path, table_id: int, part_name: str = DEFAULT_PART) -> dict[str, Any]:
    proj = projection.project_part(docx_path, part_name)
    tbl = None
    for event in proj.events:
        if isinstance(event, TableBoundaryEvent) and event.kind == "table_start" and event.table_id == table_id:
            tbl = event.element
            break
    if tbl is None:
        raise _make_error(
            ErrorCode.TABLE_NOT_FOUND,
            f"No table with table_id {table_id} (call list_tables to enumerate the current ones).",
            {"table_id": table_id},
        )
    styles = projection.list_styles_impl(docx_path)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    row_count, col_count, has_merged, has_nested = _table_stats(tbl)

    rows: list[list[dict[str, Any]]] = []
    for row_idx, tr in enumerate((r for r in tbl if projection._ln(r) == "tr"), start=1):
        cells: list[dict[str, Any]] = []
        for cell_idx, tc in enumerate(_row_cells(tr), start=1):
            lossy: list[dict[str, Any]] = []
            text = projection._table_cell_markdown(tc, styles_by_id, table_id, lossy)
            cells.append(
                {
                    "row_index": row_idx,
                    "cell_index": cell_idx,
                    "grid_span": _cell_grid_span(tc),
                    "v_merge": _cell_v_merge(tc),
                    "text": text,
                }
            )
        rows.append(cells)

    return {
        "table_id": table_id,
        "row_count": row_count,
        "col_count": col_count,
        "has_merged_cells": has_merged,
        "has_nested_table": has_nested,
        "rows": rows,
    }


def execute_get_table(path: str, table_id: int, part: str = DEFAULT_PART) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    from . import server as _server

    local_path, is_temp = _server._read_local_copy(resolved)
    try:
        result = get_table_impl(local_path, table_id, part)
        return {"path": str(resolved), "part": part, **result}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Write helpers shared by replace_table_row/replace_cell_markdown/insert_table
# ---------------------------------------------------------------------------


def _cell_content_children(tc: Any) -> list[Any]:
    return [c for c in tc if projection._ln(c) in _CELL_CONTENT_TAGS]


def _markdown_for_elements(elements: list[Any], styles_by_id: dict[str, dict], numbering_index: dict) -> str:
    return projection.markdown_from_elements(elements, styles_by_id, numbering_index)


def _render_cell_markdown(markdown: str, ctx: markdown_to_ooxml.StyleContext) -> list[Any]:
    elements = markdown_to_ooxml.render_blocks(markdown, ctx)
    if not elements:
        # Every w:tc needs at least one block-level child; an empty-string
        # cell renders as one empty paragraph rather than a structurally
        # invalid <w:tc/>.
        elements = [ET.Element(_w("p"))]
    return elements


def _replace_cell_content(
    tc: Any, new_elements: list[Any], *, track: Any | None
) -> list[Any]:
    """Replace *tc*'s own content children (w:p/w:tbl/w:sdt) with
    *new_elements*, NEVER touching w:tcPr (or anything else already in
    tc) -- the load-bearing contract both replace_table_row and
    replace_cell_markdown share. Returns the list of new_elements actually
    appended (same list, for the caller's own bookkeeping)."""
    old_children = _cell_content_children(tc)
    if track:
        tracked_changes.mark_elements_deleted(old_children, track)

        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.wrap_all_runs(new_elements, _ins_wrap)
        for el in new_elements:
            tc.append(el)
    else:
        for child in old_children:
            tc.remove(child)
        for el in new_elements:
            tc.append(el)
    return new_elements


def _evidence(
    *,
    applied: bool,
    match_count: int,
    rung: int,
    before: str,
    after: str,
    revision_before: str,
    revision_after: str,
    audit_logged: bool,
    orphaned_comment_ids: list[str],
    track: Any | None = None,
    conflict_sweep: dict[str, Any] | None = None,
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
    }
    if orphaned_comment_ids:
        evidence["orphaned_comment_ids"] = orphaned_comment_ids
    if track is not None:
        evidence["track_changes"] = True
        evidence["revision_ids"] = track.revision_ids
    if conflict_sweep is not None:
        mutations._merge_conflict_sweep(evidence, conflict_sweep)
    return evidence


# ---------------------------------------------------------------------------
# Tool: replace_table_row
# ---------------------------------------------------------------------------


def execute_replace_table_row(
    path: str,
    table_id: int,
    row_index: int,
    cells: list[str],
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    tbl = _find_table_element(document_root, table_id)

    # Whole-TABLE scope, not row scope -- mirrors GoogleDocs-MCP's own
    # _has_merged_cells refusal (tables.py:101-110/296-301 there): a
    # merged or nested-table cell ANYWHERE in this table blocks a
    # row-level rewrite, even in a different row, "as Google does" (issue
    # #28 plan WP-14). replace_cell_markdown is the escape hatch.
    if _has_merged_cells(tbl) or _has_nested_table(tbl):
        raise _make_error(
            ErrorCode.MERGED_OR_NESTED_TABLE,
            f"table_id {table_id} contains a merged cell (w:gridSpan != 1 or w:vMerge) or a "
            "nested w:tbl; replace_table_row refuses the whole table, as GoogleDocs-MCP's own "
            "replace_table_row does for a merged cell. Use replace_cell_markdown instead -- it "
            "writes one cell at a time and never touches w:tcPr.",
            {"table_id": table_id},
        )

    tr = _find_row(tbl, row_index, table_id=table_id)
    existing_cells = _row_cells(tr)
    if len(cells) != len(existing_cells):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"cells has {len(cells)} entries but table_id {table_id} row {row_index} has "
            f"{len(existing_cells)} cell(s).",
            {"table_id": table_id, "row_index": row_index, "cells_given": len(cells), "cell_count": len(existing_cells)},
        )

    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    before_cells = [
        projection._table_cell_markdown(tc, styles_by_id, table_id, [])
        for tc in existing_cells
    ]
    before_text = "\t".join(before_cells)

    hazards = mutations._scan_range_hazards(existing_cells)
    orphaned_comment_ids = mutations._check_hazards_or_raise(hazards, force, own_author)

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    per_cell_new_elements = [_render_cell_markdown(md, ctx) for md in cells]

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    for tc, new_elements in zip(existing_cells, per_cell_new_elements):
        _replace_cell_content(tc, new_elements, track=track)

    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_after_cells = [
        _markdown_for_elements(elements, styles_by_id, new_numbering_index) for elements in per_cell_new_elements
    ]

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = mutations._load_document(written_path)
        new_tbl = _find_table_element(new_document_root, table_id)
        new_tr = _find_row(new_tbl, row_index, table_id=table_id)
        new_cells = _row_cells(new_tr)
        if len(new_cells) != len(existing_cells):
            raise ValueError(f"re-read row has {len(new_cells)} cells, expected {len(existing_cells)}")
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        written_numbering_index = projection.load_numbering_index(written_path)
        for tc, new_elements, intended in zip(new_cells, per_cell_new_elements, intended_after_cells):
            children = _cell_content_children(tc)
            actual_children = children[-len(new_elements):] if new_elements else []
            actual = _markdown_for_elements(actual_children, new_styles_by_id, written_numbering_index)
            diff = mutations._diff_modulo_whitespace(intended, actual)
            if diff:
                raise ValueError(f"re-read cell does not match the intended rendering modulo whitespace: {diff}")

    conflict_sweep = mutations._write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_text = "\t".join(intended_after_cells)
    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=3,
        before=before_text,
        after=after_text,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=orphaned_comment_ids,
        track=track,
        conflict_sweep=conflict_sweep,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_table_row", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool: replace_cell_markdown
# ---------------------------------------------------------------------------


def execute_replace_cell_markdown(
    path: str,
    table_id: int,
    row_index: int,
    cell_index: int,
    markdown: str,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    tbl = _find_table_element(document_root, table_id)
    tr = _find_row(tbl, row_index, table_id=table_id)
    tc = _find_cell(tr, cell_index, table_id=table_id, row_index=row_index)

    # No merged/nested-table refusal here, deliberately -- this IS the
    # merged-cell-safe path (issue #28 plan WP-14): it never touches
    # w:tcPr, which is where w:gridSpan/w:vMerge/band-and-border
    # formatting live, so it is safe on exactly the tables
    # replace_table_row refuses.
    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    before_text = projection._table_cell_markdown(tc, styles_by_id, table_id, [])

    old_children = _cell_content_children(tc)
    hazards = mutations._scan_range_hazards(old_children)
    orphaned_comment_ids = mutations._check_hazards_or_raise(hazards, force, own_author)

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    # render_blocks carries no body-vs-cell assumption (markdown_to_ooxml.py's
    # own module docstring) -- multi-level bulleted markdown renders here
    # exactly as it would at the body level, sharing one numId/abstractNum
    # tree across levels via increasing w:ilvl (issue #28 WP-16a; see
    # StyleContext.allocate_list_root/ensure_list_level).
    new_elements = _render_cell_markdown(markdown, ctx)

    # tcPr snapshot, purely for the byte-identity acceptance check this
    # tool exists to satisfy -- never read from again, never mutated.
    tcpr_before = _cell_tcpr(tc)
    tcpr_before_bytes = ET.tostring(tcpr_before) if tcpr_before is not None else None

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    _replace_cell_content(tc, new_elements, track=track)

    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_after = _markdown_for_elements(new_elements, styles_by_id, new_numbering_index)

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = mutations._load_document(written_path)
        new_tbl = _find_table_element(new_document_root, table_id)
        new_tr = _find_row(new_tbl, row_index, table_id=table_id)
        new_tc = _find_cell(new_tr, cell_index, table_id=table_id, row_index=row_index)

        tcpr_after = _cell_tcpr(new_tc)
        tcpr_after_bytes = ET.tostring(tcpr_after) if tcpr_after is not None else None
        if tcpr_before_bytes != tcpr_after_bytes:
            raise ValueError(
                "w:tcPr changed across the write -- replace_cell_markdown must leave it byte-identical"
            )

        children = _cell_content_children(new_tc)
        actual_children = children[-len(new_elements):] if new_elements else []
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        written_numbering_index = projection.load_numbering_index(written_path)
        actual = _markdown_for_elements(actual_children, new_styles_by_id, written_numbering_index)
        diff = mutations._diff_modulo_whitespace(intended_after, actual)
        if diff:
            raise ValueError(f"re-read cell does not match the intended rendering modulo whitespace: {diff}")

    conflict_sweep = mutations._write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=3,
        before=before_text,
        after=intended_after,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=orphaned_comment_ids,
        track=track,
        conflict_sweep=conflict_sweep,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_cell_markdown", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool: insert_table
# ---------------------------------------------------------------------------

_DXA_PER_INCH = 1440


def execute_insert_table(
    path: str,
    rows: list[list[str]],
    style_id: str,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    if not rows or any(len(r) == 0 for r in rows):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "rows must be a non-empty list of non-empty cell-markdown lists.", {"rows": rows}
        )

    styles = projection.list_styles_impl(resolved)
    table_style_ids = {s["style_id"] for s in styles if s.get("type") == "table" and s.get("style_id")}
    if style_id not in table_style_ids:
        raise _make_error(
            ErrorCode.STYLE_NOT_FOUND,
            f"style_id {style_id!r} is not a table style (w:type=\"table\") in this document's styles.xml.",
            {"style_id": style_id, "available_table_styles": sorted(table_style_ids)},
        )
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}

    # New table's own eventual table_id: it is appended as the LAST
    # top-level body element (before the trailing sectPr), so in a full
    # document-order walk it gets whatever id comes right after every
    # table (top-level or nested) already present.
    new_table_id = len(list_tables_impl(resolved)) + 1

    document_root, raw_xml = mutations._load_document(resolved)
    body = mutations._find_body(document_root)
    body_children, sect_pr = mutations._split_body(body)

    col_count = max(len(r) for r in rows)

    page_sections = projection.list_page_sections_impl(resolved)
    col_width_dxa: int | None = None
    if page_sections and page_sections[0].get("column_widths_in"):
        usable_in = page_sections[0]["column_widths_in"][0]
        if usable_in:
            col_width_dxa = max(1, round((usable_in * _DXA_PER_INCH) / col_count))

    ctx = markdown_to_ooxml.StyleContext.build(resolved)

    tbl = ET.Element(_w("tbl"))
    tblpr = ET.SubElement(tbl, _w("tblPr"))
    ET.SubElement(tblpr, _w("tblStyle"), {_w("val"): style_id})
    ET.SubElement(tblpr, _w("tblW"), {_w("w"): "0", _w("type"): "auto"})
    grid = ET.SubElement(tbl, _w("tblGrid"))
    for _ in range(col_count):
        attrs = {_w("w"): str(col_width_dxa)} if col_width_dxa else {}
        ET.SubElement(grid, _w("gridCol"), attrs)

    per_row_new_elements: list[list[list[Any]]] = []
    for row_cells in rows:
        tr = ET.SubElement(tbl, _w("tr"))
        row_elements: list[list[Any]] = []
        for c in range(col_count):
            tc = ET.SubElement(tr, _w("tc"))
            tcpr = ET.SubElement(tc, _w("tcPr"))
            if col_width_dxa:
                ET.SubElement(tcpr, _w("tcW"), {_w("w"): str(col_width_dxa), _w("type"): "dxa"})
            md = row_cells[c] if c < len(row_cells) else ""
            elements = _render_cell_markdown(md, ctx)
            for el in elements:
                tc.append(el)
            row_elements.append(elements)
        per_row_new_elements.append(row_elements)

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # Nothing existing is removed by an insertion -- only the new
        # table's own runs are wrapped in w:ins (mirrors append_markdown's
        # identical treatment of a pure insertion).
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.wrap_all_runs([tbl], _ins_wrap)

    insert_at = len(body_children)
    body.insert(insert_at, tbl)
    if sect_pr is not None:
        body.remove(sect_pr)
        body.append(sect_pr)

    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_rows = [
        [_markdown_for_elements(elements, styles_by_id, new_numbering_index) for elements in row_elements]
        for row_elements in per_row_new_elements
    ]

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = mutations._load_document(written_path)
        new_tbl = _find_table_element(new_document_root, new_table_id)
        new_row_count, new_col_count, _has_merged, _has_nested = _table_stats(new_tbl)
        if new_row_count != len(rows) or new_col_count != col_count:
            raise ValueError(
                f"re-read table has {new_row_count}x{new_col_count}, expected {len(rows)}x{col_count}"
            )
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        written_numbering_index = projection.load_numbering_index(written_path)
        new_trs = [r for r in new_tbl if projection._ln(r) == "tr"]
        for tr_elem, row_elements, intended_row in zip(new_trs, per_row_new_elements, intended_rows):
            new_cells = _row_cells(tr_elem)
            for tc, elements, intended in zip(new_cells, row_elements, intended_row):
                children = _cell_content_children(tc)
                actual_children = children[-len(elements):] if elements else []
                actual = _markdown_for_elements(actual_children, new_styles_by_id, written_numbering_index)
                diff = mutations._diff_modulo_whitespace(intended, actual)
                if diff:
                    raise ValueError(f"re-read cell does not match the intended rendering modulo whitespace: {diff}")

    conflict_sweep = mutations._write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_text = "\n".join("\t".join(row) for row in intended_rows)
    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=4,
        before="",
        after=after_text,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=[],
        track=track,
        conflict_sweep=conflict_sweep,
    )
    evidence["table_id"] = new_table_id
    logged, _ = audit.append_audit(path=str(resolved), tool="insert_table", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
