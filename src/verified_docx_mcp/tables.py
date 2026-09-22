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

import re
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, locate, markdown_to_ooxml, mutations, paths, projection, tracked_changes
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
        # issue #154: tables.py has no live-mode path at all -- every
        # write here is file mode.
        "write_mode": "file",
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
# Tool: insert_table -- structured cell objects (issue #100, D1 Option 1):
# a cell may be a plain markdown string (today's behavior, unchanged) or a
# CellSpec dict {"markdown", "span", "v_merge", "fill", "color", "bold",
# "align", "valign"}. See server.py's insert_table docstring for the public
# contract; the helpers below implement it.
# ---------------------------------------------------------------------------

_DXA_PER_INCH = 1440

_HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
_ALIGN_VALUES = frozenset({"left", "center", "right", "both"})
_VALIGN_VALUES = frozenset({"top", "center", "bottom"})
_VMERGE_VALUES = frozenset({"restart", "continue"})

# Schema child order (CT_TcPrBase, subset this module writes), CT_PPrBase
# (subset: pStyle/numPr before jc), and CT_RPrBase (subset: rStyle before
# b/bCs before i/iCs before color) -- Word is picky about this even though
# ElementTree itself is not; _insert_ordered keeps every new child in the
# right slot regardless of call order.
_TCPR_CHILD_ORDER = ["tcW", "gridSpan", "vMerge", "shd", "vAlign"]
_PPR_CHILD_ORDER = ["pStyle", "numPr", "jc"]
_RPR_CHILD_ORDER = ["rStyle", "b", "bCs", "i", "iCs", "color"]


def _insert_ordered(parent: Any, new_child: Any, order: list[str]) -> None:
    """Insert *new_child* into *parent* at the position *order* (a list of
    local tag names, schema order) says it belongs, regardless of what is
    already present or the order this function is called in. A child whose
    tag is not in *order* is always appended last (never expected here --
    every tag this module inserts via this helper is in one of the three
    ORDER lists above)."""
    new_rank = order.index(projection._ln(new_child))
    insert_at = len(list(parent))
    for i, existing in enumerate(parent):
        existing_ln = projection._ln(existing)
        if existing_ln not in order:
            continue
        if order.index(existing_ln) > new_rank:
            insert_at = i
            break
    parent.insert(insert_at, new_child)


def _first_child(parent: Any, tag: str) -> Any | None:
    for child in parent:
        if projection._ln(child) == tag:
            return child
    return None


def _validate_hex(value: Any, field_name: str, row_index: int, cell_index: int) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not _HEX_COLOR_RE.match(value):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: {field_name!r} must be a 6-hex-digit color string "
            f"(e.g. \"3B3838\"), got {value!r}.",
            {"row_index": row_index, "cell_index": cell_index, field_name: value},
        )
    return value.upper()


def _normalize_cell(raw: Any, *, row_index: int, cell_index: int) -> dict[str, Any]:
    """A plain str means {"markdown": str} with every other key at its
    default (span=1, no merge/fill/formatting) -- today's behavior. A dict
    is validated and filled out to the same canonical key set."""
    if isinstance(raw, str):
        return {
            "markdown": raw,
            "span": 1,
            "v_merge": None,
            "fill": None,
            "color": None,
            "bold": None,
            "align": None,
            "valign": None,
        }
    if not isinstance(raw, dict):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: a cell must be a markdown string or a cell-spec object, "
            f"got {type(raw).__name__}.",
            {"row_index": row_index, "cell_index": cell_index},
        )
    markdown = raw.get("markdown")
    if not isinstance(markdown, str):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: a cell-spec object requires a string 'markdown' key.",
            {"row_index": row_index, "cell_index": cell_index},
        )
    span = raw.get("span", 1)
    if isinstance(span, bool) or not isinstance(span, int) or span < 1:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: 'span' must be an integer >= 1, got {span!r}.",
            {"row_index": row_index, "cell_index": cell_index, "span": span},
        )
    v_merge = raw.get("v_merge")
    if v_merge is not None and v_merge not in _VMERGE_VALUES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: 'v_merge' must be \"restart\", \"continue\", or omitted, "
            f"got {v_merge!r}.",
            {"row_index": row_index, "cell_index": cell_index, "v_merge": v_merge},
        )
    bold = raw.get("bold")
    if bold is not None and not isinstance(bold, bool):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: 'bold' must be a boolean, got {bold!r}.",
            {"row_index": row_index, "cell_index": cell_index, "bold": bold},
        )
    align = raw.get("align")
    if align is not None and align not in _ALIGN_VALUES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: 'align' must be one of {sorted(_ALIGN_VALUES)}, got {align!r}.",
            {"row_index": row_index, "cell_index": cell_index, "align": align},
        )
    valign = raw.get("valign")
    if valign is not None and valign not in _VALIGN_VALUES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"row {row_index} cell {cell_index}: 'valign' must be one of {sorted(_VALIGN_VALUES)}, got {valign!r}.",
            {"row_index": row_index, "cell_index": cell_index, "valign": valign},
        )
    return {
        "markdown": markdown,
        "span": span,
        "v_merge": v_merge,
        "fill": _validate_hex(raw.get("fill"), "fill", row_index, cell_index),
        "color": _validate_hex(raw.get("color"), "color", row_index, cell_index),
        "bold": bold,
        "align": align,
        "valign": valign,
    }


def _validate_vmerge_chain(norm_rows: list[list[dict[str, Any]]]) -> None:
    """A "continue" cell must have empty markdown, and the same grid
    column (by col_start, already annotated on every cell) in the row
    above must itself be "restart" or "continue" -- otherwise there is
    nothing for it to continue."""
    open_cols: dict[int, str] = {}
    for row_index, row in enumerate(norm_rows):
        new_open: dict[int, str] = {}
        for cell_index, cell in enumerate(row):
            v_merge = cell["v_merge"]
            if v_merge == "continue":
                if cell["markdown"] != "":
                    raise _make_error(
                        ErrorCode.INVALID_INPUT,
                        f"row {row_index} cell {cell_index}: a v_merge=\"continue\" cell must have empty markdown.",
                        {"row_index": row_index, "cell_index": cell_index},
                    )
                if open_cols.get(cell["col_start"]) not in ("restart", "continue"):
                    raise _make_error(
                        ErrorCode.INVALID_INPUT,
                        f"row {row_index} cell {cell_index}: v_merge=\"continue\" has no \"restart\" (or "
                        "\"continue\") cell at the same grid column in the row above.",
                        {"row_index": row_index, "cell_index": cell_index, "col_start": cell["col_start"]},
                    )
                new_open[cell["col_start"]] = "continue"
            elif v_merge == "restart":
                new_open[cell["col_start"]] = "restart"
        open_cols = new_open


def _normalize_rows(rows: list[list[Any]], grid_dxa: list[int] | None) -> tuple[list[list[dict[str, Any]]], int]:
    """Pure-string rows keep today's exact behavior: col_count is the
    widest row, short rows are padded with empty cells (no grid
    validation -- every span is 1, so a row's span-sum always equals
    col_count once padded). The moment ANY cell in the table uses the
    dict form, grid validation is strict instead: every row's cell spans
    must sum to col_count exactly (from grid_dxa's length when given,
    else the widest span-sum) -- no padding, a mismatch is INVALID_INPUT
    naming the row and both numbers."""
    structured = any(isinstance(c, dict) for row in rows for c in row)
    if structured:
        norm_rows = [
            [_normalize_cell(c, row_index=i, cell_index=j) for j, c in enumerate(row)]
            for i, row in enumerate(rows)
        ]
        col_count = len(grid_dxa) if grid_dxa else max(sum(c["span"] for c in row) for row in norm_rows)
        for row_index, row in enumerate(norm_rows):
            row_sum = sum(c["span"] for c in row)
            if row_sum != col_count:
                raise _make_error(
                    ErrorCode.INVALID_INPUT,
                    f"row {row_index}: cell spans sum to {row_sum}, but the table grid has {col_count} "
                    "column(s); every row's cell spans must sum to the grid column count.",
                    {"row_index": row_index, "row_span_sum": row_sum, "col_count": col_count},
                )
    else:
        col_count = len(grid_dxa) if grid_dxa else max(len(r) for r in rows)
        norm_rows = []
        for row_index, row in enumerate(rows):
            if len(row) > col_count:
                raise _make_error(
                    ErrorCode.INVALID_INPUT,
                    f"row {row_index} has {len(row)} cell(s), more than the {col_count}-column grid_dxa given.",
                    {"row_index": row_index, "row_cell_count": len(row), "col_count": col_count},
                )
            padded = list(row) + [""] * (col_count - len(row))
            norm_rows.append(
                [_normalize_cell(c, row_index=row_index, cell_index=j) for j, c in enumerate(padded)]
            )

    for row in norm_rows:
        col = 0
        for cell in row:
            cell["col_start"] = col
            col += cell["span"]

    _validate_vmerge_chain(norm_rows)
    return norm_rows, col_count


def _apply_cell_formatting(elements: list[Any], cell: dict[str, Any]) -> None:
    """align/bold/color live on the rendered runs/paragraphs themselves
    (w:pPr/w:jc, w:rPr/w:b+w:bCs+w:color) -- see the docstring caveat on
    execute_insert_table: replace_cell_markdown replaces these elements
    wholesale, so this formatting does NOT survive a later cell edit
    (unlike span/v_merge/fill/valign, which live in w:tcPr and do)."""
    align = cell["align"]
    bold = cell["bold"]
    color = cell["color"]
    if align is None and not bold and color is None:
        return
    for el in elements:
        if projection._ln(el) != "p":
            continue
        if align is not None:
            ppr = _first_child(el, "pPr")
            if ppr is None:
                ppr = ET.Element(_w("pPr"))
                el.insert(0, ppr)
            jc = _first_child(ppr, "jc")
            if jc is None:
                jc = ET.Element(_w("jc"))
                _insert_ordered(ppr, jc, _PPR_CHILD_ORDER)
            jc.set(_w("val"), align)
        if bold or color:
            for r in el:
                if projection._ln(r) != "r":
                    continue
                rpr = _first_child(r, "rPr")
                if rpr is None:
                    rpr = ET.Element(_w("rPr"))
                    r.insert(0, rpr)
                if bold:
                    if _first_child(rpr, "b") is None:
                        _insert_ordered(rpr, ET.Element(_w("b")), _RPR_CHILD_ORDER)
                    if _first_child(rpr, "bCs") is None:
                        _insert_ordered(rpr, ET.Element(_w("bCs")), _RPR_CHILD_ORDER)
                if color:
                    c = _first_child(rpr, "color")
                    if c is None:
                        c = ET.Element(_w("color"))
                        _insert_ordered(rpr, c, _RPR_CHILD_ORDER)
                    c.set(_w("val"), color)


def _texts_match_ladder(a: str, b: str) -> bool:
    """Same normalization ladder locate.py's own locate() uses (exact ->
    curly/straight quotes -> NBSP/whitespace-run collapse -> soft-hyphen
    strip), composed for a single before/after text EQUALITY check rather
    than a haystack search -- after_paragraph_text only ever needs "does
    this one candidate paragraph's text equal the requested text",
    not a position."""
    if a == b:
        return True
    qa, _ = locate._norm_quotes(a)
    qb, _ = locate._norm_quotes(b)
    if qa == qb:
        return True
    wa, _ = locate._norm_whitespace(qa)
    wb, _ = locate._norm_whitespace(qb)
    if wa == wb:
        return True
    sa, _ = locate._norm_softhyphen(wa)
    sb, _ = locate._norm_softhyphen(wb)
    return sa == sb


def _resolve_anchor_index(
    anchor: dict[str, Any] | None,
    body_children: list[Any],
    document_root: Any,
    resolved: Path,
) -> tuple[int, dict[str, Any] | None]:
    """(insert_index, anchor_resolved_evidence) -- insert_index is a body
    index into *body_children* (append-at-end when *anchor* is None,
    matching today's behavior; anchor_resolved is then None too)."""
    if anchor is None:
        return len(body_children), None
    if not isinstance(anchor, dict):
        raise _make_error(ErrorCode.INVALID_INPUT, "anchor must be an object.", {"anchor": anchor})

    after_table_id = anchor.get("after_table_id")
    section_key = anchor.get("section_key")
    position = anchor.get("position")
    after_paragraph_text = anchor.get("after_paragraph_text")

    if after_table_id is not None:
        if section_key is not None or position is not None or after_paragraph_text is not None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                "anchor.after_table_id cannot be combined with section_key/position/after_paragraph_text.",
                {"anchor": anchor},
            )
        tbl = _find_table_element(document_root, after_table_id)
        if tbl not in body_children:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"table_id {after_table_id} is nested inside another table's cell, not a top-level body "
                "table; use anchor.after_paragraph_text to place a table relative to a nested table's "
                "host cell instead.",
                {"after_table_id": after_table_id},
            )
        insert_at = body_children.index(tbl) + 1
        return insert_at, {"body_index": insert_at, "section_key": None, "after_table_id": after_table_id}

    if section_key is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "anchor must set either after_table_id, or section_key (with position or after_paragraph_text).",
            {"anchor": anchor},
        )

    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    start, end = mutations.locate_section_range(body_children, section_key, styles_by_id)

    if after_paragraph_text is not None:
        if position is not None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                "anchor cannot set both position and after_paragraph_text.",
                {"anchor": anchor},
            )
        target = after_paragraph_text.strip()
        candidates = [
            idx
            for idx in range(start, end)
            if projection._ln(body_children[idx]) == "p"
            and _texts_match_ladder(mutations._paragraph_text(body_children[idx]).strip(), target)
        ]
        if not candidates:
            raise _make_error(
                ErrorCode.ZERO_MATCH,
                f"No top-level paragraph in section {section_key!r} matches after_paragraph_text "
                f"{after_paragraph_text!r} (checked the exact/quotes/whitespace/soft-hyphen ladder).",
                {"section_key": section_key, "after_paragraph_text": after_paragraph_text},
            )
        if len(candidates) > 1:
            raise _make_error(
                ErrorCode.MATCH_COUNT_MISMATCH,
                f"{len(candidates)} top-level paragraphs in section {section_key!r} match "
                f"after_paragraph_text {after_paragraph_text!r}; expected exactly 1.",
                {
                    "section_key": section_key,
                    "after_paragraph_text": after_paragraph_text,
                    "match_count": len(candidates),
                },
            )
        insert_at = candidates[0] + 1
    else:
        pos = position if position is not None else "end"
        if pos == "start":
            insert_at = start + 1
        elif pos == "end":
            insert_at = end
        else:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"anchor.position must be \"start\" or \"end\", got {position!r}.",
                {"anchor": anchor},
            )

    return insert_at, {"body_index": insert_at, "section_key": section_key, "after_table_id": None}


def _build_table_element(
    ctx: markdown_to_ooxml.StyleContext,
    style_id: str,
    norm_rows: list[list[dict[str, Any]]],
    col_count: int,
    grid_cols_dxa: list[int] | None,
    explicit_grid: bool,
    header_rows: int,
    cant_split: bool,
) -> tuple[Any, list[list[list[Any]]]]:
    tbl = ET.Element(_w("tbl"))
    tblpr = ET.SubElement(tbl, _w("tblPr"))
    ET.SubElement(tblpr, _w("tblStyle"), {_w("val"): style_id})
    if explicit_grid:
        ET.SubElement(tblpr, _w("tblW"), {_w("w"): str(sum(grid_cols_dxa)), _w("type"): "dxa"})
    else:
        ET.SubElement(tblpr, _w("tblW"), {_w("w"): "0", _w("type"): "auto"})

    grid = ET.SubElement(tbl, _w("tblGrid"))
    for i in range(col_count):
        attrs = {_w("w"): str(grid_cols_dxa[i])} if grid_cols_dxa else {}
        ET.SubElement(grid, _w("gridCol"), attrs)

    per_row_new_elements: list[list[list[Any]]] = []
    for row_index, row in enumerate(norm_rows):
        tr = ET.SubElement(tbl, _w("tr"))
        if cant_split or row_index < header_rows:
            trpr = ET.SubElement(tr, _w("trPr"))
            if cant_split:
                ET.SubElement(trpr, _w("cantSplit"))
            if row_index < header_rows:
                ET.SubElement(trpr, _w("tblHeader"))

        row_elements: list[list[Any]] = []
        for cell in row:
            tc = ET.SubElement(tr, _w("tc"))
            tcpr = ET.SubElement(tc, _w("tcPr"))
            if grid_cols_dxa:
                tcw = sum(grid_cols_dxa[cell["col_start"] : cell["col_start"] + cell["span"]])
                _insert_ordered(tcpr, ET.Element(_w("tcW"), {_w("w"): str(tcw), _w("type"): "dxa"}), _TCPR_CHILD_ORDER)
            if cell["span"] > 1:
                _insert_ordered(tcpr, ET.Element(_w("gridSpan"), {_w("val"): str(cell["span"])}), _TCPR_CHILD_ORDER)
            if cell["v_merge"] == "restart":
                _insert_ordered(tcpr, ET.Element(_w("vMerge"), {_w("val"): "restart"}), _TCPR_CHILD_ORDER)
            elif cell["v_merge"] == "continue":
                _insert_ordered(tcpr, ET.Element(_w("vMerge")), _TCPR_CHILD_ORDER)
            if cell["fill"]:
                _insert_ordered(
                    tcpr,
                    ET.Element(_w("shd"), {_w("val"): "clear", _w("color"): "auto", _w("fill"): cell["fill"]}),
                    _TCPR_CHILD_ORDER,
                )
            if cell["valign"]:
                _insert_ordered(tcpr, ET.Element(_w("vAlign"), {_w("val"): cell["valign"]}), _TCPR_CHILD_ORDER)

            elements = _render_cell_markdown(cell["markdown"], ctx)
            _apply_cell_formatting(elements, cell)
            for el in elements:
                tc.append(el)
            row_elements.append(elements)
        per_row_new_elements.append(row_elements)

    return tbl, per_row_new_elements


def execute_insert_table(
    path: str,
    rows: list[list[Any]],
    style_id: str,
    *,
    header_rows: int = 0,
    grid_dxa: list[int] | None = None,
    cant_split: bool = False,
    anchor: dict[str, Any] | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    if not rows or any(len(r) == 0 for r in rows):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "rows must be a non-empty list of non-empty cell lists.", {"rows": rows}
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

    if not isinstance(header_rows, int) or isinstance(header_rows, bool) or header_rows < 0:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"header_rows must be an integer >= 0, got {header_rows!r}.",
            {"header_rows": header_rows},
        )
    if header_rows >= len(rows):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"header_rows ({header_rows}) must be less than the number of rows ({len(rows)}).",
            {"header_rows": header_rows, "row_count": len(rows)},
        )

    if grid_dxa and any(isinstance(w, bool) or not isinstance(w, int) or w <= 0 for w in grid_dxa):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "grid_dxa must be a list of positive integers (dxa units), one per grid column.",
            {"grid_dxa": grid_dxa},
        )

    norm_rows, col_count = _normalize_rows(rows, grid_dxa if grid_dxa else None)

    if grid_dxa:
        grid_cols_dxa: list[int] | None = list(grid_dxa)
        explicit_grid = True
    else:
        grid_cols_dxa = None
        explicit_grid = False
        page_sections = projection.list_page_sections_impl(resolved)
        if page_sections and page_sections[0].get("column_widths_in"):
            usable_in = page_sections[0]["column_widths_in"][0]
            if usable_in:
                col_width_dxa = max(1, round((usable_in * _DXA_PER_INCH) / col_count))
                grid_cols_dxa = [col_width_dxa] * col_count

    document_root, raw_xml = mutations._load_document(resolved)
    body = mutations._find_body(document_root)
    # Unlike replace_body_markdown/append_markdown, insert_table never
    # needs the trailing sectPr itself -- body.insert(insert_at, tbl)
    # below is safe as long as insert_at <= len(body_children) (always
    # true; every anchor index is derived from positions WITHIN
    # body_children), since sectPr, when present, stays body's real last
    # child either way.
    body_children, _sect_pr = mutations._split_body(body)

    insert_at, anchor_resolved = _resolve_anchor_index(anchor, body_children, document_root, resolved)

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    tbl, per_row_new_elements = _build_table_element(
        ctx, style_id, norm_rows, col_count, grid_cols_dxa, explicit_grid, header_rows, cant_split
    )

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # Nothing existing is removed by an insertion -- only the new
        # table's own runs are wrapped in w:ins (mirrors append_markdown's
        # identical treatment of a pure insertion).
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.wrap_all_runs([tbl], _ins_wrap)

    body.insert(insert_at, tbl)

    # New table's own eventual table_id: re-walk the now-live tree (the
    # table is already spliced in, at whatever body position *insert_at*
    # put it -- not necessarily last) and take the table_start event whose
    # element IS this tbl, by identity. Replaces the old "always last, so
    # count + 1" shortcut, which anchor placement breaks.
    proj = projection.project_document_root(document_root)
    new_table_id: int | None = None
    for event in proj.events:
        if isinstance(event, TableBoundaryEvent) and event.kind == "table_start" and event.element is tbl:
            new_table_id = event.table_id
            break
    if new_table_id is None:  # pragma: no cover - defensive; tbl was just inserted into this same tree
        raise RuntimeError("internal: could not locate the newly inserted table in the re-walked document tree")

    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_rows = [
        [_markdown_for_elements(elements, styles_by_id, new_numbering_index) for elements in row_elements]
        for row_elements in per_row_new_elements
    ]

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = mutations._load_document(written_path)
        new_tbl = _find_table_element(new_document_root, new_table_id)
        new_row_count, new_col_count, _has_merged, _has_nested = _table_stats(new_tbl)
        if new_row_count != len(norm_rows) or new_col_count != col_count:
            raise ValueError(
                f"re-read table has {new_row_count}x{new_col_count}, expected {len(norm_rows)}x{col_count}"
            )
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        written_numbering_index = projection.load_numbering_index(written_path)
        new_trs = [r for r in new_tbl if projection._ln(r) == "tr"]
        for row_index, (tr_elem, row_specs, row_elements, intended_row) in enumerate(
            zip(new_trs, norm_rows, per_row_new_elements, intended_rows)
        ):
            new_cells = _row_cells(tr_elem)
            if len(new_cells) != len(row_specs):
                raise ValueError(f"row {row_index}: re-read {len(new_cells)} cell(s), expected {len(row_specs)}")
            for cell_index, (tc, spec, elements, intended) in enumerate(
                zip(new_cells, row_specs, row_elements, intended_row)
            ):
                actual_span = _cell_grid_span(tc)
                if actual_span != spec["span"]:
                    raise ValueError(
                        f"row {row_index} cell {cell_index}: re-read grid_span {actual_span}, expected {spec['span']}"
                    )
                actual_v_merge = _cell_v_merge(tc)
                expected_v_merge = spec["v_merge"] or "none"
                if actual_v_merge != expected_v_merge:
                    raise ValueError(
                        f"row {row_index} cell {cell_index}: re-read v_merge {actual_v_merge!r}, "
                        f"expected {expected_v_merge!r}"
                    )
                if spec["fill"]:
                    tcpr = _cell_tcpr(tc)
                    shd = _first_child(tcpr, "shd") if tcpr is not None else None
                    actual_fill = projection._attr(shd, "fill") if shd is not None else None
                    if actual_fill is None or actual_fill.upper() != spec["fill"]:
                        raise ValueError(
                            f"row {row_index} cell {cell_index}: re-read shd fill {actual_fill!r}, "
                            f"expected {spec['fill']!r}"
                        )
                children = _cell_content_children(tc)
                actual_children = children[-len(elements):] if elements else []
                actual = _markdown_for_elements(actual_children, new_styles_by_id, written_numbering_index)
                diff = mutations._diff_modulo_whitespace(intended, actual)
                if diff:
                    raise ValueError(f"re-read cell does not match the intended rendering modulo whitespace: {diff}")

        for row_index in range(header_rows):
            trpr = _first_child(new_trs[row_index], "trPr")
            has_header = trpr is not None and _first_child(trpr, "tblHeader") is not None
            if not has_header:
                raise ValueError(f"row {row_index} was requested as a header row but w:tblHeader was not found on re-read")

    conflict_sweep = mutations._write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_text = "\n".join("\t".join(row) for row in intended_rows)
    merged_cells = sum(1 for row in norm_rows for cell in row if cell["span"] > 1 or cell["v_merge"] is not None)
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
    evidence["merged_cells"] = merged_cells
    evidence["anchor_resolved"] = anchor_resolved
    logged, _ = audit.append_audit(path=str(resolved), tool="insert_table", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
