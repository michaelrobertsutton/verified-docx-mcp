"""Formatting-preserving table structure changes (#106).

Content is rebuilt from property-only templates: IDs, revision markers, comments,
bookmarks, fields and relationships in source content are never duplicated.
"""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, markdown_to_ooxml, mutations, paths, projection, tables, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, VerifyError, _make_error
from .live import reads_live
from .live import write_mode as wm

w = tables._w


def positive(value: Any, name: str) -> int:
    if type(value) is not int or value < 1:
        raise _make_error(ErrorCode.INVALID_INPUT, f"{name} must be a positive integer")
    return value


def simple_table(root: Any, table_id: int) -> Any:
    positive(table_id, "table_id")
    table = tables._find_table_element(root, table_id)
    if tables._has_merged_cells(table) or tables._has_nested_table(table):
        raise _make_error(
            ErrorCode.MERGED_OR_NESTED_TABLE,
            "Structural edits require an unmerged table without nested tables",
        )
    return table


def inherit_properties(source: Any, dest: Any, name: str) -> None:
    """Overlay destination properties on source properties, preserving unknown tags."""
    original = source.find(w(name))
    overrides = dest.find(w(name))
    if original is None:
        return
    merged = copy.deepcopy(original)
    # Property-change histories belong to the original object, not the copy.
    for child in list(merged):
        if projection._ln(child).endswith("Change") or projection._ln(child) in {"ins", "del"}:
            merged.remove(child)
    if overrides is not None:
        for child in overrides:
            for old in list(merged):
                if old.tag == child.tag:
                    merged.remove(old)
            merged.append(copy.deepcopy(child))
        dest.remove(overrides)
    dest.insert(0, merged)


def inherit_cell(template: Any, cell: Any) -> None:
    inherit_properties(template, cell, "tcPr")
    first_p = template.find(w("p"))
    if first_p is None:
        return
    first_r = next(first_p.iter(w("r")), None)
    for paragraph in cell.iter(w("p")):
        inherit_properties(first_p, paragraph, "pPr")
        if first_r is not None:
            runs = list(paragraph.iter(w("r")))
            for run in runs:
                inherit_properties(first_r, run, "rPr")
            if not any(n.text for n in paragraph.iter(w("t"))):
                # Word stores the font of an empty paragraph on its paragraph mark.
                props = paragraph.find(w("pPr"))
                if props is None:
                    props = ET.Element(w("pPr"))
                    paragraph.insert(0, props)
                inherit_properties(first_r, props, "rPr")
                for run in runs:
                    paragraph.remove(run)


def clone_row(template: Any, cells: list[str]) -> Any:
    source_cells = tables._row_cells(template)
    if (
        not isinstance(cells, list)
        or len(cells) != len(source_cells)
        or any(not isinstance(t, str) for t in cells)
    ):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "cells must contain one string per template cell"
        )
    if any(any(ord(c) < 32 and c not in "\t\n\r" for c in text) for text in cells):
        raise _make_error(ErrorCode.INVALID_INPUT, "cells contain unsupported control characters")
    row = ET.Element(w("tr"))
    inherit_properties(template, row, "trPr")
    for source, text in zip(source_cells, cells, strict=True):
        cell = ET.SubElement(row, w("tc"))
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            p = ET.SubElement(cell, w("p"))
            r = ET.SubElement(p, w("r"))
            ET.SubElement(
                r, w("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}
            ).text = line
        inherit_cell(source, cell)
    return row


def _node_text(node: Any) -> str:
    """One element's contribution to a paragraph's text, with the projection's own character
    handling (``projection.py``): hyphens, soft hyphens and mapped symbols are text (#107)."""
    name = projection._ln(node)
    if name == "t":
        return node.text or ""
    if name == "tab":
        return "\t"
    if name in {"br", "cr"}:
        return "\n"
    if name == "noBreakHyphen":
        return "-"
    if name == "softHyphen":
        return "­"
    if name == "sym":
        mapped = projection._map_sym(projection._attr(node, "font"), projection._attr(node, "char"))
        return mapped if mapped is not None else "�"
    return ""


def paragraph_text(p: Any) -> str:
    return "".join(_node_text(node) for node in p.iter())


def plain_row(row: Any) -> str:
    def paragraph_text(p):
        return "".join(
            (node.text or "")
            if node.tag == w("t")
            else "\t"
            if node.tag == w("tab")
            else "\n"
            if node.tag in {w("br"), w("cr")}
            else ""
            for node in p.iter()
        )

    return "\t".join(
        "\n".join(paragraph_text(p) for p in cell.findall(w("p")))
        for cell in tables._row_cells(row)
    )


def tracked_row(row: Any, track: Any, deletion: bool = False) -> None:
    props = row.find(w("trPr"))
    if props is None:
        props = ET.Element(w("trPr"))
        row.insert(0, props)
    attrs = {w("id"): track.next_id(), w("author"): track.author, w("date"): track.date}
    ET.SubElement(props, w("del" if deletion else "ins"), attrs)
    if deletion:

        def wrap(run):
            return tracked_changes.wrap_deletion(
                run, rid=track.next_id(), author=track.author, date=track.date
            )
    else:

        def wrap(run):
            return tracked_changes.wrap_insertion(
                run, rid=track.next_id(), author=track.author, date=track.date
            )

    tracked_changes.wrap_all_runs([row], wrap)


def execute_row(
    path: str,
    table_id: int,
    *,
    row_index: int | None = None,
    after_row_index: int | None = None,
    before_row_index: int | None = None,
    cells: list[str] | None = None,
    template_row_index: int | None = None,
    expected_text: str | None = None,
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    deleting = row_index is not None
    positive(table_id, "table_id")
    if deleting:
        positive(row_index, "row_index")
        if not isinstance(expected_text, str):
            raise _make_error(ErrorCode.INVALID_INPUT, "expected_text is required")
    else:
        if (after_row_index is None) == (before_row_index is None):
            raise _make_error(
                ErrorCode.INVALID_INPUT, "Supply exactly one of before_row_index or after_row_index"
            )
        positive(after_row_index if after_row_index is not None else before_row_index, "position")
        if template_row_index is not None:
            positive(template_row_index, "template_row_index")
    mode = wm.resolve_write_mode(path, write_mode)
    tool = "delete_table_row" if deleting else "insert_table_row"
    if mode == "live":
        return execute_row_live(
            path,
            table_id,
            tool=tool,
            row_index=row_index,
            after_row_index=after_row_index,
            before_row_index=before_row_index,
            cells=cells,
            template_row_index=template_row_index,
            expected_text=expected_text,
            revision_before=revision_before,
            track_changes=track_changes,
        )
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre = mutations._guard_before_write(
        resolved,
        revision_before,
        allow_concurrent_editor=allow_concurrent_editor,
        live_capable=True,
    )
    root, raw = mutations._load_document(resolved)
    table = simple_table(root, table_id)
    rows = table.findall(w("tr"))
    track = (
        tracked_changes.TrackContext(root, author=resolve_author_name()) if track_changes else None
    )
    if deleting:
        row = tables._find_row(table, row_index, table_id=table_id)
        if len(rows) == 1:
            raise _make_error(
                ErrorCode.INVALID_INPUT, "Cannot delete the final row; use delete_table"
            )
        before = plain_row(row)
        if tables._norm_cell_text(before) != tables._norm_cell_text(expected_text):
            raise _make_error(ErrorCode.CELL_TEXT_MISMATCH, "Row text differs from expected_text")
        if any(projection._ln(n) in {"drawing", "pict", "object"} for n in row.iter()):
            raise _make_error(ErrorCode.ANCHORED_SHAPES, "The row contains an anchored object")
        mutations._check_hazards_or_raise(
            mutations._scan_range_hazards([row]), False, resolve_author_name()
        )
        if track:
            tracked_row(row, track, deletion=True)
        else:
            table.remove(row)
        result_index, after = row_index, ""
    else:
        adjacent = after_row_index if after_row_index is not None else before_row_index
        anchor = tables._find_row(table, adjacent, table_id=table_id)
        template = tables._find_row(table, template_row_index or adjacent, table_id=table_id)
        row = clone_row(template, cells)
        result_index = adjacent + (1 if after_row_index is not None else 0)
        if track:
            tracked_row(row, track)
        table.insert(list(table).index(anchor) + (1 if after_row_index is not None else 0), row)
        before, after = "", plain_row(row)
    expected = ET.tostring(table)

    def verify(written: Path):
        new_root, _ = mutations._load_document(written)
        if ET.tostring(tables._find_table_element(new_root, table_id)) != expected:
            raise ValueError("Table structure or formatting changed during serialization")

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    sweep = mutations._write_and_verify(resolved, root, raw, ctx, post_verify=verify)
    evidence = tables._evidence(
        applied=True,
        match_count=1,
        rung=3,
        before=before,
        after=after,
        revision_before=pre["token"],
        revision_after=sweep["revision_after"],
        audit_logged=False,
        orphaned_comment_ids=[],
        track=track,
        conflict_sweep=sweep,
    )
    evidence.update(
        table_id=table_id,
        row_index=result_index,
        rows_before=len(rows),
        rows_after=len(table.findall(w("tr"))),
    )
    evidence["audit_logged"], _ = audit.append_audit(
        path=str(resolved), tool=tool, evidence=evidence
    )
    return evidence


def property_signature(element: Any) -> tuple:
    """Prefix/order-independent direct formatting, ignoring revision history."""

    def node_key(node):
        return (
            node.tag,
            tuple(sorted(node.attrib.items())),
            node.text or "",
            tuple(
                sorted(
                    node_key(c)
                    for c in node
                    if not projection._ln(c).endswith("Change")
                    and projection._ln(c) not in {"ins", "del"}
                )
            ),
        )

    return tuple(
        node_key(n)
        for n in element.iter()
        if projection._ln(n) in {"tblPr", "trPr", "tcPr", "pPr", "rPr"}
    )


def execute_row_live(
    path: str, table_id: int, *, tool: str, revision_before: str | None, track_changes: bool, **args
) -> dict[str, Any]:
    session = wm.live_session_for(path)
    wm.require_capability(session, "table_rows_v1", feature_description=tool)
    # The live snapshot supplies independent verification of native row inheritance.
    with reads_live.read_source(path, "live") as source:
        root, _ = mutations._load_document(source.local_path)
        table = simple_table(root, table_id)
        rows = table.findall(w("tr"))
        pre = source.live["body_sha256"]
        wm.check_not_stale(revision_before, pre)
        deleting = tool == "delete_table_row"
        wire = {
            "table_index": table_id,
            "delete": deleting,
            "expectedBodySha256": pre,
            "track_changes": track_changes,
            **args,
        }
        expected_rows = [plain_row(row) for row in rows]
        expected_formats = [property_signature(row) for row in rows]
        if deleting:
            row = tables._find_row(table, args["row_index"], table_id=table_id)
            if len(rows) == 1:
                raise _make_error(ErrorCode.INVALID_INPUT, "Cannot delete the final row")
            if tables._norm_cell_text(plain_row(row)) != tables._norm_cell_text(
                args["expected_text"]
            ):
                raise _make_error(
                    ErrorCode.CELL_TEXT_MISMATCH, "Row text differs from expected_text"
                )
            expected_rows.pop(args["row_index"] - 1)
            expected_formats.pop(args["row_index"] - 1)
        else:
            adjacent = args["after_row_index"] or args["before_row_index"]
            tables._find_row(table, adjacent, table_id=table_id)
            template = tables._find_row(
                table, args["template_row_index"] or adjacent, table_id=table_id
            )
            new_row = clone_row(template, args["cells"])
            expected_rows.insert(
                adjacent if args["after_row_index"] else adjacent - 1, plain_row(new_row)
            )
            expected_formats.insert(
                adjacent if args["after_row_index"] else adjacent - 1, property_signature(new_row)
            )
        try:
            result = tables._live_cell_request(
                session, "table_row_edit", wire, expected_body_sha256=pre
            )
        except VerifyError as exc:
            audit.append_audit(
                path=path,
                tool=tool,
                evidence={"applied": None, "write_mode": "live", "error": exc.envelope.to_dict()},
            )
            raise
        actual = tables._live_table_get(session, table_id)
        actual_rows = ["\t".join(c["text"] for c in row) for row in actual["rows"]]
        if not result.get("applied") or [tables._norm_cell_text(t) for t in actual_rows] != [
            tables._norm_cell_text(t) for t in expected_rows
        ]:
            raise tables._fail_live_verification(
                tool=tool,
                path=path,
                document_name=session.document_name,
                pre_hash=pre,
                session=session,
                problems=["Row read-back differs"],
                what="Row operation failed verification",
            )
        with reads_live.read_source(path, "live") as post_source:
            post_root, _ = mutations._load_document(post_source.local_path)
            post_table = simple_table(post_root, table_id)
            observed_formats = [property_signature(row) for row in post_table.findall(w("tr"))]
        if observed_formats != expected_formats:
            raise tables._fail_live_verification(
                tool=tool,
                path=path,
                document_name=session.document_name,
                pre_hash=pre,
                session=session,
                problems=["Direct row/cell/paragraph/run formatting differs"],
                what="Row formatting failed verification",
            )
        evidence = wm.live_evidence(
            shape_result=result,
            applied=True,
            match_count=1,
            rung=3,
            before="",
            after="",
            pre_body_sha256=pre,
            post_body_sha256=result["post"],
            document_name=session.document_name,
            tool=tool,
            path=path,
            extra={
                "table_id": table_id,
                "row_index": result["row_index"],
                "rows_before": len(rows),
                "rows_after": len(actual_rows),
            },
        )
        evidence.update(
            table_id=table_id,
            row_index=result["row_index"],
            rows_before=len(rows),
            rows_after=len(actual_rows),
        )
        return evidence


def apply_table_template(
    source: Any,
    dest: Any,
    *,
    grid_explicit: bool,
    header_rows: int,
    cant_split: bool,
    font_size: int | None,
    norm_rows: list[list[dict[str, Any]]],
) -> None:
    """Copy direct layout and per-row/cell formatting; explicit fields win."""
    if tables._has_merged_cells(source) or tables._has_nested_table(source):
        raise _make_error(
            ErrorCode.MERGED_OR_NESTED_TABLE,
            "like_table_id requires an unmerged table without nesting",
        )
    source_rows = source.findall(w("tr"))
    dest_rows = dest.findall(w("tr"))
    if not source_rows or len(tables._row_cells(source_rows[0])) != len(
        tables._row_cells(dest_rows[0])
    ):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "Source and destination must have the same column count"
        )
    # The generated auto width/style are defaults, not explicit overrides.
    dest_pr = dest.find(w("tblPr"))
    if dest_pr is not None:
        dest.remove(dest_pr)
    inherit_properties(source, dest, "tblPr")
    if grid_explicit:
        props = dest.find(w("tblPr"))
        if props is None:
            props = ET.Element(w("tblPr"))
            dest.insert(0, props)
        width = props.find(w("tblW"))
        if width is None:
            width = ET.SubElement(props, w("tblW"))
        width.set(w("type"), "dxa")
        width.set(w("w"), str(sum(int(n.get(w("w"))) for n in dest.find(w("tblGrid")))))
    if not grid_explicit:
        grid = dest.find(w("tblGrid"))
        original = source.find(w("tblGrid"))
        if grid is not None and original is not None:
            dest.remove(grid)
            dest.insert(1, copy.deepcopy(original))
    for i, row in enumerate(dest_rows):
        # Preserve header formatting where requested; use the first body row
        # for additional data rows rather than copying the header's bold style.
        body_template = next(
            (r for r in source_rows if r.find(f"{w('trPr')}/{w('tblHeader')}") is None),
            source_rows[-1],
        )
        template = source_rows[min(i, len(source_rows) - 1)] if i < header_rows else body_template
        inherit_properties(template, row, "trPr")
        pr = row.find(w("trPr"))
        if pr is not None and i >= header_rows:
            for marker in list(pr.findall(w("tblHeader"))):
                pr.remove(marker)
        for c, (src, dst) in enumerate(
            zip(tables._row_cells(template), tables._row_cells(row), strict=True)
        ):
            # Generated widths are defaults unless grid_dxa was passed.
            tcpr = dst.find(w("tcPr"))
            if tcpr is not None and not grid_explicit:
                for width in list(tcpr.findall(w("tcW"))):
                    tcpr.remove(width)
            inherit_cell(src, dst)
            tables._apply_cell_formatting(
                tables._cell_content_children(dst), norm_rows[i][c], font_size
            )


def flat_package(path: Path, root: Any, raw: bytes, ctx: Any, source_flat_opc: str) -> str:
    """Patch Word's package without reserializing untouched parts/namespaces."""
    from xml.dom import minidom

    pkg = "http://schemas.microsoft.com/office/2006/xmlPackage"
    package = minidom.parseString(source_flat_opc)
    parts = {
        p.getAttributeNS(pkg, "name").lstrip("/"): p
        for p in package.getElementsByTagNameNS(pkg, "part")
    }
    source_doc = parts[projection.DEFAULT_PART].getElementsByTagNameNS(pkg, "xmlData")[0]
    original = next(n for n in source_doc.childNodes if n.nodeType == n.ELEMENT_NODE)
    # Prefixes named only in mc:Ignorable must survive the new document part.
    declarations = mutations._capture_source_namespaces(original.toxml().encode())
    overrides = mutations._build_overrides(path, root, raw, ctx)
    overrides[projection.DEFAULT_PART] = mutations._serialize_xml(root, declarations)
    for name, data in overrides.items():
        if name == "[Content_Types].xml":
            continue  # Flat OPC represents content types on each pkg:part.
        part = parts.get(name)
        if part is None:
            part = package.createElementNS(pkg, "pkg:part")
            part.setAttributeNS(pkg, "pkg:name", "/" + name)
            content_type = (
                "application/vnd.openxmlformats-package.relationships+xml"
                if name.endswith(".rels")
                else "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
            )
            part.setAttributeNS(pkg, "pkg:contentType", content_type)
            package.documentElement.appendChild(part)
        while part.firstChild:
            part.removeChild(part.firstChild)
        xml_data = package.createElementNS(pkg, "pkg:xmlData")
        xml_data.appendChild(package.importNode(minidom.parseString(data).documentElement, True))
        part.appendChild(xml_data)
    return package.toxml()


def execute_copy_live(
    path: str,
    rows: list[list[Any]],
    *,
    like_table_id: int,
    anchor: dict[str, Any] | None,
    revision_before: str | None,
    track_changes: bool,
    header_rows: int,
    grid_dxa: list[int] | None,
    cant_split: bool,
    font_size_pt: float | None,
    style_id: str | None,
    style_builtin: str | None,
    style_from_table_id: int | None,
) -> dict[str, Any]:
    if style_builtin is not None:
        raise _make_error(ErrorCode.INVALID_INPUT, "Use style_id for a faithful OOXML copy")
    session = wm.live_session_for(path)
    wm.require_capability(session, "table_copy_v1", feature_description="faithful table copy")
    if grid_dxa is not None and (
        not isinstance(grid_dxa, list)
        or not grid_dxa
        or any(type(n) is not int or n <= 0 for n in grid_dxa)
    ):
        raise _make_error(ErrorCode.INVALID_INPUT, "grid_dxa must contain positive integer widths")
    norm, cols = tables._normalize_rows(rows, grid_dxa)
    if type(header_rows) is not int or not 0 <= header_rows < len(norm):
        raise _make_error(ErrorCode.INVALID_INPUT, "header_rows must be smaller than the row count")
    wire_anchor, _ = tables._live_anchor(anchor)
    if wire_anchor:
        if "after_table_index" in wire_anchor:
            positive(wire_anchor["after_table_index"], "anchor.after_table_id")
        elif (
            not isinstance(wire_anchor.get("paragraph_text"), str)
            or not wire_anchor["paragraph_text"].strip()
            or wire_anchor.get("position") not in {"before", "after"}
        ):
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                "Anchor requires nonempty paragraph_text and before/after position",
            )
    with reads_live.read_source(path, "live") as source:
        pre = source.live["body_sha256"]
        wm.check_not_stale(revision_before, pre)
        root, raw = mutations._load_document(source.local_path)
        template = simple_table(root, like_table_id)
        ctx = markdown_to_ooxml.StyleContext.build(source.local_path)
        size = tables._font_size_half_points(font_size_pt)
        table, _ = tables._build_table_element(
            ctx, style_id or "", norm, cols, grid_dxa, bool(grid_dxa), header_rows, cant_split, size
        )
        apply_table_template(
            template,
            table,
            grid_explicit=bool(grid_dxa),
            header_rows=header_rows,
            cant_split=cant_split,
            font_size=size,
            norm_rows=norm,
        )
        override_style = style_id
        if style_from_table_id is not None:
            override_style = tables._style_id_of_table(source.local_path, style_from_table_id)
        if override_style is not None:
            if not any(
                s.get("style_id") == override_style and s.get("type") == "table"
                for s in projection.list_styles_impl(source.local_path)
            ):
                raise _make_error(ErrorCode.STYLE_NOT_FOUND, "Requested table style does not exist")
            pr = table.find(w("tblPr"))
            old = pr.find(w("tblStyle"))
            if old is not None:
                pr.remove(old)
            ET.SubElement(pr, w("tblStyle"), {w("val"): override_style})
        body = mutations._find_body(root)
        body.clear()
        body.append(table)
        ooxml = flat_package(source.local_path, root, raw, ctx, source.flat_opc)
        intended = [plain_row(row) for row in table.findall(w("tr"))]
        try:
            result = tables._live_cell_request(
                session,
                "table_copy",
                {
                    "ooxml": ooxml,
                    "anchor": wire_anchor,
                    "expectedBodySha256": pre,
                    "track_changes": track_changes,
                },
                expected_body_sha256=pre,
            )
        except VerifyError as exc:
            audit.append_audit(
                path=path,
                tool="insert_table",
                evidence={"applied": None, "write_mode": "live", "error": exc.envelope.to_dict()},
            )
            raise
        new_id = result.get("table_index")
        if type(new_id) is not int:
            problems = ["Missing inserted table index"]
        else:
            actual = tables._live_table_get(session, new_id)
            actual_rows = ["\t".join(c["text"] for c in row) for row in actual["rows"]]
            problems = (
                []
                if [tables._norm_cell_text(t) for t in intended]
                == [tables._norm_cell_text(t) for t in actual_rows]
                else ["Copied table text differs"]
            )
        if type(new_id) is int:
            with reads_live.read_source(path, "live") as post_source:
                post_root, _ = mutations._load_document(post_source.local_path)
                copied = simple_table(post_root, new_id)
                if property_signature(copied) != property_signature(table):
                    problems.append("Direct table/row/cell/paragraph/run formatting differs")
        if not result.get("applied") or problems:
            raise tables._fail_live_verification(
                tool="insert_table",
                path=path,
                document_name=session.document_name,
                pre_hash=pre,
                session=session,
                problems=problems or ["No application acknowledgment"],
                what="Table copy verification failed",
            )
        evidence = wm.live_evidence(
            shape_result=result,
            applied=True,
            match_count=1,
            rung=3,
            before="",
            after="\n".join(intended),
            pre_body_sha256=pre,
            post_body_sha256=result["post"],
            document_name=session.document_name,
            tool="insert_table",
            path=path,
            extra={"table_id": new_id, "like_table_id": like_table_id},
        )
        evidence.update(table_id=new_id, like_table_id=like_table_id)
        return evidence
