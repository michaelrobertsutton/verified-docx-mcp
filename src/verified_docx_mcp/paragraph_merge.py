"""Merge adjacent paragraphs inside one cell without flattening inline formatting."""

from __future__ import annotations

import copy
from typing import Any

from . import audit, markdown_to_ooxml, mutations, paths, tables
from . import table_structure as ts
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .live import reads_live
from .live import write_mode as wm


def prepare(root, table_id, row_index, cell_index, paragraph_index, expected_text, separator):
    for name, value in (
        ("table_id", table_id),
        ("row_index", row_index),
        ("cell_index", cell_index),
        ("paragraph_index", paragraph_index),
    ):
        ts.positive(value, name)
    table = ts.simple_table(root, table_id)
    row = tables._find_row(table, row_index, table_id=table_id)
    cell = tables._find_cell(row, cell_index, table_id=table_id, row_index=row_index)
    paragraphs = cell.findall(ts.w("p"))
    if paragraph_index >= len(paragraphs):
        raise _make_error(
            ErrorCode.STRUCTURAL_BOUNDARY, "There is no next paragraph in the same cell"
        )
    first, second = paragraphs[paragraph_index - 1 : paragraph_index + 1]
    from xml.etree import ElementTree as ET

    before_xml = [ET.tostring(p, encoding="unicode") for p in (first, second)]
    texts = [ts.plain_row(_cell_row(p)) for p in (first, second)]
    if expected_text is not None and texts != expected_text:
        raise _make_error(ErrorCode.CELL_TEXT_MISMATCH, "Paragraph texts differ from expected_text")
    mutations._check_hazards_or_raise(
        mutations._scan_range_hazards([first, second]), False, resolve_author_name()
    )
    if any(
        n.tag == ts.w(tag)
        for p in (first, second)
        for n in p.iter()
        for tag in ("drawing", "pict", "object", "sectPr")
    ):
        raise _make_error(ErrorCode.STRUCTURAL_BOUNDARY, "Paragraphs contain protected structure")
    if separator:
        r = ET.SubElement(first, ts.w("r"))
        source = next(first.iter(ts.w("r")), None)
        if source is not r and source is not None:
            ts.inherit_properties(source, r, "rPr")
        ET.SubElement(
            r, ts.w("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}
        ).text = separator
    for child in list(second):
        if child.tag != ts.w("pPr"):
            first.append(copy.deepcopy(child))
    cell.remove(second)
    return cell, first, before_xml, texts


def _cell_row(p):
    from xml.etree import ElementTree as ET

    row = ET.Element(ts.w("tr"))
    cell = ET.SubElement(row, ts.w("tc"))
    cell.append(p)
    return row


def execute(
    path: str,
    table_id: int,
    row_index: int,
    cell_index: int,
    paragraph_index: int,
    expected_text: list[str] | None = None,
    separator: str = "",
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    if not isinstance(separator, str) or any(ord(c) < 32 for c in separator):
        raise _make_error(ErrorCode.INVALID_INPUT, "separator must be inline text")
    if expected_text is not None and (
        not isinstance(expected_text, list)
        or len(expected_text) != 2
        or any(not isinstance(t, str) for t in expected_text)
    ):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "expected_text must contain the two paragraph texts"
        )
    mode = wm.resolve_write_mode(path, write_mode)
    if track_changes and mode == "file":
        raise _make_error(ErrorCode.INVALID_INPUT, "Tracked paragraph merges require live mode")
    if mode == "live":
        session = wm.live_session_for(path)
        wm.require_capability(
            session, "cell_paragraph_merge_v1", feature_description="cell paragraph merge"
        )
        with reads_live.read_source(path, "live") as source:
            root, raw = mutations._load_document(source.local_path)
            pre = source.live["body_sha256"]
            wm.check_not_stale(revision_before, pre)
            cell, merged, before_xml, texts = prepare(
                root, table_id, row_index, cell_index, paragraph_index, expected_text, separator
            )
            expected_props = ts.property_signature(cell)
            expected_cell_text = "\n".join(
                ts.plain_row(_cell_row(p)) for p in cell.findall(ts.w("p"))
            )
            expected_count = len(cell.findall(ts.w("p")))
            body = mutations._find_body(root)
            body.clear()
            body.append(copy.deepcopy(merged))
            ctx = markdown_to_ooxml.StyleContext.build(source.local_path)
            xml = ts.flat_package(source.local_path, root, raw, ctx)
            result = tables._live_cell_request(
                session,
                "cell_paragraph_merge",
                {
                    "table_index": table_id,
                    "row_index": row_index,
                    "cell_index": cell_index,
                    "paragraph_index": paragraph_index,
                    "ooxml": xml,
                    "before_xml": before_xml,
                    "expectedBodySha256": pre,
                    "track_changes": track_changes,
                },
                expected_body_sha256=pre,
            )
        with reads_live.read_source(path, "live") as source:
            new_root, _ = mutations._load_document(source.local_path)
            table = ts.simple_table(new_root, table_id)
            row = tables._find_row(table, row_index, table_id=table_id)
            cell = tables._find_cell(row, cell_index, table_id=table_id, row_index=row_index)
            actual_cell_text = "\n".join(
                ts.plain_row(_cell_row(p)) for p in cell.findall(ts.w("p"))
            )
            if (
                actual_cell_text != expected_cell_text
                or len(cell.findall(ts.w("p"))) != expected_count
                or ts.property_signature(cell) != expected_props
            ):
                raise tables._fail_live_verification(
                    tool="merge_paragraph_with_next",
                    path=path,
                    document_name=session.document_name,
                    pre_hash=pre,
                    session=session,
                    problems=["Merged paragraph count or formatting differs"],
                    what="Paragraph merge failed verification",
                )
        return wm.live_evidence(
            shape_result=result,
            applied=True,
            match_count=1,
            rung=3,
            before="\n".join(texts),
            after=texts[0] + separator + texts[1],
            pre_body_sha256=pre,
            post_body_sha256=result["post"],
            document_name=session.document_name,
            tool="merge_paragraph_with_next",
            path=path,
        )
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre = mutations._guard_before_write(
        resolved,
        revision_before,
        allow_concurrent_editor=allow_concurrent_editor,
        live_capable=True,
    )
    root, raw = mutations._load_document(resolved)
    cell, _, _, texts = prepare(
        root, table_id, row_index, cell_index, paragraph_index, expected_text, separator
    )
    from xml.etree import ElementTree as ET

    expected = ET.tostring(root)

    def verify(written):
        if ET.tostring(mutations._load_document(written)[0]) != expected:
            raise ValueError("Paragraph merge changed document structure or formatting")

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    sweep = mutations._write_and_verify(resolved, root, raw, ctx, post_verify=verify)
    evidence = tables._evidence(
        applied=True,
        match_count=1,
        rung=3,
        before="\n".join(texts),
        after=texts[0] + separator + texts[1],
        revision_before=pre["token"],
        revision_after=sweep["revision_after"],
        audit_logged=False,
        orphaned_comment_ids=[],
        conflict_sweep=sweep,
    )
    evidence["audit_logged"], _ = audit.append_audit(
        path=str(resolved), tool="merge_paragraph_with_next", evidence=evidence
    )
    return evidence
