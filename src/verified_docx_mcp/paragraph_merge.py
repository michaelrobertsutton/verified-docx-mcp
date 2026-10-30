"""Merge two adjacent paragraphs inside one table cell (issue #107).

The first paragraph's properties win; both paragraphs' runs and hyperlinks are
kept as they are (nothing is flattened to text). Both modes verify the result
independently: the cell's paragraph texts and count, the merged paragraph's
properties, every hyperlink's resolved target and text, and every OTHER cell of
the table (text and direct formatting) must read back as intended.

Tracked merges are refused in both modes: how Word records a paragraph-mark
deletion through ``insertOoxml`` is unproven, so the first version does not
offer it.
"""

from __future__ import annotations

import copy
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, markdown_to_ooxml, mutations, paths, tables
from . import table_structure as ts
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .live import reads_live
from .live import write_mode as wm

_CAPABILITY = "cell_paragraph_merge_v1"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
# Structure a merge cannot move without changing meaning: anchored objects, section
# breaks, bookmarks, fields, notes and comment anchors.
_PROTECTED_TAGS = (
    "drawing", "pict", "object", "sectPr", "bookmarkStart", "bookmarkEnd", "fldChar", "fldSimple",
    "instrText", "footnoteReference", "endnoteReference", "commentRangeStart", "commentRangeEnd",
    "commentReference",
)


def _load_rels(docx_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(docx_path) as archive:
        try:
            raw = archive.read("word/_rels/document.xml.rels")
        except KeyError:
            return {}
    return {
        rel.get("Id", ""): rel.get("Target", "")
        for rel in ET.fromstring(raw).findall(f"{{{_PKG_REL_NS}}}Relationship")
    }


def _hyperlinks(element: Any, rels: dict[str, str]) -> list[tuple[str | None, str | None, str]]:
    """(resolved target, anchor, text) per hyperlink in document order."""
    out = []
    for link in element.iter(ts.w("hyperlink")):
        rid = link.get(f"{{{_R_NS}}}id")
        out.append((rels.get(rid) if rid else None, link.get(ts.w("anchor")),
                    "".join(t.text or "" for t in link.iter(ts.w("t")))))
    return out


def _ppr_signature(paragraph: Any) -> tuple:
    props = paragraph.find(ts.w("pPr"))
    return ts.property_signature(props) if props is not None else ()


def _other_cells(table: Any, row_index: int, cell_index: int) -> dict[tuple[int, int], tuple]:
    state = {}
    for r, tr in enumerate(table.findall(ts.w("tr")), start=1):
        for c, tc in enumerate(tables._row_cells(tr), start=1):
            if (r, c) != (row_index, cell_index):
                state[(r, c)] = (
                    tuple(ts.paragraph_text(p) for p in tc.findall(ts.w("p"))),
                    ts.property_signature(tc),
                )
    return state


def _locate(root: Any, table_id: int, row_index: int, cell_index: int) -> tuple[Any, Any]:
    table = ts.simple_table(root, table_id)
    row = tables._find_row(table, row_index, table_id=table_id)
    return table, tables._find_cell(row, cell_index, table_id=table_id, row_index=row_index)


def prepare(root, rels, table_id, row_index, cell_index, paragraph_index, expected_text, separator):
    """Merge in *root* (a parsed document) and return ``(cell, merged paragraph,
    before-XML of the pair, the pair's texts, expectations for verification)``."""
    for name, value in (("table_id", table_id), ("row_index", row_index),
                        ("cell_index", cell_index), ("paragraph_index", paragraph_index)):
        ts.positive(value, name)
    table, cell = _locate(root, table_id, row_index, cell_index)
    paragraphs = cell.findall(ts.w("p"))
    if paragraph_index >= len(paragraphs):
        raise _make_error(ErrorCode.STRUCTURAL_BOUNDARY, "There is no next paragraph in the same cell")
    first, second = paragraphs[paragraph_index - 1 : paragraph_index + 1]
    texts = [ts.paragraph_text(p) for p in (first, second)]
    if expected_text is not None and texts != expected_text:
        raise _make_error(
            ErrorCode.CELL_TEXT_MISMATCH, "Paragraph texts differ from expected_text", {"actual": texts}
        )
    mutations._check_hazards_or_raise(mutations._scan_range_hazards([first, second]), False, resolve_author_name())
    found = sorted({n.tag.rsplit("}", 1)[-1] for p in (first, second) for n in p.iter()} & set(_PROTECTED_TAGS))
    if found:
        raise _make_error(
            ErrorCode.STRUCTURAL_BOUNDARY, f"Paragraphs contain protected structure ({', '.join(found)})"
        )
    before_xml = [ET.tostring(p, encoding="unicode") for p in (first, second)]
    expected = {
        "paragraphs": [
            ts.paragraph_text(p) if p is not first else texts[0] + separator + texts[1]
            for p in paragraphs
            if p is not second
        ],
        "first_ppr": _ppr_signature(first),
        "links": _hyperlinks(first, rels) + _hyperlinks(second, rels),
        "others": _other_cells(table, row_index, cell_index),
    }
    if separator:
        source = next(first.iter(ts.w("r")), None)
        run = ET.SubElement(first, ts.w("r"))
        if source is not None:
            ts.inherit_properties(source, run, "rPr")
        ET.SubElement(run, ts.w("t"), {_XML_SPACE: "preserve"}).text = separator
    for child in list(second):
        if child.tag != ts.w("pPr"):
            first.append(copy.deepcopy(child))
    cell.remove(second)
    return cell, first, before_xml, texts, expected


def check(root, rels, table_id, row_index, cell_index, paragraph_index, expected) -> list[str]:
    """Problems found when comparing *root* with the merge's expectations."""
    problems: list[str] = []
    try:
        table, cell = _locate(root, table_id, row_index, cell_index)
    except Exception as exc:  # noqa: BLE001 - any failure to find the cell is a verification problem
        return [f"cell not found on re-read: {exc}"]
    paragraphs = cell.findall(ts.w("p"))
    texts = [ts.paragraph_text(p) for p in paragraphs]
    if texts != expected["paragraphs"]:
        problems.append(f"cell paragraphs differ: expected {expected['paragraphs']!r}, found {texts!r}")
    elif _ppr_signature(paragraphs[paragraph_index - 1]) != expected["first_ppr"]:
        problems.append("the merged paragraph's properties differ from the first paragraph's")
    if _hyperlinks(cell, rels) != expected["links"]:
        problems.append("hyperlink targets or text changed")
    if _other_cells(table, row_index, cell_index) != expected["others"]:
        problems.append("another cell of the table changed")
    return problems


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
        raise _make_error(ErrorCode.INVALID_INPUT, "expected_text must contain the two paragraph texts")
    if track_changes:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "Tracked paragraph merges are not supported (how Word records them is unproven); "
            "use track_changes=false",
        )
    mode = wm.resolve_write_mode(path, write_mode)
    if mode == "live":
        return _execute_live(
            path, table_id, row_index, cell_index, paragraph_index, expected_text, separator, revision_before
        )
    return _execute_file(
        path, table_id, row_index, cell_index, paragraph_index, expected_text, separator,
        revision_before, allow_concurrent_editor,
    )


def _execute_file(path, table_id, row_index, cell_index, paragraph_index, expected_text, separator,
                  revision_before, allow_concurrent_editor) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre = mutations._guard_before_write(
        resolved, revision_before, allow_concurrent_editor=allow_concurrent_editor, live_capable=True
    )
    root, raw = mutations._load_document(resolved)
    rels = _load_rels(resolved)
    _, _, _, texts, expected = prepare(
        root, rels, table_id, row_index, cell_index, paragraph_index, expected_text, separator
    )

    def verify(written: Path) -> None:
        new_root, _ = mutations._load_document(written)
        problems = check(new_root, _load_rels(written), table_id, row_index, cell_index, paragraph_index, expected)
        if problems:
            raise ValueError("; ".join(problems))

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    sweep = mutations._write_and_verify(resolved, root, raw, ctx, post_verify=verify)
    evidence = tables._evidence(
        applied=True, match_count=1, rung=3, before="\n".join(texts), after=texts[0] + separator + texts[1],
        revision_before=pre["token"], revision_after=sweep["revision_after"],
        audit_logged=False, orphaned_comment_ids=[], conflict_sweep=sweep,
    )
    evidence["audit_logged"], _ = audit.append_audit(
        path=str(resolved), tool="merge_paragraph_with_next", evidence=evidence
    )
    return evidence


def _execute_live(path, table_id, row_index, cell_index, paragraph_index, expected_text, separator,
                  revision_before) -> dict[str, Any]:
    session = wm.live_session_for(path)  # a non-observing view of the shared session
    wm.require_capability(session, _CAPABILITY, feature_description="cell paragraph merge")
    # Reads on both sides of the write are NON-OBSERVING: a failed verification must not
    # have already re-armed this client's write baseline.
    with reads_live.read_source(path, "live", observe=False) as source:
        root, raw = mutations._load_document(source.local_path)
        rels = _load_rels(source.local_path)
        pre = source.live["body_sha256"]
        wm.check_not_stale(revision_before, pre)
        _, merged, before_xml, texts, expected = prepare(
            root, rels, table_id, row_index, cell_index, paragraph_index, expected_text, separator
        )
        body = mutations._find_body(root)
        body.clear()
        body.append(copy.deepcopy(merged))
        ctx = markdown_to_ooxml.StyleContext.build(source.local_path)
        xml = ts.flat_package(source.local_path, root, raw, ctx)
    result = tables._live_cell_request(
        session,
        "cell_paragraph_merge",
        {"table_index": table_id, "row_index": row_index, "cell_index": cell_index,
         "paragraph_index": paragraph_index, "ooxml": xml, "before_xml": before_xml,
         "expectedBodySha256": pre, "track_changes": False},
        expected_body_sha256=pre,
    )
    problems: list[str] = []
    warnings: list[str] = []
    if result.get("applied") is not True:
        problems.append("the pane did not report applied=true")
    post = result.get("post")
    try:
        with reads_live.read_source(path, "live", observe=False) as after:
            new_root, _ = mutations._load_document(after.local_path)
            problems += check(
                new_root, _load_rels(after.local_path), table_id, row_index, cell_index, paragraph_index, expected
            )
            if after.live["body_sha256"] != post:
                warnings.append(
                    "the document changed after the merge (its hash differs from the pane's reply); "
                    "verification compared the merged cell and the rest of its table only"
                )
    except Exception as exc:  # noqa: BLE001 - a failed verification read leaves the outcome unproven
        problems.append(f"verification read failed: {exc}")
    if problems:
        raise tables._fail_live_verification(
            tool="merge_paragraph_with_next", path=path, document_name=session.document_name,
            pre_hash=pre, session=session, problems=problems, what="Paragraph merge failed verification",
        )
    return wm.live_evidence(
        shape_result=result, applied=True, match_count=1, rung=3, before="\n".join(texts),
        after=texts[0] + separator + texts[1], pre_body_sha256=pre, post_body_sha256=post,
        document_name=session.document_name, tool="merge_paragraph_with_next", path=path,
        extra={"warnings": warnings} if warnings else None,
    )
