"""Property-level regression tests for #106 (real Word acceptance is separate)."""

import copy
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from verified_docx_mcp import mutations, paths, projection, tables
from verified_docx_mcp import table_structure as ts
from verified_docx_mcp.errors import VerifyError


@pytest.fixture
def doc(tmp_path, monkeypatch):
    monkeypatch.setenv(paths._ALLOWED_FILE_ROOTS_ENV, str(tmp_path))
    monkeypatch.setattr(mutations, "_QUIESCE_INTERVAL_SECONDS", 0.001)
    path = tmp_path / "rows.docx"
    shutil.copyfile(Path(__file__).parents[1] / "fixtures/tables.docx", path)
    root, raw = mutations._load_document(path)
    table = tables._find_table_element(root, 1)
    pr = table.find(ts.w("tblPr"))
    ET.SubElement(pr, ts.w("tblLayout"), {ts.w("type"): "fixed"})
    borders = ET.SubElement(pr, ts.w("tblBorders"))
    for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
        ET.SubElement(
            borders, ts.w(edge), {ts.w("val"): "single", ts.w("sz"): "4", ts.w("color"): "A6A6A6"}
        )
    for row in table.findall(ts.w("tr")):
        rp = row.find(ts.w("trPr"))
        if rp is None:
            rp = ET.Element(ts.w("trPr"))
            row.insert(0, rp)
        ET.SubElement(rp, ts.w("cantSplit"))
        for p in row.iter(ts.w("p")):
            pr = p.find(ts.w("pPr"))
            if pr is None:
                pr = ET.Element(ts.w("pPr"))
                p.insert(0, pr)
            for old in list(pr.findall(ts.w("pStyle"))):
                pr.remove(old)
            ET.SubElement(pr, ts.w("pStyle"), {ts.w("val"): "Compact"})
            ET.SubElement(pr, ts.w("keepNext"))
            ET.SubElement(pr, ts.w("spacing"), {ts.w("after"): "0"})
            for r in p.findall(ts.w("r")):
                rp = r.find(ts.w("rPr"))
                if rp is None:
                    rp = ET.Element(ts.w("rPr"))
                    r.insert(0, rp)
                ET.SubElement(rp, ts.w("sz"), {ts.w("val"): "17"})
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts[projection.DEFAULT_PART] = mutations._serialize_xml(
        root, mutations._capture_source_namespaces(raw)
    )
    with zipfile.ZipFile(path, "w") as z:
        for n, data in parts.items():
            z.writestr(n, data)
    return path


def root_table(path):
    return tables._find_table_element(mutations._load_document(path)[0], 1)


def props(row):
    return [
        ET.tostring(n) for n in row.iter() if projection._ln(n) in {"trPr", "tcPr", "pPr", "rPr"}
    ]


def test_insert_delete_preserves_format_and_neighbors(doc):
    before = copy.deepcopy(root_table(doc))
    rows = before.findall(ts.w("tr"))
    result = ts.execute_row(
        str(doc), 1, after_row_index=1, cells=["New A", "New B"], write_mode="file"
    )
    after = root_table(doc).findall(ts.w("tr"))
    assert result["rows_after"] == 3
    assert props(after[1]) == props(rows[0])
    assert ET.tostring(after[0]) == ET.tostring(rows[0])
    assert ET.tostring(after[2]) == ET.tostring(rows[1])
    ts.execute_row(str(doc), 1, row_index=2, expected_text="New A\tNew B", write_mode="file")
    assert ET.tostring(root_table(doc)) == ET.tostring(before)


@pytest.mark.parametrize(
    "kwargs",
    [
        {"after_row_index": 1, "before_row_index": 1, "cells": ["x", "y"]},
        {"after_row_index": 1, "cells": ["x"]},
        {"after_row_index": True, "cells": ["x", "y"]},
        {"row_index": 1, "expected_text": "wrong"},
        {"row_index": 99, "expected_text": "wrong"},
    ],
)
def test_refusals_leave_package_unchanged(doc, kwargs):
    before = doc.read_bytes()
    with pytest.raises(VerifyError):
        ts.execute_row(str(doc), 1, write_mode="file", **kwargs)
    assert doc.read_bytes() == before


def test_copy_inherits_direct_properties(doc):
    before = copy.deepcopy(root_table(doc))
    result = tables.execute_insert_table(
        str(doc), [["X", "Y"], ["Z", "W"]], like_table_id=1, write_mode="file"
    )
    root, _ = mutations._load_document(doc)
    after = tables._find_table_element(root, result["table_id"])
    assert ET.tostring(after.find(ts.w("tblPr"))) == ET.tostring(before.find(ts.w("tblPr")))
    assert props(after.findall(ts.w("tr"))[0]) == props(before.findall(ts.w("tr"))[0])


def test_tracked_row_has_structural_marker(doc):
    ts.execute_row(
        str(doc), 1, after_row_index=1, cells=["x", "y"], track_changes=True, write_mode="file"
    )
    row = root_table(doc).findall(ts.w("tr"))[1]
    assert row.find(f"{ts.w('trPr')}/{ts.w('ins')}") is not None
