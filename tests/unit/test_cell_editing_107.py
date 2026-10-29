"""Regression coverage for direct formatting and structural cell edits (#107)."""

import copy
import zipfile
from xml.etree import ElementTree as ET

import pytest
from test_table_structure import doc as doc  # noqa: PLC0414 - pytest fixture re-export
from test_table_structure import root_table

from verified_docx_mcp import mutations, paragraph_insert, paragraph_merge, projection, tables
from verified_docx_mcp import table_structure as ts
from verified_docx_mcp.errors import VerifyError


def rewrite(path, root):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    raw = parts[projection.DEFAULT_PART]
    parts[projection.DEFAULT_PART] = mutations._serialize_xml(
        root, mutations._capture_source_namespaces(raw)
    )
    with zipfile.ZipFile(path, "w") as z:
        for n, data in parts.items():
            z.writestr(n, data)


def test_cell_markdown_inherits_size_without_touching_neighbors(doc):
    before = copy.deepcopy(root_table(doc))
    tables.execute_replace_cell_markdown(
        str(doc), 1, 1, 1, "**Label:** new text", write_mode="file"
    )
    after = root_table(doc)
    cell = after.findall(ts.w("tr"))[0].findall(ts.w("tc"))[0]
    assert {
        r.find(f"{ts.w('rPr')}/{ts.w('sz')}").get(ts.w("val")) for r in cell.iter(ts.w("r"))
    } == {"17"}
    assert ET.tostring(after.findall(ts.w("tr"))[1]) == ET.tostring(before.findall(ts.w("tr"))[1])


def test_merge_keeps_first_paragraph_properties_and_both_run_formats(doc):
    root, _ = mutations._load_document(doc)
    table = tables._find_table_element(root, 1)
    cell = table.findall(ts.w("tr"))[0].findall(ts.w("tc"))[0]
    first = cell.find(ts.w("p"))
    second = copy.deepcopy(first)
    second.find(f"{ts.w('pPr')}/{ts.w('pStyle')}").set(ts.w("val"), "Normal")
    second.find(f"{ts.w('r')}/{ts.w('t')}").text = "second"
    ET.SubElement(second.find(f"{ts.w('r')}/{ts.w('rPr')}"), ts.w("b"))
    cell.append(second)
    rewrite(doc, root)
    props = ET.tostring(first.find(ts.w("pPr")))
    paragraph_merge.execute(
        str(doc), 1, 1, 1, 1, expected_text=["R1C1", "second"], separator=" ", write_mode="file"
    )
    cell = root_table(doc).findall(ts.w("tr"))[0].findall(ts.w("tc"))[0]
    assert len(cell.findall(ts.w("p"))) == 1
    assert ET.tostring(cell.find(f"{ts.w('p')}/{ts.w('pPr')}")) == props
    assert "".join(n.text or "" for n in cell.iter(ts.w("t"))) == "R1C1 second"
    assert len(list(cell.iter(ts.w("b")))) == 1


def test_merge_final_paragraph_refuses_without_mutation(doc):
    before = doc.read_bytes()
    with pytest.raises(VerifyError):
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, write_mode="file")
    assert doc.read_bytes() == before


def test_paragraph_after_table_and_fill_empty(doc):
    result = paragraph_insert.execute_insert_paragraphs(
        str(doc), {"after_table_id": 1}, [{"text": "Caption"}], 1, write_mode="file"
    )
    assert result["applied"]
    root, _ = mutations._load_document(doc)
    body = mutations._find_body(root)
    table = tables._find_table_element(root, 1)
    i = list(body).index(table) + 1
    assert "".join(body[i].itertext()) == "Caption"
    body[i].clear()
    rewrite(doc, root)
    paragraph_insert.execute_insert_paragraphs(
        str(doc),
        {"empty_paragraph_after_table_id": 1},
        [{"text": "Filled"}],
        1,
        write_mode="file",
        allow_concurrent_editor=True,
    )
    root, _ = mutations._load_document(doc)
    body = mutations._find_body(root)
    assert "".join(body[i].itertext()) == "Filled"
    assert sum("Filled" in "".join(p.itertext()) for p in body) == 1


def test_batch_forwards_each_row_anchor_and_chains_revisions(monkeypatch, tmp_path):
    from types import SimpleNamespace

    from verified_docx_mcp import text_edit
    from verified_docx_mcp.live import write_mode

    pane = SimpleNamespace(
        document_name="batch.docx",
        hello=SimpleNamespace(
            capabilities={"replacement_formatting", "comment_loss_guard", "row_scope"}
        ),
        request_threadsafe=lambda *args, **kwargs: {"bodySha256": "a" * 64},
    )
    monkeypatch.setattr(write_mode, "live_session_for", lambda path: pane)
    calls = []

    def apply(session, **kwargs):
        calls.append(kwargs)
        return {"post": ("b" if len(calls) == 1 else "c") * 64, "match_count": 1, "matches": []}

    monkeypatch.setattr(text_edit, "_live_replace_once", apply)
    result = text_edit.execute_replace_text_batch_live(
        str(tmp_path / "batch.docx"),
        [
            {
                "find": "same",
                "replace": "one",
                "expected_matches": 1,
                "within_row_containing": "First row",
            },
            {
                "find": "same",
                "replace": "two",
                "expected_matches": 1,
                "within_row_containing": "Second row",
            },
        ],
        "live:sha256:" + "a" * 64,
    )
    assert result["batch_status"] == "complete"
    assert [c["within_row_containing"] for c in calls] == ["First row", "Second row"]
    assert [c["pre_hash"] for c in calls] == ["a" * 64, "b" * 64]
