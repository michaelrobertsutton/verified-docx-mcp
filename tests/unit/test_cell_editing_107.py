"""File-mode coverage for issue #107: cell font inheritance, cell paragraph merge,
structural paragraph anchors and per-edit row scope. Every case runs against a real
table with 8.5pt direct formatting (the #106 fixture shape) and asserts that runs the
edit does not touch keep their size. Live-mode (fake pane) coverage is in test_live_107.py;
real Word acceptance is separate (docs/acceptance/107-cell-editing.md)."""

import copy
import shutil
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree as ET

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from verified_docx_mcp import mutations, paragraph_insert, paragraph_merge, paths, projection, tables
from verified_docx_mcp import table_structure as ts
from verified_docx_mcp.errors import ErrorCode, VerifyError

W = ts.w
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FONT = "Arial"
COLOR = "1F3864"


def rewrite(path, root):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts[projection.DEFAULT_PART] = mutations._serialize_xml(
        root, mutations._capture_source_namespaces(parts[projection.DEFAULT_PART])
    )
    with zipfile.ZipFile(path, "w") as z:
        for n, data in parts.items():
            z.writestr(n, data)


def add_hyperlink_relationship(path, rid="rIdLink107", target="https://example.test/proof"):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    rels = ET.fromstring(parts["word/_rels/document.xml.rels"])
    pkg = rels.tag.split("}")[0] + "}"
    ET.SubElement(
        rels,
        pkg + "Relationship",
        {"Id": rid, "Type": REL_NS + "/hyperlink", "Target": target, "TargetMode": "External"},
    )
    parts["word/_rels/document.xml.rels"] = ET.tostring(rels, encoding="utf-8")
    with zipfile.ZipFile(path, "w") as z:
        for n, data in parts.items():
            z.writestr(n, data)


@pytest.fixture
def doc(tmp_path, monkeypatch):
    monkeypatch.setenv(paths._ALLOWED_FILE_ROOTS_ENV, str(tmp_path))
    monkeypatch.setattr(mutations, "_QUIESCE_INTERVAL_SECONDS", 0.001)
    path = tmp_path / "cells.docx"
    shutil.copyfile(Path(__file__).parents[1] / "fixtures/tables.docx", path)
    root, _ = mutations._load_document(path)
    table = tables._find_table_element(root, 1)
    for p in table.iter(W("p")):
        pr = p.find(W("pPr"))
        if pr is None:
            pr = ET.Element(W("pPr"))
            p.insert(0, pr)
        for old in pr.findall(W("pStyle")):
            pr.remove(old)
        ET.SubElement(pr, W("pStyle"), {W("val"): "Compact"})
        ET.SubElement(pr, W("spacing"), {W("after"): "0"})
        for r in p.findall(W("r")):
            rp = r.find(W("rPr"))
            if rp is None:
                rp = ET.Element(W("rPr"))
                r.insert(0, rp)
            ET.SubElement(rp, W("rFonts"), {W("ascii"): FONT, W("hAnsi"): FONT})
            ET.SubElement(rp, W("color"), {W("val"): COLOR})
            ET.SubElement(rp, W("sz"), {W("val"): "17"})
            ET.SubElement(rp, W("szCs"), {W("val"): "17"})
    rewrite(path, root)
    return path


def table_of(path):
    return tables._find_table_element(mutations._load_document(path)[0], 1)


def cell_of(path, row=1, col=1):
    return table_of(path).findall(W("tr"))[row - 1].findall(W("tc"))[col - 1]


def sizes(element):
    return {
        (r.find(f"{W('rPr')}/{W('sz')}").get(W("val")) if r.find(f"{W('rPr')}/{W('sz')}") is not None else None)
        for r in element.iter(W("r"))
    }


def add_second_paragraph(path, text="second", bold=True, link=False, same_ppr=False):
    """A second paragraph in cell (1,1). By default its paragraph style differs from the first's;
    ``same_ppr=True`` keeps the paragraph properties identical (what a live merge requires)."""
    root, _ = mutations._load_document(path)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    first = cell.find(W("p"))
    second = copy.deepcopy(first)
    if not same_ppr:
        ppr = second.find(W("pPr"))
        if ppr is None:
            ppr = ET.Element(W("pPr"))
            second.insert(0, ppr)
        style = ppr.find(W("pStyle"))
        if style is None:
            style = ET.Element(W("pStyle"))
            ppr.insert(0, style)
        style.set(W("val"), "Normal")
    run = second.find(W("r"))
    run.find(W("t")).text = text
    if bold:
        rpr = run.find(W("rPr"))
        if rpr is None:
            rpr = ET.Element(W("rPr"))
            run.insert(0, rpr)
        ET.SubElement(rpr, W("b"))
    if link:
        second.remove(run)
        hyperlink = ET.SubElement(second, W("hyperlink"), {f"{{{REL_NS}}}id": "rIdLink107"})
        hyperlink.append(run)
    cell.append(second)
    rewrite(path, root)
    if link:
        add_hyperlink_relationship(path)
    return first.find(W("pPr"))


# ---- 1. font inheritance ---------------------------------------------------


def test_cell_markdown_inherits_the_first_run_font_without_touching_neighbors(doc):
    before = copy.deepcopy(table_of(doc))
    tables.execute_replace_cell_markdown(str(doc), 1, 1, 1, "**Label:** new text", write_mode="file")
    after = table_of(doc)
    cell = after.findall(W("tr"))[0].findall(W("tc"))[0]
    assert sizes(cell) == {"17"}
    for run in cell.iter(W("r")):
        rpr = run.find(W("rPr"))
        assert rpr.find(W("rFonts")).get(W("ascii")) == FONT
        assert rpr.find(W("color")).get(W("val")) == COLOR
        assert rpr.find(W("szCs")).get(W("val")) == "17"
    assert len(list(cell.iter(W("b")))) == 1  # the bold the markdown asked for is kept
    # Every other row, and every other cell of row 1, is byte-identical (8.5pt untouched).
    assert ET.tostring(after.findall(W("tr"))[1]) == ET.tostring(before.findall(W("tr"))[1])
    assert ET.tostring(after.findall(W("tr"))[0].findall(W("tc"))[1]) == ET.tostring(
        before.findall(W("tr"))[0].findall(W("tc"))[1]
    )


def test_inherit_format_none_keeps_the_style_font(doc):
    tables.execute_replace_cell_markdown(
        str(doc), 1, 1, 1, "plain", write_mode="file", inherit_format="none"
    )
    assert sizes(cell_of(doc)) == {None}


def test_a_run_property_set_by_the_markdown_is_not_overridden():
    font = tables._inherited_cell_font(
        ET.fromstring(
            f'<w:tc xmlns:w="{projection.W_NS}"><w:p><w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t>x</w:t></w:r></w:p></w:tc>'
        )
    )
    run = ET.fromstring(
        f'<w:p xmlns:w="{projection.W_NS}"><w:r><w:rPr><w:sz w:val="30"/></w:rPr><w:t>big</w:t></w:r></w:p>'
    )
    tables._apply_inherited_font([run], font)
    assert run.find(f"{W('r')}/{W('rPr')}/{W('sz')}").get(W("val")) == "30"


@pytest.mark.parametrize(
    "xml, want",
    [
        # a hyperlink first run is looked through
        ('<w:p><w:hyperlink><w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t>a</w:t></w:r></w:hyperlink></w:p>', {"size": 8.5}),
        # an empty first paragraph falls back to the paragraph mark
        ('<w:p><w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr></w:p>', {"size": 9.0}),
        # a font set only by the style supplies nothing to inherit
        ('<w:p><w:r><w:t>a</w:t></w:r></w:p>', None),
        # name and colour are carried; "auto" is not
        (
            '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:color w:val="auto"/></w:rPr><w:t>a</w:t></w:r></w:p>',
            {"name": "Calibri"},
        ),
    ],
)
def test_first_run_definition(xml, want):
    tc = ET.fromstring(f'<w:tc xmlns:w="{projection.W_NS}">{xml}</w:tc>')
    font = tables._inherited_cell_font(tc)
    assert (font["wire"] if font else None) == want


def test_a_link_keeps_its_link_colour_and_paragraph_marks_follow_the_cells_size(doc):
    tables.execute_replace_cell_markdown(
        str(doc), 1, 1, 1, "see [proof](https://x.test)\n\n- a\n- b", write_mode="file"
    )
    cell = cell_of(doc)
    link_runs = [r for link in cell.iter(W("hyperlink")) for r in link.iter(W("r"))]
    assert link_runs
    for run in link_runs:  # a direct colour would beat the Hyperlink style
        assert run.find(f"{W('rPr')}/{W('color')}") is None
        assert run.find(f"{W('rPr')}/{W('sz')}").get(W("val")) == "17"
    marks = [p.find(f"{W('pPr')}/{W('rPr')}") for p in cell.iter(W("p"))]
    assert all(m is not None and m.find(W("sz")).get(W("val")) == "17" for m in marks)  # bullets too


def test_a_text_run_without_a_direct_font_is_not_overridden_by_a_stray_paragraph_mark():
    tc = ET.fromstring(
        f'<w:tc xmlns:w="{projection.W_NS}"><w:p><w:pPr><w:rPr><w:sz w:val="17"/></w:rPr></w:pPr>'
        "<w:r><w:t>style-sized</w:t></w:r></w:p></w:tc>"
    )
    assert tables._inherited_cell_font(tc) is None  # the style supplies the text's size


@pytest.mark.parametrize(
    "reported, drop_colour, want",
    [
        ({"name": "Arial", "size": 8.5, "color": "#1F3864"}, False, {"name": "Arial", "size": 8.5, "color": "1F3864"}),
        ({"name": "Arial", "size": 8.5, "color": "#1F3864"}, True, {"name": "Arial", "size": 8.5}),
        # Word reports Automatic as black: forcing it would pin a colour the cell never had
        ({"name": "Calibri", "size": 9, "color": "#000000"}, False, {"name": "Calibri", "size": 9.0}),
        # a mixed value comes back empty / 0: only what is known is applied
        ({"name": "", "size": 0, "color": ""}, False, None),
        ({"name": None, "size": None, "color": None}, False, None),
        ({"size": True}, False, None),
        (None, False, None),
        ("not a dict", False, None),
    ],
)
def test_the_font_word_reports_becomes_the_font_to_apply(reported, drop_colour, want):
    assert tables._wire_font(reported, drop_colour=drop_colour) == want


def test_live_font_read_asks_for_the_cell_scoped_font_not_ooxml():
    from types import SimpleNamespace as NS

    sent = []

    def request(op, payload=None, **kw):
        sent.append((op, dict(payload or {})))
        return {"text": "Header", "paragraphs": ["Header"], "font": {"name": "Arial", "size": 8.5, "color": "#000000"}}

    session = NS(hello=NS(capabilities={"cell_font_inheritance_v1"}), document_name="x.docx", request_threadsafe=request)
    address = {"table_index": 1, "row_index": 1, "cell_index": 1}
    _, wire = tables.live_cell_read_with_font(session, address, "first_run")
    assert wire == {"name": "Arial", "size": 8.5}
    assert sent[-1][1].get("include_font") is True and "include_ooxml" not in sent[-1][1]
    _, none = tables.live_cell_read_with_font(session, address, "none")
    assert none is None and "include_font" not in sent[-1][1]


def test_merge_never_introduces_a_run_format_the_cell_did_not_have(doc):
    add_second_paragraph(doc, text="second", bold=True)
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    rels = paragraph_merge._load_rels(doc)
    _, _, _, _, expected = paragraph_merge.prepare(root, rels, 1, 1, 1, 1, None, " ")
    fresh, _ = mutations._load_document(doc)
    # the document as a (hypothetical) live import would leave it: the merged paragraph, one run shrunk
    prepared_root, _ = mutations._load_document(doc)
    paragraph_merge.prepare(prepared_root, rels, 1, 1, 1, 1, None, " ")
    bad = tables._find_table_element(prepared_root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    next(bad.iter(W("sz"))).set(W("val"), "21")
    problems = paragraph_merge.check(prepared_root, rels, 1, 1, 1, 1, expected)
    assert any("run formatting changed" in p for p in problems)
    assert fresh is not None and cell is not None


def test_invalid_inherit_format_is_refused(doc):
    with pytest.raises(VerifyError) as exc:
        tables.execute_replace_cell_markdown(str(doc), 1, 1, 1, "x", write_mode="file", inherit_format="all")
    assert exc.value.envelope.error_code == ErrorCode.INVALID_INPUT


# ---- 3. merge --------------------------------------------------------------


def test_merge_keeps_first_paragraph_properties_and_both_run_formats(doc):
    props = ET.tostring(add_second_paragraph(doc))
    other_before = ET.tostring(cell_of(doc, 1, 2))
    paragraph_merge.execute(
        str(doc), 1, 1, 1, 1, expected_text=["R1C1", "second"], separator=" ", write_mode="file"
    )
    cell = cell_of(doc)
    assert len(cell.findall(W("p"))) == 1
    assert ET.tostring(cell.find(f"{W('p')}/{W('pPr')}")) == props
    assert "".join(n.text or "" for n in cell.iter(W("t"))) == "R1C1 second"
    assert len(list(cell.iter(W("b")))) == 1
    assert sizes(cell) == {"17"}  # no run changed size
    assert ET.tostring(cell_of(doc, 1, 2)) == other_before


def test_merge_keeps_a_hyperlink_target_and_text(doc):
    add_second_paragraph(doc, text="proof", bold=False, link=True)
    paragraph_merge.execute(str(doc), 1, 1, 1, 1, separator=" ", write_mode="file")
    cell = cell_of(doc)
    links = list(cell.iter(W("hyperlink")))
    assert len(links) == 1
    assert paragraph_merge._hyperlinks(cell, paragraph_merge._load_rels(doc)) == [
        ("https://example.test/proof", None, "proof")
    ]


def test_merge_of_a_cells_last_paragraph_refuses_without_mutation(doc):
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, write_mode="file")
    assert exc.value.envelope.error_code == ErrorCode.STRUCTURAL_BOUNDARY
    assert doc.read_bytes() == before


def test_merge_refuses_a_text_mismatch_and_protected_structure(doc):
    add_second_paragraph(doc)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, expected_text=["R1C1", "WRONG"], write_mode="file")
    assert exc.value.envelope.error_code == ErrorCode.CELL_TEXT_MISMATCH
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    ET.SubElement(cell.find(W("p")), W("bookmarkStart"), {W("id"): "9", W("name"): "keep"})
    rewrite(doc, root)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, write_mode="file")
    assert exc.value.envelope.error_code == ErrorCode.STRUCTURAL_BOUNDARY
    assert doc.read_bytes() == before


@pytest.mark.parametrize("markup", ["ins", "pPrChange", "del"])
def test_merge_refuses_any_revision_markup_even_the_configured_authors(doc, markup, monkeypatch):
    monkeypatch.setenv("VERIFIED_DOCX_MCP_AUTHOR", "Sutton")
    add_second_paragraph(doc)
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    second = cell.findall(W("p"))[1]
    attrs = {W("id"): "41", W("author"): "Sutton", W("date"): "2026-10-29T00:00:00Z"}
    if markup == "pPrChange":
        ET.SubElement(second.find(W("pPr")), W("pPrChange"), attrs).append(ET.Element(W("pPr")))
    else:
        run = second.find(W("r"))
        second.remove(run)
        ET.SubElement(second, W(markup), attrs).append(run)
    rewrite(doc, root)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, write_mode="file", allow_concurrent_editor=True)
    assert exc.value.envelope.error_code == ErrorCode.TRACKED_CHANGES_PRESENT
    assert doc.read_bytes() == before


def test_merge_refuses_a_pair_separated_by_other_cell_content(doc):
    add_second_paragraph(doc)
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    first, second = cell.findall(W("p"))
    middle = copy.deepcopy(first)
    middle.find(W("r")).find(W("t")).text = "middle"
    sdt = ET.Element(W("sdt"))
    ET.SubElement(sdt, W("sdtContent")).append(middle)
    cell.insert(list(cell).index(second), sdt)  # first / content control / second
    rewrite(doc, root)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_merge.execute(str(doc), 1, 1, 1, 1, write_mode="file")
    assert exc.value.envelope.error_code == ErrorCode.STRUCTURAL_BOUNDARY
    assert "not adjacent" in exc.value.envelope.message
    assert doc.read_bytes() == before


def test_paragraph_text_matches_the_projections_character_handling():
    para = ET.fromstring(
        f'<w:p xmlns:w="{projection.W_NS}"><w:r><w:t>non</w:t><w:noBreakHyphen/><w:t>breaking</w:t>'
        '<w:softHyphen/><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>'
    )
    text = ts.paragraph_text(para)
    assert text == "non-breaking­•"
    assert text.replace("­", "") == projection.project_document_root(
        ET.fromstring(f'<w:document xmlns:w="{projection.W_NS}"><w:body>' + ET.tostring(para, encoding="unicode") + "</w:body></w:document>")
    ).text.replace("­", "").strip()


def test_merge_text_guard_accepts_the_text_a_read_returned(doc):
    add_second_paragraph(doc, text="x", bold=False)
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    cell.find(W("p")).find(W("r")).append(ET.Element(W("noBreakHyphen")))
    rewrite(doc, root)
    paragraph_merge.execute(str(doc), 1, 1, 1, 1, expected_text=["R1C1-", "x"], write_mode="file")
    assert "".join(ts.paragraph_text(p) for p in cell_of(doc).findall(W("p"))) == "R1C1-x"


@pytest.mark.parametrize("child", ["sym", "softHyphen"])
def test_fill_refuses_a_paragraph_that_displays_a_symbol_or_hyphen(doc, child):
    paragraph_insert.execute_insert_paragraphs(
        str(doc), {"after_table_id": 1}, [{"text": "Caption"}], 1, write_mode="file"
    )
    root, _ = mutations._load_document(doc)
    children = list(mutations._find_body(root))
    target = children[children.index(tables._find_table_element(root, 1)) + 1]
    for old in list(target):
        target.remove(old)
    run = ET.SubElement(target, W("r"))
    attrs = {W("font"): "Symbol", W("char"): "F0B7"} if child == "sym" else {}
    ET.SubElement(run, W(child), attrs)
    rewrite(doc, root)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_insert.execute_insert_paragraphs(
            str(doc), {"empty_paragraph_after_table_id": 1}, [{"text": "Nope"}], 1,
            write_mode="file", allow_concurrent_editor=True,
        )
    assert exc.value.envelope.error_code == ErrorCode.STRUCTURAL_BOUNDARY
    assert doc.read_bytes() == before


def test_a_hyperlink_in_a_paragraph_outside_the_merged_pair_does_not_fail_verification(doc):
    add_second_paragraph(doc, text="second", bold=False)
    root, _ = mutations._load_document(doc)
    cell = tables._find_table_element(root, 1).findall(W("tr"))[0].findall(W("tc"))[0]
    third = copy.deepcopy(cell.find(W("p")))
    run = third.find(W("r"))
    run.find(W("t")).text = "outside"
    third.remove(run)
    ET.SubElement(third, W("hyperlink"), {W("anchor"): "elsewhere"}).append(run)
    cell.append(third)
    rewrite(doc, root)
    paragraph_merge.execute(str(doc), 1, 1, 1, 1, separator=" ", write_mode="file")
    cell = cell_of(doc)
    assert [ts.paragraph_text(p) for p in cell.findall(W("p"))] == ["R1C1 second", "outside"]
    assert len(list(cell.iter(W("hyperlink")))) == 1


def test_a_fill_that_keeps_the_paragraphs_style_does_not_fail_the_live_style_check():
    from verified_docx_mcp.live import paragraphs_live

    specs = [{"text": "Filled", "style": None, "color": None}]
    inserted = [{"text": "Filled", "style": "Heading 2", "style_builtin": "", "color": "",
                 "requested_style": {"kind": "builtin", "value": "Normal"}}]
    assert paragraphs_live._spec_problems(specs, inserted)  # an ordinary insert would read Normal
    assert paragraphs_live._spec_problems(specs, inserted, keeps_first_style=True) == []
    # a named style, or a colour, is still checked on the filled paragraph
    named = [{"text": "Filled", "style": "Heading 2", "color": None}]
    assert paragraphs_live._spec_problems(named, [{**inserted[0], "requested_style": {"kind": "name", "value": "Body"}, "style": "Body"}], keeps_first_style=True)
    coloured = [{"text": "Filled", "style": None, "color": "7030A0"}]
    assert paragraphs_live._spec_problems(coloured, inserted, keeps_first_style=True)


def test_tracked_merge_is_refused_in_both_modes(doc):
    add_second_paragraph(doc)
    before = doc.read_bytes()
    for mode in ("file", "auto"):
        with pytest.raises(VerifyError) as exc:
            paragraph_merge.execute(str(doc), 1, 1, 1, 1, track_changes=True, write_mode=mode)
        assert exc.value.envelope.error_code == ErrorCode.INVALID_INPUT
    assert doc.read_bytes() == before


# ---- 4. structural paragraph anchors ---------------------------------------


def body_after_table(path):
    root, _ = mutations._load_document(path)
    body = mutations._find_body(root)
    table = tables._find_table_element(root, 1)
    children = list(body)
    return children, children.index(table)


def test_paragraph_after_table_and_fill_empty(doc):
    children_before, at = body_after_table(doc)
    result = paragraph_insert.execute_insert_paragraphs(
        str(doc), {"after_table_id": 1}, [{"text": "Caption"}], 1, write_mode="file"
    )
    assert result["applied"]
    children, at = body_after_table(doc)
    assert len(children) == len(children_before) + 1
    assert "".join(children[at + 1].itertext()) == "Caption"
    assert "".join(children[at + 2].itertext()) == "".join(children_before[at + 1].itertext())

    # Make the paragraph after the table empty, then fill it: the paragraph is kept, not replaced.
    root, _ = mutations._load_document(doc)
    body = mutations._find_body(root)
    target = list(body)[at + 1]
    for child in list(target):
        if child.tag != W("pPr"):
            target.remove(child)
    ppr = target.find(W("pPr"))
    if ppr is None:
        ppr = ET.Element(W("pPr"))
        target.insert(0, ppr)
    ET.SubElement(ppr, W("jc"), {W("val"): "center"})
    props = ET.tostring(ppr)
    count = len(list(body))
    rewrite(doc, root)
    paragraph_insert.execute_insert_paragraphs(
        str(doc), {"empty_paragraph_after_table_id": 1}, [{"text": "Filled"}, {"text": "Extra"}], 1,
        write_mode="file", allow_concurrent_editor=True,
    )
    children, at = body_after_table(doc)
    assert len(children) == count + 1  # one paragraph filled, one inserted
    assert [("".join(p.itertext())) for p in children[at + 1 : at + 3]] == ["Filled", "Extra"]
    assert ET.tostring(children[at + 1].find(W("pPr"))) == props  # the paragraph's own properties survive


@pytest.mark.parametrize("problem", ["text", "bookmark", "hyperlink", "tab"])
def test_fill_refuses_a_paragraph_that_is_not_empty(doc, problem):
    paragraph_insert.execute_insert_paragraphs(
        str(doc), {"after_table_id": 1}, [{"text": "Caption"}], 1, write_mode="file"
    )
    root, _ = mutations._load_document(doc)
    body = mutations._find_body(root)
    children = list(body)
    target = children[children.index(tables._find_table_element(root, 1)) + 1]
    for child in list(target):
        target.remove(child)
    if problem == "text":
        ET.SubElement(ET.SubElement(target, W("r")), W("t")).text = "kept"
    elif problem == "bookmark":
        ET.SubElement(target, W("bookmarkStart"), {W("id"): "7", W("name"): "b"})
    elif problem == "hyperlink":
        ET.SubElement(target, W("hyperlink"), {W("anchor"): "x"})
    else:
        ET.SubElement(ET.SubElement(target, W("r")), W("tab"))
    rewrite(doc, root)
    before = doc.read_bytes()
    with pytest.raises(VerifyError) as exc:
        paragraph_insert.execute_insert_paragraphs(
            str(doc), {"empty_paragraph_after_table_id": 1}, [{"text": "Nope"}], 1,
            write_mode="file", allow_concurrent_editor=True,
        )
    assert exc.value.envelope.error_code == ErrorCode.STRUCTURAL_BOUNDARY
    assert doc.read_bytes() == before


@pytest.mark.parametrize(
    "anchor, position, message",
    [
        ({"after_table_id": 1, "empty_paragraph_after_table_id": 1}, "after", "exactly one"),
        ({"after_table_id": 0}, "after", "positive"),
        ({"after_table_id": True}, "after", "positive"),
        ({"before_table": 1}, "after", "exactly one"),
        ({"after_table_id": 1}, "before", "after"),
    ],
)
def test_structural_anchor_validation(anchor, position, message):
    with pytest.raises(VerifyError) as exc:
        paragraph_insert.validate_insert_request(anchor, position, [{"text": "x"}], 1)
    assert exc.value.envelope.error_code == ErrorCode.INVALID_INPUT
    assert message in exc.value.envelope.message


def test_missing_table_is_table_not_found(doc):
    with pytest.raises(VerifyError) as exc:
        paragraph_insert.execute_insert_paragraphs(
            str(doc), {"after_table_id": 99}, [{"text": "x"}], 1, write_mode="file"
        )
    assert exc.value.envelope.error_code == ErrorCode.TABLE_NOT_FOUND


# ---- 5. per-edit row scope in a batch ---------------------------------------


def test_batch_forwards_each_row_anchor_and_chains_revisions(monkeypatch, tmp_path):
    from verified_docx_mcp import text_edit
    from verified_docx_mcp.live import write_mode

    pane = SimpleNamespace(
        document_name="batch.docx",
        hello=SimpleNamespace(capabilities={"replacement_formatting", "comment_loss_guard", "row_scope"}),
        request_threadsafe=lambda *args, **kwargs: {"bodySha256": "a" * 64},
    )
    monkeypatch.setattr(write_mode, "live_session_for", lambda path: pane)
    calls = []

    def apply(session, **kwargs):
        calls.append(kwargs)
        return {"post": ("b" if len(calls) == 1 else "c") * 64, "match_count": 1, "matches": []}

    monkeypatch.setattr(text_edit, "_live_replace_once", apply)
    edits = [
        {"find": "same", "replace": "one", "expected_matches": 1, "within_row_containing": "First row"},
        {"find": "same", "replace": "two", "expected_matches": 1, "within_row_containing": "Second row"},
    ]
    result = text_edit.execute_replace_text_batch_live(str(tmp_path / "batch.docx"), edits, "live:sha256:" + "a" * 64)
    assert result["batch_status"] == "complete"
    assert [c["within_row_containing"] for c in calls] == ["First row", "Second row"]
    assert [c["pre_hash"] for c in calls] == ["a" * 64, "b" * 64]


def test_batch_row_scope_needs_the_pane_capability_and_text(monkeypatch, tmp_path):
    from verified_docx_mcp import text_edit
    from verified_docx_mcp.live import write_mode

    pane = SimpleNamespace(
        document_name="batch.docx",
        hello=SimpleNamespace(capabilities={"replacement_formatting"}),  # no row_scope
        request_threadsafe=lambda *args, **kwargs: {"bodySha256": "a" * 64},
    )
    monkeypatch.setattr(write_mode, "live_session_for", lambda path: pane)
    scoped = [{"find": "x", "replace": "y", "expected_matches": 1, "within_row_containing": "Row"}]
    with pytest.raises(VerifyError) as exc:
        text_edit.execute_replace_text_batch_live(str(tmp_path / "batch.docx"), scoped, "live:sha256:" + "a" * 64)
    assert exc.value.envelope.error_code == ErrorCode.LIVE_CAPABILITY_MISSING
    for bad in ("", "  ", 3):
        edits = [{"find": "x", "replace": "y", "expected_matches": 1, "within_row_containing": bad}]
        with pytest.raises(VerifyError) as exc:
            text_edit._validate_batch_edits(edits)
        assert exc.value.envelope.error_code == ErrorCode.INVALID_INPUT
