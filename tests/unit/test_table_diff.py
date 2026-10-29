"""Issue #109: diff_table_vs_markdown reports the cells that differ between a
Word table and a markdown pipe table, read-only.

Docx fixtures are built in-test: frag.docx's package with its body swapped for
generated paragraphs/tables, so each case controls the exact cell XML. The live
cases use the FakePane pattern from test_live_reads.py.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import docx_to_flat_opc
from test_live_reads import LiveReadsTestCase

from verified_docx_mcp import paths, server, table_diff
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import reads_live

FIXTURES = REPO / "tests" / "fixtures"
_BODY_RE = re.compile(r"(<w:body>).*?(<w:sectPr)", re.DOTALL)


def _text_run(text: str, rpr: str = "") -> str:
    props = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def para(text: str = "", style: str | None = None) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{ppr}{_text_run(text) if text else ''}</w:p>"


def heading(text: str, level: int = 1) -> str:
    return para(text, f"Heading{level}")


def cell(content: str, span: int = 1) -> str:
    """*content* is plain text, or raw `<w:p...` XML used as-is."""
    body = content if content.startswith("<w:p") else para(content)
    tcpr = f'<w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr>' if span > 1 else ""
    return f"<w:tc>{tcpr}{body}</w:tc>"


def table(rows: list[list[str]], *, grid_cols: int | None = None) -> str:
    """Rows of cell XML (use `cell()`); plain strings are wrapped."""
    cols = grid_cols or max(len(r) for r in rows)
    grid = "<w:tblGrid>" + "<w:gridCol/>" * cols + "</w:tblGrid>"
    trs = "".join("<w:tr>" + "".join(c if c.startswith("<w:tc>") else cell(c) for c in r) + "</w:tr>" for r in rows)
    return f"<w:tbl><w:tblPr/>{grid}{trs}</w:tbl>"


def build_docx(path: Path, body_xml: str) -> Path:
    with zipfile.ZipFile(FIXTURES / "frag.docx") as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                text = data.decode("utf-8")
                text = _BODY_RE.sub(lambda m: m.group(1) + body_xml + m.group(2), text, count=1)
                data = text.encode("utf-8")
            dst.writestr(info, data)
    return path


def md_table(rows: list[list[str]]) -> str:
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * len(rows[0])]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines) + "\n"


def _code(exc: VerifyError) -> ErrorCode:
    return exc.envelope.error_code


BASE = [
    ["Name", "Role", "Site"],
    ["Ann", "Lead", "Austin"],
    ["Bob", "Dev", "Boston"],
    ["Cy", "QA", "Denver"],
]


class TableDiffCase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def docx(self, body: str, name: str = "t.docx") -> Path:
        return build_docx(self.dir / name, body)

    def run_diff(self, docx: Path, **kwargs):
        kwargs.setdefault("source", "file")
        return table_diff.diff_table_vs_markdown_impl(str(docx), **kwargs)

    def refused(self, docx: Path, code: ErrorCode = ErrorCode.INVALID_INPUT, **kwargs) -> VerifyError:
        with self.assertRaises(VerifyError) as ctx:
            self.run_diff(docx, **kwargs)
        self.assertEqual(_code(ctx.exception), code)
        return ctx.exception


class AcceptanceTests(TableDiffCase):
    def test_exactly_the_three_differing_cells_are_reported(self) -> None:
        docx = self.docx(table(BASE))
        md_rows = [list(r) for r in BASE]
        md_rows[1][2] = "Dallas"
        md_rows[2][0] = "Rob"
        md_rows[3][1] = "SRE"
        got = self.run_diff(docx, table_id=1, markdown=md_table(md_rows))
        self.assertFalse(got["identical"])
        self.assertEqual(
            got["differences"],
            [
                {"row": 2, "col": 3, "docx_text": "Austin", "md_text": "Dallas"},
                {"row": 3, "col": 1, "docx_text": "Bob", "md_text": "Rob"},
                {"row": 4, "col": 2, "docx_text": "QA", "md_text": "SRE"},
            ],
        )
        self.assertNotIn("dimensions", got)
        self.assertEqual(got["source"], "file")
        self.assertEqual(got["warnings"], [])

    def test_identical_tables(self) -> None:
        got = self.run_diff(self.docx(table(BASE)), table_id=1, markdown=md_table(BASE))
        self.assertTrue(got["identical"])
        self.assertEqual(got["differences"], [])
        self.assertNotIn("unsupported", got)

    def test_via_the_tool_wrapper(self) -> None:
        docx = self.docx(table(BASE))
        got = server.diff_table_vs_markdown(str(docx), table_id=1, markdown=md_table(BASE), source="file")
        self.assertTrue(got["identical"])


class NormalizationTests(TableDiffCase):
    def test_quotes_dashes_nbsp_and_whitespace_fold(self) -> None:
        docx = self.docx(table([["H"], ["“quoted” – it’s   fine"]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], ['"quoted" - it\'s fine']]))
        self.assertTrue(got["identical"], got)

    def test_emphasis_is_stripped_by_default_and_compared_on_request(self) -> None:
        rich = (
            "<w:p>" + _text_run("plain ") + _text_run("bold", "<w:b/>") + _text_run(" and ")
            + _text_run("both", "<w:b/><w:i/>") + "</w:p>"
        )
        docx = self.docx(table([["H"], [cell(rich)]]))
        md = md_table([["H"], ["plain **bold** and ***both***"]])
        self.assertTrue(self.run_diff(docx, table_id=1, markdown=md)["identical"])
        self.assertTrue(self.run_diff(docx, table_id=1, markdown=md, strip_emphasis=False)["identical"])
        # Same words, but the markdown drops the bold: only visible with markers on.
        md_plain = md_table([["H"], ["plain bold and ***both***"]])
        self.assertTrue(self.run_diff(docx, table_id=1, markdown=md_plain)["identical"])
        got = self.run_diff(docx, table_id=1, markdown=md_plain, strip_emphasis=False)
        self.assertEqual([(d["row"], d["col"]) for d in got["differences"]], [(2, 1)])

    def test_bold_split_across_runs_and_whitespace_edges(self) -> None:
        rich = "<w:p>" + _text_run("a ", "<w:b/>") + _text_run("b", "<w:b/>") + _text_run(" c") + "</w:p>"
        docx = self.docx(table([["H"], [cell(rich)]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], ["**a b** c"]]), strip_emphasis=False)
        self.assertTrue(got["identical"], got)

    def test_escaped_pipes_links_and_inline_code(self) -> None:
        docx = self.docx(table([["A | B", "C"], ["see docs", "x|y and z"]]))
        md = "| A \\| B | C |\n|---|---|\n| [see](http://e.x) docs | `x\\|y` and z |\n"
        got = self.run_diff(docx, table_id=1, markdown=md)
        self.assertTrue(got["identical"], got)

    def test_hyperlink_runs_and_tabs_and_breaks_in_word(self) -> None:
        link = "<w:p>" + _text_run("go ") + "<w:hyperlink>" + _text_run("here") + "</w:hyperlink>" + "</w:p>"
        multi = "<w:p><w:r><w:t>one</w:t><w:br/><w:t>two</w:t><w:tab/><w:t>three</w:t></w:r></w:p>" + para("four")
        docx = self.docx(table([["H1", "H2"], [cell(link), cell(multi)]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H1", "H2"], ["go [here](u)", "one two three four"]]))
        self.assertTrue(got["identical"], got)

    def test_markdown_br_is_a_space(self) -> None:
        docx = self.docx(table([["H"], ["a b"]]))
        self.assertTrue(self.run_diff(docx, table_id=1, markdown="| H |\n|---|\n| a<br>b |\n")["identical"])

    def test_literal_asterisks_in_word_differ_from_markdown_emphasis(self) -> None:
        docx = self.docx(table([["H"], ["*x*"]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], ["*x*"]]))
        self.assertEqual(got["differences"], [{"row": 2, "col": 1, "docx_text": "*x*", "md_text": "x"}])
        # An escaped literal on the markdown side does match.
        self.assertTrue(self.run_diff(docx, table_id=1, markdown=md_table([["H"], ["\\*x\\*"]]))["identical"])


class MarkdownInputTests(TableDiffCase):
    def test_file_input_matches_inline_input(self) -> None:
        docx = self.docx(table(BASE))
        md_path = self.dir / "t.md"
        md_path.write_text(md_table(BASE), encoding="utf-8")
        got = self.run_diff(docx, table_id=1, markdown_path=str(md_path))
        self.assertTrue(got["identical"])

    def test_missing_file_is_invalid_input(self) -> None:
        self.refused(self.docx(table(BASE)), table_id=1, markdown_path=str(self.dir / "nope.md"))

    def test_outside_allowlist_is_invalid_input(self) -> None:
        with tempfile.TemporaryDirectory() as outside:
            md = Path(outside) / "o.md"
            md.write_text(md_table(BASE), encoding="utf-8")
            self.refused(self.docx(table(BASE)), table_id=1, markdown_path=str(md))

    def test_denylisted_path_is_invalid_input(self) -> None:
        md = self.dir / "secret.md"
        md.write_text(md_table(BASE), encoding="utf-8")
        with mock.patch.object(paths, "_is_denylisted_sensitive_path", return_value=True):
            self.refused(self.docx(table(BASE)), table_id=1, markdown_path=str(md))

    def test_directory_is_invalid_input(self) -> None:
        d = self.dir / "adir"
        d.mkdir()
        err = self.refused(self.docx(table(BASE)), table_id=1, markdown_path=str(d))
        self.assertIn("not a regular file", err.envelope.message)


class SelectorValidationTests(TableDiffCase):
    def setUp(self) -> None:
        super().setUp()
        self.target = self.docx(heading("Team") + table(BASE) + heading("Other") + table(BASE))
        self.md = md_table(BASE)

    def test_both_or_neither_markdown_input(self) -> None:
        md_path = self.dir / "x.md"
        md_path.write_text(self.md, encoding="utf-8")
        self.refused(self.target, table_id=1, markdown=self.md, markdown_path=str(md_path))
        self.refused(self.target, table_id=1)

    def test_both_or_neither_word_selector(self) -> None:
        self.refused(self.target, table_id=1, section_key="team-1", markdown=self.md)
        self.refused(self.target, markdown=self.md)

    def test_header_row_with_section_key_is_refused(self) -> None:
        self.refused(self.target, section_key="team-1", markdown=f"# Team\n\n{self.md}", md_header_row="| Name |")

    def test_zero_tables_and_several_tables(self) -> None:
        self.refused(self.target, table_id=1, markdown="no tables here")
        self.refused(self.target, table_id=1, markdown=self.md + "\n" + self.md)

    def test_md_heading_chooses_and_unknown_heading_refuses(self) -> None:
        md = f"# A\n\n{md_table([['X'], ['1']])}\n# B\n\n{self.md}"
        got = self.run_diff(self.target, table_id=1, markdown=md, md_heading="B")
        self.assertTrue(got["identical"])
        self.refused(self.target, table_id=1, markdown=md, md_heading="Missing")

    def test_heading_with_no_table_is_refused(self) -> None:
        self.refused(self.target, table_id=1, markdown=f"# Empty\n\ntext\n\n# T\n\n{self.md}", md_heading="Empty")

    def test_repeated_md_heading_is_refused(self) -> None:
        md = f"# Team\n\n{self.md}\n# Team\n\n{self.md}"
        self.refused(self.target, table_id=1, markdown=md, md_heading="Team")

    def test_md_header_row_chooses_and_ambiguity_refuses(self) -> None:
        other = md_table([["Other"], ["1"]])
        got = self.run_diff(self.target, table_id=1, markdown=other + "\n" + self.md, md_header_row="| Name | Role | Site |")
        self.assertEqual(got["md_table_index"], 1)
        self.assertTrue(got["identical"])

    def test_md_header_row_unknown_and_repeated_refuse(self) -> None:
        self.refused(self.target, table_id=1, markdown=self.md, md_header_row="| Nope |")
        self.refused(self.target, table_id=1, markdown=self.md + "\n" + self.md, md_header_row="| Name | Role | Site |")

    def test_conflicting_heading_and_header_row_is_refused(self) -> None:
        md = f"# A\n\n{md_table([['X'], ['1']])}\n# B\n\n{self.md}"
        self.refused(self.target, table_id=1, markdown=md, md_heading="A", md_header_row="| Name | Role | Site |")
        got = self.run_diff(
            self.target, table_id=1, markdown=md, md_heading="B", md_header_row="| Name | Role | Site |"
        )
        self.assertEqual(got["md_table_index"], 1)

    def test_fenced_hash_lines_do_not_split_a_section(self) -> None:
        md = f"# Real\n\n```\n# not a heading\n```\n\n{self.md}"
        got = self.run_diff(self.target, table_id=1, markdown=md, md_heading="Real")
        self.assertTrue(got["identical"])
        # ...and the fenced line is not a heading to select either.
        self.refused(self.target, table_id=1, markdown=md, md_heading="not a heading")

    def test_unknown_table_and_section(self) -> None:
        self.refused(self.target, ErrorCode.TABLE_NOT_FOUND, table_id=9, markdown=self.md)
        self.refused(self.target, ErrorCode.SECTION_NOT_FOUND, section_key="nope-1", markdown=self.md)

    def test_bad_source(self) -> None:
        self.refused(self.target, table_id=1, markdown=self.md, source="cloud")


class UnsupportedTests(TableDiffCase):
    def test_field_cell_is_unsupported_never_equal(self) -> None:
        fld = (
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>"
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + _text_run("3")
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"
        )
        docx = self.docx(table([["H"], [cell(fld)]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], ["3"]]))
        self.assertFalse(got["identical"])
        self.assertEqual(got["unsupported"][0]["table_id"], 1)
        self.assertIn("fldChar", got["unsupported"][0]["reason"])
        self.assertEqual(got["differences"], [])

    def test_textbox_drawing_cells_are_unsupported(self) -> None:
        for tag in ("w:drawing", "w:pict", "w:txbxContent", "w:fldSimple"):
            with self.subTest(tag=tag):
                bad = f"<w:p><w:r><{tag}/></w:r></w:p>"
                docx = self.docx(table([["H"], [cell(bad)]]), f"{tag.split(':')[1]}.docx")
                got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], [""]]))
                self.assertFalse(got["identical"])
                self.assertTrue(got["unsupported"])

    def test_merged_cell_table_is_unsupported(self) -> None:
        rows = [[cell("Merged", span=2)], [cell("a"), cell("b")]]
        docx = self.docx(table(rows, grid_cols=2))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["Merged", ""], ["a", "b"]]))
        self.assertFalse(got["identical"])
        self.assertIn("merged", got["unsupported"][0]["reason"])

    def test_nested_table_is_unsupported(self) -> None:
        inner = table([["i"]])
        docx = self.docx(table([["H"], [cell(para("x") + inner)]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table([["H"], ["x i"]]))
        self.assertFalse(got["identical"])
        self.assertIn("nested", got["unsupported"][0]["reason"])


class DimensionTests(TableDiffCase):
    def test_extra_docx_row_and_missing_md_column(self) -> None:
        docx = self.docx(table(BASE))
        md_rows = [r[:2] for r in BASE[:3]]
        got = self.run_diff(docx, table_id=1, markdown=md_table(md_rows))
        self.assertFalse(got["identical"])
        self.assertEqual(
            got["dimensions"], {"docx_rows": 4, "docx_cols": 3, "md_rows": 3, "md_cols": 2}
        )
        cells = {(d["row"], d["col"]): d for d in got["differences"]}
        self.assertEqual(cells[(1, 3)], {"row": 1, "col": 3, "docx_text": "Site", "md_text": None})
        self.assertEqual(cells[(4, 1)], {"row": 4, "col": 1, "docx_text": "Cy", "md_text": None})
        self.assertEqual(len(cells), 6)  # col 3 of rows 1-3, and all of row 4
        self.assertTrue(all(d["md_text"] is None for d in got["differences"]))

    def test_extra_markdown_row_and_column(self) -> None:
        docx = self.docx(table([r[:2] for r in BASE[:2]]))
        got = self.run_diff(docx, table_id=1, markdown=md_table(BASE))
        self.assertEqual(got["dimensions"], {"docx_rows": 2, "docx_cols": 2, "md_rows": 4, "md_cols": 3})
        self.assertTrue(all(d["docx_text"] is None for d in got["differences"]))
        self.assertEqual(len(got["differences"]), 8)  # col 3 of rows 1-2, and all of rows 3-4

    def test_empty_cell_versus_missing_cell_is_a_difference(self) -> None:
        docx = self.docx(table([["A", "B"], ["x", ""]]))
        got = self.run_diff(docx, table_id=1, markdown="| A |\n|---|\n| x |\n")
        self.assertFalse(got["identical"])
        self.assertEqual(got["differences"][0]["md_text"], None)


class MultiTableTests(TableDiffCase):
    T1 = [["Name", "Role"], ["Ann", "Lead"]]
    T2 = [["Item", "Qty"], ["Pen", "2"], ["Ink", "9"]]

    def body(self) -> str:
        return heading("Intro") + para("x") + heading("Team") + table(self.T1) + para("gap") + table(self.T2) + heading("End")

    def test_pairs_by_header_array_and_reports_per_table(self) -> None:
        docx = self.docx(self.body())
        md = f"# Team\n\n{md_table(self.T2[:2] + [['Ink', '8']])}\n{md_table(self.T1)}"
        got = self.run_diff(docx, section_key="team-1", markdown=md)
        self.assertFalse(got["identical"])
        by_id = {t["table_id"]: t for t in got["tables"]}
        self.assertEqual(by_id[1]["differences"], [])
        self.assertEqual(by_id[1]["md_table_index"], 1)
        self.assertEqual(by_id[2]["differences"], [{"row": 3, "col": 2, "docx_text": "9", "md_text": "8"}])
        self.assertEqual(by_id[2]["header"], ["Item", "Qty"])
        self.assertEqual((got["unmatched_docx"], got["unmatched_md"], got["unsupported"]), ([], [], []))

    def test_identical_section(self) -> None:
        md = f"# Team\n\n{md_table(self.T1)}\n{md_table(self.T2)}"
        got = self.run_diff(self.docx(self.body()), section_key="team-1", markdown=md)
        self.assertTrue(got["identical"], got)

    def test_md_heading_override(self) -> None:
        md = f"# Staff\n\n{md_table(self.T1)}\n{md_table(self.T2)}"
        got = self.run_diff(self.docx(self.body()), section_key="team-1", markdown=md, md_heading="Staff")
        self.assertTrue(got["identical"], got)
        self.refused(self.docx(self.body(), "b.docx"), section_key="team-1", markdown=md)  # default heading "Team"

    def test_unmatched_tables_on_both_sides(self) -> None:
        md = f"# Team\n\n{md_table(self.T1)}\n{md_table([['Other'], ['z']])}"
        got = self.run_diff(self.docx(self.body()), section_key="team-1", markdown=md)
        self.assertFalse(got["identical"])
        self.assertEqual(got["unmatched_docx"], [2])
        self.assertEqual(got["unmatched_md"], [1])
        self.assertEqual([t["table_id"] for t in got["tables"]], [1])

    def test_unsupported_table_blocks_identical(self) -> None:
        merged = table([[cell("Item", span=2)], [cell("a"), cell("b")]], grid_cols=2)
        docx = self.docx(heading("Team") + table(self.T1) + merged)
        got = self.run_diff(docx, section_key="team-1", markdown=f"# Team\n\n{md_table(self.T1)}")
        self.assertFalse(got["identical"])
        self.assertEqual([u["table_id"] for u in got["unsupported"]], [2])

    def test_duplicate_headers_are_refused(self) -> None:
        md_dup = f"# Team\n\n{md_table(self.T1)}\n{md_table(self.T1)}"
        self.refused(self.docx(self.body()), section_key="team-1", markdown=md_dup)
        docx_dup = self.docx(heading("Team") + table(self.T1) + table(self.T1), "d.docx")
        self.refused(docx_dup, section_key="team-1", markdown=f"# Team\n\n{md_table(self.T1)}")

    def test_zero_tables_either_side_is_refused_not_identical(self) -> None:
        self.refused(self.docx(self.body()), section_key="team-1", markdown="# Team\n\nno tables\n")
        no_tables = self.docx(heading("Team") + para("text"), "n.docx")
        self.refused(no_tables, section_key="team-1", markdown=f"# Team\n\n{md_table(self.T1)}")

    def test_repeated_or_missing_md_heading_is_refused(self) -> None:
        dup = f"# Team\n\n{md_table(self.T1)}\n# Team\n\n{md_table(self.T2)}"
        self.refused(self.docx(self.body()), section_key="team-1", markdown=dup)
        self.refused(self.docx(self.body(), "m.docx"), section_key="team-1", markdown=md_table(self.T1))

    def test_other_sections_tables_are_not_in_scope(self) -> None:
        docx = self.docx(heading("A") + table(self.T1) + heading("B") + table(self.T2))
        got = self.run_diff(docx, section_key="b-1", markdown=f"# B\n\n{md_table(self.T2)}")
        self.assertTrue(got["identical"])
        self.assertEqual([t["table_id"] for t in got["tables"]], [2])


class ReadOnlyTests(TableDiffCase):
    def test_no_writes_to_docx_or_markdown(self) -> None:
        docx = self.docx(table(BASE))
        md = self.dir / "t.md"
        md.write_text(md_table(BASE[:3]), encoding="utf-8")
        before = [(p.read_bytes(), p.stat().st_mtime_ns) for p in (docx, md)]
        self.run_diff(docx, table_id=1, markdown_path=str(md))
        self.run_diff(docx, table_id=1, markdown_path=str(md))
        after = [(p.read_bytes(), p.stat().st_mtime_ns) for p in (docx, md)]
        self.assertEqual(before, after)
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["t.docx", "t.md"])

    def test_not_a_mutating_tool(self) -> None:
        from verified_docx_mcp import middleware

        self.assertNotIn("diff_table_vs_markdown", middleware.MUTATING_TOOLS)


class LiveTableDiffTests(LiveReadsTestCase):
    """The live body holds a table the placeholder file does not."""

    async def asyncSetUp(self) -> None:
        await super().asyncSetUp()
        self.live_docx = build_docx(Path(self._tmp.name) / "live-body.docx", table(BASE))

    async def test_live_read_carries_live_provenance(self) -> None:
        await self.connect_live(self.live_docx)
        got = await self.call(
            table_diff.diff_table_vs_markdown_impl, str(self.target), table_id=1, markdown=md_table(BASE), source="live"
        )
        self.assertEqual(got["source"], "live")
        self.assertIn("body_sha256", got["live"])
        self.assertTrue(got["identical"])

    async def test_auto_reads_live_and_file_read_warns(self) -> None:
        await self.connect_live(self.live_docx)
        auto = await self.call(
            table_diff.diff_table_vs_markdown_impl, str(self.target), table_id=1, markdown=md_table(BASE)
        )
        self.assertEqual(auto["source"], "live")
        # The placeholder file has no table: an explicit file read must not pass as the live one.
        with self.assertRaises(VerifyError) as ctx:
            await self.call(
                table_diff.diff_table_vs_markdown_impl,
                str(self.target),
                table_id=1,
                markdown=md_table(BASE),
                source="file",
            )
        self.assertEqual(_code(ctx.exception), ErrorCode.TABLE_NOT_FOUND)

    async def test_file_read_with_a_session_warns(self) -> None:
        await self.connect_live(self.live_docx)
        target = build_docx(self.target, table(BASE))
        got = await self.call(
            table_diff.diff_table_vs_markdown_impl, str(target), table_id=1, markdown=md_table(BASE), source="file"
        )
        self.assertEqual(got["source"], "file")
        self.assertIn(reads_live.WARNING_LIVE_SESSION_IGNORED, got["warnings"])

    async def test_live_with_no_session_is_unavailable(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            await self.call(
                table_diff.diff_table_vs_markdown_impl,
                str(self.target),
                table_id=1,
                markdown=md_table(BASE),
                source="live",
            )
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_UNAVAILABLE)


if __name__ == "__main__":
    unittest.main()
