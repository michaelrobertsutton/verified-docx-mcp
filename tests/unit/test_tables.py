"""Unit tests for src/verified_docx_mcp/tables.py (issue #28 WP-14), plus
WP-16a's multi-level w:numPr coverage inside replace_cell_markdown
(ReplaceCellMarkdownTests.test_multi_level_bullets_inside_cell_wp16a).

Fixture provenance (see tests/fixtures/README.md):
  - tables.docx: Word-authored 2x2 table (R1C1..R2C2), no merges.
  - tables-merged.docx: tables.docx with its one w:tbl replaced (via
    xml.etree.ElementTree, original namespace prefixes preserved, every
    other part byte-for-byte unchanged) by a hand-built one exercising a
    w:vMerge "restart"/"continue" pair, a w:gridSpan="2" cell, and a
    nested w:tbl -- standard, unambiguous OOXML markup this environment
    has no way to drive Word into producing via AppleScript (no GUI merge
    command is scriptable there), per that file's own provenance note.

Fixture (b) (issue #28 plan WP-14's own acceptance text): a merged-cell
table where replace_table_row refuses and replace_cell_markdown succeeds
with w:tcPr byte-identical -- see MergedTableFixtureBTests below.
"""

from __future__ import annotations

import copy
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paths, projection, tables
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

_EVIDENCE_KEYS = {
    "applied",
    "match_count",
    "rung",
    "before",
    "after",
    "revision_before",
    "revision_after",
    "audit_logged",
    "conflict_copy_detected",
}


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "tables.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()


def _document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(projection.DEFAULT_PART).decode("utf-8")


def _tc_pr_bytes(path: Path, table_id: int, row_index: int, cell_index: int) -> bytes | None:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read(projection.DEFAULT_PART))
    info = tables.get_table_impl(path, table_id)  # sanity: table exists
    assert info["table_id"] == table_id
    tbl = tables._find_table_element(root, table_id)
    tr = tables._find_row(tbl, row_index, table_id=table_id)
    tc = tables._find_cell(tr, cell_index, table_id=table_id, row_index=row_index)
    tcpr = tables._cell_tcpr(tc)
    return ET.tostring(tcpr) if tcpr is not None else None


def _find_cell_xml(path: Path, table_id: int, row_index: int, cell_index: int) -> Any:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read(projection.DEFAULT_PART))
    tbl = tables._find_table_element(root, table_id)
    tr = tables._find_row(tbl, row_index, table_id=table_id)
    return tables._find_cell(tr, cell_index, table_id=table_id, row_index=row_index)


def _row_xml(path: Path, table_id: int, row_index: int) -> Any:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read(projection.DEFAULT_PART))
    tbl = tables._find_table_element(root, table_id)
    return tables._find_row(tbl, row_index, table_id=table_id)


def _duplicate_body_paragraph(path: Path, body_index: int) -> None:
    """Test-only fixture mutation: duplicate the top-level body paragraph
    at *body_index* (0-based, into body_children) right after itself --
    used to exercise MATCH_COUNT_MISMATCH on after_paragraph_text without
    hand-building a whole new fixture (every real fixture's sections have
    exactly one non-heading paragraph, never two identical ones)."""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        parts = {n: zf.read(n) for n in names}
    raw = parts[projection.DEFAULT_PART]
    decls = mutations._capture_source_namespaces(raw)
    root = ET.fromstring(raw)
    body = mutations._find_body(root)
    body_children, _sect_pr = mutations._split_body(body)
    dup = copy.deepcopy(body_children[body_index])
    body.insert(body_index + 1, dup)
    parts[projection.DEFAULT_PART] = mutations._serialize_xml(root, decls)
    with zipfile.ZipFile(path, "w") as zf:
        for name, content in parts.items():
            zf.writestr(name, content)


class MutatingToolsRegistrationTests(unittest.TestCase):
    def test_table_tools_registered(self):
        self.assertIn("replace_table_row", MUTATING_TOOLS)
        self.assertIn("replace_cell_markdown", MUTATING_TOOLS)
        self.assertIn("insert_table", MUTATING_TOOLS)

    def test_read_tools_not_registered(self):
        self.assertNotIn("list_tables", MUTATING_TOOLS)
        self.assertNotIn("get_table", MUTATING_TOOLS)


class ListTablesReadTests(unittest.TestCase):
    def test_plain_table(self):
        result = tables.list_tables_impl(FIXTURES / "tables.docx")
        self.assertEqual(len(result), 1)
        info = result[0]
        self.assertEqual(info["table_id"], 1)
        self.assertEqual(info["row_count"], 2)
        self.assertEqual(info["col_count"], 2)
        self.assertFalse(info["has_merged_cells"])
        self.assertFalse(info["has_nested_table"])
        self.assertIsNone(info["nested_in_table_id"])

    def test_merged_and_nested_table(self):
        result = tables.list_tables_impl(FIXTURES / "tables-merged.docx")
        self.assertEqual(len(result), 2)
        outer, nested = result
        self.assertEqual(outer["table_id"], 1)
        self.assertEqual(outer["row_count"], 4)
        self.assertEqual(outer["col_count"], 2)
        self.assertTrue(outer["has_merged_cells"])
        self.assertTrue(outer["has_nested_table"])
        self.assertIsNone(outer["nested_in_table_id"])

        self.assertEqual(nested["table_id"], 2)
        self.assertEqual(nested["row_count"], 2)
        self.assertEqual(nested["col_count"], 1)
        self.assertFalse(nested["has_merged_cells"])
        self.assertFalse(nested["has_nested_table"])
        self.assertEqual(nested["nested_in_table_id"], 1)

    def test_table_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.get_table_impl(FIXTURES / "tables.docx", 99)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.TABLE_NOT_FOUND)


class GetTableReadTests(unittest.TestCase):
    def test_plain_table_cells(self):
        result = tables.get_table_impl(FIXTURES / "tables.docx", 1)
        self.assertEqual(result["row_count"], 2)
        self.assertEqual(result["col_count"], 2)
        texts = [[c["text"] for c in row] for row in result["rows"]]
        self.assertEqual(texts, [["R1C1", "R1C2"], ["R2C1", "R2C2"]])
        for row in result["rows"]:
            for cell in row:
                self.assertEqual(cell["grid_span"], 1)
                self.assertEqual(cell["v_merge"], "none")

    def test_merged_table_reports_gridspan_and_vmerge(self):
        result = tables.get_table_impl(FIXTURES / "tables-merged.docx", 1)
        rows = result["rows"]
        # Row 1: H1 / H2, no merges.
        self.assertEqual([c["text"] for c in rows[0]], ["H1", "H2"])
        self.assertEqual([c["v_merge"] for c in rows[0]], ["none", "none"])
        # Row 2: "Merged down" (vMerge restart) / B2.
        self.assertEqual(rows[1][0]["text"], "Merged down")
        self.assertEqual(rows[1][0]["v_merge"], "restart")
        self.assertEqual(rows[1][1]["v_merge"], "none")
        # Row 3: vMerge continuation cell (empty text) / B3 cell (nested
        # table flattened into its own text).
        self.assertEqual(rows[2][0]["text"], "")
        self.assertEqual(rows[2][0]["v_merge"], "continue")
        self.assertIn("B3", rows[2][1]["text"])
        self.assertIn("Nested A", rows[2][1]["text"])
        self.assertIn("Nested B", rows[2][1]["text"])
        # Row 4: single gridSpan=2 cell.
        self.assertEqual(len(rows[3]), 1)
        self.assertEqual(rows[3][0]["grid_span"], 2)
        self.assertEqual(rows[3][0]["text"], "Spanned row")

    def test_nested_table_own_cells(self):
        result = tables.get_table_impl(FIXTURES / "tables-merged.docx", 2)
        texts = [[c["text"] for c in row] for row in result["rows"]]
        self.assertEqual(texts, [["Nested A"], ["Nested B"]])


class ReplaceTableRowTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_replaces_row_content_leaving_other_rows_alone(self):
        evidence = tables.execute_replace_table_row(str(self.target), 1, 1, ["New A", "New B"])
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["match_count"], 1)

        result = tables.get_table_impl(self.target, 1)
        texts = [[c["text"] for c in row] for row in result["rows"]]
        self.assertEqual(texts, [["New A", "New B"], ["R2C1", "R2C2"]])

    def test_cell_count_mismatch_refuses(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_replace_table_row(str(self.target), 1, 1, ["only one"])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_row_out_of_range(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_replace_table_row(str(self.target), 1, 99, ["a", "b"])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.TABLE_ROW_NOT_FOUND)

    def test_track_changes_wraps_old_and_new_content(self):
        evidence = tables.execute_replace_table_row(str(self.target), 1, 1, ["New A", "New B"], track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        xml = _document_xml(self.target)
        self.assertIn("<w:ins ", xml)
        self.assertIn("<w:del ", xml)
        self.assertIn("R1C1", xml)  # old text preserved inside w:delText
        # A tracked read still projects the NEW text as current.
        result = tables.get_table_impl(self.target, 1)
        self.assertEqual(result["rows"][0][0]["text"], "New A")


class MergedTableFixtureBTests(_TempFixtureCase):
    """Fixture (b), verbatim from the plan: a merged-cell table where
    replace_table_row refuses, replace_cell_markdown succeeds, and
    w:tcPr is byte-identical afterwards."""

    fixture_name = "tables-merged.docx"

    def test_replace_table_row_refuses_whole_table(self):
        # Row 1 (H1/H2) carries no merge of its own -- the refusal must
        # still fire because SOME cell in the table is merged/nested.
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_replace_table_row(str(self.target), 1, 1, ["x", "y"])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.MERGED_OR_NESTED_TABLE)

    def test_replace_cell_markdown_succeeds_tcpr_byte_identical(self):
        # Target the vMerge "restart" cell itself (row 2, cell 1) -- its
        # w:tcPr carries the actual merge marker this test must prove
        # survives untouched, not an empty <w:tcPr/>.
        before_tcpr = _tc_pr_bytes(self.target, 1, 2, 1)
        self.assertIn(b"vMerge", before_tcpr)

        evidence = tables.execute_replace_cell_markdown(str(self.target), 1, 2, 1, "New merged-cell content")
        self.assertTrue(evidence["applied"])

        after_tcpr = _tc_pr_bytes(self.target, 1, 2, 1)
        self.assertEqual(before_tcpr, after_tcpr, "w:tcPr must be byte-identical across replace_cell_markdown")

        result = tables.get_table_impl(self.target, 1)
        self.assertEqual(result["rows"][1][0]["text"], "New merged-cell content")
        # The merge itself is still intact -- v_merge is still "restart".
        self.assertEqual(result["rows"][1][0]["v_merge"], "restart")

    def test_replace_cell_markdown_on_gridspan_cell_tcpr_byte_identical(self):
        before_tcpr = _tc_pr_bytes(self.target, 1, 4, 1)
        self.assertIn(b"gridSpan", before_tcpr)
        tables.execute_replace_cell_markdown(str(self.target), 1, 4, 1, "New spanned content")
        after_tcpr = _tc_pr_bytes(self.target, 1, 4, 1)
        self.assertEqual(before_tcpr, after_tcpr)
        result = tables.get_table_impl(self.target, 1)
        self.assertEqual(result["rows"][3][0]["grid_span"], 2)
        self.assertEqual(result["rows"][3][0]["text"], "New spanned content")


class ReplaceCellMarkdownTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_basic_replace(self):
        evidence = tables.execute_replace_cell_markdown(str(self.target), 1, 1, 1, "**Bold** content")
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        result = tables.get_table_impl(self.target, 1)
        self.assertEqual(result["rows"][0][0]["text"], "**Bold** content")

    def test_cell_out_of_range(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_replace_cell_markdown(str(self.target), 1, 1, 99, "x")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.TABLE_CELL_NOT_FOUND)

    def test_track_changes(self):
        evidence = tables.execute_replace_cell_markdown(str(self.target), 1, 1, 1, "Tracked new", track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        xml = _document_xml(self.target)
        self.assertIn("<w:ins ", xml)
        self.assertIn("<w:del ", xml)
        self.assertIn("R1C1", xml)

    def test_multi_level_bullets_inside_cell_wp16a(self):
        """issue #28 WP-16a: multi-level w:numPr inside replace_cell_markdown,
        sharing one numId/abstractNum tree across nesting levels via
        increasing w:ilvl (the PR #3 fix markdown_to_ooxml.py's own module
        docstring documents), reused as-is via render_blocks -- this test
        proves it end-to-end inside a table cell, not just at body level.
        """
        markdown = "- Top Skills\n    - Nested skill one\n    - Nested skill two\n- Second top item\n"
        evidence = tables.execute_replace_cell_markdown(str(self.target), 1, 1, 1, markdown)
        self.assertTrue(evidence["applied"])

        with zipfile.ZipFile(self.target) as zf:
            doc_root = ET.fromstring(zf.read(projection.DEFAULT_PART))
            numbering_root = ET.fromstring(zf.read("word/numbering.xml"))

        tbl = tables._find_table_element(doc_root, 1)
        tr = tables._find_row(tbl, 1, table_id=1)
        tc = tables._find_cell(tr, 1, table_id=1, row_index=1)
        paragraphs = [c for c in tc if projection._ln(c) == "p"]
        self.assertEqual(len(paragraphs), 4)

        def _num_pr(p):
            for child in p:
                if projection._ln(child) == "pPr":
                    for gc in child:
                        if projection._ln(gc) == "numPr":
                            ilvl = None
                            num_id = None
                            for ggc in gc:
                                if projection._ln(ggc) == "ilvl":
                                    ilvl = int(projection._attr(ggc, "val"))
                                elif projection._ln(ggc) == "numId":
                                    num_id = int(projection._attr(ggc, "val"))
                            return ilvl, num_id
            return None

        num_prs = [_num_pr(p) for p in paragraphs]
        ilvls = [np[0] for np in num_prs]
        num_ids = {np[1] for np in num_prs}
        self.assertEqual(ilvls, [0, 1, 1, 0])
        # One shared numId/abstractNum tree for the whole list, not a
        # fresh one per nesting level (the exact bug PR #3 fixed).
        self.assertEqual(len(num_ids), 1)
        (shared_num_id,) = num_ids

        # word/numbering.xml actually carries two w:lvl entries (ilvl 0
        # and 1) under that num's abstractNum -- the tree, not just the
        # paragraph-level ilvl/numId values, is real.
        abstract_id = None
        for num in numbering_root:
            if projection._ln(num) == "num" and projection._attr(num, "numId") == str(shared_num_id):
                for child in num:
                    if projection._ln(child) == "abstractNumId":
                        abstract_id = projection._attr(child, "val")
        self.assertIsNotNone(abstract_id)
        lvl_count = 0
        for abstract in numbering_root:
            if projection._ln(abstract) == "abstractNum" and projection._attr(abstract, "abstractNumId") == abstract_id:
                lvl_count = sum(1 for c in abstract if projection._ln(c) == "lvl")
        self.assertEqual(lvl_count, 2)

        # Round-trip via get_table's own text rendering: still readable,
        # content preserved (flattened to one line, same as any other
        # cell paragraph join -- list markers are not part of this
        # tool's own text contract, only the underlying w:numPr is).
        result = tables.get_table_impl(self.target, 1)
        text = result["rows"][0][0]["text"]
        self.assertIn("Top Skills", text)
        self.assertIn("Nested skill one", text)
        self.assertIn("Second top item", text)

    def test_three_level_nesting_wp16a(self):
        """A third nesting level (ilvl 0/1/2) inside one cell -- the KP
        resume rendering contract's own accomplishment bullets can nest
        this deep (issue #28 plan WP-16); one shared numId/abstractNum
        tree must still cover all three levels, not just two."""
        markdown = "- A\n    - B\n        - C\n"
        tables.execute_replace_cell_markdown(str(self.target), 1, 1, 1, markdown)

        with zipfile.ZipFile(self.target) as zf:
            doc_root = ET.fromstring(zf.read(projection.DEFAULT_PART))
            numbering_root = ET.fromstring(zf.read("word/numbering.xml"))
        tbl = tables._find_table_element(doc_root, 1)
        tr = tables._find_row(tbl, 1, table_id=1)
        tc = tables._find_cell(tr, 1, table_id=1, row_index=1)
        paragraphs = [c for c in tc if projection._ln(c) == "p"]

        def _ilvl_numid(p):
            for child in p:
                if projection._ln(child) == "pPr":
                    for gc in child:
                        if projection._ln(gc) == "numPr":
                            ilvl = num_id = None
                            for ggc in gc:
                                if projection._ln(ggc) == "ilvl":
                                    ilvl = int(projection._attr(ggc, "val"))
                                elif projection._ln(ggc) == "numId":
                                    num_id = int(projection._attr(ggc, "val"))
                            return ilvl, num_id
            return None

        pairs = [_ilvl_numid(p) for p in paragraphs]
        self.assertEqual([p[0] for p in pairs], [0, 1, 2])
        num_ids = {p[1] for p in pairs}
        self.assertEqual(len(num_ids), 1, "all three levels share one numId")
        (num_id,) = num_ids

        abstract_id = next(
            projection._attr(c, "val")
            for num in numbering_root
            if projection._ln(num) == "num" and projection._attr(num, "numId") == str(num_id)
            for c in num
            if projection._ln(c) == "abstractNumId"
        )
        lvl_count = sum(
            1
            for abstract in numbering_root
            if projection._ln(abstract) == "abstractNum" and projection._attr(abstract, "abstractNumId") == abstract_id
            for c in abstract
            if projection._ln(c) == "lvl"
        )
        self.assertEqual(lvl_count, 3)

    def test_two_cells_each_get_their_own_numid_wp16a(self):
        """Two SEPARATE replace_cell_markdown calls, each writing its own
        nested list into a different cell of the SAME table, must not
        collide on numId -- each call's StyleContext.build re-reads the
        document's own numbering.xml fresh, so a second write's ids are
        allocated above whatever the first write already landed on disk."""
        tables.execute_replace_cell_markdown(str(self.target), 1, 1, 1, "- X\n    - Y\n")
        tables.execute_replace_cell_markdown(str(self.target), 1, 1, 2, "- P\n    - Q\n")

        with zipfile.ZipFile(self.target) as zf:
            doc_root = ET.fromstring(zf.read(projection.DEFAULT_PART))
        tbl = tables._find_table_element(doc_root, 1)
        tr = tables._find_row(tbl, 1, table_id=1)
        tc1 = tables._find_cell(tr, 1, table_id=1, row_index=1)
        tc2 = tables._find_cell(tr, 2, table_id=1, row_index=1)

        def _num_ids(tc):
            found = set()
            for p in tc:
                if projection._ln(p) != "p":
                    continue
                for child in p:
                    if projection._ln(child) == "pPr":
                        for gc in child:
                            if projection._ln(gc) == "numPr":
                                for ggc in gc:
                                    if projection._ln(ggc) == "numId":
                                        found.add(projection._attr(ggc, "val"))
            return found

        ids1, ids2 = _num_ids(tc1), _num_ids(tc2)
        self.assertEqual(len(ids1), 1)
        self.assertEqual(len(ids2), 1)
        self.assertNotEqual(ids1, ids2, "each cell's own write must get a distinct numId, not reuse the other's")


class InsertTableTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_insert_table_appends_new_table(self):
        evidence = tables.execute_insert_table(
            str(self.target), [["A1", "B1"], ["A2", "B2"]], "TableNormal"
        )
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        self.assertEqual(evidence["table_id"], 2)  # tables.docx already has table 1

        tables_list = tables.list_tables_impl(self.target)
        self.assertEqual(len(tables_list), 2)
        result = tables.get_table_impl(self.target, 2)
        texts = [[c["text"] for c in row] for row in result["rows"]]
        self.assertEqual(texts, [["A1", "B1"], ["A2", "B2"]])

    def test_insert_table_sets_column_widths_from_page_section(self):
        tables.execute_insert_table(str(self.target), [["A", "B"]], "TableNormal")
        with zipfile.ZipFile(self.target) as zf:
            doc_root = ET.fromstring(zf.read(projection.DEFAULT_PART))
        tbl = tables._find_table_element(doc_root, 2)
        grid = next(c for c in tbl if projection._ln(c) == "tblGrid")
        widths = [projection._attr(c, "w") for c in grid if projection._ln(c) == "gridCol"]
        # 6.5in usable width (8.5in page - 1in margins each side), split
        # evenly across 2 columns -> 4680 dxa each (6.5*1440/2).
        self.assertEqual(widths, ["4680", "4680"])

    def test_style_id_required_and_validated(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(str(self.target), [["A"]], "NotARealStyle")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.STYLE_NOT_FOUND)

    def test_empty_rows_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(str(self.target), [], "TableNormal")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_track_changes_wraps_new_table_runs(self):
        evidence = tables.execute_insert_table(
            str(self.target), [["A1"]], "TableNormal", track_changes=True
        )
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        xml = _document_xml(self.target)
        self.assertIn("<w:ins ", xml)


# ---------------------------------------------------------------------------
# issue #100: insert_table structured cell objects (gridSpan/vMerge/shading/
# header rows/placement anchor). Group numbers below match the issue's
# implementation spec ("Tests" section, groups 1-11); group 1 (back-compat)
# is covered by InsertTableTests above, unchanged.
# ---------------------------------------------------------------------------


class InsertTableStructuredCellsTests(_TempFixtureCase):
    """Groups 2, 3, 4, 5: the CellSpec object form."""

    fixture_name = "tables.docx"

    def test_title_row_span_reports_gridspan_and_is_lossy_on_read(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [[{"markdown": "Program Title", "span": 2}], ["Role", "PM"]],
            "TableNormal",
        )
        table_id = evidence["table_id"]
        info = tables.get_table_impl(self.target, table_id)
        self.assertEqual(info["col_count"], 2)
        self.assertTrue(info["has_merged_cells"])
        self.assertEqual(info["rows"][0][0]["grid_span"], 2)
        self.assertEqual(info["rows"][0][0]["text"], "Program Title")
        self.assertEqual(evidence["merged_cells"], 1)

        markdown, _warnings, lossy = projection.read_document_markdown(self.target)
        self.assertEqual(markdown.count("Program Title"), 1)
        self.assertTrue(any(item["kind"] == "table_merge" for item in lossy))
        title_line = next(line for line in markdown.splitlines() if "Program Title" in line)
        self.assertEqual(title_line, "| Program Title |  |")  # text once, one empty cell

    def test_explicit_grid_dxa_widths(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [[{"markdown": "Title", "span": 2}], ["A", "B"]],
            "TableNormal",
            grid_dxa=[5935, 3415],
        )
        table_id = evidence["table_id"]
        with zipfile.ZipFile(self.target) as zf:
            root = ET.fromstring(zf.read(projection.DEFAULT_PART))
        tbl = tables._find_table_element(root, table_id)
        tblpr = next(c for c in tbl if projection._ln(c) == "tblPr")
        tblw = next(c for c in tblpr if projection._ln(c) == "tblW")
        self.assertEqual(projection._attr(tblw, "w"), "9350")
        self.assertEqual(projection._attr(tblw, "type"), "dxa")
        grid = next(c for c in tbl if projection._ln(c) == "tblGrid")
        widths = [projection._attr(c, "w") for c in grid if projection._ln(c) == "gridCol"]
        self.assertEqual(widths, ["5935", "3415"])

        title_tc = _find_cell_xml(self.target, table_id, 1, 1)
        tcpr = tables._cell_tcpr(title_tc)
        tcw = next(c for c in tcpr if projection._ln(c) == "tcW")
        self.assertEqual(projection._attr(tcw, "w"), "9350")  # spans both grid columns

    def test_header_rows_and_cant_split(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [["H1", "H2"], ["A", "B"], ["C", "D"]],
            "TableNormal",
            header_rows=1,
            cant_split=True,
        )
        table_id = evidence["table_id"]

        def _trpr_children(row_index: int) -> set[str]:
            tr = _row_xml(self.target, table_id, row_index)
            trpr = next((c for c in tr if projection._ln(c) == "trPr"), None)
            return {projection._ln(c) for c in trpr} if trpr is not None else set()

        self.assertEqual(_trpr_children(1), {"cantSplit", "tblHeader"})
        self.assertEqual(_trpr_children(2), {"cantSplit"})
        self.assertEqual(_trpr_children(3), {"cantSplit"})

    def test_cell_formatting_present_only_on_requested_cell(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [
                [
                    {
                        "markdown": "Styled",
                        "fill": "3b3838",
                        "color": "ffffff",
                        "bold": True,
                        "align": "center",
                        "valign": "center",
                    },
                    "Plain",
                ]
            ],
            "TableNormal",
        )
        table_id = evidence["table_id"]
        styled_xml = ET.tostring(_find_cell_xml(self.target, table_id, 1, 1), encoding="unicode")
        plain_xml = ET.tostring(_find_cell_xml(self.target, table_id, 1, 2), encoding="unicode")

        self.assertIn('w:fill="3B3838"', styled_xml)
        self.assertIn('w:val="FFFFFF"', styled_xml)
        self.assertIn("<w:b", styled_xml)
        self.assertIn("<w:bCs", styled_xml)
        self.assertIn('w:jc w:val="center"', styled_xml)
        self.assertIn('w:vAlign w:val="center"', styled_xml)

        self.assertNotIn("w:fill", plain_xml)
        self.assertNotIn("<w:b", plain_xml)
        self.assertNotIn("w:jc", plain_xml)
        self.assertNotIn("w:vAlign", plain_xml)


class InsertTableValidationTests(_TempFixtureCase):
    """Group 6: strict grid/hex/header_rows validation."""

    fixture_name = "tables.docx"

    def test_grid_mismatch_names_row_and_both_numbers(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [[{"markdown": "A", "span": 2}], ["a", "b", "c"]],
                "TableNormal",
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        message = str(ctx.exception)
        self.assertIn("row 0", message)
        self.assertIn("2", message)
        self.assertIn("3", message)

    def test_bad_hex_fill_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target), [[{"markdown": "A", "fill": "not-a-color"}]], "TableNormal"
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_header_rows_at_or_above_row_count_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(str(self.target), [["a", "b"]], "TableNormal", header_rows=1)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class InsertTableAnchorTests(_TempFixtureCase):
    """Group 7: anchor placement on sections.docx (Overview / Background /
    Next Steps, one body paragraph per section)."""

    fixture_name = "sections.docx"

    def test_position_start_lands_right_after_the_heading(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [["A", "B"]],
            "TableNormal",
            anchor={"section_key": "overview-1", "position": "start"},
        )
        self.assertEqual(
            evidence["anchor_resolved"], {"body_index": 1, "section_key": "overview-1", "after_table_id": None}
        )
        markdown, _w, _l = projection.read_document_markdown(self.target)
        lines = [line for line in markdown.splitlines() if line.strip()]
        self.assertLess(lines.index("# Overview"), lines.index("| A | B |"))
        self.assertLess(lines.index("| A | B |"), lines.index("Some overview text."))

    def test_position_end_lands_before_the_next_heading(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [["A", "B"]],
            "TableNormal",
            anchor={"section_key": "overview-1", "position": "end"},
        )
        self.assertEqual(evidence["anchor_resolved"]["section_key"], "overview-1")
        markdown, _w, _l = projection.read_document_markdown(self.target)
        lines = [line for line in markdown.splitlines() if line.strip()]
        self.assertLess(lines.index("Some overview text."), lines.index("| A | B |"))
        self.assertLess(lines.index("| A | B |"), lines.index("## Background"))

    def test_after_paragraph_text_lands_after_that_paragraph(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [["A", "B"]],
            "TableNormal",
            anchor={"section_key": "overview-1", "after_paragraph_text": "Some overview text."},
        )
        self.assertEqual(evidence["anchor_resolved"]["section_key"], "overview-1")
        markdown, _w, _l = projection.read_document_markdown(self.target)
        lines = [line for line in markdown.splitlines() if line.strip()]
        self.assertLess(lines.index("Some overview text."), lines.index("| A | B |"))
        self.assertLess(lines.index("| A | B |"), lines.index("## Background"))

    def test_unknown_section_key_raises_section_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [["A"]],
                "TableNormal",
                anchor={"section_key": "does-not-exist", "position": "start"},
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.SECTION_NOT_FOUND)

    def test_after_paragraph_text_zero_matches_raises_zero_match(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [["A"]],
                "TableNormal",
                anchor={"section_key": "overview-1", "after_paragraph_text": "Not in this section."},
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.ZERO_MATCH)

    def test_after_paragraph_text_matching_twice_raises_match_count_mismatch(self):
        _duplicate_body_paragraph(self.target, 3)  # "Background text." (body index 3)
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [["A"]],
                "TableNormal",
                anchor={"section_key": "background-1", "after_paragraph_text": "Background text."},
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.MATCH_COUNT_MISMATCH)

    def test_table_id_when_inserted_before_an_existing_table(self):
        """Group 8: table_id correctness when the new table lands BEFORE
        an existing one in document order."""
        first = tables.execute_insert_table(str(self.target), [["X"]], "TableNormal")
        self.assertEqual(first["table_id"], 1)
        second = tables.execute_insert_table(
            str(self.target),
            [["Y"]],
            "TableNormal",
            anchor={"section_key": "overview-1", "position": "start"},
        )
        self.assertEqual(second["table_id"], 1)
        self.assertEqual([t["table_id"] for t in tables.list_tables_impl(self.target)], [1, 2])
        self.assertEqual(tables.get_table_impl(self.target, 1)["rows"][0][0]["text"], "Y")
        self.assertEqual(tables.get_table_impl(self.target, 2)["rows"][0][0]["text"], "X")


class InsertTableAfterTableIdAnchorTests(_TempFixtureCase):
    """Group 7 (continued): the after_table_id anchor form on tables.docx."""

    fixture_name = "tables.docx"

    def test_after_table_id_places_table_right_after_and_gets_next_id(self):
        evidence = tables.execute_insert_table(
            str(self.target), [["new"]], "TableNormal", anchor={"after_table_id": 1}
        )
        self.assertEqual(evidence["table_id"], 2)
        self.assertEqual(
            evidence["anchor_resolved"], {"body_index": 2, "section_key": None, "after_table_id": 1}
        )
        markdown, _w, _l = projection.read_document_markdown(self.target)
        lines = [line for line in markdown.splitlines() if line.strip()]
        self.assertLess(lines.index("| new |"), lines.index("After the table."))

    def test_after_table_id_missing_raises_table_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target), [["x"]], "TableNormal", anchor={"after_table_id": 99}
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.TABLE_NOT_FOUND)


class InsertTableAfterNestedTableIdTests(_TempFixtureCase):
    fixture_name = "tables-merged.docx"

    def test_after_table_id_on_nested_table_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target), [["x"]], "TableNormal", anchor={"after_table_id": 2}
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("nested", str(ctx.exception).lower())


class InsertTableVMergeTests(_TempFixtureCase):
    """Group 9: a v_merge restart/continue pair."""

    fixture_name = "tables.docx"

    def test_restart_then_continue_reports_v_merge(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [
                [{"markdown": "Merged", "v_merge": "restart"}, "B1"],
                [{"markdown": "", "v_merge": "continue"}, "B2"],
            ],
            "TableNormal",
        )
        table_id = evidence["table_id"]
        info = tables.get_table_impl(self.target, table_id)
        self.assertEqual(info["rows"][0][0]["v_merge"], "restart")
        self.assertEqual(info["rows"][1][0]["v_merge"], "continue")
        self.assertEqual(info["rows"][1][0]["text"], "")
        self.assertEqual(evidence["merged_cells"], 2)

    def test_continue_with_nonempty_markdown_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [
                    [{"markdown": "Merged", "v_merge": "restart"}],
                    [{"markdown": "not empty", "v_merge": "continue"}],
                ],
                "TableNormal",
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_continue_with_no_restart_above_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            tables.execute_insert_table(
                str(self.target),
                [["plain"], [{"markdown": "", "v_merge": "continue"}]],
                "TableNormal",
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class InsertTableTrackChangesStructuredTests(_TempFixtureCase):
    """Group 10: track_changes=True extended to a dict-form row."""

    fixture_name = "tables.docx"

    def test_track_changes_wraps_every_run_in_structured_row(self):
        evidence = tables.execute_insert_table(
            str(self.target),
            [[{"markdown": "**Spanning title**", "span": 2, "fill": "3B3838"}], ["a", "b"]],
            "TableNormal",
            track_changes=True,
        )
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        xml = _document_xml(self.target)
        self.assertIn("<w:ins ", xml)
        self.assertIn("Spanning title", xml)


class InsertTableRealWorldShapeEndToEndTests(_TempFixtureCase):
    """Group 11: two real-world merged-cell table shapes, built
    end-to-end -- a contract-information table (2 columns, each program a
    3-row block whose first row spans both columns with a fill) and a
    proof table (a spanning title row, a shaded header row via
    header_rows=1, body rows)."""

    fixture_name = "tables.docx"

    def test_contract_information_table_program_blocks(self):
        programs = [("Program A", "Prime"), ("Program B", "Subcontractor")]
        rows: list[list[Any]] = []
        for name, role in programs:
            rows.append(
                [
                    {
                        "markdown": f"**{name}** - Contractor, {role}",
                        "span": 2,
                        "fill": "3B3838",
                        "color": "FFFFFF",
                        "bold": True,
                    }
                ]
            )
            rows.append(["Contract Number", "GS-00F-1234X"])
            rows.append(["Period of Performance", "2024-2029"])

        evidence = tables.execute_insert_table(str(self.target), rows, "TableNormal", grid_dxa=[5935, 3415])
        table_id = evidence["table_id"]
        info = tables.get_table_impl(self.target, table_id)
        self.assertEqual(info["row_count"], 6)
        self.assertEqual(info["col_count"], 2)
        self.assertTrue(info["has_merged_cells"])
        self.assertEqual(info["rows"][0][0]["grid_span"], 2)
        self.assertIn("Program A", info["rows"][0][0]["text"])
        self.assertEqual(info["rows"][3][0]["grid_span"], 2)
        self.assertIn("Program B", info["rows"][3][0]["text"])
        self.assertEqual(evidence["merged_cells"], 2)

        title_xml = ET.tostring(_find_cell_xml(self.target, table_id, 1, 1), encoding="unicode")
        self.assertIn('w:fill="3B3838"', title_xml)

    def test_proof_table_title_and_shaded_header_row(self):
        rows = [
            [
                {
                    "markdown": "**Table 3: Proof Points**",
                    "span": 2,
                    "fill": "3B3838",
                    "color": "FFFFFF",
                    "align": "center",
                }
            ],
            [
                {"markdown": "**Metric**", "fill": "D9D9D9", "bold": True},
                {"markdown": "**Result**", "fill": "D9D9D9", "bold": True},
            ],
            ["Uptime", "99.99%"],
            ["Response Time", "<2h"],
        ]
        evidence = tables.execute_insert_table(str(self.target), rows, "TableNormal", header_rows=1)
        table_id = evidence["table_id"]
        info = tables.get_table_impl(self.target, table_id)
        self.assertEqual(info["row_count"], 4)
        self.assertEqual(info["col_count"], 2)
        self.assertEqual(info["rows"][0][0]["grid_span"], 2)
        self.assertIn("Table 3", info["rows"][0][0]["text"])

        title_row = _row_xml(self.target, table_id, 1)
        trpr = next((c for c in title_row if projection._ln(c) == "trPr"), None)
        self.assertIsNotNone(trpr)
        self.assertTrue(any(projection._ln(c) == "tblHeader" for c in trpr))

        header_cell_xml = ET.tostring(_find_cell_xml(self.target, table_id, 2, 1), encoding="unicode")
        self.assertIn('w:fill="D9D9D9"', header_cell_xml)
        self.assertIn("<w:b", header_cell_xml)


if __name__ == "__main__":
    unittest.main()
