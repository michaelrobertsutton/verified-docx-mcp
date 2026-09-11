"""Unit tests for src/verified_docx_mcp/tables.py (issue #28 WP-14).
WP-16a's multi-level w:numPr coverage inside replace_cell_markdown is
added by its own dedicated commit, in this same file.

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

import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
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

    # Multi-level w:numPr inside a cell (issue #28 WP-16a) is covered in
    # its own dedicated commit/test (test_multi_level_bullets_inside_cell_wp16a,
    # WP-16a's own acceptance coverage) rather than here.


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


if __name__ == "__main__":
    unittest.main()
