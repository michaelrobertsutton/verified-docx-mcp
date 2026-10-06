"""File-mode insert_paragraphs (issue #77): styled, colored, optionally tracked
paragraphs inserted before/after an anchored body paragraph in the .docx itself."""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paragraph_insert, paths, projection
from verified_docx_mcp.author import resolve_author_name
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

SPECS = [
    {"text": "Proposed heading", "style": "Heading 2"},
    {"text": "Purple instruction", "color": "7030a0"},
]


class _Case(unittest.TestCase):
    fixture_name = "sections.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / self.fixture_name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def insert(self, anchor="Some overview text.", paragraphs=SPECS, **kwargs):
        return paragraph_insert.execute_insert_paragraphs(str(self.target), anchor, paragraphs, 1, **kwargs)

    def refused(self, code, **kwargs):
        before = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            self.insert(**kwargs)
        self.assertEqual(cm.exception.envelope.error_code, code, cm.exception.envelope.message)
        self.assertEqual(self.target.read_bytes(), before)

    def document_xml(self) -> str:
        with zipfile.ZipFile(self.target) as z:
            return z.read("word/document.xml").decode("utf-8")

    def lines(self) -> list[str]:
        return projection.read_document_text(self.target).split("\n")


class InsertParagraphsFileTests(_Case):
    def test_registered_as_mutating(self):
        self.assertIn("insert_paragraphs", MUTATING_TOOLS)

    def test_insert_after_with_style_and_color(self):
        evidence = self.insert()
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "file")
        lines = self.lines()
        at = lines.index("Some overview text.")
        self.assertEqual(lines[at + 1 : at + 3], ["Proposed heading", "Purple instruction"])
        self.assertEqual(lines[at + 3], "Background")
        xml = self.document_xml()
        self.assertIn('w:val="Heading2"', xml)
        self.assertIn('<w:color w:val="7030A0"', xml)
        self.assertEqual(evidence["inserted"][1]["color"], "7030A0")
        self.assertTrue(evidence["audit_logged"])

    def test_insert_before(self):
        self.insert(paragraphs=SPECS[:1], position="before")
        lines = self.lines()
        self.assertEqual(lines[lines.index("Proposed heading") + 1], "Some overview text.")

    def test_tracked_runs_are_wrapped_in_ins(self):
        evidence = self.insert(track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertEqual(len(evidence["revision_ids"]), 2)
        xml = self.document_xml()
        self.assertEqual(xml.count("<w:ins "), 2)
        self.assertIn(f'w:author="{resolve_author_name()}"', xml)

    def test_style_by_id_and_case_insensitive_name(self):
        with zipfile.ZipFile(FIXTURES / self.fixture_name) as z:
            original = z.read("word/document.xml").decode("utf-8").count('<w:pStyle w:val="Heading')
        self.insert(paragraphs=[{"text": "A", "style": "heading 2"}, {"text": "B", "style": "Heading3"}])
        self.assertEqual(self.document_xml().count('<w:pStyle w:val="Heading'), original + 2)

    def test_refusals_leave_the_file_untouched(self):
        self.refused(ErrorCode.STYLE_NOT_FOUND, paragraphs=[{"text": "t", "style": "No Such Style"}])
        self.refused(ErrorCode.UNSUPPORTED_STYLE_TYPE, paragraphs=[{"text": "t", "style": "Default Paragraph Font"}])
        self.refused(ErrorCode.ZERO_MATCH, anchor="absent text")
        self.refused(ErrorCode.INVALID_INPUT, paragraphs=[])

    def test_duplicate_anchor_is_a_count_mismatch(self):
        self.refused(ErrorCode.MATCH_COUNT_MISMATCH, anchor="text.")


class InsertParagraphsTableAnchorTests(_Case):
    fixture_name = "tables.docx"

    def test_table_cell_anchor_is_a_structural_boundary(self):
        self.refused(ErrorCode.STRUCTURAL_BOUNDARY, anchor="R1C1")


if __name__ == "__main__":
    unittest.main()
