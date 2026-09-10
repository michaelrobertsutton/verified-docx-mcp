"""Unit tests for the WP-03 read tools wired up in server.py: list_parts,
read_document, find_sections, list_page_sections, list_styles.

Covers what test_projection.py does not: path resolution through
paths.resolve_allowed_docx_path, the "reads never refuse" rule (a Word
owner file present triggers paths.snapshot_docx_package rather than
DOCX_LOCKED — core/document-backend-protocol.md §4), format validation,
and the VerifyError -> ErrorCode mapping for a bad part name.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import paths, server
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"


class ReadToolsTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "frag.docx"
        shutil.copyfile(FIXTURES / "frag.docx", self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()


class ListPartsToolTests(ReadToolsTestCase):
    def test_returns_document_part(self):
        result = server.execute_list_parts(str(self.target))
        parts = {p["part"] for p in result["parts"]}
        self.assertIn("word/document.xml", parts)


class ReadDocumentToolTests(ReadToolsTestCase):
    def test_text_format(self):
        result = server.execute_read_document(str(self.target), format="text")
        self.assertEqual(result["text"], "The quick brown fox jumps over the lazy dog.")
        self.assertIn("revision", result)
        self.assertIn(":", result["revision"])
        self.assertIn("document_sha256", result["revision_detail"])

    def test_runs_format(self):
        result = server.execute_read_document(str(self.target), format="runs")
        self.assertTrue(any(r.get("text") == "brown" and r["rPr"]["bold"] for r in result["runs"]))

    def test_markdown_format_is_default(self):
        result = server.execute_read_document(str(self.target))
        self.assertEqual(result["format"], "markdown")
        self.assertIn("**brown**", result["markdown"])

    def test_invalid_format_rejected(self):
        with self.assertRaises(VerifyError) as ctx:
            server.execute_read_document(str(self.target), format="pdf")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_unknown_part_raises_part_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            server.execute_read_document(str(self.target), format="text", part="word/nope.xml")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.PART_NOT_FOUND)


class FindSectionsToolTests(ReadToolsTestCase):
    def test_no_headings_in_frag_docx(self):
        result = server.execute_find_sections(str(self.target))
        self.assertEqual(result["sections"], [])

    def test_headings_in_sections_docx(self):
        target = Path(self._tmp.name) / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", target)
        result = server.execute_find_sections(str(target))
        self.assertEqual(len(result["sections"]), 3)


class ListPageSectionsToolTests(ReadToolsTestCase):
    def test_returns_page_geometry(self):
        result = server.execute_list_page_sections(str(self.target))
        self.assertEqual(len(result["page_sections"]), 1)
        self.assertAlmostEqual(result["page_sections"][0]["page_width_in"], 8.5)


class ListStylesToolTests(ReadToolsTestCase):
    def test_returns_normal_style(self):
        result = server.execute_list_styles(str(self.target))
        ids = {s["style_id"] for s in result["styles"]}
        self.assertIn("Normal", ids)


class ReadsNeverRefuseTests(ReadToolsTestCase):
    """core/document-backend-protocol.md §4: DOCX_LOCKED gates writes
    only. Every WP-03 read tool must still succeed — via a validated
    snapshot — when Word's owner file is present."""

    def _plant_owner_file(self):
        (Path(self._tmp.name) / "~$frag.docx").write_bytes(b"\x00Someone\x00")

    def test_list_parts_uses_snapshot_when_owner_file_present(self):
        self._plant_owner_file()
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            result = server.execute_list_parts(str(self.target))
        spy.assert_called_once()
        self.assertTrue(any(p["part"] == "word/document.xml" for p in result["parts"]))

    def test_read_document_uses_snapshot_when_owner_file_present(self):
        self._plant_owner_file()
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            result = server.execute_read_document(str(self.target), format="text")
        spy.assert_called_once()
        self.assertEqual(result["text"], "The quick brown fox jumps over the lazy dog.")
        # The reported path is the real target (symlink-resolved, e.g.
        # macOS's /var -> /private/var), never the temp snapshot.
        self.assertEqual(Path(result["path"]).resolve(), self.target.resolve())

    def test_find_sections_uses_snapshot_when_owner_file_present(self):
        self._plant_owner_file()
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            server.execute_find_sections(str(self.target))
        spy.assert_called_once()

    def test_list_page_sections_uses_snapshot_when_owner_file_present(self):
        self._plant_owner_file()
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            server.execute_list_page_sections(str(self.target))
        spy.assert_called_once()

    def test_list_styles_uses_snapshot_when_owner_file_present(self):
        self._plant_owner_file()
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            server.execute_list_styles(str(self.target))
        spy.assert_called_once()

    def test_no_owner_file_does_not_snapshot(self):
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            server.execute_read_document(str(self.target), format="text")
        spy.assert_not_called()

    def test_snapshot_temp_file_is_cleaned_up(self):
        self._plant_owner_file()
        before = set(Path(tempfile.gettempdir()).glob("verified-docx-mcp-snapshot-*"))
        server.execute_read_document(str(self.target), format="text")
        after = set(Path(tempfile.gettempdir()).glob("verified-docx-mcp-snapshot-*"))
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
