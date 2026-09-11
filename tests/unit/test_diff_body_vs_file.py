"""Unit tests for diff_body_vs_file (server.execute_diff_body_vs_file) --
issue #28 WP-11a.

Mirrors GoogleDocs-MCP's tests/unit/test_diff_tab_vs_file.py coverage list,
adapted to this server's docx-body-only shape:

- Returns a structured diff when the body and file differ.
- Returns identical=True when the body and file are the same.
- Returns INVALID_INPUT when the docx path does not exist.
- Returns INVALID_INPUT when the file does not exist.
- Returns INVALID_INPUT when file_path resolves outside the allowlist.
- Returns INVALID_INPUT when file_path is a directory, not a regular file.
- Reads via a validated snapshot when Word's owner file is present (core/
  document-backend-protocol.md §4's "reads never refuse" rule), same as
  every other WP-03 read tool (test_read_tools.py's ReadsNeverRefuseTests).
- lossy_elements is surfaced (present only when non-empty), the concrete
  hook for this tool's affirmative-only caveat.
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

FRAG_BODY_MARKDOWN = "The quick **brown** fox jumps over the lazy dog."


class DiffBodyVsFileTestCase(unittest.TestCase):
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

    def _write_md(self, name: str, text: str) -> Path:
        md_path = Path(self._tmp.name) / name
        md_path.write_text(text, encoding="utf-8")
        return md_path


class IdenticalTests(DiffBodyVsFileTestCase):
    def test_identical_when_file_matches_body_projection(self):
        md_path = self._write_md("match.md", FRAG_BODY_MARKDOWN)
        result = server.execute_diff_body_vs_file(str(self.target), str(md_path))

        self.assertTrue(result["identical"])
        self.assertEqual(result["hunks"], [{
            "tag": "equal",
            "body_lines": [FRAG_BODY_MARKDOWN],
            "file_lines": [FRAG_BODY_MARKDOWN],
            "body_range": [1, 1],
            "file_range": [1, 1],
        }])
        self.assertEqual(result["unified_diff"], "")
        self.assertNotIn("lossy_elements", result)
        self.assertEqual(result["path"], str(self.target.resolve()))
        self.assertEqual(result["file_path"], str(md_path))
        self.assertIn(":", result["revision"])
        self.assertEqual(result["warnings"], [])


class DiffersTests(DiffBodyVsFileTestCase):
    def test_structured_diff_when_file_differs(self):
        md_path = self._write_md("differ.md", "The quick **brown** fox leaps over the lazy dog.")
        result = server.execute_diff_body_vs_file(str(self.target), str(md_path))

        self.assertFalse(result["identical"])
        tags = [h["tag"] for h in result["hunks"]]
        self.assertIn("replace", tags)
        self.assertNotEqual(result["unified_diff"], "")
        self.assertTrue(result["unified_diff"].startswith("--- docx:"))
        self.assertIn(f"+++ {md_path}", result["unified_diff"])


class MissingInputTests(DiffBodyVsFileTestCase):
    def test_missing_docx_path_raises_invalid_input(self):
        missing = Path(self._tmp.name) / "nope.docx"
        with self.assertRaises(VerifyError) as ctx:
            server.execute_diff_body_vs_file(str(missing), str(self._write_md("x.md", "x")))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_missing_file_path_raises_invalid_input(self):
        missing = Path(self._tmp.name) / "nope.md"
        with self.assertRaises(VerifyError) as ctx:
            server.execute_diff_body_vs_file(str(self.target), str(missing))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_file_path_outside_allowlist_raises_invalid_input(self):
        with tempfile.TemporaryDirectory() as outside:
            outside_md = Path(outside) / "outside.md"
            outside_md.write_text("x", encoding="utf-8")
            with self.assertRaises(VerifyError) as ctx:
                server.execute_diff_body_vs_file(str(self.target), str(outside_md))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_file_path_directory_raises_invalid_input(self):
        directory = Path(self._tmp.name) / "adir"
        directory.mkdir()
        with self.assertRaises(VerifyError) as ctx:
            server.execute_diff_body_vs_file(str(self.target), str(directory))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("not a regular file", ctx.exception.envelope.message)


class ReadsNeverRefuseTests(DiffBodyVsFileTestCase):
    """core/document-backend-protocol.md §4: DOCX_LOCKED gates writes
    only -- diff_body_vs_file must still succeed, via a validated
    snapshot, when Word's owner file is present. Same pattern as
    test_read_tools.py's ReadsNeverRefuseTests."""

    def test_uses_snapshot_when_owner_file_present(self):
        (Path(self._tmp.name) / "~$frag.docx").write_bytes(b"\x00Someone\x00")
        md_path = self._write_md("match.md", FRAG_BODY_MARKDOWN)
        with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
            result = server.execute_diff_body_vs_file(str(self.target), str(md_path))
        spy.assert_called_once()
        self.assertTrue(result["identical"])


class LossyElementsTests(unittest.TestCase):
    """lossy_elements is the concrete backing for this tool's
    affirmative-only caveat: a merged/nested table drops out of the
    markdown projection entirely, so identical=True cannot be trusted
    when this key is present."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "tables-merged.docx"
        shutil.copyfile(FIXTURES / "tables-merged.docx", self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def test_lossy_elements_surfaced_when_body_has_merged_table(self):
        from verified_docx_mcp import projection

        body_markdown, _, lossy = projection.read_document_markdown(self.target)
        self.assertTrue(lossy)
        md_path = Path(self._tmp.name) / "match.md"
        md_path.write_text(body_markdown, encoding="utf-8")

        result = server.execute_diff_body_vs_file(str(self.target), str(md_path))
        self.assertTrue(result["identical"])
        self.assertIn("lossy_elements", result)
        self.assertEqual(result["lossy_elements"], lossy)


if __name__ == "__main__":
    unittest.main()
