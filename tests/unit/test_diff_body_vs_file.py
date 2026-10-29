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
        self.assertTrue(result["identical_raw"])
        self.assertTrue(result["normalized"])
        # Identical: no hunks at all, never the document echoed back (#110).
        self.assertEqual(result["hunks"], [])
        self.assertNotIn("unified_diff", result)
        self.assertFalse(result["truncated"])
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
        self.assertNotIn("unified_diff", result)

    def test_unified_format_returns_only_the_unified_diff(self):
        md_path = self._write_md("differ.md", "The quick **brown** fox leaps over the lazy dog.")
        result = server.execute_diff_body_vs_file(str(self.target), str(md_path), format="unified")

        self.assertNotIn("hunks", result)
        self.assertTrue(result["unified_diff"].startswith("--- docx:"))
        self.assertIn(f"+++ {md_path}", result["unified_diff"])

    def test_bad_format_and_max_hunks_are_invalid_input(self):
        md_path = self._write_md("x.md", "x")
        for kwargs in ({"format": "both"}, {"max_hunks": 0}):
            with self.assertRaises(VerifyError) as ctx:
                server.execute_diff_body_vs_file(str(self.target), str(md_path), **kwargs)
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class NormalizedDiffTests(DiffBodyVsFileTestCase):
    """#110: smart quotes, dashes, NBSP, NFC and trailing spaces are not edits."""

    def _sections_docx(self) -> tuple[Path, str]:
        from verified_docx_mcp import projection

        docx = Path(self._tmp.name) / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", docx)
        body_md, _, _ = projection.read_document_markdown(docx)
        return docx, body_md

    def test_typographic_differences_are_identical_but_not_identical_raw(self):
        variant = "The quick **brown** fox jumps over the lazy dog."
        md_path = self._write_md("nbsp.md", variant + "  \n")
        result = server.execute_diff_body_vs_file(str(self.target), str(md_path))
        self.assertTrue(result["identical"])
        self.assertFalse(result["identical_raw"])
        self.assertEqual(result["hunks"], [])

        raw = server.execute_diff_body_vs_file(str(self.target), str(md_path), normalize=False)
        self.assertFalse(raw["identical"])
        self.assertFalse(raw["normalized"])
        self.assertTrue(raw["hunks"])

    def test_normalize_for_compare_folds_quotes_dashes_softhyphen_nfc(self):
        from verified_docx_mcp import locate

        folded = locate.normalize_for_compare("“A” – it’s­ ok  \ncafé")
        self.assertEqual(folded, '"A" - it\'s ok\ncafé')

    def test_three_real_edits_give_exactly_three_non_equal_hunks(self):
        docx, body_md = self._sections_docx()
        edited = (
            body_md.replace("Some overview text.", "Some NEW overview text.")
            .replace("Background text.", "Background text, changed.")
            .replace("Next steps text.", "Next steps text!")
        )
        result = server.execute_diff_body_vs_file(str(docx), str(self._write_md("edited.md", edited)))
        changed = [h for h in result["hunks"] if h["tag"] != "equal"]
        self.assertEqual(len(changed), 3)
        self.assertNotIn("unified_diff", result)
        equal = [h for h in result["hunks"] if h["tag"] == "equal"]
        self.assertTrue(equal)
        for hunk in equal:
            self.assertEqual(set(hunk), {"tag", "body_lines_count", "file_lines_count", "body_range", "file_range"})
        self.assertFalse(result["truncated"])

    def test_max_hunks_truncates_and_reports_the_remainder(self):
        docx, body_md = self._sections_docx()
        edited = body_md.replace("text.", "text, edited.")
        result = server.execute_diff_body_vs_file(str(docx), str(self._write_md("edited.md", edited)), max_hunks=1)
        self.assertEqual(len([h for h in result["hunks"] if h["tag"] != "equal"]), 1)
        self.assertTrue(result["truncated"])
        self.assertEqual(result["remaining_hunks"], 2)

    def test_section_key_scopes_both_sides(self):
        docx, body_md = self._sections_docx()
        # Only Background differs; Overview and Next Steps are untouched.
        md_path = self._write_md("edited.md", body_md.replace("Background text.", "Background text, changed."))

        scoped = server.execute_diff_body_vs_file(str(docx), str(md_path), section_key="background-1")
        self.assertFalse(scoped["identical"])
        self.assertEqual(scoped["section_key"], "background-1")
        self.assertEqual(len([h for h in scoped["hunks"] if h["tag"] != "equal"]), 1)
        # The Background heading is file line 5 -> 0-based offset 4.
        self.assertEqual(scoped["file_line_offset"], 4)

        clean = server.execute_diff_body_vs_file(str(docx), str(md_path), section_key="overview-1")
        self.assertTrue(clean["identical"])
        self.assertEqual(clean["hunks"], [])

    def test_appending_a_paragraph_does_not_make_the_unchanged_last_line_an_edit(self):
        """Review finding: the last unchanged line used to come back as 'replace' because only the
        file side kept its line ending; and the unified diff glued two lines together."""
        docx, body_md = self._sections_docx()
        md_path = self._write_md("longer.md", body_md + "\n\nExtra closing paragraph.\n")
        result = server.execute_diff_body_vs_file(str(docx), str(md_path))
        changed = [h for h in result["hunks"] if h["tag"] != "equal"]
        self.assertEqual([h["tag"] for h in changed], ["insert"])
        self.assertEqual(changed[0]["body_lines"], [])
        self.assertIn("Extra closing paragraph.\n", changed[0]["file_lines"])

        unified = server.execute_diff_body_vs_file(str(docx), str(md_path), format="unified")["unified_diff"]
        for line in unified.splitlines():
            self.assertFalse(line.startswith("-") and "+" in line[1:] and "Next steps" in line, line)
        self.assertIn("+Extra closing paragraph.", unified.splitlines())

        edited = self._write_md("edited.md", body_md.replace("Next steps text.", "Next steps text, CHANGED."))
        text = server.execute_diff_body_vs_file(str(docx), str(edited), format="unified")["unified_diff"]
        self.assertIn("-Next steps text.", text.splitlines())
        self.assertIn("+Next steps text, CHANGED.", text.splitlines())

    def test_file_section_slice_heading_levels_and_folded_headings(self):
        # Outline levels 6-8 render as 7-9 '#'; a converted file may have lost a soft hyphen.
        lines = ["# One\n", "text\n", "####### Deep\n", "deep text\n", "# inter­national\n", "more\n"]
        chunk, start = server._file_section_slice(lines, "deep-1")
        self.assertEqual((chunk, start), (["####### Deep\n", "deep text\n"], 2))
        # The docx key was slugged from the raw text, soft hyphen included ("inter-national-1"); the
        # file keeps it. Matching only the folded text used to refuse this.
        self.assertEqual(server._file_section_slice(lines, "inter-national-1")[1], 4)

    def test_section_key_unknown_in_body_or_file(self):
        docx, _ = self._sections_docx()
        md_path = self._write_md("nohead.md", "no headings here\n")
        with self.assertRaises(VerifyError) as ctx:
            server.execute_diff_body_vs_file(str(docx), str(md_path), section_key="nope-1")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.SECTION_NOT_FOUND)
        with self.assertRaises(VerifyError) as ctx:
            server.execute_diff_body_vs_file(str(docx), str(md_path), section_key="overview-1")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


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
