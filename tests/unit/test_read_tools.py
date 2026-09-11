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

    def test_section_key_echoed_none_by_default(self):
        result = server.execute_read_document(str(self.target), format="text")
        self.assertIsNone(result["section_key"])

    def test_field_markdown_has_no_leaked_instr_or_doubled_result(self):
        """The exact regression a probe review caught: markdown for
        fields.docx must not leak "[FIELD:...]"/"[field:...]" or double
        the field's already-rendered result text."""
        target = Path(self._tmp.name) / "fields.docx"
        shutil.copyfile(FIXTURES / "fields.docx", target)
        result = server.execute_read_document(str(target), format="markdown")
        self.assertNotIn("[FIELD:", result["markdown"])
        self.assertNotIn("[field:", result["markdown"])
        self.assertNotIn("MERGEFORMAT", result["markdown"])
        self.assertEqual(result["markdown"].count("Target paragraph"), 2)


class TextboxReachabilityToolTests(ReadToolsTestCase):
    """A probe review found text-box content unreachable through any tool
    call (iter_textbox_scopes existed but nothing in server.py called it).
    These exercise the actual fix: read_document(section_key=...)."""

    def setUp(self):
        super().setUp()
        self.textbox_target = Path(self._tmp.name) / "textbox.docx"
        shutil.copyfile(FIXTURES / "textbox.docx", self.textbox_target)

    def test_find_sections_lists_the_textbox_key(self):
        result = server.execute_find_sections(str(self.textbox_target))
        textbox_entries = [s for s in result["sections"] if s["kind"] == "textbox"]
        self.assertEqual([e["section_key"] for e in textbox_entries], ["textbox-1"])

    def test_read_document_section_key_returns_textbox_text(self):
        result = server.execute_read_document(str(self.textbox_target), format="text", section_key="textbox-1")
        self.assertEqual(result["text"], "Text inside the text box.")
        self.assertEqual(result["section_key"], "textbox-1")

    def test_read_document_section_key_runs_and_markdown(self):
        runs_result = server.execute_read_document(str(self.textbox_target), format="runs", section_key="textbox-1")
        self.assertTrue(any(r.get("text") == "Text inside the text box." for r in runs_result["runs"]))

        md_result = server.execute_read_document(
            str(self.textbox_target), format="markdown", section_key="textbox-1"
        )
        self.assertIn("Text inside the text box.", md_result["markdown"])

    def test_host_part_excludes_textbox_text(self):
        """The text box's content is reachable via section_key, but NOT
        double-counted into the host part's own read (no section_key)."""
        result = server.execute_read_document(str(self.textbox_target), format="text")
        self.assertNotIn("Text inside the text box.", result["text"])

    def test_unknown_section_key_raises_invalid_input_with_available_keys(self):
        with self.assertRaises(VerifyError) as ctx:
            server.execute_read_document(str(self.textbox_target), format="text", section_key="textbox-99")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(ctx.exception.envelope.diagnostics["available_textbox_keys"], ["textbox-1"])

    def test_section_key_on_document_with_no_textbox_raises_invalid_input(self):
        with self.assertRaises(VerifyError) as ctx:
            server.execute_read_document(str(self.target), format="text", section_key="textbox-1")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(ctx.exception.envelope.diagnostics["available_textbox_keys"], [])


class FindSectionsToolTests(ReadToolsTestCase):
    def test_no_headings_in_frag_docx(self):
        result = server.execute_find_sections(str(self.target))
        self.assertEqual(result["sections"], [])

    def test_headings_in_sections_docx(self):
        target = Path(self._tmp.name) / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", target)
        result = server.execute_find_sections(str(target))
        self.assertEqual(len(result["sections"]), 3)
        self.assertTrue(all(s["kind"] == "heading" for s in result["sections"]))


class ListPageSectionsToolTests(ReadToolsTestCase):
    def test_returns_page_geometry(self):
        result = server.execute_list_page_sections(str(self.target))
        self.assertEqual(len(result["page_sections"]), 1)
        self.assertAlmostEqual(result["page_sections"][0]["page_width_in"], 8.5)

    def test_default_single_column_reports_real_width(self):
        target = Path(self._tmp.name) / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", target)
        result = server.execute_list_page_sections(str(target))
        widths = result["page_sections"][0]["column_widths_in"]
        self.assertEqual(len(widths), 1)
        self.assertAlmostEqual(widths[0], 6.5)


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


class ReadHeaderFooterToolTests(ReadToolsTestCase):
    """issue #28 WP-15a: read_header_footer reads every header/footer part
    as markdown in one call -- headers.docx is Word-authored (see
    tests/fixtures/README.md), with three headers and three footers, each
    a distinct header_footer_type ("default"/"first"/"even")."""

    def setUp(self):
        super().setUp()
        self.target = Path(self._tmp.name) / "headers.docx"
        shutil.copyfile(FIXTURES / "headers.docx", self.target)

    def test_reads_every_header_and_footer_part(self):
        result = server.execute_read_header_footer(str(self.target))
        entries = result["headers_and_footers"]
        kinds = {e["kind"] for e in entries}
        self.assertEqual(kinds, {"header", "footer"})
        self.assertEqual(len({e["part"] for e in entries}), len(entries))  # each part reported once
        types_by_kind: dict[str, set] = {"header": set(), "footer": set()}
        for e in entries:
            types_by_kind[e["kind"]].add(e["header_footer_type"])
        self.assertEqual(types_by_kind["header"], {"default", "first", "even"})
        self.assertEqual(types_by_kind["footer"], {"default", "first", "even"})

    def test_default_header_and_footer_content(self):
        result = server.execute_read_header_footer(str(self.target))
        default_header = next(
            e for e in result["headers_and_footers"] if e["kind"] == "header" and e["header_footer_type"] == "default"
        )
        default_footer = next(
            e for e in result["headers_and_footers"] if e["kind"] == "footer" and e["header_footer_type"] == "default"
        )
        self.assertIn("Header text, distinct from the body.", default_header["markdown"])
        self.assertIn("Footer text, also distinct.", default_footer["markdown"])

    def test_document_with_no_headers_or_footers_returns_empty_list(self):
        no_hf_target = Path(self._tmp.name) / "frag.docx"
        shutil.copyfile(FIXTURES / "frag.docx", no_hf_target)
        result = server.execute_read_header_footer(str(no_hf_target))
        self.assertEqual(result["headers_and_footers"], [])

    def test_not_gated_by_docx_locked(self):
        owner_file = self.target.with_name("~$" + self.target.name)
        owner_file.write_bytes(b"\x00Someone\x00")
        try:
            with mock.patch.object(paths, "snapshot_docx_package", wraps=paths.snapshot_docx_package) as spy:
                result = server.execute_read_header_footer(str(self.target))
            spy.assert_called_once()
            self.assertTrue(result["headers_and_footers"])
        finally:
            owner_file.unlink(missing_ok=True)


class ReadHeaderFooterNotMutatingTests(unittest.TestCase):
    def test_not_registered(self):
        from verified_docx_mcp.middleware import MUTATING_TOOLS

        self.assertNotIn("read_header_footer", MUTATING_TOOLS)


if __name__ == "__main__":
    unittest.main()
