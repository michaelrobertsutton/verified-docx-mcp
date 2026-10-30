"""Issue #110: ranged read_document (start_para_ref/end_para_ref), to_file,
and the chars / sha12 additions to find_sections."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import paths, projection, server
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"


class RangedReadTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.sections = Path(self._tmp.name) / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", self.sections)
        self.tables = Path(self._tmp.name) / "tables-merged.docx"
        shutil.copyfile(FIXTURES / "tables-merged.docx", self.tables)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def read(self, target: Path, **kwargs):
        return server.execute_read_document(str(target), **kwargs)


class RangeTests(RangedReadTestCase):
    def test_text_range_returns_only_those_paragraphs(self):
        result = self.read(self.sections, format="text", start_para_ref="p2", end_para_ref="p3")
        self.assertEqual(result["text"], "Background\nBackground text.")
        self.assertEqual((result["start_para_ref"], result["end_para_ref"]), ("p2", "p3"))

    def test_markdown_range(self):
        result = self.read(self.sections, format="markdown", start_para_ref="p2", end_para_ref="p3")
        self.assertEqual(result["markdown"], "## Background\n\nBackground text.")
        self.assertNotIn("lossy_elements", result)

    def test_runs_range_only_has_in_range_paragraphs(self):
        result = self.read(self.sections, format="runs", start_para_ref="p4", end_para_ref="p5")
        refs = {r["para_ref"] for r in result["runs"] if "para_ref" in r}
        self.assertEqual(refs, {"p4", "p5"})

    def test_single_paragraph_range(self):
        result = self.read(self.sections, format="text", start_para_ref="p1", end_para_ref="p1")
        self.assertEqual(result["text"], "Some overview text.")

    def test_range_matches_find_sections_section(self):
        sec = next(s for s in projection.find_sections_impl(self.sections) if s["section_key"] == "next-steps-1")
        result = self.read(
            self.sections, format="text", start_para_ref=sec["start_para_ref"], end_para_ref=sec["end_para_ref"]
        )
        self.assertEqual(len(result["text"]), sec["chars"])

    def test_bad_ranges_are_invalid_input(self):
        for kwargs in (
            {"start_para_ref": "p3", "end_para_ref": "p1"},  # inverted
            {"start_para_ref": "p0", "end_para_ref": "p99"},  # out of range
            {"start_para_ref": "p1"},  # only one bound
            {"start_para_ref": "x1", "end_para_ref": "p2"},  # malformed
            {"start_para_ref": "p0", "end_para_ref": "p1", "section_key": "textbox-1"},
        ):
            with self.assertRaises(VerifyError, msg=kwargs) as ctx:
                self.read(self.sections, format="text", **kwargs)
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_range_cutting_a_table_is_marked_partial(self):
        # p5/p6 are row 3 of table 1 only (see the fixture's container chains).
        result = self.read(self.tables, format="markdown", start_para_ref="p5", end_para_ref="p6")
        self.assertIn({"kind": "partial_table", "table_id": 1}, result["lossy_elements"])
        self.assertNotIn("Merged down", result["markdown"])
        self.assertIn("B3", result["markdown"])

    def test_range_covering_a_whole_table_is_not_partial(self):
        result = self.read(self.tables, format="markdown", start_para_ref="p0", end_para_ref="p10")
        kinds = [e["kind"] for e in result.get("lossy_elements", [])]
        self.assertNotIn("partial_table", kinds)
        whole = self.read(self.tables, format="markdown")
        self.assertEqual(result["markdown"], whole["markdown"])

    def test_full_range_equals_whole_read_in_every_format(self):
        for fmt in ("text", "markdown", "runs"):
            whole = self.read(self.sections, format=fmt)
            ranged = self.read(self.sections, format=fmt, start_para_ref="p0", end_para_ref="p5")
            self.assertEqual(whole[fmt], ranged[fmt], fmt)


class ToFileTests(RangedReadTestCase):
    def test_to_file_returns_summary_without_body_and_matching_sha(self):
        out = Path(self._tmp.name) / "out.md"
        result = self.read(self.sections, format="markdown", to_file=str(out))
        self.assertNotIn("markdown", result)
        self.assertNotIn("text", result)
        written = out.read_bytes()
        self.assertEqual(result["sha256"], hashlib.sha256(written).hexdigest())
        self.assertEqual(result["chars"], len(written.decode("utf-8")))
        self.assertEqual(result["written_to"], str(out.resolve()))
        self.assertEqual(result["paragraph_count"], 6)
        # Compact rows: [section_key, start_para_ref, end_para_ref, chars].
        self.assertEqual(
            [row[0] for row in result["sections"]], ["overview-1", "background-1", "next-steps-1"]
        )
        self.assertEqual(result["sections"][0][1:3], ["p0", "p1"])
        self.assertEqual(result["section_count"], 3)
        self.assertNotIn("sections_truncated", result)
        self.assertIn("revision", result)
        self.assertLess(len(json.dumps(result)), 1500)

    def test_to_file_caps_the_section_list_and_says_so(self):
        from unittest import mock

        out = Path(self._tmp.name) / "capped.md"
        with mock.patch.object(server, "_TO_FILE_MAX_SECTIONS", 1):
            result = self.read(self.sections, format="markdown", to_file=str(out))
        self.assertEqual(len(result["sections"]), 1)
        self.assertEqual(result["section_count"], 3)
        self.assertTrue(result["sections_truncated"])

    def test_to_file_runs_is_json(self):
        out = Path(self._tmp.name) / "runs.json"
        result = self.read(self.sections, format="runs", to_file=str(out))
        self.assertTrue(json.loads(out.read_text(encoding="utf-8")))
        self.assertNotIn("runs", result)

    def test_to_file_with_range_writes_only_the_range(self):
        out = Path(self._tmp.name) / "range.txt"
        result = self.read(
            self.sections, format="text", start_para_ref="p4", end_para_ref="p5", to_file=str(out)
        )
        self.assertEqual(out.read_text(encoding="utf-8"), "Next Steps\nNext steps text.")
        self.assertEqual(result["paragraph_count"], 2)
        self.assertEqual(result["start_para_ref"], "p4")

    def test_to_file_outside_allowed_roots_is_refused_before_reading(self):
        with tempfile.TemporaryDirectory() as outside:
            with self.assertRaises(VerifyError) as ctx:
                self.read(self.sections, format="text", to_file=str(Path(outside) / "x.txt"))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
            self.assertFalse((Path(outside) / "x.txt").exists())

    def test_to_file_missing_parent_and_directory_target_refused(self):
        for target in (Path(self._tmp.name) / "nope" / "x.txt", Path(self._tmp.name)):
            with self.assertRaises(VerifyError, msg=str(target)) as ctx:
                self.read(self.sections, format="text", to_file=str(target))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_to_file_never_overwrites_an_existing_file_unless_asked(self):
        out = Path(self._tmp.name) / "existing.md"
        out.write_text("precious", encoding="utf-8")
        with self.assertRaises(VerifyError) as ctx:
            self.read(self.sections, format="text", to_file=str(out))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(out.read_text(encoding="utf-8"), "precious")
        result = self.read(self.sections, format="text", to_file=str(out), overwrite=True)
        self.assertEqual(result["written_to"], str(out.resolve()))
        self.assertNotEqual(out.read_text(encoding="utf-8"), "precious")

    def test_to_file_cannot_replace_the_document_being_read_or_any_non_text_file(self):
        """Review finding: to_file=<the .docx itself> used to destroy it."""
        before = self.sections.read_bytes()
        for target in (self.sections, Path(self._tmp.name) / ".zshrc", Path(self._tmp.name) / "notes"):
            for overwrite in (False, True):
                with self.assertRaises(VerifyError, msg=f"{target.name} overwrite={overwrite}") as ctx:
                    self.read(self.sections, format="text", to_file=str(target), overwrite=overwrite)
                self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(self.sections.read_bytes(), before)

    def test_to_file_leaves_no_temp_files_behind(self):
        out = Path(self._tmp.name) / "clean.txt"
        self.read(self.sections, format="text", to_file=str(out))
        leftovers = [p.name for p in Path(self._tmp.name).iterdir() if p.name.startswith(".clean.txt.")]
        self.assertEqual(leftovers, [])


class FindSectionsSizeAndDigestTests(RangedReadTestCase):
    def test_every_entry_has_chars_and_digest_is_opt_in(self):
        plain = server.execute_find_sections(str(self.sections))["sections"]
        self.assertTrue(all("chars" in s for s in plain))
        self.assertTrue(all("sha12" not in s for s in plain))
        digested = server.execute_find_sections(str(self.sections), digests=True)["sections"]
        self.assertTrue(all(len(s["sha12"]) == 12 for s in digested))

    def test_sha12_is_the_pane_definition(self):
        digested = {s["section_key"]: s for s in server.execute_find_sections(str(self.sections), digests=True)["sections"]}
        expected = hashlib.sha256(b"Background\nBackground text.").hexdigest()[:12]
        self.assertEqual(digested["background-1"]["sha12"], expected)
        self.assertEqual(digested["background-1"]["chars"], len("Background\nBackground text."))

    def test_live_body_trailing_empty_paragraph_is_ignored_for_digests(self):
        """Real Word (#110): Body.getOoxml() appends one empty paragraph Body.paragraphs lacks,
        so a live body projection's last section must be digested without it."""
        import zipfile

        padded = Path(self._tmp.name) / "padded.docx"
        with zipfile.ZipFile(self.sections) as src, zipfile.ZipFile(padded, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "word/document.xml":
                    xml = data.decode("utf-8")
                    marker = "<w:sectPr"
                    self.assertIn(marker, xml)
                    xml = xml.replace(marker, "<w:p/>" + marker, 1)
                    data = xml.encode("utf-8")
                dst.writestr(item, data)
        plain = {s["section_key"]: s["sha12"] for s in projection.find_sections_impl(self.sections, digests=True)}
        raw = {s["section_key"]: s["sha12"] for s in projection.find_sections_impl(padded, digests=True)}
        trimmed = {
            s["section_key"]: s["sha12"]
            for s in projection.find_sections_impl(padded, digests=True, ignore_trailing_empty=True)
        }
        self.assertNotEqual(raw["next-steps-1"], plain["next-steps-1"])  # the extra paragraph shows
        self.assertEqual(trimmed, plain)  # and is ignored when asked

    def test_a_soft_line_break_is_digested_as_words_vertical_tab(self):
        """Real Word (#110 review): paragraph.text reports a w:br as U+000B, and the digest
        must hash what Word reports. Verified against real Word on 2026-10-29."""
        doc = Path(self._tmp.name) / "br.docx"
        shutil.copyfile(self.sections, doc)
        from verified_docx_mcp import mutations

        mutations._QUIESCE_INTERVAL_SECONDS = 0.02
        mutations.execute_replace_body_markdown(str(doc), "# Alpha\n\nLine one  \nLine two\n", force=True)
        texts = projection.paragraph_texts(projection.project_part(doc))
        self.assertEqual(texts, ["Alpha", "Line one\u000bLine two"])
        digest = projection.find_sections_impl(doc, digests=True)[0]["sha12"]
        self.assertEqual(digest, hashlib.sha256("Alpha\nLine one\u000bLine two".encode()).hexdigest()[:12])

    def test_textbox_entries_carry_chars(self):
        box = Path(self._tmp.name) / "textbox.docx"
        shutil.copyfile(FIXTURES / "textbox.docx", box)
        entries = server.execute_find_sections(str(box))["sections"]
        boxes = [s for s in entries if s["kind"] == "textbox"]
        self.assertTrue(boxes)
        self.assertTrue(all(isinstance(s["chars"], int) for s in boxes))


if __name__ == "__main__":
    unittest.main()


class HeadingSectionReadTests(RangedReadTestCase):
    def _sections(self):
        return [s for s in projection.find_sections_impl(self.sections) if s["kind"] == "heading"]

    def test_heading_read_equals_the_range_read_in_every_format(self):
        for sec in self._sections():
            for fmt in ("text", "markdown", "runs"):
                by_key = self.read(self.sections, format=fmt, section_key=sec["section_key"])
                by_range = self.read(
                    self.sections, format=fmt, start_para_ref=sec["start_para_ref"], end_para_ref=sec["end_para_ref"]
                )
                self.assertEqual(by_key[fmt], by_range[fmt], (sec["section_key"], fmt))
                self.assertEqual(by_key["section_key"], sec["section_key"])
                self.assertEqual(by_key["start_para_ref"], sec["start_para_ref"])
                self.assertEqual(by_key["end_para_ref"], sec["end_para_ref"])

    def test_last_section_runs_to_the_end_of_the_body(self):
        sec = self._sections()[-1]
        result = self.read(self.sections, format="text", section_key=sec["section_key"])
        self.assertEqual(result["text"], "Next Steps\nNext steps text.")
        whole = self.read(self.sections, format="text")
        self.assertTrue(whole["text"].endswith(result["text"]))

    def test_original_para_refs_are_kept(self):
        result = self.read(self.sections, format="runs", section_key="next-steps-1")
        refs = {r["para_ref"] for r in result["runs"] if "para_ref" in r}
        self.assertEqual(refs, {"p4", "p5"})

    def test_unknown_key_lists_heading_and_textbox_keys(self):
        with self.assertRaises(VerifyError) as ctx:
            self.read(self.sections, format="text", section_key="nope-1")
        env = ctx.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("next-steps-1", env.diagnostics["available_heading_keys"])
        self.assertEqual(env.diagnostics["available_textbox_keys"], [])

    def test_textbox_key_still_reads_the_textbox(self):
        box = Path(self._tmp.name) / "textbox.docx"
        shutil.copyfile(FIXTURES / "textbox.docx", box)
        result = self.read(box, format="text", section_key="textbox-1")
        self.assertIn("Text inside the text box.", result["text"])
        self.assertNotIn("start_para_ref", result)

    def test_heading_key_with_a_range_is_invalid_input(self):
        with self.assertRaises(VerifyError) as ctx:
            self.read(self.sections, format="text", section_key="next-steps-1", start_para_ref="p0", end_para_ref="p1")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_to_file_writes_only_the_section(self):
        out = Path(self._tmp.name) / "section.txt"
        result = self.read(self.sections, format="text", section_key="next-steps-1", to_file=str(out))
        self.assertEqual(out.read_text(encoding="utf-8"), "Next Steps\nNext steps text.")
        self.assertEqual(result["section_key"], "next-steps-1")
        self.assertEqual((result["start_para_ref"], result["end_para_ref"]), ("p4", "p5"))
        self.assertNotIn("text", result)

    def test_section_with_a_table_matches_the_range_read(self):
        sections = [s for s in projection.find_sections_impl(self.tables) if s["kind"] == "heading"]
        if not sections:
            self.skipTest("fixture has no headings")
        sec = sections[0]
        by_key = self.read(self.tables, format="markdown", section_key=sec["section_key"])
        by_range = self.read(
            self.tables, format="markdown", start_para_ref=sec["start_para_ref"], end_para_ref=sec["end_para_ref"]
        )
        self.assertEqual(by_key["markdown"], by_range["markdown"])
