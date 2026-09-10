"""Unit tests for src/verified_docx_mcp/projection.py (issue #28 WP-03).

Covers the WP's four named acceptance tests against the Word-authored
fixtures in tests/fixtures/ (see tests/fixtures/README.md for provenance):

  - the fragmented phrase (frag.docx) is contiguous in the projection
  - field instruction text is absent and result text is present (fields.docx)
  - table-cell text carries its container chain (tables.docx)
  - the revision token: unchanged after open-and-close without edits,
    changed after one typed character, changed after a comment, unchanged
    after `touch` (tests/fixtures/revision/*.docx)

Plus targeted coverage of list_parts, find_sections, list_page_sections,
list_styles, the drawing/table_start/table_end runs-format records, the
w:txbxContent sub-scope projection, and the PART_NOT_FOUND error path.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import projection
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"
REVISION_FIXTURES = FIXTURES / "revision"


class FragmentedPhraseTests(unittest.TestCase):
    """Acceptance: the fragmented phrase is contiguous in the projection."""

    def test_phrase_is_contiguous(self):
        text = projection.read_document_text(FIXTURES / "frag.docx")
        self.assertIn("The quick brown fox jumps over the lazy dog.", text)

    def test_bold_word_is_its_own_run(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        run_events = [e for e in proj.events if isinstance(e, projection.RunEvent)]
        bold_runs = [e for e in run_events if e.rpr.get("bold")]
        self.assertEqual([e.text for e in bold_runs], ["brown"])
        # 3+ runs total (before / bold / after), per the WP's fixture spec.
        self.assertGreaterEqual(len(run_events), 3)


class FieldTests(unittest.TestCase):
    """Acceptance: field instruction text is absent and result text is
    present."""

    def test_instruction_text_absent_result_present(self):
        text = projection.read_document_text(FIXTURES / "fields.docx")
        # Field instruction codes (both the complex PAGE field's instrText
        # and the simple REF field's w:instr attribute) never appear.
        self.assertNotIn("MERGEFORMAT", text)
        self.assertNotIn("REF Target1", text)
        self.assertNotIn(" PAGE ", text)
        # Result text (the PAGE field's rendered page number, and the REF
        # field's resolved cross-reference text) IS present.
        self.assertIn("of the document. Cross-reference:", text)
        self.assertTrue(text.rstrip().endswith("Target paragraph"))

    def test_field_records_in_runs_format(self):
        records = projection.read_document_runs(FIXTURES / "fields.docx")
        fields = [r for r in records if r.get("type") == "field"]
        self.assertEqual(len(fields), 2)
        page_field = next(f for f in fields if f["instr"].startswith("PAGE"))
        ref_field = next(f for f in fields if f["instr"].startswith("REF"))
        self.assertEqual(page_field["result_text"], "1")
        self.assertEqual(ref_field["result_text"], "Target paragraph ")


class TableContainerChainTests(unittest.TestCase):
    """Acceptance: table-cell text carries its container chain."""

    def test_cell_paragraphs_carry_container_chain(self):
        proj = projection.project_part(FIXTURES / "tables.docx")
        meta_by_ref = {m.para_ref: m for m in proj.paragraphs}

        def chain_for(cell_text: str) -> list[dict]:
            run = next(
                e
                for e in proj.events
                if isinstance(e, projection.RunEvent) and e.text == cell_text
            )
            return meta_by_ref[run.para_ref].container_chain

        self.assertEqual(chain_for("R1C1"), [{"table_id": 1, "row": 1, "cell": 1}])
        self.assertEqual(chain_for("R1C2"), [{"table_id": 1, "row": 1, "cell": 2}])
        self.assertEqual(chain_for("R2C1"), [{"table_id": 1, "row": 2, "cell": 1}])
        self.assertEqual(chain_for("R2C2"), [{"table_id": 1, "row": 2, "cell": 2}])

        # Body paragraphs outside the table carry no container chain.
        before = next(
            e for e in proj.events if isinstance(e, projection.RunEvent) and e.text == "Before the table."
        )
        self.assertEqual(meta_by_ref[before.para_ref].container_chain, [])

    def test_table_start_end_records_bracket_cells(self):
        records = projection.read_document_runs(FIXTURES / "tables.docx")
        types = [r.get("type") for r in records if "type" in r]
        self.assertEqual(types, ["table_start", "table_end"])
        start = next(r for r in records if r.get("type") == "table_start")
        end = next(r for r in records if r.get("type") == "table_end")
        self.assertEqual(start["table_id"], 1)
        self.assertEqual(end["table_id"], 1)

        start_idx = records.index(start)
        end_idx = records.index(end)
        cell_texts = {r["text"] for r in records[start_idx + 1 : end_idx] if "text" in r}
        self.assertEqual(cell_texts, {"R1C1", "R1C2", "R2C1", "R2C2"})


class RevisionTokenTests(unittest.TestCase):
    """Acceptance: revision token unchanged after open-and-close without
    edits, changed after one typed character, changed after a comment,
    unchanged after `touch`."""

    def test_unchanged_after_open_and_resave_without_edits(self):
        base = projection.compute_revision(REVISION_FIXTURES / "base.docx")
        reopened = projection.compute_revision(REVISION_FIXTURES / "reopened.docx")
        self.assertEqual(base["token"], reopened["token"])
        self.assertEqual(base["detail"]["document_sha256"], reopened["detail"]["document_sha256"])

    def test_changed_after_one_typed_character(self):
        base = projection.compute_revision(REVISION_FIXTURES / "base.docx")
        typed = projection.compute_revision(REVISION_FIXTURES / "typed.docx")
        self.assertNotEqual(base["token"], typed["token"])
        self.assertNotEqual(base["detail"]["document_sha256"], typed["detail"]["document_sha256"])

    def test_changed_after_a_comment(self):
        base = projection.compute_revision(REVISION_FIXTURES / "base.docx")
        commented = projection.compute_revision(REVISION_FIXTURES / "commented.docx")
        self.assertNotEqual(base["token"], commented["token"])
        # Specifically the comments half changed (new comment parts +
        # rels + [Content_Types].xml), not just the document half.
        self.assertNotEqual(base["detail"]["comments_sha256"], commented["detail"]["comments_sha256"])

    def test_unchanged_after_touch(self):
        with tempfile.TemporaryDirectory() as tmp:
            copy_path = Path(tmp) / "touched.docx"
            shutil.copyfile(REVISION_FIXTURES / "base.docx", copy_path)

            before = projection.compute_revision(copy_path)

            # Change mtime only — content and size are untouched.
            new_time = copy_path.stat().st_mtime + 120
            os.utime(copy_path, (new_time, new_time))

            after = projection.compute_revision(copy_path)

            self.assertEqual(before["token"], after["token"])
            self.assertEqual(before["detail"]["document_sha256"], after["detail"]["document_sha256"])
            self.assertEqual(before["detail"]["comments_sha256"], after["detail"]["comments_sha256"])
            self.assertEqual(before["detail"]["size"], after["detail"]["size"])
            # mtime_ns DID change — proves the test actually touched the
            # file rather than trivially comparing a no-op.
            self.assertNotEqual(before["detail"]["mtime_ns"], after["detail"]["mtime_ns"])

    def test_token_equality_is_hash_only_not_full_tuple(self):
        """A same-content file with a different size/mtime on disk still
        equals the original token (guards against an __eq__ that
        accidentally compares all four detail fields)."""
        with tempfile.TemporaryDirectory() as tmp:
            copy_path = Path(tmp) / "copy.docx"
            shutil.copyfile(REVISION_FIXTURES / "base.docx", copy_path)
            base = projection.compute_revision(REVISION_FIXTURES / "base.docx")
            copy = projection.compute_revision(copy_path)
            self.assertEqual(base["token"], copy["token"])


class DrawingRecordTests(unittest.TestCase):
    """Orchestrator carry-forward: read_document(format="runs") must emit
    the drawing record (real example: textbox.docx's shape)."""

    def test_drawing_record_present(self):
        records = projection.read_document_runs(FIXTURES / "textbox.docx")
        drawings = [r for r in records if r.get("type") == "drawing"]
        self.assertEqual(len(drawings), 1)
        drawing = drawings[0]
        self.assertIn("para_ref", drawing)
        self.assertIn("extent_in", drawing)
        self.assertEqual(len(drawing["extent_in"]), 2)
        # A shape (not a picture) carries no blip — both keys are still
        # present, just None, never silently dropped.
        self.assertIn("blip_rid", drawing)
        self.assertIn("media_part", drawing)

    def test_alternate_content_fallback_not_double_walked(self):
        """mc:Fallback's legacy w:pict/v:textbox carries the SAME text as
        mc:Choice's w:drawing/txbxContent — only one drawing record, and
        the text box's own text (projected as a separate sub-scope) is
        not duplicated in the host part's flat text."""
        records = projection.read_document_runs(FIXTURES / "textbox.docx")
        self.assertEqual(len([r for r in records if r.get("type") == "drawing"]), 1)
        host_text = projection.read_document_text(FIXTURES / "textbox.docx")
        self.assertNotIn("Text inside the text box.", host_text)


class TextboxSubScopeTests(unittest.TestCase):
    def test_textbox_projected_as_separate_scope(self):
        scopes = projection.iter_textbox_scopes(FIXTURES / "textbox.docx")
        self.assertEqual(len(scopes), 1)
        self.assertEqual(scopes[0]["section_key"], "textbox-1")
        self.assertEqual(scopes[0]["text"], "Text inside the text box.")

    def test_document_with_no_textbox_has_no_scopes(self):
        self.assertEqual(projection.iter_textbox_scopes(FIXTURES / "frag.docx"), [])


class ListPartsTests(unittest.TestCase):
    def test_headers_docx_parts(self):
        parts = projection.list_parts_impl(FIXTURES / "headers.docx")
        by_part = {p["part"]: p for p in parts}
        self.assertEqual(by_part["word/document.xml"]["kind"], "document")
        self.assertEqual(by_part["word/header2.xml"]["header_footer_type"], "default")
        self.assertEqual(by_part["word/footer2.xml"]["header_footer_type"], "default")
        self.assertIn("word/footnotes.xml", by_part)
        self.assertIn("word/endnotes.xml", by_part)

    def test_header_and_footer_text_are_independent_scopes(self):
        parts = projection.list_parts_impl(FIXTURES / "headers.docx")
        default_header = next(
            p["part"] for p in parts if p["kind"] == "header" and p["header_footer_type"] == "default"
        )
        default_footer = next(
            p["part"] for p in parts if p["kind"] == "footer" and p["header_footer_type"] == "default"
        )
        header_text = projection.read_document_text(FIXTURES / "headers.docx", default_header)
        footer_text = projection.read_document_text(FIXTURES / "headers.docx", default_footer)
        body_text = projection.read_document_text(FIXTURES / "headers.docx")
        self.assertEqual(header_text, "Header text, distinct from the body.")
        self.assertEqual(footer_text, "Footer text, also distinct.")
        self.assertEqual(body_text, "Main body text of the document.")


class FindSectionsTests(unittest.TestCase):
    def test_heading_ranges(self):
        sections = projection.find_sections_impl(FIXTURES / "sections.docx")
        self.assertEqual([s["heading_text"] for s in sections], ["Overview", "Background", "Next Steps"])
        self.assertEqual([s["outline_level"] for s in sections], [0, 1, 0])
        keys = [s["section_key"] for s in sections]
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(keys, ["overview-1", "background-1", "next-steps-1"])

    def test_document_with_no_headings_has_no_sections(self):
        self.assertEqual(projection.find_sections_impl(FIXTURES / "frag.docx"), [])

    def test_slugify(self):
        self.assertEqual(projection._slugify("Overview"), "overview")
        self.assertEqual(projection._slugify("Section 3: Next Steps!"), "section-3-next-steps")


class ListPageSectionsTests(unittest.TestCase):
    def test_default_letter_page_geometry_in_inches(self):
        sections = projection.list_page_sections_impl(FIXTURES / "tables.docx")
        self.assertEqual(len(sections), 1)
        section = sections[0]
        self.assertAlmostEqual(section["page_width_in"], 8.5)
        self.assertAlmostEqual(section["page_height_in"], 11.0)
        self.assertAlmostEqual(section["margin_top_in"], 1.0)
        self.assertAlmostEqual(section["margin_bottom_in"], 1.0)
        self.assertAlmostEqual(section["margin_left_in"], 1.0)
        self.assertAlmostEqual(section["margin_right_in"], 1.0)


class ListStylesTests(unittest.TestCase):
    def test_heading_styles_carry_outline_level(self):
        styles = projection.list_styles_impl(FIXTURES / "sections.docx")
        by_id = {s["style_id"]: s for s in styles}
        self.assertEqual(by_id["Heading1"]["outline_lvl"], 0)
        self.assertEqual(by_id["Heading2"]["outline_lvl"], 1)
        self.assertEqual(by_id["Normal"]["type"], "paragraph")


class PartNotFoundTests(unittest.TestCase):
    def test_read_document_text_raises_part_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            projection.project_part(FIXTURES / "frag.docx", "word/does-not-exist.xml")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.PART_NOT_FOUND)

    def test_list_page_sections_raises_part_not_found(self):
        with self.assertRaises(VerifyError) as ctx:
            projection.list_page_sections_impl(FIXTURES / "frag.docx", "word/does-not-exist.xml")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.PART_NOT_FOUND)


class MarkdownRenderTests(unittest.TestCase):
    def test_headings_rendered_as_markdown_headers(self):
        markdown, warnings = projection.read_document_markdown(FIXTURES / "sections.docx")
        self.assertIn("# Overview", markdown)
        self.assertIn("## Background", markdown)
        self.assertEqual(warnings, [])

    def test_bold_run_rendered(self):
        markdown, _ = projection.read_document_markdown(FIXTURES / "frag.docx")
        self.assertIn("**brown**", markdown)

    def test_table_and_graphic_placeholders(self):
        table_md, _ = projection.read_document_markdown(FIXTURES / "tables.docx")
        self.assertIn("[TABLE]", table_md)
        graphic_md, _ = projection.read_document_markdown(FIXTURES / "textbox.docx")
        self.assertIn("[GRAPHIC]", graphic_md)


if __name__ == "__main__":
    unittest.main()
