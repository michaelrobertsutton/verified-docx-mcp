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
import zipfile
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
    """Text-box content is reachable: iter_textbox_scopes / find_sections
    discover the textbox-<n> section_key, and project_textbox_scope (the
    function server.py's read_document(section_key=...) calls) actually
    returns that scope's content — not just dead internal machinery."""

    def test_textbox_projected_as_separate_scope(self):
        scopes = projection.iter_textbox_scopes(FIXTURES / "textbox.docx")
        self.assertEqual(len(scopes), 1)
        self.assertEqual(scopes[0]["section_key"], "textbox-1")
        self.assertEqual(scopes[0]["text"], "Text inside the text box.")

    def test_document_with_no_textbox_has_no_scopes(self):
        self.assertEqual(projection.iter_textbox_scopes(FIXTURES / "frag.docx"), [])

    def test_find_sections_lists_textbox_scope(self):
        sections = projection.find_sections_impl(FIXTURES / "textbox.docx")
        textbox_entries = [s for s in sections if s["kind"] == "textbox"]
        self.assertEqual(len(textbox_entries), 1)
        entry = textbox_entries[0]
        self.assertEqual(entry["section_key"], "textbox-1")
        self.assertIsNone(entry["heading_text"])
        self.assertIsNone(entry["outline_level"])
        self.assertEqual(entry["paragraph_count"], 1)

    def test_heading_sections_are_tagged_kind_heading(self):
        sections = projection.find_sections_impl(FIXTURES / "sections.docx")
        self.assertTrue(all(s["kind"] == "heading" for s in sections))

    def test_project_textbox_scope_returns_full_projection(self):
        proj = projection.project_textbox_scope(FIXTURES / "textbox.docx", projection.DEFAULT_PART, "textbox-1")
        self.assertIsNotNone(proj)
        self.assertEqual(proj.text, "Text inside the text box.")
        self.assertEqual(len(proj.paragraphs), 1)

    def test_project_textbox_scope_unknown_key_returns_none(self):
        result = projection.project_textbox_scope(FIXTURES / "textbox.docx", projection.DEFAULT_PART, "textbox-99")
        self.assertIsNone(result)

    def test_runs_and_markdown_from_a_textbox_projection(self):
        proj = projection.project_textbox_scope(FIXTURES / "textbox.docx", projection.DEFAULT_PART, "textbox-1")
        runs = projection.runs_from_projection(proj)
        self.assertTrue(any(r.get("text") == "Text inside the text box." for r in runs))
        markdown, _, _ = projection.markdown_from_projection(FIXTURES / "textbox.docx", proj)
        self.assertIn("Text inside the text box.", markdown)


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

    def test_default_single_column_reports_real_width(self):
        """sections.docx's real <w:cols w:space="720"/> carries no w:num
        and no explicit w:col children (Word's own default single-column
        form) — this must report ONE column of page width minus both side
        margins, not an empty list."""
        with zipfile.ZipFile(FIXTURES / "sections.docx") as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('<w:cols w:space="720"/>', document_xml)
        self.assertNotIn("w:num", document_xml)

        sections = projection.list_page_sections_impl(FIXTURES / "sections.docx")
        self.assertEqual(len(sections), 1)
        widths = sections[0]["column_widths_in"]
        self.assertEqual(len(widths), 1)
        self.assertAlmostEqual(widths[0], 6.5)  # 8.5in page - 1in - 1in margins


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
        markdown, warnings, lossy = projection.read_document_markdown(FIXTURES / "sections.docx")
        self.assertIn("# Overview", markdown)
        self.assertIn("## Background", markdown)
        self.assertEqual(warnings, [])
        self.assertEqual(lossy, [])

    def test_bold_run_rendered(self):
        markdown, _, _ = projection.read_document_markdown(FIXTURES / "frag.docx")
        self.assertIn("**brown**", markdown)

    def test_table_renders_as_gfm_pipe_table_not_a_placeholder(self):
        """issue #28 WP-03b-a: tables.docx's 2x2 table (R1C1..R2C2, no
        merges) renders as a real pipe table — first row header, "---"
        separator, no [TABLE] placeholder anywhere — matching
        GoogleDocs-MCP markdown.py's _table/_pipe_row conventions. No
        merge/nesting in this fixture, so lossy_elements is empty."""
        table_md, _, lossy = projection.read_document_markdown(FIXTURES / "tables.docx")
        self.assertNotIn("[TABLE]", table_md)
        self.assertIn("| R1C1 | R1C2 |", table_md)
        self.assertIn("| --- | --- |", table_md)
        self.assertIn("| R2C1 | R2C2 |", table_md)
        self.assertIn("Before the table.", table_md)
        self.assertIn("After the table.", table_md)
        self.assertEqual(lossy, [])

    def test_merged_and_nested_table_reports_lossy_elements(self):
        """tables-merged.docx (derived from the Word-authored tables.docx —
        see fixtures/README.md) exercises the three OOXML table constructs
        a GFM pipe table cannot represent: a w:vMerge pair (a "restart"
        cell keeps its real text, its "continue" cell renders empty), a
        w:gridSpan cell (its text once, then an empty filler cell), and a
        nested w:tbl (flattened to inline text). Each is recorded in
        lossy_elements rather than silently reproduced as an ordinary grid
        cell."""
        md, _, lossy = projection.read_document_markdown(FIXTURES / "tables-merged.docx")
        self.assertNotIn("[TABLE]", md)
        self.assertIn("| H1 | H2 |", md)
        self.assertIn("| Merged down | B2 |", md)
        # vMerge continuation cell: empty, not a copy of "Merged down".
        self.assertIn("|  | B3 Nested A Nested B |", md)
        # Nested table flattened inline, not a placeholder or silent drop.
        self.assertIn("Nested A", md)
        self.assertIn("Nested B", md)
        # gridSpan=2 cell: its text once, filled out with an empty cell.
        self.assertIn("| Spanned row |  |", md)

        kinds = [entry["kind"] for entry in lossy]
        self.assertEqual(kinds.count("table_merge"), 3)  # vMerge restart + continue + gridSpan
        self.assertEqual(kinds.count("nested_table"), 1)
        self.assertTrue(all(entry["table_id"] == 1 for entry in lossy))

    def test_drawing_renders_as_image_placeholder_not_graphic(self):
        """textbox.docx's shape (not a picture) carries no blip_rid, so
        the placeholder falls back to "[image:unknown]" — still an
        [image:...] token, never the old [GRAPHIC]."""
        graphic_md, _, _ = projection.read_document_markdown(FIXTURES / "textbox.docx")
        self.assertNotIn("[GRAPHIC]", graphic_md)
        self.assertIn("[image:", graphic_md)

    def test_field_result_not_doubled_and_instr_not_leaked(self):
        """fields.docx has both a complex PAGE field and a simple REF
        cross-reference field. The result text (already carried into the
        markdown as ordinary runs) must appear exactly once each, and
        neither field's instruction code may leak into the rendering —
        in any casing."""
        markdown, _, lossy = projection.read_document_markdown(FIXTURES / "fields.docx")
        self.assertNotIn("[FIELD:", markdown)
        self.assertNotIn("[field:", markdown)
        self.assertNotIn("MERGEFORMAT", markdown)
        self.assertNotIn("REF Target1", markdown)
        # The result text appears exactly once each — not doubled by both
        # the ordinary run stream AND a re-appended FieldEvent.
        self.assertEqual(markdown.count("1 of the document"), 1)
        self.assertEqual(markdown.count("Target paragraph"), 2)  # heading + resolved ref, not 3
        # Same paragraph text as read_document(format="text") modulo the
        # blank-line block separation WP-03b-a added to match
        # GoogleDocs-MCP's tight-list/loose-block convention (fields.docx
        # has no lists/tables, so only that separator differs).
        flat_text = projection.read_document_text(FIXTURES / "fields.docx")
        self.assertEqual(markdown.replace("\n\n", "\n"), flat_text)
        self.assertEqual(lossy, [])

    def test_field_with_no_result_gets_lowercase_placeholder(self):
        """A field with an instr but a genuinely empty result_text (never
        produced by a real Word-authored fixture — Word always resolves a
        field's result before saving) still renders SOMETHING, not
        silence. Exercises the real markdown_from_projection against a
        synthetic single-event Projection (the only way to reach this
        defensive branch at all); docx_path is only needed for a style
        lookup that this synthetic paragraph's style_id=None makes moot,
        so any real .docx satisfies it honestly."""
        meta = projection.ParagraphMeta(para_ref="p0", style_id=None, outline_lvl=None, container_chain=[])
        leading_run = projection.RunEvent(text="Author: ", rpr={}, para_ref="p0", run_ref="p0/r0")
        field_event = projection.FieldEvent(instr="AUTHOR", result_text="", para_ref="p0")
        synthetic = projection.Projection(
            part="synthetic-for-this-test-only",
            text="Author: ",
            offset_map=[(0, 8, "p0", "p0/r0")],
            events=[leading_run, field_event],
            paragraphs=[meta],
            deleted_spans=[],
            warnings=[],
        )
        markdown, warnings, lossy = projection.markdown_from_projection(FIXTURES / "frag.docx", synthetic)
        self.assertEqual(markdown, "Author: [field:AUTHOR]")
        self.assertEqual(warnings, [])
        self.assertEqual(lossy, [])


if __name__ == "__main__":
    unittest.main()
