"""Unit tests for src/verified_docx_mcp/geometry.py (issue #102).

Pure-function coverage, no Word and no MCP server involved:

  - ordinal_of()/build_probe_ordinals(): the para_ref -> Word paragraph
    ordinal mapping, cross-checked against a direct count of <w:p>
    elements in a fixture's word/document.xml (independent of
    projection.py's own machinery, so a projection.py bug in that count
    would not silently validate itself).
  - assemble_sections(): the page-span math (two sections on one page,
    one section spanning two pages, the last-section end approximation
    including its +0.03/min(1.0, ...) cap, and rounding), and the
    verification-failure return contract ((None, "<detail>"), never a
    raise -- see the module docstring for why).
  - normalize_probe_text(): exercised indirectly through the
    verification-mismatch cases above, and directly for the curly-quote/
    soft-hyphen/control-character/whitespace rules its own docstring
    documents.
"""

from __future__ import annotations

import re
import sys
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import geometry, projection

FIXTURE_DIR = REPO / "tests" / "fixtures"

_WP_OPEN_RE = re.compile(r"<w:p[ >]")


def _count_w_p_elements(docx_path: Path) -> int:
    """A direct count of <w:p ...> / <w:p> elements in word/document.xml,
    computed independently of projection.py -- this is the ground truth
    ordinal_of()/build_probe_ordinals() are checked against, not a second
    call into the same machinery under test."""
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    return len(_WP_OPEN_RE.findall(xml))


# ---------------------------------------------------------------------------
# Ordinal mapping (spec test 1)
# ---------------------------------------------------------------------------


class OrdinalMappingTests(unittest.TestCase):
    def test_sections_docx_headings_map_to_direct_w_p_count(self):
        docx_path = FIXTURE_DIR / "sections.docx"
        total_paragraphs = _count_w_p_elements(docx_path)
        self.assertEqual(total_paragraphs, 6)  # 3 headings x 2 paragraphs each

        proj = projection.project_part(docx_path)
        all_para_refs = [m.para_ref for m in proj.paragraphs]
        self.assertEqual(len(all_para_refs), total_paragraphs)

        sections = [s for s in projection.find_sections_impl(docx_path) if s["kind"] == "heading"]
        self.assertEqual([s["heading_text"] for s in sections], ["Overview", "Background", "Next Steps"])

        # Overview is the 1st paragraph, Background the 3rd, Next Steps
        # the 5th -- each heading is followed by exactly one body
        # paragraph before the next heading.
        expected_start_ordinals = [1, 3, 5]
        for section, expected in zip(sections, expected_start_ordinals):
            self.assertEqual(
                geometry.ordinal_of(section["start_para_ref"], all_para_refs), expected
            )

        # The document's last paragraph (Next Steps' own end_para_ref) is
        # the 6th and final <w:p> in the document.
        self.assertEqual(
            geometry.ordinal_of(sections[-1]["end_para_ref"], all_para_refs), total_paragraphs
        )

    def test_tables_docx_paragraph_after_table_counts_cell_paragraphs(self):
        docx_path = FIXTURE_DIR / "tables.docx"
        total_paragraphs = _count_w_p_elements(docx_path)
        self.assertEqual(total_paragraphs, 6)  # 1 before + 4 cells + 1 after

        proj = projection.project_part(docx_path)
        all_para_refs = [m.para_ref for m in proj.paragraphs]
        self.assertEqual(len(all_para_refs), total_paragraphs)

        # tables.docx has no headings (confirmed: none of its paragraphs
        # carry a heading style), so this exercises ordinal_of() directly
        # against the document's own last paragraph -- "After the table.",
        # whose ordinal must count the table's 4 cell paragraphs that
        # precede it in document order.
        self.assertEqual(
            [s for s in projection.find_sections_impl(docx_path) if s["kind"] == "heading"], []
        )
        last_para_ref = all_para_refs[-1]
        self.assertEqual(geometry.ordinal_of(last_para_ref, all_para_refs), total_paragraphs)
        self.assertEqual(geometry.ordinal_of(last_para_ref, all_para_refs), 6)


class BuildProbeOrdinalsTests(unittest.TestCase):
    def test_empty_sections_needs_no_probes(self):
        self.assertEqual(geometry.build_probe_ordinals([], []), [])

    def test_distinct_starts_plus_last_end_only(self):
        sections = [
            {"start_para_ref": "p0", "end_para_ref": "p1", "next_start_para_ref": "p2"},
            {"start_para_ref": "p2", "end_para_ref": "p4", "next_start_para_ref": None},
        ]
        all_para_refs = ["p0", "p1", "p2", "p3", "p4"]
        # Starts: ordinal_of(p0)=1, ordinal_of(p2)=3. First section's end
        # borrows the second's start (already ordinal 3, no extra probe).
        # The second section has no next_start_para_ref (it IS the
        # document's last heading), so its own end (p4)=5 is probed
        # instead. p1/p3 are never probed at all.
        self.assertEqual(geometry.build_probe_ordinals(sections, all_para_refs), [1, 3, 5])

    def test_next_start_para_ref_probed_even_when_not_in_filtered_list(self):
        # A single, explicitly-filtered section (section_keys=[...]) whose
        # next_start_para_ref points at a heading that is NOT itself
        # present in *sections* at all -- the follow-up behavior (issue
        # #102, 2026-09-16): the borrowed next heading's start ordinal
        # must still be probed even though only one section was passed
        # in and it never appears in the output.
        sections = [{"start_para_ref": "p0", "end_para_ref": "p1", "next_start_para_ref": "p2"}]
        all_para_refs = ["p0", "p1", "p2", "p3"]
        self.assertEqual(geometry.build_probe_ordinals(sections, all_para_refs), [1, 3])


# ---------------------------------------------------------------------------
# Geometry math (spec test 2)
# ---------------------------------------------------------------------------


class AssembleSectionsGeometryTests(unittest.TestCase):
    def test_two_sections_on_one_page(self):
        sections = [
            {
                "section_key": "intro-1",
                "kind": "heading",
                "heading_text": "Intro",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": "p2",
            },
            {
                "section_key": "body-1",
                "kind": "heading",
                "heading_text": "Body",
                "start_para_ref": "p2",
                "end_para_ref": "p3",
                "next_start_para_ref": None,
            },
        ]
        all_para_refs = ["p0", "p1", "p2", "p3"]
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Intro"},
            3: {"page": 1, "vpos_pt": 300.0, "text": "Body"},
            4: {"page": 1, "vpos_pt": 400.0, "text": "last line of Body"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        self.assertEqual(len(result), 2)

        intro, body = result
        self.assertEqual(intro["section_key"], "intro-1")
        self.assertEqual(intro["start_page"], 1)
        self.assertEqual(intro["start_fraction"], 0.09)
        # Intro's end borrows Body's own start (contiguous sections).
        self.assertEqual(intro["end_page"], 1)
        self.assertEqual(intro["end_fraction"], 0.38)
        self.assertEqual(intro["pages"], 0.29)
        self.assertEqual(intro["start_paragraph"], 1)
        self.assertEqual(intro["end_paragraph"], 2)
        self.assertTrue(intro["text_verified"])

        self.assertEqual(body["section_key"], "body-1")
        self.assertEqual(body["start_page"], 1)
        self.assertEqual(body["start_fraction"], 0.38)
        # Body is the last section: its end comes from its OWN
        # end_para_ref probe (ordinal 4), nudged by +0.03.
        self.assertEqual(body["end_page"], 1)
        self.assertEqual(body["end_fraction"], 0.54)
        self.assertEqual(body["pages"], 0.16)
        self.assertEqual(body["start_paragraph"], 3)
        self.assertEqual(body["end_paragraph"], 4)

    def test_one_section_spanning_two_pages(self):
        sections = [
            {
                "section_key": "a-1",
                "kind": "heading",
                "heading_text": "A",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": "p2",
            },
            {
                "section_key": "b-1",
                "kind": "heading",
                "heading_text": "B",
                "start_para_ref": "p2",
                "end_para_ref": "p3",
                "next_start_para_ref": None,
            },
        ]
        all_para_refs = ["p0", "p1", "p2", "p3"]
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "A"},
            3: {"page": 2, "vpos_pt": 100.0, "text": "B"},
            4: {"page": 2, "vpos_pt": 200.0, "text": "last line of B"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        section_a = result[0]
        self.assertEqual(section_a["start_page"], 1)
        self.assertEqual(section_a["start_fraction"], 0.09)
        self.assertEqual(section_a["end_page"], 2)
        self.assertEqual(section_a["end_fraction"], 0.13)
        self.assertEqual(section_a["pages"], 1.04)

    def test_last_section_end_approximation_with_792pt_page_height(self):
        # Mirrors the issue's own live-probe numbers (Word 16.112.4,
        # 2026-09-16): start_fraction 0.14 for a paragraph at page 5,
        # vpos 110.88pt on a 792.0pt-tall page.
        sections = [
            {
                "section_key": "factor-1-mission-focused-corporate-experience-5-pages-total-1",
                "kind": "heading",
                "heading_text": "Factor 1",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": None,
            }
        ]
        all_para_refs = ["p0", "p1"]
        paragraph_geometry = {
            1: {"page": 5, "vpos_pt": 110.88, "text": "Factor 1"},
            2: {"page": 10, "vpos_pt": 341.2, "text": "last line"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        entry = result[0]
        self.assertEqual(entry["start_page"], 5)
        self.assertEqual(entry["start_fraction"], 0.14)
        self.assertEqual(entry["end_page"], 10)
        # 341.2 / 792.0 + 0.03 = 0.4605... -> rounded to 0.46
        self.assertEqual(entry["end_fraction"], 0.46)
        self.assertEqual(entry["pages"], 5.32)

    def test_last_section_end_fraction_capped_at_one(self):
        sections = [
            {
                "section_key": "tail-1",
                "kind": "heading",
                "heading_text": "Tail",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": None,
            }
        ]
        all_para_refs = ["p0", "p1"]
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Tail"},
            # 780.0 / 792.0 + 0.03 == 1.0148... -- must cap at 1.0, never
            # report a fraction > 1.0.
            2: {"page": 1, "vpos_pt": 780.0, "text": "last line"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        self.assertEqual(result[0]["end_fraction"], 1.0)

    def test_empty_sections_returns_empty_list_no_error(self):
        result, error = geometry.assemble_sections([], [], {}, None)
        self.assertEqual(result, [])
        self.assertIsNone(error)

    def test_single_filtered_section_borrows_full_document_next_heading_start(self):
        # Follow-up (issue #102, 2026-09-16): an explicit, single-section
        # section_keys=[...] call still borrows the FULL document's next
        # heading's start for its end geometry, even though that next
        # heading is not itself part of *sections* and never appears in
        # the returned list -- this is the skills' primary page-budget
        # call shape. "Background" (ordinal 3, on page 2) is the next
        # heading; it supplies end_page/end_fraction directly, with NO
        # +0.03 approximation and NO text verification against it.
        sections = [
            {
                "section_key": "overview-1",
                "kind": "heading",
                "heading_text": "Overview",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": "p2",
            }
        ]
        all_para_refs = ["p0", "p1", "p2", "p3"]
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Overview"},
            3: {"page": 2, "vpos_pt": 100.0, "text": "Background"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        self.assertEqual(len(result), 1)
        entry = result[0]
        self.assertEqual(entry["section_key"], "overview-1")
        self.assertEqual(entry["start_page"], 1)
        self.assertEqual(entry["start_fraction"], 0.09)
        # end == the next heading's own start, exactly -- no +0.03 nudge.
        self.assertEqual(entry["end_page"], 2)
        self.assertEqual(entry["end_fraction"], round(100.0 / 792.0, 2))
        self.assertEqual(entry["end_fraction"], 0.13)
        self.assertEqual(entry["pages"], round((2 + 0.13) - (1 + 0.09), 2))
        # start_paragraph/end_paragraph stay the section's OWN structural
        # ordinals (ordinal_of(p1)=2), never the borrowed heading's (3).
        self.assertEqual(entry["start_paragraph"], 1)
        self.assertEqual(entry["end_paragraph"], 2)

    def test_single_filtered_last_heading_still_uses_approximation(self):
        # The other half of the same follow-up: when the ONE requested
        # section IS the document's own last heading (next_start_para_ref
        # is None even though it came from an explicit, single-section
        # section_keys=[...] filter), the +0.03 last-paragraph
        # approximation still applies -- there is no next heading to
        # borrow from, full document or otherwise.
        sections = [
            {
                "section_key": "next-steps-1",
                "kind": "heading",
                "heading_text": "Next Steps",
                "start_para_ref": "p4",
                "end_para_ref": "p5",
                "next_start_para_ref": None,
            }
        ]
        all_para_refs = ["p0", "p1", "p2", "p3", "p4", "p5"]
        paragraph_geometry = {
            5: {"page": 1, "vpos_pt": 400.0, "text": "Next Steps"},
            6: {"page": 1, "vpos_pt": 600.0, "text": "last line of Next Steps"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(error)
        entry = result[0]
        self.assertEqual(entry["end_fraction"], round(min(1.0, 600.0 / 792.0 + 0.03), 2))


# ---------------------------------------------------------------------------
# Verification / probe-error failure contract (spec test 3)
# ---------------------------------------------------------------------------


class AssembleSectionsFailureTests(unittest.TestCase):
    def _one_section(self):
        return [
            {
                "section_key": "background-1",
                "kind": "heading",
                "heading_text": "Background",
                "start_para_ref": "p0",
                "end_para_ref": "p1",
                "next_start_para_ref": None,
            }
        ], ["p0", "p1"]

    def test_heading_text_mismatch_returns_none_and_detail(self):
        sections, all_para_refs = self._one_section()
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Backgroundx"},  # mismatch
            2: {"page": 1, "vpos_pt": 200.0, "text": "last line"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(result)
        self.assertIsNotNone(error)
        self.assertIn("background-1", error)
        self.assertIn("Background", error)

    def test_probe_error_entry_returns_none_and_detail(self):
        sections, all_para_refs = self._one_section()
        paragraph_geometry = {
            1: {"error": "paragraph 1 of active document is out of range"},
            2: {"page": 1, "vpos_pt": 200.0, "text": "last line"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(result)
        self.assertIn("probe error", error)
        self.assertIn("out of range", error)

    def test_missing_ordinal_returns_none_and_detail(self):
        sections, all_para_refs = self._one_section()
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Background"},
            # ordinal 2 (the section's own end_para_ref) never came back.
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, 792.0)
        self.assertIsNone(result)
        self.assertIn("missing", error)

    def test_page_height_none_returns_none_and_detail(self):
        sections, all_para_refs = self._one_section()
        paragraph_geometry = {
            1: {"page": 1, "vpos_pt": 72.0, "text": "Background"},
            2: {"page": 1, "vpos_pt": 200.0, "text": "last line"},
        }
        result, error = geometry.assemble_sections(sections, all_para_refs, paragraph_geometry, None)
        self.assertIsNone(result)
        self.assertIn("page_height_pt", error)


# ---------------------------------------------------------------------------
# normalize_probe_text()
# ---------------------------------------------------------------------------


class NormalizeProbeTextTests(unittest.TestCase):
    def test_strips_and_collapses_whitespace(self):
        self.assertEqual(geometry.normalize_probe_text("  Factor   1  "), "Factor 1")

    def test_straightens_curly_quotes(self):
        self.assertEqual(geometry.normalize_probe_text("“Quoted”"), '"Quoted"')
        self.assertEqual(geometry.normalize_probe_text("Mike’s"), "Mike's")

    def test_drops_soft_hyphen(self):
        self.assertEqual(geometry.normalize_probe_text("Back­ground"), "Background")

    def test_drops_control_characters_without_merging_words(self):
        # TAB (U+0009) is folded to a space by whitespace collapse, not
        # dropped outright -- see the module's _CONTROL_RE comment.
        self.assertEqual(geometry.normalize_probe_text("Factor 1\tOverview"), "Factor 1 Overview")
        # A field/tab-stop marker (U+0013, within the C0 drop-set) IS
        # dropped outright.
        self.assertEqual(geometry.normalize_probe_text("Factor1"), "Factor1")

    def test_comparison_is_case_sensitive(self):
        self.assertNotEqual(
            geometry.normalize_probe_text("background"), geometry.normalize_probe_text("Background")
        )

    def test_none_and_empty_normalize_to_empty_string(self):
        self.assertEqual(geometry.normalize_probe_text(None), "")
        self.assertEqual(geometry.normalize_probe_text(""), "")


if __name__ == "__main__":
    unittest.main()
