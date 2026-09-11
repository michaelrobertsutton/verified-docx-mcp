"""Unit tests for src/verified_docx_mcp/locate.py (issue #28 WP-06):
locate()'s 4-rung normalization ladder, STRUCTURAL_BOUNDARY, the
crosses_comment_range/crosses_revision/touches_field_result warnings, and
expected_matches enforcement (D4: required, no default).

RUNG_FIELD (the plan text's own name) is deliberately NOT a fifth match
rung here -- see locate.py's own header comment and TouchesFieldResultTests
below for the corrected-projection reasoning and the corruption risk a
literal implementation would have carried.

Fixture (a) (frag.docx) is the load-bearing acceptance test named in the
WP: the fragmented phrase must be found at the EXACT rung with every
original w:rPr intact on its original characters -- a higher rung would
mean the projection itself is wrong (see projection.py's own module
docstring and tests/fixtures/README.md for frag.docx's provenance).
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import locate, projection
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"


class ExactRungFragmentedPhraseTests(unittest.TestCase):
    """Acceptance (a): frag.docx's fragmented phrase locates at RUNG_EXACT,
    and every original run's w:rPr survives on its own original characters
    (verified here via the projection's own RunEvents, which text_edit.py's
    runs_before/runs_after evidence is built from directly)."""

    def test_locates_at_exact_rung_not_higher(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        result = locate.locate("The quick brown fox jumps over the lazy dog.", proj, 1)
        self.assertEqual(result.rung, locate.RUNG_EXACT)
        self.assertEqual(result.match_count, 1)
        self.assertEqual(result.spans, [(0, 44)])

    def test_original_rpr_intact_per_run(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        run_events = [e for e in proj.events if isinstance(e, projection.RunEvent)]
        # Three runs, matching tests/fixtures/README.md's frag.docx spec:
        # before (unbold) / "brown" (bold) / after (unbold), each carrying
        # its OWN original characters.
        self.assertEqual([e.text for e in run_events], ["The quick ", "brown", " fox jumps over the lazy dog."])
        self.assertEqual([e.rpr["bold"] for e in run_events], [False, True, False])


class StructuralBoundaryTests(unittest.TestCase):
    """Acceptance: a match crossing a w:p/w:tbl/w:tc boundary refuses."""

    def test_crosses_table_cell_boundary(self):
        proj = projection.project_part(FIXTURES / "tables.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("R1C1\nR1C2", proj, 1)
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.STRUCTURAL_BOUNDARY)
        self.assertEqual(set(env.diagnostics["boundary_kinds"]), {"w:p", "w:tbl", "w:tc"})

    def test_within_one_cell_does_not_cross(self):
        proj = projection.project_part(FIXTURES / "tables.docx")
        result = locate.locate("R1C1", proj, 1)
        self.assertEqual(result.rung, locate.RUNG_EXACT)

    def test_crosses_ordinary_paragraph_boundary_body_text(self):
        proj = projection.project_part(FIXTURES / "tables.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("Before the table.\nR1C1", proj, 1)
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.STRUCTURAL_BOUNDARY)
        # Crosses into the table but the "before" paragraph itself is not
        # inside one -- w:tc must not be claimed for a paragraph that was
        # never in any cell.
        self.assertIn("w:p", env.diagnostics["boundary_kinds"])
        self.assertIn("w:tbl", env.diagnostics["boundary_kinds"])
        self.assertNotIn("w:tc", env.diagnostics["boundary_kinds"])


class ExpectedMatchesTests(unittest.TestCase):
    """D4: expected_matches is required (this test calls locate() directly
    with an explicit value; the required-ness of the *tool* parameter is
    covered in test_text_edit.py against the actual function signature)."""

    def test_match_count_mismatch(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("o", proj, 1)
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.MATCH_COUNT_MISMATCH)
        self.assertEqual(env.diagnostics["expected"], 1)
        self.assertGreater(env.diagnostics["actual"], 1)

    def test_zero_match_reports_near_miss(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("The quick brown fox jumps over the lazy catt.", proj, 1)
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.ZERO_MATCH)
        self.assertIn("ladder_report", env.diagnostics)
        self.assertIn("near_miss", env.diagnostics)

    def test_empty_needle_is_invalid_input(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("", proj, 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class NormalizationLadderTests(unittest.TestCase):
    def test_curly_quote_matches_straight_needle(self):
        # frag.docx has no quotes at all; confirm the ladder still reaches
        # rung 2 correctly using a synthetic haystack via a hand-built
        # Projection is out of scope here -- exercised instead through the
        # normalizer helpers directly, which is what rung 2+ actually is.
        haystack = "He said ‘hello’ today."
        normalized = haystack.translate(locate._QUOTE_MAP)
        self.assertEqual(normalized, "He said 'hello' today.")

    def test_whitespace_collapse_orig_pos_roundtrips(self):
        text = "a   b\tc"
        normalized, orig_pos = locate._norm_whitespace(text)
        self.assertEqual(normalized, "a b c")
        # Every normalized position maps back into the original string
        # (or to the sentinel at len(text)).
        for pos in orig_pos:
            self.assertLessEqual(pos, len(text))

    def test_soft_hyphen_stripped(self):
        text = "auto­matic"
        normalized, _ = locate._norm_softhyphen(text)
        self.assertEqual(normalized, "automatic")


class TouchesFieldResultTests(unittest.TestCase):
    """RUNG_FIELD deviation (see locate.py's own header comment): rebasing
    onto PR #3 (feat/28-wp-04-markdown-mutations @ 7e5aecf) fixed the WP-03
    defect the original RUNG_FIELD design depended on -- a field's markdown
    rendering used to emit its live result text AND a "[FIELD:instr]"
    placeholder back to back; corrected, a result-bearing field renders as
    plain live text (no placeholder at all) and a no-result field renders
    as a lowercase "[field:instr]" placeholder while contributing ZERO
    RunEvents (a zero-width point, not a span) to the projection. Building
    a haystack-substitution rung for either case is therefore either
    unnecessary (the result text already matches at RUNG_EXACT) or unsafe
    (splicing across a no-result field's unmodeled w:fldChar/w:fldSimple
    skeleton would corrupt it) -- so this WP surfaces
    WARNING_TOUCHES_FIELD_RESULT instead of a fifth rung. fields.docx's own
    two fields (PAGE, REF Target1) both have a result; it carries no
    no-result field, so the scope limit below is demonstrated by a
    verbatim-copied placeholder legitimately failing to locate, not by a
    fixture that could exercise the no-result case positively."""

    def test_result_bearing_field_locates_at_exact_rung_with_warning(self):
        proj = projection.project_part(FIXTURES / "fields.docx")
        # Unique in fields.docx's projected text, and entirely inside one
        # paragraph (no STRUCTURAL_BOUNDARY risk) -- spans the REF field's
        # own live result text ("Target paragraph ") at the tail.
        result = locate.locate("Cross-reference: .Target paragraph ", proj, 1)
        self.assertEqual(result.rung, locate.RUNG_EXACT)
        self.assertIn(locate.WARNING_TOUCHES_FIELD_RESULT, result.warnings)

    def test_ordinary_text_with_no_field_overlap_carries_no_field_warning(self):
        proj = projection.project_part(FIXTURES / "fields.docx")
        result = locate.locate("Target paragraph for cross-reference.", proj, 1)
        self.assertEqual(result.rung, locate.RUNG_EXACT)
        self.assertNotIn(locate.WARNING_TOUCHES_FIELD_RESULT, result.warnings)

    def test_old_uppercase_placeholder_now_zero_matches(self):
        # The ORIGINAL (pre-deviation) design's own needle shape -- proof
        # the corrected projection no longer produces what RUNG_FIELD was
        # built to substitute for.
        proj = projection.project_part(FIXTURES / "fields.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("[FIELD:PAGE  \\* MERGEFORMAT]", proj, 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ZERO_MATCH)

    def test_lowercase_no_result_placeholder_also_zero_matches(self):
        # fields.docx has no no-result field to locate positively (see
        # class docstring) -- this documents the scope limit itself: even
        # the CORRECT lowercase placeholder shape for a field that doesn't
        # exist in this fixture legitimately fails to locate, rather than
        # silently matching something wrong.
        proj = projection.project_part(FIXTURES / "fields.docx")
        with self.assertRaises(VerifyError) as cm:
            locate.locate("[field:PAGE  \\* MERGEFORMAT]", proj, 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ZERO_MATCH)


class WarningsTests(unittest.TestCase):
    """crosses_comment_range / crosses_revision are non-fatal warnings at
    this WP (WP-07 adds the actual refusal on top of this detection)."""

    def test_crosses_comment_range_warning(self):
        proj = projection.project_part(FIXTURES / "revision" / "commented.docx")
        result = locate.locate("This", proj, 1)
        self.assertIn("crosses_comment_range", result.warnings)

    def test_crosses_revision_warning(self):
        proj = projection.project_part(FIXTURES / "revision" / "tracked.docx")
        result = locate.locate("X", proj, 1)
        self.assertIn("crosses_revision", result.warnings)

    def test_no_warnings_on_ordinary_text(self):
        proj = projection.project_part(FIXTURES / "frag.docx")
        result = locate.locate("brown", proj, 1)
        self.assertEqual(result.warnings, [])


if __name__ == "__main__":
    unittest.main()
