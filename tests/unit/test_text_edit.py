"""Unit tests for src/verified_docx_mcp/text_edit.py (issue #28 WP-06):
replace_text and format_text end-to-end -- run splitting, w:rPr
inheritance/cloning, the eight standard evidence keys plus runs_before/
runs_after, expected_matches enforcement (D4: required, no default on the
tool signature itself), and MUTATING_TOOLS registration.

Fixture provenance: frag.docx and tables.docx are the Word-authored
fixtures from WP-03 (tests/fixtures/README.md); frag.docx's own spec is
this WP's fixture (a).
"""

from __future__ import annotations

import inspect
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paths, projection, text_edit, tracked_changes
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

_EVIDENCE_KEYS = {
    "applied",
    "match_count",
    "rung",
    "before",
    "after",
    "revision_before",
    "revision_after",
    "audit_logged",
    # issue #28 WP-10: layer 3's conflict-copy sweep result, always present
    # (True/False) on every successful write.
    "conflict_copy_detected",
}


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "frag.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()


class MutatingToolsRegistrationTests(unittest.TestCase):
    def test_replace_text_and_format_text_are_registered(self):
        self.assertIn("replace_text", MUTATING_TOOLS)
        self.assertIn("format_text", MUTATING_TOOLS)


class ExpectedMatchesRequiredTests(unittest.TestCase):
    """D4: expected_matches has no default on either tool's own signature."""

    def test_replace_text_has_no_default(self):
        sig = inspect.signature(text_edit.execute_replace_text)
        self.assertEqual(sig.parameters["expected_matches"].default, inspect.Parameter.empty)

    def test_format_text_has_no_default(self):
        sig = inspect.signature(text_edit.execute_format_text)
        self.assertEqual(sig.parameters["expected_matches"].default, inspect.Parameter.empty)


class ReplaceTextWholeRunTests(_TempFixtureCase):
    """Fixture (a): replacing the whole fragmented phrase -- every original
    run's w:rPr must be visible (and correctly attributed) in runs_before
    before the run boundaries disappear into one replacement run."""

    def test_whole_phrase_replace_at_exact_rung_runs_before_shows_original_rpr(self):
        evidence = text_edit.execute_replace_text(
            str(self.target), "The quick brown fox jumps over the lazy dog.", "Done.", 1
        )
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        self.assertEqual(evidence["rung"], "exact")
        self.assertEqual(evidence["match_count"], 1)
        self.assertTrue(evidence["applied"])
        self.assertIn("runs_before", evidence)
        self.assertIn("runs_after", evidence)
        # The ORIGINAL 3 runs, each with its own original rPr and text,
        # not flattened into one.
        span_runs = evidence["runs_before"][0]
        self.assertEqual([r["text"] for r in span_runs], ["The quick ", "brown", " fox jumps over the lazy dog."])
        self.assertEqual([r["bold"] for r in span_runs], [False, True, False])
        self.assertEqual(projection.read_document_text(self.target), "Done.")


class ReplaceTextRunSplittingTests(_TempFixtureCase):
    def test_mid_run_replace_splits_boundary_run_and_preserves_neighbors(self):
        # "fox" sits inside the third (unbold) run; replacing it must leave
        # the bold "brown" run and the surrounding unbold text untouched.
        evidence = text_edit.execute_replace_text(str(self.target), "fox", "wolf", 1)
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(projection.read_document_text(self.target), "The quick brown wolf jumps over the lazy dog.")

        runs = [r for r in projection.read_document_runs(self.target) if "rPr" in r]
        bold_runs = [r for r in runs if r["rPr"]["bold"]]
        self.assertEqual([r["text"] for r in bold_runs], ["brown"])
        self.assertEqual("".join(r["text"] for r in runs), "The quick brown wolf jumps over the lazy dog.")

    def test_replacement_run_inherits_first_touched_runs_rpr(self):
        # A match spanning the bold "brown" run and into the trailing
        # unbold run: the ONE new replacement run must carry "brown"'s
        # (the first touched run's) bold rPr, not the trailing run's.
        text_edit.execute_replace_text(str(self.target), "brown fox", "wolf", 1)
        self.assertEqual(projection.read_document_text(self.target), "The quick wolf jumps over the lazy dog.")
        runs = [r for r in projection.read_document_runs(self.target) if "rPr" in r]
        wolf_run = next(r for r in runs if r["text"] == "wolf")
        self.assertTrue(wolf_run["rPr"]["bold"], "replacement must inherit the FIRST touched run's (bold) rPr")

    def test_post_write_verification_reads_the_replacement(self):
        # execute_replace_text's own post_verify already asserts this on
        # every call (it would raise VERIFICATION_FAILED otherwise); assert
        # it positively here too, against the file on disk.
        text_edit.execute_replace_text(str(self.target), "lazy", "sleepy", 1)
        self.assertIn("sleepy", projection.read_document_text(self.target))
        self.assertNotIn("lazy", projection.read_document_text(self.target))


class ExpectedMatchesEnforcementTests(_TempFixtureCase):
    def test_match_count_mismatch_no_write(self):
        before_text = projection.read_document_text(self.target)
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_replace_text(str(self.target), "o", "0", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.MATCH_COUNT_MISMATCH)
        self.assertEqual(projection.read_document_text(self.target), before_text)


class StructuralBoundaryRefusalTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_replace_text_refuses_across_cell_boundary(self):
        before_text = projection.read_document_text(self.target)
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_replace_text(str(self.target), "R1C1\nR1C2", "X", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.STRUCTURAL_BOUNDARY)
        self.assertEqual(projection.read_document_text(self.target), before_text, "a refused match must not write")


class FormatTextTests(_TempFixtureCase):
    def test_bold_whole_phrase_evidence_shape_and_runs(self):
        evidence = text_edit.execute_format_text(
            str(self.target), "The quick brown fox jumps over the lazy dog.", {"italic": True}, 1
        )
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        self.assertEqual(evidence["rung"], "exact")
        before_runs = evidence["runs_before"][0]
        after_runs = evidence["runs_after"][0]
        self.assertEqual([r["bold"] for r in before_runs], [False, True, False])
        self.assertEqual([r["italic"] for r in before_runs], [False, False, False])
        self.assertEqual([r["italic"] for r in after_runs], [True, True, True])
        self.assertEqual([r["bold"] for r in after_runs], [False, True, False], "bold must survive the italic-only request")
        # Content is untouched -- format_text never mutates text.
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")

    def test_mid_run_split_preserves_untouched_text_and_style(self):
        text_edit.execute_format_text(str(self.target), "fox", {"bold": True}, 1)
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")
        runs = [r for r in projection.read_document_runs(self.target) if "rPr" in r]
        fox_run = next(r for r in runs if r["text"] == "fox")
        self.assertTrue(fox_run["rPr"]["bold"])
        brown_run = next(r for r in runs if r["text"] == "brown")
        self.assertTrue(brown_run["rPr"]["bold"], "the pre-existing bold run must be unaffected")
        space_run = next(r for r in runs if r["text"] == " ")
        self.assertFalse(space_run["rPr"]["bold"], "the split-off unmatched space must keep its original (unbold) rPr")

    def test_noop_when_style_already_matches_skips_write(self):
        text_edit.execute_format_text(str(self.target), "brown", {"bold": True}, 1)
        # bold is already true on "brown" -- re-running must be a no-op:
        # same revision before and after (no new write/revision created).
        evidence = text_edit.execute_format_text(str(self.target), "brown", {"bold": True}, 1)
        self.assertEqual(evidence["revision_before"], evidence["revision_after"])

    def test_invalid_style_rejected(self):
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_format_text(str(self.target), "brown", {"comic_sans": True}, 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_empty_style_rejected(self):
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_format_text(str(self.target), "brown", {}, 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class WarningsSurfaceOnEvidenceTests(_TempFixtureCase):
    fixture_name = "revision/commented.docx"

    def test_crosses_comment_range_warning_does_not_block_wp06(self):
        # WP-06 only warns; it does not refuse (WP-07 adds the refusal).
        evidence = text_edit.execute_replace_text(str(self.target), "This", "That", 1)
        self.assertTrue(evidence["applied"])
        self.assertIn("crosses_comment_range", evidence.get("warnings", []))


class TrackChangesTests(_TempFixtureCase):
    """Issue #28 WP-07b-a: replace_text/format_text's track_changes=True.

    Patches author.resolve_author_name to a fixed name (rather than
    relying on this machine's own resolved fallback) so w:author
    assertions are deterministic regardless of whose machine runs the
    suite -- see test_author.py for the resolution logic itself.
    """

    def setUp(self):
        super().setUp()
        patcher = mock.patch("verified_docx_mcp.text_edit.resolve_author_name", return_value="Jane Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def _document_xml(self) -> str:
        with zipfile.ZipFile(self.target) as zf:
            return zf.read("word/document.xml").decode("utf-8")

    def test_replace_text_tracked_evidence_and_xml_shape(self):
        evidence = text_edit.execute_replace_text(str(self.target), "fox", "wolf", 1, track_changes=True)
        self.assertTrue(evidence["applied"])
        self.assertTrue(evidence["track_changes"])
        self.assertIsInstance(evidence["revision_ids"], list)
        self.assertEqual(len(evidence["revision_ids"]), 2, "one w:del + one w:ins for a single-run replace")

        # Content read back as current text is the NEW text (w:del
        # excluded, w:ins included) -- unchanged read-side contract.
        self.assertEqual(projection.read_document_text(self.target), "The quick brown wolf jumps over the lazy dog.")

        doc = self._document_xml()
        self.assertIn('<w:del w:id="', doc)
        self.assertIn("<w:delText>fox</w:delText>", doc)
        self.assertIn('<w:ins w:id="', doc)
        self.assertIn("<w:t>wolf</w:t>", doc)
        self.assertIn('w:author="Jane Reviewer"', doc)
        for rid in evidence["revision_ids"]:
            self.assertIn(f'w:id="{rid}"', doc)

    def test_replace_text_untracked_has_no_track_changes_key(self):
        evidence = text_edit.execute_replace_text(str(self.target), "fox", "wolf", 1)
        self.assertNotIn("track_changes", evidence)
        self.assertNotIn("revision_ids", evidence)
        self.assertNotIn("<w:del", self._document_xml())
        self.assertNotIn("<w:ins", self._document_xml())

    def test_format_text_tracked_produces_rprchange_with_old_style(self):
        evidence = text_edit.execute_format_text(str(self.target), "fox", {"bold": True}, 1, track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertEqual(len(evidence["revision_ids"]), 1)

        doc = self._document_xml()
        self.assertIn("<w:rPrChange", doc)
        self.assertIn('w:author="Jane Reviewer"', doc)
        # The recorded PRE-change state for "fox" (no rPr at all before)
        # is an empty <w:rPr /> nested inside rPrChange.
        rprchange_start = doc.find("<w:rPrChange")
        rprchange_snippet = doc[rprchange_start : rprchange_start + 200]
        self.assertIn("<w:rPr", rprchange_snippet)

        # Content is unaffected; only style + rPrChange metadata changed.
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")
        runs = [r for r in projection.read_document_runs(self.target) if "rPr" in r]
        fox_run = next(r for r in runs if r["text"] == "fox")
        self.assertTrue(fox_run["rPr"]["bold"])

    def test_reject_tracked_changes_restores_replace_text(self):
        evidence = text_edit.execute_replace_text(str(self.target), "fox", "wolf", 1, track_changes=True)
        self.assertEqual(projection.read_document_text(self.target), "The quick brown wolf jumps over the lazy dog.")

        reject_evidence = tracked_changes.execute_reject_tracked_changes(str(self.target), evidence["revision_ids"])
        self.assertTrue(reject_evidence["applied"])
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")

    def test_own_author_tracked_edit_does_not_deadlock_a_second_tracked_edit(self):
        # First tracked edit creates a w:ins authored by "Jane Reviewer"
        # (the patched own_author for this test class).
        text_edit.execute_replace_text(str(self.target), "fox", "wolf", 1, track_changes=True)
        # A second edit over that SAME (own-authored) insertion must not
        # refuse -- WP-07b-a's own-author exclusion.
        second = text_edit.execute_replace_text(str(self.target), "wolf", "coyote", 1, track_changes=True)
        self.assertTrue(second["applied"])
        self.assertEqual(projection.read_document_text(self.target), "The quick brown coyote jumps over the lazy dog.")

    def test_foreign_author_tracked_change_still_refuses(self):
        # tracked.docx's own real Word-authored revisions are authored
        # "Michael Sutton" -- different from this test's patched own_author
        # ("Jane Reviewer") -- so this must still refuse.
        target = Path(self._tmp.name) / "tracked.docx"
        shutil.copyfile(FIXTURES / "revision" / "tracked.docx", target)
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_replace_text(str(target), "X", "Y", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.TRACKED_CHANGES_PRESENT)


if __name__ == "__main__":
    unittest.main()
