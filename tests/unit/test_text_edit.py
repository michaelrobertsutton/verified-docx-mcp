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


class ApplyStyleRegistrationTests(unittest.TestCase):
    def test_apply_style_is_registered(self):
        self.assertIn("apply_style", MUTATING_TOOLS)


class ApplyStyleCharacterTests(_TempFixtureCase):
    """issue #28 WP-15a: apply_style on a w:type="character" style --
    same locate/run-splitting/track_changes machinery as format_text."""

    fixture_name = "frag.docx"

    def test_applies_rstyle_to_matched_run(self):
        evidence = text_edit.execute_apply_style(str(self.target), "brown", "IntenseEmphasis", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["style_id"], "IntenseEmphasis")
        self.assertEqual(evidence["style_type"], "character")
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('<w:rStyle w:val="IntenseEmphasis"', xml)
        # rStyle is rPr's FIRST child per the OOXML schema's fixed element
        # order, ahead of the run's own pre-existing <w:b/>.
        self.assertIn('<w:rStyle w:val="IntenseEmphasis" /><w:b', xml)

    def test_unknown_style_id_raises_style_not_found(self):
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_apply_style(str(self.target), "brown", "NoSuchStyle", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.STYLE_NOT_FOUND)

    def test_table_style_id_rejected_as_unsupported_type(self):
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_apply_style(str(self.target), "brown", "TableNormal", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.UNSUPPORTED_STYLE_TYPE)

    def test_track_changes_wraps_in_rprchange(self):
        patcher = mock.patch("verified_docx_mcp.text_edit.resolve_author_name", return_value="Jane Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)
        evidence = text_edit.execute_apply_style(
            str(self.target), "brown", "IntenseEmphasis", 1, track_changes=True
        )
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn("<w:rPrChange ", xml)
        self.assertIn('w:author="Jane Reviewer"', xml)
        # Content itself is untouched by a style-only change.
        self.assertEqual(
            projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog."
        )


class ApplyStyleParagraphTests(_TempFixtureCase):
    fixture_name = "frag.docx"

    def test_applies_pstyle_to_enclosing_paragraph(self):
        evidence = text_edit.execute_apply_style(str(self.target), "brown", "Heading2", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["style_type"], "paragraph")
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('<w:pStyle w:val="Heading2"', xml)
        # Whole-paragraph text is unchanged -- pStyle carries no
        # sub-paragraph granularity.
        self.assertEqual(
            projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog."
        )

    def test_track_changes_writes_rprchange_shaped_pprchange(self):
        # frag.docx's own paragraph has no explicit w:pPr at all before
        # this call -- the same starting shape as the real Word-authored
        # fixture (tests/fixtures/revision/pstyle-tracked.docx), so the
        # produced XML should match it structurally: w:pPrChange as
        # w:pPr's LAST child, right after w:pStyle, with an empty
        # <w:pPr/> snapshot.
        evidence = text_edit.execute_apply_style(str(self.target), "brown", "Heading2", 1, track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertEqual(len(evidence["revision_ids"]), 1)
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('<w:pStyle w:val="Heading2" /><w:pPrChange ', xml)
        self.assertIn("<w:pPr />", xml)
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)


class ApplyStyleParagraphTrackedTests(_TempFixtureCase):
    """issue #7: w:pPrChange for apply_style(track_changes=True) on a
    paragraph style, verified against real Word-authored fixtures
    (tests/fixtures/revision/pstyle-tracked*.docx) rather than an assumed
    shape -- see tests/fixtures/README.md for how they were produced."""

    fixture_name = "sections.docx"

    def _pprchange_xml_around(self, needle: str) -> str:
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        idx = xml.find(needle)
        self.assertNotEqual(idx, -1, f"{needle!r} not found in document.xml")
        return xml[max(0, idx - 400) : idx + len(needle) + 40]

    def test_pprchange_shape_matches_word_authored_fixture(self):
        evidence = text_edit.execute_apply_style(
            str(self.target), "Background text.", "Heading2", 1, track_changes=True
        )
        self.assertEqual(evidence["revision_ids"], ["1"])
        snippet = self._pprchange_xml_around("Background text.")
        self.assertIn('<w:pStyle w:val="Heading2" />', snippet)
        self.assertIn("<w:pPrChange ", snippet)
        self.assertIn("<w:pPr />", snippet)
        # w:pPrChange is pPr's LAST child, right after w:pStyle -- same
        # order the real Word-authored fixture uses.
        self.assertIn('<w:pPr><w:pStyle w:val="Heading2" /><w:pPrChange', snippet)
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)

    def test_noop_when_style_already_matches_writes_no_pprchange(self):
        # "Overview" already carries an explicit w:pStyle val="Heading1".
        evidence = text_edit.execute_apply_style(str(self.target), "Overview", "Heading1", 1, track_changes=True)
        self.assertEqual(evidence["revision_ids"], [])
        self.assertEqual(evidence["revision_before"], evidence["revision_after"])
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertNotIn("pPrChange", xml)

    def test_second_tracked_edit_reuses_existing_pprchange(self):
        # Real Word behavior (tests/fixtures/revision/pstyle-tracked-twice.docx):
        # a second same-author tracked style change on the same paragraph,
        # before any accept/reject, updates pStyle in place and leaves the
        # FIRST change record's id/date/snapshot untouched.
        first = text_edit.execute_apply_style(
            str(self.target), "Background text.", "Heading2", 1, track_changes=True
        )
        second = text_edit.execute_apply_style(
            str(self.target), "Background text.", "Heading1", 1, track_changes=True
        )
        self.assertEqual(first["revision_ids"], ["1"])
        self.assertEqual(second["revision_ids"], [])  # no NEW id consumed -- reused
        snippet = self._pprchange_xml_around("Background text.")
        self.assertIn('<w:pStyle w:val="Heading1" />', snippet)
        self.assertIn('<w:pPrChange w:id="1" w:author="Michael Sutton"', snippet)
        self.assertIn("<w:pPr />", snippet)  # original (pre-first-edit) snapshot, unchanged
        self.assertEqual(snippet.count("<w:pPrChange "), 1)  # never stacked/nested
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)

    def test_foreign_author_pprchange_refuses_and_force_overrides(self):
        # Real Word does NOT itself protect a foreign author's pending
        # pPrChange here (tests/fixtures/revision/pstyle-tracked-foreign.docx
        # -- a second Word session under a different user name silently
        # overwrites it) -- this tool's own, stricter safety policy does.
        text_edit.execute_apply_style(str(self.target), "Background text.", "Heading2", 1, track_changes=True)
        with mock.patch("verified_docx_mcp.text_edit.resolve_author_name", return_value="Jordan Author"):
            with self.assertRaises(VerifyError) as cm:
                text_edit.execute_apply_style(str(self.target), "Background text.", "Heading1", 1, track_changes=True)
            self.assertEqual(cm.exception.envelope.error_code, ErrorCode.TRACKED_CHANGES_PRESENT)

            forced = text_edit.execute_apply_style(
                str(self.target), "Background text.", "Heading1", 1, track_changes=True, force=True
            )
        self.assertEqual(forced["revision_ids"], ["2"])
        snippet = self._pprchange_xml_around("Background text.")
        self.assertIn('w:author="Jordan Author"', snippet)
        self.assertEqual(snippet.count("<w:pPrChange "), 1)  # the foreign one was replaced, not stacked
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)

    def test_revision_id_allocator_sees_existing_change_record_ids(self):
        # Regression for the RevisionIdAllocator fix: a fresh id must not
        # collide with an existing w:rPrChange/w:pPrChange id already in
        # the package. Simulate by tracking a run-level format_text
        # change first (consumes id "1"), then a paragraph-style change
        # (must allocate "2", not reuse "1").
        run_evidence = text_edit.execute_format_text(
            str(self.target), "Overview", {"bold": True}, 1, track_changes=True
        )
        para_evidence = text_edit.execute_apply_style(
            str(self.target), "Background text.", "Heading2", 1, track_changes=True
        )
        self.assertEqual(run_evidence["revision_ids"], ["1"])
        self.assertEqual(para_evidence["revision_ids"], ["2"])

    def test_no_nested_change_record_on_repeated_run_level_edit(self):
        # Pre-existing defect this issue also fixes: two consecutive
        # track_changes=True format_text calls on the same run, same
        # author, must produce exactly ONE w:rPrChange (via the same
        # own-author-reuse rule _apply_style_to_run now applies -- see
        # its docstring), never a second, doubly-nested one. A FOREIGN-
        # author repeat (apply_rpr_change's own remove-existing-record
        # safety net, exercised via force=True) is covered by the
        # paragraph-level equivalent above
        # (test_foreign_author_pprchange_refuses_and_force_overrides) --
        # same underlying tracked_changes.apply_rpr_change/apply_ppr_change
        # mechanism, only the element tag differs.
        text_edit.execute_format_text(str(self.target), "Overview", {"bold": True}, 1, track_changes=True)
        text_edit.execute_format_text(str(self.target), "Overview", {"italic": True}, 1, track_changes=True)
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        idx = xml.find("Overview")
        snippet = xml[max(0, idx - 400) : idx]
        self.assertEqual(snippet.count("<w:rPrChange "), 1)
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)


if __name__ == "__main__":
    unittest.main()
