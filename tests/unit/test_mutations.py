"""Unit tests for src/verified_docx_mcp/mutations.py (issue #28 WP-04):
the write guard (lock/sync/revision), the comment-anchor/tracked-change
hazard scan against REAL Word-authored fixtures, the atomic write's two
failure surfaces (OPC_INVALID pre-replace, VERIFICATION_FAILED
post-replace with a .jsbak restore), and the three mutating tools
end-to-end (evidence shape, round trip, range targeting, append).

Fixture provenance: tests/fixtures/revision/commented.docx and
tracked.docx are real Word-authored packages (see
tests/fixtures/README.md) — the hazard-detection tests run the actual
guard logic against genuine w:commentRangeStart/End/commentReference and
w:ins/w:del elements, not hand-built XML. tests/fixtures/word/empty-shell.docx
is tests/fixtures/word/one-page.docx (WP-02's Word-authored fixture) with
its w:body's content children programmatically removed (keeping its
sectPr, styles, and numbering.xml intact) — an empty template shell for
the round-trip acceptance test, not a claim about a Word-output nuance,
so it does not need to be Word-authored itself (see fixtures/README.md).
"""

from __future__ import annotations

import glob
import hashlib
import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any
from unittest import mock
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paths, projection
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

# Every guarded call runs lock_status's two-sample quiesce check; shrink it
# for the whole test module so the suite does not eat ~1.5s per call.
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class _TempFixtureCase(unittest.TestCase):
    """Copies one fixture into an isolated allowed-roots temp dir per test."""

    fixture_name = "word/empty-shell.docx"

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
    def test_all_three_wp04_tools_are_registered(self):
        # WP-06 (issue #28) adds replace_text/format_text to MUTATING_TOOLS
        # in the same commit that adds those tools -- see
        # tests/unit/test_text_edit.py's own registration test for that
        # pair; this test stays scoped to WP-04's three markdown-mutation
        # tools, which this module owns.
        self.assertLessEqual(
            frozenset({"replace_body_markdown", "replace_range_markdown", "append_markdown"}),
            MUTATING_TOOLS,
        )


# ---------------------------------------------------------------------------
# Atomic write: OPC_INVALID (pre-replace) and VERIFICATION_FAILED (post-
# replace, .jsbak restore) — exercised directly against atomic_replace_docx_parts
# so each failure surface is deterministic rather than depending on a real
# filesystem fault or a genuine content mismatch.
# ---------------------------------------------------------------------------


class AtomicWriteTests(_TempFixtureCase):
    def test_corrupted_temp_write_is_caught_by_opc_valid_original_untouched(self):
        before_hash = _sha256(self.target)

        def _post_verify(_path):
            self.fail("post_verify must not run when opc_valid already rejected the temp file")

        with self.assertRaises(VerifyError) as cm:
            mutations.atomic_replace_docx_parts(
                self.target,
                {},
                post_verify=_post_verify,
                _corrupt_temp_for_test=True,
            )
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.OPC_INVALID)
        self.assertEqual(_sha256(self.target), before_hash, "original must be byte-identical after a rejected write")
        self.assertEqual(glob.glob(str(self.target) + "*.jsbak"), [])
        self.assertFalse((self.target.parent / (self.target.name + ".jsbak")).exists())

    def test_failed_post_verify_restores_original_from_jsbak(self):
        before_hash = _sha256(self.target)

        with zipfile.ZipFile(self.target) as zf:
            doc_bytes = zf.read(projection.DEFAULT_PART)
        # A legitimate (opc_valid-passing) but different rewrite of
        # document.xml, so the write genuinely replaces the file before
        # post_verify deliberately fails it.
        overrides = {projection.DEFAULT_PART: doc_bytes}

        def _post_verify_always_fails(_path):
            raise ValueError("simulated verification failure")

        with self.assertRaises(VerifyError) as cm:
            mutations.atomic_replace_docx_parts(self.target, overrides, post_verify=_post_verify_always_fails)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)
        self.assertEqual(_sha256(self.target), before_hash, "original must be restored byte-identical from .jsbak")
        self.assertFalse((self.target.parent / (self.target.name + ".jsbak")).exists(), "jsbak must be consumed by the restore")

    def test_successful_write_deletes_jsbak(self):
        with zipfile.ZipFile(self.target) as zf:
            doc_bytes = zf.read(projection.DEFAULT_PART)
        overrides = {projection.DEFAULT_PART: doc_bytes}
        calls = []

        def _post_verify_ok(path):
            calls.append(path)

        mutations.atomic_replace_docx_parts(self.target, overrides, post_verify=_post_verify_ok)
        self.assertEqual(len(calls), 1)
        self.assertFalse((self.target.parent / (self.target.name + ".jsbak")).exists())


# ---------------------------------------------------------------------------
# Lock guard layers 0/4 (acquire_lock/release_lock/remote_checkout) and
# layer 3 (conflict_copy_sweep) — issue #28 WP-10.
# ---------------------------------------------------------------------------


class LockGuardLayersTests(_TempFixtureCase):
    def _overrides_for_noop_write(self) -> dict[str, bytes]:
        with zipfile.ZipFile(self.target) as zf:
            doc_bytes = zf.read(projection.DEFAULT_PART)
        return {projection.DEFAULT_PART: doc_bytes}

    def test_remote_checkout_is_a_documented_noop(self):
        # WP-10's own instruction: a seam for a future Microsoft Graph
        # checkout call, MUST NOT call Graph today. Asserting its return
        # shape (rather than the absence of a network call, which cannot
        # be observed from here) locks in that it is inert: a plain dict,
        # no side effects, no exception.
        result = mutations.remote_checkout(self.target)
        self.assertEqual(result, {"remote_checkout": "not_implemented", "path": str(self.target)})

    def test_successful_write_leaves_no_jsclaim_behind(self):
        claim_path = self.target.parent / (self.target.name + ".jsclaim")
        mutations.atomic_replace_docx_parts(self.target, self._overrides_for_noop_write(), post_verify=lambda _p: None)
        self.assertFalse(claim_path.exists())

    def test_jsclaim_released_even_when_post_verify_fails(self):
        claim_path = self.target.parent / (self.target.name + ".jsclaim")

        def _post_verify_always_fails(_path):
            raise ValueError("simulated verification failure")

        with self.assertRaises(VerifyError):
            mutations.atomic_replace_docx_parts(
                self.target, self._overrides_for_noop_write(), post_verify=_post_verify_always_fails
            )
        self.assertFalse(claim_path.exists(), "release_lock must run in a finally regardless of write outcome")

    def test_preexisting_jsclaim_refuses_as_docx_locked_without_touching_original(self):
        # Simulates a genuine concurrent write (or a stale claim left by a
        # crashed prior process): acquire_lock must fail atomically via
        # O_EXCL rather than racing a check-then-create.
        claim_path = self.target.parent / (self.target.name + ".jsclaim")
        claim_path.write_text("12345")
        before_hash = _sha256(self.target)
        try:
            with self.assertRaises(VerifyError) as cm:
                mutations.atomic_replace_docx_parts(
                    self.target, self._overrides_for_noop_write(), post_verify=lambda _p: None
                )
            self.assertEqual(cm.exception.envelope.error_code, ErrorCode.DOCX_LOCKED)
            self.assertEqual(_sha256(self.target), before_hash, "original must be untouched when the claim is refused")
        finally:
            claim_path.unlink(missing_ok=True)

    def test_acquire_lock_calls_remote_checkout_first(self):
        calls = []
        with mock.patch("verified_docx_mcp.mutations.remote_checkout", side_effect=calls.append):
            claim_path = mutations.acquire_lock(self.target)
        self.assertEqual(calls, [self.target])
        mutations.release_lock(claim_path)


class ConflictCopyPatternTests(unittest.TestCase):
    """Direct tests of _matches_conflict_copy_pattern against every naming
    form core/document-backend-protocol.md §4 lists verbatim, plus the
    "matches none of them" case that must fall through to
    sibling_files_changed instead.

    machine_names is passed EXPLICITLY (never the real
    mutations._local_machine_names()) so these tests are deterministic
    regardless of what machine/CI runs them -- the real resolver has its
    own dedicated tests below (LocalMachineNamesTests)."""

    stem = "Proposal-Final"
    machine_names = frozenset({"JMS-MacBook"})

    def test_machine_suffix_form(self):
        self.assertTrue(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final-JMS-MacBook.docx", self.machine_names)
        )

    def test_machine_suffix_form_is_case_insensitive(self):
        self.assertTrue(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final-jms-macbook.docx", self.machine_names)
        )

    def test_numbered_paren_form(self):
        self.assertTrue(mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final (2).docx", frozenset()))

    def test_dash_copy_form(self):
        self.assertTrue(mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final-Copy.docx", frozenset()))

    def test_space_dash_copy_form(self):
        self.assertTrue(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final - Copy.docx", frozenset())
        )

    def test_conflict_substring_anywhere(self):
        self.assertTrue(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final-Conflict1.docx", frozenset())
        )

    def test_conflicted_copy_substring(self):
        self.assertTrue(
            mutations._matches_conflict_copy_pattern(
                self.stem, "Proposal-Final's conflicted copy 2026-09-10.docx", frozenset()
            )
        )

    def test_dash_suffix_not_matching_any_known_machine_name_does_not_match(self):
        # issue #28 WP-10 review fix: a plain "-<anything>.docx" sibling
        # is NOT this pattern any more unless <anything> IS a known
        # machine name -- "proposal-final.docx" next to "proposal.docx"
        # is the exact false positive the plan's own text warns about by
        # name, and must fall through to sibling_files_changed instead.
        self.assertFalse(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final-final.docx", self.machine_names)
        )

    def test_dash_suffix_matching_a_DIFFERENT_machines_name_does_not_match(self):
        # Proves this is genuinely gated on THIS machine's own name(s),
        # not "any token" -- a name that happens to look machine-ish but
        # is not in machine_names still does not match.
        self.assertFalse(
            mutations._matches_conflict_copy_pattern(
                self.stem, "Proposal-Final-Some-Other-Laptop.docx", self.machine_names
            )
        )

    def test_underscore_suffix_does_not_match_any_form(self):
        self.assertFalse(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final_backup.docx", self.machine_names)
        )

    def test_different_stem_entirely_does_not_match(self):
        self.assertFalse(
            mutations._matches_conflict_copy_pattern(self.stem, "Other-Document (2).docx", self.machine_names)
        )

    def test_own_name_unchanged_does_not_match(self):
        self.assertFalse(
            mutations._matches_conflict_copy_pattern(self.stem, "Proposal-Final.docx", self.machine_names)
        )


class NormalizeMachineNameTests(unittest.TestCase):
    def test_strips_apostrophe_and_collapses_spaces(self):
        self.assertEqual(mutations._normalize_machine_name("Michael's MacBook Pro"), "Michaels-MacBook-Pro")

    def test_curly_apostrophe_also_stripped(self):
        self.assertEqual(mutations._normalize_machine_name("Michael’s MacBook Pro"), "Michaels-MacBook-Pro")

    def test_already_hyphenated_hostname_is_unchanged(self):
        self.assertEqual(mutations._normalize_machine_name("Michaels-MacBook-Pro"), "Michaels-MacBook-Pro")

    def test_underscore_runs_collapse_to_one_dash(self):
        self.assertEqual(mutations._normalize_machine_name("host__name"), "host-name")


class LocalMachineNamesTests(unittest.TestCase):
    """_local_machine_names() itself is best-effort and platform-specific
    (macOS's scutil); these tests exercise its own failure-tolerance
    (never raises) without depending on what this test happens to run
    on."""

    def test_never_raises_when_scutil_and_hostname_both_fail(self):
        with (
            mock.patch("verified_docx_mcp.mutations.subprocess.run", side_effect=OSError("no such command")),
            mock.patch("verified_docx_mcp.mutations.socket.gethostname", side_effect=OSError("no hostname")),
        ):
            self.assertEqual(mutations._local_machine_names(), frozenset())

    def test_uses_scutil_output_when_available(self):
        def _fake_run(args, **_kwargs):
            result = mock.Mock()
            if args[-1] == "ComputerName":
                result.returncode = 0
                result.stdout = "Michael's MacBook Pro\n"
            else:
                result.returncode = 1
                result.stdout = ""
            return result

        with (
            mock.patch("verified_docx_mcp.mutations.subprocess.run", side_effect=_fake_run),
            mock.patch("verified_docx_mcp.mutations.socket.gethostname", return_value="unrelated-host"),
        ):
            names = mutations._local_machine_names()
        self.assertIn("Michaels-MacBook-Pro", names)
        self.assertIn("unrelated-host", names)

    def test_falls_back_to_hostname_when_scutil_is_absent(self):
        with (
            mock.patch("verified_docx_mcp.mutations.subprocess.run", side_effect=OSError("no such command")),
            mock.patch("verified_docx_mcp.mutations.socket.gethostname", return_value="my-host.local"),
        ):
            names = mutations._local_machine_names()
        self.assertEqual(names, frozenset({"my-host"}))


class ConflictCopySweepTests(_TempFixtureCase):
    def test_no_siblings_reports_clean(self):
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)
        self.assertEqual(sweep, {"conflict_copy_detected": False, "conflict_copies": [], "sibling_files_changed": []})

    def test_matching_sibling_flags_detected_without_raising(self):
        conflict_sibling = self.target.parent / (self.target.stem + "-Copy.docx")
        conflict_sibling.write_bytes(b"stub")
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)
        self.assertTrue(sweep["conflict_copy_detected"])
        self.assertEqual(sweep["conflict_copies"], [conflict_sibling.name])
        self.assertEqual(sweep["sibling_files_changed"], [])

    def test_non_matching_newer_sibling_is_info_only(self):
        info_sibling = self.target.parent / (self.target.stem + "_backup.docx")
        info_sibling.write_bytes(b"stub")
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)
        self.assertFalse(sweep["conflict_copy_detected"])
        self.assertEqual(sweep["conflict_copies"], [])
        self.assertEqual(sweep["sibling_files_changed"], [info_sibling.name])

    def test_real_machine_name_form_flags_but_a_plain_dash_suffix_does_not(self):
        # Issue #28 WP-10 review fix, both directions in one test so a
        # future change cannot pass by only checking one side again:
        # <stem>-<Machine>.docx flags ONLY when <Machine> is this
        # (mocked) machine's own name; an ordinary same-stem working file
        # like "<stem>-final.docx" -- the plan's own named counter-
        # example -- must land in sibling_files_changed, never
        # conflict_copies, even though both share the same "-suffix"
        # shape.
        with mock.patch("verified_docx_mcp.mutations._local_machine_names", return_value=frozenset({"TESTHOST"})):
            real_conflict = self.target.parent / (self.target.stem + "-TESTHOST.docx")
            real_conflict.write_bytes(b"stub")
            ordinary_sibling = self.target.parent / (self.target.stem + "-final.docx")
            ordinary_sibling.write_bytes(b"stub")

            sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)

        self.assertTrue(sweep["conflict_copy_detected"])
        self.assertEqual(sweep["conflict_copies"], [real_conflict.name])
        self.assertEqual(sweep["sibling_files_changed"], [ordinary_sibling.name])

    def test_sibling_older_than_since_ns_is_ignored(self):
        # since_ns in the far future -> even a genuine conflict-pattern
        # sibling sitting there from before this write began is not
        # attributed to this write.
        conflict_sibling = self.target.parent / (self.target.stem + "-Copy.docx")
        conflict_sibling.write_bytes(b"stub")
        far_future_ns = (int(conflict_sibling.stat().st_mtime) + 3600) * 1_000_000_000
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=far_future_ns)
        self.assertFalse(sweep["conflict_copy_detected"])
        self.assertEqual(sweep["conflict_copies"], [])
        self.assertEqual(sweep["sibling_files_changed"], [])

    def test_jsbak_jsclaim_and_tmp_siblings_never_flagged(self):
        (self.target.parent / (self.target.name + ".jsbak")).write_bytes(b"stub")
        (self.target.parent / (self.target.name + ".jsclaim")).write_bytes(b"stub")
        (self.target.parent / f"{self.target.stem}.tmp-abc123.docx").write_bytes(b"stub")
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)
        self.assertEqual(sweep, {"conflict_copy_detected": False, "conflict_copies": [], "sibling_files_changed": []})

    def test_owner_files_never_flagged(self):
        (self.target.parent / ("~$" + self.target.name[2:])).write_bytes(b"stub")
        (self.target.parent / f".~lock.{self.target.name}#").write_bytes(b"stub")
        sweep = mutations.conflict_copy_sweep(self.target, since_ns=0)
        self.assertEqual(sweep, {"conflict_copy_detected": False, "conflict_copies": [], "sibling_files_changed": []})

    def test_a_real_mutating_tool_call_surfaces_the_flag_and_still_applies(self):
        # Wires layer 3 all the way through a real mutating tool
        # (execute_append_markdown) rather than only unit-testing
        # conflict_copy_sweep in isolation: patches time.time_ns so
        # since_ns is far in the past, so a sibling already sitting next
        # to the target (planted before this call, for determinism) is
        # picked up exactly as a genuinely concurrent conflict copy
        # appearing mid-write would be. core/document-backend-protocol.md
        # §4: "the write itself succeeded" — applied must still be True.
        conflict_sibling = self.target.parent / (self.target.stem + "-Copy.docx")
        conflict_sibling.write_bytes(b"stub")
        with mock.patch("verified_docx_mcp.mutations.time.time_ns", return_value=0):
            evidence = mutations.execute_append_markdown(str(self.target), "New paragraph.\n")
        self.assertTrue(evidence["applied"])
        self.assertTrue(evidence["conflict_copy_detected"])
        self.assertEqual(evidence["conflict_copies"], [conflict_sibling.name])

    def test_a_real_mutating_tool_call_with_no_conflict_reports_flag_false(self):
        # The common case: conflict_copy_detected is always present (not
        # just on a hit), so a caller can assert its absence positively.
        evidence = mutations.execute_append_markdown(str(self.target), "New paragraph.\n")
        self.assertTrue(evidence["applied"])
        self.assertFalse(evidence["conflict_copy_detected"])
        self.assertNotIn("conflict_copies", evidence)
        self.assertNotIn("sibling_files_changed", evidence)


# ---------------------------------------------------------------------------
# Guard: DOCX_LOCKED and REVISION_CONFLICT (SYNC_IN_FLIGHT is covered by
# test_server.py's existing lock_status sync-quiesce tests at the
# lower level; this WP only adds the write-side consumption of that
# signal, exercised here via the owner-file / revision paths).
# ---------------------------------------------------------------------------


class GuardTests(_TempFixtureCase):
    def test_locked_file_refuses_before_any_temp_file_is_written(self):
        owner_file = self.target.parent / ("~$" + self.target.name[2:])
        owner_file.write_bytes(b"Michael Sutton")
        try:
            with self.assertRaises(VerifyError) as cm:
                mutations.execute_append_markdown(str(self.target), "New paragraph.\n")
            self.assertEqual(cm.exception.envelope.error_code, ErrorCode.DOCX_LOCKED)
            # No temp/.jsbak file was left behind by the (never-started) write.
            self.assertEqual(
                sorted(p.name for p in self.target.parent.iterdir()),
                sorted([self.target.name, owner_file.name]),
            )
        finally:
            owner_file.unlink()

    def test_stale_revision_before_refuses(self):
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_append_markdown(str(self.target), "New paragraph.\n", revision_before="deadbeef:deadbeef")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.REVISION_CONFLICT)

    def test_matching_revision_before_proceeds(self):
        current = projection.compute_revision(self.target)
        evidence = mutations.execute_append_markdown(str(self.target), "New paragraph.\n", revision_before=current["token"])
        self.assertTrue(evidence["applied"])


# ---------------------------------------------------------------------------
# Hazard scan against REAL Word-authored comment/tracked-change fixtures.
# ---------------------------------------------------------------------------


class CommentAnchorHazardTests(_TempFixtureCase):
    fixture_name = "revision/commented.docx"

    def test_refuses_without_force(self):
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_body_markdown(str(self.target), "Replacement text.\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.COMMENT_ANCHORS_IN_RANGE)
        self.assertEqual(cm.exception.envelope.diagnostics["comment_ids"], ["0"])

    def test_force_removes_anchors_and_reports_orphaned_ids(self):
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Replacement text.\n", force=True)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["orphaned_comment_ids"], ["0"])
        markdown_after, _, _ = projection.read_document_markdown(self.target)
        self.assertEqual(markdown_after.strip(), "Replacement text.")


class TrackedChangeHazardTests(_TempFixtureCase):
    """Patches author.resolve_author_name to a name DIFFERENT from
    tracked.docx's real Word-authored author ("Michael Sutton") -- WP-07b-a's
    own-author exclusion (_check_hazards_or_raise) would otherwise treat
    this fixture's revisions as the server's own prior work whenever this
    suite happens to run on Michael Sutton's own machine, silently
    defeating the refusal these tests exist to prove — see
    test_tracked_changes.py's OverlappingWriteRefusalTests for the
    identical rationale."""

    fixture_name = "revision/tracked.docx"

    def setUp(self):
        super().setUp()
        # mutations.py imports resolve_author_name lazily (a deferred,
        # function-local import — see execute_replace_body_markdown), so
        # patching the SOURCE (author.resolve_author_name) is what every
        # such deferred import actually resolves at call time.
        patcher = mock.patch("verified_docx_mcp.author.resolve_author_name", return_value="A Different Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_refuses_without_force(self):
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_body_markdown(str(self.target), "Replacement text.\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.TRACKED_CHANGES_PRESENT)

    def test_force_proceeds(self):
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Replacement text.\n", force=True)
        self.assertTrue(evidence["applied"])
        # No comment anchors were involved (only tracked changes), so
        # orphaned_comment_ids is omitted rather than reported as an empty list.
        self.assertEqual(evidence.get("orphaned_comment_ids", []), [])


# ---------------------------------------------------------------------------
# replace_body_markdown: the acceptance round trip.
# ---------------------------------------------------------------------------


class ReplaceBodyMarkdownRoundTripTests(_TempFixtureCase):
    fixture_name = "word/empty-shell.docx"

    def test_round_trip_modulo_whitespace(self):
        section_md = (FIXTURES / "markdown" / "section.md").read_text(encoding="utf-8")
        evidence = mutations.execute_replace_body_markdown(str(self.target), section_md)

        for key in ("applied", "match_count", "rung", "before", "after", "revision_before", "revision_after", "audit_logged"):
            self.assertIn(key, evidence, f"missing evidence key: {key}")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(evidence["rung"], 4)
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertTrue(evidence["audit_logged"])

        markdown_after, _, _ = projection.read_document_markdown(self.target)
        normalize = lambda s: " ".join(s.split())
        self.assertEqual(normalize(markdown_after), normalize(section_md))
        # The evidence's own "after" must match what a fresh read reports.
        self.assertEqual(normalize(evidence["after"]), normalize(markdown_after))

    # STYLE_NOT_FOUND itself is exercised directly against StyleContext in
    # test_markdown_to_ooxml.py — CommonMark ATX headings cap at level 6,
    # and empty-shell.docx's source (one-page.docx) defines Heading1..
    # Heading9, so no markdown heading this fixture can express is
    # actually missing a style; a real "no style at this level" case needs
    # a document that omits one of levels 1-6, which none of this WP's
    # fixtures do.


# ---------------------------------------------------------------------------
# replace_range_markdown: targets only the named section.
# ---------------------------------------------------------------------------


class ReplaceRangeMarkdownTests(_TempFixtureCase):
    fixture_name = "sections.docx"

    def test_replaces_only_the_target_section(self):
        sections_before = projection.find_sections_impl(self.target)
        keys = {s["section_key"] for s in sections_before}
        self.assertEqual(keys, {"overview-1", "background-1", "next-steps-1"})

        new_md = "## Background\n\nUpdated background content for the WP-04 test.\n"
        evidence = mutations.execute_replace_range_markdown(str(self.target), "background-1", new_md)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], 3)
        self.assertIn("Background text.", evidence["before"])
        self.assertIn("Updated background content", evidence["after"])

        sections_after = projection.find_sections_impl(self.target)
        keys_after = {s["section_key"] for s in sections_after}
        self.assertEqual(keys_after, {"overview-1", "background-1", "next-steps-1"})

        full_markdown, _, _ = projection.read_document_markdown(self.target)
        self.assertIn("Updated background content", full_markdown)
        self.assertNotIn("Background text.", full_markdown)
        self.assertIn("Some overview text.", full_markdown)
        self.assertIn("Next steps text.", full_markdown)

    def test_unknown_section_key_raises_section_not_found(self):
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_range_markdown(str(self.target), "nope-1", "# X\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.SECTION_NOT_FOUND)


class ReplaceRangeMarkdownTextboxKeyTests(_TempFixtureCase):
    """find_sections (PR #2) lists text-box sub-scopes alongside heading
    sections, both keyed by section_key. A textbox-<n> key must never
    silently fall through to a body-range write — this tool has no write
    path for text-box content, so it must refuse explicitly rather than
    let locate_section_range's heading-only scan near-miss its way to an
    ambiguous result."""

    fixture_name = "textbox.docx"

    def test_textbox_section_key_is_refused_not_silently_written(self):
        sections = projection.find_sections_impl(self.target)
        self.assertEqual([s["section_key"] for s in sections], ["textbox-1"])

        before_markdown, _, _ = projection.read_document_markdown(self.target)
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_range_markdown(str(self.target), "textbox-1", "Replacement.\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("text box", cm.exception.envelope.message)

        # Refused before any write: the document is untouched.
        after_markdown, _, _ = projection.read_document_markdown(self.target)
        self.assertEqual(before_markdown, after_markdown)


# ---------------------------------------------------------------------------
# append_markdown
# ---------------------------------------------------------------------------


class AppendMarkdownTests(_TempFixtureCase):
    fixture_name = "sections.docx"

    def test_appends_after_existing_content(self):
        evidence = mutations.execute_append_markdown(str(self.target), "## Appendix\n\nAppended content.\n")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], 4)
        full_markdown, _, _ = projection.read_document_markdown(self.target)
        self.assertTrue(full_markdown.strip().endswith("Appended content."))
        self.assertIn("Next steps text.", full_markdown)  # existing content preserved


# ---------------------------------------------------------------------------
# Numbering created from scratch when word/numbering.xml is absent.
# ---------------------------------------------------------------------------


class MissingNumberingPartTests(_TempFixtureCase):
    fixture_name = "word/empty-shell.docx"

    def setUp(self):
        super().setUp()
        # Strip numbering.xml, its rels entry, and its content-types
        # override from the copied fixture so StyleContext.build sees no
        # existing numbering part.
        with zipfile.ZipFile(self.target) as zin:
            items = {i.filename: zin.read(i) for i in zin.infolist()}
        del items["word/numbering.xml"]

        rels_bytes = items[projection._rels_path_for(projection.DEFAULT_PART)]
        mutations._capture_source_namespaces(rels_bytes)
        rels_root = ET.fromstring(rels_bytes)
        for rel in list(rels_root):
            if rel.get("Target") == "numbering.xml":
                rels_root.remove(rel)
        items[projection._rels_path_for(projection.DEFAULT_PART)] = ET.tostring(rels_root, encoding="utf-8")

        ct_bytes = items["[Content_Types].xml"]
        mutations._capture_source_namespaces(ct_bytes)
        ct_root = ET.fromstring(ct_bytes)
        for child in list(ct_root):
            if child.get("PartName") == "/word/numbering.xml":
                ct_root.remove(child)
        items["[Content_Types].xml"] = ET.tostring(ct_root, encoding="utf-8")

        stripped = self.target.with_suffix(".stripped.docx")
        with zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in items.items():
                zout.writestr(name, data)
        stripped.replace(self.target)

    def test_bullet_list_creates_numbering_part_and_stays_opc_valid(self):
        with zipfile.ZipFile(self.target) as zf:
            self.assertNotIn("word/numbering.xml", zf.namelist())

        evidence = mutations.execute_replace_body_markdown(str(self.target), "- one\n- two\n")
        self.assertTrue(evidence["applied"])

        with zipfile.ZipFile(self.target) as zf:
            names = set(zf.namelist())
            self.assertIn("word/numbering.xml", names)
            ct_root = ET.fromstring(zf.read("[Content_Types].xml"))
            self.assertTrue(
                any(c.get("PartName") == "/word/numbering.xml" for c in ct_root if c.tag.endswith("Override"))
            )
            rels_root = ET.fromstring(zf.read(projection._rels_path_for(projection.DEFAULT_PART)))
            self.assertTrue(any(r.get("Target") == "numbering.xml" for r in rels_root))

        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)


# ---------------------------------------------------------------------------
# PR #3 review, BLOCKING fix: namespace declarations must survive a write.
#
# Word 16 refused a written file: mc:Ignorable survived on the root naming
# prefixes (w14, w15, w16se, ...) that ElementTree's serializer no longer
# declared anywhere, because nothing in the WRITTEN tree happened to still
# use those namespaces. Reproduced here against the same two real,
# Word-authored fixtures the review used (tracked.docx, sections.docx),
# through the real tools, with explicit before/after declaration counts —
# plus a dedicated opc_valid test proving the new rule actually rejects a
# deliberately broken file (a check that cannot fail is not a check).
# ---------------------------------------------------------------------------

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _root_namespace_decl_count(xml_bytes: bytes) -> int:
    # Skip the "<?xml ...?>" declaration first — it contains its own ">"
    # (the one closing "?>"), which would otherwise be mistaken for the
    # end of the ROOT element's opening tag.
    after_prolog = xml_bytes.split(b"?>", 1)[-1] if xml_bytes.startswith(b"<?xml") else xml_bytes
    end = after_prolog.find(b">")
    open_tag = after_prolog[: end + 1] if end != -1 else after_prolog
    return len(re.findall(rb'xmlns(?::[A-Za-z0-9]+)?="[^"]*"', open_tag))


def _mc_ignorable_undeclared(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    ignorable = root.get(f"{{{_MC_NS}}}Ignorable")
    if not ignorable:
        return []
    declared = mutations._declared_prefixes(xml_bytes)
    return [tok for tok in ignorable.split() if tok not in declared]


class NamespaceDeclarationPreservationTests(_TempFixtureCase):
    fixture_name = "revision/tracked.docx"

    def _read_document_xml(self) -> bytes:
        with zipfile.ZipFile(self.target) as zf:
            return zf.read(projection.DEFAULT_PART)

    def test_replace_body_markdown_keeps_every_mc_ignorable_prefix_declared(self):
        before_bytes = self._read_document_xml()
        before_count = _root_namespace_decl_count(before_bytes)
        self.assertGreater(before_count, 20, "sanity: the real fixture's root should carry many xmlns decls")
        self.assertEqual(_mc_ignorable_undeclared(before_bytes), [])

        mutations.execute_replace_body_markdown(str(self.target), "Replacement text.\n", force=True)

        after_bytes = self._read_document_xml()
        after_count = _root_namespace_decl_count(after_bytes)
        undeclared = _mc_ignorable_undeclared(after_bytes)
        self.assertEqual(
            undeclared, [], f"mc:Ignorable names undeclared prefixes after the write: {undeclared} "
            f"(decl count {before_count} -> {after_count})"
        )
        # opc_valid itself must now catch this class of bug (the second half
        # of this fix, per the review: the missing CHECK, not just the fix).
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)


class NamespaceDeclarationPreservationSectionsTests(_TempFixtureCase):
    fixture_name = "sections.docx"

    def _read_document_xml(self) -> bytes:
        with zipfile.ZipFile(self.target) as zf:
            return zf.read(projection.DEFAULT_PART)

    def test_replace_range_markdown_keeps_every_mc_ignorable_prefix_declared(self):
        before_bytes = self._read_document_xml()
        before_count = _root_namespace_decl_count(before_bytes)
        self.assertGreater(before_count, 20)
        self.assertEqual(_mc_ignorable_undeclared(before_bytes), [])

        mutations.execute_replace_range_markdown(
            str(self.target), "background-1", "## Background\n\nUpdated.\n"
        )

        after_bytes = self._read_document_xml()
        after_count = _root_namespace_decl_count(after_bytes)
        undeclared = _mc_ignorable_undeclared(after_bytes)
        self.assertEqual(
            undeclared, [], f"mc:Ignorable names undeclared prefixes after the write: {undeclared} "
            f"(decl count {before_count} -> {after_count})"
        )
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)

    def test_append_markdown_keeps_every_mc_ignorable_prefix_declared(self):
        before_bytes = self._read_document_xml()
        before_count = _root_namespace_decl_count(before_bytes)

        mutations.execute_append_markdown(str(self.target), "## Appendix\n\nMore.\n")

        after_bytes = self._read_document_xml()
        after_count = _root_namespace_decl_count(after_bytes)
        undeclared = _mc_ignorable_undeclared(after_bytes)
        self.assertEqual(
            undeclared, [], f"mc:Ignorable names undeclared prefixes after the write: {undeclared} "
            f"(decl count {before_count} -> {after_count})"
        )
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)

    def test_the_r_namespace_specifically_survives_a_write(self):
        # The review's second-order finding: "r" (relationships) was in
        # the lost set independent of mc:Ignorable — every hyperlink/image
        # r:id depends on it being declared, whether or not anything in
        # THIS particular write happens to use one.
        before_bytes = self._read_document_xml()
        before_decls = mutations._declared_prefixes(before_bytes)
        self.assertIn("r", before_decls)

        mutations.execute_replace_range_markdown(
            str(self.target), "background-1", "## Background\n\nUpdated, no hyperlink here.\n"
        )

        after_bytes = self._read_document_xml()
        after_decls = mutations._declared_prefixes(after_bytes)
        self.assertIn("r", after_decls, "the relationships namespace prefix must survive even when unused by this write")


class OpcValidCatchesUndeclaredMcIgnorablePrefixesTests(_TempFixtureCase):
    fixture_name = "revision/tracked.docx"

    def test_opc_valid_fails_on_a_deliberately_broken_namespace_declaration(self):
        with zipfile.ZipFile(self.target) as zf:
            raw = zf.read(projection.DEFAULT_PART)
        # Reproduce the ORIGINAL bug shape directly: serialize via bare
        # ET.tostring with NO namespace preservation, so mc:Ignorable
        # survives while most of its prefixes lose their declaration —
        # exactly what _ensure_namespace_declarations now prevents in the
        # real write path. This proves opc_valid's new rule actually
        # fires on a file it did not check before this PR (a check that
        # cannot fail is not a check).
        root = ET.fromstring(raw)
        broken = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(root, encoding="utf-8")
        undeclared_in_broken = _mc_ignorable_undeclared(broken)
        self.assertTrue(undeclared_in_broken, "the naive re-serialization must reproduce the original bug")

        with zipfile.ZipFile(self.target) as zin:
            items = {i.filename: zin.read(i) for i in zin.infolist()}
        items[projection.DEFAULT_PART] = broken
        broken_path = self.target.with_name("broken.docx")
        with zipfile.ZipFile(broken_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in items.items():
                zout.writestr(name, data)

        valid, problems = mutations.opc_valid(broken_path)
        self.assertFalse(valid)
        self.assertTrue(
            any("mc:Ignorable" in p and "undeclared" in p for p in problems), problems
        )

    def test_opc_valid_passes_when_declarations_are_intact(self):
        valid, problems = mutations.opc_valid(self.target)
        self.assertTrue(valid, problems)


# ---------------------------------------------------------------------------
# PR #3 review, should-fix #2: half-anchor pairs when a comment's span
# crosses the target range's boundary.
# ---------------------------------------------------------------------------


class HalfAnchorCrossingTests(unittest.TestCase):
    """No real Word-authored fixture combines a heading-delimited section
    with a comment whose span crosses it (deliberately not fabricated —
    same reasoning as tests/fixtures/README.md's "not produced this WP"
    note: hand-assembling that combination would be exactly the
    hand-built-OOXML anti-pattern this repo's fixtures avoid). Instead,
    this takes the REAL commentRangeStart/commentRangeEnd/commentReference
    elements from commented.docx — genuine Word output, moved intact, not
    retyped — and splits them across two separate paragraph elements to
    reproduce the topology a real multi-paragraph comment span crossing a
    section boundary would have. Only the paragraph ARRANGEMENT is a test
    construction; every anchor element's own shape is real."""

    def _load_real_comment_children(self) -> list[Any]:
        with zipfile.ZipFile(FIXTURES / "revision" / "commented.docx") as zf:
            raw = zf.read(projection.DEFAULT_PART)
        doc_root = ET.fromstring(raw)
        body = next(c for c in doc_root if projection._ln(c) == "body")
        paragraph = next(c for c in body if projection._ln(c) == "p")
        children = list(paragraph)
        # Sanity: the known real shape (commentRangeStart, "This" run,
        # commentRangeEnd, the commentReference run, the trailing text
        # run) — five children. If commented.docx's own shape ever
        # changes this fixture-derived test should fail loudly here
        # rather than silently testing something else.
        self.assertEqual(len(children), 5)
        return children

    def _build_two_paragraphs(self) -> tuple[Any, Any]:
        children = self._load_real_comment_children()
        w_p = f"{{{projection.W_NS}}}p"
        target_para = ET.Element(w_p)  # simulates: inside the section being replaced
        surviving_para = ET.Element(w_p)  # simulates: a paragraph elsewhere, left alone
        for child in children[:2]:  # commentRangeStart, "This"
            target_para.append(child)
        for child in children[2:]:  # commentRangeEnd, the commentReference run, trailing text
            surviving_para.append(child)
        return target_para, surviving_para

    def test_hazard_scan_finds_the_id_from_only_the_start_half(self):
        target_para, _surviving_para = self._build_two_paragraphs()
        hazards = mutations._scan_range_hazards([target_para])
        self.assertEqual(hazards["comment_ids"], ["0"])

    def test_strip_removes_the_other_half_from_the_surviving_range(self):
        target_para, surviving_para = self._build_two_paragraphs()
        hazards = mutations._scan_range_hazards([target_para])

        before_tags = {
            projection._ln(n) for n in surviving_para.iter() if projection._ln(n) in mutations._COMMENT_ANCHOR_TAGS
        }
        self.assertEqual(before_tags, {"commentRangeEnd", "commentReference"})

        mutations._strip_comment_anchors_by_id([surviving_para], set(hazards["comment_ids"]))

        after_tags = {
            projection._ln(n) for n in surviving_para.iter() if projection._ln(n) in mutations._COMMENT_ANCHOR_TAGS
        }
        self.assertEqual(after_tags, set(), "no unpaired anchor may remain in the surviving range")
        # The trailing text run (not part of the comment anchor at all)
        # must survive untouched.
        self.assertTrue(any((t.text or "").strip() for t in surviving_para.iter(f"{{{projection.W_NS}}}t")))

    def test_strip_is_a_no_op_when_the_id_does_not_appear(self):
        _target_para, surviving_para = self._build_two_paragraphs()
        before = ET.tostring(surviving_para)
        mutations._strip_comment_anchors_by_id([surviving_para], {"does-not-exist"})
        after = ET.tostring(surviving_para)
        self.assertEqual(before, after)


class HalfAnchorCrossingIntegrationTests(_TempFixtureCase):
    """Same fix, exercised end-to-end through execute_replace_range_markdown
    against a real, headed document (sections.docx) with the real comment
    anchors from commented.docx spliced in — one half inside the
    "background-1" section, the other half in "Next Steps" (outside it) —
    so the section boundary genuinely crosses the comment's span."""

    fixture_name = "sections.docx"

    def setUp(self):
        super().setUp()
        with zipfile.ZipFile(FIXTURES / "revision" / "commented.docx") as zf:
            comment_raw = zf.read(projection.DEFAULT_PART)
        comment_root = ET.fromstring(comment_raw)
        comment_body = next(c for c in comment_root if projection._ln(c) == "body")
        comment_paragraph = next(c for c in comment_body if projection._ln(c) == "p")
        comment_children = list(comment_paragraph)
        self.assertEqual(len(comment_children), 5)

        w_p = f"{{{projection.W_NS}}}p"
        start_para = ET.Element(w_p)
        for child in comment_children[:2]:
            start_para.append(child)
        end_para = ET.Element(w_p)
        for child in comment_children[2:]:
            end_para.append(child)

        with zipfile.ZipFile(self.target) as zf:
            names = {i.filename: i for i in zf.infolist()}
            doc_bytes = zf.read(projection.DEFAULT_PART)
            other_parts = {n: zf.read(n) for n in names if n != projection.DEFAULT_PART}
        doc_decls = mutations._capture_source_namespaces(doc_bytes)
        doc_root = ET.fromstring(doc_bytes)
        body = next(c for c in doc_root if projection._ln(c) == "body")
        body_children = list(body)
        # sections.docx: [Overview(h1), Overview text, Background(h1... h2),
        # Background text, Next Steps(h1), Next steps text, sectPr].
        background_text_idx = next(
            i for i, c in enumerate(body_children)
            if projection._ln(c) == "p" and "Background text." in "".join(t.text or "" for t in c.iter(f"{{{projection.W_NS}}}t"))
        )
        next_steps_text_idx = next(
            i for i, c in enumerate(body_children)
            if projection._ln(c) == "p" and "Next steps text." in "".join(t.text or "" for t in c.iter(f"{{{projection.W_NS}}}t"))
        )
        # start_para goes INSIDE the background-1 range (right after its
        # own text paragraph); end_para goes AFTER "Next steps text.",
        # i.e. outside background-1's range entirely.
        body.insert(background_text_idx + 1, start_para)
        body_children = list(body)  # indices shifted by the insert above
        next_steps_text_idx = next(
            i for i, c in enumerate(body_children)
            if projection._ln(c) == "p" and "Next steps text." in "".join(t.text or "" for t in c.iter(f"{{{projection.W_NS}}}t"))
        )
        body.insert(next_steps_text_idx + 1, end_para)

        new_doc_bytes = mutations._serialize_xml(doc_root, doc_decls)
        with zipfile.ZipFile(self.target, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in names.items():
                zout.writestr(data, new_doc_bytes if name == projection.DEFAULT_PART else other_parts[name])

    def test_force_removes_both_halves_and_reports_the_id(self):
        sections_before = projection.find_sections_impl(self.target)
        self.assertIn("background-1", {s["section_key"] for s in sections_before})

        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_range_markdown(str(self.target), "background-1", "## Background\n\nUpdated.\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.COMMENT_ANCHORS_IN_RANGE)

        evidence = mutations.execute_replace_range_markdown(
            str(self.target), "background-1", "## Background\n\nUpdated.\n", force=True
        )
        self.assertEqual(evidence["orphaned_comment_ids"], ["0"])

        with zipfile.ZipFile(self.target) as zf:
            final_doc = ET.fromstring(zf.read(projection.DEFAULT_PART))
        remaining = [
            projection._ln(n) for n in final_doc.iter() if projection._ln(n) in mutations._COMMENT_ANCHOR_TAGS
        ]
        self.assertEqual(remaining, [], "no unpaired (or paired) anchor may remain for the removed comment")


# ---------------------------------------------------------------------------
# track_changes=True (issue #28 WP-07b-a) on the three markdown-mutation
# tools: old content marked deleted (kept, wrapped in w:del) rather than
# removed, new content wrapped in w:ins, evidence gains track_changes/
# revision_ids.
# ---------------------------------------------------------------------------


class MarkdownTrackChangesTests(_TempFixtureCase):
    fixture_name = "word/empty-shell.docx"

    def setUp(self):
        super().setUp()
        patcher = mock.patch("verified_docx_mcp.author.resolve_author_name", return_value="Jane Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def _document_xml(self) -> str:
        with zipfile.ZipFile(self.target) as zf:
            return zf.read(projection.DEFAULT_PART).decode("utf-8")

    def test_replace_body_markdown_tracked_wraps_new_content_in_ins(self):
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Hello world.\n", track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertEqual(len(evidence["revision_ids"]), 1)
        doc = self._document_xml()
        self.assertIn('<w:ins w:id="', doc)
        self.assertIn('w:author="Jane Reviewer"', doc)
        self.assertNotIn("<w:del", doc, "an empty shell has no prior content to mark deleted")
        after_md, _, _ = projection.read_document_markdown(self.target)
        self.assertEqual(after_md.strip(), "Hello world.")

    def test_replace_body_markdown_tracked_marks_old_content_deleted_and_keeps_it(self):
        mutations.execute_replace_body_markdown(str(self.target), "First version.\n")
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Second version.\n", track_changes=True)
        self.assertTrue(evidence["applied"])
        doc = self._document_xml()
        self.assertIn("<w:delText>First version.</w:delText>", doc, "old content stays, marked deleted")
        self.assertIn("<w:t>Second version.</w:t>", doc)
        # Reading back as CURRENT text shows only the new content -- w:del
        # excluded, w:ins included, unchanged read-side contract.
        after_md, _, _ = projection.read_document_markdown(self.target)
        self.assertEqual(after_md.strip(), "Second version.")

    def test_append_markdown_tracked_wraps_only_new_content(self):
        mutations.execute_replace_body_markdown(str(self.target), "Existing paragraph.\n")
        evidence = mutations.execute_append_markdown(str(self.target), "Appended paragraph.\n", track_changes=True)
        self.assertTrue(evidence["track_changes"])
        doc = self._document_xml()
        self.assertIn("<w:t>Existing paragraph.</w:t>", doc, "pre-existing content is untouched by an append")
        self.assertNotIn("<w:del", doc)
        self.assertIn('<w:ins w:id="', doc)
        after_md, _, _ = projection.read_document_markdown(self.target)
        self.assertIn("Appended paragraph.", after_md)
        self.assertIn("Existing paragraph.", after_md)

    def test_replace_range_markdown_tracked_marks_section_deleted_and_inserts_after(self):
        section_md = (FIXTURES / "markdown" / "section.md").read_text(encoding="utf-8")
        mutations.execute_replace_body_markdown(str(self.target), section_md)
        sections = projection.find_sections_impl(self.target)
        key = next(s["section_key"] for s in sections if s["heading_text"] == "Overview")

        evidence = mutations.execute_replace_range_markdown(
            str(self.target), key, "## Overview\n\nRewritten overview.\n", track_changes=True
        )
        self.assertTrue(evidence["track_changes"])
        self.assertGreater(len(evidence["revision_ids"]), 0)
        doc = self._document_xml()
        self.assertIn("<w:del", doc)
        self.assertIn("<w:ins", doc)
        after_md, _, _ = projection.read_document_markdown(self.target)
        self.assertIn("Rewritten overview.", after_md)
        # The other, untouched sections survive unaffected.
        self.assertIn("Background", after_md)
        self.assertIn("Next Steps", after_md)

    def test_untracked_writes_have_no_track_changes_key(self):
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Plain write.\n")
        self.assertNotIn("track_changes", evidence)
        self.assertNotIn("revision_ids", evidence)


class TrackedChangeOwnAuthorExclusionTests(_TempFixtureCase):
    """WP-07b-a load-bearing note: a hazard scan must exclude the
    server's own previously-authored tracked change, or a second
    track_changes=True write over the first one's own content deadlocks."""

    fixture_name = "word/empty-shell.docx"

    def setUp(self):
        super().setUp()
        patcher = mock.patch("verified_docx_mcp.author.resolve_author_name", return_value="Jane Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_second_tracked_write_over_own_prior_tracked_change_does_not_deadlock(self):
        mutations.execute_replace_body_markdown(str(self.target), "First draft.\n", track_changes=True)
        # A second track_changes write over the document (which now
        # contains this server's own w:ins/w:del, all authored "Jane
        # Reviewer") must not refuse.
        evidence = mutations.execute_replace_body_markdown(str(self.target), "Second draft.\n", track_changes=True)
        self.assertTrue(evidence["applied"])


if __name__ == "__main__":
    unittest.main()
