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
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
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
        self.assertEqual(
            MUTATING_TOOLS,
            frozenset({"replace_body_markdown", "replace_range_markdown", "append_markdown"}),
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
        markdown_after, _ = projection.read_document_markdown(self.target)
        self.assertEqual(markdown_after.strip(), "Replacement text.")


class TrackedChangeHazardTests(_TempFixtureCase):
    fixture_name = "revision/tracked.docx"

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

        markdown_after, _ = projection.read_document_markdown(self.target)
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

        full_markdown, _ = projection.read_document_markdown(self.target)
        self.assertIn("Updated background content", full_markdown)
        self.assertNotIn("Background text.", full_markdown)
        self.assertIn("Some overview text.", full_markdown)
        self.assertIn("Next steps text.", full_markdown)

    def test_unknown_section_key_raises_section_not_found(self):
        with self.assertRaises(VerifyError) as cm:
            mutations.execute_replace_range_markdown(str(self.target), "nope-1", "# X\n")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.SECTION_NOT_FOUND)


# ---------------------------------------------------------------------------
# append_markdown
# ---------------------------------------------------------------------------


class AppendMarkdownTests(_TempFixtureCase):
    fixture_name = "sections.docx"

    def test_appends_after_existing_content(self):
        evidence = mutations.execute_append_markdown(str(self.target), "## Appendix\n\nAppended content.\n")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], 4)
        full_markdown, _ = projection.read_document_markdown(self.target)
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
        mutations._register_source_namespaces(rels_bytes)
        rels_root = ET.fromstring(rels_bytes)
        for rel in list(rels_root):
            if rel.get("Target") == "numbering.xml":
                rels_root.remove(rel)
        items[projection._rels_path_for(projection.DEFAULT_PART)] = ET.tostring(rels_root, encoding="utf-8")

        ct_bytes = items["[Content_Types].xml"]
        mutations._register_source_namespaces(ct_bytes)
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


if __name__ == "__main__":
    unittest.main()
