"""Unit tests for src/verified_docx_mcp/tracked_changes.py (issue #28
WP-07): list_open_items' Google-shaped output, the overlapping-write
TRACKED_CHANGES_PRESENT refusal (text_edit.py's guard, exercised here end
to end), accept/reject clearing the way for a subsequent write, and
REVISION_ID_NOT_FOUND.

Fixture provenance: tests/fixtures/revision/tracked.docx is Word-authored
(tests/fixtures/README.md) and already carries exactly two edits -- one
w:ins ("X" inserted) and one w:del ("he re" deleted), both with real
w:author/w:date -- so it satisfies this WP's "two edits" acceptance
requirement without a new fixture. tests/fixtures/revision/commented.docx
(also Word-authored, WP-03) covers the comments half of list_open_items'
Google-shaped output; comment reply THREADING is WP-08 scope (see that
fixture's own golden-comment.docx, untouched here) and is not tested.
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

from verified_docx_mcp import mutations, paths, projection, text_edit, tracked_changes
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "revision/tracked.docx"

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
    def test_accept_reject_registered_list_open_items_is_not(self):
        self.assertIn("accept_tracked_changes", MUTATING_TOOLS)
        self.assertIn("reject_tracked_changes", MUTATING_TOOLS)
        self.assertNotIn("list_open_items", MUTATING_TOOLS, "list_open_items is read-only")


class ListOpenItemsTrackedChangesTests(_TempFixtureCase):
    """Acceptance: both edits in tracked.docx are listed, Google's shape."""

    def test_both_edits_listed_in_googles_shape(self):
        result = tracked_changes.execute_list_open_items(str(self.target))
        self.assertIn("comments", result)
        self.assertIn("pending_suggestions", result)
        suggestions = result["pending_suggestions"]
        self.assertEqual(len(suggestions), 2)

        by_kind = {s["kind"]: s for s in suggestions}
        self.assertEqual(set(by_kind), {"insertion", "deletion"})

        insertion = by_kind["insertion"]
        self.assertEqual(insertion["text"], "X")
        self.assertEqual(insertion["author"], "Michael Sutton")
        self.assertTrue(insertion["date"])
        self.assertIn("suggestion_id", insertion)
        self.assertIn("anchor_context", insertion)

        deletion = by_kind["deletion"]
        self.assertEqual(deletion["text"], "he re")
        self.assertEqual(deletion["author"], "Michael Sutton")


class ListOpenItemsCommentsTests(_TempFixtureCase):
    fixture_name = "revision/commented.docx"

    def test_comment_listed_in_googles_shape(self):
        result = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(len(result["comments"]), 1)
        comment = result["comments"][0]
        for key in ("comment_id", "content", "resolved", "reply_count", "replies", "quoted_text", "author", "scope"):
            self.assertIn(key, comment)
        self.assertEqual(comment["content"], "This is a test comment.")
        self.assertEqual(comment["quoted_text"], "This")
        self.assertFalse(comment["resolved"])
        self.assertEqual(comment["scope"], "document")
        self.assertEqual(result["pending_suggestions"], [])

    def test_comment_id_is_the_durable_id_not_the_raw_w_id(self):
        # WP-08 interop fix: comment_id must be commentsIds.xml's own
        # durableId (Google's shape: a durable, opaque id), not the raw
        # w:id -- which is only ever max-existing-plus-1 and therefore not
        # stable/round-trippable across edits. The raw w:id is still
        # available, under its own clearly-named field.
        result = tracked_changes.execute_list_open_items(str(self.target))
        comment = result["comments"][0]
        self.assertEqual(comment["comment_id"], "2A21A9D2")  # commented.docx's own real durableId
        self.assertEqual(comment["w_id"], "0")


class OverlappingWriteRefusalTests(_TempFixtureCase):
    """Acceptance: overlapping replace_text refuses; accept clears it and
    the write then proceeds.

    Patches text_edit.resolve_author_name to a name DIFFERENT from
    tracked.docx's real Word-authored author ("Michael Sutton") for the
    whole class: WP-07b-a's own-author exclusion (text_edit.py's
    _check_tracked_changes_guard) would otherwise treat tracked.docx's
    revisions as this server's own prior work whenever this suite happens
    to run on Michael Sutton's own machine (resolve_author_name's own
    fallback path -- see test_author.py), silently defeating exactly the
    refusal these tests exist to prove. Pinning a fixed, different name
    makes the "foreign author" case deterministic regardless of whose
    machine runs the suite.
    """

    def setUp(self):
        super().setUp()
        patcher = mock.patch("verified_docx_mcp.text_edit.resolve_author_name", return_value="A Different Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_replace_text_over_an_insertion_refuses_then_succeeds_after_accept(self):
        before_bytes = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_replace_text(str(self.target), "X", "Y", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.TRACKED_CHANGES_PRESENT)
        self.assertIn("crosses_revision", cm.exception.envelope.diagnostics.get("warnings", []))
        self.assertEqual(self.target.read_bytes(), before_bytes, "a refused write must not touch the file")

        accept_evidence = tracked_changes.execute_accept_tracked_changes(str(self.target))
        self.assertTrue(accept_evidence["applied"])
        self.assertEqual(accept_evidence["match_count"], 2)
        self.assertEqual(accept_evidence["rung"], "all")
        self.assertEqual(sorted(accept_evidence["revision_ids"]), ["0", "1"])

        # The write proceeds now that no tracked change remains.
        replace_evidence = text_edit.execute_replace_text(str(self.target), "X", "Y", 1)
        self.assertTrue(replace_evidence["applied"])
        self.assertIn("Y", projection.read_document_text(self.target))

    def test_force_bypasses_the_refusal(self):
        evidence = text_edit.execute_replace_text(str(self.target), "X", "Y", 1, force=True)
        self.assertTrue(evidence["applied"])
        self.assertIn("Y", projection.read_document_text(self.target))


class AcceptTests(_TempFixtureCase):
    def test_accept_all_makes_insertion_permanent_and_deletion_final(self):
        tracked_changes.execute_accept_tracked_changes(str(self.target))
        text = projection.read_document_text(self.target)
        # "X" (the accepted insertion) is now ordinary live text; "he re"
        # (the accepted deletion) stays gone -- both already true of the
        # LIVE projection even before accepting (w:ins content is live,
        # w:del content is excluded), so the real assertion is that no
        # tracked change remains afterward.
        self.assertIn("X", text)
        items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(items["pending_suggestions"], [])

    def test_accept_by_id_leaves_the_other_pending(self):
        evidence = tracked_changes.execute_accept_tracked_changes(str(self.target), revision_ids=["0"])
        self.assertEqual(evidence["rung"], "by_id")
        items = tracked_changes.execute_list_open_items(str(self.target))
        remaining = [s["suggestion_id"] for s in items["pending_suggestions"]]
        self.assertEqual(remaining, ["1"])

    def test_unknown_revision_id_raises_with_available_ids(self):
        with self.assertRaises(VerifyError) as cm:
            tracked_changes.execute_accept_tracked_changes(str(self.target), revision_ids=["nope"])
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.REVISION_ID_NOT_FOUND)
        self.assertEqual(env.diagnostics["missing_ids"], ["nope"])
        self.assertEqual(sorted(env.diagnostics["available_ids"]), ["0", "1"])


class RejectTests(_TempFixtureCase):
    def test_reject_all_restores_original_text(self):
        tracked_changes.execute_reject_tracked_changes(str(self.target))
        text = projection.read_document_text(self.target)
        base_text = projection.read_document_text(FIXTURES / "revision" / "base.docx")
        self.assertEqual(text, base_text)
        items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(items["pending_suggestions"], [])


if __name__ == "__main__":
    unittest.main()
