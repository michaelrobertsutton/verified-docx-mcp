"""Behavior tests for addin/taskpane.js's comments_list op (issue #31), run
under node against a stub Word context. Skipped when node is not installed."""

from __future__ import annotations

import json
import shutil
import subprocess
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
HARNESS = Path(__file__).resolve().parent / "js" / "comments_list_harness.mjs"
NODE = shutil.which("node")


@unittest.skipUnless(NODE, "node not found on PATH")
class OpCommentsListTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"harness failed:\n{proc.stderr}")
        cls.out = json.loads(proc.stdout.strip().splitlines()[-1])

    def test_sync_count_does_not_grow_with_comment_count(self):
        self.assertLessEqual(self.out["syncsBig"], 3)
        self.assertEqual(self.out["syncsSmall"], self.out["syncsBig"])

    def test_wire_shape(self):
        self.assertEqual(self.out["bigCount"], 150)
        self.assertEqual(
            self.out["firstKeys"],
            ["id", "content", "authorName", "creationDate", "resolved", "anchorText", "replies"],
        )
        self.assertEqual(self.out["firstReplyCount"], 1)
        self.assertTrue(self.out["hasTiming"])

    def test_ids_filter_only_touches_matching_comments(self):
        self.assertEqual(self.out["filteredIds"], ["c7", "c9"])
        self.assertEqual(self.out["filteredGetRange"], 2)

    def test_include_anchor_false_skips_get_range(self):
        self.assertEqual(self.out["noAnchorGetRange"], 0)
        self.assertFalse(self.out["noAnchorHasAnchorText"])


BODY_OOXML_HARNESS = Path(__file__).resolve().parent / "js" / "body_ooxml_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class OpBodyOoxmlTests(unittest.TestCase):
    """issue #33. A stub DOM stands in for the browser's (node has none), so
    these pin the op's control flow; real-DOM behavior is the live Word spike."""

    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(BODY_OOXML_HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"harness failed:\n{proc.stderr}")
        cls.out = json.loads(proc.stdout.strip().splitlines()[-1])

    def test_pane_advertises_the_capability(self):
        self.assertIn("body_ooxml", self.out["capabilities"])

    def test_reply_shape(self):
        self.assertEqual(self.out["keys"], ["ooxml", "bodySha256", "documentUrl", "strippedParts"])
        self.assertEqual(self.out["documentUrl"], "https://x.example/Doc.docx")
        # sha256("hello"), computed the way opDescribe computes the body hash
        self.assertEqual(
            self.out["bodySha256"],
            "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824",
        )

    def test_binary_parts_are_found_by_namespace_and_emptied(self):
        self.assertEqual(self.out["seenNs"], "http://schemas.microsoft.com/office/2006/xmlPackage")
        self.assertEqual(self.out["seenLocal"], "binaryData")
        self.assertEqual(self.out["strippedParts"], ["/word/media/a.png", "/word/media/b.png"])
        self.assertEqual(
            self.out["emptied"], [["/word/media/a.png", ""], ["/word/media/b.png", ""]]
        )
        self.assertEqual(self.out["mime"], "application/xml")

    def test_unparseable_export_is_refused_not_sent_unstripped(self):
        self.assertFalse(self.out["badOk"])
        self.assertEqual(self.out["badCode"], "LIVE_OP_FAILED")

    def test_oversized_reply_is_refused_with_too_large(self):
        self.assertEqual(self.out["tooLargeCode"], "too_large")


TEXTBOX_HARNESS = Path(__file__).resolve().parent / "js" / "textbox_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class TextboxOpTests(unittest.TestCase):
    """issue #35. Real taskpane.js against a stubbed shape object model: pins
    the pane's control flow (traversal, capability gate, compare-and-set
    before any write). What a real Word returns for body.shapes is the manual
    sideload check in docs/live-mode.md."""

    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(TEXTBOX_HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"harness failed:\n{proc.stderr}")
        cls.out = json.loads(proc.stdout.strip().splitlines()[-1])

    def test_capability_only_with_wordapidesktop_1_2(self):
        self.assertIn("textbox_scope", self.out["capsWithDesktop"])
        self.assertNotIn("textbox_scope", self.out["capsWithoutDesktop"])
        self.assertEqual(self.out["unsupportedCode"], "capability_missing")

    def test_traversal_recurses_groups_and_canvases_and_keeps_geometric_shapes(self):
        self.assertEqual(self.out["listedIds"], ["1", "3", "5"])
        self.assertEqual(self.out["listedTypes"], ["textbox", "textbox", "geometricshape"])
        self.assertEqual(self.out["deepGroupPath"], ["2", "4"])
        self.assertEqual(self.out["skipped"], [{"shape_id": "6", "type": "picture"}])
        self.assertTrue(self.out["incompleteEmpty"])
        self.assertEqual(self.out["firstParagraphCount"], 2)
        self.assertTrue(self.out["firstShaMatches"])

    def test_duplicate_ids_and_unreadable_bodies_are_reported_not_skipped(self):
        self.assertEqual(self.out["messyIds"], ["8"])
        self.assertEqual(
            self.out["messyIncomplete"],
            [["7", "duplicate shape id"], ["7", "duplicate shape id"], ["9", "body not readable"]],
        )

    def test_exhaustive_scope_refuses_on_incomplete_coverage_but_a_single_shape_works(self):
        self.assertEqual(self.out["incompleteCode"], "incomplete_coverage")
        self.assertTrue(self.out["singleOk"])
        self.assertEqual(self.out["singlePostText"], "OK")
        self.assertEqual(self.out["singleAfter"], "OK")

    def test_stale_or_vanished_target_refuses_before_any_write(self):
        self.assertEqual(self.out["staleCode"], "stale_target")
        self.assertEqual(self.out["vanishedCode"], "stale_target")
        self.assertEqual(self.out["insertsAfterStale"], 0)
        self.assertFalse(self.out["noExpectOk"])

    def test_count_gate_refuses_before_any_write(self):
        self.assertIn("found 1", self.out["wrongCountMessage"])
        self.assertEqual(self.out["insertsAfterWrongCount"], 0)

    def test_happy_path_reads_back_what_word_reports(self):
        self.assertTrue(self.out["goodOk"])
        self.assertEqual(
            self.out["goodShapes"],
            [{"shape_id": "1", "match_count": 1, "pre_text": "alpha beta", "post_text": "alpha B"}],
        )
        self.assertEqual(self.out["goodMatch"], {"before": "beta", "after": "B", "shape_id": "1"})
        self.assertEqual(self.out["goodBodyMatchCount"], 0)
        self.assertTrue(self.out["goodPreEqualsPost"])  # a text-box edit leaves the body hash alone

    def test_invalid_scope_is_refused_and_body_scope_never_touches_shapes(self):
        self.assertFalse(self.out["invalidScopeOk"])
        self.assertEqual(self.out["bodySearchCalls"], 1)
        self.assertEqual(self.out["insertsOnBodyPath"], 0)
