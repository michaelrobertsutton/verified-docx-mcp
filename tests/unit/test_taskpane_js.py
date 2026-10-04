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
            ["id", "content", "authorName", "creationDate", "resolved", "anchorText", "anchorParagraphText", "replies"],
        )
        self.assertEqual(self.out["firstReplyCount"], 1)
        self.assertTrue(self.out["hasTiming"])

    def test_ids_filter_only_touches_matching_comments(self):
        self.assertEqual(self.out["filteredIds"], ["c7", "c9"])
        self.assertEqual(self.out["filteredGetRange"], 2)

    def test_counts_reflect_whole_collection_even_when_ids_filter_applies(self):
        # makeDoc resolves every even-indexed comment: 150 total, 75 open.
        self.assertEqual(self.out["bigCounts"], {"total": 150, "open": 75})
        self.assertEqual(self.out["filteredCounts"], {"total": 150, "open": 75})

    def test_scope_and_observed_at_reported(self):
        self.assertEqual(self.out["scope"], "body")
        self.assertTrue(self.out["hasObservedAt"])

    def test_include_anchor_false_skips_get_range(self):
        self.assertEqual(self.out["noAnchorGetRange"], 0)
        self.assertFalse(self.out["noAnchorHasAnchorText"])


BODY_OOXML_HARNESS = Path(__file__).resolve().parent / "js" / "body_ooxml_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class OpBodyOoxmlTests(unittest.TestCase):
    """issue #33. A stub DOM stands in for the browser's (node has none), so
    these pin the op's control flow; the body_ooxml op itself was checked
    against real Word (docs/live-mode.md, "Live reads (issue #33)")."""

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
            self.out["bodySha256"], "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824"
        )

    def test_binary_parts_are_found_by_namespace_and_emptied(self):
        self.assertEqual(self.out["seenNs"], "http://schemas.microsoft.com/office/2006/xmlPackage")
        self.assertEqual(self.out["seenLocal"], "binaryData")
        self.assertEqual(self.out["strippedParts"], ["/word/media/a.png", "/word/media/b.png"])
        self.assertEqual(self.out["emptied"], [["/word/media/a.png", ""], ["/word/media/b.png", ""]])
        self.assertEqual(self.out["mime"], "application/xml")

    def test_unparseable_export_is_refused_not_sent_unstripped(self):
        self.assertFalse(self.out["badOk"])
        self.assertEqual(self.out["badCode"], "LIVE_OP_FAILED")

    def test_oversized_reply_is_refused_with_too_large(self):
        self.assertEqual(self.out["tooLargeCode"], "too_large")


TABLES_HARNESS = Path(__file__).resolve().parent / "js" / "tables_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class LiveTableOpsTests(unittest.TestCase):
    """Issue #34: table_get / table_insert / cells_set (and the refactored
    cell_set) against a MOCK Word object model that enforces load()+sync()
    ordering, ClientResult.value-after-sync and queued writes. Real Word is
    still unverified (docs/live-mode.md)."""

    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(TABLES_HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"tables harness failed:\n{proc.stderr}")
        cls.results = json.loads(proc.stdout.strip().splitlines()[-1])["results"]

    def test_harness_ran_every_scenario(self):
        self.assertGreaterEqual(len(self.results), 50)

    def test_every_check_passes(self):
        for result in self.results:
            with self.subTest(result["name"]):
                self.assertTrue(result["pass"], result["detail"])
