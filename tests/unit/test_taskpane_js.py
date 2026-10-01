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
