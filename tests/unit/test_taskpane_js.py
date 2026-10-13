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

    def test_ooxml_crosscheck_counts_open_root_threads_only(self):
        # 4 comments: roots 1 (open), 2 (done=1), 3 (done=false); 4 is a reply to 1.
        self.assertEqual(
            self.out["ooxml"],
            {"available": True, "total_threads": 3, "open_threads": 2, "ids": ["1", "3"]},
        )

    def test_ooxml_crosscheck_unavailable_without_a_comments_part(self):
        for key in ("ooxmlNoExtended", "ooxmlNoComments", "ooxmlGarbage"):
            block = self.out[key]
            self.assertFalse(block["available"], key)
            self.assertIsNone(block["open_threads"], key)
            self.assertEqual(block["ids"], [], key)

    def test_ooxml_crosscheck_skipped_for_an_ids_filter_and_survives_a_missing_reader(self):
        self.assertIsNone(self.out["ooxmlFiltered"])
        self.assertFalse(self.out["ooxmlNoReader"]["available"])  # body.getOoxml absent -> caught, list still returned

    def test_ooxml_crosscheck_adds_one_sync_and_a_capability(self):
        self.assertEqual(self.out["ooxmlSyncs"], 3)
        self.assertTrue(self.out["capability"])


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


DELETE_HARNESS = Path(__file__).resolve().parent / "js" / "comment_delete_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class OpCommentDeleteTests(unittest.TestCase):
    """issue #80: the pane's comment_delete op."""

    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(DELETE_HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"harness failed:\n{proc.stderr}")
        cls.out = json.loads(proc.stdout.strip().splitlines()[-1])

    def test_deletes_and_reports_reply_count_and_absence(self):
        # `resolved` is read only after an explicit load, or the stub throws.
        self.assertEqual(self.out["ok"]["result"], {"deleted": True, "reply_count": 1, "still_present": False})
        self.assertEqual(self.out["ok"]["deleted"], ["c1"])

    def test_require_resolved_refuses_an_open_comment_before_deleting(self):
        self.assertIn("not resolved", self.out["openRefused"]["error"])
        self.assertEqual(self.out["openRefused"]["deleted"], [])
        self.assertTrue(self.out["openAllowedWhenNotRequired"]["result"]["deleted"])

    def test_missing_stale_and_changed_comments_are_refused(self):
        self.assertIn("no comment with id", self.out["missing"]["error"])
        self.assertEqual(self.out["staleEpoch"]["code"], "COMMENT_ID_STALE")
        self.assertEqual(self.out["identityMismatch"]["code"], "COMMENT_ID_STALE")
        for key in ("missing", "staleEpoch", "identityMismatch"):
            self.assertEqual(self.out[key]["deleted"], [], key)

    def test_still_present_is_reported_not_thrown(self):
        self.assertTrue(self.out["notDeleted"]["result"]["still_present"])


BOOT_DIAGNOSTICS_HARNESS = Path(__file__).resolve().parent / "js" / "boot_diagnostics_harness.mjs"


@unittest.skipUnless(NODE, "node not found on PATH")
class BootDiagnosticsTests(unittest.TestCase):
    """Issue #84: lifecycle capture, previous_events, boot_seq, decoupled connect."""

    @classmethod
    def setUpClass(cls):
        proc = subprocess.run(
            [NODE, str(BOOT_DIAGNOSTICS_HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if proc.returncode != 0:
            raise AssertionError(f"harness failed:\n{proc.stderr}")
        cls.out = json.loads(proc.stdout.strip().splitlines()[-1])

    def test_ring_buffer_is_capped_and_persisted(self):
        self.assertEqual(self.out["ringLength"], 50)
        self.assertEqual((self.out["ringFirst"], self.out["ringLast"]), ("e70", "e119"))
        self.assertEqual(self.out["ringStackLen"], 400)
        self.assertEqual(self.out["ringKeys"], ["kind", "message", "stack", "t"])

    def test_previous_events_reported_then_cleared(self):
        self.assertEqual(
            self.out["previousKinds"],
            ["boot", "error", "unhandledrejection", "pagehide", "beforeunload", "visibilitychange"],
        )
        self.assertEqual(self.out["previousErrorMessage"], "kaboom")
        # the key was cleared at load: only this load's own boot event is stored
        self.assertEqual(self.out["afterSecondLoadStored"], ["boot"])
        self.assertEqual(self.out["thirdPreviousKinds"], ["boot"])

    def test_report_keeps_existing_fields_and_adds_diagnostics(self):
        self.assertEqual(self.out["reportKeepsFields"], ["issue-106-wp1", "https://x.example/Row01.docx", 1, None])
        self.assertTrue(self.out["reportHasInstance"])
        self.assertEqual(self.out["reportBootMsKeys"], ["office_ready", "report_build", "ws_connect"])

    def test_boot_seq_increments(self):
        self.assertEqual(self.out["bootSeqs"], [1, 2, 3])

    def test_throwing_storage_does_not_break_boot(self):
        for key in ("throwingMethods", "throwingAccess"):
            entry = self.out[key]
            self.assertEqual(entry["posted"], 1, key)
            self.assertIsNone(entry["bootSeq"], key)
            self.assertEqual(entry["previous"], [], key)
            self.assertEqual((entry["comments"], entry["sockets"]), (1, 1), key)

    def test_connect_starts_before_and_regardless_of_report_build(self):
        for mode in ("reject", "hang"):
            self.assertEqual(self.out[f"{mode}SocketsImmediately"], 1, mode)
            self.assertEqual(self.out[f"{mode}PostsImmediately"], 0, mode)

    def test_failed_build_posts_minimal_report(self):
        for mode in ("reject", "hang"):
            entry = self.out[mode]
            self.assertEqual(entry["posts"], 1, mode)
            self.assertEqual(
                entry["keys"],
                ["boot_ms", "boot_seq", "documentUrl", "generated_at", "instance_id", "load_error", "previous_events"],
            )
            self.assertEqual(entry["documentUrl"], "https://x.example/Row01.docx")
            self.assertEqual(entry["bootSeq"], 1)
            self.assertIn("load_error", entry["storedKinds"])
        self.assertIn("getComments blew up", self.out["reject"]["loadError"])
        self.assertIn("timed out", self.out["hang"]["loadError"])

