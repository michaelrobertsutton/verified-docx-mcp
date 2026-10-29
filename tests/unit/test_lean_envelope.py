"""Issue #112: the lean write envelope drops constants from what the model sees."""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from fastmcp import Client

from verified_docx_mcp import audit, mutations, paths, server
from verified_docx_mcp.lean_envelope import lean_result

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

LIVE_FULL = {
    "applied": True,
    "match_count": 1,
    "rung": 1,
    "before": "beta",
    "after": "BETA",
    "revision_before": "live:sha256:" + "a" * 64,
    "revision_after": "live:sha256:" + "b" * 64,
    "audit_logged": True,
    "write_mode": "live",
    "verified_via": "word-addin",
    "document_name": "Doc.docx",
    "track_changes_author": "word-signed-in-user",
    "orphaned_comment_ids": [],
    "evidence_ref": "abc",
}


class LeanResultTests(unittest.TestCase):
    def test_live_constants_go_and_information_stays(self):
        out = lean_result("replace_text", {"write_mode": "live", "revision_before": LIVE_FULL["revision_before"]}, LIVE_FULL)
        self.assertEqual(
            set(out), {"applied", "match_count", "rung", "before", "after", "revision_after", "evidence_ref"}
        )
        self.assertEqual(out["revision_after"], LIVE_FULL["revision_after"])

    def test_write_mode_is_kept_when_it_was_not_requested_explicitly(self):
        for args in ({}, {"write_mode": "auto"}):
            self.assertEqual(lean_result("replace_text", args, LIVE_FULL)["write_mode"], "live")

    def test_revision_before_is_kept_when_it_differs_or_was_not_passed(self):
        self.assertIn("revision_before", lean_result("replace_text", {"revision_before": "live:sha256:" + "c" * 64}, LIVE_FULL))
        self.assertIn("revision_before", lean_result("replace_text", {}, LIVE_FULL))

    def test_the_batch_always_keeps_revision_before(self):
        batch = {"applied": True, "revision_before": "rev1", "revision_after": "rev2", "batch_status": "complete"}
        self.assertEqual(lean_result("replace_text_batch", {"revision_before": "rev1"}, batch)["revision_before"], "rev1")

    def test_false_and_nonempty_flags_are_information_and_stay(self):
        result = {
            "applied": True,
            "audit_logged": False,
            "ledger_logged": False,
            "conflict_copy_detected": True,
            "orphaned_comment_ids": ["c1"],
            "conflict_copies": ["x (conflicted copy).docx"],
        }
        self.assertEqual(lean_result("replace_text", {}, result), result)

    def test_file_mode_constants_go(self):
        result = {"applied": True, "write_mode": "file", "audit_logged": True, "ledger_logged": True,
                  "conflict_copy_detected": False, "revision_after": "r2", "revision_before": "r1"}
        out = lean_result("append_markdown", {"write_mode": "file", "revision_before": "r1"}, result)
        self.assertEqual(out, {"applied": True, "revision_after": "r2"})

    def test_single_match_comment_returns_comment_id_only(self):
        result = {"applied": True, "before": "the quote", "after": "the quote", "comment_id": "c1", "comment_ids": ["c1"]}
        out = lean_result("add_anchored_comment", {"quote": "the quote"}, result)
        self.assertEqual(out, {"applied": True, "comment_id": "c1"})
        many = {"applied": True, "comment_ids": ["c1", "c2"], "before": "q", "after": "q"}
        self.assertEqual(lean_result("add_anchored_comment", {"quote": "q"}, many)["comment_ids"], ["c1", "c2"])
        # A before that is only an excerpt around the quote (file mode) is information: kept.
        excerpt = {"applied": True, "before": "context the quote context", "after": "context the quote context"}
        self.assertIn("before", lean_result("add_anchored_comment", {"quote": "the quote"}, excerpt))

    def test_errors_and_non_evidence_dicts_pass_through_untouched(self):
        error = {"error_code": "ZERO_MATCH", "audit_logged": True}
        self.assertIs(lean_result("replace_text", {}, error), error)
        plain = {"audit_logged": True}
        self.assertIs(lean_result("replace_text", {}, plain), plain)

    def test_input_is_not_mutated(self):
        snapshot = json.loads(json.dumps(LIVE_FULL))
        lean_result("replace_text", {"write_mode": "live"}, LIVE_FULL)
        self.assertEqual(LIVE_FULL, snapshot)


class EndToEndTests(unittest.TestCase):
    """Through the real FastMCP server: the model-facing result is lean; the audit log is not."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "frag.docx"
        shutil.copyfile(FIXTURES / "frag.docx", self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def call(self, tool: str, args: dict) -> dict:
        async def run():
            async with Client(server.mcp) as client:
                return (await client.call_tool(tool, args)).structured_content

        return asyncio.run(run())

    def test_replace_text_result_is_lean_but_the_audit_record_is_full(self):
        result = self.call(
            "replace_text",
            {"path": str(self.target), "find": "quick", "replace": "slow", "expected_matches": 1, "write_mode": "file"},
        )
        self.assertTrue(result["applied"])
        self.assertEqual(result["match_count"], 1)
        self.assertIn("revision_after", result)
        for gone in ("audit_logged", "conflict_copy_detected", "ledger_logged", "write_mode"):
            self.assertNotIn(gone, result)
        path = audit._state_dir() / "audit.jsonl"
        record = json.loads(path.read_text().splitlines()[-1])
        self.assertEqual(record["tool"], "replace_text")
        self.assertEqual(record["evidence"]["write_mode"], "file")
        self.assertIn("conflict_copy_detected", record["evidence"])

    def test_the_enforcement_contract_still_holds(self):
        result = self.call("append_markdown", {"path": str(self.target), "markdown": "More."})
        self.assertIn("applied", result)


if __name__ == "__main__":
    unittest.main()
