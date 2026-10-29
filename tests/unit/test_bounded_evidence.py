"""Issue #112: whole-document before/after are bounded in the response and kept in the
audit log, fetchable by evidence_ref."""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import audit, evidence, mutations, paths

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


class BoundedFieldsTests(unittest.TestCase):
    def test_append_excerpt_sits_at_the_end_of_the_old_text(self):
        before = "alpha\n\nbeta"
        after = before + "\n\ngamma"
        fields = evidence.bounded_fields(before, after)
        self.assertEqual(fields["before_chars"], len(before))
        self.assertEqual(fields["after_chars"], len(after))
        self.assertEqual(fields["paragraphs_before"], 2)
        self.assertEqual(fields["paragraphs_after"], 3)
        self.assertEqual(fields["before_excerpt"], before)
        self.assertTrue(fields["after_excerpt"].endswith("gamma"))
        self.assertNotEqual(fields["before_sha256"], fields["after_sha256"])

    def test_excerpt_is_a_200_character_window_around_the_first_difference(self):
        before = "x" * 5000 + "OLD" + "y" * 5000
        after = "x" * 5000 + "NEW" + "y" * 5000
        fields = evidence.bounded_fields(before, after)
        self.assertEqual(len(fields["before_excerpt"]), 200)
        self.assertIn("OLD", fields["before_excerpt"])
        self.assertIn("NEW", fields["after_excerpt"])

    def test_identical_and_empty_inputs(self):
        same = evidence.bounded_fields("same text", "same text")
        self.assertEqual(same["before_sha256"], same["after_sha256"])
        empty = evidence.bounded_fields("", "")
        self.assertEqual(empty["paragraphs_before"], 0)
        self.assertEqual(empty["before_excerpt"], "")

    def test_compact_drops_whole_text_and_only_includes_a_ref_that_exists(self):
        full = {"applied": True, "before": "a" * 9000, "after": "b" * 9000}
        out = evidence.compact_whole_document(dict(full), evidence_ref="abc123")
        self.assertNotIn("before", out)
        self.assertNotIn("after", out)
        self.assertEqual(out["evidence_ref"], "abc123")
        dangling = evidence.compact_whole_document(dict(full), evidence_ref=None)
        self.assertNotIn("evidence_ref", dangling)
        self.assertLess(len(json.dumps(out)), 1200)


class WholeDocumentToolsTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "doc.docx"
        shutil.copyfile(FIXTURES / "sections.docx", self.target)
        # A body big enough that echoing it would be expensive (~60 sections, ~8K words).
        words = " ".join(["program", "office", "delivery", "milestone"] * 33)
        big = "\n\n".join(f"# Section {n:03d}\n\n{words}." for n in range(60))
        mutations.execute_replace_body_markdown(str(self.target), big, force=True)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def test_append_one_paragraph_is_small_and_the_full_text_is_fetchable(self):
        result = mutations.execute_append_markdown(str(self.target), "One more paragraph.")
        self.assertTrue(result["applied"])
        self.assertLess(len(json.dumps(result)) // 4, 400)  # tokens, chars / 4
        self.assertGreater(result["after_chars"], 40_000)
        self.assertTrue(result["after_excerpt"].endswith("One more paragraph."))
        record = audit.find_audit_record(result["evidence_ref"])
        self.assertEqual(record["tool"], "append_markdown")
        self.assertTrue(record["evidence"]["after"].endswith("One more paragraph."))
        self.assertEqual(len(record["evidence"]["after"]), result["after_chars"])

    def test_replace_body_is_small_and_the_old_body_is_fetchable(self):
        body = mutations.execute_replace_body_markdown(
            str(self.target), "# Only\n\nSmall replacement body.", force=True
        )
        self.assertLess(len(json.dumps(body)) // 4, 500)
        self.assertGreater(body["before_chars"], 40_000)
        record = audit.find_audit_record(body["evidence_ref"])
        self.assertEqual(len(record["evidence"]["before"]), body["before_chars"])
        self.assertEqual(record["tool"], "replace_body_markdown")

    def test_replace_range_is_small_and_keeps_both_sides_in_the_audit_log(self):
        result = mutations.execute_replace_range_markdown(
            str(self.target), "section-007-1", "# Section 007\n\nRewritten section body."
        )
        self.assertTrue(result["applied"])
        self.assertLess(len(json.dumps(result)) // 4, 500)
        record = audit.find_audit_record(result["evidence_ref"])
        self.assertIn("program office", record["evidence"]["before"])
        self.assertIn("Rewritten section body.", record["evidence"]["after"])
        self.assertNotIn("before", result)

    def test_no_evidence_ref_when_the_audit_record_was_not_written(self):
        with mock.patch.object(audit, "append_audit", return_value=(False, "disk full")):
            result = mutations.execute_append_markdown(str(self.target), "Another paragraph.")
        self.assertFalse(result["audit_logged"])
        self.assertNotIn("evidence_ref", result)
        self.assertIn("after_sha256", result)

    def test_redacted_audit_still_returns_bounded_evidence(self):
        with mock.patch.dict(os.environ, {audit._AUDIT_EXCERPTS_ENV: "0"}):
            result = mutations.execute_append_markdown(str(self.target), "Redacted case.")
        record = audit.find_audit_record(result["evidence_ref"])
        self.assertTrue(record["evidence"]["after"].startswith("[redacted;"))
        self.assertIn("after_excerpt", result)

    def test_unknown_ref_is_none(self):
        self.assertIsNone(audit.find_audit_record("0" * 32))
        self.assertIsNone(audit.find_audit_record(""))

    def test_read_audit_evidence_returns_the_full_text_a_write_hid(self):
        from verified_docx_mcp import server
        from verified_docx_mcp.errors import ErrorCode, VerifyError

        written = mutations.execute_append_markdown(str(self.target), "Fetch me back.")
        ref = written["evidence_ref"]

        whole = server.execute_read_audit_evidence(ref)
        self.assertEqual(whole["tool"], "append_markdown")
        self.assertTrue(whole["evidence"]["after"].endswith("Fetch me back."))

        one = server.execute_read_audit_evidence(ref, field="after")
        self.assertEqual(one["field"], "after")
        self.assertTrue(one["value"].endswith("Fetch me back."))
        self.assertNotIn("evidence", one)

        out = Path(self._tmp.name) / "after.md"
        summary = server.execute_read_audit_evidence(ref, field="after", to_file=str(out))
        self.assertNotIn("value", summary)
        self.assertEqual(out.read_text(encoding="utf-8"), one["value"])
        self.assertEqual(summary["chars"], len(one["value"]))
        again = server.execute_read_audit_evidence(ref, field="after", to_file=str(out), overwrite=True)
        self.assertEqual(again["written_to"], summary["written_to"])

        # The output guard is shared with read_document: no overwrite, no non-text names.
        for bad in (
            {"evidence_ref": ref, "to_file": str(out)},  # exists now
            {"evidence_ref": ref, "to_file": str(self.target)},  # the .docx itself
            {"evidence_ref": "0" * 32},
            {"evidence_ref": ref, "field": "nope"},
        ):
            with self.assertRaises(VerifyError) as ctx:
                server.execute_read_audit_evidence(**bad)
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_read_audit_evidence_is_registered_read_only(self):
        import asyncio

        from verified_docx_mcp import server
        from verified_docx_mcp.middleware import MUTATING_TOOLS

        tool = asyncio.run(server.mcp.get_tool("read_audit_evidence"))
        self.assertTrue(tool.annotations.readOnlyHint)
        self.assertNotIn("read_audit_evidence", MUTATING_TOOLS)


if __name__ == "__main__":
    unittest.main()
