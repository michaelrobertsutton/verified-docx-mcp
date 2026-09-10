"""Unit tests for src/verified_docx_mcp/audit.py (adapted from
GoogleDocs-MCP's verify.py audit-writer section — see that module's
header comment)."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import audit


class AppendAuditTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_xdg = os.environ.get("XDG_STATE_HOME")
        os.environ["XDG_STATE_HOME"] = self._tmp.name
        self._old_excerpts = os.environ.get(audit._AUDIT_EXCERPTS_ENV)
        os.environ.pop(audit._AUDIT_EXCERPTS_ENV, None)

    def tearDown(self):
        if self._old_xdg is None:
            os.environ.pop("XDG_STATE_HOME", None)
        else:
            os.environ["XDG_STATE_HOME"] = self._old_xdg
        if self._old_excerpts is None:
            os.environ.pop(audit._AUDIT_EXCERPTS_ENV, None)
        else:
            os.environ[audit._AUDIT_EXCERPTS_ENV] = self._old_excerpts
        self._tmp.cleanup()

    def _read_records(self) -> list[dict]:
        audit_path = audit._state_dir() / "audit.jsonl"
        lines = audit_path.read_text(encoding="utf-8").splitlines()
        return [json.loads(line) for line in lines]

    def test_appends_one_jsonl_line_with_owner_only_perms(self):
        ok, reason = audit.append_audit(
            path="/tmp/doc.docx", tool="export_pdf", evidence={"applied": True}
        )
        self.assertTrue(ok)
        self.assertEqual(reason, "")

        audit_path = audit._state_dir() / "audit.jsonl"
        self.assertTrue(audit_path.is_file())
        mode = oct(audit_path.stat().st_mode & 0o777)
        self.assertEqual(mode, "0o600")

        records = self._read_records()
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["path"], "/tmp/doc.docx")
        self.assertEqual(records[0]["tool"], "export_pdf")
        self.assertEqual(records[0]["evidence"], {"applied": True})
        self.assertIn("timestamp", records[0])

    def test_second_call_appends_not_overwrites(self):
        audit.append_audit(path="a.docx", tool="t1", evidence={"applied": True})
        audit.append_audit(path="b.docx", tool="t2", evidence={"applied": True})
        records = self._read_records()
        self.assertEqual(len(records), 2)
        self.assertEqual([r["path"] for r in records], ["a.docx", "b.docx"])

    def test_excerpts_redacted_when_disabled_via_env(self):
        os.environ[audit._AUDIT_EXCERPTS_ENV] = "0"
        audit.append_audit(
            path="a.docx",
            tool="replace_text",
            evidence={"applied": True, "before": "secret text", "after": "new text"},
        )
        record = self._read_records()[0]
        self.assertTrue(record["evidence"]["before"].startswith("[redacted;"))
        self.assertTrue(record["evidence"]["after"].startswith("[redacted;"))
        self.assertTrue(record["evidence"]["applied"])

    def test_excerpts_kept_by_default(self):
        audit.append_audit(
            path="a.docx",
            tool="replace_text",
            evidence={"applied": True, "before": "secret text", "after": "new text"},
        )
        record = self._read_records()[0]
        self.assertEqual(record["evidence"]["before"], "secret text")
        self.assertEqual(record["evidence"]["after"], "new text")

    def test_env_overrides_explicit_default_true(self):
        os.environ[audit._AUDIT_EXCERPTS_ENV] = "off"
        audit.append_audit(
            path="a.docx",
            tool="replace_text",
            evidence={"before": "x"},
            audit_excerpts=True,
        )
        record = self._read_records()[0]
        self.assertTrue(record["evidence"]["before"].startswith("[redacted;"))

    def test_never_raises_on_unwritable_state_dir(self):
        # Point XDG_STATE_HOME at a path that cannot be created as a
        # directory (a regular file in its place) to force append_audit's
        # mkdir to fail; it must report failure, not raise.
        blocker = Path(self._tmp.name) / "blocked"
        blocker.write_bytes(b"not a directory")
        os.environ["XDG_STATE_HOME"] = str(blocker)
        ok, reason = audit.append_audit(path="a.docx", tool="t", evidence={"applied": True})
        self.assertFalse(ok)
        self.assertNotEqual(reason, "")


if __name__ == "__main__":
    unittest.main()
