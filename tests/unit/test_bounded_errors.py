"""Issue #112: errors say what a caller can act on and nothing else."""

from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from fastmcp import Client

from verified_docx_mcp import locate, mutations, paths, server, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError, short_hash
from verified_docx_mcp.live import write_mode

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


class MatchDiagnosticsTests(unittest.TestCase):
    def test_five_contexts_and_the_total(self):
        text = " ".join(f"item {n:02d}: hit" for n in range(40))
        proj = SimpleNamespace(text=text)
        spans = [(m.start(), m.end()) for m in re.finditer("hit", text)]
        diag = locate._match_diagnostics(proj, spans)
        self.assertEqual(diag["total_matches"], 40)
        self.assertEqual(len(diag["matches"]), 5)
        for match in diag["matches"]:
            self.assertLessEqual(len(match["context"]), 40)
            self.assertEqual(set(match), {"start", "end", "context"})
        contexts = [m["context"] for m in diag["matches"]]
        self.assertEqual(len(set(contexts)), 5, "each hit must read differently")

    def test_context_never_runs_off_the_start_and_newlines_are_spaces(self):
        diag = locate._match_diagnostics(SimpleNamespace(text="hit\nnext line"), [(0, 3)])
        self.assertEqual(diag["matches"][0]["context"], "hit next line")


class ToolErrorShapeTests(unittest.TestCase):
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

    def call_error(self, tool: str, args: dict) -> dict:
        async def run():
            async with Client(server.mcp) as client:
                return await client.call_tool(tool, args, raise_on_error=False)

        result = asyncio.run(run())
        self.assertTrue(result.is_error, result)
        text = result.content[0].text
        return json.loads(text[text.index("{") :])

    def replace_args(self, **extra):
        return {"path": self.target.as_posix(), "find": "no such text", "replace": "x", "expected_matches": 1,
                "write_mode": "file", **extra}

    def test_zero_match_keeps_near_miss_and_drops_the_ladder_unless_verbose(self):
        lean = self.call_error("replace_text", self.replace_args())
        self.assertEqual(lean["error_code"], "ZERO_MATCH")
        self.assertNotIn("ladder_report", lean["diagnostics"])
        self.assertIn("near_miss", lean["diagnostics"])
        self.assertNotIn("retryable", lean)

        full = self.call_error("replace_text", self.replace_args(verbose=True))
        self.assertIn("ladder_report", full["diagnostics"])

    def test_count_mismatch_lists_contexts_not_bare_offsets(self):
        try:
            text_edit.execute_replace_text(
                self.target.as_posix(), "o", "0", 1, write_mode="file"
            )
            self.fail("expected a MATCH_COUNT_MISMATCH")
        except VerifyError as exc:
            diag = exc.envelope.diagnostics
        self.assertEqual(diag["total_matches"], 4)
        self.assertNotIn("spans", diag)
        self.assertTrue(all(m["context"] for m in diag["matches"]))


class StaleAndUnavailableTests(unittest.TestCase):
    def test_short_hash(self):
        self.assertEqual(short_hash("live:sha256:" + "ab" * 32), "ab" * 6)
        self.assertEqual(short_hash("cd" * 32), "cd" * 6)
        self.assertIsNone(short_hash(None))
        self.assertIsNone(short_hash(""))

    def test_live_stale_carries_twelve_hex_prefixes_not_two_full_tokens(self):
        with self.assertRaises(VerifyError) as ctx:
            write_mode.check_not_stale("live:sha256:" + "a" * 64, "b" * 64)
        env = ctx.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.LIVE_STALE)
        self.assertEqual(env.diagnostics, {"expected": "a" * 12, "actual": "b" * 12})
        self.assertNotRegex(env.message, r"[0-9a-f]{20}")
        self.assertIn("live_revision", env.message)

    def test_live_unavailable_is_the_code_the_document_and_one_hint(self):
        with tempfile.TemporaryDirectory() as tmp, mock.patch.dict(
            os.environ, {paths._ALLOWED_FILE_ROOTS_ENV: tmp}
        ), self.assertRaises(VerifyError) as ctx:
            write_mode.resolve_write_mode(str(Path(tmp) / "NoPane.docx"), "live")
        env = ctx.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.LIVE_UNAVAILABLE)
        self.assertEqual(env.diagnostics, {"document_name": "NoPane.docx"})
        self.assertLess(len(env.message), 100)


if __name__ == "__main__":
    unittest.main()
