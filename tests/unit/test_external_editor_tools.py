"""Issue #27, through the real MCP tool layer.

test_external_editor.py calls the execute_* functions directly, which
skips the ~15 hand-edited @mcp.tool wrappers in server.py -- exactly where
a mistyped pass-through of ``allow_concurrent_editor`` would hide. These
tests drive every guarded tool in-process through a FastMCP ``Client``, so
the published parameter schema, the wrapper's pass-through, the error
envelope, and the middleware are all on the path.

Also here: ledger races on ONE document (an observation write racing
record_write) and the thread-isolation of the guard-to-write handoff.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastmcp import Client

from verified_docx_mcp import mutations, paths, projection, server, write_ledger

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

# tool name -> arguments (path is added). Every one of these reaches
# mutations._guard_before_write before it does anything that could fail for
# an unrelated reason; a tool that fails LATER with a different error still
# proves the override passed the guard.
GUARDED_TOOLS: dict[str, dict[str, Any]] = {
    "replace_body_markdown": {"markdown": "x"},
    "replace_range_markdown": {"section_key": "no-such-section", "markdown": "x"},
    "append_markdown": {"markdown": "x"},
    "replace_text": {"find": "R1C2", "replace": "z", "expected_matches": 1, "write_mode": "file"},
    "format_text": {"find": "R1C2", "style": {"bold": True}, "expected_matches": 1, "write_mode": "file"},
    "apply_style": {"find": "R1C2", "style_id": "Normal", "expected_matches": 1},
    "accept_tracked_changes": {},
    "reject_tracked_changes": {},
    "add_anchored_comment": {"quote": "R1C2", "text": "hi", "expected_matches": 1, "write_mode": "file"},
    "reply_to_comment": {"comment_id": "nope", "text": "hi", "write_mode": "file"},
    "resolve_comment": {"comment_id": "nope", "write_mode": "file"},
    "replace_table_row": {"table_id": 1, "row_index": 1, "cells": ["a", "b"]},
    "replace_cell_markdown": {
        "table_id": 1, "row_index": 1, "cell_index": 2, "markdown": "z", "write_mode": "file",
    },
    "insert_table": {"rows": [["a", "b"]], "style_id": "TableGrid"},
    "insert_image": {"image_path": str(FIXTURES / "images" / "sample.png")},
}


def _external_overwrite(path: Path) -> None:
    tmp = path.with_name(path.name + ".ext")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item)
            if item.filename == projection.DEFAULT_PART:
                data = data.replace(b"R2C2", b"EXTERNAL")
            dst.writestr(item, data)
    os.replace(tmp, path)


class _Env(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old = {
            k: os.environ.get(k) for k in (paths._ALLOWED_FILE_ROOTS_ENV, "XDG_STATE_HOME")
        }
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        os.environ["XDG_STATE_HOME"] = str(Path(self._tmp.name) / "state")
        self.target = Path(self._tmp.name) / "tables.docx"
        shutil.copyfile(FIXTURES / "tables.docx", self.target)

    async def asyncTearDown(self):
        for k, v in self._old.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        self._tmp.cleanup()

    async def call(self, client: Client, name: str, **args: Any):
        return await client.call_tool(name, {"path": str(self.target), **args}, raise_on_error=False)

    @staticmethod
    def error_code(result) -> str | None:
        if not result.is_error:
            return None
        text = "".join(getattr(c, "text", "") for c in result.content)
        try:
            return json.loads(text).get("error_code")
        except ValueError:
            return text  # not a VerifyError envelope; return raw text for the assertion message


class GuardedToolSchemaTests(unittest.IsolatedAsyncioTestCase):
    async def test_every_guarded_tool_publishes_allow_concurrent_editor(self):
        tools = {t.name: t for t in await server.mcp.list_tools()}
        for name in GUARDED_TOOLS:
            with self.subTest(tool=name):
                props = tools[name].parameters["properties"]
                self.assertIn("allow_concurrent_editor", props)
                self.assertEqual(props["allow_concurrent_editor"].get("default"), False)

    async def test_replace_cell_markdown_publishes_write_mode(self):
        tools = {t.name: t for t in await server.mcp.list_tools()}
        self.assertIn("write_mode", tools["replace_cell_markdown"].parameters["properties"])


class GuardedToolPassThroughTests(_Env):
    async def test_each_tool_refuses_then_accepts_the_override(self):
        async with Client(server.mcp) as client:
            for name, args in GUARDED_TOOLS.items():
                with self.subTest(tool=name):
                    # A fresh document per tool: copying the fixture back over a
                    # path that already has a ledger entry is itself (correctly)
                    # a divergence.
                    self.target = Path(self._tmp.name) / f"{name}.docx"
                    shutil.copyfile(FIXTURES / "tables.docx", self.target)
                    seed = await self.call(
                        client, "replace_cell_markdown",
                        table_id=1, row_index=1, cell_index=1, markdown="seed", write_mode="file",
                    )
                    self.assertFalse(seed.is_error, seed)
                    _external_overwrite(self.target)

                    refused = await self.call(client, name, **args)
                    self.assertEqual(
                        self.error_code(refused), "EXTERNAL_EDITOR_ACTIVE",
                        f"{name} did not refuse a divergent file",
                    )

                    allowed = await self.call(client, name, **args, allow_concurrent_editor=True)
                    self.assertNotEqual(
                        self.error_code(allowed), "EXTERNAL_EDITOR_ACTIVE",
                        f"{name} ignored allow_concurrent_editor=True",
                    )

    async def test_override_success_carries_evidence_through_the_tool_layer(self):
        async with Client(server.mcp) as client:
            await self.call(
                client, "replace_cell_markdown",
                table_id=1, row_index=1, cell_index=1, markdown="seed", write_mode="file",
            )
            _external_overwrite(self.target)
            result = await self.call(
                client, "replace_cell_markdown",
                table_id=1, row_index=1, cell_index=1, markdown="again",
                write_mode="file", allow_concurrent_editor=True,
            )
            self.assertFalse(result.is_error, result)
            data = result.structured_content or json.loads(result.content[0].text)
            self.assertTrue(data["applied"])
            self.assertIn("concurrent_editor_override", data)
            self.assertTrue(data["ledger_logged"])

    async def test_lock_status_tool_reports_external_activity(self):
        async with Client(server.mcp) as client:
            await self.call(
                client, "replace_cell_markdown",
                table_id=1, row_index=1, cell_index=1, markdown="seed", write_mode="file",
            )
            ok = await self.call(client, "lock_status")
            data = ok.structured_content or json.loads(ok.content[0].text)
            self.assertTrue(data["external_activity"]["still_current"])
            _external_overwrite(self.target)
            bad = await self.call(client, "lock_status")
            data = bad.structured_content or json.loads(bad.content[0].text)
            self.assertFalse(data["external_activity"]["still_current"])
            self.assertTrue(data["external_activity"]["recent"])


class LedgerRaceTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old = os.environ.get("XDG_STATE_HOME")
        os.environ["XDG_STATE_HOME"] = self._tmp.name
        self.doc = Path(self._tmp.name) / "doc.docx"
        with zipfile.ZipFile(self.doc, "w") as zf:
            zf.writestr("word/document.xml", b"<x/>")

    def tearDown(self):
        if self._old is None:
            os.environ.pop("XDG_STATE_HOME", None)
        else:
            os.environ["XDG_STATE_HOME"] = self._old
        self._tmp.cleanup()

    def test_observations_racing_record_write_never_damage_the_main_record(self):
        """external_activity (a read tool, no claim) writes the .obs sidecar
        while record_write (under the claim) rewrites the main record. They
        are different files on purpose; hammer both on ONE document."""
        write_ledger.record_write(self.doc, "fp-0", "tok-0")  # ledger != the file => always divergent
        stop = threading.Event()
        bad: list[str] = []

        def observe():
            while not stop.is_set():
                write_ledger.external_activity(self.doc)

        def read_main():
            while not stop.is_set():
                status, _rec, reason = write_ledger._read_record(self.doc)
                if status == write_ledger.LEDGER_UNREADABLE:
                    bad.append(reason)

        threads = [threading.Thread(target=observe), threading.Thread(target=observe),
                   threading.Thread(target=read_main)]
        for t in threads:
            t.start()
        try:
            for i in range(150):
                ok, reason = write_ledger.record_write(self.doc, f"fp-{i}", f"tok-{i}")
                self.assertTrue(ok, reason)
        finally:
            stop.set()
            for t in threads:
                t.join()
        self.assertEqual(bad, [])
        status, record, _ = write_ledger._read_record(self.doc)
        self.assertEqual(status, write_ledger.LEDGER_OK)
        self.assertEqual(record["fingerprint"], "fp-149")

    def test_observation_is_bound_to_the_ledger_generation(self):
        """A new record_write must not inherit the previous divergence's
        first-seen time."""
        write_ledger.record_write(self.doc, "fp-old", "t")
        first = write_ledger.external_activity(self.doc, now=1000.0)
        self.assertEqual(first["divergence_first_observed_at"], 1000.0)
        write_ledger.record_write(self.doc, "fp-new", "t")
        second = write_ledger.external_activity(self.doc, now=5000.0)
        self.assertEqual(second["divergence_first_observed_at"], 5000.0)


class PendingHandoffThreadTests(unittest.TestCase):
    def test_guard_context_is_thread_local(self):
        path = Path("/tmp/x-issue27.docx")
        mutations._pending_put(path, "fp", {"age_s": 1})
        seen: dict[str, Any] = {}
        t = threading.Thread(target=lambda: seen.update(mutations._pending_pop(path)))
        t.start()
        t.join()
        self.assertEqual(seen, {}, "another thread consumed this thread's guard context")
        mine = mutations._pending_pop(path)
        self.assertEqual(mine["source_fingerprint"], "fp")
        self.assertEqual(mutations._pending_pop(path), {}, "context must be consumed exactly once")

    def test_stale_guard_context_is_ignored(self):
        path = Path("/tmp/y-issue27.docx")
        old = mutations._PENDING_MAX_AGE_S
        mutations._PENDING_MAX_AGE_S = -1.0
        try:
            mutations._pending_put(path, "fp", None)
            self.assertEqual(mutations._pending_pop(path), {})
        finally:
            mutations._PENDING_MAX_AGE_S = old


if __name__ == "__main__":
    unittest.main()
