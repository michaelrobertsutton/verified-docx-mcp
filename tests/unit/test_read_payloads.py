"""Issue #111 part B: lean read envelope, lock_status trimming, tool annotations."""

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

from verified_docx_mcp import paths, server
from verified_docx_mcp.lean_envelope import READ_TOOLS, lean_read_result
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

LIVE_HASH = "a" * 64
LIVE_READ = {
    "path": "/x/Doc.docx",
    "part": "word/document.xml",
    "format": "text",
    "section_key": None,
    "revision": f"live:sha256:{LIVE_HASH}",
    "revision_detail": None,
    "text": "hello",
    "warnings": [],
    "source": "live",
    "live": {"document_url": "https://x/Doc.docx", "body_sha256": LIVE_HASH, "stripped_parts": []},
}

READ_ONLY_TOOLS = frozenset(
    ["read_document", "find_sections", "list_page_sections", "list_tables", "get_table", "list_open_items", "get_comment_thread", "list_styles", "list_parts", "list_shapes", "list_textboxes", "read_textbox", "read_header_footer", "diff_body_vs_file", "live_status", "live_list_sections", "live_revision", "lock_status", "export_pdf", "read_audit_evidence", "diff_table_vs_markdown"]
)
DESTRUCTIVE_TOOLS = frozenset(
    ["delete_paragraph", "delete_table", "delete_comment", "delete_resolved_comments", "replace_body_markdown", "accept_tracked_changes", "reject_tracked_changes", "delete_table_row"]
)


class LeanReadResultTests(unittest.TestCase):
    def test_constants_and_empties_go(self):
        out = lean_read_result("read_document", {"path": "/x/Doc.docx", "format": "text"}, LIVE_READ)
        for gone in ("section_key", "revision_detail", "warnings"):
            self.assertNotIn(gone, out)
        self.assertNotIn("body_sha256", out["live"])
        self.assertEqual(out["live"], {"document_url": "https://x/Doc.docx", "stripped_parts": []})
        # JennyStack reads these from read_document: all stay.
        for kept in ("path", "source", "format", "text", "revision"):
            self.assertIn(kept, out)

    def test_live_dropped_when_it_becomes_empty(self):
        result = dict(LIVE_READ, live={"body_sha256": LIVE_HASH})
        self.assertNotIn("live", lean_read_result("read_document", {}, result))

    def test_live_hash_kept_when_it_is_not_the_revision(self):
        result = dict(LIVE_READ, revision="live:sha256:" + "b" * 64)
        out = lean_read_result("read_document", {}, result)
        self.assertEqual(out["live"]["body_sha256"], LIVE_HASH)

    def test_non_empty_values_stay(self):
        result = {
            "path": "/x/a.docx", "revision": "r", "section_key": "textbox-1", "markdown": "m",
            "warnings": ["live_styles_missing"], "lossy_elements": [{"kind": "nested_table", "table_id": 1}],
        }
        out = lean_read_result("read_document", {"path": "/x/a.docx"}, result)
        self.assertEqual(out, result)

    def test_revision_detail_only_with_verbose(self):
        detail = {"document_sha256": "d" * 64, "comments_sha256": "c" * 64, "size": 1, "mtime_ns": 2}
        result = {"path": "/x/a.docx", "revision": "r", "revision_detail": detail}
        self.assertNotIn("revision_detail", lean_read_result("read_document", {"path": "/x/a.docx"}, result))
        self.assertNotIn(
            "revision_detail", lean_read_result("read_document", {"verbose": False}, result)
        )
        self.assertEqual(lean_read_result("read_document", {"verbose": True}, result)["revision_detail"], detail)

    def test_path_dropped_only_when_equal_and_not_a_jennystack_tool(self):
        result = {"path": "/x/a.docx", "sections": []}
        self.assertNotIn("path", lean_read_result("find_sections", {"path": "/x/a.docx"}, result))
        self.assertIn("path", lean_read_result("find_sections", {"path": "other.docx"}, result))
        for keeper in ("read_document", "list_open_items"):
            self.assertIn("path", lean_read_result(keeper, {"path": "/x/a.docx"}, result))

    def test_live_session_is_shortened(self):
        result = {
            "path": "/x/a.docx",
            "warnings": ["live_session_ignored"],
            "live_session": {"document_url": "u", "reason": "requested_file", "action": "long guidance"},
        }
        out = lean_read_result("find_sections", {"path": "/x/a.docx"}, result)
        self.assertEqual(out["live_session"], {"document_url": "u", "reason": "requested_file"})
        self.assertEqual(out["warnings"], ["live_session_ignored"])

    def test_header_footer_parts_drop_empty_warnings_and_lossy(self):
        result = {
            "path": "/x/a.docx",
            "headers_and_footers": [
                {"part": "word/header1.xml", "markdown": "H", "warnings": [], "lossy_elements": []},
                {"part": "word/footer1.xml", "markdown": "F", "warnings": ["w"], "lossy_elements": [{"kind": "k"}]},
            ],
        }
        out = lean_read_result("read_header_footer", {"path": "/x/a.docx"}, result)
        first, second = out["headers_and_footers"]
        self.assertEqual(first, {"part": "word/header1.xml", "markdown": "H"})
        self.assertEqual(second["warnings"], ["w"])
        self.assertEqual(second["lossy_elements"], [{"kind": "k"}])

    def test_error_results_and_input_are_untouched(self):
        error = {"error_code": "ZERO_MATCH", "path": "/x/a.docx", "warnings": []}
        self.assertIs(lean_read_result("read_document", {"path": "/x/a.docx"}, error), error)
        snapshot = json.loads(json.dumps(LIVE_READ))
        lean_read_result("read_document", {}, LIVE_READ)
        self.assertEqual(LIVE_READ, snapshot)

    def test_read_tools_never_overlap_mutating_tools(self):
        self.assertEqual(READ_TOOLS, READ_ONLY_TOOLS)
        self.assertFalse(READ_TOOLS & MUTATING_TOOLS)


class _FileBase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "frag.docx"
        shutil.copyfile(FIXTURES / "frag.docx", self.target)

    def tearDown(self):
        if self._old is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old
        self._tmp.cleanup()

    def call(self, tool: str, **args) -> dict:
        async def run():
            async with Client(server.mcp) as client:
                return (await client.call_tool(tool, {"path": str(self.target), **args})).structured_content

        return asyncio.run(run())


class ReadEnvelopeEndToEndTests(_FileBase):
    def test_read_document_default_is_lean_and_verbose_restores_detail(self):
        lean = self.call("read_document", format="text")
        self.assertIn("path", lean)  # JennyStack reads it
        self.assertEqual(lean["source"], "file")
        self.assertIn("text", lean)
        self.assertIn("revision", lean)
        for gone in ("section_key", "revision_detail", "warnings"):
            self.assertNotIn(gone, lean)
        full = self.call("read_document", format="text", verbose=True)
        self.assertIn("revision_detail", full)

    def test_find_sections_drops_the_echoed_path(self):
        result = self.call("find_sections")
        self.assertNotIn("path", result)
        self.assertIn("sections", result)

    def test_read_header_footer_parts_have_no_empty_noise(self):
        result = self.call("read_header_footer")
        for part in result["headers_and_footers"]:
            for key in ("warnings", "lossy_elements"):
                self.assertTrue(part.get(key, True), (key, part))

    def test_error_results_are_untouched(self):
        async def run():
            async with Client(server.mcp) as client:
                return await client.call_tool(
                    "read_document", {"path": str(self.target), "format": "bogus"}, raise_on_error=False
                )

        result = asyncio.run(run())
        self.assertTrue(result.is_error)
        self.assertIn("INVALID_INPUT", result.content[0].text)

    def test_list_open_items_keeps_path_and_source(self):
        result = self.call("list_open_items")
        self.assertIn("path", result)
        self.assertIn("source", result)


class LockStatusToolTests(_FileBase):
    def test_default_is_compact_and_verbose_restores_everything(self):
        compact = self.call("lock_status")
        self.assertEqual(set(compact), {"owner_file", "desktop_word", "sync_quiesced", "external_activity"})
        self.assertEqual(set(compact["external_activity"]), {"divergent", "still_current", "recent", "ledger"})
        full = self.call("lock_status", verbose=True)
        self.assertIn("sync_detail", full)
        self.assertNotIn("path", full)  # the lean read envelope drops an echoed path
        self.assertIn("note", full["external_activity"])
        self.assertIn("ledger", full["external_activity"])

    def test_internal_callers_still_get_the_full_dict(self):
        result = server.execute_lock_status(str(self.target), quiesce_interval=0.01)
        for key in ("path", "sync_detail", "external_activity", "sync_quiesced"):
            self.assertIn(key, result)
        self.assertIn("note", result["external_activity"])


class AnnotationTests(unittest.TestCase):
    def tools(self) -> dict:
        async def run():
            async with Client(server.mcp) as client:
                return {t.name: t for t in await client.list_tools()}

        return asyncio.run(run())

    def test_every_read_tool_is_read_only_and_idempotent(self):
        tools = self.tools()
        for name in READ_ONLY_TOOLS:
            ann = tools[name].annotations
            self.assertIsNotNone(ann, name)
            self.assertIs(ann.readOnlyHint, True, name)
            self.assertIs(ann.idempotentHint, True, name)

    def test_no_mutating_tool_is_marked_read_only(self):
        tools = self.tools()
        for name in set(tools) - READ_TOOLS:
            ann = tools[name].annotations
            self.assertFalse(ann is not None and ann.readOnlyHint, name)
        for name in MUTATING_TOOLS:
            self.assertNotIn(name, READ_TOOLS)

    def test_destructive_tools_say_so(self):
        tools = self.tools()
        for name in DESTRUCTIVE_TOOLS:
            self.assertIs(tools[name].annotations.destructiveHint, True, name)
        for name in set(tools) - DESTRUCTIVE_TOOLS:
            ann = tools[name].annotations
            self.assertFalse(ann is not None and ann.destructiveHint, name)

    def test_instructions_carry_the_workflow_lines(self):
        text = server.mcp.instructions or ""
        for fragment in ('format="markdown"', "replace_text_batch", "revision_after", "live_revision", "write_mode"):
            self.assertIn(fragment, text)


if __name__ == "__main__":
    unittest.main()
