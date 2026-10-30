"""Issue #111: pin every field JennyStack reads from list_open_items / list_tables / list_styles
in DEFAULT mode, through the real FastMCP server.

JennyStack (scripts/live_doc_stamp.py, live_doc_guard.py, comment_dispositions.py,
live_guard_state.py, live_revision.py, section_identity.py) consumes these results. Payload
trimming must never remove a field it reads from a default response. What it reads, with the
JennyStack source of each read:

list_open_items (live or file):
  top level  `comments` (a list; live_doc_stamp.record_list_items drops the response otherwise),
             `source` ("live"/"file"; live_doc_stamp stamps it), `path`, `pending_suggestions`
  per comment `comment_id` (comment_dispositions.load_items), `author`, `resolved` (`is True` rows are
             skipped), the anchor as `quoted_text` OR `anchor_text` OR `anchor`; `reply_count` and
             `replies` are OPTIONAL there (comment_dispositions.py:171-175, 189: a missing count is 0,
             a missing `replies` is [], the thread size is max(count, len(replies))), so full mode
             omits `reply_count` and an empty `replies` (#111).
  `fields="summary"` is columnar (comment_columns + row lists), which that loader does NOT read:
             JennyStack must keep asking for `full` until it learns the columnar shape.
  tracked-change rows are recognised by `suggestion_id`, or `kind` in (insertion, deletion), or
  `type` (live: Inserted/Deleted...).
list_tables / list_styles: no JennyStack script reads either today; the long-standing fields are
pinned anyway so a future reader is not surprised.

If one of these assertions has to change, JennyStack's reader has to change in the same breath.
"""

from __future__ import annotations

import asyncio
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeComment, FakeDocument, FakeReply, docx_to_flat_opc
from fastmcp import Client
from test_live_comments import LiveCommentsTestCase

from verified_docx_mcp import mutations, paths, server

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

COMMENT_KEYS = {"comment_id", "author", "resolved"}
SUMMARY_COLUMNS = ["comment_id", "author", "resolved", "anchor", "reply_count"]


def jennystack_reply_count(item: dict) -> int:
    """JennyStack comment_dispositions.py load_items, lines 171-175 and 189, reduced to the reply
    count: `_count(item.get("reply_count"))` (0 when absent) maxed with the ids found in a
    `replies` list that is only used when it is a list."""
    nested = item["replies"] if isinstance(item.get("replies"), list) else []
    declared = int(item.get("reply_count") or 0)
    return max(declared, len(nested))


async def call(tool: str, args: dict) -> dict:
    async with Client(server.mcp) as client:
        return (await client.call_tool(tool, args)).structured_content


def anchor_of(item: dict) -> str:
    """comment_dispositions.load_items' anchor lookup."""
    return item.get("quoted_text") or item.get("anchor_text") or item.get("anchor") or ""


class _FileCase(unittest.TestCase):
    fixture_name = "comments/golden-comment.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def call(self, tool: str, args: dict | None = None) -> dict:
        return asyncio.run(call(tool, {"path": str(self.target), **(args or {})}))


class ListOpenItemsFileContractTests(_FileCase):
    def test_default_comment_rows_carry_every_field_jennystack_reads(self):
        result = self.call("list_open_items")
        self.assertEqual(result["source"], "file")
        self.assertIsInstance(result["comments"], list)
        self.assertIn("path", result)
        self.assertIn("pending_suggestions", result)
        root = result["comments"][0]
        self.assertTrue(COMMENT_KEYS <= set(root), set(root))
        self.assertEqual(root["comment_id"], "654503C0")
        self.assertTrue(anchor_of(root))
        self.assertIn("quoted_text", root)  # file mode: the anchor JennyStack reads first
        self.assertIs(root["resolved"], False)
        self.assertNotIn("reply_count", root)
        self.assertEqual(jennystack_reply_count(root), 1)  # still correct without reply_count
        self.assertTrue(root["replies"][0].get("id") or root["replies"][0].get("comment_id"))

    def test_a_comment_without_replies_still_reads_as_zero(self):
        shutil.copyfile(FIXTURES / "revision/commented.docx", self.target)
        row = self.call("list_open_items")["comments"][0]
        self.assertNotIn("replies", row)
        self.assertNotIn("reply_count", row)
        self.assertEqual(jennystack_reply_count(row), 0)

    def test_summary_is_columnar_with_complete_ids(self):
        result = self.call("list_open_items", {"fields": "summary"})
        self.assertEqual(result["comment_columns"], SUMMARY_COLUMNS)
        row = dict(zip(result["comment_columns"], result["comments"][0], strict=True))
        self.assertEqual(row["comment_id"], "654503C0")
        self.assertEqual(row["reply_count"], 1)
        self.assertTrue(row["anchor"])
        self.assertEqual(result["source"], "file")


class ListOpenItemsTrackedChangeFileContractTests(_FileCase):
    fixture_name = "revision/tracked.docx"

    def test_suggestion_rows_keep_the_keys_that_mark_a_tracked_change(self):
        result = self.call("list_open_items")
        rows = result["pending_suggestions"]
        self.assertEqual({r["kind"] for r in rows}, {"insertion", "deletion"})
        for row in rows:
            self.assertIn("suggestion_id", row)
            self.assertIn("author", row)
        summary = self.call("list_open_items", {"fields": "summary"})
        self.assertEqual(summary["suggestion_columns"], ["suggestion_id", "kind", "author", "anchor"])
        self.assertEqual({r[1] for r in summary["suggestions"]}, {"insertion", "deletion"})


class ListOpenItemsLiveContractTests(LiveCommentsTestCase):
    fixture_name = "sections.docx"

    def _document(self) -> FakeDocument:
        doc = FakeDocument(text="", author_name="Pat Reviewer")
        doc.session_epoch = "ep1"
        doc.comments = [
            FakeComment(
                id="c1", content="please check", author_name="Pat Reviewer",
                creation_date="2026-10-01T12:00:00.000Z", anchor_text="overview",
                anchor_paragraph_text="Some overview text.",
                replies=[FakeReply(id="r1", content="agreed", author_name="Lee", creation_date="2026-10-01T13:00:00.000Z")],
            )
        ]
        return doc

    async def test_default_live_rows_carry_every_field_jennystack_reads(self):
        await self.connect_pane(document=self._document())
        result = await call("list_open_items", {"path": str(self.target), "source": "live"})
        self.assertEqual(result["source"], "live")
        self.assertIsInstance(result["comments"], list)
        self.assertIn("path", result)
        self.assertIn("pending_suggestions", result)
        row = result["comments"][0]
        self.assertTrue(COMMENT_KEYS <= set(row), set(row))
        self.assertEqual(row["comment_id"], "live:ep1:c1")
        self.assertEqual(anchor_of(row), "overview")  # live: anchor_text (quoted_text is gone)
        self.assertIs(row["resolved"], False)
        self.assertNotIn("reply_count", row)
        self.assertEqual(jennystack_reply_count(row), 1)
        self.assertTrue(row["replies"][0].get("id") or row["replies"][0].get("comment_id"))

    async def test_live_summary_is_columnar_with_complete_ids(self):
        await self.connect_pane(document=self._document())
        result = await call("list_open_items", {"path": str(self.target), "source": "live", "fields": "summary"})
        self.assertEqual(result["comment_columns"], SUMMARY_COLUMNS)
        row = dict(zip(result["comment_columns"], result["comments"][0], strict=True))
        self.assertEqual(
            (row["comment_id"], row["anchor"], row["reply_count"], row["resolved"]), ("live:ep1:c1", "overview", 1, False)
        )
        self.assertEqual(result["source"], "live")


class ListTablesAndStylesContractTests(_FileCase):
    fixture_name = "approved-changes.docx"

    def test_list_tables_keeps_its_long_standing_fields(self):
        table = self.call("list_tables")["tables"][0]
        for key in ("table_id", "row_count", "col_count", "has_merged_cells", "has_nested_table",
                    "nested_in_table_id", "header_row", "anchor_section_key"):
            self.assertIn(key, table)

    def test_list_styles_default_returns_every_style_with_its_fields(self):
        styles = self.call("list_styles")["styles"]
        self.assertGreater(len(styles), 25)
        for key in ("style_id", "name", "type", "outline_lvl"):
            self.assertIn(key, styles[0])
        filtered = self.call("list_styles", {"type": "paragraph", "name_contains": "Heading"})["styles"]
        self.assertTrue(filtered)
        self.assertLess(len(filtered), len(styles))


class ReadDocumentContractTests(_FileCase):
    """JennyStack live_doc_stamp.py / live_guard_state.resolve_read_response read, from read_document:
    `path` (must be a str naming the document), `source`, `format`, `text` (or `content` / `markdown`)
    and `revision`, plus `truncated` / `is_truncated` / `has_more` / `isError` (all absent = fine).
    The lean read envelope must never drop them."""

    fixture_name = "frag.docx"

    def test_text_read_keeps_every_field_jennystack_reads(self):
        result = self.call("read_document", {"format": "text"})
        for key in ("path", "source", "format", "text", "revision"):
            self.assertIn(key, result)
        self.assertIsInstance(result["path"], str)
        self.assertEqual(result["source"], "file")
        self.assertEqual(result["format"], "text")
        self.assertIsInstance(result["text"], str)

    def test_markdown_read_keeps_format_and_markdown(self):
        result = self.call("read_document")
        for key in ("path", "source", "format", "markdown", "revision"):
            self.assertIn(key, result)

    def test_nothing_jennystack_reads_is_behind_verbose(self):
        lean, full = self.call("read_document", {"format": "text"}), self.call("read_document", {"format": "text", "verbose": True})
        self.assertTrue(set(lean) <= set(full))
        self.assertEqual(set(full) - set(lean), {"revision_detail", "section_key", "warnings"} & set(full))


class ReadDocumentLiveContractTests(LiveCommentsTestCase):
    fixture_name = "sections.docx"

    async def test_live_text_read_keeps_text_revision_source_path_format(self):
        await self.connect_pane(body_ooxml=docx_to_flat_opc(FIXTURES / "sections.docx"))
        result = await call("read_document", {"path": str(self.target), "source": "live", "format": "text"})
        for key in ("path", "source", "format", "text", "revision"):
            self.assertIn(key, result)
        self.assertEqual(result["source"], "live")
        self.assertTrue(result["revision"].startswith("live:sha256:"))
        self.assertNotIn("body_sha256", result.get("live", {}))


class OtherReadToolContractTests(_FileCase):
    """No JennyStack script reads find_sections / list_tables / get_table / read_header_footer /
    live_status / lock_status / live_revision / live_list_sections results today (surveyed in
    scripts/ and cli/tests/test_live_doc_*.py). Their payload fields that matter are pinned anyway."""

    fixture_name = "approved-changes.docx"

    def test_find_sections_and_tables_keep_their_payload(self):
        self.assertIn("sections", self.call("find_sections"))
        self.assertIn("tables", self.call("list_tables"))
        self.assertIn("headers_and_footers", self.call("read_header_footer"))

    def test_lock_status_default_keeps_the_guard_booleans(self):
        result = self.call("lock_status")
        self.assertIn("present", result["owner_file"])
        self.assertIn("sync_quiesced", result)
        self.assertEqual(set(result["external_activity"]), {"divergent", "still_current"})


if __name__ == "__main__":
    unittest.main()
