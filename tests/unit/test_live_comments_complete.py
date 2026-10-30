"""Issue #39: `list_open_items(source="live")` `complete` flag, `counts.ooxml_open` and the
`pane_collection_incomplete` warning, driven by the fake pane's configurable `ooxml` block.

The pane's `ooxml` block (open threads counted from the body OOXML) is the only independent
check on `body.getComments()`; these tests pin the server's reading of it: true / false (with a
single retry) / null (old pane or an unreadable comments part), and that the change is additive.
Comment ids: Office.js ids and OOXML `w:id` values are different id spaces on desktop Word, so
`missing_comment_ids` is only a diff when the two spaces overlap (see comments_live).
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeComment, FakeDocument
from test_live_comments import LiveCommentsTestCase

from verified_docx_mcp.live import comments_live


def _doc(*ids: str) -> FakeDocument:
    doc = FakeDocument(text="body")
    doc.comments = [
        FakeComment(
            id=i, content=f"note {i}", author_name="A", creation_date="2026-09-16T14:06:00Z",
            anchor_text="body", resolved=False,
        )
        for i in ids
    ]
    return doc


def _ooxml(open_threads: int, ids: list[str], *, total: int | None = None) -> dict:
    return {"available": True, "total_threads": total if total is not None else open_threads,
            "open_threads": open_threads, "ids": ids}


UNAVAILABLE = {"available": False, "total_threads": None, "open_threads": None, "ids": [],
               "reason": "commentsExtended part missing"}


class CompleteFlagTests(LiveCommentsTestCase):
    async def _list(self) -> dict:
        return await self.call(comments_live.execute_list_open_items_live, str(self.target))

    async def test_complete_true_when_counts_agree(self) -> None:
        doc = _doc("1", "2")
        doc.ooxml = _ooxml(2, ["1", "2"])
        pane = await self.connect_pane(document=doc)

        result = await self._list()
        self.assertIs(result["complete"], True)
        self.assertEqual(result["counts"], {"live_total": 2, "live_open": 2, "returned": 2, "ooxml_open": 2})
        self.assertNotIn("pane_collection_incomplete", [w["code"] for w in result.get("warnings", [])])
        self.assertEqual(len(pane.comments_list_payloads), 1)  # no retry when nothing is wrong

    async def test_fewer_ooxml_threads_than_listed_is_still_complete(self) -> None:
        doc = _doc("1", "2")
        doc.ooxml = _ooxml(1, ["1"])
        await self.connect_pane(document=doc)
        self.assertIs((await self._list())["complete"], True)

    async def test_complete_false_warns_with_counts_and_missing_ids_after_one_retry(self) -> None:
        doc = _doc("1", "2")
        doc.ooxml = _ooxml(3, ["1", "2", "9"])
        pane = await self.connect_pane(document=doc)

        result = await self._list()
        self.assertIs(result["complete"], False)
        self.assertEqual(result["counts"]["ooxml_open"], 3)
        self.assertEqual(result["counts"]["live_open"], 2)
        warning = next(w for w in result.get("warnings", []) if w["code"] == "pane_collection_incomplete")
        self.assertIn("2", warning["message"])
        self.assertIn("3", warning["message"])
        self.assertEqual(warning["missing_comment_ids"], ["9"])
        self.assertEqual(warning["ooxml_open_ids"], ["1", "2", "9"])
        self.assertTrue(warning["retried"])
        self.assertEqual(len(pane.comments_list_payloads), 2)  # exactly one retry, not a loop

    async def test_retry_that_heals_reports_complete(self) -> None:
        doc = _doc("1", "2")
        doc.ooxml = [_ooxml(3, ["1", "2", "9"]), _ooxml(2, ["1", "2"])]
        pane = await self.connect_pane(document=doc)

        result = await self._list()
        self.assertIs(result["complete"], True)
        self.assertEqual(len(pane.comments_list_payloads), 2)
        self.assertNotIn("pane_collection_incomplete", [w["code"] for w in result.get("warnings", [])])

    async def test_uncorrelated_id_spaces_leave_missing_ids_unknown(self) -> None:
        doc = _doc("guid-a")
        doc.ooxml = _ooxml(2, ["7", "8"])
        await self.connect_pane(document=doc)

        warning = next(w for w in (await self._list())["warnings"] if w["code"] == "pane_collection_incomplete")
        self.assertIsNone(warning["missing_comment_ids"])
        self.assertEqual(warning["ooxml_open_ids"], ["7", "8"])

    async def test_no_returned_comments_but_open_threads_lists_all_ids_as_missing(self) -> None:
        doc = _doc()
        doc.ooxml = _ooxml(2, ["4", "5"])
        await self.connect_pane(document=doc)

        result = await self._list()
        self.assertIs(result["complete"], False)
        warning = next(w for w in result.get("warnings", []) if w["code"] == "pane_collection_incomplete")
        self.assertEqual(warning["missing_comment_ids"], ["4", "5"])

    async def test_complete_null_for_a_pane_without_the_crosscheck(self) -> None:
        pane = await self.connect_pane(document=_doc("1"), capabilities=[])

        result = await self._list()
        self.assertIsNone(result["complete"])
        self.assertNotIn("ooxml_open", result["counts"])
        self.assertNotIn("pane_collection_incomplete", [w["code"] for w in result.get("warnings", [])])
        self.assertEqual(len(pane.comments_list_payloads), 1)

    async def test_complete_null_when_the_pane_could_not_read_the_comments_parts(self) -> None:
        doc = _doc("1")
        doc.ooxml = dict(UNAVAILABLE)
        await self.connect_pane(document=doc)

        result = await self._list()
        self.assertIsNone(result["complete"])
        self.assertNotIn("ooxml_open", result["counts"])

    async def test_existing_fields_are_unchanged_and_new_ones_additive(self) -> None:
        doc = _doc("1")
        doc.ooxml = _ooxml(1, ["1"])
        await self.connect_pane(document=doc)

        result = await self._list()
        self.assertEqual(list(result)[:4], ["counts", "scope", "pane", "path"])
        for key in ("path", "source", "comments", "pending_suggestions", "file_side_available", "revision_coverage"):
            self.assertIn(key, result)
        # null / empty top-level keys are omitted (#111): no epoch on this pane, nothing file-only
        for key in ("session_epoch", "file_only_open_comments", "warnings", "pending_suggestions_source"):
            self.assertNotIn(key, result)
        self.assertNotIn("correlation", result)  # opt-in since #111
        self.assertEqual([c["comment_id"] for c in result["comments"]], ["live:1"])
        self.assertEqual(result["scope"], "body")


if __name__ == "__main__":
    unittest.main()
