"""Issue #79: live table clean-up -- numbering inherited by a new table, clear_numbering,
delete_table and table-scoped accept/reject, through the real bridge against tests/unit/fake_pane.py.

Pins the server's side of each contract (validation before anything is sent, capability gates,
independent read-back, evidence). The pane's Office.js side is covered by tests/unit/js/tables_harness.mjs.
"""

from __future__ import annotations

import asyncio
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, FakePane
from test_live_tables import _TableCase

from verified_docx_mcp import server
from verified_docx_mcp.errors import ErrorCode
from verified_docx_mcp.live import paragraphs_live, revisions_live
from verified_docx_mcp.middleware import MUTATING_TOOLS

ROWS = [["R1C1", "R1C2"], ["R2C1", "R2C2"]]
TABLE_TEXT = "R1C1\tR1C2\nR2C1\tR2C2"


def _two_tables() -> FakeDocument:
    doc = FakeDocument(text="Intro\nProof\nTail")
    doc.tables = [[["A1", "A2"]], [["B1", "B2"]]]
    return doc


def _table_doc() -> FakeDocument:
    doc = FakeDocument(text="Intro")
    doc.tables = [[list(r) for r in ROWS]]
    return doc


def _caps(*, drop: tuple[str, ...] = (), add: tuple[str, ...] = ()) -> list[str]:
    return [*(c for c in FakePane().capabilities if c not in drop), *add]


class _CleanupCase(_TableCase):
    async def clear(self, **kwargs):
        return await asyncio.to_thread(paragraphs_live.clear_numbering, str(self.target), **kwargs)

    async def delete(self, table_id, expected_text, **kwargs):
        return await asyncio.to_thread(
            paragraphs_live.delete_table, str(self.target), table_id, expected_text, **kwargs
        )

    async def mutate(self, action, **kwargs):
        ids = kwargs.pop("ids", None)
        revision_before = kwargs.pop("revision_before", None)
        return await asyncio.to_thread(
            revisions_live.mutate_revisions, str(self.target), action, ids, revision_before, **kwargs
        )


class InsertTableNumberingTests(_CleanupCase):
    async def test_pane_reporting_inherited_numbering_fails_verification(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.inherit_numbering = True
        await self.connect_pane(doc)
        envelope = await self.refused(ErrorCode.VERIFICATION_FAILED, self.insert, ROWS)
        self.assertIn("list items", envelope.message)
        self.assertIn("clear_numbering", envelope.message)

    async def test_pane_reporting_zero_list_items_passes_and_reports_it(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        await self.connect_pane(doc)
        evidence = await self.insert(ROWS)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["list_items_in_table"], 0)

    async def test_pane_that_does_not_report_the_field_is_not_failed(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.inherit_numbering = True  # would be caught, but this pane cannot say so
        doc.report_list_items = False
        await self.connect_pane(doc)
        evidence = await self.insert(ROWS)
        self.assertTrue(evidence["applied"])
        self.assertNotIn("list_items_in_table", evidence)


class ClearNumberingTests(_CleanupCase):
    async def test_table_scope_detaches_and_reports_the_count(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.tables = [[["a", "b"], ["c", "d"]]]
        doc.list_items = {0: 4}
        pane = await self.connect_pane(doc)
        before = doc.sha256()
        evidence = await self.clear(table_id=1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["unlisted"], 4)
        self.assertEqual(evidence["scope"], {"table_id": 1})
        self.assertEqual(doc.list_items, {})
        self.assertEqual(evidence["revision_before"], f"live:sha256:{before}")
        self.assertEqual(evidence["revision_after"], f"live:sha256:{doc.sha256()}")
        self.assertEqual(pane.op_payloads["paragraphs_unlist"][0]["table_index"], 1)

    async def test_paragraph_scope(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.list_paragraphs = {"Proof"}
        await self.connect_pane(doc)
        evidence = await self.clear(paragraph_text="Proof")
        self.assertEqual(evidence["unlisted"], 1)
        self.assertEqual(doc.list_paragraphs, set())

    async def test_nothing_to_detach_is_a_verified_noop(self) -> None:
        await self.connect_pane(_two_tables())
        evidence = await self.clear(table_id=2)
        self.assertEqual(evidence["unlisted"], 0)

    async def test_exactly_one_scope_is_required(self) -> None:
        pane = await self.connect_pane(_two_tables())
        await self.refused(ErrorCode.INVALID_INPUT, self.clear)
        await self.refused(ErrorCode.INVALID_INPUT, self.clear, table_id=1, paragraph_text="Proof")
        await self.refused(ErrorCode.INVALID_INPUT, self.clear, table_id=0)
        self.assertNotIn("paragraphs_unlist", pane.op_payloads)

    async def test_ambiguous_or_missing_paragraph_is_match_count_mismatch(self) -> None:
        await self.connect_pane(FakeDocument(text="Dup\nDup\nTail"))
        await self.refused(ErrorCode.MATCH_COUNT_MISMATCH, self.clear, paragraph_text="Dup")
        await self.refused(ErrorCode.MATCH_COUNT_MISMATCH, self.clear, paragraph_text="Nope")

    async def test_missing_table_and_tracked_changes_are_refused(self) -> None:
        doc = _two_tables()
        doc.revisions = [{"kind": "ins", "author": "Colleague", "text": "x", "table": 0}]
        await self.connect_pane(doc)
        await self.refused(ErrorCode.TABLE_NOT_FOUND, self.clear, table_id=9)
        await self.refused(ErrorCode.TRACKED_CHANGES_PRESENT, self.clear, table_id=1)

    async def test_a_detach_that_did_not_take_fails_verification(self) -> None:
        doc = _two_tables()
        doc.list_items = {0: 2}
        doc.drop_unlist = True
        await self.connect_pane(doc)
        await self.refused(ErrorCode.VERIFICATION_FAILED, self.clear, table_id=1)

    async def test_old_pane_is_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect_pane(_two_tables(), capabilities=_caps(drop=("paragraphs_unlist",)))
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.clear, table_id=1)
        self.assertNotIn("paragraphs_unlist", pane.op_payloads)

    async def test_stale_revision_is_refused(self) -> None:
        await self.connect_pane(_two_tables())
        await self.refused(ErrorCode.LIVE_STALE, self.clear, table_id=1, revision_before="live:sha256:deadbeef")

    async def test_file_mode_is_refused(self) -> None:
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, self.clear, table_id=1, write_mode="file")
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, self.clear, table_id=1, write_mode="auto")


class DeleteTableTests(_CleanupCase):
    async def test_matching_text_deletes_and_neighbours_survive(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.tables = [[list(r) for r in ROWS], [["keep"]]]
        pane = await self.connect_pane(doc)
        evidence = await self.delete(1, TABLE_TEXT)
        self.assertTrue(evidence["applied"])
        self.assertEqual((evidence["tables_before"], evidence["tables_after"]), (2, 1))
        self.assertEqual(doc.tables, [[["keep"]]])
        self.assertEqual(doc.text, "Intro\nProof\nTail")
        self.assertEqual(evidence["before"], TABLE_TEXT)
        self.assertEqual(pane.op_payloads["table_delete"][0]["table_index"], 1)

    async def test_whitespace_differences_in_expected_text_are_tolerated(self) -> None:
        doc = _table_doc()
        await self.connect_pane(doc)
        evidence = await self.delete(1, "R1C1  R1C2\n\nR2C1 R2C2 ")
        self.assertEqual(evidence["tables_after"], 0)

    async def test_text_mismatch_refuses_and_writes_nothing(self) -> None:
        doc = _table_doc()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.CELL_TEXT_MISMATCH, self.delete, 1, "something else")
        self.assertEqual(len(doc.tables), 1)

    async def test_anchored_shape_and_missing_table_are_refused(self) -> None:
        doc = _table_doc()
        doc.shape_tables = {0}
        await self.connect_pane(doc)
        await self.refused(ErrorCode.ANCHORED_SHAPES, self.delete, 1, TABLE_TEXT)
        await self.refused(ErrorCode.TABLE_NOT_FOUND, self.delete, 5, TABLE_TEXT)
        self.assertEqual(len(doc.tables), 1)

    async def test_input_validation_happens_before_anything_is_sent(self) -> None:
        pane = await self.connect_pane(_table_doc())
        await self.refused(ErrorCode.INVALID_INPUT, self.delete, 0, TABLE_TEXT)
        await self.refused(ErrorCode.INVALID_INPUT, self.delete, 1, "  ")
        self.assertNotIn("table_delete", pane.op_payloads)

    async def test_a_delete_the_pane_claims_but_did_not_do_fails_verification(self) -> None:
        doc = _table_doc()
        doc.drop_table_delete = True
        await self.connect_pane(doc)
        await self.refused(ErrorCode.VERIFICATION_FAILED, self.delete, 1, TABLE_TEXT)

    async def test_old_pane_is_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect_pane(_table_doc(), capabilities=_caps(drop=("table_delete",)))
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.delete, 1, TABLE_TEXT)
        self.assertNotIn("table_delete", pane.op_payloads)

    async def test_stale_revision_is_refused(self) -> None:
        doc = _table_doc()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.LIVE_STALE, self.delete, 1, TABLE_TEXT, revision_before="live:sha256:deadbeef")
        self.assertEqual(len(doc.tables), 1)

    async def test_file_mode_is_refused(self) -> None:
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, self.delete, 1, TABLE_TEXT, write_mode="file")
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, self.delete, 1, TABLE_TEXT, write_mode="auto")


class ScopedRevisionTests(_CleanupCase):
    def _doc(self) -> FakeDocument:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.tables = [[["A1"]], [["B1"]]]
        doc.revisions = [
            {"kind": "ins", "author": "Colleague", "text": "elsewhere", "table": None},
            {"kind": "ins", "author": "Colleague", "text": "in table 2", "table": 1},
            {"kind": "del", "author": "Colleague", "text": "also table 2", "table": 1},
        ]
        return doc

    async def test_scoped_accept_touches_only_that_table(self) -> None:
        doc = self._doc()
        pane = await self.connect_pane(doc, capabilities=_caps(add=("live_revisions",)))
        evidence = await self.mutate("accept", table_id=2)
        self.assertEqual(evidence["changed_count"], 2)
        self.assertFalse(evidence["table_removed"])
        self.assertEqual([r["text"] for r in doc.revisions], ["elsewhere"])
        self.assertEqual(pane.op_payloads["revisions_accept"][0]["table_index"], 2)

    async def test_rejecting_a_tracked_table_insertion_removes_the_table(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        doc.revisions = [{"kind": "ins", "author": "Colleague", "text": "elsewhere", "table": None}]
        await self.connect_pane(doc, capabilities=_caps(add=("live_revisions",)))
        inserted = await self.insert(ROWS, track_changes=True, anchor={"paragraph_text": "Proof"})
        self.assertEqual(inserted["table_id"], 1)
        self.assertEqual(len(doc.revisions), 2)
        text_before = doc.text
        evidence = await self.mutate("reject", table_id=1)
        self.assertTrue(evidence["table_removed"])
        self.assertEqual(doc.tables, [])
        self.assertEqual([r["text"] for r in doc.revisions], ["elsewhere"])  # the colleague's change survives
        self.assertEqual(doc.text, text_before)  # the anchor paragraph's text is untouched

    async def test_scoped_request_needs_the_revisions_scoped_capability(self) -> None:
        doc = self._doc()
        pane = await self.connect_pane(doc, capabilities=_caps(drop=("revisions_scoped",), add=("live_revisions",)))
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.mutate, "reject", table_id=2)
        self.assertNotIn("revisions_reject", pane.op_payloads)
        self.assertEqual(len(doc.revisions), 3)

    async def test_table_id_and_revision_ids_are_mutually_exclusive(self) -> None:
        pane = await self.connect_pane(self._doc(), capabilities=_caps(add=("live_revisions",)))
        await self.refused(ErrorCode.INVALID_INPUT, self.mutate, "reject", ids=["revision:x"], table_id=2)
        await self.refused(ErrorCode.INVALID_INPUT, self.mutate, "reject", table_id=0)
        self.assertNotIn("revisions_reject", pane.op_payloads)

    async def test_unscoped_payload_carries_no_table_index(self) -> None:
        doc = self._doc()
        pane = await self.connect_pane(doc, capabilities=_caps(drop=("revisions_scoped",), add=("live_revisions",)))
        await self.mutate("accept")
        self.assertNotIn("table_index", pane.op_payloads["revisions_accept"][0])
        self.assertEqual(doc.revisions, [])

    async def test_file_mode_refuses_table_id(self) -> None:
        tool = getattr(server.reject_tracked_changes, "fn", server.reject_tracked_changes)
        with self.assertRaises(Exception) as cm:
            await asyncio.to_thread(tool, str(self.target), table_id=1, write_mode="file")
        self.assertIn("live only", str(cm.exception))


class RegistrationTests(unittest.TestCase):
    def test_new_tools_are_mutating_and_registered(self) -> None:
        self.assertIn("clear_numbering", MUTATING_TOOLS)
        self.assertIn("delete_table", MUTATING_TOOLS)
        for name in ("clear_numbering", "delete_table"):
            self.assertTrue(hasattr(server, name), name)

    def test_fake_pane_reports_the_new_capabilities(self) -> None:
        for name in ("table_delete", "paragraphs_unlist", "revisions_scoped"):
            self.assertIn(name, FakePane().capabilities)


if __name__ == "__main__":
    unittest.main()
