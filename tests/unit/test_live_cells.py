"""Unit tests for issue #27's live table-cell edit: replace_cell_markdown
with write_mode="live" over the WSS ops channel (cell_get / cell_set), the
context-free markdown subset parser it uses, and the two new protocol
payloads.

Same harness as test_live_write_mode.py: a real bridge on EPHEMERAL ports
with a connected fake_pane.FakePane, exercised end to end. The fake's table
model proves the SERVER's behavior (routing, capability gate, subset
parsing, compare-and-set handling, independent read-back); it cannot prove
the real Office JS object-model path in addin/taskpane.js -- that needs the
manual sideload check in docs/live-mode.md.
"""

from __future__ import annotations

import asyncio
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import markdown_to_ooxml, mutations, tables
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import protocol

mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _two_by_two() -> FakeDocument:
    doc = FakeDocument(text="body")
    doc.tables = [[["R1C1", "R1C2"], ["R2C1", "R2C2"]]]
    return doc


class ReplaceCellLiveTests(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def replace(self, markdown: str, *, row: int = 1, cell: int = 1, **kwargs):
        return await asyncio.to_thread(
            tables.execute_replace_cell_markdown, str(self.target), 1, row, cell, markdown, **kwargs
        )

    async def assertRefusedWith(self, code: ErrorCode, markdown: str = "x", **kwargs):
        with self.assertRaises(VerifyError) as cm:
            await self.replace(markdown, **kwargs)
        self.assertEqual(cm.exception.envelope.error_code, code)
        return cm.exception.envelope

    async def test_auto_routes_to_live_and_leaves_the_file_alone(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        evidence = await self.replace("**Bold** and *italic*")

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["before"], "R1C1")
        self.assertEqual(evidence["after"], "Bold and italic")
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertEqual(doc.tables[0][0][0], "Bold and italic")
        self.assertEqual(doc.tables[0][1][1], "R2C2")  # neighbours untouched
        self.assertEqual(self.target.read_bytes(), before_bytes)

    async def test_multiple_paragraphs(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        evidence = await self.replace("one\n\ntwo", row=2, cell=2)
        self.assertTrue(evidence["applied"])
        self.assertEqual(doc.tables[0][1][1], "one\ntwo")

    async def test_wire_payload_carries_runs(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        seen: list[dict] = []
        real = doc.cell_set

        def spy(payload):
            seen.append(payload)
            return real(payload)

        doc.cell_set = spy
        await self.replace("plain **b** [site](https://example.com/x)")
        runs = seen[0]["paragraphs"][0]
        self.assertEqual(
            [(r["text"], r["bold"], r["link"]) for r in runs],
            [("plain ", False, None), ("b", True, None), (" ", False, None), ("site", False, "https://example.com/x")],
        )
        self.assertEqual(seen[0]["expected_before_text"], "R1C1")

    async def test_block_structure_is_refused_before_anything_is_sent(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        for markdown in ("- a\n- b", "1. a\n2. b", "# Heading", "| a | b |\n|---|---|\n| 1 | 2 |", "> quote", "---"):
            with self.subTest(markdown=markdown):
                envelope = await self.assertRefusedWith(ErrorCode.INVALID_INPUT, markdown)
                self.assertIn("write_mode='file'", envelope.message)
        self.assertEqual(doc.tables[0][0][0], "R1C1")

    async def test_pane_without_cell_edit_capability_is_refused(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc, capabilities=["row_scope"])
        await self.assertRefusedWith(ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertEqual(doc.tables[0][0][0], "R1C1")

    async def test_concurrent_edit_to_the_same_cell_is_refused_not_overwritten(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        # A co-author types in the cell between the server's read and its write.
        doc.before_cell_set = lambda: doc.tables[0][0].__setitem__(0, "co-author was here")
        envelope = await self.assertRefusedWith(ErrorCode.LIVE_OP_FAILED, "mine")
        self.assertIn("changed since it was read", envelope.message)
        self.assertEqual(doc.tables[0][0][0], "co-author was here")

    async def test_a_write_the_pane_reports_but_did_not_apply_fails_the_independent_readback(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        doc.dropped_cell_writes.add((0, 0, 0))  # applied=true in the reply, but the cell is unchanged
        envelope = await self.assertRefusedWith(ErrorCode.VERIFICATION_FAILED, "mine")
        self.assertIn("Nothing to roll back", envelope.message)
        self.assertEqual(envelope.diagnostics["actual"], "R1C1")

    async def test_nested_table_document_is_refused(self) -> None:
        doc = _two_by_two()
        doc.has_nested_table = True
        await self.connect_pane(doc)
        envelope = await self.assertRefusedWith(ErrorCode.LIVE_OP_FAILED)
        self.assertIn("nested table", envelope.message)

    async def test_out_of_range_addresses_are_refused(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.assertRefusedWith(ErrorCode.LIVE_OP_FAILED, row=9)
        await self.assertRefusedWith(ErrorCode.LIVE_OP_FAILED, cell=9)
        with self.assertRaises(VerifyError):
            await asyncio.to_thread(tables.execute_replace_cell_markdown, str(self.target), 7, 1, 1, "x")

    async def test_stale_live_revision_before_is_refused(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.assertRefusedWith(ErrorCode.LIVE_STALE, revision_before="live:sha256:" + "0" * 64)
        self.assertEqual(doc.tables[0][0][0], "R1C1")

    async def test_chained_live_revisions(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        first = await self.replace("one")
        second = await self.replace("two", revision_before=first["revision_after"])
        self.assertTrue(second["applied"])
        self.assertEqual(doc.tables[0][0][0], "two")

    async def test_explicit_file_mode_with_a_session_still_refuses(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.assertRefusedWith(ErrorCode.LIVE_SESSION_ACTIVE, write_mode="file")

    async def test_explicit_live_mode_without_a_session_is_unavailable(self) -> None:
        await self.assertRefusedWith(ErrorCode.LIVE_UNAVAILABLE, write_mode="live")

    async def test_auto_without_a_session_writes_the_file(self) -> None:
        evidence = await self.replace("file written")
        self.assertEqual(evidence["write_mode"], "file")
        self.assertTrue(evidence["ledger_logged"])


class ReplaceCellToolLayerTests(LiveWriteBridgeTestCase):
    """Through the real @mcp.tool wrapper (FastMCP Client), so a wrapper that
    forgets to pass write_mode through is caught -- the execute_* tests above
    never touch it."""

    fixture_name = "tables.docx"

    async def _call(self, **args):
        from fastmcp import Client

        from verified_docx_mcp import server

        async with Client(server.mcp) as client:
            return await client.call_tool(
                "replace_cell_markdown",
                {"path": str(self.target), "table_id": 1, "row_index": 1, "cell_index": 1,
                 "markdown": "via tool", **args},
                raise_on_error=False,
            )

    async def test_write_mode_live_reaches_the_live_route(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        result = await self._call(write_mode="live")
        self.assertFalse(result.is_error, result)
        self.assertEqual(doc.tables[0][0][0], "via tool")

    async def test_write_mode_file_is_honoured_and_refused_under_a_session(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        result = await self._call(write_mode="file")
        self.assertTrue(result.is_error)
        self.assertIn("LIVE_SESSION_ACTIVE", "".join(c.text for c in result.content))
        self.assertEqual(doc.tables[0][0][0], "R1C1")

    async def test_bad_write_mode_is_rejected(self) -> None:
        result = await self._call(write_mode="sideways")
        self.assertTrue(result.is_error)
        self.assertIn("INVALID_INPUT", "".join(c.text for c in result.content))


class ParagraphRunParserTests(unittest.TestCase):
    def test_paragraphs_and_inline_marks(self):
        paragraphs = markdown_to_ooxml.parse_paragraph_runs("a **b** *c*\n\nsecond")
        self.assertEqual(len(paragraphs), 2)
        self.assertEqual([(r.text, r.bold, r.italic) for r in paragraphs[0]],
                         [("a ", False, False), ("b", True, False), (" ", False, False), ("c", False, True)])
        self.assertEqual([r.text for r in paragraphs[1]], ["second"])

    def test_empty_markdown_is_no_paragraphs(self):
        self.assertEqual(markdown_to_ooxml.parse_paragraph_runs(""), [])

    def test_unsupported_block_is_invalid_input(self):
        with self.assertRaises(VerifyError) as cm:
            markdown_to_ooxml.parse_paragraph_runs("```\ncode\n```")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class CellPayloadTests(unittest.TestCase):
    def test_cell_get_round_trip(self):
        payload = protocol.CellGetPayload(table_index=2, row_index=3, cell_index=4)
        self.assertEqual(protocol.CellGetPayload.from_json(payload.to_json()), payload)

    def test_cell_set_round_trip(self):
        payload = protocol.CellSetPayload(
            table_index=1,
            row_index=1,
            cell_index=2,
            paragraphs=[[{"text": "x", "bold": True}]],
            expected_before_text="old",
            track_changes=True,
            expected_body_sha256="abc",
        )
        self.assertEqual(protocol.CellSetPayload.from_json(payload.to_json()), payload)

    def test_indexes_are_one_based(self):
        for bad in (0, -1):
            with self.assertRaises(protocol.ProtocolError):
                protocol.CellGetPayload.from_json({"table_index": bad, "row_index": 1, "cell_index": 1})

    def test_paragraphs_must_be_lists_of_run_objects(self):
        with self.assertRaises(protocol.ProtocolError):
            protocol.CellSetPayload.from_json(
                {
                    "table_index": 1,
                    "row_index": 1,
                    "cell_index": 1,
                    "paragraphs": ["not a list"],
                    "expected_before_text": "",
                }
            )

    def test_ops_are_registered(self):
        self.assertIn("cell_get", protocol.VALID_OPS)
        self.assertIn("cell_set", protocol.VALID_OPS)


if __name__ == "__main__":
    unittest.main()
