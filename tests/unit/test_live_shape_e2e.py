"""Issue #44 end to end: the anchored-shape guard seen through the live bridge.

`list_shapes`, the `shapes_before`/`shapes_after` counts in tool evidence and the audit log, and the
`ANCHORED_SHAPES` refusals, all against `tests/unit/fake_pane.py` (whose shape model mirrors the
pane's `inventoryShapes` / `dispatchOp`). The JS side is covered by tests/unit/js/shapes_harness.mjs;
real Word is checked by docs/acceptance/44-shape-guards.md. The old-pane case (no `shape_guard`
capability -> LIVE_CAPABILITY_MISSING) lives in the per-feature live test files.
"""

from __future__ import annotations

import asyncio
import json
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, mutations, server, tables, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import paragraphs_live

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

SHAPE = {
    "id": "7",
    "name": "Sidebar",
    "kind": "drawingml",
    "anchor_paragraph": {"index": 2, "text": "Proof"},
    "size_emu": {"width": "100", "height": "200"},
    "horizontal_position": {"relative_to": "page", "value": "25"},
    "vertical_position": {"relative_to": "paragraph", "value": "0"},
    "legacy_style": None,
}


def _doc() -> FakeDocument:
    doc = FakeDocument(text="Intro\nProof\nTail")
    doc.shapes = [dict(SHAPE)]
    return doc


def _audit_records() -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text().splitlines() if line.strip()]


class ShapeGuardE2ETests(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def replace(self, find: str, new: str = "X"):
        return await asyncio.to_thread(
            text_edit.execute_replace_text, str(self.target), find, new, 1, write_mode="live"
        )

    async def test_list_shapes_returns_the_pane_inventory(self) -> None:
        await self.connect_pane(_doc())
        result = await asyncio.to_thread(server.list_shapes, str(self.target))
        self.assertEqual(result["shapes"], [SHAPE])
        self.assertEqual(result["coverage"], "document_body")

    async def test_list_shapes_on_a_shapeless_document_is_empty_not_an_error(self) -> None:
        await self.connect_pane(FakeDocument(text="Intro"))
        result = await asyncio.to_thread(server.list_shapes, str(self.target))
        self.assertEqual(result["shapes"], [])

    async def test_permitted_replace_reports_equal_shape_counts_in_evidence_and_audit(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        before = len(_audit_records())

        evidence = await self.replace("Tail", "End")

        self.assertTrue(evidence["applied"])
        self.assertEqual(doc.text, "Intro\nProof\nEnd")
        self.assertEqual(evidence["shapes_before"], 1)
        self.assertEqual(evidence["shapes_after"], 1)
        record = _audit_records()[before:][-1]
        self.assertEqual(record["tool"], "replace_text")
        self.assertEqual(record["evidence"]["shapes_before"], 1)
        self.assertEqual(record["evidence"]["shapes_after"], 1)

    async def test_shapeless_document_reports_zero_counts(self) -> None:
        await self.connect_pane(FakeDocument(text="Intro\nTail"))
        evidence = await self.replace("Tail", "End")
        self.assertEqual((evidence["shapes_before"], evidence["shapes_after"]), (0, 0))

    async def test_replace_on_an_anchored_paragraph_is_refused_and_writes_nothing(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        before = len(_audit_records())

        with self.assertRaises(VerifyError) as cm:
            await self.replace("Proof", "Changed")

        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ANCHORED_SHAPES)
        self.assertEqual(doc.text, "Intro\nProof\nTail")
        self.assertEqual(len(doc.shapes), 1)
        self.assertEqual([r for r in _audit_records()[before:] if r["tool"] == "replace_text"], [])

    async def test_cell_edit_on_an_anchored_cell_is_refused(self) -> None:
        doc = FakeDocument(text="body")
        doc.tables = [[["Proof", "R1C2"], ["R2C1", "R2C2"]]]
        doc.shapes = [dict(SHAPE)]
        await self.connect_pane(doc)

        with self.assertRaises(VerifyError) as cm:
            await asyncio.to_thread(
                tables.execute_replace_cell_markdown, str(self.target), 1, 1, 1, "new", write_mode="live"
            )

        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ANCHORED_SHAPES)
        self.assertEqual(doc.tables[0][0][0], "Proof")

    async def test_delete_paragraph_refuses_an_anchored_paragraph_but_deletes_a_plain_one(self) -> None:
        doc = _doc()
        await self.connect_pane(doc, capabilities=["shape_guard", "delete_paragraph"])
        revision = f"live:sha256:{doc.sha256()}"

        with self.assertRaises(VerifyError) as cm:
            await asyncio.to_thread(paragraphs_live.delete_paragraph, str(self.target), "Proof", revision)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ANCHORED_SHAPES)
        self.assertEqual(doc.text, "Intro\nProof\nTail")

        evidence = await asyncio.to_thread(paragraphs_live.delete_paragraph, str(self.target), "Tail", revision)
        self.assertEqual(doc.text, "Intro\nProof")
        self.assertEqual((evidence["shapes_before"], evidence["shapes_after"]), (1, 1))


if __name__ == "__main__":
    unittest.main()
