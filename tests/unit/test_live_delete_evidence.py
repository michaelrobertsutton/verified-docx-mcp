"""Issue #112: live delete_paragraph returns counts and the two neighbouring paragraphs,
never every paragraph of the section twice, for a current pane (compact_evidence_v1)
and for an older pane that still sends the full lists (the server trims them and keeps
them in the audit log)."""

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

from verified_docx_mcp import audit, mutations
from verified_docx_mcp.live import paragraphs_live

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

CAPS = ["delete_paragraph", "shape_guard"]


def _records(tool: str) -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    rows = [json.loads(line) for line in path.read_text().splitlines() if line.strip()]
    return [r for r in rows if r["tool"] == tool]


def _big_doc(legacy: bool) -> FakeDocument:
    doc = FakeDocument(text="\n".join(f"Paragraph {n} " + "word " * 40 for n in range(300)) + "\nTarget paragraph\nTail")
    doc.legacy_delete_lists = legacy
    return doc


class LiveDeleteEvidenceTests(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def delete(self, doc: FakeDocument):
        await self.connect_pane(doc, capabilities=CAPS)
        return await asyncio.to_thread(
            paragraphs_live.delete_paragraph, str(self.target), "Target paragraph", f"live:sha256:{doc.sha256()}"
        )

    async def test_current_pane_is_compact(self) -> None:
        doc = _big_doc(legacy=False)
        result = await self.delete(doc)
        self.assertNotIn("before_paragraphs", result)
        self.assertNotIn("after_paragraphs", result)
        self.assertEqual(result["paragraphs_before"], 302)
        self.assertEqual(result["paragraphs_after"], 301)
        self.assertEqual(result["before_excerpt"], "Target paragraph")
        self.assertTrue(result["neighbours"]["previous"].startswith("Paragraph 299 "))
        self.assertEqual(result["neighbours"]["next"], "Tail")
        self.assertLess(len(json.dumps(result)) // 4, 500)

    async def test_older_pane_lists_are_trimmed_but_kept_in_the_audit_log(self) -> None:
        doc = _big_doc(legacy=True)
        result = await self.delete(doc)
        self.assertNotIn("before_paragraphs", result)
        self.assertNotIn("after_paragraphs", result)
        self.assertEqual((result["paragraphs_before"], result["paragraphs_after"]), (302, 301))
        self.assertEqual(result["neighbours"]["next"], "Tail")
        self.assertLess(len(json.dumps(result)) // 4, 500)
        record = _records("delete_paragraph")[-1]["evidence"]
        self.assertEqual(len(record["before_paragraphs"]), 302)
        self.assertEqual(len(record["after_paragraphs"]), 301)

    async def test_older_pane_lists_are_redacted_when_the_operator_disables_excerpts(self) -> None:
        """Review finding: the lists routed into the audit record bypassed VERIFIED_DOCX_MCP_AUDIT_EXCERPTS=0."""
        import os
        from unittest import mock

        doc = FakeDocument(text="Secret one\nSecret target\nSecret three\nTail")
        doc.legacy_delete_lists = True
        await self.connect_pane(doc, capabilities=CAPS)
        with mock.patch.dict(os.environ, {audit._AUDIT_EXCERPTS_ENV: "0"}):
            await asyncio.to_thread(
                paragraphs_live.delete_paragraph, str(self.target), "Secret target", f"live:sha256:{doc.sha256()}"
            )
        raw = (audit._state_dir() / "audit.jsonl").read_text(encoding="utf-8")
        self.assertNotIn("Secret one", raw)
        self.assertNotIn("Secret three", raw)
        record = _records("delete_paragraph")[-1]["evidence"]
        self.assertTrue(record["before_paragraphs"].startswith("[redacted;"))
        self.assertTrue(record["after_paragraphs"].startswith("[redacted;"))

    async def test_long_neighbours_are_clipped(self) -> None:
        doc = FakeDocument(text="x" * 900 + "\nTarget paragraph\n" + "y" * 900 + "\nTail")
        result = await self.delete(doc)
        self.assertEqual(len(result["neighbours"]["previous"]), 201)
        self.assertTrue(result["neighbours"]["previous"].endswith("…"))


if __name__ == "__main__":
    unittest.main()
