"""Live insert_paragraphs (issue #77) over the WSS ops channel, against
tests/unit/fake_pane.py (no Word). The pane's Office.js calls are covered by
tests/unit/js/paragraph_insert_harness.mjs; these pin the server's side: validation
before anything is sent, the capability/stale gates, the independent read-back, the
failure audit record, and routing/error mapping."""

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

from verified_docx_mcp import audit, mutations, paragraph_insert
from verified_docx_mcp.errors import ErrorCode, VerifyError

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

SPECS = [
    {"text": "Proposed heading", "style": "Heading 2"},
    {"text": "Purple instruction", "color": "7030a0"},
    {"text": "Green evaluation", "color": "00B050"},
]


def _doc() -> FakeDocument:
    return FakeDocument(text="Intro\nProof\nTail")


def _audit_records() -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text().splitlines() if line.strip()]


class InsertParagraphsLiveTests(LiveWriteBridgeTestCase):
    async def insert(self, anchor="Proof", paragraphs=SPECS, expected_matches=1, **kwargs):
        return await asyncio.to_thread(
            paragraph_insert.execute_insert_paragraphs, str(self.target), anchor, paragraphs, expected_matches, **kwargs)

    async def refused(self, code: ErrorCode, *args, **kwargs):
        with self.assertRaises(VerifyError) as cm:
            await self.insert(*args, **kwargs)
        self.assertEqual(cm.exception.envelope.error_code, code, cm.exception.envelope.message)
        return cm.exception.envelope

    async def test_tracked_insert_after_the_anchor(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        evidence = await self.insert(track_changes=True)

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["position"], "after")
        self.assertTrue(evidence["track_changes"])
        self.assertEqual(doc.text, "Intro\nProof\nProposed heading\nPurple instruction\nGreen evaluation\nTail")
        self.assertEqual([i["color"] for i in evidence["inserted"]], ["", "#7030A0", "#00B050"])
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertTrue(evidence["audit_logged"])
        self.assertEqual(self.target.read_bytes(), before_bytes)
        sent = pane.paragraph_insert_payloads[0]
        self.assertTrue(sent["track_changes"])
        self.assertEqual(sent["paragraphs"][1], {"text": "Purple instruction", "color": "7030A0"})
        self.assertEqual(sent["paragraphs"][0], {"text": "Proposed heading", "style": "Heading 2"})

    async def test_insert_before_the_anchor(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.insert(paragraphs=SPECS[:1], position="before")
        self.assertEqual(doc.text, "Intro\nProposed heading\nProof\nTail")

    async def test_anchor_count_mismatch_and_zero_match(self) -> None:
        doc = FakeDocument(text="x one\nx two")
        await self.connect_pane(doc)
        await self.refused(ErrorCode.MATCH_COUNT_MISMATCH, anchor="x")
        await self.refused(ErrorCode.ZERO_MATCH, anchor="absent")
        self.assertEqual(doc.text, "x one\nx two")

    async def test_invalid_requests_are_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect_pane(_doc())
        bad = [
            {"paragraphs": []},
            {"paragraphs": [{"text": "t", "color": "purple"}]},
            {"paragraphs": [{"text": "line one\nline two"}]},
            {"paragraphs": [{"text": "t", "extra": 1}]},
            {"paragraphs": [{"text": "t", "style": " "}]},
            {"expected_matches": 2},
            {"position": "middle"},
            {"anchor": "x" * 256},
            {"anchor": ""},
        ]
        for kwargs in bad:
            await self.refused(ErrorCode.INVALID_INPUT, **kwargs)
        self.assertEqual(pane.paragraph_insert_payloads, [])

    async def test_pane_without_the_capability_is_refused(self) -> None:
        pane = await self.connect_pane(_doc(), capabilities=["shape_guard"])
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertEqual(pane.paragraph_insert_payloads, [])

    async def test_stale_revision_before_is_refused(self) -> None:
        pane = await self.connect_pane(_doc())
        await self.refused(ErrorCode.LIVE_STALE, revision_before="live:sha256:" + "0" * 64)
        self.assertEqual(pane.paragraph_insert_payloads, [])

    async def test_unknown_style_inserts_nothing(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.STYLE_NOT_FOUND, paragraphs=[{"text": "t", "style": "No Such Style"}])
        self.assertEqual(doc.text, "Intro\nProof\nTail")

    async def test_dropped_color_is_caught_and_audited(self) -> None:
        doc = _doc()
        doc.drop_paragraph_color = True
        await self.connect_pane(doc)
        envelope = await self.refused(ErrorCode.VERIFICATION_FAILED)
        self.assertIn("color", envelope.message)
        failures = [r for r in _audit_records() if r.get("tool") == "insert_paragraphs:verification_failed"]
        self.assertTrue(failures)

    async def test_explicit_file_mode_under_a_session_still_refuses(self) -> None:
        await self.connect_pane(_doc())
        await self.refused(ErrorCode.LIVE_SESSION_ACTIVE, write_mode="file")

    async def test_explicit_live_without_a_session_is_unavailable(self) -> None:
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, write_mode="live")


if __name__ == "__main__":
    unittest.main()
