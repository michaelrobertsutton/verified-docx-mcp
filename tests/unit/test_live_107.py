"""Live-mode (fake pane, no Word) coverage for issue #107: cell font inheritance,
the cell paragraph merge, and structural paragraph anchors over the ops channel.

The fake proves the SERVER's behaviour: routing, capability gates, what is sent,
and the independent read-back. It cannot prove Word's own range, import or font
behaviour; that is the real-Word acceptance in docs/acceptance/107-cell-editing.md.
"""

from __future__ import annotations

import asyncio
import json
import shutil
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
FIXTURES = REPO / "tests" / "fixtures"
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, OpRefused, docx_to_flat_opc
from test_cell_editing_107 import add_second_paragraph
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, mutations, paragraph_insert, paragraph_merge, tables
from verified_docx_mcp.errors import ErrorCode, VerifyError

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

FONT = {"name": "Arial", "size": 8.5, "color": "#1F3864"}  # what Word reports at the cell's start


def _audit_tools() -> list[str]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    return [json.loads(line).get("tool") for line in path.read_text().splitlines() if line.strip()]


def _two_by_two() -> FakeDocument:
    doc = FakeDocument(text="body")
    doc.tables = [[["R1C1", "R1C2"], ["R2C1", "R2C2"]]]
    return doc


class CellFontLiveTests(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def replace(self, markdown="new", **kwargs):
        return await asyncio.to_thread(
            tables.execute_replace_cell_markdown, str(self.target), 1, 1, 1, markdown, **kwargs
        )

    async def test_the_first_run_font_is_read_from_the_cell_and_sent_with_the_write(self) -> None:
        doc = _two_by_two()
        doc.cell_font[(0, 0, 0)] = FONT
        pane = await self.connect_pane(doc)
        evidence = await self.replace("**Bold** text")
        self.assertTrue(evidence["applied"])
        self.assertEqual(doc.applied_cell_fonts[(0, 0, 0)], {"name": "Arial", "size": 8.5, "color": "1F3864"})
        self.assertEqual([p.get("include_font") for p in doc.cell_get_payloads][:1], [True])
        # a cell the edit does not touch is not given a font
        self.assertNotIn((0, 0, 1), doc.applied_cell_fonts)
        self.assertIsNotNone(pane)

    async def test_markdown_with_a_link_does_not_send_a_colour_that_would_beat_the_link_style(self) -> None:
        doc = _two_by_two()
        doc.cell_font[(0, 0, 0)] = FONT
        await self.connect_pane(doc)
        await self.replace("see [proof](https://x.test)")
        self.assertEqual(doc.applied_cell_fonts[(0, 0, 0)], {"name": "Arial", "size": 8.5})

    async def test_a_font_readback_failure_is_audited_because_the_text_may_have_landed(self) -> None:
        doc = _two_by_two()
        doc.cell_font[(0, 0, 0)] = FONT
        doc.font_readback_mismatch.add((0, 0, 0))
        await self.connect_pane(doc)
        records = len(_audit_tools())
        with self.assertRaises(VerifyError):
            await self.replace()
        self.assertIn("replace_cell_markdown:verification_failed", _audit_tools()[records:])

    async def test_a_cell_whose_font_comes_from_its_style_sends_nothing(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        self.assertTrue((await self.replace())["applied"])
        self.assertEqual(doc.applied_cell_fonts, {})

    async def test_inherit_format_none_neither_reads_nor_applies_the_font(self) -> None:
        doc = _two_by_two()
        doc.cell_font[(0, 0, 0)] = FONT
        await self.connect_pane(doc)
        self.assertTrue((await self.replace(inherit_format="none"))["applied"])
        self.assertEqual(doc.applied_cell_fonts, {})
        self.assertFalse(any(p.get("include_font") for p in doc.cell_get_payloads))

    async def test_a_font_that_fails_the_pane_readback_is_a_verification_failure(self) -> None:
        doc = _two_by_two()
        doc.cell_font[(0, 0, 0)] = FONT
        doc.font_readback_mismatch.add((0, 0, 0))
        await self.connect_pane(doc)
        with self.assertRaises(VerifyError) as cm:
            await self.replace()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)


class StructuralAnchorLiveTests(LiveWriteBridgeTestCase):
    async def insert(self, anchor, **kwargs):
        return await asyncio.to_thread(
            paragraph_insert.execute_insert_paragraphs,
            str(self.target), anchor, [{"text": "Caption"}], 1, write_mode="live", **kwargs,
        )

    async def test_after_table_sends_the_structural_anchor_and_no_hazard_list(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        await self.connect_pane(doc)
        evidence = await self.insert({"after_table_id": 1})
        self.assertTrue(evidence["applied"])
        sent = doc.positional_payloads[-1]
        self.assertEqual(sent["anchor"], {"after_table_id": 1})
        self.assertNotIn("forbidden_empty_tags", sent)

    async def test_fill_empty_sends_the_shared_empty_paragraph_policy(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        await self.connect_pane(doc)
        await self.insert({"empty_paragraph_after_table_id": 1})
        sent = doc.positional_payloads[-1]
        self.assertEqual(sent["anchor"], {"empty_paragraph_after_table_id": 1})
        self.assertEqual(sent["forbidden_empty_tags"], list(paragraph_insert.EMPTY_PARAGRAPH_FORBIDDEN_TAGS))

    async def test_a_pane_without_positions_refuses_before_sending(self) -> None:
        doc = FakeDocument(text="Intro\nProof\nTail")
        await self.connect_pane(doc, capabilities=["paragraph_insert", "shape_guard", "body_hash_v1"])
        with self.assertRaises(VerifyError) as cm:
            await self.insert({"after_table_id": 1})
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertEqual(doc.positional_payloads, [])
        # a text anchor still works on that older pane
        self.assertTrue((await self.insert("Proof"))["applied"])


class MergeLiveTests(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def asyncSetUp(self) -> None:
        await super().asyncSetUp()
        add_second_paragraph(self.target, text="second", bold=True, same_ppr=True)
        self.doc = _two_by_two()

    async def connect(self, merged_body: bool = True, **kwargs):
        pane = await self.connect_pane(
            self.doc, body_ooxml=docx_to_flat_opc(self.target), **kwargs
        )
        if merged_body:
            merged = self.target.with_name("merged.docx")
            shutil.copyfile(self.target, merged)
            paragraph_merge.execute(str(merged), 1, 1, 1, 1, separator=" ", write_mode="file")

            def land(p, payload):
                p.body_ooxml = docx_to_flat_opc(merged)

            pane.on_cell_paragraph_merge = land
        return pane

    async def merge(self, **kwargs):
        kwargs.setdefault("separator", " ")
        return await asyncio.to_thread(
            paragraph_merge.execute, str(self.target), 1, 1, 1, 1, write_mode="live", **kwargs
        )

    async def test_merge_sends_the_merged_paragraph_and_verifies_the_cell(self) -> None:
        pane = await self.connect()
        evidence = await self.merge(expected_text=["R1C1", "second"])
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["after"], "R1C1 second")
        sent = pane.merge_payloads[0]
        self.assertEqual((sent["table_index"], sent["row_index"], sent["cell_index"], sent["paragraph_index"]), (1, 1, 1, 1))
        self.assertEqual(sent["expected_text"], ["R1C1", "second"])  # text, not XML: see the pane comment
        self.assertNotIn("before_xml", sent)
        # no OOXML is imported: Word for Mac throws GeneralException on a cell insertOoxml
        self.assertNotIn("ooxml", sent)
        self.assertEqual(sent["separator"], " ")
        self.assertFalse(sent["track_changes"])
        self.assertNotIn("warnings", evidence)

    async def test_a_write_the_pane_reports_but_that_did_not_land_fails_verification_and_is_audited(self) -> None:
        pane = await self.connect(merged_body=False)  # the pane replies applied=true; the body is unchanged
        records = len(_audit_tools())
        with self.assertRaises(VerifyError) as cm:
            await self.merge()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)
        self.assertIn("paragraphs differ", cm.exception.envelope.message)
        self.assertEqual(len(pane.merge_payloads), 1)
        self.assertIn("merge_paragraph_with_next:verification_failed", _audit_tools()[records:])

    async def test_a_pane_reported_verification_failure_after_the_write_is_audited(self) -> None:
        pane = await self.connect(merged_body=False)

        def refuse(p, payload):
            raise OpRefused("VERIFICATION_FAILED", "Paragraph merge count differs")

        pane.on_cell_paragraph_merge = refuse
        records = len(_audit_tools())
        with self.assertRaises(VerifyError) as cm:
            await self.merge()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)
        self.assertIn("merge_paragraph_with_next:verification_failed", _audit_tools()[records:])

    async def test_a_pane_that_does_not_report_applied_fails_verification(self) -> None:
        pane = await self.connect()
        pane.on_cell_paragraph_merge = lambda p, payload: {"applied": False}
        with self.assertRaises(VerifyError) as cm:
            await self.merge()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)

    async def test_a_document_that_moved_after_the_write_is_a_warning_not_silence(self) -> None:
        pane = await self.connect()
        merged_land = pane.on_cell_paragraph_merge

        def land_then_move(p, payload):
            merged_land(p, payload)
            return {"post": "f" * 64}  # the pane's reply names a hash the next read will not match

        pane.on_cell_paragraph_merge = land_then_move
        evidence = await self.merge()
        self.assertTrue(evidence["applied"])
        self.assertIn("changed after the merge", evidence["warnings"][0])

    async def test_a_pair_with_different_paragraph_properties_is_refused_before_anything_is_sent(self) -> None:
        # Real Word keeps the SECOND paragraph's properties when it joins two (16.114.3: the merged
        # paragraph came back centred), so a live merge of a differing pair cannot keep the first's.
        shutil.copyfile(FIXTURES / "tables.docx", self.target)
        add_second_paragraph(self.target, text="second", bold=True)  # second paragraph: a different style
        pane = await self.connect()
        with self.assertRaises(VerifyError) as cm:
            await self.merge()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.STRUCTURAL_BOUNDARY)
        self.assertIn("different paragraph properties", cm.exception.envelope.message)
        self.assertEqual(pane.merge_payloads, [])

    async def test_tracked_merge_is_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect()
        with self.assertRaises(VerifyError) as cm:
            await self.merge(track_changes=True)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(pane.merge_payloads, [])

    async def test_a_stale_revision_before_is_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect()
        with self.assertRaises(VerifyError) as cm:
            await self.merge(revision_before="live:sha256:" + "0" * 64)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_STALE)
        self.assertEqual(pane.merge_payloads, [])

    async def test_a_pane_without_the_capability_is_refused_before_anything_is_sent(self) -> None:
        pane = await self.connect(capabilities=["cell_edit", "shape_guard", "body_ooxml", "body_hash_v1"])
        with self.assertRaises(VerifyError) as cm:
            await self.merge()
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertEqual(pane.merge_payloads, [])

    async def test_a_text_mismatch_sends_nothing(self) -> None:
        pane = await self.connect()
        with self.assertRaises(VerifyError) as cm:
            await self.merge(expected_text=["R1C1", "WRONG"])
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.CELL_TEXT_MISMATCH)
        self.assertEqual(pane.merge_payloads, [])


if __name__ == "__main__":
    unittest.main()
