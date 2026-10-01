"""Unit tests for issue #35: live ``replace_text`` / ``format_text`` reaching
Word text boxes via ``scope``, and the live ``list_textboxes`` read.

Same harness as test_live_cells.py: a real bridge on EPHEMERAL ports with a
connected ``fake_pane.FakePane``. The fake's shape model proves the SERVER's
behavior (scope validation, capability gate, compare-and-set, exact-text and
independent-re-read verification, evidence/audit). It cannot prove what a real
Word returns for ``body.shapes`` -- that is the manual sideload check in
docs/live-mode.md; ``test_taskpane_js.py`` covers the pane's own control flow.
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

from fake_pane import FakeDocument, FakeShape
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import protocol, textboxes_live


def _doc() -> FakeDocument:
    doc = FakeDocument(text="Body copy mentions proven results.")
    doc.shapes = [
        FakeShape("101", "Why Team Skyward: proven delivery\rSecond paragraph"),
        FakeShape("102", "Callout about cost savings"),
        FakeShape("103", "A cat sat here", type="geometricshape"),
    ]
    return doc


def _audit_records() -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    return [
        json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()
    ]


class TextboxLiveTestCase(LiveWriteBridgeTestCase):
    async def replace(self, find: str, replace: str, n: int = 1, **kwargs):
        kwargs.setdefault("write_mode", "live")
        kwargs.setdefault("scope", "textboxes")
        return await asyncio.to_thread(
            text_edit.execute_replace_text, str(self.target), find, replace, n, **kwargs
        )

    async def fmt(self, find: str, style: dict, n: int = 1, **kwargs):
        kwargs.setdefault("write_mode", "live")
        kwargs.setdefault("scope", "textboxes")
        return await asyncio.to_thread(
            text_edit.execute_format_text, str(self.target), find, style, n, **kwargs
        )

    async def assertRefused(self, code: ErrorCode, coro):
        with self.assertRaises(VerifyError) as cm:
            await coro
        self.assertEqual(cm.exception.envelope.error_code, code)
        return cm.exception.envelope


class ReplaceInTextBoxTests(TextboxLiveTestCase):
    async def test_replace_in_one_box_is_verified_and_leaves_others_alone(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        body_before = doc.text

        evidence = await self.replace("proven delivery", "verified delivery")

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["rung"], 3)
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(evidence["before"], "proven delivery")
        self.assertEqual(evidence["after"], "verified delivery")
        self.assertEqual(evidence["scope"], "textboxes")
        self.assertEqual(evidence["verification"], "exact")
        self.assertTrue(evidence["second_read"])
        self.assertEqual(
            doc.shapes[0].text, "Why Team Skyward: verified delivery\rSecond paragraph"
        )
        self.assertEqual(doc.shapes[1].text, "Callout about cost savings")
        self.assertEqual(doc.text, body_before)  # the body was never touched
        hashes = evidence["textbox_text_sha256"]
        self.assertNotEqual(hashes["before"]["101"], hashes["after"]["101"])
        self.assertEqual(hashes["before"]["102"], hashes["after"]["102"])

    async def test_shape_scope_targets_one_box_by_id(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.replace("Callout", "Box", scope="shape:102")
        self.assertEqual(evidence["scope"], "shape:102")
        self.assertEqual(doc.shapes[1].text, "Box about cost savings")

    async def test_geometric_shape_text_is_reachable(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.replace("sat", "stood")
        self.assertEqual(doc.shapes[2].text, "A cat stood here")

    async def test_default_body_scope_still_cannot_see_box_text(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.assertRefused(
            ErrorCode.ZERO_MATCH,
            asyncio.to_thread(
                text_edit.execute_replace_text,
                str(self.target),
                "Why Team Skyward",
                "x",
                1,
                write_mode="live",
            ),
        )
        self.assertTrue(doc.shapes[0].text.startswith("Why Team Skyward"))

    async def test_all_scope_edits_body_and_boxes_with_one_total_gate(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.replace("proven", "verified", 2, scope="all")
        self.assertEqual(evidence["match_count"], 2)
        self.assertIn("verified results", doc.text)
        self.assertIn("verified delivery", doc.shapes[0].text)
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])

    async def test_replacement_containing_find_is_verified_exactly(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.replace("cat", "cats")
        self.assertEqual(doc.shapes[2].text, "A cats sat here")

    async def test_deletion_is_verified_exactly(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.replace(" about cost savings", "")
        self.assertEqual(doc.shapes[1].text, "Callout")

    async def test_track_changes_uses_relaxed_verification_and_says_so(self) -> None:
        doc = _doc()
        doc.track_changes_keeps_deleted = True  # Word leaves tracked-deleted text in the box
        await self.connect_pane(doc)
        evidence = await self.replace("proven delivery", "verified delivery", track_changes=True)
        self.assertEqual(evidence["verification"], "relaxed_track_changes")
        self.assertIn("verified delivery", doc.shapes[0].text)

    async def test_audit_entry_carries_the_scope_fields(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.replace("proven delivery", "verified delivery")
        record = _audit_records()[-1]
        self.assertEqual(record["tool"], "replace_text")
        self.assertEqual(record["evidence"]["scope"], "textboxes")
        self.assertEqual(record["evidence"]["textbox_ids"], ["101", "102", "103"])
        self.assertIn("textbox_text_sha256", record["evidence"])


class RefusalTests(TextboxLiveTestCase):
    async def test_missing_capability_refuses_before_any_op(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(
            doc, capabilities=["row_scope", "cell_edit", "comments_by_id"]
        )
        await self.assertRefused(ErrorCode.LIVE_CAPABILITY_MISSING, self.replace("proven", "x"))
        self.assertEqual(pane.textbox_list_requests, 0)
        self.assertTrue(doc.shapes[0].text.startswith("Why Team Skyward"))

    async def test_invalid_scope_and_combinations(self) -> None:
        await self.connect_pane(_doc())
        await self.assertRefused(ErrorCode.INVALID_INPUT, self.replace("a", "b", scope="nope"))
        await self.assertRefused(ErrorCode.INVALID_INPUT, self.replace("a", "b", scope="shape:"))
        await self.assertRefused(
            ErrorCode.INVALID_INPUT, self.replace("a", "b", within_row_containing="anchor")
        )

    async def test_find_equal_to_replace_is_refused_up_front(self) -> None:
        pane = await self.connect_pane(_doc())
        await self.assertRefused(ErrorCode.INVALID_INPUT, self.replace("cat", "cat"))
        self.assertEqual(pane.textbox_list_requests, 0)

    async def test_unknown_shape_id_lists_what_exists(self) -> None:
        await self.connect_pane(_doc())
        envelope = await self.assertRefused(
            ErrorCode.INVALID_INPUT, self.replace("a", "b", scope="shape:999")
        )
        self.assertEqual(envelope.diagnostics["available_shape_ids"], ["101", "102", "103"])

    async def test_zero_and_wrong_counts_map_to_the_usual_codes(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.assertRefused(ErrorCode.ZERO_MATCH, self.replace("not in any box", "x"))
        await self.assertRefused(
            ErrorCode.MATCH_COUNT_MISMATCH, self.replace("proven delivery", "x", 2)
        )
        self.assertTrue(doc.shapes[0].text.startswith("Why Team Skyward"))

    async def test_box_edited_between_list_and_write_is_stale_and_nothing_is_written(self) -> None:
        doc = _doc()

        def co_author_types() -> None:
            doc.shapes[1].text += " (edited by a co-author)"

        doc.before_shape_write = co_author_types
        await self.connect_pane(doc)
        await self.assertRefused(
            ErrorCode.LIVE_STALE, self.replace("proven delivery", "verified delivery")
        )
        self.assertIn("proven delivery", doc.shapes[0].text)  # the target was never written
        self.assertTrue(doc.shapes[1].text.endswith("(edited by a co-author)"))

    async def test_stale_body_revision_is_refused_before_listing(self) -> None:
        pane = await self.connect_pane(_doc())
        await self.assertRefused(
            ErrorCode.LIVE_STALE,
            self.replace("proven delivery", "x", revision_before="live:sha256:" + "0" * 64),
        )
        self.assertEqual(pane.textbox_list_requests, 0)

    async def test_exhaustive_scope_refuses_when_coverage_is_incomplete_but_shape_scope_works(
        self,
    ) -> None:
        doc = _doc()
        doc.incomplete_shapes = [
            {"shape_id": "7", "type": "textbox", "reason": "duplicate shape id"}
        ]
        await self.connect_pane(doc)
        envelope = await self.assertRefused(
            ErrorCode.LIVE_OP_FAILED, self.replace("proven delivery", "x")
        )
        self.assertEqual(envelope.diagnostics["incomplete"][0]["shape_id"], "7")
        await self.replace("proven delivery", "verified delivery", scope="shape:101")
        self.assertIn("verified delivery", doc.shapes[0].text)

    async def test_no_text_boxes_at_all(self) -> None:
        doc = FakeDocument(text="body only")
        await self.connect_pane(doc)
        await self.assertRefused(ErrorCode.ZERO_MATCH, self.replace("x", "y"))

    async def test_malformed_listing_is_refused(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(doc)
        pane.malformed_textbox_list = True
        await self.assertRefused(ErrorCode.LIVE_OP_FAILED, self.replace("proven delivery", "x"))

    async def test_file_mode_refuses_every_non_body_scope(self) -> None:
        # No pane connected: write_mode auto resolves to file, and so does an explicit "file".
        for mode in ("auto", "file"):
            with self.subTest(write_mode=mode):
                await self.assertRefused(
                    ErrorCode.INVALID_INPUT,
                    self.replace("x", "y", write_mode=mode, scope="textboxes"),
                )
                await self.assertRefused(
                    ErrorCode.INVALID_INPUT,
                    self.fmt("x", {"bold": True}, write_mode=mode, scope="all"),
                )


class VerificationTests(TextboxLiveTestCase):
    async def test_a_dropped_write_fails_verification_and_is_audited_as_uncertain(self) -> None:
        doc = _doc()
        doc.dropped_shape_writes = {"101"}
        await self.connect_pane(doc)
        envelope = await self.assertRefused(
            ErrorCode.VERIFICATION_FAILED, self.replace("proven delivery", "verified delivery")
        )
        self.assertIn("may have been applied", envelope.message)
        self.assertIn("proven delivery", doc.shapes[0].text)
        record = _audit_records()[-1]
        self.assertEqual(record["evidence"]["applied"], "unknown")
        self.assertEqual(record["evidence"]["scope"], "textboxes")

    async def test_all_scope_partial_success_fails(self) -> None:
        doc = _doc()
        doc.dropped_shape_writes = {"101"}  # the body edit lands, the box edit is dropped
        await self.connect_pane(doc)
        await self.assertRefused(
            ErrorCode.VERIFICATION_FAILED, self.replace("proven", "verified", 2, scope="all")
        )
        self.assertIn("verified results", doc.text)

    async def test_dropped_write_to_one_of_several_boxes_is_caught(self) -> None:
        doc = _doc()
        doc.shapes[1].text = "cat in the second box"
        doc.dropped_shape_writes = {"102"}
        await self.connect_pane(doc)
        await self.assertRefused(ErrorCode.VERIFICATION_FAILED, self.replace("cat", "dog", 2))


class FormatInTextBoxTests(TextboxLiveTestCase):
    async def test_every_requested_property_is_read_back_and_verified(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.fmt(
            "proven delivery",
            {"bold": True, "italic": True, "underline": True, "strike": True, "color": "C00000"},
        )
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], 2)
        self.assertEqual(evidence["scope"], "textboxes")
        self.assertEqual(
            doc.shapes[0].formatting,
            {"bold": True, "italic": True, "underline": True, "strike": True, "color": "C00000"},
        )
        self.assertEqual(doc.shapes[0].text, "Why Team Skyward: proven delivery\rSecond paragraph")

    async def test_a_dropped_format_write_is_caught_by_the_read_back(self) -> None:
        doc = _doc()
        doc.dropped_shape_writes = {"101"}
        await self.connect_pane(doc)
        await self.assertRefused(
            ErrorCode.VERIFICATION_FAILED, self.fmt("proven delivery", {"bold": True})
        )

    async def test_a_dropped_color_write_is_caught(self) -> None:
        doc = _doc()
        doc.dropped_shape_writes = {"101"}
        await self.connect_pane(doc)
        await self.assertRefused(
            ErrorCode.VERIFICATION_FAILED, self.fmt("proven delivery", {"color": "C00000"})
        )

    async def test_stale_target_refuses_a_format_too(self) -> None:
        doc = _doc()
        doc.before_shape_write = lambda: setattr(doc.shapes[0], "text", "changed under us")
        await self.connect_pane(doc)
        await self.assertRefused(ErrorCode.LIVE_STALE, self.fmt("proven delivery", {"bold": True}))


class BodyFormatReadBackTests(LiveWriteBridgeTestCase):
    """The body path now verifies bold/italic/underline read-back too."""

    async def test_body_format_still_passes_with_the_fake_pane(self) -> None:
        doc = FakeDocument(text="alpha beta gamma")
        await self.connect_pane(doc)
        evidence = await asyncio.to_thread(
            text_edit.execute_format_text,
            str(self.target),
            "beta",
            {"bold": True, "underline": True},
            1,
            write_mode="live",
        )
        self.assertTrue(evidence["applied"])


class ListTextboxesTests(TextboxLiveTestCase):
    async def test_lists_every_box_with_ids_hashes_and_notes(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        result = await asyncio.to_thread(textboxes_live.list_textboxes_live, str(self.target))
        self.assertEqual(result["source"], "live")
        self.assertEqual([t["shape_id"] for t in result["textboxes"]], ["101", "102", "103"])
        self.assertEqual(result["textboxes"][0]["paragraph_count"], 2)
        self.assertEqual(result["textboxes"][2]["type"], "geometricshape")
        self.assertTrue(result["revision"].startswith("live:sha256:"))
        self.assertTrue(any("NOT a file-mode" in n for n in result["notes"]))
        self.assertEqual(result["incomplete"], [])

    async def test_surfaces_incomplete_and_skipped(self) -> None:
        doc = _doc()
        doc.incomplete_shapes = [
            {"shape_id": "9", "type": "textbox", "reason": "body not readable"}
        ]
        doc.skipped_shapes = [{"shape_id": "6", "type": "picture"}]
        await self.connect_pane(doc)
        result = await asyncio.to_thread(textboxes_live.list_textboxes_live, str(self.target))
        self.assertEqual(result["incomplete"][0]["shape_id"], "9")
        self.assertEqual(result["skipped"][0]["shape_id"], "6")

    async def test_needs_the_capability(self) -> None:
        await self.connect_pane(_doc(), capabilities=[])
        with self.assertRaises(VerifyError) as cm:
            await asyncio.to_thread(textboxes_live.list_textboxes_live, str(self.target))
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_no_session_is_live_unavailable(self) -> None:
        with self.assertRaises(VerifyError) as cm:
            await asyncio.to_thread(textboxes_live.list_textboxes_live, str(self.target))
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)


class ProtocolTests(unittest.TestCase):
    def test_textbox_list_is_a_valid_op(self) -> None:
        self.assertIn("textbox_list", protocol.VALID_OPS)

    def test_body_scope_payloads_keep_their_wire_shape(self) -> None:
        wire = protocol.ReplacePayload(find="a", expected_matches=1, replace="b").to_json()
        self.assertNotIn("scope", wire)
        self.assertNotIn("expect", wire)

    def test_scoped_payloads_round_trip(self) -> None:
        for cls, extra in (
            (protocol.ReplacePayload, {"replace": "b"}),
            (protocol.FormatPayload, {"bold": True}),
        ):
            payload = cls(
                find="a", expected_matches=1, scope="shape:12", expect={"12": "ab" * 32}, **extra
            )
            again = cls.from_json(payload.to_json())
            self.assertEqual(again, payload)

    def test_invalid_scope_and_expect_are_rejected(self) -> None:
        base = {"find": "a", "expected_matches": 1, "replace": "b"}
        with self.assertRaises(protocol.ProtocolError):
            protocol.ReplacePayload.from_json({**base, "scope": "everything"})
        with self.assertRaises(protocol.ProtocolError):
            protocol.ReplacePayload.from_json({**base, "scope": "shape:"})
        with self.assertRaises(protocol.ProtocolError):
            protocol.ReplacePayload.from_json({**base, "scope": "textboxes", "expect": {"1": 5}})

    def test_textbox_list_result_must_be_an_object_with_a_list(self) -> None:
        good = {
            "textboxes": [
                {
                    "shape_id": "1",
                    "type": "textbox",
                    "group_path": [],
                    "text": "t",
                    "text_sha256": "x",
                    "paragraph_count": 1,
                }
            ]
        }
        parsed = protocol.TextboxListResult.from_json(good)
        self.assertEqual(parsed.textboxes[0].shape_id, "1")
        self.assertEqual(parsed.incomplete, [])
        with self.assertRaises(protocol.ProtocolError):
            protocol.TextboxListResult.from_json(
                [good["textboxes"][0]]
            )  # a bare array is not a reply
        with self.assertRaises(protocol.ProtocolError):
            protocol.TextboxListResult.from_json({"textboxes": ["nope"]})


if __name__ == "__main__":
    unittest.main()
