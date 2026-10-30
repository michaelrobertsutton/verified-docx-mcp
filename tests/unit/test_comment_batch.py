"""Issue #113 part B: add_anchored_comment_batch (file mode, fake-pane live mode, registration)."""

from __future__ import annotations

import asyncio
import os
import shutil
import sys
import tempfile
import time
import unittest
import unittest.mock
from pathlib import Path
from types import SimpleNamespace

import pytest

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, docx_to_flat_opc
from fastmcp import Client
from test_live_comments import LiveCommentsTestCase

from verified_docx_mcp import (
    audit,
    comment_batch,
    mutations,
    paths,
    projection,
    server,
    tracked_changes,
)
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.lean_envelope import _KEEPS_REVISION_BEFORE
from verified_docx_mcp.live import comments_live, shared
from verified_docx_mcp.live.session import LiveOpFailed
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

PARAGRAPHS = [f"Finding {n} is stated here." for n in range(1, 13)]
QUOTES = [f"Finding {n} is" for n in range(1, 13)]
BODY_TEXT = "\n".join(PARAGRAPHS)


def items(n: int = 12, **extra):
    return [{"quote": QUOTES[i], "content": f"Comment {i + 1}", **extra} for i in range(n)]


def code_of(exc: BaseException) -> ErrorCode:
    return exc.envelope.error_code  # type: ignore[attr-defined]


class RegistrationTests(unittest.TestCase):
    def test_registered_as_mutating_and_keeps_revision_before(self):
        self.assertIn("add_anchored_comment_batch", MUTATING_TOOLS)
        self.assertIn("add_anchored_comment_batch", _KEEPS_REVISION_BEFORE)

    def test_validation_up_front(self):
        bad = [
            [],
            items(1) * 51,
            [{"quote": "", "content": "c"}],
            [{"quote": "q", "content": ""}],
            [{"quote": "q", "content": "c", "expected_matches": 0}],
            [{"quote": "q", "content": "c", "expected_matches": True}],
            [{"quote": "q", "content": "c", "section_key": ""}],
            [{"quote": "q", "content": "c", "extra": 1}],
            ["nope"],
        ]
        for value in bad:
            with self.subTest(value=str(value)[:40]):
                with self.assertRaises(VerifyError) as ctx:
                    comment_batch.validate_comments(value)
                self.assertEqual(code_of(ctx.exception), ErrorCode.INVALID_INPUT)
        self.assertEqual(len(comment_batch.validate_comments(items(1) * 50)), 50)
        explicit_null = comment_batch.validate_comments([{"quote": "q", "content": "c", "expected_matches": None}])
        self.assertEqual(explicit_null[0]["expected_matches"], 1)


class _FileCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.root = Path(self._tmp.name)

    def tearDown(self):
        if self._old is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old
        self._tmp.cleanup()

    def twelve_paragraph_doc(self) -> Path:
        target = self.root / "twelve.docx"
        # sections.docx declares the full Word namespace set the comment parts need
        # (empty-shell.docx does not, so any comment on it fails OPC validation).
        shutil.copyfile(FIXTURES / "sections.docx", target)
        mutations.execute_replace_body_markdown(str(target), "\n\n".join(PARAGRAPHS))
        return target

    def sections_doc(self) -> Path:
        target = self.root / "sections.docx"
        shutil.copyfile(FIXTURES / "sections.docx", target)
        return target

    @staticmethod
    def open_comments(target: Path) -> list[dict]:
        return tracked_changes.execute_list_open_items(str(target))["comments"]


class FileModeTests(_FileCase):
    def test_twelve_comments_one_call_one_save(self):
        target = self.twelve_paragraph_doc()
        result = server.add_anchored_comment_batch(str(target), items(), write_mode="file")
        self.assertTrue(result["applied"])
        self.assertEqual(result["batch_status"], "complete")
        self.assertEqual((result["applied_count"], result["item_count"]), (12, 12))
        self.assertEqual([r["status"] for r in result["results"]], ["applied"] * 12)
        self.assertEqual([r["index"] for r in result["results"]], list(range(12)))
        self.assertTrue(all(r["comment_id"] for r in result["results"]))
        self.assertNotEqual(result["revision_before"], result["revision_after"])
        comments = self.open_comments(target)
        self.assertEqual(len(comments), 12)
        self.assertEqual({c["quoted_text"] for c in comments}, set(QUOTES))
        self.assertEqual({c["content"] for c in comments}, {f"Comment {n}" for n in range(1, 13)})

    def test_mid_batch_zero_match_persists_nothing(self):
        target = self.twelve_paragraph_doc()
        before = target.read_bytes()
        batch = items(8)
        batch[4] = {"quote": "not in the document", "content": "x"}
        result = server.add_anchored_comment_batch(str(target), batch, write_mode="file")
        self.assertFalse(result["applied"])
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(result["applied_count"], 0)
        statuses = [r["status"] for r in result["results"]]
        self.assertEqual(statuses, ["not_persisted"] * 4 + ["refused_before_write"] + ["not_attempted"] * 3)
        self.assertEqual(result["results"][4]["code"], "ZERO_MATCH")
        self.assertEqual(target.read_bytes(), before)
        self.assertEqual(self.open_comments(target), [])

    def test_expected_matches_comments_every_match(self):
        target = self.sections_doc()
        result = server.add_anchored_comment_batch(
            str(target), [{"quote": "text.", "content": "all", "expected_matches": 3}], write_mode="file"
        )
        self.assertEqual(result["batch_status"], "complete")
        self.assertEqual(len(result["results"][0]["comment_ids"]), 3)
        self.assertEqual(len(self.open_comments(target)), 3)

    def test_section_key_scopes_the_search(self):
        target = self.sections_doc()
        # "text." occurs in three sections: unscoped with expected_matches=1 is a count mismatch.
        refused = server.add_anchored_comment_batch(
            str(target), [{"quote": "text.", "content": "c"}], write_mode="file"
        )
        self.assertEqual(refused["results"][0]["code"], "MATCH_COUNT_MISMATCH")
        scoped = server.add_anchored_comment_batch(
            str(target), [{"quote": "text.", "content": "c", "section_key": "background-1"}], write_mode="file"
        )
        self.assertEqual(scoped["batch_status"], "complete")
        comments = self.open_comments(target)
        self.assertEqual(len(comments), 1)
        proj = projection.project_part(target)
        anchored = [e for e in proj.events if isinstance(e, projection.RunEvent) and e.comment_ids]
        self.assertTrue(anchored)
        paragraph = next(m for m in proj.paragraphs if m.para_ref == anchored[0].para_ref)
        self.assertEqual(paragraph.para_ref, "p3")  # "Background text." sits in Background (p2-p3)

    def test_unknown_section_key_is_section_not_found(self):
        target = self.sections_doc()
        before = target.read_bytes()
        result = server.add_anchored_comment_batch(
            str(target), [{"quote": "text.", "content": "c", "section_key": "nope-1"}], write_mode="file"
        )
        self.assertEqual(result["results"][0]["code"], "SECTION_NOT_FOUND")
        self.assertEqual(target.read_bytes(), before)

    def test_one_audit_record_holds_every_item(self):
        target = self.twelve_paragraph_doc()
        batch = items(3)
        result = server.add_anchored_comment_batch(str(target), batch, write_mode="file")
        record = audit.find_audit_record(result["evidence_ref"])
        self.assertEqual(record["tool"], "add_anchored_comment_batch")
        ev = record["evidence"]
        self.assertEqual((ev["batch_status"], ev["applied_count"], ev["item_count"]), ("complete", 3, 3))
        self.assertEqual(len(ev["results"]), 3)
        for i, row in enumerate(ev["results"]):
            self.assertEqual(row["status"], "applied")
            self.assertEqual(row["quote"], batch[i]["quote"])
            self.assertEqual(row["content"], batch[i]["content"])
            self.assertTrue(row["comment_id"])
            self.assertEqual(row["match_count"], 1)

    def test_audit_keeps_no_text_when_excerpts_are_redacted(self):
        target = self.twelve_paragraph_doc()
        os.environ["VERIFIED_DOCX_MCP_AUDIT_EXCERPTS"] = "0"
        try:
            result = server.add_anchored_comment_batch(str(target), items(2), write_mode="file")
        finally:
            del os.environ["VERIFIED_DOCX_MCP_AUDIT_EXCERPTS"]
        row = audit.find_audit_record(result["evidence_ref"])["evidence"]["results"][0]
        self.assertNotIn("quote", row)
        self.assertNotIn("content", row)
        self.assertNotIn("before", row)
        self.assertEqual(row["quote_chars"], len(QUOTES[0]))

    def test_stale_revision_before_raises_before_anything(self):
        target = self.twelve_paragraph_doc()
        with self.assertRaises(Exception) as ctx:
            server.add_anchored_comment_batch(str(target), items(2), write_mode="file", revision_before="stale")
        self.assertIn("REVISION_CONFLICT", str(ctx.exception))

    def test_lean_envelope_keeps_revision_before_and_drops_constants(self):
        target = self.twelve_paragraph_doc()
        first = server.add_anchored_comment_batch(str(target), items(1), write_mode="file")
        token = first["revision_after"]

        async def run():
            async with Client(server.mcp) as client:
                return (
                    await client.call_tool(
                        "add_anchored_comment_batch",
                        {"path": str(target), "comments": items(2)[1:], "write_mode": "file", "revision_before": token},
                    )
                ).structured_content

        out = asyncio.run(run())
        self.assertEqual(out["revision_before"], token)
        for gone in ("write_mode", "audit_logged", "conflict_copy_detected"):
            self.assertNotIn(gone, out)
        self.assertEqual(out["batch_status"], "complete", out)


class LiveBatchTests(LiveCommentsTestCase):
    def doc(self) -> FakeDocument:
        return FakeDocument(text=BODY_TEXT)

    async def batch(self, batch, **kwargs):
        kwargs.setdefault("write_mode", "live")
        return await self.call(comment_batch.execute_add_anchored_comment_batch, str(self.target), batch, **kwargs)

    async def test_twelve_comments_in_one_call(self):
        pane = await self.connect_pane(document=self.doc())
        result = await self.batch(items())
        self.assertTrue(result["applied"])
        self.assertEqual(result["batch_status"], "complete")
        self.assertEqual((result["applied_count"], result["item_count"]), (12, 12))
        self.assertEqual([r["status"] for r in result["results"]], ["applied"] * 12)
        self.assertTrue(all(r["comment_id"].startswith("live:") for r in result["results"]))
        self.assertEqual([c.content for c in pane.document.comments], [f"Comment {n}" for n in range(1, 13)])
        self.assertTrue(result["revision_after"].startswith("live:sha256:"))

    async def test_mid_batch_zero_match_stops_and_keeps_earlier_comments(self):
        pane = await self.connect_pane(document=self.doc())
        batch = items(8)
        batch[3] = {"quote": "not in the document", "content": "x"}
        result = await self.batch(batch)
        self.assertEqual(result["batch_status"], "partial")
        self.assertEqual(result["applied_count"], 3)
        self.assertEqual(
            [r["status"] for r in result["results"]],
            ["applied"] * 3 + ["refused_before_write"] + ["not_attempted"] * 4,
        )
        self.assertEqual(result["results"][3]["code"], "ZERO_MATCH")
        self.assertEqual(len(pane.document.comments), 3)

    async def test_reply_without_comment_id_is_unknown_and_still_audited(self):
        await self.connect_pane(document=self.doc())
        real = comment_batch._send_comment_add
        calls = []

        def flaky(session, item, current):
            calls.append(item)
            reply = real(session, item, current)
            return {k: v for k, v in reply.items() if k != "comment_id"} if len(calls) == 2 else reply

        with unittest.mock.patch.object(comment_batch, "_send_comment_add", flaky):
            result = await self.batch(items(4))
        self.assertEqual(result["batch_status"], "unknown")
        self.assertEqual([r["status"] for r in result["results"]],
                         ["applied", "unknown", "not_attempted", "not_attempted"])
        self.assertIn("evidence_ref", result)

    async def test_first_item_refused_is_none_applied(self):
        pane = await self.connect_pane(document=self.doc())
        result = await self.batch([{"quote": "absent", "content": "x"}, *items(1)])
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(pane.document.comments, [])

    async def test_expected_matches_above_one_is_invalid_input_and_sends_nothing(self):
        pane = await self.connect_pane(document=self.doc())
        batch = items(2)
        batch[1]["expected_matches"] = 2
        with self.assertRaises(VerifyError) as ctx:
            await self.batch(batch)
        self.assertEqual(code_of(ctx.exception), ErrorCode.INVALID_INPUT)
        self.assertEqual(pane.document.comments, [])

    async def test_stale_revision_before_refuses_the_whole_batch(self):
        pane = await self.connect_pane(document=self.doc())
        result = await self.batch(items(3), revision_before="live:sha256:" + "0" * 64)
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual([r["status"] for r in result["results"]], ["refused_before_write", "not_attempted", "not_attempted"])
        self.assertEqual(result["results"][0]["code"], "LIVE_STALE")
        self.assertEqual(pane.document.comments, [])
        self.assertEqual(result["revision_before"], "live:sha256:" + "0" * 64)

    async def test_audit_record_holds_all_items(self):
        await self.connect_pane(document=self.doc())
        result = await self.batch(items(4))
        record = audit.find_audit_record(result["evidence_ref"])
        ev = record["evidence"]
        self.assertEqual(record["tool"], "add_anchored_comment_batch")
        self.assertEqual(ev["write_mode"], "live")
        self.assertEqual([r["status"] for r in ev["results"]], ["applied"] * 4)
        self.assertEqual([r["quote"] for r in ev["results"]], QUOTES[:4])

    async def test_no_section_key_means_no_body_read(self):
        pane = await self.connect_pane(document=self.doc())
        await self.batch(items(2))
        self.assertEqual(pane.body_ooxml_requests, 0)


SECTIONS_BODY = FIXTURES / "sections.docx"


class LiveSectionKeyTests(LiveCommentsTestCase):
    async def connect_sections(self):
        text = projection.read_document_text(SECTIONS_BODY)
        return await self.connect_pane(document=FakeDocument(text=text), body_ooxml=docx_to_flat_opc(SECTIONS_BODY))

    async def batch(self, batch):
        return await self.call(
            comment_batch.execute_add_anchored_comment_batch, str(self.target), batch, write_mode="live"
        )

    async def test_unique_quote_inside_its_section_is_commented(self):
        pane = await self.connect_sections()
        result = await self.batch([{"quote": "Background text.", "content": "c", "section_key": "background-1"}])
        self.assertEqual(result["batch_status"], "complete")
        self.assertEqual([c.anchor_text for c in pane.document.comments], ["Background text."])
        self.assertEqual(pane.body_ooxml_requests, 1)

    async def test_quote_also_outside_the_section_is_ambiguous(self):
        pane = await self.connect_sections()
        result = await self.batch([{"quote": "text.", "content": "c", "section_key": "background-1"}])
        self.assertEqual(result["results"][0]["status"], "refused_before_write")
        self.assertEqual(result["results"][0]["code"], "AMBIGUOUS_SPAN")
        self.assertEqual(pane.document.comments, [])

    async def test_quote_only_outside_the_section_is_zero_match(self):
        pane = await self.connect_sections()
        result = await self.batch([{"quote": "Some overview text.", "content": "c", "section_key": "background-1"}])
        self.assertEqual(result["results"][0]["code"], "ZERO_MATCH")
        self.assertEqual(pane.document.comments, [])

    async def test_unknown_section_stops_the_batch_after_earlier_items(self):
        pane = await self.connect_sections()
        result = await self.batch(
            [
                {"quote": "Background text.", "content": "a", "section_key": "background-1"},
                {"quote": "Next steps text.", "content": "b", "section_key": "gone-1"},
                {"quote": "Some overview text.", "content": "c"},
            ]
        )
        self.assertEqual(result["batch_status"], "partial")
        self.assertEqual(
            [(r["status"], r.get("code")) for r in result["results"]],
            [("applied", None), ("refused_before_write", "SECTION_NOT_FOUND"), ("not_attempted", None)],
        )
        self.assertEqual(len(pane.document.comments), 1)


# --- a stale shared-bridge client is not re-armed by the section_key read ----------------

from test_shared_bridge import Pane

from verified_docx_mcp.live import bridge, write_mode
from verified_docx_mcp.live.protocol import HelloMessage


class _OoxmlPane(Pane):
    hello = HelloMessage(
        Pane.document_url, "Word", "Mac", {}, "initial", frozenset({"shared_queue", "body_ooxml"}), "pane-1"
    )

    def request_threadsafe(self, op, payload=None, **kwargs):
        result = super().request_threadsafe(op, payload, **kwargs)
        if op == "body_ooxml":
            result = {**result, "ooxml": docx_to_flat_opc(SECTIONS_BODY)}
        return result


def test_section_key_read_does_not_rearm_a_stale_shared_bridge_client(monkeypatch):
    pane = _OoxmlPane()
    registry = SimpleNamespace(
        get=lambda name: pane if name == pane.document_name else None, list=lambda: [pane], collisions=list
    )
    broker = shared.Broker(registry)
    a = broker.dispatch({"action": "attach"})["client_id"]
    b = broker.dispatch({"action": "attach"})["client_id"]

    def call(client, op="describe", **kwargs):
        return broker.dispatch(
            {
                "action": "request", "client_id": client, "document_name": pane.document_name,
                "document_url": pane.document_url, "connected_since": pane.connected_since, "op": op, **kwargs,
            }
        )

    call(a)
    call(b)
    call(a, "replace")  # b's baseline is now behind

    client = SimpleNamespace(rpc=lambda action, **kw: broker.dispatch({"action": action, "client_id": b, **kw}))
    remote = shared.RemoteSession(client, broker.snapshot(pane))
    monkeypatch.setattr(bridge, "current_registry", lambda: SimpleNamespace(get=lambda n: remote))
    monkeypatch.setattr(write_mode, "_check_session_identity", lambda path, session: None)
    monkeypatch.setattr(write_mode, "_document_name", lambda path: "Proposal.docx")
    monkeypatch.setattr(comments_live, "_document_name_and_path", lambda path: (Path(path), "Proposal.docx"))

    result = comment_batch.execute_add_anchored_comment_batch(
        "/x/Proposal.docx",
        [{"quote": "Background text.", "content": "c", "section_key": "background-1"}],
        write_mode="live",
    )
    assert result["results"][0]["status"] == "refused_before_write"
    assert result["results"][0]["code"] == "LIVE_STALE"
    assert result["batch_status"] == "none_applied"
    assert any(op == "body_ooxml" for op, _ in pane.calls), "the section check really read the body"
    assert not any(op == "comment_add" for op, _ in pane.calls)
    time.sleep(0)
    # The section_key read did not re-arm b: a plain write is still refused.
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")


if __name__ == "__main__":
    unittest.main()
