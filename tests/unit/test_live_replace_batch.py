"""replace_text_batch, allow_comment_loss_ids, LIVE_STALE_AFTER_WRITE and open-pane
diagnostics (issue #82), against tests/unit/fake_pane.py (no Word).

The real pane's Office.js is covered by tests/unit/js/edit_harness.mjs; these tests pin the
server side: validation before anything is sent, the hash chain, the outcome classification
(refused_before_write vs unknown), the one-audit-record-per-batch rule and the error mapping.
"""

from __future__ import annotations

import asyncio
import json
import sys
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeComment, FakeDocument, FakeReply
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import comments_live, open_pane_live, open_pane_mac, shared, write_mode
from verified_docx_mcp.live.protocol import HelloMessage
from verified_docx_mcp.live.session import LiveOpFailed, LiveStale, LiveStaleAfterWrite

TEXT = "alpha beta gamma delta epsilon zeta eta theta iota kappa"


def _rev(doc: FakeDocument) -> str:
    return f"live:sha256:{doc.sha256()}"


def _audit_records(tool: str | None = None) -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    records = [json.loads(line) for line in path.read_text().splitlines() if line.strip()]
    return [r for r in records if tool is None or r["tool"] == tool]


def _edit(find: str, replace: str, expected: int = 1, **extra) -> dict:
    return {"find": find, "replace": replace, "expected_matches": expected, **extra}


class _BatchCase(LiveWriteBridgeTestCase):
    async def batch(self, edits, revision_before, **kwargs):
        return await asyncio.to_thread(
            text_edit.execute_replace_text_batch_live, str(self.target), edits, revision_before, **kwargs
        )

    async def replace(self, *args, **kwargs):
        return await asyncio.to_thread(text_edit.execute_replace_text, str(self.target), *args, **kwargs)


class ReplaceTextBatchTests(_BatchCase):
    async def test_complete_batch_lands_in_order_with_one_audit_record(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        before = len(_audit_records("replace_text_batch"))
        result = await self.batch(
            [_edit("alpha", "A"), _edit("gamma", "G"), _edit("kappa", "K")], _rev(doc)
        )
        self.assertEqual(result["batch_status"], "complete")
        self.assertTrue(result["applied"])
        self.assertEqual([r["status"] for r in result["results"]], ["applied"] * 3)
        self.assertEqual(doc.text, "A beta G delta epsilon zeta eta theta iota K")
        self.assertEqual(result["revision_after"], _rev(doc))
        self.assertEqual(result["results"][2]["evidence"]["revision_after"], _rev(doc))
        records = _audit_records("replace_text_batch")[before:]
        self.assertEqual(len(records), 1)  # one record per batch, not per edit
        self.assertEqual(records[0]["evidence"]["batch_status"], "complete")
        self.assertEqual(len(records[0]["evidence"]["results"]), 3)
        self.assertNotIn("before", records[0]["evidence"]["results"][0].get("evidence", {}))

    async def test_each_edit_is_guarded_by_the_previous_post_hash(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        seen: list[str] = []
        original = pane._dispatch

        def spy(op, payload):
            if op == "replace":
                seen.append(payload["expectedBodySha256"])
            return original(op, payload)

        pane._dispatch = spy
        first = doc.sha256()
        await self.batch([_edit("alpha", "A"), _edit("beta", "B")], _rev(doc))
        self.assertEqual(seen[0], first)
        self.assertNotEqual(seen[1], first)
        self.assertEqual(seen[1], FakeDocument(text="A beta gamma delta epsilon zeta eta theta iota kappa").sha256())

    async def test_mid_batch_refusal_is_partial(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        result = await self.batch(
            [_edit("alpha", "A"), _edit("missing", "X"), _edit("kappa", "K")], _rev(doc)
        )
        self.assertEqual(result["batch_status"], "partial")
        self.assertFalse(result["applied"])
        self.assertEqual(
            [r["status"] for r in result["results"]], ["applied", "refused_before_write", "not_attempted"]
        )
        self.assertEqual(result["results"][1]["code"], ErrorCode.ZERO_MATCH.value)
        self.assertEqual(result["revision_after"], _rev(doc))
        self.assertTrue(doc.text.startswith("A beta"))

    async def test_first_edit_refusal_is_none_applied(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        start = _rev(doc)
        result = await self.batch([_edit("alpha", "A", expected=2), _edit("beta", "B")], start)
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(result["results"][0]["code"], ErrorCode.MATCH_COUNT_MISMATCH.value)
        self.assertEqual(result["results"][1]["status"], "not_attempted")
        self.assertEqual(result["revision_after"], start)
        self.assertEqual(doc.text, TEXT)

    async def test_stale_at_start_is_none_applied(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        result = await self.batch([_edit("alpha", "A"), _edit("beta", "B")], "live:sha256:" + "0" * 64)
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(result["results"][0]["status"], "refused_before_write")
        self.assertEqual(result["results"][0]["code"], ErrorCode.LIVE_STALE.value)
        self.assertEqual(result["results"][1]["status"], "not_attempted")
        self.assertEqual(result["revision_after"], _rev(doc))
        self.assertEqual(pane.replace_count, 0)

    async def test_pane_disconnect_mid_batch_is_unknown(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        pane.disconnect_on_replace = 2
        result = await self.batch(
            [_edit("alpha", "A"), _edit("beta", "B"), _edit("gamma", "G")], _rev(doc)
        )
        self.assertEqual(result["batch_status"], "unknown")
        self.assertIsNone(result["revision_after"])
        self.assertFalse(result["applied"])
        self.assertEqual(
            [r["status"] for r in result["results"]], ["applied", "unknown", "not_attempted"]
        )
        self.assertEqual(result["results"][1]["code"], ErrorCode.LIVE_DISCONNECTED.value)

    async def test_post_write_hash_mismatch_is_unknown_not_refused(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        doc.report_pre = "f" * 64
        result = await self.batch([_edit("alpha", "A"), _edit("beta", "B")], _rev(doc))
        self.assertEqual(result["batch_status"], "unknown")
        self.assertEqual(result["results"][0]["status"], "unknown")
        self.assertEqual(result["results"][0]["code"], ErrorCode.LIVE_STALE_AFTER_WRITE.value)
        self.assertEqual(result["results"][1]["status"], "not_attempted")
        self.assertTrue(doc.text.startswith("A beta"))  # the write really landed

    async def test_allow_comment_loss_inside_an_edit_is_refused(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        with self.assertRaises(VerifyError) as ctx:
            await self.batch([_edit("alpha", "A", allow_comment_loss=True)], _rev(doc))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(pane.replace_count, 0)

    async def test_edit_count_bounds(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        for edits in ([], [_edit("alpha", "A")] * 51):
            with self.assertRaises(VerifyError) as ctx:
                await self.batch(edits, _rev(doc))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(pane.replace_count, 0)

    async def test_file_write_mode_and_bad_revision_are_refused(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        with self.assertRaises(VerifyError) as ctx:
            await self.batch([_edit("alpha", "A")], _rev(doc), write_mode="file")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        with self.assertRaises(VerifyError) as ctx:
            await self.batch([_edit("alpha", "A")], "sha-of-a-file-revision")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    async def test_no_pane_is_live_unavailable(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            await self.batch([_edit("alpha", "A")], "live:sha256:" + "0" * 64)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    async def test_old_pane_without_comment_loss_ids_is_capability_missing(self) -> None:
        doc = FakeDocument(text=TEXT)
        caps = ["replacement_formatting", "comment_loss_guard", "shape_guard"]
        await self.connect_pane(doc, capabilities=caps)
        with self.assertRaises(VerifyError) as ctx:
            await self.batch([_edit("alpha", "A", allow_comment_loss_ids=[])], _rev(doc))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_comment_loss_ids_in_a_batch_edit(self) -> None:
        doc = FakeDocument(text=TEXT)
        doc.comments.append(FakeComment("c1", "q", "A", "d", "beta", resolved=False))
        await self.connect_pane(doc)
        result = await self.batch([_edit("beta", "B", allow_comment_loss_ids=[])], _rev(doc))
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(result["results"][0]["code"], ErrorCode.WOULD_DELETE_COMMENTS.value)
        ok = await self.batch(
            [_edit("alpha", "A"), _edit("beta", "B", allow_comment_loss_ids=[{"comment_id": "c1", "reply_count": 0}])],
            _rev(doc),
        )
        self.assertEqual(ok["batch_status"], "partial")  # c1 is unresolved: refused after alpha applied


class AllowCommentLossIdsTests(_BatchCase):
    def _doc(self, *, resolved: bool = True, replies: int = 1) -> FakeDocument:
        doc = FakeDocument(text=TEXT)
        doc.comments.append(
            FakeComment(
                "c1", "q", "A", "d", "beta", resolved=resolved,
                replies=[FakeReply(f"r{i}", "a", "B", "d") for i in range(replies)],
            )
        )
        return doc

    async def test_exact_match_is_allowed_and_reported(self) -> None:
        doc = self._doc()
        pane = await self.connect_pane(doc)
        result = await self.replace(
            "beta", "B", 1, write_mode="live", allow_comment_loss_ids=[{"comment_id": "c1", "reply_count": 1}]
        )
        self.assertTrue(result["applied"])
        self.assertEqual(result["comments_removed"][0]["id"], "c1")
        self.assertEqual(pane.document.comments, [])

    async def test_unlisted_unresolved_and_wrong_count_are_refused(self) -> None:
        pane = await self.connect_pane(self._doc())
        cases = [
            (self._doc(), []),
            (self._doc(resolved=False), [{"comment_id": "c1", "reply_count": 1}]),
            (self._doc(), [{"comment_id": "c1", "reply_count": 0}]),
        ]
        for doc, ids in cases:
            with self.subTest(ids=ids, resolved=doc.comments[0].resolved):
                pane.document = doc
                with self.assertRaises(VerifyError) as ctx:
                    await self.replace("beta", "B", 1, write_mode="live", allow_comment_loss_ids=ids)
                self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.WOULD_DELETE_COMMENTS)
                self.assertEqual(doc.text, TEXT)

    async def test_payload_carries_the_ids_only_when_given(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        payloads: list[dict] = []
        original = pane._dispatch
        pane._dispatch = lambda op, payload: (payloads.append(payload) if op == "replace" else None) or original(op, payload)
        await self.replace("alpha", "A", 1, write_mode="live")
        await self.replace("gamma", "G", 1, write_mode="live", allow_comment_loss_ids=[])
        self.assertNotIn("allow_comment_loss_ids", payloads[0])
        self.assertEqual(payloads[1]["allow_comment_loss_ids"], [])

    async def test_validation_exclusivity_file_mode_and_capability(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc, capabilities=["replacement_formatting", "comment_loss_guard", "shape_guard"])
        bad = [
            {"allow_comment_loss_ids": [{"comment_id": "c1"}]},
            {"allow_comment_loss_ids": [{"comment_id": "c1", "reply_count": True}]},
            {"allow_comment_loss_ids": "c1"},
            {"allow_comment_loss_ids": [], "allow_comment_loss": True},
        ]
        for kwargs in bad:
            with self.subTest(kwargs=kwargs), self.assertRaises(VerifyError) as ctx:
                await self.replace("alpha", "A", 1, write_mode="live", **kwargs)
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        with self.assertRaises(VerifyError) as ctx:  # old pane
            await self.replace("alpha", "A", 1, write_mode="live", allow_comment_loss_ids=[])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)


class AllowCommentLossIdsFileModeTests(unittest.TestCase):
    def test_file_mode_refuses_with_live_only_message(self) -> None:
        with (
            patch.object(text_edit.live_write_mode, "resolve_write_mode", return_value="file"),
            self.assertRaises(VerifyError) as ctx,
        ):
            text_edit.execute_replace_text("/x.docx", "a", "b", 1, write_mode="file", allow_comment_loss_ids=[])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("live-only", ctx.exception.envelope.message)


class StaleAfterWriteTests(_BatchCase):
    async def test_replace_text_surfaces_stale_after_write_and_audits(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        doc.report_pre = "f" * 64
        before = len(_audit_records("replace_text:verification_failed"))
        with self.assertRaises(VerifyError) as ctx:
            await self.replace("alpha", "A", 1, write_mode="live")
        env = ctx.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.LIVE_STALE_AFTER_WRITE)
        self.assertFalse(env.retryable)
        self.assertTrue(doc.text.startswith("A beta"))
        records = _audit_records("replace_text:verification_failed")[before:]
        self.assertEqual(len(records), 1)
        self.assertTrue(records[0]["evidence"]["applied"])

    async def test_pane_side_stale_is_live_stale_and_writes_nothing(self) -> None:
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(doc)
        original = pane._dispatch

        def moving(op, payload):
            if op == "replace":
                doc.text += " moved"  # a co-author typed between describe and replace
            return original(op, payload)

        pane._dispatch = moving
        with self.assertRaises(VerifyError) as ctx:
            await self.replace("alpha", "A", 1, write_mode="live")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_STALE)
        self.assertTrue(doc.text.startswith("alpha"))


@pytest.mark.parametrize("code", ["stale", "LIVE_STALE", "Stale"])
def test_pane_stale_maps_to_live_stale_everywhere(code):
    exc = LiveOpFailed(code, "the document changed")
    assert write_mode.classify_op_failed(exc) == ErrorCode.LIVE_STALE
    with pytest.raises(VerifyError) as err:
        comments_live._raise_from_live_error(exc)
    assert err.value.envelope.error_code == ErrorCode.LIVE_STALE


def test_comment_paragraph_and_revision_paths_share_the_stale_mapping():
    stub = SimpleNamespace(request_threadsafe=lambda op, payload=None, **kw: (_ for _ in ()).throw(LiveOpFailed("stale", "moved")))
    with pytest.raises(VerifyError) as err:
        comments_live._request(stub, "paragraph_delete", {})  # delete_paragraph / insert_paragraphs / revisions
    assert err.value.envelope.error_code == ErrorCode.LIVE_STALE


def test_session_layer_error_classes():
    assert issubclass(LiveStaleAfterWrite, LiveStale)
    with pytest.raises(VerifyError) as err:
        comments_live._raise_from_live_error(LiveStaleAfterWrite("a", "b"))
    assert err.value.envelope.error_code == ErrorCode.LIVE_STALE_AFTER_WRITE
    assert err.value.envelope.retryable is False


# -- shared mode ------------------------------------------------------------


class _SharedPane:
    document_name = "Proposal.docx"
    document_url = "https://example.test/site/Proposal.docx"
    connected_since = 1.0
    last_body_sha256 = "initial"
    hello = HelloMessage(document_url, "Word", "Mac", {}, "initial", frozenset({"shared_queue"}), "pane-1")

    def heartbeat_age(self):
        return 0

    def request_threadsafe(self, op, payload=None, **kwargs):
        if op == "describe":
            return {"bodySha256": self.last_body_sha256}
        raise LiveStaleAfterWrite("initial", "moved")


def test_shared_broker_serializes_stale_after_write_and_client_maps_it_back(tmp_path):
    import tempfile

    pane = _SharedPane()
    registry = SimpleNamespace(
        get=lambda name: pane if name == pane.document_name else None, list=lambda: [pane], collisions=list
    )
    broker = shared.Broker(registry)
    short = tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-s82-")
    path = Path(short.name) / "rpc.sock"
    server = shared.Server(str(path), shared.Handler)
    server.broker = broker
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    client = shared.Client(path)
    try:
        session = client.get(pane.document_name)
        session.request_threadsafe("describe")
        with pytest.raises(LiveStaleAfterWrite) as raised:
            session.request_threadsafe("replace", {"expectedBodySha256": "initial"})
        assert (raised.value.expected, raised.value.actual) == ("initial", "moved")
        # ... and text_edit turns it into the caller-visible code, with the write flagged as landed.
        session.request_threadsafe("describe")
        with pytest.raises(VerifyError) as err:
            text_edit._live_replace_once(
                session, find="a", replace="b", expected_matches=1, pre_hash="initial"
            )
        assert err.value.envelope.error_code == ErrorCode.LIVE_STALE_AFTER_WRITE
    finally:
        client.close()
        server.shutdown()
        server.server_close()
        thread.join()
        short.cleanup()


# -- live_open_pane diagnostics ----------------------------------------------


def _open_with(error: Exception):
    registry = SimpleNamespace(get=lambda name: None)
    with (
        patch.object(open_pane_live.live_bridge, "start_in_background", return_value=registry),
        patch.object(open_pane_live, "_open_documents", return_value=["doc.docx"]),
        patch.object(open_pane_live.open_pane_mac, "open_pane", side_effect=error),
        pytest.raises(VerifyError) as err,
    ):
        open_pane_live.execute_open_pane(str(Path.home() / "doc.docx"))
    return err.value.envelope


@pytest.mark.parametrize(
    ("reason", "retryable"),
    [("ribbon_unreachable", True), ("front_failed", True), ("not_mac", False), ("other", False)],
)
def test_open_pane_error_reason_hint_and_retryable(reason, retryable):
    env = _open_with(open_pane_mac.OpenPaneError("boom", reason))
    assert env.error_code == ErrorCode.LIVE_UNAVAILABLE
    assert env.diagnostics["reason"] == reason
    assert "Unlock the screen" in env.diagnostics["hint"]
    assert env.retryable is retryable


def test_open_pane_non_pane_errors_are_other_and_not_retryable():
    env = _open_with(OSError("no such app"))
    assert env.diagnostics["reason"] == "other"
    assert env.retryable is False


def test_open_pane_mac_classifies_failures():
    def run(osascript, system="Darwin"):
        with patch.object(open_pane_mac, "_osascript", osascript), patch.object(open_pane_mac.time, "sleep"), \
             patch.object(open_pane_mac.platform, "system", return_value=system), pytest.raises(open_pane_mac.OpenPaneError) as err:
            open_pane_mac.open_pane("doc.docx", lambda: False, 1e9)
        return err.value

    assert run(lambda *a, **k: "ready", system="Linux").reason == "not_mac"

    def no_front(script, *args, **kw):
        if script is open_pane_mac._PREPARE:
            raise open_pane_mac.OpenPaneError("could not bring doc.docx to the front")
        return "ready"

    assert run(no_front).reason == "front_failed"

    def no_ribbon(script, *args, **kw):
        return "absent" if script in (open_pane_mac._PRESS_LIVE_PANE, open_pane_mac._ADDINS_CENTRE) else "ready"

    error = run(no_ribbon)
    assert error.reason == "ribbon_unreachable"
    assert "Unlock the screen" in str(error)


def test_unlisted_codes_are_not_globally_retryable():
    from verified_docx_mcp.errors import _make_error

    assert _make_error(ErrorCode.LIVE_UNAVAILABLE, "x").envelope.retryable is False


class CompactBatchReportTests(_BatchCase):
    """Issue #112: the report names what callers read and nothing else."""

    async def test_fifty_edit_success_is_compact_and_keeps_the_hook_fields(self) -> None:
        words = [f"w{n:02d}" for n in range(50)]
        doc = FakeDocument(text=" ".join(words))
        await self.connect_pane(doc)
        before = _rev(doc)
        result = await self.batch([_edit(w, w.upper()) for w in words], before)

        # JennyStack's hook validates exactly these (live_doc_stamp.record_batch_result).
        self.assertEqual(
            set(result),
            {"applied", "batch_status", "applied_count", "edit_count", "revision_before", "revision_after", "results",
             "evidence_ref"},
        )
        self.assertIs(result["applied"], True)
        self.assertEqual((result["batch_status"], result["applied_count"], result["edit_count"]), ("complete", 50, 50))
        self.assertEqual(result["revision_before"], before)
        self.assertEqual(result["revision_after"], _rev(doc))
        self.assertEqual([r["index"] for r in result["results"]], list(range(50)))
        for r in result["results"]:
            self.assertEqual(set(r), {"index", "status", "evidence"})
            self.assertEqual(r["status"], "applied")
            self.assertEqual(r["evidence"]["match_count"], 1)
            self.assertRegex(r["evidence"]["revision_after"], r"^live:sha256:[0-9a-f]{64}$")
        self.assertEqual(result["results"][-1]["evidence"]["revision_after"], result["revision_after"])
        for gone in ("before", "after", "rung", "match_count", "write_mode", "verified_via", "document_name"):
            self.assertNotIn(gone, result)
        # Recorded in docs/acceptance/112-write-tokens.md (chars / 4). The per-edit full
        # revision_after is what JennyStack's hook chains, so it is the floor: ~157 chars per edit.
        compact_tokens = len(json.dumps(result, separators=(",", ":"))) // 4
        legacy = await self.batch([_edit(w.upper(), w) for w in words], _rev(doc), verbose=True)
        legacy_tokens = len(json.dumps(legacy, separators=(",", ":"))) // 4
        print(f"\n50-edit batch: compact {compact_tokens} tokens, verbose (previous shape) {legacy_tokens} tokens")
        self.assertLess(compact_tokens, legacy_tokens)
        self.assertLess(compact_tokens, 2200)

    async def test_mid_batch_refusal_names_the_failing_index_code_and_message(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        result = await self.batch(
            [_edit("alpha", "A"), _edit("nowhere", "X"), _edit("gamma", "G")], _rev(doc)
        )
        self.assertEqual(result["batch_status"], "partial")
        self.assertEqual(result["results"][0]["status"], "applied")
        failed = result["results"][1]
        self.assertEqual((failed["index"], failed["status"], failed["code"]), (1, "refused_before_write", "ZERO_MATCH"))
        self.assertTrue(failed["message"])
        self.assertLessEqual(len(failed["message"]), 240)
        self.assertEqual(result["results"][2], {"index": 2, "status": "not_attempted"})
        self.assertEqual(result["revision_after"], result["results"][0]["evidence"]["revision_after"])

    async def test_stale_start_is_reported_with_its_message(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        result = await self.batch([_edit("alpha", "A")], "live:sha256:" + "0" * 64)
        self.assertEqual(result["batch_status"], "none_applied")
        self.assertEqual(result["results"][0]["code"], ErrorCode.LIVE_STALE.value)
        self.assertTrue(result["results"][0]["message"])
        self.assertEqual(result["revision_before"], "live:sha256:" + "0" * 64)

    async def test_verbose_restores_per_edit_before_and_after(self) -> None:
        doc = FakeDocument(text=TEXT)
        await self.connect_pane(doc)
        result = await self.batch([_edit("alpha", "A")], _rev(doc), verbose=True)
        evidence = result["results"][0]["evidence"]
        self.assertEqual((evidence["before"], evidence["after"]), ("alpha", "A"))
        self.assertIn("comments_removed", evidence)
        self.assertIn("write_mode", result)  # the full live envelope, as before
