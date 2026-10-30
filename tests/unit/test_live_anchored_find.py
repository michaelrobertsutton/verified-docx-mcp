"""Issue #113 part A, live mode: anchored find and para_ref targets against
tests/unit/fake_pane.py (no Word), the shared-bridge stale-client rule, and approved
runs. The pane's Office.js side is tests/unit/js/anchored_find_harness.mjs."""

from __future__ import annotations

import asyncio
import os
import sys
import tempfile
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, FakePane, OpRefused  # noqa: E402
from test_anchored_find import FIND, LONG  # noqa: E402
from test_live_write_mode import LiveWriteBridgeTestCase  # noqa: E402

from verified_docx_mcp import locate, paths, text_edit  # noqa: E402
from verified_docx_mcp.errors import ErrorCode, VerifyError  # noqa: E402
from verified_docx_mcp.live import bridge, paragraphs_live, shared  # noqa: E402
from verified_docx_mcp.live.protocol import HelloMessage  # noqa: E402
from verified_docx_mcp.live.session import LiveOpFailed  # noqa: E402

DEFAULT_CAPS = FakePane().capabilities
CAPS = [*DEFAULT_CAPS, "anchored_find_v1", "delete_paragraph"]
SHORT = "The short closing paragraph ends right here."
TEXT = f"Heading\n{LONG}\n{SHORT}"


def _rev(doc: FakeDocument) -> str:
    return f"live:sha256:{doc.sha256()}"


class LiveAnchoredTests(LiveWriteBridgeTestCase):
    async def open(self, text: str = TEXT, caps: list[str] = CAPS) -> tuple[FakeDocument, FakePane]:
        doc = FakeDocument(text=text)
        pane = await self.connect_pane(doc, capabilities=caps, body_ooxml_from_text=True)
        return doc, pane

    async def replace(self, *args, **kwargs):
        kwargs.setdefault("write_mode", "live")
        return await asyncio.to_thread(text_edit.execute_replace_text, str(self.target), *args, **kwargs)

    async def code_of(self, coro) -> ErrorCode:
        with self.assertRaises(VerifyError) as ctx:
            await coro
        return ctx.exception.envelope.error_code

    async def test_long_anchored_replace_sends_prefix_suffix_and_hash(self) -> None:
        doc, pane = await self.open()
        evidence = await self.replace(FIND, "Rewritten.", None, expected_length=412, revision_before=_rev(doc))
        self.assertEqual(doc.text, f"Heading\nRewritten.\n{SHORT}")
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(evidence["before"], LONG)
        payload = pane.op_payloads["replace"][-1]
        self.assertNotIn("find", payload)
        self.assertEqual(payload["prefix"], "Skyward will modernize")
        self.assertEqual(payload["suffix"], "batch jobs one workload at a time by 2027.")
        self.assertEqual(payload["expected_text_sha256"], locate.span_sha256(LONG))
        self.assertEqual(payload["expected_matches"], 1)

    async def test_expected_length_400_is_refused_before_anything_is_sent(self) -> None:
        doc, pane = await self.open()
        code = await self.code_of(self.replace(FIND, "x", None, expected_length=400))
        self.assertEqual(code, ErrorCode.ZERO_MATCH)
        self.assertEqual(doc.text, TEXT)
        self.assertNotIn("replace", pane.op_payloads)

    async def test_a_short_anchored_span_goes_as_plain_find_and_works_on_an_old_pane(self) -> None:
        doc, pane = await self.open(caps=DEFAULT_CAPS)
        await self.replace("The short closing…ends right here.", "Done.", None, expected_length=len(SHORT))
        self.assertEqual(doc.text, f"Heading\n{LONG}\nDone.")
        self.assertEqual(pane.op_payloads["replace"][-1]["find"], SHORT)

    async def test_a_long_span_on_an_old_pane_is_capability_missing(self) -> None:
        doc, pane = await self.open(caps=DEFAULT_CAPS)
        code = await self.code_of(self.replace(FIND, "x", None, expected_length=412))
        self.assertEqual(code, ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertEqual(doc.text, TEXT)
        self.assertNotIn("replace", pane.op_payloads)

    async def test_para_ref_replaces_a_paragraph_and_refuses_a_stale_hash(self) -> None:
        doc, _pane = await self.open()
        code = await self.code_of(self.replace(None, "x", None, para_ref="p1", expected_sha256=locate.span_sha256("old")))
        self.assertEqual(code, ErrorCode.PARA_HASH_MISMATCH)
        await self.replace(None, "New middle.", None, para_ref="p1", expected_sha256=locate.span_sha256(LONG))
        self.assertEqual(doc.text, f"Heading\nNew middle.\n{SHORT}")

    async def test_a_duplicated_paragraph_is_ambiguous_in_live_mode(self) -> None:
        for body in (f"{LONG}\n{LONG}", f"{SHORT}\nmiddle\n{SHORT}"):
            doc, pane = await self.open(text=body)
            target = LONG if body.startswith(LONG) else SHORT
            code = await self.code_of(self.replace(None, "x", None, para_ref="p0", expected_sha256=locate.span_sha256(target)))
            self.assertEqual(code, ErrorCode.AMBIGUOUS_SPAN)
            self.assertEqual(doc.text, body)
            await pane.close()

    async def test_anchored_format_text_live(self) -> None:
        doc, pane = await self.open()
        evidence = await asyncio.to_thread(
            text_edit.execute_format_text, str(self.target), FIND, {"bold": True}, None,
            find_sha256=locate.span_sha256(LONG), write_mode="live",
        )
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(pane.op_payloads["format"][-1]["prefix"], "Skyward will modernize")

    async def test_batch_resolves_each_edit_against_the_previous_edits_output(self) -> None:
        doc, _pane = await self.open()
        rewritten = "Closing paragraph rewritten with enough words to anchor."
        result = await asyncio.to_thread(
            text_edit.execute_replace_text_batch_live, str(self.target), [
                {"para_ref": "p2", "expected_sha256": locate.span_sha256(SHORT), "replace": rewritten},
                {"find": "Closing paragraph…enough words to anchor.", "expected_length": len(rewritten),
                 "replace": "Final."},
                {"find": FIND, "find_sha256": locate.span_sha256(LONG), "replace": "Middle."},
                {"find": "Heading", "expected_matches": 1, "replace": "Title"},
            ], _rev(doc),
        )
        self.assertEqual(result["batch_status"], "complete", result)
        self.assertEqual(doc.text, "Title\nMiddle.\nFinal.")

    async def test_batch_resolution_failure_is_refused_before_write(self) -> None:
        doc, _pane = await self.open()
        result = await asyncio.to_thread(
            text_edit.execute_replace_text_batch_live, str(self.target), [
                {"find": "Heading", "expected_matches": 1, "replace": "Title"},
                {"para_ref": "p1", "expected_sha256": locate.span_sha256("stale"), "replace": "x"},
                {"find": SHORT, "expected_matches": 1, "replace": "y"},
            ], _rev(doc),
        )
        self.assertEqual(result["batch_status"], "partial")
        self.assertEqual([r["status"] for r in result["results"]], ["applied", "refused_before_write", "not_attempted"])
        self.assertEqual(result["results"][1]["code"], "PARA_HASH_MISMATCH")
        self.assertEqual(doc.text, f"Title\n{LONG}\n{SHORT}")

    async def test_batch_edit_validation(self) -> None:
        doc, _pane = await self.open()
        for edit in ({"find": FIND, "replace": "x"}, {"para_ref": "p1", "replace": "x"},
                     {"find": "Heading", "replace": "x", "expected_length": 7, "expected_matches": 1}):
            with self.assertRaises(VerifyError) as ctx:
                await asyncio.to_thread(text_edit.execute_replace_text_batch_live, str(self.target), [edit], _rev(doc))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    async def test_delete_paragraph_with_an_anchored_anchor_text(self) -> None:
        doc, pane = await self.open()
        evidence = await asyncio.to_thread(
            paragraphs_live.delete_paragraph, str(self.target), FIND, _rev(doc), False, 412, None
        )
        self.assertTrue(evidence["applied"])
        self.assertEqual(doc.text, f"Heading\n{SHORT}")
        with self.assertRaises(VerifyError) as ctx:  # a span that is not a whole paragraph
            await asyncio.to_thread(
                paragraphs_live.delete_paragraph, str(self.target), "The short closing…ends right", _rev(doc),
                False, len(SHORT) - 6, None,
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


# ---------------------------------------------------------------------------
# Safety rule: resolving a target never re-arms a stale shared-bridge client.
# ---------------------------------------------------------------------------


class _BrokerPane:
    """The broker's pane session: a FakePane's dispatcher, called in-process."""

    document_name = "anchored.docx"
    document_url = "https://example.test/site/anchored.docx"
    connected_since = 1.0
    body_hash_state = "sampled"
    heartbeat_count = 0
    hash_sample_count = 0

    def __init__(self, doc: FakeDocument) -> None:
        self.fake = FakePane(document=doc, body_ooxml_from_text=True)
        self.calls: list[str] = []
        self.last_body_sha256 = doc.sha256()
        self.body_hash_sampled_monotonic = time.monotonic()
        caps = frozenset([*CAPS, "shared_queue"])
        self.hello = HelloMessage(self.document_url, "Word", "Mac", {}, doc.sha256(), caps, "pane-1")

    def heartbeat_age(self):
        return 0

    def request_threadsafe(self, op, payload=None, **_kwargs):
        self.calls.append(op)
        try:
            return self.fake._dispatch(op, payload or {})
        except OpRefused as exc:
            raise LiveOpFailed(exc.code, exc.message) from exc


class _ClientSession:
    """What the shared-bridge client process sees (RemoteSession's contract)."""

    def __init__(self, broker: shared.Broker, client_id: str, pane: _BrokerPane, observe: bool = True) -> None:
        self.broker, self.client_id, self.pane, self._observe = broker, client_id, pane, observe
        self.hello, self.document_name, self.document_url = pane.hello, pane.document_name, pane.document_url

    def status_view(self):
        return _ClientSession(self.broker, self.client_id, self.pane, observe=False)

    def request_threadsafe(self, op, payload=None, *, timeout=None, expected_body_sha256=None, request_id=None):
        return self.broker.dispatch({
            "action": "request", "client_id": self.client_id, "document_name": self.pane.document_name,
            "document_url": self.pane.document_url, "connected_since": self.pane.connected_since,
            "op": op, "payload": payload, "expected_body_sha256": expected_body_sha256, "observe": self._observe,
        })


@pytest.fixture
def stale_client(monkeypatch):
    tmp = tempfile.TemporaryDirectory()
    monkeypatch.setenv(paths._ALLOWED_FILE_ROOTS_ENV, tmp.name)
    monkeypatch.setenv("XDG_STATE_HOME", os.path.join(tmp.name, "state"))
    doc = FakeDocument(text=TEXT)
    pane = _BrokerPane(doc)
    registry = SimpleNamespace(get=lambda name: pane if name == pane.document_name else None,
                               list=lambda: [pane], collisions=list)
    broker = shared.Broker(registry)
    a = _ClientSession(broker, broker.dispatch({"action": "attach"})["client_id"], pane)
    b = _ClientSession(broker, broker.dispatch({"action": "attach"})["client_id"], pane)
    a.request_threadsafe("describe")
    b.request_threadsafe("describe")
    # Another client writes: b's baseline is now stale.
    a.request_threadsafe("replace", {"find": "Heading", "expected_matches": 1, "replace": "Title"})
    monkeypatch.setattr(bridge, "current_registry",
                        lambda: SimpleNamespace(get=lambda name: b if name == pane.document_name else None,
                                                collisions=list))
    yield str(Path(tmp.name) / pane.document_name), doc, pane, broker, b
    tmp.cleanup()


@pytest.mark.parametrize("target", [
    {"find": FIND, "expected_length": 412},
    {"find": None, "para_ref": "p1", "expected_sha256": locate.span_sha256(LONG)},
])
def test_a_stale_client_is_refused_and_not_rearmed(stale_client, target):
    path, doc, pane, broker, b = stale_client
    key = pane.document_url
    seen_before = broker.seen.get((b.client_id, key))
    for _ in range(2):  # the retry must stay refused too
        with pytest.raises(VerifyError) as ctx:
            text_edit.execute_replace_text(path, replace="x", expected_matches=None, write_mode="live", **target)
        assert ctx.value.envelope.error_code == ErrorCode.LIVE_STALE
    assert "body_ooxml" in pane.calls  # the target was resolved -- without re-arming b
    assert broker.seen.get((b.client_id, key)) == seen_before
    assert doc.text == f"Title\n{LONG}\n{SHORT}"
    # An explicit (observing) read re-arms b; then the same anchored write lands.
    b.request_threadsafe("describe")
    text_edit.execute_replace_text(path, replace="x", expected_matches=None, write_mode="live", **target)
    assert doc.text == f"Title\nx\n{SHORT}"
