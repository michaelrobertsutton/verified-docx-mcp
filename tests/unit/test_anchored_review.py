"""Issue #113 part A, review findings: the uncapped candidate set, crossing spans that pass,
literal ellipses in plain finds, normalized expected_length, UTF-16 anchor limits and the
pane's 50-hit prefix cap."""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_anchored_find import E, _FileCase, _p, _proj
from test_approved_anchored import op, setup  # noqa: F401
from test_shared_bridge import Pane as BrokerPane

from verified_docx_mcp import locate, server, text_edit
from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import shared
from verified_docx_mcp.live.session import LiveOpFailed

NBSP = chr(0xA0)
FIND = f"The Contractor shall{E}to the COR within five days."


def _code(ctx: pytest.ExceptionInfo) -> str:
    return ctx.value.envelope.error_code.value


def _sixty():
    paras = []
    for i in range(60):
        if i == 10:
            mid = "pay the invoice for all " + "x" * 30
        elif i == 55:
            mid = "deliver the security plan " + "y" * 28
        else:
            mid = "z" * (60 + i * 3)
        paras.append(f"The Contractor shall {mid} to the COR within five days.")
    assert len(paras[10]) == len(paras[55])
    return paras


# 1. Every candidate feeds the uniqueness decision (no 50-pair cap).

def test_sixty_shared_prefixes_same_length_twice_is_ambiguous():
    paras = _sixty()
    proj = _proj(*[_p(t) for t in paras])
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=len(paras[55]))
    assert _code(ctx) == "AMBIGUOUS_SPAN"
    assert ctx.value.envelope.diagnostics["total"] == 2
    # find_sha256 picks #55 itself, past the 50th candidate.
    result = locate.locate_anchored(FIND, proj, find_sha256=locate.span_sha256(paras[55]))
    assert proj.text[slice(*result.spans[0])] == paras[55]


def test_identical_copies_at_10_and_55_are_ambiguous_with_find_sha256():
    paras = _sixty()
    paras[55] = paras[10]
    proj = _proj(*[_p(t) for t in paras])
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, find_sha256=locate.span_sha256(paras[10]))
    assert _code(ctx) == "AMBIGUOUS_SPAN"
    assert ctx.value.envelope.diagnostics["total"] == 2
    assert len(ctx.value.envelope.diagnostics["candidates"]) == 2


def test_the_pane_cap_of_fifty_prefix_hits_is_refused_before_sending():
    suffix = "deliver the security plan " + "with named reviewers and dated sign-offs " * 6 + "to the COR."
    target = f"The Contractor shall {suffix}"
    paras = [f"The Contractor shall do task {i}." for i in range(59)] + [target]
    proj = _proj(*[_p(t) for t in paras])
    result = locate.locate_anchored(f"The Contractor shall{E}{suffix}", proj, expected_length=len(target))
    with pytest.raises(VerifyError) as ctx:
        text_edit.pane_target(None, proj, result)
    assert _code(ctx) == "AMBIGUOUS_SPAN"
    assert ctx.value.envelope.diagnostics["prefix_hits"] == 60
    assert "lengthen the prefix" in ctx.value.envelope.message


def test_the_pane_prefix_reaches_toward_the_suffix():
    paras = [f"The Contractor shall do task {i}." for i in range(59)]
    long_target = "The Contractor shall deliver " + "a unique plan " * 30 + "to the COR within five days."
    proj = _proj(*[_p(t) for t in [*paras, long_target]])
    result = locate.locate_anchored(FIND, proj, expected_length=len(long_target))
    target = text_edit.pane_target(None, proj, result)
    assert target["prefix"] == long_target[:230]  # one hit, not 60
    assert target["suffix"] == "to the COR within five days."


# 2. A crossing span that passes is never passed over for an inside one.

def test_a_passing_crossing_span_and_a_passing_inside_span_are_ambiguous():
    p1 = "Skyward will deliver the transition plan within thirty days"
    p2 = "of award and report weekly to the program office."
    cross = len(p1) + 1 + len(p2)
    p3 = "Skyward will deliver the help desk staffing model and then report weekly to the program office."
    p3 = p3.replace("help desk", "help desk" + "!" * (cross - len(p3)))
    assert len(p3) == cross
    proj = _proj(_p(p1), _p(p2), _p(p3))
    find = f"Skyward will deliver{E}weekly to the program office."
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(find, proj, expected_length=cross)
    assert _code(ctx) == "AMBIGUOUS_SPAN"
    assert ctx.value.envelope.diagnostics["crossing_boundaries"] == 1
    # Only the crossing one passes: STRUCTURAL_BOUNDARY, as before.
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(find, _proj(_p(p1), _p(p2), _p(p3.replace("help", "the help"))), expected_length=cross)
    assert _code(ctx) == "STRUCTURAL_BOUNDARY"


# 3. A literal ellipsis is plain text unless expected_length/find_sha256 is passed.

def test_a_literal_ellipsis_is_a_plain_find():
    proj = _proj(_p(f"Wait{E} what did the RFP [...] say?"))
    assert locate.locate_find(f"Wait{E} what", proj, 1).rung == "exact"
    assert locate.locate_find("RFP [...] say", proj, 1).rung == "exact"
    assert text_edit.check_target_args(f"Wait{E} what", 1) == "plain"
    assert text_edit.check_target_args(f"Wait{E} what{E} now", 1) == "plain"
    with pytest.raises(VerifyError) as ctx:  # anchored arguments, but no delimiter
        text_edit.check_target_args("no delimiter here at all", None, expected_length=24)
    assert _code(ctx) == "INVALID_INPUT"


class LiteralEllipsisFileTests(_FileCase):
    def setUp(self) -> None:
        super().setUp()
        self.write(_p(f"Wait{E} what did the RFP say?"), _p("Second paragraph, unchanged."))

    def test_replace_text_and_the_mcp_tool_take_a_literal_ellipsis(self):
        text_edit.execute_replace_text(str(self.target), f"Wait{E} what", "So", 1, write_mode="file")
        self.assertEqual(self.text(), "So did the RFP say?\nSecond paragraph, unchanged.")

        async def call(args):
            from fastmcp import Client
            async with Client(server.mcp) as client:
                return await client.call_tool("replace_text", args, raise_on_error=False)

        plain = asyncio.run(call({"path": str(self.target), "find": "So did", "replace": f"Hm{E} did",
                                  "expected_matches": 1, "write_mode": "file"}))
        self.assertFalse(plain.is_error, plain)
        para = asyncio.run(call({"path": str(self.target), "para_ref": "p1", "replace": "Second, changed.",
                                 "expected_sha256": locate.span_sha256("Second paragraph, unchanged."),
                                 "write_mode": "file"}))
        self.assertFalse(para.is_error, para)
        self.assertEqual(self.text(), f"Hm{E} did the RFP say?\nSecond, changed.")
        missing = asyncio.run(call({"path": str(self.target), "find": "Second", "expected_matches": 1}))
        self.assertTrue(missing.is_error)
        self.assertEqual(json.loads(missing.content[0].text)["error_code"], "INVALID_INPUT")


def test_an_approved_plain_find_may_contain_an_ellipsis(setup):  # noqa: F811
    path, pane = setup
    pane.sections[0]["paragraphs"][1] = f"Proof{E} is current."
    prepared = runs.prepare(path, [op(find=f"Proof{E} is", replace="Proof is", expected_matches=1)])
    assert runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")["status"] == "complete"
    assert pane.sections[0]["paragraphs"][1] == "Proof is current."


# 6. expected_length counts the normalized span.

def test_expected_length_counts_the_normalized_span():
    text = "Skyward  will   deliver" + NBSP + " the  plan  on  time  and  report  weekly  to  the  office."
    proj = _proj(_p(text))
    find = f"Skyward  will   deliver{E}weekly  to  the  office."
    normalized = len(locate.normalize_span_text(text))
    assert len(text) - normalized > 4
    assert locate.locate_anchored(find, proj, expected_length=normalized).match_count == 1
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(find, proj, expected_length=len(text))
    assert _code(ctx) == "ZERO_MATCH"
    assert ctx.value.envelope.diagnostics["candidates"][0]["length"] == normalized


# 4. Anchor limits are UTF-16 units, as the pane's String.length counts.

def test_anchor_limits_are_utf16_units():
    smile = chr(0x1F600)
    span = smile * 5 + "a" * 300
    prefix, suffix = locate.pane_anchors(span, 290)
    assert locate.utf16_len(prefix) <= 230 and locate.utf16_len(suffix) <= 230
    assert prefix.startswith(smile * 5) and len(prefix) == 225  # 5 astral (10 units) + 220
    # 200 code points but 240 UTF-16 units: too long for one Word search, so it goes anchored.
    text = "The emoji paragraph opens here " + smile * 40 + "b" * 129 + " and it closes right here."
    assert len(text) < 230 < locate.utf16_len(text)
    proj = _proj(_p(text))
    result = locate.locate_paragraph(proj, "p0", locate.span_sha256(text))
    target = text_edit.pane_target(None, proj, result)
    assert "find" not in target and locate.utf16_len(target["prefix"]) <= 230


# 5. A pane refusal raised before any write does not stale other shared-bridge clients.

@pytest.mark.parametrize("code", ["AMBIGUOUS_SPAN", "INVALID_INPUT"])
def test_a_pre_write_refusal_does_not_stale_other_clients(code):
    class Refusing(BrokerPane):
        def request_threadsafe(self, op, payload=None, **kwargs):
            if op == "replace" and payload.get("prefix"):
                self.calls.append((op, payload))
                raise LiveOpFailed(code, "refused before any write")
            return super().request_threadsafe(op, payload, **kwargs)

    pane = Refusing()
    broker = shared.Broker(SimpleNamespace(get=lambda name: pane, list=lambda: [pane], collisions=list))
    a, b = (broker.dispatch({"action": "attach"})["client_id"] for _ in range(2))

    def call(client, op="describe", payload=None):
        return broker.dispatch({"action": "request", "client_id": client, "document_name": pane.document_name,
                                "document_url": pane.document_url, "connected_since": pane.connected_since,
                                "op": op, "payload": payload or {}})

    call(a)
    call(b)
    with pytest.raises(LiveOpFailed):
        call(a, "replace", {"prefix": "x" * 12, "suffix": "y" * 12, "expected_text_sha256": "0" * 64})
    call(b, "replace", {"find": "z", "expected_matches": 1, "replace": "w"})  # not stale
