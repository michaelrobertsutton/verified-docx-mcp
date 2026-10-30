"""Issue #108: partial replacement_segments formats inherit from the first replaced run.

Server side against tests/unit/fake_pane.py (direct live replace_text) and the approved-run
pane double from test_approved_changes.py; the pane's own Office.js is covered by
tests/unit/js/partial_font_harness.mjs.
"""

from __future__ import annotations

import asyncio
import copy
import json
import shutil
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import DEFAULT_FONT, FakeDocument
from test_approved_changes import Pane as ApprovedPane
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.formatted_text import FIELDS, is_partial, resolve_segments, validate_segments
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import write_mode as wm
from verified_docx_mcp.live.session import LiveOpFailed

REPO = Path(__file__).resolve().parents[2]
CAPS = ["formatted_replacements_v1", "partial_font_segments_v1"]
GEORGIA = {**DEFAULT_FONT, "name": "Georgia", "size": 14, "color": "#C00000", "italic": True}


def _code(exc: pytest.ExceptionInfo) -> str:
    return exc.value.envelope.error_code.value


# -- formatted_text ---------------------------------------------------------------------------


def test_partial_formats_validate_only_when_asked():
    segments = [{"text": "Bold", "format": {"bold": True}}, {"text": " rest", "format": {}}]
    out = validate_segments(segments, "Bold rest", partial=True)
    assert out[0]["format"] == {"bold": True} and out[1]["format"] == {}
    assert is_partial(out)
    with pytest.raises(VerifyError):
        validate_segments(segments, "Bold rest")  # full dicts stay mandatory by default
    full = [{"text": "x", "format": dict(DEFAULT_FONT, color="#abcdef")}]
    assert validate_segments(full, "x", partial=True) == validate_segments(full, "x")
    assert not is_partial(validate_segments(full, "x"))


@pytest.mark.parametrize("font", [{"style": "Normal"}, {"bold": 1}, {"underline": "Wavy"}, {"color": "red"},
                                  {"size": 0}, {"name": ""}, {"subscript": True, "superscript": True}])
def test_partial_formats_keep_unknown_key_and_value_rules(font):
    with pytest.raises(VerifyError) as exc:
        validate_segments([{"text": "x", "format": font}], "x", partial=True)
    assert _code(exc) == "INVALID_INPUT"


def test_resolve_fills_missing_fields_from_the_baseline():
    segments = validate_segments([{"text": "A", "format": {"bold": True, "color": "#00ff00"}},
                                  {"text": "b", "format": {}}], "Ab", partial=True)
    out = resolve_segments(segments, GEORGIA, "Ab")
    assert set(out[0]["format"]) == FIELDS
    assert out[0]["format"] == {**GEORGIA, "bold": True, "color": "#00FF00"}
    assert out[1]["format"] == GEORGIA


def test_resolve_refuses_to_inherit_an_unrepresentable_field():
    baseline = {**GEORGIA, "underline": "Dotted", "color": ""}
    with pytest.raises(VerifyError) as exc:
        resolve_segments([{"text": "x", "format": {"bold": True}}], baseline, "x")
    assert _code(exc) == "INVALID_INPUT" and "specify" in exc.value.envelope.message
    # Naming the field explicitly avoids inheriting it.
    out = resolve_segments([{"text": "x", "format": {"underline": "None", "color": "#000000"}}], baseline, "x")
    assert out[0]["format"]["name"] == "Georgia"


# -- direct live replace_text ------------------------------------------------------------------


class PartialLiveTests(LiveWriteBridgeTestCase):
    async def connect(self, text="alpha beta gamma delta", caps=CAPS):
        doc = FakeDocument(text=text)
        doc.fonts.append(("gamma", GEORGIA))
        pane = await self.connect_pane(doc, capabilities=[*FakeDocPaneCaps, *caps])
        sent: list[tuple[str, dict]] = []
        original = pane._dispatch

        def spy(op, payload):
            sent.append((op, dict(payload)))
            return original(op, payload)

        pane._dispatch = spy
        return doc, pane, sent

    async def replace(self, find, replace, segments, **kwargs):
        return await asyncio.to_thread(
            text_edit.execute_replace_text, str(self.target), find, replace, 1,
            write_mode="live", replacement_segments=segments, **kwargs)

    async def test_bold_only_segment_keeps_the_inherited_font(self):
        doc, _, sent = await self.connect()
        result = await self.replace("gamma", "GAMMA", [{"text": "GAMMA", "format": {"bold": True}}])
        self.assertEqual(doc.text, "alpha beta GAMMA delta")
        lookups = [p for op, p in sent if op == "replacement_font"]
        self.assertEqual(len(lookups), 1)
        self.assertEqual(lookups[0]["find"], "gamma")
        write = [p for op, p in sent if op == "replace"][-1]
        self.assertEqual(write["expected_first_font"], GEORGIA)
        self.assertEqual(write["replacement_segments"][0]["format"], {**GEORGIA, "bold": True})
        self.assertTrue(result["applied"])

    async def test_full_segments_need_no_lookup_and_no_new_capability(self):
        doc, _, sent = await self.connect(caps=["formatted_replacements_v1"])
        await self.replace("gamma", "G", [{"text": "G", "format": dict(DEFAULT_FONT, bold=True)}])
        self.assertEqual(doc.text, "alpha beta G delta")
        self.assertFalse([op for op, _ in sent if op == "replacement_font"])
        self.assertNotIn("expected_first_font", [p for op, p in sent if op == "replace"][-1])

    async def test_old_pane_refuses_partial_segments_before_sending_anything(self):
        doc, _, sent = await self.connect(caps=["formatted_replacements_v1"])
        with self.assertRaises(VerifyError) as ctx:
            await self.replace("gamma", "G", [{"text": "G", "format": {"bold": True}}])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)
        self.assertFalse([op for op, _ in sent if op in {"replacement_font", "replace"}])
        self.assertEqual(doc.text, "alpha beta gamma delta")

    async def test_matches_with_different_first_run_fonts_refuse(self):
        doc, _, sent = await self.connect(text="gamma one, gamma two")
        doc.fonts[:] = [("gamma one", GEORGIA)]
        with self.assertRaises(VerifyError) as ctx:
            await asyncio.to_thread(
                text_edit.execute_replace_text, str(self.target), "gamma", "G", 2,
                write_mode="live", replacement_segments=[{"text": "G", "format": {"bold": True}}])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("specify every field", ctx.exception.envelope.message)
        self.assertFalse([op for op, _ in sent if op == "replace"])

    async def test_reformat_between_lookup_and_write_refuses_format_drift(self):
        doc, _, _ = await self.connect()

        def reformat(document):
            document.fonts.append(("gamma", {**GEORGIA, "size": 18}))

        doc.before_replace = reformat
        with self.assertRaises(VerifyError) as ctx:
            await self.replace("gamma", "GAMMA", [{"text": "GAMMA", "format": {"bold": True}}])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.FORMAT_DRIFT)
        self.assertEqual(doc.text, "alpha beta gamma delta")

    async def test_file_mode_still_refuses_segments(self):
        with self.assertRaises(VerifyError) as ctx:
            await asyncio.to_thread(
                text_edit.execute_replace_text, str(self.target), "x", "y", 1, write_mode="file",
                replacement_segments=[{"text": "y", "format": {"bold": True}}])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)


FakeDocPaneCaps = ["row_scope", "comment_loss_guard", "replacement_formatting", "format_readback", "shape_guard",
                   "comment_counts", "comment_loss_ids", "body_hash_v1"]


def test_format_drift_is_a_pre_write_batch_refusal():
    from verified_docx_mcp.errors import _make_error

    exc = _make_error(ErrorCode.FORMAT_DRIFT, "drift")
    assert text_edit._batch_refused_before_write(exc)
    pane = _make_error(ErrorCode.LIVE_OP_FAILED, "drift", {"pane_code": "FORMAT_DRIFT"})
    assert text_edit._batch_refused_before_write(pane)


def test_broker_treats_format_drift_as_pre_mutation():
    from verified_docx_mcp.live import shared

    assert shared.refused_before_mutation(LiveOpFailed("FORMAT_DRIFT", "drift"))
    assert "replacement_font" in shared.READ_OPS
    from verified_docx_mcp.live.protocol import VALID_OPS

    assert "replacement_font" in VALID_OPS


# -- approved runs -----------------------------------------------------------------------------


class FontPane(ApprovedPane):
    def __init__(self):
        super().__init__()
        self.hello = SimpleNamespace(capabilities=[*ApprovedPane.hello.capabilities, *CAPS])
        self.first_font = dict(GEORGIA)
        self.sent: list[tuple[str, dict]] = []
        self.before_replace = None

    def request_threadsafe(self, op, payload=None, **kwargs):
        payload = payload or {}
        if op == "replacement_font":
            self.calls[op] += 1
            self.sent.append((op, dict(payload)))
            return {"font": dict(self.first_font), "match_count": 1}
        if op == "replace":
            self.calls[op] += 1
            self.sent.append((op, dict(payload)))
            if payload["executionSections"] != self.sections:
                raise LiveOpFailed("LIVE_STALE", "Protected section changed; nothing written")
            if self.before_replace:
                self.before_replace(self)
            expected = payload.get("expected_first_font")
            if expected is not None and expected != self.first_font:
                raise LiveOpFailed("FORMAT_DRIFT", "first run changed; nothing written")
            paragraphs = self.sections[0]["paragraphs"]
            index = next(i for i, p in enumerate(paragraphs) if payload["find"] in p)
            paragraphs[index] = paragraphs[index].replace(payload["find"], payload["replace"], 1)
            return {"applied": True, "pre": "a" * 64, "post": runs.scope_sha(self.sections),
                    "matches": [{"before": payload["find"], "after": payload["replace"],
                                 "segments": payload.get("replacement_segments")}],
                    "shapes_before": 0, "shapes_after": 0}
        return super().request_threadsafe(op, payload, **kwargs)


@pytest.fixture
def font_run(tmp_path, monkeypatch):
    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    pane = FontPane()
    monkeypatch.setattr(wm, "live_session_for", lambda path: runs.scoped_session(pane))
    return str(tmp_path / pane.document_name), pane


def _text_op(find="Proof is current.", replace="Proof: current.", segments=None):
    segments = segments or [{"text": "Proof:", "format": {"bold": True}}, {"text": " current.", "format": {}}]
    return {"op": "replace_text", "section": "target",
            "args": {"find": find, "replace": replace, "expected_matches": 1, "replacement_segments": segments}}


def test_prepare_resolves_partial_segments_and_binds_the_baseline(font_run):
    path, pane = font_run
    prepared = runs.prepare(path, [_text_op()])
    manifest = json.loads(Path(prepared["manifest_path"]).read_text())
    operation = manifest["operations"][0]
    assert operation["font_baseline"] == GEORGIA
    resolved = operation["args"]["replacement_segments"]
    assert resolved[0]["format"] == {**GEORGIA, "bold": True}
    assert resolved[1]["format"] == GEORGIA
    assert manifest["formatting_policy"] == "explicit-segments-v1+first-run-baseline-v1"
    assert prepared["preview"][0]["replacement_segments"] == resolved
    assert prepared["preview"][0]["font_baseline"] == GEORGIA
    lookup = next(p for op, p in pane.sent if op == "replacement_font")
    assert lookup["executionTargetSection"] == "target" and lookup["find"] == "Proof is current."
    assert runs.sha(manifest) == prepared["manifest_sha256"]
    for mutate in (lambda m: m["operations"][0]["font_baseline"].update(size=12),
                   lambda m: m["operations"][0]["args"]["replacement_segments"][1]["format"].update(bold=True)):
        changed = copy.deepcopy(manifest)
        mutate(changed)
        assert runs.sha(changed) != prepared["manifest_sha256"]


def test_execute_sends_the_baseline_and_completes(font_run):
    path, pane = font_run
    prepared = runs.prepare(path, [_text_op()])
    report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")
    assert report["status"] == "complete", report
    write = [p for op, p in pane.sent if op == "replace"][-1]
    assert write["expected_first_font"] == GEORGIA
    assert write["replacement_segments"][0]["format"] == {**GEORGIA, "bold": True}
    assert pane.sections[0]["paragraphs"][1] == "Proof: current."


def test_reformat_after_prepare_is_a_format_drift_conflict_and_nothing_is_written(font_run):
    path, pane = font_run
    prepared = runs.prepare(path, [_text_op()])
    pane.first_font = {**GEORGIA, "size": 18}  # a manual reformat between prepare and execute
    report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")
    assert report["status"] != "complete"
    assert report["issues"][0]["status"] == "conflict"
    assert report["issues"][0]["code"] == "FORMAT_DRIFT"
    assert pane.sections[0]["paragraphs"][1] == "Proof is current."


def test_partial_segments_after_an_earlier_rewrite_of_the_same_text_refuse(font_run):
    path, pane = font_run
    first = {"op": "replace_text", "section": "target",
             "args": {"find": "Proof is", "replace": "Evidence is", "expected_matches": 1}}
    second = _text_op(find="Evidence is current.", replace="Evidence: current.",
                      segments=[{"text": "Evidence:", "format": {"bold": True}}, {"text": " current.", "format": {}}])
    with pytest.raises(VerifyError) as exc:
        runs.prepare(path, [first, second])
    assert _code(exc) == "INVALID_INPUT" and "explicitly" in exc.value.envelope.message
    assert not [op for op, _ in pane.sent if op == "replacement_font"]
    # Full fields need no inheritance, so the same run prepares.
    full = copy.deepcopy(second)
    for segment in full["args"]["replacement_segments"]:
        segment["format"] = {**DEFAULT_FONT, **segment["format"]}
    runs.prepare(path, [first, full])


def test_old_pane_refuses_partial_segments_at_prepare(font_run):
    path, pane = font_run
    pane.hello = SimpleNamespace(capabilities=[c for c in pane.hello.capabilities if c != "partial_font_segments_v1"])
    with pytest.raises(VerifyError) as exc:
        runs.prepare(path, [_text_op()])
    assert _code(exc) == "LIVE_CAPABILITY_MISSING"
    full = _text_op(segments=[{"text": "Proof:", "format": dict(DEFAULT_FONT, bold=True)},
                              {"text": " current.", "format": dict(DEFAULT_FONT)}])
    prepared = runs.prepare(path, [full])
    manifest = json.loads(Path(prepared["manifest_path"]).read_text())
    assert "font_baseline" not in manifest["operations"][0]
    assert manifest["formatting_policy"] == "explicit-segments-v1"


def test_validate_alone_still_requires_full_fields():
    with pytest.raises(VerifyError):
        runs.validate([_text_op()])
    runs.validate([_text_op()], partial=True)
    caller_baseline = {**_text_op(), "font_baseline": GEORGIA}
    with pytest.raises(VerifyError):
        runs.validate([caller_baseline], partial=True)  # only prepare may bind a baseline


# -- pane (Office.js) --------------------------------------------------------------------------


@pytest.mark.skipif(shutil.which("node") is None, reason="node not found on PATH")
def test_pane_partial_font_harness():
    proc = subprocess.run(["node", str(REPO / "tests/unit/js/partial_font_harness.mjs"), str(REPO / "addin/taskpane.js")],
                          capture_output=True, text=True, timeout=60, check=False)
    assert proc.returncode == 0, proc.stderr + proc.stdout
