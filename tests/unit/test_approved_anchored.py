"""Issue #113 part A: anchored find and para_ref replace_text operations in approved runs."""

import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from test_approved_changes import Pane

from verified_docx_mcp import locate
from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import write_mode as wm
from verified_docx_mcp.live.session import LiveOpFailed

E = locate.ELLIPSIS
PARA = "Skyward will operate the help desk " + "with tiered escalation and measured response times, " * 11 + "reported monthly to the program office."
assert len(PARA) == 646, len(PARA)
FIND = f"Skyward will operate the help desk{E}reported monthly to the program office."


class ReplacePane(Pane):
    """Pane double whose replace honours both the plain and the anchored payload."""

    def __init__(self, capabilities=("approved_changes_v1", "anchored_find_v1")):
        super().__init__()
        self.hello = SimpleNamespace(capabilities=list(capabilities))
        self.sections[0]["paragraphs"].insert(2, PARA)
        self.replace_payloads = []

    def request_threadsafe(self, op, payload=None, **kwargs):
        if op != "replace":
            return super().request_threadsafe(op, payload, **kwargs)
        self.calls[op] += 1
        self.replace_payloads.append(dict(payload))
        if payload["executionSections"] != self.sections:
            raise LiveOpFailed("LIVE_STALE", "Protected section changed; nothing written")
        section = next(s for s in self.sections if s["slug"] == payload["executionTargetSection"])
        paragraphs = section["paragraphs"]
        if "prefix" in payload:
            assert "find" not in payload
            hits = []
            for i, text in enumerate(paragraphs):
                start = text.find(payload["prefix"])
                end = text.find(payload["suffix"], start + len(payload["prefix"])) if start >= 0 else -1
                if end >= 0 and locate.span_sha256(text[start:end + len(payload["suffix"])]) == payload["expected_text_sha256"]:
                    hits.append((i, text[start:end + len(payload["suffix"])]))
            if len(hits) != 1:
                raise LiveOpFailed("zero_match" if not hits else "AMBIGUOUS_SPAN", "no unique span")
            index, found = hits[0]
        else:
            matches = [i for i, text in enumerate(paragraphs) if payload["find"] in text]
            if len(matches) != 1:
                raise LiveOpFailed("zero_match", "found 0")
            index, found = matches[0], payload["find"]
        paragraphs[index] = paragraphs[index].replace(found, payload["replace"], 1)
        return {"applied": True, "matches": [{"before": found, "after": payload["replace"]}],
                "pre": "a" * 64, "post": runs.scope_sha(self.sections)}


@pytest.fixture
def setup(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path / "state"))
    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    pane = ReplacePane()
    monkeypatch.setattr(wm, "live_session_for", lambda path: runs.scoped_session(pane))
    return str(tmp_path / pane.document_name), pane


def op(**args):
    return {"op": "replace_text", "section": "target", "args": args}


def test_a_600_character_ellipsis_replacement_prepares_previews_and_executes(setup):
    path, pane = setup
    prepared = runs.prepare(path, [op(find=FIND, replace="Skyward runs the help desk.", expected_length=646)])
    preview = prepared["preview"][0]
    assert preview["find"] == PARA[:200] + E and preview["find_chars"] == 646
    assert preview["find_sha256"] == locate.span_sha256(PARA)
    args = json.loads(Path(prepared["manifest_path"]).read_text())["operations"][0]["args"]
    assert args["find"] == PARA
    assert args["prefix"] == "Skyward will operate the help desk"
    assert args["suffix"] == "reported monthly to the program office."
    assert args["expected_text_sha256"] == locate.span_sha256(PARA)
    report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")
    assert report["status"] == "complete", report
    assert pane.sections[0]["paragraphs"][2] == "Skyward runs the help desk."
    assert "find" not in pane.replace_payloads[-1] and pane.replace_payloads[-1]["prefix"] == args["prefix"]


def test_an_operation_on_an_earlier_operations_output_resolves(setup):
    path, pane = setup
    rewritten = "Proof is updated for the anchored follow-up test."
    prepared = runs.prepare(path, [
        op(find="Proof is current.", replace=rewritten, expected_matches=1),
        op(find=f"Proof is updated{E}anchored follow-up test.", replace="Proof is final.", expected_length=len(rewritten)),
    ])
    args = json.loads(Path(prepared["manifest_path"]).read_text())["operations"][1]["args"]
    assert args["find"] == rewritten
    report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")
    assert report["status"] == "complete", report
    assert pane.sections[0]["paragraphs"][1] == "Proof is final."
    assert pane.replace_payloads[-1]["find"] == rewritten  # <= 230: the plain path


def test_para_ref_counts_within_the_section(setup):
    path, pane = setup
    prepared = runs.prepare(path, [op(para_ref="p1", expected_sha256=locate.span_sha256("Proof is current."),
                                      replace="Proof replaced by para_ref.")])
    assert runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead")["status"] == "complete"
    assert pane.sections[0]["paragraphs"][1] == "Proof replaced by para_ref."


@pytest.mark.parametrize(("operation", "code"), [
    (op(para_ref="p1", expected_sha256=locate.span_sha256("stale"), replace="x"), "PARA_HASH_MISMATCH"),
    (op(para_ref="p0", expected_sha256=locate.span_sha256("Target"), replace="x"), "INVALID_INPUT"),
    (op(find=FIND, replace="x", expected_length=600), "ZERO_MATCH"),
    (op(find=FIND, replace="x"), "INVALID_INPUT"),
    (op(find=FIND, replace="x", expected_length=646, expected_matches=2), "INVALID_INPUT"),
    (op(find="Proof is current.", replace="x", expected_matches=1, expected_length=17), "INVALID_INPUT"),
])
def test_unresolvable_targets_are_refused_at_prepare(setup, operation, code):
    path, _pane = setup
    with pytest.raises(VerifyError) as ctx:
        runs.prepare(path, [operation])
    assert ctx.value.envelope.error_code.value == code


def test_two_matching_spans_are_ambiguous_at_prepare(setup):
    path, pane = setup
    pane.sections[0]["paragraphs"].append(PARA)
    with pytest.raises(VerifyError) as ctx:
        runs.prepare(path, [op(find=FIND, replace="x", expected_length=646)])
    assert ctx.value.envelope.error_code.value == "AMBIGUOUS_SPAN"
    assert ctx.value.envelope.diagnostics["total"] == 2


def test_an_old_pane_is_capability_missing_at_prepare(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path / "state"))
    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    pane = ReplacePane(capabilities=("approved_changes_v1",))
    monkeypatch.setattr(wm, "live_session_for", lambda path: runs.scoped_session(pane))
    with pytest.raises(VerifyError) as ctx:
        runs.prepare(str(tmp_path / pane.document_name), [op(find=FIND, replace="x", expected_length=646)])
    assert ctx.value.envelope.error_code.value == "LIVE_CAPABILITY_MISSING"
