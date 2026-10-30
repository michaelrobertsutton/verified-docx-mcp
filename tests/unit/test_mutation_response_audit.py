"""Issue #108: every mutating tool's compact response keeps the essentials, and the audit
record is written from the full evidence before anything is compacted.

Each file-mode-capable tool in ``middleware.MUTATING_TOOLS`` runs through
``fastmcp.Client(server.mcp)`` -- so the lean-envelope and evidence-enforcement middleware
both apply -- on a copy of a Word-authored fixture. Live-only tools are listed in
``LIVE_ONLY`` with where their compact envelopes are pinned; the batch's partial/unknown
outcomes are driven here through a fake pane.
"""

from __future__ import annotations

import asyncio
import json
import shutil
import sys
from pathlib import Path
from typing import Any

import pytest
from fastmcp import Client

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, mutations, server
from verified_docx_mcp.middleware import MUTATING_TOOLS

REPO = Path(__file__).resolve().parents[2]
FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02

LIVE_ONLY = {
    "replace_text_batch": "live-only; driven through a fake pane below and in test_live_replace_batch.py",
    "live_save": "live-only; needs a pane (test_live_write_mode.py)",
    "live_changes_execute": "approved runs need a pane (test_approved_report_compact.py)",
    "live_changes_resume": "approved runs need a pane (test_approved_report_compact.py)",
    "delete_paragraph": "live-only; needs a pane (test_live_delete_evidence.py)",
    "clear_numbering": "live-only; needs a pane (test_live_table_cleanup.py)",
    "delete_table": "live-only; needs a pane (test_live_table_cleanup.py)",
}

PARA = "The quick brown fox jumps over the lazy dog."
COMMENT = "2A21A9D2"

# tool -> (fixture, setup calls run first, arguments, keys the compact response must carry
# besides `applied`). Every file-mode write also returns revision_after.
CASES: dict[str, tuple[str, list[tuple[str, dict[str, Any]]], dict[str, Any], set[str]]] = {
    "replace_text": ("frag.docx", [], {"find": "quick", "replace": "slow", "expected_matches": 1},
                     {"match_count", "before", "after"}),
    "format_text": ("frag.docx", [], {"find": "lazy", "style": {"bold": True}, "expected_matches": 1},
                    {"match_count", "format_delta"}),
    "apply_style": ("sections.docx", [], {"find": "Overview", "style_id": "Heading2", "expected_matches": 1},
                    {"match_count", "style_id"}),
    "append_markdown": ("frag.docx", [], {"markdown": "More text."}, {"after_chars", "evidence_ref"}),
    "replace_body_markdown": ("frag.docx", [], {"markdown": "# New\n\nBody."}, {"after_chars", "evidence_ref"}),
    "replace_range_markdown": ("sections.docx", [],
                               {"section_key": "background-1", "markdown": "## Background\n\nNew."},
                               {"after_chars", "evidence_ref"}),
    "accept_tracked_changes": ("revision/tracked.docx", [], {}, {"match_count", "revision_ids"}),
    "reject_tracked_changes": ("revision/tracked.docx", [], {}, {"match_count", "revision_ids"}),
    "add_anchored_comment": ("frag.docx", [], {"quote": "lazy dog", "text": "Why?", "expected_matches": 1},
                             {"match_count", "comment_id"}),
    "reply_to_comment": ("revision/commented.docx", [], {"comment_id": COMMENT, "text": "Answer."},
                         {"comment_id", "parent_comment_id"}),
    "resolve_comment": ("revision/commented.docx", [], {"comment_id": COMMENT}, {"comment_id", "after"}),
    "delete_comment": ("revision/commented.docx", [], {"comment_id": COMMENT},
                       {"deleted_thread_count", "deleted_comment_ids", "comments_after"}),
    "delete_resolved_comments": ("revision/commented.docx", [("resolve_comment", {"comment_id": COMMENT})],
                                 {"comment_ids": [COMMENT]}, {"deleted_comment_ids", "deleted_thread_count"}),
    "replace_table_row": ("tables.docx", [], {"table_id": 1, "row_index": 1, "cells": ["a", "b"]},
                          {"match_count", "after"}),
    "replace_cell_markdown": ("tables.docx", [], {"table_id": 1, "row_index": 1, "cell_index": 1, "markdown": "x"},
                              {"match_count", "after"}),
    "insert_table": ("tables.docx", [], {"rows": [["h1", "h2"], ["a", "b"]], "style_id": "TableNormal"},
                     {"table_id", "match_count"}),
    "insert_table_row": ("tables.docx", [], {"table_id": 1, "cells": ["c", "d"], "after_row_index": 2},
                         {"table_id", "row_index", "rows_after"}),
    "delete_table_row": ("tables.docx", [], {"table_id": 1, "row_index": 2, "expected_text": "R2C1\tR2C2"},
                         {"table_id", "row_index", "rows_after"}),
    "insert_image": ("frag.docx", [], {"image_path": "sample.png"}, {"match_count", "blip_rid"}),
    "insert_paragraphs": ("frag.docx", [], {"anchor": "lazy dog", "paragraphs": [{"text": "Inserted."}],
                                            "expected_matches": 1}, {"inserted"}),
    "merge_paragraph_with_next": (
        "tables.docx",
        [("replace_cell_markdown", {"table_id": 1, "row_index": 1, "cell_index": 1, "markdown": "one\n\ntwo"})],
        {"table_id": 1, "row_index": 1, "cell_index": 1, "paragraph_index": 1}, {"match_count", "after"}),
}


def _call(tool: str, args: dict[str, Any]) -> tuple[dict[str, Any] | None, bool]:
    async def run():
        async with Client(server.mcp) as client:
            result = await client.call_tool(tool, args, raise_on_error=False)
            if result.structured_content is None:  # a typed error arrives as JSON text
                text = "".join(getattr(block, "text", "") for block in result.content)
                return json.loads(text), result.is_error
            return result.structured_content, result.is_error

    return asyncio.run(run())


def _records(tool: str | None = None) -> list[dict[str, Any]]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    records = [json.loads(line) for line in path.read_text().splitlines() if line.strip()]
    return [r for r in records if tool is None or r["tool"] == tool]


@pytest.fixture
def workdir(tmp_path, monkeypatch):
    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    shutil.copyfile(FIXTURES / "images" / "sample.png", tmp_path / "sample.png")
    return tmp_path


def _run_case(workdir: Path, tool: str) -> tuple[dict[str, Any], set[str]]:
    fixture, setup, args, essentials = CASES[tool]
    target = workdir / Path(fixture).name
    shutil.copyfile(FIXTURES / fixture, target)
    for setup_tool, setup_args in setup:
        out, is_error = _call(setup_tool, {"path": str(target), **setup_args})
        assert not is_error and out["applied"], (setup_tool, out)
    args = json.loads(json.dumps(args))
    if "image_path" in args:
        args["image_path"] = str(workdir / args["image_path"])
    out, is_error = _call(tool, {"path": str(target), **args})
    assert not is_error, (tool, out)
    return out, essentials


def test_every_mutating_tool_is_covered_or_explicitly_skipped():
    assert set(CASES) | set(LIVE_ONLY) == set(MUTATING_TOOLS)
    assert not set(CASES) & set(LIVE_ONLY)


@pytest.mark.parametrize("tool", sorted(CASES))
def test_compact_response_keeps_the_essentials(workdir, tool):
    before = len(_records())
    out, essentials = _run_case(workdir, tool)
    assert out["applied"] is True
    missing = ({"revision_after"} | essentials) - set(out)
    assert not missing, (tool, missing, out)
    assert out["revision_after"]
    # Uncertainty and loss are never compacted away: whatever the full evidence recorded as a
    # warning, an orphaned comment, a conflict copy or a failed log write is in the response.
    records = _records()[before:]
    assert records, f"{tool} wrote no audit record"
    evidence = records[-1]["evidence"]
    for key in ("warnings", "orphaned_comment_ids", "conflict_copies", "comments_removed"):
        if evidence.get(key):
            assert out.get(key) == evidence[key], (tool, key)
    for flag in ("audit_logged", "ledger_logged"):
        assert out.get(flag, True) is True  # absent means the constant True was dropped
    if "revision_after" in evidence:
        assert out["revision_after"] == evidence["revision_after"]


def test_skipped_tools_are_live_only():
    for tool in LIVE_ONLY:
        assert tool in MUTATING_TOOLS, tool


# -- the audit record ---------------------------------------------------------------------------


def test_success_audit_record_holds_the_full_evidence_the_response_compacts(workdir):
    target = workdir / "frag.docx"
    shutil.copyfile(FIXTURES / "frag.docx", target)
    out, is_error = _call("append_markdown", {"path": str(target), "markdown": "More text."})
    assert not is_error and out["applied"]
    record = audit.find_audit_record(out["evidence_ref"])
    assert record is not None and record["tool"] == "append_markdown"
    evidence = record["evidence"]
    # The whole before/after text is in the audit record only; the response has bounded fields.
    assert evidence["after"].endswith("More text.")
    assert "after" not in out and "before" not in out
    assert out["after_chars"] == len(evidence["after"])
    assert evidence["revision_after"] == out["revision_after"]
    # Constants the lean envelope drops were recorded before it ran.
    for key in ("write_mode", "revision_before"):
        assert key in evidence

    out, _ = _call("replace_text", {"path": str(target), "find": "quick", "replace": "slow",
                                    "expected_matches": 1, "write_mode": "file"})
    evidence = _records("replace_text")[-1]["evidence"]
    assert evidence["write_mode"] == "file" and "write_mode" not in out
    assert "conflict_copy_detected" in evidence and "conflict_copy_detected" not in out
    assert evidence["match_count"] == out["match_count"]


def test_refused_write_is_a_typed_error_and_leaves_no_applied_record(workdir):
    target = workdir / "frag.docx"
    shutil.copyfile(FIXTURES / "frag.docx", target)
    before = len(_records())
    out, is_error = _call("replace_text", {"path": str(target), "find": "absent", "replace": "x",
                                           "expected_matches": 1, "write_mode": "file"})
    assert is_error and out["error_code"] == "ZERO_MATCH"
    assert _records()[before:] == []  # nothing was written, so nothing is claimed


def test_audit_failure_is_reported_in_the_compact_response(workdir, monkeypatch):
    target = workdir / "frag.docx"
    shutil.copyfile(FIXTURES / "frag.docx", target)
    blocker = workdir / "not-a-dir"
    blocker.write_text("x")
    monkeypatch.setattr(audit, "_state_dir", lambda: blocker / "state")  # mkdir fails: audit write fails
    for tool, args in [("replace_text", {"find": "quick", "replace": "slow", "expected_matches": 1}),
                       ("append_markdown", {"markdown": "More."})]:
        out, is_error = _call(tool, {"path": str(target), "write_mode": "file", **args} if tool == "replace_text"
                              else {"path": str(target), **args})
        assert not is_error and out["applied"] is True
        assert out["audit_logged"] is False, (tool, out)
        assert "evidence_ref" not in out  # never a ref to a record that does not exist


# -- live: partial / unknown batches and a write that landed but failed verification -------------


class LiveAuditTests(LiveWriteBridgeTestCase):
    async def call(self, tool, args):
        return await asyncio.to_thread(_call, tool, {"path": str(self.target), **args})

    async def test_partial_batch_keeps_status_counts_codes_and_revision(self):
        doc = FakeDocument(text="alpha beta gamma")
        await self.connect_pane(doc)
        rev = f"live:sha256:{doc.sha256()}"
        out, is_error = await self.call("replace_text_batch", {
            "revision_before": rev, "edits": [{"find": "alpha", "replace": "A", "expected_matches": 1},
                                              {"find": "missing", "replace": "X", "expected_matches": 1}]})
        self.assertFalse(is_error)
        self.assertFalse(out["applied"])
        self.assertEqual(out["batch_status"], "partial")
        self.assertEqual((out["applied_count"], out["edit_count"]), (1, 2))
        self.assertEqual(out["revision_before"], rev)
        self.assertEqual(out["revision_after"], f"live:sha256:{doc.sha256()}")
        self.assertEqual([r["status"] for r in out["results"]], ["applied", "refused_before_write"])
        self.assertEqual(out["results"][1]["code"], "ZERO_MATCH")
        record = _records("replace_text_batch")[-1]["evidence"]
        self.assertEqual(record["batch_status"], "partial")
        self.assertEqual([r["status"] for r in record["results"]], ["applied", "refused_before_write"])

    async def test_unknown_batch_says_so_and_has_no_revision_after(self):
        doc = FakeDocument(text="alpha beta gamma")
        pane = await self.connect_pane(doc)
        pane.disconnect_on_replace = 1
        out, is_error = await self.call("replace_text_batch", {
            "revision_before": f"live:sha256:{doc.sha256()}",
            "edits": [{"find": "alpha", "replace": "A", "expected_matches": 1}]})
        self.assertFalse(is_error)
        self.assertEqual(out["batch_status"], "unknown")
        self.assertFalse(out["applied"])
        self.assertIn("revision_after", out)
        self.assertIsNone(out["revision_after"])
        self.assertEqual(out["results"][0]["status"], "unknown")
        self.assertEqual(_records("replace_text_batch")[-1]["evidence"]["batch_status"], "unknown")

    async def test_landed_but_unverified_write_is_audited_and_reported(self):
        doc = FakeDocument(text="alpha beta gamma")
        await self.connect_pane(doc)
        doc.report_pre = "0" * 64  # the pane applied the op against a body that had moved
        out, is_error = await self.call("replace_text", {"find": "beta", "replace": "B", "expected_matches": 1,
                                                         "write_mode": "live"})
        self.assertTrue(is_error)
        self.assertEqual(out["error_code"], "LIVE_STALE_AFTER_WRITE")
        self.assertIn("ALREADY LANDED", out["message"])
        record = _records("replace_text:verification_failed")[-1]["evidence"]
        self.assertTrue(record["applied"])
        self.assertFalse(record["verified"])
        self.assertTrue(record["revision_before"].startswith("live:sha256:"))
        self.assertIn("verification_failure", record)

