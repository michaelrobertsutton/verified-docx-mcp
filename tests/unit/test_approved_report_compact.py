"""Issue #112: approved-run reports are compact and the prepare preview is reviewable."""

import json
from pathlib import Path

from test_approved_changes import Pane, cell, execute, setup  # noqa: F401

from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live.session import LiveOpFailed


def tokens(value) -> int:
    return len(json.dumps(value, separators=(",", ":"))) // 4


def conflicted_run(setup, count=12):  # noqa: F811
    """Every cell is refused before write, and a co-author rewrites a long paragraph each time,
    so every failure carries a large section_diff."""
    path, pane = setup
    prepared = runs.prepare(path, [cell(r, 1, f"new {r}") for r in range(1, count + 1)])
    original = pane.request_threadsafe
    state = {"n": 0}

    def request(op, payload=None, **kw):
        if op == "cell_set":
            state["n"] += 1
            pane.sections[0]["paragraphs"][1] = f"{state['n']} " + "co-author text " * 200
            raise LiveOpFailed("WOULD_DELETE_COMMENTS", "target cannot be safely replaced")
        return original(op, payload, **kw)

    pane.request_threadsafe = request
    return path, pane, prepared


def test_report_does_not_echo_the_callers_own_arguments(setup):  # noqa: F811
    path, _pane = setup
    prepared = runs.prepare(path, [cell(1), cell(2)])
    report = execute(path, prepared)
    assert report["status"] == "complete"
    assert "run_id" not in report
    assert "manifest_sha256" not in report
    assert "performance" not in report
    assert report["remaining"] == []
    assert report["remaining_ranges"] == ""


def test_verbose_brings_performance_back(setup):  # noqa: F811
    path, _pane = setup
    prepared = runs.prepare(path, [cell(1)])
    report = runs.execute(path, prepared["run_id"], prepared["manifest_sha256"], "Lead", verbose=True)
    assert set(report["performance"]) == {"operation_elapsed_ms", "office_sync_count", "bridge_requests"}
    assert "run_id" not in report  # verbose never restores the echoes


def test_twelve_prewrite_conflicts_stay_under_three_thousand_tokens(setup):  # noqa: F811
    path, _pane, prepared = conflicted_run(setup)
    report = execute(path, prepared)
    assert report["status"] in {"partial", "unknown"}
    assert len(report["issues"]) == 12
    assert tokens(report) < 3000, tokens(report)
    print(f"\n12 prewrite conflicts: {tokens(report)} tokens (4000-character section_diff each before the cap)")
    # remaining stays a list of integers (JennyStack's guard and the resume flow read it)...
    assert report["remaining"] == list(range(12))
    assert all(type(i) is int for i in report["remaining"])
    # ...and the same set is also a compact string.
    assert report["remaining_ranges"] == "0-11"


def test_section_diff_is_cut_and_the_journal_keeps_the_rest(setup):  # noqa: F811
    path, _pane, prepared = conflicted_run(setup, count=3)
    report = execute(path, prepared)
    cut = [i for i in report["issues"] if i.get("section_diff_truncated")]
    assert cut, report["issues"]
    for issue in cut:
        assert len(issue["section_diff"]) == 600
    journal = json.loads(Path(report["journal_path"]).read_text())
    kept = max(len(v.get("section_diff", "")) for v in journal["deferred"].values())
    assert kept > 600  # the journal still holds the full (4000-character) diff


def test_remaining_ranges_compacts_runs():
    assert runs._ranges([]) == ""
    assert runs._ranges([4]) == "4"
    assert runs._ranges([7, 8, 9, 12, 14, 15]) == "7-9,12,14-15"


def test_prepare_preview_shows_what_will_change_truncated_to_200(setup):  # noqa: F811
    path, _pane = setup
    long_text = "word " * 100
    ops = [cell(r, c, f"**{r}:{c}** {long_text}") for r in range(1, 13) for c in range(1, 4)]
    prepared = runs.prepare(path, ops)
    preview = prepared["preview"]
    assert len(preview) == 36
    first = preview[0]
    assert (first["table_id"], first["row_index"], first["cell_index"]) == (1, 1, 1)
    assert first["markdown"].startswith("**1:1** word")
    assert first["markdown"].endswith("…") and len(first["markdown"]) == 201
    # The manifest on disk, not the preview, is the authority: it holds the whole text.
    manifest = json.loads(Path(prepared["manifest_path"]).read_text())
    assert manifest["operations"][0]["args"]["markdown"].endswith(long_text)
    assert len(manifest["operations"]) == 36


def test_preview_covers_text_rows_and_deletes(setup):  # noqa: F811
    path, _pane = setup
    ops = [
        {"op": "replace_text", "section": "target",
         "args": {"find": "Proof is current.", "replace": "Proof is updated.", "expected_matches": 1}},
        {"op": "delete_paragraph", "section": "target", "args": {"anchor_text": "Delete the whole paragraph."}},
        {"op": "replace_table_row", "section": "target",
         "args": {"table_id": 1, "row_index": 2, "cells": ["a", "b" * 300, "c", "d"]}},
    ]
    preview = runs.prepare(path, ops)["preview"]
    assert preview[0] == {"op": "replace_text", "section": "target", "find": "Proof is current.",
                          "replace": "Proof is updated.", "expected_matches": 1}
    assert preview[1]["anchor_text"] == "Delete the whole paragraph."
    assert preview[2]["cells"][0] == "a"
    assert len(preview[2]["cells"][1]) == 201
    assert preview[2]["row_index"] == 2
