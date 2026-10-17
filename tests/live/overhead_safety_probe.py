"""Manual #101/#102 safety probe, after optimized overhead_probe on its disposable copy."""

import json
import subprocess
import time
from pathlib import Path

from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge
from verified_docx_mcp.live import write_mode as wm

bridge.start_in_background()
path = "/private/tmp/word-101-optimized-1.docx"
pane = wm.live_session_for(path)


def cell(r, c, text):
    return {
        "op": "replace_cell_markdown",
        "section": "target",
        "args": {"table_id": 1, "row_index": r, "cell_index": c, "markdown": text},
    }


save = {"op": "live_save", "section": "", "args": {}}


def execute(ops, tracked=False):
    p = runs.prepare(path, ops, track_changes=tracked)
    r = runs.execute(
        path, p["run_id"], p["manifest_sha256"], "User-approved disposable safety matrix"
    )
    print(json.dumps(r), flush=True)
    return r


results = {}
results["cells"] = execute(
    [
        cell(1, 1, ""),
        cell(1, 2, "**Top**\n\nBottom"),
        cell(1, 3, "Repeated"),
        cell(2, 1, "Repeated"),
        save,
    ]
)
assert results["cells"]["status"] == "complete"


# Genuine native edits bypass the broker's revision counters, idle and busy.
def human(r, c, text):
    subprocess.run(
        [
            "osascript",
            "-e",
            f'tell application "Microsoft Word" to set content of text object of cell {c} of row {r} of table 1 of document "word-101-optimized-1.docx" to "{text}"',
        ],
        check=True,
        capture_output=True,
        text=True,
    )


ops = [
    cell(1, 1, "Must not overwrite idle edit"),
    cell(1, 2, "Independent one"),
    cell(1, 3, "Must not overwrite busy edit"),
    cell(2, 1, "Independent two"),
    save,
]
p = runs.prepare(path, ops)
human(1, 1, "Idle human retained")
time.sleep(6)
original = runs.apply
calls = 0


def interleaved(*args, **kwargs):
    global calls
    result = original(*args, **kwargs)
    if args[1]["op"] != "live_save":
        calls += 1
        if calls == 1:
            human(1, 3, "Busy human retained")
    return result


runs.apply = interleaved
try:
    results["human"] = runs.execute(
        path, p["run_id"], p["manifest_sha256"], "User-approved disposable human edit preservation"
    )
finally:
    runs.apply = original
print(json.dumps(results["human"]), flush=True)
assert results["human"]["status"] == "partial" and results["human"]["remaining"] == [0, 2], results[
    "human"
]
assert results["human"]["applied_count"] == 3
for r, c, text in [
    (1, 1, "Idle human retained"),
    (1, 3, "Busy human retained"),
    (1, 2, "Independent one"),
    (2, 1, "Independent two"),
]:
    assert (
        pane.request_threadsafe(
            "cell_get",
            {"executionTargetSection": "target", "table_index": 1, "row_index": r, "cell_index": c},
        )["text"].strip()
        == text
    )
results["tracked"] = execute([cell(2, 2, "Tracked candidate"), save], True)
assert results["tracked"]["status"] == "complete"
results["revision_guard"] = execute(
    [cell(2, 2, "Do not overwrite revision"), cell(2, 3, "Independent of revision"), save]
)
assert (
    results["revision_guard"]["status"] == "partial"
    and results["revision_guard"]["issues"][0]["code"] == "TRACKED_CHANGES_PRESENT"
), results["revision_guard"]
sections = pane.request_threadsafe("execution_snapshot", {"section_slugs": ["target"]})["sections"]
pane.request_threadsafe(
    "comment_add",
    {
        "find": "Accepted 3:1",
        "text": "Preserve this cell comment",
        "expected_matches": 1,
        "scope": "body",
        "executionSections": sections,
        "executionTargetSection": "target",
    },
)
results["comment_guard"] = execute(
    [cell(3, 1, "Do not remove comment"), cell(3, 2, "Independent of comment"), save]
)
assert (
    results["comment_guard"]["status"] == "partial"
    and results["comment_guard"]["issues"][0]["code"] == "WOULD_DELETE_COMMENTS"
), results["comment_guard"]
Path("/private/tmp/word-101-safety-evidence.json").write_text(json.dumps(results, indent=2))
print("Native safety matrix passed", flush=True)
