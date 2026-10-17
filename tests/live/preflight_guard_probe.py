"""Real #101 boundary refusals on the disposable three-pane benchmark copy."""

import copy
import json
from pathlib import Path

from verified_docx_mcp.live import approved_changes as runs
from verified_docx_mcp.live import bridge, shared
from verified_docx_mcp.live.session import LiveOpFailed

registry = bridge.start_in_background()
path = "/private/tmp/word-101-optimized-three-1.docx"
pane = registry.get(Path(path).name)
assert pane is not None and "approved_cell_stream_v1" in pane.hello.capabilities
sections = pane.request_threadsafe("execution_snapshot", {"section_slugs": ["target"]})["sections"]
before = sections[0]["tables"][0]["rows"][0][0]
payload = {
    "executionSections": sections,
    "executionTargetSection": "target",
    "table_index": 1,
    "row_index": 1,
    "cell_index": 1,
    "expected_before_text": before,
    "expected_before_paragraph_count": 1,
    "paragraphs": [[{"text": "Must never apply"}]],
}
results = {}
changed = copy.deepcopy(payload)
changed["executionSections"][0]["tables"][0]["rows"].append(["Changed structure"] * 4)
changed["executionSections"][0]["tables"][0]["paragraph_counts"].append([1] * 4)
missing = {**payload, "executionTargetSection": "missing-section"}
for label, candidate, code in [
    ("table_structure", changed, "LIVE_STALE"),
    ("section_identity", missing, "LOCK_SCOPE_UNRESOLVED"),
]:
    try:
        pane.request_threadsafe("approved_cell_set", candidate)
        raise AssertionError("Unsafe preflight unexpectedly wrote")
    except LiveOpFailed as exc:
        assert exc.code == code, exc
        results[label] = code
    assert (
        pane.request_threadsafe(
            "cell_get",
            {"executionTargetSection": "target", "table_index": 1, "row_index": 1, "cell_index": 1},
        )["text"].strip()
        == before
    )
# The fixture's existing outside paragraph is entirely comment-anchored. Add a
# separate plain paragraph so the independent target is actually eligible.
other = pane.request_threadsafe("execution_snapshot", {"section_slugs": ["other"]})["sections"]
if "Independent guard proof." not in other[0]["paragraphs"]:
    baseline = pane.request_threadsafe("describe", {})
    inserted = pane.request_threadsafe(
        "paragraph_insert",
        {
            "anchor": other[0]["paragraphs"][0],
            "expected_matches": 1,
            "position": "after",
            "paragraphs": [{"text": "Independent guard proof."}],
            "expectedBodySha256": baseline["bodySha256"],
        },
    )
    assert inserted["applied"] is True
ops = [
    {
        "op": "replace_cell_markdown",
        "section": "target",
        "args": {
            "table_id": 1,
            "row_index": 1,
            "cell_index": 1,
            "markdown": "Blocked by peer lock",
        },
    },
    {
        "op": "replace_text",
        "section": "other",
        "args": {
            "find": "Independent guard proof.",
            "replace": "Independent guard verified.",
            "expected_matches": 1,
        },
    },
    {"op": "live_save", "section": "", "args": {}},
]
prepared = runs.prepare(path, ops)
peer = shared.Client(shared.runtime_dir() / "broker.sock")
peer_pane = peer.get(Path(path).name)
listed = peer_pane.request_threadsafe("sections_list", {})["sections"]
key = next(s["section_key"] for s in listed if s["slug"] == "target")
peer.rpc("section_lock", **peer_pane.target(), sections=[key])
try:
    report = runs.execute(
        path,
        prepared["run_id"],
        prepared["manifest_sha256"],
        "User-approved disposable protected-target acceptance",
    )
    assert (
        report["status"] == "partial"
        and report["applied_count"] == 2
        and report["remaining"] == [0]
    ), report
    assert report["issues"][0]["code"] == "LOCKED_BY_OTHER_CLIENT", report
    assert (
        pane.request_threadsafe(
            "cell_get",
            {"executionTargetSection": "target", "table_index": 1, "row_index": 1, "cell_index": 1},
        )["text"].strip()
        == before
    )
    results["protected_target"] = report
finally:
    peer.rpc("section_unlock", **peer_pane.target())
    peer.close()
Path("/private/tmp/word-101-preflight-guards-evidence.json").write_text(
    json.dumps(results, indent=2)
)
print(json.dumps(results), flush=True)
