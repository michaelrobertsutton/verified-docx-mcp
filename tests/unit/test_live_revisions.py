"""Offline revision state and pane behavior; never contacts Word."""
import shutil
import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from verified_docx_mcp.live import comments_live, revisions_live


def test_old_pane_never_reports_placeholder_revisions():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities=set(), instance_id=None, platform=None, host=None))
    with patch.object(comments_live, "_request", return_value={}), \
         patch.object(comments_live, "_session_for", return_value=session), \
         patch.object(comments_live, "_document_name_and_path", return_value=(Path("/tmp/doc.docx"), "doc")), \
         patch.object(comments_live, "_live_state", return_value=([], [])), \
         patch.object(comments_live, "_request", return_value={"comments": [], "counts": {"total": 0, "open": 0}, "session_epoch": "epoch"}), \
         patch.object(comments_live, "_file_comments_and_suggestions", return_value=([], [{"id": "fake"}])):
        result = comments_live.execute_list_open_items_live("/tmp/doc.docx")
    assert result["pending_suggestions"] is None
    assert "pane lacks" in result["pending_suggestions_reason"]


def test_live_list_does_not_substitute_file_state():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    with patch.object(comments_live, "_request", return_value={"revisions": [{"text": "live"}], "coverage": "body"}):
        assert revisions_live.list_revisions(session)["revisions"][0]["text"] == "live"


def test_partial_coverage_is_reported_when_markup_exceeds_api_list():
    session = SimpleNamespace(hello=SimpleNamespace(capabilities={"live_revisions"}))
    reply = {"revisions": [{"text": "a"}] * 2, "coverage": "body", "ooxml_revision_count": 5}
    with patch.object(comments_live, "_request", return_value=reply):
        state = revisions_live.list_revisions(session)
    assert state["coverage"] == "partial"
    assert "2 revisions" in state["reason"] and "has 5" in state["reason"]
    full = {**reply, "ooxml_revision_count": 2}
    with patch.object(comments_live, "_request", return_value=full):
        assert revisions_live.list_revisions(session)["coverage"] == "body"


def test_revision_js_behavior():
    node = shutil.which("node")
    assert node, "Node is required for new pane behavior tests"
    root = Path(__file__).resolve().parents[2]
    subprocess.run([node, str(root / "tests/unit/js/revisions_harness.mjs"),
                    str(root / "addin/taskpane.js")], check=True, capture_output=True, text=True)
