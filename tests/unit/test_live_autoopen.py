"""Auto-open tagging of the Live pane (#66); never contacts Word."""
import re
import shutil
import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest

from verified_docx_mcp.errors import VerifyError
from verified_docx_mcp.live import autoopen_live, comments_live, write_mode
from verified_docx_mcp.live.protocol import VALID_OPS

ROOT = Path(__file__).resolve().parents[2]


def _session(capabilities):
    return SimpleNamespace(hello=SimpleNamespace(capabilities=set(capabilities)), document_name="doc.docx")


def test_ops_are_registered():
    assert {"autoopen_get", "autoopen_set"} <= VALID_OPS


def test_read_does_not_change_the_document():
    with patch.object(write_mode, "live_session_for", return_value=_session({"autoopen"})), \
         patch.object(comments_live, "_request", return_value={"enabled": True, "supported": True}) as request:
        result = autoopen_live.execute_autoopen("/x/doc.docx")
    request.assert_called_once()
    assert request.call_args.args[1] == "autoopen_get"
    assert result["enabled"] is True and result["applied"] is False


def test_set_sends_the_boolean_and_reports_the_change():
    with patch.object(write_mode, "live_session_for", return_value=_session({"autoopen"})), \
         patch.object(comments_live, "_request", return_value={"enabled": True, "was_enabled": False, "supported": True}) as request:
        result = autoopen_live.execute_autoopen("/x/doc.docx", True)
    assert request.call_args.args[1:] == ("autoopen_set", {"enabled": True})
    assert result["applied"] is True and result["enabled"] is True and "modified" in result["note"]


def test_old_pane_refuses_before_any_op():
    with patch.object(write_mode, "live_session_for", return_value=_session(set())), \
         patch.object(comments_live, "_request") as request, pytest.raises(VerifyError) as err:
        autoopen_live.execute_autoopen("/x/doc.docx", True)
    request.assert_not_called()
    assert "'autoopen' capability" in str(err.value) and "Reload the Live pane" in str(err.value)


def test_manifest_uses_the_autoshow_task_pane_id_and_pane_advertises_it():
    manifest = (ROOT / "addin/manifest.xml").read_text()
    assert "<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>" in manifest
    assert re.search(r'PANE_CAPABILITIES = \[[^\]]*"autoopen"', (ROOT / "addin/taskpane.js").read_text())


def test_autoopen_js_behavior():
    node = shutil.which("node")
    assert node, "Node is required for pane behavior tests"
    subprocess.run([node, str(ROOT / "tests/unit/js/autoopen_harness.mjs"), str(ROOT / "addin/taskpane.js")],
                   check=True, capture_output=True, text=True)
