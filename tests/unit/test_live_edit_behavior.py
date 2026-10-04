"""New pane harnesses are mandatory and run without a Word session."""
import shutil
import subprocess
from pathlib import Path


def run_harness(name):
    node = shutil.which("node")
    assert node, "Node is required for live editing behavior tests"
    root = Path(__file__).resolve().parents[2]
    subprocess.run([node, str(root / "tests/unit/js" / name), str(root / "addin/taskpane.js")],
                   check=True, capture_output=True, text=True)


def test_comment_loss_guard():
    run_harness("edit_harness.mjs")
