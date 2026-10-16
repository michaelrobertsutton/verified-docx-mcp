import subprocess
from pathlib import Path


def test_pane_anchored_text_harness():
    root=Path(__file__).resolve().parents[2]
    subprocess.run(['node', str(root/'tests/unit/js/anchored_text_harness.mjs'), str(root/'addin/taskpane.js')],check=True,capture_output=True,text=True)
