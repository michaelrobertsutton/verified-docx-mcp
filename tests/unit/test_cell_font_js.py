"""Runs tests/unit/js/cell_font_harness.mjs under node (issue #107's pane-side pure checks).
Skipped when node is not installed."""

from __future__ import annotations

import shutil
import subprocess
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
HARNESS = Path(__file__).resolve().parent / "js" / "cell_font_harness.mjs"
NODE = shutil.which("node")


@unittest.skipUnless(NODE, "node not found on PATH")
class CellFontHarnessTests(unittest.TestCase):
    def test_harness_passes(self):
        proc = subprocess.run(
            [NODE, str(HARNESS), str(REPO / "addin" / "taskpane.js")],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("cell font harness passed", proc.stdout)


if __name__ == "__main__":
    unittest.main()
