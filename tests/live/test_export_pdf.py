"""Live acceptance for export_pdf against the real Word render path.

LEAD:-run only — needs Word installed and a macOS Automation grant for the
app hosting this process (see README.md's `verified-docx-mcp doctor`).
Run with: uv run pytest tests/live --run-live
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.live

REPO = Path(__file__).resolve().parents[2]
FIXTURE = REPO / "tests" / "fixtures" / "word" / "one-page.docx"


def test_export_pdf_renders_bundled_one_page_fixture(word_available):
    from verified_docx_mcp import paths, server

    assert FIXTURE.is_file(), f"bundled fixture missing: {FIXTURE}"

    old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
    with tempfile.TemporaryDirectory() as tmp:
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = f"{tmp}{os.pathsep}{REPO}"
        try:
            out_pdf = Path(tmp) / "one-page.pdf"
            result = server.execute_export_pdf(str(FIXTURE), str(out_pdf))
        finally:
            if old_allowed is None:
                os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
            else:
                os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = old_allowed

    assert result["engine"] == "word"
    assert result["page_count"] == 1
    assert Path(result["pdf_path"]).is_file()
    assert len(result["sha256"]) == 64


def test_doctor_passes(word_available):
    from verified_docx_mcp.server import doctor

    assert doctor() == 0
