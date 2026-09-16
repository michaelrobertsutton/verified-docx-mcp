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
SECTIONS_FIXTURE = REPO / "tests" / "fixtures" / "sections.docx"


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


def test_export_pdf_reports_section_spans_for_bundled_sections_fixture(word_available):
    """Issue #102's live acceptance: sections.docx (three headings —
    Overview, Background, Next Steps) renders to 1 page, and export_pdf's
    "sections" list reports all three, each with a verified start
    paragraph."""
    from verified_docx_mcp import paths, server

    assert SECTIONS_FIXTURE.is_file(), f"bundled fixture missing: {SECTIONS_FIXTURE}"

    old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
    with tempfile.TemporaryDirectory() as tmp:
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = f"{tmp}{os.pathsep}{REPO}"
        try:
            out_pdf = Path(tmp) / "sections.pdf"
            result = server.execute_export_pdf(str(SECTIONS_FIXTURE), str(out_pdf))
        finally:
            if old_allowed is None:
                os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
            else:
                os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = old_allowed

    assert result["page_count"] == 1
    assert result["sections_error"] is None
    assert result["page_height_pt"] is not None and result["page_height_pt"] > 0

    sections = result["sections"]
    assert [s["section_key"] for s in sections] == ["overview-1", "background-1", "next-steps-1"]

    starts = []
    for section in sections:
        assert section["start_page"] == 1
        assert 0.0 <= section["start_fraction"] <= 1.0
        assert 0.0 <= section["end_fraction"] <= 1.0
        assert section["text_verified"] is True
        starts.append(section["start_page"] + section["start_fraction"])
    # Monotonically increasing starts -- headings appear in document order
    # and never overlap.
    assert starts == sorted(starts)
    assert len(set(starts)) == len(starts)

    # sum(pages) should be within 0.1 of the distance from the first
    # section's start to the last section's end (they are contiguous, so
    # this is mostly a sanity check on the rounding budget, not an exact
    # identity).
    total_pages = sum(s["pages"] for s in sections)
    first = sections[0]
    last = sections[-1]
    span = (last["end_page"] + last["end_fraction"]) - (first["start_page"] + first["start_fraction"])
    assert abs(total_pages - span) < 0.1


def test_doctor_passes(word_available):
    from verified_docx_mcp.server import doctor

    assert doctor() == 0
