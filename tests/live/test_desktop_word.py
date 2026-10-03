"""Live acceptance for the desktop-Word open-document query (issue #29).

LEAD:-run only; needs Word installed and an Automation grant. The unit tests
mock osascript, which is how a Word-version AppleScript incompatibility
(`repeat with d in documents` -> -1708 on 16.113) went unnoticed.
Run with: uv run pytest tests/live --run-live
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.live

REPO = Path(__file__).resolve().parents[2]


def test_status_checks_successfully_against_running_word(word_available):
    from verified_docx_mcp import desktop_word

    observed = desktop_word.status(REPO / "tests" / "fixtures" / "sections.docx")
    assert observed["supported"] is True
    assert observed["checked"] is True, observed
    assert observed["open"] is False
