"""Quarantine machinery for the live acceptance suite (needs a real Word
install and a macOS Automation grant for the invoking app). Mirrors
verified-googledocs-mcp's tests/live/conftest.py --run-live pattern.

  * `pytest` with no flag  -> live tests are skipped.
  * `pytest --run-live`    -> live tests run, if Word is actually present.

LEAD:-run only (issue #28 WP-02): this repo's own gate is
`uv run pytest tests/unit`.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--run-live",
        action="store_true",
        default=False,
        help="Run the live acceptance suite against a real Word install.",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if config.getoption("--run-live"):
        return
    skip_live = pytest.mark.skip(
        reason="live acceptance suite — pass --run-live (and have Word installed, "
        "an Automation grant, and be running from the app hosting the MCP server) to run"
    )
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip_live)


@pytest.fixture(scope="session")
def word_available() -> None:
    from verified_docx_mcp.render import _word_sandbox_root

    if not _word_sandbox_root().is_dir():
        pytest.skip("Word's sandbox container directory does not exist — launch Word once.")
