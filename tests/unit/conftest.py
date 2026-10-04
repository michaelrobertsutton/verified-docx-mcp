"""Shared unit-test fixtures."""

from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _isolated_state_dir(tmp_path_factory, monkeypatch):
    """Point XDG_STATE_HOME at a per-test temp dir (issue #27).

    Every file-mode write now records a write-ledger entry, and every
    mutating call appends to audit.jsonl, both under the state dir. Without
    this, the unit suite writes into the developer's real
    ~/.local/state/verified-docx-mcp/, and ledger records from one test can
    leak into the next. Tests that need to inspect the state dir read
    ``os.environ["XDG_STATE_HOME"]``.
    """
    monkeypatch.setenv("XDG_STATE_HOME", str(tmp_path_factory.mktemp("state")))


@pytest.fixture(autouse=True)
def _isolate_desktop_word_query(monkeypatch, request):
    """Unit writes must never query the developer's actual Word session."""
    if request.node.module.__name__ == 'test_desktop_word':
        return
    from verified_docx_mcp import desktop_word
    monkeypatch.setattr(desktop_word, 'status', lambda path:
                        {'supported': True, 'checked': True, 'open': False})
