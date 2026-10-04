from pathlib import Path
from types import SimpleNamespace

import pytest

from verified_docx_mcp import desktop_word
from verified_docx_mcp.errors import VerifyError


def test_open_without_pane(monkeypatch):
    monkeypatch.setattr(desktop_word.sys, 'platform', 'darwin')
    monkeypatch.setattr(desktop_word.subprocess, 'run', lambda *a, **k:
                        SimpleNamespace(returncode=0, stdout='/tmp/example.docx\n', stderr=''))
    with pytest.raises(VerifyError):
        desktop_word.raise_if_open(Path('/tmp/example.docx'))
    desktop_word.raise_if_open(Path('/tmp/other.docx'))


def test_automation_failure_refuses(monkeypatch):
    monkeypatch.setattr(desktop_word.sys, 'platform', 'darwin')
    monkeypatch.setattr(desktop_word.subprocess, 'run', lambda *a, **k:
                        SimpleNamespace(returncode=1, stdout='', stderr='not authorized'))
    with pytest.raises(VerifyError):
        desktop_word.raise_if_open(Path('/tmp/example.docx'))


def test_cloud_url_name_is_decoded(monkeypatch):
    monkeypatch.setattr(desktop_word.sys, 'platform', 'darwin')
    monkeypatch.setattr(desktop_word.subprocess, 'run', lambda *a, **k:
                        SimpleNamespace(returncode=0,
                            stdout='https://example.sharepoint.com/docs/My%20Document.docx?web=1',
                            stderr=''))
    with pytest.raises(VerifyError):
        desktop_word.raise_if_open(Path('/tmp/My Document.docx'))


def test_script_enumerates_documents_by_index():
    # Word 16.113 rejects `repeat with d in documents` (-1708); the indexed
    # form is the only one verified against desktop Word.
    assert '    repeat with d in documents\n' not in desktop_word._SCRIPT
    assert 'full name of document i' in desktop_word._SCRIPT
