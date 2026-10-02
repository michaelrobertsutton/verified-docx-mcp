from unittest.mock import Mock, patch

import pytest
from test_live_edit_behavior import run_harness

from verified_docx_mcp.errors import ErrorCode, VerifyError, _make_error
from verified_docx_mcp.live import scopes_live


def test_textbox_js():
    run_harness("textboxes_harness.mjs")


def test_transient_read_retries_once_and_preserves_empty():
    error = _make_error(ErrorCode.HOST_SHAPE_READ_FAILED, "host failed")
    assert error.envelope.retryable
    with patch.object(scopes_live.write_mode, "live_session_for", return_value=Mock()), \
         patch.object(scopes_live.write_mode, "require_capability"), \
         patch.object(scopes_live.comments_live, "_request", side_effect=[error, {"text": ""}]) as request:
        assert scopes_live.read_scope("doc.docx", "textbox:epoch:1")["text"] == ""
        assert request.call_count == 2


def test_persistent_failure_is_not_empty():
    error = _make_error(ErrorCode.HOST_SHAPE_READ_FAILED, "host failed")
    with patch.object(scopes_live.write_mode, "live_session_for", return_value=Mock()), \
         patch.object(scopes_live.write_mode, "require_capability"), \
         patch.object(scopes_live.comments_live, "_request", side_effect=error) as request, \
         pytest.raises(VerifyError):
        scopes_live.read_scope("doc.docx")
    assert request.call_count == 2
