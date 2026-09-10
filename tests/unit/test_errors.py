"""Unit tests for src/verified_docx_mcp/errors.py — the lifted
ErrorCode/_make_error envelope (see that module's header comment for the
GoogleDocs-MCP source)."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp.errors import ErrorCode, VerifyError, _make_error


class MakeErrorTests(unittest.TestCase):
    def test_make_error_builds_envelope(self):
        exc = _make_error(ErrorCode.DOCX_LOCKED, "the file is locked", {"owner": "A. Reviewer"})
        self.assertIsInstance(exc, VerifyError)
        self.assertEqual(exc.envelope.error_code, ErrorCode.DOCX_LOCKED)
        self.assertEqual(exc.envelope.message, "the file is locked")
        self.assertEqual(exc.envelope.diagnostics, {"owner": "A. Reviewer"})

    def test_make_error_defaults_diagnostics_to_empty_dict(self):
        exc = _make_error(ErrorCode.INVALID_INPUT, "bad input")
        self.assertEqual(exc.envelope.diagnostics, {})

    def test_to_dict_shape(self):
        exc = _make_error(ErrorCode.RENDER_FAILED, "boom", {"detail": "stderr text"})
        d = exc.envelope.to_dict()
        self.assertEqual(
            set(d.keys()), {"error_code", "message", "diagnostics", "retryable"}
        )
        self.assertEqual(d["error_code"], "RENDER_FAILED")
        self.assertEqual(d["message"], "boom")
        self.assertEqual(d["diagnostics"], {"detail": "stderr text"})
        self.assertFalse(d["retryable"])

    def test_no_code_is_retryable_yet(self):
        # WP-02 has no revision-stamped write to race against — see
        # errors.py's _RETRYABLE_CODES comment. Every current member is
        # non-retryable.
        for code in ErrorCode:
            with self.subTest(code=code):
                exc = _make_error(code, "x")
                self.assertFalse(exc.envelope.retryable)

    def test_render_path_codes_present(self):
        # These four map 1:1 onto render.py's RenderError.code values
        # (server.py's execute_export_pdf relies on this).
        for name in (
            "AUTOMATION_NOT_GRANTED",
            "WORD_SANDBOX_UNAVAILABLE",
            "RENDER_FAILED",
            "RENDER_ENGINE_UNAVAILABLE",
        ):
            with self.subTest(name=name):
                self.assertIn(name, ErrorCode.__members__)

    def test_image_source_unsupported_deliberately_absent(self):
        # The Google server's 18th/19th error code — dropped here because a
        # local path is exactly what this server reads natively. See
        # errors.py's header comment.
        self.assertNotIn("IMAGE_SOURCE_UNSUPPORTED", ErrorCode.__members__)


if __name__ == "__main__":
    unittest.main()
