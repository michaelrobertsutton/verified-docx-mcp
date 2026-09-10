# Adapted from GoogleDocs-MCP src/verified_googledocs_mcp/verify.py's error
# envelope (ErrorCode, ErrorEnvelope, VerifyError, _make_error), commit
# 374bcf8, lines ~30-115. Lifted per issue #28 WP-02 ("Lift, do not
# rewrite" — "the ErrorCode/_make_error envelope"). The MECHANISM is
# copied as-is (frozen dataclass envelope, VerifyError exception carrying
# it, _make_error() constructor); the ErrorCode MEMBERS are docx's own —
# the Google server's 19 members describe Docs-API-specific failures
# (STALE_RANGE, TAB_NOT_FOUND, SUGGESTIONS_PRESENT, ...) that mostly do
# not apply here, and docx has failure modes (a locked file, a sync
# client mid-upload, Word's Automation grant, its render sandbox) the
# Google server has none of. One code is a deliberate, documented
# omission: IMAGE_SOURCE_UNSUPPORTED (source must be a public URL, not a
# local path) does not exist here because a local path is exactly what
# this server reads natively (core/document-backend-protocol.md was
# never meant to be read by this repo directly — this comment states the
# rationale inline instead).
#
# Members below come from core/document-backend-protocol.md §4's docx
# runtime error codes (the authoritative source for their names and
# meaning) plus the render-path codes documented in the vendored
# render.py's own module docstring. Not every member is raised by a WP-02
# tool yet: SYNC_IN_FLIGHT and CONFLICT_COPY_DETECTED are write-path
# codes that first get raised by the write guard landing in WP-04/WP-10;
# they are declared here now so the enum is the single, stable
# vocabulary every later WP extends rather than redefines. Extend this
# enum in the same commit that first raises a new member — same
# discipline as MUTATING_TOOLS in middleware.py.

from __future__ import annotations

import dataclasses
from enum import Enum
from typing import Any


class ErrorCode(Enum):
    # Generic
    INVALID_INPUT = "INVALID_INPUT"

    # Path safety (paths.py: the $DOCS containment rule + allowlist)
    DOCX_PATH_ESCAPE = "DOCX_PATH_ESCAPE"
    DOCX_ROOT_NOT_FOUND = "DOCX_ROOT_NOT_FOUND"

    # Lock / consistency window (core/document-backend-protocol.md §4, §9)
    DOCX_LOCKED = "DOCX_LOCKED"  # writes only; reads snapshot instead (WP-02)
    SNAPSHOT_FAILED = "SNAPSHOT_FAILED"  # read-path snapshot validation exhausted its retries (WP-02)
    SYNC_IN_FLIGHT = "SYNC_IN_FLIGHT"  # reserved: write guard, WP-04/WP-10
    CONFLICT_COPY_DETECTED = "CONFLICT_COPY_DETECTED"  # reserved: post-write sweep, WP-10

    # Projection / read tools (projection.py, WP-03)
    PART_NOT_FOUND = "PART_NOT_FOUND"  # read_document(part=...) names a package part that does not exist

    # Render path (render.py's module docstring; core doc §4)
    AUTOMATION_NOT_GRANTED = "AUTOMATION_NOT_GRANTED"
    WORD_SANDBOX_UNAVAILABLE = "WORD_SANDBOX_UNAVAILABLE"
    RENDER_FAILED = "RENDER_FAILED"
    RENDER_ENGINE_UNAVAILABLE = "RENDER_ENGINE_UNAVAILABLE"


# Which codes signal a transient condition worth a single retry by the
# caller. Empty for now — WP-02 has no revision-stamped write to race
# against (that arrives with REVISION_CONFLICT-style handling in WP-04).
# SYNC_IN_FLIGHT is deliberately NOT here: core/document-backend-protocol.md
# §4 specifies exactly one wait (<=10s) performed by the tool itself before
# raising, not a caller-side retry loop.
_RETRYABLE_CODES: frozenset[ErrorCode] = frozenset()


@dataclasses.dataclass(frozen=True)
class ErrorEnvelope:
    """Typed error structure returned by every verification failure."""

    error_code: ErrorCode
    message: str
    diagnostics: dict[str, Any]
    retryable: bool

    def to_dict(self) -> dict[str, Any]:
        return {
            "error_code": self.error_code.value,
            "message": self.message,
            "diagnostics": self.diagnostics,
            "retryable": self.retryable,
        }


class VerifyError(Exception):
    """Exception that carries a fully-typed ErrorEnvelope."""

    def __init__(self, envelope: ErrorEnvelope) -> None:
        super().__init__(envelope.message)
        self.envelope = envelope


def _make_error(
    code: ErrorCode,
    message: str,
    diagnostics: dict[str, Any] | None = None,
) -> VerifyError:
    return VerifyError(
        ErrorEnvelope(
            error_code=code,
            message=message,
            diagnostics=diagnostics or {},
            retryable=code in _RETRYABLE_CODES,
        )
    )
