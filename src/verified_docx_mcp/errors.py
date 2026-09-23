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
# tool yet: CONFLICT_COPY_DETECTED is a write-path code that first gets
# raised by the lock guard landing in WP-10 (layers 3-4); it is declared
# here now so the enum is the single, stable vocabulary every later WP
# extends rather than redefines. SYNC_IN_FLIGHT was reserved the same way
# in WP-02 and is now actually raised, by WP-04's write guard (see
# mutations.py). Extend this enum in the same commit that first raises a
# new member — same discipline as MUTATING_TOOLS in middleware.py.

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
    DOCX_LOCKED = "DOCX_LOCKED"  # writes only; reads snapshot instead (WP-02). Also raised by acquire_lock's own layer-4 .jsclaim O_EXCL mutex (WP-10) when this same server has a write to the same file already in flight -- a same-machine case, distinct from (but reported with the same code as) the owner-file (layer 0) case.
    SNAPSHOT_FAILED = "SNAPSHOT_FAILED"  # read-path snapshot validation exhausted its retries (WP-02)
    SYNC_IN_FLIGHT = "SYNC_IN_FLIGHT"  # reserved: write guard, WP-04/WP-10
    CONFLICT_COPY_DETECTED = "CONFLICT_COPY_DETECTED"  # WP-10's post-write conflict-copy sweep: an EVIDENCE FLAG (evidence["conflict_copy_detected"]=True + the sibling's name), never raised -- the write itself still succeeded and is reported as applied

    # Projection / read tools (projection.py, WP-03)
    PART_NOT_FOUND = "PART_NOT_FOUND"  # read_document(part=...) names a package part that does not exist

    # Render path (render.py's module docstring; core doc §4)
    AUTOMATION_NOT_GRANTED = "AUTOMATION_NOT_GRANTED"
    WORD_SANDBOX_UNAVAILABLE = "WORD_SANDBOX_UNAVAILABLE"
    RENDER_FAILED = "RENDER_FAILED"
    RENDER_ENGINE_UNAVAILABLE = "RENDER_ENGINE_UNAVAILABLE"
    SECTION_GEOMETRY_UNAVAILABLE = "SECTION_GEOMETRY_UNAVAILABLE"  # export_pdf(section_keys=[...]) (issue #102): a requested section's page-span could not be verified -- a probed heading's text did not match find_sections_impl's heading_text, a probed ordinal came back {"error": ...}, or page_height_pt was missing. Raised only in explicit mode (section_keys given); default mode (section_keys=None) degrades instead, returning sections: null + sections_error alongside the PDF and page_count, which are unaffected either way -- see geometry.assemble_sections and execute_export_pdf.

    # Markdown mutations (mutations.py, markdown_to_ooxml.py; WP-04)
    REVISION_CONFLICT = "REVISION_CONFLICT"  # a passed revision_before no longer matches the file on disk
    COMMENT_ANCHORS_IN_RANGE = "COMMENT_ANCHORS_IN_RANGE"  # w:commentRangeStart/End/commentReference inside the target range; needs force=true
    TRACKED_CHANGES_PRESENT = "TRACKED_CHANGES_PRESENT"  # w:ins/w:del inside the target range; needs force=true
    STYLE_NOT_FOUND = "STYLE_NOT_FOUND"  # a markdown heading level has no matching style in list_styles
    SECTION_NOT_FOUND = "SECTION_NOT_FOUND"  # replace_range_markdown's section_key does not match find_sections' output.
    # Not named in the issue #28 plan text for WP-04 (which lists STYLE_NOT_FOUND
    # but not this one); added because replace_range_markdown needs SOME code
    # for "the section_key does not exist" and reusing STYLE_NOT_FOUND for that
    # would conflate two different failures. Flagged in this WP's PR notes.
    OPC_INVALID = "OPC_INVALID"  # the rendered .docx failed OPC validation before the atomic write (original untouched)
    VERIFICATION_FAILED = "VERIFICATION_FAILED"  # the post-write re-read/re-project did not confirm the write; restored from .jsbak

    # Text location + targeted edits (locate.py, text_edit.py; WP-06)
    ZERO_MATCH = "ZERO_MATCH"  # `find` not located after the full normalization ladder; near-miss in diagnostics
    MATCH_COUNT_MISMATCH = "MATCH_COUNT_MISMATCH"  # match count != expected_matches (D4: expected_matches is required, no default)
    STRUCTURAL_BOUNDARY = "STRUCTURAL_BOUNDARY"  # a match crosses a w:p/w:tbl/w:tc boundary

    # Tracked changes (tracked_changes.py; WP-07)
    REVISION_ID_NOT_FOUND = "REVISION_ID_NOT_FOUND"  # accept_tracked_changes/reject_tracked_changes named a w:ins/w:del id not present in the document

    # Comments (comments.py; WP-09). COMMENT_STILL_OPEN is adapted, not
    # lifted verbatim, from GoogleDocs-MCP's own member of the same name
    # (verify.py) -- see comments.py's resolve_comment docstring for the
    # divergence this server's local, synchronous write means it can only
    # ever reach via a genuine bug in this server's own code, unlike
    # Google's version, which guards a real Drive-API eventual-consistency
    # hazard this backend has no analogue of.
    COMMENT_STILL_OPEN = "COMMENT_STILL_OPEN"  # resolve_comment's own post-write re-read did not confirm w15:done="1"

    # Tables (tables.py; issue #28 WP-14)
    TABLE_NOT_FOUND = "TABLE_NOT_FOUND"  # table_id does not match any table (call list_tables)
    TABLE_ROW_NOT_FOUND = "TABLE_ROW_NOT_FOUND"  # row_index out of range for the named table
    TABLE_CELL_NOT_FOUND = "TABLE_CELL_NOT_FOUND"  # cell_index out of range for the named row
    MERGED_OR_NESTED_TABLE = "MERGED_OR_NESTED_TABLE"  # replace_table_row's refusal: the target table has a w:gridSpan/w:vMerge cell or a nested w:tbl anywhere in it -- use replace_cell_markdown instead

    # Images (images.py; issue #28 WP-15a). GoogleDocs-MCP's 18th code,
    # IMAGE_SOURCE_UNSUPPORTED (source must be a public URL, not a local
    # path), is a deliberate omission here -- a local path is exactly what
    # insert_image reads natively (see errors.py's own header comment).
    UNSUPPORTED_IMAGE_FORMAT = "UNSUPPORTED_IMAGE_FORMAT"  # insert_image's path is not a .png/.svg, or the file's own bytes do not parse as one
    SVG_RASTERIZATION_FAILED = "SVG_RASTERIZATION_FAILED"  # insert_image could not produce the PNG fallback part Word requires for an SVG (macOS `sips` failed or is unavailable)

    # Style application (text_edit.py; issue #28 WP-15a)
    UNSUPPORTED_STYLE_TYPE = "UNSUPPORTED_STYLE_TYPE"  # apply_style's style_id names a style whose w:type is neither "paragraph" nor "character" (e.g. "table"/"numbering")

    # Live mode (live/session.py, live/bridge.py; issue #106 WP-2). Every
    # member here maps 1:1 to a live/session.py exception of the same
    # shape (LiveUnavailable/LiveDisconnected/LiveStale/LiveOpFailed);
    # server.py's WP-3/WP-4 live write path is what actually catches
    # those and raises these. Declared now, in the same commit as the
    # session-layer exceptions they mirror, per this file's own header
    # comment's rule ("extend this enum in the same commit that first
    # raises a new member") -- not yet raised by any tool in this WP
    # (write_mode="live" lands in WP-3/WP-4), same status as
    # CONFLICT_COPY_DETECTED/SYNC_IN_FLIGHT above when THEY were first
    # declared.
    LIVE_UNAVAILABLE = "LIVE_UNAVAILABLE"  # no connected pane session for the target document (live/session.py LiveUnavailable)
    LIVE_DISCONNECTED = "LIVE_DISCONNECTED"  # the pane's WebSocket closed mid-request, or a reply never arrived within the op timeout (live/session.py LiveDisconnected)
    LIVE_STALE = "LIVE_STALE"  # the pane's body hash before the op (result["pre"]) did not match a caller-supplied expected_body_sha256 (live/session.py LiveStale)
    LIVE_OP_FAILED = "LIVE_OP_FAILED"  # the pane replied ok=false (e.g. an expected_matches count mismatch it refused to act on) (live/session.py LiveOpFailed)

    # Issue #154: a Live pane is connected for the target document, so the
    # server-side "lock_status owner file" signal that used to gate a
    # file-mode write is unreliable (Word for Mac + a SharePoint/OneDrive
    # sync never writes that owner file) -- the connected-pane check below
    # is the independent second signal that actually protects a file-mode
    # write from racing Word's own autosave.
    LIVE_SESSION_ACTIVE = "LIVE_SESSION_ACTIVE"  # mutations._guard_before_write refused a FILE-MODE write because a pane session is connected for this document -- use write_mode="live" (where the tool has one), or save and close the document in Word (not just the pane) and retry
    LIVE_SESSION_MISMATCH = "LIVE_SESSION_MISMATCH"  # live/write_mode.py's auto/live routing found a session matching this document's basename, but that session's own document_url resolves to a different local file -- refusing rather than risk mutating the wrong document

    # Issue #22 B2: within_row_containing/rowAnchor requires the connected
    # pane to have reported the "row_scope" capability in its own hello --
    # an already-connected OLD pane (from before this feature existed)
    # would otherwise silently ignore an unrecognized rowAnchor payload
    # key and run an UNSCOPED op instead of refusing, which is exactly the
    # "wrote to all N identical cells" bug this feature exists to prevent,
    # just moved onto the wire instead of fixed. live/write_mode.py's
    # require_capability raises this BEFORE the op is ever sent.
    LIVE_CAPABILITY_MISSING = "LIVE_CAPABILITY_MISSING"  # the connected pane did not report a capability this call requires (live/write_mode.py require_capability)


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
