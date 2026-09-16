"""FastMCP server: tool registration, doctor, and entry point.

No OAuth, no auth step of any kind: this backend reads and writes a local
.docx package directly (pure-Python OOXML manipulation for the tools WP-03/
WP-04 add later), so reading and writing need no Word installation at all.
Word is required only for the render path (``export_pdf``), which
additionally needs macOS Automation permission granted for the calling app
(see ``AUTOMATION_NOT_GRANTED`` below) — a machine without Word can still
read and write this backend, it just cannot render a PDF. This mirrors
core/document-backend-protocol.md §4 in the JennyStack repo (the consumer of
this server), word for word.

Entry point dispatch
---------------------
  verified-docx-mcp            -> start the stdio MCP server
  verified-docx-mcp doctor     -> run the local diagnostic (see doctor())

Tools registered here (WP-02): ``export_pdf``, ``lock_status``. WP-03 adds
five more READ tools — ``list_parts``, ``read_document``, ``find_sections``,
``list_page_sections``, ``list_styles`` — built on ``projection.py``. None
of those seven is a mutating tool (export_pdf writes a PDF, never the
source .docx; every other one only reads). WP-04 adds the first three
mutating tools — ``replace_body_markdown``, ``replace_range_markdown``,
``append_markdown`` (built on ``mutations.py`` + ``markdown_to_ooxml.py``)
— each gated by a write guard (lock/sync/revision) that runs BEFORE any
temp file is written, and each is in MUTATING_TOOLS (middleware.py).

Every WP-03 read tool honors core/document-backend-protocol.md §4's "reads
never refuse" rule via ``_read_local_copy`` below: when Word's owner file is
present, the tool reads a validated snapshot (``paths.snapshot_docx_package``)
instead of the live file, rather than returning DOCX_LOCKED (which gates
writes only).
"""

from __future__ import annotations

import difflib
import glob
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, NoReturn

from fastmcp import FastMCP

from . import comments, images, mutations, paths, projection, tables, text_edit, tracked_changes
from . import render as render_module
from .errors import ErrorCode, VerifyError, _make_error
from .middleware import EvidenceEnforcementMiddleware

mcp = FastMCP(
    "verified-docx-mcp",
    instructions=(
        "MCP server for local .docx files with verified writes. Reading and "
        "writing need no Word installation (pure-Python OOXML manipulation); "
        "Word is required only for export_pdf, which additionally needs a "
        "macOS Automation grant for the calling app. export_pdf writes a "
        "local PDF file only and never modifies the source .docx. lock_status "
        "reports Word/LibreOffice owner-file and sync-quiesce state as data "
        "only — it never refuses a call by itself."
    ),
)

mcp.add_middleware(EvidenceEnforcementMiddleware())


def _raise_tool_error(exc: VerifyError) -> NoReturn:
    """Raise a FastMCP ToolError carrying a JSON error envelope.

    Mirrors verified-googledocs-mcp's server.py helper of the same name.
    """
    from fastmcp.exceptions import ToolError

    raise ToolError(json.dumps(exc.envelope.to_dict(), ensure_ascii=False)) from exc


# ---------------------------------------------------------------------------
# Tool: export_pdf
# ---------------------------------------------------------------------------


def _resolve_export_output_path(output_path: str) -> Path:
    """Validate *output_path* for export_pdf: allowlisted, not denylisted,
    parent directory must already exist (no implicit mkdir — mirrors
    verified-googledocs-mcp's export_pdf validation), and an existing
    target must be a regular file."""
    if not output_path.strip():
        raise _make_error(ErrorCode.INVALID_INPUT, "output_path must not be empty")

    resolved = paths.resolve_allowed_docx_path(output_path, must_exist=False)

    parent = resolved.parent
    if not parent.is_dir():
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"Parent directory does not exist: {parent}",
            {"output_path": output_path, "resolved_path": str(resolved), "parent": str(parent)},
        )

    if resolved.exists() and not resolved.is_file():
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"Output path exists and is not a regular file: {output_path!r}",
            {"output_path": output_path, "resolved_path": str(resolved)},
        )

    return resolved


def execute_export_pdf(path: str, output_path: str, *, timeout: int = render_module.DEFAULT_TIMEOUT) -> dict[str, Any]:
    """Render *path* (.docx) to *output_path* (.pdf) via Word automation.

    Returns {pdf_path, sha256, page_count, page_count_source, engine,
    left_open_document}. page_count is Word's own count when available
    (authoritative — see render.py's module docstring), cross-checked
    against pdfinfo/regex; None (never a guessed 0) when neither source can
    determine it. Raises VerifyError(RENDER_ENGINE_UNAVAILABLE),
    (AUTOMATION_NOT_GRANTED), (WORD_SANDBOX_UNAVAILABLE), or (RENDER_FAILED)
    — see errors.py and render.py's module docstring for exactly when each
    fires.

    Not gated by DOCX_LOCKED (core/document-backend-protocol.md §4:
    "Reads and export_pdf are not gated by this code") — render_word()
    itself already stages a private copy of the source into Word's sandbox
    container before opening it, so a concurrently-open Word session on
    the original file is not disturbed.
    """
    source = paths.resolve_allowed_docx_path(path, must_exist=True)
    target = _resolve_export_output_path(output_path)

    try:
        result = render_module.render_word(str(source), str(target), timeout=timeout)
    except render_module.RenderError as exc:
        # 1:1 name mapping between render.py's RenderError codes and this
        # server's ErrorCode members (errors.py's docstring records why).
        try:
            code = ErrorCode(exc.code)
        except ValueError:
            code = ErrorCode.RENDER_FAILED
        raise _make_error(code, exc.message, {"detail": exc.detail} if exc.detail else {}) from exc

    pdf_path = result["pdf"]
    word_pages = result.get("pages")
    cross_check_pages, cross_check_source = render_module.page_count(pdf_path)

    page_count_value: int | None
    page_count_source: str | None
    if word_pages is not None:
        page_count_value, page_count_source = word_pages, "word"
    else:
        page_count_value, page_count_source = cross_check_pages, cross_check_source

    return {
        "pdf_path": pdf_path,
        "sha256": _sha256_file(pdf_path),
        "page_count": page_count_value,
        "page_count_source": page_count_source,
        "engine": "word",
        # Extra, informational field beyond the WP-02 return shape: Word's
        # AppleScript surface has no `close` command (render.py's module
        # docstring), so every successful render leaves its document open.
        # Surfaced here rather than silently dropped.
        "left_open_document": result.get("left_open_document"),
    }


def _sha256_file(path: str) -> str:
    import hashlib

    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@mcp.tool()
def export_pdf(path: str, output_path: str) -> dict[str, Any]:
    """Render a local .docx to PDF via Microsoft Word and report its page count.

    output_path must fall inside VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS
    (defaults to the user's home directory plus the Claude Code scratch
    root, /private/tmp/claude-<uid>, when that directory exists), must not
    resolve to a credential path, and its parent directory must already
    exist. This is a read/render tool: the source .docx is never modified
    (Word opens a private staged copy — see render.py), so the return value
    has no "applied" key.

    Returns pdf_path, sha256, page_count (best-effort; None — never a
    guessed 0 — when it cannot be determined), page_count_source
    ("word"|"pdfinfo"|"regex"|None), engine ("word"; the only engine, D2:
    no LibreOffice), and left_open_document (the rendered document's
    window title in Word — see render.py's module docstring for why it is
    never auto-closed).

    Errors:
      INVALID_INPUT             - a bad path or output_path
      RENDER_ENGINE_UNAVAILABLE - no render engine available (Word only)
      AUTOMATION_NOT_GRANTED    - macOS declined Automation control of Word
                                   for this app; run `verified-docx-mcp
                                   doctor` for the fix, scoped to the app
                                   hosting this MCP server's own process
      WORD_SANDBOX_UNAVAILABLE  - Word has never been launched on this
                                   machine (its sandbox container does not
                                   exist yet)
      RENDER_FAILED             - any other Word automation failure
    """
    try:
        return execute_export_pdf(path, output_path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tool: lock_status
# ---------------------------------------------------------------------------

_TMP_SIBLING_GLOBS = ("*.tmp", ".~*")


def _word_owner_file_matches(candidate_name: str, target_name: str) -> bool:
    """True if *candidate_name* (a "~$*" hit in the target's directory) is
    plausibly Word's owner file for *target_name*.

    Word's owner file is not literally "~$" + the full filename
    (core/document-backend-protocol.md §9): for a short target name it
    prefixes the FULL name ("test.docx" -> "~$test.docx", 0 characters
    dropped); for a longer one it replaces the first two characters
    instead of prefixing them ("Proposal-Final.docx" ->
    "~$oposal-Final.docx", 2 characters dropped) — so detection matches by
    shared filename suffix rather than constructing one exact expected
    name. Generalized here as: after stripping "~$", the remainder must be
    a suffix of target_name with either 0 or exactly 2 characters dropped
    from target_name's front.
    """
    if not candidate_name.startswith("~$"):
        return False
    remainder = candidate_name[2:]
    dropped = len(target_name) - len(remainder)
    return dropped in (0, 2) and target_name[dropped:] == remainder


def _parse_word_owner_name(data: bytes) -> str | None:
    """Best-effort extraction of the owner's display name from a Word
    owner-file's binary header.

    VERIFIED against a real Word-generated owner file (issue #28 WP-02
    review): opening a copy of a .docx from ~/Documents in Word 16.112.3
    created "~$oposal-Final.docx" for a source named "Proposal-Final.docx"
    — the first two characters replaced rather than prefixed, exactly as
    _word_owner_file_matches's docstring predicts for a base name over 8
    characters — and lock_status() against it correctly returned
    format="word", owner_name="Michael Sutton", sync_quiesced=true.
    Implemented against the commonly documented layout: a short header of
    padding/control bytes followed by the plain-text owner name with no
    explicit terminator. Decodes permissively (latin-1: the name is
    typically ASCII/Latin text) and returns the longest printable run of
    length >= 2, or None if nothing printable is found.
    """
    text = data.decode("latin-1", errors="replace")
    runs = [r.strip() for r in re.findall(r"[ -~ -ÿ]{2,}", text)]
    candidates = [r for r in runs if r and not r.isspace()]
    if not candidates:
        return None
    return max(candidates, key=len)


def _parse_libreoffice_owner_name(data: bytes) -> str | None:
    """Best-effort extraction of the owner's display name from a
    LibreOffice ``.~lock.<name>#`` file.

    STILL UNVERIFIED against a real LibreOffice-generated lock file — this
    is a separate gap from _parse_word_owner_name above, whose Word path
    WAS verified against real bytes in WP-02 review; that verification
    does not extend here. This machine has no LibreOffice install (D2
    declines LibreOffice only as a RENDER engine, not as a lock-file
    signal, but this repo still has no fixture to test against).
    Implemented against the commonly documented comma-separated layout
    (user name, user@host, host, a file:// URL, a timestamp); returns the
    second comma-separated field if present and non-empty, else the
    first, else None.
    """
    try:
        text = data.decode("utf-8", errors="replace").strip().rstrip(";")
    except Exception:  # noqa: BLE001
        return None
    fields = [f.strip() for f in text.split(",")]
    for field in (fields[1] if len(fields) > 1 else None, fields[0] if fields else None):
        if field:
            return field
    return None


def _find_owner_file(directory: Path, target_name: str) -> dict[str, Any]:
    """Return {"present", "path", "format", "owner_name"} — data only, per
    core/document-backend-protocol.md §4: this never gates a call by
    itself (DOCX_LOCKED gates writes only, and only in the write-guard
    landing in WP-04)."""
    lo_lock = directory / f".~lock.{target_name}#"
    if lo_lock.is_file():
        try:
            data = lo_lock.read_bytes()
        except OSError:
            data = b""
        return {
            "present": True,
            "path": str(lo_lock),
            "format": "libreoffice",
            "owner_name": _parse_libreoffice_owner_name(data),
        }

    for candidate in sorted(glob.glob(str(directory / "~$*"))):
        candidate_path = Path(candidate)
        if not candidate_path.is_file():
            continue
        if _word_owner_file_matches(candidate_path.name, target_name):
            try:
                data = candidate_path.read_bytes()
            except OSError:
                data = b""
            return {
                "present": True,
                "path": str(candidate_path),
                "format": "word",
                "owner_name": _parse_word_owner_name(data),
            }

    return {"present": False, "path": None, "format": None, "owner_name": None}


def _stat_sample(target: Path) -> dict[str, int]:
    st = target.stat()
    return {"size": st.st_size, "mtime_ns": st.st_mtime_ns}


def _find_tmp_siblings(directory: Path, target_name: str) -> list[str]:
    hits: list[str] = []
    for pattern in _TMP_SIBLING_GLOBS:
        for candidate in sorted(glob.glob(str(directory / pattern))):
            if Path(candidate).is_file():
                hits.append(candidate)
    return hits


def execute_lock_status(
    path: str,
    *,
    quiesce_interval: float = 1.5,
    sleep=time.sleep,
) -> dict[str, Any]:
    """Report owner-file and sync-quiesce state for *path*, as data only.

    Never refuses (core/document-backend-protocol.md §4: DOCX_LOCKED gates
    writes only, and that gate lives in the write guard landing in
    WP-04/WP-10 — lock_status itself only reports what it observed).
    """
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    directory = resolved.parent
    target_name = resolved.name

    owner_file = _find_owner_file(directory, target_name)

    sample_1 = _stat_sample(resolved)
    sleep(quiesce_interval)
    sample_2 = _stat_sample(resolved)
    tmp_siblings = _find_tmp_siblings(directory, target_name)

    sync_quiesced = sample_1 == sample_2 and not tmp_siblings

    return {
        "path": str(resolved),
        "owner_file": owner_file,
        "sync_quiesced": sync_quiesced,
        "sync_detail": {
            "sample_1": sample_1,
            "sample_2": sample_2,
            "interval_seconds": quiesce_interval,
            "tmp_siblings": tmp_siblings,
        },
    }


@mcp.tool()
def lock_status(path: str) -> dict[str, Any]:
    """Report Word/LibreOffice owner-file presence and sync-quiesce state.

    Data only — never refuses. Owner-file detection recognizes Word's
    "~$*" owner file (matched by shared filename suffix, not by
    constructing one exact expected name — see
    _word_owner_file_matches's docstring) and LibreOffice's
    ".~lock.<name>#". Sync-quiesce takes two (size, mtime_ns) samples
    ~1.5s apart and also checks for "*.tmp"/".~*" siblings in the same
    directory (OneDrive/iCloud staging artifacts); "sync_quiesced" is
    false if either signal suggests the file is still being written.

    Returns path, owner_file ({present, path, format, owner_name}),
    sync_quiesced, sync_detail ({sample_1, sample_2, interval_seconds,
    tmp_siblings}).

    Errors:
      INVALID_INPUT - path does not exist or is outside the allowed roots
    """
    try:
        return execute_lock_status(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Read tools (WP-03): list_parts, read_document, find_sections,
# list_page_sections, list_styles — all built on projection.py.
# ---------------------------------------------------------------------------


def _read_local_copy(resolved: Path) -> tuple[Path, bool]:
    """Return (path_to_read, is_temp) for *resolved* per
    core/document-backend-protocol.md §4's "reads never refuse" rule:
    when Word's (or LibreOffice's) owner file is present, read a validated
    snapshot instead of the live file (DOCX_LOCKED gates writes only — see
    lock_status's own docstring above). The caller is responsible for
    deleting the temp file (is_temp=True) once done, typically in a
    ``finally`` block.
    """
    owner_file = _find_owner_file(resolved.parent, resolved.name)
    if owner_file["present"]:
        snapshot = paths.snapshot_docx_package(resolved)
        return snapshot, True
    return resolved, False


_VALID_READ_FORMATS = frozenset({"markdown", "text", "runs"})


def execute_list_parts(path: str) -> dict[str, Any]:
    """List the package parts read_document(part=...) may target."""
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        return {"path": str(resolved), "parts": projection.list_parts_impl(local_path)}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def list_parts(path: str) -> dict[str, Any]:
    """List the package parts of a .docx that read_document(part=...) can
    target: the document body (word/document.xml), every header/footer,
    and footnotes/endnotes if present.

    Returns path, parts (list of {part, kind, header_footer_type}).
    header_footer_type ("default"|"first"|"even") is resolved via
    word/document.xml's own header/footer references where possible; None
    when it cannot be determined.

    Not gated by DOCX_LOCKED — reads a validated snapshot instead when
    Word's owner file is present (core/document-backend-protocol.md §4).

    Errors:
      INVALID_INPUT   - path does not exist or is outside the allowed roots
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_list_parts(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_read_document(
    path: str,
    format: str = "markdown",
    part: str = projection.DEFAULT_PART,
    section_key: str | None = None,
) -> dict[str, Any]:
    if format not in _VALID_READ_FORMATS:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"format must be one of {sorted(_VALID_READ_FORMATS)}, got {format!r}",
            {"format": format},
        )
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        revision = projection.compute_revision(local_path)
        result: dict[str, Any] = {
            "path": str(resolved),
            "part": part,
            "format": format,
            "section_key": section_key,
            "revision": revision["token"],
            "revision_detail": revision["detail"],
        }

        if section_key is not None:
            # Only a textbox-<n> sub-scope resolves here (WP-03 scope —
            # a heading section_key, listed by find_sections alongside
            # textbox ones, is not yet readable this way; that arrives
            # with a later WP). None means section_key named no text box
            # in this part — INVALID_INPUT with the keys that DO exist,
            # never a silent empty read.
            scoped = projection.project_textbox_scope(local_path, part, section_key)
            if scoped is None:
                available = [s["section_key"] for s in projection.iter_textbox_scopes(local_path, part)]
                raise _make_error(
                    ErrorCode.INVALID_INPUT,
                    f"section_key {section_key!r} does not name a text box in part {part!r}.",
                    {"section_key": section_key, "part": part, "available_textbox_keys": available},
                )
            if format == "text":
                result["text"] = scoped.text
            elif format == "runs":
                result["runs"] = projection.runs_from_projection(scoped)
            else:  # markdown
                markdown, _, lossy_elements = projection.markdown_from_projection(local_path, scoped)
                result["markdown"] = markdown
                if lossy_elements:
                    result["lossy_elements"] = lossy_elements
            result["warnings"] = scoped.warnings
            return result

        if format == "text":
            result["text"] = projection.read_document_text(local_path, part)
            result["warnings"] = projection.project_part(local_path, part).warnings
        elif format == "runs":
            result["runs"] = projection.read_document_runs(local_path, part)
            result["warnings"] = projection.project_part(local_path, part).warnings
        else:  # markdown
            markdown, warnings, lossy_elements = projection.read_document_markdown(local_path, part)
            result["markdown"] = markdown
            result["warnings"] = warnings
            # Same response shape as GoogleDocs-MCP's read_document (issue
            # #28 WP-03b-a): a lossy_elements key, present only when the
            # rendering actually lost something (a merged/nested table).
            if lossy_elements:
                result["lossy_elements"] = lossy_elements
        return result
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def read_document(
    path: str,
    format: str = "markdown",
    part: str = projection.DEFAULT_PART,
    section_key: str | None = None,
) -> dict[str, Any]:
    """Read a .docx part's content as markdown, flat text, or a run-level
    structural list.

    part scopes the read to one package part — word/document.xml's body
    (the default) is a completely separate scope from a header, footer,
    footnotes, or endnotes part. Call list_parts first to enumerate the
    parts actually present in this .docx.

    A text box's own w:txbxContent is a further, separate sub-scope never
    merged into its host part's text (see projection.py's module
    docstring) — call find_sections first to get its section_key
    ("textbox-<n>"), then pass that section_key here to read JUST that
    text box's content, scoped to it exactly as if it were its own part.
    Omitting section_key (the default) reads the whole part named by
    part, in which a text box's own text is absent (by design — it would
    otherwise be double-counted between the host part and the text box's
    own scope). section_key currently resolves a text-box scope only; a
    heading section_key (also listed by find_sections) is not yet
    readable this way.

    format="text" returns the flat projected string (runs concatenated in
    document order; a paragraph break is "\\n"). format="runs" returns, in
    document order, run records {text, rPr, para_ref} interleaved with
    structural records: {"type":"table_start"|"table_end", table_id},
    {"type":"drawing", blip_rid, media_part, extent_in, para_ref} for every
    w:drawing/a:blip, and {"type":"field", instr, result_text}.
    format="markdown" (default) renders headings, bold/italic, bulleted/
    numbered lists (nested by indent), and GFM pipe tables — matching
    GoogleDocs-MCP's markdown.py conventions exactly (issue #28 WP-03b-a)
    — plus stable placeholder tokens ("[image:rId]", "[field:instr]") for
    a drawing / a field with no result. WP-04's inverse (markdown ->
    OOXML) is not implemented here. A merged table cell (w:gridSpan/
    w:vMerge) or a nested w:tbl has no pipe-table representation and is
    reported in lossy_elements ({"kind": "table_merge"|"nested_table",
    "table_id"}) instead of silently reproduced as an ordinary grid cell.

    A deleted span (w:del/w:delText) and a field's own instruction text
    (w:instrText) are excluded from every format; a field's RESULT text
    (between fldChar "separate" and "end", or all of a w:fldSimple's
    nested runs) IS included — in format="markdown" it flows in as
    ordinary rendered text, never duplicated by a "[field:...]" token
    (that placeholder appears only when a field carries NO result at all).

    Returns path, part, format, section_key (echoed back, None unless
    passed), revision (the "<doc8>:<cmt8>" token), revision_detail (the
    full {document_sha256, comments_sha256, size, mtime_ns} tuple),
    warnings, lossy_elements (format="markdown" only, present only when
    non-empty), plus text|runs|markdown per format.

    Not gated by DOCX_LOCKED — see list_parts' docstring.

    Errors:
      INVALID_INPUT   - a bad path, an unrecognized format, or a
                         section_key that names no text box in this part
      PART_NOT_FOUND  - part names a package part absent from this .docx
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_read_document(path, format, part, section_key)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_find_sections(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        return {"path": str(resolved), "part": part, "sections": projection.find_sections_impl(local_path, part)}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def find_sections(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    """List heading-delimited section ranges AND text-box sub-scopes in a
    .docx part — this is the one call that discovers every section_key
    read_document(section_key=...) can then address.

    DISAMBIGUATION: this is about DOCUMENT SECTIONS (heading ranges, by
    style/outline level) — see list_page_sections for PAGE-LAYOUT sections
    (w:sectPr page size/margins/columns). The two are unrelated OOXML
    concepts that happen to share the English word "section".

    A heading paragraph's w:pStyle is mapped through list_styles to an
    outline level, falling back to the paragraph's own direct
    w:outlineLvl. A section runs from one heading (inclusive) to the next
    heading at any level, or the end of the document. section_key =
    slug(heading text) + a 1-based ordinal disambiguating duplicate
    headings. Headings inside a table cell are out of scope (not a real
    document section boundary).

    Every w:txbxContent in the part (see read_document's docstring) is
    also listed, each as its own entry with kind="textbox" and
    section_key="textbox-<n>" (heading_text/outline_level/start_para_ref/
    end_para_ref are None on these entries; paragraph_count is still the
    text box's own paragraph count). Pass one of these section_keys to
    read_document to read that text box's content.

    Returns path, part, sections (list of {section_key, kind
    ("heading"|"textbox"), heading_text, outline_level, start_para_ref,
    end_para_ref, paragraph_count}).

    Not gated by DOCX_LOCKED — see list_parts' docstring.

    Errors:
      INVALID_INPUT   - a bad path
      PART_NOT_FOUND  - part names a package part absent from this .docx
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_find_sections(path, part)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_list_page_sections(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        return {
            "path": str(resolved),
            "part": part,
            "page_sections": projection.list_page_sections_impl(local_path, part),
        }
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def list_page_sections(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    """List w:sectPr PAGE-LAYOUT sections in a .docx part.

    DISAMBIGUATION: this is about PAGE-LAYOUT sections (page size,
    margins, column widths) — see find_sections for DOCUMENT sections
    (heading ranges). The two are unrelated OOXML concepts that happen to
    share the English word "section".

    Page size, margins, and text-column widths are reported in inches,
    converted from OOXML's own unit for page geometry — twentieths of a
    point (dxa), NOT EMU (dxa: 1440 = 1 inch; EMU, used for a drawing's own
    extent instead, is 914400 = 1 inch — see projection.py's
    list_page_sections_impl for the note on why these are not the same
    conversion).

    Returns path, part, page_sections (list of {page_width_in,
    page_height_in, orientation, margin_top_in, margin_bottom_in,
    margin_left_in, margin_right_in, column_widths_in}).

    Not gated by DOCX_LOCKED — see list_parts' docstring.

    Errors:
      INVALID_INPUT   - a bad path
      PART_NOT_FOUND  - part names a package part absent from this .docx
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_list_page_sections(path, part)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_list_styles(path: str) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        return {"path": str(resolved), "styles": projection.list_styles_impl(local_path)}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def list_styles(path: str) -> dict[str, Any]:
    """Enumerate word/styles.xml's styles: ids, names, types, outline
    levels.

    find_sections' heading detection and read_document(format="markdown")'s
    heading rendering both resolve a paragraph's outline level through
    this same list (style-level w:outlineLvl, falling back to a
    "Heading1".."Heading9" style-id pattern) — call this tool directly
    when you need to know what styles a .docx actually defines before
    targeting one.

    Returns path, styles (list of {style_id, name, type, outline_lvl}).

    Not gated by DOCX_LOCKED — see list_parts' docstring.

    Errors:
      INVALID_INPUT   - path does not exist or is outside the allowed roots
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_list_styles(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Markdown mutations (WP-04): replace_body_markdown, replace_range_markdown,
# append_markdown. All three are in middleware.MUTATING_TOOLS (added in the
# same commit) and are gated by mutations._guard_before_write BEFORE any
# temp file is written — DOCX_LOCKED / SYNC_IN_FLIGHT / REVISION_CONFLICT —
# unlike every WP-03 read tool above, which never refuses for a lock
# (core/document-backend-protocol.md §4: "reads never refuse"; the
# reads-vs-writes asymmetry is the point of this WP, not a regression of
# it — see mutations.py's module docstring).
#
# Guard layers 3-4 (issue #28 WP-10, landed in this same PR):
# mutations.atomic_replace_docx_parts now wraps every write in
# acquire_lock/release_lock (layer 0's remote_checkout no-op seam for a
# future Microsoft Graph checkout, plus layer 4's same-machine ``.jsclaim``
# O_EXCL mutex) and runs conflict_copy_sweep (layer 3) after a successful
# write, surfacing conflict_copy_detected/conflict_copies/
# sibling_files_changed on every mutating tool's own evidence dict (never
# raised — the write already succeeded). What is still true, and always
# will be per core/document-backend-protocol.md §9's own framing
# ("detection-only, not prevention"): a write that lands while another
# client has the file open is not PREVENTED by anything above, only
# detected after the fact by the sweep; the naming patterns the sweep
# matches on are themselves client- and locale-dependent, so their
# absence is not proof no conflict occurred.
# ---------------------------------------------------------------------------


@mcp.tool()
def replace_body_markdown(
    path: str, markdown: str, revision_before: str | None = None, force: bool = False, track_changes: bool = False
) -> dict[str, Any]:
    """Replace an entire document body's content with markdown, atomically.

    Rung 4 (last resort) of core/document-backend-protocol.md's backend-
    neutral edit ladder — the whole-document rewrite. Prefer
    replace_range_markdown when find_sections can locate the target
    section.

    Markdown is rendered against THIS document's own styles: headings
    resolve through list_styles (STYLE_NOT_FOUND if a level has no
    matching style — never a hardcoded "Heading2"), tables pick up the
    document's own table style if one exists, and every bulleted/ordered
    list gets a freshly allocated numbering definition. Supports
    bold/italic/***both***, links, lists (nested), and pipe tables; a
    thematic break, blockquote, or code fence degrades gracefully (see
    markdown_to_ooxml's module docstring) rather than failing the call.

    Guard order (this WP's whole point — Codex defect: an earlier draft
    let writes ship before the guard): lock_status runs FIRST, before any
    temp file exists. An owner file present -> DOCX_LOCKED (report the
    owner, no retry). Not sync-quiesced -> one bounded wait (<=10s total)
    then SYNC_IN_FLIGHT. A revision_before that no longer matches the
    file's current revision -> REVISION_CONFLICT. If the CURRENT body
    contains comment anchors or tracked changes, the call refuses
    (COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT) unless
    force=True; with force, they are removed and their comment ids are
    reported as orphaned_comment_ids (nothing disappears silently ahead
    of WP-08's comment tools).

    The write itself is atomic: a temp file is built and OPC-validated
    (XML well-formed, every r:id resolves, [Content_Types].xml covers
    every part) BEFORE the original is ever touched (a failure here ->
    OPC_INVALID, original untouched); the original is then swapped in via
    os.replace() with a .jsbak kept until a post-write re-read/re-project
    confirms the change (a failure here restores from .jsbak and raises
    VERIFICATION_FAILED).

    Residual risk (layers 3-4, issue #28 WP-10, wrap the atomic write
    itself -- see the module-level comment above): a write that lands
    while another client has the file open is not PREVENTED by anything
    above, only detected after the fact, by the post-write conflict-copy
    sweep on THIS SAME call (conflict_copy_detected in the evidence
    below) or a later one.

    track_changes=True (issue #28 WP-07b-a): the OLD body content is kept
    (not removed) with its runs wrapped in w:del/w:delText, and the NEW
    content is wrapped in w:ins -- both carrying w:author (author.
    resolve_author_name())/w:date/an above-package-maximum w:id. Reading
    the projection is unaffected either way, so after_text still reads as
    the new content alone. Scope limit: this wraps RUNS, not whole table/
    list structures -- untested against a target range containing a
    table. The comment/tracked-change hazard scan's TRACKED_CHANGES_PRESENT
    excludes a revision authored by the configured author (the server's
    own prior tracked write), so a second track_changes call never
    deadlocks on the first one's own tracked change.

    Returns the eight evidence keys: applied, match_count (always 1),
    rung (4), before, after, revision_before, revision_after,
    audit_logged; plus orphaned_comment_ids when force removed anchors,
    and (track_changes=True only) revision_ids/track_changes.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (see above)
      REVISION_CONFLICT                 - revision_before is stale
      COMMENT_ANCHORS_IN_RANGE          - comment anchors present, no force
      TRACKED_CHANGES_PRESENT           - w:ins/w:del present, no force
      STYLE_NOT_FOUND                   - a heading level has no style
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return mutations.execute_replace_body_markdown(
            path, markdown, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def replace_range_markdown(
    path: str,
    section_key: str,
    markdown: str,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Replace one heading-delimited section (by section_key, from
    find_sections) with markdown, atomically.

    Rung 3 of core/document-backend-protocol.md's edit ladder: use this
    when find_sections can locate the target section, ahead of
    replace_body_markdown's whole-document rewrite. section_key is keyed
    by heading slug rather than a revision-stamped range — STALE_RANGE is
    a gdoc-only hazard (the protocol doc's footnote on this rung).

    Same guard, hazard scan, atomic-write, and markdown rendering rules as
    replace_body_markdown — see that tool's docstring for the full guard
    order, hazard-refusal/force semantics, and atomic-write mechanics; the
    only difference is the target range (one section's top-level body
    children, from its own heading up to but not including the next
    top-level heading of any level, or the end of the document) instead
    of the whole body.

    track_changes=True: same as replace_body_markdown (WP-07b-a) -- the
    section's OLD content is kept (w:del-wrapped), the NEW content is
    wrapped in w:ins and inserted right after it.

    Returns the eight evidence keys (rung is always 3 here), plus
    orphaned_comment_ids when force removed anchors, and (track_changes=
    True only) revision_ids/track_changes.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors: as replace_body_markdown, plus:
      SECTION_NOT_FOUND - section_key does not match any current section
                           (call find_sections to see the current ones;
                           not named in the issue #28 plan text for this
                           WP, added because replace_range_markdown needs
                           some code for this case)
    """
    try:
        return mutations.execute_replace_range_markdown(
            path, section_key, markdown, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def append_markdown(
    path: str, markdown: str, revision_before: str | None = None, force: bool = False, track_changes: bool = False
) -> dict[str, Any]:
    """Append markdown to the end of a document body (before its trailing
    w:sectPr, if any), atomically.

    Nothing existing is removed, so there is no comment-anchor/tracked-
    change hazard to scan and force has no effect here — it is accepted
    only so a caller's generic retry code can pass it uniformly across
    all three mutating tools. Same guard order (lock_status first, then
    revision_before) and atomic-write mechanics as replace_body_markdown
    — see that tool's docstring.

    track_changes=True (WP-07b-a): nothing existing is removed by an
    append, so there is no w:del side here -- only the newly appended
    content is wrapped in w:ins (w:author/w:date/w:id, same source as
    replace_text's).

    Returns the eight evidence keys: before/after are the whole body's
    markdown immediately before/after the append; rung is reported as 4
    (append_markdown is not itself a rung on
    core/document-backend-protocol.md's 4-rung table, which only names
    replace_body_markdown at rung 4 — grouped with it here as the other
    whole-document-scoped write). Plus (track_changes=True only)
    revision_ids/track_changes.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      STYLE_NOT_FOUND                   - a heading level has no style
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return mutations.execute_append_markdown(
            path, markdown, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Targeted text edits (WP-06): replace_text, format_text. Both are in
# middleware.MUTATING_TOOLS (added in the same commit) and share
# text_edit.py's guard (mutations._guard_before_write, same as the WP-04
# tools above) and locate.locate()'s normalization ladder + STRUCTURAL_
# BOUNDARY refusal. expected_matches is REQUIRED here (no default) --
# issue #28 plan ruling D4, this server only: the shipped GoogleDocs-MCP
# server defaults it to 1 (its server.py replace_text/format_text
# signatures) but this server's contract deliberately does not carry that
# default across (every caller must say how many matches it expects).
#
# WP-06 only WARNS on a match crossing a comment range or a tracked change
# (the "warnings" evidence key, when non-empty) -- it does not refuse.
# WP-07 (tracked_changes.py) adds the actual TRACKED_CHANGES_PRESENT
# refusal on top of this same detection.
# ---------------------------------------------------------------------------


@mcp.tool()
def replace_text(
    path: str,
    find: str,
    replace: str,
    expected_matches: int,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Replace every occurrence of `find` with `replace`, atomically.

    Locates `find` via a 4-rung normalization ladder (exact -> curly/
    straight quote equivalence -> NBSP/whitespace-run collapse -> soft-
    hyphen strip), stopping at the first rung with at least one match. A
    match overlapping a RESULT-bearing field's cached text surfaces
    "touches_field_result" in the evidence's `warnings` (see locate.py's
    own header comment for why this is a warning, not a match rung, on
    the corrected field-markdown projection).

    expected_matches is REQUIRED (no default, D4): the call refuses with
    MATCH_COUNT_MISMATCH if the actual count differs, listing every span
    found at the matching rung.

    A run whose text a match's boundary falls in the middle of splits into
    up to three pieces: the unmatched prefix and suffix keep the run's
    ORIGINAL w:rPr (cloned verbatim onto a fresh sibling run when a suffix
    survives); exactly one new run is inserted for `replace`, its own
    w:rPr inherited from the FIRST run the match touched. A run entirely
    inside the match is removed outright. A match crossing a w:p/w:tbl/
    w:tc boundary refuses with STRUCTURAL_BOUNDARY instead.

    track_changes=True (issue #28 WP-07b-a): the deleted text is wrapped
    in w:del/w:delText and the replacement in w:ins, each carrying
    w:author (author.resolve_author_name() -- ~/.jennystack/config.json's
    author_name, falling back to the macOS full name) and w:date, with
    w:id values allocated above the package maximum. Reading the
    projection is unaffected (w:del text stays excluded, w:ins text stays
    included), so the write is visible to the next read as current text
    either way. The evidence then also carries revision_ids (the ids just
    created) and track_changes: true. A match crossing a tracked change
    authored by someone OTHER than the configured author refuses with
    TRACKED_CHANGES_PRESENT unless force=True; a match crossing the
    server's OWN prior tracked change never refuses (accept/reject it via
    tracked_changes.py, or just keep editing under track_changes).

    Same atomic-write mechanics as replace_body_markdown (temp file, OPC-
    validated before the original is touched, .jsbak-backed post-write
    verification) -- see that tool's docstring.

    Returns the eight evidence keys (before/after are ±200-character
    excerpts around the first match, not the whole document), plus
    runs_before/runs_after (each match's overlapping run(s), clipped to the
    span, before and after), `warnings` when non-empty, and (track_changes=
    True only) revision_ids/track_changes.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or an empty find
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      ZERO_MATCH                        - find not located after the full ladder
      MATCH_COUNT_MISMATCH              - the located count != expected_matches
      STRUCTURAL_BOUNDARY               - a match crosses a w:p/w:tbl/w:tc boundary
      TRACKED_CHANGES_PRESENT           - a match crosses a FOREIGN-authored tracked change; no force
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return text_edit.execute_replace_text(
            path, find, replace, expected_matches, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def format_text(
    path: str,
    find: str,
    style: dict[str, bool],
    expected_matches: int,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Apply character styling (bold/italic/underline/strike) to a matched
    text span, without touching its content.

    style maps any of "bold"/"italic"/"underline"/"strike" to true/false;
    every requested field's value is applied verbatim (including false, so
    {"bold": false} actually clears bold). Same locate()/expected_matches
    contract as replace_text (see that tool's docstring for the
    normalization ladder and STRUCTURAL_BOUNDARY refusal).

    Idempotent: if every located run already carries every requested
    field's value, the call skips the write entirely (revision_before ==
    revision_after in the returned evidence) rather than creating a new,
    no-op revision -- with track_changes=True too: nothing to record in a
    w:rPrChange when no style is actually changing.

    Same run-splitting rule as replace_text for a boundary run (up to
    three pieces; unmatched prefix/suffix keep the ORIGINAL w:rPr cloned
    verbatim), except the matched middle SURVIVES here (as a new run
    carrying the requested style) rather than being deleted -- this tool
    never changes character counts.

    track_changes=True (issue #28 WP-07b-a): the changed run's w:rPr gains
    a w:rPrChange recording its PRE-change formatting (w:author/w:date/
    w:id, same source and allocation as replace_text's track_changes).
    Same own-author-vs-foreign-author TRACKED_CHANGES_PRESENT refusal
    rule as replace_text.

    Returns the eight evidence keys, plus runs_before/runs_after (each
    match's overlapping run(s) and their style flags, before and after),
    `warnings` when non-empty, and (track_changes=True only) revision_ids/
    track_changes.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors: as replace_text, plus:
      INVALID_INPUT - style is empty, not an object, names an unknown key,
                       or a value is not a literal true/false boolean
    """
    try:
        return text_edit.execute_format_text(
            path, find, style, expected_matches, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tracked changes (WP-07): list_open_items (read-only), accept_tracked_
# changes / reject_tracked_changes (mutating, in MUTATING_TOOLS). Also
# where replace_text/format_text's own TRACKED_CHANGES_PRESENT refusal
# (text_edit._check_tracked_changes_guard) is wired in -- see that
# function and each tool's own updated docstring above.
# ---------------------------------------------------------------------------


@mcp.tool()
def list_open_items(path: str) -> dict[str, Any]:
    """List every open comment and pending tracked change (w:ins/w:del) in
    a .docx.

    Returns `comments` (from word/comments.xml + commentsExtended.xml's
    resolved flag; each with comment_id, content, resolved, reply_count,
    replies, quoted_text, author, created_time, modified_time, scope --
    the same shape the GoogleDocs-MCP server returns, "scope": "document"
    always since a docx comment anchor is not tab-scoped) and
    `pending_suggestions` (every w:ins/w:del, with suggestion_id, kind
    ("insertion"|"deletion"), text, author, date, anchor_context -- the
    owning paragraph's own live text).

    Scope limit: comment REPLY THREADING (commentsExtended's parent/child
    linking) is not resolved here -- every comment reports reply_count=0/
    replies=[]. reply_count/replies are still present, structurally, for
    forward compatibility with a later WP that resolves them.

    Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
    Word's owner file is present, like every other read tool.

    Errors:
      INVALID_INPUT   - path does not exist or is outside the allowed roots
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return tracked_changes.execute_list_open_items(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def accept_tracked_changes(
    path: str, revision_ids: list[str] | None = None, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    """Accept tracked changes (w:ins/w:del), atomically -- all of them, or
    only the ids named in revision_ids (from list_open_items'
    pending_suggestions[].suggestion_id).

    Accepting a w:ins makes its inserted text ordinary, permanent content
    (the wrapper is removed, the text stays). Accepting a w:del makes the
    deletion permanent (the w:del element and its w:delText content are
    removed outright).

    Same guard/atomic-write mechanics as replace_text (lock_status first,
    then revision_before; a temp file OPC-validated before the original is
    touched; .jsbak-backed post-write verification).

    Returns the eight evidence keys (before/after are the WHOLE document's
    plain text, not an excerpt -- there is no single match span here;
    rung is "all" when revision_ids is omitted, else "by_id"; match_count
    is the number of w:ins/w:del elements processed), plus revision_ids
    (the ids actually processed).

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      REVISION_ID_NOT_FOUND             - a named id is not present (available ids listed)
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return tracked_changes.execute_accept_tracked_changes(
            path, revision_ids, revision_before=revision_before, force=force
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def reject_tracked_changes(
    path: str, revision_ids: list[str] | None = None, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    """Reject tracked changes (w:ins/w:del), atomically -- all of them, or
    only the ids named in revision_ids.

    Rejecting a w:ins undoes the insertion (the element and its content
    are removed outright). Rejecting a w:del undoes the deletion (every
    w:delText inside it is renamed back to w:t and the w:del wrapper is
    removed, so the previously-deleted text becomes live again).

    Same guard/atomic-write mechanics and evidence shape as
    accept_tracked_changes -- see that tool's docstring.

    Errors: as accept_tracked_changes.
    """
    try:
        return tracked_changes.execute_reject_tracked_changes(
            path, revision_ids, revision_before=revision_before, force=force
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Comments (WP-08): add_anchored_comment (mutating, in MUTATING_TOOLS),
# get_comment_thread (read-only). Reply threading and resolve are WP-09,
# not this WP -- get_comment_thread reads whatever threading a document
# already carries (e.g. authored in Word desktop); it never creates a
# reply.
# ---------------------------------------------------------------------------


@mcp.tool()
def add_anchored_comment(path: str, quote: str, text: str, expected_matches: int) -> dict[str, Any]:
    """Add a comment anchored to a quoted passage, atomically.

    Locates `quote` via the same locate()/expected_matches contract as
    replace_text (normalization ladder, STRUCTURAL_BOUNDARY refusal,
    MATCH_COUNT_MISMATCH on the wrong count) -- see that tool's docstring.
    Every matched span gets its own new comment (same `text`, separate
    w:id/paraId/durableId each) when expected_matches > 1.

    Builds all five comment-related package parts a real Word comment
    needs (word/comments.xml, commentsExtended.xml, commentsIds.xml,
    commentsExtensible.xml, people.xml -- created fresh on a document's
    first-ever comment, otherwise appended to) plus the anchor itself in
    word/document.xml (w:commentRangeStart/End and a w:commentReference
    run) and the matching [Content_Types].xml override + word/_rels/
    document.xml.rels relationship for each newly created part.
    w:id = one past the highest existing comment id (0 for a document's
    first comment); paraId/durableId = a random 8-hex-uppercase value
    below 0x80000000, unique within the package. Author (w:author, and
    word/people.xml's w:15:person) comes from author.resolve_author_name()
    (~/.jennystack/config.json's author_name, falling back to the macOS
    full name).

    A run whose text the quote's boundary falls in the middle of splits
    the same way replace_text/format_text's boundary runs do (ORIGINAL
    w:rPr cloned verbatim onto every surviving piece) -- but nothing here
    wraps content the way w:ins/w:del does: commentRangeStart/End are
    self-closing position markers, so a boundary run only ever needs
    splitting to expose a clean insertion point, never any content change.

    After the write, the target file is re-read from disk and the
    anchored range is confirmed to bracket the requested quote (modulo
    whitespace) before the call returns -- a mismatch raises
    VERIFICATION_FAILED via the same .jsbak-backed atomic-write path
    every other mutating tool in this server uses.

    Returns the eight evidence keys (before/after are unchanged -- a
    comment never edits document text), plus comment_ids (every durableId
    created, one per matched span) and comment_id (the singular durableId,
    only when exactly one span matched).

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or an empty quote
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      ZERO_MATCH                        - quote not located after the full ladder
      MATCH_COUNT_MISMATCH              - the located count != expected_matches
      STRUCTURAL_BOUNDARY               - a match crosses a w:p/w:tbl/w:tc boundary
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return comments.execute_add_anchored_comment(path, quote, text, expected_matches)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def get_comment_thread(path: str, comment_id: str) -> dict[str, Any]:
    """Read a comment and its direct replies, by durableId (from
    add_anchored_comment's own evidence, or word/commentsIds.xml's
    w16cid:durableId directly).

    Returns comment_id, content, author, created_time, resolved
    (commentsExtended.xml's w15:done), quoted_text (the live text its
    range currently brackets), reply_count, and replies (each reply in
    the same shape, one level deep -- a reply-of-reply is not walked
    further).

    Scope limit: if comment_id itself names a reply (it has its own
    w15:paraIdParent), this returns that reply alone with replies=[] --
    it does not walk upward to find and return the whole thread's root.
    Reply creation and resolve are reply_to_comment/resolve_comment; this
    tool only ever reads whatever threading state already exists.

    Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
    Word's owner file is present, like every other read tool.

    Errors:
      INVALID_INPUT   - path does not exist or is outside the allowed roots,
                         the document has no comments at all, or comment_id
                         does not match any commentsIds.xml durableId
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return comments.execute_get_comment_thread(path, comment_id)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def reply_to_comment(path: str, comment_id: str, text: str) -> dict[str, Any]:
    """Reply to an existing comment (durableId), atomically -- issue #28
    WP-09.

    Creates a NEW, independent comment (its own w:comment / w:id / paraId
    / durableId, and its own full commentRangeStart/End/commentReference
    anchor triplet in word/document.xml) whose commentsExtended.xml entry
    carries w15:paraIdParent pointing at the PARENT's own paraId -- this
    is exactly how a real Word-authored reply is shaped (verified against
    tests/fixtures/comments/golden-comment.docx's own real reply). The
    reply's anchor brackets the SAME live text the parent's own anchor
    currently does; no new text is located or matched.

    Returns the eight evidence keys (before/after are the parent's own
    quoted text, unchanged -- a reply never edits document text; rung is
    the fixed label "reply", since no text search is performed), plus
    comment_id (the new reply's own durableId) and parent_comment_id.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or
                         comment_id does not match any existing comment,
                         or that comment has no live anchor to reply against
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return comments.execute_reply_to_comment(path, comment_id, text)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def resolve_comment(path: str, comment_id: str) -> dict[str, Any]:
    """Resolve a comment thread (durableId), atomically -- issue #28
    WP-09. Sets w15:done="1" on the comment's own commentsExtended.xml
    entry; list_open_items excludes it afterward (it is no longer an
    "open" item), but get_comment_thread still fetches it by id --
    resolved is not deleted, only marked.

    COMMENT_STILL_OPEN is adapted, not lifted verbatim, from
    GoogleDocs-MCP's own member of the same name: that server's version
    guards genuine Drive-API eventual consistency (a resolve action that
    does not durably stick server-side) by re-querying the comment from
    the API after the write. This backend's write is a local, synchronous,
    atomically-verified file replace with no such external-consistency
    hazard -- so this code can only ever fire here via a bug in this
    server's own code, not a real runtime race. Kept anyway (re-reading
    the fresh file from disk and checking w15:done="1" independently of
    the write's own success) so the error VOCABULARY still matches
    Google's for this exact failure mode.

    Idempotent: resolving an already-resolved comment succeeds again
    (not an error).

    Returns the eight evidence keys (before="open", after="resolved";
    rung is the fixed label "resolve", since no text search is
    performed), plus comment_id.

    Always also carries conflict_copy_detected (issue #28 WP-10's
    post-write conflict-copy sweep -- never raised, an evidence flag on
    an already-successful write); conflict_copies/sibling_files_changed
    are added only when non-empty.

    Errors:
      INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or
                         comment_id does not match any existing comment
      DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
      REVISION_CONFLICT                 - revision_before is stale
      COMMENT_STILL_OPEN                - the post-write re-read did not confirm w15:done="1" (see this tool's own docstring)
      OPC_INVALID                       - the rendered .docx failed OPC validation
      VERIFICATION_FAILED               - post-write verification failed; rolled back
    """
    try:
        return comments.execute_resolve_comment(path, comment_id)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tables (issue #28 WP-14): list_tables, get_table (read-only), and
# replace_table_row / replace_cell_markdown / insert_table (mutating, in
# MUTATING_TOOLS). table_id/row_index/cell_index are 1-based and match
# read_document(format="runs")'s own container_chain addressing.
# ---------------------------------------------------------------------------


@mcp.tool()
def list_tables(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    """Enumerate every w:tbl in a .docx part, including tables nested
    inside a cell (each gets its own table_id, in document order).

    Returns path, part, tables (list of {table_id, row_count, col_count,
    has_merged_cells, has_nested_table, nested_in_table_id}). table_id is
    1-based, document order, the same numbering read_document(format=
    "runs")'s table_start/table_end records and container_chain use.

    Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
    Word's owner file is present (core/document-backend-protocol.md §4).

    Errors:
      INVALID_INPUT  - path does not exist or is outside the allowed roots
      PART_NOT_FOUND - part names a package part absent from this .docx
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return tables.execute_list_tables(path, part)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def get_table(path: str, table_id: int, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    """Full row/cell detail for one table, addressed by the table_id
    list_tables reports.

    Returns path, part, table_id, row_count, col_count, has_merged_cells,
    has_nested_table, rows (list of list of {row_index, cell_index,
    grid_span, v_merge, text}). grid_span is w:tcPr/w:gridSpan (1 when
    absent); v_merge is "none"/"restart"/"continue" from w:tcPr/w:vMerge
    (a continuation cell's own text is always ""). text is the same
    per-cell markdown rendering read_document(format="markdown") uses for
    a pipe-table cell (a nested w:tbl inside a cell flattens to inline
    text, same as there).

    Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
    Word's owner file is present.

    Errors:
      INVALID_INPUT   - path does not exist or is outside the allowed roots
      PART_NOT_FOUND  - part names a package part absent from this .docx
      TABLE_NOT_FOUND - table_id does not match any table (call list_tables)
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return tables.execute_get_table(path, table_id, part)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def replace_table_row(
    path: str,
    table_id: int,
    row_index: int,
    cells: list[str],
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Replace one table row's cell content wholesale, one markdown string
    per cell (cells must have exactly as many entries as the row has
    cells).

    Refuses -- for the WHOLE table, not just this row -- the moment
    table_id names a table containing a merged cell (w:gridSpan != 1 or a
    w:vMerge) or a nested w:tbl anywhere in it, mirroring
    GoogleDocs-MCP's own replace_table_row refusal on a merged cell.
    replace_cell_markdown is the escape hatch for that case: a cell-scoped
    write that never touches w:tcPr.

    Same guard (lock/sync/revision, before any temp file), comment-anchor/
    tracked-change hazard scan (force=True to proceed; COMMENT_ANCHORS_IN_RANGE
    / TRACKED_CHANGES_PRESENT otherwise), atomic write with OPC validation
    and a .jsbak rollback, and audit log as every other mutating tool.

    track_changes=True wraps the row's OLD cell content in w:del and the
    NEW content in w:ins (per cell), under the configured author, rather
    than rewriting directly.

    Returns the eight evidence keys (before/after are the row's cells
    tab-joined), plus (track_changes=True only) revision_ids/track_changes,
    and conflict_copy_detected (+ conflict_copies/sibling_files_changed
    when non-empty).

    Errors:
      INVALID_INPUT          - a bad path, or cells' length != the row's own cell count
      DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
      REVISION_CONFLICT      - revision_before is stale
      TABLE_NOT_FOUND         - table_id does not match any table
      TABLE_ROW_NOT_FOUND     - row_index out of range for this table
      MERGED_OR_NESTED_TABLE  - this table has a merged/nested-table cell; use replace_cell_markdown
      COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT - a hazard in the row; force=True to proceed
      OPC_INVALID             - the rendered .docx failed OPC validation
      VERIFICATION_FAILED     - post-write verification failed; rolled back
    """
    try:
        return tables.execute_replace_table_row(
            path, table_id, row_index, cells, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def replace_cell_markdown(
    path: str,
    table_id: int,
    row_index: int,
    cell_index: int,
    markdown: str,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Replace one table cell's content with rendered markdown, leaving
    its own w:tcPr byte-identical -- the only write path safe on a merged
    (w:gridSpan/w:vMerge) cell, and the intended path for an Appendix-A
    style band-and-border table's own cell content.

    Supports multi-level bulleted/numbered markdown inside the cell
    (issue #28 WP-16a): rendered through the SAME markdown_to_ooxml
    machinery replace_body_markdown/append_markdown use, one numId/
    abstractNum tree per top-level list shared across nesting levels via
    increasing w:ilvl.

    Same guard/hazard-scan/atomic-write/audit machinery as
    replace_table_row. track_changes=True wraps the OLD cell content in
    w:del and the NEW content in w:ins, under the configured author.

    Returns the eight evidence keys (before/after are the cell's own
    markdown), plus (track_changes=True only) revision_ids/track_changes,
    and conflict_copy_detected (+ conflict_copies/sibling_files_changed
    when non-empty).

    Errors:
      INVALID_INPUT           - a bad path
      DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
      REVISION_CONFLICT       - revision_before is stale
      TABLE_NOT_FOUND          - table_id does not match any table
      TABLE_ROW_NOT_FOUND      - row_index out of range for this table
      TABLE_CELL_NOT_FOUND     - cell_index out of range for this row
      COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT - a hazard in the cell; force=True to proceed
      OPC_INVALID              - the rendered .docx failed OPC validation
      VERIFICATION_FAILED      - post-write verification failed (including w:tcPr drift); rolled back
    """
    try:
        return tables.execute_replace_cell_markdown(
            path, table_id, row_index, cell_index, markdown,
            revision_before=revision_before, force=force, track_changes=track_changes,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def insert_table(
    path: str,
    rows: list[list[str]],
    style_id: str,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Append a new table at the end of the document body, one markdown
    string per cell (rows is a list of rows, each a list of cell markdown;
    short rows are padded with empty cells up to the widest row).

    style_id is REQUIRED and must name an existing w:type="table" style in
    this document's styles.xml (list_styles reports style type) -- unlike
    GoogleDocs-MCP's insert_table, which relies on the Docs API's own
    default table style, raw OOXML has no sensible default to fall back
    to. Column widths are dxa, split evenly across the text-column width
    list_page_sections reports for this document's first page section.

    Same guard/atomic-write/audit machinery as append_markdown.
    track_changes=True wraps the new table's own runs in w:ins (nothing
    existing is removed by an insertion, so there is no w:del side).

    Returns the eight evidence keys (before="", after is the new table's
    rows tab/newline-joined), plus table_id (the new table's own id, for a
    follow-up get_table/replace_table_row/replace_cell_markdown call),
    (track_changes=True only) revision_ids/track_changes, and
    conflict_copy_detected (+ conflict_copies/sibling_files_changed when
    non-empty).

    Errors:
      INVALID_INPUT      - a bad path, or rows is empty or contains an empty row
      DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
      REVISION_CONFLICT  - revision_before is stale
      STYLE_NOT_FOUND    - style_id does not name a table style in this document
      OPC_INVALID        - the rendered .docx failed OPC validation
      VERIFICATION_FAILED - post-write verification failed; rolled back
    """
    try:
        return tables.execute_insert_table(
            path, rows, style_id, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Images (issue #28 WP-15a): insert_image, apply_style (both mutating, in
# MUTATING_TOOLS), read_header_footer (read-only).
# ---------------------------------------------------------------------------


@mcp.tool()
def insert_image(
    path: str,
    image_path: str,
    width_in: float | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Append a new inline picture at the end of the document body, from a
    LOCAL .png or .svg file (image_path is read natively -- no
    IMAGE_SOURCE_UNSUPPORTED, unlike GoogleDocs-MCP's URL-only tool).

    width_in defaults to the text-column width list_page_sections reports
    for this document's first page section (its own height follows the
    image's own aspect ratio). An .svg is embedded natively (Word 2016+'s
    own a:blip/a:extLst SVG extension) WITH a PNG fallback part Word
    requires for any older/non-SVG-aware consumer -- rasterized from the
    SVG's own native pixel size via macOS `sips`.

    Records design_width_in/design_height_in (a PNG's own intrinsic pixel
    size / 96 DPI; an SVG's own root width/height or viewBox),
    placed_width_in/placed_height_in (what was actually written to
    wp:extent), and effective_scale = placed_width_in / design_width_in --
    a real measured ratio, consumed downstream by JennyStack's own
    readability gate, not a placeholder.

    Same guard/atomic-write/audit machinery as insert_table.
    track_changes=True wraps the new run in w:ins (nothing existing is
    removed by an insertion).

    Returns the eight evidence keys (before="", after="[image:<rId>]"),
    plus format, design_width_in, design_height_in, placed_width_in,
    placed_height_in, effective_scale, blip_rid (svg_rid when the source
    was an .svg), (track_changes=True only) revision_ids/track_changes,
    and conflict_copy_detected (+ conflict_copies/sibling_files_changed
    when non-empty).

    Errors:
      INVALID_INPUT           - a bad path, image_path is not a regular file, or width_in <= 0
      DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
      REVISION_CONFLICT       - revision_before is stale
      UNSUPPORTED_IMAGE_FORMAT - image_path is not .png/.svg, or its bytes do not parse as one
      SVG_RASTERIZATION_FAILED - the PNG fallback part could not be produced (macOS `sips` failed/unavailable)
      OPC_INVALID              - the rendered .docx failed OPC validation
      VERIFICATION_FAILED      - post-write verification failed; rolled back
    """
    try:
        return images.execute_insert_image(
            path, image_path, width_in, revision_before=revision_before, force=force, track_changes=track_changes
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def apply_style(
    path: str,
    find: str,
    style_id: str,
    expected_matches: int,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Apply a NAMED style (from list_styles) to text located via find --
    the named-style counterpart to format_text's four boolean toggles.

    A CHARACTER style (w:type="character") applies to the matched run(s)
    exactly like format_text (same locate()/normalization-ladder/
    STRUCTURAL_BOUNDARY contract, same run-splitting rule, same
    track_changes=True support via w:rPrChange). A PARAGRAPH style
    (w:type="paragraph") applies w:pStyle to every paragraph CONTAINING a
    matched run (the whole paragraph, not just the matched substring) --
    but does NOT support track_changes=True: WP-07b-a's contract defines
    w:rPrChange for run formatting only, never a paragraph-level tracked
    change, so a paragraph-style call with track_changes=True raises
    INVALID_INPUT naming this gap rather than silently applying it
    untracked or inventing an untested w:pPrChange shape.

    Returns the eight evidence keys, plus runs_before/runs_after (as
    format_text), style_id, style_type ("paragraph"/"character"),
    `warnings` when non-empty, and (character style, track_changes=True
    only) revision_ids/track_changes.

    Errors: as format_text, plus:
      STYLE_NOT_FOUND        - style_id is not a style in this document's styles.xml
      UNSUPPORTED_STYLE_TYPE - style_id names neither a paragraph nor a character style
      INVALID_INPUT          - track_changes=True on a paragraph style (see above)
    """
    try:
        return text_edit.execute_apply_style(
            path, find, style_id, expected_matches,
            revision_before=revision_before, force=force, track_changes=track_changes,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


_HEADER_FOOTER_KINDS = frozenset({"header", "footer"})


def execute_read_header_footer(path: str) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        results: list[dict[str, Any]] = []
        for part_info in projection.list_parts_impl(local_path):
            if part_info["kind"] not in _HEADER_FOOTER_KINDS:
                continue
            markdown, warnings, lossy_elements = projection.read_document_markdown(local_path, part_info["part"])
            results.append(
                {
                    "part": part_info["part"],
                    "kind": part_info["kind"],
                    "header_footer_type": part_info["header_footer_type"],
                    "markdown": markdown,
                    "warnings": warnings,
                    "lossy_elements": lossy_elements,
                }
            )
        return {"path": str(resolved), "headers_and_footers": results}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def read_header_footer(path: str) -> dict[str, Any]:
    """Read every header/footer part's content as markdown, in one call --
    list_parts already discovers these parts (kind "header"/"footer",
    header_footer_type "default"/"first"/"even") and read_document(part=
    ...) already reads any one of them; this is the convenience wrapper
    that reads all of them without the caller enumerating parts first.

    Returns path, headers_and_footers (list of {part, kind,
    header_footer_type, markdown, warnings, lossy_elements} -- same shape
    read_document(format="markdown") returns per part). Empty list for a
    document with no header/footer parts.

    Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
    Word's owner file is present.

    Errors:
      INVALID_INPUT  - path does not exist or is outside the allowed roots
      SNAPSHOT_FAILED - the read-path snapshot could not be validated
    """
    try:
        return execute_read_header_footer(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tool: diff_body_vs_file
# ---------------------------------------------------------------------------


def execute_diff_body_vs_file(path: str, file_path: str) -> dict[str, Any]:
    """Diff the docx body's markdown projection against a local markdown
    file -- issue #28 WP-11a, same ``difflib`` mechanism as
    GoogleDocs-MCP's ``diff_tab_vs_file`` (markdown_mutations.py
    ``execute_diff_tab_vs_file``).

    This is a READ tool -- no mutation, no audit, no evidence envelope
    required (mirrors Google's own "no changes to the document or file"
    framing for that tool).
    """
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)

    # Reuses the same allowlist/denylist paths.resolve_allowed_docx_path
    # already enforces for every docx path this server touches (its own
    # docstring: "the same shape the Google server uses for
    # diff_tab_vs_file") -- one allowlist for both sides of the diff,
    # rather than a second, parallel resolver the way Google's
    # markdown_mutations.py keeps a diff-specific
    # _resolve_allowed_diff_file next to its general one.
    file_resolved = paths.resolve_allowed_docx_path(file_path, must_exist=True)
    if not file_resolved.is_file():
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"Path is not a regular file: {file_path!r}",
            {"file_path": file_path, "resolved_path": str(file_resolved)},
        )

    local_path, is_temp = _read_local_copy(resolved)
    try:
        revision = projection.compute_revision(local_path)
        body_markdown, warnings, lossy_elements = projection.read_document_markdown(
            local_path, projection.DEFAULT_PART
        )
        file_content = file_resolved.read_text(encoding="utf-8")

        body_lines = body_markdown.splitlines(keepends=True)
        file_lines = file_content.splitlines(keepends=True)

        matcher = difflib.SequenceMatcher(None, body_lines, file_lines)
        hunks: list[dict[str, Any]] = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            hunks.append(
                {
                    "tag": tag,
                    "body_lines": body_lines[i1:i2],
                    "file_lines": file_lines[j1:j2],
                    "body_range": [i1 + 1, i2],
                    "file_range": [j1 + 1, j2],
                }
            )

        unified = list(
            difflib.unified_diff(
                body_lines,
                file_lines,
                fromfile=f"docx:{path}#body",
                tofile=file_path,
            )
        )

        result: dict[str, Any] = {
            "path": str(resolved),
            "file_path": file_path,
            "revision": revision["token"],
            "identical": body_markdown == file_content,
            "hunks": hunks,
            "unified_diff": "".join(unified),
            "warnings": warnings,
        }
        # Same convention as read_document: present only when the body's
        # own markdown rendering actually lost something (a merged/nested
        # table). A non-empty lossy_elements here is a direct reason an
        # identical=True verdict below cannot be fully trusted -- see this
        # tool's own docstring.
        if lossy_elements:
            result["lossy_elements"] = lossy_elements
        return result
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool()
def diff_body_vs_file(path: str, file_path: str) -> dict[str, Any]:
    """Export the docx body as markdown and diff against a local file.

    Use this tool when you need to compare a .docx's body against a local
    markdown file -- e.g. to decide whether a push/sync is needed. The
    docx side is rendered through the exact same projection
    ``read_document(format="markdown")`` uses (``part`` fixed to the body,
    ``word/document.xml``), so this compares against the same rendering
    every other skill and tool sees, never a second, divergent one. The
    server reads the local file directly (it runs locally); both ``path``
    and ``file_path`` are resolved through the same allowlist/denylist
    (VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS; see paths.py).

    This is a read-only tool -- it makes no changes to the document or
    file. It does not join MUTATING_TOOLS and does not return the eight
    evidence keys.

    Both sides are split with splitlines(keepends=True) before diffing --
    matching GoogleDocs-MCP's diff_tab_vs_file exactly, on purpose, so a
    file on disk that ends with a trailing newline (the common case) will
    surface as a one-line difference against the body projection, which
    does not end with one; a caller comparing against a projection should
    account for that rather than read it as real drift.

    AFFIRMATIVE-ONLY, same as GoogleDocs-MCP's diff_tab_vs_file: a
    returned hunk, or identical=False, is solid evidence a difference
    exists. identical=True is NOT proof no difference exists -- it only
    means this particular markdown rendering of the docx body found none
    against this particular reading of the file, which can under-report a
    real difference for docx-specific reasons a plain text diff has no
    way to see:
      - The docx side is a PROJECTION, not the OOXML itself. Content the
        markdown renderer cannot express as GFM (a merged/spanned table
        cell, a nested table -- see lossy_elements) is dropped from the
        comparison entirely, so two docx files that differ only in ways
        the renderer discards compare identical here even though the
        underlying .docx bytes are not. A non-empty lossy_elements in the
        result is a direct signal to distrust an identical=True verdict.
      - Only the BODY is compared. Headers, footers, footnotes/endnotes,
        and comments are out of scope by design (see diff_body_vs_file's
        own name) -- a real difference confined to one of those parts
        never surfaces here.
      - When Word's owner file is present, this reads a validated
        snapshot rather than the live file (core/document-backend-
        protocol.md §4's "reads never refuse" rule) -- the comparison
        is then only as fresh as the last flush/save the snapshot
        captured, not necessarily the very latest keystrokes.
    A skill deciding whether a sync is needed should treat identical=True
    as "no difference found by this projection," not as a guarantee the
    two are in sync.

    Returns path, file_path, revision (the "<doc8>:<cmt8>" token),
    identical (bool), hunks (tagged body_lines/file_lines/body_range/
    file_range per difflib.SequenceMatcher.get_opcodes()), unified_diff
    (a unified diff string), warnings, and lossy_elements (present only
    when non-empty).

    Errors:
      INVALID_INPUT - a bad path, file_path not found, or file_path is
                       not a regular file
    """
    try:
        return execute_diff_body_vs_file(path, file_path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Doctor
# ---------------------------------------------------------------------------

_WORD_APP_PATH = "/Applications/Microsoft Word.app"
_DOCTOR_FIXTURE = Path(__file__).resolve().parent.parent.parent / "tests" / "fixtures" / "word" / "one-page.docx"


def _detect_host_app() -> str:
    """Walk the process ancestry to the nearest ".app/Contents/MacOS/<name>"
    bundle — the app macOS actually needs the Automation grant against
    (Terminal, or whichever app is hosting this process). Ported from
    scripts/docx-doctor.sh's detect_host_app(): the immediate parent is
    usually a shell, `/usr/bin/login` is typically one hop further up, and
    only then the actual host .app — verified on this machine giving
    "Ghostty" via: claude -> -/bin/zsh -> /usr/bin/login -> Ghostty.app ->
    launchd."""
    pid = os.getppid()
    for _ in range(8):
        if not pid or pid <= 1:
            break
        try:
            comm = subprocess.run(
                ["ps", "-o", "comm=", "-p", str(pid)],
                capture_output=True, text=True, timeout=5,
             check=False,
            ).stdout.strip()
        except (OSError, subprocess.TimeoutExpired):
            break
        if ".app/Contents/MacOS/" in comm:
            return comm.split(".app/Contents/MacOS/")[0].rsplit("/", 1)[-1]
        try:
            ppid_out = subprocess.run(
                ["ps", "-o", "ppid=", "-p", str(pid)],
                capture_output=True, text=True, timeout=5,
             check=False,
            ).stdout.strip()
            pid = int(ppid_out) if ppid_out else 0
        except (OSError, subprocess.TimeoutExpired, ValueError):
            break
    return "(unknown host app — could not walk the process ancestry to an .app bundle)"


def _doctor_print(status: str, message: str) -> None:
    print(f"{status:<5} {message}")


def doctor() -> int:
    """`verified-docx-mcp doctor` — one-machine, one-host-app diagnostic for
    the Word PDF render path. Run from the SAME host application that runs
    the MCP server (macOS grants Automation, and file access, per hosting
    app — a grant to Terminal.app does not authorize iTerm, Ghostty, VS
    Code, or Claude Desktop).

    Checks, in order (adapted from scripts/docx-doctor.sh in JennyStack,
    issue #30 WP-04 — "the #30 doctor checks"; pandoc/uv presence are
    intentionally NOT carried over, since this server has no runtime
    dependency on either — see this WP's PR notes for the correction):
      0. NOTE (read-only, never fails): the effective allowed file roots and
         whether they came from VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS or the
         default (issue #101)
      1. Microsoft Word present
      2. Word's sandbox container directory exists
      3. pdfinfo present (ADVISORY — page counts fall back to a regex)
      4. AppleScript round trip: `osascript -e '... version of application'`
      5. Leftover "jennystack-render-*" document windows (ADVISORY,
         read-only — never closes anything; see render.py's module
         docstring for why every render leaves its document open)
      6. A render of the bundled 1-page fixture (tests/fixtures/word/
         one-page.docx), asserting page_count == 1

    Returns 0 (PASS) or 1 (FAIL); prints the detected host app on an
    automation failure.
    """
    fail = False
    host_app = _detect_host_app()

    def automation_fix() -> None:
        print(
            f'      macOS has not granted "{host_app}" permission to control Microsoft Word:\n'
            f'      System Settings -> Privacy & Security -> Automation -> "{host_app}" -> Microsoft Word,\n'
            "      then rerun."
        )

    print("== verified-docx-mcp doctor: Word PDF render path ==\n")

    # 0. Effective allowed file roots (read-only; issue #101)
    allowed_roots_source = "env" if os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV) is not None else "default"
    effective_roots = ", ".join(str(root) for root in paths._allowed_file_roots())
    _doctor_print(
        "NOTE",
        f"allowed file roots ({allowed_roots_source}): {effective_roots}",
    )

    # 1. Microsoft Word present
    word_present = Path(_WORD_APP_PATH).is_dir()
    if word_present:
        _doctor_print("PASS", f'Microsoft Word present ("{_WORD_APP_PATH}")')
    else:
        _doctor_print("FAIL", f'Microsoft Word not found at "{_WORD_APP_PATH}". Fix: install Microsoft Word.')
        fail = True

    # 2. Word's sandbox container directory
    sandbox_root = render_module._word_sandbox_root()
    if sandbox_root.is_dir():
        _doctor_print("PASS", f"Word's sandbox container directory exists ({sandbox_root})")
    else:
        _doctor_print(
            "FAIL",
            f"Word's sandbox container directory does not exist ({sandbox_root}). "
            "Fix: launch Microsoft Word once (so macOS creates its container), then rerun.",
        )
        fail = True

    # 3. pdfinfo (advisory)
    if render_module._find_pdfinfo():
        _doctor_print("PASS", f"pdfinfo present ({render_module._find_pdfinfo()})")
    else:
        _doctor_print("NOTE", "pdfinfo not found (ADVISORY: page counts fall back to a regex, less reliable). Fix: brew install poppler")

    # 4. AppleScript round trip
    if word_present:
        try:
            proc = subprocess.run(
                [render_module._osascript_bin(), "-e", 'version of application "Microsoft Word"'],
                capture_output=True, text=True, timeout=20,
             check=False,
            )
        except subprocess.TimeoutExpired:
            _doctor_print("FAIL", "AppleScript round trip timed out after 20s (a stuck native dialog is the most common cause).")
            automation_fix()
            fail = True
            proc = None
        if proc is not None:
            if proc.returncode == 0 and proc.stdout.strip():
                _doctor_print("PASS", f'AppleScript round trip: version of application "Microsoft Word" -> {proc.stdout.strip()}')
            elif render_module._classify_automation_error(proc.stderr) or render_module._classify_automation_error(proc.stdout):
                _doctor_print("FAIL", "AppleScript round trip failed: automation not granted.")
                automation_fix()
                fail = True
            else:
                _doctor_print("FAIL", f"AppleScript round trip failed: {proc.stderr.strip() or proc.stdout.strip()}")
                fail = True
    else:
        _doctor_print("NOTE", "skipping AppleScript round trip — Word is not installed.")

    # 5. Leftover jennystack-render-* windows (advisory, read-only)
    if word_present:
        try:
            proc = subprocess.run(
                [render_module._osascript_bin(), "-e", 'tell application "Microsoft Word" to get name of every window'],
                capture_output=True, text=True, timeout=10,
             check=False,
            )
            if proc.returncode == 0 and proc.stdout.strip():
                names = [n.strip() for n in proc.stdout.split(",")]
                leftover = [n for n in names if "jennystack-render-" in n]
                if leftover:
                    _doctor_print("NOTE", f"{len(leftover)} leftover \"jennystack-render-*\" document window(s) open in Word (read-only count; never auto-closed — see render.py's module docstring). Safe to close by hand.")
                else:
                    _doctor_print("NOTE", 'no leftover "jennystack-render-*" document windows open in Word.')
        except (OSError, subprocess.TimeoutExpired):
            pass

    # 6. 1-page fixture render
    if fail:
        _doctor_print("NOTE", "skipping the fixture render — an earlier check already failed.")
    elif not _DOCTOR_FIXTURE.is_file():
        _doctor_print("FAIL", f"bundled fixture not found: {_DOCTOR_FIXTURE}")
        fail = True
    else:
        import shutil
        import tempfile

        # NOT the system default tempfile location (/tmp or
        # /var/folders/...): that falls outside
        # VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS' default (the user's home
        # directory), so execute_export_pdf's own path allowlist check
        # would reject the smoke render's own output path. Stay under
        # the home directory, which is what a caller running this doctor
        # with default settings actually has allowed.
        doctor_cache_root = Path.home() / ".cache" / "verified-docx-mcp" / "doctor"
        doctor_cache_root.mkdir(parents=True, exist_ok=True)
        tmp = tempfile.mkdtemp(prefix="smoke-", dir=str(doctor_cache_root))
        try:
            out_pdf = Path(tmp) / "doctor-smoke.pdf"
            try:
                result = execute_export_pdf(str(_DOCTOR_FIXTURE), str(out_pdf))
            except VerifyError as exc:
                code = exc.envelope.error_code
                if code == ErrorCode.AUTOMATION_NOT_GRANTED:
                    _doctor_print("FAIL", "fixture render: automation not granted.")
                    automation_fix()
                elif code == ErrorCode.WORD_SANDBOX_UNAVAILABLE:
                    _doctor_print("FAIL", "fixture render: Word's sandbox container directory does not exist. Fix: launch Microsoft Word once, then rerun.")
                else:
                    _doctor_print("FAIL", f"fixture render failed: {exc.envelope.to_dict()}")
                fail = True
            else:
                if result.get("page_count") == 1:
                    _doctor_print("PASS", f"fixture render: 1-page document rendered, page_count == 1 (pdf: {result['pdf_path']})")
                else:
                    _doctor_print("FAIL", f"fixture render: expected page_count == 1, got: {result}")
                    fail = True
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    print()
    if fail:
        print("verified-docx-mcp doctor: FAIL — see above for the fix(es).")
        return 1
    print("verified-docx-mcp doctor: PASS — the Word render path is ready.")
    return 0


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Entry point for the `verified-docx-mcp` command.

    `verified-docx-mcp doctor`  — run the local render-path diagnostic.
    `verified-docx-mcp`         — start the stdio MCP server.
    """
    if len(sys.argv) > 1 and sys.argv[1] == "doctor":
        sys.exit(doctor())

    mcp.run()


if __name__ == "__main__":
    main()
