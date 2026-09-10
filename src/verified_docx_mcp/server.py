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
of the seven is a mutating tool (export_pdf writes a PDF, never the source
.docx; every other tool here only reads), so MUTATING_TOOLS (middleware.py)
stays empty until WP-04's markdown writers land.

Every WP-03 read tool honors core/document-backend-protocol.md §4's "reads
never refuse" rule via ``_read_local_copy`` below: when Word's owner file is
present, the tool reads a validated snapshot (``paths.snapshot_docx_package``)
instead of the live file, rather than returning DOCX_LOCKED (which gates
writes only).
"""

from __future__ import annotations

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

from . import paths, projection
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
    (defaults to the user's home directory), must not resolve to a
    credential path, and its parent directory must already exist. This is
    a read/render tool: the source .docx is never modified (Word opens a
    private staged copy — see render.py), so the return value has no
    "applied" key.

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
                markdown, _ = projection.markdown_from_projection(local_path, scoped)
                result["markdown"] = markdown
            result["warnings"] = scoped.warnings
            return result

        if format == "text":
            result["text"] = projection.read_document_text(local_path, part)
            result["warnings"] = projection.project_part(local_path, part).warnings
        elif format == "runs":
            result["runs"] = projection.read_document_runs(local_path, part)
            result["warnings"] = projection.project_part(local_path, part).warnings
        else:  # markdown
            markdown, warnings = projection.read_document_markdown(local_path, part)
            result["markdown"] = markdown
            result["warnings"] = warnings
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
    format="markdown" (default) renders headings/bold/italic and stable
    placeholder tokens ([TABLE], [GRAPHIC], [field:instr]) for constructs
    markdown cannot represent — WP-04's inverse (markdown -> OOXML) is not
    implemented here.

    A deleted span (w:del/w:delText) and a field's own instruction text
    (w:instrText) are excluded from every format; a field's RESULT text
    (between fldChar "separate" and "end", or all of a w:fldSimple's
    nested runs) IS included — in format="markdown" it flows in as
    ordinary rendered text, never duplicated by a "[field:...]" token
    (that placeholder appears only when a field carries NO result at all).

    Returns path, part, format, section_key (echoed back, None unless
    passed), revision (the "<doc8>:<cmt8>" token), revision_detail (the
    full {document_sha256, comments_sha256, size, mtime_ns} tuple),
    warnings, plus text|runs|markdown per format.

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
