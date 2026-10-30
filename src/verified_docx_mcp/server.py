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

Tools registered here (WP-02): ``export_pdf``, ``lock_status``. WP-2 of
issue #106 (https://github.com/michaelrobertsutton/JennyStack/issues/106)
adds ``live_status`` (read-only; starts the local live-mode bridge lazily
and reports connected Word task-pane sessions -- see ``live/bridge.py``).
WP-03 adds
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
import hashlib
import itertools
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Literal, NoReturn

from fastmcp import FastMCP

from . import (
    audit,
    comment_batch,
    comments,
    geometry,
    images,
    locate,
    mutations,
    paragraph_insert,
    paths,
    projection,
    table_diff,
    table_structure,
    tables,
    text_edit,
    tracked_changes,
    write_ledger,
)
from . import render as render_module
from .errors import ErrorCode, VerifyError, _make_error
from .lean_envelope import LeanEnvelopeMiddleware
from .live import bridge as live_bridge
from .live import comments_live, reads_live
from .markdown_input import resolve_markdown_input, resolve_table_rows
from .middleware import EvidenceEnforcementMiddleware

_READ_ONLY: dict[str, bool] = {"readOnlyHint": True, "idempotentHint": True}
_DESTRUCTIVE: dict[str, bool] = {"destructiveHint": True}

mcp = FastMCP(
    "verified-docx-mcp",
    instructions=(
        "MCP server for local .docx files with verified writes. Reading and "
        "writing need no Word installation (pure-Python OOXML manipulation); "
        "Word is required only for export_pdf, which additionally needs a "
        "macOS Automation grant for the calling app. export_pdf writes a "
        "local PDF file only and never modifies the source .docx. lock_status "
        "reports Word/LibreOffice owner-file and sync-quiesce state as data "
        "only — it never refuses a call by itself. live_status starts the "
        "local live-mode bridge lazily and reports connected Word task-pane "
        "sessions; empty sessions is normal, not an error. Workflow: prefer "
        'format="markdown" for reads (format="runs" is 5-10x larger); use '
        "replace_text_batch for two or more edits; chain each result's "
        "revision_after into the next call's revision_before; "
        'source="auto" / write_mode="auto" route to a connected pane without '
        "a live_status call; live_revision is the cheap way to get a fresh "
        "live revision (and re-arms writes after LIVE_STALE)."
    ),
)

mcp.add_middleware(EvidenceEnforcementMiddleware())
mcp.add_middleware(LeanEnvelopeMiddleware())


def _raise_tool_error(exc: VerifyError, verbose: bool = False) -> NoReturn:
    """Raise a FastMCP ToolError carrying a JSON error envelope.

    Mirrors verified-googledocs-mcp's server.py helper of the same name. A ZERO_MATCH's
    ``ladder_report`` (a count of zero per normalization rung) is dropped unless
    ``verbose``: ``near_miss`` is the useful part (issue #112).
    """
    from fastmcp.exceptions import ToolError

    payload = exc.envelope.to_dict()
    if not verbose and payload.get("error_code") == ErrorCode.ZERO_MATCH.value:
        payload["diagnostics"] = {k: v for k, v in payload["diagnostics"].items() if k != "ladder_report"}
    raise ToolError(json.dumps(payload, ensure_ascii=False)) from exc


# ---------------------------------------------------------------------------
# Tool: export_pdf
# ---------------------------------------------------------------------------


def _resolve_export_output_path(output_path: str) -> Path:
    """Validate *output_path* for export_pdf (see paths.resolve_output_path;
    mirrors verified-googledocs-mcp's export_pdf validation)."""
    return paths.resolve_output_path(output_path)


def _section_keys_error(missing: list[str]) -> Any:
    return _make_error(
        ErrorCode.SECTION_NOT_FOUND,
        f"section_keys named section(s) not found: {missing!r} "
        "(call find_sections to enumerate the current ones).",
        {"section_keys": missing},
    )


def _plan_section_probes(
    source: Path, section_keys: list[str] | None
) -> tuple[list[dict[str, Any]], list[str], list[int]]:
    """(headings, all_para_refs, probe_ordinals) for export_pdf's section
    geometry, computed BEFORE any render happens.

    headings is projection.find_sections_impl(source)'s "heading"-kind
    entries, in document order, each augmented with a
    "next_start_para_ref" key (start_para_ref of the FULL document's next
    heading, or None when this heading is the document's last one),
    filtered to *section_keys* when given. That augmentation happens
    BEFORE filtering, deliberately: the skills' primary call is
    export_pdf(section_keys=[<one section>]) for a page-budget check, and
    that section's true end is the FULL document's next heading's start
    when one exists — even though that next heading is not itself part
    of the filtered/requested set and never appears in the returned
    "sections" list. Only a heading with no next one AT ALL (the
    document's last heading) falls back to the +0.03 last-paragraph
    approximation in geometry.assemble_sections, regardless of whether
    section_keys narrowed the request down to just that one heading. An
    unknown key in *section_keys* raises SECTION_NOT_FOUND here — before
    render_word() ever runs — since the caller asked for that section by
    name (mirrors mutations.locate_section_range's SECTION_NOT_FOUND).
    all_para_refs is projection.project_part(source)'s main-story
    paragraph order (only computed when there is at least one heading to
    map, since it walks the whole document). probe_ordinals is
    geometry.build_probe_ordinals(headings, all_para_refs) — [] when
    headings is empty, so export_pdf passes probe_paragraphs=None and
    render_word() does no extra work for a document (or a filtered
    section_keys view of one) with nothing to probe.
    """
    all_sections = projection.find_sections_impl(source)
    all_headings = [s for s in all_sections if s["kind"] == "heading"]

    headings = [
        {
            **heading,
            "next_start_para_ref": (
                all_headings[idx + 1]["start_para_ref"] if idx + 1 < len(all_headings) else None
            ),
        }
        for idx, heading in enumerate(all_headings)
    ]

    if section_keys is not None:
        known = {s["section_key"] for s in headings}
        missing = [k for k in section_keys if k not in known]
        if missing:
            raise _section_keys_error(missing)
        wanted = set(section_keys)
        headings = [s for s in headings if s["section_key"] in wanted]

    if not headings:
        return [], [], []

    proj = projection.project_part(source)
    all_para_refs = [m.para_ref for m in proj.paragraphs]
    probe_ordinals = geometry.build_probe_ordinals(headings, all_para_refs)
    return headings, all_para_refs, probe_ordinals


def execute_export_pdf(
    path: str,
    output_path: str,
    *,
    timeout: int = render_module.DEFAULT_TIMEOUT,
    close_after: bool = True,
    section_keys: list[str] | None = None,
) -> dict[str, Any]:
    """Render *path* (.docx) to *output_path* (.pdf) via Word automation.

    Returns {pdf_path, sha256, page_count, page_count_source, engine,
    left_open_document, closed_after, close_error, page_height_pt,
    sections, sections_error}. page_count is Word's own count when
    available (authoritative — see render.py's module docstring),
    cross-checked against pdfinfo/regex; None (never a guessed 0) when
    neither source can determine it. Raises VerifyError(RENDER_ENGINE_UNAVAILABLE),
    (AUTOMATION_NOT_GRANTED), (WORD_SANDBOX_UNAVAILABLE), (RENDER_FAILED),
    (SECTION_NOT_FOUND) (an unknown section_keys entry, checked before any
    render happens), or (SECTION_GEOMETRY_UNAVAILABLE) (section_keys was
    given and a requested section's page-span could not be verified) — see
    errors.py and render.py's module docstring for exactly when each
    fires.

    Not gated by DOCX_LOCKED (core/document-backend-protocol.md §4:
    "Reads and export_pdf are not gated by this code") — render_word()
    itself already stages a private copy of the source into Word's sandbox
    container before opening it, so a concurrently-open Word session on
    the original file is not disturbed.
    """
    source = paths.resolve_allowed_docx_path(path, must_exist=True)
    target = _resolve_export_output_path(output_path)

    headings, all_para_refs, probe_ordinals = _plan_section_probes(source, section_keys)

    try:
        result = render_module.render_word(
            str(source),
            str(target),
            timeout=timeout,
            close_after=close_after,
            probe_paragraphs=probe_ordinals or None,
        )
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

    sections_result, sections_error = geometry.assemble_sections(
        headings, all_para_refs, result.get("paragraph_geometry") or {}, result.get("page_height_pt")
    )
    if sections_result is None and section_keys is not None:
        # Explicit mode: the caller asked for this section's geometry by
        # name, so a verification failure is not something export_pdf can
        # silently paper over — raise rather than degrade. The PDF is
        # already written to *target* by this point; only the tool call
        # itself fails.
        raise _make_error(
            ErrorCode.SECTION_GEOMETRY_UNAVAILABLE,
            sections_error or "section geometry unavailable",
            {"detail": sections_error, "section_keys": section_keys},
        )

    return {
        "pdf_path": pdf_path,
        "sha256": _sha256_file(pdf_path),
        "page_count": page_count_value,
        "page_count_source": page_count_source,
        "engine": "word",
        # left_open_document is None when the staged copy's window was
        # closed; otherwise it is that window's title. closed_after and
        # close_error report whether the close itself (attempted only when
        # close_after=True) succeeded — see render.py's module docstring.
        "left_open_document": result.get("left_open_document"),
        "closed_after": result.get("closed_after", False),
        "close_error": result.get("close_error"),
        # page_height_pt / sections / sections_error (issue #102): see
        # geometry.py and this function's docstring. page_height_pt is
        # None when no section had to be probed (no headings, or an empty
        # section_keys filter). sections is [] (never null) when there
        # were no headings to probe in the first place; it is null only
        # on a default-mode (section_keys=None) verification failure,
        # paired with sections_error explaining why — page_count and
        # every field above are unaffected either way.
        "page_height_pt": result.get("page_height_pt"),
        "sections": sections_result,
        "sections_error": sections_error,
    }


def _sha256_file(path: str) -> str:
    import hashlib

    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@mcp.tool(annotations=_READ_ONLY)
def export_pdf(
    path: str,
    output_path: str,
    close_after: bool = True,
    section_keys: list[str] | None = None,
) -> dict[str, Any]:
    """Render a local .docx to PDF via Microsoft Word and report its page count and section page spans.

    output_path must be inside VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS, not a credential path, parent
    directory existing. The source .docx is never modified (Word renders a private staged copy), so
    there is no "applied" key. close_after (default True) closes that Word window after the PDF is
    written; a close failure goes in close_error, not raised.

    section_keys (default None = every heading section): limit "sections" to these find_sections
    heading keys. Entries have section_key, heading_text, start_page/start_fraction,
    end_page/end_fraction (fraction of page_height_pt), pages, start_paragraph/end_paragraph,
    text_verified. A section ends at the next heading in the whole document; the last section's end
    is approximate. Each probed heading must verify: in default mode a failure gives sections null
    plus sections_error; with section_keys given it raises SECTION_GEOMETRY_UNAVAILABLE.

    Returns pdf_path, sha256, page_count (None when unknown, never a guessed 0), page_count_source,
    engine, left_open_document, closed_after, close_error, page_height_pt, sections, sections_error.
    Errors: INVALID_INPUT, RENDER_ENGINE_UNAVAILABLE, AUTOMATION_NOT_GRANTED (macOS declined
    Automation of Word; run `verified-docx-mcp doctor`), WORD_SANDBOX_UNAVAILABLE (Word never
    launched), RENDER_FAILED, SECTION_NOT_FOUND (checked before rendering),
    SECTION_GEOMETRY_UNAVAILABLE.
    Full contract: docs/tool-contracts.md#export_pdf.
    """
    try:
        return execute_export_pdf(
            path, output_path, close_after=close_after, section_keys=section_keys
        )
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


# recent / ledger are what the write guard refuses on (EXTERNAL_EDITOR_ACTIVE), so they stay visible.
_EXTERNAL_ACTIVITY_COMPACT_KEYS = ("divergent", "still_current", "recent", "ledger")


def execute_lock_status(
    path: str,
    *,
    quiesce_interval: float = 1.5,
    sleep=time.sleep,
    include_external_activity: bool = True,
) -> dict[str, Any]:
    """Report owner-file and sync-quiesce state for *path*, as data only.

    Never refuses (core/document-backend-protocol.md §4: DOCX_LOCKED gates
    writes only, and that gate lives in the write guard landing in
    WP-04/WP-10 — lock_status itself only reports what it observed).

    *include_external_activity* (issue #27): the write guard passes False
    because it computes ``write_ledger.external_activity`` itself, with
    the fingerprint it also hands to the write window.
    """
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    directory = resolved.parent
    target_name = resolved.name

    from . import desktop_word

    desktop_status = desktop_word.status(resolved)
    owner_file = _find_owner_file(directory, target_name)

    sample_1 = _stat_sample(resolved)
    sleep(quiesce_interval)
    sample_2 = _stat_sample(resolved)
    tmp_siblings = _find_tmp_siblings(directory, target_name)

    sync_quiesced = sample_1 == sample_2 and not tmp_siblings

    result: dict[str, Any] = {
        "path": str(resolved),
        "owner_file": owner_file,
        "desktop_word": desktop_status,
        "sync_quiesced": sync_quiesced,
        "sync_detail": {
            "sample_1": sample_1,
            "sample_2": sample_2,
            "interval_seconds": quiesce_interval,
            "tmp_siblings": tmp_siblings,
        },
    }
    if include_external_activity:
        # Issue #27. Observing a divergence persists its first-seen time
        # (best-effort), so this read tool has a small state side effect.
        result["external_activity"] = write_ledger.external_activity(resolved)
    return result


@mcp.tool(annotations=_READ_ONLY)
def lock_status(path: str, verbose: bool = False) -> dict[str, Any]:
    """Report Word/LibreOffice owner-file presence and sync-quiesce state. Data only; never refuses.

    Returns owner_file {present, path, format, owner_name}, desktop_word, sync_quiesced (false if the
    file looks still being written, e.g. OneDrive/iCloud staging files) and external_activity reduced
    to {divergent, still_current, recent, ledger}. verbose=True adds path, sync_detail and the full
    external_activity: how the file on disk compares with what THIS server last wrote. Use it after a
    write has had time to sync: still_current false means another writer changed the package since,
    so this server's edits may have been overwritten. It shows that bytes changed, not that an editor
    is still open; no divergence is not proof of safety, and a file this server never wrote reports
    ledger "none". Reading it persists the divergence's first-seen time.

    Errors: INVALID_INPUT (path missing or outside the allowed roots).
    Full contract: docs/tool-contracts.md#lock_status.
    """
    try:
        result = execute_lock_status(path)
    except VerifyError as exc:
        _raise_tool_error(exc)
    if verbose:
        return result
    result.pop("sync_detail", None)
    result.pop("path", None)
    activity = result.get("external_activity")
    if isinstance(activity, dict):
        result["external_activity"] = {k: activity.get(k) for k in _EXTERNAL_ACTIVITY_COMPACT_KEYS}
    return result


# ---------------------------------------------------------------------------
# Tool: live_status (issue #106 WP-2:
# https://github.com/michaelrobertsutton/JennyStack/issues/106)
# ---------------------------------------------------------------------------


def _revision_status(session: Any) -> dict[str, Any]:
    from .live.revisions_live import list_revisions

    state = list_revisions(session)
    return {"revision_count": len(state["revisions"]) if state["revisions"] is not None else None,
            "revision_coverage": state["coverage"], "revision_reason": state.get("reason")}


_LIVE_STATUS_VERBOSE_SESSION_FIELDS = ("heartbeat_count", "hash_sample_count", "requirement_sets")


def execute_live_status(path: str | None = None, verbose: bool = False) -> dict[str, Any]:
    """Start the live bridge lazily (idempotent) and report its state.

    *path* scopes ``sessions`` to that document's basename; default (non-verbose)
    output omits the per-session counters / ``requirement_sets`` and the top-level
    ``recent_boots`` and ``shared`` blocks (issue #111); ``locks`` (and this client's ``client_id``)
    are always included in shared mode.

    Never raises for "no pane connected yet" -- an empty ``sessions``
    list is the normal, expected answer before the lead opens the task
    pane in Word. It DOES raise (``LIVE_UNAVAILABLE``) if the bridge
    itself could not start -- e.g. no ``openssl`` on PATH to generate a
    first-run cert, or its HTTPS/WSS ports are already bound by
    something else. That is a real failure distinct from "nothing has
    connected to a working bridge yet".
    """
    try:
        registry = live_bridge.start_in_background()
    except Exception as exc:
        raise _make_error(
            ErrorCode.LIVE_UNAVAILABLE,
            f"live bridge failed to start: {exc}",
            {"exception_type": type(exc).__name__},
        ) from exc

    ports = live_bridge.current_ports()
    port, ops_port = ports if ports is not None else (None, None)

    from .live import write_mode as _write_mode

    wanted_name: str | None = None
    if path:
        try:
            wanted_name = _write_mode._document_name(path)
        except VerifyError:
            # A path outside the allowed roots can still name a pane's document.
            wanted_name = os.path.basename(path.replace("\\", "/").rstrip("/"))
    sessions: list[dict[str, Any]] = []
    for session in registry.list():
        if wanted_name is not None and session.document_name != wanted_name:
            continue
        entry = (
            {
                "document_name": session.document_name,
                "document_url": session.document_url,
                "connected_since": session.connected_since,
                "last_heartbeat_age_s": session.heartbeat_age(),
                "body_sha256": session.last_body_sha256,
                "body_hash_state": session.body_hash_state,
                "body_hash_age_s": (time.monotonic() - session.body_hash_sampled_monotonic
                                    if session.body_hash_sampled_monotonic is not None else None),
                "heartbeat_count": session.heartbeat_count,
                "hash_sample_count": session.hash_sample_count,
                "requirement_sets": session.hello.requirement_sets,
                "instance_id": session.hello.instance_id,
                "platform": session.hello.platform,
                **_revision_status(session.status_view() if hasattr(session, "status_view") else session),
            }
        )
        if not verbose:
            for key in _LIVE_STATUS_VERBOSE_SESSION_FIELDS:
                entry.pop(key, None)
            for key in ("revision_reason", "revision_count"):
                if entry.get(key) is None:
                    entry.pop(key, None)
            for key in ("last_heartbeat_age_s", "body_hash_age_s"):
                if isinstance(entry.get(key), float):
                    entry[key] = round(entry[key], 1)
        sessions.append(entry)

    shared_status = registry.status() if hasattr(registry, "status") else None
    pid = os.getpid()
    owner_pid = shared_status.get("owner_pid", pid) if shared_status is not None else pid
    # Issue #84: pane boots are recorded where the HTTP server runs (the owner), so a
    # shared client reads them from the broker status instead of its own process.
    recent_boots = (
        list(shared_status.get("recent_boots", []))
        if shared_status is not None
        else live_bridge.recent_boots()
    )
    result: dict[str, Any] = {
        "bridge_running": True,
        "pid": pid,
        "owner_pid": owner_pid,
        "is_owner": owner_pid == pid,
        "mode": "shared" if shared_status is not None else "local",
        "port": port,
        "ops_port": ops_port,
        "sessions": sessions,
        # Issue #22 B3 follow-up: SessionRegistry keys purely on basename
        # (deliberate -- see that class's own docstring), so two DIFFERENT
        # documents sharing a file name are indistinguishable by name
        # alone. Never refused (the common, harmless case is the SAME
        # document reconnecting under a new LiveSession object), but every
        # such basename collision is recorded here so it is visible
        # rather than a silent misroute -- empty when none have happened.
        "session_collisions": registry.collisions(),
    }
    # Who holds a document lease / section locks is the only way to see the holder behind
    # LOCKED_BY_OTHER_CLIENT, so `locks` (and this client's id) are always surfaced; the rest of
    # `shared` stays behind verbose.
    if shared_status is not None:
        locks = list(shared_status.get("locks") or [])
        if wanted_name is not None:
            urls = {s.document_url for s in registry.list() if s.document_name == wanted_name}
            locks = [lk for lk in locks if lk.get("document_url") in urls]
        for lk in locks:
            if isinstance(lk.get("remaining_s"), float):
                lk["remaining_s"] = round(lk["remaining_s"], 1)
        result["locks"] = locks
        if shared_status.get("client_id"):
            result["client_id"] = shared_status["client_id"]
    if verbose:
        result["shared"] = shared_status
        result["recent_boots"] = recent_boots
    boot_warnings = _pane_reload_loop_warnings(recent_boots)
    if boot_warnings:
        result.setdefault("warnings", []).extend(boot_warnings)
    return result


_RELOAD_LOOP_BOOTS = 3
_RELOAD_LOOP_WINDOW_S = 120


def _pane_reload_loop_warnings(recent_boots: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Issue #84: warn when one document's pane booted >= 3 times within 120 s."""
    warnings: list[dict[str, Any]] = []
    for group in recent_boots:
        boots = [b for b in group.get("boots", []) if b.get("age_s", 1e9) <= _RELOAD_LOOP_WINDOW_S]
        if len(boots) < _RELOAD_LOOP_BOOTS:
            continue
        ages = sorted((b["age_s"] for b in boots), reverse=True)
        gaps = [a - b for a, b in itertools.pairwise(ages)]
        avg = round(sum(gaps) / len(gaps), 1)
        kinds = sorted({k for b in boots for k in b.get("previous_event_kinds", [])})
        name = group.get("document_name", "(unknown)")
        warnings.append(
            {
                "code": "pane_reload_loop",
                "document_name": name,
                "boots": len(boots),
                "avg_interval_s": avg,
                "previous_event_kinds": kinds,
                "message": (
                    f"The task pane for {name} booted {len(boots)} times in the last "
                    f"{_RELOAD_LOOP_WINDOW_S} s (average interval {avg} s; previous_event_kinds "
                    f"seen: {', '.join(kinds) or 'none'}). Word is reloading the pane. Check "
                    "previous_events in ~/.cache/verified-docx-mcp/live-report/latest.json, try a "
                    "local copy of the document, and report on issue #84."
                ),
            }
        )
    return warnings


@mcp.tool()
def live_document_lock(path: str, release: bool = False, lease_s: float = 60) -> dict[str, Any]:
    """Acquire/renew or release this client's shared live DOCUMENT lock (path, release, lease_s).

    Requires shared bridge mode (the POSIX default; VERIFIED_DOCX_SHARED_BRIDGE=0 disables it).
    Lease is 1-300 seconds; reads stay open to every client. Whole-document locks only. While the
    lease is held, other clients' writes fail LOCKED_BY_OTHER_CLIENT.
    """
    from .live import comments_live, write_mode
    from .live.session import LiveError

    try:
        session = write_mode.live_session_for(path)
        if not hasattr(session, "client"):
            raise _make_error(ErrorCode.LIVE_UNAVAILABLE, "Enable shared bridge and call live_status first")
        return session.client.rpc("unlock" if release else "lock", **session.target(), lease_s=lease_s)
    except LiveError as exc:
        try:
            comments_live._raise_from_live_error(exc)
        except VerifyError as error:
            _raise_tool_error(error)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def live_list_sections(path: str, compact: bool = False, content_digests: bool = False) -> dict[str, Any]:
    """List the live document's heading sections as the Word pane sees them. Needs a pane that reports
    section_locks; works with or without the shared bridge.

    Each entry has section_key (as find_sections), heading_text, level, paragraph_count, slug_unique
    and sectionSha256. Only slug_unique sections can be locked with live_section_lock.
    compact=true returns {sections: [[section_key, heading_text, sha12], ...], body_sha12} (sha12 =
    first 12 hex of sectionSha256, equal to find_sections(digests=true) for unchanged text).
    content_digests=true adds content_sections [[section_key, content_sha12], ...] and
    preamble_content_sha12 (see find_sections), computed from one body_ooxml export taken after the
    section list (LIVE_OP_FAILED if the body changed between the two). If the server's section keys
    differ from the pane's, warnings carries content_sections_key_mismatch. This never re-arms a
    shared-bridge write baseline.
    Full contract: docs/tool-contracts.md#live_list_sections.
    """
    from .live import comments_live, reads_live, write_mode
    from .live.session import LiveError

    try:
        session = write_mode.live_session_for(path)
        write_mode.require_capability(session, "section_locks", feature_description="live_list_sections")
        reply = comments_live._request(session, "sections_list", {})
        content: reads_live.BodyDigests | None = None
        if content_digests:
            write_mode.require_capability(
                session, reads_live.BODY_OOXML_CAPABILITY, feature_description="live_list_sections(content_digests=true)"
            )
            content = reads_live.body_digests_from_ooxml(session, write_mode._document_name(path), content=True)
            if reply.get("bodySha256") != content.body_sha256:
                # The body moved between the list and the export: re-list once (cheap).
                reply = comments_live._request(session, "sections_list", {})
                if reply.get("bodySha256") != content.body_sha256:
                    raise _make_error(
                        ErrorCode.LIVE_OP_FAILED,
                        "the live document changed while its sections were being digested; retry.",
                        {"stage": "content_digests_snapshot"},
                    )
        result: dict[str, Any]
        if compact:
            result = {
                "sections": [
                    [sec["section_key"], sec["heading_text"], sec["sectionSha256"][:12]]
                    for sec in reply.get("sections", [])
                ],
                "body_sha12": str(reply.get("bodySha256", ""))[:12],
            }
        else:
            result = reply
        if content is not None:
            result = {
                **result,
                "content_sections": content.content_sections,
                "preamble_content_sha12": content.preamble_content_sha12,
            }
            warnings = list(content.warnings)
            pane_keys = [sec.get("section_key") for sec in reply.get("sections", [])]
            server_keys = [key for key, _sha in content.content_sections]
            if pane_keys != server_keys:
                warnings.append("content_sections_key_mismatch")
                result["content_sections_key_mismatch"] = {
                    "pane_only": [k for k in pane_keys if k not in server_keys],
                    "server_only": [k for k in server_keys if k not in pane_keys],
                    "order_differs": sorted(map(str, pane_keys)) == sorted(server_keys),
                }
            if warnings:
                result["warnings"] = warnings
        return result
    except LiveError as exc:
        try:
            comments_live._raise_from_live_error(exc)
        except VerifyError as error:
            _raise_tool_error(error)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def live_revision(path: str, content_digests: bool = False) -> dict[str, Any]:
    """The live document's current revision and per-section digests, without reading the body.

    Returns {revision: "live:sha256:<64 hex>", body_sha12, sections: [[section_key, sha12], ...],
    observed_at, via}. revision is the token read_document(source="live") reports and live writes
    accept as revision_before. Calling this re-arms your write baseline like a live read_document
    (shared-bridge mode): use it to resume after LIVE_STALE, but re-read the text if sections shows
    a section you care about changed. Pane capability body_hash_v1 gives one cheap round trip; an
    older pane transfers the whole body (via "body_ooxml_fallback"): reload the pane.
    content_digests=true also returns content_sections and preamble_content_sha12 (see
    find_sections) from one body_ooxml export (via "body_ooxml"); costs a body transfer.

    Errors: LIVE_UNAVAILABLE (no pane), LIVE_SESSION_MISMATCH, LIVE_OP_FAILED,
    LIVE_CAPABILITY_MISSING (no body_hash_v1 or body_ooxml, or no body_ooxml with
    content_digests=true), INVALID_INPUT (unreadable comment threading, content_digests=true only).
    Full contract: docs/tool-contracts.md#live_revision.
    """
    from .live import comments_live, reads_live
    from .live.session import LiveError

    try:
        return reads_live.live_revision(path, content_digests=content_digests)
    except LiveError as exc:
        try:
            comments_live._raise_from_live_error(exc)
        except VerifyError as error:
            _raise_tool_error(error)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def live_section_lock(
    path: str,
    section_keys: list[str],
    release: bool = False,
    lease_s: float = 60,
    include_subsections: bool = False,
) -> dict[str, Any]:
    """Acquire/renew or release this client's lock on heading SECTIONS of a shared live document.
    Requires shared bridge mode (the POSIX default).

    section_keys come from live_list_sections; include_subsections also locks deeper sections that
    follow. While you hold a lock, other clients' writes touching the section fail
    LOCKED_BY_OTHER_CLIENT (reads stay open); your own writes are confined to your locked sections
    (OUTSIDE_LOCKED_SECTION otherwise) and are NOT staled by others' edits elsewhere. Do not pass
    revision_before while locked: freshness is checked per section. A section heading must have
    unique text; if renamed or duplicated the lock fails closed (LOCK_SCOPE_UNRESOLVED) until the
    owner releases it. Leases last 1-300 s (lease_s) and renew by locking again; renewing never
    re-baselines. Revision, text-box-scope and other document-wide operations are refused under
    section locks. release=true releases the named sections (all yours if section_keys is empty).
    """
    from .live import comments_live, write_mode
    from .live.session import LiveError

    try:
        session = write_mode.live_session_for(path)
        if not hasattr(session, "client"):
            raise _make_error(ErrorCode.LIVE_UNAVAILABLE, "Enable shared bridge and call live_status first")
        if release:
            return session.client.rpc("section_unlock", **session.target(), sections=section_keys)
        return session.client.rpc(
            "section_lock",
            **session.target(),
            sections=section_keys,
            lease_s=lease_s,
            include_subsections=include_subsections,
            rpc_timeout=30,
        )
    except LiveError as exc:
        try:
            comments_live._raise_from_live_error(exc)
        except VerifyError as error:
            _raise_tool_error(error)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def live_status(path: str | None = None, verbose: bool = False) -> dict[str, Any]:
    """Report live-bridge state: whether it is running, its ports, and the connected Word task-pane
    sessions (all, or just path's). Read-only; never modifies a .docx.

    Starts the bridge lazily and idempotently; a pane need not be connected (sessions: [] is normal).
    path scopes sessions to that file name. verbose=True adds per-session heartbeat_count,
    hash_sample_count, requirement_sets plus recent_boots and broker shared blocks. In shared-bridge
    mode the result always carries locks (document leases and section locks, filtered to path's
    document when given) and this client's client_id: read them to see who holds a lock after
    LOCKED_BY_OTHER_CLIENT.

    Returns bridge_running, pid, owner_pid, is_owner, mode ("local"|"shared"), port, ops_port,
    sessions (each {document_name, document_url, connected_since, last_heartbeat_age_s, body_sha256,
    body_hash_state, body_hash_age_s, instance_id, platform, revision_count, revision_coverage}),
    session_collisions (two DIFFERENT documents that connected under the same file name; never a
    refusal itself, but check it before trusting a live write's routing) and warnings (e.g.
    pane_reload_loop) only when non-empty.

    Errors: LIVE_UNAVAILABLE (the bridge itself failed to start, e.g. no openssl on PATH or ports
    already bound); "no pane connected yet" is not an error.
    Full contract: docs/tool-contracts.md#live_status.
    """
    try:
        return execute_live_status(path, verbose)
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
        result: dict[str, Any] = {"path": str(resolved), "parts": projection.list_parts_impl(local_path)}
        warnings, live_session = reads_live.file_read_warnings(path)
        if warnings:
            result["warnings"] = warnings
            result["live_session"] = live_session
        return result
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool(annotations=_READ_ONLY)
def list_parts(path: str) -> dict[str, Any]:
    """List the package parts of a .docx that read_document(part=...) can target: the document body
    (word/document.xml), every header/footer, and footnotes/endnotes if present.

    Returns path, parts (list of {part, kind, header_footer_type}); header_footer_type is
    "default"|"first"|"even", None when unknown. Not gated by DOCX_LOCKED: reads a validated snapshot
    when Word's owner file is present.

    Errors: INVALID_INPUT (path missing or outside the allowed roots), SNAPSHOT_FAILED (snapshot
    could not be validated).
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
    source: str = "auto",
    start_para_ref: str | None = None,
    end_para_ref: str | None = None,
    to_file: str | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    if format not in _VALID_READ_FORMATS:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"format must be one of {sorted(_VALID_READ_FORMATS)}, got {format!r}",
            {"format": format},
        )
    if (start_para_ref is None) != (end_para_ref is None):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "start_para_ref and end_para_ref must be given together (an inclusive pN..pM range).",
            {"start_para_ref": start_para_ref, "end_para_ref": end_para_ref},
        )
    if start_para_ref is not None and section_key is not None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "A paragraph range cannot be combined with section_key (a section key already selects its own paragraphs).",
            {"section_key": section_key, "start_para_ref": start_para_ref},
        )
    # Resolve the destination BEFORE the read so a bad path costs nothing and
    # a live read is never half-performed for nothing.
    target = (
        paths.resolve_text_output_path(to_file, param="to_file", overwrite=overwrite) if to_file is not None else None
    )
    with reads_live.read_source(path, source, part) as rs:
        result, paragraph_count = _read_document_from(
            rs.local_path,
            rs.resolved,
            format,
            part,
            section_key,
            start_para_ref,
            end_para_ref,
            live=rs.source == "live",
        )
        if target is not None:
            result = _written_read_summary(rs.local_path, result, paragraph_count, format, part, target)
        return rs.annotate(result)


_TO_FILE_MAX_SECTIONS = 25


def _written_read_summary(
    local_path: Path, result: dict[str, Any], paragraph_count: int, format: str, part: str, target: Path
) -> dict[str, Any]:
    """Write the rendered body of a read to *target* and return only the
    summary (no body), plus the heading sections with their sizes."""
    rendered = result.pop(format)
    text = rendered if isinstance(rendered, str) else json.dumps(rendered, ensure_ascii=False)
    # Read the sections BEFORE writing: nothing may fail after the file is on disk.
    headings = [sec for sec in projection.find_sections_impl(local_path, part) if sec["kind"] == "heading"]
    data = paths.write_text_atomic(target, text)
    # Compact rows, capped: a 150-section document would otherwise make the
    # "no body" summary itself thousands of tokens. find_sections has the rest.
    sections = [
        [sec["section_key"], sec["start_para_ref"], sec["end_para_ref"], sec["chars"]]
        for sec in headings[:_TO_FILE_MAX_SECTIONS]
    ]
    summary: dict[str, Any] = {
        "path": result["path"],
        "part": part,
        "format": format,
        "section_key": result["section_key"],
        "revision": result["revision"],
        "revision_detail": result["revision_detail"],
        "written_to": str(target),
        "sha256": hashlib.sha256(data).hexdigest(),
        "chars": len(text),
        "paragraph_count": paragraph_count,
        "sections": sections,
        "section_count": len(headings),
        "warnings": result["warnings"],
    }
    if len(headings) > len(sections):
        summary["sections_truncated"] = True
    for key in ("lossy_elements", "start_para_ref", "end_para_ref"):
        if key in result:
            summary[key] = result[key]
    return summary


def _read_document_from(
    local_path: Path,
    resolved: Path,
    format: str,
    part: str,
    section_key: str | None,
    start_para_ref: str | None = None,
    end_para_ref: str | None = None,
    *,
    live: bool = False,
) -> tuple[dict[str, Any], int]:
    """The read body plus the number of paragraphs it covers."""
    revision = projection.compute_revision(local_path)
    result: dict[str, Any] = {
        "path": str(resolved),
        "part": part,
        "format": format,
        "section_key": section_key,
        "revision": revision["token"],
        "revision_detail": revision["detail"],
    }

    def render(proj: projection.Projection, extra_lossy: list[dict[str, Any]]) -> tuple[dict[str, Any], int]:
        if format == "text":
            result["text"] = proj.text
        elif format == "runs":
            result["runs"] = projection.runs_from_projection(proj)
        else:  # markdown
            markdown, _, lossy_elements = projection.markdown_from_projection(local_path, proj)
            result["markdown"] = markdown
            lossy_elements = lossy_elements + extra_lossy
            if lossy_elements:
                result["lossy_elements"] = lossy_elements
        result["warnings"] = proj.warnings
        return result, len(proj.paragraphs)

    if section_key is not None:
        # A textbox-<n> key reads that text box's own scope.
        scoped = projection.project_textbox_scope(local_path, part, section_key)
        # Otherwise a heading section_key: read that section's paragraph range
        # exactly as find_sections reports it (original para refs / table ids kept).
        sections = projection.find_sections_impl(local_path, part, ignore_trailing_empty=live)
        heading = next((s for s in sections if s["kind"] == "heading" and s["section_key"] == section_key), None)
        if scoped is not None and heading is not None:
            # A heading titled "Textbox" slugs to the same key as the first text box.
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"section_key {section_key!r} names both a heading section and a text box; read the heading "
                "with start_para_ref/end_para_ref from find_sections instead.",
                {"section_key": section_key, "part": part, "start_para_ref": heading["start_para_ref"],
                 "end_para_ref": heading["end_para_ref"]},
            )
        if scoped is not None:
            return render(scoped, [])
        if heading is None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"section_key {section_key!r} does not name a heading section or text box in part {part!r}.",
                {
                    "section_key": section_key,
                    "part": part,
                    "available_heading_keys": [s["section_key"] for s in sections if s["kind"] == "heading"],
                    "available_textbox_keys": [s["section_key"] for s in sections if s["kind"] == "textbox"],
                },
            )
        whole = projection.project_part(local_path, part)
        first, last = projection.resolve_para_range(whole, heading["start_para_ref"], heading["end_para_ref"])
        sliced, partial = projection.slice_projection(whole, first, last)
        result["start_para_ref"] = heading["start_para_ref"]
        result["end_para_ref"] = heading["end_para_ref"]
        return render(sliced, partial)

    whole = projection.project_part(local_path, part)
    if start_para_ref is not None and end_para_ref is not None:
        first, last = projection.resolve_para_range(whole, start_para_ref, end_para_ref)
        sliced, partial = projection.slice_projection(whole, first, last)
        result["start_para_ref"] = start_para_ref
        result["end_para_ref"] = end_para_ref
        return render(sliced, partial)

    return render(whole, [])


@mcp.tool(annotations=_READ_ONLY)
def read_document(
    path: str,
    format: str = "markdown",
    part: str = projection.DEFAULT_PART,
    section_key: str | None = None,
    source: str = "auto",
    start_para_ref: str | None = None,
    end_para_ref: str | None = None,
    to_file: str | None = None,
    overwrite: bool = False,
    verbose: bool = False,
) -> dict[str, Any]:
    """Read a .docx part's content as markdown, flat text, or a run-level structural list.

    format "markdown" (default; headings, bold/italic, nested lists, GFM tables, placeholders
    "[image:rId]", "[field:instr]"), "text", or "runs" (run records {text, rPr, para_ref} plus
    table_start/table_end, drawing and field records). Deleted text and field instructions are
    excluded. A merged cell or nested table is reported in lossy_elements, not reproduced.

    source "auto" (default) | "file" | "live". "live" reads the body the connected pane holds (default
    part only; needs body_ooxml; no images; never falls back to the file, a pane failure is an
    error; revision is "live:sha256:<hash>", the token live writes accept as revision_before). "auto"
    reads live when a pane session exists for this file name and can serve it, else the file. A file
    read while a pane session exists carries warning "live_session_ignored" and live_session
    {document_url, reason, action} (reason: requested_file, part_not_in_live, pane_missing_capability
    (reload the pane), basename_collision (confirm which document is open, then pass source="live"),
    session_mismatch). para_ref and table_id are positional: use them only with the same source and
    revision.

    part (default word/document.xml; see list_parts). section_key (find_sections) reads one heading
    section (any source; adds start_para_ref/end_para_ref; not proof of whole-document or comment
    freshness) or "textbox-<n>" alone; a text box's text is absent from a whole-part read; an
    unknown key is INVALID_INPUT listing the keys. start_para_ref + end_para_ref (both, inclusive,
    "pN") return that range; a cut table is lossy_elements partial_table; bad range INVALID_INPUT;
    not with section_key.
    to_file writes the rendering to an allowed path (.md/.txt/.json, parent must exist, existing file
    refused unless overwrite=true) and returns NO body: {path, part, format, revision, written_to,
    sha256, chars, paragraph_count, sections (first 25), section_count, sections_truncated}.

    Returns path, part, format, revision ("<doc8>:<cmt8>"), text|runs|markdown; warnings,
    lossy_elements when non-empty; verbose=True adds revision_detail.
    Errors: INVALID_INPUT (bad path, format or section_key), PART_NOT_FOUND, SNAPSHOT_FAILED. Not
    gated by DOCX_LOCKED.
    Full contract: docs/tool-contracts.md#read_document.
    """
    try:
        return execute_read_document(
            path, format, part, section_key, source, start_para_ref, end_para_ref, to_file, overwrite
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_find_sections(
    path: str, part: str = projection.DEFAULT_PART, source: str = "auto", digests: bool = False
) -> dict[str, Any]:
    with reads_live.read_source(path, source, part) as rs:
        trailing = rs.source == "live"
        unavailable: VerifyError | None = None
        try:
            sections, preamble = projection.find_sections_and_preamble(
                rs.local_path, part, digests=digests, content_digests=digests, ignore_trailing_empty=trailing
            )
        except VerifyError as exc:
            if not digests or exc.envelope.diagnostics.get("stage") != "content_digest":
                raise
            # A caller who wants sha12 must not lose it to comment/table canonicalization.
            unavailable = exc
            sections, preamble = projection.find_sections_and_preamble(
                rs.local_path, part, digests=True, ignore_trailing_empty=trailing
            )
            for sec in sections:
                if sec["kind"] == "heading":
                    sec["content_sha12"] = None
        result: dict[str, Any] = {"path": str(rs.resolved), "part": part, "sections": sections}
        if digests:
            result["preamble_content_sha12"] = preamble
        if unavailable is not None:
            result["warnings"] = ["content_digest_unavailable"]
            result["content_digest_unavailable"] = {
                "reason": unavailable.envelope.message,
                "diagnostics": unavailable.envelope.diagnostics,
            }
        return rs.annotate(result)


@mcp.tool(annotations=_READ_ONLY)
def find_sections(
    path: str, part: str = projection.DEFAULT_PART, source: str = "auto", digests: bool = False
) -> dict[str, Any]:
    """List heading-delimited section ranges AND text-box sub-scopes in a .docx part: the call that
    discovers every section_key read_document(section_key=...) can address. (For page-layout
    sections use list_page_sections.)

    source "auto"|"file"|"live" works as in read_document; the response carries "source". A section
    runs from one heading to the next heading at any level (or the end); headings inside table cells
    do not count. section_key = slug(heading text) + a 1-based ordinal for duplicates. Each text box
    is its own entry, kind "textbox", section_key "textbox-<n>" (heading_text, outline_level and
    para refs are None).

    Returns path, part, sections [{section_key, kind, heading_text, outline_level, start_para_ref,
    end_para_ref, paragraph_count, chars}]; chars is the plain-text size, so you can judge whether
    reading it is affordable. digests=true adds to each heading entry sha12 (first 12 hex of sha256
    of the section's paragraph texts joined by newline; agrees with live_list_sections compact and
    live_revision for unchanged text; what section locks compare) and content_sha12, plus top-level
    preamble_content_sha12: a digest of what a sync acts on (paragraph texts, tracked deletions,
    table structure, root comments with replies). Use them for "did the prose / content change".
    If comments or tables cannot be canonicalized, the content digests are null, sha12 is still
    returned, and warnings has "content_digest_unavailable".

    Errors: INVALID_INPUT, PART_NOT_FOUND, SNAPSHOT_FAILED. Not gated by DOCX_LOCKED.
    Full contract: docs/tool-contracts.md#find_sections.
    """
    try:
        return execute_find_sections(path, part, source, digests)
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


@mcp.tool(annotations=_READ_ONLY)
def list_page_sections(path: str, part: str = projection.DEFAULT_PART) -> dict[str, Any]:
    """List w:sectPr PAGE-LAYOUT sections in a .docx part (page size, margins, column widths, in
    inches). For heading sections use find_sections; the two are unrelated.

    Returns path, part, page_sections [{page_width_in, page_height_in, orientation, margin_top_in,
    margin_bottom_in, margin_left_in, margin_right_in, column_widths_in}]. Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT, PART_NOT_FOUND, SNAPSHOT_FAILED.
    """
    try:
        return execute_list_page_sections(path, part)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_list_styles(
    path: str, type: str | None = None, name_contains: str | None = None
) -> dict[str, Any]:
    if type is not None and type not in projection.STYLE_TYPES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"type must be one of {projection.STYLE_TYPES}, got {type!r}",
            {"type": type},
        )
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    local_path, is_temp = _read_local_copy(resolved)
    try:
        styles = projection.list_styles_impl(local_path)
        if type is not None or name_contains:
            styles = projection.filter_styles(styles, style_type=type, name_contains=name_contains)
        return {"path": str(resolved), "styles": styles}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


@mcp.tool(annotations=_READ_ONLY)
def list_styles(
    path: str,
    type: Literal["paragraph", "character", "table", "numbering"] | None = None,
    name_contains: str | None = None,
) -> dict[str, Any]:
    """Enumerate word/styles.xml's styles: ids, names, types, outline levels. Heading detection in
    find_sections and read_document resolves outline levels through this same list.

    type narrows to one style type; name_contains keeps styles whose name or style_id contains it
    (case-insensitive). Both default to no filter. Returns path, styles [{style_id, name, type,
    outline_lvl}]. Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT (path missing or outside the allowed roots), SNAPSHOT_FAILED.
    """
    try:
        return execute_list_styles(path, type, name_contains)
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


@mcp.tool(annotations=_DESTRUCTIVE)
def replace_body_markdown(
    path: str,
    markdown: str | None = None,
    markdown_file: str | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Replace an entire document body's content with markdown, atomically. Last resort: prefer
    replace_range_markdown when find_sections can locate the section.

    markdown or markdown_file (a .md/.txt path, UTF-8, <= 2 MB, under the allowed roots else
    DOCX_PATH_ESCAPE): exactly one (else INVALID_INPUT); the response adds markdown_sha256 and
    markdown_chars. Markdown renders against THIS document's styles (STYLE_NOT_FOUND if a heading
    level has no style); bold, italic, links, nested lists and pipe tables are supported.

    Guards: owner file -> DOCX_LOCKED (no retry); not sync-quiesced -> SYNC_IN_FLIGHT; stale
    revision_before -> REVISION_CONFLICT; file changed since this server's last write ->
    EXTERNAL_EDITOR_ACTIVE (bypass only that with allow_concurrent_editor=True). If the body has
    comment anchors or tracked changes it refuses (COMMENT_ANCHORS_IN_RANGE /
    TRACKED_CHANGES_PRESENT) unless force=True, which removes them and reports orphaned_comment_ids.
    The write is atomic: OPC-validated first (OPC_INVALID, original untouched), restored from .jsbak
    if post-write verification fails (VERIFICATION_FAILED). track_changes=True keeps old content as
    w:del and wraps new content in w:ins (untested with a table in range).

    Returns applied, match_count (1), rung (4), before, after, revision_before, revision_after,
    audit_logged, conflict_copy_detected, orphaned_comment_ids, revision_ids/track_changes.
    Errors: INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND, DOCX_LOCKED, SYNC_IN_FLIGHT,
    REVISION_CONFLICT (nothing written), EXTERNAL_EDITOR_ACTIVE, COMMENT_ANCHORS_IN_RANGE,
    TRACKED_CHANGES_PRESENT, STYLE_NOT_FOUND, OPC_INVALID, VERIFICATION_FAILED (rolled back, unless
    another writer changed the file: details.rollback_skipped).
    Full contract: docs/tool-contracts.md#replace_body_markdown.
    """
    try:
        text, extra = resolve_markdown_input(markdown, markdown_file)
        return extra | mutations.execute_replace_body_markdown(
            path,
            text,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def replace_range_markdown(
    path: str,
    section_key: str,
    markdown: str | None = None,
    markdown_file: str | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Replace one heading-delimited section (section_key from find_sections) with markdown,
    atomically: from its heading up to, not including, the next heading of any level.

    markdown / markdown_file as in replace_body_markdown (exactly one; markdown_sha256 and
    markdown_chars in the response). Same guards, hazard refusal (COMMENT_ANCHORS_IN_RANGE /
    TRACKED_CHANGES_PRESENT unless force=True), atomic write, revision_before,
    allow_concurrent_editor and track_changes as replace_body_markdown (tracked: old content kept as
    w:del, new as w:ins after it). Returns the eight evidence keys (rung 3), orphaned_comment_ids
    when force removed anchors, revision_ids/track_changes when tracked.

    Errors: as replace_body_markdown, plus SECTION_NOT_FOUND (section_key matches no current
    section; call find_sections).
    Full contract: docs/tool-contracts.md#replace_range_markdown.
    """
    try:
        text, extra = resolve_markdown_input(markdown, markdown_file)
        return extra | mutations.execute_replace_range_markdown(
            path,
            section_key,
            text,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def append_markdown(
    path: str,
    markdown: str | None = None,
    markdown_file: str | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Append markdown to the end of a document body (before its trailing w:sectPr), atomically.

    markdown / markdown_file as in replace_body_markdown (exactly one). Nothing existing is removed,
    so there is no hazard scan and force has no effect (accepted only for uniform caller code).
    Guards, revision_before, allow_concurrent_editor as in replace_body_markdown. track_changes=True
    wraps only the new content in w:ins. before/after are the whole body's markdown; rung is 4.
    Returns the eight evidence keys plus revision_ids/track_changes when tracked.

    Errors: INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND, DOCX_LOCKED, SYNC_IN_FLIGHT,
    REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE, STYLE_NOT_FOUND, OPC_INVALID, VERIFICATION_FAILED
    (rolled back).
    Full contract: docs/tool-contracts.md#append_markdown.
    """
    try:
        text, extra = resolve_markdown_input(markdown, markdown_file)
        return extra | mutations.execute_append_markdown(
            path,
            text,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
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
    find: str | None = None,
    replace: str | None = None,
    expected_matches: int | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
    within_row_containing: str | None = None,
    allow_concurrent_editor: bool = False,
    allow_comment_loss: bool = False,
    inherit_format: str = "replaced",
    replacement_segments: list[dict[str, Any]] | None = None,
    scope: str = "body",
    allow_comment_loss_ids: list[dict[str, Any]] | None = None,
    verbose: bool = False,
    *,
    expected_length: int | None = None,
    find_sha256: str | None = None,
    para_ref: str | None = None,
    expected_sha256: str | None = None,
) -> dict[str, Any]:
    """Replace every occurrence of `find` with `replace` (required), atomically.

    Target, exactly one way: plain `find` + `expected_matches` (required; MATCH_COUNT_MISMATCH on a
    different count); anchored `find` "<prefix>…<suffix>" (each side >= 12 chars) + `expected_length`
    and/or `find_sha256` (one span must pass, else AMBIGUOUS_SPAN / ZERO_MATCH; over 230 chars needs
    pane capability "anchored_find_v1"); or `para_ref` "pN" + `expected_sha256` (replaces the whole
    paragraph; PARA_HASH_MISMATCH). A match crossing a paragraph/table/cell boundary
    refuses (STRUCTURAL_BOUNDARY).

    `write_mode` "auto" (default) | "file" | "live": live edits the open document via the
    pane (LIVE_UNAVAILABLE if none); file refuses with LIVE_SESSION_ACTIVE while a pane is connected;
    auto goes live when connected. `revision_before` stale -> REVISION_CONFLICT (file) / LIVE_STALE
    (live), nothing written. `track_changes=True` writes w:del/w:ins; crossing someone else's tracked
    change refuses TRACKED_CHANGES_PRESENT unless `force=True`. `within_row_containing` limits find to
    the table row holding that other, unique text (live needs "row_scope").
    `allow_concurrent_editor=True` bypasses only EXTERNAL_EDITOR_ACTIVE. Live only:
    `replacement_segments` (formatted segments; omitted fields inherit from the first replaced run,
    pane "partial_font_segments_v1"; FORMAT_DRIFT if it was reformatted); `allow_comment_loss_ids`
    [{"comment_id","reply_count"}] consents to deleting those RESOLVED threads; other overlapping
    threads refuse WOULD_DELETE_COMMENTS.

    Other errors: INVALID_INPUT, DOCX_LOCKED, SYNC_IN_FLIGHT, OPC_INVALID, VERIFICATION_FAILED
    (file: rolled back), LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_CAPABILITY_MISSING,
    LIVE_STALE_AFTER_WRITE (write LANDED; do not retry).
    Full contract: docs/tool-contracts.md#replace_text.
    """
    try:
        if replace is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "replace is required (pass \"\" to delete the text)")
        return text_edit.execute_replace_text(
            path,
            find,
            replace,
            expected_matches,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            write_mode=write_mode,
            within_row_containing=within_row_containing,
            scope=scope,
            allow_comment_loss=allow_comment_loss,
            allow_comment_loss_ids=allow_comment_loss_ids,
            inherit_format=inherit_format,
            replacement_segments=replacement_segments,
            allow_concurrent_editor=allow_concurrent_editor,
            expected_length=expected_length,
            find_sha256=find_sha256,
            para_ref=para_ref,
            expected_sha256=expected_sha256,
        )
    except VerifyError as exc:
        _raise_tool_error(exc, verbose)


@mcp.tool()
def replace_text_batch(
    path: str,
    edits: list[dict[str, Any]],
    revision_before: str,
    write_mode: str = "live",
    track_changes: bool = False,
    inherit_format: str = "replaced",
    verbose: bool = False,
) -> dict[str, Any]:
    """Apply 1..50 replace_text edits to the LIVE Word document in order, in one call.

    LIVE ONLY: needs a connected pane (live_open_pane first); write_mode "live" or "auto", never
    "file". revision_before is REQUIRED ("live:sha256:<hex>" from a live read). `edits`: [{find,
    replace, expected_matches, allow_comment_loss_ids?, within_row_containing?}] with per-edit
    semantics as replace_text; an edit may instead use an anchored find (expected_length/find_sha256)
    or para_ref + expected_sha256, resolved against the body as the previous edit left it;
    allow_comment_loss inside an edit is refused. track_changes and inherit_format apply to all.

    One LIVE_STALE check runs first; each edit is then guarded by the body hash the previous one
    produced. The first failure stops the batch; NOTHING IS EVER RETRIED. Under a section lock the
    broker drops body-hash checks for the lock holder.

    Returns batch_status: "complete", "partial" (>=1 applied, then a known pre-write refusal),
    "none_applied" (document unchanged) or "unknown" (failure after dispatch that cannot be proven:
    disconnect, timeout, LIVE_OP_FAILED, VERIFICATION_FAILED, LIVE_STALE_AFTER_WRITE: re-read before
    doing anything). results [{index, status ("applied"|"refused_before_write"|"unknown"|
    "not_attempted"), evidence?, code?, message?}]; also revision_before, revision_after (null when
    unknown), applied (true only for "complete"), applied_count, edit_count. verbose=true adds
    per-edit before/after. One audit record per batch.
    Errors (nothing sent): INVALID_INPUT (bad edits, >50, non-live write_mode, malformed
    revision_before), LIVE_UNAVAILABLE, LIVE_CAPABILITY_MISSING, LIVE_DISCONNECTED.
    Full contract: docs/tool-contracts.md#replace_text_batch.
    """
    try:
        return text_edit.execute_replace_text_batch_live(
            path,
            edits,
            revision_before,
            write_mode=write_mode,
            track_changes=track_changes,
            inherit_format=inherit_format,
            verbose=verbose,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def format_text(
    path: str,
    find: str,
    style: dict[str, bool | str],
    expected_matches: int | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
    within_row_containing: str | None = None,
    allow_concurrent_editor: bool = False,
    scope: str = "body",
    verbose: bool = False,
    expected_length: int | None = None,
    find_sha256: str | None = None,
) -> dict[str, Any]:
    """Apply character styling to a matched text span without touching its content.

    `find` plain with expected_matches, or anchored with expected_length/find_sha256, exactly as in
    replace_text (same normalization ladder, STRUCTURAL_BOUNDARY, within_row_containing, scope).
    `style` maps "bold"/"italic"/"underline"/"strike" to true/false and/or "color" to six hex digits
    ("#" optional). Every requested field is applied verbatim (false clears); an explicit color also
    clears theme color attributes. Empty style, unknown key, non-boolean value or bad color is
    INVALID_INPUT. Idempotent: if every run already has the values, nothing is written
    (revision_before == revision_after).

    write_mode "auto"|"file"|"live" (default "auto") routes as in replace_text. In live mode a
    format-only op never changes the body hash, so revision_before == revision_after is expected,
    not a failure. track_changes=True records w:rPrChange with the prior formatting; foreign-authored
    change refuses TRACKED_CHANGES_PRESENT unless force=True. revision_before, allow_concurrent_editor
    as in replace_text.

    Returns the eight evidence keys with text_unchanged: true, one excerpt, format_delta (direct
    formatting that moved), warnings, revision_ids/track_changes when tracked; verbose=true adds
    runs_before/runs_after.

    Errors: as replace_text, plus INVALID_INPUT for a bad style.
    Full contract: docs/tool-contracts.md#format_text.
    """
    try:
        return text_edit.execute_format_text(
            path,
            find,
            style,
            expected_matches,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            write_mode=write_mode,
            within_row_containing=within_row_containing,
            scope=scope,
            allow_concurrent_editor=allow_concurrent_editor,
            verbose=verbose,
            expected_length=expected_length,
            find_sha256=find_sha256,
        )
    except VerifyError as exc:
        _raise_tool_error(exc, verbose)


# ---------------------------------------------------------------------------
# Tool: live_save (issue #106 WP-3:
# https://github.com/michaelrobertsutton/JennyStack/issues/106)
# ---------------------------------------------------------------------------


@mcp.tool()
def live_open_pane(path: str, timeout_s: float = 60) -> dict[str, Any]:
    """Make sure the Live pane is open and connected for this document (Word for Mac).

    Opens the document in Word if needed and drives Word's UI to open the pane; returns once
    connected, or at once if already connected. Needs macOS Accessibility permission for the calling
    app. For documents you will keep working on, follow with live_autoopen(path, enabled=true).
    timeout_s bounds the wait.

    Returns document_name, method, opened. Errors: LIVE_UNAVAILABLE (message says what to do by
    hand; diagnostics.reason is ribbon_unreachable | front_failed | not_mac | other, plus a hint;
    retryable only for ribbon_unreachable and front_failed: unlock the screen, bring Word to the
    front, retry).
    """
    from .live.open_pane_live import execute_open_pane

    try:
        return execute_open_pane(path, timeout_s)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def live_autoopen(path: str, enabled: bool | None = None) -> dict[str, Any]:
    """Read or set whether Word opens the Live pane automatically with this document.

    enabled=true (from a connected pane) writes Office.AutoShowTaskpaneWithDocument into the .docx so
    later opens show the pane. Omit enabled to only read. The setting marks the document as
    modified, so ask before tagging a document that is not disposable. It only works if the add-in
    is sideloaded or centrally deployed.

    Returns enabled, supported, was_enabled (when set), applied. Errors: LIVE_UNAVAILABLE (no pane
    connected), LIVE_CAPABILITY_MISSING (older pane or host without AddinCommands 1.1),
    VERIFICATION_FAILED (setting did not read back).
    """
    from .live.autoopen_live import execute_autoopen

    try:
        return execute_autoopen(path, enabled)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def live_save(path: str) -> dict[str, Any]:
    """Ask the connected Word task pane to save the live document. Word owns the open document during a
    live session: live edits reach the .docx only here.

    Returns applied, saved, document_name, revision_after ("live:sha256:<hex>"), file_revision (the
    file-mode revision token of the saved file: pass it as the next file-mode call's
    revision_before; null when path names no local file, e.g. an unsynced SharePoint document) and
    audit_logged. It confirms the save only, not the file's contents: verify structural edits
    afterwards with read_document/find_sections/diff_body_vs_file.

    Errors: LIVE_UNAVAILABLE (no pane), LIVE_DISCONNECTED (socket closed or save timed out),
    LIVE_OP_FAILED (pane replied ok=false), VERIFICATION_FAILED (ok but saved not true).
    Full contract: docs/tool-contracts.md#live_save.
    """
    try:
        return text_edit.execute_live_save(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tracked changes (WP-07): list_open_items (read-only), accept_tracked_
# changes / reject_tracked_changes (mutating, in MUTATING_TOOLS). Also
# where replace_text/format_text's own TRACKED_CHANGES_PRESENT refusal
# (text_edit._check_tracked_changes_guard) is wired in -- see that
# function and each tool's own updated docstring above.
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
def list_open_items(
    path: str,
    source: str = "auto",
    fields: Literal["summary", "full"] = "full",
    section_key: str | None = None,
    within_text: str | None = None,
    include_correlation: bool = False,
) -> dict[str, Any]:
    """List every open comment and pending tracked change (w:ins/w:del) in a .docx.

    fields "full" (default) | "summary" (columnar: counts, comment_columns with comment_rows,
    suggestion_columns with suggestion_rows; never `comments`, `suggestions` or
    `pending_suggestions` keys). source "auto" (default) | "file" | "live" (pane read;
    LIVE_UNAVAILABLE if none; auto prefers live when a session exists). section_key (find_sections)
    and/or within_text (must match exactly once) scope the result: counts adds in_scope, scoped is
    true, and items that cannot be placed are kept with scope_uncertain.

    Returns `comments` (ROOT comments only: comment_id, content, resolved, replies, quoted_text,
    author, created_time; resolved threads omitted) and `pending_suggestions` (suggestion_id, kind
    insertion|deletion, text, author, date, para_ref). comment_id is the durableId (raw w:id
    fallback), accepted by get_comment_thread/reply_to_comment/resolve_comment. Live mode: comment_id
    is "live:<id>" (does not survive a Word restart), with anchor_text, replies, and full reply
    threads. READ BEFORE TELLING A USER A COMMENT DOES NOT EXIST: a live read covers the document
    BODY only (not headers, footers or text boxes); counts {live_total, live_open, returned}, scope,
    pane and warnings are advisory and cannot certify completeness. If a user sees a comment not
    listed, say it is "not visible to the pane" and quote counts/scope/observed_at; never say it does
    not exist. A malformed pane reply raises LIVE_OP_FAILED. include_correlation=true adds
    `correlation` (best file-mode comment_id match per live comment; advisory only). With no local
    file pending_suggestions is null (never []) and file_side_available is false. Not gated by
    DOCX_LOCKED.

    Errors: INVALID_INPUT (bad path, source, fields or section_key), ZERO_MATCH /
    MATCH_COUNT_MISMATCH (within_text not exactly once), SNAPSHOT_FAILED, LIVE_UNAVAILABLE.
    Full contract: docs/tool-contracts.md#list_open_items.
    """
    try:
        return comments_live.execute_list_open_items(
            path,
            source,
            fields=fields,
            section_key=section_key,
            within_text=within_text,
            include_correlation=include_correlation,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def accept_tracked_changes(
    path: str,
    revision_ids: list[str] | None = None,
    revision_before: str | None = None,
    force: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
    table_id: int | None = None,
) -> dict[str, Any]:
    """Accept tracked changes (w:ins/w:del), atomically: all of them, or only revision_ids (from
    list_open_items pending_suggestions[].suggestion_id). Accepting an insertion makes it plain
    content; accepting a deletion removes the text permanently.

    table_id (live only): accept only changes inside that table (list_tables numbering); exclusive
    with revision_ids; file mode refuses it (LIVE_UNAVAILABLE); a pane without "revisions_scoped"
    refuses (LIVE_CAPABILITY_MISSING); refused under a section lock. Same guard, revision_before,
    force, write_mode and allow_concurrent_editor rules as replace_text.

    Returns the eight evidence keys (before/after are the WHOLE document text; rung "all" or
    "by_id"; match_count = elements processed) plus revision_ids processed.

    Errors: INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND, DOCX_LOCKED, SYNC_IN_FLIGHT,
    REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE, REVISION_ID_NOT_FOUND (available ids listed),
    OPC_INVALID, VERIFICATION_FAILED (rolled back).
    Full contract: docs/tool-contracts.md#accept_tracked_changes.
    """
    try:
        from .live.revisions_live import mutate_revisions

        if table_id is not None:
            from .live import write_mode as live_write_mode

            live_write_mode.require_live_only(path, write_mode, tool="accept_tracked_changes(table_id=...)")
        if comments_live._resolve_write_mode(path, write_mode) == "live":
            return mutate_revisions(path, "accept", revision_ids, revision_before, table_id)
        return tracked_changes.execute_accept_tracked_changes(
            path,
            revision_ids,
            revision_before=revision_before,
            force=force,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def reject_tracked_changes(
    path: str,
    revision_ids: list[str] | None = None,
    revision_before: str | None = None,
    force: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
    table_id: int | None = None,
) -> dict[str, Any]:
    """Reject tracked changes (w:ins/w:del), atomically: all, or only revision_ids. Rejecting an
    insertion removes it; rejecting a deletion restores the deleted text.

    table_id (live only): reject only changes inside that table; rejecting a tracked table insertion
    removes the table (evidence table_removed=true). Other parameters, guards, evidence and errors
    are as accept_tracked_changes.
    """
    try:
        from .live.revisions_live import mutate_revisions

        if table_id is not None:
            from .live import write_mode as live_write_mode

            live_write_mode.require_live_only(path, write_mode, tool="reject_tracked_changes(table_id=...)")
        if comments_live._resolve_write_mode(path, write_mode) == "live":
            return mutate_revisions(path, "reject", revision_ids, revision_before, table_id)
        return tracked_changes.execute_reject_tracked_changes(
            path,
            revision_ids,
            revision_before=revision_before,
            force=force,
            allow_concurrent_editor=allow_concurrent_editor,
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
def add_anchored_comment(
    path: str,
    quote: str,
    text: str,
    expected_matches: int,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Add a comment (`text`) anchored to a quoted passage (`quote`), atomically.

    `quote` is located like replace_text's find (normalization ladder, STRUCTURAL_BOUNDARY,
    MATCH_COUNT_MISMATCH) with `expected_matches`. File mode adds one comment per matched span
    (separate ids); the anchor is re-read and verified (VERIFICATION_FAILED, rolled back). Author
    comes from the configured author name.
    write_mode "auto" (default) | "file" | "live". "live" sends comment_add to the pane
    (LIVE_UNAVAILABLE if none) and creates exactly ONE comment on the FIRST match even when
    expected_matches > 1 (the count is still verified first); the author is always the signed-in
    Word user; the id is a "live:<id>" handle that does not survive a Word restart. "file" refuses
    with LIVE_SESSION_ACTIVE while a pane is connected for this document; "auto" goes live when one
    is. allow_concurrent_editor=True bypasses only EXTERNAL_EDITOR_ACTIVE.

    Returns the eight evidence keys (before/after unchanged), comment_ids (durableIds, one per
    span), comment_id (only when exactly one), conflict_copy_detected (file mode). Live evidence
    adds write_mode, verified_via, document_name, author, orphaned_comment_ids [] and rung "exact".

    Errors: INVALID_INPUT (bad path or empty quote), DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND,
    ZERO_MATCH, MATCH_COUNT_MISMATCH, LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH,
    LIVE_DISCONNECTED, LIVE_STALE. File mode only: DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT,
    EXTERNAL_EDITOR_ACTIVE, STRUCTURAL_BOUNDARY, OPC_INVALID, VERIFICATION_FAILED.
    Full contract: docs/tool-contracts.md#add_anchored_comment.
    """
    try:
        mode = comments_live._resolve_write_mode(path, write_mode)
        if mode == "live":
            return comments_live.execute_add_anchored_comment_live(path, quote, text, expected_matches)
        return comments.execute_add_anchored_comment(
            path, quote, text, expected_matches, allow_concurrent_editor=allow_concurrent_editor
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def add_anchored_comment_batch(
    path: str,
    comments: list[dict[str, Any]],
    write_mode: str = "auto",
    revision_before: str | None = None,
    allow_concurrent_editor: bool = False,
    verbose: bool = False,
) -> dict[str, Any]:
    """Add 1-50 anchored comments in ONE call instead of one add_anchored_comment per finding.

    comments: [{quote, content, expected_matches (default 1), section_key (optional; restricts the
    item to that find_sections section)}], validated up front. The first failure stops the batch and
    NOTHING IS EVER RETRIED. One audit record per call. write_mode, revision_before,
    allow_concurrent_editor as add_anchored_comment / replace_text.
    LIVE: one comment_add per item, chained on the body hash; earlier comments stay if a later item
    fails. Every item needs expected_matches == 1 (the pane comments only the FIRST match;
    INVALID_INPUT otherwise, nothing sent). With section_key the quote must occur in that section and
    nowhere else (ZERO_MATCH / AMBIGUOUS_SPAN). revision_before is checked once at the start
    (LIVE_STALE).
    FILE: all comments are staged and saved ONCE, guarded and verified; expected_matches comments
    every match; SECTION_NOT_FOUND for an unknown key. All-or-nothing: if an item fails nothing is
    saved (earlier items "not_persisted", the failing one its code, the rest "not_attempted",
    "none_applied"); a save that may have landed gives "unknown".

    Returns {applied (true only for "complete"), batch_status ("complete"|"partial"|"none_applied"|
    "unknown"), applied_count, item_count, revision_before, revision_after, results [{index, status
    ("applied"|"refused_before_write"|"unknown"|"not_attempted"|"not_persisted"), comment_id?, code?,
    message?}], evidence_ref}. "unknown": re-read the document before doing anything.
    Errors (nothing written): INVALID_INPUT, DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT,
    EXTERNAL_EDITOR_ACTIVE, LIVE_SESSION_ACTIVE, LIVE_UNAVAILABLE, LIVE_CAPABILITY_MISSING,
    LIVE_DISCONNECTED.
    Full contract: docs/tool-contracts.md#add_anchored_comment_batch.
    """
    try:
        return comment_batch.execute_add_anchored_comment_batch(
            path, comments, write_mode=write_mode, revision_before=revision_before,
            allow_concurrent_editor=allow_concurrent_editor, verbose=verbose,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def get_comment_thread(path: str, comment_id: str | None = None,
                       source: str = "auto", session_epoch: str | None = None,
                       match_spec: dict[str, str] | None = None) -> dict[str, Any]:
    """Read a comment and its direct replies by durableId (or a raw w:comment/@w:id, as list_open_items
    reports for files without commentsIds.xml). A multi-paragraph comment is keyed on its LAST
    paragraph.

    Returns comment_id, content, author, created_time, resolved, quoted_text, reply_count, replies
    (one level deep) and comment_id_resolved_via ("durableId"|"w_id"). If comment_id names a reply,
    only that reply is returned (replies=[]); the root is not walked to. Resolved comments can still
    be read. source ("auto"|"file"|"live"), session_epoch and match_spec address live comments.
    Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT (bad path, no comments, or comment_id matches nothing), SNAPSHOT_FAILED.
    """
    try:
        if comments_live._resolve_source(path, source) == "live":
            return comments_live.execute_get_comment_thread_live(path, comment_id, match_spec, session_epoch)
        if comment_id is None or match_spec is not None or session_epoch is not None:
            raise _make_error(ErrorCode.INVALID_INPUT, "File comment reads require comment_id only")
        return comments.execute_get_comment_thread(path, comment_id)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def reply_to_comment(
    path: str,
    comment_id: str | None = None,
    text: str = "",
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
    session_epoch: str | None = None,
    match_spec: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Reply to an existing comment (comment_id: durableId, raw w:id, or live handle) with `text`,
    atomically. The reply is a new comment anchored over the same text as the parent.

    write_mode "auto" (default) | "file" | "live". "live" sends comment_reply to the pane and accepts
    a "live:<id>" handle or a durableId/w:id resolved through list_open_items(source="live")
    correlation (INVALID_INPUT naming both id spaces if neither resolves); the reply is verified by
    re-listing; the author is always the signed-in Word user. "file" refuses with LIVE_SESSION_ACTIVE
    while a pane is connected; "auto" goes live when one is. session_epoch and match_spec are
    live-mode addressing. allow_concurrent_editor=True bypasses only EXTERNAL_EDITOR_ACTIVE.

    Returns the eight evidence keys (before/after are the parent's quoted text; rung "reply"),
    comment_id (the new reply), parent_comment_id, comment_id_resolved_via, conflict_copy_detected
    (file mode). Live evidence adds write_mode, verified_via, document_name, author; its revision_
    before/after are usually equal.

    Errors: INVALID_INPUT (bad path, unknown comment_id, parent has no anchor, or a live id that
    resolves nowhere), DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND, LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE,
    LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_OP_FAILED (reply not found on re-list). File mode
    only: DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE, OPC_INVALID,
    VERIFICATION_FAILED.
    Full contract: docs/tool-contracts.md#reply_to_comment.
    """
    try:
        mode = comments_live._resolve_write_mode(path, write_mode)
        if mode == "live":
            handle, address = comments_live.address_comment(path, comment_id, match_spec, session_epoch)
            return comments_live.execute_reply_to_comment_live(path, handle, text, address)
        if comment_id is None or match_spec is not None or session_epoch is not None:
            raise _make_error(ErrorCode.INVALID_INPUT, "File comment writes require comment_id only")
        return comments.execute_reply_to_comment(
            path, comment_id, text, allow_concurrent_editor=allow_concurrent_editor
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def resolve_comment(
    path: str,
    comment_id: str | None = None,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
    session_epoch: str | None = None,
    match_spec: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Resolve a comment thread (comment_id: durableId, raw w:id, or live handle), atomically. The
    thread is marked done, not deleted: list_open_items omits it afterwards, get_comment_thread still
    reads it. Idempotent: resolving an already-resolved comment succeeds.

    write_mode "auto" (default) | "file" | "live". "live" sends comment_resolve to the pane and
    accepts a "live:<id>" handle or a durableId/w:id resolved through list_open_items(source="live")
    correlation (INVALID_INPUT naming both id spaces if neither resolves); a re-list verifies it,
    and the resolver is the signed-in Word user. "file" refuses with LIVE_SESSION_ACTIVE while a pane
    is connected; "auto" goes live when one is. session_epoch and match_spec are live-mode
    addressing. allow_concurrent_editor=True bypasses only EXTERNAL_EDITOR_ACTIVE.

    Returns the eight evidence keys (before "open", after "resolved", rung "resolve"), comment_id,
    comment_id_resolved_via, conflict_copy_detected (file mode). Live evidence adds write_mode,
    verified_via, document_name, author.

    Errors: INVALID_INPUT (bad path, unknown comment_id, or an unresolvable live id),
    DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND, COMMENT_STILL_OPEN (the post-write re-read did not show
    it resolved), LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED.
    File mode only: DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE,
    OPC_INVALID, VERIFICATION_FAILED.
    Full contract: docs/tool-contracts.md#resolve_comment.
    """
    try:
        mode = comments_live._resolve_write_mode(path, write_mode)
        if mode == "live":
            handle, address = comments_live.address_comment(path, comment_id, match_spec, session_epoch)
            return comments_live.execute_resolve_comment_live(path, handle, address)
        if comment_id is None or match_spec is not None or session_epoch is not None:
            raise _make_error(ErrorCode.INVALID_INPUT, "File comment writes require comment_id only")
        return comments.execute_resolve_comment(
            path, comment_id, allow_concurrent_editor=allow_concurrent_editor
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def delete_comment(
    path: str,
    comment_id: str | None = None,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
    session_epoch: str | None = None,
    match_spec: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Delete a comment thread (root and every reply), atomically. The comment need not be resolved
    (resolve_comment only marks it done).

    Thread roots only: a reply id is INVALID_INPUT naming its root. comment_id: durableId, raw w:id or
    live "live:<id>"; session_epoch and match_spec are live addressing. write_mode and
    allow_concurrent_editor as resolve_comment. File mode re-reads and rolls back unless the thread is
    gone and every other comment and anchor is unchanged. Live mode sends comment_delete, verifies by
    re-listing, and CANNOT roll back.

    Returns the eight evidence keys (rung "delete"), comment_id, deleted_thread_count,
    deleted_reply_count, deleted_comment_ids, other_comments_unchanged; live adds scope "body"
    (comments outside the main body are not covered).
    Errors: INVALID_INPUT (bad path, unknown id, reply id; nothing deleted), DOCX_PATH_ESCAPE,
    DOCX_ROOT_NOT_FOUND, DOCX_LOCKED, SYNC_IN_FLIGHT, EXTERNAL_EDITOR_ACTIVE, OPC_INVALID /
    VERIFICATION_FAILED (file: rolled back; live: another comment changed, NOT rolled back),
    COMMENT_STILL_PRESENT (live re-list still shows it), LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE,
    LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_CAPABILITY_MISSING (reopen the pane),
    COMMENT_ID_STALE (handle from an earlier pane epoch, or identity changed), LIVE_OP_FAILED
    (diagnostics.deleted/remaining/unknown say what is known; a part-way delete is never rolled back).
    Full contract: docs/tool-contracts.md#delete_comment.
    """
    try:
        mode = comments_live._resolve_write_mode(path, write_mode)
        if mode == "live":
            handle, address = comments_live.address_comment(path, comment_id, match_spec, session_epoch)
            return comments_live.execute_delete_comment_live(path, comment_id, handle, address)
        if comment_id is None or match_spec is not None or session_epoch is not None:
            raise _make_error(ErrorCode.INVALID_INPUT, "File comment writes require comment_id only")
        return comments.execute_delete_comments(
            path,
            [comment_id],
            tool_name="delete_comment",
            require_resolved=False,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def delete_resolved_comments(
    path: str,
    comment_ids: list[str],
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Delete several RESOLVED comment threads (comment_ids) in one call; use after resolve_comment.

    Every id is resolved, deduplicated by thread root and validated BEFORE anything is deleted: an
    empty list, unknown id, reply id or still-open thread refuses the whole call (INVALID_INPUT,
    diagnostics.offenders lists every one; nothing deleted). File mode deletes all in one atomic
    write. Live mode resolves against one pane snapshot, deletes in turn and verifies with one
    re-list; if a delete fails part-way the error's diagnostics list deleted, remaining and unknown
    handles and nothing is rolled back. write_mode and allow_concurrent_editor as delete_comment.
    Evidence is delete_comment's, with match_count/deleted_thread_count = threads deleted.

    Errors: as delete_comment (INVALID_INPUT also covers an empty list or a still-open thread).
    Full contract: docs/tool-contracts.md#delete_resolved_comments.
    """
    try:
        mode = comments_live._resolve_write_mode(path, write_mode)
        if mode == "live":
            return comments_live.execute_delete_resolved_comments_live(path, comment_ids)
        return comments.execute_delete_comments(
            path,
            comment_ids,
            tool_name="delete_resolved_comments",
            require_resolved=True,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tables (issue #28 WP-14): list_tables, get_table (read-only), and
# replace_table_row / replace_cell_markdown / insert_table (mutating, in
# MUTATING_TOOLS). table_id/row_index/cell_index are 1-based and match
# read_document(format="runs")'s own container_chain addressing.
# ---------------------------------------------------------------------------


@mcp.tool(annotations=_READ_ONLY)
def list_tables(
    path: str, part: str = projection.DEFAULT_PART, source: str = "auto"
) -> dict[str, Any]:
    """Enumerate every w:tbl in a .docx part, including nested tables (each gets its own table_id).

    source "auto"|"file"|"live" works as in read_document (response carries "source"; a file read
    with a live pane session carries "live_session_ignored"). table_id is positional (1-based,
    document order): use it only with the same source and revision.
    Returns path, part, tables [{table_id, row_count, col_count, has_merged_cells,
    has_nested_table, nested_in_table_id}]. Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT, PART_NOT_FOUND, SNAPSHOT_FAILED.
    """
    try:
        return tables.execute_list_tables(path, part, source)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def get_table(
    path: str, table_id: int, part: str = projection.DEFAULT_PART, source: str = "auto", compact: bool = True
) -> dict[str, Any]:
    """One table's content, addressed by the table_id list_tables reports.

    compact=true (default): rows is a list of rows of cell-text strings and merged lists only merged
    cells {row_index, cell_index, grid_span, v_merge}. compact=false: rows of {row_index,
    cell_index, grid_span, v_merge ("none"|"restart"|"continue"), text} per cell, no merged key.
    Cell text is per-cell markdown (a nested table flattens to inline text).
    source "auto" (default) | "file" | "live". LIVE READS DIFFER: main body only (a non-default part
    is INVALID_INPUT with source="live"; "auto" falls back to the file); text is Word's plain text;
    per-cell grid_span/v_merge are unavailable (use has_merged_cells; live compact reads of merged
    tables give merged null plus warning "live_merge_metadata_unavailable"); a document with a nested
    table is refused; the pane must report "table_edit". The result adds source, style,
    header_row_count. A file read with a pane session carries "live_session_ignored" and a
    live_session object. Live table_id agrees with list_tables only when there is no nested table.

    Returns path, part, table_id, row_count, col_count, has_merged_cells, has_nested_table, rows.
    Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT, PART_NOT_FOUND, TABLE_NOT_FOUND (call list_tables), SNAPSHOT_FAILED,
    LIVE_UNAVAILABLE / LIVE_CAPABILITY_MISSING / LIVE_DISCONNECTED / LIVE_OP_FAILED (source="live").
    Full contract: docs/tool-contracts.md#get_table.
    """
    try:
        return tables.execute_get_table(path, table_id, part, source=source, compact=compact)
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
    allow_concurrent_editor: bool = False,
    write_mode: str = "auto",
) -> dict[str, Any]:
    """Replace one table row's content, one markdown string per cell in `cells` (exactly the row's cell
    count). row_index is 1-based (row 1 is usually the header).

    Refuses for the WHOLE table if it has a merged cell or a nested table anywhere
    (MERGED_OR_NESTED_TABLE); use replace_cell_markdown. Hazard scan (COMMENT_ANCHORS_IN_RANGE /
    TRACKED_CHANGES_PRESENT unless force=True), revision_before, allow_concurrent_editor,
    track_changes (old w:del, new w:ins) as the other write tools.
    write_mode "auto" (default) | "file" | "live". Live is a SUBSET: bold/italic/link cells only;
    force does not apply; before/after are plain text; needs "table_edit". The row is one Word batch
    guarded by a body-hash check, then verified by re-reading; batches are not transactions, so under
    heavy co-authoring a cell edited in the last instant can still be overwritten. A write that lands
    but fails read-back is audit-logged before VERIFICATION_FAILED.

    Returns the eight evidence keys (before/after = the row's cells tab-joined).
    Errors: INVALID_INPUT (bad path or wrong cells length), DOCX_LOCKED, SYNC_IN_FLIGHT,
    REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE, TABLE_NOT_FOUND, TABLE_ROW_NOT_FOUND,
    MERGED_OR_NESTED_TABLE, COMMENT_ANCHORS_IN_RANGE, TRACKED_CHANGES_PRESENT, OPC_INVALID,
    VERIFICATION_FAILED, LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH,
    LIVE_DISCONNECTED, LIVE_STALE, LIVE_CAPABILITY_MISSING, LIVE_OP_FAILED (also a cell edited
    between read and write, or nested tables).
    Full contract: docs/tool-contracts.md#replace_table_row.
    """
    try:
        return tables.execute_replace_table_row(
            path,
            table_id,
            row_index,
            cells,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
            write_mode=write_mode,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def replace_cell_markdown(
    path: str,
    table_id: int,
    row_index: int,
    cell_index: int,
    markdown: str | None = None,
    markdown_file: str | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
    write_mode: str = "auto",
    expected_before: str | None = None,
    inherit_format: str = "first_run",
) -> dict[str, Any]:
    """Replace one table cell's content with markdown, leaving its w:tcPr byte-identical: the only write
    path safe on a merged cell. row_index and cell_index are 1-based.

    markdown / markdown_file: exactly one (a .md/.txt path under the allowed roots, <= 2 MB; response
    adds markdown_sha256, markdown_chars). Lists may nest. expected_before (the cell's current text,
    modulo marks and whitespace) guards a wrong index or stale read: CELL_TEXT_MISMATCH, nothing
    written. inherit_format "first_run" (default) | "none": rendered runs take the font name, size
    and colour of the cell's first text run where the markdown sets none (live needs
    "cell_font_inheritance_v1", or pass "none").
    write_mode "auto" (default) | "file" | "live" (live edits a cell in a document a co-author has
    open; a file write races their autosave). LIVE IS A SUBSET: paragraphs, hard line breaks,
    bold/italic/links only (lists, headings, tables -> INVALID_INPUT, nothing sent); more than one
    paragraph or a line break needs "cell_multiline"; compare-and-set against the text just read, so
    a concurrent edit is refused (LIVE_OP_FAILED); a failed read-back restores the old content
    (VERIFICATION_FAILED, diagnostics.rolled_back=true; false means look at the cell); a nested
    table is refused; force does not apply; needs "cell_edit"; revisions are "live:sha256:<hex>".
    revision_before, force (COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT hazards),
    track_changes (old w:del, new w:ins), allow_concurrent_editor as replace_table_row.

    Returns the eight evidence keys (before/after = cell markdown).
    Errors: INVALID_INPUT, DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE,
    TABLE_NOT_FOUND, TABLE_ROW_NOT_FOUND, TABLE_CELL_NOT_FOUND, CELL_TEXT_MISMATCH,
    COMMENT_ANCHORS_IN_RANGE, TRACKED_CHANGES_PRESENT, OPC_INVALID, VERIFICATION_FAILED (file: rolled
    back, incl. tcPr drift), LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH,
    LIVE_DISCONNECTED, LIVE_STALE, LIVE_CAPABILITY_MISSING, LIVE_OP_FAILED.
    Full contract: docs/tool-contracts.md#replace_cell_markdown.
    """
    try:
        text, extra = resolve_markdown_input(markdown, markdown_file)
        return extra | tables.execute_replace_cell_markdown(
            path, table_id, row_index, cell_index, text,
            revision_before=revision_before, force=force, track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor, write_mode=write_mode,
            expected_before=expected_before, inherit_format=inherit_format,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def insert_table(
    path: str,
    rows: list[list[str | dict[str, Any]]] | None = None,
    markdown_file: str | None = None,
    style_id: str | None = None,
    header_rows: int = 0,
    grid_dxa: list[int] | None = None,
    cant_split: bool = False,
    anchor: dict[str, Any] | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
    write_mode: str = "auto",
    style_from_table_id: int | None = None,
    style_builtin: str | None = None,
    font_size_pt: float | None = None,
    like_table_id: int | None = None,
) -> dict[str, Any]:
    """Insert a new table, one markdown string or cell-spec object per cell.

    rows: list of rows of cells; a string cell means {"markdown": s}; short rows are padded. A cell
    object: {"markdown" (required), "span" (int >= 1), "v_merge" ("restart"|"continue", markdown
    must be ""), "fill" (6-hex), "color" (6-hex), "bold", "align" (left|center|right|both), "valign"
    (top|center|bottom)}. With any object cell, every row's spans must sum to the same column count
    (INVALID_INPUT). fill/valign/span/v_merge survive a later replace_cell_markdown;
    bold/color/align/font_size_pt are replaced by it. markdown_file: a .md/.txt GFM pipe-table file
    (UTF-8, <= 2 MB, allowed roots) instead of rows; exactly one of the two.
    style_id (or style_from_table_id / like_table_id) is REQUIRED in file mode and must name a table
    style (list_styles; STYLE_NOT_FOUND); style_builtin is live only. header_rows (default 0),
    grid_dxa (column widths), cant_split, font_size_pt. anchor (default: append at end):
    {"section_key", "position": "start"|"end"}, {"section_key", "after_paragraph_text"} (exactly one
    top-level paragraph else ZERO_MATCH / MATCH_COUNT_MISMATCH) or {"after_table_id": N}.
    write_mode "auto" (default) | "file" | "live" (live adds a table to a document a co-author has
    open). LIVE IS A SUBSET: bold/italic/link cells only (lists, headings, nested tables, span,
    v_merge, cant_split -> INVALID_INPUT, nothing sent); anchor is {"paragraph_text", "position"} or
    {"after_table_id"} or omitted; style_id is a Word table style NAME, resolved before inserting
    (STYLE_NOT_FOUND leaves the document untouched); force does not apply; needs
    "table_edit"; refuses LIVE_STALE if the document moved. Every property is read back: a table
    that lands but fails is audit-logged and raised VERIFICATION_FAILED (not rolled back; fix or
    delete it).
    revision_before, force, track_changes (w:ins), allow_concurrent_editor as append_markdown.

    Returns the eight evidence keys, table_id (the new table), merged_cells, anchor_resolved.
    Errors: INVALID_INPUT, DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE,
    STYLE_NOT_FOUND, SECTION_NOT_FOUND, TABLE_NOT_FOUND, ZERO_MATCH, MATCH_COUNT_MISMATCH,
    OPC_INVALID, VERIFICATION_FAILED (file: rolled back), LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE,
    LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_STALE, LIVE_CAPABILITY_MISSING, LIVE_OP_FAILED.
    Full contract: docs/tool-contracts.md#insert_table.
    """
    try:
        table_rows, extra = resolve_table_rows(rows, markdown_file)
        return extra | tables.execute_insert_table(
            path,
            table_rows,
            style_id,
            header_rows=header_rows,
            grid_dxa=grid_dxa,
            cant_split=cant_split,
            anchor=anchor,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
            write_mode=write_mode,
            style_from_table_id=style_from_table_id,
            style_builtin=style_builtin,
            font_size_pt=font_size_pt,
            like_table_id=like_table_id,
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
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Append a new inline picture at the end of the document body from a LOCAL .png or .svg file
    (image_path). width_in defaults to the first page section's text-column width; height follows the
    aspect ratio. An .svg is embedded natively with a PNG fallback rasterized via macOS `sips`.

    Returns the eight evidence keys (before "", after "[image:<rId>]"), format, design_width_in,
    design_height_in, placed_width_in, placed_height_in, effective_scale (placed / design), blip_rid
    (svg_rid for an svg), revision_ids/track_changes, conflict_copy_detected. Guard,
    revision_before, force, track_changes (w:ins) and allow_concurrent_editor as insert_table.

    Errors: INVALID_INPUT (bad path, image_path not a regular file, width_in <= 0), DOCX_LOCKED,
    SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE, UNSUPPORTED_IMAGE_FORMAT,
    SVG_RASTERIZATION_FAILED, OPC_INVALID, VERIFICATION_FAILED (rolled back).
    """
    try:
        return images.execute_insert_image(
            path,
            image_path,
            width_in,
            revision_before=revision_before,
            force=force,
            track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def apply_style(
    path: str,
    find: str,
    style_id: str,
    expected_matches: int | None = None,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    allow_concurrent_editor: bool = False,
    verbose: bool = False,
    expected_length: int | None = None,
    find_sha256: str | None = None,
) -> dict[str, Any]:
    """Apply a NAMED style (style_id from list_styles) to text located via find: the named-style
    counterpart to format_text.

    `find` plain with expected_matches, or anchored with expected_length/find_sha256, as in
    replace_text. A CHARACTER style applies to the matched runs like format_text. A PARAGRAPH style
    applies w:pStyle to every paragraph CONTAINING a match (the whole paragraph). The text never
    changes: the result has text_unchanged: true, one excerpt, format_delta and style_id;
    verbose=true adds runs_before/runs_after.
    NO write_mode: it always writes the .docx and refuses with LIVE_SESSION_ACTIVE while a pane is
    connected for this document (a live pane's autosave would silently revert it). For a live edit
    use format_text/replace_text(write_mode="live"), or save and close the document in Word first.
    track_changes=True records w:rPrChange (character) or w:pPrChange (paragraph, shaped from real
    Word output). A paragraph or run already carrying someone else's change record refuses
    TRACKED_CHANGES_PRESENT unless force=True. accept_/reject_tracked_changes do NOT handle
    w:rPrChange/w:pPrChange. revision_before, allow_concurrent_editor as format_text.

    Returns the eight evidence keys, runs_before/runs_after, style_id, style_type, warnings,
    revision_ids/track_changes (empty when it reused an own change record).

    Errors: as format_text's file-mode errors, plus STYLE_NOT_FOUND, UNSUPPORTED_STYLE_TYPE (neither
    paragraph nor character style), TRACKED_CHANGES_PRESENT, LIVE_SESSION_ACTIVE.
    Full contract: docs/tool-contracts.md#apply_style.
    """
    try:
        return text_edit.execute_apply_style(
            path, find, style_id, expected_matches,
            revision_before=revision_before, force=force, track_changes=track_changes,
            allow_concurrent_editor=allow_concurrent_editor, verbose=verbose,
            expected_length=expected_length, find_sha256=find_sha256,
        )
    except VerifyError as exc:
        _raise_tool_error(exc, verbose)


@mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
def read_audit_evidence(
    evidence_ref: str, field: str | None = None, to_file: str | None = None, overwrite: bool = False
) -> dict[str, Any]:
    """Fetch the full audit record behind an evidence_ref a write returned.

    Large writes return hashes, sizes and a short excerpt; the text is in the append-only audit log
    ($XDG_STATE_HOME/verified-docx-mcp/audit.jsonl, default ~/.local/state/verified-docx-mcp/).
    If the operator set VERIFIED_DOCX_MCP_AUDIT_EXCERPTS=0 the text was never stored and
    before/after read "[redacted; N chars]". field (e.g. "before", "after") returns just that field.
    to_file writes the result to an allowed path (.md, .txt or .json; an existing file is refused
    unless overwrite=true) and returns only {written_to, sha256, chars}.

    Errors: INVALID_INPUT (unknown evidence_ref or field, bad to_file).
    """
    try:
        return execute_read_audit_evidence(evidence_ref, field, to_file, overwrite)
    except VerifyError as exc:
        _raise_tool_error(exc)


def execute_read_audit_evidence(
    evidence_ref: str, field: str | None = None, to_file: str | None = None, overwrite: bool = False
) -> dict[str, Any]:
    target = (
        paths.resolve_text_output_path(to_file, param="to_file", overwrite=overwrite) if to_file is not None else None
    )
    record = audit.find_audit_record(evidence_ref)
    if record is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"No audit record with evidence_ref {evidence_ref!r}. The ref is only valid on the machine "
            "and state directory that wrote it, and only while its audit log still holds the record.",
            {"evidence_ref": evidence_ref},
        )
    evidence = record.get("evidence", {})
    if field is not None and field not in evidence:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"The audit record has no evidence field {field!r}.",
            {"evidence_ref": evidence_ref, "available_fields": sorted(evidence)},
        )
    result: dict[str, Any] = {
        "evidence_ref": evidence_ref,
        "tool": record.get("tool"),
        "path": record.get("path"),
        "timestamp": record.get("timestamp"),
    }
    payload: Any = evidence if field is None else evidence[field]
    if field is None:
        result["evidence"] = payload
    else:
        result["field"] = field
        result["value"] = payload
    if target is None:
        return result
    text = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)
    data = paths.write_text_atomic(target, text)
    return {
        **{k: result[k] for k in ("evidence_ref", "tool", "path", "timestamp")},
        "written_to": str(target),
        "sha256": hashlib.sha256(data).hexdigest(),
        "chars": len(text),
    }


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


@mcp.tool(annotations=_READ_ONLY)
def read_header_footer(path: str) -> dict[str, Any]:
    """Read every header/footer part's content as markdown in one call (the wrapper over list_parts +
    read_document(part=...)). Returns path, headers_and_footers [{part, kind, header_footer_type,
    markdown, warnings, lossy_elements}]; empty list when there are none. Not gated by DOCX_LOCKED.

    Errors: INVALID_INPUT, SNAPSHOT_FAILED.
    """
    try:
        return execute_read_header_footer(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


# ---------------------------------------------------------------------------
# Tool: diff_body_vs_file
# ---------------------------------------------------------------------------


_VALID_DIFF_FORMATS = frozenset({"hunks", "unified"})
# Outline levels 6-8 render as 7-9 "#" characters (level + 1), so accept up to nine.
_MD_HEADING_RE = re.compile(r"^ {0,3}#{1,9}[ \t]+(.+?)[ \t]*#*[ \t]*$")


def _file_section_slice(lines: list[str], section_key: str) -> tuple[list[str], int]:
    """The lines of the markdown file's section matching *section_key* (a
    find_sections heading key: slug-ordinal) and the 0-based index of its
    first line. Headings match by folded text (slug), the same ordinal
    rule find_sections uses for duplicates; the section runs to the next
    heading of any level. INVALID_INPUT when the file has no such heading."""
    base, _, ordinal_text = section_key.rpartition("-")
    ordinal = int(ordinal_text) if ordinal_text.isdigit() else 1
    base = base or section_key
    headings: list[int] = []
    in_fence = False
    for idx, line in enumerate(lines):
        if line.lstrip().startswith(("```", "~~~")):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        m = _MD_HEADING_RE.match(line.rstrip("\r\n"))
        if m:
            headings.append(idx)
    def slug_matches(idx: int) -> bool:
        text = _MD_HEADING_RE.match(lines[idx].rstrip("\r\n")).group(1)
        # find_sections slugged the docx's raw heading text; a converted file may have folded it
        # (soft hyphens, NFC), so accept either spelling.
        return base in (projection._slugify(text), projection._slugify(locate.normalize_for_compare(text)))

    matching = [idx for idx in headings if slug_matches(idx)]
    if ordinal < 1 or ordinal > len(matching):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"The markdown file has no heading matching section_key {section_key!r}.",
            {"section_key": section_key},
        )
    start = matching[ordinal - 1]
    later = [h for h in headings if h > start]
    return lines[start : later[0] if later else len(lines)], start


def execute_diff_body_vs_file(
    path: str,
    file_path: str,
    normalize: bool = True,
    format: str = "hunks",
    section_key: str | None = None,
    max_hunks: int = 50,
) -> dict[str, Any]:
    """Diff the docx body's markdown projection against a local markdown
    file -- issue #28 WP-11a, same ``difflib`` mechanism as
    GoogleDocs-MCP's ``diff_tab_vs_file`` (markdown_mutations.py
    ``execute_diff_tab_vs_file``), reshaped by issue #110 so the response
    is O(diff) rather than O(document).

    This is a READ tool -- no mutation, no audit, no evidence envelope
    required (mirrors Google's own "no changes to the document or file"
    framing for that tool).
    """
    if format not in _VALID_DIFF_FORMATS:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"format must be one of {sorted(_VALID_DIFF_FORMATS)}, got {format!r}",
            {"format": format},
        )
    if max_hunks < 1:
        raise _make_error(ErrorCode.INVALID_INPUT, "max_hunks must be at least 1", {"max_hunks": max_hunks})
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
        whole = projection.project_part(local_path, projection.DEFAULT_PART)
        file_content = file_resolved.read_text(encoding="utf-8")
        file_line_offset = 0
        extra_lossy: list[dict[str, Any]] = []

        if section_key is None:
            body_markdown, warnings, lossy_elements = projection.markdown_from_projection(local_path, whole)
        else:
            entry = next(
                (
                    sec
                    for sec in projection.find_sections_impl(local_path)
                    if sec["kind"] == "heading" and sec["section_key"] == section_key
                ),
                None,
            )
            if entry is None:
                raise _make_error(
                    ErrorCode.SECTION_NOT_FOUND,
                    f"section_key {section_key!r} does not name a heading section (call find_sections).",
                    {"section_key": section_key},
                )
            sliced, extra_lossy = projection.slice_projection(
                whole, projection.para_index(entry["start_para_ref"]), projection.para_index(entry["end_para_ref"])
            )
            body_markdown, warnings, lossy_elements = projection.markdown_from_projection(local_path, sliced)
            section_lines, file_line_offset = _file_section_slice(file_content.splitlines(keepends=True), section_key)
            file_content = "".join(section_lines)
        lossy_elements = lossy_elements + extra_lossy

        identical_raw = body_markdown == file_content
        if normalize:
            body_cmp = locate.normalize_for_compare(body_markdown).rstrip("\n")
            file_cmp = locate.normalize_for_compare(file_content).rstrip("\n")
            # Give both sides the same final newline, so the last line of an otherwise
            # unchanged file never differs only by its line ending (and a unified diff
            # never glues two lines together).
            body_cmp += "\n" if body_cmp else ""
            file_cmp += "\n" if file_cmp else ""
        else:
            body_cmp, file_cmp = body_markdown, file_content
        identical = body_cmp == file_cmp

        body_lines = body_cmp.splitlines(keepends=True)
        file_lines = file_cmp.splitlines(keepends=True)
        matcher = difflib.SequenceMatcher(None, body_lines, file_lines)

        result: dict[str, Any] = {
            "path": str(resolved),
            "file_path": file_path,
            "revision": revision["token"],
            "identical": identical,
            "identical_raw": identical_raw,
            "normalized": normalize,
            "warnings": warnings,
        }
        if section_key is not None:
            result["section_key"] = section_key
            result["file_line_offset"] = file_line_offset

        if format == "hunks":
            opcodes = matcher.get_opcodes()
            changed = [op for op in opcodes if op[0] != "equal"]
            hunks: list[dict[str, Any]] = []
            if changed:
                emitted = 0
                for tag, i1, i2, j1, j2 in opcodes:
                    if tag == "equal":
                        hunks.append(
                            {
                                "tag": "equal",
                                "body_lines_count": i2 - i1,
                                "file_lines_count": j2 - j1,
                                "body_range": [i1 + 1, i2],
                                "file_range": [j1 + 1, j2],
                            }
                        )
                        continue
                    if emitted >= max_hunks:
                        break
                    emitted += 1
                    hunks.append(
                        {
                            "tag": tag,
                            "body_lines": body_lines[i1:i2],
                            "file_lines": file_lines[j1:j2],
                            "body_range": [i1 + 1, i2],
                            "file_range": [j1 + 1, j2],
                        }
                    )
                # A trailing run of equal hunks after the cap adds nothing.
                while hunks and hunks[-1]["tag"] == "equal" and emitted >= max_hunks and emitted < len(changed):
                    hunks.pop()
                result["truncated"] = len(changed) > max_hunks
                if result["truncated"]:
                    result["remaining_hunks"] = len(changed) - max_hunks
            else:
                result["truncated"] = False
            result["hunks"] = hunks
        else:
            unified = list(
                difflib.unified_diff(
                    body_lines,
                    file_lines,
                    fromfile=f"docx:{path}#body",
                    tofile=file_path,
                )
            )
            groups: list[list[str]] = []
            header: list[str] = []
            for line in unified:
                if line.startswith("@@"):
                    groups.append([line])
                elif groups:
                    groups[-1].append(line)
                else:
                    header.append(line)
            kept = groups[:max_hunks]
            result["unified_diff"] = "".join(header) + "".join("".join(g) for g in kept)
            result["truncated"] = len(groups) > max_hunks
            if result["truncated"]:
                result["remaining_hunks"] = len(groups) - max_hunks
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


@mcp.tool(annotations=_READ_ONLY)
def diff_body_vs_file(
    path: str,
    file_path: str,
    normalize: bool = True,
    format: str = "hunks",
    section_key: str | None = None,
    max_hunks: int = 50,
) -> dict[str, Any]:
    """Export the docx body as markdown and diff it against a local file (file_path). Read-only; no
    evidence keys. Both paths go through the allowed-roots check.

    AFFIRMATIVE-ONLY: a hunk or identical=False is solid evidence of a difference; identical=True is
    NOT proof of none. The docx side is a markdown projection of the BODY only: content it cannot
    express (merged cells, nested tables: see lossy_elements) is dropped from the comparison,
    headers/footers/footnotes/comments are out of scope, and with Word's owner file present it reads
    a snapshot that may lag. Treat identical=True as "no difference found by this projection", and a
    non-empty lossy_elements as a reason to distrust it.
    normalize=true (default) folds quotes, dashes, NBSP, soft hyphens, NFC and trailing whitespace
    first; identical is judged on the folded text, identical_raw is the byte-for-byte comparison
    (use it, or normalize=false, for "nothing differs at all"). format "hunks" (default) | "unified"
    (never both); equal hunks carry no text. section_key restricts both sides to one heading section
    (file side: the heading whose folded text matches; file_line_offset reports where it starts).
    max_hunks (default 50) caps changed hunks; truncated and remaining_hunks say what was cut.

    Returns path, file_path, revision, identical, identical_raw, normalized, truncated, hunks or
    unified_diff, warnings, lossy_elements (when non-empty).

    Errors: INVALID_INPUT (bad path, file_path missing or not a regular file, bad format/max_hunks,
    no file heading matching section_key), SECTION_NOT_FOUND.
    Full contract: docs/tool-contracts.md#diff_body_vs_file.
    """
    try:
        return execute_diff_body_vs_file(path, file_path, normalize, format, section_key, max_hunks)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations={"readOnlyHint": True, "idempotentHint": True})
def diff_table_vs_markdown(
    path: str,
    table_id: int | None = None,
    markdown: str | None = None,
    markdown_path: str | None = None,
    md_heading: str | None = None,
    md_header_row: str | None = None,
    section_key: str | None = None,
    source: str = "auto",
    strip_emphasis: bool = True,
) -> dict[str, Any]:
    """Report the cells that differ between a Word table and a markdown pipe table. Read-only.

    Give exactly one markdown input (`markdown`, or `markdown_path` under the allowed roots) and
    exactly one Word selector: `table_id` (single mode; the markdown table is chosen by `md_heading`
    (first table under it) and/or `md_header_row` (a pipe row like "| Name | Role |"); with neither,
    the markdown must hold exactly one table; zero or several candidates is INVALID_INPUT, never
    guessed) or `section_key` (multi mode: every Word table in that section is paired with the
    markdown tables under `md_heading`, default the section heading, by header-row cells; repeated
    headers, a missing heading or no tables is refused).
    Cells compare as visible text, folded (quotes, dashes, NBSP, NFC) and whitespace-collapsed;
    bold/italic are ignored unless strip_emphasis=false. A table with merged cells, a nested table,
    or a cell holding a field, drawing or text box is NOT compared: it is listed under `unsupported`
    and identical is false. source "auto"|"file"|"live" as read_document.

    Single mode returns path, table_id, md_table_index (0-based), identical, differences [{row, col,
    docx_text, md_text}] (1-based, header = row 1), dimensions (if they differ), unsupported,
    warnings. Multi mode returns path, identical, tables [{table_id, md_table_index, header,
    differences, dimensions?}], unmatched_docx, unmatched_md, unsupported.

    Errors: INVALID_INPUT, TABLE_NOT_FOUND, SECTION_NOT_FOUND, LIVE_UNAVAILABLE /
    LIVE_CAPABILITY_MISSING / LIVE_DISCONNECTED / LIVE_OP_FAILED (source="live").
    Full contract: docs/tool-contracts.md#diff_table_vs_markdown.
    """
    try:
        return table_diff.diff_table_vs_markdown_impl(
            path, table_id, markdown, markdown_path, md_heading, md_header_row, section_key, source, strip_emphasis
        )
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
                    _doctor_print("NOTE", f"{len(leftover)} leftover \"jennystack-render-*\" document window(s) open in Word (read-only count; renders now close their own window by default — leftovers predate that change or had close_error set). Safe to close by hand.")
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


@mcp.tool(annotations=_READ_ONLY)
def list_textboxes(path: str) -> dict[str, Any]:
    """List live text-bearing shapes with epoch handles, coverage and revisions."""
    from .live.scopes_live import read_scope

    try:
        return read_scope(path)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def read_textbox(path: str, textbox_id: str) -> dict[str, Any]:
    """Read one live text box. Host failures are distinct from verified empty text."""
    from .live.scopes_live import read_scope

    try:
        return read_scope(path, textbox_id)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def delete_paragraph(path: str, anchor_text: str, revision_before: str,
                     track_changes: bool = False, expected_length: int | None = None,
                     find_sha256: str | None = None) -> dict[str, Any]:
    """Delete exactly one whole live paragraph (body or table cell). Refuses anchors, comments,
    revisions, section boundaries, the final body paragraph, a nested table's paragraph and a cell's
    LAST paragraph (use merge_paragraph_with_next to join). `anchor_text` must equal exactly one
    paragraph in the whole body: its whole text, or "<prefix>…<suffix>" covering it with
    expected_length and/or find_sha256 (as replace_text's anchored find). revision_before and
    track_changes as other live writes.
    """
    from .live.paragraphs_live import delete_paragraph as execute

    try:
        return execute(path, anchor_text, revision_before, track_changes, expected_length, find_sha256)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def merge_paragraph_with_next(
    path: str,
    table_id: int,
    row_index: int,
    cell_index: int,
    paragraph_index: int,
    expected_text: list[str] | None = None,
    separator: str = "",
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Merge two adjacent paragraphs of ONE table cell into one. All indices are 1-based;
    paragraph_index is the FIRST of the pair and must not be the cell's last (STRUCTURAL_BOUNDARY).

    The first paragraph's properties win; runs and hyperlinks are kept. LIVE: Word keeps the SECOND
    paragraph's properties, so a live merge of paragraphs whose properties differ is refused before
    anything is written (STRUCTURAL_BOUNDARY): make them match, or use write_mode="file".
    expected_text [first, second] guards both texts (CELL_TEXT_MISMATCH, nothing written).
    separator (default "") goes between the texts. Refuses (nothing written) comments, tracked
    changes, anchored shapes, fields, bookmarks, footnote references and section breaks in either
    paragraph, and tables with merged or nested cells. track_changes=true is refused in both modes
    (INVALID_INPUT; how Word records it is unproven).
    write_mode "auto"|"file"|"live"; live needs pane capability "cell_paragraph_merge_v1"
    (LIVE_CAPABILITY_MISSING: reload the pane). Both modes verify texts, count, the merged
    paragraph's properties, every hyperlink and every other cell. VERIFICATION_FAILED means the
    write may have landed: re-read the table before retrying. A document that moved after a live
    write adds a warnings entry. Audited. Returns the eight evidence keys.
    Full contract: docs/tool-contracts.md#merge_paragraph_with_next.
    """
    from .paragraph_merge import execute

    try:
        return execute(
            path, table_id, row_index, cell_index, paragraph_index, expected_text, separator,
            revision_before, track_changes, write_mode, allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def clear_numbering(
    path: str,
    table_id: int | None = None,
    paragraph_text: str | None = None,
    revision_before: str | None = None,
    write_mode: str = "live",
) -> dict[str, Any]:
    """Remove list numbering/bullets from existing live content. LIVE ONLY: file mode refuses
    (LIVE_UNAVAILABLE).

    Pass exactly one of table_id (every list-item paragraph in that table) or paragraph_text (the one
    paragraph whose WHOLE text equals it; MATCH_COUNT_MISMATCH unless exactly one). Text is never
    changed. Refused while tracked changes touch the scope (TRACKED_CHANGES_PRESENT) or a section
    lock covers it. The server fails VERIFICATION_FAILED unless no list item remains. Evidence adds
    `unlisted` (paragraphs detached; 0 is a verified no-op).
    """
    from .live.paragraphs_live import clear_numbering as execute

    try:
        return execute(path, table_id, paragraph_text, revision_before, write_mode)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_DESTRUCTIVE)
def delete_table(
    path: str,
    table_id: int,
    expected_text: str,
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "live",
) -> dict[str, Any]:
    """Delete one live top-level table. LIVE ONLY (file mode refuses with LIVE_UNAVAILABLE).

    table_id is the list_tables / get_table number. expected_text must equal the table's current text
    as get_table reports it (cells of a row joined by a tab, rows by a newline; whitespace- and
    quote-insensitive); a mismatch refuses with CELL_TEXT_MISMATCH and nothing is written. Refused
    for tables that anchor a shape, overlap comments, or contain tracked changes. The pane verifies
    the table count dropped by exactly one and no outside paragraph changed. With track_changes=true
    the table is only marked deleted until the change is accepted.
    """
    from .live.paragraphs_live import delete_table as execute

    try:
        return execute(path, table_id, expected_text, revision_before, track_changes, write_mode)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def insert_paragraphs(
    path: str,
    anchor: str | dict[str, Any],
    paragraphs: list[dict[str, Any]],
    expected_matches: int,
    position: str = "after",
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    """Insert new paragraphs (and headings) before or after the body paragraph containing `anchor`.

    `paragraphs`: 1-50 objects {"text" (no newlines), "style" (paragraph style name, e.g. "Heading 2";
    omitted = Normal, never the anchor's), "color" (six hex; omitted = unset, live Word may carry the
    anchor's colour, so pass one to pin it)}. Each inserted text, style and colour is read back and
    listed under `inserted`.
    `anchor` text must match exactly once (expected_matches must be 1) in a top-level body paragraph,
    not a table cell. position "after" (default) | "before" (live section locks allow only "after").
    Structural anchors: {"after_table_id": N} or {"empty_paragraph_after_table_id": N} (fills the
    empty paragraph right after table N; STRUCTURAL_BOUNDARY if none or it holds text, a field, bookmark,
    tab, break, hyperlink, drawing, comment anchor or tracked change); both need position="after". Live mode needs "paragraph_positions_v1".
    track_changes=True: live records as the signed-in Word user; file mode wraps new runs in w:ins.
    write_mode "auto"|"file"|"live", revision_before, allow_concurrent_editor as replace_text. A live
    write that fails read-back is deleted again when untracked (the message says whether that
    worked); a tracked one is left for Word's Review pane.

    Errors: INVALID_INPUT, ZERO_MATCH, MATCH_COUNT_MISMATCH, STRUCTURAL_BOUNDARY, STYLE_NOT_FOUND,
    UNSUPPORTED_STYLE_TYPE, OUTSIDE_LOCKED_SECTION, VERIFICATION_FAILED, LIVE_UNAVAILABLE,
    LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_STALE,
    LIVE_CAPABILITY_MISSING, DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT, EXTERNAL_EDITOR_ACTIVE,
    OPC_INVALID.
    Full contract: docs/tool-contracts.md#insert_paragraphs.
    """
    try:
        return paragraph_insert.execute_insert_paragraphs(
            path,
            anchor,
            paragraphs,
            expected_matches,
            position=position,
            revision_before=revision_before,
            track_changes=track_changes,
            write_mode=write_mode,
            allow_concurrent_editor=allow_concurrent_editor,
        )
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations=_READ_ONLY)
def list_shapes(path: str) -> dict[str, Any]:
    """List live body anchored shapes with paragraph anchors and OOXML geometry.

    Includes table and textbox paragraphs. Excludes headers/footers and inline
    pictures. DrawingML geometry uses EMUs; legacy VML retains its style units.
    """
    from .live import comments_live, write_mode
    try:
        session = write_mode.live_session_for(path)
        write_mode.require_capability(session, "shape_guard", feature_description="shape inventory")
        return comments_live._request(session, "shapes_list")
    except VerifyError as exc:
        _raise_tool_error(exc)


def main() -> None:
    """Entry point for the `verified-docx-mcp` command.

    `verified-docx-mcp doctor`  — run the local render-path diagnostic.
    `verified-docx-mcp`         — start the stdio MCP server.
    """
    if len(sys.argv) > 1 and sys.argv[1] == "doctor":
        sys.exit(doctor())

    from .cancellation import install_cancel_guard
    from .shutdown import install_shutdown_handlers

    install_cancel_guard()
    install_shutdown_handlers(live_bridge.stop)
    mcp.run()


@mcp.tool()
def live_changes_prepare(path: str, operations: list[dict[str, Any]], track_changes: bool = False) -> dict[str, Any]:
    """Prepare a scoped approved run without modifying Word. Returns a preview, run_id and
    manifest_sha256.

    operations: [{op, section, args}]; section is a unique heading slug; ops are replace_cell_markdown,
    replace_table_row, replace_text, delete_paragraph and a final live_save. The preview shows each
    operation's find/replace/markdown/cells/anchor_text (cut to 200 characters), table address and
    expected_matches, so the lead can approve from it; it is DERIVED from the manifest, and approving
    the digest binds the exact operations, not the clipped preview.
    replace_text args may add replacement_segments (as replace_text; concatenated text must equal
    replace) and may name the target as an anchored find (expected_length/find_sha256) or {para_ref,
    expected_sha256} where pN counts the section's paragraphs (p0 = its heading); prepare resolves
    it against the section as earlier operations leave it and stores the resolved full find (over
    230 characters needs anchored_find_v1). replace_cell_markdown args may carry markdown_file
    instead of markdown: the manifest stores the CONTENT, so the digest binds the text, not the
    path. Needs approved_changes_v1; explicit segments also need formatted_replacements_v1; partial
    formats (partial_font_segments_v1) resolve against the first replaced run, the manifest binds
    that font_baseline, and execute refuses FORMAT_DRIFT if the run was reformatted since.
    Full contract: docs/tool-contracts.md#live_changes_prepare.
    """
    from .live.approved_changes import prepare
    try:
        return prepare(path, operations, track_changes)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def live_changes_execute(
    path: str, run_id: str, manifest_sha256: str, approved_by: str, verbose: bool = False
) -> dict[str, Any]:
    """Execute one immutable, lead-approved change set. Code handles all writes,
    internal verification, bounded pre-write retries and the durable progress
    journal. Unrelated section edits do not invalidate the run. Returns counts
    and exceptions only; full evidence stays local. Unknown outcomes never retry.
    The report does not echo your run_id or manifest_sha256; `remaining` lists
    operation indices and `remaining_ranges` is the same as a string ("7-36");
    an issue's `section_diff` is cut to 600 characters (`journal_path` has the
    rest). `verbose=true` adds `performance`.
    """
    from .live.approved_changes import execute
    try:
        return execute(path, run_id, manifest_sha256, approved_by, verbose)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool()
def live_changes_resume(path: str, run_id: str, manifest_sha256: str, verbose: bool = False) -> dict[str, Any]:
    """Resume independent remaining operations of an already approved run. Verifies
    document, digest and protected targets; skips verified operations and defers
    uncertain targets without replaying them. Does not broaden approval.
    """
    from .live.approved_changes import execute
    try:
        return execute(path, run_id, manifest_sha256, None, verbose)
    except VerifyError as exc:
        _raise_tool_error(exc)

@mcp.tool()
def live_changes_reconcile(path: str, run_id: str, manifest_sha256: str,
                           disposition: str, decision_ref: str, operation_index: int | None = None,
                           verbose: bool = False) -> dict[str, Any]:
    """Record a lead's explicit applied/not-applied decision on an uncertain run.
    Exact live section readback must support the decision. Does not write Word.
    """
    from .live.approved_changes import reconcile
    try:
        return reconcile(path, run_id, manifest_sha256, disposition, decision_ref, operation_index, verbose)
    except VerifyError as exc:
        _raise_tool_error(exc)



@mcp.tool()
def insert_table_row(path: str, table_id: int, cells: list[str],
                     after_row_index: int | None = None, before_row_index: int | None = None,
                     template_row_index: int | None = None, revision_before: str | None = None,
                     track_changes: bool = False, write_mode: str = "auto",
                     allow_concurrent_editor: bool = False) -> dict[str, Any]:
    """Insert a row of plain-text cells, preserving template formatting. Indices are 1-based.

    Supply exactly one of after_row_index/before_row_index. The adjacent row is
    the default template. Merged/nested tables are refused. Re-read positional IDs
    after structure changes. Writes are revision-guarded, verified and audited.
    """
    try:
        return table_structure.execute_row(path, table_id, cells=cells,
            after_row_index=after_row_index, before_row_index=before_row_index,
            template_row_index=template_row_index, revision_before=revision_before,
            track_changes=track_changes, write_mode=write_mode,
            allow_concurrent_editor=allow_concurrent_editor)
    except VerifyError as exc:
        _raise_tool_error(exc)


@mcp.tool(annotations={"destructiveHint": True})
def delete_table_row(path: str, table_id: int, row_index: int, expected_text: str,
                     revision_before: str | None = None, track_changes: bool = False,
                     write_mode: str = "auto", allow_concurrent_editor: bool = False) -> dict[str, Any]:
    """Delete one row after matching its tab-joined cell text. Indices are 1-based.

    Refuses the final row, mismatched text, protected content, and merged/nested
    tables. Re-read row indices after deletion. Verified and audit-logged.
    """
    try:
        return table_structure.execute_row(path, table_id, row_index=row_index,
            expected_text=expected_text, revision_before=revision_before,
            track_changes=track_changes, write_mode=write_mode,
            allow_concurrent_editor=allow_concurrent_editor)
    except VerifyError as exc:
        _raise_tool_error(exc)

if __name__ == "__main__":
    main()
