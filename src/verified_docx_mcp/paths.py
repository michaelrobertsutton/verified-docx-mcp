"""Path safety: the $DOCS containment rule, the file-root allowlist, and the
read-path package snapshot.

Three lifted/adapted pieces, per issue #28 WP-02 ("paths.py = the lifted
allowlist plus the $DOCS containment rule from WP-01"):

1. ``DocxPathError`` / ``resolve_docx_pointer`` — mirrored VERBATIM (per
   that section's own instruction: "Any skill or tool that resolves a
   docx: pointer mirrors this exact snippet inline rather than restating
   the rule loosely") from JennyStack core/document-backend-protocol.md
   §3, "Hard guardrail — reject an escape, not a naive prefix match". That
   doc is the canonical source; if this ever disagrees with it, the core
   doc wins (JennyStack CLAUDE.md's single-source-of-truth rule) and this
   copy needs updating to match.

2. ``_allowed_file_roots`` / ``_is_denylisted_sensitive_path`` — adapted
   from GoogleDocs-MCP src/verified_googledocs_mcp/markdown_mutations.py
   (commit 374bcf8, ~lines 1151-1200), lifted per issue #28 WP-02 ("the
   path allowlist in markdown_mutations.py (_allowed_file_roots,
   _is_denylisted_sensitive_path)"). Same logic; only the env var prefix
   changes (VERIFIED_GOOGLEDOCS_MCP_ -> VERIFIED_DOCX_MCP_). This is a
   SEPARATE guardrail from resolve_docx_pointer above: that one confines a
   docx: pointer's relative path to $DOCS; this one is the server
   process's own general file-access allowlist (defaults to the user's
   home directory), the same shape the Google server uses for
   diff_tab_vs_file. Both apply — $DOCS containment first (when a docx:
   pointer is in play), then this allowlist as the server-wide floor.

3. ``snapshot_docx_package`` — new for this WP (no Google-side analogue;
   .docx is a local zip package, Google Docs has no such concept). Backs
   the "reads never refuse" rule (core/document-backend-protocol.md §4):
   when a Word owner file is present, a read snapshots the package to a
   temp dir first rather than refusing. Snapshot is NOT atomic (Codex
   defect noted in the WP): the copy can race a concurrent Word
   autosave/flush, so it is validated (zipfile.testzip() + a parse of
   word/document.xml) and retried up to 3 times at 0.5s backoff before
   raising SNAPSHOT_FAILED.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree
from xml.etree.ElementTree import ParseError

from .errors import ErrorCode, _make_error

# ---------------------------------------------------------------------------
# 1. $DOCS containment (mirrored verbatim from
#    core/document-backend-protocol.md §3 — see that doc for the "why, so
#    nobody rebuilds it" rationale; not restated here to avoid drifting
#    from the canonical copy).
# ---------------------------------------------------------------------------


class DocxPathError(ValueError):
    """A docx: pointer's path cannot be safely resolved under $DOCS."""


def resolve_docx_pointer(relative_path: str, docs_root: str) -> Path:
    if not relative_path or relative_path in (".", ""):
        raise DocxPathError("DOCX_PATH_ESCAPE: empty or '.' resolves to $DOCS itself")

    if Path(relative_path).is_absolute():
        raise DocxPathError("DOCX_PATH_ESCAPE: relative_path must not be absolute")

    if "\x00" in relative_path:
        # Without this, a NUL byte reaches resolve() below and raises a raw
        # ValueError ("embedded null character in path") instead of
        # DocxPathError — fine for a caller that catches ValueError (the
        # parent class), but a crash for one that catches DocxPathError only,
        # which this snippet's own existence invites a caller to do.
        raise DocxPathError("DOCX_PATH_ESCAPE: relative_path contains a NUL byte")

    # Reject traversal before resolution — belt, not just suspenders.
    if ".." in Path(relative_path).parts:
        raise DocxPathError("DOCX_PATH_ESCAPE: relative path contains '..'")

    root = Path(docs_root).resolve(strict=False)
    if not root.is_dir():
        # resolve(strict=False) alone would let a typo'd root pass containment
        # silently — a nonexistent root must fail loudly here, not let a write
        # land in a phantom tree.
        raise DocxPathError(f"DOCX_ROOT_NOT_FOUND: $DOCS does not exist: {root}")

    # NOTE: pathlib's `/` operator silently DISCARDS the left side when the right
    # side is absolute — root / "/etc/passwd" resolves to "/etc/passwd", not
    # "<root>/etc/passwd". The is_absolute() rejection above exists so a caller
    # gets a clear DOCX_PATH_ESCAPE instead of a confusing pass-through; even
    # without it, "/etc/passwd" would still be caught below by is_relative_to()
    # (the containment check), not by the ".." check above — it contains no "..".
    target = (root / relative_path).resolve(strict=False)

    # is_relative_to, not a string prefix test — a prefix match treats
    # "/docs-evil" as inside "/docs".
    if not target.is_relative_to(root):
        raise DocxPathError("DOCX_PATH_ESCAPE: target resolves outside $DOCS")

    return target


# ---------------------------------------------------------------------------
# 2. Server-wide file-access allowlist (adapted from markdown_mutations.py's
#    _allowed_file_roots / _is_denylisted_sensitive_path).
# ---------------------------------------------------------------------------

_ALLOWED_FILE_ROOTS_ENV = "VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS"

# The Claude Code per-user scratch root (issue #101): a proposal build staged
# in the harness-mandated scratchpad (a subdirectory of this) could not be
# read back or rendered without copying it into a pursuit repo first, which
# is exactly the round trip the verify-before-copy workflow
# (core/document-backend-protocol.md) exists to avoid. Folded into the
# DEFAULT allowed roots only (below); an explicit
# VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS is used verbatim and is not widened.
_CLAUDE_SCRATCH_ROOT_TEMPLATE = "/private/tmp/claude-{uid}"


def _claude_code_scratch_root() -> Path | None:
    """The Claude Code per-user scratch root, or None when it does not exist."""
    candidate = Path(_CLAUDE_SCRATCH_ROOT_TEMPLATE.format(uid=os.getuid()))
    return candidate if candidate.is_dir() else None


# Well-known credential/secrets locations, relative to the user's home
# directory. Denylisted unconditionally — regardless of the configured
# allowed roots — because the risk this guards against is a document's own
# content tricking an agent into reading credentials (prompt injection), not
# an operator deliberately misconfiguring VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS.
# Widening the default allowed root to the home directory (below) makes this
# denylist load-bearing: it is what keeps that default from exposing
# ~/.ssh, ~/.aws, etc. to a doc-driven read.
_SENSITIVE_HOME_RELATIVE_DENYLIST = (
    ".ssh",
    ".aws",
    ".gnupg",
    ".netrc",
    ".git-credentials",
    ".config/gh",
    ".docker/config.json",
    ".npmrc",
)


def _allowed_file_roots() -> list[Path]:
    """Directories this server may read/write a .docx file under.

    Defaults to the user's home directory (covers iCloud Drive / OneDrive
    sync roots under it without extra configuration), not the server
    process's working directory, PLUS the Claude Code scratch root
    (/private/tmp/claude-<uid>) when that directory exists (issue #101) — a
    scratch build can be read back and rendered without first being copied
    into a synced repo. Home still excludes system paths and other users'
    home directories, and _is_denylisted_sensitive_path below unconditionally
    blocks well-known credential locations under it. Set
    VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS explicitly to use that list verbatim
    instead — an explicit setting is never silently widened with the scratch
    root or narrowed to it.
    """
    raw = os.environ.get(_ALLOWED_FILE_ROOTS_ENV)
    if raw is not None:
        root_values = [p for p in raw.split(os.pathsep) if p]
    else:
        root_values = [str(Path.home())]
        scratch_root = _claude_code_scratch_root()
        if scratch_root is not None:
            root_values.append(str(scratch_root))
    return [Path(value).expanduser().resolve(strict=False) for value in root_values]


def _is_denylisted_sensitive_path(resolved: Path) -> bool:
    """True if *resolved* is (or is inside) a well-known credential location
    under the user's home directory — see _SENSITIVE_HOME_RELATIVE_DENYLIST.
    """
    home = Path.home().resolve(strict=False)
    for entry in _SENSITIVE_HOME_RELATIVE_DENYLIST:
        denied = (home / entry).resolve(strict=False)
        if resolved == denied or resolved.is_relative_to(denied):
            return True
    return False


def resolve_allowed_docx_path(
    file_path: str, *, must_exist: bool = True, escape_code: ErrorCode = ErrorCode.INVALID_INPUT
) -> Path:
    """Resolve *file_path* and enforce the server-wide allowlist + denylist.

    Raises INVALID_INPUT (via _make_error) on a denylisted or
    outside-allowlist path, matching the Google server's diff_tab_vs_file
    error shape. Does not enforce $DOCS containment — that is
    resolve_docx_pointer's job for an actual docx: pointer; this is the
    floor every resolved path must clear regardless of how it was named.
    ``escape_code`` is the code for a denylisted/outside-allowlist path
    (``resolve_input_text_file`` passes DOCX_PATH_ESCAPE).
    """
    requested = Path(file_path).expanduser()
    if must_exist:
        try:
            resolved = requested.resolve(strict=True)
        except FileNotFoundError as exc:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"File not found: {file_path!r}",
                {"path": file_path},
            ) from exc
    else:
        resolved = requested.resolve(strict=False)

    if _is_denylisted_sensitive_path(resolved):
        raise _make_error(
            escape_code,
            (
                "File path resolves to a well-known credential/secrets location "
                "and is never allowed, regardless of "
                f"{_ALLOWED_FILE_ROOTS_ENV}."
            ),
            {"path": file_path, "resolved_path": str(resolved)},
        )

    allowed_roots = _allowed_file_roots()
    if not any(resolved.is_relative_to(root) for root in allowed_roots):
        raise _make_error(
            escape_code,
            (
                "File path is outside the allowed roots for this server. "
                "By default this is the user's home directory plus the "
                "Claude Code scratch root (/private/tmp/claude-<uid>) when "
                "that directory exists — set "
                f"{_ALLOWED_FILE_ROOTS_ENV} on the server process to a "
                f"{os.pathsep!r}-separated list of directories to widen it, "
                "then restart the server."
            ),
            {
                "path": file_path,
                "resolved_path": str(resolved),
                "allowed_roots": [str(root) for root in allowed_roots],
                "env_var": _ALLOWED_FILE_ROOTS_ENV,
            },
        )

    return resolved


def resolve_output_path(output_path: str, *, param: str = "output_path") -> Path:
    """Validate an output file path (export_pdf, read_document(to_file=...)):
    allowlisted, not denylisted, parent directory must already exist (no
    implicit mkdir), and an existing target must be a regular file."""
    if not output_path.strip():
        raise _make_error(ErrorCode.INVALID_INPUT, f"{param} must not be empty")

    resolved = resolve_allowed_docx_path(output_path, must_exist=False)

    parent = resolved.parent
    if not parent.is_dir():
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"Parent directory does not exist: {parent}",
            {param: output_path, "resolved_path": str(resolved), "parent": str(parent)},
        )

    if resolved.exists() and not resolved.is_file():
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"Output path exists and is not a regular file: {output_path!r}",
            {param: output_path, "resolved_path": str(resolved)},
        )

    return resolved


_TEXT_OUTPUT_SUFFIXES = frozenset({".md", ".txt", ".json"})


def resolve_text_output_path(output_path: str, *, param: str = "to_file", overwrite: bool = False) -> Path:
    """Like ``resolve_output_path`` for a tool that writes document TEXT to a file the caller
    names (``read_document(to_file=...)``): only ``.md``/``.txt``/``.json``, and never an
    existing file unless ``overwrite`` is true. A read tool must not be able to replace a
    .docx (including the one it is reading), a dotfile or a source file inside the allowed
    roots just because the caller passed that path."""
    resolved = resolve_output_path(output_path, param=param)
    if resolved.suffix.lower() not in _TEXT_OUTPUT_SUFFIXES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"{param} must end in .md, .txt or .json (got {resolved.suffix or 'no extension'!r}).",
            {param: output_path},
        )
    if resolved.exists() and not overwrite:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"{param} already exists: {resolved.name!r}. Choose a new path or pass overwrite=true.",
            {param: output_path},
        )
    return resolved


_INPUT_TEXT_SUFFIXES = frozenset({".md", ".txt"})
MAX_INPUT_TEXT_BYTES = 2 * 1024 * 1024


def resolve_input_text_file(path: str, *, param: str = "markdown_file") -> Path:
    """Resolve a caller-named markdown/text INPUT file (``markdown_file=``) without reading it.

    Same roots and credential denylist as ``resolve_allowed_docx_path`` (symlinks are
    resolved first, so a link pointing outside the roots is refused): outside them is
    DOCX_PATH_ESCAPE. Anything else wrong (missing, not a regular file, suffix other than
    ``.md``/``.txt``, over 2 MB) is INVALID_INPUT.
    """
    if not isinstance(path, str) or not path.strip():
        raise _make_error(ErrorCode.INVALID_INPUT, f"{param} must not be empty")
    resolved = resolve_allowed_docx_path(path, escape_code=ErrorCode.DOCX_PATH_ESCAPE)
    details = {param: path, "resolved_path": str(resolved)}
    if not resolved.is_file():
        raise _make_error(ErrorCode.INVALID_INPUT, f"{param} is not a regular file: {path!r}", details)
    if resolved.suffix.lower() not in _INPUT_TEXT_SUFFIXES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"{param} must end in .md or .txt (got {resolved.suffix or 'no extension'!r}).",
            details,
        )
    return resolved


def read_input_text_file(path: str, *, param: str = "markdown_file") -> tuple[Path, str]:
    """``resolve_input_text_file`` plus ONE read: at most 2 MB, strict UTF-8 (else
    INVALID_INPUT); a leading BOM is dropped. Returns ``(resolved_path, text)``."""
    resolved = resolve_input_text_file(path, param=param)
    details = {param: path, "resolved_path": str(resolved)}
    try:
        with open(resolved, "rb") as handle:
            data = handle.read(MAX_INPUT_TEXT_BYTES + 1)
    except OSError as exc:
        raise _make_error(ErrorCode.INVALID_INPUT, f"Could not read {param}: {exc}", details) from exc
    if len(data) > MAX_INPUT_TEXT_BYTES:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"{param} is larger than {MAX_INPUT_TEXT_BYTES // (1024 * 1024)} MB.",
            details,
        )
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise _make_error(ErrorCode.INVALID_INPUT, f"{param} is not valid UTF-8: {exc}", details) from exc
    return resolved, text.removeprefix("\ufeff")


def write_text_atomic(target: Path, text: str) -> bytes:
    """Write *text* (UTF-8) to *target* via a temp file in the same directory
    and ``os.replace``; returns the bytes written so the caller can hash
    exactly what landed on disk."""
    data = text.encode("utf-8")
    fd, tmp_name = tempfile.mkstemp(prefix=f".{target.name}.", dir=str(target.parent))
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
        os.replace(tmp_name, target)
    except BaseException:
        try:
            os.unlink(tmp_name)
        except FileNotFoundError:
            pass
        raise
    return data


# ---------------------------------------------------------------------------
# 3. Read-path package snapshot, with validated retry.
# ---------------------------------------------------------------------------

_SNAPSHOT_MAX_ATTEMPTS = 3
_SNAPSHOT_RETRY_BACKOFF_SECONDS = 0.5


def _validate_docx_snapshot(snapshot_path: Path) -> None:
    """Raise ValueError (any subclass) if *snapshot_path* is not a readable,
    well-formed .docx package: zipfile.testzip() must report no bad
    entries, and word/document.xml must parse as XML."""
    with zipfile.ZipFile(snapshot_path) as zf:
        bad_entry = zf.testzip()
        if bad_entry is not None:
            raise ValueError(f"corrupt zip entry: {bad_entry}")
        try:
            data = zf.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("missing word/document.xml") from exc
        try:
            ElementTree.fromstring(data)
        except ParseError as exc:
            raise ValueError(f"word/document.xml did not parse: {exc}") from exc


def snapshot_docx_package(
    source_path: Path,
    *,
    max_attempts: int = _SNAPSHOT_MAX_ATTEMPTS,
    backoff_seconds: float = _SNAPSHOT_RETRY_BACKOFF_SECONDS,
    sleep=time.sleep,
) -> Path:
    """Copy *source_path* to a validated temp file and return its path.

    Used on the read path when a Word owner file is present (DOCX_LOCKED
    gates writes only — see core/document-backend-protocol.md §4): rather
    than refuse the read, snapshot the package first so a concurrent Word
    autosave cannot hand back a half-written zip.

    The copy is NOT atomic (Codex defect noted in issue #28 WP-02): it can
    race Word's own flush. Each attempt is therefore validated
    (_validate_docx_snapshot: zipfile.testzip() + a word/document.xml
    parse) before being trusted; a failed attempt is retried up to
    ``max_attempts`` times with ``backoff_seconds`` between tries. The
    caller owns cleanup of the returned temp file.

    Raises the source's own OSError subclasses if *source_path* cannot be
    read at all (distinct from SNAPSHOT_FAILED, which means "copies
    happened but never validated clean").
    """
    last_error: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        fd, tmp_name = tempfile.mkstemp(
            prefix="verified-docx-mcp-snapshot-", suffix=".docx"
        )
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            shutil.copyfile(source_path, tmp_path)
            _validate_docx_snapshot(tmp_path)
            return tmp_path
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            last_error = exc
            tmp_path.unlink(missing_ok=True)
            if attempt < max_attempts:
                sleep(backoff_seconds)

    raise _make_error(
        ErrorCode.SNAPSHOT_FAILED,
        (
            f"Could not produce a validated read snapshot of {source_path} after "
            f"{max_attempts} attempts."
        ),
        {"path": str(source_path), "attempts": max_attempts, "last_error": str(last_error)},
    )
