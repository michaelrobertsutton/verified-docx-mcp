# Adapted from GoogleDocs-MCP src/verified_googledocs_mcp/verify.py's audit
# writer section (commit 374bcf8, ~lines 867-940: _AUDIT_REDACTED_KEYS,
# _audit_excerpts_enabled, _state_dir, append_audit). Lifted per issue #28
# WP-02 ("Lift, do not rewrite" — "verify.py (... audit writer,
# _AUDIT_REDACTED_KEYS :867)"). Logic is unchanged; only the package name
# (verified-googledocs-mcp -> verified-docx-mcp) and its env-var prefix
# (VERIFIED_GOOGLEDOCS_MCP_ -> VERIFIED_DOCX_MCP_) differ, and the record
# shape's doc/tab identity fields are replaced with this backend's own
# addressing (a filesystem path has no tab dimension).
"""Audit writer: appends a JSONL record for every mutating call.

Mirrors verified-googledocs-mcp's audit log exactly: one line per call,
owner-only permissions on both the state directory and the file, and the
same before/after excerpt-redaction toggle (env var, not a tool parameter —
see _audit_excerpts_enabled).
"""

from __future__ import annotations

import json
import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

_AUDIT_REDACTED_KEYS = frozenset({"before", "after", "runs_before", "runs_after"})

# Operator control surface for audit-excerpt redaction. No tool exposes the
# toggle, so this environment variable is the only way to reach the
# redaction path end-to-end (mirrors the Google server's own rationale).
_AUDIT_EXCERPTS_ENV = "VERIFIED_DOCX_MCP_AUDIT_EXCERPTS"
_AUDIT_EXCERPTS_FALSEY = frozenset({"0", "false", "no", "off", ""})


def _audit_excerpts_enabled(default: bool) -> bool:
    """Resolve whether before/after excerpts are kept in the audit log.

    ``VERIFIED_DOCX_MCP_AUDIT_EXCERPTS`` is authoritative when set: falsey
    values (0/false/no/off, case-insensitive, plus the empty string) redact
    the excerpts, any other value keeps them. When unset, the caller's
    ``default`` applies.
    """
    raw = os.environ.get(_AUDIT_EXCERPTS_ENV)
    if raw is None:
        return default
    return raw.strip().lower() not in _AUDIT_EXCERPTS_FALSEY


def _state_dir() -> Path:
    """Return the XDG state directory for this package."""
    xdg = os.environ.get("XDG_STATE_HOME", "")
    if xdg:
        base = Path(xdg)
    else:
        base = Path.home() / ".local" / "state"
    return base / "verified-docx-mcp"


def append_audit(
    *,
    path: str,
    tool: str,
    evidence: dict[str, Any],
    audit_excerpts: bool = True,
) -> tuple[bool, str]:
    """Append a mutation record to the audit log.

    Best-effort: never raises. Returns (logged: bool, reason: str); reason
    is empty on success. ``path`` is the resolved .docx path the call acted
    on — the docx-backend analogue of the Google server's (doc, tab) pair
    (a local file has no tab dimension).

    Excerpt redaction: when excerpts are disabled the ``before``/``after``
    (and ``runs_before``/``runs_after``) fields are replaced with
    ``"[redacted; N chars]"`` while all other evidence keys are preserved.
    Set ``VERIFIED_DOCX_MCP_AUDIT_EXCERPTS`` to a falsey value (0/false/no/
    off) to redact; it overrides the ``audit_excerpts`` argument when set.
    """
    try:
        state_dir = _state_dir()
        # The audit log records document content excerpts; keep it owner-only.
        state_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
        audit_path = state_dir / "audit.jsonl"

        payload = evidence.copy()
        if not _audit_excerpts_enabled(audit_excerpts):
            for key in _AUDIT_REDACTED_KEYS:
                if key in payload:
                    payload[key] = f"[redacted; {len(str(payload[key]))} chars]"

        record = {
            "timestamp": datetime.now(UTC).isoformat(),
            "path": path,
            "tool": tool,
            "evidence": payload,
        }

        # Open append-only, creating with owner-only perms; chmod each time so
        # a pre-existing loose-permission file is also tightened. Volume is low.
        fd = os.open(audit_path, os.O_WRONLY | os.O_CREAT | os.O_APPEND, 0o600)
        with os.fdopen(fd, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, ensure_ascii=False) + "\n")
        os.chmod(audit_path, 0o600)

        return True, ""

    except Exception as exc:  # noqa: BLE001
        return False, str(exc)
