# New for issue #28 WP-07b-a. No GoogleDocs-MCP analogue: that server's
# comment/suggestion author comes from the authenticated Google account
# (the Drive/Docs API's own identity), which this server has no equivalent
# of (a local .docx has no signed-in account) -- author identity here is
# resolved from JennyStack's own config file, with a live-machine fallback.
"""Resolve the author name stamped on every w:ins/w:del/w:rPrChange/
w:comment this server creates: ``author_name`` from
``~/.jennystack/config.json`` (the same file/key WP-08's comments use),
falling back to the current macOS user's full name when the key is absent
-- as it is on a machine that has only set ``docs_root`` so far (issue #28
plan WP-07b-a's own note: "author_name is currently absent... your fallback
path is the live one").

Every step is independently overridable/mockable for tests:
``VERIFIED_DOCX_MCP_JENNYSTACK_CONFIG`` points at an alternate config file
(default ``~/.jennystack/config.json``); ``_macos_full_name`` is a separate,
patchable function rather than inlined, so a test can force the fallback
path without depending on the real machine's actual account name.
"""

from __future__ import annotations

import json
import os
import pwd
import subprocess
from pathlib import Path

_CONFIG_PATH_ENV = "VERIFIED_DOCX_MCP_JENNYSTACK_CONFIG"
_DEFAULT_CONFIG_PATH = Path.home() / ".jennystack" / "config.json"
_FALLBACK_AUTHOR = "Unknown Author"


def _config_path() -> Path:
    override = os.environ.get(_CONFIG_PATH_ENV)
    return Path(override).expanduser() if override else _DEFAULT_CONFIG_PATH


def _read_config_author_name() -> str | None:
    """``author_name`` from the JennyStack config file, or None if the file
    is absent, unparseable, or the key is missing/empty. Never raises --
    a malformed or missing config is exactly the "not configured yet" case
    this function's caller falls back past, not an error."""
    path = _config_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    name = data.get("author_name")
    if isinstance(name, str) and name.strip():
        return name.strip()
    return None


def _macos_full_name() -> str | None:
    """The current user's full name via the password database's GECOS
    field (pw_gecos) -- verified on this machine to resolve to "Michael
    Sutton", matching macOS's own `id -F` / `dscl . -read RealName` output.
    GECOS is comma-separated (full name, office, work phone, home phone);
    only the first field is the name. Falls back to `id -F` (a macOS-
    specific command with the identical output) if GECOS is empty --
    defense in depth, not expected to differ from pw_gecos on a real
    macOS account, but pw_gecos being blank is not unheard of on other
    *nix account provisioning. Returns None (never raises) if both are
    unavailable, e.g. a non-macOS CI runner with no configured full name.
    """
    try:
        gecos = pwd.getpwuid(os.getuid()).pw_gecos
    except (KeyError, OSError):
        gecos = ""
    name = gecos.split(",", 1)[0].strip()
    if name:
        return name

    try:
        proc = subprocess.run(["id", "-F"], capture_output=True, text=True, timeout=5, check=False)
    except (OSError, subprocess.TimeoutExpired):
        return None
    name = proc.stdout.strip()
    return name or None


def resolve_author_name() -> str:
    """``author_name`` from the JennyStack config, else the macOS full
    name, else the literal string "Unknown Author" (never raises, never
    returns an empty string -- every w:author/w:ins author attribute this
    server writes needs SOME non-empty value)."""
    return _read_config_author_name() or _macos_full_name() or _FALLBACK_AUTHOR
