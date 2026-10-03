"""Detect open Word documents independently of the optional task pane."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

from .errors import ErrorCode, _make_error

_SCRIPT = '''
if application "Microsoft Word" is not running then return ""
tell application "Microsoft Word"
    set paths to ""
    repeat with d in documents
        set documentPath to full name of d
        try
            set documentPath to POSIX path of (documentPath as alias)
        end try
        set paths to paths & documentPath & linefeed
    end repeat
    return paths
end tell
'''


def status(path: Path) -> dict:
    """macOS host query; an unavailable query is explicit, never 'closed'."""
    if sys.platform != "darwin":
        return {"supported": False, "open": False, "checked": False}
    try:
        result = subprocess.run(["/usr/bin/osascript", "-e", _SCRIPT],
                                capture_output=True, text=True, timeout=10, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"supported": True, "checked": False, "open": None, "detail": str(exc)}
    if result.returncode:
        return {"supported": True, "checked": False, "open": None,
                "detail": result.stderr.strip()}
    target = path.resolve()
    candidates = [unquote(urlparse(p).path) if "://" in p else p
                  for p in result.stdout.splitlines()]
    # If Word returns a cloud URL or an unresolved HFS path, a matching
    # basename is sufficient to refuse; it is never evidence of closure.
    opened = any(Path(p).resolve() == target or
                 p.replace(":", "/").rsplit("/", 1)[-1] == target.name
                 for p in candidates if p.strip())
    return {"supported": True, "checked": True, "open": opened}


def raise_if_open(path: Path) -> None:
    observed = status(path)
    if observed["supported"] and (observed["open"] or not observed["checked"]):
        raise _make_error(ErrorCode.EXTERNAL_EDITOR_ACTIVE,
                          "Close this document in Word before a file-mode write. "
                          "The desktop query found it open or could not confirm it is closed.",
                          {"desktop_word": observed})
