"""``live_open_pane``: get a connected Live pane for a document, opening it if needed (#66)."""
import subprocess
import time
from typing import Any

from ..errors import ErrorCode, _make_error
from . import bridge as live_bridge
from . import open_pane_mac, write_mode


def _open_documents() -> list[str]:
    out = open_pane_mac._osascript('tell application "Microsoft Word" to get name of every document')
    return [n.strip() for n in out.split(",") if n.strip()]


def execute_open_pane(path: str, timeout_s: float = 60) -> dict[str, Any]:
    registry = live_bridge.start_in_background()
    name = write_mode._document_name(path)
    if registry.get(name) is not None:
        return {"opened": False, "already_connected": True, "document_name": name, "method": "already-connected"}
    try:
        if name not in _open_documents():
            subprocess.run(["open", "-a", "Microsoft Word", path], check=True, timeout=30)
            end = time.monotonic() + 30
            while name not in _open_documents():
                if time.monotonic() > end:
                    raise open_pane_mac.OpenPaneError(f"Word did not open {name!r}")
                time.sleep(1)
        method = open_pane_mac.open_pane(name, lambda: registry.get(name) is not None, timeout_s)
    except (open_pane_mac.OpenPaneError, subprocess.SubprocessError, OSError) as exc:
        raise _make_error(ErrorCode.LIVE_UNAVAILABLE, str(exc), {"document_name": name}) from exc
    return {"opened": method != "already-connected", "already_connected": method == "already-connected",
            "document_name": name, "method": method}
