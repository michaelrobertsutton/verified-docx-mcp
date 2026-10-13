"""SIGTERM/SIGHUP handling so a terminated MCP process releases the live bridge.

The handler takes no locks itself: it starts the teardown on a daemon thread,
waits briefly, then exits unconditionally so a wedged teardown cannot hang it.
"""

from __future__ import annotations

import os
import signal
import sys
import threading
from collections.abc import Callable

JOIN_TIMEOUT_S = 2.0


def make_shutdown_handler(stop: Callable[[], None]) -> Callable[[int, object], None]:
    def handler(signum: int, _frame: object) -> None:
        worker = threading.Thread(target=stop, daemon=True, name="verified-docx-mcp-shutdown")
        worker.start()
        worker.join(JOIN_TIMEOUT_S)
        os._exit(128 + signum)

    return handler


def install_shutdown_handlers(stop: Callable[[], None]) -> None:
    if sys.platform == "win32":
        return
    handler = make_shutdown_handler(stop)
    for name in ("SIGTERM", "SIGHUP"):
        signal.signal(getattr(signal, name), handler)
