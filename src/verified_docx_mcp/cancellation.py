"""Keep a cancelled tool call from ending the MCP server (#93).

When the client cancels a request, the mcp SDK answers it at once with
"Request cancelled". A sync tool runs in a worker thread that the cancel cannot
interrupt, so the handler still returns normally afterwards and the SDK tries to
respond a second time. That trips ``assert not self._completed`` in
``RequestResponder.respond``, which escapes the server's task group and exits
the process -- taking the shared live bridge down with it when this process owns
it. The guard drops the late, duplicate response instead.
"""

from __future__ import annotations

import logging

from mcp.shared.session import RequestResponder

logger = logging.getLogger(__name__)


def install_cancel_guard() -> None:
    original = RequestResponder.respond
    if getattr(original, "_verified_docx_cancel_guard", False):
        return

    async def respond(self, response):  # type: ignore[no-untyped-def]
        if getattr(self, "_completed", False):
            logger.info("Request %s already answered (cancelled); late response dropped", self.request_id)
            return
        await original(self, response)

    respond._verified_docx_cancel_guard = True  # type: ignore[attr-defined]
    RequestResponder.respond = respond  # type: ignore[method-assign]
