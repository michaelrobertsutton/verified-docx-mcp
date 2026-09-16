"""Session model for the WSS ops channel (issue #106 WP-2:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

One ``LiveSession`` per connected Word task pane, keyed in a
``SessionRegistry`` by the document's file name (the basename of the
``documentUrl`` the pane reports in its ``hello`` message -- see
``protocol.HelloMessage``). ``live/bridge.py``'s WSS connection handler
owns the actual ``websockets`` connection object and the receive loop that
feeds this module; everything here is transport-agnostic (it only needs
something with an async ``send(str)`` and a way to fail pending requests
on disconnect), which is what lets ``tests/unit/fake_pane.py`` and
``tests/unit/test_live_session.py`` exercise it against the real bridge
without a browser or Word.

Thread-safety: ``SessionRegistry``'s registration/lookup methods take a
plain ``threading.Lock``, not an asyncio one, because they are called from
two different threads in the real bridge -- the WSS connection handler
(running on the bridge's own asyncio event loop, in a background thread)
registers and unregisters sessions, while ``live_status`` and any future
``write_mode="live"`` tool call look sessions up from whatever thread
FastMCP runs the tool call on. A plain lock is correct here because every
critical section is synchronous, short, and never awaits.
"""

from __future__ import annotations

import asyncio
import json
import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Any, Protocol
from urllib.parse import unquote, urlparse

from .protocol import VALID_OPS, HelloMessage, OpError, OpReply, OpRequest

DEFAULT_HEARTBEAT_INTERVAL = 5.0
DEFAULT_MISSED_HEARTBEATS = 3
DEFAULT_REQUEST_TIMEOUT = 15.0


class LiveError(Exception):
    """Base class for every exception this module raises. server.py's
    tool layer (WP-3/WP-4) maps each subclass below to the matching
    ``ErrorCode`` in errors.py (LIVE_UNAVAILABLE / LIVE_DISCONNECTED /
    LIVE_STALE / LIVE_OP_FAILED)."""


class LiveUnavailable(LiveError):
    """No connected pane session exists for the requested document name."""


class LiveDisconnected(LiveError):
    """The pane's WebSocket closed -- or a reply never arrived within the
    timeout, which this module treats the same way: an unresponsive pane
    is operationally indistinguishable from a disconnected one, and a
    caller waiting on ``request()`` should not have to tell the two
    apart."""


class LiveOpFailed(LiveError):
    """The pane replied ``ok=false``. Carries the pane's own error code
    and message (``protocol.OpError``) rather than inventing one."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message


class LiveStale(LiveError):
    """The caller supplied ``expected_body_sha256`` and the pane's own
    ``result["pre"]`` (its body hash immediately before running the op)
    did not match it -- someone (the lead, typing) changed the document
    between the caller's last read and this op."""

    def __init__(self, expected: str, actual: str) -> None:
        super().__init__(f"expected body sha256 {expected!r}, pane reported {actual!r} before the op")
        self.expected = expected
        self.actual = actual


class PaneTransport(Protocol):
    """The minimum surface ``LiveSession`` needs from a WebSocket-like
    object. ``websockets.asyncio.server.ServerConnection`` satisfies this
    structurally; so does any test double with the same two methods."""

    async def send(self, message: str) -> None: ...

    async def close(self, code: int = 1000, reason: str = "") -> None: ...


def document_name_from_url(document_url: str) -> str:
    """Basename of a ``documentUrl``'s path component, URL-decoded.

    ``Office.context.document.url`` reports a local file path or a
    file:// / https:// (OneDrive/SharePoint) URL depending on where the
    document lives; in every case the basename is what the plan calls
    "the document's file name" and what ``lock_status``'s own path
    resolution already keys on for the file-mode side of ``write_mode``
    (WP-3). Falls back to the raw string if it has no path separator at
    all (defensive -- a real pane always reports a path-shaped URL).
    """
    parsed = urlparse(document_url)
    path = unquote(parsed.path) if parsed.scheme else document_url
    name = path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    return name or document_url


@dataclass
class LiveSession:
    """One connected pane. Not constructed directly by tool code --
    ``live/bridge.py``'s connection handler builds one per WSS connection
    from the pane's ``hello`` message and registers it."""

    document_name: str
    document_url: str
    hello: HelloMessage
    transport: PaneTransport
    loop: asyncio.AbstractEventLoop
    connected_since: float = field(default_factory=time.time)
    last_heartbeat_monotonic: float = field(default_factory=time.monotonic)
    last_body_sha256: str = ""
    _pending: dict[str, asyncio.Future[OpReply]] = field(default_factory=dict, repr=False)
    _closed: bool = field(default=False, repr=False)

    def __post_init__(self) -> None:
        if not self.last_body_sha256:
            self.last_body_sha256 = self.hello.body_sha256

    # -- heartbeat / liveness -------------------------------------------------

    def touch_heartbeat(self, body_sha256: str | None = None) -> None:
        self.last_heartbeat_monotonic = time.monotonic()
        if body_sha256:
            self.last_body_sha256 = body_sha256

    def heartbeat_age(self) -> float:
        return time.monotonic() - self.last_heartbeat_monotonic

    # -- request/reply ----------------------------------------------------
    #
    # `_request_on_loop` is the actual engine: it creates an
    # `asyncio.Future` via `self.loop.create_future()`, so it may only
    # ever run as a coroutine scheduled ON `self.loop` (the bridge's own
    # ops-thread event loop) -- `resolve_reply`, called from that same
    # loop's receive loop, resolves that future directly. `request()` and
    # `request_threadsafe()` below are the two public entry points that
    # get a caller from *any* thread/loop onto `self.loop` safely:
    #
    #   - `request()` (async): awaitable from any event loop, including
    #     `self.loop` itself. Detects whether the caller is already on
    #     `self.loop` (common once WP-3/4 wire a live tool call that
    #     happens to run there, and inside this module's own tests) and
    #     awaits the coroutine directly in that case; otherwise it
    #     schedules the coroutine onto `self.loop` via
    #     `asyncio.run_coroutine_threadsafe` and awaits THAT future
    #     instead (`asyncio.wrap_future` makes a `concurrent.futures.
    #     Future` awaitable).
    #   - `request_threadsafe()` (sync): for a plain, non-async caller --
    #     every `@mcp.tool()` function in server.py today is a plain
    #     `def`, not `async def`. Blocks the calling thread on the
    #     scheduled coroutine's `concurrent.futures.Future.result()`.
    #     Must NOT be called from `self.loop`'s own thread (it would
    #     deadlock: the coroutine it schedules can never run while this
    #     call blocks that same loop's thread).
    #
    # This split exists because of a real bug caught while smoke-testing
    # this module during WP-2: `future = self.loop.create_future()`
    # followed by `await asyncio.wait_for(future, ...)` from a coroutine
    # running on a DIFFERENT loop than `self.loop` raises `RuntimeError:
    # ... attached to a different loop` (or, for a sync caller, cannot be
    # awaited at all) -- exactly the situation both the fake-pane tests
    # and any future MCP tool call are in, since the bridge's ops
    # listener runs in its own background thread with its own loop
    # (`bridge.start_in_background`), never the caller's.

    async def request(
        self,
        op: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float = DEFAULT_REQUEST_TIMEOUT,
        expected_body_sha256: str | None = None,
        request_id: str | None = None,
    ) -> dict[str, Any]:
        """Send ``op`` with ``payload`` and await the pane's reply, from
        any event loop (see the block comment above). Raises
        ``LiveDisconnected`` if the socket closes while sending or before
        a reply arrives (including a plain timeout -- see that
        exception's docstring), ``LiveOpFailed`` if the pane replies
        ``ok=false``, and ``LiveStale`` if ``expected_body_sha256`` is
        given and disagrees with the reply's own ``result["pre"]``.
        Returns the reply's ``result`` dict (``{}`` if the pane sent
        none, e.g. a bare ``ping``).
        """
        coro = self._request_on_loop(
            op, payload, timeout=timeout, expected_body_sha256=expected_body_sha256, request_id=request_id
        )
        if asyncio.get_running_loop() is self.loop:
            return await coro
        concurrent_future = asyncio.run_coroutine_threadsafe(coro, self.loop)
        return await asyncio.wrap_future(concurrent_future)

    def request_threadsafe(
        self,
        op: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float = DEFAULT_REQUEST_TIMEOUT,
        expected_body_sha256: str | None = None,
        request_id: str | None = None,
    ) -> dict[str, Any]:
        """Synchronous counterpart to ``request()`` for a plain (non-
        async) caller -- see the block comment above. Blocks the calling
        thread; do not call from ``self.loop``'s own thread."""
        coro = self._request_on_loop(
            op, payload, timeout=timeout, expected_body_sha256=expected_body_sha256, request_id=request_id
        )
        concurrent_future = asyncio.run_coroutine_threadsafe(coro, self.loop)
        return concurrent_future.result(timeout=timeout + 5)

    async def _request_on_loop(
        self,
        op: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float = DEFAULT_REQUEST_TIMEOUT,
        expected_body_sha256: str | None = None,
        request_id: str | None = None,
    ) -> dict[str, Any]:
        """The actual send-and-await engine. MUST run as a coroutine
        scheduled on ``self.loop`` -- see the block comment above
        ``request()``. Not called directly by anything outside this
        class."""
        if op not in VALID_OPS:
            raise ValueError(f"unknown op {op!r}; must be one of {sorted(VALID_OPS)}")
        if self._closed:
            raise LiveDisconnected(f"session for {self.document_name!r} is already closed")

        rid = request_id or uuid.uuid4().hex
        message = OpRequest(request_id=rid, op=op, payload=payload or {})
        future: asyncio.Future[OpReply] = self.loop.create_future()
        self._pending[rid] = future
        try:
            await self.transport.send(_dump(message.to_json()))
        except Exception as exc:
            self._pending.pop(rid, None)
            raise LiveDisconnected(f"failed sending {op!r} to {self.document_name!r}: {exc}") from exc

        try:
            reply = await asyncio.wait_for(future, timeout=timeout)
        except TimeoutError as exc:
            raise LiveDisconnected(
                f"{op!r} to {self.document_name!r} timed out after {timeout}s waiting for a reply"
            ) from exc
        except asyncio.CancelledError:
            raise
        finally:
            self._pending.pop(rid, None)

        if not reply.ok:
            error = reply.error or OpError(code="LIVE_OP_FAILED", message="pane reported failure with no detail")
            raise LiveOpFailed(error.code, error.message)

        result = reply.result or {}
        if expected_body_sha256 is not None:
            pre = result.get("pre")
            if pre is not None and pre != expected_body_sha256:
                raise LiveStale(expected_body_sha256, pre)
        return result

    def resolve_reply(self, reply: OpReply) -> bool:
        """Called by the bridge's receive loop when a ``reply`` message
        arrives. Returns True if it matched a pending request (the normal
        case); False for a stray/late reply (already timed out, or a
        duplicate), which the bridge logs but does not treat as fatal."""
        future = self._pending.get(reply.request_id)
        if future is None or future.done():
            return False
        future.set_result(reply)
        return True

    def fail_pending(self, exc: BaseException) -> None:
        """Called on disconnect: unblocks every in-flight ``request()``
        immediately instead of making each one wait out its own timeout."""
        self._closed = True
        for rid, future in list(self._pending.items()):
            if not future.done():
                future.set_exception(exc)
            self._pending.pop(rid, None)


def _dump(obj: dict[str, Any]) -> str:
    return json.dumps(obj, ensure_ascii=False)


class SessionRegistry:
    """Thread-safe map of document name -> ``LiveSession``, plus stale
    eviction. One instance per running bridge (``bridge.start_in_background``
    owns it); ``live_status`` and (from WP-3 onward) the live write path
    both go through this registry rather than holding a session reference
    of their own.
    """

    def __init__(
        self,
        *,
        heartbeat_interval: float = DEFAULT_HEARTBEAT_INTERVAL,
        missed_heartbeats: int = DEFAULT_MISSED_HEARTBEATS,
    ) -> None:
        self.heartbeat_interval = heartbeat_interval
        self.missed_heartbeats = missed_heartbeats
        self.stale_after = heartbeat_interval * missed_heartbeats
        self._sessions: dict[str, LiveSession] = {}
        self._lock = threading.Lock()

    def register(self, session: LiveSession) -> None:
        with self._lock:
            self._sessions[session.document_name] = session

    def unregister(self, document_name: str) -> None:
        with self._lock:
            self._sessions.pop(document_name, None)

    def get(self, document_name: str) -> LiveSession | None:
        with self._lock:
            self._evict_stale_locked()
            return self._sessions.get(document_name)

    def list(self) -> list[LiveSession]:
        with self._lock:
            self._evict_stale_locked()
            return list(self._sessions.values())

    def touch_heartbeat(self, document_name: str, body_sha256: str | None = None) -> None:
        with self._lock:
            session = self._sessions.get(document_name)
            if session is not None:
                session.touch_heartbeat(body_sha256)

    def _evict_stale_locked(self) -> None:
        """Drop sessions that missed ``missed_heartbeats`` heartbeats at
        ``heartbeat_interval`` seconds each -- 3 x 5s = 15s by default.
        This only forgets the bookkeeping entry; the bridge's own
        receive-loop ``finally`` clause (on an actual socket close) is
        what calls ``fail_pending`` for any request still in flight, so a
        session evicted here for silence, not a socket error, simply
        stops being visible to new lookups."""
        now = time.monotonic()
        stale = [
            name
            for name, session in self._sessions.items()
            if now - session.last_heartbeat_monotonic > self.stale_after
        ]
        for name in stale:
            del self._sessions[name]

    async def request(
        self,
        document_name: str,
        op: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float = DEFAULT_REQUEST_TIMEOUT,
        expected_body_sha256: str | None = None,
    ) -> dict[str, Any]:
        """Convenience: look up ``document_name`` and forward to its
        session's ``request()``. Raises ``LiveUnavailable`` if no session
        is connected for that document (including one just evicted as
        stale)."""
        session = self.get(document_name)
        if session is None:
            raise LiveUnavailable(f"no connected pane session for document {document_name!r}")
        return await session.request(
            op, payload, timeout=timeout, expected_body_sha256=expected_body_sha256
        )

    def request_threadsafe(
        self,
        document_name: str,
        op: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float = DEFAULT_REQUEST_TIMEOUT,
        expected_body_sha256: str | None = None,
    ) -> dict[str, Any]:
        """Synchronous counterpart to ``request()`` -- see
        ``LiveSession.request_threadsafe``'s docstring for why both
        exist. This is the entry point a plain (non-async) FastMCP tool
        function calls."""
        session = self.get(document_name)
        if session is None:
            raise LiveUnavailable(f"no connected pane session for document {document_name!r}")
        return session.request_threadsafe(
            op, payload, timeout=timeout, expected_body_sha256=expected_body_sha256
        )
