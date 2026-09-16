"""Local HTTPS + WSS bridge (issue #106:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

Two things live here now:

  1. The WP-1 static HTTPS server (unchanged): serves `addin/` (the
     task-pane files) plus `GET /ping` and `POST /report`, stdlib
     `http.server` + `ssl`. Still reachable standalone via
     `--serve-only` for the lead's manual runbook
     (docs/live-mode.md) -- see `main()` below.

  2. WP-2's WSS ops channel: `websockets.asyncio.server.serve` running
     on its own asyncio event loop, in its own background thread,
     answering pane connections on `/ops`. Every connection speaks the
     `live/protocol.py` request/reply protocol; `_handle_pane_connection`
     turns each one into a `live/session.py` `LiveSession` registered in
     a shared `SessionRegistry`.

Listener layout: the WSS ops channel is a SEPARATE port (`ops_port`,
default `DEFAULT_PORT + 1`), not the same port as the static HTTPS
server. `http.server.HTTPServer` is a synchronous, blocking-accept
server; folding an asyncio `websockets` listener onto the same listening
socket would mean either rewriting the static side onto asyncio too (out
of scope for this WP -- WP-1's `/ping`/`/report` handlers are simple
stdlib code with their own passing tests) or running two disjoint I/O
loops fighting over one `accept()`, which stdlib does not support
cleanly. Two ports, two independent listener threads, is the boring
option and keeps WP-1's static server byte-for-byte as it was. The
manifest's `AppDomains` list gets a second entry for the ops port
(`addin/manifest.xml`); `taskpane.js` connects to `wss://localhost:<ops_
port>/ops` using a constant one line away from the existing `https://
localhost:<port>` constants it already hardcodes for `/ping`/`/report`.

`start_in_background()` is what `verified_docx_mcp.server` calls lazily,
on the first live-aware tool call (`live_status` in this WP; WP-3/4's
`write_mode="live"` tools later) -- idempotent (a second call while
already running just returns the existing `SessionRegistry`), so no tool
needs to track whether it already started the bridge. `stop()` tears both
listeners down; mainly for tests (each test starts and stops its own
bridge instance against ephemeral ports; the module-level singleton
tracked by `start_in_background`/`stop` is for the real MCP process,
where exactly one bridge should ever run).

Importing this module never starts a server (see the package's own
`__init__.py` docstring) -- everything here is a function some caller
invokes explicitly, whether that's `main()` (the CLI) or `server.py`'s
`live_status` tool (`start_in_background`).

Binds to 127.0.0.1 only, on both listeners. Office requires HTTPS for
add-in resources even on localhost, hence the cert step; there is no HTTP
fallback on the static side, and the ops channel is WSS (TLS) using the
same cert, never plain WS.
"""

from __future__ import annotations

import argparse
import asyncio
import dataclasses
import functools
import http.server
import json
import ssl
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

from websockets.asyncio.server import Server as WsServer
from websockets.asyncio.server import ServerConnection
from websockets.asyncio.server import serve as ws_serve

from .protocol import HeartbeatMessage, OpReply, ProtocolError, peek_type
from .session import LiveDisconnected, LiveSession, SessionRegistry, document_name_from_url

DEFAULT_PORT = 53135
DEFAULT_OPS_PORT = DEFAULT_PORT + 1
DEFAULT_HOST = "127.0.0.1"

# bridge.py lives at src/verified_docx_mcp/live/bridge.py; the addin/
# directory this WP-1 spike serves lives at the repo root, four levels
# up. addin/ is deliberately NOT in pyproject.toml's sdist include list
# (dev/spike tooling, not a package to ship in the wheel), so this path
# only resolves inside a checkout -- see main()'s explicit error when it
# does not, rather than a bare FileNotFoundError.
REPO_ADDIN_DIR = Path(__file__).resolve().parents[3] / "addin"

DEFAULT_CERT_DIR = Path.home() / ".cache" / "verified-docx-mcp" / "live-cert"
DEFAULT_REPORT_DIR = Path.home() / ".cache" / "verified-docx-mcp" / "live-report"
CERT_FILENAME = "localhost.pem"
KEY_FILENAME = "localhost-key.pem"


def trust_command(cert_path: Path) -> str:
    """The exact command the lead runs once to trust the cert -- LOGIN
    keychain, not System, so no sudo (deliverable 3's requirement)."""
    return f"security add-trusted-cert -d -r trustRoot -k ~/Library/Keychains/login.keychain-db {cert_path}"


def make_cert(cert_dir: Path, *, openssl_bin: str = "openssl") -> tuple[Path, Path]:
    """Generate a self-signed cert+key for `localhost` (SAN
    `DNS:localhost, IP:127.0.0.1`) via the `openssl` CLI. Returns
    `(cert_path, key_path)`. Always overwrites any existing cert in
    `cert_dir`; `main()` is the layer that decides whether to skip
    regeneration when one already exists."""
    cert_dir.mkdir(parents=True, exist_ok=True)
    cert_path = cert_dir / CERT_FILENAME
    key_path = cert_dir / KEY_FILENAME
    cmd = [
        openssl_bin,
        "req",
        "-x509",
        "-newkey",
        "rsa:2048",
        "-sha256",
        "-days",
        "825",
        "-nodes",
        "-keyout",
        str(key_path),
        "-out",
        str(cert_path),
        "-subj",
        "/CN=localhost",
        "-addext",
        "subjectAltName=DNS:localhost,IP:127.0.0.1",
    ]
    subprocess.run(cmd, check=True, capture_output=True, text=True)
    return cert_path, key_path


class _AddinRequestHandler(http.server.SimpleHTTPRequestHandler):
    """Serves `addin_dir` statically, plus a synthetic `/ping` JSON
    route the static handler cannot produce. `log_message` is silenced
    by default -- the lead runs this in a foreground terminal per
    docs/live-mode.md and does not need a request log for the spike."""

    def __init__(self, *args, addin_dir: Path, report_dir: Path = DEFAULT_REPORT_DIR, **kwargs):
        self._addin_dir = addin_dir
        self._report_dir = report_dir
        super().__init__(*args, directory=str(addin_dir), **kwargs)

    def do_POST(self) -> None:  # stdlib override name
        """`POST /report` (WP-1 convenience): the pane hands its JSON report
        straight to the bridge, so the lead never has to copy it out of Word
        by hand. Written to `<report_dir>/latest.json` (overwritten) and a
        timestamped sibling; replies with where it landed."""
        if not (self.path == "/report" or self.path.startswith("/report?")):
            self.send_error(404, "unknown POST route")
            return
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            report = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            self._send_json(400, {"ok": False, "error": f"body is not JSON: {exc}"})
            return
        self._report_dir.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%dT%H%M%S")
        stamped = self._report_dir / f"report-{stamp}.json"
        latest = self._report_dir / "latest.json"
        text = json.dumps(report, indent=2, ensure_ascii=False)
        stamped.write_text(text, encoding="utf-8")
        latest.write_text(text, encoding="utf-8")
        self._send_json(200, {"ok": True, "saved": str(latest), "stamped": str(stamped)})

    def _send_json(self, status: int, obj: dict) -> None:
        payload = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:  # stdlib override name, not our naming convention to control
        if self.path == "/ping" or self.path.startswith("/ping?"):
            self._send_ping()
            return
        super().do_GET()

    def _send_ping(self) -> None:
        payload = json.dumps({"ok": True, "server": "verified-docx-mcp", "time": time.time()}).encode(
            "utf-8"
        )
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args: object) -> None:  # stdlib override signature
        pass


def make_server(
    addin_dir: Path,
    *,
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    certfile: Path,
    keyfile: Path,
    report_dir: Path = DEFAULT_REPORT_DIR,
) -> http.server.HTTPServer:
    """Build (but do not start) an HTTPS `HTTPServer` bound to `host:port`
    serving `addin_dir`. `port=0` binds an ephemeral port -- the unit
    tests use that to avoid colliding with a real WP-1 run on
    DEFAULT_PORT; read the actual port back from `httpd.server_address`."""
    handler = functools.partial(_AddinRequestHandler, addin_dir=addin_dir, report_dir=report_dir)
    httpd = http.server.HTTPServer((host, port), handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=str(certfile), keyfile=str(keyfile))
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
    return httpd


# ---------------------------------------------------------------------------
# WP-2: WSS ops channel
# ---------------------------------------------------------------------------


def _make_ssl_context(certfile: Path, keyfile: Path) -> ssl.SSLContext:
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=str(certfile), keyfile=str(keyfile))
    return ctx


async def _handle_pane_connection(connection: ServerConnection, *, registry: SessionRegistry) -> None:
    """Per-connection handler passed to `websockets.asyncio.server.serve`.

    Rejects anything not on `/ops`, then requires the pane's first
    message to be a `hello` (protocol.HelloMessage) -- any other first
    message, or a connection that closes before sending one, is dropped
    without registering a session. Once registered, every subsequent
    message is either a `heartbeat` (updates the session's liveness) or
    a `reply` (resolves a pending `LiveSession.request()` future);
    anything else is ignored rather than treated as fatal, since a
    forward-compatible pane sending an unrecognized message type should
    not tear down an otherwise-working session.
    """
    request = getattr(connection, "request", None)
    path = getattr(request, "path", None)
    if path is not None and path.split("?", 1)[0] not in ("/ops", ""):
        await connection.close(code=1008, reason=f"unknown path {path!r}, expected /ops")
        return

    try:
        raw_hello = await connection.recv()
    except Exception:  # noqa: BLE001 - connection closed before sending anything
        return

    try:
        hello = _hello_from_raw(raw_hello)
    except ProtocolError as exc:
        await connection.close(code=1002, reason=f"first message must be 'hello': {exc}")
        return

    document_name = document_name_from_url(hello.document_url)
    session = LiveSession(
        document_name=document_name,
        document_url=hello.document_url,
        hello=hello,
        transport=connection,
        loop=asyncio.get_running_loop(),
    )
    registry.register(session)
    try:
        async for raw_msg in connection:
            try:
                msg_type = peek_type(raw_msg)
            except ProtocolError:
                continue
            if msg_type == "heartbeat":
                try:
                    hb = HeartbeatMessage.from_json(raw_msg)
                except ProtocolError:
                    continue
                registry.touch_heartbeat(document_name, hb.body_sha256)
            elif msg_type == "reply":
                try:
                    reply = OpReply.from_json(raw_msg)
                except ProtocolError:
                    continue
                session.resolve_reply(reply)
            # any other message type: ignored, not fatal (forward compat)
    finally:
        registry.unregister(document_name)
        session.fail_pending(LiveDisconnected(f"pane for {document_name!r} disconnected"))


def _hello_from_raw(raw: str | bytes):
    from .protocol import HelloMessage

    if peek_type(raw) != "hello":
        raise ProtocolError(f"expected 'hello' as the first message, got {peek_type(raw)!r}")
    return HelloMessage.from_json(raw)


@dataclasses.dataclass
class _RunningBridge:
    registry: SessionRegistry
    httpd: http.server.HTTPServer
    http_thread: threading.Thread
    ops_loop: asyncio.AbstractEventLoop
    ops_thread: threading.Thread
    ops_server: WsServer
    host: str
    port: int
    ops_port: int


_singleton_lock = threading.Lock()
_singleton: _RunningBridge | None = None


def start_in_background(
    *,
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    ops_port: int | None = None,
    cert_dir: Path = DEFAULT_CERT_DIR,
    addin_dir: Path = REPO_ADDIN_DIR,
    report_dir: Path = DEFAULT_REPORT_DIR,
    heartbeat_interval: float = 5.0,
    missed_heartbeats: int = 3,
    ready_timeout: float = 10.0,
) -> SessionRegistry:
    """Start the static HTTPS server and the WSS ops channel in two
    background daemon threads, generating a cert first if one is not
    already in `cert_dir`. Idempotent: a second call while a bridge is
    already running just returns the existing `SessionRegistry` (the
    `live_status` tool and any WP-3/4 live tool call both call this
    unconditionally, lazily, on their own first invocation -- none of
    them should need to track "did I already start this"). Not
    idempotent across `stop()`: call this again after `stop()` to start a
    fresh bridge (e.g. a new SessionRegistry, all prior sessions gone).
    """
    global _singleton
    with _singleton_lock:
        if _singleton is not None:
            return _singleton.registry

        resolved_ops_port = ops_port if ops_port is not None else port + 1
        cert_dir = Path(cert_dir).expanduser()
        cert_path = cert_dir / CERT_FILENAME
        key_path = cert_dir / KEY_FILENAME
        if not cert_path.exists() or not key_path.exists():
            cert_path, key_path = make_cert(cert_dir)

        httpd = make_server(addin_dir, host=host, port=port, certfile=cert_path, keyfile=key_path, report_dir=report_dir)
        http_thread = threading.Thread(
            target=httpd.serve_forever, daemon=True, name="verified-docx-mcp-live-http"
        )
        http_thread.start()

        registry = SessionRegistry(heartbeat_interval=heartbeat_interval, missed_heartbeats=missed_heartbeats)

        ready = threading.Event()
        state: dict[str, Any] = {}

        def _run_ops_loop() -> None:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            state["loop"] = loop
            ctx = _make_ssl_context(cert_path, key_path)

            async def _serve() -> None:
                server = await ws_serve(
                    functools.partial(_handle_pane_connection, registry=registry),
                    host,
                    resolved_ops_port,
                    ssl=ctx,
                )
                state["server"] = server
                state["ops_port"] = server.sockets[0].getsockname()[1]  # resolve an ephemeral (0) port
                ready.set()
                await server.wait_closed()

            try:
                loop.run_until_complete(_serve())
            finally:
                loop.close()

        ops_thread = threading.Thread(target=_run_ops_loop, daemon=True, name="verified-docx-mcp-live-ops")
        ops_thread.start()

        if not ready.wait(timeout=ready_timeout):
            # Best-effort cleanup of the half-started bridge before raising --
            # a caller that catches this can retry cleanly.
            httpd.shutdown()
            httpd.server_close()
            raise RuntimeError(
                f"WSS ops listener on {host}:{resolved_ops_port} did not start within {ready_timeout}s"
            )

        _singleton = _RunningBridge(
            registry=registry,
            httpd=httpd,
            http_thread=http_thread,
            ops_loop=state["loop"],
            ops_thread=ops_thread,
            ops_server=state["server"],
            host=host,
            port=httpd.server_address[1],
            ops_port=state["ops_port"],
        )
        return registry


def stop(*, timeout: float = 5.0) -> None:
    """Stop the running bridge (both listeners) started by
    `start_in_background`. A no-op if nothing is running."""
    global _singleton
    with _singleton_lock:
        running = _singleton
        _singleton = None
    if running is None:
        return

    try:
        running.httpd.shutdown()
        running.httpd.server_close()
    except Exception:  # noqa: BLE001, S110 - best-effort teardown, nothing to log to
        pass
    running.ops_loop.call_soon_threadsafe(running.ops_server.close)
    running.http_thread.join(timeout=timeout)
    running.ops_thread.join(timeout=timeout)


def current_registry() -> SessionRegistry | None:
    """The running bridge's `SessionRegistry`, or None if no bridge is
    running -- used by `live_status` to report `bridge_running: False`
    without itself starting a bridge just to answer the question is
    deliberately NOT what this is for; `live_status` per the WP-2 spec
    starts the bridge lazily, so it calls `start_in_background()`
    directly and only falls back to this for symmetry in tests that want
    to check bridge state without starting one."""
    with _singleton_lock:
        return _singleton.registry if _singleton is not None else None


def current_ports() -> tuple[int, int] | None:
    """(port, ops_port) of the running bridge, or None if not running."""
    with _singleton_lock:
        return (_singleton.port, _singleton.ops_port) if _singleton is not None else None


def _build_argparser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m verified_docx_mcp.live.bridge",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--serve-only",
        action="store_true",
        help=(
            "Serve addin/ (+/ping,/report) over local HTTPS and the WSS "
            "/ops channel, and block until Ctrl+C."
        ),
    )
    parser.add_argument(
        "--make-cert",
        metavar="DIR",
        nargs="?",
        const=str(DEFAULT_CERT_DIR),
        default=None,
        help=(
            "Generate a self-signed localhost cert in DIR "
            f"(default: {DEFAULT_CERT_DIR}) and print the trust command."
        ),
    )
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"default {DEFAULT_PORT}")
    parser.add_argument(
        "--ops-port",
        type=int,
        default=None,
        help=f"WSS /ops port for --serve-only (default: --port + 1, i.e. {DEFAULT_OPS_PORT}).",
    )
    parser.add_argument(
        "--cert",
        metavar="DIR",
        default=None,
        help=f"Cert directory to serve with (default: {DEFAULT_CERT_DIR}).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_argparser()
    args = parser.parse_args(argv)

    if args.make_cert:
        cert_dir = Path(args.make_cert).expanduser()
        cert_path, key_path = make_cert(cert_dir)
        print(f"wrote {cert_path}")
        print(f"wrote {key_path}")
        print()
        print("Trust it once (login keychain, no sudo required):")
        print(f"  {trust_command(cert_path)}")
        return 0

    if args.serve_only:
        cert_dir = Path(args.cert).expanduser() if args.cert else DEFAULT_CERT_DIR
        cert_path = cert_dir / CERT_FILENAME
        key_path = cert_dir / KEY_FILENAME
        if not cert_path.exists() or not key_path.exists():
            print(f"No cert found in {cert_dir}.", file=sys.stderr)
            print("Generate one first:", file=sys.stderr)
            print(f"  python -m verified_docx_mcp.live.bridge --make-cert {cert_dir}", file=sys.stderr)
            return 1
        if not REPO_ADDIN_DIR.exists():
            print(f"addin/ not found at {REPO_ADDIN_DIR}.", file=sys.stderr)
            print("Run this from a verified-docx-mcp checkout, not an installed wheel.", file=sys.stderr)
            return 1
        start_in_background(
            port=args.port,
            ops_port=args.ops_port,
            cert_dir=cert_dir,
            addin_dir=REPO_ADDIN_DIR,
        )
        bound_port, bound_ops_port = current_ports()  # type: ignore[misc]
        print(f"Serving {REPO_ADDIN_DIR} on https://{DEFAULT_HOST}:{bound_port}/ (Ctrl+C to stop)")
        print(f"WSS ops channel on wss://{DEFAULT_HOST}:{bound_ops_port}/ops")
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            pass
        finally:
            stop()
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
