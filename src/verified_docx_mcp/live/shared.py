"""Experimental same-user shared bridge for macOS/POSIX (#47).

A private Unix socket carries JSON only, never pickle. flock elects one owner
before TLS listeners/certificates are created. All clients, including the
owner's MCP process, use the same broker path. No mutation is ever replayed.
"""

from __future__ import annotations

import asyncio
import copy
import json
import os
import socket
import socketserver
import stat
import threading
import time
import uuid
from pathlib import Path
from typing import Any

from .protocol import VALID_OPS, HelloMessage
from .session import LiveDisconnected, LiveOpFailed, LiveStale, resolve_timeout

READ_OPS = frozenset(
    {
        "ping",
        "describe",
        "body_ooxml",
        "search",
        "comments_list",
        "cell_get",
        "table_get",
        "shapes_list",
        "textboxes_list",
        "textboxes_read",
        "scope_describe",
        "revisions_list",
        "autoopen_get",
    }
)
MAX_FRAME = 72 * 2**20  # body_ooxml can return 48 MiB, with JSON escaping
CLIENT_TTL = 300.0
_client = None
_owner = None
_init_lock = threading.Lock()


def enabled() -> bool:
    return os.environ.get("VERIFIED_DOCX_SHARED_BRIDGE") == "1"


def runtime_dir() -> Path:
    # Keep below macOS's short AF_UNIX path limit, independent of repo/iCloud.
    return Path(f"/tmp/verified-docx-live-{os.getuid()}")


def private_dir(path: Path) -> None:
    path.mkdir(mode=0o700, parents=True, exist_ok=True)
    info = path.lstat()
    if not stat.S_ISDIR(info.st_mode) or info.st_uid != os.getuid() or info.st_mode & 0o077:
        raise RuntimeError(
            f"Shared bridge directory must be owned by this user and mode 0700: {path}"
        )


def _frame(stream) -> dict:
    raw = stream.readline(MAX_FRAME + 1)
    if len(raw) > MAX_FRAME or not raw.endswith(b"\n"):
        raise ValueError("Invalid/oversized shared bridge frame")
    value = json.loads(raw)
    if not isinstance(value, dict):
        raise TypeError("Shared bridge frame must be an object")
    return value


class Broker:
    def __init__(self, registry):
        self.registry = registry
        self.clients: dict[str, dict] = {}
        self.locks: dict[str, dict] = {}
        self.mutex = threading.RLock()
        self.doc_mutexes: dict[str, threading.Lock] = {}
        self.versions: dict[str, int] = {}
        self.epochs: dict[str, float] = {}
        self.seen: dict[tuple[str, str], int] = {}
        self.hashes: dict[tuple[str, str], str] = {}

    def _expire(self):
        now = time.monotonic()
        expired = [
            key for key, value in self.clients.items() if now - value["last_seen"] > CLIENT_TTL
        ]
        for key in expired:
            self.clients.pop(key)
        self.locks = {
            key: value
            for key, value in self.locks.items()
            if value["expires"] > now and value["client_id"] in self.clients
        }
        self.seen = {key: value for key, value in self.seen.items() if key[0] in self.clients}
        self.hashes = {key: value for key, value in self.hashes.items() if key[0] in self.clients}

    def dispatch(self, request: dict) -> Any:
        action = request.get("action")
        with self.mutex:
            self._expire()
            if action == "attach":
                client_id = uuid.uuid4().hex
                self.clients[client_id] = {"last_seen": time.monotonic()}
                return {"client_id": client_id, "protocol": 1}
            client_id = request.get("client_id", "")
            if client_id not in self.clients:
                raise LiveOpFailed(
                    "LIVE_DISCONNECTED", "Shared client expired; call live_status to reattach"
                )
            self.clients[client_id]["last_seen"] = time.monotonic()
            if action == "detach":
                self.clients.pop(client_id)
                self._expire()
                return {}
            if action == "status":
                return {
                    "clients": [
                        {"client_id": key, "idle_s": time.monotonic() - v["last_seen"]}
                        for key, v in self.clients.items()
                    ],
                    "locks": [
                        {
                            "document_url": key,
                            "client_id": v["client_id"],
                            "scope": "document",
                            "remaining_s": v["expires"] - time.monotonic(),
                        }
                        for key, v in self.locks.items()
                    ],
                }
        if action == "list":
            return [self.snapshot(s) for s in self.registry.list()]
        if action == "collisions":
            return self.registry.collisions()
        if action not in {"request", "lock", "unlock"}:
            raise ValueError("Unknown shared bridge action")
        session = self.registry.get(request["document_name"])
        if session is None:
            raise LiveOpFailed("LIVE_UNAVAILABLE", "No pane connected for this document")
        if (
            session.document_url != request["document_url"]
            or session.connected_since != request["connected_since"]
        ):
            raise LiveOpFailed("LIVE_SESSION_MISMATCH", "Pane changed; obtain a fresh session")
        if any(
            c.get("kind") == "basename_collision"
            and c.get("document_name") == session.document_name
            for c in self.registry.collisions()
        ):
            raise LiveOpFailed(
                "LIVE_SESSION_MISMATCH",
                "Ambiguous filename collision; use uniquely named documents and restart bridge",
            )
        key = session.document_url
        with self.mutex:
            gate = self.doc_mutexes.setdefault(key, threading.Lock())
        timeout = resolve_timeout(request.get("op", "describe"))
        if not gate.acquire(timeout=timeout):
            raise LiveOpFailed("LIVE_OP_FAILED", "Document queue busy; no operation dispatched")
        try:
            # Recheck identity after waiting; never route into a displaced pane.
            if self.registry.get(session.document_name) is not session:
                raise LiveOpFailed("LIVE_SESSION_MISMATCH", "Pane changed while queued")
            with self.mutex:
                self._expire()
                if self.epochs.get(key) != session.connected_since:
                    self.epochs[key] = session.connected_since
                    self.versions[key] = self.versions.get(key, 0) + 1
                    self.seen = {k: v for k, v in self.seen.items() if k[1] != key}
                    self.hashes = {k: v for k, v in self.hashes.items() if k[1] != key}
                    self.locks.pop(key, None)
                lease = self.locks.get(key)
                if action in {"lock", "unlock"}:
                    if lease and lease["client_id"] != client_id:
                        raise LiveOpFailed(
                            "LOCKED_BY_OTHER_CLIENT", "Document is locked by another client"
                        )
                    if action == "unlock":
                        self.locks.pop(key, None)
                        return {"released": True}
                    duration = request.get("lease_s", 60)
                    if not isinstance(duration, (int, float)) or not 1 <= duration <= 300:
                        raise ValueError("lease_s must be between 1 and 300")
                    self.locks[key] = {
                        "client_id": client_id,
                        "expires": time.monotonic() + duration,
                    }
                    return {"client_id": client_id, "scope": "document", "lease_s": duration}
            op = request["op"]
            if op not in VALID_OPS:
                raise ValueError("Unknown pane operation")
            write = op not in READ_OPS
            payload = dict(request.get("payload") or {})
            with self.mutex:
                version = self.versions.get(key, 0)
                if write:
                    if lease and lease["client_id"] != client_id:
                        raise LiveOpFailed(
                            "LOCKED_BY_OTHER_CLIENT", "Document is locked by another client"
                        )
                    if self.seen.get((client_id, key)) != version:
                        raise LiveOpFailed(
                            "LIVE_STALE", "Read the live document after the other client's write"
                        )
            if write:
                if "shared_queue" not in session.hello.capabilities:
                    raise LiveOpFailed(
                        "LIVE_CAPABILITY_MISSING",
                        "Reopen the pane to load shared_queue before shared writes",
                    )
                expected = request.get("expected_body_sha256") or payload.get("expectedBodySha256")
                # Never silently refresh a client's old read baseline at write time.
                if not expected:
                    with self.mutex:
                        expected = self.hashes.get((client_id, key))
                if not expected:
                    raise LiveOpFailed(
                        "LIVE_STALE", "Read the live document body before a shared write"
                    )
                payload["expectedBodySha256"] = expected
            payload["clientId"] = client_id
            try:
                result = session.request_threadsafe(
                    op,
                    payload,
                    timeout=timeout,
                    expected_body_sha256=request.get("expected_body_sha256"),
                )
            finally:
                if write:
                    # Failure may be post-mutation. Invalidate other baselines
                    # even when the pane times out or reports verification failure.
                    with self.mutex:
                        self.versions[key] = version + 1
            with self.mutex:
                if write or request.get("observe", True):
                    self.seen[(client_id, key)] = self.versions.get(key, 0)
                    body_hash = result.get("bodySha256") or result.get("post")
                    if body_hash:
                        self.hashes[(client_id, key)] = body_hash
            return result
        finally:
            gate.release()

    @staticmethod
    def snapshot(session):
        return {
            "hello": session.hello.to_json(),
            "document_name": session.document_name,
            "connected_since": session.connected_since,
            "heartbeat_age": session.heartbeat_age(),
            "last_body_sha256": session.last_body_sha256,
        }


class Handler(socketserver.StreamRequestHandler):
    server: Server

    def handle(self):
        self.request.settimeout(300)
        try:
            result = self.server.broker.dispatch(_frame(self.rfile))
            response = {"ok": True, "result": result}
        except LiveStale as exc:
            response = {"ok": False, "code": "LIVE_STALE", "message": str(exc)}
        except LiveOpFailed as exc:
            response = {"ok": False, "code": exc.code, "message": exc.message}
        except LiveDisconnected as exc:
            response = {"ok": False, "code": "LIVE_DISCONNECTED", "message": str(exc)}
        except Exception as exc:  # noqa: BLE001 - RPC boundary must return a typed failure
            response = {"ok": False, "code": "LIVE_OP_FAILED", "message": str(exc)}
        self.wfile.write(json.dumps(response).encode() + b"\n")


class Server(socketserver.ThreadingUnixStreamServer):
    broker: Broker
    daemon_threads = True


class RemoteSession:
    def __init__(self, client, snapshot):
        self.client = client
        self._observe = True
        self.hello = HelloMessage.from_json(snapshot["hello"])
        self.document_name = snapshot["document_name"]
        self.document_url = self.hello.document_url
        self.connected_since = snapshot["connected_since"]
        self.last_body_sha256 = snapshot["last_body_sha256"]
        self._age = snapshot["heartbeat_age"]
        self._at = time.monotonic()

    def status_view(self):
        view = copy.copy(self)
        view._observe = False
        return view

    def heartbeat_age(self):
        return self._age + time.monotonic() - self._at

    def target(self):
        return {
            "document_name": self.document_name,
            "document_url": self.document_url,
            "connected_since": self.connected_since,
        }

    def request_threadsafe(
        self, op, payload=None, *, timeout=None, expected_body_sha256=None, request_id=None
    ):
        if request_id is not None:
            raise ValueError("Shared bridge assigns request IDs; caller IDs are unsupported")
        return self.client.rpc(
            "request",
            **self.target(),
            op=op,
            payload=payload,
            expected_body_sha256=expected_body_sha256,
            observe=self._observe,
            rpc_timeout=(timeout or resolve_timeout(op)) * 2 + 10,
        )

    async def request(self, op, payload=None, **kwargs):
        return await asyncio.to_thread(self.request_threadsafe, op, payload, **kwargs)


class Client:
    def __init__(self, path: Path):
        self.path = path
        self.client_id = ""
        attached = self.rpc("attach")
        if attached["protocol"] != 1:
            raise RuntimeError("Incompatible shared bridge protocol; restart all MCP servers")
        self.client_id = attached["client_id"]
        self._closed = threading.Event()
        threading.Thread(target=self._heartbeat, daemon=True).start()

    def _heartbeat(self):
        while not self._closed.wait(30):
            try:
                self.rpc("status")
            except (LiveDisconnected, LiveOpFailed):
                return

    def close(self):
        self._closed.set()
        try:
            self.rpc("detach")
        except (LiveDisconnected, LiveOpFailed):
            pass

    def rpc(self, action, *, rpc_timeout=10, **params):
        try:
            with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as sock:
                sock.settimeout(rpc_timeout)
                sock.connect(str(self.path))
                sock.sendall(
                    json.dumps({"action": action, "client_id": self.client_id, **params}).encode()
                    + b"\n"
                )
                with sock.makefile("rb") as stream:
                    response = _frame(stream)
        except (OSError, ValueError, TypeError) as exc:
            raise LiveDisconnected(
                "Shared bridge unavailable; result may be unknown. Do not replay writes. "
                "Call live_status to reconnect."
            ) from exc
        if not response["ok"]:
            raise LiveOpFailed(response["code"], response["message"])
        return response["result"]

    def list(self):
        return [RemoteSession(self, value) for value in self.rpc("list")]

    def get(self, name):
        return next((s for s in self.list() if s.document_name == name), None)

    def collisions(self):
        return self.rpc("collisions")

    def status(self):
        return {"client_id": self.client_id, **self.rpc("status")}


def connect(start_local, *, directory=None, ready_timeout=10):
    """Elect owner or attach. Called only by explicit bridge startup."""
    import fcntl

    global _client, _owner
    with _init_lock:
        if _client is not None:
            try:
                _client.status()
                return _client
            except (LiveDisconnected, LiveOpFailed):
                _client.close()
                _client = None
        directory = directory or runtime_dir()
        private_dir(directory)
        socket_path = directory / "broker.sock"
        # Refuse symlinks before opening an election file in the private dir.
        fd = os.open(directory / "owner.lock", os.O_CREAT | os.O_RDWR | os.O_NOFOLLOW, 0o600)
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            os.close(fd)
            deadline = time.monotonic() + ready_timeout
            while True:
                try:
                    _client = Client(socket_path)
                    return _client
                except LiveDisconnected:
                    if time.monotonic() >= deadline:
                        raise
                    time.sleep(0.1)
        try:
            registry = start_local()
            socket_path.unlink(missing_ok=True)  # only elected owner removes stale socket
            server = Server(str(socket_path), Handler)
            os.chmod(socket_path, 0o600)
            server.broker = Broker(registry)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            _owner = (server, thread, fd, socket_path)
            _client = Client(socket_path)
            return _client
        except Exception:
            os.close(fd)
            raise


def current():
    return _client


def stop():
    global _client, _owner
    with _init_lock:
        if _client:
            _client.close()
            _client = None
        if _owner:
            server, thread, fd, socket_path = _owner
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)
            socket_path.unlink(missing_ok=True)
            os.close(fd)
            _owner = None
