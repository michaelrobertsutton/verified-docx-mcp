"""Broker safety and actual cross-process Unix-socket transport; no real Word."""

import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from types import SimpleNamespace

import pytest

from verified_docx_mcp.live import shared
from verified_docx_mcp.live.protocol import HelloMessage
from verified_docx_mcp.live.session import LiveOpFailed


class Pane:
    document_name = "Proposal.docx"
    document_url = "https://example.test/site/Proposal.docx"
    connected_since = 1.0
    last_body_sha256 = "initial"
    hello = HelloMessage(
        document_url, "Word", "Mac", {}, "initial", frozenset({"shared_queue"}), "pane-1"
    )

    def __init__(self):
        self.calls = []
        self.active = 0
        self.max_active = 0
        self.fail = False

    def heartbeat_age(self):
        return 0

    def request_threadsafe(self, op, payload=None, **kwargs):
        self.calls.append((op, payload))
        self.active += 1
        self.max_active = max(self.active, self.max_active)
        try:
            time.sleep(0.01)
            if self.fail:
                raise LiveOpFailed("VERIFICATION_FAILED", "May have applied")
            if op in shared.READ_OPS:
                return {"bodySha256": self.last_body_sha256}
            assert payload["expectedBodySha256"] == self.last_body_sha256
            self.last_body_sha256 = "changed"
            return {"pre": "initial", "post": "changed"}
        finally:
            self.active -= 1


@pytest.fixture
def setup_broker():
    pane = Pane()
    registry = SimpleNamespace(
        get=lambda name: pane if name == pane.document_name else None,
        list=lambda: [pane],
        collisions=list,
    )
    broker = shared.Broker(registry)
    a = broker.dispatch({"action": "attach"})["client_id"]
    b = broker.dispatch({"action": "attach"})["client_id"]

    def call(client, op="describe", action="request", **kwargs):
        return broker.dispatch(
            {
                "action": action,
                "client_id": client,
                "document_name": pane.document_name,
                "document_url": pane.document_url,
                "connected_since": pane.connected_since,
                "op": op,
                **kwargs,
            }
        )

    return pane, broker, a, b, call


def test_stale_other_client_refuses_before_dispatch(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")
    assert sum(op == "replace" for op, _ in pane.calls) == 1
    call(b)
    call(b, "replace")
    assert pane.calls[-1][1]["clientId"] == b


def test_lock_blocks_writes_but_allows_reads_and_expires(setup_broker):
    pane, broker, a, b, call = setup_broker
    call(a, action="lock")
    call(b)
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, "replace")
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, action="unlock")
    broker.locks[pane.document_url]["expires"] = time.monotonic() - 1
    call(b, "replace")


def test_failed_write_invalidates_other_clients(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    pane.fail = True
    with pytest.raises(LiveOpFailed, match="VERIFICATION_FAILED"):
        call(a, "replace")
    pane.fail = False
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")


def test_missing_baseline_old_pane_and_wrong_identity_refuse(setup_broker):
    pane, _, a, _, call = setup_broker
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(a, "replace")
    call(a)
    with pytest.raises(LiveOpFailed, match="LIVE_SESSION_MISMATCH"):
        call(a, document_url="https://other.test/Proposal.docx")
    pane.hello = HelloMessage(pane.document_url, "Word", "Mac", {}, "initial")
    with pytest.raises(LiveOpFailed, match="LIVE_CAPABILITY_MISSING"):
        call(a, "replace")
    assert all(op == "describe" for op, _ in pane.calls)


def test_concurrent_operations_are_serial(setup_broker):
    pane, _, a, b, call = setup_broker
    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(lambda n: call(a if n % 2 else b), range(16)))
    assert pane.max_active == 1


def test_unix_rpc_between_processes_and_client_detach(tmp_path, setup_broker):
    pane, broker, _, _, _ = setup_broker
    short_dir = tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-rpc-")
    path = __import__("pathlib").Path(short_dir.name) / "rpc.sock"
    server = shared.Server(str(path), shared.Handler)
    server.broker = broker
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    client = shared.Client(path)
    try:
        session = client.get(pane.document_name)
        session.request_threadsafe("describe")
        script = """import json, sys
from pathlib import Path
from verified_docx_mcp.live.shared import Client
c = Client(Path(sys.argv[1]))
s = c.get("Proposal.docx")
s.request_threadsafe("describe")
print(json.dumps({"id":c.client_id,"result":s.request_threadsafe("replace")}))
c.close()
"""
        result = subprocess.run(
            [sys.executable, "-c", script, str(path)],
            capture_output=True,
            text=True,
            check=True,
            timeout=15,
        )
        assert json.loads(result.stdout)["id"] != client.client_id
        with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
            session.request_threadsafe("replace")
        assert len(client.status()["clients"]) == 3  # two fixture clients + this client
    finally:
        client.close()
        server.shutdown()
        server.server_close()
        thread.join()
        short_dir.cleanup()


def test_private_directory_rejects_world_access_and_symlink(tmp_path):
    directory = tmp_path / "public"
    directory.mkdir(mode=0o755)
    os.chmod(directory, 0o755)
    with pytest.raises(RuntimeError, match="0700"):
        shared.private_dir(directory)
    link = tmp_path / "link"
    link.symlink_to(directory)
    with pytest.raises(RuntimeError):
        shared.private_dir(link)


def test_pane_queue_and_stale_preflight():
    import shutil
    from pathlib import Path

    node = shutil.which("node")
    if not node:
        pytest.skip("node unavailable")
    root = Path(__file__).resolve().parents[2]
    subprocess.run(
        [
            node,
            str(root / "tests/unit/js/shared_queue_harness.mjs"),
            str(root / "addin/taskpane.js"),
        ],
        check=True,
        capture_output=True,
        timeout=15,
    )


def test_reconnect_invalidates_client_baselines(setup_broker):
    pane, _, a, _, call = setup_broker
    call(a)
    pane.connected_since = 2.0
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(a, "replace")


def test_owner_election_across_processes():
    import select
    from pathlib import Path

    script = """import sys
from pathlib import Path
from types import SimpleNamespace
from verified_docx_mcp.live import shared
root = Path(sys.argv[1])
def start():
    with (root / 'starts').open('a') as f: f.write('started\\n')
    return SimpleNamespace(list=list, get=lambda name:None, collisions=list)
c = shared.connect(start, directory=root)
print(c.client_id, flush=True)
sys.stdin.readline()
shared.stop()
"""
    with tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-elect-") as directory:
        children = [
            subprocess.Popen(
                [sys.executable, "-c", script, directory],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            for _ in range(2)
        ]
        try:
            ids = []
            for child in children:
                assert select.select([child.stdout], [], [], 15)[0], "election timed out"
                ids.append(child.stdout.readline().strip())
            assert ids[0] and ids[1] and ids[0] != ids[1]
            assert (Path(directory) / "starts").read_text().splitlines() == ["started"]
        finally:
            for child in children:
                _out, err = child.communicate("\n", timeout=15)
                assert child.returncode == 0, err


def test_bridge_integration_returns_proxy_even_for_owner(monkeypatch, setup_broker):
    from pathlib import Path

    from verified_docx_mcp.live import bridge

    pane, broker, _, _, _ = setup_broker
    with tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-owner-") as directory:
        monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", "1")
        monkeypatch.setattr(shared, "runtime_dir", lambda: Path(directory))
        monkeypatch.setattr(bridge, "_start_local", lambda: broker.registry)
        try:
            registry = bridge.start_in_background()
            assert isinstance(registry, shared.Client)
            assert bridge.current_registry() is registry
            assert bridge.start_in_background() is registry
            remote = registry.get(pane.document_name)
            remote.request_threadsafe("describe")
            remote.request_threadsafe("replace")
            assert pane.calls[-1][1]["clientId"] == registry.client_id
        finally:
            bridge.stop()
        assert shared.current() is None


def test_status_probe_does_not_acknowledge_another_write(setup_broker):
    _, _, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")
    call(b, observe=False)
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")


def test_lease_rejects_bool_and_out_of_range(setup_broker):
    _, _, a, _, call = setup_broker
    for bad in (True, 0, 301, "60"):
        with pytest.raises(ValueError, match="lease_s"):
            call(a, action="lock", lease_s=bad)
    assert call(a, action="lock", lease_s=5)["lease_s"] == 5


def test_write_tool_session_does_not_acknowledge_other_writes(setup_broker, monkeypatch):
    """A write tool's internal body read must not re-baseline a stale client."""
    from verified_docx_mcp.live import bridge, write_mode

    pane, broker, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")

    client = SimpleNamespace(rpc=lambda action, **kw: broker.dispatch({"action": action, "client_id": b, **kw}))
    remote = shared.RemoteSession(client, broker.snapshot(pane))
    monkeypatch.setattr(bridge, "current_registry", lambda: SimpleNamespace(get=lambda n: remote))
    monkeypatch.setattr(write_mode, "_check_session_identity", lambda path, session: None)
    monkeypatch.setattr(write_mode, "_document_name", lambda path: "Proposal.docx")
    session = write_mode.live_session_for("/x/Proposal.docx")
    session.request_threadsafe("body_ooxml")  # tool-internal read
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        session.request_threadsafe("replace")
    assert sum(op == "replace" for op, _ in pane.calls) == 1
