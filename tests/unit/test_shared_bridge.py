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


def test_pane_section_guard():
    import shutil
    from pathlib import Path

    node = shutil.which("node")
    if not node:
        pytest.skip("node unavailable")
    root = Path(__file__).resolve().parents[2]
    subprocess.run(
        [node, str(root / "tests/unit/js/section_guard_harness.mjs"), str(root / "addin/taskpane.js")],
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


# -- section locks (#47) ---------------------------------------------------------


class SectionPane(Pane):
    hello = HelloMessage(
        Pane.document_url,
        "Word",
        "Mac",
        {},
        "initial",
        frozenset({"shared_queue", "section_locks"}),
        "pane-1",
    )

    def __init__(self):
        super().__init__()
        self.refuse = None
        self.sent = []
        self.sha = {"intro": "s-intro", "scope": "s-scope", "scope-detail": "s-detail", "dup": "s-dup"}
        self.sections = [
            ("intro", 1, True),
            ("scope", 1, True),
            ("scope-detail", 2, True),
            ("dup", 1, False),
        ]

    def request_threadsafe(self, op, payload=None, **kwargs):
        if op == "sections_list":
            self.calls.append((op, payload))
            return {
                "bodySha256": self.last_body_sha256,
                "sections": [
                    {
                        "section_key": f"{slug}-1",
                        "slug": slug,
                        "heading_text": slug,
                        "level": level,
                        "slug_unique": unique,
                        "sectionSha256": self.sha[slug],
                    }
                    for slug, level, unique in self.sections
                ],
            }
        if self.refuse and op not in shared.READ_OPS:
            self.calls.append((op, payload))
            self.sent.append((op, dict(payload)))
            raise LiveOpFailed(self.refuse, "refused")
        self.kwargs = kwargs
        self.sent.append((op, dict(payload)))
        if op not in shared.READ_OPS:
            payload = {**payload}
            payload.setdefault("expectedBodySha256", self.last_body_sha256)
        return super().request_threadsafe(op, payload, **kwargs)


@pytest.fixture
def sections():
    pane = SectionPane()
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


def _last_write(pane):
    return [p for op, p in pane.sent if op == "replace"][-1]


def test_section_locks_conflict_and_forward_scopes(sections):
    pane, _, a, b, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    call(b, action="section_lock", sections=["intro-1"])
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, action="section_lock", sections=["scope-1"])
    call(a, "replace")
    own = _last_write(pane)
    assert own["ownSections"] == [{"slug": "scope", "expectedSha256": "s-scope"}]
    assert own["forbiddenSections"] == ["intro"]
    assert "expectedBodySha256" not in own


def test_section_holder_is_not_staled_by_other_clients_writes(sections):
    pane, _, a, b, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    call(b)  # B has no lock: explicit read baseline
    call(b, "replace")
    call(a, "replace")  # A never read, but is not stale: checked per section
    assert pane.kwargs["expected_body_sha256"] is None
    call(a, "replace")


def test_lock_requires_unique_heading_known_key_and_capability(sections):
    pane, _, a, _, call = sections
    with pytest.raises(LiveOpFailed, match="LOCK_SCOPE_UNRESOLVED"):
        call(a, action="section_lock", sections=["dup-1"])
    with pytest.raises(LiveOpFailed, match="LOCK_SCOPE_UNRESOLVED"):
        call(a, action="section_lock", sections=["nope-1"])
    with pytest.raises(ValueError, match="sections"):
        call(a, action="section_lock", sections=[])
    pane.hello = HelloMessage(pane.document_url, "Word", "Mac", {}, "initial", frozenset({"shared_queue"}))
    with pytest.raises(LiveOpFailed, match="LIVE_CAPABILITY_MISSING"):
        call(a, action="section_lock", sections=["intro-1"])


def test_include_subsections_stops_at_same_or_higher_level(sections):
    _, _, a, _, call = sections
    result = call(a, action="section_lock", sections=["scope-1"], include_subsections=True)
    assert [s["section_key"] for s in result["sections"]] == ["scope-1", "scope-detail-1"]


def test_document_lease_and_section_locks_exclude_each_other(sections):
    _, _, a, b, call = sections
    call(a, action="section_lock", sections=["intro-1"])
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, action="lock")
    with pytest.raises(ValueError, match="Release your section locks"):
        call(a, action="lock")
    call(a, action="section_unlock")
    call(b, action="lock")
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(a, action="section_lock", sections=["intro-1"])
    call(b, action="unlock")
    call(a, action="lock")
    with pytest.raises(ValueError, match="Release your document lease"):
        call(a, action="section_lock", sections=["intro-1"])


def test_unlock_by_key_foreign_refused_and_status_lists_locks(sections):
    _, broker, a, b, call = sections
    call(a, action="section_lock", sections=["intro-1", "scope-1"])
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, action="section_unlock", sections=["intro-1"])
    status = broker.dispatch({"action": "status", "client_id": a})["locks"]
    assert {lock["section_key"] for lock in status if lock["scope"] == "section"} == {
        "intro-1",
        "scope-1",
    }
    assert call(a, action="section_unlock", sections=["intro-1"])["released"] == ["intro-1"]
    assert call(a, action="section_unlock")["released"] == ["scope-1"]


def test_section_lock_expiry_and_reconnect_release(sections):
    pane, broker, a, b, call = sections
    call(a, action="section_lock", sections=["intro-1"], lease_s=5)
    broker.section_locks[pane.document_url]["intro"]["expires"] = time.monotonic() - 1
    call(b, action="section_lock", sections=["intro-1"])
    pane.connected_since = 2.0
    call(b)
    assert pane.document_url not in broker.section_locks


def test_caller_supplied_scopes_are_stripped(sections):
    pane, _, a, _, call = sections
    call(a)
    call(a, "replace", payload={"forbiddenSections": [], "ownSections": [{"slug": "x"}]})
    sent = _last_write(pane)
    assert "ownSections" not in sent and "forbiddenSections" not in sent


def test_pre_mutation_refusal_does_not_stale_anyone_but_failure_does(sections):
    pane, _, a, b, call = sections
    call(a)
    call(b)
    pane.refuse = "LOCKED_BY_OTHER_CLIENT"
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(a, "replace")
    pane.refuse = None
    call(b, "replace")  # B's baseline survived A's refused attempt
    call(a)
    pane.refuse = "VERIFICATION_FAILED"
    with pytest.raises(LiveOpFailed, match="VERIFICATION_FAILED"):
        call(a, "replace")
    pane.refuse = None
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")


def test_holder_baseline_refreshes_on_own_write_and_explicit_read(sections):
    pane, _, a, _, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    pane.sha["scope"] = "after-write"
    call(a, "replace")
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "after-write"
    pane.sha["scope"] = "human-edit"
    call(a, "body_ooxml")
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "human-edit"
    pane.sha["scope"] = "unseen-edit"
    call(a, "describe")  # not an explicit body read: baseline stays
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "human-edit"


def test_renewing_a_lock_does_not_rebaseline(sections):
    pane, _, a, _, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    pane.sha["scope"] = "changed-behind-our-back"
    call(a, action="section_lock", sections=["scope-1"])
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "s-scope"


def test_basename_collision_refuses_every_request(setup_broker):
    pane, broker, a, _, call = setup_broker
    call(a)
    broker.registry.collisions = lambda: [
        {"kind": "basename_collision", "document_name": pane.document_name}
    ]
    for op in ("describe", "replace"):
        with pytest.raises(LiveOpFailed, match="LIVE_SESSION_MISMATCH"):
            call(a, op)
    with pytest.raises(LiveOpFailed, match="LIVE_SESSION_MISMATCH"):
        call(a, action="lock")
