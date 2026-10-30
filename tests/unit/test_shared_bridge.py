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
    body_hash_state = "sampled"
    body_hash_sampled_monotonic = time.monotonic()
    heartbeat_count = 0
    hash_sample_count = 0
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
            if op == "save":  # issue #107: persistence carries no body baseline
                return {"saved": True}
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


def test_body_hash_read_rearms_a_stale_client_like_any_observed_read(setup_broker):
    """#110 live_revision: an observed body_hash is an explicit read."""
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")
    call(b, "body_hash")
    call(b, "replace")
    assert pane.calls[-1][0] == "replace"


def test_non_observing_body_hash_never_rearms_a_refused_client(setup_broker):
    """A lookup made while refusing a write (observe=False, the status view) must not
    re-arm the client that was just refused: the retry has to stay refused."""
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")
    call(b, "body_hash", observe=False)
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")
    assert sum(op == "replace" for op, _ in pane.calls) == 1


def test_section_holder_baseline_refreshes_on_observed_body_hash_only(sections):
    pane, _, a, _, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    pane.sha["scope"] = "seen-by-live-revision"
    call(a, "body_hash", observe=False)  # error-path lookup: baseline stays
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "s-scope"
    pane.sha["scope"] = "seen-by-live-revision-2"
    call(a, "body_hash")  # live_revision: an explicit read re-baselines
    call(a, "replace")
    assert _last_write(pane)["ownSections"][0]["expectedSha256"] == "seen-by-live-revision-2"


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


# --- #83: default-on, ownership identity, re-election, shutdown -------------


@pytest.mark.parametrize(
    ("value", "platform", "expected"),
    [
        (None, "darwin", True),
        (None, "linux", True),
        ("1", "darwin", True),
        ("0", "darwin", False),
        ("false", "linux", False),
        ("OFF", "linux", False),
        ("no", "linux", False),
        (None, "win32", False),
        ("1", "win32", False),
    ],
)
def test_enabled_matrix(monkeypatch, value, platform, expected):
    if value is None:
        monkeypatch.delenv("VERIFIED_DOCX_SHARED_BRIDGE", raising=False)
    else:
        monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", value)
    monkeypatch.setattr(sys, "platform", platform)
    assert shared.enabled() is expected


def test_start_in_background_kwargs_stay_local(monkeypatch):
    from verified_docx_mcp.live import bridge

    monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", "1")
    monkeypatch.setattr(shared, "connect", lambda *a, **k: pytest.fail("shared used"))
    monkeypatch.setattr(bridge, "_start_local", lambda **kw: ("local", kw))
    assert bridge.start_in_background(port=0) == ("local", {"port": 0})


def test_owner_pid_file_and_broker_status(monkeypatch):
    from pathlib import Path

    from verified_docx_mcp.live import bridge

    registry = SimpleNamespace(list=list, get=lambda name: None, collisions=list)
    with tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-pid-") as directory:
        monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", "1")
        monkeypatch.setattr(shared, "runtime_dir", lambda: Path(directory))
        monkeypatch.setattr(bridge, "_start_local", lambda: registry)
        try:
            client = bridge.start_in_background()
            pid_file = Path(directory) / "owner.pid"
            assert pid_file.read_text().strip() == str(os.getpid())
            assert pid_file.stat().st_mode & 0o777 == 0o600
            assert client.status()["owner_pid"] == os.getpid()

            from verified_docx_mcp import server

            status = server.execute_live_status()
            assert status["pid"] == status["owner_pid"] == os.getpid()
            assert status["is_owner"] is True
            assert status["mode"] == "shared"
        finally:
            bridge.stop()


def test_dead_owner_triggers_reelection(monkeypatch):
    from pathlib import Path

    from verified_docx_mcp.live import bridge

    script = """import os, sys
from pathlib import Path
from types import SimpleNamespace
from verified_docx_mcp.live import shared
def start():
    return SimpleNamespace(list=list, get=lambda name:None, collisions=list)
shared.connect(start, directory=Path(sys.argv[1]))
print('ready', flush=True)
sys.stdin.readline()
os._exit(0)
"""
    registry = SimpleNamespace(list=list, get=lambda name: None, collisions=list)
    with tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-reelect-") as directory:
        monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", "1")
        monkeypatch.setattr(shared, "runtime_dir", lambda: Path(directory))
        monkeypatch.setattr(bridge, "_start_local", lambda: registry)
        owner = subprocess.Popen(
            [sys.executable, "-c", script, directory],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            text=True,
        )
        try:
            assert owner.stdout.readline().strip() == "ready"
            client = bridge.start_in_background()
            assert client.status()["owner_pid"] == owner.pid
            assert bridge.current_registry() is client
            owner.communicate("\n", timeout=15)
            with pytest.raises(shared.LiveDisconnected, match="re-attaches automatically"):
                client.status()
            assert client.dead
            fresh = bridge.current_registry()
            assert fresh is not client and not fresh.dead
            assert fresh.status()["owner_pid"] == os.getpid()
        finally:
            owner.kill()
            bridge.stop()


def test_heartbeat_takes_over_from_dead_owner_without_a_tool_call(monkeypatch):
    """#93: an idle client re-elects itself, so the pane ports come back."""
    from pathlib import Path

    from verified_docx_mcp.live import bridge

    script = """import os, sys
from pathlib import Path
from types import SimpleNamespace
from verified_docx_mcp.live import shared
shared.connect(lambda: SimpleNamespace(list=list, get=lambda n: None, collisions=list),
               directory=Path(sys.argv[1]))
print('ready', flush=True)
sys.stdin.readline()
os._exit(0)
"""
    registry = SimpleNamespace(list=list, get=lambda name: None, collisions=list)
    with tempfile.TemporaryDirectory(dir="/tmp", prefix="docx-takeover-") as directory:
        monkeypatch.setenv("VERIFIED_DOCX_SHARED_BRIDGE", "1")
        monkeypatch.setattr(shared, "runtime_dir", lambda: Path(directory))
        monkeypatch.setattr(shared, "HEARTBEAT_S", 0.2)
        started = []
        monkeypatch.setattr(bridge, "_start_local", lambda: started.append(1) or registry)
        owner = subprocess.Popen(
            [sys.executable, "-c", script, directory],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            text=True,
        )
        try:
            assert owner.stdout.readline().strip() == "ready"
            client = bridge.start_in_background()
            assert client.status()["owner_pid"] == owner.pid and not started
            owner.communicate("\n", timeout=15)
            deadline = time.monotonic() + 10
            while shared.current() is client and time.monotonic() < deadline:
                time.sleep(0.05)
            fresh = shared.current()
            assert fresh is not client and client.dead
            assert started == [1]
            assert fresh.status()["owner_pid"] == os.getpid()
        finally:
            owner.kill()
            bridge.stop()


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX signals")
def test_shutdown_handler_exits_with_signal_code(monkeypatch):
    from verified_docx_mcp import shutdown

    exits = []
    stopped = threading.Event()
    monkeypatch.setattr(shutdown.os, "_exit", exits.append)
    shutdown.make_shutdown_handler(stopped.set)(15, None)
    assert stopped.is_set() and exits == [143]

    # A wedged teardown must not block the exit past the join timeout.
    monkeypatch.setattr(shutdown, "JOIN_TIMEOUT_S", 0.1)
    block = threading.Event()
    shutdown.make_shutdown_handler(block.wait)(1, None)
    assert exits == [143, 129]
    block.set()


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX signals")
def test_sigterm_terminates_stdio_server():
    import signal

    env = {**os.environ, "VERIFIED_DOCX_SHARED_BRIDGE": "0"}
    proc = subprocess.Popen(
        [sys.executable, "-m", "verified_docx_mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    try:
        time.sleep(3)  # let the server install handlers and block on stdin
        assert proc.poll() is None, proc.stderr.read()
        proc.send_signal(signal.SIGTERM)
        assert proc.wait(timeout=5) == 128 + signal.SIGTERM
    finally:
        if proc.poll() is None:
            proc.kill()
        proc.communicate()


def test_cancelled_tool_call_does_not_end_stdio_server(tmp_path):
    """#93: the late result of a cancelled sync tool must not crash the server."""
    script = tmp_path / "slow_server.py"
    script.write_text(
        "import sys, time\n"
        "from verified_docx_mcp import server\n"
        "from verified_docx_mcp.live import open_pane_live\n"
        "open_pane_live.execute_open_pane = lambda path, timeout_s=60: time.sleep(2) or {}\n"
        "sys.argv = ['verified-docx-mcp']\n"
        "server.main()\n"
    )
    env = {**os.environ, "VERIFIED_DOCX_SHARED_BRIDGE": "0"}
    proc = subprocess.Popen(
        [sys.executable, str(script)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )

    def send(message):
        proc.stdin.write((json.dumps({"jsonrpc": "2.0", **message}) + "\n").encode())
        proc.stdin.flush()

    def reply(request_id):
        while True:
            line = proc.stdout.readline()
            assert line, proc.stderr.read().decode()
            message = json.loads(line)
            if message.get("id") == request_id:
                return message

    try:
        send({"id": 1, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "t", "version": "1"}}})
        reply(1)
        send({"method": "notifications/initialized"})
        send({"id": 2, "method": "tools/call",
              "params": {"name": "live_open_pane", "arguments": {"path": "/tmp/x.docx"}}})
        time.sleep(0.5)
        send({"method": "notifications/cancelled", "params": {"requestId": 2}})
        assert reply(2)["error"]["message"] == "Request cancelled"
        time.sleep(3)  # the worker thread finishes after the cancel
        assert proc.poll() is None, proc.stderr.read().decode()
        send({"id": 3, "method": "tools/list"})
        assert "result" in reply(3)
    finally:
        proc.kill()
        proc.communicate()


def test_preflight_guard_refusal_does_not_invalidate_baselines(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    real = pane.request_threadsafe

    def refuse(op, payload=None, **kwargs):
        if op == "replace":
            raise LiveOpFailed("ANCHORED_SHAPES", "paragraph anchors a shape")
        return real(op, payload, **kwargs)

    pane.request_threadsafe = refuse
    with pytest.raises(LiveOpFailed, match="ANCHORED_SHAPES"):
        call(a, "replace")
    pane.request_threadsafe = real
    call(a, "replace")  # nothing was written, so neither client needs a fresh read


def test_lone_client_may_write_without_a_prior_read():
    pane = Pane()
    registry = SimpleNamespace(
        get=lambda name: pane if name == pane.document_name else None,
        list=lambda: [pane],
        collisions=list,
    )
    broker = shared.Broker(registry)
    only = broker.dispatch({"action": "attach"})["client_id"]
    request = {
        "action": "request",
        "client_id": only,
        "document_name": pane.document_name,
        "document_url": pane.document_url,
        "connected_since": pane.connected_since,
        "op": "replace",
    }
    broker.dispatch(request)
    assert [op for op, _ in pane.calls] == ["describe", "replace"]
    assert pane.calls[-1][1]["expectedBodySha256"] == "initial"
    # A second client attaching restores the strict rule for a client that never read.
    other = broker.dispatch({"action": "attach"})["client_id"]
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        broker.dispatch({**request, "client_id": other})


def test_approved_run_keeps_owned_lock_membership_without_stale_hash_refresh(sections):
    import dataclasses
    pane,broker,a,_b,call=sections
    pane.hello=dataclasses.replace(pane.hello,capabilities=pane.hello.capabilities | {'approved_changes_v1'})
    call(a,action='section_lock',sections=['scope-1'])
    before=len([1 for op,_ in pane.calls if op=='sections_list'])
    for n in range(3):
        call(a,'replace',payload={'executionSections':[{'slug':'scope','paragraphs':[str(n)],'tables':[]}]})
        assert _last_write(pane)['ownSections']==[{'slug':'scope'}]
        assert 'expectedBodySha256' not in _last_write(pane)
    assert len([1 for op,_ in pane.calls if op=='sections_list'])==before
    assert not broker.hashes


def test_scoped_hash_never_becomes_legacy_body_baseline(sections):
    import dataclasses
    pane,broker,a,_b,call=sections
    pane.hello=dataclasses.replace(pane.hello,capabilities=pane.hello.capabilities | {'approved_changes_v1'})
    call(a)
    call(a,'replace',payload={'executionSections':[{'slug':'scope','paragraphs':[],'tables':[]}]})
    assert not broker.hashes
    with pytest.raises(LiveOpFailed,match='LIVE_STALE'):
        call(a,'replace')
    call(a)  # an explicit body read restores the legacy baseline
    call(a,'replace')


def test_scoped_deadline_matches_broker_and_remote_rpc(sections):
    import dataclasses
    pane,_broker,a,_b,call=sections
    pane.hello=dataclasses.replace(pane.hello,capabilities=pane.hello.capabilities | {'approved_changes_v1'})
    original=pane.request_threadsafe
    budgets=[]
    def capture(op,payload=None,**kw):
        budgets.append(kw['timeout'])
        return original(op,payload,**kw)
    pane.request_threadsafe=capture
    payload={'executionSections':[{'slug':'scope','paragraphs':[],'tables':[]}]}
    call(a,'replace',payload=payload)
    assert budgets==[60]
    remote=object.__new__(shared.RemoteSession)
    remote.target=dict
    remote._observe=True
    requests=[]
    remote.client=SimpleNamespace(rpc=lambda action,**kw: requests.append(kw))
    remote.request_threadsafe('replace',payload)
    assert requests[0]['rpc_timeout']==130


def test_heartbeat_sample_cannot_refresh_client_read_baseline(setup_broker):
    pane, broker, a, b, call = setup_broker
    call(a)
    # A human edit observed by a background sample is not an explicit read.
    pane.last_body_sha256 = 'human-edit'
    pane.body_hash_state = 'sampled'
    pane.hash_sample_count += 1
    assert broker.snapshot(pane)['last_body_sha256'] == 'human-edit'
    assert broker.hashes[(a, pane.document_url)] == 'initial'
    def checked(op, payload=None, **kwargs):
        assert payload['expectedBodySha256'] == 'initial'
        raise LiveOpFailed('LIVE_STALE', 'Human content changed before mutation')
    pane.request_threadsafe = checked
    with pytest.raises(LiveOpFailed, match='Human content changed'):
        call(a, op='replace')
    assert broker.hashes[(a, pane.document_url)] == 'initial'


# -- save is persistence, not a content write or a read (#107) --------------------


def test_save_needs_no_baseline_and_is_not_refused_because_someone_else_wrote(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")  # b is now stale for content writes
    assert call(b, "save")["saved"] is True  # but may still save (no read, no baseline needed)
    assert "expectedBodySha256" not in pane.calls[-1][1]
    # A client that never read anything may save too.
    c = _attach_third(setup_broker)
    assert call(c, "save")["saved"] is True


def _attach_third(setup):
    _, broker, _, _, _ = setup
    return broker.dispatch({"action": "attach"})["client_id"]


def test_save_neither_acknowledges_a_read_nor_stales_other_clients(setup_broker):
    pane, broker, a, b, call = setup_broker
    call(a)
    call(b)
    call(a, "replace")
    versions = dict(broker.versions)
    seen = dict(broker.seen)
    hashes = dict(broker.hashes)
    call(b, "save")
    # The stale client was not re-armed by saving ...
    assert (broker.versions, broker.seen, broker.hashes) == (versions, seen, hashes)
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")
    # ... and the client that wrote earlier is not staled by b's save.
    call(a, "save")
    call(a, "replace")


def test_a_failing_save_does_not_stale_other_clients(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    pane.fail = True
    with pytest.raises(LiveOpFailed, match="VERIFICATION_FAILED"):
        call(a, "save")
    pane.fail = False
    call(b, "replace")  # a save never changes content, so nobody's baseline moved


def test_save_still_honours_the_document_lease(setup_broker):
    _, _, a, b, call = setup_broker
    call(a, action="lock")
    with pytest.raises(LiveOpFailed, match="LOCKED_BY_OTHER_CLIENT"):
        call(b, "save")
    assert call(a, "save")["saved"] is True


def test_save_still_carries_other_clients_section_locks(sections):
    pane, _, a, b, call = sections
    call(a, action="section_lock", sections=["scope-1"])
    call(b)
    call(b, "save")
    forwarded = [p for op, p in pane.sent if op == "save"][-1]
    assert forwarded["forbiddenSections"] == ["scope"]
    # The lock holder's save is confined to its own sections and needs no whole-body hash.
    call(a, "save")
    own = [p for op, p in pane.sent if op == "save"][-1]
    assert own["ownSections"] == [{"slug": "scope", "expectedSha256": "s-scope"}]


def test_save_needs_the_shared_queue_capability_like_any_write(setup_broker):
    pane, _, a, _, call = setup_broker
    pane.hello = HelloMessage(pane.document_url, "Word", "Mac", {}, "initial")
    with pytest.raises(LiveOpFailed, match="LIVE_CAPABILITY_MISSING"):
        call(a, "save")
    assert all(op != "save" for op, _ in pane.calls)


# -- issue #108: cheap re-sync after partial / unknown batches -------------------------------


class DocPane:
    """A broker-registered pane over a real FakeDocument, so text_edit's batch and single
    replace run end to end through the broker's stale/lock/unknown barriers."""

    document_name = "Proposal.docx"
    document_url = "/x/Proposal.docx"
    connected_since = 1.0
    body_hash_state = "sampled"
    body_hash_sampled_monotonic = time.monotonic()
    heartbeat_count = 0
    hash_sample_count = 0
    hello = HelloMessage(
        document_url, "Word", "Mac", {}, "initial",
        frozenset({"shared_queue", "section_locks", "body_hash_v1", "replacement_formatting",
                   "comment_loss_guard", "formatted_replacements_v1", "partial_font_segments_v1"}),
        "pane-1",
    )

    def __init__(self, text):
        from fake_pane import FakeDocument

        self.doc = FakeDocument(text=text)
        self.calls = []
        self.fail_after_apply = False  # apply the replace, then report VERIFICATION_FAILED

    @property
    def last_body_sha256(self):
        return self.doc.sha256()

    def heartbeat_age(self):
        return 0

    def request_threadsafe(self, op, payload=None, **kwargs):
        from fake_pane import OpRefused

        payload = payload or {}
        self.calls.append((op, dict(payload)))
        if op in {"describe", "body_hash"}:
            return {"bodySha256": self.doc.sha256(), "sections": []}
        if op == "sections_list":
            return {"bodySha256": self.doc.sha256(), "sections": [
                {"section_key": "scope-1", "slug": "scope", "heading_text": "Scope", "level": 1,
                 "slug_unique": True, "sectionSha256": "s"}]}
        if op == "replacement_font":
            try:
                return self.doc.replacement_font(payload["find"], int(payload["expected_matches"]))
            except OpRefused as exc:
                raise LiveOpFailed(exc.code, exc.message) from None
        if op != "replace":
            raise AssertionError(op)
        if payload.get("forbiddenSections"):
            raise LiveOpFailed("LOCKED_BY_OTHER_CLIENT", "target lies in another client's locked section")
        try:
            result = self.doc.replace(
                payload["find"], int(payload["expected_matches"]), payload["replace"],
                expected_body_sha256=payload.get("expectedBodySha256"),
                replacement_segments=payload.get("replacement_segments"),
                expected_first_font=payload.get("expected_first_font"),
            )
        except OpRefused as exc:
            raise LiveOpFailed(exc.code, exc.message) from None
        if self.fail_after_apply:
            raise LiveOpFailed("VERIFICATION_FAILED", "Replacement may have applied")
        return result


_DOC = {"path": ""}


@pytest.fixture
def doc_broker(monkeypatch, tmp_path):
    """(pane, broker, a, b, use): use(client_id) routes every tool call through that client."""
    from verified_docx_mcp.live import bridge, write_mode

    monkeypatch.setenv("VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS", str(tmp_path))
    monkeypatch.setitem(_DOC, "path", str(tmp_path / "Proposal.docx"))

    pane = DocPane("alpha beta gamma delta epsilon")
    broker = shared.Broker(SimpleNamespace(
        get=lambda name: pane if name == pane.document_name else None, list=lambda: [pane], collisions=list))
    current = {}

    def use(client_id):
        client = SimpleNamespace(
            rpc=lambda action, **kw: broker.dispatch({"action": action, "client_id": client_id, **kw}))
        current["remote"] = shared.RemoteSession(client, broker.snapshot(pane))

    monkeypatch.setattr(bridge, "current_registry", lambda: SimpleNamespace(
        get=lambda name: current["remote"], collisions=list))
    monkeypatch.setattr(write_mode, "_check_session_identity", lambda path, session: None)
    monkeypatch.setattr(write_mode, "_document_name", lambda path: "Proposal.docx")
    a = broker.dispatch({"action": "attach"})["client_id"]
    b = broker.dispatch({"action": "attach"})["client_id"]
    return pane, broker, a, b, use


def _revision():
    from verified_docx_mcp.live import reads_live

    return reads_live.live_revision(_DOC["path"])["revision"]


def _batch(edits, revision_before):
    from verified_docx_mcp import text_edit

    return text_edit.execute_replace_text_batch_live(
        _DOC["path"], [{"find": f, "replace": r, "expected_matches": 1} for f, r in edits], revision_before)


def _replace(find, replace, revision_before):
    from verified_docx_mcp import text_edit

    return text_edit.execute_replace_text_live(_DOC["path"], find, replace, 1, revision_before=revision_before)


def _reads(pane):
    return sum(op in {"body_hash", "body_ooxml"} for op, _ in pane.calls)


def _code(exc_info):
    return exc_info.value.envelope.error_code.value


def test_partial_batch_revision_after_is_accepted_without_a_body_read(doc_broker):
    from verified_docx_mcp.errors import VerifyError

    pane, _, a, b, use = doc_broker
    use(b)
    _revision()
    use(a)
    start = _revision()
    report = _batch([("alpha", "A"), ("missing", "X"), ("gamma", "G")], start)
    assert report["batch_status"] == "partial"
    assert [r["status"] for r in report["results"]] == ["applied", "refused_before_write", "not_attempted"]
    reads = _reads(pane)
    result = _replace("gamma", "G", report["revision_after"])
    assert result["revision_after"] == f"live:sha256:{pane.doc.sha256()}"
    assert _reads(pane) == reads  # the batch's own revision_after was proof enough
    assert pane.doc.text == "A beta G delta epsilon"
    # ... and the other client is still stale: it never saw A's writes.
    use(b)
    with pytest.raises(VerifyError) as exc:
        _replace("delta", "D", result["revision_after"])
    assert _code(exc) == "LIVE_STALE"
    assert pane.doc.text == "A beta G delta epsilon"


def test_pane_match_count_refusal_codes_never_stale_anyone(setup_broker):
    """A count-gate refusal is pre-mutation whether the pane names it (zero_match /
    match_count_mismatch) or, like older panes, sends LIVE_OP_FAILED with the gate's message."""
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    original = pane.request_threadsafe
    for code, message in [("zero_match", "x"), ("match_count_mismatch", "x"),
                          ("LIVE_OP_FAILED", 'expected 1 match(es) for "missing", found 0'),
                          ("LIVE_OP_FAILED", 'expected 1 match(es) for "a", found 2')]:
        def refuse(op, payload=None, _code=code, _message=message, **kwargs):
            if op == "replace":
                raise LiveOpFailed(_code, _message)
            return original(op, payload, **kwargs)

        pane.request_threadsafe = refuse
        with pytest.raises(LiveOpFailed):
            call(a, "replace")
        pane.request_threadsafe = original
        call(b, "replace")  # b was not staled by a's refused write
        call(a)
        call(b)


def test_unrecognised_live_op_failed_still_stales_other_clients(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    original = pane.request_threadsafe

    def boom(op, payload=None, **kwargs):
        if op == "replace":
            raise LiveOpFailed("LIVE_OP_FAILED", "GeneralException after insertText")
        return original(op, payload, **kwargs)

    pane.request_threadsafe = boom
    with pytest.raises(LiveOpFailed):
        call(a, "replace")
    pane.request_threadsafe = original
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        call(b, "replace")


def test_unknown_batch_outcome_refuses_until_live_revision(doc_broker):
    from verified_docx_mcp.errors import VerifyError
    from verified_docx_mcp.live import bridge

    pane, _, a, _, use = doc_broker
    use(a)
    start = _revision()
    pane.fail_after_apply = True
    report = _batch([("alpha", "A"), ("gamma", "G")], start)
    pane.fail_after_apply = False
    assert report["batch_status"] == "unknown"
    assert report["revision_after"] is None
    # Not even the true current hash re-arms a write: only an observed read does.
    with pytest.raises(VerifyError) as exc:
        _replace("gamma", "G", f"live:sha256:{pane.doc.sha256()}")
    assert _code(exc) == "LIVE_STALE"
    # A non-observing lookup (live_status / an error path's status view) never re-arms.
    bridge.current_registry().get("Proposal.docx").status_view().request_threadsafe("body_hash")
    with pytest.raises(VerifyError) as exc:
        _replace("gamma", "G", f"live:sha256:{pane.doc.sha256()}")
    assert _code(exc) == "LIVE_STALE"
    assert pane.doc.text == "A beta gamma delta epsilon"
    revision = _revision()  # live_revision: one body_hash, no body
    _replace("gamma", "G", revision)
    assert pane.doc.text == "A beta G delta epsilon"


def test_resync_keeps_lease_and_section_lock_barriers(doc_broker):
    from verified_docx_mcp.errors import VerifyError

    pane, broker, a, b, use = doc_broker
    use(a)
    start = _revision()
    report = _batch([("alpha", "A"), ("missing", "X")], start)
    assert report["batch_status"] == "partial"
    target = {"client_id": b, "document_name": pane.document_name, "document_url": pane.document_url,
              "connected_since": pane.connected_since}
    broker.dispatch({"action": "lock", **target})
    with pytest.raises(VerifyError) as exc:
        _replace("gamma", "G", report["revision_after"])
    assert _code(exc) == "LOCKED_BY_OTHER_CLIENT"
    broker.dispatch({"action": "unlock", **target})
    broker.dispatch({"action": "section_lock", "sections": ["scope-1"], **target})
    with pytest.raises(VerifyError) as exc:
        _replace("gamma", "G", report["revision_after"])
    assert _code(exc) == "LOCKED_BY_OTHER_CLIENT"
    assert pane.doc.text == "A beta gamma delta epsilon"


def test_partial_segment_font_lookup_never_rearms_a_stale_client(doc_broker):
    """#108: replace_text's replacement_font lookup is read-only and non-observing, so a
    client staled by another client's write is still refused, and nothing is written."""
    from verified_docx_mcp import text_edit
    from verified_docx_mcp.errors import VerifyError

    pane, _, a, b, use = doc_broker
    use(b)
    stale = _revision()
    use(a)
    _replace("alpha", "A", _revision())
    use(b)
    with pytest.raises(VerifyError) as exc:
        text_edit.execute_replace_text_live(
            _DOC["path"], "gamma", "G", 1, revision_before=f"live:sha256:{pane.doc.sha256()}",
            replacement_segments=[{"text": "G", "format": {"bold": True}}])
    assert _code(exc) == "LIVE_STALE"
    assert [op for op, _ in pane.calls].count("replacement_font") == 1
    assert pane.doc.text == "A beta gamma delta epsilon"
    assert stale != f"live:sha256:{pane.doc.sha256()}"
    _revision()  # an explicit read re-arms; the partial segment then resolves and lands
    text_edit.execute_replace_text_live(
        _DOC["path"], "gamma", "G", 1, revision_before=f"live:sha256:{pane.doc.sha256()}",
        replacement_segments=[{"text": "G", "format": {"bold": True}}])
    assert pane.doc.text == "A beta G delta epsilon"


def test_format_drift_refusal_does_not_stale_other_clients(setup_broker):
    pane, _, a, b, call = setup_broker
    call(a)
    call(b)
    original = pane.request_threadsafe

    def drift(op, payload=None, **kwargs):
        if op == "replace":
            raise LiveOpFailed("FORMAT_DRIFT", "first run changed")
        return original(op, payload, **kwargs)

    pane.request_threadsafe = drift
    with pytest.raises(LiveOpFailed, match="FORMAT_DRIFT"):
        call(a, "replace")
    pane.request_threadsafe = original
    call(b, "replace")
