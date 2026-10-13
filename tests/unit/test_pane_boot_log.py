"""Issue #84: bridge-side pane boot log, `recent_boots`, shared-mode status, and
the `pane_reload_loop` warning in `live_status`. No real Word."""

from __future__ import annotations

import os
from types import SimpleNamespace

import pytest

from verified_docx_mcp import server
from verified_docx_mcp.live import bridge, shared

DOC_A = "https://contoso.sharepoint.com/sites/x/Row01.docx"
DOC_B = "/Users/me/Scratch.docx"


@pytest.fixture(autouse=True)
def clean_boot_log():
    with bridge._BOOT_LOG_LOCK:
        bridge._BOOT_LOG.clear()
    yield
    with bridge._BOOT_LOG_LOCK:
        bridge._BOOT_LOG.clear()


def boot(url=DOC_A, age_s=0.0, **extra):
    """Record a boot as if it arrived `age_s` seconds ago."""
    bridge.record_boot({"documentUrl": url, "boot_seq": 1, **extra})
    with bridge._BOOT_LOG_LOCK:
        bridge._BOOT_LOG[-1]["monotonic"] -= age_s


def test_record_boot_ignores_non_dict_and_bad_shapes():
    bridge.record_boot("nope")
    bridge.record_boot({"documentUrl": 5, "previous_events": "x", "load_error": None})
    boots = bridge.recent_boots()
    assert len(boots) == 1
    assert boots[0]["document_name"] == "(unknown)"
    assert boots[0]["boots"][0]["load_error"] is False


def test_boot_log_is_bounded():
    for _ in range(250):
        bridge.record_boot({"documentUrl": DOC_A})
    assert len(bridge._BOOT_LOG) == 200


def test_recent_boots_groups_by_basename_and_windows():
    boot(DOC_A, 100, boot_seq=1, previous_events=[{"kind": "pagehide"}, {"kind": "boot"}])
    boot(DOC_A, 40, boot_seq=2, load_error="boom")
    boot(DOC_B, 10, boot_seq=7)
    boot(DOC_A, 1000, boot_seq=0)  # outside the 300 s window
    groups = {g["document_name"]: g for g in bridge.recent_boots()}
    assert set(groups) == {"Row01.docx", "Scratch.docx"}
    row = groups["Row01.docx"]
    assert row["count"] == 2
    assert [b["boot_seq"] for b in row["boots"]] == [1, 2]
    assert [b["load_error"] for b in row["boots"]] == [False, True]
    assert row["previous_event_kinds"] == ["boot", "pagehide"]
    assert row["avg_interval_s"] == pytest.approx(60, abs=1)
    assert groups["Scratch.docx"]["avg_interval_s"] is None
    wide = {g["document_name"]: g for g in bridge.recent_boots(window_s=5000)}
    assert wide["Row01.docx"]["count"] == 3


def fake_registry(**extra):
    return SimpleNamespace(list=list, collisions=list, **extra)


@pytest.fixture
def local_bridge(monkeypatch):
    monkeypatch.setattr(bridge, "start_in_background", lambda **kw: fake_registry())
    monkeypatch.setattr(bridge, "current_ports", lambda: None)


def test_live_status_has_recent_boots_and_no_warning_by_default(local_bridge):
    status = server.execute_live_status()
    assert status["recent_boots"] == []
    assert "warnings" not in status


def test_pane_reload_loop_warning_at_three_boots_within_120s(local_bridge):
    boot(DOC_A, 100, previous_events=[{"kind": "pagehide"}])
    boot(DOC_A, 50, previous_events=[{"kind": "beforeunload"}])
    boot(DOC_A, 0)
    status = server.execute_live_status()
    (warning,) = status["warnings"]
    assert warning["code"] == "pane_reload_loop"
    assert warning["boots"] == 3
    assert warning["avg_interval_s"] == pytest.approx(50, abs=1)
    assert warning["previous_event_kinds"] == ["beforeunload", "pagehide"]
    for needle in ("Row01.docx", "3 times", "50", "pagehide", "latest.json", "local copy", "#84"):
        assert needle in warning["message"], needle
    assert status["recent_boots"][0]["count"] == 3


def test_no_warning_for_two_boots_or_wide_spread(local_bridge):
    boot(DOC_A, 10)
    boot(DOC_A, 0)
    assert "warnings" not in server.execute_live_status()
    boot(DOC_A, 200)  # three boots overall, but only two inside 120 s
    assert "warnings" not in server.execute_live_status()


def test_boots_for_different_documents_do_not_combine(local_bridge):
    boot(DOC_A, 10)
    boot(DOC_B, 5)
    boot(DOC_A, 0)
    assert "warnings" not in server.execute_live_status()


def test_broker_status_carries_recent_boots():
    boot(DOC_A, 5, boot_seq=4)
    broker = shared.Broker(fake_registry(get=lambda name: None))
    client = broker.dispatch({"action": "attach"})["client_id"]
    status = broker.dispatch({"action": "status", "client_id": client})
    assert status["owner_pid"] == os.getpid()
    assert status["recent_boots"][0]["document_name"] == "Row01.docx"
    assert status["recent_boots"][0]["boots"][0]["boot_seq"] == 4


def test_live_status_in_shared_mode_uses_the_owners_boots(monkeypatch):
    owner_boots = [
        {
            "document_name": "Row01.docx",
            "count": 3,
            "avg_interval_s": 10.0,
            "previous_event_kinds": ["pagehide"],
            "boots": [{"age_s": a, "previous_event_kinds": ["pagehide"]} for a in (20, 10, 0)],
        }
    ]
    registry = fake_registry(status=lambda: {"owner_pid": os.getpid() + 1, "recent_boots": owner_boots})
    monkeypatch.setattr(bridge, "start_in_background", lambda **kw: registry)
    monkeypatch.setattr(bridge, "current_ports", lambda: None)
    status = server.execute_live_status()  # this process's own (empty) log is ignored
    assert status["mode"] == "shared"
    assert status["recent_boots"] == owner_boots
    assert status["warnings"][0]["code"] == "pane_reload_loop"


def test_old_owner_without_recent_boots_still_works(monkeypatch):
    registry = fake_registry(status=lambda: {"owner_pid": os.getpid()})
    monkeypatch.setattr(bridge, "start_in_background", lambda **kw: registry)
    monkeypatch.setattr(bridge, "current_ports", lambda: None)
    status = server.execute_live_status()
    assert status["recent_boots"] == []
    assert "warnings" not in status
