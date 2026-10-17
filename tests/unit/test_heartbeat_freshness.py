"""Negotiated liveness never refreshes a sampled digest or a client baseline."""

import asyncio
import subprocess
from pathlib import Path
from unittest.mock import Mock

import pytest

from verified_docx_mcp.live.protocol import HeartbeatMessage, HelloMessage, ProtocolError
from verified_docx_mcp.live.session import LiveSession


def test_actual_pane_heartbeat_and_preflight():
    root = Path(__file__).resolve().parents[2]
    subprocess.run(
        [
            "node",
            str(root / "tests/unit/js/heartbeat_preflight_harness.mjs"),
            str(root / "addin/taskpane.js"),
        ],
        check=True,
        capture_output=True,
        text=True,
    )


def test_liveness_does_not_refresh_digest_age():
    session = LiveSession(
        "fixture.docx",
        "/fixture.docx",
        HelloMessage("/fixture.docx", "Word", "Mac", {}, "old"),
        Mock(),
        asyncio.new_event_loop(),
    )
    try:
        sampled = session.body_hash_sampled_monotonic
        heartbeat = session.last_heartbeat_monotonic
        session.touch_heartbeat()
        assert session.last_heartbeat_monotonic > heartbeat
        assert session.body_hash_sampled_monotonic == sampled
        assert session.last_body_sha256 == "old"
        assert session.body_hash_state == "retained"
        session.touch_heartbeat("new")
        assert session.body_hash_sampled_monotonic > sampled
        assert session.body_hash_state == "sampled"
    finally:
        session.loop.close()


@pytest.mark.parametrize("state,digest", [("unavailable", None), ("sampled", "abc")])
def test_new_heartbeat_roundtrip(state, digest):
    hb = HeartbeatMessage("/fixture.docx", digest, state)
    assert HeartbeatMessage.from_json(hb.to_json()) == hb


@pytest.mark.parametrize(
    "fields",
    [
        {"digestState": "fresh", "bodySha256": "abc"},
        {"digestState": "unavailable", "bodySha256": "abc"},
        {"digestState": "sampled"},
        {},
    ],
)
def test_ambiguous_hash_refused(fields):
    with pytest.raises(ProtocolError):
        HeartbeatMessage.from_json(dict(type="heartbeat", documentUrl="/fixture.docx", **fields))


def test_old_broker_has_unknown_sample_age():
    from verified_docx_mcp.live.shared import RemoteSession

    hello = HelloMessage("/fixture.docx", "Word", "Mac", {}, "old")
    session = RemoteSession(
        Mock(),
        {
            "hello": hello.to_json(),
            "document_name": "fixture.docx",
            "connected_since": 1,
            "last_body_sha256": "old",
            "heartbeat_age": 0,
        },
    )
    assert session.body_hash_state == "retained"
    assert session.body_hash_sampled_monotonic is None
