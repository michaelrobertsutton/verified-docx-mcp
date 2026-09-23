"""Unit tests for live/session.py + the WSS wiring it depends on in
live/bridge.py (issue #106 WP-2:
https://github.com/michaelrobertsutton/JennyStack/issues/106).

Exercised end to end through `tests/unit/fake_pane.py`'s `FakePane` --
a real WSS client over a real (ephemeral-port, self-signed) TLS socket,
answering ops against an in-memory document. No Word, no AppleScript, no
real pane: this is the WP-2 acceptance harness the plan calls for.

Each test gets its own bridge on ephemeral ports (`asyncSetUp`/
`asyncTearDown` below) -- `live/bridge.py`'s `start_in_background`/`stop`
are a process-wide singleton, so tests must not share one running
instance (and must not each try to bind the real default ports 53135/
53136, which could collide with a developer's own `--serve-only` run).
"""

from __future__ import annotations

import asyncio
import ssl
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, FakePane

from verified_docx_mcp.live import bridge
from verified_docx_mcp.live import session as live_session


def _client_ssl_context() -> ssl.SSLContext:
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


class LiveBridgeTestCase(unittest.IsolatedAsyncioTestCase):
    """Base class: a fresh bridge on ephemeral ports, torn down after
    each test. Subclasses may override `heartbeat_interval`/
    `missed_heartbeats` (bridge-side eviction thresholds) at the class
    level."""

    heartbeat_interval = 5.0
    missed_heartbeats = 3

    async def asyncSetUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        bridge.stop()  # defensive: no bridge should be running between tests, but don't assume it
        self.registry = bridge.start_in_background(
            host="127.0.0.1",
            port=0,
            ops_port=0,
            cert_dir=tmp / "cert",
            addin_dir=bridge.REPO_ADDIN_DIR,
            report_dir=tmp / "reports",
            heartbeat_interval=self.heartbeat_interval,
            missed_heartbeats=self.missed_heartbeats,
        )
        self.port, self.ops_port = bridge.current_ports()
        self.ssl_context = _client_ssl_context()
        self._panes: list[FakePane] = []

    async def asyncTearDown(self) -> None:
        for pane in self._panes:
            await pane.close()
        bridge.stop()
        self._tmp.cleanup()

    async def connect_pane(self, **kwargs) -> FakePane:
        pane = FakePane(**kwargs)
        await pane.connect(f"wss://127.0.0.1:{self.ops_port}/ops", ssl_context=self.ssl_context)
        self._panes.append(pane)
        await asyncio.sleep(0.1)  # let the bridge's receive loop process 'hello' and register
        return pane


class HelloRegistrationTests(LiveBridgeTestCase):
    async def test_hello_registers_session_keyed_by_document_name(self):
        await self.connect_pane(document_url="/Users/lead/Desktop/Report Final.docx")
        sessions = self.registry.list()
        self.assertEqual([s.document_name for s in sessions], ["Report Final.docx"])
        self.assertEqual(sessions[0].document_url, "/Users/lead/Desktop/Report Final.docx")

    async def test_get_returns_none_for_unregistered_document(self):
        self.assertIsNone(self.registry.get("Nope.docx"))

    async def test_two_panes_register_two_independent_sessions(self):
        await self.connect_pane(document_url="/tmp/One.docx")
        await self.connect_pane(document_url="/tmp/Two.docx")
        names = sorted(s.document_name for s in self.registry.list())
        self.assertEqual(names, ["One.docx", "Two.docx"])


class SessionCollisionTests(LiveBridgeTestCase):
    """Issue #22 B3 follow-up: SessionRegistry keys purely on basename, so
    two DIFFERENT documents sharing a file name are indistinguishable by
    name alone. Never refused (the common case -- the SAME document
    reconnecting -- must not be flagged), but recorded so it is visible."""

    async def test_same_url_reconnect_is_not_a_collision(self):
        await self.connect_pane(document_url="/tmp/Shared/Proposal.docx")
        await self.connect_pane(document_url="/tmp/Shared/Proposal.docx")  # reconnect, same doc
        self.assertEqual(self.registry.collisions(), [])

    async def test_different_url_same_basename_is_recorded_not_refused(self):
        await self.connect_pane(document_url="https://contoso.sharepoint.com/sites/A/Proposal.docx")
        await self.connect_pane(document_url="https://contoso.sharepoint.com/sites/B/Proposal.docx")

        # Not refused: the second connection still registers and takes
        # over routing for "Proposal.docx" -- today's existing behavior.
        sessions = self.registry.list()
        self.assertEqual([s.document_name for s in sessions], ["Proposal.docx"])
        self.assertEqual(
            self.registry.get("Proposal.docx").document_url,
            "https://contoso.sharepoint.com/sites/B/Proposal.docx",
        )

        collisions = self.registry.collisions()
        self.assertEqual(len(collisions), 1)
        self.assertEqual(collisions[0]["document_name"], "Proposal.docx")
        self.assertEqual(
            collisions[0]["displaced_document_url"], "https://contoso.sharepoint.com/sites/A/Proposal.docx"
        )
        self.assertEqual(collisions[0]["new_document_url"], "https://contoso.sharepoint.com/sites/B/Proposal.docx")
        self.assertIn("at", collisions[0])

    async def test_collision_log_is_capped(self):
        for i in range(25):
            await self.connect_pane(document_url=f"https://contoso.sharepoint.com/sites/{i}/Cap.docx")
        self.assertEqual(len(self.registry.collisions()), 20)


class OpRoundTripTests(LiveBridgeTestCase):
    async def test_describe_round_trips(self):
        doc = FakeDocument(text="Hello world.")
        await self.connect_pane(document=doc, document_url="/tmp/D.docx")
        result = await self.registry.request("D.docx", "describe", {})
        self.assertEqual(result["bodySha256"], doc.sha256())
        self.assertEqual(result["changeTrackingMode"], "Off")
        self.assertTrue(result["saved"])

    async def test_search_round_trips(self):
        doc = FakeDocument(text="alpha beta alpha")
        await self.connect_pane(document=doc, document_url="/tmp/S.docx")
        result = await self.registry.request("S.docx", "search", {"find": "alpha"})
        self.assertEqual(len(result["matches"]), 2)
        self.assertEqual(result["matches"][0]["text"], "alpha")
        self.assertEqual(result["matches"][1]["contextBefore"], "alpha beta ")

    async def test_replace_round_trips_and_mutates_document(self):
        doc = FakeDocument(text="alpha beta")
        await self.connect_pane(document=doc, document_url="/tmp/R.docx")
        pre_hash = doc.sha256()
        result = await self.registry.request(
            "R.docx", "replace", {"find": "alpha", "expected_matches": 1, "replace": "gamma"}
        )
        self.assertTrue(result["applied"])
        self.assertEqual(result["match_count"], 1)
        self.assertEqual(result["matches"], [{"before": "alpha", "after": "gamma"}])
        self.assertEqual(doc.text, "gamma beta")
        self.assertEqual(result["pre"], pre_hash)
        self.assertEqual(result["post"], doc.sha256())
        self.assertNotEqual(result["pre"], result["post"])

    async def test_format_round_trips_without_mutating_text(self):
        doc = FakeDocument(text="alpha beta")
        await self.connect_pane(document=doc, document_url="/tmp/F.docx")
        result = await self.registry.request(
            "F.docx", "format", {"find": "alpha", "expected_matches": 1, "bold": True}
        )
        self.assertTrue(result["applied"])
        self.assertEqual(doc.text, "alpha beta")
        self.assertEqual(result["pre"], result["post"])

    async def test_comment_add_reply_resolve_round_trip(self):
        doc = FakeDocument(text="alpha beta")
        await self.connect_pane(document=doc, document_url="/tmp/C.docx")

        added = await self.registry.request(
            "C.docx", "comment_add", {"find": "alpha", "expected_matches": 1, "text": "note"}
        )
        comment_id = added["comment_id"]
        self.assertEqual(added["pre"], added["post"])  # body text unchanged by a comment insert

        listed = await self.registry.request("C.docx", "comments_list", {})
        self.assertEqual(len(listed["comments"]), 1)
        self.assertEqual(listed["comments"][0]["id"], comment_id)
        self.assertEqual(listed["comments"][0]["anchorText"], "alpha")
        self.assertEqual(listed["comments"][0]["replies"], [])

        replied = await self.registry.request(
            "C.docx", "comment_reply", {"comment_id": comment_id, "text": "reply text"}
        )
        self.assertIn("reply_id", replied)

        listed_again = await self.registry.request("C.docx", "comments_list", {})
        self.assertEqual(len(listed_again["comments"][0]["replies"]), 1)
        self.assertEqual(listed_again["comments"][0]["replies"][0]["content"], "reply text")

        resolved = await self.registry.request(
            "C.docx", "comment_resolve", {"comment_id": comment_id, "resolved": True}
        )
        self.assertTrue(resolved["resolved"])

    async def test_save_round_trips(self):
        doc = FakeDocument(text="x")
        doc.saved = False
        await self.connect_pane(document=doc, document_url="/tmp/Sv.docx")
        result = await self.registry.request("Sv.docx", "save", {})
        self.assertTrue(result["saved"])
        self.assertTrue(doc.saved)


class ErrorPathTests(LiveBridgeTestCase):
    async def test_no_session_raises_live_unavailable(self):
        with self.assertRaises(live_session.LiveUnavailable):
            await self.registry.request("Nope.docx", "ping", {}, timeout=1)

    async def test_request_threadsafe_no_session_raises_live_unavailable(self):
        # Sync entry point (the one a non-async FastMCP tool would call) --
        # run off the test's own event loop via a worker thread.
        with self.assertRaises(live_session.LiveUnavailable):
            await asyncio.to_thread(
                self.registry.request_threadsafe, "Nope.docx", "ping", {}, timeout=1
            )

    async def test_expected_matches_mismatch_raises_live_op_failed_with_pane_message(self):
        doc = FakeDocument(text="alpha")
        await self.connect_pane(document=doc, document_url="/tmp/M.docx")
        with self.assertRaises(live_session.LiveOpFailed) as ctx:
            await self.registry.request(
                "M.docx", "replace", {"find": "zzz", "expected_matches": 1, "replace": "y"}
            )
        self.assertEqual(ctx.exception.code, "LIVE_OP_FAILED")
        self.assertIn("found 0", ctx.exception.message)
        self.assertEqual(doc.text, "alpha")  # a refusal must never touch the document

    async def test_disconnect_by_timeout_raises_live_disconnected(self):
        doc = FakeDocument(text="alpha")
        pane = await self.connect_pane(document=doc, document_url="/tmp/Dc.docx")
        pane.drop_ops.add("ping")
        with self.assertRaises(live_session.LiveDisconnected):
            await self.registry.request("Dc.docx", "ping", {}, timeout=0.3)

    async def test_socket_close_mid_request_raises_live_disconnected(self):
        doc = FakeDocument(text="alpha")
        pane = await self.connect_pane(document=doc, document_url="/tmp/Sc.docx")
        pane.drop_ops.add("ping")

        async def close_soon() -> None:
            await asyncio.sleep(0.1)
            await pane.close()

        closer = asyncio.create_task(close_soon())
        try:
            with self.assertRaises(live_session.LiveDisconnected):
                await self.registry.request("Sc.docx", "ping", {}, timeout=5)
        finally:
            await closer

    async def test_stale_body_hash_raises_live_stale(self):
        doc = FakeDocument(text="alpha")
        await self.connect_pane(document=doc, document_url="/tmp/St.docx")
        actual_pre_hash = doc.sha256()
        with self.assertRaises(live_session.LiveStale) as ctx:
            await self.registry.request(
                "St.docx",
                "replace",
                {"find": "alpha", "expected_matches": 1, "replace": "beta"},
                expected_body_sha256="not-the-real-hash",
            )
        self.assertEqual(ctx.exception.expected, "not-the-real-hash")
        self.assertEqual(ctx.exception.actual, actual_pre_hash)
        # The pane does not act on expected_body_sha256 (protocol.py's own
        # docstring); staleness is a caller-side check on the reply, so
        # the op already applied by the time LiveStale is raised.
        self.assertEqual(doc.text, "beta")


class HeartbeatEvictionTests(LiveBridgeTestCase):
    heartbeat_interval = 0.1
    missed_heartbeats = 2  # stale_after = 0.2s

    async def test_session_evicted_after_missed_heartbeats(self):
        await self.connect_pane(document_url="/tmp/Hb.docx", heartbeat_interval=1000)
        self.assertIsNotNone(self.registry.get("Hb.docx"))
        await asyncio.sleep(0.4)
        self.assertIsNone(self.registry.get("Hb.docx"))
        with self.assertRaises(live_session.LiveUnavailable):
            await self.registry.request("Hb.docx", "ping", {}, timeout=1)

    async def test_regular_heartbeats_keep_session_alive(self):
        await self.connect_pane(document_url="/tmp/Hb2.docx", heartbeat_interval=0.05)
        await asyncio.sleep(0.4)  # several heartbeats land within this window
        self.assertIsNotNone(self.registry.get("Hb2.docx"))


class LiveStatusShapeTests(LiveBridgeTestCase):
    async def test_live_status_shape_before_and_after_a_pane_connects(self):
        # execute_live_status() calls start_in_background() with no
        # overrides; because asyncSetUp already started a bridge (on
        # ephemeral ports), start_in_background's idempotent guard
        # returns that SAME registry rather than trying to bind the real
        # default ports -- see bridge.start_in_background's docstring.
        from verified_docx_mcp import server as server_module

        status_before = server_module.execute_live_status()
        self.assertTrue(status_before["bridge_running"])
        self.assertEqual(status_before["port"], self.port)
        self.assertEqual(status_before["ops_port"], self.ops_port)
        self.assertEqual(status_before["sessions"], [])
        self.assertEqual(status_before["session_collisions"], [])

        doc = FakeDocument(text="hello")
        await self.connect_pane(
            document=doc, document_url="/tmp/LS.docx", requirement_sets={"1.4": True, "1.5": False}
        )

        status_after = server_module.execute_live_status()
        self.assertEqual(len(status_after["sessions"]), 1)
        entry = status_after["sessions"][0]
        self.assertEqual(entry["document_name"], "LS.docx")
        self.assertEqual(entry["document_url"], "/tmp/LS.docx")
        self.assertIn("connected_since", entry)
        self.assertGreaterEqual(entry["last_heartbeat_age_s"], 0)
        self.assertEqual(entry["body_sha256"], doc.sha256())
        self.assertEqual(entry["requirement_sets"], {"1.4": True, "1.5": False})
        self.assertEqual(status_after["session_collisions"], [])

    async def test_live_status_surfaces_session_collisions(self):
        from verified_docx_mcp import server as server_module

        await self.connect_pane(document_url="https://contoso.sharepoint.com/sites/A/Dup.docx")
        await self.connect_pane(document_url="https://contoso.sharepoint.com/sites/B/Dup.docx")

        status = server_module.execute_live_status()
        self.assertEqual(len(status["session_collisions"]), 1)
        self.assertEqual(status["session_collisions"][0]["document_name"], "Dup.docx")


if __name__ == "__main__":
    unittest.main()
