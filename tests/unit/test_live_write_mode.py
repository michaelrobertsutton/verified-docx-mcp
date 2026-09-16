"""Unit tests for src/verified_docx_mcp/live/write_mode.py and the
write_mode="live" path on replace_text/format_text/live_save (issue #106
WP-3: https://github.com/michaelrobertsutton/JennyStack/issues/106).

Two styles of test here:

  - ResolverTruthTableTests: pure unit tests of
    write_mode.resolve_write_mode's file/live/auto selection rule against
    a mocked current_registry()/execute_lock_status() -- no bridge, no
    sockets, no fixture file mutation.
  - LiveWriteBridgeTestCase (subclasses the same pattern
    test_live_session.py's own LiveBridgeTestCase uses): a real bridge on
    EPHEMERAL ports with a connected tests/unit/fake_pane.py FakePane,
    exercising replace_text/format_text/live_save end to end, plus
    write_mode="file"/"auto" parity checks against the pre-WP-3 file
    path.

Hard rule respected throughout: every bridge started here binds
EPHEMERAL ports (host="127.0.0.1", port=0, ops_port=0) -- never the real
default ports 53135/53136 (a real lead's live bridge may be running on
those on this machine; see this WP's own task description).
"""

from __future__ import annotations

import asyncio
import os
import shutil
import ssl
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, FakePane

from verified_docx_mcp import mutations, paths, projection, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import bridge, write_mode
from verified_docx_mcp.live import session as live_session

FIXTURES = REPO / "tests" / "fixtures"

# Same rationale as test_text_edit.py's own module-level tweak: shrink the
# file-mode write guard's sync-quiesce sample interval so a write_mode=
# "file"/"auto"-resolved-to-file call in this suite does not pay the
# (default 1.5s) real-world quiesce wait.
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _client_ssl_context() -> ssl.SSLContext:
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


# ---------------------------------------------------------------------------
# Resolver truth table -- no bridge, no sockets.
# ---------------------------------------------------------------------------


class ResolverTruthTableTests(unittest.TestCase):
    """resolve_write_mode's file/live/auto x owner-file-present/absent x
    session-present/absent truth table, against mocked
    live_bridge.current_registry()/server.execute_lock_status()."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "frag.docx"
        shutil.copyfile(FIXTURES / "frag.docx", self.target)

    def tearDown(self) -> None:
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def _owner_status(self, *, word_present: bool) -> dict:
        return {
            "path": str(self.target),
            "owner_file": {
                "present": word_present,
                "path": str(self.target.parent / "~$ag.docx") if word_present else None,
                "format": "word" if word_present else None,
                "owner_name": "Someone" if word_present else None,
            },
            "sync_quiesced": True,
            "sync_detail": {},
        }

    def _mock_registry(self, *, has_session: bool):
        registry = mock.Mock()
        registry.get.return_value = mock.Mock(spec=live_session.LiveSession) if has_session else None
        return registry

    def test_write_mode_file_is_always_file_with_no_lookups(self) -> None:
        with (
            mock.patch("verified_docx_mcp.live.bridge.current_registry") as mock_registry,
            mock.patch("verified_docx_mcp.server.execute_lock_status") as mock_lock,
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "file")
        self.assertEqual(mode, "file")
        mock_registry.assert_not_called()
        mock_lock.assert_not_called()

    def test_write_mode_live_with_session_is_live(self) -> None:
        with mock.patch(
            "verified_docx_mcp.live.bridge.current_registry",
            return_value=self._mock_registry(has_session=True),
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "live")
        self.assertEqual(mode, "live")

    def test_write_mode_live_without_session_raises_live_unavailable(self) -> None:
        with (
            mock.patch("verified_docx_mcp.live.bridge.current_registry", return_value=None),
            self.assertRaises(VerifyError) as ctx,
        ):
            write_mode.resolve_write_mode(str(self.target), "live")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    def test_auto_with_owner_and_session_is_live(self) -> None:
        with (
            mock.patch(
                "verified_docx_mcp.live.bridge.current_registry",
                return_value=self._mock_registry(has_session=True),
            ),
            mock.patch(
                "verified_docx_mcp.server.execute_lock_status",
                return_value=self._owner_status(word_present=True),
            ),
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "live")

    def test_auto_with_session_but_no_owner_is_file(self) -> None:
        with (
            mock.patch(
                "verified_docx_mcp.live.bridge.current_registry",
                return_value=self._mock_registry(has_session=True),
            ),
            mock.patch(
                "verified_docx_mcp.server.execute_lock_status",
                return_value=self._owner_status(word_present=False),
            ) as mock_lock,
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "file")
        mock_lock.assert_called_once()

    def test_auto_with_owner_but_no_session_is_file_and_skips_lock_status(self) -> None:
        """No session -> "auto" never even bothers checking the owner
        file (short-circuits before it), a deliberate perf property for
        the overwhelmingly common "nothing live is happening" call."""
        with (
            mock.patch(
                "verified_docx_mcp.live.bridge.current_registry",
                return_value=self._mock_registry(has_session=False),
            ),
            mock.patch("verified_docx_mcp.server.execute_lock_status") as mock_lock,
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "file")
        mock_lock.assert_not_called()

    def test_auto_with_neither_owner_nor_session_is_file(self) -> None:
        with (
            mock.patch("verified_docx_mcp.live.bridge.current_registry", return_value=None),
            mock.patch("verified_docx_mcp.server.execute_lock_status") as mock_lock,
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "file")
        mock_lock.assert_not_called()

    def test_invalid_write_mode_raises_invalid_input(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            write_mode.resolve_write_mode(str(self.target), "sometimes")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


# ---------------------------------------------------------------------------
# End-to-end over a real (ephemeral-port) bridge + fake pane.
# ---------------------------------------------------------------------------


class LiveWriteBridgeTestCase(unittest.IsolatedAsyncioTestCase):
    heartbeat_interval = 5.0
    missed_heartbeats = 3
    fixture_name = "frag.docx"

    async def asyncSetUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(tmp)
        self.target = tmp / self.fixture_name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

        bridge.stop()  # defensive: no bridge should be running between tests
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
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    async def connect_pane(self, document: FakeDocument, **kwargs) -> FakePane:
        pane = FakePane(document=document, document_url=str(self.target), **kwargs)
        await pane.connect(f"wss://127.0.0.1:{self.ops_port}/ops", ssl_context=self.ssl_context)
        self._panes.append(pane)
        await asyncio.sleep(0.1)  # let the bridge's receive loop process 'hello' and register
        return pane

    def _write_owner_file(self) -> Path:
        """Simulate a desktop Word owner file for self.target (>8-char
        base name: first TWO characters dropped, per
        _word_owner_file_matches's own documented rule -- "frag.docx" ->
        "~$ag.docx")."""
        owner = self.target.parent / "~$ag.docx"
        owner.write_bytes(b"\x00\x00Someone")
        return owner


class ReplaceTextLiveTests(LiveWriteBridgeTestCase):
    async def test_happy_path_full_evidence(self) -> None:
        doc = FakeDocument(text="alpha beta gamma")
        await self.connect_pane(doc)

        # execute_replace_text is a plain blocking call (LiveSession.
        # request_threadsafe internally) -- run it off this test's own
        # event loop via asyncio.to_thread, or the FakePane's own
        # recv-loop task (scheduled on THIS SAME loop by connect_pane
        # above) would never get a chance to run and reply, deadlocking
        # until the op's own timeout.
        evidence = await asyncio.to_thread(
            text_edit.execute_replace_text, str(self.target), "beta", "OMEGA", 1, write_mode="live"
        )

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(evidence["rung"], 2)
        self.assertEqual(evidence["before"], "beta")
        self.assertEqual(evidence["after"], "OMEGA")
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["verified_via"], "word-addin")
        self.assertEqual(evidence["document_name"], self.target.name)
        self.assertEqual(evidence["track_changes_author"], "word-signed-in-user")
        self.assertEqual(evidence["orphaned_comment_ids"], [])
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertTrue(evidence["revision_after"].startswith("live:sha256:"))
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertNotIn("conflict_copy_detected", evidence)
        self.assertIn("audit_logged", evidence)
        self.assertEqual(doc.text, "alpha OMEGA gamma")

    async def test_match_count_mismatch(self) -> None:
        doc = FakeDocument(text="alpha alpha")
        await self.connect_pane(doc)

        with self.assertRaises(VerifyError) as ctx:
            await asyncio.to_thread(
                text_edit.execute_replace_text, str(self.target), "alpha", "x", 1, write_mode="live"
            )

        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.MATCH_COUNT_MISMATCH)
        self.assertEqual(doc.text, "alpha alpha")  # a refusal must never touch the document

    async def test_zero_match(self) -> None:
        doc = FakeDocument(text="alpha")
        await self.connect_pane(doc)

        with self.assertRaises(VerifyError) as ctx:
            await asyncio.to_thread(
                text_edit.execute_replace_text, str(self.target), "zzz", "x", 1, write_mode="live"
            )

        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.ZERO_MATCH)
        self.assertEqual(doc.text, "alpha")

    async def test_live_stale_on_stale_revision_before(self) -> None:
        doc = FakeDocument(text="alpha")
        await self.connect_pane(doc)

        with self.assertRaises(VerifyError) as ctx:
            await asyncio.to_thread(
                text_edit.execute_replace_text,
                str(self.target),
                "alpha",
                "beta",
                1,
                write_mode="live",
                revision_before="live:sha256:deadbeef",
            )

        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_STALE)
        # refused BEFORE the op was ever sent -- the document is untouched.
        self.assertEqual(doc.text, "alpha")

    async def test_live_unavailable_when_write_mode_live_has_no_session(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_replace_text(str(self.target), "alpha", "beta", 1, write_mode="live")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    async def test_pane_disconnect_mid_op_raises_live_disconnected(self) -> None:
        doc = FakeDocument(text="alpha")
        pane = await self.connect_pane(doc)
        pane.drop_ops.add("replace")  # 'describe' still answers; 'replace' never does

        async def close_soon() -> None:
            await asyncio.sleep(0.1)
            await pane.close()

        closer = asyncio.create_task(close_soon())
        try:
            with self.assertRaises(VerifyError) as ctx:
                # execute_replace_text is a plain blocking call (it uses
                # LiveSession.request_threadsafe internally) -- run it off
                # this test's own event loop so 'closer' above can still
                # run concurrently and actually close the pane's socket.
                await asyncio.to_thread(
                    text_edit.execute_replace_text,
                    str(self.target),
                    "alpha",
                    "beta",
                    1,
                    write_mode="live",
                )
        finally:
            await closer
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_DISCONNECTED)

    async def test_write_mode_file_takes_todays_path_byte_for_byte(self) -> None:
        """A pane is connected (so "live" would be available), but an
        explicit write_mode="file" must still take the exact pre-WP-3
        file-mode path -- same rung/evidence shape as
        test_text_edit.py's own ReplaceTextWholeRunTests, and no
        write_mode/verified_via/document_name keys (those are live-only)."""
        doc = FakeDocument(text="unused -- write_mode='file' never talks to the pane")
        await self.connect_pane(doc)

        evidence = text_edit.execute_replace_text(
            str(self.target),
            "The quick brown fox jumps over the lazy dog.",
            "Done.",
            1,
            write_mode="file",
        )

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], "exact")
        self.assertEqual(evidence["match_count"], 1)
        self.assertIn("conflict_copy_detected", evidence)
        self.assertNotIn("write_mode", evidence)
        self.assertNotIn("verified_via", evidence)
        self.assertEqual(doc.text, "unused -- write_mode='file' never talks to the pane")

    async def test_auto_with_session_but_no_owner_file_uses_file_path_and_sends_no_ops(self) -> None:
        doc = FakeDocument(text="unused")
        pane = await self.connect_pane(doc)

        ops_received: list[str] = []
        original_handle_op = pane._handle_op

        def _tracking_handle_op(request_id, op, payload):
            ops_received.append(op)
            return original_handle_op(request_id, op, payload)

        pane._handle_op = _tracking_handle_op  # instance-level monkeypatch

        # write_mode left at its default ("auto"); no ~$ owner file exists
        # next to self.target, so this must resolve to "file" even though
        # a pane session is connected for this document name.
        evidence = text_edit.execute_replace_text(
            str(self.target), "The quick brown fox jumps over the lazy dog.", "Done.", 1
        )

        self.assertTrue(evidence["applied"])
        self.assertNotIn("write_mode", evidence)
        self.assertEqual(ops_received, [])  # the fake pane never received any op
        self.assertEqual(doc.text, "unused")

    async def test_auto_with_session_and_owner_file_goes_live(self) -> None:
        doc = FakeDocument(text="alpha beta")
        await self.connect_pane(doc)
        owner = self._write_owner_file()
        try:
            evidence = await asyncio.to_thread(
                text_edit.execute_replace_text, str(self.target), "beta", "OMEGA", 1
            )
        finally:
            owner.unlink(missing_ok=True)

        self.assertEqual(evidence.get("write_mode"), "live")
        self.assertEqual(doc.text, "alpha OMEGA")


class FormatTextLiveTests(LiveWriteBridgeTestCase):
    async def test_happy_path(self) -> None:
        doc = FakeDocument(text="alpha beta")
        await self.connect_pane(doc)

        evidence = await asyncio.to_thread(
            text_edit.execute_format_text, str(self.target), "beta", {"bold": True}, 1, write_mode="live"
        )

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(evidence["rung"], 1)
        self.assertEqual(evidence["before"], "beta")
        self.assertEqual(evidence["after"], "beta")  # format never changes content
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["verified_via"], "word-addin")
        self.assertEqual(evidence["document_name"], self.target.name)
        # format never mutates body.text -- pre/post hash equal is EXPECTED here
        # (unlike replace_text, where that would be a verification failure).
        self.assertEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertEqual(doc.text, "alpha beta")


class LiveSaveTests(LiveWriteBridgeTestCase):
    async def test_returns_both_revisions(self) -> None:
        doc = FakeDocument(text="alpha")
        doc.saved = False
        await self.connect_pane(doc)

        evidence = await asyncio.to_thread(text_edit.execute_live_save, str(self.target))

        self.assertTrue(evidence["applied"])
        self.assertTrue(evidence["saved"])
        self.assertTrue(doc.saved)
        self.assertEqual(evidence["document_name"], self.target.name)
        self.assertTrue(evidence["revision_after"].startswith("live:sha256:"))
        self.assertIn("file_revision", evidence)
        # file_revision is the plain file-mode revision TOKEN STRING
        # (projection.compute_revision(...)["token"]) -- the same shape
        # replace_text/format_text/etc. use as revision_before/
        # revision_after in file mode -- not a nested {"token": ...} dict.
        self.assertIsInstance(evidence["file_revision"], str)
        self.assertEqual(
            evidence["file_revision"],
            projection.compute_revision(self.target)["token"],
        )

    async def test_live_unavailable_with_no_session(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_live_save(str(self.target))
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)


if __name__ == "__main__":
    unittest.main()
