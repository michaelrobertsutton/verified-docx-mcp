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

from verified_docx_mcp import mutations, paths, projection, tables, text_edit
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

    def _mock_registry(self, *, has_session: bool, document_url: str | None = None):
        """*document_url* defaults to ``str(self.target)`` -- i.e. the
        mocked session names the SAME document the caller is targeting,
        so issue #154 WP-1b's ``_check_session_identity`` passes by
        default in every existing test here. Tests for the mismatch case
        itself pass an explicit *document_url* naming a different file."""
        registry = mock.Mock()
        if not has_session:
            registry.get.return_value = None
            return registry
        session = mock.Mock(spec=live_session.LiveSession)
        session.document_name = self.target.name
        session.document_url = document_url if document_url is not None else str(self.target)
        registry.get.return_value = session
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

    def test_auto_with_session_is_live_regardless_of_owner_file(self) -> None:
        """issue #154: a connected pane session is sufficient on its own
        -- the owner-file signal is no longer consulted at all, so this
        stays "live" whether or not lock_status would report one."""
        with mock.patch(
            "verified_docx_mcp.live.bridge.current_registry",
            return_value=self._mock_registry(has_session=True),
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "live")

    def test_auto_with_session_is_live_and_skips_lock_status(self) -> None:
        """issue #154 regression pin: this is the exact Mac + SharePoint
        shape of the original bug -- a pane session connected, no Word
        owner file (Word for Mac + a SharePoint/OneDrive sync never
        writes one). "auto" must resolve to "live" and must never even
        call lock_status to check -- the owner-file signal is gone from
        this path entirely, not merely satisfied."""
        with (
            mock.patch(
                "verified_docx_mcp.live.bridge.current_registry",
                return_value=self._mock_registry(has_session=True),
            ),
            mock.patch("verified_docx_mcp.server.execute_lock_status") as mock_lock,
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "live")
        mock_lock.assert_not_called()

    def test_auto_with_session_mismatched_local_document_url_raises_live_session_mismatch(self) -> None:
        """issue #154 WP-1b: a session sharing this document's basename
        but naming a DIFFERENT local file via its own document_url must
        refuse rather than silently route the write into the wrong
        document."""
        other = Path(self._tmp.name) / "other-dir"
        other.mkdir()
        other_file = other / self.target.name  # same basename, different path
        other_file.write_bytes(b"unused")
        with (
            mock.patch(
                "verified_docx_mcp.live.bridge.current_registry",
                return_value=self._mock_registry(has_session=True, document_url=str(other_file)),
            ),
            self.assertRaises(VerifyError) as ctx,
        ):
            write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_MISMATCH)

    def test_auto_with_session_web_document_url_falls_back_to_basename_match(self) -> None:
        """issue #154 WP-1b: a session whose document_url is a
        SharePoint/OneDrive web URL (not a local path -- the actual
        incident's shape) carries no local path to compare against, so
        identity falls back to the basename match the registry lookup
        already performed, and "auto" still resolves to "live"."""
        with mock.patch(
            "verified_docx_mcp.live.bridge.current_registry",
            return_value=self._mock_registry(
                has_session=True,
                document_url=f"https://contoso.sharepoint.com/sites/Team/Shared%20Documents/{self.target.name}",
            ),
        ):
            mode = write_mode.resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "live")

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

    async def test_write_mode_file_parity_with_no_pane_connected(self) -> None:
        """No pane connected at all -- explicit write_mode="file" (and
        the guard's new live-session check, which finds nothing) takes
        the exact pre-#154 file-mode path: same rung/evidence shape as
        test_text_edit.py's own ReplaceTextWholeRunTests, no verified_via/
        document_name keys (those are live-only), and write_mode == "file"
        (issue #154 WP-3 -- the one new key this evidence carries)."""
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
        self.assertEqual(evidence["write_mode"], "file")
        self.assertNotIn("verified_via", evidence)

    async def test_write_mode_file_refuses_under_connected_pane(self) -> None:
        """issue #154 WP-2: a pane IS connected for this document, so an
        explicit write_mode="file" must now refuse with
        LIVE_SESSION_ACTIVE rather than write to disk underneath Word's
        own open, autosaving buffer -- no escape hatch, per the ruling
        this session made on the issue."""
        doc = FakeDocument(text="unused -- write_mode='file' must refuse before ever reading this")
        await self.connect_pane(doc)

        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_replace_text(
                str(self.target),
                "The quick brown fox jumps over the lazy dog.",
                "Done.",
                1,
                write_mode="file",
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_ACTIVE)
        self.assertEqual(doc.text, "unused -- write_mode='file' must refuse before ever reading this")

    async def test_auto_with_session_sends_live_ops_and_skips_lock_status(self) -> None:
        """issue #154 regression pin, end to end: a pane is connected, no
        Word owner file exists (Word for Mac + a SharePoint/OneDrive sync
        never writes one) -- "auto" must still route to the pane, send it
        the op, and report write_mode == "live". Pre-fix this sent no ops
        at all and silently wrote to disk instead."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        pane = await self.connect_pane(doc)

        ops_received: list[str] = []
        original_handle_op = pane._handle_op

        def _tracking_handle_op(request_id, op, payload):
            ops_received.append(op)
            return original_handle_op(request_id, op, payload)

        pane._handle_op = _tracking_handle_op  # instance-level monkeypatch

        # write_mode left at its default ("auto"); no ~$ owner file exists
        # next to self.target -- issue #154: that must no longer matter.
        evidence = await asyncio.to_thread(
            text_edit.execute_replace_text, str(self.target), "The quick brown fox jumps over the lazy dog.", "Done.", 1
        )

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertIn("replace", ops_received)
        self.assertEqual(doc.text, "Done.")


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
        # issue #154 WP-3: live_save is a LIVE operation -- its write_mode
        # key must read "live", never "file" (it doesn't go through
        # atomic_replace_docx_parts at all).
        self.assertEqual(evidence["write_mode"], "live")
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


# ---------------------------------------------------------------------------
# LIVE_SESSION_ACTIVE guard (issue #154 WP-2): a connected pane must refuse
# every FILE-MODE mutating write, not just the ones with a write_mode
# parameter -- this is what actually protects apply_style (the tool the
# original incident hit; it has no write_mode kwarg at all) and every other
# write_mode-less tool that only reaches file mode through
# mutations._guard_before_write.
# ---------------------------------------------------------------------------


class LiveSessionActiveGuardTests(LiveWriteBridgeTestCase):
    """fixture_name stays "frag.docx" (the class default) for the
    apply_style case -- the literal #154 repro."""

    async def test_apply_style_refuses_under_connected_pane(self) -> None:
        """The literal #154 repro: apply_style has no write_mode
        parameter at all, so it only ever reached file mode through
        mutations._guard_before_write. Pre-fix this returned
        applied: true and lost the edit to Word's next autosave."""
        doc = FakeDocument(text="unused -- apply_style must refuse before ever reading this")
        await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_apply_style(str(self.target), "brown", "IntenseEmphasis", 1)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_ACTIVE)
        self.assertEqual(self.target.read_bytes(), before_bytes)  # file bytes unchanged

    async def test_apply_style_succeeds_with_no_pane_connected(self) -> None:
        """The guard is inert with no session connected -- apply_style
        takes its ordinary file-mode path exactly as before #154."""
        evidence = text_edit.execute_apply_style(str(self.target), "brown", "IntenseEmphasis", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "file")


class LiveSessionActiveMarkdownRungGuardTests(LiveWriteBridgeTestCase):
    """A rung-3 markdown tool (no write_mode parameter at all) also
    refuses under a connected pane -- proving the guard's single
    chokepoint (mutations._guard_before_write) reaches beyond
    text_edit.py's own tools. One test does not prove all 14 call sites;
    this is one representative site, not exhaustive coverage."""

    fixture_name = "sections.docx"

    async def test_replace_range_markdown_refuses_under_connected_pane(self) -> None:
        doc = FakeDocument(text="unused -- replace_range_markdown never talks to the pane")
        await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        with self.assertRaises(VerifyError) as ctx:
            mutations.execute_replace_range_markdown(str(self.target), "background-1", "## Background\n\nX.\n")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_ACTIVE)
        self.assertEqual(self.target.read_bytes(), before_bytes)


class LiveSessionActiveTableRungGuardTests(LiveWriteBridgeTestCase):
    """A table tool (also no write_mode parameter) refuses under a
    connected pane -- a second representative site distinct from
    text_edit.py/mutations.py's own module."""

    fixture_name = "tables.docx"

    async def test_replace_table_row_refuses_under_connected_pane(self) -> None:
        doc = FakeDocument(text="unused -- replace_table_row never talks to the pane")
        await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        with self.assertRaises(VerifyError) as ctx:
            tables.execute_replace_table_row(str(self.target), 1, 1, ["New A", "New B"])
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_ACTIVE)
        self.assertEqual(self.target.read_bytes(), before_bytes)


class NoBridgeGuardTests(unittest.TestCase):
    """issue #154 WP-2: an ordinary file-mode write, with NO bridge
    running at all (the overwhelming common case), must never start one
    as a side effect of checking for a live session -- the guard's whole
    point is a cheap registry lookup, never an attempt to bind the
    bridge's real ports (53135/53136 in production). This is exactly
    the property the ephemeral-port bridge fixture (LiveWriteBridgeTestCase)
    CANNOT prove, since it always starts a real bridge in asyncSetUp."""

    fixture_name = "frag.docx"

    def setUp(self) -> None:
        bridge.stop()  # defensive: no bridge should be running between tests
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / self.fixture_name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self) -> None:
        bridge.stop()
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()

    def test_file_mode_write_never_starts_the_bridge(self) -> None:
        with mock.patch.object(bridge, "start_in_background") as mock_start:
            evidence = text_edit.execute_apply_style(str(self.target), "brown", "IntenseEmphasis", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "file")
        mock_start.assert_not_called()


if __name__ == "__main__":
    unittest.main()
