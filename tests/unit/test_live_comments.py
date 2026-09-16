"""Unit tests for src/verified_docx_mcp/live/comments_live.py (issue #106
WP-4: https://github.com/michaelrobertsutton/JennyStack/issues/106) --
write_mode="live" for add_anchored_comment/reply_to_comment/
resolve_comment, and source="live" for list_open_items, including the
correlation table that bridges a live comment (Office.js Comment.id) back
to the OOXML durableId/w:id `list_open_items`'s own file mode reports.

Exercised through `tests/unit/fake_pane.py`'s `FakePane`/`FakeDocument` --
a real WSS client over a real (ephemeral-port, self-signed) TLS socket, on
a fresh bridge per test (mirrors `test_live_session.py`'s own
`LiveBridgeTestCase`). No Word, no AppleScript, no real pane; never
connects to the real default ports 53135/53136.

Correlation fixture data: `tests/fixtures/comments/multipara-comment.docx`
(see tests/fixtures/README.md's provenance row; also used by
tests/unit/test_comments.py's own MultiParagraphCommentIdentityTests) has
two real, Word-authored comments this module's own file-mode read
(`tracked_changes._parse_comments`) already reports as
`comment_id="58A3F864"` (a two-paragraph comment anchored to "Fixture",
content "First paragraph of a two-paragraph comment.Second paragraph of
the same comment.") and `comment_id="22F779FB"` (anchored to "econd ",
content "A single-paragraph comment for contrast."). The fake pane's own
comment content for the SAME two comments is set up with a "\\r" between
the root's two paragraphs, per WP-1's real finding (docs/live-mode.md)
that the pane's own Comment.content joins multi-paragraph content that
way while file mode's own join has no separator at all.
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

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeComment, FakeDocument, FakePane

from verified_docx_mcp import comments, mutations, paths
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import bridge, comments_live

FIXTURES = REPO / "tests" / "fixtures"
MULTIPARA = FIXTURES / "comments" / "multipara-comment.docx"
FRAG = FIXTURES / "frag.docx"

ROOT_DURABLE_ID = "58A3F864"
ROOT_ANCHOR = "Fixture"
ROOT_CONTENT_LIVE = "First paragraph of a two-paragraph comment.\rSecond paragraph of the same comment."
ROOT_AUTHOR = "Michael Sutton"
ROOT_DATE = "2026-09-16T14:06:00Z"

SINGLE_DURABLE_ID = "22F779FB"
SINGLE_ANCHOR = "econd "
SINGLE_CONTENT_LIVE = "A single-paragraph comment for contrast."

mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _client_ssl_context() -> ssl.SSLContext:
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


class LiveCommentsTestCase(unittest.IsolatedAsyncioTestCase):
    """Fresh bridge on ephemeral ports + a fresh allowed-roots temp dir,
    per test -- combines test_live_session.py's LiveBridgeTestCase
    (bridge bring-up) with test_comments.py's _TempFixtureCase (a private
    copy of a fixture under a temp VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS)."""

    fixture_name = "frag.docx"

    async def asyncSetUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)

        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(tmp)
        self.target = tmp / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

        bridge.stop()  # defensive: no bridge should be running between tests
        self.registry = bridge.start_in_background(
            host="127.0.0.1",
            port=0,
            ops_port=0,
            cert_dir=tmp / "cert",
            addin_dir=bridge.REPO_ADDIN_DIR,
            report_dir=tmp / "reports",
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

    async def connect_pane(self, **kwargs) -> FakePane:
        kwargs.setdefault("document_url", str(self.target))
        pane = FakePane(**kwargs)
        await pane.connect(f"wss://127.0.0.1:{self.ops_port}/ops", ssl_context=self.ssl_context)
        self._panes.append(pane)
        await asyncio.sleep(0.1)  # let the bridge's receive loop process 'hello' and register
        return pane

    @staticmethod
    async def call(fn, *args, **kwargs):
        """Run a synchronous comments_live.execute_*_live call (which
        blocks on LiveSession.request_threadsafe) in a worker thread --
        the fake pane's own WebSocket connection/recv loop runs on THIS
        test's event loop (the main thread), so calling a blocking
        request_threadsafe() directly from a test coroutine would starve
        that same loop of the chance to ever answer the request it is
        waiting on. A real FastMCP tool call is dispatched on its own
        thread for the same underlying reason (request_threadsafe must
        not run on the bridge's own ops-loop thread, and here must not
        run on the fake pane's loop thread either) -- this mirrors that."""
        return await asyncio.to_thread(fn, *args, **kwargs)


# ---------------------------------------------------------------------------
# list_open_items(source="live") + correlation, against the multipara fixture.
# ---------------------------------------------------------------------------


class ListOpenItemsLiveCorrelationTests(LiveCommentsTestCase):
    fixture_name = "comments/multipara-comment.docx"

    def _fixture_comments(self) -> FakeDocument:
        doc = FakeDocument(text="", author_name=ROOT_AUTHOR)
        doc.comments = [
            FakeComment(
                id="c-root",
                content=ROOT_CONTENT_LIVE,
                author_name=ROOT_AUTHOR,
                creation_date=ROOT_DATE,
                anchor_text=ROOT_ANCHOR,
                resolved=False,
            ),
            FakeComment(
                id="c-single",
                content=SINGLE_CONTENT_LIVE,
                author_name=ROOT_AUTHOR,
                creation_date=ROOT_DATE,
                anchor_text=SINGLE_ANCHOR,
                resolved=False,
            ),
        ]
        return doc

    async def test_shape_and_exact_correlation_for_both_fixture_comments(self):
        doc = self._fixture_comments()
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(result["source"], "live")
        self.assertEqual(len(result["comments"]), 2)

        by_id = {c["comment_id"]: c for c in result["comments"]}
        root = by_id["live:c-root"]
        self.assertEqual(root["w_id"], None)
        self.assertEqual(root["content"], ROOT_CONTENT_LIVE)
        self.assertEqual(root["anchor_text"], ROOT_ANCHOR)
        self.assertEqual(root["author"], ROOT_AUTHOR)
        self.assertFalse(root["resolved"])
        self.assertEqual(root["replies"], [])

        correlation_by_live_id = {e["live_comment_id"]: e for e in result["correlation"]}
        self.assertEqual(correlation_by_live_id["live:c-root"]["comment_id"], ROOT_DURABLE_ID)
        self.assertEqual(correlation_by_live_id["live:c-root"]["confidence"], "exact")
        self.assertEqual(correlation_by_live_id["live:c-single"]["comment_id"], SINGLE_DURABLE_ID)
        self.assertEqual(correlation_by_live_id["live:c-single"]["confidence"], "exact")

    async def test_resolved_live_comment_filtered_from_comments_and_correlation(self):
        doc = self._fixture_comments()
        doc.comments[1].resolved = True  # the single-paragraph one
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        ids = {c["comment_id"] for c in result["comments"]}
        self.assertEqual(ids, {"live:c-root"})
        self.assertEqual({e["live_comment_id"] for e in result["correlation"]}, {"live:c-root"})

    async def test_correlation_none_when_content_differs(self):
        doc = self._fixture_comments()
        doc.comments.append(
            FakeComment(
                id="c-unrelated",
                content="Nothing in the fixture says this.",
                author_name=ROOT_AUTHOR,
                creation_date=ROOT_DATE,
                anchor_text="unrelated anchor",
                resolved=False,
            )
        )
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        correlation_by_live_id = {e["live_comment_id"]: e for e in result["correlation"]}
        entry = correlation_by_live_id["live:c-unrelated"]
        self.assertEqual(entry["confidence"], "none")
        self.assertIsNone(entry["comment_id"])
        self.assertIsNone(entry["w_id"])
        self.assertIsNone(entry["date_skew_s"])

    async def test_exact_confidence_survives_real_pane_date_skew(self):
        """issue #106 WP-6 real-pane finding: Word for Mac wrote this
        fixture's own w:date as '2026-09-16T14:06:00Z' (local wall-clock
        time with a 'Z' suffix), while Office.js's real creationDate for
        the SAME comment was true UTC, '2026-09-16T18:06:00.000Z' -- a
        4-hour gap the old 2-minute date-tolerance window could never
        close, so every real match came back 'content-only'. Confidence
        must be 'exact' regardless (anchor + content + author all still
        match); date_skew_s reports the gap, informationally, never
        gates it."""
        doc = self._fixture_comments()
        doc.comments[0].creation_date = "2026-09-16T18:06:00.000Z"  # the root comment's live-side date
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        correlation_by_live_id = {e["live_comment_id"]: e for e in result["correlation"]}
        root = correlation_by_live_id["live:c-root"]
        self.assertEqual(root["confidence"], "exact")
        self.assertEqual(root["comment_id"], ROOT_DURABLE_ID)
        # file-mode's own w:date for this fixture is "2026-09-16T14:06:00Z";
        # 18:06 - 14:06 = 4 hours = 14400s, live side later (positive).
        self.assertEqual(root["date_skew_s"], 14400)
        self.assertFalse(root["w_id_stable"])

    async def test_w_id_stable_false_on_every_correlation_entry(self):
        """issue #106 WP-6 real-pane finding: a real Word for Mac save
        renumbered a comment's w:id (1 -> 3), so w_id is only ever
        advisory within one open session -- every correlation entry must
        say so, regardless of confidence."""
        doc = self._fixture_comments()
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertTrue(result["correlation"], "expected at least one correlation entry")
        for entry in result["correlation"]:
            self.assertIs(entry["w_id_stable"], False)


# ---------------------------------------------------------------------------
# add_anchored_comment(write_mode="live")
# ---------------------------------------------------------------------------


class AddAnchoredCommentLiveTests(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def test_happy_path(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        evidence = await self.call(
            comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "Nice color choice.", 1
        )
        for key in ("applied", "match_count", "rung", "before", "after", "revision_before", "revision_after", "audit_logged"):
            self.assertIn(key, evidence)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["comment_id"], "live:c1")
        self.assertEqual(evidence["comment_ids"], ["live:c1"])
        self.assertEqual(evidence["rung"], "exact")
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["verified_via"], "word-addin")
        self.assertEqual(evidence["author"], "word-signed-in-user")
        self.assertEqual(evidence["orphaned_comment_ids"], [])
        self.assertNotIn("conflict_copy_detected", evidence)
        self.assertEqual(len(doc.comments), 1)
        self.assertEqual(doc.comments[0].content, "Nice color choice.")

    async def test_rung_equals_file_mode_add_anchored_comment(self):
        """issue #106 WP-6 real-pane finding: live evidence used to report
        the placeholder string "live" for `rung`; it must instead equal
        whatever file mode's own `add_anchored_comment` reports for an
        ordinary single-pass match (`locate.RUNG_EXACT`, "exact") -- not
        the unrelated numeric edit-ladder `rung` replace_text/format_text
        report."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        live_evidence = await self.call(
            comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1
        )
        # A separate quote on the SAME real fixture file -- file mode
        # writes to self.target directly; the fake pane only ever
        # touches the in-memory FakeDocument, so the two do not collide.
        file_evidence = comments.execute_add_anchored_comment(str(self.target), "lazy dog", "y", 1)

        self.assertIsInstance(live_evidence["rung"], str)
        self.assertEqual(live_evidence["rung"], file_evidence["rung"])
        self.assertEqual(live_evidence["rung"], "exact")

    async def test_match_count_mismatch(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 2)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.MATCH_COUNT_MISMATCH)

    async def test_zero_match(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "nonexistent phrase", "x", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ZERO_MATCH)


# ---------------------------------------------------------------------------
# reply_to_comment(write_mode="live")
# ---------------------------------------------------------------------------


class ReplyToCommentLiveTests(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def test_reply_via_live_handle(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]  # "live:c1"

        evidence = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), live_id, "a reply")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["comment_id_resolved_via"], "live-handle")
        self.assertEqual(evidence["comment_id"], live_id)
        self.assertEqual(evidence["parent_comment_id"], live_id)
        self.assertEqual(len(doc.comments[0].replies), 1)
        self.assertEqual(doc.comments[0].replies[0].content, "a reply")

    async def test_evidence_carries_live_sha256_revision_tokens(self):
        """issue #106 WP-6 real-pane finding: reply evidence used to
        report revision_before/revision_after as a bare None; they must
        instead be "live:sha256:<hex>" describe()-sourced tokens, the
        same shape replace_text/format_text/add_anchored_comment's live
        evidence already uses. A reply never edits body text, so the two
        are expected to be equal here (mirrors format_text's own
        before==after case), not a sign nothing happened."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]

        evidence = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), live_id, "a reply")
        self.assertIsNotNone(evidence["revision_before"])
        self.assertIsNotNone(evidence["revision_after"])
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertTrue(evidence["revision_after"].startswith("live:sha256:"))
        self.assertEqual(evidence["revision_before"], evidence["revision_after"])


class ReplyToCommentLiveCorrelationTests(LiveCommentsTestCase):
    fixture_name = "comments/multipara-comment.docx"

    async def test_reply_via_durable_id_correlation(self):
        doc = FakeDocument(text="", author_name=ROOT_AUTHOR)
        doc.comments = [
            FakeComment(
                id="c-root",
                content=ROOT_CONTENT_LIVE,
                author_name=ROOT_AUTHOR,
                creation_date=ROOT_DATE,
                anchor_text=ROOT_ANCHOR,
                resolved=False,
            )
        ]
        await self.connect_pane(document=doc)

        evidence = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), ROOT_DURABLE_ID, "a matched reply")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["comment_id_resolved_via"], "durableId-correlation")
        self.assertEqual(evidence["comment_id"], "live:c-root")
        self.assertEqual(evidence["parent_comment_id"], ROOT_DURABLE_ID)
        self.assertEqual(doc.comments[0].replies[0].content, "a matched reply")

    async def test_unresolvable_comment_id_raises_invalid_input(self):
        doc = FakeDocument(text="", author_name=ROOT_AUTHOR)
        doc.comments = [
            FakeComment(
                id="c-root",
                content=ROOT_CONTENT_LIVE,
                author_name=ROOT_AUTHOR,
                creation_date=ROOT_DATE,
                anchor_text=ROOT_ANCHOR,
                resolved=False,
            )
        ]
        await self.connect_pane(document=doc)

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_reply_to_comment_live, str(self.target), "nonexistent-id", "x")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


# ---------------------------------------------------------------------------
# resolve_comment(write_mode="live")
# ---------------------------------------------------------------------------


class ResolveCommentLiveTests(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def test_resolve_via_live_handle(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]

        evidence = await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["before"], "open")
        self.assertEqual(evidence["after"], "resolved")
        self.assertEqual(evidence["comment_id_resolved_via"], "live-handle")
        self.assertTrue(doc.comments[0].resolved)

    async def test_evidence_carries_live_sha256_revision_tokens(self):
        """issue #106 WP-6 real-pane finding: resolve evidence used to
        report revision_before/revision_after as a bare None; they must
        instead be "live:sha256:<hex>" describe()-sourced tokens, same as
        reply_to_comment's own fix -- see that test's docstring."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]

        evidence = await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        self.assertIsNotNone(evidence["revision_before"])
        self.assertIsNotNone(evidence["revision_after"])
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertTrue(evidence["revision_after"].startswith("live:sha256:"))
        self.assertEqual(evidence["revision_before"], evidence["revision_after"])

    async def test_comment_still_open_when_pane_refuses_to_flip(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]
        raw_id = live_id.split(":", 1)[1]
        doc.stuck_ids.add(raw_id)

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.COMMENT_STILL_OPEN)
        self.assertFalse(doc.comments[0].resolved)


# ---------------------------------------------------------------------------
# write_mode resolution: LIVE_UNAVAILABLE / auto without an owner file.
# ---------------------------------------------------------------------------


class WriteModeResolutionTests(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def test_live_unavailable_with_no_connected_session(self):
        with self.assertRaises(VerifyError) as cm:
            comments_live._resolve_write_mode(str(self.target), "live")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    async def test_auto_without_owner_file_takes_the_file_path(self):
        # A pane IS connected, but no Word/LibreOffice owner file exists
        # for self.target -- "auto" must still resolve to "file" (both
        # conditions are required, not either) and the write must go
        # through comments.py's own file-mode path, untouched.
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        mode = comments_live._resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "file")

        evidence = comments.execute_add_anchored_comment(str(self.target), "brown fox", "file-mode text", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(doc.comments, [])  # the fake pane received no op


if __name__ == "__main__":
    unittest.main()
