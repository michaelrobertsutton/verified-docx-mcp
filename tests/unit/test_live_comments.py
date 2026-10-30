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
import json
import os
import shutil
import ssl
import sys
import tempfile
import unittest
import unittest.mock
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeComment, FakeDocument, FakePane, FakeReply, docx_to_flat_opc

from verified_docx_mcp import comments, mutations, paths, server
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import bridge, comments_live
from verified_docx_mcp.live.protocol import HelloMessage
from verified_docx_mcp.live.session import LiveOpFailed

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

    async def _list_with_correlation(self):
        return await self.call(
            lambda: comments_live.execute_list_open_items_live(str(self.target), include_correlation=True)
        )

    async def test_shape_and_exact_correlation_for_both_fixture_comments(self):
        doc = self._fixture_comments()
        await self.connect_pane(document=doc)

        result = await self._list_with_correlation()
        self.assertEqual(result["source"], "live")
        self.assertEqual(len(result["comments"]), 2)

        by_id = {c["comment_id"]: c for c in result["comments"]}
        root = by_id["live:c-root"]
        for gone in ("w_id", "quoted_text", "modified_time", "scope", "session_epoch"):
            self.assertNotIn(gone, root)  # constant or duplicate fields are not sent (#111)
        self.assertEqual(root["content"], ROOT_CONTENT_LIVE)
        self.assertEqual(root["anchor_text"], ROOT_ANCHOR)
        self.assertEqual(root["author"], ROOT_AUTHOR)
        self.assertFalse(root["resolved"])
        self.assertNotIn("replies", root)  # empty list omitted (#111)
        self.assertNotIn("reply_count", root)

        correlation_by_live_id = {e["live_comment_id"]: e for e in result["correlation"]}
        self.assertEqual(correlation_by_live_id["live:c-root"]["comment_id"], ROOT_DURABLE_ID)
        self.assertEqual(correlation_by_live_id["live:c-root"]["confidence"], "exact")
        self.assertEqual(correlation_by_live_id["live:c-single"]["comment_id"], SINGLE_DURABLE_ID)
        self.assertEqual(correlation_by_live_id["live:c-single"]["confidence"], "exact")

    async def test_resolved_live_comment_filtered_from_comments_and_correlation(self):
        doc = self._fixture_comments()
        doc.comments[1].resolved = True  # the single-paragraph one
        await self.connect_pane(document=doc)

        result = await self._list_with_correlation()
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

        result = await self._list_with_correlation()
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

        result = await self._list_with_correlation()
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

        result = await self._list_with_correlation()
        self.assertTrue(result["correlation"], "expected at least one correlation entry")
        for entry in result["correlation"]:
            self.assertIs(entry["w_id_stable"], False)

    # -- issue #39: counts / scope / pane identity / warnings ---------------

    async def test_counts_scope_and_pane_identity_come_first(self):
        doc = self._fixture_comments()
        doc.comments[1].resolved = True
        await self.connect_pane(document=doc, instance_id="pane-A", platform="Mac")

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(list(result)[:4], ["counts", "scope", "pane", "path"])
        self.assertEqual(result["counts"], {"live_total": 2, "live_open": 1, "returned": 1})
        self.assertEqual(result["scope"], "body")
        self.assertEqual(result["pane"]["instance_id"], "pane-A")
        self.assertEqual(result["pane"]["platform"], "Mac")
        self.assertEqual(result["pane"]["observed_at"], "2026-10-01T12:00:00.000Z")
        self.assertNotIn("scope_body_only", [w["code"] for w in result.get("warnings", [])])  # `scope` says it (#111)
        self.assertNotIn("pane_note", result)

    async def test_older_pane_without_counts_gets_computed_counts_and_note(self):
        doc = self._fixture_comments()
        pane = await self.connect_pane(document=doc, capabilities=[])
        pane.comments_list_transform = lambda r: {"comments": r["comments"]}

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(result["counts"], {"live_total": 2, "live_open": 2, "returned": 2})
        self.assertEqual(result["pane_note"], comments_live.STALE_PANE_NOTE_COUNTS)

    async def test_missing_comments_list_is_refused_not_reported_empty(self):
        pane = await self.connect_pane(document=self._fixture_comments())
        pane.comments_list_transform = lambda r: {"counts": r["counts"]}

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_OP_FAILED)

    async def test_counts_that_disagree_with_returned_comments_are_refused(self):
        pane = await self.connect_pane(document=self._fixture_comments())
        pane.comments_list_transform = lambda r: {**r, "counts": {"total": 7, "open": 7}}

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_OP_FAILED)

    async def test_file_open_comment_missing_from_pane_is_advisory_warning(self):
        # The pane reports only the single-paragraph comment; the saved
        # file's multi-paragraph root comment has no live counterpart.
        doc = self._fixture_comments()
        doc.comments = [doc.comments[1]]
        await self.connect_pane(document=doc)

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertIn("file_has_unlisted_comments", [w["code"] for w in result.get("warnings", [])])
        self.assertEqual(
            [c["comment_id"] for c in result["file_only_open_comments"]], [ROOT_DURABLE_ID]
        )

    async def test_second_pane_on_same_document_warns(self):
        doc = self._fixture_comments()
        await self.connect_pane(document=doc, instance_id="pane-old", platform="Mac")
        await self.connect_pane(document=doc, instance_id="pane-new", platform="OfficeOnline")

        result = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        self.assertEqual(result["pane"]["instance_id"], "pane-new")
        dup = [w for w in result.get("warnings", []) if w["code"] == "same_document_duplicate"]
        self.assertEqual(len(dup), 1)
        self.assertEqual(dup[0]["displaced_instance_id"], "pane-old")


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
        report.

        issue #154: the file-mode half of this comparison must run with
        NO pane connected -- since a connected pane now makes
        mutations._guard_before_write refuse every file-mode write with
        LIVE_SESSION_ACTIVE (this test's own pane is connected for this
        same document/basename), so the live half runs first and its
        pane is closed before the file-mode half runs."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        pane = await self.connect_pane(document=doc)

        live_evidence = await self.call(
            comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1
        )
        await pane.close()
        self._panes.remove(pane)
        await asyncio.sleep(0.1)  # let the bridge's receive loop process the close and unregister

        # A separate quote on the SAME real fixture file, now with no pane
        # connected -- file mode writes to self.target directly; the fake
        # pane only ever touched the in-memory FakeDocument, so the two
        # never collided on content, only on the (now-closed) session.
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

    async def test_listed_epoch_handle_resolves_and_replies(self):
        """A real-Word probe for #117 found list_open_items' live:<epoch>:<id>
        handle reached the pane as "<epoch>:<id>" and was refused; the listed
        handle must work as-is for reply and resolve."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        doc.session_epoch = "ep9"
        await self.connect_pane(document=doc)
        await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        listed = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        handle = listed["comments"][0]["comment_id"]
        self.assertTrue(handle.startswith("live:ep9:"))

        replied = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), handle, "ok")
        self.assertTrue(replied["applied"])
        resolved = await self.call(comments_live.execute_resolve_comment_live, str(self.target), handle)
        self.assertTrue(resolved["applied"])
        self.assertTrue(doc.comments[0].resolved)

    async def test_listed_handle_from_an_earlier_epoch_is_refused(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        doc.session_epoch = "ep9"
        await self.connect_pane(document=doc)
        await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        listed = await self.call(comments_live.execute_list_open_items_live, str(self.target))
        stale = listed["comments"][0]["comment_id"].replace("live:ep9:", "live:ep8:")

        with self.assertRaises(VerifyError) as ctx:
            await self.call(comments_live.execute_resolve_comment_live, str(self.target), stale)
        self.assertEqual(ctx.exception.envelope.error_code.value, "COMMENT_ID_STALE")
        self.assertFalse(doc.comments[0].resolved)

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


class DefaultAutoDispatchRegressionTests(LiveCommentsTestCase):
    """issue #154's actual repro for the comment tools: every test above
    calls comments_live.execute_*_live DIRECTLY, bypassing the
    write_mode="auto" DISPATCH that a real caller (and the original
    incident) actually goes through -- server.reply_to_comment/
    resolve_comment, which resolve the mode themselves via
    comments_live._resolve_write_mode before picking an implementation.
    None of the tests above would have caught the original bug: on Mac +
    SharePoint (no owner file), "auto" used to resolve to "file", so a
    live:<id> handle from list_open_items(source="live") reached
    comments.execute_reply_to_comment's file-mode durableId lookup and
    failed with INVALID_INPUT: comment_id 'live:<id>' not found -- exactly
    the failure the incident report quotes verbatim. These tests exercise
    the dispatch layer itself, with write_mode left at its default."""

    fixture_name = "frag.docx"

    async def test_reply_to_comment_default_auto_dispatches_live_handle(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]  # "live:c1" -- the exact handle shape the incident's comment_id was

        evidence = await self.call(server.reply_to_comment, str(self.target), live_id, "a reply")
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(len(doc.comments[0].replies), 1)

    async def test_resolve_comment_default_auto_dispatches_live_handle(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        live_id = added["comment_id"]

        evidence = await self.call(server.resolve_comment, str(self.target), live_id)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertTrue(doc.comments[0].resolved)


# ---------------------------------------------------------------------------
# delete_comment / delete_resolved_comments (issue #80)
# ---------------------------------------------------------------------------

TEXT = "The quick brown fox jumps over the lazy dog."


class _DeleteLiveMixin:
    fixture_name = "frag.docx"

    async def expect_error(self, fn, *args, **kwargs):
        """server.* tools raise a FastMCP ToolError carrying the JSON envelope."""
        import json

        from fastmcp.exceptions import ToolError

        with self.assertRaises(ToolError) as cm:
            await self.call(fn, *args, **kwargs)
        return json.loads(str(cm.exception))

    async def _two(self):
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(document=doc)
        a = (await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "A", 1))["comment_id"]
        b = (await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "lazy dog", "B", 1))["comment_id"]
        return doc, pane, a, b


class DeleteCommentLiveTests(_DeleteLiveMixin, LiveCommentsTestCase):
    async def test_delete_via_live_handle_with_replies_and_evidence(self):
        doc, _pane, a, b = await self._two()
        await self.call(comments_live.execute_reply_to_comment_live, str(self.target), a, "reply")

        ev = await self.call(server.delete_comment, str(self.target), a)

        self.assertTrue(ev["applied"])
        self.assertEqual(ev["write_mode"], "live")
        self.assertEqual(ev["rung"], "delete")
        self.assertEqual(ev["after"], "deleted")
        self.assertEqual(ev["before"], "brown fox")
        self.assertEqual(ev["comment_id"], a)
        self.assertEqual(ev["comment_id_resolved_via"], "live-handle")
        self.assertEqual(ev["deleted_comment_ids"], [a])
        self.assertEqual(ev["deleted_reply_count"], 1)
        self.assertEqual((ev["comments_before"], ev["comments_after"]), (2, 1))
        self.assertTrue(ev["other_comments_unchanged"])
        self.assertEqual(ev["scope"], "body")
        self.assertTrue(ev["revision_before"].startswith("live:sha256:"))
        self.assertEqual([c.id for c in doc.comments], [b[len("live:"):]])

    async def test_open_comment_may_be_deleted_by_the_single_tool(self):
        doc, _pane, a, _b = await self._two()
        self.assertFalse(doc.comments[0].resolved)
        ev = await self.call(server.delete_comment, str(self.target), a)
        self.assertTrue(ev["applied"])

    async def test_pane_without_the_capability_is_refused_before_any_op(self):
        doc = FakeDocument(text=TEXT)
        pane = await self.connect_pane(document=doc, capabilities=["comments_by_id", "shape_guard"])
        doc.comments.append(FakeComment("c9", "x", "A", "d", "brown fox"))
        env = await self.expect_error(server.delete_comment, str(self.target), "live:c9")
        self.assertEqual(env["error_code"], ErrorCode.LIVE_CAPABILITY_MISSING.value)
        self.assertEqual(pane.comment_delete_payloads, [])
        self.assertEqual(len(doc.comments), 1)

    async def test_pane_that_does_not_delete_raises_comment_still_present(self):
        doc, _pane, a, _b = await self._two()
        doc.undeletable_ids.add(a[len("live:"):])
        env = await self.expect_error(server.delete_comment, str(self.target), a)
        self.assertEqual(env["error_code"], ErrorCode.COMMENT_STILL_PRESENT.value)
        self.assertEqual(env["diagnostics"]["remaining"], [a])
        self.assertEqual(env["diagnostics"]["deleted"], [])

    async def test_another_comments_reply_changing_is_verification_failed(self):
        doc, _pane, a, b = await self._two()
        await self.call(comments_live.execute_reply_to_comment_live, str(self.target), b, "keep me")

        def corrupt(document):
            other = next(c for c in document.comments if c.id == b[len("live:"):])
            other.replies[0].content = "silently edited"

        doc.after_comment_delete = corrupt
        env = await self.expect_error(server.delete_comment, str(self.target), a)
        self.assertEqual(env["error_code"], ErrorCode.VERIFICATION_FAILED.value)
        self.assertEqual(env["diagnostics"]["drifted_comments"], [b])
        self.assertEqual(env["diagnostics"]["deleted"], [a])

    async def test_error_after_the_delete_applied_is_reconciled_not_reported_as_remaining(self):
        doc, _pane, a, _b = await self._two()
        doc.fail_after_delete_ids.add(a[len("live:"):])
        env = await self.expect_error(server.delete_comment, str(self.target), a)
        self.assertEqual(env["error_code"], ErrorCode.VERIFICATION_FAILED.value)
        self.assertEqual(env["diagnostics"]["deleted"], [a])
        self.assertEqual(env["diagnostics"]["remaining"], [])

    async def test_untrusted_readback_leaves_the_thread_unknown_never_deleted(self):
        doc, pane, a, _b = await self._two()
        calls = {"n": 0}

        def malformed_after_first(reply):
            calls["n"] += 1
            # 1: address_comment, 2: pre-snapshot, 3+: the post re-list.
            return {k: v for k, v in reply.items() if k != "comments"} if calls["n"] >= 3 else reply

        pane.comments_list_transform = malformed_after_first
        env = await self.expect_error(server.delete_comment, str(self.target), a)
        self.assertEqual(env["error_code"], ErrorCode.LIVE_OP_FAILED.value)
        self.assertEqual(env["diagnostics"]["unknown"], [a])
        self.assertEqual(env["diagnostics"]["deleted"], [])

    async def test_handle_from_another_pane_epoch_is_stale(self):
        doc, _pane, a, _b = await self._two()
        doc.session_epoch = "epoch-2"
        env = await self.expect_error(server.delete_comment, str(self.target), a, session_epoch="epoch-1")
        self.assertEqual(env["error_code"], ErrorCode.COMMENT_ID_STALE.value)
        self.assertEqual(len(doc.comments), 2)

    async def test_a_reply_cannot_be_addressed(self):
        doc, _pane, a, _b = await self._two()
        await self.call(comments_live.execute_reply_to_comment_live, str(self.target), a, "reply")
        reply_handle = f"live:{doc.comments[0].replies[0].id}"
        env = await self.expect_error(server.delete_comment, str(self.target), reply_handle)
        self.assertEqual(env["error_code"], ErrorCode.INVALID_INPUT.value)
        self.assertEqual(len(doc.comments), 2)


class DeleteResolvedCommentsLiveTests(_DeleteLiveMixin, LiveCommentsTestCase):
    async def test_batch_deletes_resolved_threads_with_one_snapshot_and_one_relist(self):
        doc, pane, a, b = await self._two()
        for c in doc.comments:
            c.resolved = True
        before = len(pane.comments_list_payloads)

        ev = await self.call(server.delete_resolved_comments, str(self.target), [a, b, a])

        self.assertEqual(ev["deleted_thread_count"], 2)
        self.assertEqual(ev["match_count"], 2)
        self.assertEqual(sorted(ev["deleted_comment_ids"]), sorted([a, b]))
        self.assertEqual(len(ev["resolutions"]), 3)
        self.assertEqual(ev["before"], ["brown fox", "lazy dog"])
        self.assertEqual(doc.comments, [])
        self.assertEqual(len(pane.comments_list_payloads) - before, 2)
        self.assertEqual(len(pane.comment_delete_payloads), 2)
        self.assertTrue(all(p["require_resolved"] for p in pane.comment_delete_payloads))

    async def test_open_or_unknown_ids_refuse_the_whole_batch(self):
        doc, pane, a, b = await self._two()
        doc.comments[0].resolved = True
        env = await self.expect_error(server.delete_resolved_comments, str(self.target), [a, b, "live:nope"])
        self.assertEqual(env["error_code"], ErrorCode.INVALID_INPUT.value)
        reasons = {o["comment_id"]: o["reason"] for o in env["diagnostics"]["offenders"]}
        self.assertEqual(reasons, {b: "not_resolved", "live:nope": "unknown_or_reply"})
        self.assertEqual(pane.comment_delete_payloads, [])
        self.assertEqual(len(doc.comments), 2)

    async def test_empty_list_is_refused(self):
        await self._two()
        env = await self.expect_error(server.delete_resolved_comments, str(self.target), [])
        self.assertEqual(env["error_code"], ErrorCode.INVALID_INPUT.value)

    async def test_failure_part_way_reports_deleted_and_remaining(self):
        doc, pane, a, b = await self._two()
        for c in doc.comments:
            c.resolved = True
        doc.fail_after_delete_ids.add(a[len("live:"):])
        env = await self.expect_error(server.delete_resolved_comments, str(self.target), [a, b])
        self.assertEqual(env["diagnostics"]["deleted"], [a])
        self.assertEqual(env["diagnostics"]["remaining"], [b])
        self.assertEqual(len(pane.comment_delete_payloads), 1)
        self.assertEqual([c.id for c in doc.comments], [b[len("live:"):]])


class DeleteResolvedCommentsLiveCorrelationTests(LiveCommentsTestCase):
    fixture_name = "comments/multipara-comment.docx"

    async def test_batch_resolves_a_durable_id_through_correlation(self):
        doc = FakeDocument(text="", author_name=ROOT_AUTHOR)
        doc.comments = [
            FakeComment("c-root", ROOT_CONTENT_LIVE, ROOT_AUTHOR, ROOT_DATE, ROOT_ANCHOR, resolved=True),
            FakeComment("c-single", SINGLE_CONTENT_LIVE, ROOT_AUTHOR, ROOT_DATE, SINGLE_ANCHOR, resolved=False),
        ]
        await self.connect_pane(document=doc)

        ev = await self.call(server.delete_resolved_comments, str(self.target), [ROOT_DURABLE_ID])

        self.assertEqual(ev["resolutions"][0]["resolved_via"], "durableId-correlation")
        self.assertEqual(ev["comment_id"], "live:c-root")
        self.assertEqual([c.id for c in doc.comments], ["c-single"])


# ---------------------------------------------------------------------------
# write_mode resolution: LIVE_UNAVAILABLE / auto without an owner file.
# ---------------------------------------------------------------------------


class WriteModeResolutionTests(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def test_live_unavailable_with_no_connected_session(self):
        with self.assertRaises(VerifyError) as cm:
            comments_live._resolve_write_mode(str(self.target), "live")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    async def test_auto_without_owner_file_resolves_live_and_file_mode_refuses(self):
        """issue #154: a pane IS connected, and no Word/LibreOffice owner
        file exists for self.target (Word for Mac + a SharePoint/OneDrive
        sync never writes one) -- "auto" must resolve to "live" on the
        connected-session signal alone, and an explicit file-mode write
        for this same document must now refuse with LIVE_SESSION_ACTIVE
        rather than silently write to disk underneath Word's own open
        buffer. This inverts the pre-#154 version of this test, which
        asserted the opposite of both."""
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        mode = comments_live._resolve_write_mode(str(self.target), "auto")
        self.assertEqual(mode, "live")

        with self.assertRaises(VerifyError) as ctx:
            comments.execute_add_anchored_comment(str(self.target), "brown fox", "file-mode text", 1)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_SESSION_ACTIVE)
        self.assertEqual(doc.comments, [])  # the fake pane received no op either


class NoLocalFileLiveTests(LiveCommentsTestCase):
    """Issue #22 B3: every live comment tool must work when `path` names
    NO local file at all -- the real incident's own shape (a SharePoint/
    OneDrive document with no local sync). asyncSetUp's base class copies
    the fixture to self.target first (so document_url's basename still
    matches what a real pane would report); this class deletes it right
    after, before connecting any pane, so every call below genuinely has
    no local file to fall back to."""

    fixture_name = "frag.docx"

    async def asyncSetUp(self) -> None:
        await super().asyncSetUp()
        self.target.unlink()
        self.assertFalse(self.target.exists())

    async def test_list_open_items_live_reports_file_side_unavailable(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        result = await self.call(
            lambda: comments_live.execute_list_open_items_live(str(self.target), include_correlation=True)
        )
        self.assertEqual(result["source"], "live")
        self.assertIsNone(result["pending_suggestions"])
        self.assertFalse(result["file_side_available"])
        self.assertEqual(result["correlation"], [])

    async def test_add_reply_resolve_all_work_via_live_handle_with_no_local_file(self):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        added = await self.call(
            comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1
        )
        self.assertTrue(added["applied"])
        live_id = added["comment_id"]

        reply = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), live_id, "a reply")
        self.assertTrue(reply["applied"])

        resolved = await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        self.assertTrue(resolved["applied"])

    async def test_durable_id_resolution_fails_closed_with_no_local_file(self):
        # Without a local file there is no file-side correlation table at
        # all -- a durableId/w:id handle (as opposed to a live:<id> one)
        # can never resolve, and must say so rather than silently
        # matching the wrong thing.
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        await self.connect_pane(document=doc)

        with self.assertRaises(VerifyError) as cm:
            await self.call(comments_live.execute_reply_to_comment_live, str(self.target), "58A3F864", "x")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


if __name__ == "__main__":
    unittest.main()


class TargetedVerifyTests(LiveCommentsTestCase):
    """Issue #31: reply/resolve must not re-list every comment."""

    fixture_name = "frag.docx"

    async def _two_comments(self, **pane_kwargs):
        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        pane = await self.connect_pane(document=doc, **pane_kwargs)
        first = await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "lazy dog", "y", 1)
        pane.comments_list_payloads.clear()
        return doc, pane, first["comment_id"]

    async def test_reply_with_live_handle_only_asks_for_the_target(self):
        _doc, pane, live_id = await self._two_comments()
        evidence = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), live_id, "r")
        raw_id = live_id.split(":", 1)[1]
        self.assertEqual(
            pane.comments_list_payloads,
            [
                {"ids": [raw_id], "include_anchor": True},  # pre: parent anchor for evidence
                {"ids": [raw_id], "include_anchor": False},  # post: verify
            ],
        )
        self.assertEqual(evidence["before"], "brown fox")
        self.assertNotIn("pane_note", evidence)

    async def test_resolve_with_live_handle_only_verifies_the_target(self):
        doc, pane, live_id = await self._two_comments()
        evidence = await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        raw_id = live_id.split(":", 1)[1]
        self.assertEqual(pane.comments_list_payloads, [{"ids": [raw_id], "include_anchor": False}])
        self.assertTrue(doc.comments[0].resolved)
        self.assertNotIn("pane_note", evidence)

    async def test_pane_without_capability_falls_back_and_flags_stale_pane(self):
        doc, pane, live_id = await self._two_comments(capabilities=["row_scope", "shape_guard"])
        reply = await self.call(comments_live.execute_reply_to_comment_live, str(self.target), live_id, "r")
        resolve = await self.call(comments_live.execute_resolve_comment_live, str(self.target), live_id)
        self.assertTrue(all("ids" not in p for p in pane.comments_list_payloads))
        self.assertIn("close and reopen", reply["pane_note"])
        self.assertIn("close and reopen", resolve["pane_note"])
        self.assertTrue(doc.comments[0].resolved)


# ---------------------------------------------------------------------------
# Issue #111: list_open_items(source="live") fields/scoping/trim.
# ---------------------------------------------------------------------------

SECTIONS = FIXTURES / "sections.docx"
PARA_OVERVIEW = "Some overview text."
PARA_BACKGROUND = "Background text."
PARA_NEXT = "Next steps text."


def _flat_body(replace: dict[str, str] | None = None) -> str:
    """Flat OPC of sections.docx's body, optionally with literal text replaced in its document.xml
    (to make two paragraphs identical)."""
    if not replace:
        return docx_to_flat_opc(SECTIONS)
    out = Path(tempfile.mkdtemp()) / "sections.docx"
    with zipfile.ZipFile(SECTIONS) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                text = data.decode("utf-8")
                for old, new in replace.items():
                    assert old in text, old
                    text = text.replace(old, new)
                data = text.encode("utf-8")
            zout.writestr(info.filename, data)
    return docx_to_flat_opc(out)


def _comment(cid: str, anchor: str, paragraph: str | None, *, content: str = "please check", replies: int = 0):
    return FakeComment(
        id=cid,
        content=content,
        author_name="Pat Reviewer",
        creation_date="2026-10-01T12:00:00.000Z",
        anchor_text=anchor,
        anchor_paragraph_text=paragraph,
        replies=[
            FakeReply(id=f"{cid}-r{n}", content="agreed", author_name="Lee", creation_date="2026-10-01T13:00:00.000Z")
            for n in range(replies)
        ],
    )


def twenty_seven_comment_document() -> FakeDocument:
    """A realistic review pass: 27 open comments, a few with replies, mixed anchor lengths."""
    anchors = [
        "Staffing plan", "Transition schedule", "the incumbent's key personnel retention rate",
        "Quality control", "Phase II option", "past performance reference 3",
        "Skyward will deliver the first increment within 30 days of award and every 30 days after that",
    ]
    doc = FakeDocument(text="", author_name="Pat Reviewer")
    doc.session_epoch = "ep7a1c"
    for n in range(1, 28):
        anchor = anchors[n % len(anchors)]
        doc.comments.append(
            _comment(
                f"c{n}",
                anchor,
                f"{anchor} is described in this paragraph of the technical volume, section {n}.",
                content=f"Comment {n}: please tighten this language and cite the source for the figure quoted here.",
                replies=1 if n % 9 == 0 else 0,
            )
        )
    return doc


class ListOpenItemsLiveTrimTests(LiveCommentsTestCase):
    fixture_name = "sections.docx"

    def _document(self) -> FakeDocument:
        doc = FakeDocument(text="", author_name="Pat Reviewer")
        doc.session_epoch = "ep1"
        doc.comments = [
            _comment("c-ov", "overview", PARA_OVERVIEW, content="overview note", replies=1),
            _comment("c-bg", "Background", PARA_BACKGROUND, content="background note"),
            _comment("c-nx", "steps", PARA_NEXT, content="next note"),
        ]
        return doc

    async def _list(self, **kwargs):
        return await self.call(
            lambda: comments_live.execute_list_open_items_live(str(self.target), **kwargs)
        )

    async def test_full_drops_constant_and_duplicate_fields(self):
        await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        result = await self._list()
        self.assertEqual(result["session_epoch"], "ep1")  # once, at the top level
        self.assertEqual(result["scope"], "body")
        self.assertNotIn("scope_body_only", [w["code"] for w in result.get("warnings", [])])
        self.assertNotIn("correlation", result)
        for comment in result["comments"]:
            for gone in ("quoted_text", "w_id", "modified_time", "scope", "session_epoch", "reply_count"):
                self.assertNotIn(gone, comment)
            for kept in ("comment_id", "content", "resolved", "anchor_text",
                         "anchor_paragraph_text", "author", "created_time"):
                self.assertIn(kept, comment)
        by_id = {c["comment_id"]: c for c in result["comments"]}
        self.assertEqual(len(by_id["live:ep1:c-ov"]["replies"]), 1)  # only threads with replies carry the key
        self.assertNotIn("replies", by_id["live:ep1:c-bg"])

    async def test_include_correlation_restores_the_correlation_list(self):
        await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        result = await self._list(include_correlation=True)
        self.assertEqual(len(result["correlation"]), 3)

    async def test_summary_shape_and_kept_completeness_fields(self):
        doc = self._document()
        doc.ooxml = {"available": True, "total_threads": 4, "open_threads": 4, "ids": []}
        await self.connect_pane(document=doc, body_ooxml=_flat_body())
        result = await self._list(fields="summary")
        self.assertEqual(result["comment_columns"], ["comment_id", "author", "resolved", "anchor", "reply_count"])
        self.assertEqual(result["comment_rows"][0][4], 1)
        self.assertNotIn("suggestion_columns", result)
        self.assertIsNone(result["suggestion_rows"])  # fake pane has no live revisions
        for absent in ("comments", "suggestions", "pending_suggestions"):
            self.assertNotIn(absent, result)  # JennyStack's recorder keys on a `comments` list
        self.assertIn("path", result)
        for key in ("complete", "scope", "warnings", "source", "counts"):
            self.assertIn(key, result)
        self.assertEqual(result["source"], "live")
        self.assertIs(result["complete"], False)  # ooxml says 4 open, the pane reported 3
        self.assertIn("pane_collection_incomplete", [w["code"] for w in result.get("warnings", [])])
        self.assertNotIn("correlation", result)

    async def test_summary_anchor_is_clipped_to_60_chars(self):
        doc = self._document()
        doc.comments[0].anchor_text = "x" * 200
        await self.connect_pane(document=doc, body_ooxml=_flat_body())
        result = await self._list(fields="summary")
        self.assertEqual(result["comment_rows"][0][3], "x" * 60)

    async def test_unknown_fields_value_refused(self):
        await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        with self.assertRaises(VerifyError) as ctx:
            await self._list(fields="brief")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    async def test_section_key_scopes_live_comments_without_correlating_ids(self):
        pane = await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        result = await self._list(section_key="background-1")
        self.assertEqual([c["comment_id"] for c in result["comments"]], ["live:ep1:c-bg"])
        self.assertNotIn("scope_uncertain", result["comments"][0])
        self.assertEqual(result["counts"]["live_open"], 3)
        self.assertEqual(result["counts"]["returned"], 1)
        self.assertEqual(result["counts"]["open_comments"], 3)
        self.assertEqual(result["counts"]["in_scope"]["comments"], 1)
        self.assertEqual(pane.body_ooxml_requests, 1)

    async def test_within_text_scopes_live_comments(self):
        await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        result = await self._list(within_text="Next steps text")
        self.assertEqual([c["comment_id"] for c in result["comments"]], ["live:ep1:c-nx"])

    async def test_unscoped_call_never_reads_the_body(self):
        pane = await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        await self._list()
        await self._list(fields="summary")
        self.assertEqual(pane.body_ooxml_requests, 0)

    async def test_unknown_section_key_lists_available_keys(self):
        await self.connect_pane(document=self._document(), body_ooxml=_flat_body())
        with self.assertRaises(VerifyError) as ctx:
            await self._list(section_key="nope")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(
            ctx.exception.envelope.diagnostics["available_section_keys"],
            ["overview-1", "background-1", "next-steps-1"],
        )

    async def test_scoping_needs_a_pane_that_can_read_the_body(self):
        await self.connect_pane(
            document=self._document(),
            body_ooxml=_flat_body(),
            capabilities=["comment_counts"],
        )
        with self.assertRaises(VerifyError) as ctx:
            await self._list(section_key="background-1")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_identical_paragraphs_straddling_the_scope_are_flagged_uncertain(self):
        # Next Steps' paragraph now reads exactly like Background's.
        body = _flat_body({PARA_NEXT: PARA_BACKGROUND})
        doc = self._document()
        doc.comments[2].anchor_paragraph_text = PARA_BACKGROUND
        await self.connect_pane(document=doc, body_ooxml=body)
        result = await self._list(section_key="background-1")
        by_id = {c["comment_id"]: c for c in result["comments"]}
        self.assertEqual(set(by_id), {"live:ep1:c-bg", "live:ep1:c-nx"})
        self.assertTrue(by_id["live:ep1:c-bg"]["scope_uncertain"])
        self.assertTrue(by_id["live:ep1:c-nx"]["scope_uncertain"])
        summary = await self._list(section_key="background-1", fields="summary")
        self.assertEqual(summary["comment_columns"][-1], "scope_uncertain")
        self.assertTrue(all(c[-1] for c in summary["comment_rows"]))

    async def test_comment_with_no_paragraph_text_is_kept_and_flagged(self):
        doc = self._document()
        doc.comments[1].anchor_paragraph_text = None  # an older pane / unreadable anchor
        await self.connect_pane(document=doc, body_ooxml=_flat_body())
        result = await self._list(section_key="next-steps-1")
        by_id = {c["comment_id"]: c for c in result["comments"]}
        self.assertEqual(set(by_id), {"live:ep1:c-bg", "live:ep1:c-nx"})
        self.assertTrue(by_id["live:ep1:c-bg"]["scope_uncertain"])
        self.assertNotIn("scope_uncertain", by_id["live:ep1:c-nx"])

    async def test_scoped_list_does_not_rearm_a_stale_shared_bridge_client(self):
        """live_session_for hands back a non-observing view: the body read made for scoping must not
        acknowledge another client's write (see test_shared_bridge.py for the broker-level pattern)."""
        from types import SimpleNamespace

        from verified_docx_mcp.live import shared, write_mode

        target = self.target

        class Pane:
            document_name = target.name
            document_url = str(target)
            connected_since = 1.0
            last_body_sha256 = "initial"
            heartbeat_count = hash_sample_count = 0
            body_hash_state = "sampled"
            body_hash_sampled_monotonic = 0.0
            hello = HelloMessage(
                str(target), "Word", "Mac", {}, "initial",
                frozenset({"shared_queue", "body_ooxml", "comment_counts"}), "pane-1",
            )

            def __init__(self):
                self.calls: list[str] = []

            def heartbeat_age(self):
                return 0

            def request_threadsafe(self, op, payload=None, **kwargs):
                self.calls.append(op)
                if op == "comments_list":
                    return {
                        "comments": [_comment("c-bg", "Background", PARA_BACKGROUND).to_dict()],
                        "counts": {"total": 1, "open": 1}, "scope": "body", "observed_at": "t",
                        "session_epoch": "ep1",
                    }
                if op == "body_ooxml":
                    return {"ooxml": _flat_body(), "bodySha256": "initial"}
                if op == "describe":
                    return {"session_epoch": "ep1", "bodySha256": self.last_body_sha256}
                if op in shared.READ_OPS:
                    return {"bodySha256": "initial"}
                self.last_body_sha256 = "changed"
                return {"pre": "initial", "post": "changed"}

        pane = Pane()
        registry = SimpleNamespace(
            get=lambda name: pane if name == pane.document_name else None, list=lambda: [pane], collisions=list
        )
        broker = shared.Broker(registry)
        a = broker.dispatch({"action": "attach"})["client_id"]
        b = broker.dispatch({"action": "attach"})["client_id"]

        def call(client, op, **kw):
            return broker.dispatch(
                {"action": "request", "client_id": client, "document_name": pane.document_name,
                 "document_url": pane.document_url, "connected_since": pane.connected_since, "op": op, **kw}
            )

        call(a, "describe")
        call(b, "describe")
        call(a, "replace")  # A writes: B's baseline is now stale

        client = SimpleNamespace(rpc=lambda action, **kw: broker.dispatch({"action": action, "client_id": b, **kw}))
        remote = shared.RemoteSession(client, broker.snapshot(pane))
        with (
            unittest.mock.patch.object(
                bridge, "current_registry", lambda: SimpleNamespace(get=lambda n: remote, collisions=list)
            ),
            unittest.mock.patch.object(write_mode, "_check_session_identity", lambda path, session: None),
        ):
            result = await self._list(section_key="background-1", fields="summary")
        self.assertEqual([c[0] for c in result["comment_rows"]], ["live:ep1:c-bg"])
        self.assertIn("body_ooxml", pane.calls)
        with self.assertRaises(LiveOpFailed) as ctx:
            call(b, "replace")  # still stale: the scoped list did not re-arm it
        self.assertIn("LIVE_STALE", str(ctx.exception))


class ListOpenItemsLivePayloadSizeTests(LiveCommentsTestCase):
    """Issue #111 acceptance numbers on the 27-comment fixture (chars/4 ~ tokens)."""

    fixture_name = "sections.docx"

    @staticmethod
    def _size(value) -> int:
        return len(json.dumps(value, separators=(",", ":")))

    async def _old_shape_size(self, doc: FakeDocument) -> int:
        """The pre-#111 full response: per-comment constants/duplicates, per-comment epoch, the
        correlation list and the scope_body_only warning."""
        result = await self.call(
            lambda: comments_live.execute_list_open_items_live(str(self.target), include_correlation=True)
        )
        old = {"pending_suggestions_source": None, "file_only_open_comments": [], **result}  # keys now omitted
        old["comments"] = [
            comments_live._live_comment_record({**c.to_dict(), "session_epoch": doc.session_epoch})
            for c in doc.comments
        ]
        old["warnings"] = [
            *result.get("warnings", []),
            {
                "code": "scope_body_only",
                "message": "only comments in the main document body are visible to the pane; comments "
                "anchored in headers, footers or text boxes are not listed",
            },
        ]
        return self._size(old)

    async def test_summary_and_full_are_much_smaller_than_the_old_full_shape(self):
        """Measured on the 27-comment fixture above (compact JSON chars; ~chars/4 tokens):
        old full 19099, new full 11482 (60.1% of old: 23 chars over the 60% target), columnar
        summary 2308 (keeps `path`; drops summary-only diagnostics). Closing the last 23 chars of full would mean dropping
        information. What remains in full's comments array (~10997): anchor_paragraph_text 3451,
        content 2718, anchor_text 1354, created_time 1134, comment_id 828, author 648, resolved 459,
        replies 350; summary rows ~1997 (anchor 844, comment_id 477, author 405, resolved 162,
        reply_count 54). Summary omits `pane` (and keeps `path`); both modes omit null/empty top-level keys."""
        doc = twenty_seven_comment_document()
        await self.connect_pane(document=doc)
        summary = await self.call(
            lambda: comments_live.execute_list_open_items_live(str(self.target), fields="summary")
        )
        full = await self.call(lambda: comments_live.execute_list_open_items_live(str(self.target)))
        old = await self._old_shape_size(doc)
        self.assertEqual(len(summary["comment_rows"]), 27)
        print(f"\n#111 sizes (chars): old full={old} new full={self._size(full)} summary={self._size(summary)}")
        self.assertLessEqual(self._size(full), old * 0.605)
        self.assertLess(self._size(summary), 2400)
        for gone in ("pane", "warnings", "file_only_open_comments", "suggestion_columns", "comments", "suggestions"):
            self.assertNotIn(gone, summary)
        self.assertIn("path", summary)
        for kept in ("pane", "path"):
            self.assertIn(kept, full)


class ListOpenItemsLiveSuggestionScopeTests(LiveCommentsTestCase):
    """Issue #111 review: a live suggestion's paragraph_context carries w:delText, which the
    projection-based paragraph index does not; scoping must still place it."""

    fixture_name = "revision/tracked.docx"

    def _patched_revisions(self, revisions):
        from verified_docx_mcp.live import revisions_live

        return unittest.mock.patch.object(
            revisions_live, "list_revisions",
            lambda session: {"revisions": revisions, "coverage": "body" if revisions is not None else "unavailable",
                             "reason": None},
        )

    async def test_whole_document_scope_places_insertion_and_deletion_in(self):
        from verified_docx_mcp.live import revisions_live

        flat = docx_to_flat_opc(self.target)
        inventory = revisions_live._markup_revisions(flat, "ep1")
        self.assertEqual({r["markup_type"] for r in inventory}, {"ins", "del"})
        self.assertIn("he re", inventory[1]["paragraph_context"][0])  # deleted text is in the context
        await self.connect_pane(document=FakeDocument(text="", author_name="Pat"), body_ooxml=flat)
        with self._patched_revisions(inventory):
            result = await self.call(
                lambda: comments_live.execute_list_open_items_live(
                    str(self.target), within_text="vision-token base fixture"
                )
            )
        self.assertEqual(len(result["pending_suggestions"]), 2)
        self.assertFalse(any(s.get("scope_uncertain") for s in result["pending_suggestions"]))
        self.assertEqual(result["counts"]["in_scope"]["suggestions"], 2)
        self.assertIs(result["scoped"], True)

    async def test_in_scope_suggestion_count_is_unknown_without_a_revision_inventory(self):
        flat = docx_to_flat_opc(self.target)
        await self.connect_pane(document=FakeDocument(text="", author_name="Pat"), body_ooxml=flat)
        with self._patched_revisions(None):
            result = await self.call(
                lambda: comments_live.execute_list_open_items_live(
                    str(self.target), within_text="vision-token base fixture"
                )
            )
        self.assertIsNone(result["counts"]["in_scope"]["suggestions"])
        self.assertIsNone(result["pending_suggestions"])


class CommentToolsWithListedLiveHandleTests(LiveCommentsTestCase):
    """Issue #111 review: the MCP tools (not just execute_*) accept the `live:<epoch>:<id>` handle
    that list_open_items hands out."""

    fixture_name = "frag.docx"

    async def test_reply_and_resolve_tools_accept_a_listed_handle(self):
        from fastmcp import Client

        doc = FakeDocument(text="The quick brown fox jumps over the lazy dog.")
        doc.session_epoch = "ep9"
        await self.connect_pane(document=doc)
        await self.call(comments_live.execute_add_anchored_comment_live, str(self.target), "brown fox", "x", 1)
        async with Client(server.mcp) as client:
            listed = (
                await client.call_tool("list_open_items", {"path": str(self.target), "source": "live"})
            ).structured_content
            handle = listed["comments"][0]["comment_id"]
            self.assertTrue(handle.startswith("live:ep9:"))
            replied = await client.call_tool(
                "reply_to_comment",
                {"path": str(self.target), "comment_id": handle, "text": "ok", "write_mode": "live"},
                raise_on_error=False,
            )
            self.assertFalse(replied.is_error, replied)
            resolved = await client.call_tool(
                "resolve_comment",
                {"path": str(self.target), "comment_id": handle, "write_mode": "live"},
                raise_on_error=False,
            )
            self.assertFalse(resolved.is_error, resolved)
        self.assertTrue(doc.comments[0].resolved)
        self.assertEqual(len(doc.comments[0].replies), 1)
