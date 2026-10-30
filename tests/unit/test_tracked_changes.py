"""Unit tests for src/verified_docx_mcp/tracked_changes.py (issue #28
WP-07): list_open_items' Google-shaped output, the overlapping-write
TRACKED_CHANGES_PRESENT refusal (text_edit.py's guard, exercised here end
to end), accept/reject clearing the way for a subsequent write, and
REVISION_ID_NOT_FOUND.

Fixture provenance: tests/fixtures/revision/tracked.docx is Word-authored
(tests/fixtures/README.md) and already carries exactly two edits -- one
w:ins ("X" inserted) and one w:del ("he re" deleted), both with real
w:author/w:date -- so it satisfies this WP's "two edits" acceptance
requirement without a new fixture. tests/fixtures/revision/commented.docx
(also Word-authored, WP-03) covers the comments half of list_open_items'
Google-shaped output; comment reply THREADING is WP-08 scope (see that
fixture's own golden-comment.docx, untouched here) and is not tested.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import mutations, paths, projection, text_edit, tracked_changes
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "revision/tracked.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()


class MutatingToolsRegistrationTests(unittest.TestCase):
    def test_accept_reject_registered_list_open_items_is_not(self):
        self.assertIn("accept_tracked_changes", MUTATING_TOOLS)
        self.assertIn("reject_tracked_changes", MUTATING_TOOLS)
        self.assertNotIn("list_open_items", MUTATING_TOOLS, "list_open_items is read-only")


class ListOpenItemsTrackedChangesTests(_TempFixtureCase):
    """Acceptance: both edits in tracked.docx are listed, Google's shape."""

    def test_both_edits_listed_in_googles_shape(self):
        result = tracked_changes.execute_list_open_items(str(self.target))
        self.assertIn("comments", result)
        self.assertIn("pending_suggestions", result)
        suggestions = result["pending_suggestions"]
        self.assertEqual(len(suggestions), 2)

        by_kind = {s["kind"]: s for s in suggestions}
        self.assertEqual(set(by_kind), {"insertion", "deletion"})

        insertion = by_kind["insertion"]
        self.assertEqual(insertion["text"], "X")
        self.assertEqual(insertion["author"], "Michael Sutton")
        self.assertTrue(insertion["date"])
        self.assertIn("suggestion_id", insertion)
        self.assertNotIn("anchor_context", insertion)  # sent once per paragraph in anchor_contexts (#111)
        self.assertIn("revision-token base fixture", result["anchor_contexts"][insertion["para_ref"]])

        deletion = by_kind["deletion"]
        self.assertEqual(deletion["text"], "he re")
        self.assertEqual(deletion["author"], "Michael Sutton")


class ListOpenItemsCommentsTests(_TempFixtureCase):
    fixture_name = "revision/commented.docx"

    def test_comment_listed_in_googles_shape(self):
        result = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(len(result["comments"]), 1)
        comment = result["comments"][0]
        for key in ("comment_id", "content", "resolved", "quoted_text", "author"):
            self.assertIn(key, comment)
        for gone in ("scope", "modified_time", "reply_count", "replies"):  # constant/derivable/empty (#111)
            self.assertNotIn(gone, comment)
        self.assertEqual(comment["content"], "This is a test comment.")
        self.assertEqual(comment["quoted_text"], "This")
        self.assertFalse(comment["resolved"])
        self.assertEqual(result["pending_suggestions"], [])

    def test_comment_id_is_the_durable_id_not_the_raw_w_id(self):
        # WP-08 interop fix: comment_id must be commentsIds.xml's own
        # durableId (Google's shape: a durable, opaque id), not the raw
        # w:id -- which is only ever max-existing-plus-1 and therefore not
        # stable/round-trippable across edits. The raw w:id is still
        # available, under its own clearly-named field.
        result = tracked_changes.execute_list_open_items(str(self.target))
        comment = result["comments"][0]
        self.assertEqual(comment["comment_id"], "2A21A9D2")  # commented.docx's own real durableId
        self.assertEqual(comment["w_id"], "0")


def _add_reply(docx: Path, *, reply_done: str = "0", root_done: str = "0", extra_reply: bool = False) -> None:
    """Rewrite commented.docx's zip in place: add a reply (and optionally a
    reply-to-reply) to its one comment (paraId 79387F2F), no binary fixture."""
    import zipfile

    def comment(w_id: str, para: str, text: str, author: str) -> str:
        return (
            f'<w:comment w:id="{w_id}" w:author="{author}" w:date="2026-09-11T10:00:00Z" w:initials="X">'
            f'<w:p w14:paraId="{para}"><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>'
        )

    comments = comment("1", "1111AAAA", "first reply", "Alice")
    ex = f'<w15:commentEx w15:paraId="1111AAAA" w15:paraIdParent="79387F2F" w15:done="{reply_done}"/>'
    ids = '<w16cid:commentId w16cid:paraId="1111AAAA" w16cid:durableId="0000AAAA"/>'
    if extra_reply:
        comments += comment("2", "2222BBBB", "second reply", "Bob")
        ex += '<w15:commentEx w15:paraId="2222BBBB" w15:paraIdParent="79387F2F" w15:done="0"/>'
        ids += '<w16cid:commentId w16cid:paraId="2222BBBB" w16cid:durableId="0000BBBB"/>'
    with zipfile.ZipFile(docx) as zf:
        parts = {n: zf.read(n) for n in zf.namelist()}
        infos = zf.infolist()
    parts["word/comments.xml"] = parts["word/comments.xml"].replace(b"</w:comments>", comments.encode() + b"</w:comments>")
    ext = parts["word/commentsExtended.xml"].replace(b'w15:done="0"/>', f'w15:done="{root_done}"/>'.encode(), 1)
    parts["word/commentsExtended.xml"] = ext.replace(b"</w15:commentsEx>", ex.encode() + b"</w15:commentsEx>")
    parts["word/commentsIds.xml"] = parts["word/commentsIds.xml"].replace(
        b"</w16cid:commentsIds>", ids.encode() + b"</w16cid:commentsIds>"
    )
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as out:
        for info in infos:
            out.writestr(info.filename, parts[info.filename])


class ListOpenItemsThreadingTests(_TempFixtureCase):
    fixture_name = "revision/commented.docx"

    def test_reply_is_nested_under_root_not_listed_top_level(self):
        _add_reply(self.target)
        result = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(len(result["comments"]), 1)
        root = result["comments"][0]
        self.assertEqual(root["content"], "This is a test comment.")
        self.assertNotIn("reply_count", root)  # derivable: len(replies) (#111)
        self.assertEqual(
            root["replies"],
            [{"id": "0000AAAA", "author": "Alice", "date": "2026-09-11T10:00:00Z", "text": "first reply"}],
        )

    def test_replies_keep_comment_order_and_count_matches(self):
        _add_reply(self.target, extra_reply=True)
        root = tracked_changes.execute_list_open_items(str(self.target))["comments"][0]
        self.assertEqual(len(root["replies"]), 2)
        self.assertEqual([r["text"] for r in root["replies"]], ["first reply", "second reply"])

    def test_root_resolved_state_governs_the_thread(self):
        _add_reply(self.target, root_done="1", reply_done="0")
        self.assertEqual(tracked_changes.execute_list_open_items(str(self.target))["comments"], [])

    def test_resolved_reply_does_not_resolve_open_root(self):
        _add_reply(self.target, root_done="0", reply_done="1")
        comments = tracked_changes.execute_list_open_items(str(self.target))["comments"]
        self.assertEqual(len(comments), 1)
        self.assertFalse(comments[0]["resolved"])
        self.assertEqual(len(comments[0]["replies"]), 1)

    def test_unthreaded_comment_reports_zero_replies(self):
        root = tracked_changes.execute_list_open_items(str(self.target))["comments"][0]
        self.assertNotIn("replies", root)  # empty list omitted (#111)
        self.assertNotIn("reply_count", root)


class ListOpenItemsSummaryAndScopeTests(_TempFixtureCase):
    """Issue #111 (file mode): fields="summary", the trimmed full shape, section_key/within_text.

    sections.docx: Overview (p0-p1), Background (p2-p3), Next Steps (p4-p5). Comments sit in p1, p3
    and p5; tracked changes in p1, p4 and p5 (an insertion + a deletion each, so every paragraph's
    text would repeat per run in the old shape).
    """

    fixture_name = "sections.docx"

    def setUp(self):
        super().setUp()
        from verified_docx_mcp import comments

        path = str(self.target)
        comments.execute_add_anchored_comment(path, "overview", "c-overview", 1)
        comments.execute_add_anchored_comment(path, "Background text", "c-background", 1)
        comments.execute_add_anchored_comment(path, "steps text", "c-next", 1)
        text_edit.execute_replace_text(path, "Some ", "Any ", 1, track_changes=True)
        text_edit.execute_replace_text(path, "Next", "Upcoming", 2, track_changes=True)

    def _list(self, **kwargs):
        return tracked_changes.execute_list_open_items(str(self.target), **kwargs)

    def test_full_drops_constants_and_sends_each_paragraph_once(self):
        result = self._list()
        for comment in result["comments"]:
            self.assertNotIn("modified_time", comment)
            self.assertNotIn("scope", comment)
            # kept: JennyStack reads these in file mode
            for key in ("comment_id", "w_id", "quoted_text", "author", "resolved"):
                self.assertIn(key, comment)
            self.assertNotIn("reply_count", comment)
            self.assertNotIn("replies", comment)  # none of these has replies
        suggestions = result["pending_suggestions"]
        self.assertEqual(len(suggestions), 6)
        self.assertEqual(sorted(result["anchor_contexts"]), ["p1", "p4", "p5"])  # one entry per paragraph
        for suggestion in suggestions:
            self.assertNotIn("anchor_context", suggestion)
            self.assertIn(suggestion["para_ref"], result["anchor_contexts"])
        self.assertIn("overview text.", result["anchor_contexts"]["p1"])
        self.assertNotIn("counts", result)  # only when scoped or summary

    def test_summary_rows_and_counts(self):
        result = self._list(fields="summary")
        self.assertEqual(result["counts"], {"comments": 3, "suggestions": 6})
        self.assertEqual(result["comment_columns"], ["comment_id", "author", "resolved", "anchor", "reply_count"])
        self.assertEqual([c[3] for c in result["comments"]], ["overview", "Background text", "steps text"])
        self.assertEqual(result["suggestion_columns"], ["suggestion_id", "kind", "author", "anchor"])
        self.assertEqual(len(result["suggestions"][0]), 4)
        self.assertNotIn("pending_suggestions", result)
        full = self._list()
        self.assertLess(len(json.dumps(result)), len(json.dumps(full)))
        # summary ids are exactly the full ids
        self.assertEqual(
            [c[0] for c in result["comments"]], [c["comment_id"] for c in full["comments"]]
        )
        self.assertEqual(
            [s[0] for s in result["suggestions"]],
            [s["suggestion_id"] for s in full["pending_suggestions"]],
        )

    def test_section_key_scopes_comments_and_suggestions(self):
        result = self._list(section_key="background-1")
        self.assertEqual([c["content"] for c in result["comments"]], ["c-background"])
        self.assertEqual(result["pending_suggestions"], [])
        self.assertEqual(result["counts"], {"comments": 3, "suggestions": 6, "in_scope": {"comments": 1, "suggestions": 0}})

        result = self._list(section_key="overview-1")
        self.assertEqual([c["content"] for c in result["comments"]], ["c-overview"])
        self.assertEqual({s["para_ref"] for s in result["pending_suggestions"]}, {"p1"})
        self.assertEqual(result["counts"]["in_scope"], {"comments": 1, "suggestions": 2})

    def test_section_key_scope_applies_to_summary_too(self):
        result = self._list(fields="summary", section_key="upcoming-steps-1")
        self.assertEqual([c[3] for c in result["comments"]], ["steps text"])
        self.assertEqual(len(result["suggestions"]), 4)
        self.assertEqual(result["counts"]["in_scope"], {"comments": 1, "suggestions": 4})

    def test_within_text_scopes_to_the_matched_paragraph(self):
        result = self._list(within_text="overview text")
        self.assertEqual([c["content"] for c in result["comments"]], ["c-overview"])
        self.assertEqual({s["para_ref"] for s in result["pending_suggestions"]}, {"p1"})

    def test_section_key_and_within_text_intersect(self):
        result = self._list(section_key="background-1", within_text="overview text")
        self.assertEqual(result["comments"], [])
        self.assertEqual(result["counts"]["in_scope"], {"comments": 0, "suggestions": 0})

    def test_unknown_section_key_lists_available_keys(self):
        with self.assertRaises(VerifyError) as ctx:
            self._list(section_key="nope")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(
            ctx.exception.envelope.diagnostics["available_section_keys"],
            ["overview-1", "background-1", "upcoming-steps-1"],
        )

    def test_within_text_must_match_once(self):
        with self.assertRaises(VerifyError) as ctx:
            self._list(within_text="no such text anywhere")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.ZERO_MATCH)

    def test_unknown_fields_value_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            self._list(fields="brief")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class ListOpenItemsGoldenThreadingTests(_TempFixtureCase):
    fixture_name = "comments/golden-comment.docx"

    def test_golden_root_has_its_real_reply_and_resolved_thread_is_hidden(self):
        result = tracked_changes.execute_list_open_items(str(self.target))
        # golden: id=0 root (open) + id=1 reply; id=2 is a resolved root (filtered).
        self.assertEqual(len(result["comments"]), 1)
        root = result["comments"][0]
        self.assertEqual(len(root["replies"]), 1)
        self.assertEqual(root["replies"][0]["text"], "reply.")
        self.assertEqual(root["replies"][0]["id"] != root["comment_id"], True)


class OverlappingWriteRefusalTests(_TempFixtureCase):
    """Acceptance: overlapping replace_text refuses; accept clears it and
    the write then proceeds.

    Patches text_edit.resolve_author_name to a name DIFFERENT from
    tracked.docx's real Word-authored author ("Michael Sutton") for the
    whole class: WP-07b-a's own-author exclusion (text_edit.py's
    _check_tracked_changes_guard) would otherwise treat tracked.docx's
    revisions as this server's own prior work whenever this suite happens
    to run on Michael Sutton's own machine (resolve_author_name's own
    fallback path -- see test_author.py), silently defeating exactly the
    refusal these tests exist to prove. Pinning a fixed, different name
    makes the "foreign author" case deterministic regardless of whose
    machine runs the suite.
    """

    def setUp(self):
        super().setUp()
        patcher = mock.patch("verified_docx_mcp.text_edit.resolve_author_name", return_value="A Different Reviewer")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_replace_text_over_an_insertion_refuses_then_succeeds_after_accept(self):
        before_bytes = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            text_edit.execute_replace_text(str(self.target), "X", "Y", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.TRACKED_CHANGES_PRESENT)
        self.assertIn("crosses_revision", cm.exception.envelope.diagnostics.get("warnings", []))
        self.assertEqual(self.target.read_bytes(), before_bytes, "a refused write must not touch the file")

        accept_evidence = tracked_changes.execute_accept_tracked_changes(str(self.target))
        self.assertTrue(accept_evidence["applied"])
        self.assertEqual(accept_evidence["match_count"], 2)
        self.assertEqual(accept_evidence["rung"], "all")
        self.assertEqual(sorted(accept_evidence["revision_ids"]), ["0", "1"])

        # The write proceeds now that no tracked change remains.
        replace_evidence = text_edit.execute_replace_text(str(self.target), "X", "Y", 1)
        self.assertTrue(replace_evidence["applied"])
        self.assertIn("Y", projection.read_document_text(self.target))

    def test_force_bypasses_the_refusal(self):
        evidence = text_edit.execute_replace_text(str(self.target), "X", "Y", 1, force=True)
        self.assertTrue(evidence["applied"])
        self.assertIn("Y", projection.read_document_text(self.target))


class AcceptTests(_TempFixtureCase):
    def test_accept_all_makes_insertion_permanent_and_deletion_final(self):
        evidence = tracked_changes.execute_accept_tracked_changes(str(self.target))
        text = projection.read_document_text(self.target)
        # "X" (the accepted insertion) is now ordinary live text; "he re"
        # (the accepted deletion) stays gone -- both already true of the
        # LIVE projection even before accepting (w:ins content is live,
        # w:del content is excluded), so the real assertion is that no
        # tracked change remains afterward.
        self.assertIn("X", text)
        items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(items["pending_suggestions"], [])
        # issue #28 WP-10: layer 3's conflict-copy sweep result, wired
        # through accept/reject_tracked_changes' own evidence dict too.
        self.assertIn("conflict_copy_detected", evidence)
        self.assertFalse(evidence["conflict_copy_detected"])

    def test_accept_returns_bounded_evidence_and_keeps_the_full_text_in_the_audit_log(self):
        from verified_docx_mcp import audit

        evidence = tracked_changes.execute_accept_tracked_changes(str(self.target))
        self.assertNotIn("before", evidence)
        self.assertNotIn("after", evidence)
        for key in ("before_sha256", "after_sha256", "before_chars", "after_chars", "before_excerpt", "evidence_ref"):
            self.assertIn(key, evidence)
        record = audit.find_audit_record(evidence["evidence_ref"])
        self.assertEqual(record["tool"], "accept_tracked_changes")
        self.assertEqual(len(record["evidence"]["after"]), evidence["after_chars"])
        self.assertEqual(record["evidence"]["after"], projection.read_document_text(self.target))

    def test_accept_by_id_leaves_the_other_pending(self):
        evidence = tracked_changes.execute_accept_tracked_changes(str(self.target), revision_ids=["0"])
        self.assertEqual(evidence["rung"], "by_id")
        items = tracked_changes.execute_list_open_items(str(self.target))
        remaining = [s["suggestion_id"] for s in items["pending_suggestions"]]
        self.assertEqual(remaining, ["1"])

    def test_unknown_revision_id_raises_with_available_ids(self):
        with self.assertRaises(VerifyError) as cm:
            tracked_changes.execute_accept_tracked_changes(str(self.target), revision_ids=["nope"])
        env = cm.exception.envelope
        self.assertEqual(env.error_code, ErrorCode.REVISION_ID_NOT_FOUND)
        self.assertEqual(env.diagnostics["missing_ids"], ["nope"])
        self.assertEqual(sorted(env.diagnostics["available_ids"]), ["0", "1"])


class RejectTests(_TempFixtureCase):
    def test_reject_all_restores_original_text(self):
        tracked_changes.execute_reject_tracked_changes(str(self.target))
        text = projection.read_document_text(self.target)
        base_text = projection.read_document_text(FIXTURES / "revision" / "base.docx")
        self.assertEqual(text, base_text)
        items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(items["pending_suggestions"], [])


if __name__ == "__main__":
    unittest.main()
