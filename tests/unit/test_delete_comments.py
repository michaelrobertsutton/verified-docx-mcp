"""File-mode delete_comment / delete_resolved_comments (issue #80)."""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import comments, mutations, paths, projection
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {n: zf.read(n) for n in zf.namelist()}


def _ids(path: Path, part: str, tag: str, attr: str) -> list[str]:
    root = ET.fromstring(_parts(path)[part])
    return [projection._attr(c, attr) for c in root if projection._ln(c) == tag]


class _Base(unittest.TestCase):
    fixture = "frag.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "doc.docx"
        shutil.copyfile(FIXTURES / self.fixture, self.target)

    def tearDown(self):
        if self._old is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old
        self._tmp.cleanup()

    def add(self, quote="brown fox", text="root"):
        return comments.execute_add_anchored_comment(str(self.target), quote, text, 1)["comment_id"]

    def reply(self, parent, text="reply"):
        return comments.execute_reply_to_comment(str(self.target), parent, text)["comment_id"]

    def resolve(self, cid):
        comments.execute_resolve_comment(str(self.target), cid)

    def delete(self, ids, require_resolved=False):
        return comments.execute_delete_comments(
            str(self.target), ids, tool_name="delete_resolved_comments", require_resolved=require_resolved
        )

    def comment_count(self):
        return len(_ids(self.target, "word/comments.xml", "comment", "id"))

    def anchor_count(self, tag):
        return _parts(self.target)["word/document.xml"].decode().count(f"<w:{tag}")


class DeleteThreadTests(_Base):
    def test_registered_as_mutating(self):
        self.assertIn("delete_comment", MUTATING_TOOLS)
        self.assertIn("delete_resolved_comments", MUTATING_TOOLS)

    def test_deleting_root_removes_replies_nested_replies_and_all_entries(self):
        keep = self.add("lazy dog", "keep me")
        root = self.add("brown fox", "root")
        r1 = self.reply(root, "r1")
        self.reply(r1, "r1 child")
        self.resolve(root)
        self.assertEqual(self.comment_count(), 4)
        before = _parts(self.target)

        ev = self.delete([root], require_resolved=True)

        self.assertTrue(ev["applied"])
        self.assertEqual(ev["rung"], "delete")
        self.assertEqual(ev["deleted_thread_count"], 1)
        self.assertEqual(ev["deleted_reply_count"], 2)
        self.assertEqual(ev["comments_before"], 4)
        self.assertEqual(ev["comments_after"], 1)
        self.assertTrue(ev["other_comments_unchanged"])
        self.assertEqual(ev["before"], "brown fox")
        self.assertEqual(self.comment_count(), 1)
        for part, tag, attr in (
            ("word/commentsExtended.xml", "commentEx", "paraId"),
            ("word/commentsIds.xml", "commentId", "paraId"),
            ("word/commentsExtensible.xml", "commentExtensible", "durableId"),
        ):
            self.assertEqual(len(_ids(self.target, part, tag, attr)), 1, part)
        for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            self.assertEqual(self.anchor_count(tag), 1, tag)
        after = _parts(self.target)
        for untouched in ("word/people.xml", "[Content_Types].xml", "word/_rels/document.xml.rels"):
            self.assertEqual(before[untouched], after[untouched], untouched)
        thread = comments.execute_get_comment_thread(str(self.target), keep)
        self.assertEqual(thread["content"], "keep me")
        self.assertEqual(
            projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog."
        )

    def test_durable_id_and_w_id_alias_and_duplicates_delete_once(self):
        root = self.add()
        self.resolve(root)
        ev = self.delete([root, "0", root], require_resolved=True)
        self.assertEqual(ev["deleted_thread_count"], 1)
        self.assertEqual(len(ev["resolutions"]), 3)
        self.assertEqual(self.comment_count(), 0)

    def test_reply_id_is_refused_and_names_root(self):
        root = self.add()
        rep = self.reply(root)
        before = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            self.delete([rep])
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        offender = cm.exception.envelope.diagnostics["offenders"][0]
        self.assertEqual(offender["reason"], "reply")
        self.assertEqual(offender["root_comment_id"], root)
        self.assertEqual(self.target.read_bytes(), before)

    def test_unknown_unresolved_and_empty_refuse_without_writing(self):
        open_root = self.add()
        resolved_root = self.add("lazy dog", "x")
        self.resolve(resolved_root)
        before = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            self.delete([resolved_root, open_root, "FFFFFFFF"], require_resolved=True)
        reasons = {o["comment_id"]: o["reason"] for o in cm.exception.envelope.diagnostics["offenders"]}
        self.assertEqual(reasons, {open_root: "not_resolved", "FFFFFFFF": "unknown"})
        with self.assertRaises(VerifyError) as cm:
            self.delete([])
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(self.target.read_bytes(), before)

    def test_single_delete_does_not_require_resolved(self):
        root = self.add()
        ev = self.delete([root], require_resolved=False)
        self.assertEqual(ev["comment_id"], root)
        self.assertEqual(self.comment_count(), 0)

    def test_verification_failure_rolls_back(self):
        root = self.add()
        before = self.target.read_bytes()
        with mock.patch.object(mutations, "_strip_comment_anchors_recursive", lambda *a, **k: None):
            with self.assertRaises(VerifyError) as cm:
                self.delete([root])
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.VERIFICATION_FAILED)
        self.assertEqual(self.target.read_bytes(), before)

    def test_self_parented_entry_is_a_reply_and_refused_without_looping(self):
        root = self.add()
        parts = _parts(self.target)
        ext = ET.fromstring(parts["word/commentsExtended.xml"])
        w15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"
        for c in ext:
            c.set(w15 + "paraIdParent", projection._attr(c, "paraId"))
        tmp = self.target.with_suffix(".new")
        with zipfile.ZipFile(self.target) as zin, zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/commentsExtended.xml":
                    data = ET.tostring(ext, encoding="utf-8")
                zout.writestr(item, data)
        tmp.replace(self.target)
        with self.assertRaises(VerifyError) as cm:
            comments.execute_delete_comments(
                str(self.target), [root], tool_name="delete_comment", require_resolved=False,
                allow_concurrent_editor=True,
            )
        self.assertEqual(cm.exception.envelope.diagnostics["offenders"][0]["reason"], "reply")


class GoldenDeleteTests(_Base):
    fixture = "comments/golden-comment.docx"

    def test_delete_golden_resolved_thread_keeps_the_rest(self):
        self.assertEqual(sorted(_ids(self.target, "word/comments.xml", "comment", "id")), ["0", "1", "2"])
        before = _parts(self.target)
        ev = self.delete(["2"], require_resolved=True)
        self.assertTrue(ev["applied"])
        after = _parts(self.target)
        self.assertEqual(sorted(_ids(self.target, "word/comments.xml", "comment", "id")), ["0", "1"])
        self.assertEqual(before["word/people.xml"], after["word/people.xml"])
        thread = comments.execute_get_comment_thread(str(self.target), "0")
        self.assertEqual(thread["reply_count"], 1)


if __name__ == "__main__":
    unittest.main()
