"""Issue #117: content_sha12, a per-section digest over what sync acts on
(paragraph text, tracked deletions, table structure, anchored comments),
computed only server-side from OOXML -- the .docx in file mode, the pane's
body_ooxml export in live mode.

Fixture: tests/fixtures/content-digest.docx (Word-saved; see its README row).
Variants edit the XML of a temp copy, one concern at a time.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from collections.abc import Callable
from pathlib import Path
from types import SimpleNamespace

import pytest

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import docx_to_flat_opc
from test_live_revision_hash import LiveRevisionTestCase, pane_sections

from verified_docx_mcp import paths, projection, server
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import reads_live

FIXTURE = REPO / "tests" / "fixtures" / "content-digest.docx"
DOC = "word/document.xml"
COMMENTS = "word/comments.xml"
EXT = "word/commentsExtended.xml"
KEYS = ["overview-1", "background-1", "next-steps-1"]


def variant(dest: Path, edits: dict[str, Callable[[str], str]], src: Path = FIXTURE) -> Path:
    """Copy *src* to *dest*, rewriting the named parts' XML text."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in edits:
                text = data.decode("utf-8")
                changed = edits[item.filename](text)
                assert changed != text, f"edit to {item.filename} changed nothing"
                data = changed.encode("utf-8")
            zout.writestr(item, data)
    return dest


def sub1(old: str, new: str) -> Callable[[str], str]:
    def edit(text: str) -> str:
        assert text.count(old) == 1, (old, text.count(old))
        return text.replace(old, new)

    return edit


def digests(path: Path, **kwargs) -> tuple[dict[str, str], dict[str, str], str | None]:
    """({key: sha12}, {key: content_sha12}, preamble_content_sha12)."""
    sections, preamble = projection.find_sections_and_preamble(path, digests=True, content_digests=True, **kwargs)
    heads = [s for s in sections if s["kind"] == "heading"]
    return (
        {s["section_key"]: s["sha12"] for s in heads},
        {s["section_key"]: s["content_sha12"] for s in heads},
        preamble,
    )


def items_of(path: Path, key: str) -> list:
    return dict(projection.section_content_items(path)["sections"])[key]


class ContentDigestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.base = digests(FIXTURE)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def assert_only_content_changed(self, path: Path, changed: set[str], *, preamble_changed: bool = False) -> None:
        sha, content, preamble = digests(path)
        base_sha, base_content, base_preamble = self.base
        self.assertEqual(sha, base_sha, "no section's sha12 may change")
        self.assertEqual({k for k in KEYS if content[k] != base_content[k]}, changed)
        self.assertEqual(preamble != base_preamble, preamble_changed)


class CanonicalProjectionTests(ContentDigestCase):
    def test_fixture_items(self) -> None:
        overview = items_of(FIXTURE, "overview-1")
        self.assertEqual(
            overview,
            [
                ["p", "Overview"],
                ["p", "Some overview text."],
                ["c", "Michael Sutton", "Root comment one", False, [["Michael Sutton", "Reply to one"]], False],
            ],
        )
        background = items_of(FIXTURE, "background-1")
        self.assertEqual(background[:2], [["p", "Background"], ["p", "Background text."]])
        self.assertEqual(
            background[2],
            [
                "tbl",
                [
                    [[1, "none", []], [1, "none", []]],
                    [[1, "restart", []], [1, "none", []]],
                    [[1, "continue", []], [1, "none", [[[[1, "none", []]], [[1, "none", []]]]]]],
                    [[2, "none", []]],
                ],
            ],
        )
        self.assertEqual(
            background[-1],
            ["c", "Michael Sutton", "Root comment two", True, [["Michael Sutton", "Reply to two"]], False],
        )
        self.assertEqual(sum(1 for item in background if item[0] == "tbl"), 1, "nested tables never at top level")
        self.assertEqual(
            items_of(FIXTURE, "next-steps-1"),
            [["p", "Next Steps"], ["p", "Next phase text."], ["del", 1, len("Next phase"), "steps"]],
        )

    def test_hash_is_the_documented_canonical_json(self) -> None:
        import hashlib
        import json

        content = projection.section_content_items(FIXTURE)
        expected = {
            key: hashlib.sha256(
                json.dumps(items, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
            ).hexdigest()[:12]
            for key, items in content["sections"]
        }
        self.assertEqual(self.base[1], expected)
        self.assertEqual(projection.section_content_digests(FIXTURE)["sections"], [[k, expected[k]] for k in KEYS])

    def test_find_sections_impl_shape_is_unchanged_without_content_digests(self) -> None:
        plain = projection.find_sections_impl(FIXTURE, digests=True)
        self.assertTrue(all("content_sha12" not in s for s in plain))

    def test_deleted_span_offset_is_recorded(self) -> None:
        proj = projection.project_part(FIXTURE)
        (span,) = proj.deleted_spans
        self.assertEqual((span.text, span.offset), ("steps", len("Next phase")))


class SingleConcernEditTests(ContentDigestCase):
    def test_comment_text_only(self) -> None:
        path = variant(self.tmp / "a.docx", {COMMENTS: sub1("Root comment one", "Root comment 1")})
        self.assert_only_content_changed(path, {"overview-1"})

    def test_comment_author_only(self) -> None:
        def edit(text: str) -> str:
            return text.replace('w:id="2" w:author="Michael Sutton"', 'w:id="2" w:author="Someone Else"', 1)

        path = variant(self.tmp / "a.docx", {COMMENTS: edit})
        self.assert_only_content_changed(path, {"background-1"})

    def test_reply_text_only(self) -> None:
        path = variant(self.tmp / "a.docx", {COMMENTS: sub1("Reply to two", "Reply to 2")})
        self.assert_only_content_changed(path, {"background-1"})

    def test_resolved_state_only(self) -> None:
        path = variant(
            self.tmp / "a.docx",
            {EXT: sub1('w15:paraId="3FADA680" w15:done="0"', 'w15:paraId="3FADA680" w15:done="1"')},
        )
        self.assert_only_content_changed(path, {"overview-1"})

    def test_a_replys_own_done_flag_does_not_govern_the_thread(self) -> None:
        path = variant(
            self.tmp / "a.docx",
            {EXT: sub1('w15:paraId="667E31A7" w15:paraIdParent="3FADA680" w15:done="0"',
                       'w15:paraId="667E31A7" w15:paraIdParent="3FADA680" w15:done="1"')},
        )
        self.assert_only_content_changed(path, set())

    def test_grid_span_only(self) -> None:
        path = variant(self.tmp / "a.docx", {DOC: sub1('<w:gridSpan w:val="2"/>', '<w:gridSpan w:val="1"/>')})
        self.assert_only_content_changed(path, {"background-1"})

    def test_vmerge_only(self) -> None:
        def edit(text: str) -> str:
            return text.replace('<w:vMerge w:val="restart"/>', "").replace("<w:vMerge/>", "")

        path = variant(self.tmp / "a.docx", {DOC: edit})
        self.assert_only_content_changed(path, {"background-1"})

    def test_nested_table_structure_only(self) -> None:
        # Two one-cell rows become one two-cell row: same paragraphs, same order.
        def edit(text: str) -> str:
            out, n = re.subn(
                r'(Nested A</w:t></w:r></w:p></w:tc>)</w:tr><w:tr [^>]*>(<w:tc>)', r"\1\2", text
            )
            assert n == 1
            return out

        path = variant(self.tmp / "a.docx", {DOC: edit})
        self.assertEqual(
            projection.paragraph_texts(projection.project_part(path)),
            projection.paragraph_texts(projection.project_part(FIXTURE)),
        )
        self.assert_only_content_changed(path, {"background-1"})

    def test_tracked_deletion_text_only(self) -> None:
        path = variant(self.tmp / "a.docx", {DOC: sub1("<w:delText>steps</w:delText>", "<w:delText>stops</w:delText>")})
        self.assert_only_content_changed(path, {"next-steps-1"})

    def test_tracked_deletion_position_only(self) -> None:
        # Move the w:del before the w:ins: same live text, deletion now at offset 5.
        def edit(text: str) -> str:
            m = re.search(r"<w:del w:id=\"6\".*?</w:del>", text)
            assert m
            deletion = m.group(0)
            text = text.replace(deletion, "", 1)
            return text.replace('<w:ins w:id="4"', deletion + '<w:ins w:id="4"', 1)

        path = variant(self.tmp / "a.docx", {DOC: edit})
        self.assertEqual(items_of(path, "next-steps-1")[-1], ["del", 1, len("Next "), "steps"])
        self.assert_only_content_changed(path, {"next-steps-1"})

    def test_a_deletion_split_across_runs_digests_like_one(self) -> None:
        split_runs = sub1(
            '<w:r w:rsidDel="00812B71"><w:delText>steps</w:delText></w:r>',
            '<w:r w:rsidDel="00812B71"><w:delText>st</w:delText></w:r>'
            '<w:r w:rsidDel="00812B71"><w:rPr><w:b/></w:rPr><w:delText>eps</w:delText></w:r>',
        )
        path = variant(self.tmp / "a.docx", {DOC: split_runs})
        self.assertEqual(len(projection.project_part(path).deleted_spans), 2)
        self.assertEqual(digests(path), self.base)

    def test_a_deletion_split_across_del_elements_digests_like_one(self) -> None:
        split_dels = sub1(
            '<w:r w:rsidDel="00812B71"><w:delText>steps</w:delText></w:r></w:del>',
            '<w:r w:rsidDel="00812B71"><w:delText>ste</w:delText></w:r></w:del>'
            '<w:del w:id="66" w:author="Other" w:date="2026-10-30T00:00:00Z">'
            '<w:r><w:delText>ps</w:delText></w:r></w:del>',
        )
        path = variant(self.tmp / "a.docx", {DOC: split_dels})
        self.assertEqual(digests(path), self.base)


class NoiseTests(ContentDigestCase):
    def test_dates_ids_rsids_paraids_proof_bookmarks_change_nothing(self) -> None:
        # Comment ids renumbered consistently (body markers + comments.xml).
        def renumber(text: str) -> str:
            return re.sub(
                r'(<w:(?:commentRangeStart|commentRangeEnd|commentReference|comment) w:id=")(\d)"',
                lambda m: f'{m.group(1)}{int(m.group(2)) + 40}"',
                text,
            )

        def body(text: str) -> str:
            text = renumber(text)
            text = re.sub(r'w:rsid(\w*)="[0-9A-F]{8}"', r'w:rsid\1="00ABCDEF"', text)
            text = text.replace('w14:paraId="5A43F5BC"', 'w14:paraId="1234ABCD"')
            text = text.replace('w14:textId="77777777"', 'w14:textId="11111111"')
            text = text.replace(
                "<w:r><w:t>Some overview text.</w:t></w:r>",
                '<w:bookmarkStart w:id="90" w:name="bm"/><w:proofErr w:type="spellStart"/>'
                "<w:r><w:lastRenderedPageBreak/><w:t>Some overview text.</w:t></w:r>"
                '<w:proofErr w:type="spellEnd"/><w:bookmarkEnd w:id="90"/>',
            )
            return text

        def comments(text: str) -> str:
            text = renumber(text)
            return text.replace('w:date="2026-10-29T20:06:00Z"', 'w:date="2030-01-01T00:00:00Z"')

        path = variant(self.tmp / "a.docx", {DOC: body, COMMENTS: comments})
        self.assertEqual(digests(path), self.base)

    def test_character_formatting_changes_nothing(self) -> None:
        path = variant(
            self.tmp / "a.docx",
            {DOC: sub1("<w:r><w:t>Some overview text.</w:t></w:r>",
                       "<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Some overview text.</w:t></w:r>")},
        )
        self.assertEqual(digests(path), self.base)


class AnchoringTests(ContentDigestCase):
    def _add_comment(self, text: str, w_id: str, para_id: str, *, parent: str | None = None) -> Callable[[str], str]:
        def edit(xml: str) -> str:
            comment = (
                f'<w:comment w:id="{w_id}" w:author="Reviewer" w:date="2026-10-29T20:06:00Z" w:initials="R">'
                f'<w:p w14:paraId="{para_id}" w14:textId="77777777"><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>'
            )
            return xml.replace("</w:comments>", comment + "</w:comments>")

        return edit

    def test_zero_width_anchor_is_placed(self) -> None:
        marker = '<w:commentRangeStart w:id="9"/><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r>'
        path = variant(
            self.tmp / "a.docx",
            {
                DOC: sub1("<w:r><w:t>Some overview text.</w:t></w:r>", "<w:r><w:t>Some overview text.</w:t></w:r>" + marker),
                COMMENTS: self._add_comment("Zero width", "9", "0A0B0C0D"),
            },
        )
        proj = projection.project_part(path)
        self.assertFalse(any("9" in e.comment_ids for e in proj.events if isinstance(e, projection.RunEvent)))
        self.assertIn(["c", "Reviewer", "Zero width", False, [], False], items_of(path, "overview-1"))
        self.assert_only_content_changed(path, {"overview-1"})

    def test_block_level_range_markers_are_placed(self) -> None:
        def edit(xml: str) -> str:
            xml = xml.replace("<w:tbl>", '<w:commentRangeStart w:id="9"/><w:tbl>', 1)
            return xml.replace("</w:tbl><w:p", '</w:tbl><w:commentRangeEnd w:id="9"/><w:p', 1)

        path = variant(self.tmp / "a.docx", {DOC: edit, COMMENTS: self._add_comment("On the table", "9", "0A0B0C0D")})
        self.assertIn(["c", "Reviewer", "On the table", False, [], False], items_of(path, "background-1"))
        self.assert_only_content_changed(path, {"background-1"})

    def test_a_comment_spanning_two_sections_counts_in_both(self) -> None:
        def edit(xml: str) -> str:
            xml = xml.replace('<w:commentRangeEnd w:id="2"/>', "", 1)
            return xml.replace(
                '<w:r><w:t xml:space="preserve"> text.</w:t></w:r>',
                '<w:r><w:t xml:space="preserve"> text.</w:t></w:r><w:commentRangeEnd w:id="2"/>',
                1,
            )

        path = variant(self.tmp / "a.docx", {DOC: edit})
        root_two = ["c", "Michael Sutton", "Root comment two", True, [["Michael Sutton", "Reply to two"]], False]
        self.assertIn(root_two, items_of(path, "background-1"))
        self.assertIn(root_two, items_of(path, "next-steps-1"))
        self.assertNotIn(root_two, items_of(path, "overview-1"))

    def test_orphan_reply_is_its_own_root(self) -> None:
        def drop_root(xml: str) -> str:
            out, n = re.subn(r'<w:comment w:id="0" .*?</w:comment>', "", xml, flags=re.DOTALL)
            assert n == 1
            return out

        def drop_anchor(xml: str) -> str:
            return re.sub(r'<w:commentRange(?:Start|End) w:id="0"/>', "", xml)

        path = variant(self.tmp / "a.docx", {COMMENTS: drop_root, DOC: drop_anchor})
        overview = items_of(path, "overview-1")
        self.assertEqual(overview[-1], ["c", "Michael Sutton", "Reply to one", False, [], True])

    def test_a_replys_start_does_not_order_its_thread(self) -> None:
        path = variant(
            self.tmp / "a.docx",
            {DOC: sub1('<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>',
                       '<w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="0"/>')},
        )
        # Comment 1 is a reply (placed through its root), so only the root's start counts.
        self.assert_only_content_changed(path, set())

    def test_two_roots_order_by_range_start(self) -> None:
        """A second root in the same heading paragraph: whose commentRangeStart comes
        first decides the order of the two "c" items, and so the digest."""
        reference = '<w:r><w:commentReference w:id="9"/></w:r>'

        def second_root(start_before_root_one: bool) -> Callable[[str], str]:
            def edit(xml: str) -> str:
                start = '<w:commentRangeStart w:id="9"/>'
                anchor = '<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>'
                xml = xml.replace(anchor, start + anchor if start_before_root_one else anchor + start, 1)
                end = '<w:commentRangeEnd w:id="0"/>'
                return xml.replace(end, end + '<w:commentRangeEnd w:id="9"/>' + reference, 1)

            return edit

        first = variant(self.tmp / "first.docx", {DOC: second_root(True), COMMENTS: self._add_comment("Second root", "9", "0A0B0C0D")})
        later = variant(self.tmp / "later.docx", {DOC: second_root(False), COMMENTS: self._add_comment("Second root", "9", "0A0B0C0D")})
        texts = lambda path: [i[2] for i in items_of(path, "overview-1") if i[0] == "c"]
        self.assertEqual(texts(first), ["Second root", "Root comment one"])
        self.assertEqual(texts(later), ["Root comment one", "Second root"])
        self.assertNotEqual(digests(first)[1]["overview-1"], digests(later)[1]["overview-1"])
        self.assertEqual(digests(first)[0], digests(later)[0])

    def test_comment_anchored_inside_a_tracked_deletion_is_placed(self) -> None:
        deleted = (
            '<w:del w:id="80" w:author="R" w:date="2026-10-29T20:06:00Z"><w:commentRangeStart w:id="9"/>'
            '<w:r><w:delText>gone</w:delText></w:r><w:commentRangeEnd w:id="9"/>'
            '<w:r><w:commentReference w:id="9"/></w:r></w:del>'
        )
        body = sub1("<w:r><w:t>Some overview text.</w:t></w:r>", "<w:r><w:t>Some overview text.</w:t></w:r>" + deleted)
        path = variant(self.tmp / "a.docx", {DOC: body, COMMENTS: self._add_comment("In a deletion", "9", "0A0B0C0D")})
        self.assertIn(["c", "Reviewer", "In a deletion", False, [], False], items_of(path, "overview-1"))
        edited = variant(self.tmp / "b.docx", {DOC: body, COMMENTS: self._add_comment("In a deletion!", "9", "0A0B0C0D")})
        base_sha, base_content, _ = digests(path)
        sha, content, _ = digests(edited)
        self.assertEqual(sha, base_sha)
        self.assertEqual({k for k in KEYS if content[k] != base_content[k]}, {"overview-1"})

    def test_reference_only_inside_a_deletion_is_placed(self) -> None:
        deleted = (
            '<w:del w:id="80" w:author="R" w:date="2026-10-29T20:06:00Z">'
            '<w:r><w:commentReference w:id="9"/></w:r></w:del>'
        )
        path = variant(
            self.tmp / "a.docx",
            {
                DOC: sub1("<w:r><w:t>Some overview text.</w:t></w:r>", "<w:r><w:t>Some overview text.</w:t></w:r>" + deleted),
                COMMENTS: self._add_comment("Reference only", "9", "0A0B0C0D"),
            },
        )
        self.assertIn(["c", "Reviewer", "Reference only", False, [], False], items_of(path, "overview-1"))

    def test_reply_order_follows_comments_xml_order(self) -> None:
        def add_second_reply(xml: str) -> str:
            reply = (
                '<w:comment w:id="7" w:author="B" w:date="2026-10-29T20:06:00Z" w:initials="B">'
                '<w:p w14:paraId="0E0E0E0E" w14:textId="77777777"><w:r><w:t>Second reply</w:t></w:r></w:p></w:comment>'
            )
            return xml.replace("</w:comments>", reply + "</w:comments>")

        def ext(xml: str) -> str:
            return xml.replace(
                "</w15:commentsEx>",
                '<w15:commentEx w15:paraId="0E0E0E0E" w15:paraIdParent="3FADA680" w15:done="0"/></w15:commentsEx>',
            )

        path = variant(self.tmp / "a.docx", {COMMENTS: add_second_reply, EXT: ext})
        self.assertEqual(
            items_of(path, "overview-1")[-1][4], [["Michael Sutton", "Reply to one"], ["B", "Second reply"]]
        )


class MalformedLinkageTests(ContentDigestCase):
    def assert_invalid(self, path: Path) -> None:
        with self.assertRaises(VerifyError) as ctx:
            projection.section_content_digests(path)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertEqual(ctx.exception.envelope.diagnostics.get("stage"), "content_digest")

    def test_parent_cycle(self) -> None:
        path = variant(
            self.tmp / "a.docx",
            {EXT: sub1('w15:paraId="3FADA680" w15:done="0"', 'w15:paraId="3FADA680" w15:paraIdParent="667E31A7" w15:done="0"')},
        )
        self.assert_invalid(path)

    def test_duplicate_comment_id(self) -> None:
        path = variant(self.tmp / "a.docx", {COMMENTS: sub1('<w:comment w:id="3"', '<w:comment w:id="2"')})
        self.assert_invalid(path)

    def test_ambiguous_parent_para_id(self) -> None:
        path = variant(self.tmp / "a.docx", {COMMENTS: sub1('w14:paraId="37317D53"', 'w14:paraId="3FADA680"')})
        self.assert_invalid(path)

    def test_non_numeric_grid_span(self) -> None:
        path = variant(self.tmp / "a.docx", {DOC: sub1('<w:gridSpan w:val="2"/>', '<w:gridSpan w:val="two"/>')})
        self.assert_invalid(path)


class ThreadingScaleTests(unittest.TestCase):
    def test_three_thousand_comments_thread_in_under_a_second(self) -> None:
        """Review of #117: the known-paraId lookup is built once, not scanned per comment."""
        import time
        import xml.etree.ElementTree as ET

        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
        w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
        w16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
        comments = ET.Element(f"{{{w}}}comments")
        ext = ET.Element(f"{{{w15}}}commentsEx")
        ids = ET.Element(f"{{{w16cid}}}commentsIds")
        n = 3000
        for i in range(n):
            pid = f"{i + 1:08X}"
            c = ET.SubElement(comments, f"{{{w}}}comment", {f"{{{w}}}id": str(i), f"{{{w}}}author": "A"})
            p = ET.SubElement(c, f"{{{w}}}p", {f"{{{w14}}}paraId": pid})
            ET.SubElement(ET.SubElement(p, f"{{{w}}}r"), f"{{{w}}}t").text = f"comment {i}"
            attrs = {f"{{{w15}}}paraId": pid, f"{{{w15}}}done": "0"}
            if i % 2:
                attrs[f"{{{w15}}}paraIdParent"] = f"{i:08X}"  # reply to the previous root
            ET.SubElement(ext, f"{{{w15}}}commentEx", attrs)
            ET.SubElement(ids, f"{{{w16cid}}}commentId", {f"{{{w16cid}}}paraId": pid, f"{{{w16cid}}}durableId": pid})
        t0 = time.perf_counter()
        threads = projection._content_comment_threads(comments, ext, ids)
        elapsed = time.perf_counter() - t0
        self.assertEqual(len(threads), n // 2)
        self.assertTrue(all(len(item[3]) == 1 for item in threads.values()))
        self.assertLess(elapsed, 1.0)


class PreambleTests(ContentDigestCase):
    def test_preamble_covers_content_before_the_first_heading(self) -> None:
        self.assertEqual(projection.section_content_items(FIXTURE)["preamble"], [])
        path = variant(
            self.tmp / "a.docx",
            {DOC: sub1("<w:body>", "<w:body><w:p><w:r><w:t>Cover note</w:t></w:r></w:p>")},
        )
        self.assertEqual(projection.section_content_items(path)["preamble"], [["p", "Cover note"]])
        self.assert_only_content_changed(path, set(), preamble_changed=True)

    def test_headingless_document_is_all_preamble(self) -> None:
        def edit(xml: str) -> str:
            return re.sub(r'<w:pStyle w:val="Heading\d"/>', "", xml)

        path = variant(self.tmp / "a.docx", {DOC: edit})
        sha, content, _preamble = digests(path)
        self.assertEqual((sha, content), ({}, {}))
        items = projection.section_content_items(path)["preamble"]
        self.assertEqual(items[0], ["p", "Overview"])
        self.assertEqual(sum(1 for item in items if item[0] == "c"), 2)
        self.assertEqual(sum(1 for item in items if item[0] == "tbl"), 1)

    def test_ignore_trailing_empty_drops_words_appended_paragraph(self) -> None:
        path = variant(self.tmp / "a.docx", {DOC: sub1("<w:sectPr", "<w:p/><w:sectPr")})
        self.assertNotEqual(digests(path)[1]["next-steps-1"], self.base[1]["next-steps-1"])
        self.assertEqual(digests(path, ignore_trailing_empty=True), self.base)


class FindSectionsToolTests(ContentDigestCase):
    def setUp(self) -> None:
        super().setUp()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(self.tmp)
        self.doc = self.tmp / "content-digest.docx"
        shutil.copyfile(FIXTURE, self.doc)

    def tearDown(self) -> None:
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        super().tearDown()

    def test_digests_true_adds_content_sha12_and_preamble(self) -> None:
        result = server.execute_find_sections(str(self.doc), digests=True)
        heads = [s for s in result["sections"] if s["kind"] == "heading"]
        self.assertEqual({s["section_key"]: s["content_sha12"] for s in heads}, self.base[1])
        self.assertEqual({s["section_key"]: s["sha12"] for s in heads}, self.base[0])
        self.assertEqual(result["preamble_content_sha12"], self.base[2])

    def test_uncanonicalizable_content_keeps_sha12_and_warns(self) -> None:
        variant(self.tmp / "dup.docx", {COMMENTS: sub1('<w:comment w:id="3"', '<w:comment w:id="2"')})
        result = server.execute_find_sections(str(self.tmp / "dup.docx"), digests=True)
        heads = [s for s in result["sections"] if s["kind"] == "heading"]
        self.assertEqual({s["section_key"]: s["sha12"] for s in heads}, self.base[0])
        self.assertTrue(all(s["content_sha12"] is None for s in heads))
        self.assertIsNone(result["preamble_content_sha12"])
        self.assertIn("content_digest_unavailable", result["warnings"])
        self.assertIn("w:id", result["content_digest_unavailable"]["reason"])
        self.assertEqual(result["content_digest_unavailable"]["diagnostics"]["stage"], "content_digest")

    def test_digests_false_adds_nothing(self) -> None:
        result = server.execute_find_sections(str(self.doc))
        self.assertNotIn("preamble_content_sha12", result)
        self.assertTrue(all("content_sha12" not in s for s in result["sections"]))


class LiveContentDigestTests(LiveRevisionTestCase):
    fixture_name = "content-digest.docx"

    async def connect_live(self, **kwargs):
        kwargs.setdefault("body_ooxml", docx_to_flat_opc(FIXTURE))
        kwargs.setdefault("sections", pane_sections(FIXTURE))
        return await self.connect_pane(**kwargs)

    def file_side(self):
        return digests(FIXTURE)

    async def test_live_find_sections_equals_file_mode(self) -> None:
        await self.connect_live()
        live = await self.call(server.execute_find_sections, str(self.target), source="live", digests=True)
        self.assertEqual(live["source"], "live")
        heads = [s for s in live["sections"] if s["kind"] == "heading"]
        sha, content, preamble = self.file_side()
        self.assertEqual({s["section_key"]: s["content_sha12"] for s in heads}, content)
        self.assertEqual({s["section_key"]: s["sha12"] for s in heads}, sha)
        self.assertEqual(live["preamble_content_sha12"], preamble)

    async def test_live_with_words_appended_paragraph_equals_file_mode(self) -> None:
        padded = variant(Path(self._tmp.name) / "padded.docx", {DOC: sub1("<w:sectPr", "<w:p/><w:sectPr")})
        pane = await self.connect_live(body_ooxml=docx_to_flat_opc(padded))
        result = await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        sha, content, preamble = self.file_side()
        self.assertEqual(result["content_sections"], [[k, content[k]] for k in KEYS])
        self.assertEqual(result["sections"], [[k, sha[k]] for k in KEYS])
        self.assertEqual(result["preamble_content_sha12"], preamble)
        self.assertEqual(pane.body_ooxml_requests, 1)

    async def test_live_revision_content_digests_is_one_body_ooxml_export(self) -> None:
        pane = await self.connect_live()
        result = await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        self.assertEqual(pane.body_ooxml_requests, 1)
        self.assertEqual(pane.body_hash_requests, 0)
        self.assertEqual(result["via"], "body_ooxml")
        self.assertEqual(result["revision"], f"live:sha256:{pane.document.sha256()}")
        self.assertEqual(result["body_sha12"], pane.document.sha256()[:12])
        sha, content, preamble = self.file_side()
        self.assertEqual(result["sections"], [[k, sha[k]] for k in KEYS])
        self.assertEqual(result["content_sections"], [[k, content[k]] for k in KEYS])
        self.assertEqual(result["preamble_content_sha12"], preamble)

    async def test_live_revision_default_is_unchanged(self) -> None:
        pane = await self.connect_live()
        result = await self.call(server.live_revision, str(self.target))
        self.assertEqual(result["via"], "body_hash")
        self.assertEqual(set(result), {"revision", "body_sha12", "sections", "observed_at", "via"})
        self.assertEqual((pane.body_hash_requests, pane.body_ooxml_requests), (1, 0))

    async def test_live_revision_content_digests_needs_body_ooxml(self) -> None:
        await self.connect_live(capabilities=["body_hash_v1"])
        with self.assertRaises(VerifyError) as ctx:
            await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_live_list_sections_content_digests_keeps_row_arity(self) -> None:
        pane = await self.connect_live(capabilities=["section_locks", "body_ooxml", "body_hash_v1"])
        compact = await self.call(server.live_list_sections, str(self.target), compact=True, content_digests=True)
        self.assertTrue(all(len(row) == 3 for row in compact["sections"]))
        sha, content, preamble = self.file_side()
        self.assertEqual([row[2] for row in compact["sections"]], [sha[k] for k in KEYS])
        self.assertEqual(compact["content_sections"], [[k, content[k]] for k in KEYS])
        self.assertEqual(compact["preamble_content_sha12"], preamble)
        self.assertEqual((pane.sections_list_requests, pane.body_ooxml_requests), (1, 1))

        full = await self.call(server.live_list_sections, str(self.target), content_digests=True)
        self.assertEqual([s["section_key"] for s in full["sections"]], KEYS)
        self.assertEqual(full["content_sections"], compact["content_sections"])

        plain = await self.call(server.live_list_sections, str(self.target), compact=True)
        self.assertEqual(set(plain), {"sections", "body_sha12"})
        self.assertNotIn("warnings", compact)

    async def test_live_list_sections_reports_a_key_mismatch_never_pairs_silently(self) -> None:
        pane_view = pane_sections(FIXTURE)
        pane_view[1] = {**pane_view[1], "section_key": "background-renamed-1"}
        await self.connect_live(capabilities=["section_locks", "body_ooxml"], sections=pane_view)
        compact = await self.call(server.live_list_sections, str(self.target), compact=True, content_digests=True)
        self.assertIn("content_sections_key_mismatch", compact["warnings"])
        self.assertEqual(
            compact["content_sections_key_mismatch"],
            {"pane_only": ["background-renamed-1"], "server_only": ["background-1"], "order_differs": False},
        )
        self.assertEqual([k for k, _ in compact["content_sections"]], KEYS)  # the server's own keys

    async def test_export_fidelity_warnings_are_passed_through(self) -> None:
        flat = docx_to_flat_opc(FIXTURE)
        flat, n = re.subn(r'<pkg:part pkg:name="/word/styles\.xml".*?</pkg:part>', "", flat, flags=re.DOTALL)
        self.assertEqual(n, 1)
        await self.connect_live(capabilities=["section_locks", "body_ooxml"], body_ooxml=flat)
        revision = await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        self.assertIn("live_styles_missing", revision["warnings"])
        listed = await self.call(server.live_list_sections, str(self.target), compact=True, content_digests=True)
        self.assertIn("live_styles_missing", listed["warnings"])

    async def test_explicit_content_digests_still_raise_on_bad_threading(self) -> None:
        dup = variant(Path(self._tmp.name) / "dup.docx", {COMMENTS: sub1('<w:comment w:id="3"', '<w:comment w:id="2"')})
        await self.connect_live(body_ooxml=docx_to_flat_opc(dup))
        with self.assertRaises(VerifyError) as ctx:
            await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        live = await self.call(server.execute_find_sections, str(self.target), source="live", digests=True)
        self.assertIn("content_digest_unavailable", live["warnings"])
        self.assertTrue(all(len(s["sha12"]) == 12 for s in live["sections"] if s["kind"] == "heading"))


# -- shared-bridge observe semantics ---------------------------------------------


class _OoxmlPane:
    """A broker-side pane that answers body_ooxml and sections_list (see
    test_shared_bridge.Pane for the shape the broker expects)."""

    document_name = "Proposal.docx"
    document_url = "https://example.test/site/Proposal.docx"
    connected_since = 1.0
    last_body_sha256 = "initial"
    heartbeat_count = 0
    hash_sample_count = 0

    def __init__(self) -> None:
        import time

        from verified_docx_mcp.live.protocol import HelloMessage

        self.body_hash_sampled_monotonic = time.monotonic()
        self.body_hash_state = "sampled"
        self.hello = HelloMessage(
            self.document_url, "Word", "Mac", {}, "initial",
            frozenset({"shared_queue", "section_locks", "body_ooxml"}), "pane-1",
        )
        self.calls: list[str] = []
        self.ooxml = docx_to_flat_opc(FIXTURE)
        self.sections = pane_sections(FIXTURE)

    def heartbeat_age(self) -> float:
        return 0

    def request_threadsafe(self, op, payload=None, **kwargs):
        from verified_docx_mcp.live import shared

        self.calls.append(op)
        if op == "body_ooxml":
            return {"ooxml": self.ooxml, "bodySha256": self.last_body_sha256, "strippedParts": []}
        if op == "sections_list":
            return {"sections": self.sections, "bodySha256": self.last_body_sha256}
        if op in shared.READ_OPS:
            return {"bodySha256": self.last_body_sha256}
        assert payload["expectedBodySha256"] == self.last_body_sha256
        self.last_body_sha256 = "changed"
        return {"pre": "initial", "post": "changed"}


@pytest.fixture
def shared_pane(monkeypatch, tmp_path):
    from verified_docx_mcp.live import bridge, shared, write_mode

    pane = _OoxmlPane()
    broker = shared.Broker(SimpleNamespace(get=lambda n: pane, list=lambda: [pane], collisions=list))
    a = broker.dispatch({"action": "attach"})["client_id"]
    b = broker.dispatch({"action": "attach"})["client_id"]

    def remote_for(client_id):
        client = SimpleNamespace(rpc=lambda action, **kw: broker.dispatch({"action": action, "client_id": client_id, **kw}))
        return shared.RemoteSession(client, broker.snapshot(pane))

    remote_b = remote_for(b)
    monkeypatch.setenv(paths._ALLOWED_FILE_ROOTS_ENV, str(tmp_path))
    monkeypatch.setattr(bridge, "current_registry", lambda: SimpleNamespace(get=lambda n: remote_b, collisions=list))
    monkeypatch.setattr(write_mode, "_check_session_identity", lambda path, session: None)
    remote_a = remote_for(a)
    remote_a.request_threadsafe("describe")
    remote_b.request_threadsafe("describe")
    remote_a.request_threadsafe("replace")  # b is now stale
    return pane, remote_b, str(tmp_path / "Proposal.docx")


def test_live_revision_content_digests_rearms_like_an_observed_read(shared_pane):
    from verified_docx_mcp.live.session import LiveOpFailed

    pane, remote_b, path = shared_pane
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        remote_b.request_threadsafe("replace")
    result = reads_live.live_revision(path, content_digests=True)
    assert result["via"] == "body_ooxml"
    assert pane.calls.count("body_ooxml") == 1 and "body_hash" not in pane.calls
    remote_b.request_threadsafe("replace")  # re-armed by the observed body_ooxml read
    assert pane.calls[-1] == "replace"


def test_live_list_sections_content_digests_never_rearms(shared_pane):
    from verified_docx_mcp.live.session import LiveOpFailed

    pane, remote_b, path = shared_pane
    result = server.live_list_sections(path, compact=True, content_digests=True)
    assert len(result["content_sections"]) == len(KEYS)
    assert pane.calls.count("body_ooxml") == 1
    with pytest.raises(LiveOpFailed, match="LIVE_STALE"):
        remote_b.request_threadsafe("replace")
    assert pane.calls.count("replace") == 1  # only client a's


if __name__ == "__main__":
    unittest.main()
