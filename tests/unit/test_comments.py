"""Unit tests for src/verified_docx_mcp/comments.py (issue #28 WP-08):
add_anchored_comment/get_comment_thread, and the part-by-part comparison
against tests/fixtures/comments/golden-comment.docx the WP's own "Method"
section calls for -- same parts present, same relationship types, same
content types, same namespace declarations, equivalent element structure
(ignoring ids, dates, text). This file only ever READS the golden fixture
(a fresh copy per test, never the tracked file itself, and never written
to even in copy) -- its sha256 must never change; see
tests/fixtures/README.md and this module's own GoldenShaUnchangedTest.

Comparison scope: this module builds its OWN docx (via
add_anchored_comment against a plain, comment-free fixture) and compares
its five NEW comment-related parts against the golden's comment id=0
("comment.", the plainest -- unresolved, unthreaded-root -- shape the
golden carries) rather than id=1 (a reply -- WP-09 territory) or id=2
(resolved -- also WP-09). The comparison ignores id-shaped values (w:id,
paraId, durableId, textId, author, date/dateUtc, initials, rsid*,
providerId/userId -- the last two because this module cannot authentically
reproduce a real signed-in AD identity, a documented simplification in
comments.py's own _ensure_person) and all element TEXT content, per the
WP's own "ignoring ids, dates, text" instruction -- extended here to
author/initials/rsid*/presence-identity, which are exactly the same kind
of per-invocation-variable value in spirit.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import comments, mutations, paths, projection, tracked_changes
from verified_docx_mcp.errors import ErrorCode, VerifyError

FIXTURES = REPO / "tests" / "fixtures"
GOLDEN = FIXTURES / "comments" / "golden-comment.docx"
GOLDEN_SHA256 = "7839eddc95812fd0fb1b658a45085ec2738cd0f01c04c499f64be993e7dd03e6"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

# Attributes whose VALUE is ignored (per-invocation-variable: an id, a
# date, an author string) but whose PRESENCE is still required to match --
# a comment must always carry SOME author/date/id, just not a specific one.
_IGNORE_VALUE_ONLY = frozenset(
    {
        "id",
        "author",
        "date",
        "initials",
        "paraId",
        "paraIdParent",
        "textId",
        "durableId",
        "dateUtc",
        "userId",
        "providerId",
    }
)

# Attributes ignored ENTIRELY, including their mere presence: Word's own
# internal revision-session markers (rsid*), which this module never
# emits (documented in comments.py's own _append_comment) -- their
# absence is not a structural difference, just this module not tracking
# an rsid concept at all.
_IGNORE_COMPLETELY = frozenset({"rsidR", "rsidRDefault", "rsidP", "rsid"})


def _local(tag_or_attr: str) -> str:
    return tag_or_attr.rsplit("}", 1)[-1] if "}" in tag_or_attr else tag_or_attr


def _normalize(elem) -> tuple:
    """(tag, sorted (attr, value) pairs for non-ignored attrs, sorted set
    of ignored-attr LOCAL NAMES present, tuple of normalized children) --
    deliberately excludes .text entirely (WP-08: "ignoring ids, dates,
    text")."""
    attrs_kept = []
    ignored_keys_present = []
    for key, value in elem.attrib.items():
        local = _local(key)
        if local in _IGNORE_COMPLETELY:
            continue
        if local in _IGNORE_VALUE_ONLY:
            ignored_keys_present.append(local)
        else:
            attrs_kept.append((key, value))
    children = tuple(_normalize(c) for c in elem)
    return (elem.tag, tuple(sorted(attrs_kept)), tuple(sorted(ignored_keys_present)), children)


def assert_structurally_equivalent(test: unittest.TestCase, mine: ET.Element, golden: ET.Element, msg: str) -> None:
    test.assertEqual(_normalize(mine), _normalize(golden), msg)


def _declared_namespaces(xml_bytes: bytes) -> dict[str, str]:
    import io

    decls: dict[str, str] = {}
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        decls[prefix] = uri
    return decls


class GoldenShaUnchangedTest(unittest.TestCase):
    """Guards against ever accidentally re-saving/rewriting the golden
    fixture -- if this fails, something touched the file and the lead
    must re-author it (per tests/fixtures/README.md and this WP's own
    "never re-save or rewrite it")."""

    def test_golden_sha256_unchanged(self):
        actual = __import__("hashlib").sha256(GOLDEN.read_bytes()).hexdigest()
        self.assertEqual(actual, GOLDEN_SHA256, "golden-comment.docx has been modified -- do not touch it")


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "frag.docx"

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


class AddAnchoredCommentTests(_TempFixtureCase):
    def test_evidence_shape_and_range_brackets_the_quote(self):
        evidence = comments.execute_add_anchored_comment(str(self.target), "brown fox", "Nice color choice.", 1)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["rung"], "exact")
        self.assertEqual(evidence["comment_ids"], [evidence["comment_id"]])
        for key in ("applied", "match_count", "rung", "before", "after", "revision_before", "revision_after", "audit_logged"):
            self.assertIn(key, evidence)

        # "Verify after write that the range brackets the quoted span" --
        # re-read from disk (not the tool's own in-memory belief).
        proj = projection.project_part(self.target)
        by_id = {}
        for e in proj.events:
            if isinstance(e, projection.RunEvent):
                for cid in e.comment_ids:
                    by_id.setdefault(cid, []).append(e.text)
        # comment w:id is "0" for the first comment in a fresh document.
        self.assertEqual("".join(by_id.get("0", [])), "brown fox")

    def test_mid_run_quote_splits_without_losing_original_rpr(self):
        # "brown fox" straddles the bold "brown" run and the plain " fox"
        # tail of the third run -- the bold run must survive intact and
        # unsplit-unless-necessary.
        comments.execute_add_anchored_comment(str(self.target), "brown fox", "x", 1)
        runs = [r for r in projection.read_document_runs(self.target) if "rPr" in r]
        bold_runs = [r for r in runs if r["rPr"]["bold"]]
        self.assertEqual([r["text"] for r in bold_runs], ["brown"])
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")

    def test_quote_not_found_raises_zero_match_and_does_not_write(self):
        before_bytes = self.target.read_bytes()
        with self.assertRaises(VerifyError) as cm:
            comments.execute_add_anchored_comment(str(self.target), "nonexistent phrase", "x", 1)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.ZERO_MATCH)
        self.assertEqual(self.target.read_bytes(), before_bytes)

    def test_second_comment_reuses_existing_parts_and_increments_ids(self):
        first = comments.execute_add_anchored_comment(str(self.target), "brown", "first", 1)
        second = comments.execute_add_anchored_comment(str(self.target), "lazy", "second", 1)
        self.assertNotEqual(first["comment_id"], second["comment_id"])
        with zipfile.ZipFile(self.target) as zf:
            comments_root = ET.fromstring(zf.read("word/comments.xml"))
        ids = [projection._attr(c, "id") for c in comments_root if projection._ln(c) == "comment"]
        self.assertEqual(ids, ["0", "1"])


class GetCommentThreadTests(_TempFixtureCase):
    def test_reads_back_own_created_comment(self):
        evidence = comments.execute_add_anchored_comment(str(self.target), "lazy dog", "A comment.", 1)
        thread = comments.execute_get_comment_thread(str(self.target), evidence["comment_id"])
        self.assertEqual(thread["comment_id"], evidence["comment_id"])
        self.assertEqual(thread["content"], "A comment.")
        self.assertEqual(thread["quoted_text"], "lazy dog")
        self.assertFalse(thread["resolved"])
        self.assertEqual(thread["replies"], [])

    def test_unknown_comment_id_raises_invalid_input(self):
        comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        with self.assertRaises(VerifyError) as cm:
            comments.execute_get_comment_thread(str(self.target), "FFFFFFFF")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class GetCommentThreadAgainstGoldenTests(unittest.TestCase):
    """Reads (never writes) a fresh COPY of the golden fixture -- confirms
    get_comment_thread correctly resolves REAL Word-authored threading
    (a genuine reply) and a REAL resolved thread, neither of which this
    WP creates itself."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "golden.docx"
        shutil.copyfile(GOLDEN, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()
        # Belt and suspenders: assert the ORIGINAL golden file (not the
        # copy) is untouched by this test having run.
        actual = __import__("hashlib").sha256(GOLDEN.read_bytes()).hexdigest()
        assert actual == GOLDEN_SHA256, "golden-comment.docx was modified by a test -- this must never happen"

    def test_root_comment_with_real_reply(self):
        thread = comments.execute_get_comment_thread(str(self.target), "654503C0")
        self.assertEqual(thread["content"], "comment.")
        self.assertFalse(thread["resolved"])
        self.assertEqual(thread["reply_count"], 1)
        self.assertEqual(thread["replies"][0]["comment_id"], "7E7E1D27")
        self.assertEqual(thread["replies"][0]["content"], "reply.")

    def test_resolved_thread_with_no_replies(self):
        thread = comments.execute_get_comment_thread(str(self.target), "5CE3201C")
        self.assertEqual(thread["content"], "test and resolve")
        self.assertTrue(thread["resolved"])
        self.assertEqual(thread["reply_count"], 0)

    def test_list_open_items_then_get_comment_thread_chain_succeeds(self):
        # The exact caller chain WP-11b's resolve-gdoc-comments-on-docx
        # will do -- list the open comments, then fetch one's thread by
        # the comment_id list_open_items itself just handed back. This is
        # the regression test for the comment_id interop bug found in
        # review: list_open_items used to expose the raw w:id (e.g. "0"),
        # while get_comment_thread only ever accepts a durableId (e.g.
        # "654503C0") -- the chain failed outright until
        # tracked_changes._parse_comments was fixed to expose durableId as
        # comment_id (matching WP-08's own spec) with the raw w:id
        # available separately as w_id.
        open_items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertGreaterEqual(len(open_items["comments"]), 1)
        for summary in open_items["comments"]:
            thread = comments.execute_get_comment_thread(str(self.target), summary["comment_id"])
            self.assertEqual(thread["comment_id"], summary["comment_id"])
            self.assertEqual(thread["content"], summary["content"])


class PartByPartGoldenComparisonTests(_TempFixtureCase):
    """The WP-08 deliverable: add_anchored_comment's output compared PART
    BY PART to golden-comment.docx -- same parts present, same
    relationship types, same content types, same namespace declarations,
    equivalent element structure (ignoring ids, dates, text). Compared
    against golden's comment id=0 ("comment.") -- the plainest fresh,
    unresolved, unthreaded comment shape the golden carries (id=1 is a
    reply and id=2 is resolved; both WP-09 concerns)."""

    _COMMENT_PARTS = (
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/commentsExtensible.xml",
        "word/people.xml",
    )

    def setUp(self):
        super().setUp()
        comments.execute_add_anchored_comment(str(self.target), "brown fox", "comment.", 1)
        with zipfile.ZipFile(self.target) as zf:
            self.mine = {name: zf.read(name) for name in zf.namelist()}
        with zipfile.ZipFile(GOLDEN) as zf:
            self.golden = {name: zf.read(name) for name in zf.namelist()}

    def test_same_comment_parts_present(self):
        for part in self._COMMENT_PARTS:
            self.assertIn(part, self.mine, f"{part} missing from the generated package")

    def test_same_relationship_types_for_each_new_part(self):
        mine_rels = ET.fromstring(self.mine["word/_rels/document.xml.rels"])
        golden_rels = ET.fromstring(self.golden["word/_rels/document.xml.rels"])

        def type_by_target(root):
            out = {}
            for rel in root:
                target = rel.get("Target")
                out[target] = rel.get("Type")
            return out

        mine_by_target = type_by_target(mine_rels)
        golden_by_target = type_by_target(golden_rels)
        for part in self._COMMENT_PARTS:
            target = part.split("/", 1)[1]
            self.assertIn(target, mine_by_target, f"no relationship for {part}")
            self.assertEqual(
                mine_by_target[target], golden_by_target[target], f"relationship Type mismatch for {part}"
            )

    def test_same_content_type_overrides(self):
        mine_ct = ET.fromstring(self.mine["[Content_Types].xml"])
        golden_ct = ET.fromstring(self.golden["[Content_Types].xml"])

        def overrides_by_partname(root):
            return {
                child.get("PartName"): child.get("ContentType")
                for child in root
                if child.tag.rsplit("}", 1)[-1] == "Override"
            }

        mine_overrides = overrides_by_partname(mine_ct)
        golden_overrides = overrides_by_partname(golden_ct)
        for part in self._COMMENT_PARTS:
            part_name = f"/{part}"
            self.assertIn(part_name, mine_overrides)
            self.assertEqual(mine_overrides[part_name], golden_overrides[part_name])

    def test_same_namespace_declarations_per_part(self):
        # The namespace-declaration half of the comparison is the one
        # most likely to catch a real defect (PR #3's own bug: 31 of 35
        # root declarations silently dropped on re-serialization) --
        # explicitly comparing the DECLARED SET here, not assuming it
        # survived.
        for part in self._COMMENT_PARTS:
            mine_decls = _declared_namespaces(self.mine[part])
            golden_decls = _declared_namespaces(self.golden[part])
            self.assertEqual(mine_decls, golden_decls, f"namespace declaration mismatch in {part}")

    def test_equivalent_element_structure_ignoring_ids_dates_text(self):
        golden_comments_root = ET.fromstring(self.golden["word/comments.xml"])
        golden_comment_0 = next(
            c for c in golden_comments_root if projection._ln(c) == "comment" and projection._attr(c, "id") == "0"
        )
        mine_comments_root = ET.fromstring(self.mine["word/comments.xml"])
        mine_comment_0 = next(c for c in mine_comments_root if projection._ln(c) == "comment")
        assert_structurally_equivalent(self, mine_comment_0, golden_comment_0, "word/comments.xml's <w:comment> shape differs")

        golden_ext_root = ET.fromstring(self.golden["word/commentsExtended.xml"])
        golden_ext_0 = next(
            c
            for c in golden_ext_root
            if projection._ln(c) == "commentEx" and projection._attr(c, "paraId") == "4F84A13A"
        )
        mine_ext_root = ET.fromstring(self.mine["word/commentsExtended.xml"])
        mine_ext_0 = next(c for c in mine_ext_root if projection._ln(c) == "commentEx")
        assert_structurally_equivalent(self, mine_ext_0, golden_ext_0, "commentsExtended.xml's <w15:commentEx> shape differs")

        golden_ids_root = ET.fromstring(self.golden["word/commentsIds.xml"])
        golden_id_0 = next(c for c in golden_ids_root if projection._ln(c) == "commentId")
        mine_ids_root = ET.fromstring(self.mine["word/commentsIds.xml"])
        mine_id_0 = next(c for c in mine_ids_root if projection._ln(c) == "commentId")
        assert_structurally_equivalent(self, mine_id_0, golden_id_0, "commentsIds.xml's <w16cid:commentId> shape differs")

        golden_cex_root = ET.fromstring(self.golden["word/commentsExtensible.xml"])
        golden_cex_0 = next(c for c in golden_cex_root if projection._ln(c) == "commentExtensible")
        mine_cex_root = ET.fromstring(self.mine["word/commentsExtensible.xml"])
        mine_cex_0 = next(c for c in mine_cex_root if projection._ln(c) == "commentExtensible")
        assert_structurally_equivalent(
            self, mine_cex_0, golden_cex_0, "commentsExtensible.xml's <w16cex:commentExtensible> shape differs"
        )

        golden_people_root = ET.fromstring(self.golden["word/people.xml"])
        golden_person = next(c for c in golden_people_root if projection._ln(c) == "person")
        mine_people_root = ET.fromstring(self.mine["word/people.xml"])
        mine_person = next(c for c in mine_people_root if projection._ln(c) == "person")
        assert_structurally_equivalent(self, mine_person, golden_person, "people.xml's <w15:person> shape differs")

    def test_document_xml_anchor_shape(self):
        # commentRangeStart/End + the commentReference run's own rPr
        # shape (rStyle + sz + szCs), ignoring w:id.
        golden_doc = ET.fromstring(self.golden["word/document.xml"])
        mine_doc = ET.fromstring(self.mine["word/document.xml"])

        def anchor_tags(root):
            return [
                _local(el.tag)
                for el in root.iter()
                if _local(el.tag) in ("commentRangeStart", "commentRangeEnd", "commentReference")
            ]

        # Golden has 3 comments (2 overlapping anchors for the thread +
        # one standalone) -- compare shapes rather than counts; every
        # commentRangeStart has a matching commentRangeEnd and exactly
        # one commentReference each, in both packages.
        mine_tags = anchor_tags(mine_doc)
        self.assertEqual(mine_tags.count("commentRangeStart"), mine_tags.count("commentRangeEnd"))
        self.assertEqual(mine_tags.count("commentRangeStart"), mine_tags.count("commentReference"))
        golden_tags = anchor_tags(golden_doc)
        self.assertEqual(golden_tags.count("commentRangeStart"), 3)
        self.assertEqual(golden_tags.count("commentRangeEnd"), 3)
        self.assertEqual(golden_tags.count("commentReference"), 3)

        # Structural check on the commentReference run's OWN rPr siblings:
        # rStyle val + sz val + szCs val, ignoring w:id everywhere.
        def first_reference_rpr(root):
            for r in root.iter():
                if _local(r.tag) != "r":
                    continue
                if any(_local(c.tag) == "commentReference" for c in r):
                    for c in r:
                        if _local(c.tag) == "rPr":
                            return c
            return None

        mine_rpr = first_reference_rpr(mine_doc)
        golden_rpr = first_reference_rpr(golden_doc)
        self.assertIsNotNone(mine_rpr)
        self.assertIsNotNone(golden_rpr)
        assert_structurally_equivalent(self, mine_rpr, golden_rpr, "commentReference run's own rPr shape differs")


if __name__ == "__main__":
    unittest.main()
