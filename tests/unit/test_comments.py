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
from verified_docx_mcp.middleware import MUTATING_TOOLS

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


class MutatingToolsRegistrationTests(unittest.TestCase):
    def test_add_anchored_comment_and_wp09_tools_registered_get_comment_thread_is_not(self):
        self.assertIn("add_anchored_comment", MUTATING_TOOLS)
        self.assertIn("reply_to_comment", MUTATING_TOOLS)
        self.assertIn("resolve_comment", MUTATING_TOOLS)
        self.assertNotIn("get_comment_thread", MUTATING_TOOLS, "get_comment_thread is read-only")


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
        # issue #28 WP-10: layer 3's conflict-copy sweep result, wired
        # through add_anchored_comment's own evidence dict too.
        self.assertIn("conflict_copy_detected", evidence)
        self.assertFalse(evidence["conflict_copy_detected"])

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


# ---------------------------------------------------------------------------
# reply_to_comment / resolve_comment -- issue #28 WP-09.
# ---------------------------------------------------------------------------


class ReplyToCommentTests(_TempFixtureCase):
    def test_reply_links_to_parent_and_shares_its_anchor(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown fox", "root comment", 1)
        reply = comments.execute_reply_to_comment(str(self.target), root["comment_id"], "a reply")
        self.assertTrue(reply["applied"])
        self.assertEqual(reply["parent_comment_id"], root["comment_id"])
        self.assertNotEqual(reply["comment_id"], root["comment_id"])
        # issue #28 WP-10: layer 3's conflict-copy sweep result, wired
        # through reply_to_comment's own evidence dict too.
        self.assertIn("conflict_copy_detected", reply)
        self.assertFalse(reply["conflict_copy_detected"])
        # A reply never changes document text.
        self.assertEqual(projection.read_document_text(self.target), "The quick brown fox jumps over the lazy dog.")

        thread = comments.execute_get_comment_thread(str(self.target), root["comment_id"])
        self.assertEqual(thread["reply_count"], 1)
        self.assertEqual(thread["replies"][0]["comment_id"], reply["comment_id"])
        self.assertEqual(thread["replies"][0]["content"], "a reply")
        # The reply brackets the SAME live text as its parent.
        self.assertEqual(thread["replies"][0]["quoted_text"], "brown fox")

    def test_reply_gets_its_own_full_anchor_triplet_in_document_xml(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown fox", "root comment", 1)
        comments.execute_reply_to_comment(str(self.target), root["comment_id"], "a reply")
        with zipfile.ZipFile(self.target) as zf:
            doc = zf.read("word/document.xml").decode("utf-8")
        self.assertEqual(doc.count("<w:commentRangeStart"), 2)
        self.assertEqual(doc.count("<w:commentRangeEnd"), 2)
        self.assertEqual(doc.count("<w:commentReference"), 2)

    def test_unknown_parent_comment_id_raises_invalid_input(self):
        comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        with self.assertRaises(VerifyError) as cm:
            comments.execute_reply_to_comment(str(self.target), "FFFFFFFF", "a reply")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class ReplyStructuralComparisonAgainstGoldenTests(_TempFixtureCase):
    """Extends WP-08's part-by-part comparison to a generated REPLY,
    against the golden fixture's own real reply (paraId=5FA1F0F0,
    paraIdParent=4F84A13A) -- per this WP's own orchestrator notes."""

    def test_reply_commentex_shape_matches_golden_reply(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown fox", "root comment", 1)
        comments.execute_reply_to_comment(str(self.target), root["comment_id"], "a reply")

        with zipfile.ZipFile(self.target) as zf:
            mine_ext_root = ET.fromstring(zf.read("word/commentsExtended.xml"))
        with zipfile.ZipFile(GOLDEN) as zf:
            golden_ext_root = ET.fromstring(zf.read("word/commentsExtended.xml"))

        mine_reply_ex = next(
            c for c in mine_ext_root if projection._ln(c) == "commentEx" and projection._attr(c, "paraIdParent")
        )
        golden_reply_ex = next(
            c
            for c in golden_ext_root
            if projection._ln(c) == "commentEx" and projection._attr(c, "paraId") == "5FA1F0F0"
        )
        assert_structurally_equivalent(
            self, mine_reply_ex, golden_reply_ex, "a generated reply's <w15:commentEx> shape differs from the golden's real reply"
        )


class ResolveCommentTests(_TempFixtureCase):
    def test_resolve_sets_done_and_is_reflected_in_thread(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        evidence = comments.execute_resolve_comment(str(self.target), root["comment_id"])
        self.assertTrue(evidence["applied"])
        thread = comments.execute_get_comment_thread(str(self.target), root["comment_id"])
        self.assertTrue(thread["resolved"])
        # issue #28 WP-10: layer 3's conflict-copy sweep result, wired
        # through resolve_comment's own evidence dict too.
        self.assertIn("conflict_copy_detected", evidence)
        self.assertFalse(evidence["conflict_copy_detected"])

    def test_resolving_an_already_resolved_comment_is_idempotent(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        comments.execute_resolve_comment(str(self.target), root["comment_id"])
        evidence = comments.execute_resolve_comment(str(self.target), root["comment_id"])
        self.assertTrue(evidence["applied"])

    def test_unknown_comment_id_raises_invalid_input(self):
        comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        with self.assertRaises(VerifyError) as cm:
            comments.execute_resolve_comment(str(self.target), "FFFFFFFF")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)


class ListOpenItemsFilterResolvedTests(_TempFixtureCase):
    """Orchestrator note C: filtering resolved must not disturb the
    comment_id/w_id interop fix, and a resolved comment must still be
    fetchable by id even when filtered out of the open list."""

    def test_resolved_comment_excluded_from_list_open_items(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        comments.execute_add_anchored_comment(str(self.target), "lazy", "y", 1)
        comments.execute_resolve_comment(str(self.target), root["comment_id"])

        open_items = tracked_changes.execute_list_open_items(str(self.target))
        open_ids = {c["comment_id"] for c in open_items["comments"]}
        self.assertNotIn(root["comment_id"], open_ids, "a resolved comment must not appear in list_open_items")
        self.assertEqual(len(open_items["comments"]), 1)

    def test_resolved_comment_still_fetchable_by_id_even_though_filtered(self):
        root = comments.execute_add_anchored_comment(str(self.target), "brown", "x", 1)
        comments.execute_resolve_comment(str(self.target), root["comment_id"])

        open_items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertEqual(open_items["comments"], [], "the only comment is resolved -- the open list must be empty")

        # Filtered out of the OPEN list does not mean gone from the
        # package -- get_comment_thread must still resolve it by id.
        thread = comments.execute_get_comment_thread(str(self.target), root["comment_id"])
        self.assertEqual(thread["comment_id"], root["comment_id"])
        self.assertTrue(thread["resolved"])


class GoldenGoldenResolvedFilterTests(unittest.TestCase):
    """The golden fixture's own resolved thread (id=2, done=1) must be
    filtered from list_open_items but remain fetchable -- read-only,
    never writes to the golden."""

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
        actual = __import__("hashlib").sha256(GOLDEN.read_bytes()).hexdigest()
        assert actual == GOLDEN_SHA256, "golden-comment.docx was modified by a test -- this must never happen"

    def test_resolved_thread_excluded_but_still_fetchable(self):
        open_items = tracked_changes.execute_list_open_items(str(self.target))
        open_ids = {c["comment_id"] for c in open_items["comments"]}
        self.assertNotIn("5CE3201C", open_ids)
        self.assertIn("654503C0", open_ids)

        thread = comments.execute_get_comment_thread(str(self.target), "5CE3201C")
        self.assertTrue(thread["resolved"])
        self.assertEqual(thread["content"], "test and resolve")


# ---------------------------------------------------------------------------
# Issue #108: a multi-paragraph comment's identity is its LAST w:p's paraId
# (what Word itself keys commentsIds.xml/commentsExtended.xml on), not its
# first -- and every comment_id list_open_items can emit (durableId, or a
# raw w:id fallback) is accepted by get_comment_thread/reply_to_comment/
# resolve_comment.
# ---------------------------------------------------------------------------

MULTIPARA = FIXTURES / "comments" / "multipara-comment.docx"
MULTIPARA_SHA256 = "d3ad193433969b995895c1b407ed4ffb94baa720419a551220adb4cb68edfcf2"


class MultiparaFixtureShaUnchangedTest(unittest.TestCase):
    """Guards against ever accidentally re-saving/rewriting the
    multi-paragraph comment fixture -- see tests/fixtures/README.md's own
    provenance row for it."""

    def test_multipara_sha256_unchanged(self):
        actual = __import__("hashlib").sha256(MULTIPARA.read_bytes()).hexdigest()
        self.assertEqual(actual, MULTIPARA_SHA256, "multipara-comment.docx has been modified -- do not touch it")


class MultiParagraphCommentIdentityTests(_TempFixtureCase):
    """tests/fixtures/comments/multipara-comment.docx (Word 16.112.4, see
    tests/fixtures/README.md's provenance row): w:id="0" has two
    paragraphs (paraIds 5C83D4AC then 53D84979 -- Word keyed
    commentsIds.xml/commentsExtended.xml on 53D84979 only, the LAST one);
    w:id="1" has one paragraph (paraId 439EAAC1), keyed normally.
    """

    fixture_name = "comments/multipara-comment.docx"

    ROOT_DURABLE_ID = "58A3F864"  # w:id="0" -- keyed on its LAST paraId, 53D84979
    ROOT_LAST_PARA_ID = "53D84979"
    SINGLE_DURABLE_ID = "22F779FB"  # w:id="1"
    ROOT_CONTENT = "First paragraph of a two-paragraph comment.Second paragraph of the same comment."

    def test_list_open_items_reports_durable_ids_with_joined_content_and_unresolved(self):
        # Spec test 1: two comments, both comment_ids are 8-hex durableIds
        # (no raw-w:id fallback firing on a Word-authored multi-paragraph
        # comment), the root's content joins both paragraphs, both open.
        open_items = tracked_changes.execute_list_open_items(str(self.target))
        ids = {c["comment_id"] for c in open_items["comments"]}
        self.assertEqual(ids, {self.ROOT_DURABLE_ID, self.SINGLE_DURABLE_ID})
        for cid in ids:
            self.assertEqual(len(cid), 8)
            int(cid, 16)  # every char is hex -- confirms this is a durableId, not a raw w:id like "0"/"1"

        root = next(c for c in open_items["comments"] if c["comment_id"] == self.ROOT_DURABLE_ID)
        self.assertEqual(root["content"], self.ROOT_CONTENT)
        self.assertFalse(root["resolved"])
        single = next(c for c in open_items["comments"] if c["comment_id"] == self.SINGLE_DURABLE_ID)
        self.assertFalse(single["resolved"])

    def test_get_comment_thread_returns_both_paragraphs_and_no_replies(self):
        # Spec test 2.
        thread = comments.execute_get_comment_thread(str(self.target), self.ROOT_DURABLE_ID)
        self.assertEqual(thread["content"], self.ROOT_CONTENT)
        self.assertEqual(thread["replies"], [])
        self.assertEqual(thread["comment_id_resolved_via"], "durableId")

    def test_resolve_marks_the_last_paragraphs_commentex_and_creates_no_new_one(self):
        # Spec test 3.
        with zipfile.ZipFile(self.target) as zf:
            before_entries = [
                c for c in ET.fromstring(zf.read("word/commentsExtended.xml")) if projection._ln(c) == "commentEx"
            ]
        before_count = len(before_entries)

        evidence = comments.execute_resolve_comment(str(self.target), self.ROOT_DURABLE_ID)
        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["comment_id_resolved_via"], "durableId")

        with zipfile.ZipFile(self.target) as zf:
            after_entries = [
                c for c in ET.fromstring(zf.read("word/commentsExtended.xml")) if projection._ln(c) == "commentEx"
            ]
        self.assertEqual(len(after_entries), before_count, "resolve must not create a new commentEx entry")
        target_ex = next(c for c in after_entries if projection._attr(c, "paraId") == self.ROOT_LAST_PARA_ID)
        self.assertEqual(projection._attr(target_ex, "done"), "1")
        # Never touches the FIRST paragraph's paraId -- it has no commentEx
        # of its own (Word never wrote one for it either).
        self.assertIsNone(next((c for c in after_entries if projection._attr(c, "paraId") == "5C83D4AC"), None))

        open_items = tracked_changes.execute_list_open_items(str(self.target))
        self.assertNotIn(self.ROOT_DURABLE_ID, {c["comment_id"] for c in open_items["comments"]})
        thread = comments.execute_get_comment_thread(str(self.target), self.ROOT_DURABLE_ID)
        self.assertTrue(thread["resolved"])

    def test_reply_links_paraidparent_to_the_last_paragraph(self):
        # Spec test 4.
        reply = comments.execute_reply_to_comment(str(self.target), self.ROOT_DURABLE_ID, "a reply")
        self.assertEqual(reply["comment_id_resolved_via"], "durableId")

        with zipfile.ZipFile(self.target) as zf:
            ext_root = ET.fromstring(zf.read("word/commentsExtended.xml"))
        reply_ex = next(c for c in ext_root if projection._attr(c, "paraIdParent"))
        self.assertEqual(projection._attr(reply_ex, "paraIdParent"), self.ROOT_LAST_PARA_ID)

        thread = comments.execute_get_comment_thread(str(self.target), self.ROOT_DURABLE_ID)
        self.assertEqual(thread["reply_count"], 1)
        self.assertEqual(thread["replies"][0]["comment_id"], reply["comment_id"])

    def test_w_id_fallback_on_resolve_and_get_comment_thread(self):
        # Spec test 5, first two parts: resolve_comment(path, "1") and
        # get_comment_thread(path, "0") both resolve via the raw w:id and
        # report comment_id_resolved_via: "w_id".
        resolve_evidence = comments.execute_resolve_comment(str(self.target), "1")
        self.assertTrue(resolve_evidence["applied"])
        self.assertEqual(resolve_evidence["comment_id_resolved_via"], "w_id")
        single_thread = comments.execute_get_comment_thread(str(self.target), self.SINGLE_DURABLE_ID)
        self.assertTrue(single_thread["resolved"])

        root_thread = comments.execute_get_comment_thread(str(self.target), "0")
        self.assertEqual(root_thread["comment_id_resolved_via"], "w_id")
        self.assertEqual(root_thread["content"], self.ROOT_CONTENT)

    def test_unknown_comment_id_lists_both_durable_ids_and_w_ids(self):
        # Spec test 5, third part: a comment_id matching neither kind
        # raises INVALID_INPUT with both id kinds in the diagnostics.
        with self.assertRaises(VerifyError) as cm:
            comments.execute_get_comment_thread(str(self.target), "nonexistent")
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        diagnostics = cm.exception.envelope.diagnostics
        self.assertIn(self.ROOT_DURABLE_ID, diagnostics["available_durable_ids"])
        self.assertIn(self.SINGLE_DURABLE_ID, diagnostics["available_durable_ids"])
        self.assertIn("0", diagnostics["available_w_ids"])
        self.assertIn("1", diagnostics["available_w_ids"])

        with self.assertRaises(VerifyError) as cm2:
            comments.execute_resolve_comment(str(self.target), "nonexistent")
        self.assertEqual(cm2.exception.envelope.error_code, ErrorCode.INVALID_INPUT)
        self.assertIn("available_durable_ids", cm2.exception.envelope.diagnostics)
        self.assertIn("available_w_ids", cm2.exception.envelope.diagnostics)


class GoldenFixtureListOpenItemsRegressionTests(unittest.TestCase):
    """Spec test 7: the identity fix must not change list_open_items'
    output for the golden fixture -- it carries only single-paragraph
    comments, so a first-vs-last-paragraph identity change is a no-op for
    it. Pinned as an exact snapshot rather than a relative comparison
    (there is no "before this change" binary to diff against here); a
    single-paragraph comment has only one paraId, so this snapshot is
    equally a regression pin against ANY future accidental identity
    change, not just this one. Read-only, never writes to the golden."""

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
        actual = __import__("hashlib").sha256(GOLDEN.read_bytes()).hexdigest()
        assert actual == GOLDEN_SHA256, "golden-comment.docx was modified by a test -- this must never happen"

    def test_list_open_items_output_unchanged(self):
        open_items = tracked_changes.execute_list_open_items(str(self.target))
        expected = [
            {
                "comment_id": "654503C0",
                "w_id": "0",
                "content": "comment.",
                "resolved": False,
                "reply_count": 0,
                "replies": [],
                "quoted_text": "role in making cities healthier and more livable. Parks, gardens, and "
                "tree-lined streets help soften the hard edges ",
                "author": "Michael Sutton",
                "created_time": "2026-09-10T21:33:00Z",
                "modified_time": "",
                "scope": "document",
            },
            {
                "comment_id": "7E7E1D27",
                "w_id": "1",
                "content": "reply.",
                "resolved": False,
                "reply_count": 0,
                "replies": [],
                "quoted_text": "role in making cities healthier and more livable. Parks, gardens, and "
                "tree-lined streets help soften the hard edges ",
                "author": "Michael Sutton",
                "created_time": "2026-09-10T21:33:00Z",
                "modified_time": "",
                "scope": "document",
            },
        ]
        self.assertEqual(open_items["comments"], expected)


if __name__ == "__main__":
    unittest.main()
