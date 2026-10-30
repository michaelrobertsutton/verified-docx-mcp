"""Issue #110: live_revision (pane op body_hash, with a body_ooxml fallback for
panes that predate body_hash_v1), compact live_list_sections, and the rule that a
non-observing lookup never re-arms a refused client.

Pattern from test_live_reads.py: a real bridge, a FakePane, a sections.docx body.
"""

from __future__ import annotations

import hashlib
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument, FakePane, docx_to_flat_opc
from test_live_comments import LiveCommentsTestCase

from verified_docx_mcp import projection, server, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import reads_live, write_mode

FIXTURES = REPO / "tests" / "fixtures"
LIVE_BODY = FIXTURES / "sections.docx"


def pane_sections(path: Path) -> list[dict]:
    """What the pane's sections_list would report for *path*, with real full hashes."""
    proj = projection.project_part(path)
    texts = projection.paragraph_texts(proj)
    out = []
    for sec in projection.find_sections_impl(path):
        if sec["kind"] != "heading":
            continue
        start = projection.para_index(sec["start_para_ref"])
        end = projection.para_index(sec["end_para_ref"]) + 1
        slug = sec["section_key"].rsplit("-", 1)[0]
        out.append(
            {
                "section_key": sec["section_key"],
                "slug": slug,
                "heading_text": sec["heading_text"],
                "level": sec["outline_level"] + 1,
                "paragraph_count": sec["paragraph_count"],
                "slug_unique": True,
                "sectionSha256": hashlib.sha256("\n".join(texts[start:end]).encode("utf-8")).hexdigest(),
            }
        )
    return out


class LiveRevisionTestCase(LiveCommentsTestCase):
    fixture_name = "frag.docx"

    async def connect_live(self, **kwargs) -> FakePane:
        kwargs.setdefault("body_ooxml", docx_to_flat_opc(LIVE_BODY))
        kwargs.setdefault("sections", pane_sections(LIVE_BODY))
        return await self.connect_pane(**kwargs)

    async def revision(self):
        return await self.call(reads_live.live_revision, str(self.target))


class LiveRevisionTests(LiveRevisionTestCase):
    async def test_native_pane_is_one_round_trip_and_ships_no_body(self) -> None:
        pane = await self.connect_live()
        result = await self.revision()
        self.assertEqual(result["via"], "body_hash")
        self.assertEqual(result["revision"], f"live:sha256:{pane.document.sha256()}")
        self.assertEqual(result["body_sha12"], pane.document.sha256()[:12])
        self.assertEqual(pane.body_hash_requests, 1)
        self.assertEqual(pane.body_ooxml_requests, 0)
        self.assertIn("T", result["observed_at"])
        write_mode.check_not_stale(result["revision"], pane.document.sha256())

    async def test_section_digests_agree_with_find_sections_digests(self) -> None:
        await self.connect_live()
        result = await self.revision()
        file_side = [
            [s["section_key"], s["sha12"]]
            for s in projection.find_sections_impl(LIVE_BODY, digests=True)
            if s["kind"] == "heading"
        ]
        self.assertEqual(result["sections"], file_side)

    async def test_old_pane_falls_back_to_body_ooxml_and_says_so(self) -> None:
        pane = await self.connect_live(capabilities=["body_ooxml"])
        result = await self.revision()
        self.assertEqual(result["via"], "body_ooxml_fallback")
        self.assertEqual(pane.body_hash_requests, 0)
        self.assertEqual(pane.body_ooxml_requests, 1)
        self.assertEqual(result["revision"], f"live:sha256:{pane.document.sha256()}")
        file_side = [
            [s["section_key"], s["sha12"]]
            for s in projection.find_sections_impl(LIVE_BODY, digests=True)
            if s["kind"] == "heading"
        ]
        self.assertEqual(result["sections"], file_side)

    async def test_content_digests_take_one_body_ooxml_and_no_body_hash(self) -> None:
        """Issue #117: every value comes from one body_ooxml snapshot, even on a body_hash_v1 pane."""
        pane = await self.connect_live()
        result = await self.call(reads_live.live_revision, str(self.target), content_digests=True)
        self.assertEqual((pane.body_ooxml_requests, pane.body_hash_requests), (1, 0))
        self.assertEqual(result["via"], "body_ooxml")
        self.assertEqual(result["revision"], f"live:sha256:{pane.document.sha256()}")
        file_sections, preamble = projection.find_sections_and_preamble(LIVE_BODY, digests=True, content_digests=True)
        heads = [s for s in file_sections if s["kind"] == "heading"]
        self.assertEqual(result["sections"], [[s["section_key"], s["sha12"]] for s in heads])
        self.assertEqual(result["content_sections"], [[s["section_key"], s["content_sha12"]] for s in heads])
        self.assertEqual(result["preamble_content_sha12"], preamble)

    async def test_pane_with_neither_capability_is_refused(self) -> None:
        await self.connect_live(capabilities=[])
        with self.assertRaises(VerifyError) as ctx:
            await self.revision()
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_a_malformed_pane_reply_is_live_op_failed_not_a_crash(self) -> None:
        from unittest import mock

        from verified_docx_mcp.live import comments_live

        await self.connect_live()
        for bad in ({"bodySha256": 7, "sections": []}, {"bodySha256": "a" * 64, "sections": "nope"},
                    {"bodySha256": "a" * 64, "sections": [["only-one"]]}):
            with mock.patch.object(comments_live, "_request", return_value=bad):
                with self.assertRaises(VerifyError, msg=str(bad)) as ctx:
                    await self.revision()
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_OP_FAILED)

    async def test_no_session_is_live_unavailable(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            await self.revision()
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.LIVE_UNAVAILABLE)

    async def test_revision_is_accepted_as_revision_before_for_a_write(self) -> None:
        doc = FakeDocument(text="alpha beta gamma")
        pane = await self.connect_live(document=doc)
        result = await self.revision()
        edit = {"find": "beta", "replace": "BETA", "expected_matches": 1}
        written = await self.call(
            text_edit.execute_replace_text_batch_live, str(self.target), [edit], result["revision"]
        )
        self.assertTrue(written["applied"])
        self.assertEqual(pane.document.text, "alpha BETA gamma")
        self.assertEqual(written["revision_before"], result["revision"])

    async def test_a_stale_revision_is_still_refused(self) -> None:
        doc = FakeDocument(text="alpha beta gamma")
        await self.connect_live(document=doc)
        result = await self.revision()
        doc.text = "alpha beta gamma!"  # the body moves on after the read
        edit = {"find": "beta", "replace": "BETA", "expected_matches": 1}
        refused = await self.call(
            text_edit.execute_replace_text_batch_live, str(self.target), [edit], result["revision"]
        )
        self.assertEqual(refused["batch_status"], "none_applied")
        self.assertEqual(refused["results"][0]["code"], ErrorCode.LIVE_STALE.value)
        self.assertEqual(doc.text, "alpha beta gamma!")


class CompactListSectionsTests(LiveRevisionTestCase):
    async def test_compact_returns_key_heading_and_sha12_only(self) -> None:
        pane = await self.connect_live(
            capabilities=["section_locks", "body_ooxml", "body_hash_v1"],
        )
        full = await self.call(server.live_list_sections, str(self.target))
        compact = await self.call(server.live_list_sections, str(self.target), compact=True)
        self.assertEqual(
            compact["sections"],
            [[s["section_key"], s["heading_text"], s["sectionSha256"][:12]] for s in full["sections"]],
        )
        self.assertEqual(compact["body_sha12"], pane.document.sha256()[:12])
        file_side = [
            [s["section_key"], s["heading_text"], s["sha12"]]
            for s in projection.find_sections_impl(LIVE_BODY, digests=True)
            if s["kind"] == "heading"
        ]
        self.assertEqual(compact["sections"], file_side)
        self.assertLess(len(str(compact)), len(str(full)) / 2)


if __name__ == "__main__":
    unittest.main()
