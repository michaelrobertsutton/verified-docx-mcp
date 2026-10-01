"""Issue #33: read_document / find_sections / list_tables / get_table read
the connected live pane (``source``), and a file read made while a live
session exists says so instead of passing a placeholder off as the live body.

Bridge pattern from test_live_comments.py: a real bridge on ephemeral ports and
a FakePane that answers ``body_ooxml`` with Flat OPC built from a real .docx
(``fake_pane.docx_to_flat_opc``). The placeholder file at the document's path
is ``frag.docx``; the "live" body is a different fixture, so which one a read
returned is unambiguous.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakePane, docx_to_flat_opc
from test_live_comments import LiveCommentsTestCase

from verified_docx_mcp import projection, server, tables
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import reads_live, write_mode

FIXTURES = REPO / "tests" / "fixtures"
LIVE_BODY = FIXTURES / "sections.docx"
LIVE_TABLES = FIXTURES / "tables.docx"
SHAREPOINT_URL = "https://contoso.sharepoint.com/sites/team/Shared%20Documents/frag.docx"


def _code(exc: VerifyError) -> ErrorCode:
    return exc.envelope.error_code


class LiveReadsTestCase(LiveCommentsTestCase):
    """Placeholder frag.docx on disk; a live pane holding a different body."""

    fixture_name = "frag.docx"

    def live_text(self) -> str:
        return projection.read_document_text(LIVE_BODY)

    def placeholder_text(self) -> str:
        return projection.read_document_text(self.target)

    async def connect_live(self, body: Path = LIVE_BODY, **kwargs) -> FakePane:
        kwargs.setdefault("body_ooxml", docx_to_flat_opc(body))
        return await self.connect_pane(**kwargs)

    async def read(self, **kwargs):
        kwargs.setdefault("format", "text")
        return await self.call(server.execute_read_document, str(self.target), **kwargs)


class IssueReproTests(LiveReadsTestCase):
    async def test_placeholder_and_live_bodies_differ(self) -> None:
        self.assertNotEqual(self.live_text(), self.placeholder_text())

    async def test_auto_reads_live_not_the_placeholder(self) -> None:
        await self.connect_live()
        result = await self.read()
        self.assertEqual(result["source"], "live")
        self.assertEqual(result["text"], self.live_text())
        self.assertNotEqual(result["text"], self.placeholder_text())
        self.assertNotIn(reads_live.WARNING_LIVE_SESSION_IGNORED, result["warnings"])

    async def test_explicit_file_read_warns_that_a_live_session_exists(self) -> None:
        pane = await self.connect_live()
        result = await self.read(source="file")
        self.assertEqual(result["source"], "file")
        self.assertEqual(result["text"], self.placeholder_text())
        self.assertIn(reads_live.WARNING_LIVE_SESSION_IGNORED, result["warnings"])
        self.assertEqual(result["live_session"]["reason"], reads_live.REASON_REQUESTED_FILE)
        self.assertEqual(result["live_session"]["document_url"], str(self.target))
        self.assertTrue(result["live_session"]["action"])
        self.assertEqual(pane.body_ooxml_requests, 0)

    async def test_file_read_without_a_session_is_unchanged_and_has_no_warning(self) -> None:
        result = await self.read()
        self.assertEqual(result["source"], "file")
        self.assertEqual(result["warnings"], [])
        self.assertNotIn("live_session", result)


class LiveReadShapeTests(LiveReadsTestCase):
    async def test_live_revision_is_the_live_write_token(self) -> None:
        pane = await self.connect_live()
        result = await self.read()
        self.assertEqual(result["revision"], f"live:sha256:{pane.document.sha256()}")
        self.assertIsNone(result["revision_detail"])
        self.assertEqual(result["live"]["body_sha256"], pane.document.sha256())
        self.assertEqual(result["live"]["document_url"], str(self.target))
        # The token round-trips through the stale check live writes use.
        write_mode.check_not_stale(result["revision"], pane.document.sha256())
        with self.assertRaises(VerifyError) as ctx:
            write_mode.check_not_stale(result["revision"], "0" * 64)
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_STALE)

    async def test_live_read_works_with_no_local_file(self) -> None:
        self.target.unlink()
        await self.connect_live()
        result = await self.read()
        self.assertEqual(result["source"], "live")
        self.assertEqual(result["text"], self.live_text())

    async def test_every_format_matches_the_file_read_of_the_same_document(self) -> None:
        await self.connect_live()
        for fmt in ("markdown", "text", "runs"):
            with self.subTest(format=fmt):
                live = await self.read(format=fmt)
                self.assertEqual(live["source"], "live")
                self.assertEqual(live[fmt], projection_read(LIVE_BODY, fmt))

    async def test_find_sections_over_live(self) -> None:
        await self.connect_live()
        result = await self.call(server.execute_find_sections, str(self.target))
        self.assertEqual(result["source"], "live")
        self.assertEqual(result["sections"], projection.find_sections_impl(LIVE_BODY))

    async def test_list_tables_and_get_table_over_live(self) -> None:
        await self.connect_live(LIVE_TABLES)
        listed = await self.call(tables.execute_list_tables, str(self.target))
        self.assertEqual(listed["source"], "live")
        self.assertEqual(listed["tables"], tables.list_tables_impl(LIVE_TABLES))
        first = listed["tables"][0]["table_id"]
        got = await self.call(tables.execute_get_table, str(self.target), first)
        self.assertEqual(got["source"], "live")
        self.assertEqual(got["rows"], tables.get_table_impl(LIVE_TABLES, first)["rows"])

    async def test_a_non_pkg_namespace_prefix_is_fine(self) -> None:
        await self.connect_live(body_ooxml=docx_to_flat_opc(LIVE_BODY, prefix="x1"))
        result = await self.read()
        self.assertEqual(result["text"], self.live_text())

    async def test_a_reply_over_one_mib_gets_through(self) -> None:
        padding = "x" * (2 * 1024 * 1024)
        big = docx_to_flat_opc(LIVE_BODY).replace(
            "</pkg:package>",
            '<pkg:part pkg:name="/customXml/big.xml" pkg:contentType="application/xml">'
            f"<pkg:xmlData><big>{padding}</big></pkg:xmlData></pkg:part></pkg:package>",
        )
        await self.connect_live(body_ooxml=big)
        result = await self.read()
        self.assertEqual(result["text"], self.live_text())


class SourceSelectionTests(LiveReadsTestCase):
    async def test_invalid_source_is_rejected(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            await self.read(source="cloud")
        self.assertEqual(_code(ctx.exception), ErrorCode.INVALID_INPUT)

    async def test_live_with_no_session_is_unavailable(self) -> None:
        with self.assertRaises(VerifyError) as ctx:
            await self.read(source="live")
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_UNAVAILABLE)

    async def test_live_with_a_non_body_part_is_invalid_input(self) -> None:
        await self.connect_live()
        with self.assertRaises(VerifyError) as ctx:
            await self.read(source="live", part="word/header1.xml")
        self.assertEqual(_code(ctx.exception), ErrorCode.INVALID_INPUT)

    async def test_auto_with_a_non_body_part_reads_the_file_and_warns(self) -> None:
        # A non-body part can't be served live, so `auto` takes the file path.
        # frag.docx has no header, so the file path's PART_NOT_FOUND is the
        # proof it was taken (a live read would have been INVALID_INPUT).
        await self.connect_live()
        with self.assertRaises(VerifyError) as ctx:
            await self.read(part="word/header1.xml")
        self.assertEqual(_code(ctx.exception), ErrorCode.PART_NOT_FOUND)

    async def test_old_pane_auto_falls_back_with_a_warning_and_live_refuses(self) -> None:
        await self.connect_live(capabilities=[])
        result = await self.read()
        self.assertEqual(result["source"], "file")
        self.assertIn(reads_live.WARNING_LIVE_SESSION_IGNORED, result["warnings"])
        self.assertEqual(result["live_session"]["reason"], reads_live.REASON_PANE_MISSING_CAPABILITY)
        with self.assertRaises(VerifyError) as ctx:
            await self.read(source="live")
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_CAPABILITY_MISSING)

    async def test_basename_collision_auto_falls_back_but_explicit_live_works(self) -> None:
        await self.connect_live()  # local-path url
        await self.connect_live(document_url=SHAREPOINT_URL)  # same basename, different document
        self.assertTrue(self.registry.collisions())
        auto = await self.read()
        self.assertEqual(auto["source"], "file")
        self.assertEqual(auto["live_session"]["reason"], reads_live.REASON_BASENAME_COLLISION)
        self.assertIn(reads_live.WARNING_LIVE_SESSION_IGNORED, auto["warnings"])
        live = await self.read(source="live")
        self.assertEqual(live["source"], "live")
        self.assertEqual(live["text"], self.live_text())

    async def test_session_for_a_different_local_file_auto_reads_the_named_file(self) -> None:
        other = Path(self._tmp.name) / "elsewhere" / self.target.name
        await self.connect_live(document_url=str(other))
        auto = await self.read()
        self.assertEqual(auto["source"], "file")
        self.assertEqual(auto["live_session"]["reason"], reads_live.REASON_SESSION_MISMATCH)
        with self.assertRaises(VerifyError) as ctx:
            await self.read(source="live")
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_SESSION_MISMATCH)

    async def test_list_parts_warns_when_a_session_exists(self) -> None:
        await self.connect_live()
        result = await self.call(server.execute_list_parts, str(self.target))
        self.assertIn(reads_live.WARNING_LIVE_SESSION_IGNORED, result["warnings"])
        self.assertEqual(result["live_session"]["document_url"], str(self.target))


class NoFallbackAfterLiveIsChosenTests(LiveReadsTestCase):
    async def asyncSetUp(self) -> None:
        await super().asyncSetUp()
        self._old_timeout = os.environ.get("VERIFIED_DOCX_LIVE_TIMEOUT_BODY_OOXML_S")
        os.environ["VERIFIED_DOCX_LIVE_TIMEOUT_BODY_OOXML_S"] = "0.5"

    async def asyncTearDown(self) -> None:
        if self._old_timeout is None:
            os.environ.pop("VERIFIED_DOCX_LIVE_TIMEOUT_BODY_OOXML_S", None)
        else:
            os.environ["VERIFIED_DOCX_LIVE_TIMEOUT_BODY_OOXML_S"] = self._old_timeout
        await super().asyncTearDown()

    async def test_an_unresponsive_pane_errors_instead_of_reading_the_file(self) -> None:
        pane = await self.connect_live()
        pane.drop_ops.add("body_ooxml")
        with self.assertRaises(VerifyError) as ctx:
            await self.read()  # auto
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_DISCONNECTED)

    async def test_too_large_is_an_error_that_points_at_source_file(self) -> None:
        await self.connect_live(body_ooxml_error=("too_large", "body OOXML is 99 bytes, over the limit"))
        with self.assertRaises(VerifyError) as ctx:
            await self.read()
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_OP_FAILED)
        self.assertIn('source="file"', ctx.exception.envelope.message)

    async def test_a_malformed_export_is_an_error_not_a_file_read(self) -> None:
        await self.connect_live(body_ooxml="<not-xml")
        with self.assertRaises(VerifyError) as ctx:
            await self.read()
        self.assertEqual(_code(ctx.exception), ErrorCode.LIVE_OP_FAILED)
        self.assertEqual(ctx.exception.envelope.diagnostics["stage"], "flat_opc")


def projection_read(docx: Path, fmt: str):
    if fmt == "text":
        return projection.read_document_text(docx)
    if fmt == "runs":
        return projection.read_document_runs(docx)
    return projection.read_document_markdown(docx)[0]


class FlatOpcConversionTests(unittest.TestCase):
    def convert(self, ooxml: str):
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "out.docx"
            result = reads_live.flat_opc_to_docx(ooxml, dest)
            names = zipfile.ZipFile(dest).namelist() if dest.exists() else []
            return result, names, dest, tmp

    def test_round_trip_projects_the_same_text_as_the_source_docx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "out.docx"
            reads_live.flat_opc_to_docx(docx_to_flat_opc(LIVE_BODY), dest)
            self.assertEqual(
                projection.read_document_markdown(dest)[0], projection.read_document_markdown(LIVE_BODY)[0]
            )
            self.assertEqual(projection.find_sections_impl(dest), projection.find_sections_impl(LIVE_BODY))

    def test_tables_and_merged_tables_round_trip(self) -> None:
        for fixture in ("tables.docx", "tables-merged.docx"):
            with self.subTest(fixture=fixture), tempfile.TemporaryDirectory() as tmp:
                dest = Path(tmp) / "out.docx"
                reads_live.flat_opc_to_docx(docx_to_flat_opc(FIXTURES / fixture), dest)
                self.assertEqual(
                    tables.list_tables_impl(dest), tables.list_tables_impl(FIXTURES / fixture)
                )

    def assertRefused(self, ooxml: str, needle: str) -> None:
        with tempfile.TemporaryDirectory() as tmp, self.assertRaises(VerifyError) as ctx:
            reads_live.flat_opc_to_docx(ooxml, Path(tmp) / "out.docx")
        envelope = ctx.exception.envelope
        self.assertEqual(envelope.error_code, ErrorCode.LIVE_OP_FAILED)
        self.assertEqual(envelope.diagnostics["stage"], "flat_opc")
        self.assertIn(needle, envelope.message)

    def test_refuses_malformed_xml(self) -> None:
        self.assertRefused("<pkg:package", "not well-formed")

    def test_refuses_a_dtd(self) -> None:
        self.assertRefused('<!DOCTYPE x [<!ENTITY a "b">]><x/>', "DTD")

    def test_refuses_the_wrong_root(self) -> None:
        self.assertRefused("<html/>", "expected pkg:package")

    def test_refuses_a_package_with_no_body_part(self) -> None:
        good = docx_to_flat_opc(LIVE_BODY)
        self.assertRefused(good.replace("/word/document.xml", "/word/other.xml"), "no /word/document.xml")

    def test_refuses_a_duplicate_part(self) -> None:
        good = docx_to_flat_opc(LIVE_BODY)
        start = good.index("<pkg:part ")
        end = good.index("</pkg:part>", start) + len("</pkg:part>")
        self.assertRefused(good[:end] + good[start:end] + good[end:], "same part twice")

    def test_refuses_a_part_without_a_content_type(self) -> None:
        good = docx_to_flat_opc(LIVE_BODY)
        stripped = good.replace('pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"', "", 1)
        self.assertRefused(stripped, "no pkg:contentType")

    def test_refuses_a_part_with_no_data(self) -> None:
        ooxml = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            '<pkg:part pkg:name="/a.xml" pkg:contentType="application/xml"/></pkg:package>'
        )
        self.assertRefused(ooxml, "neither pkg:xmlData nor pkg:binaryData")

    def test_warns_when_referenced_styles_or_numbering_are_missing(self) -> None:
        # sections.docx's body uses heading styles; drop the styles part from the export.
        good = docx_to_flat_opc(LIVE_BODY)
        self.assertNotIn(reads_live.WARNING_LIVE_STYLES_MISSING, self.convert(good)[0].warnings)
        start = good.index('<pkg:part pkg:name="/word/styles.xml"')
        end = good.index("</pkg:part>", start) + len("</pkg:part>")
        result = self.convert(good[:start] + good[end:])[0]
        self.assertIn(reads_live.WARNING_LIVE_STYLES_MISSING, result.warnings)


if __name__ == "__main__":
    unittest.main()
