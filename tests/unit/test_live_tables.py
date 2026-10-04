"""Live table tools (issue #34): insert_table / get_table / replace_table_row
over the WSS ops channel, against tests/unit/fake_pane.py (no Word).

The real pane's Office.js calls are NOT exercised here (see
test_taskpane_js.py for the mocked-Word checks, and docs/live-mode.md for the
manual runbook): what these tests pin is the server's side of the contract --
validation before anything is sent, the capability / stale gates, the
independent read-back of every requested property, the failure audit record,
and the routing / error mapping.
"""

from __future__ import annotations

import asyncio
import json
import math
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_pane import FakeDocument
from test_live_write_mode import LiveWriteBridgeTestCase

from verified_docx_mcp import audit, mutations, tables
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import protocol, session

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

STYLE = "Grid Table 4 - Accent 1"


def _doc(text: str = "Intro\nProof\nTail") -> FakeDocument:
    return FakeDocument(text=text)


def _two_by_two() -> FakeDocument:
    doc = FakeDocument(text="Intro\nProof\nTail")
    doc.tables = [[["R1C1", "R1C2"], ["R2C1", "R2C2"]]]
    return doc


def _audit_records() -> list[dict]:
    path = audit._state_dir() / "audit.jsonl"
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text().splitlines() if line.strip()]


class _TableCase(LiveWriteBridgeTestCase):
    fixture_name = "tables.docx"

    async def insert(self, rows, **kwargs):
        kwargs.setdefault("style_id", STYLE)
        return await asyncio.to_thread(tables.execute_insert_table, str(self.target), rows, **kwargs)

    async def get(self, table_id=1, part="word/document.xml", **kwargs):
        return await asyncio.to_thread(tables.execute_get_table, str(self.target), table_id, part, **kwargs)

    async def row(self, cells, *, table_id=1, row_index=1, **kwargs):
        return await asyncio.to_thread(
            tables.execute_replace_table_row, str(self.target), table_id, row_index, cells, **kwargs
        )

    async def refused(self, code: ErrorCode, call, *args, **kwargs):
        with self.assertRaises(VerifyError) as cm:
            await call(*args, **kwargs)
        self.assertEqual(cm.exception.envelope.error_code, code, cm.exception.envelope.message)
        return cm.exception.envelope


class InsertTableLiveTests(_TableCase):
    async def test_auto_routes_live_and_leaves_the_file_alone(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()

        evidence = await self.insert(
            [["**Metric**", "**Value**"], ["Uptime", "99.9%"]],
            header_rows=1,
            anchor={"paragraph_text": "Proof", "position": "after"},
        )

        self.assertTrue(evidence["applied"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["table_id"], 1)
        self.assertEqual(evidence["merged_cells"], 0)
        self.assertEqual(evidence["anchor_resolved"], {"paragraph_text": "Proof", "position": "after"})
        self.assertEqual(evidence["before"], "")
        self.assertEqual(evidence["after"], "Metric\tValue\nUptime\t99.9%")
        self.assertTrue(evidence["revision_before"].startswith("live:sha256:"))
        self.assertNotEqual(evidence["revision_before"], evidence["revision_after"])
        self.assertEqual(doc.tables[0], [["Metric", "Value"], ["Uptime", "99.9%"]])
        self.assertEqual(doc.table_meta[0]["header"], 1)
        self.assertEqual(doc.table_meta[0]["style"], STYLE)
        self.assertEqual(self.target.read_bytes(), before_bytes)
        # inline bold reaches the wire as a run, not as literal asterisks
        sent = pane.table_insert_payloads[0]
        self.assertTrue(sent["rows"][0][0]["paragraphs"][0][0]["bold"])
        self.assertEqual(sent["rows"][0][0]["paragraphs"][0][0]["text"], "Metric")
        self.assertFalse(sent["rows"][1][1]["paragraphs"][0][0]["bold"])

    async def test_new_table_id_is_read_back_not_counted(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        evidence = await self.insert([["new"]], anchor={"paragraph_text": "Proof"})
        # inserted before the existing (end-of-document) table
        self.assertEqual(evidence["table_id"], 1)
        self.assertEqual(doc.tables[0], [["new"]])
        self.assertEqual(doc.tables[1][0][0], "R1C1")
        self.assertEqual(evidence["anchor_resolved"]["position"], "after")  # default

    async def test_no_anchor_appends_at_the_end(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        evidence = await self.insert([["tail"]])
        self.assertEqual(evidence["table_id"], 2)
        self.assertIsNone(evidence["anchor_resolved"])
        self.assertEqual(doc.tables[1], [["tail"]])

    async def test_position_before_and_after_table_id(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        e1 = await self.insert([["after-1"]], anchor={"after_table_id": 1})
        self.assertEqual(e1["table_id"], 2)
        self.assertEqual(e1["anchor_resolved"], {"after_table_id": 1})
        e2 = await self.insert([["before"]], anchor={"paragraph_text": "Tail", "position": "before"})
        self.assertEqual(e2["table_id"], 1)

    async def test_every_requested_property_is_applied(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.insert(
            [
                [{"markdown": "Head", "fill": "1F3864", "color": "FFFFFF", "bold": True, "align": "center", "valign": "center"}, "Value"],
                ["a", "b"],
            ],
            header_rows=1,
            grid_dxa=[2000, 4000],
            font_size_pt=10,
            anchor={"paragraph_text": "Proof", "position": "after"},
        )
        self.assertTrue(evidence["applied"])
        fmt = doc.table_meta[0]["cells"][(0, 0)]
        self.assertEqual(fmt["fill"], "#1F3864")
        self.assertEqual(fmt["color"], "#FFFFFF")
        self.assertTrue(fmt["bold"])
        self.assertEqual(fmt["align"], "center")
        self.assertEqual(fmt["valign"], "center")
        self.assertEqual(fmt["width"], 100.0)  # 2000 dxa -> points
        self.assertEqual(doc.table_meta[0]["cells"][(0, 1)]["width"], 200.0)
        self.assertEqual(fmt["size"], 10.0)

    async def test_chained_live_revisions(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        first = await self.insert([["one"]], anchor={"paragraph_text": "Proof"})
        second = await self.insert(
            [["two"]], anchor={"paragraph_text": "Proof"}, revision_before=first["revision_after"]
        )
        self.assertEqual(second["revision_before"], first["revision_after"])

    async def test_style_from_table_id_copies_the_source_style(self) -> None:
        doc = _two_by_two()
        doc._ensure_meta()
        doc.table_meta[0]["style"] = "Fancy"
        pane = await self.connect_pane(doc)
        evidence = await self.insert([["x"]], style_id=None, style_from_table_id=1, anchor={"after_table_id": 1})
        self.assertEqual(doc.table_meta[evidence["table_id"] - 1]["style"], "Fancy")
        # a concrete name is sent (so it can be verified), not the source index
        self.assertEqual(pane.table_insert_payloads[0]["style"], "Fancy")
        self.assertIsNone(pane.table_insert_payloads[0]["style_from_table_index"])

    async def test_style_from_table_without_a_style_is_style_not_found(self) -> None:
        doc = _two_by_two()
        doc._ensure_meta()
        doc.table_meta[0]["style"] = ""
        pane = await self.connect_pane(doc)
        await self.refused(ErrorCode.STYLE_NOT_FOUND, self.insert, [["x"]], style_id=None, style_from_table_id=1)
        self.assertEqual(pane.table_insert_payloads, [])

    async def test_builtin_style(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.insert([["x"]], style_id=None, style_builtin="GridTable4_Accent1")
        self.assertEqual(doc.table_meta[evidence["table_id"] - 1]["style"], "Grid Table 4 - Accent 1")
        await self.refused(ErrorCode.STYLE_NOT_FOUND, self.insert, [["y"]], style_id=None, style_builtin="Nope")
        self.assertEqual(len(doc.tables), 1)

    async def test_style_is_optional_in_live_mode(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        evidence = await self.insert([["x"]], style_id=None)
        self.assertTrue(evidence["applied"])

    async def test_unknown_style_inserts_nothing(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.STYLE_NOT_FOUND, self.insert, [["x"]], style_id="No Such Style")
        self.assertEqual(doc.tables, [])

    async def test_anchor_must_match_exactly_one_paragraph(self) -> None:
        doc = _doc("Dup\nDup\nOnce")
        await self.connect_pane(doc)
        await self.refused(ErrorCode.MATCH_COUNT_MISMATCH, self.insert, [["x"]], anchor={"paragraph_text": "Dup"})
        await self.refused(ErrorCode.ZERO_MATCH, self.insert, [["x"]], anchor={"paragraph_text": "Nope"})
        self.assertEqual(doc.tables, [])
        ok = await self.insert([["x"]], anchor={"paragraph_text": "  Once "})  # whitespace-normalized
        self.assertTrue(ok["applied"])

    async def test_after_table_id_out_of_range(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.TABLE_NOT_FOUND, self.insert, [["x"]], anchor={"after_table_id": 9})
        self.assertEqual(len(doc.tables), 1)

    async def test_invalid_requests_are_refused_before_anything_is_sent(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(doc)
        cases = {
            "span": dict(rows=[[{"markdown": "a", "span": 2}]]),
            "v_merge": dict(rows=[[{"markdown": "a", "v_merge": "restart"}], [{"markdown": "", "v_merge": "continue"}]]),
            "list markdown": dict(rows=[["- a\n- b"]]),
            "heading markdown": dict(rows=[["# h"]]),
            "section_key anchor": dict(rows=[["a"]], anchor={"section_key": "s1", "position": "end"}),
            "after_paragraph_text anchor": dict(rows=[["a"]], anchor={"after_paragraph_text": "Proof"}),
            "anchor with both forms": dict(rows=[["a"]], anchor={"paragraph_text": "Proof", "after_table_id": 1}),
            "anchor with unknown key": dict(rows=[["a"]], anchor={"paragraph_text": "Proof", "bogus": 1}),
            "two style sources": dict(rows=[["a"]], style_from_table_id=1),
            "header_rows too big": dict(rows=[["a"]], header_rows=1),
            "negative header_rows": dict(rows=[["a"], ["b"]], header_rows=-1),
            "cant_split": dict(rows=[["a"]], cant_split=True),
            "grid width mismatch": dict(rows=[["a", "b"]], grid_dxa=[100]),
            "non-positive width": dict(rows=[["a"]], grid_dxa=[0]),
            "font size zero": dict(rows=[["a"]], font_size_pt=0),
            "font size nan": dict(rows=[["a"]], font_size_pt=math.nan),
            "font size bool": dict(rows=[["a"]], font_size_pt=True),
            "empty rows": dict(rows=[]),
            "bad fill": dict(rows=[[{"markdown": "a", "fill": "zzz"}]]),
        }
        for name, kwargs in cases.items():
            with self.subTest(name):
                await self.refused(ErrorCode.INVALID_INPUT, self.insert, **kwargs)
        self.assertEqual(pane.table_insert_payloads, [])
        self.assertEqual(doc.tables, [])

    async def test_pane_without_table_edit_is_refused(self) -> None:
        doc = _doc()
        pane = await self.connect_pane(doc, capabilities=["cell_edit"])
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.insert, [["x"]])
        self.assertEqual(pane.table_insert_payloads, [])

    async def test_stale_revision_before_is_refused(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.LIVE_STALE, self.insert, [["x"]], revision_before="live:sha256:deadbeef")
        self.assertEqual(doc.tables, [])

    async def test_pane_refuses_before_writing_when_the_document_moved(self) -> None:
        # The server's own pre-check passes, then a co-author types before the
        # op lands: the pane's body-hash precondition must refuse, WITHOUT
        # having inserted anything (the session's post-hoc check could only
        # report staleness after the write).
        doc = _doc()
        await self.connect_pane(doc)
        original = doc.table_insert

        def racing(payload):
            doc.text += " (a co-author typed)"
            return original(payload)

        doc.table_insert = racing
        await self.refused(ErrorCode.LIVE_STALE, self.insert, [["x"]])
        self.assertEqual(doc.tables, [])

    async def test_nested_table_document_is_refused(self) -> None:
        doc = _doc()
        doc.has_nested_table = True
        await self.connect_pane(doc)
        await self.refused(ErrorCode.LIVE_OP_FAILED, self.insert, [["x"]])

    async def test_dropped_formatting_is_caught_and_audited(self) -> None:
        for dropped, rows, kwargs in [
            ("fill", [[{"markdown": "a", "fill": "FF0000"}]], {}),
            ("color", [[{"markdown": "a", "color": "00FF00"}]], {}),
            ("bold", [[{"markdown": "a", "bold": True}]], {}),
            ("align", [[{"markdown": "a", "align": "right"}]], {}),
            ("valign", [[{"markdown": "a", "valign": "bottom"}]], {}),
            ("size", [["a"]], {"font_size_pt": 9}),
            ("width", [["a"]], {"grid_dxa": [3000]}),
            ("header", [["a"], ["b"]], {"header_rows": 1}),
        ]:
            with self.subTest(dropped):
                doc = _doc()
                doc.dropped_table_formatting = {dropped}
                pane = await self.connect_pane(doc)
                records_before = len(_audit_records())
                envelope = await self.refused(ErrorCode.VERIFICATION_FAILED, self.insert, rows, **kwargs)
                self.assertTrue(envelope.diagnostics["problems"])
                # the table DID land (nothing to roll back in live mode) ...
                self.assertEqual(len(doc.tables), 1)
                # ... and the failure left an audit record
                new = _audit_records()[records_before:]
                failures = [r for r in new if r["tool"] == "insert_table:verification_failed"]
                self.assertEqual(len(failures), 1, new)
                self.assertTrue(failures[0]["evidence"]["verification_failure"])
                self.assertTrue(failures[0]["evidence"]["revision_before"].startswith("live:sha256:"))
                await pane.close()
                self._panes.remove(pane)

    async def test_wrong_style_read_back_is_caught(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        original = doc.table_get

        def lying(payload):
            result = original(payload)
            result["style"] = "Some Other Style"
            return result

        doc.table_get = lying
        envelope = await self.refused(ErrorCode.VERIFICATION_FAILED, self.insert, [["x"]])
        self.assertIn("style", envelope.message)

    async def test_color_read_back_is_canonicalized(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        original = doc.table_get

        def lower_no_hash(payload):
            result = original(payload)
            for row in result["rows"]:
                for cell in row:
                    cell["fill"] = cell["fill"].lstrip("#").lower()
            return result

        doc.table_get = lower_no_hash
        evidence = await self.insert([[{"markdown": "a", "fill": "AABBCC"}]])
        self.assertTrue(evidence["applied"])

    async def test_explicit_file_mode_under_a_session_still_refuses(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        envelope = await self.refused(ErrorCode.LIVE_SESSION_ACTIVE, self.insert, [["x"]], write_mode="file")
        self.assertIn('write_mode="live"', envelope.message)

    async def test_explicit_live_without_a_session_is_unavailable(self) -> None:
        await self.refused(ErrorCode.LIVE_UNAVAILABLE, self.insert, [["x"]], write_mode="live")

    async def test_auto_without_a_session_writes_the_file(self) -> None:
        evidence = await self.insert([["x"]], style_id="TableNormal")
        self.assertEqual(evidence["write_mode"], "file")

    async def test_bad_write_mode(self) -> None:
        await self.refused(ErrorCode.INVALID_INPUT, self.insert, [["x"]], write_mode="sideways")


class GetTableLiveTests(_TableCase):
    async def test_live_read_shape(self) -> None:
        doc = _two_by_two()
        doc._ensure_meta()
        doc.table_meta[0]["style"] = "Fancy"
        doc.table_meta[0]["header"] = 1
        await self.connect_pane(doc)
        result = await self.get(1)
        self.assertEqual(result["source"], "live")
        self.assertEqual(result["table_id"], 1)
        self.assertEqual((result["row_count"], result["col_count"]), (2, 2))
        self.assertFalse(result["has_merged_cells"])
        self.assertFalse(result["has_nested_table"])
        self.assertEqual(result["style"], "Fancy")
        self.assertEqual(result["header_row_count"], 1)
        self.assertEqual([[c["text"] for c in row] for row in result["rows"]], [["R1C1", "R1C2"], ["R2C1", "R2C2"]])
        self.assertEqual(result["rows"][1][0], {"row_index": 2, "cell_index": 1, "grid_span": 1, "v_merge": "none", "text": "R2C1"})

    async def test_merged_flag(self) -> None:
        doc = _two_by_two()
        doc.merged_tables = {0}
        await self.connect_pane(doc)
        self.assertTrue((await self.get(1))["has_merged_cells"])

    async def test_missing_table_and_capability(self) -> None:
        doc = _two_by_two()
        pane = await self.connect_pane(doc)
        await self.refused(ErrorCode.TABLE_NOT_FOUND, self.get, 5, source="live")
        await pane.close()
        self._panes.remove(pane)
        await self.connect_pane(_two_by_two(), capabilities=[])
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.get, 1, source="live")

    async def test_source_validation_and_routing(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.INVALID_INPUT, self.get, 1, source="sideways")
        # an explicit live read of another part is refused ...
        await self.refused(ErrorCode.INVALID_INPUT, self.get, 1, "word/header1.xml", source="live")
        # ... while "auto" falls back to the file for it (which has no such part)
        await self.refused(ErrorCode.PART_NOT_FOUND, self.get, 1, "word/header1.xml")
        # an explicit file read ignores the pane -- and says so (issue #33)
        file_result = await self.get(1, source="file")
        self.assertEqual(file_result["source"], "file")
        self.assertIn("live_session_ignored", file_result["warnings"])
        self.assertEqual(file_result["live_session"]["reason"], "requested_file")

    async def test_auto_without_a_session_reads_the_file(self) -> None:
        result = await self.get(1)
        self.assertEqual(result["source"], "file")
        self.assertNotIn("warnings", result)  # no session, nothing to warn about
        self.assertEqual(result["table_id"], 1)


class ReplaceTableRowLiveTests(_TableCase):
    async def test_auto_routes_live(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        before_bytes = self.target.read_bytes()
        evidence = await self.row(["**New A**", "New B"])
        self.assertEqual(evidence["write_mode"], "live")
        self.assertEqual(evidence["before"], "R1C1\tR1C2")
        self.assertEqual(evidence["after"], "New A\tNew B")
        self.assertEqual(doc.tables[0][0], ["New A", "New B"])
        self.assertEqual(doc.tables[0][1], ["R2C1", "R2C2"])
        self.assertEqual(self.target.read_bytes(), before_bytes)

    async def test_merged_table_is_refused(self) -> None:
        doc = _two_by_two()
        doc.merged_tables = {0}
        await self.connect_pane(doc)
        await self.refused(ErrorCode.MERGED_OR_NESTED_TABLE, self.row, ["a", "b"])
        self.assertEqual(doc.tables[0][0], ["R1C1", "R1C2"])

    async def test_shape_errors(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.INVALID_INPUT, self.row, ["only one"])
        await self.refused(ErrorCode.TABLE_ROW_NOT_FOUND, self.row, ["a", "b"], row_index=9)
        await self.refused(ErrorCode.TABLE_NOT_FOUND, self.row, ["a", "b"], table_id=9)
        await self.refused(ErrorCode.INVALID_INPUT, self.row, ["- a\n- b", "c"])
        self.assertEqual(doc.tables[0][0], ["R1C1", "R1C2"])

    async def test_concurrent_edit_refuses_the_whole_row(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        doc.before_cell_set = lambda: doc.tables[0][0].__setitem__(1, "a co-author typed here")
        await self.refused(ErrorCode.LIVE_OP_FAILED, self.row, ["a", "b"])
        # nothing written -- not even the cell that had NOT changed
        self.assertEqual(doc.tables[0][0], ["R1C1", "a co-author typed here"])

    async def test_pane_refuses_before_writing_when_the_document_moved(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        original = doc.cells_set

        def racing(payload):
            doc.text += " (typed)"
            return original(payload)

        doc.cells_set = racing
        await self.refused(ErrorCode.LIVE_STALE, self.row, ["a", "b"])
        self.assertEqual(doc.tables[0][0], ["R1C1", "R1C2"])

    async def test_dropped_write_is_caught_and_audited(self) -> None:
        doc = _two_by_two()
        doc.dropped_cell_writes = {(0, 0, 1)}
        await self.connect_pane(doc)
        before = len(_audit_records())
        await self.refused(ErrorCode.VERIFICATION_FAILED, self.row, ["a", "b"])
        failures = [r for r in _audit_records()[before:] if r["tool"] == "replace_table_row:verification_failed"]
        self.assertEqual(len(failures), 1)

    async def test_stale_revision_and_capability(self) -> None:
        doc = _two_by_two()
        pane = await self.connect_pane(doc)
        await self.refused(ErrorCode.LIVE_STALE, self.row, ["a", "b"], revision_before="live:sha256:deadbeef")
        await pane.close()
        self._panes.remove(pane)
        await self.connect_pane(_two_by_two(), capabilities=["cell_edit"])
        await self.refused(ErrorCode.LIVE_CAPABILITY_MISSING, self.row, ["a", "b"])

    async def test_explicit_modes(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        await self.refused(ErrorCode.LIVE_SESSION_ACTIVE, self.row, ["a", "b"], write_mode="file")


class ToolLayerTests(_TableCase):
    """Through the real @mcp.tool wrappers (FastMCP Client), so a wrapper that
    forgets to pass a new parameter through is caught."""

    async def _call(self, tool: str, **args):
        from fastmcp import Client

        from verified_docx_mcp import server

        async with Client(server.mcp) as client:
            return await client.call_tool(tool, {"path": str(self.target), **args}, raise_on_error=False)

    async def test_insert_table_write_mode_style_and_font_reach_the_live_route(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        result = await self._call(
            "insert_table",
            rows=[["**a**", "b"]],
            style_id=STYLE,
            write_mode="live",
            font_size_pt=10,
            anchor={"paragraph_text": "Proof", "position": "after"},
        )
        self.assertFalse(result.is_error, result)
        self.assertEqual(doc.tables[0], [["a", "b"]])
        self.assertEqual(doc.table_meta[0]["cells"][(0, 0)]["size"], 10.0)

    async def test_insert_table_style_builtin_and_from_table(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        result = await self._call("insert_table", rows=[["a"]], style_builtin="GridTable4_Accent1")
        self.assertFalse(result.is_error, result)
        result = await self._call("insert_table", rows=[["b"]], style_from_table_id=1)
        self.assertFalse(result.is_error, result)

    async def test_insert_table_file_mode_is_refused_under_a_session(self) -> None:
        doc = _doc()
        await self.connect_pane(doc)
        result = await self._call("insert_table", rows=[["a"]], style_id="TableNormal", write_mode="file")
        self.assertTrue(result.is_error)
        self.assertIn("LIVE_SESSION_ACTIVE", "".join(c.text for c in result.content))

    async def test_get_table_source_reaches_the_live_route(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        result = await self._call("get_table", table_id=1, source="live")
        self.assertFalse(result.is_error, result)
        self.assertEqual(result.structured_content["source"], "live")

    async def test_replace_table_row_write_mode_reaches_the_live_route(self) -> None:
        doc = _two_by_two()
        await self.connect_pane(doc)
        result = await self._call("replace_table_row", table_id=1, row_index=1, cells=["x", "y"], write_mode="live")
        self.assertFalse(result.is_error, result)
        self.assertEqual(doc.tables[0][0], ["x", "y"])
        result = await self._call("replace_table_row", table_id=1, row_index=1, cells=["x", "y"], write_mode="file")
        self.assertTrue(result.is_error)


class TimeoutAndCodeMappingTests(unittest.TestCase):
    def test_table_ops_get_more_than_the_default_timeout(self) -> None:
        for op in ("table_get", "table_insert", "cells_set"):
            self.assertGreater(session.resolve_timeout(op), session.DEFAULT_REQUEST_TIMEOUT, op)

    def test_pane_codes_map_to_typed_errors(self) -> None:
        from verified_docx_mcp.live import write_mode

        for pane_code, expected in [
            ("stale", ErrorCode.LIVE_STALE),
            ("style_not_found", ErrorCode.STYLE_NOT_FOUND),
            ("table_not_found", ErrorCode.TABLE_NOT_FOUND),
            ("zero_match", ErrorCode.ZERO_MATCH),
            ("something_else", ErrorCode.LIVE_OP_FAILED),
        ]:
            self.assertEqual(write_mode.classify_op_failed(session.LiveOpFailed(pane_code, "x")), expected)


class ProtocolTests(unittest.TestCase):
    def _cell(self, **kw):
        return {"paragraphs": [[{"text": "a", "bold": False}]], **kw}

    def test_ops_are_registered(self) -> None:
        for op in ("table_get", "table_insert", "cells_set"):
            self.assertIn(op, protocol.VALID_OPS)

    def test_table_get_round_trip(self) -> None:
        payload = protocol.TableGetPayload(3)
        self.assertEqual(protocol.TableGetPayload.from_json(payload.to_json()), payload)
        with self.assertRaises(protocol.ProtocolError):
            protocol.TableGetPayload.from_json({"table_index": 0})

    def test_table_insert_round_trip(self) -> None:
        payload = protocol.TableInsertPayload(
            rows=[[self._cell(fill="AABBCC", align="center")], [self._cell()]],
            anchor={"paragraph_text": "Proof", "position": "after"},
            style=STYLE,
            header_rows=1,
            column_widths_pt=[100.0],
            font_size_pt=10,
            track_changes=True,
            expected_body_sha256="abc",
        )
        wire = payload.to_json()
        self.assertEqual(wire["expectedBodySha256"], "abc")
        self.assertEqual(protocol.TableInsertPayload.from_json(wire), payload)

    def test_table_insert_validation(self) -> None:
        good = [[self._cell()], [self._cell()]]
        bad = {
            "empty": dict(rows=[]),
            "empty row": dict(rows=[[]]),
            "ragged": dict(rows=[[self._cell(), self._cell()], [self._cell()]]),
            "cell not an object": dict(rows=[["text"]]),
            "unknown cell key": dict(rows=[[self._cell(span=2)]]),
            "bad fill": dict(rows=[[self._cell(fill="red")]]),
            "bad align": dict(rows=[[self._cell(align="middle")]]),
            "bad valign": dict(rows=[[self._cell(valign="both")]]),
            "bad bold": dict(rows=[[self._cell(bold="yes")]]),
            "bad run": dict(rows=[[{"paragraphs": [["not a run"]]}]]),
            "bad run text": dict(rows=[[{"paragraphs": [[{"text": 1}]]}]]),
            "anchor both": dict(rows=good, anchor={"paragraph_text": "a", "position": "after", "after_table_index": 1}),
            "anchor neither": dict(rows=good, anchor={}),
            "anchor no position": dict(rows=good, anchor={"paragraph_text": "a"}),
            "anchor bad position": dict(rows=good, anchor={"paragraph_text": "a", "position": "middle"}),
            "anchor blank": dict(rows=good, anchor={"paragraph_text": "  ", "position": "after"}),
            "anchor table zero": dict(rows=good, anchor={"after_table_index": 0}),
            "anchor table with position": dict(rows=good, anchor={"after_table_index": 1, "position": "after"}),
            "anchor unknown key": dict(rows=good, anchor={"paragraph_text": "a", "position": "after", "x": 1}),
            "two styles": dict(rows=good, style="a", style_builtin="b"),
            "style from zero": dict(rows=good, style_from_table_index=0),
            "header too big": dict(rows=good, header_rows=2),
            "header negative": dict(rows=good, header_rows=-1),
            "header bool": dict(rows=good, header_rows=True),
            "widths wrong count": dict(rows=good, column_widths_pt=[1.0, 2.0]),
            "widths nonpositive": dict(rows=good, column_widths_pt=[0]),
            "widths inf": dict(rows=good, column_widths_pt=[math.inf]),
            "font nan": dict(rows=good, font_size_pt=math.nan),
            "font negative": dict(rows=good, font_size_pt=-1),
        }
        for name, kwargs in bad.items():
            with self.subTest(name), self.assertRaises(protocol.ProtocolError):
                protocol.TableInsertPayload(**kwargs)
        protocol.TableInsertPayload(rows=good)  # the control case is accepted

    def test_cells_set_round_trip_and_validation(self) -> None:
        cell = {"table_index": 1, "row_index": 1, "cell_index": 2, "paragraphs": [[{"text": "x"}]], "expected_before_text": "y"}
        payload = protocol.CellsSetPayload(cells=[cell], track_changes=True, expected_body_sha256="h")
        self.assertEqual(protocol.CellsSetPayload.from_json(payload.to_json()), payload)
        for name, cells in {
            "empty": [],
            "not an object": ["x"],
            "zero index": [{**cell, "row_index": 0}],
            "missing expected": [{k: v for k, v in cell.items() if k != "expected_before_text"}],
            "duplicate address": [cell, dict(cell)],
            "bad paragraphs": [{**cell, "paragraphs": ["x"]}],
        }.items():
            with self.subTest(name), self.assertRaises(protocol.ProtocolError):
                protocol.CellsSetPayload(cells=cells)


if __name__ == "__main__":
    unittest.main()
