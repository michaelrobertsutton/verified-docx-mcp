"""Issue #113 part B: ``markdown_file=`` on the markdown write tools."""

from __future__ import annotations

import asyncio
import copy
import hashlib
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from fastmcp import Client
from fastmcp.exceptions import ToolError

from verified_docx_mcp import mutations, paths, server
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.live import approved_changes
from verified_docx_mcp.markdown_input import parse_pipe_table

FIXTURES = REPO / "tests" / "fixtures"
mutations._QUIESCE_INTERVAL_SECONDS = 0.02


def _document_xml(path: Path) -> bytes:
    with zipfile.ZipFile(path) as zf:
        return zf.read("word/document.xml")


class _Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name).resolve()
        self._old = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = str(self.root)

    def tearDown(self):
        if self._old is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old
        self._tmp.cleanup()

    def doc(self, fixture: str, name: str) -> str:
        target = self.root / name
        shutil.copyfile(FIXTURES / fixture, target)
        return str(target)

    def md(self, text: str, name: str = "in.md") -> str:
        file = self.root / name
        file.write_text(text, encoding="utf-8")
        return str(file)

    def err(self, fn, *args, **kwargs) -> ErrorCode:
        """The error code from a VerifyError or a server tool's ToolError(JSON envelope)."""
        with self.assertRaises((VerifyError, ToolError)) as ctx:
            fn(*args, **kwargs)
        exc = ctx.exception
        if isinstance(exc, VerifyError):
            return exc.envelope.error_code
        return ErrorCode(json.loads(str(exc))["error_code"])


class MarkdownFileParityTests(_Base):
    """markdown= and markdown_file= leave byte-identical word/document.xml."""

    TEXT = "## Background\n\nUpdated **bold** text.\n\n- one\n- two\n"

    def _pair(self, fixture, call):
        inline, viafile = self.doc(fixture, "a.docx"), self.doc(fixture, "b.docx")
        first = call(inline, {"markdown": self.TEXT})
        second = call(viafile, {"markdown_file": self.md(self.TEXT)})
        self.assertEqual(_document_xml(Path(inline)), _document_xml(Path(viafile)))
        return first, second

    def _check_evidence(self, first, second):
        self.assertNotIn("markdown_sha256", first)
        self.assertEqual(second["markdown_sha256"], hashlib.sha256(self.TEXT.encode()).hexdigest())
        self.assertEqual(second["markdown_chars"], len(self.TEXT))
        self.assertTrue(second["applied"])

    def test_replace_range_markdown(self):
        self._check_evidence(
            *self._pair("sections.docx", lambda p, kw: server.replace_range_markdown(p, "background-1", **kw))
        )

    def test_replace_body_markdown(self):
        self._check_evidence(*self._pair("word/empty-shell.docx", lambda p, kw: server.replace_body_markdown(p, **kw)))

    def test_append_markdown(self):
        self._check_evidence(*self._pair("word/empty-shell.docx", lambda p, kw: server.append_markdown(p, **kw)))

    def test_replace_cell_markdown(self):
        text = "**Bold** cell"
        a, b = self.doc("tables.docx", "a.docx"), self.doc("tables.docx", "b.docx")
        server.replace_cell_markdown(a, 1, 1, 1, markdown=text)
        out = server.replace_cell_markdown(b, 1, 1, 1, markdown_file=self.md(text))
        self.assertEqual(_document_xml(Path(a)), _document_xml(Path(b)))
        self.assertEqual(out["markdown_chars"], len(text))
        self.assertEqual(out["markdown_sha256"], hashlib.sha256(text.encode()).hexdigest())

    def test_insert_table_parses_a_pipe_table_into_rows(self):
        table = "| A | B |\n|---|:-:|\n| a1 | b1 |\n| a2 \\| x | b2 |\n"
        rows = [["A", "B"], ["a1", "b1"], ["a2 | x", "b2"]]
        a, b = self.doc("tables.docx", "a.docx"), self.doc("tables.docx", "b.docx")
        server.insert_table(a, rows=rows, style_id="TableNormal")
        out = server.insert_table(b, markdown_file=self.md(table), style_id="TableNormal")
        self.assertEqual(_document_xml(Path(a)), _document_xml(Path(b)))
        self.assertEqual(out["markdown_chars"], len(table))
        self.assertIn("markdown_sha256", out)

    def test_pipe_table_parser_refuses_non_tables(self):
        for bad in ("just prose", "| a |\nnot a separator\n", "| a |\n|---|\nprose row"):
            with self.assertRaises(VerifyError):
                parse_pipe_table(bad)


class MarkdownFileInputTests(_Base):
    def setUp(self):
        super().setUp()
        self.target = self.doc("sections.docx", "t.docx")

    def call(self, **kwargs):
        return server.replace_range_markdown(self.target, "background-1", **kwargs)

    def test_both_or_neither_is_invalid_input(self):
        file = self.md("x")
        self.assertEqual(self.err(self.call, markdown="x", markdown_file=file), ErrorCode.INVALID_INPUT)
        self.assertEqual(self.err(self.call), ErrorCode.INVALID_INPUT)
        self.assertEqual(self.err(server.insert_table, self.target, rows=[["a"]], markdown_file=file), ErrorCode.INVALID_INPUT)
        self.assertEqual(self.err(server.insert_table, self.target), ErrorCode.INVALID_INPUT)

    def test_outside_roots_is_path_escape(self):
        with tempfile.TemporaryDirectory() as other:
            outside = Path(other).resolve() / "x.md"
            outside.write_text("x")
            self.assertEqual(self.err(self.call, markdown_file=str(outside)), ErrorCode.DOCX_PATH_ESCAPE)

    def test_missing_outside_roots_is_path_escape_too(self):
        # Containment is checked before existence, so the code never says what exists outside.
        with tempfile.TemporaryDirectory() as other:
            missing = Path(other).resolve() / "nope.md"
            self.assertEqual(self.err(self.call, markdown_file=str(missing)), ErrorCode.DOCX_PATH_ESCAPE)

    def test_file_used_as_a_directory_is_invalid_input(self):
        (self.root / "a.md").write_text("x")
        self.assertEqual(self.err(self.call, markdown_file=str(self.root / "a.md" / "x.md")), ErrorCode.INVALID_INPUT)

    def test_symlink_to_outside_roots_is_refused(self):
        with tempfile.TemporaryDirectory() as other:
            secret = Path(other).resolve() / "secret.md"
            secret.write_text("secret")
            link = self.root / "link.md"
            link.symlink_to(secret)
            self.assertEqual(self.err(self.call, markdown_file=str(link)), ErrorCode.DOCX_PATH_ESCAPE)

    def test_other_failures_are_invalid_input(self):
        big = self.root / "big.md"
        big.write_bytes(b"a" * (paths.MAX_INPUT_TEXT_BYTES + 1))
        bad_utf8 = self.root / "bad.md"
        bad_utf8.write_bytes(b"\xff\xfe\xfa")
        wrong_suffix = self.root / "x.docx"
        wrong_suffix.write_text("x")
        (self.root / "dir.md").mkdir()
        for file in (big, bad_utf8, wrong_suffix, self.root / "missing.md", self.root / "dir.md"):
            with self.subTest(file=file.name):
                self.assertEqual(self.err(self.call, markdown_file=str(file)), ErrorCode.INVALID_INPUT)

    def test_exactly_two_megabytes_is_allowed(self):
        edge = self.root / "edge.md"
        edge.write_bytes(b"a" * paths.MAX_INPUT_TEXT_BYTES)
        _, text = paths.read_input_text_file(str(edge))
        self.assertEqual(len(text), paths.MAX_INPUT_TEXT_BYTES)

    def test_lean_envelope_keeps_the_markdown_evidence(self):
        async def run():
            async with Client(server.mcp) as client:
                return (
                    await client.call_tool(
                        "append_markdown", {"path": self.doc("word/empty-shell.docx", "e.docx"), "markdown_file": self.md("Hello.")}
                    )
                ).structured_content

        result = asyncio.run(run())
        self.assertEqual(result["markdown_chars"], 6)
        self.assertEqual(result["markdown_sha256"], hashlib.sha256(b"Hello.").hexdigest())


class ApprovedPrepareMarkdownFileTests(_Base):
    def _op(self, **args):
        return {"op": "replace_cell_markdown", "section": "s", "args": {"table_id": 1, "row_index": 1, "cell_index": 1, **args}}

    def test_content_replaces_the_path(self):
        file = self.md("Cell **text**")
        ops = [self._op(markdown_file=file)]
        original = copy.deepcopy(ops)
        out = approved_changes.resolve_markdown_files(ops)
        self.assertEqual(ops, original, "input is not mutated")
        self.assertEqual(out[0]["args"]["markdown"], "Cell **text**")
        self.assertNotIn("markdown_file", out[0]["args"])
        approved_changes.validate(out)
        # The digest binds the text: same text, different path -> same digest.
        other = approved_changes.resolve_markdown_files([self._op(markdown_file=self.md("Cell **text**", "other.md"))])
        self.assertEqual(approved_changes.sha(out), approved_changes.sha(other))
        self.assertNotIn(self.root.name, json.dumps(out))

    def test_both_markdown_and_file_refused_and_outside_roots_escapes(self):
        with self.assertRaises(VerifyError):
            approved_changes.resolve_markdown_files([self._op(markdown="x", markdown_file=self.md("y"))])
        with tempfile.TemporaryDirectory() as other:
            outside = Path(other).resolve() / "o.md"
            outside.write_text("x")
            self.assertEqual(
                self.err(approved_changes.resolve_markdown_files, [self._op(markdown_file=str(outside))]),
                ErrorCode.DOCX_PATH_ESCAPE,
            )

    def test_ops_without_markdown_keep_markdown_file_and_fail_validation(self):
        op = {"op": "replace_text", "section": "s", "args": {"find": "a", "replace": "b", "expected_matches": 1, "markdown_file": "x.md"}}
        out = approved_changes.resolve_markdown_files([op])
        with self.assertRaises(VerifyError):
            approved_changes.validate(out)


if __name__ == "__main__":
    unittest.main()
