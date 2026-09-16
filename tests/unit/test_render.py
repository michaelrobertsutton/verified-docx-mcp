#!/usr/bin/env python3
# Vendored from JennyStack scripts/test_docx_render.py at commit 678aac1
# (PR #105), source sha256
# 7d1eed51e5fe0470d0409787da9ff550eba3de8fd2829a2ac478844f452f478c, 19
# tests. Adapted (issue #28 WP-02) for this repo's layout only: the
# module under test is `verified_docx_mcp.render` here (the source file
# is `scripts/docx_render.py` there — the scaffold names it `render.py`),
# so `import docx_render` becomes `from verified_docx_mcp import render`
# and every `docx_render.` reference becomes `render.`; the two fixture
# PDFs move from `fixtures/md-render/pdf/` to `tests/fixtures/pdf/`
# (copied verbatim, same sha256, see that directory). Test logic,
# assertions, and comments are otherwise unchanged from the source file.
"""test_render.py — unit tests for src/verified_docx_mcp/render.py
(vendored from JennyStack's scripts/docx_render.py; see the header above).
stdlib `unittest`, no third-party dependencies. Run:

  uv run pytest tests/unit
  # or directly:
  uv run python -m unittest tests.unit.test_render

Covers:
  - page_count() against the two committed PDFs in tests/fixtures/pdf/
    (plain.pdf: pdfinfo and the regex fallback both find 2 pages directly;
    objstm.pdf: pdfinfo correctly finds 3 pages compressed inside an
    object stream, while the regex fallback — used only when pdfinfo is
    absent — finds ZERO literal "/Type /Page" matches and must return
    None, never 0, per the plan's explicit callout that a 0 would flow
    into a page-budget check as if it were real).
  - page_count() on an unreadable/missing path returns (None, None).
  - render_word()'s error-code mapping: AUTOMATION_NOT_GRANTED (-1712,
    -1743) and RENDER_FAILED (everything else), exercised WITHOUT a real
    Word instance by injecting a fake `osascript` earlier on PATH that
    prints a recorded stderr and exits 1 — an agent cannot revoke a real
    macOS Automation grant to exercise the failure path for real, so this
    is the documented substitute (WP-02 acceptance, agent side).
  - render_word()'s close_after / closed_after / close_error handling
    (issue #103, D3: close_after defaults to True), exercised via a fake
    `osascript` that prints the "OK <pages> closed=<0|1> <err>" form the
    real AppleScript worker emits.
  - probe_paragraphs / page_height_pt / paragraph_geometry parsing
    (issue #102): PROBE/PAGEH line parsing, a per-ordinal PROBE ERR
    line, a missing PAGEH line when probes were requested, no extra
    argv when no ordinals were requested, and an invalid ordinal
    failing before osascript is ever invoked.
  - WORD_SANDBOX_UNAVAILABLE when Word's container directory is absent.
"""

from __future__ import annotations

import os
import shlex
import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import render

FIXTURE_DIR = REPO / "tests" / "fixtures" / "pdf"
PLAIN_PDF = FIXTURE_DIR / "plain.pdf"
OBJSTM_PDF = FIXTURE_DIR / "objstm.pdf"


class PageCountTests(unittest.TestCase):
    def test_plain_via_pdfinfo(self):
        if not render._find_pdfinfo():
            self.skipTest("pdfinfo not installed on this machine")
        count, source = render.page_count(str(PLAIN_PDF))
        self.assertEqual(count, 2)
        self.assertEqual(source, "pdfinfo")

    def test_objstm_via_pdfinfo(self):
        if not render._find_pdfinfo():
            self.skipTest("pdfinfo not installed on this machine")
        count, source = render.page_count(str(OBJSTM_PDF))
        self.assertEqual(count, 3)
        self.assertEqual(source, "pdfinfo")

    def test_plain_regex_fallback(self):
        with mock.patch.object(render, "_find_pdfinfo", return_value=None):
            count, source = render.page_count(str(PLAIN_PDF))
        self.assertEqual(count, 2)
        self.assertEqual(source, "regex")

    def test_objstm_regex_fallback_returns_none_not_zero(self):
        # The whole point of committing objstm.pdf: its Page objects are
        # compressed inside an object stream, so a literal byte regex for
        # "/Type /Page" finds NOTHING. That must surface as (None, None)
        # — never (0, "regex"), which would look like a real, tiny
        # document to a caller (e.g. a page-budget check).
        with mock.patch.object(render, "_find_pdfinfo", return_value=None):
            count, source = render.page_count(str(OBJSTM_PDF))
        self.assertIsNone(count)
        self.assertIsNone(source)

    def test_missing_file_returns_none(self):
        count, source = render.page_count("/nonexistent/path/does-not-exist.pdf")
        self.assertIsNone(count)
        self.assertIsNone(source)


def _write_fake_osascript(bin_dir: Path, stderr_text: str, exit_code: int = 1) -> None:
    script = bin_dir / "osascript"
    script.write_text(
        "#!/bin/sh\n"
        f"echo {stderr_text!r} 1>&2\n"
        f"exit {exit_code}\n",
        encoding="utf-8",
    )
    script.chmod(script.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


class ErrorMappingTests(unittest.TestCase):
    """Exercises render_word()'s error-code mapping by injecting a fake
    `osascript` earlier on PATH. Guarded on Word's sandbox container
    directory actually existing — render_word() checks that BEFORE ever
    invoking osascript, so on a machine without Word installed at all
    these tests correctly skip rather than false-failing on the wrong
    error code."""

    def setUp(self):
        if not render._word_sandbox_root().is_dir():
            self.skipTest(
                "Word's sandbox container directory does not exist on this "
                "machine (Word never launched) — render_word() would raise "
                "WORD_SANDBOX_UNAVAILABLE before ever reaching osascript."
            )
        self._tmp = tempfile.TemporaryDirectory()
        self._bin_dir = Path(self._tmp.name) / "bin"
        self._bin_dir.mkdir()
        self._old_path = os.environ.get("PATH", "")
        os.environ["PATH"] = f"{self._bin_dir}{os.pathsep}{self._old_path}"

        self._in_dir = Path(self._tmp.name) / "in"
        self._in_dir.mkdir()
        self._in_path = self._in_dir / "test-input.docx"
        self._in_path.write_bytes(b"not a real docx, but osascript never reads it")
        self._out_path = Path(self._tmp.name) / "out" / "test-output.pdf"

    def tearDown(self):
        os.environ["PATH"] = self._old_path
        self._tmp.cleanup()

    def test_automation_not_granted_1743(self):
        # Recorded shape of a real macOS "Automation declined" AppleEvent
        # error (osascript's own message text for -1743, "not authorized
        # to send Apple events").
        _write_fake_osascript(
            self._bin_dir,
            "execution error: Not authorized to send Apple events to Microsoft Word. (-1743)",
        )
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(ctx.exception.code, "AUTOMATION_NOT_GRANTED")

    def test_automation_not_granted_1712(self):
        _write_fake_osascript(
            self._bin_dir,
            "execution error: Microsoft Word got an error: Application isn't running. (-1712)",
        )
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(ctx.exception.code, "AUTOMATION_NOT_GRANTED")

    def test_other_failure_maps_to_render_failed(self):
        # Recorded shape of this WP's own real, reproduced finding (see
        # the vendored module's docstring): Word rejecting saveAs with a
        # generic AppleEvent dispatch error that is NOT one of the two
        # automation-grant codes.
        _write_fake_osascript(self._bin_dir, "Error: Message not understood.")
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(ctx.exception.code, "RENDER_FAILED")
        self.assertIn("Message not understood", ctx.exception.detail or "")

    def test_stage_dir_cleaned_up_on_failure(self):
        _write_fake_osascript(self._bin_dir, "Error: Message not understood.")
        sandbox_root = render._word_sandbox_root()
        before = set(p.name for p in sandbox_root.iterdir())
        with self.assertRaises(render.RenderError):
            render.render_word(str(self._in_path), str(self._out_path))
        after = set(p.name for p in sandbox_root.iterdir())
        self.assertEqual(
            before, after, "render_word() left a stray staging directory behind on failure"
        )


def _write_fake_osascript_success(bin_dir: Path, stdout_text: str, args_log: Path | None = None) -> None:
    """A fake `osascript` that simulates a SUCCESSFUL render: it `touch`es
    the staged-output path (argv[3] from render_word()'s own subprocess
    call — the sh script's own $3, since $1 is the AppleScript file and
    $2 is the staged input) so render_word()'s staged_out.is_file() check
    passes without a real Word instance, then prints stdout_text (the
    "OK <pages> closed=<0|1> <err>" line under test) and exits 0. When
    args_log is given, the fake also writes its own argv (space-joined)
    there, so a test can assert the 5th argument (the close_after
    "close"/"keep" flag added for issue #103) took the expected value."""
    lines = ["#!/bin/sh", 'touch "$3"']
    if args_log is not None:
        lines.append(f'echo "$@" > {shlex.quote(str(args_log))}')
    lines.append(f"echo {stdout_text!r}")
    lines.append("exit 0")
    script = bin_dir / "osascript"
    script.write_text("\n".join(lines) + "\n", encoding="utf-8")
    script.chmod(script.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


class CloseAfterTests(unittest.TestCase):
    """Exercises render_word()'s close_after / closed_after / close_error
    handling (issue #103, D3: close_after defaults to True) via a fake
    `osascript` that prints the "OK <pages> closed=<0|1> <err>" form the
    real AppleScript worker now emits — no real Word instance needed.
    Guarded on Word's sandbox container directory existing, same as
    ErrorMappingTests above, for the same reason."""

    def setUp(self):
        if not render._word_sandbox_root().is_dir():
            self.skipTest(
                "Word's sandbox container directory does not exist on this "
                "machine (Word never launched) — render_word() would raise "
                "WORD_SANDBOX_UNAVAILABLE before ever reaching osascript."
            )
        self._tmp = tempfile.TemporaryDirectory()
        self._bin_dir = Path(self._tmp.name) / "bin"
        self._bin_dir.mkdir()
        self._old_path = os.environ.get("PATH", "")
        os.environ["PATH"] = f"{self._bin_dir}{os.pathsep}{self._old_path}"

        self._in_dir = Path(self._tmp.name) / "in"
        self._in_dir.mkdir()
        self._in_path = self._in_dir / "test-input.docx"
        self._in_path.write_bytes(b"not a real docx, but osascript never reads it")
        self._out_path = Path(self._tmp.name) / "out" / "test-output.pdf"

    def tearDown(self):
        os.environ["PATH"] = self._old_path
        self._tmp.cleanup()

    def test_closed_after_true(self):
        # Spec test 1: "OK 3 closed=1" -> pages 3, closed_after True,
        # left_open_document None, close_error None.
        _write_fake_osascript_success(self._bin_dir, "OK 3 closed=1 ")
        result = render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(result["pages"], 3)
        self.assertTrue(result["closed_after"])
        self.assertIsNone(result["left_open_document"])
        self.assertIsNone(result["close_error"])

    def test_closed_after_false_with_error_and_no_exception(self):
        # Spec test 2: "OK 3 closed=0 <error text>" -> closed_after False,
        # left_open_document == staged name, close_error contains the
        # error text, and NO exception is raised (a close failure is
        # reported, never fatal — the PDF is already written).
        _write_fake_osascript_success(
            self._bin_dir,
            "OK 3 closed=0 AppleEvent error -1708 while closing the render window",
        )
        result = render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(result["pages"], 3)
        self.assertFalse(result["closed_after"])
        self.assertTrue(result["left_open_document"])
        self.assertIn(self._in_path.stem, result["left_open_document"])
        self.assertIn("-1708", result["close_error"])

    def test_close_after_false_passes_keep_as_fourth_applescript_argument(self):
        # Spec test 3: close_after=False -> the fourth AppleScript
        # argument (closeMode) is "keep", not "close" — asserted via a
        # fake osascript that echoes its own argv to a file. Result is
        # closed_after False with no close_error (kept open on purpose,
        # not a close failure).
        args_log = Path(self._tmp.name) / "argv.log"
        _write_fake_osascript_success(self._bin_dir, "OK 3 closed=0 ", args_log=args_log)
        result = render.render_word(
            str(self._in_path), str(self._out_path), close_after=False
        )
        self.assertFalse(result["closed_after"])
        self.assertIsNone(result["close_error"])
        logged_argv = args_log.read_text().strip().split()
        self.assertEqual(logged_argv[-1], "keep", logged_argv)

    def test_legacy_ok_only_form_is_render_failed(self):
        # Spec test 4: the legacy bare "OK 3" (no "closed=" token) is now
        # RENDER_FAILED "unexpected form" — the script and parser ship
        # together and the old form is never silently accepted.
        _write_fake_osascript_success(self._bin_dir, "OK 3")
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(ctx.exception.code, "RENDER_FAILED")


def _write_fake_osascript_multiline(
    bin_dir: Path, stdout_lines: list[str], args_log: Path | None = None
) -> None:
    """A fake `osascript` like `_write_fake_osascript_success` above, but
    emitting several REAL newline-separated lines of stdout (PROBE/PAGEH
    lines followed by the OK line — issue #102), rather than the single
    line that helper's `!r`-based echo is limited to. Each line is
    printed with its own shell-quoted `echo` statement so embedded tabs
    survive intact."""
    lines = ["#!/bin/sh", 'touch "$3"']
    if args_log is not None:
        lines.append(f'echo "$@" > {shlex.quote(str(args_log))}')
    for out_line in stdout_lines:
        lines.append(f"echo {shlex.quote(out_line)}")
    lines.append("exit 0")
    script = bin_dir / "osascript"
    script.write_text("\n".join(lines) + "\n", encoding="utf-8")
    script.chmod(script.stat().st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)


class ProbeParagraphsTests(unittest.TestCase):
    """Exercises render_word()'s probe_paragraphs / page_height_pt /
    paragraph_geometry handling (issue #102) via a fake `osascript` that
    prints the PROBE/PAGEH/OK lines the real AppleScript worker now
    emits when ordinals are requested — no real Word instance needed.
    Guarded on Word's sandbox container directory existing, same as
    ErrorMappingTests/CloseAfterTests above, for the same reason."""

    def setUp(self):
        if not render._word_sandbox_root().is_dir():
            self.skipTest(
                "Word's sandbox container directory does not exist on this "
                "machine (Word never launched) — render_word() would raise "
                "WORD_SANDBOX_UNAVAILABLE before ever reaching osascript."
            )
        self._tmp = tempfile.TemporaryDirectory()
        self._bin_dir = Path(self._tmp.name) / "bin"
        self._bin_dir.mkdir()
        self._old_path = os.environ.get("PATH", "")
        os.environ["PATH"] = f"{self._bin_dir}{os.pathsep}{self._old_path}"

        self._in_dir = Path(self._tmp.name) / "in"
        self._in_dir.mkdir()
        self._in_path = self._in_dir / "test-input.docx"
        self._in_path.write_bytes(b"not a real docx, but osascript never reads it")
        self._out_path = Path(self._tmp.name) / "out" / "test-output.pdf"

    def tearDown(self):
        os.environ["PATH"] = self._old_path
        self._tmp.cleanup()

    def test_a_two_probes_and_pageh_are_parsed(self):
        _write_fake_osascript_multiline(
            self._bin_dir,
            [
                "PROBE\t1\t1\t72.0\tHello world",
                "PROBE\t2\t5\t540.0\tSecond probed paragraph",
                "PAGEH\t792.0",
                "OK 4 closed=1 ",
            ],
        )
        result = render.render_word(
            str(self._in_path), str(self._out_path), probe_paragraphs=[1, 2]
        )
        self.assertEqual(result["pages"], 4)
        self.assertEqual(result["page_height_pt"], 792.0)
        self.assertIsInstance(result["page_height_pt"], float)
        geometry = result["paragraph_geometry"]
        self.assertEqual(geometry[1], {"page": 1, "vpos_pt": 72.0, "text": "Hello world"})
        self.assertIsInstance(geometry[1]["page"], int)
        self.assertIsInstance(geometry[1]["vpos_pt"], float)
        self.assertEqual(
            geometry[2],
            {"page": 5, "vpos_pt": 540.0, "text": "Second probed paragraph"},
        )

    def test_b_probe_error_line_is_captured_without_exception(self):
        _write_fake_osascript_multiline(
            self._bin_dir,
            [
                "PROBE\t7\tERR\tparagraph 7 of active document is out of range",
                "PAGEH\t792.0",
                "OK 4 closed=1 ",
            ],
        )
        result = render.render_word(
            str(self._in_path), str(self._out_path), probe_paragraphs=[7]
        )
        self.assertEqual(
            result["paragraph_geometry"][7],
            {"error": "paragraph 7 of active document is out of range"},
        )

    def test_c_probes_requested_but_no_pageh_line_is_render_failed(self):
        _write_fake_osascript_multiline(self._bin_dir, ["OK 4 closed=1 "])
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(
                str(self._in_path), str(self._out_path), probe_paragraphs=[1]
            )
        self.assertEqual(ctx.exception.code, "RENDER_FAILED")

    def test_d_no_probes_requested_adds_no_extra_argv(self):
        args_log = Path(self._tmp.name) / "argv.log"
        _write_fake_osascript_multiline(
            self._bin_dir, ["OK 4 closed=1 "], args_log=args_log
        )
        result = render.render_word(str(self._in_path), str(self._out_path))
        self.assertEqual(result["page_height_pt"], None)
        self.assertEqual(result["paragraph_geometry"], {})
        logged_argv = args_log.read_text().strip().split()
        # $1 script, $2 staged-in, $3 staged-out, $4 staged-in-name, $5
        # close_mode — nothing after it when no ordinals were requested.
        self.assertEqual(len(logged_argv), 5, logged_argv)
        self.assertEqual(logged_argv[-1], "close", logged_argv)

    def test_e_invalid_ordinal_fails_before_osascript_runs(self):
        args_log = Path(self._tmp.name) / "argv.log"
        _write_fake_osascript_multiline(
            self._bin_dir, ["OK 4 closed=1 "], args_log=args_log
        )
        with self.assertRaises(render.RenderError) as ctx:
            render.render_word(
                str(self._in_path), str(self._out_path), probe_paragraphs=[0]
            )
        self.assertEqual(ctx.exception.code, "RENDER_FAILED")
        self.assertIn("invalid probe ordinal", ctx.exception.message)
        self.assertFalse(
            args_log.exists(), "osascript (fake) was invoked despite the bad ordinal"
        )


class WordSandboxUnavailableTests(unittest.TestCase):
    def test_missing_container_dir_raises_named_error(self):
        with tempfile.TemporaryDirectory() as td:
            fake_in = Path(td) / "in.docx"
            fake_in.write_bytes(b"placeholder")
            missing_root = Path(td) / "no-such-container"
            with mock.patch.object(render, "_word_sandbox_root", return_value=missing_root):
                with self.assertRaises(render.RenderError) as ctx:
                    render.render_word(str(fake_in), str(Path(td) / "out.pdf"))
            self.assertEqual(ctx.exception.code, "WORD_SANDBOX_UNAVAILABLE")
            self.assertIn(str(missing_root), ctx.exception.message)


if __name__ == "__main__":
    unittest.main()
