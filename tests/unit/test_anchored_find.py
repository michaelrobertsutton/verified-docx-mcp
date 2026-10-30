"""Issue #113 part A: anchored ("ellipsis") find and para_ref targets -- locate.py, file
mode, and the pane's normalization parity. Live mode and approved runs are in
test_live_anchored_find.py."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import pytest

from verified_docx_mcp import locate, paths, projection, text_edit
from verified_docx_mcp.errors import ErrorCode, VerifyError

REPO = Path(__file__).resolve().parents[2]
FIXTURES = REPO / "tests" / "fixtures"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
E = locate.ELLIPSIS
NBSP = chr(0xA0)
LDQ, RDQ, RSQ = chr(0x201C), chr(0x201D), chr(0x2019)

# A 412-character paragraph (the issue's acceptance shape).
LONG = (
    "Skyward will modernize the claim platform "
    + "with measured, incremental releases that keep each state partner informed "
    * 4
    + "and retire the legacy mainframe batch jobs one workload at a time by 2027."
)
assert len(LONG) == 412, len(LONG)
FIND = f"Skyward will modernize{E}batch jobs one workload at a time by 2027."


def _p(text: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>' if text else "<w:p/>"


def _table(rows: list[list[str]]) -> str:
    body = "".join("<w:tr>" + "".join(f"<w:tc>{_p(c)}</w:tc>" for c in row) + "</w:tr>" for row in rows)
    return f"<w:tbl>{body}</w:tbl>"


def _document(*blocks: str) -> str:
    return f'<w:document xmlns:w="{W}"><w:body>{"".join(blocks)}<w:sectPr/></w:body></w:document>'


def _proj(*blocks: str) -> projection.Projection:
    return projection.project_document_root(ET.fromstring(_document(*blocks)))


def _code(ctx: pytest.ExceptionInfo) -> str:
    return ctx.value.envelope.error_code.value


# ---------------------------------------------------------------------------
# locate_anchored / locate_find / locate_paragraph
# ---------------------------------------------------------------------------


def test_expected_length_412_resolves_and_400_is_refused():
    proj = _proj(_p("Intro."), _p(LONG), _p("Outro."))
    result = locate.locate_anchored(FIND, proj, expected_length=412)
    start, end = result.spans[0]
    assert proj.text[start:end] == LONG
    assert result.match_count == 1 and result.rung == locate.RUNG_ANCHORED
    assert locate.locate_anchored(FIND, proj, expected_length=414).spans == result.spans  # +-2
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=400)
    assert _code(ctx) == "ZERO_MATCH"
    assert ctx.value.envelope.diagnostics["candidates"][0]["length"] == 412


def test_find_sha256_path_and_both_must_hold():
    proj = _proj(_p(LONG))
    digest = locate.span_sha256(LONG)
    assert locate.locate_anchored(FIND, proj, find_sha256=digest.upper()).match_count == 1
    assert locate.locate_anchored(FIND, proj, expected_length=412, find_sha256=digest).match_count == 1
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=412, find_sha256="0" * 64)
    assert _code(ctx) == "ZERO_MATCH"
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=300, find_sha256=digest)
    assert _code(ctx) == "ZERO_MATCH"


@pytest.mark.parametrize(
    ("find", "kwargs"),
    [
        (FIND, {}),  # neither expected_length nor find_sha256
        (f"Skyward{E}by 2027.", {"expected_length": 412}),  # prefix < 12
        (f"Skyward will modernize{E}2027.", {"expected_length": 412}),  # suffix < 12
        (f"Skyward will modernize{E}incremental releases{E}by 2027 the end", {"expected_length": 412}),
        (f"Skyward will modernize{E}incremental [...] the end of it", {"expected_length": 412}),
        (FIND, {"find_sha256": "not-a-hash"}),
        (FIND, {"expected_length": 0}),
    ],
)
def test_malformed_anchored_finds_are_invalid_input(find, kwargs):
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(find, _proj(_p(LONG)), **kwargs)
    assert _code(ctx) == "INVALID_INPUT"


def test_ascii_delimiter_and_normalized_whitespace_and_quotes():
    doc_text = f"The team{RSQ}s {LDQ}claims{RDQ} platform{NBSP} will  ship in measured releases by the end of 2027."
    proj = _proj(_p(doc_text))
    find = "The team's \"claims\" platform [...] in measured releases by the end of 2027."
    result = locate.locate_anchored(find, proj, find_sha256=locate.span_sha256(doc_text))
    start, end = result.spans[0]
    assert proj.text[start:end] == doc_text
    assert locate.span_sha256(doc_text) == locate.span_sha256("The team's \"claims\" platform will ship in measured releases by the end of 2027.")


def test_two_spans_are_ambiguous_with_two_contexts():
    other = LONG.replace("partner", "parties")
    proj = _proj(_p(LONG), _p("Between."), _p(other))
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=412)
    assert _code(ctx) == "AMBIGUOUS_SPAN"
    diag = ctx.value.envelope.diagnostics
    assert diag["total"] == 2 and len(diag["candidates"]) == 2
    assert {"start", "end", "context"} <= set(diag["candidates"][0])
    # find_sha256 tells the two apart.
    result = locate.locate_anchored(FIND, proj, find_sha256=locate.span_sha256(other))
    assert proj.text[slice(*result.spans[0])] == other


def test_a_span_across_paragraphs_is_a_structural_boundary():
    first, second = LONG[:200], LONG[200:]
    proj = _proj(_p(first), _p(second))
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=413)
    assert _code(ctx) == "STRUCTURAL_BOUNDARY"
    assert ctx.value.envelope.diagnostics["boundary_kinds"] == ["w:p"]


def test_a_span_across_table_cells_is_refused():
    proj = _proj(_table([["Skyward will modernize the platform", "batch jobs one workload at a time by 2027."]]))
    with pytest.raises(VerifyError) as ctx:
        locate.locate_anchored(FIND, proj, expected_length=79)
    assert _code(ctx) == "STRUCTURAL_BOUNDARY"
    assert "w:tc" in ctx.value.envelope.diagnostics["boundary_kinds"]


def test_span_filter_applies_before_the_check():
    proj = _proj(_p(LONG), _p(LONG))
    second_start = proj.text.index(LONG, 1)
    result = locate.locate_anchored(FIND, proj, expected_length=412, span_filter=lambda s, _e: s >= second_start)
    assert result.spans[0][0] == second_start


def test_locate_find_dispatch_rules():
    proj = _proj(_p(LONG))
    assert locate.locate_find("claim platform", proj, 1).rung == "exact"
    assert locate.locate_find(FIND, proj, None, expected_length=412).rung == locate.RUNG_ANCHORED
    for args, kwargs in [
        (("claims platform", proj, None), {}),  # plain needs expected_matches
        (("claims platform", proj, 1), {"expected_length": 15}),  # anchored-only argument
        ((FIND, proj, 2), {"expected_length": 412}),  # anchored names one span
    ]:
        with pytest.raises(VerifyError) as ctx:
            locate.locate_find(*args, **kwargs)
        assert _code(ctx) == "INVALID_INPUT"


def test_locate_paragraph_hash_ok_stale_and_out_of_range():
    proj = _proj(_p("Heading"), _p(LONG), _p(""))
    result = locate.locate_paragraph(proj, "p1", locate.span_sha256(LONG))
    assert proj.text[slice(*result.spans[0])] == LONG
    with pytest.raises(VerifyError) as ctx:
        locate.locate_paragraph(proj, "p1", locate.span_sha256(LONG + "x"))
    assert _code(ctx) == "PARA_HASH_MISMATCH"
    for ref in ("p9", "x1", "p2"):  # out of range, malformed, empty paragraph
        with pytest.raises(VerifyError) as ctx:
            locate.locate_paragraph(proj, ref, locate.span_sha256(LONG))
        assert _code(ctx) == "INVALID_INPUT"


def test_pane_target_short_long_and_duplicate():
    short = _proj(_p("A short paragraph that fits one Word search."))
    result = locate.locate_paragraph(short, "p0", locate.span_sha256("A short paragraph that fits one Word search."))
    assert text_edit.pane_target(None, short, result) == {
        "text": "A short paragraph that fits one Word search.",
        "find": "A short paragraph that fits one Word search.",
    }
    long_proj = _proj(_p(LONG))
    target = text_edit.pane_target(None, long_proj, locate.locate_anchored(FIND, long_proj, expected_length=412))
    assert target["prefix"] == LONG[:230] and target["suffix"] == "batch jobs one workload at a time by 2027."
    assert target["expected_text_sha256"] == locate.span_sha256(LONG) and target["text"] == LONG
    dup = _proj(_p(LONG), _p(LONG))
    with pytest.raises(VerifyError) as ctx:
        text_edit.pane_target(None, dup, locate.locate_paragraph(dup, "p0", locate.span_sha256(LONG)))
    assert _code(ctx) == "AMBIGUOUS_SPAN"


# ---------------------------------------------------------------------------
# File mode
# ---------------------------------------------------------------------------


class _FileCase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self._old = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._tmp.name
        self.target = Path(self._tmp.name) / "anchored.docx"
        self.write(_p("Heading"), _p(LONG), _p("Short closing paragraph."))

    def tearDown(self) -> None:
        if self._old is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old
        self._tmp.cleanup()

    def write(self, *blocks: str) -> None:
        tmp = self.target.with_suffix(".tmp")
        with zipfile.ZipFile(FIXTURES / "frag.docx") as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = _document(*blocks).encode() if info.filename == "word/document.xml" else src.read(info.filename)
                dst.writestr(info, data)
        shutil.move(tmp, self.target)

    def text(self) -> str:
        return projection.read_document_text(self.target)


class FileModeTests(_FileCase):
    def test_anchored_replace_text_replaces_the_whole_span(self):
        evidence = text_edit.execute_replace_text(
            str(self.target), FIND, "Rewritten.", None, expected_length=412, write_mode="file"
        )
        self.assertEqual(evidence["match_count"], 1)
        self.assertIn("Skyward will modernize", evidence["before"])
        self.assertEqual(self.text(), "Heading\nRewritten.\nShort closing paragraph.")

    def test_anchored_replace_with_400_writes_nothing(self):
        before = self.target.read_bytes()
        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_replace_text(str(self.target), FIND, "x", None, expected_length=400, write_mode="file")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.ZERO_MATCH)
        self.assertEqual(self.target.read_bytes(), before)

    def test_para_ref_replaces_one_paragraph_and_a_stale_hash_is_refused(self):
        with self.assertRaises(VerifyError) as ctx:
            text_edit.execute_replace_text(
                str(self.target), None, "x", None, para_ref="p1", expected_sha256=locate.span_sha256("old"),
                write_mode="file",
            )
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.PARA_HASH_MISMATCH)
        evidence = text_edit.execute_replace_text(
            str(self.target), None, "New second paragraph.", None, para_ref="p1",
            expected_sha256=locate.span_sha256(LONG), write_mode="file",
        )
        self.assertEqual(evidence["match_count"], 1)
        self.assertEqual(self.text(), "Heading\nNew second paragraph.\nShort closing paragraph.")

    def test_target_argument_mixes_are_invalid(self):
        for args, kwargs in [
            ((FIND, "x", None), {"para_ref": "p1", "expected_sha256": "0" * 64}),
            ((None, "x", None), {"para_ref": "p1"}),
            (("Heading", "x", None), {}),
            (("Heading", "x", 1), {"find_sha256": "0" * 64}),
        ]:
            with self.assertRaises(VerifyError) as ctx:
                text_edit.execute_replace_text(str(self.target), *args, write_mode="file", **kwargs)
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT, (args, kwargs))

    def test_anchored_format_text_and_apply_style(self):
        evidence = text_edit.execute_format_text(
            str(self.target), FIND, {"bold": True}, None, find_sha256=locate.span_sha256(LONG), write_mode="file",
            verbose=True,
        )
        self.assertEqual(evidence["match_count"], 1)
        self.assertTrue(all(r["bold"] for span in evidence["runs_after"] for r in span))
        styled = text_edit.execute_apply_style(str(self.target), FIND, "Heading2", None, expected_length=412)
        self.assertEqual(styled["match_count"], 1)
        self.assertEqual(self.text(), f"Heading\n{LONG}\nShort closing paragraph.")


# ---------------------------------------------------------------------------
# The pane's normalizeSpanText must hash exactly like locate.span_sha256.
# ---------------------------------------------------------------------------


NODE = shutil.which("node")


@unittest.skipUnless(NODE, "node not found on PATH")
def test_pane_anchored_find_harness_and_hash_parity():
    proc = subprocess.run(
        [NODE, str(REPO / "tests/unit/js/anchored_find_harness.mjs"), str(REPO / "addin/taskpane.js")],
        capture_output=True, text=True, timeout=60, check=False,
    )
    assert proc.returncode == 0, proc.stderr
    out = json.loads(proc.stdout.strip().splitlines()[-1])
    assert out["parityNormalized"] == locate.normalize_span_text(out["parityText"])
    assert out["parityHash"] == locate.span_sha256(out["parityText"])
    for key in ("longReplaced", "ambiguousRefused", "zeroMatchRefused", "invalidRefused", "racedRefused",
                "hashOnlyRaceRefused", "formatRaceRefused", "formatApplied"):
        assert out[key] is True, key


def test_pane_reports_anchored_find_v1_and_cache_buster_moved():
    js = (REPO / "addin/taskpane.js").read_text()
    assert '"anchored_find_v1"' in js.split("const PANE_CAPABILITIES = ", 1)[1].split(";", 1)[0]
    assert 'taskpane.js?v=68"' in (REPO / "addin/taskpane.html").read_text()
