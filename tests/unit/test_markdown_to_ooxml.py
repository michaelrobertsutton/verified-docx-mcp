"""Unit tests for src/verified_docx_mcp/markdown_to_ooxml.py (issue #28 WP-04).

Covers the markdown -> OOXML block builder directly (no atomic write, no
guard) so structural output (headings resolved through list_styles,
bold/italic, links + relationships, lists + numbering, tables) can be
asserted precisely, plus STYLE_NOT_FOUND's "never a hardcoded style"
guarantee.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import markdown_to_ooxml
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.projection import R_NS, W_NS

FIXTURES = REPO / "tests" / "fixtures"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


class StyleContextBuildTests(unittest.TestCase):
    def test_resolves_heading_styles_and_table_style_from_target_document(self):
        ctx = markdown_to_ooxml.StyleContext.build(FIXTURES / "word" / "one-page.docx")
        self.assertEqual(ctx.heading_styles.get(0), "Heading1")
        self.assertEqual(ctx.heading_styles.get(1), "Heading2")
        self.assertIsNotNone(ctx.table_style_id)
        self.assertTrue(ctx.numbering_part_exists)

    def test_missing_heading_level_raises_style_not_found_not_a_hardcode(self):
        ctx = markdown_to_ooxml.StyleContext(
            heading_styles={},
            table_style_id=None,
            hyperlink_rstyle_id=None,
            next_abstract_num_id=0,
            next_num_id=0,
            numbering_part_exists=False,
        )
        with self.assertRaises(VerifyError) as cm:
            ctx.heading_style_for_level(2)
        self.assertEqual(cm.exception.envelope.error_code, ErrorCode.STYLE_NOT_FOUND)
        self.assertEqual(cm.exception.envelope.diagnostics["level"], 2)


class RenderBlocksTests(unittest.TestCase):
    def setUp(self):
        self.ctx = markdown_to_ooxml.StyleContext.build(FIXTURES / "word" / "one-page.docx")

    def test_heading_uses_resolved_style_never_hardcoded(self):
        elements = markdown_to_ooxml.render_blocks("## A heading\n", self.ctx)
        self.assertEqual(len(elements), 1)
        p = elements[0]
        pstyle = p.find(f"{_w('pPr')}/{_w('pStyle')}")
        self.assertIsNotNone(pstyle)
        self.assertEqual(pstyle.get(_w("val")), self.ctx.heading_styles[1])

    def test_bold_italic_and_plain_runs(self):
        elements = markdown_to_ooxml.render_blocks("Plain **bold** *italic* ***both***.\n", self.ctx)
        p = elements[0]
        runs = p.findall(_w("r"))
        texts_by_style = []
        for r in runs:
            rpr = r.find(_w("rPr"))
            bold = rpr is not None and rpr.find(_w("b")) is not None
            italic = rpr is not None and rpr.find(_w("i")) is not None
            t = r.find(_w("t"))
            texts_by_style.append((t.text if t is not None else None, bold, italic))
        self.assertIn(("bold", True, False), texts_by_style)
        self.assertIn(("italic", False, True), texts_by_style)
        self.assertIn(("both", True, True), texts_by_style)
        self.assertTrue(any(text == "Plain " and not bold and not italic for text, bold, italic in texts_by_style))

    def test_link_creates_hyperlink_and_relationship(self):
        elements = markdown_to_ooxml.render_blocks("See [our site](https://example.com/page).\n", self.ctx)
        p = elements[0]
        hyperlink = p.find(_w("hyperlink"))
        self.assertIsNotNone(hyperlink)
        rid = hyperlink.get(f"{{{R_NS}}}id")
        self.assertIsNotNone(rid)
        self.assertEqual(len(self.ctx.new_relationships), 1)
        got_rid, rel_type, target = self.ctx.new_relationships[0]
        self.assertEqual(got_rid, rid)
        self.assertEqual(target, "https://example.com/page")
        self.assertIn("hyperlink", rel_type)
        run_text = hyperlink.find(f"{_w('r')}/{_w('t')}").text
        self.assertEqual(run_text, "our site")

    def test_bulleted_list_gets_numpr_and_a_fresh_numbering_definition(self):
        elements = markdown_to_ooxml.render_blocks("- one\n- two\n", self.ctx)
        self.assertEqual(len(elements), 2)
        num_ids = set()
        for p in elements:
            numid_el = p.find(f"{_w('pPr')}/{_w('numPr')}/{_w('numId')}")
            self.assertIsNotNone(numid_el)
            num_ids.add(numid_el.get(_w("val")))
        self.assertEqual(len(num_ids), 1)  # both items share the same list's numId
        self.assertEqual(len(self.ctx.new_abstract_nums), 1)
        self.assertEqual(len(self.ctx.new_nums), 1)
        fmt = self.ctx.new_abstract_nums[0].find(f"{_w('lvl')}/{_w('numFmt')}")
        self.assertEqual(fmt.get(_w("val")), "bullet")

    def test_ordered_list_uses_decimal_format(self):
        markdown_to_ooxml.render_blocks("1. first\n2. second\n", self.ctx)
        fmt = self.ctx.new_abstract_nums[-1].find(f"{_w('lvl')}/{_w('numFmt')}")
        self.assertEqual(fmt.get(_w("val")), "decimal")

    def test_nested_list_allocates_a_second_numbering_definition(self):
        markdown_to_ooxml.render_blocks("- top\n    - nested\n", self.ctx)
        self.assertEqual(len(self.ctx.new_abstract_nums), 2)

    def test_table_becomes_wtbl_with_style_and_grid(self):
        md = "| A | B |\n| --- | --- |\n| 1 | 2 |\n"
        elements = markdown_to_ooxml.render_blocks(md, self.ctx)
        self.assertEqual(len(elements), 1)
        tbl = elements[0]
        self.assertEqual(tbl.tag, _w("tbl"))
        tblstyle = tbl.find(f"{_w('tblPr')}/{_w('tblStyle')}")
        self.assertIsNotNone(tblstyle)
        self.assertEqual(tblstyle.get(_w("val")), self.ctx.table_style_id)
        grid_cols = tbl.findall(f"{_w('tblGrid')}/{_w('gridCol')}")
        self.assertEqual(len(grid_cols), 2)
        rows = tbl.findall(_w("tr"))
        self.assertEqual(len(rows), 2)
        first_cell_text = rows[0].find(f"{_w('tc')}/{_w('p')}/{_w('r')}/{_w('t')}")
        self.assertEqual(first_cell_text.text, "A")

    def test_unsupported_constructs_degrade_with_a_warning_not_an_error(self):
        markdown_to_ooxml.render_blocks("---\n\n> quoted text\n\n```\ncode\n```\n", self.ctx)
        self.assertIn("hr_skipped", self.ctx.warnings)
        self.assertIn("blockquote_flattened", self.ctx.warnings)
        self.assertIn("code_block_as_plain_paragraph", self.ctx.warnings)

    def test_softbreak_becomes_a_space_not_a_hard_line_break(self):
        # PR #3 review, should-fix #3: a bare single newline inside a
        # markdown paragraph is a CommonMark/GFM softbreak and renders as
        # a space, not a hard line break — emitting w:br here turned every
        # wrapped markdown line into a hard break in Word.
        md = "Line one\nline two continues the same paragraph.\n"
        elements = markdown_to_ooxml.render_blocks(md, self.ctx)
        self.assertEqual(len(elements), 1)
        p = elements[0]
        self.assertIsNone(p.find(f"{_w('r')}/{_w('br')}"), "a softbreak must not become w:br")
        joined = "".join(t.text or "" for t in p.iter(_w("t")))
        self.assertEqual(joined, "Line one line two continues the same paragraph.")

    def test_hardbreak_still_becomes_a_real_line_break(self):
        # An explicit hard break (trailing two spaces before the newline)
        # is unaffected by the softbreak fix above.
        md = "Line one  \nLine two (hard break above).\n"
        elements = markdown_to_ooxml.render_blocks(md, self.ctx)
        p = elements[0]
        brs = p.findall(f"{_w('r')}/{_w('br')}")
        self.assertEqual(len(brs), 1)


if __name__ == "__main__":
    unittest.main()
