"""Unit tests for src/verified_docx_mcp/images.py: insert_image (issue #28
WP-15a).

Fixture provenance: tests/fixtures/images/sample.png and sample.svg are
synthetic, not Word-authored -- images.py's own module docstring explains
why that is the right call here (unlike every other fixture in this repo,
these are USER-SUPPLIED assets insert_image reads verbatim; they are not
a claim about a Word-output nuance this repo could get wrong by
guessing). sample.png (300x150px, solid color) is built with pure-Python
struct/zlib (no external imaging library is installed in this
environment); sample.svg (200x100 viewBox, a rect + text) is hand-authored
XML text. Both are committed so a test run needs no image tooling to
regenerate them.
"""

from __future__ import annotations

import os
import shutil
import struct
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "src"))

from verified_docx_mcp import images, mutations, paths, projection
from verified_docx_mcp.errors import ErrorCode, VerifyError
from verified_docx_mcp.middleware import MUTATING_TOOLS

FIXTURES = REPO / "tests" / "fixtures"
IMAGES = FIXTURES / "images"

mutations._QUIESCE_INTERVAL_SECONDS = 0.02

_EVIDENCE_KEYS = {
    "applied",
    "match_count",
    "rung",
    "before",
    "after",
    "revision_before",
    "revision_after",
    "audit_logged",
    "conflict_copy_detected",
}


class _TempFixtureCase(unittest.TestCase):
    fixture_name = "tables.docx"

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_allowed = os.environ.get(paths._ALLOWED_FILE_ROOTS_ENV)
        # Two roots: the isolated temp dir (the .docx under test) and the
        # repo's own image fixtures directory (the source images -- never
        # written to, only read).
        os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = os.pathsep.join([self._tmp.name, str(IMAGES)])
        self.target = Path(self._tmp.name) / Path(self.fixture_name).name
        shutil.copyfile(FIXTURES / self.fixture_name, self.target)

    def tearDown(self):
        if self._old_allowed is None:
            os.environ.pop(paths._ALLOWED_FILE_ROOTS_ENV, None)
        else:
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = self._old_allowed
        self._tmp.cleanup()


class MutatingToolsRegistrationTests(unittest.TestCase):
    def test_insert_image_registered(self):
        self.assertIn("insert_image", MUTATING_TOOLS)


class PngIntrinsicSizeTests(unittest.TestCase):
    def test_ihdr_dimensions(self):
        data = (IMAGES / "sample.png").read_bytes()
        w, h = images._png_dimensions_px(data)
        self.assertEqual((w, h), (300, 150))

    def test_rejects_non_png_bytes(self):
        with self.assertRaises(VerifyError) as ctx:
            images._png_dimensions_px(b"not a png")
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.UNSUPPORTED_IMAGE_FORMAT)


class SvgIntrinsicSizeTests(unittest.TestCase):
    def test_width_height_attrs(self):
        data = (IMAGES / "sample.svg").read_bytes()
        w_in, h_in = images._svg_dimensions_in(data)
        self.assertAlmostEqual(w_in, 200 / 96)
        self.assertAlmostEqual(h_in, 100 / 96)

    def test_viewbox_fallback_when_width_absent(self):
        svg = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200"><rect/></svg>'
        w_in, h_in = images._svg_dimensions_in(svg)
        self.assertAlmostEqual(w_in, 400 / 96)
        self.assertAlmostEqual(h_in, 200 / 96)

    def test_no_usable_size_raises(self):
        svg = b'<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>'
        with self.assertRaises(VerifyError) as ctx:
            images._svg_dimensions_in(svg)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.UNSUPPORTED_IMAGE_FORMAT)


class InsertImagePngTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_default_width_uses_text_column_width(self):
        evidence = images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"))
        self.assertTrue(_EVIDENCE_KEYS.issubset(evidence.keys()))
        self.assertEqual(evidence["format"], "png")
        # tables.docx: 8.5in page, 1in margins each side -> 6.5in usable,
        # single column -> list_page_sections' own column_widths_in[0].
        self.assertAlmostEqual(evidence["placed_width_in"], 6.5)
        self.assertAlmostEqual(evidence["design_width_in"], 300 / 96)
        self.assertAlmostEqual(evidence["design_height_in"], 150 / 96)
        expected_scale = 6.5 / (300 / 96)
        self.assertAlmostEqual(evidence["effective_scale"], round(expected_scale, 6))

        with zipfile.ZipFile(self.target) as zf:
            self.assertIn("word/media/image1.png", zf.namelist())
            self.assertEqual(zf.read("word/media/image1.png"), (IMAGES / "sample.png").read_bytes())
        self.assertTrue(mutations.opc_valid(self.target)[0])

    def test_explicit_width_in_changes_placed_size_not_design_size(self):
        evidence = images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"), 2.0)
        self.assertAlmostEqual(evidence["placed_width_in"], 2.0)
        self.assertAlmostEqual(evidence["design_width_in"], 300 / 96)
        self.assertAlmostEqual(evidence["effective_scale"], round(2.0 / (300 / 96), 6))
        # Height follows the image's own aspect ratio (2:1 here).
        self.assertAlmostEqual(evidence["placed_height_in"], 1.0)

    def test_extent_round_trips_through_emu(self):
        images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"), 3.0)
        proj = projection.project_part(self.target)
        drawing_events = [e for e in proj.events if isinstance(e, projection.DrawingEvent)]
        self.assertEqual(len(drawing_events), 1)
        # projection.py's own EMU->inch conversion (914400/inch) applied
        # to the exact EMU this module wrote must round-trip to the
        # placed size -- proves insert_image used EMU (not dxa) for the
        # drawing's own wp:extent.
        self.assertAlmostEqual(drawing_events[0].extent_in[0], 3.0, places=3)
        self.assertAlmostEqual(drawing_events[0].extent_in[1], 1.5, places=3)

    def test_unsupported_extension_refuses(self):
        with tempfile.TemporaryDirectory() as td:
            bogus = Path(td) / "not-an-image.gif"
            bogus.write_bytes(b"GIF89a")
            os.environ[paths._ALLOWED_FILE_ROOTS_ENV] = os.pathsep.join(
                [os.environ[paths._ALLOWED_FILE_ROOTS_ENV], td]
            )
            with self.assertRaises(VerifyError) as ctx:
                images.execute_insert_image(str(self.target), str(bogus))
            self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.UNSUPPORTED_IMAGE_FORMAT)

    def test_negative_width_refuses(self):
        with self.assertRaises(VerifyError) as ctx:
            images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"), -1.0)
        self.assertEqual(ctx.exception.envelope.error_code, ErrorCode.INVALID_INPUT)

    def test_track_changes_wraps_new_run(self):
        evidence = images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"), track_changes=True)
        self.assertTrue(evidence["track_changes"])
        self.assertTrue(evidence["revision_ids"])
        with zipfile.ZipFile(self.target) as zf:
            xml = zf.read(projection.DEFAULT_PART).decode("utf-8")
        self.assertIn("<w:ins ", xml)

    def test_second_insert_gets_a_fresh_media_index_and_rid(self):
        ev1 = images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"))
        ev2 = images.execute_insert_image(str(self.target), str(IMAGES / "sample.png"))
        self.assertNotEqual(ev1["blip_rid"], ev2["blip_rid"])
        with zipfile.ZipFile(self.target) as zf:
            media = sorted(n for n in zf.namelist() if n.startswith("word/media/"))
        self.assertEqual(media, ["word/media/image1.png", "word/media/image2.png"])


class InsertImageSvgTests(_TempFixtureCase):
    fixture_name = "tables.docx"

    def test_svg_embeds_natively_with_png_fallback(self):
        evidence = images.execute_insert_image(str(self.target), str(IMAGES / "sample.svg"), 4.0)
        self.assertEqual(evidence["format"], "svg")
        self.assertIn("svg_rid", evidence)
        self.assertNotEqual(evidence["svg_rid"], evidence["blip_rid"])
        self.assertAlmostEqual(evidence["design_width_in"], round(200 / 96, 4))
        self.assertAlmostEqual(evidence["placed_width_in"], 4.0)
        self.assertAlmostEqual(evidence["effective_scale"], round(4.0 / (200 / 96), 6))

        with zipfile.ZipFile(self.target) as zf:
            names = zf.namelist()
            self.assertIn("word/media/image1.svg", names)
            self.assertIn("word/media/image2.png", names)  # the fallback
            self.assertEqual(zf.read("word/media/image1.svg"), (IMAGES / "sample.svg").read_bytes())
            fallback_bytes = zf.read("word/media/image2.png")
            fallback_w, fallback_h = struct.unpack(">II", fallback_bytes[16:24])
            self.assertEqual((fallback_w, fallback_h), (200, 100))
            ct_bytes = zf.read("[Content_Types].xml")
        self.assertIn(b'Extension="svg"', ct_bytes)
        self.assertIn(b'Extension="png"', ct_bytes)
        self.assertTrue(mutations.opc_valid(self.target)[0])

    def test_blip_extension_points_at_svg_relationship(self):
        evidence = images.execute_insert_image(str(self.target), str(IMAGES / "sample.svg"))
        with zipfile.ZipFile(self.target) as zf:
            doc_xml = zf.read(projection.DEFAULT_PART)
        root = ET.fromstring(doc_xml)
        svg_ns = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
        r_ns = projection.R_NS
        found_svg_blip_rid = None
        for el in root.iter():
            if el.tag == f"{{{svg_ns}}}svgBlip":
                found_svg_blip_rid = el.get(f"{{{r_ns}}}embed")
        self.assertEqual(found_svg_blip_rid, evidence["svg_rid"])


if __name__ == "__main__":
    unittest.main()
