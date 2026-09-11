# New for issue #28 WP-15a. No GoogleDocs-MCP analogue to lift: that
# server's insert_image (markdown_mutations.py:986-1090) compiles a Docs
# API insertInlineImage request against a PUBLIC URL the API itself
# fetches -- a hardcoded 300x200pt objectSize, no width defaulting, no
# EMU math (the Docs API's own objectSize is already in points), and no
# SVG handling anywhere in that server. This module reads a LOCAL PNG/SVG
# file (issue #28's corrections section: "docx drops IMAGE_SOURCE_UNSUPPORTED
# and documents why -- a local path is exactly what this backend reads
# natively") and writes real DrawingML (``w:drawing``/``wp:inline``/
# ``a:graphic``/``pic:pic``) plus the media part + relationship +
# [Content_Types].xml entries a real Word inline picture needs -- none of
# which exists anywhere else in this repo to reuse; markdown_to_ooxml.py's
# StyleContext.relationship_for_link is the closest existing pattern (a
# new external hyperlink relationship) but that relationship is
# TargetMode="External"; an image relationship targets an internal part
# and carries no TargetMode attribute at all, so it is not reused as-is.
"""``insert_image`` -- issue #28 WP-15a.

``insert_image(path, image_path, width_in=None)`` appends a new paragraph
holding one inline picture at the end of the document body. PNG is
embedded directly; SVG is embedded natively (Word 2016+'s
``a14:svgBlip``-in-``a:blip``-``a:extLst`` extension) WITH a PNG fallback
part, because Word requires one for any consumer that does not understand
the SVG extension -- the fallback is rasterized from the SVG's own native
pixel size via macOS's ``sips`` (the same "drive a real macOS tool via
subprocess" posture ``render.py``'s Word-automation path already takes;
this environment has no ``cairosvg``/``rsvg-convert``/Inkscape installed,
and `sips` is confirmed to rasterize SVG correctly, including a
viewBox-only SVG with no explicit width/height -- see this module's own
``_rasterize_svg_to_png``).

Units (verified against real code, not the plan's literal wording -- see
projection.py's own correction comment): ``width_in`` and every size this
module computes in inches are converted to a drawing's own ``wp:extent``
in EMU (914400 per inch) -- NEVER dxa/twips (1440 per inch), which is
page-geometry-only. ``list_page_sections`` (dxa-derived, already in
inches) supplies the DEFAULT ``width_in`` when the caller omits one.

``design_width``/``effective_scale``: for a PNG, ``design_width_in`` is
the file's own intrinsic pixel width (its ``IHDR`` chunk) divided by 96
(the standard CSS/Office "no DPI metadata" assumption -- this server
does not parse a PNG's ``pHYs`` chunk). For an SVG, ``design_width_in``
is the root ``<svg>`` element's own ``width`` attribute (converted to
inches; unitless/``px`` treated as 96 DPI), falling back to its
``viewBox`` width (also treated as px at 96 DPI) when ``width`` is
absent or relative (a `%`). ``effective_scale = placed_width_in /
design_width_in`` is a REAL measured ratio, not a placeholder --
consumed downstream by JennyStack's WP-15b as the mechanical readability
gate ``native_pt * effective_scale >= 8``.
"""

from __future__ import annotations

import re
import struct
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .projection import DEFAULT_PART, R_NS, W_NS

_EMU_PER_INCH = 914400
_PNG_DPI_ASSUMPTION = 96.0

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_SVG_BLIP_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
_SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"  # Word's own extension GUID for an SVG blip

_IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_PICTURE_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"

for _prefix, _uri in (("wp", _WP_NS), ("a", _A_NS), ("pic", _PIC_NS), ("asvg", _SVG_BLIP_NS)):
    ET.register_namespace(_prefix, _uri)


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _wp(tag: str) -> str:
    return f"{{{_WP_NS}}}{tag}"


def _a(tag: str) -> str:
    return f"{{{_A_NS}}}{tag}"


def _pic(tag: str) -> str:
    return f"{{{_PIC_NS}}}{tag}"


# ---------------------------------------------------------------------------
# PNG intrinsic size (IHDR chunk, no external dependency).
# ---------------------------------------------------------------------------

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def _png_dimensions_px(data: bytes) -> tuple[int, int]:
    if not data.startswith(_PNG_SIGNATURE) or len(data) < 24:
        raise _make_error(
            ErrorCode.UNSUPPORTED_IMAGE_FORMAT,
            "File does not start with a valid PNG signature.",
            {},
        )
    width, height = struct.unpack(">II", data[16:24])
    if width <= 0 or height <= 0:
        raise _make_error(ErrorCode.UNSUPPORTED_IMAGE_FORMAT, "PNG IHDR reports a non-positive width/height.", {})
    return width, height


# ---------------------------------------------------------------------------
# SVG intrinsic size (root width/height, falling back to viewBox).
# ---------------------------------------------------------------------------

_SVG_UNIT_TO_IN = {"": 1 / 96, "px": 1 / 96, "in": 1.0, "pt": 1 / 72, "pc": 1 / 6, "mm": 1 / 25.4, "cm": 1 / 2.54}
_SVG_LENGTH_RE = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*([a-zA-Z%]*)\s*$")


def _svg_length_to_in(raw: str | None) -> float | None:
    if not raw:
        return None
    m = _SVG_LENGTH_RE.match(raw)
    if not m:
        return None
    value = float(m.group(1))
    unit = m.group(2).lower()
    if unit == "%":
        return None  # relative to its container -- no absolute design size to derive
    factor = _SVG_UNIT_TO_IN.get(unit)
    return value * factor if factor is not None else None


def _svg_dimensions_in(data: bytes) -> tuple[float, float]:
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise _make_error(ErrorCode.UNSUPPORTED_IMAGE_FORMAT, f"File is not well-formed SVG/XML: {exc}", {}) from exc
    if projection._ln(root) != "svg":
        raise _make_error(ErrorCode.UNSUPPORTED_IMAGE_FORMAT, "File's root element is not <svg>.", {})

    w_in = _svg_length_to_in(root.get("width"))
    h_in = _svg_length_to_in(root.get("height"))
    if w_in is None or h_in is None:
        view_box = root.get("viewBox")
        if view_box:
            parts = view_box.replace(",", " ").split()
            if len(parts) == 4:
                try:
                    vb_w, vb_h = float(parts[2]), float(parts[3])
                except ValueError:
                    vb_w = vb_h = None  # type: ignore[assignment]
                if vb_w and vb_h:
                    w_in = w_in if w_in is not None else vb_w / 96.0
                    h_in = h_in if h_in is not None else vb_h / 96.0
    if not w_in or not h_in:
        raise _make_error(
            ErrorCode.UNSUPPORTED_IMAGE_FORMAT,
            "SVG has no usable absolute width/height or viewBox to derive a design size from.",
            {},
        )
    return w_in, h_in


# ---------------------------------------------------------------------------
# SVG -> PNG fallback rasterization (macOS `sips`; see module docstring).
# ---------------------------------------------------------------------------


def _rasterize_svg_to_png(svg_path: Path) -> bytes:
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "fallback.png"
        try:
            result = subprocess.run(
                ["sips", "-s", "format", "png", str(svg_path), "--out", str(out_path)],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            raise _make_error(
                ErrorCode.SVG_RASTERIZATION_FAILED,
                f"Could not run macOS `sips` to rasterize the SVG PNG fallback Word requires: {exc}",
                {},
            ) from exc
        if result.returncode != 0 or not out_path.exists():
            raise _make_error(
                ErrorCode.SVG_RASTERIZATION_FAILED,
                f"`sips` failed to rasterize the SVG PNG fallback: {result.stderr.strip() or result.stdout.strip()}",
                {"returncode": result.returncode},
            )
        return out_path.read_bytes()


# ---------------------------------------------------------------------------
# Relationship / media / content-type bookkeeping.
# ---------------------------------------------------------------------------

_MEDIA_INDEX_RE = re.compile(r"^image(\d+)\.")
_REL_ID_RE = re.compile(r"^rId(\d+)$")


def _next_media_index(media_names: set[str]) -> int:
    best = 0
    for name in media_names:
        base = name.rsplit("/", 1)[-1]
        m = _MEDIA_INDEX_RE.match(base)
        if m:
            best = max(best, int(m.group(1)))
    return best + 1


def _max_rel_id(rels_root: Any) -> int:
    best = 0
    if rels_root is None:
        return best
    for rel in rels_root:
        m = _REL_ID_RE.match(rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


def _add_relationship(rels_root: Any, *, rid: str, target: str) -> None:
    ET.SubElement(rels_root, "Relationship", {"Id": rid, "Type": _IMAGE_REL_TYPE, "Target": target})


def _ensure_default_extension(ct_root: Any, extension: str, content_type: str) -> None:
    for child in ct_root:
        if projection._ln(child) == "Default" and (child.get("Extension") or "").lower() == extension.lower():
            return
    ET.SubElement(ct_root, "Default", {"Extension": extension, "ContentType": content_type})


def _next_doc_pr_id(document_root: Any) -> int:
    best = 0
    for el in document_root.iter():
        if projection._ln(el) == "docPr":
            raw = projection._attr(el, "id")
            if raw is not None and raw.isdigit():
                best = max(best, int(raw))
    return best + 1


# ---------------------------------------------------------------------------
# DrawingML element construction.
# ---------------------------------------------------------------------------


def _build_drawing(
    *, cx: int, cy: int, blip_rid: str, svg_rid: str | None, doc_pr_id: int, name: str
) -> Any:
    drawing = ET.Element(_w("drawing"))
    inline = ET.SubElement(
        drawing, _wp("inline"), {"distT": "0", "distB": "0", "distL": "0", "distR": "0"}
    )
    ET.SubElement(inline, _wp("extent"), {"cx": str(cx), "cy": str(cy)})
    ET.SubElement(inline, _wp("effectExtent"), {"l": "0", "t": "0", "r": "0", "b": "0"})
    ET.SubElement(inline, _wp("docPr"), {"id": str(doc_pr_id), "name": name})
    cnv_frame_pr = ET.SubElement(inline, _wp("cNvGraphicFramePr"))
    ET.SubElement(cnv_frame_pr, _a("graphicFrameLocks"), {"noChangeAspect": "1"})

    graphic = ET.SubElement(inline, _a("graphic"))
    graphic_data = ET.SubElement(graphic, _a("graphicData"), {"uri": _PICTURE_GRAPHIC_DATA_URI})
    pic = ET.SubElement(graphic_data, _pic("pic"))

    nv_pic_pr = ET.SubElement(pic, _pic("nvPicPr"))
    ET.SubElement(nv_pic_pr, _pic("cNvPr"), {"id": str(doc_pr_id), "name": name})
    ET.SubElement(nv_pic_pr, _pic("cNvPicPr"))

    blip_fill = ET.SubElement(pic, _pic("blipFill"))
    blip = ET.SubElement(blip_fill, _a("blip"), {f"{{{R_NS}}}embed": blip_rid})
    if svg_rid is not None:
        # Word 2016+'s native SVG extension: the a:blip's OWN r:embed still
        # points at the PNG fallback (what every older/non-SVG-aware
        # consumer renders); this extension is what a modern Word actually
        # paints instead, per-spec -- "SVG embedded natively with a PNG
        # fallback part as Word requires" (issue #28 plan WP-15a).
        ext_lst = ET.SubElement(blip, _a("extLst"))
        ext_el = ET.SubElement(ext_lst, _a("ext"), {"uri": _SVG_EXT_URI})
        ET.SubElement(ext_el, f"{{{_SVG_BLIP_NS}}}svgBlip", {f"{{{R_NS}}}embed": svg_rid})
    stretch = ET.SubElement(blip_fill, _a("stretch"))
    ET.SubElement(stretch, _a("fillRect"))

    sp_pr = ET.SubElement(pic, _pic("spPr"))
    xfrm = ET.SubElement(sp_pr, _a("xfrm"))
    ET.SubElement(xfrm, _a("off"), {"x": "0", "y": "0"})
    ET.SubElement(xfrm, _a("ext"), {"cx": str(cx), "cy": str(cy)})
    prst_geom = ET.SubElement(sp_pr, _a("prstGeom"), {"prst": "rect"})
    ET.SubElement(prst_geom, _a("avLst"))

    return drawing


# ---------------------------------------------------------------------------
# Tool: insert_image
# ---------------------------------------------------------------------------

_SUPPORTED_EXTENSIONS = frozenset({"png", "svg"})


def execute_insert_image(
    path: str,
    image_path: str,
    width_in: float | None = None,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    # Same allowlist floor diff_body_vs_file already applies to its own
    # second, non-docx local path (paths.resolve_allowed_docx_path is the
    # server-wide allowlist/denylist, not $DOCS containment -- see that
    # function's own docstring).
    image_resolved = paths.resolve_allowed_docx_path(image_path, must_exist=True)
    if not image_resolved.is_file():
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"Path is not a regular file: {image_path!r}", {"image_path": image_path}
        )
    if width_in is not None and width_in <= 0:
        raise _make_error(ErrorCode.INVALID_INPUT, "width_in must be positive.", {"width_in": width_in})

    ext = image_resolved.suffix.lower().lstrip(".")
    if ext not in _SUPPORTED_EXTENSIONS:
        raise _make_error(
            ErrorCode.UNSUPPORTED_IMAGE_FORMAT,
            f"insert_image supports .png and .svg only, got {image_resolved.suffix!r}.",
            {"image_path": image_path},
        )

    image_bytes = image_resolved.read_bytes()
    if ext == "png":
        design_w_px, design_h_px = _png_dimensions_px(image_bytes)
        design_width_in = design_w_px / _PNG_DPI_ASSUMPTION
        design_height_in = design_h_px / _PNG_DPI_ASSUMPTION
    else:
        design_width_in, design_height_in = _svg_dimensions_in(image_bytes)

    if width_in is None:
        page_sections = projection.list_page_sections_impl(resolved)
        column_widths = page_sections[0].get("column_widths_in") if page_sections else None
        width_in = column_widths[0] if column_widths else design_width_in

    aspect = design_height_in / design_width_in
    placed_width_in = width_in
    placed_height_in = placed_width_in * aspect
    effective_scale = placed_width_in / design_width_in

    cx = max(1, round(placed_width_in * _EMU_PER_INCH))
    cy = max(1, round(placed_height_in * _EMU_PER_INCH))

    document_root, raw_xml = mutations._load_document(resolved)
    body = mutations._find_body(document_root)
    body_children, sect_pr = mutations._split_body(body)

    with zipfile.ZipFile(resolved) as zf:
        names = set(zf.namelist())
        rels_path = projection._rels_path_for(DEFAULT_PART)
        rels_bytes = zf.read(rels_path) if rels_path in names else None
        content_types_bytes = zf.read("[Content_Types].xml")
        media_names = {n for n in names if n.startswith("word/media/")}

    rels_decls = mutations._capture_source_namespaces(rels_bytes) if rels_bytes is not None else {}
    rels_root = (
        ET.fromstring(rels_bytes)
        if rels_bytes is not None
        else ET.Element("Relationships", {"xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"})
    )

    next_rid = _max_rel_id(rels_root) + 1
    next_media_index = _next_media_index(media_names)
    media_overrides: dict[str, bytes] = {}
    svg_rid: str | None = None

    if ext == "png":
        media_filename = f"image{next_media_index}.png"
        media_overrides[f"word/media/{media_filename}"] = image_bytes
        blip_rid = f"rId{next_rid}"
        _add_relationship(rels_root, rid=blip_rid, target=f"media/{media_filename}")
        next_rid += 1
    else:
        svg_filename = f"image{next_media_index}.svg"
        media_overrides[f"word/media/{svg_filename}"] = image_bytes
        png_fallback_bytes = _rasterize_svg_to_png(image_resolved)
        png_filename = f"image{next_media_index + 1}.png"
        media_overrides[f"word/media/{png_filename}"] = png_fallback_bytes

        blip_rid = f"rId{next_rid}"
        _add_relationship(rels_root, rid=blip_rid, target=f"media/{png_filename}")
        next_rid += 1
        svg_rid = f"rId{next_rid}"
        _add_relationship(rels_root, rid=svg_rid, target=f"media/{svg_filename}")
        next_rid += 1

    doc_pr_id = _next_doc_pr_id(document_root)
    drawing = _build_drawing(
        cx=cx, cy=cy, blip_rid=blip_rid, svg_rid=svg_rid, doc_pr_id=doc_pr_id, name=image_resolved.name
    )
    run = ET.Element(_w("r"))
    run.append(drawing)
    p = ET.Element(_w("p"))
    p.append(run)

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # Nothing existing is removed by an insertion -- only the new
        # run's own content is wrapped in w:ins (mirrors insert_table's
        # identical treatment of a pure insertion).
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.wrap_all_runs([p], _ins_wrap)

    insert_at = len(body_children)
    body.insert(insert_at, p)
    if sect_pr is not None:
        body.remove(sect_pr)
        body.append(sect_pr)

    document_decls = mutations._capture_source_namespaces(raw_xml)
    overrides: dict[str, bytes] = {DEFAULT_PART: mutations._serialize_xml(document_root, document_decls)}
    overrides[rels_path] = mutations._serialize_xml(rels_root, rels_decls)
    overrides.update(media_overrides)

    ct_decls = mutations._capture_source_namespaces(content_types_bytes)
    ct_root = ET.fromstring(content_types_bytes)
    _ensure_default_extension(ct_root, "png", "image/png")
    if ext == "svg":
        _ensure_default_extension(ct_root, "svg", "image/svg+xml")
    overrides["[Content_Types].xml"] = mutations._serialize_xml(ct_root, ct_decls)

    def _post_verify(written_path: Path) -> None:
        with zipfile.ZipFile(written_path) as zf:
            written_names = set(zf.namelist())
        missing = [name for name in media_overrides if name not in written_names]
        if missing:
            raise ValueError(f"media part(s) missing after write: {missing}")
        proj = projection.project_part(written_path)
        found = any(
            isinstance(event, projection.DrawingEvent) and event.blip_rid == blip_rid for event in proj.events
        )
        if not found:
            raise ValueError(f"re-read document does not show a drawing with blip_rid {blip_rid!r}")

    conflict_sweep = mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": 1,
        "rung": 4,
        "before": "",
        "after": f"[image:{blip_rid}]",
        "revision_before": pre_revision["token"],
        "revision_after": post_revision["token"],
        "audit_logged": False,
    }
    if track is not None:
        evidence["track_changes"] = True
        evidence["revision_ids"] = track.revision_ids
    mutations._merge_conflict_sweep(evidence, conflict_sweep)
    evidence["format"] = ext
    evidence["design_width_in"] = round(design_width_in, 4)
    evidence["design_height_in"] = round(design_height_in, 4)
    evidence["placed_width_in"] = round(placed_width_in, 4)
    evidence["placed_height_in"] = round(placed_height_in, 4)
    evidence["effective_scale"] = round(effective_scale, 6)
    evidence["blip_rid"] = blip_rid
    if svg_rid is not None:
        evidence["svg_rid"] = svg_rid

    logged, _ = audit.append_audit(path=str(resolved), tool="insert_image", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
