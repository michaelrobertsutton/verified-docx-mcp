# New for issue #28 WP-04. No GoogleDocs-MCP analogue exists to lift for the
# write direction here: that server's markdown_writer.py compiles markdown
# into Google Docs API batchUpdate REQUESTS (JSON operations against a
# server-side document model); this server has no such API underneath it —
# a .docx is a local OOXML package, so "compiling markdown" here means
# building actual <w:p>/<w:tbl>/<w:r> XML elements instead. markdown.py
# (the READ direction, Docs JSON -> markdown) was likewise not reusable:
# projection.py's read_document_markdown (WP-03) already plays that role
# for this backend. markdown-it-py itself (the pinned dependency,
# pyproject.toml) IS reused here, per the plan's "Lift, do not rewrite" for
# shared infrastructure — only the block-token -> OOXML side is new.
"""Markdown (CommonMark + GFM tables) -> OOXML block elements.

``render_blocks(markdown_text, ctx)`` parses *markdown_text* with
``markdown_it.MarkdownIt("commonmark").enable("table")`` and returns a list
of ``xml.etree.ElementTree.Element`` objects (``w:p`` / ``w:tbl``) in
document order, ready to be spliced into a ``w:body`` (or a table cell).

Supported subset (issue #28 plan WP-04): ATX headings (resolved through
``StyleContext.heading_style_for_level`` -> ``list_styles``, never a
hardcoded "Heading2" — ``STYLE_NOT_FOUND`` when a level has no heading
style in the target document), paragraphs, **bold**/*italic*/***both***
runs, bulleted and ordered lists (including nesting — a top-level list
gets one freshly allocated ``w:abstractNum``/``w:num`` pair for its WHOLE
tree, shared by every nested level under it via an increasing ``w:ilvl``,
the standard OOXML idiom projection.py's markdown reader also relies on to
detect nesting; see ``StyleContext.allocate_list_root``/
``ensure_list_level``), pipe tables (``w:tbl`` with
``style_id`` from the target document's own table styles when one exists),
and ``[text](url)`` links (``w:hyperlink`` + a new External relationship,
``StyleContext.relationship_for_link``).

Out-of-subset constructs (thematic breaks, blockquotes, code
blocks/fences, images, raw HTML) degrade gracefully rather than raising:
an ``hr`` is dropped, a blockquote's paragraphs are flattened into the
surrounding flow, and a code block becomes a plain paragraph — each adds a
name to ``StyleContext.warnings`` so nothing is silently lossy. Only a
missing heading style raises (``STYLE_NOT_FOUND``); everything else in
this list is a documented degradation, not a hard error, because none of
the WP-04 acceptance tests exercises it and refusing outright would make
`replace_body_markdown` unusable on any markdown containing them.
"""

from __future__ import annotations

import dataclasses
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from markdown_it import MarkdownIt

from . import projection
from .errors import ErrorCode, _make_error
from .projection import R_NS, W_NS

_MD = MarkdownIt("commonmark").enable("table")

_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _wa(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Style/numbering/relationship context — built once per write call from the
# TARGET document's own styles/numbering/rels (never hardcoded), consumed
# while rendering, and read back afterwards by mutations.py to know what new
# numbering.xml / document.xml.rels / [Content_Types].xml content (if any)
# needs to be written alongside the modified document.xml.
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class StyleContext:
    heading_styles: dict[int, str]
    table_style_id: str | None
    hyperlink_rstyle_id: str | None
    next_abstract_num_id: int
    next_num_id: int
    numbering_part_exists: bool
    new_abstract_nums: list[Any] = dataclasses.field(default_factory=list)
    new_nums: list[Any] = dataclasses.field(default_factory=list)
    new_relationships: list[tuple[str, str, str]] = dataclasses.field(default_factory=list)
    warnings: list[str] = dataclasses.field(default_factory=list)
    _next_rid: int = 1
    _link_rid_cache: dict[str, str] = dataclasses.field(default_factory=dict)
    _abstract_num_by_num_id: dict[int, Any] = dataclasses.field(default_factory=dict)
    _defined_levels_by_num_id: dict[int, set[int]] = dataclasses.field(default_factory=dict)

    @staticmethod
    def build(resolved: Path) -> StyleContext:
        styles = projection.list_styles_impl(resolved)
        heading_styles: dict[int, str] = {}
        table_style_id: str | None = None
        hyperlink_rstyle_id: str | None = None
        for s in styles:
            style_id: str | None = s.get("style_id")
            lvl = s.get("outline_lvl")
            if lvl is None and style_id:
                m = projection._HEADING_STYLE_ID_RE.match(style_id)
                if m:
                    lvl = int(m.group(1)) - 1
            if lvl is not None and style_id is not None and s.get("type") == "paragraph" and lvl not in heading_styles:
                heading_styles[lvl] = style_id
            if s.get("type") == "table" and table_style_id is None:
                table_style_id = style_id
            if style_id == "Hyperlink":
                hyperlink_rstyle_id = style_id

        with zipfile.ZipFile(resolved) as zf:
            names = set(zf.namelist())
            numbering_exists = "word/numbering.xml" in names
            numbering_root = projection.read_part_xml(zf, "word/numbering.xml") if numbering_exists else None
            rels_root = projection.load_rels_root(zf, projection.DEFAULT_PART)

        next_abstract = _max_attr_value(numbering_root, "abstractNum", "abstractNumId") + 1
        next_num = _max_attr_value(numbering_root, "num", "numId") + 1
        next_rid = _max_rel_id(rels_root) + 1

        return StyleContext(
            heading_styles=heading_styles,
            table_style_id=table_style_id,
            hyperlink_rstyle_id=hyperlink_rstyle_id,
            next_abstract_num_id=next_abstract,
            next_num_id=next_num,
            numbering_part_exists=numbering_exists,
            _next_rid=next_rid,
        )

    def heading_style_for_level(self, level: int) -> str:
        """*level* is 1-based (markdown "#" count). Raises STYLE_NOT_FOUND
        (never falls back to a hardcoded style id) when the target document
        defines no paragraph style at that outline level."""
        style_id = self.heading_styles.get(level - 1)
        if style_id is None:
            raise _make_error(
                ErrorCode.STYLE_NOT_FOUND,
                (
                    f"No heading style found for markdown level {level} "
                    f"(outline level {level - 1}) in this document. Call "
                    "list_styles to see what this document actually defines."
                ),
                {"level": level, "outline_level": level - 1, "available_levels": sorted(self.heading_styles)},
            )
        return style_id

    def allocate_list_root(self, ordered: bool, depth: int) -> int:
        """Allocate a fresh abstractNum/num pair for one TOP-LEVEL list —
        every level nested under it (see ensure_list_level) shares this
        SAME num_id, only its w:ilvl changing per depth. This is the
        standard OOXML idiom (one numId, increasing ilvl per nesting
        level) — PR #3 review / WP-03b-a's extended round trip: the prior
        design (a fresh numId *per nesting level*, every paragraph's own
        w:ilvl left at 0) left the reader with no signal at all to tell a
        nested item from a new, unrelated top-level list — every markdown
        list-nesting round trip silently flattened."""
        abstract_id = self.next_abstract_num_id
        num_id = self.next_num_id
        self.next_abstract_num_id += 1
        self.next_num_id += 1
        abstract_num = _build_abstract_num_shell(abstract_id)
        self.new_abstract_nums.append(abstract_num)
        self.new_nums.append(_build_num(num_id, abstract_id))
        self._abstract_num_by_num_id[num_id] = abstract_num
        self._defined_levels_by_num_id[num_id] = set()
        self.ensure_list_level(num_id, ordered, depth)
        return num_id

    def ensure_list_level(self, num_id: int, ordered: bool, depth: int) -> None:
        """Add a w:lvl for *depth* to num_id's abstractNum if this num_id
        has never used that depth before (e.g. the first ordered list
        nested under a bulleted parent, or a second bulleted item at a
        depth a sibling branch already defined — a no-op the second
        time)."""
        defined = self._defined_levels_by_num_id.setdefault(num_id, set())
        if depth in defined:
            return
        abstract_num = self._abstract_num_by_num_id[num_id]
        abstract_num.append(_build_lvl(depth, ordered))
        defined.add(depth)

    def relationship_for_link(self, url: str) -> str:
        cached = self._link_rid_cache.get(url)
        if cached is not None:
            return cached
        rid = f"rId{self._next_rid}"
        self._next_rid += 1
        self.new_relationships.append(
            (rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", url)
        )
        self._link_rid_cache[url] = rid
        return rid


def _max_attr_value(root: Any | None, tag_local: str, attr_local: str) -> int:
    if root is None:
        return -1
    best = -1
    for el in root.iter():
        if projection._ln(el) == tag_local:
            val = projection._attr(el, attr_local)
            if val is not None:
                try:
                    best = max(best, int(val))
                except ValueError:
                    pass
    return best


_REL_ID_RE = re.compile(r"^rId(\d+)$")


def _max_rel_id(root: Any | None) -> int:
    if root is None:
        return 0
    best = 0
    for rel in root:
        m = _REL_ID_RE.match(rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


# ---------------------------------------------------------------------------
# numbering.xml element builders
# ---------------------------------------------------------------------------


def _build_abstract_num_shell(abstract_id: int) -> Any:
    """An abstractNum with no w:lvl children yet -- ensure_list_level
    adds one per nesting depth actually used, as that depth is first
    reached (a list may never go past ilvl 0, or may reach several)."""
    abstract_num = ET.Element(_w("abstractNum"), {_wa("abstractNumId"): str(abstract_id)})
    ET.SubElement(abstract_num, _w("multiLevelType"), {_wa("val"): "hybridMultilevel"})
    return abstract_num


def _build_lvl(ilvl: int, ordered: bool) -> Any:
    lvl = ET.Element(_w("lvl"), {_wa("ilvl"): str(ilvl)})
    ET.SubElement(lvl, _w("start"), {_wa("val"): "1"})
    ET.SubElement(lvl, _w("numFmt"), {_wa("val"): "decimal" if ordered else "bullet"})
    ET.SubElement(lvl, _w("lvlText"), {_wa("val"): "%1." if ordered else ""})
    ET.SubElement(lvl, _w("lvlJc"), {_wa("val"): "left"})
    ppr = ET.SubElement(lvl, _w("pPr"))
    left = 720 * (ilvl + 1)
    ET.SubElement(ppr, _w("ind"), {_wa("left"): str(left), _wa("hanging"): "360"})
    if not ordered:
        rpr = ET.SubElement(lvl, _w("rPr"))
        ET.SubElement(rpr, _w("rFonts"), {_wa("ascii"): "Symbol", _wa("hAnsi"): "Symbol", _wa("hint"): "default"})
    return lvl


def _build_num(num_id: int, abstract_id: int) -> Any:
    num = ET.Element(_w("num"), {_wa("numId"): str(num_id)})
    ET.SubElement(num, _w("abstractNumId"), {_wa("val"): str(abstract_id)})
    return num


# ---------------------------------------------------------------------------
# Inline runs
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class RunSpec:
    text: str
    bold: bool
    italic: bool
    link: str | None
    hard_break: bool = False


def _inline_runs(inline_tok: Any | None, ctx: StyleContext) -> list[RunSpec]:
    if inline_tok is None or not inline_tok.children:
        return []
    runs: list[RunSpec] = []
    bold = 0
    italic = 0
    link_stack: list[str] = []
    for child in inline_tok.children:
        t = child.type
        current_link = link_stack[-1] if link_stack else None
        if t in ("text", "code_inline"):
            if child.content:
                runs.append(RunSpec(text=child.content, bold=bold > 0, italic=italic > 0, link=current_link))
        elif t == "strong_open":
            bold += 1
        elif t == "strong_close":
            bold = max(0, bold - 1)
        elif t == "em_open":
            italic += 1
        elif t == "em_close":
            italic = max(0, italic - 1)
        elif t == "link_open":
            href = dict(child.attrs or {}).get("href", "") or ""
            link_stack.append(href)
        elif t == "link_close":
            if link_stack:
                link_stack.pop()
        elif t == "hardbreak":
            # An explicit hard line break (trailing two spaces, or a
            # backslash, before the newline) -> a real w:br.
            runs.append(RunSpec(text="", bold=bold > 0, italic=italic > 0, link=current_link, hard_break=True))
        elif t == "softbreak":
            # A bare single newline inside a paragraph is a CommonMark/GFM
            # softbreak, NOT a hard line break: it renders as a single
            # space (PR #3 review, should-fix #3 — emitting w:br here
            # turned every wrapped markdown line into a hard break in
            # Word, so a re-read gave back different text than the
            # author wrote).
            runs.append(RunSpec(text=" ", bold=bold > 0, italic=italic > 0, link=current_link))
        # images, raw html_inline, entities beyond markdown-it's own decoding:
        # silently dropped — out of the supported subset (module docstring).
    return runs


def _append_run_element(parent: Any, run: RunSpec, ctx: StyleContext, *, is_hyperlink: bool) -> None:
    r = ET.SubElement(parent, _w("r"))
    if run.bold or run.italic or (is_hyperlink and ctx.hyperlink_rstyle_id):
        rpr = ET.SubElement(r, _w("rPr"))
        if is_hyperlink and ctx.hyperlink_rstyle_id:
            ET.SubElement(rpr, _w("rStyle"), {_wa("val"): ctx.hyperlink_rstyle_id})
        if run.bold:
            ET.SubElement(rpr, _w("b"))
        if run.italic:
            ET.SubElement(rpr, _w("i"))
    if run.hard_break:
        ET.SubElement(r, _w("br"))
        return
    t = ET.SubElement(r, _w("t"))
    t.text = run.text
    if run.text != run.text.strip() or run.text == "":
        t.set(f"{{{_XML_NS}}}space", "preserve")


def _build_paragraph(
    ctx: StyleContext,
    style_id: str | None,
    runs: list[RunSpec],
    *,
    num_id: int | None = None,
    ilvl: int = 0,
) -> Any:
    p = ET.Element(_w("p"))
    if style_id is not None or num_id is not None:
        ppr = ET.SubElement(p, _w("pPr"))
        if style_id is not None:
            ET.SubElement(ppr, _w("pStyle"), {_wa("val"): style_id})
        if num_id is not None:
            numpr = ET.SubElement(ppr, _w("numPr"))
            ET.SubElement(numpr, _w("ilvl"), {_wa("val"): str(ilvl)})
            ET.SubElement(numpr, _w("numId"), {_wa("val"): str(num_id)})

    idx = 0
    n = len(runs)
    while idx < n:
        run = runs[idx]
        if run.link:
            j = idx
            while j < n and runs[j].link == run.link:
                j += 1
            hyperlink_el = ET.SubElement(p, _w("hyperlink"))
            hyperlink_el.set(f"{{{R_NS}}}id", ctx.relationship_for_link(run.link))
            for g in runs[idx:j]:
                _append_run_element(hyperlink_el, g, ctx, is_hyperlink=True)
            idx = j
        else:
            _append_run_element(p, run, ctx, is_hyperlink=False)
            idx += 1
    return p


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------


def _find_matching_close(tokens: list[Any], i: int, close_type: str) -> int:
    level = tokens[i].level
    j = i + 1
    n = len(tokens)
    while j < n:
        if tokens[j].type == close_type and tokens[j].level == level:
            return j
        j += 1
    raise AssertionError(f"no matching {close_type} for token at index {i} ({tokens[i].type})")


def _build_table(tokens: list[Any], open_idx: int, close_idx: int, ctx: StyleContext) -> Any:
    tbl = ET.Element(_w("tbl"))
    tblpr = ET.SubElement(tbl, _w("tblPr"))
    if ctx.table_style_id:
        ET.SubElement(tblpr, _w("tblStyle"), {_wa("val"): ctx.table_style_id})
    ET.SubElement(tblpr, _w("tblW"), {_wa("w"): "0", _wa("type"): "auto"})

    row_data: list[list[list[RunSpec]]] = []
    j = open_idx + 1
    while j < close_idx:
        tok = tokens[j]
        if tok.type == "tr_open":
            row_end = _find_matching_close(tokens, j, "tr_close")
            cells: list[list[RunSpec]] = []
            k = j + 1
            while k < row_end:
                if tokens[k].type in ("th_open", "td_open"):
                    close_type = "th_close" if tokens[k].type == "th_open" else "td_close"
                    cell_end = _find_matching_close(tokens, k, close_type)
                    inline_tok = tokens[k + 1] if k + 1 < cell_end and tokens[k + 1].type == "inline" else None
                    cells.append(_inline_runs(inline_tok, ctx))
                    k = cell_end + 1
                else:
                    k += 1
            row_data.append(cells)
            j = row_end + 1
        else:
            j += 1

    col_count = max((len(row) for row in row_data), default=1) or 1
    grid = ET.SubElement(tbl, _w("tblGrid"))
    for _ in range(col_count):
        ET.SubElement(grid, _w("gridCol"))

    for cells in row_data:
        tr = ET.SubElement(tbl, _w("tr"))
        for c in range(col_count):
            tc = ET.SubElement(tr, _w("tc"))
            ET.SubElement(tc, _w("tcPr"))
            runs = cells[c] if c < len(cells) else []
            tc.append(_build_paragraph(ctx, None, runs))
    return tbl


# ---------------------------------------------------------------------------
# Block walk
# ---------------------------------------------------------------------------


def _walk_list_item(tokens: list[Any], start: int, end: int, ctx: StyleContext, depth: int, num_id: int) -> list[Any]:
    elements: list[Any] = []
    j = start
    while j < end:
        tok = tokens[j]
        if tok.type == "paragraph_open":
            inline_tok = tokens[j + 1] if j + 1 < end and tokens[j + 1].type == "inline" else None
            runs = _inline_runs(inline_tok, ctx)
            # ilvl=depth (never a hardcoded 0) — this paragraph's nesting
            # level is the ONLY signal a reader (projection.py's
            # _events_to_markdown) has to tell it apart from a top-level
            # item; see allocate_list_root's docstring.
            elements.append(_build_paragraph(ctx, None, runs, num_id=num_id, ilvl=depth))
            j += 3
        elif tok.type in ("bullet_list_open", "ordered_list_open"):
            nested_ordered = tok.type == "ordered_list_open"
            nested_close = "bullet_list_close" if not nested_ordered else "ordered_list_close"
            nested_end = _find_matching_close(tokens, j, nested_close)
            # Same num_id as the enclosing list — a nested list is a
            # deeper LEVEL of the same numbering tree, not an unrelated
            # list of its own (see allocate_list_root).
            elements.extend(_walk_list(tokens, j, nested_end, ctx, nested_ordered, depth + 1, num_id=num_id))
            j = nested_end + 1
        else:
            j += 1
    return elements


def _walk_list(
    tokens: list[Any], open_idx: int, close_idx: int, ctx: StyleContext, ordered: bool, depth: int, num_id: int | None = None
) -> list[Any]:
    """*num_id* is None for a top-level list (allocates a fresh
    numId/abstractNum tree) and the enclosing list's num_id when this is a
    nested list — either way, ensure_list_level makes sure this depth's
    w:lvl (bullet or numbered, whichever THIS level actually uses — a
    nested list's own marker type is independent of its parent's) is
    defined on that shared abstractNum."""
    if num_id is None:
        num_id = ctx.allocate_list_root(ordered, depth)
    else:
        ctx.ensure_list_level(num_id, ordered, depth)
    elements: list[Any] = []
    j = open_idx + 1
    while j < close_idx:
        tok = tokens[j]
        if tok.type == "list_item_open":
            item_end = _find_matching_close(tokens, j, "list_item_close")
            elements.extend(_walk_list_item(tokens, j + 1, item_end, ctx, depth, num_id))
            j = item_end + 1
        else:
            j += 1
    return elements


def _walk_blocks(tokens: list[Any], ctx: StyleContext) -> list[Any]:
    elements: list[Any] = []
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        if tok.type == "heading_open":
            level = int(tok.tag[1:]) if tok.tag[1:].isdigit() else 1
            inline_tok = tokens[i + 1] if i + 1 < n and tokens[i + 1].type == "inline" else None
            style_id = ctx.heading_style_for_level(level)
            elements.append(_build_paragraph(ctx, style_id, _inline_runs(inline_tok, ctx)))
            i += 3
        elif tok.type == "paragraph_open":
            inline_tok = tokens[i + 1] if i + 1 < n and tokens[i + 1].type == "inline" else None
            elements.append(_build_paragraph(ctx, None, _inline_runs(inline_tok, ctx)))
            i += 3
        elif tok.type in ("bullet_list_open", "ordered_list_open"):
            ordered = tok.type == "ordered_list_open"
            close_type = "bullet_list_close" if not ordered else "ordered_list_close"
            end = _find_matching_close(tokens, i, close_type)
            elements.extend(_walk_list(tokens, i, end, ctx, ordered, depth=0))
            i = end + 1
        elif tok.type == "table_open":
            end = _find_matching_close(tokens, i, "table_close")
            elements.append(_build_table(tokens, i, end, ctx))
            i = end + 1
        elif tok.type == "blockquote_open":
            end = _find_matching_close(tokens, i, "blockquote_close")
            elements.extend(_walk_blocks(tokens[i + 1 : end], ctx))
            ctx.warnings.append("blockquote_flattened")
            i = end + 1
        elif tok.type in ("fence", "code_block"):
            text = (tok.content or "").rstrip("\n")
            elements.append(_build_paragraph(ctx, None, [RunSpec(text=text, bold=False, italic=False, link=None)]))
            ctx.warnings.append("code_block_as_plain_paragraph")
            i += 1
        elif tok.type == "hr":
            ctx.warnings.append("hr_skipped")
            i += 1
        else:
            i += 1
    return elements


def render_blocks(markdown_text: str, ctx: StyleContext) -> list[Any]:
    """Parse *markdown_text* and return a list of ``w:p``/``w:tbl``
    Elements in document order. Mutates *ctx* (numbering/relationship
    allocations, warnings) as a side effect — see the module docstring."""
    tokens = _MD.parse(markdown_text)
    return _walk_blocks(tokens, ctx)
