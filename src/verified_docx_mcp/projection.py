# New for issue #28 WP-03. No GoogleDocs-MCP analogue exists to lift here:
# the Google server projects the Docs API's own JSON tree (docs.py), which
# has no OOXML package underneath it. This module is the docx-side
# equivalent — it walks the raw XML of one package part into a flat text
# projection, an offset map, and a document-order block list — and is what
# WP-06's locate() (lifted from GoogleDocs-MCP verify.py) will run against
# once it lands. The revision-token, find_sections/list_page_sections/
# list_styles/list_parts logic below is new; it has no Google-side
# counterpart either (Google Docs carries its own server-side revision id).
"""OOXML projection: walk one package part into a flat string + offset map
+ document-order block list, plus the read tools built on top of it
(list_parts, read_document, find_sections, list_page_sections,
list_styles) and the revision token.

Scope = one part (issue #28 plan WP-03). ``word/document.xml``'s ``w:body``
is the default; a header, footer, footnotes, or endnotes part is addressed
by its own part name (``part=`` on every tool below). A text box's
``w:txbxContent`` is projected as its own sub-scope with its own
``section_key``, never merged into the body's offsets — see
``iter_textbox_scopes``.

Paragraph walk order (issue #28 plan WP-03, "Codex defects folded in"):
every ``w:p`` in document order, including paragraphs inside
``w:tbl/w:tr/w:tc``, ``w:sdt/w:sdtContent`` (block-level), and text that is
wrapped (at the run level, inside an otherwise-ordinary paragraph) in
``w:smartTag``, ``w:hyperlink``, ``w:ins``, an inline ``w:sdt``, or
``w:fldSimple``. Each table-cell paragraph records its container chain
(table id, row, cell) so a future structural-boundary check (WP-06) can be
computed from it.

Text mapping: ``w:t`` verbatim; ``w:tab`` -> ``\\t``; ``w:br``/``w:cr`` ->
``\\n``; ``w:noBreakHyphen`` -> ``-``; ``w:softHyphen`` -> U+00AD.
``w:del``/``w:delText`` is EXCLUDED from the projection (its span is
recorded separately, ``deleted_spans``); ``w:instrText`` (a field's
instruction code) is EXCLUDED entirely; a field's RESULT runs (the ones
between ``w:fldChar fldCharType="separate"`` and ``"end"``, or all of a
``w:fldSimple``'s nested runs) ARE included, each tagged
``field_result=True`` internally for locate.py's WARNING_TOUCHES_FIELD_RESULT
(WP-06; NOT a fifth match rung — see locate.py's own header comment for why
a literal "RUNG_FIELD" as the plan text names it turned out to be unsafe to
build once this file's own field-markdown-rendering fix landed).
``w:sym`` is mapped to a character only when ``w:font`` is Symbol,
Wingdings, or Webdings, via the small table in ``_SYM_TABLE`` below (NOT a
complete 256-glyph mapping for any of the three — see that table's own
docstring); every other case (an unmapped font, or a code point this
module's table does not cover) emits U+FFFD and adds ``"unmapped_symbol"``
to the read warnings.

Only the DrawingML branch of a drawing (``w:drawing``, inside
``mc:Choice``) is walked, never the legacy VML fallback (``w:pict``, inside
``mc:Fallback``) — ``mc:AlternateContent`` gives the SAME visual object in
both branches (confirmed against a real Word-authored text box in this
WP's own fixture generation: both branches carried byte-identical text),
so walking both would double-count every drawing and every text box.
"""

from __future__ import annotations

import dataclasses
import hashlib
import re
import zipfile
from pathlib import Path
from typing import Any

from .errors import ErrorCode, _make_error

# ---------------------------------------------------------------------------
# Namespaces / tag helpers
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DEFAULT_PART = "word/document.xml"


def _ln(elem: Any) -> str:
    """Local (unprefixed) tag name of *elem* — namespace-agnostic dispatch.

    Every tag this module matches (r, t, tab, br, p, tbl, tc, sdt, ...) is
    unambiguous by local name alone within a .docx package; comparing only
    the local name (not the full ``{ns}tag``) keeps the dispatch tables
    below readable and robust to a producer using an unexpected namespace
    prefix (prefixes are cosmetic in XML; only the URI + local name pair is
    semantically meaningful, and collisions across namespaces on these
    specific local names do not occur in practice in a .docx package).
    """
    tag = elem.tag
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _attr(elem: Any, local_name: str) -> str | None:
    """Get an attribute by local name regardless of its namespace prefix
    (``w:val``, ``r:embed``, ``r:id`` all reach here from different
    namespaces depending on the element)."""
    for key, value in elem.attrib.items():
        key_local = key.rsplit("}", 1)[-1] if "}" in key else key
        if key_local == local_name:
            return value
    return None


def identity_para_id(comment_elem: Any, *, ids_root: Any | None = None, ext_root: Any | None = None) -> str | None:
    """Issue #108: the ``w14:paraId`` Word itself keys ``word/commentsIds.xml``
    (``w16cid:commentId/@w16cid:paraId``) and ``word/commentsExtended.xml``
    (``w15:commentEx/@w15:paraId``) on for *comment_elem* (a
    ``word/comments.xml`` ``w:comment`` element) -- the LAST ``w:p`` inside
    it, not the first. A single-paragraph comment has only one candidate, so
    this only matters for a multi-paragraph one; verified against
    ``tests/fixtures/comments/multipara-comment.docx`` (Word 16.112.4):
    a two-paragraph comment's first ``w:p`` paraId appears in NEITHER
    ``commentsIds.xml`` nor ``commentsExtended.xml``, only its last one does.

    *ids_root*/*ext_root* (the parsed roots of those two parts, either or
    both omittable) are used defensively: if the last paragraph's own
    paraId is not actually present in either part while an EARLIER
    paragraph's paraId of this same comment is, that earlier paraId is
    returned instead -- a client that keys differently than Word does
    should still resolve, rather than this helper insisting on Word's own
    convention against evidence in the file itself. With neither root
    given, the last paragraph's paraId is returned unconditionally.

    Returns None if *comment_elem* has no ``w:p`` children with a paraId at
    all (malformed input)."""
    para_ids: list[str] = []
    for p in comment_elem:
        if _ln(p) != "p":
            continue
        pid = _attr(p, "paraId")
        if pid:
            para_ids.append(pid)
    if not para_ids:
        return None

    last = para_ids[-1]
    if ids_root is None and ext_root is None:
        return last

    def _known(pid: str) -> bool:
        if ids_root is not None:
            for child in ids_root:
                if _ln(child) == "commentId" and _attr(child, "paraId") == pid:
                    return True
        if ext_root is not None:
            for child in ext_root:
                if _ln(child) == "commentEx" and _attr(child, "paraId") == pid:
                    return True
        return False

    if _known(last):
        return last
    for pid in para_ids[:-1]:
        if _known(pid):
            return pid
    return last


def _bool_toggle(pr_elem: Any | None, tag: str) -> bool:
    """OOXML boolean-toggle convention: the element's mere presence means
    true UNLESS it carries an explicit ``w:val="false"``/``"0"``."""
    if pr_elem is None:
        return False
    for child in pr_elem:
        if _ln(child) == tag:
            val = _attr(child, "val")
            if val is None:
                return True
            return val.strip().lower() not in ("false", "0", "off")
    return False


def _run_properties(rpr_elem: Any | None) -> dict[str, Any]:
    if rpr_elem is None:
        return {"bold": False, "italic": False, "underline": None, "strike": False, "color": None}
    underline = None
    color = None
    for child in rpr_elem:
        if _ln(child) == "u":
            underline = _attr(child, "val") or "single"
        elif _ln(child) == "color":
            # issue #22: a themeColor-only w:color (val="auto" or absent,
            # theme attributes present) has no explicit RGB to report --
            # left as None rather than "auto", since "auto" is not a
            # settable hex value format_text's own color key accepts.
            val = _attr(child, "val")
            if val and val.lower() != "auto":
                color = val.upper()
    return {
        "bold": _bool_toggle(rpr_elem, "b"),
        "italic": _bool_toggle(rpr_elem, "i"),
        "underline": underline,
        "strike": _bool_toggle(rpr_elem, "strike"),
        "color": color,
    }


# ---------------------------------------------------------------------------
# w:sym mapping — SMALL, deliberately partial (see module docstring).
# ---------------------------------------------------------------------------

# Keyed by (font family lowercased, char code as a 4-hex-digit string,
# e.g. "F0B7" as Word emits it in w:char). Only the handful of glyphs a
# JennyStack-authored or reviewed proposal has actually been seen to use
# (bullets, arrows, a check mark) are covered; anything else — including a
# genuinely unmapped code point in one of these three fonts — falls back to
# U+FFFD + the "unmapped_symbol" warning, per the module docstring.
_SYM_TABLE: dict[tuple[str, str], str] = {
    ("symbol", "F0B7"): "•",  # bullet
    ("symbol", "F0E0"): "←",  # left arrow
    ("symbol", "F0E8"): "→",  # right arrow (Symbol's "Þ")
    ("wingdings", "F0FC"): "✓",  # check mark
    ("wingdings", "F0A7"): "▪",  # black small square
    ("webdings", "F050"): "•",  # webdings ball -> bullet (approximate)
}


def _map_sym(font: str | None, char_code: str | None) -> str | None:
    if not font or not char_code:
        return None
    key = (font.strip().lower(), char_code.strip().upper())
    return _SYM_TABLE.get(key)


# ---------------------------------------------------------------------------
# Internal event model
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class RunEvent:
    text: str
    rpr: dict[str, Any]
    para_ref: str
    run_ref: str
    field_result: bool = False
    # New for issue #28 WP-06 (text_edit.py's locate()/replace_text/
    # format_text). in_revision/comment_ids let locate() report the
    # "crosses_revision"/"crosses_comment_range" warnings without a second
    # walk of the live tree: True/non-empty when this run sits inside a
    # w:ins element, or between a still-open commentRangeStart/End pair,
    # at the moment this event was emitted. r_elem/text_elem/parent_elem
    # are the LIVE Element objects this event came from (the owning w:r,
    # its specific text-bearing child, and that child's own parent
    # container) — None for a synthetic paragraph-break event, or when the
    # caller never needed them (every WP-03 read-path caller passes
    # positional args only and never asks for these, so they stay None
    # there; text_edit.py's run-splitter is the only WP-06 consumer). Not
    # used for equality anywhere in this repo (no test constructs a
    # RunEvent to compare by value) so carrying a live Element reference
    # here is safe.
    in_revision: bool = False
    # New for issue #28 WP-07b-a: the w:author of the (innermost) w:ins this
    # run sits inside, or None when in_revision is False. Lets a caller
    # (text_edit.py's TRACKED_CHANGES_PRESENT guard) exclude the server's
    # OWN previously-authored revisions from a refusal check — see that
    # module's own comment for why ("a second proposed edit deadlocks on
    # the first one's tracked change" otherwise).
    revision_author: str | None = None
    comment_ids: frozenset[str] = dataclasses.field(default_factory=frozenset)
    r_elem: Any = None
    text_elem: Any = None
    parent_elem: Any = None


@dataclasses.dataclass
class TableBoundaryEvent:
    kind: str  # "table_start" | "table_end"
    table_id: int
    # The raw w:tbl Element, set only on "table_start" (None on
    # "table_end") — WP-03b-a's markdown renderer needs direct XML access
    # (w:tcPr/w:gridSpan, w:vMerge, a nested w:tbl) that the flat
    # paragraph/run event stream below does not carry; runs_from_projection
    # never includes this field in its dict output, so the "runs" format's
    # contract is unaffected.
    element: Any = None


@dataclasses.dataclass
class DrawingEvent:
    blip_rid: str | None
    media_part: str | None
    extent_in: list[float] | None
    para_ref: str


@dataclasses.dataclass
class FieldEvent:
    instr: str
    result_text: str
    para_ref: str


@dataclasses.dataclass
class DeletedSpan:
    text: str
    para_ref: str
    author: str | None
    date: str | None


Event = RunEvent | TableBoundaryEvent | DrawingEvent | FieldEvent


# ---------------------------------------------------------------------------
# Ids
# ---------------------------------------------------------------------------


class _Ids:
    def __init__(self) -> None:
        self._para_counter = 0
        self._table_counter = 0

    def next_para_ref(self) -> str:
        ref = f"p{self._para_counter}"
        self._para_counter += 1
        return ref

    def next_table_id(self) -> int:
        self._table_counter += 1
        return self._table_counter


# Non-content wrapper/marker tags a walker explicitly skips (present but
# carry no projectable text of their own — either formatting properties or
# empty markers). Listed explicitly, rather than relying on an unmatched
# tag silently falling through to "recurse into children" in every call
# site, so a genuinely new/unexpected tag is recursed into defensively
# (better to over-discover nested w:r/w:t than silently drop text) while
# these known-empty ones are not descended into pointlessly.
_SKIP_TAGS = frozenset(
    {
        "pPr",
        "rPr",
        "bookmarkStart",
        "bookmarkEnd",
        "commentReference",
        "proofErr",
        "lastRenderedPageBreak",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "permStart",
        "permEnd",
        "tblPr",
        "tblGrid",
        "trPr",
        "tcPr",
        "sdtPr",
        "sdtEndPr",
    }
)


@dataclasses.dataclass
class ParagraphMeta:
    para_ref: str
    style_id: str | None
    outline_lvl: int | None
    container_chain: list[dict[str, int]]
    # w:pPr/w:numPr — None/None for a non-list paragraph. WP-03b-a's list
    # rendering resolves (num_id, ilvl) through _load_numbering_index to
    # decide bullet vs numbered and to track per-level nesting/indent,
    # mirroring GoogleDocs-MCP markdown.py's (listId, nestingLevel).
    num_id: int | None = None
    ilvl: int | None = None


class _PartWalker:
    """Walks one part's block content (a ``w:body``, ``w:hdr``, ``w:ftr``,
    or a ``w:tc``/``w:sdtContent`` fragment) into a flat ``events`` list, in
    document order, plus per-paragraph metadata."""

    def __init__(self) -> None:
        self.ids = _Ids()
        self.events: list[Event] = []
        self.paragraphs: list[ParagraphMeta] = []
        self.deleted_spans: list[DeletedSpan] = []
        self.warnings: list[str] = []
        self._first_para = True
        # New for issue #28 WP-06/WP-07b-a: shared walker state (not
        # threaded as parameters — every recursive _walk_runs call already
        # shares one walker instance) backing RunEvent.in_revision/
        # comment_ids/revision_author. A w:commentRangeStart/End pair can
        # nest (overlapping comments), so this is a running set, not a
        # single id; w:ins can nest too (rare, but the schema allows an
        # w:ins inside another), so this is a STACK of each nested
        # w:ins's own w:author (its innermost entry is the one a run
        # currently inside reports), not a depth counter.
        self._open_comment_ids: set[str] = set()
        self._revision_author_stack: list[str | None] = []

    # -- block level -----------------------------------------------------

    def walk_block_container(self, container: Any, chain: list[dict[str, int]]) -> None:
        for child in container:
            tag = _ln(child)
            if tag == "p":
                self._walk_paragraph(child, chain)
            elif tag == "tbl":
                self._walk_table(child, chain)
            elif tag == "sdt":
                content = self._find_child(child, "sdtContent")
                if content is not None:
                    self.walk_block_container(content, chain)
            elif tag in _SKIP_TAGS:
                continue
            else:
                # Defensive: an unrecognized block-level wrapper might still
                # carry paragraphs (e.g. a future/unknown OOXML extension).
                self.walk_block_container(child, chain)

    def _walk_table(self, tbl: Any, chain: list[dict[str, int]]) -> None:
        table_id = self.ids.next_table_id()
        self.events.append(TableBoundaryEvent("table_start", table_id, element=tbl))
        row_idx = 0
        for row in tbl:
            if _ln(row) != "tr":
                continue
            row_idx += 1
            cell_idx = 0
            for cell in row:
                if _ln(cell) != "tc":
                    continue
                cell_idx += 1
                cell_chain = chain + [{"table_id": table_id, "row": row_idx, "cell": cell_idx}]
                self.walk_block_container(cell, cell_chain)
        self.events.append(TableBoundaryEvent("table_end", table_id))

    def _find_child(self, elem: Any, local_name: str) -> Any | None:
        for child in elem:
            if _ln(child) == local_name:
                return child
        return None

    # -- paragraph / run level --------------------------------------------

    def _walk_paragraph(self, p: Any, chain: list[dict[str, int]]) -> None:
        para_ref = self.ids.next_para_ref()
        ppr = self._find_child(p, "pPr")
        style_id = None
        outline_lvl = None
        num_id = None
        ilvl = None
        if ppr is not None:
            pstyle = self._find_child(ppr, "pStyle")
            if pstyle is not None:
                style_id = _attr(pstyle, "val")
            outline = self._find_child(ppr, "outlineLvl")
            if outline is not None:
                raw = _attr(outline, "val")
                if raw is not None and raw.isdigit():
                    outline_lvl = int(raw)
            num_pr = self._find_child(ppr, "numPr")
            if num_pr is not None:
                num_id_el = self._find_child(num_pr, "numId")
                ilvl_el = self._find_child(num_pr, "ilvl")
                if num_id_el is not None:
                    raw_num_id = _attr(num_id_el, "val")
                    # numId="0" is OOXML's own "no list" convention (used to
                    # cancel inherited list formatting from a style) — not a
                    # real list membership, so it is deliberately excluded
                    # here rather than rendered as list item #0's bullet.
                    if raw_num_id is not None and raw_num_id.isdigit() and int(raw_num_id) != 0:
                        num_id = int(raw_num_id)
                if ilvl_el is not None:
                    raw_ilvl = _attr(ilvl_el, "val")
                    if raw_ilvl is not None and raw_ilvl.isdigit():
                        ilvl = int(raw_ilvl)
        self.paragraphs.append(ParagraphMeta(para_ref, style_id, outline_lvl, chain, num_id=num_id, ilvl=ilvl))

        if not self._first_para:
            self.events.append(RunEvent("\n", {}, para_ref, f"{para_ref}/break"))
        self._first_para = False

        run_counter = [0]
        field_stack: list[dict[str, Any]] = []
        self._walk_runs(p, para_ref, run_counter, field_stack)
        # A field left open at paragraph end (malformed input) still gets
        # flushed as a FieldEvent so no data is silently dropped.
        while field_stack:
            self._flush_field(field_stack.pop(), para_ref)

    def _next_run_ref(self, para_ref: str, counter: list[int]) -> str:
        ref = f"{para_ref}/r{counter[0]}"
        counter[0] += 1
        return ref

    def _flush_field(self, frame: dict[str, Any], para_ref: str) -> None:
        self.events.append(
            FieldEvent(
                instr=frame["instr"].strip(),
                result_text="".join(frame["result_parts"]),
                para_ref=para_ref,
            )
        )

    def _emit_text(
        self,
        text: str,
        rpr: dict[str, Any],
        para_ref: str,
        counter: list[int],
        field_stack: list[dict[str, Any]],
        *,
        r_elem: Any = None,
        text_elem: Any = None,
        parent_elem: Any = None,
    ) -> None:
        if not text:
            return
        in_result = bool(field_stack) and field_stack[-1]["mode"] == "result"
        run_ref = self._next_run_ref(para_ref, counter)
        self.events.append(
            RunEvent(
                text,
                rpr,
                para_ref,
                run_ref,
                field_result=in_result,
                in_revision=bool(self._revision_author_stack),
                revision_author=self._revision_author_stack[-1] if self._revision_author_stack else None,
                comment_ids=frozenset(self._open_comment_ids),
                r_elem=r_elem,
                text_elem=text_elem,
                parent_elem=parent_elem,
            )
        )
        if in_result:
            field_stack[-1]["result_parts"].append(text)

    def _walk_runs(
        self,
        container: Any,
        para_ref: str,
        counter: list[int],
        field_stack: list[dict[str, Any]],
    ) -> None:
        for child in container:
            tag = _ln(child)
            if tag == "r":
                self._walk_run(child, para_ref, counter, field_stack, container)
            elif tag == "ins":
                # New for issue #28 WP-06: track live w:ins nesting (a
                # stack, not a bool/counter, so a nested w:ins reports its
                # own, innermost author) so every RunEvent emitted while
                # inside one is flagged in_revision=True (locate()'s
                # "crosses_revision" warning) and carries revision_author
                # (WP-07b-a: the TRACKED_CHANGES_PRESENT guard's own-
                # author exclusion needs to know WHOSE revision this is,
                # not just that one exists). Content is otherwise walked
                # exactly like hyperlink/smartTag below — an insertion's
                # text is LIVE text (not yet accepted), so it belongs in
                # the projection like any other run.
                self._revision_author_stack.append(_attr(child, "author"))
                try:
                    self._walk_runs(child, para_ref, counter, field_stack)
                finally:
                    self._revision_author_stack.pop()
            elif tag in ("hyperlink", "smartTag"):
                self._walk_runs(child, para_ref, counter, field_stack)
            elif tag == "commentRangeStart":
                # New for issue #28 WP-06: track open comment ids (a set,
                # not a single id — comment ranges can nest) so every
                # RunEvent emitted while one is open carries it in
                # comment_ids (locate()'s "crosses_comment_range" warning).
                # No text of its own; still excluded from the projection.
                cid = _attr(child, "id")
                if cid is not None:
                    self._open_comment_ids.add(cid)
            elif tag == "commentRangeEnd":
                cid = _attr(child, "id")
                if cid is not None:
                    self._open_comment_ids.discard(cid)
            elif tag == "sdt":
                content = self._find_child(child, "sdtContent")
                if content is not None:
                    self._walk_runs(content, para_ref, counter, field_stack)
            elif tag == "del":
                self._record_deleted(child, para_ref)
            elif tag == "fldSimple":
                self._walk_fld_simple(child, para_ref, counter)
            elif tag in _SKIP_TAGS:
                continue
            else:
                self._walk_runs(child, para_ref, counter, field_stack)

    def _record_deleted(self, del_elem: Any, para_ref: str) -> None:
        author = _attr(del_elem, "author")
        date = _attr(del_elem, "date")
        for run in del_elem:
            if _ln(run) != "r":
                continue
            for piece in run:
                if _ln(piece) == "delText":
                    self.deleted_spans.append(DeletedSpan(piece.text or "", para_ref, author, date))

    def _walk_fld_simple(self, fld: Any, para_ref: str, counter: list[int]) -> None:
        instr = _attr(fld, "instr") or ""
        result_parts: list[str] = []
        # A w:fldSimple's own nested runs ARE the field result (there is no
        # separate begin/separate/end sequence for the simple form).
        nested_field_stack: list[dict[str, Any]] = [{"instr": instr, "mode": "result", "result_parts": result_parts}]
        for child in fld:
            if _ln(child) == "r":
                self._walk_run(child, para_ref, counter, nested_field_stack, fld)
            elif _ln(child) not in _SKIP_TAGS:
                self._walk_runs(child, para_ref, counter, nested_field_stack)
        self.events.append(FieldEvent(instr=instr.strip(), result_text="".join(result_parts), para_ref=para_ref))

    def _walk_run(
        self,
        run: Any,
        para_ref: str,
        counter: list[int],
        field_stack: list[dict[str, Any]],
        parent_elem: Any = None,
    ) -> None:
        rpr_elem = self._find_child(run, "rPr")
        rpr = _run_properties(rpr_elem)
        for child in run:
            tag = _ln(child)
            if tag == "rPr":
                continue
            elif tag == "t":
                self._emit_text(
                    child.text or "", rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag == "tab":
                self._emit_text(
                    "\t", rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag in ("br", "cr"):
                self._emit_text(
                    "\n", rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag == "noBreakHyphen":
                self._emit_text(
                    "-", rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag == "softHyphen":
                self._emit_text(
                    "­", rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag == "sym":
                font = _attr(child, "font")
                code = _attr(child, "char")
                mapped = _map_sym(font, code)
                if mapped is None:
                    mapped = "�"
                    self.warnings.append("unmapped_symbol")
                self._emit_text(
                    mapped, rpr, para_ref, counter, field_stack,
                    r_elem=run, text_elem=child, parent_elem=parent_elem,
                )
            elif tag == "fldChar":
                fld_type = _attr(child, "fldCharType")
                if fld_type == "begin":
                    field_stack.append({"instr": "", "mode": "instr", "result_parts": []})
                elif fld_type == "separate" and field_stack:
                    field_stack[-1]["mode"] = "result"
                elif fld_type == "end" and field_stack:
                    self._flush_field(field_stack.pop(), para_ref)
            elif tag == "instrText":
                if field_stack:
                    field_stack[-1]["instr"] += child.text or ""
                # else: an instrText outside any begin/separate — malformed
                # input; excluded from the projection either way.
            elif tag == "delText":
                # A delText directly under a bare w:r (not wrapped in
                # w:del) is malformed, but exclude it from the projection
                # defensively rather than emit deleted text as if live.
                self.deleted_spans.append(DeletedSpan(child.text or "", para_ref, None, None))
            elif tag == "drawing":
                self._emit_drawing(child, para_ref)
            elif tag == "AlternateContent":
                # Real Word wraps a text box's/picture's w:drawing here:
                # <w:r><w:rPr/><mc:AlternateContent><mc:Choice Requires="wps">
                # <w:drawing>...</w:drawing></mc:Choice><mc:Fallback><w:pict>
                # ...</w:pict></mc:Fallback></mc:AlternateContent></w:r>
                # (confirmed against this WP's own Word-authored
                # textbox.docx fixture). Choice and Fallback encode the
                # SAME visual object for different Word versions — only
                # Choice (the modern w:drawing branch) is walked; Fallback
                # (the legacy w:pict/VML branch) is skipped entirely, or
                # every drawing/text box would be double-counted. See the
                # module docstring.
                choice = self._find_child(child, "Choice")
                if choice is not None:
                    for grandchild in choice:
                        if _ln(grandchild) == "drawing":
                            self._emit_drawing(grandchild, para_ref)
            elif tag in _SKIP_TAGS:
                continue
            # Anything else inside a run (proofErr, lastRenderedPageBreak,
            # etc.) carries no text and is silently skipped.

    def _emit_drawing(self, drawing: Any, para_ref: str) -> None:
        blip_rid: str | None = None
        extent_in: list[float] | None = None
        for extent in drawing.iter():
            if _ln(extent) == "extent":
                cx = extent.get("cx")
                cy = extent.get("cy")
                if cx is not None and cy is not None:
                    extent_in = [round(int(cx) / 914400, 4), round(int(cy) / 914400, 4)]
                break
        for blip in drawing.iter():
            if _ln(blip) == "blip":
                blip_rid = _attr(blip, "embed")
                break
        self.events.append(
            DrawingEvent(blip_rid=blip_rid, media_part=None, extent_in=extent_in, para_ref=para_ref)
        )


# ---------------------------------------------------------------------------
# Projection result
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class Projection:
    part: str
    text: str
    # Sorted, non-overlapping spans: (start, end, para_ref, run_ref).
    offset_map: list[tuple[int, int, str, str]]
    events: list[Event]
    paragraphs: list[ParagraphMeta]
    deleted_spans: list[DeletedSpan]
    warnings: list[str]

    def locate_offset(self, offset: int) -> tuple[str, str, int] | None:
        """(para_ref, run_ref, run_offset) for *offset* in ``text``, or
        None if *offset* falls in a paragraph break / outside any run."""
        for start, end, para_ref, run_ref in self.offset_map:
            if start <= offset < end:
                return para_ref, run_ref, offset - start
        return None


def _rels_path_for(part_name: str) -> str:
    if "/" in part_name:
        directory, base = part_name.rsplit("/", 1)
        return f"{directory}/_rels/{base}.rels"
    return f"_rels/{part_name}.rels"


def _load_rels(zf: zipfile.ZipFile, part_name: str) -> dict[str, str]:
    """rId -> Target for *part_name*'s own relationships part. Empty dict
    if the .rels part does not exist (a part with no relationships)."""
    rels_path = _rels_path_for(part_name)
    if rels_path not in zf.namelist():
        return {}
    data = zf.read(rels_path)
    import xml.etree.ElementTree as ET

    root = ET.fromstring(data)
    out: dict[str, str] = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            out[rid] = target
    return out


def _resolve_media_part(part_name: str, rels: dict[str, str], rid: str | None) -> str | None:
    if not rid or rid not in rels:
        return None
    target = rels[rid]
    if target.startswith("/"):
        return target.lstrip("/")
    # Relative to the part's own directory (word/ for document.xml/headerN.xml).
    directory = part_name.rsplit("/", 1)[0] if "/" in part_name else ""
    return f"{directory}/{target}" if directory else target


def load_rels_root(zf: zipfile.ZipFile, part_name: str) -> Any | None:
    """Parsed root element of *part_name*'s own .rels part, or None if it
    does not exist. Complements _load_rels (rId -> Target map) with the
    full element tree — mutations.py (WP-04) needs the tree itself so it
    can append a new <Relationship> (e.g. for a markdown link's
    hyperlink) rather than just read existing targets."""
    rels_path = _rels_path_for(part_name)
    if rels_path not in zf.namelist():
        return None
    import xml.etree.ElementTree as ET

    return ET.fromstring(zf.read(rels_path))


def read_part_xml(zf: zipfile.ZipFile, part_name: str) -> Any | None:
    """Parsed root element of *part_name*, or None if the part is absent
    from the package (a header/footer/footnotes/endnotes part is optional;
    a caller asking for one that does not exist gets PART_NOT_FOUND, not a
    silent empty projection — see project_part)."""
    if part_name not in zf.namelist():
        return None
    import xml.etree.ElementTree as ET

    return ET.fromstring(zf.read(part_name))


def project_part(docx_path: Path, part_name: str = DEFAULT_PART) -> Projection:
    """Walk *part_name* (default ``word/document.xml``'s body) into a
    Projection. Raises VerifyError(PART_NOT_FOUND) if the part is absent.
    """
    with zipfile.ZipFile(docx_path) as zf:
        root = read_part_xml(zf, part_name)
        if root is None:
            raise _make_error(
                ErrorCode.PART_NOT_FOUND,
                f"Part not found in package: {part_name!r}",
                {"part": part_name, "available_parts": sorted(zf.namelist())},
            )
        rels = _load_rels(zf, part_name)

        walker = _PartWalker()
        root_tag = _ln(root)
        if root_tag == "document":
            body = walker._find_child(root, "body")
            if body is not None:
                walker.walk_block_container(body, [])
        elif root_tag in ("footnotes", "endnotes"):
            # Each w:footnote/w:endnote is its own mini-story; concatenated
            # here into one flat projection for the part (WP-03 scope —
            # per-story addressing is not required by any acceptance test
            # in this WP). Separator/continuation-separator stories (ids
            # -1 and 0) carry no real content and are skipped.
            child_tag = "footnote" if root_tag == "footnotes" else "endnote"
            for story in root:
                if _ln(story) != child_tag:
                    continue
                story_id = _attr(story, "id")
                if story_id in ("-1", "0"):
                    continue
                walker.walk_block_container(story, [])
        else:
            # w:hdr / w:ftr (headers/footers): block content is direct.
            walker.walk_block_container(root, [])

        # Resolve media parts on drawing events now that rels are loaded.
        for event in walker.events:
            if isinstance(event, DrawingEvent) and event.blip_rid:
                event.media_part = _resolve_media_part(part_name, rels, event.blip_rid)

    return _finalize_projection(walker, part_name)


def _finalize_projection(walker: _PartWalker, part_name: str) -> Projection:
    """Build the flat text + offset map (from RunEvents only) and wrap
    *walker*'s accumulated state into a Projection. Shared by project_part
    (a package part), project_textbox_scope (a w:txbxContent sub-scope),
    and project_document_root (issue #28 WP-06: an already-parsed, LIVE
    document root, reused so text_edit.py's replace_text/format_text can
    splice the very same Element objects rather than a throwaway
    re-parse) — the three differ only in WHICH element was walked and
    whether it came from a fresh parse or a live tree already in memory,
    not in how the walk's output is assembled.
    """
    text_parts: list[str] = []
    offset_map: list[tuple[int, int, str, str]] = []
    cursor = 0
    for event in walker.events:
        if isinstance(event, RunEvent):
            length = len(event.text)
            if length:
                offset_map.append((cursor, cursor + length, event.para_ref, event.run_ref))
                cursor += length
            text_parts.append(event.text)

    return Projection(
        part=part_name,
        text="".join(text_parts),
        offset_map=offset_map,
        events=walker.events,
        paragraphs=walker.paragraphs,
        deleted_spans=walker.deleted_spans,
        warnings=walker.warnings,
    )


def project_document_root(document_root: Any, part_name: str = DEFAULT_PART) -> Projection:
    """Like project_part, but walks an already-parsed, LIVE document root
    (e.g. from mutations._load_document) instead of reading *part_name*
    fresh from a zip on disk.

    New for issue #28 WP-06: text_edit.py's replace_text/format_text need
    the actual live Element objects (the owning w:r, its text-bearing
    child, that child's own parent) to splice, not a throwaway parse — every
    RunEvent's r_elem/text_elem/parent_elem here are elements from
    *document_root* itself, so mutating them (or inserting siblings next to
    them) edits the tree the caller is about to serialize. Scoped to
    word/document.xml's w:body (this WP's tools only ever target the body);
    unlike project_part, this never resolves DrawingEvent.media_part (no
    rels are loaded here — no WP-06 caller needs it).
    """
    walker = _PartWalker()
    body = walker._find_child(document_root, "body")
    if body is not None:
        walker.walk_block_container(body, [])
    return _finalize_projection(walker, part_name)


# ---------------------------------------------------------------------------
# Text boxes: sub-scopes, never merged into body offsets.
# ---------------------------------------------------------------------------


def _read_part_root_or_raise(docx_path: Path, part_name: str) -> Any:
    with zipfile.ZipFile(docx_path) as zf:
        root = read_part_xml(zf, part_name)
        if root is None:
            raise _make_error(
                ErrorCode.PART_NOT_FOUND,
                f"Part not found in package: {part_name!r}",
                {"part": part_name},
            )
        return root


def _iter_textbox_content_elements(root: Any) -> list[tuple[str, Any]]:
    """[(section_key, w:txbxContent element), ...] in document order, from
    *root*'s DrawingML branch only (``w:drawing``/``mc:Choice`` — see
    module docstring on why the legacy VML ``mc:Fallback`` is never
    walked). ``section_key`` is ``textbox-<n>``, 1-based in document
    order — the shared enumeration behind both iter_textbox_scopes (list)
    and project_textbox_scope (read one).
    """
    found: list[tuple[str, Any]] = []
    n = 0
    for drawing in root.iter():
        if _ln(drawing) != "drawing":
            continue
        for txbx_content in drawing.iter():
            if _ln(txbx_content) != "txbxContent":
                continue
            n += 1
            found.append((f"textbox-{n}", txbx_content))
    return found


def iter_textbox_scopes(docx_path: Path, part_name: str = DEFAULT_PART) -> list[dict[str, Any]]:
    """One entry per ``w:txbxContent`` found in *part_name*'s DrawingML
    branch, each projected as an independent sub-scope with its own flat
    text and ``section_key`` (``textbox-<n>``), never merged into the host
    part's own offsets. Reachable through a tool call via
    read_document(..., section_key=<one of these>) and listed alongside
    heading sections by find_sections (see both for how a caller
    discovers and then addresses one).
    """
    root = _read_part_root_or_raise(docx_path, part_name)
    scopes: list[dict[str, Any]] = []
    for section_key, txbx_content in _iter_textbox_content_elements(root):
        walker = _PartWalker()
        walker.walk_block_container(txbx_content, [])
        text_parts = [e.text for e in walker.events if isinstance(e, RunEvent)]
        scopes.append(
            {
                "section_key": section_key,
                "text": "".join(text_parts),
                "paragraph_count": len(walker.paragraphs),
            }
        )
    return scopes


def project_textbox_scope(docx_path: Path, part_name: str, section_key: str) -> Projection | None:
    """The full Projection (text, offset_map, runs, warnings — same shape
    project_part returns for a whole part) for one text box's own
    ``w:txbxContent``, addressed by the ``section_key`` iter_textbox_scopes
    and find_sections both report (``textbox-<n>``). None if *section_key*
    does not name a text box found in *part_name* — the caller (server.py)
    turns that into INVALID_INPUT with the actual available keys attached,
    not a silent empty read.
    """
    root = _read_part_root_or_raise(docx_path, part_name)
    for key, txbx_content in _iter_textbox_content_elements(root):
        if key != section_key:
            continue
        walker = _PartWalker()
        walker.walk_block_container(txbx_content, [])
        return _finalize_projection(walker, f"{part_name}#{section_key}")
    return None


# ---------------------------------------------------------------------------
# Revision token
# ---------------------------------------------------------------------------

# Fixed order (issue #28 plan WP-03's revision-token spec). The rels and
# [Content_Types].xml parts are included because adding the FIRST comment
# to a document creates these new parts/relationships without necessarily
# touching word/document.xml's own text — a document-only hash would miss
# exactly that case.
_COMMENT_PARTS_ORDER: tuple[str, ...] = (
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
    "word/_rels/document.xml.rels",
    "[Content_Types].xml",
)


def _sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def compute_package_fingerprint(source: Path | bytes) -> str:
    """Full sha256 over EVERY zip entry of a .docx (issue #27).

    Distinct from ``compute_revision``: that token is 2x32 bits over
    word/document.xml plus the comment parts only, which is right for the
    caller-facing ``revision_before`` handshake but is not a file
    fingerprint -- a styles/numbering/header/footer/media change leaves it
    untouched. This one is what the write ledger and the write-window
    rechecks compare, so "did anyone else change this package" has a
    defined scope.

    Entries are hashed sorted by name, each framed as (name length, name,
    data length, decompressed data), so compression level, entry order,
    and zip timestamps -- which a co-authoring editor's re-save changes
    freely -- do not matter, but any content change does. A path is read
    with ONE ``read_bytes()`` and the zip parsed from that buffer, so the
    hash never mixes two versions of a file that changed mid-read.
    """
    import io

    data = source if isinstance(source, bytes) else Path(source).read_bytes()
    hasher = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in sorted(zf.namelist()):
            encoded = name.encode("utf-8")
            payload = zf.read(name)
            hasher.update(len(encoded).to_bytes(8, "big"))
            hasher.update(encoded)
            hasher.update(len(payload).to_bytes(8, "big"))
            hasher.update(payload)
    return hasher.hexdigest()


def compute_revision(docx_path: Path) -> dict[str, Any]:
    """Revision token for *docx_path*: ``{"token": "<doc8>:<cmt8>",
    "detail": {...}}``.

    ``document`` hash = sha256(word/document.xml). ``comments`` hash =
    sha256 over a LENGTH-PREFIXED concatenation of the seven parts in
    ``_COMMENT_PARTS_ORDER`` (each part missing from the package
    contributes a zero-length frame, i.e. is "hashed as the empty string")
    — length-prefixed rather than naive concatenation so that content
    shifting across a part boundary (e.g. comments.xml growing by exactly
    as much as commentsIds.xml shrinks) cannot produce a hash collision.

    Equality is defined on ``token`` (the two hashes) ONLY. ``detail``
    additionally carries ``size``/``mtime_ns`` of *docx_path* itself —
    informational, for a caller that wants to skip recomputing the hashes
    when neither has changed; this function itself always recomputes both
    hashes from the actual bytes, so a `touch` (mtime changes, content and
    size do not) still yields the SAME token.
    """
    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        doc_bytes = zf.read(DEFAULT_PART) if DEFAULT_PART in names else b""
        doc_hash = _sha256_hex(doc_bytes)

        hasher = hashlib.sha256()
        for part in _COMMENT_PARTS_ORDER:
            data = zf.read(part) if part in names else b""
            hasher.update(len(data).to_bytes(8, "big"))
            hasher.update(data)
        cmt_hash = hasher.hexdigest()

    st = docx_path.stat()
    doc8 = doc_hash[:8]
    cmt8 = cmt_hash[:8]
    return {
        "token": f"{doc8}:{cmt8}",
        "detail": {
            "document_sha256": doc_hash,
            "comments_sha256": cmt_hash,
            "size": st.st_size,
            "mtime_ns": st.st_mtime_ns,
        },
    }


# ---------------------------------------------------------------------------
# list_parts
# ---------------------------------------------------------------------------

_HEADER_FOOTER_RE = re.compile(r"^word/(header|footer)(\d+)\.xml$")


def list_parts_impl(docx_path: Path) -> list[dict[str, Any]]:
    """Enumerate the parts a read_document(part=...) call may target:
    the document body, every header/footer, and footnotes/endnotes if
    present. header_footer_type ("default"/"first"/"even") is resolved via
    word/document.xml's own w:headerReference/w:footerReference elements
    where possible; None when it cannot be determined (a header/footer
    part present but not referenced from any sectPr — unusual but not
    invalid OOXML).
    """
    with zipfile.ZipFile(docx_path) as zf:
        names = zf.namelist()
        parts: list[dict[str, Any]] = []
        if DEFAULT_PART in names:
            parts.append({"part": DEFAULT_PART, "kind": "document", "header_footer_type": None})

        ref_type_by_target: dict[str, str] = {}
        if DEFAULT_PART in names:
            root = read_part_xml(zf, DEFAULT_PART)
            rels = _load_rels(zf, DEFAULT_PART)
            if root is not None:
                for ref in root.iter():
                    ln = _ln(ref)
                    if ln in ("headerReference", "footerReference"):
                        rid = _attr(ref, "id")
                        ref_type = _attr(ref, "type")
                        target = rels.get(rid) if rid else None
                        if target and ref_type:
                            target_part = f"word/{target}" if not target.startswith("word/") else target
                            ref_type_by_target[target_part] = ref_type

        for name in sorted(names):
            m = _HEADER_FOOTER_RE.match(name)
            if m:
                kind = "header" if m.group(1) == "header" else "footer"
                parts.append(
                    {
                        "part": name,
                        "kind": kind,
                        "header_footer_type": ref_type_by_target.get(name),
                    }
                )

        if "word/footnotes.xml" in names:
            parts.append({"part": "word/footnotes.xml", "kind": "footnotes", "header_footer_type": None})
        if "word/endnotes.xml" in names:
            parts.append({"part": "word/endnotes.xml", "kind": "endnotes", "header_footer_type": None})

    return parts


# ---------------------------------------------------------------------------
# list_styles
# ---------------------------------------------------------------------------


def list_styles_impl(docx_path: Path) -> list[dict[str, Any]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = read_part_xml(zf, "word/styles.xml")
        if root is None:
            return []
        styles: list[dict[str, Any]] = []
        for style in root:
            if _ln(style) != "style":
                continue
            style_id = _attr(style, "styleId")
            style_type = _attr(style, "type")
            name = None
            outline_lvl = None
            for child in style:
                if _ln(child) == "name":
                    name = _attr(child, "val")
                elif _ln(child) == "pPr":
                    for grandchild in child:
                        if _ln(grandchild) == "outlineLvl":
                            raw = _attr(grandchild, "val")
                            if raw is not None and raw.isdigit():
                                outline_lvl = int(raw)
            styles.append(
                {
                    "style_id": style_id,
                    "name": name,
                    "type": style_type,
                    "outline_lvl": outline_lvl,
                }
            )
        return styles


# ---------------------------------------------------------------------------
# numbering.xml — w:numId -> {ilvl: numFmt}, WP-03b-a's list rendering.
# ---------------------------------------------------------------------------


def numbering_index_from_elements(
    abstract_num_elements: list[Any], num_elements: list[Any]
) -> dict[int, dict[int, str]]:
    """The same ``w:num``/``w:abstractNum`` -> {numId: {ilvl: numFmt}}
    join load_numbering_index does, but over already-in-memory Elements
    rather than a part read from a docx on disk — the on-disk numbering.xml
    doesn't exist yet for a list a write is ABOUT to add (mutations.py's
    ``intended_preview`` renders new_elements — freshly built by
    markdown_to_ooxml.render_blocks — before anything is written; those
    paragraphs' w:numId values only resolve against
    ``ctx.new_abstract_nums``/``ctx.new_nums``, never an existing numId, so
    this is the correct — and only available — index for that preview).
    """
    abstract_fmt: dict[int, dict[int, str]] = {}
    for el in abstract_num_elements:
        raw_aid = _attr(el, "abstractNumId")
        if raw_aid is None or not raw_aid.isdigit():
            continue
        aid = int(raw_aid)
        fmts: dict[int, str] = {}
        for lvl in el:
            if _ln(lvl) != "lvl":
                continue
            raw_ilvl = _attr(lvl, "ilvl")
            if raw_ilvl is None or not raw_ilvl.isdigit():
                continue
            for child in lvl:
                if _ln(child) == "numFmt":
                    val = _attr(child, "val")
                    if val is not None:
                        fmts[int(raw_ilvl)] = val
        abstract_fmt[aid] = fmts

    num_to_abstract: dict[int, int] = {}
    overrides: dict[int, dict[int, str]] = {}
    for el in num_elements:
        raw_nid = _attr(el, "numId")
        if raw_nid is None or not raw_nid.isdigit():
            continue
        nid = int(raw_nid)
        for child in el:
            if _ln(child) == "abstractNumId":
                val = _attr(child, "val")
                if val is not None and val.isdigit():
                    num_to_abstract[nid] = int(val)
            elif _ln(child) == "lvlOverride":
                raw_ilvl = _attr(child, "ilvl")
                if raw_ilvl is None or not raw_ilvl.isdigit():
                    continue
                for sub in child:
                    if _ln(sub) != "lvl":
                        continue
                    for grandchild in sub:
                        if _ln(grandchild) == "numFmt":
                            val = _attr(grandchild, "val")
                            if val is not None:
                                overrides.setdefault(nid, {})[int(raw_ilvl)] = val

    index: dict[int, dict[int, str]] = {}
    for nid, aid in num_to_abstract.items():
        fmts = dict(abstract_fmt.get(aid, {}))
        fmts.update(overrides.get(nid, {}))
        index[nid] = fmts
    return index


def load_numbering_index(docx_path: Path) -> dict[int, dict[int, str]]:
    """``word/numbering.xml``'s ``w:num`` (numId -> abstractNumId) joined
    with ``w:abstractNum`` (abstractNumId -> {ilvl: numFmt}), a per-numId
    ``w:num/w:lvlOverride/w:lvl/w:numFmt`` taking precedence over the
    abstractNum's own value at that level when present. Returns {} when the
    part is absent (no lists anywhere in the document).

    Used by the markdown renderer to tell a bullet list from a numbered one
    (a "bullet" numFmt renders "- "; anything else this table recognizes as
    numeric renders sequential "1. " — see _ORDERED_NUM_FMTS) — never
    inferred from the paragraph's own lvlText/glyph, which OOXML does not
    guarantee carries that distinction directly on the paragraph itself.
    """
    with zipfile.ZipFile(docx_path) as zf:
        if "word/numbering.xml" not in zf.namelist():
            return {}
        root = read_part_xml(zf, "word/numbering.xml")
    if root is None:
        return {}
    abstract_num_elements = [el for el in root if _ln(el) == "abstractNum"]
    num_elements = [el for el in root if _ln(el) == "num"]
    return numbering_index_from_elements(abstract_num_elements, num_elements)


# numFmt values rendered as a sequential "1. " marker — anything else
# (most commonly "bullet", but also "none" or an unrecognized future
# value) renders "- " instead. An explicit allowlist, never a denylist:
# GoogleDocs-MCP markdown.py's _ORDERED_GLYPH_TYPES makes the identical
# choice for the same reason (never GUESS a value is numeric).
_ORDERED_NUM_FMTS = frozenset(
    {
        "decimal",
        "decimalZero",
        "decimalEnclosedCircle",
        "decimalEnclosedFullstop",
        "decimalEnclosedParen",
        "lowerRoman",
        "upperRoman",
        "lowerLetter",
        "upperLetter",
        "ordinal",
    }
)


# Built-in heading style ids Word emits even when styles.xml's own <w:pPr>
# carries no explicit <w:outlineLvl> (common for a document whose styles
# part was never customized) — "Heading1".."Heading9" -> outline level 0-8.
_HEADING_STYLE_ID_RE = re.compile(r"^Heading(\d)$", re.IGNORECASE)


def _resolve_outline_level(style_id: str | None, direct_outline_lvl: int | None, styles_by_id: dict[str, dict]) -> int | None:
    if direct_outline_lvl is not None:
        return direct_outline_lvl
    if style_id and style_id in styles_by_id:
        style_outline = styles_by_id[style_id].get("outline_lvl")
        if style_outline is not None:
            return style_outline
    if style_id:
        m = _HEADING_STYLE_ID_RE.match(style_id)
        if m:
            return int(m.group(1)) - 1
    return None


# ---------------------------------------------------------------------------
# find_sections
# ---------------------------------------------------------------------------


def _slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.strip().lower()).strip("-")
    return slug or "section"


def find_sections_impl(docx_path: Path, part_name: str = DEFAULT_PART) -> list[dict[str, Any]]:
    """Heading ranges PLUS text-box sub-scopes, both keyed by
    ``section_key`` — this is the one discovery call a caller uses to find
    every ``section_key`` read_document(section_key=...) can then address,
    of either kind (each entry's ``"kind"`` is ``"heading"`` or
    ``"textbox"``).

    Heading ranges: a heading paragraph's w:pStyle mapped through
    list_styles to an outline level (falling back to the paragraph's own
    direct w:outlineLvl). A section runs from one heading paragraph
    (inclusive) up to, but not including, the NEXT heading paragraph at
    any level, or the end of the document for the last heading —
    WP-03 scope does not nest sections by level; a future WP may.

    section_key = slug(heading text) + "-" + a 1-based ordinal disambiguating
    duplicate headings (two paragraphs both literally "Introduction" get
    "introduction-1" and "introduction-2").

    Text-box sub-scopes: one entry per iter_textbox_scopes' own
    ``textbox-<n>`` keys (never merged into the heading list's ordinals —
    a text box's key never collides with a heading slug). heading_text,
    outline_level, start_para_ref, end_para_ref are None on a textbox
    entry; paragraph_count is still meaningful (the text box's own
    paragraph count).
    """
    styles = list_styles_impl(docx_path)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}

    projection = project_part(docx_path, part_name)
    heading_paras: list[tuple[ParagraphMeta, int, str]] = []
    para_text_by_ref: dict[str, list[str]] = {}
    for event in projection.events:
        if isinstance(event, RunEvent):
            para_text_by_ref.setdefault(event.para_ref, []).append(event.text)

    for meta in projection.paragraphs:
        level = _resolve_outline_level(meta.style_id, meta.outline_lvl, styles_by_id)
        if level is None or meta.container_chain:
            # Headings inside a table cell are out of scope for WP-03's
            # section model (a table-cell "heading" style is common for
            # emphasis, not a real document section boundary).
            continue
        heading_text = "".join(para_text_by_ref.get(meta.para_ref, [])).strip()
        heading_paras.append((meta, level, heading_text))

    slug_counts: dict[str, int] = {}
    sections: list[dict[str, Any]] = []
    all_para_refs = [m.para_ref for m in projection.paragraphs]

    for idx, (meta, level, heading_text) in enumerate(heading_paras):
        base_slug = _slugify(heading_text)
        slug_counts[base_slug] = slug_counts.get(base_slug, 0) + 1
        section_key = f"{base_slug}-{slug_counts[base_slug]}"

        start_idx = all_para_refs.index(meta.para_ref)
        if idx + 1 < len(heading_paras):
            next_meta = heading_paras[idx + 1][0]
            end_idx = all_para_refs.index(next_meta.para_ref)
        else:
            end_idx = len(all_para_refs)

        sections.append(
            {
                "section_key": section_key,
                "kind": "heading",
                "heading_text": heading_text,
                "outline_level": level,
                "start_para_ref": meta.para_ref,
                "end_para_ref": all_para_refs[end_idx - 1] if end_idx > start_idx else meta.para_ref,
                "paragraph_count": max(end_idx - start_idx, 1),
            }
        )

    for textbox_scope in iter_textbox_scopes(docx_path, part_name):
        sections.append(
            {
                "section_key": textbox_scope["section_key"],
                "kind": "textbox",
                "heading_text": None,
                "outline_level": None,
                "start_para_ref": None,
                "end_para_ref": None,
                "paragraph_count": textbox_scope["paragraph_count"],
            }
        )

    return sections


# ---------------------------------------------------------------------------
# list_page_sections
# ---------------------------------------------------------------------------

# OOXML page geometry (w:pgSz, w:pgMar, w:cols/w:col) is expressed in
# twentieths of a point (dxa) — 1440 dxa = 1 inch — NOT in EMU. EMU (the
# 914400-per-inch unit the issue #28 plan's WP-03 text names) is what
# DrawingML extents (a w:drawing's wp:extent, handled in _emit_drawing
# above) use instead. Converting page geometry with the plan's literal
# "EMU / 914400" phrasing would silently produce numbers 635x too small;
# see this module's PR notes for the correction against the plan text.
_DXA_PER_INCH = 1440


def _dxa_to_in(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return round(int(value) / _DXA_PER_INCH, 4)
    except ValueError:
        return None


def list_page_sections_impl(docx_path: Path, part_name: str = DEFAULT_PART) -> list[dict[str, Any]]:
    with zipfile.ZipFile(docx_path) as zf:
        root = read_part_xml(zf, part_name)
        if root is None:
            raise _make_error(
                ErrorCode.PART_NOT_FOUND,
                f"Part not found in package: {part_name!r}",
                {"part": part_name},
            )
        sections: list[dict[str, Any]] = []
        for sect_pr in root.iter():
            if _ln(sect_pr) != "sectPr":
                continue
            # Two passes, not one: real Word emits pgSz/pgMar before cols,
            # but nothing in the schema guarantees that order, and cols'
            # own (no explicit w:col children) branch needs both pgSz and
            # pgMar already resolved to compute an even column width.
            pg_sz = next((c for c in sect_pr if _ln(c) == "pgSz"), None)
            pg_mar = next((c for c in sect_pr if _ln(c) == "pgMar"), None)
            cols_elem = next((c for c in sect_pr if _ln(c) == "cols"), None)

            # w:pgSz/w:pgMar attributes (w:w, w:h, w:top, ...) are in the
            # w: namespace, NOT bare — unlike DrawingML's cx/cy (see
            # _emit_drawing). _attr() (namespace-agnostic by local name)
            # is required here; a bare .get("w") always returns None.
            page_w = _dxa_to_in(_attr(pg_sz, "w")) if pg_sz is not None else None
            page_h = _dxa_to_in(_attr(pg_sz, "h")) if pg_sz is not None else None
            orientation = _attr(pg_sz, "orient") if pg_sz is not None else None
            margin_top = _dxa_to_in(_attr(pg_mar, "top")) if pg_mar is not None else None
            margin_bottom = _dxa_to_in(_attr(pg_mar, "bottom")) if pg_mar is not None else None
            margin_left = _dxa_to_in(_attr(pg_mar, "left")) if pg_mar is not None else None
            margin_right = _dxa_to_in(_attr(pg_mar, "right")) if pg_mar is not None else None

            cols_widths: list[float] = []
            explicit_cols = [c for c in cols_elem if _ln(c) == "col"] if cols_elem is not None else []
            if explicit_cols:
                for col in explicit_cols:
                    w = _dxa_to_in(_attr(col, "w"))
                    if w is not None:
                        cols_widths.append(w)
            else:
                # No explicit w:col children — the common case. w:cols'
                # own w:num attribute DEFAULTS TO 1 when absent; it is not
                # "no column information". Confirmed against a real
                # Word-authored fixture (sections.docx): Word's default
                # single-column section emits a bare <w:cols w:space="720"/>
                # with no w:num at all, not <w:cols w:num="1" .../>. Both
                # forms, and cols_elem being entirely absent from sectPr,
                # are treated identically here: one column spanning the
                # full usable width (page width minus both side margins).
                num_raw = _attr(cols_elem, "num") if cols_elem is not None else None
                space = _dxa_to_in(_attr(cols_elem, "space")) if cols_elem is not None else None
                n = int(num_raw) if num_raw and num_raw.isdigit() else 1
                if page_w is not None and margin_left is not None and margin_right is not None and n > 0:
                    usable = page_w - margin_left - margin_right
                    gap_total = (space or 0.0) * (n - 1)
                    cols_widths = [round((usable - gap_total) / n, 4)] * n

            sections.append(
                {
                    "page_width_in": page_w,
                    "page_height_in": page_h,
                    "orientation": orientation or "portrait",
                    "margin_top_in": margin_top,
                    "margin_bottom_in": margin_bottom,
                    "margin_left_in": margin_left,
                    "margin_right_in": margin_right,
                    "column_widths_in": cols_widths,
                }
            )
        return sections


# ---------------------------------------------------------------------------
# read_document: text / runs / markdown
# ---------------------------------------------------------------------------


def _run_event_to_dict(event: RunEvent) -> dict[str, Any]:
    return {"text": event.text, "rPr": event.rpr, "para_ref": event.para_ref}


def runs_from_projection(proj: Projection) -> list[dict[str, Any]]:
    """The ``read_document(format="runs")`` list, built from an
    already-resolved Projection — shared by read_document_runs (a whole
    package part) and server.py's textbox-scoped read
    (project_textbox_scope), so the two never carry independently
    maintained copies of this record-shape logic. Run records
    ``{text, rPr, para_ref}`` interleaved, in document order, with
    structural records — ``{"type": "table_start"|"table_end",
    "table_id"}``, ``{"type": "drawing", blip_rid, media_part, extent_in,
    para_ref}`` (emitted for every ``w:drawing``/``a:blip``), and
    ``{"type": "field", instr, result_text}``.
    """
    out: list[dict[str, Any]] = []
    for event in proj.events:
        if isinstance(event, RunEvent):
            if event.run_ref.endswith("/break"):
                continue  # internal paragraph-break marker, not a real run
            out.append(_run_event_to_dict(event))
        elif isinstance(event, TableBoundaryEvent):
            out.append({"type": event.kind, "table_id": event.table_id})
        elif isinstance(event, DrawingEvent):
            out.append(
                {
                    "type": "drawing",
                    "blip_rid": event.blip_rid,
                    "media_part": event.media_part,
                    "extent_in": event.extent_in,
                    "para_ref": event.para_ref,
                }
            )
        elif isinstance(event, FieldEvent):
            out.append({"type": "field", "instr": event.instr, "result_text": event.result_text})
    return out


def read_document_runs(docx_path: Path, part_name: str = DEFAULT_PART) -> list[dict[str, Any]]:
    """``read_document(format="runs")`` for a whole package part — see
    runs_from_projection for the record shapes."""
    return runs_from_projection(project_part(docx_path, part_name))


def read_document_text(docx_path: Path, part_name: str = DEFAULT_PART) -> str:
    return project_part(docx_path, part_name).text


@dataclasses.dataclass
class _ListFrame:
    """Sequential-numbering state for one active nesting level while
    converting a run of list-item paragraphs back to markdown — ported
    from GoogleDocs-MCP markdown.py's _ListFrame (num_id/ilvl standing in
    for that module's list_id/nestingLevel; see that module's
    _advance_list_item docstring for the indent-width rationale)."""

    num_id: int
    ilvl: int
    ordered: bool
    counter: int
    indent: int  # leading-space count for THIS level's own marker


def _advance_list_item(
    stack: list[_ListFrame], num_id: int, ilvl: int, ordered: bool
) -> tuple[str, int]:
    """Advance (or start) the counter for this (num_id, ilvl) and return
    (leading-space indent, 1-based counter) for the current item —
    identical algorithm to GoogleDocs-MCP markdown.py's
    _Converter._advance_list_item; see that module for why indent is the
    cumulative width of every ancestor's own marker rather than a fixed
    per-level indent."""
    while stack and stack[-1].ilvl >= ilvl:
        if stack[-1].ilvl == ilvl and stack[-1].num_id == num_id:
            break
        stack.pop()

    if stack and stack[-1].ilvl == ilvl:
        frame = stack[-1]
        frame.counter += 1
        frame.ordered = ordered
    else:
        parent_indent = 0
        if stack:
            parent = stack[-1]
            marker = f"{parent.counter}. " if parent.ordered else "- "
            parent_indent = parent.indent + len(marker)
        frame = _ListFrame(num_id=num_id, ilvl=ilvl, ordered=ordered, counter=1, indent=parent_indent)
        stack.append(frame)

    return " " * frame.indent, frame.counter


def _pipe_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def _table_cell_markdown(
    tc: Any, styles_by_id: dict[str, dict], table_id: int, lossy: list[dict[str, Any]]
) -> str:
    """One cell's markdown text — GoogleDocs-MCP markdown.py's ``_table``
    cell handling (:350-359): the cell's own paragraphs rendered through
    the SAME renderer as ordinary body content (so a bold run, a field, an
    image inside a cell all still work), multiple paragraphs joined with a
    single space (never their own newlines — a pipe-table row is one
    line), any nested w:tbl flattened to inline text and recorded as a
    "nested_table" lossy element. A literal "|" is escaped so it cannot be
    mistaken for a column boundary — GoogleDocs-MCP gets this for free by
    running every text run through _escape_markdown; this backend's
    run-level renderer does not escape at all (a pre-existing WP-03 gap
    out of this WP's scope), so it is done narrowly at this cell boundary
    instead, where an unescaped "|" would otherwise corrupt the row.
    """
    paragraphs_only: list[Any] = []
    nested_texts: list[str] = []
    for child in tc:
        tag = _ln(child)
        if tag == "p":
            paragraphs_only.append(child)
        elif tag == "tbl":
            lossy.append({"kind": "nested_table", "table_id": table_id})
            nested_texts.append(_flatten_table_to_text(child, styles_by_id, table_id, lossy))
        elif tag == "sdt":
            content = None
            for grandchild in child:
                if _ln(grandchild) == "sdtContent":
                    content = grandchild
                    break
            if content is not None:
                for c in content:
                    if _ln(c) == "p":
                        paragraphs_only.append(c)
                    elif _ln(c) == "tbl":
                        lossy.append({"kind": "nested_table", "table_id": table_id})
                        nested_texts.append(_flatten_table_to_text(c, styles_by_id, table_id, lossy))
        # tcPr and anything else: not block content, skipped.

    text = markdown_from_elements(paragraphs_only, styles_by_id) if paragraphs_only else ""
    text = text.replace("\n", " ")
    if nested_texts:
        extra = " ".join(t for t in nested_texts if t)
        text = f"{text} {extra}".strip() if text else extra
    text = " ".join(text.split())
    return text.replace("|", "\\|")


def _flatten_table_to_text(
    tbl: Any, styles_by_id: dict[str, dict], table_id: int, lossy: list[dict[str, Any]]
) -> str:
    """A nested w:tbl has no pipe-table representation at its host cell's
    position (GFM does not nest tables inside a cell) — every cell's text
    is flattened into one inline, space-joined run so the outer cell is
    not silently emptied. The caller already recorded the "nested_table"
    lossy element; this only renders the fallback text."""
    parts: list[str] = []
    for tr in tbl:
        if _ln(tr) != "tr":
            continue
        for tc in tr:
            if _ln(tc) != "tc":
                continue
            parts.append(_table_cell_markdown(tc, styles_by_id, table_id, lossy))
    return " ".join(p for p in parts if p)


def _table_to_markdown(
    tbl: Any, table_id: int, styles_by_id: dict[str, dict]
) -> tuple[str, list[dict[str, Any]]]:
    """The ``read_document(format="markdown")`` rendering of one w:tbl — a
    GFM pipe table matching GoogleDocs-MCP markdown.py's ``_table``/
    ``_pipe_row`` conventions exactly (see that module's :341-381): first
    row as header, a separator row of "---", column count normalised to
    the widest row (short rows padded with empty cells).

    Two OOXML-specific structural cases that pipe-table markdown has no
    room for, both recorded as a lossy_elements entry rather than
    reproduced (or silently dropped):
      - w:gridSpan (a horizontally merged cell): its text is emitted once,
        followed by (span - 1) empty cells to keep the grid rectangular.
      - w:vMerge (a vertically merged cell): the "restart" cell (the top
        of the merge) renders its own text normally; every continuation
        cell (no w:val, or w:val != "restart") renders as an empty cell —
        it carries no content of its own in the source OOXML either.
    """
    lossy: list[dict[str, Any]] = []
    rows: list[list[str]] = []
    for tr in tbl:
        if _ln(tr) != "tr":
            continue
        cells: list[str] = []
        for tc in tr:
            if _ln(tc) != "tc":
                continue
            tc_pr = None
            for child in tc:
                if _ln(child) == "tcPr":
                    tc_pr = child
                    break
            grid_span = 1
            v_merge_present = False
            v_merge_continue = False
            if tc_pr is not None:
                for child in tc_pr:
                    if _ln(child) == "gridSpan":
                        val = _attr(child, "val")
                        if val is not None and val.isdigit():
                            grid_span = int(val)
                    elif _ln(child) == "vMerge":
                        v_merge_present = True
                        val = _attr(child, "val")
                        v_merge_continue = (val is None) or (val.strip().lower() != "restart")
            if v_merge_present or grid_span > 1:
                lossy.append({"kind": "table_merge", "table_id": table_id})
            text = "" if v_merge_continue else _table_cell_markdown(tc, styles_by_id, table_id, lossy)
            cells.append(text)
            cells.extend([""] * (grid_span - 1))
        rows.append(cells)

    if not rows:
        return "", lossy

    col_count = max((len(r) for r in rows), default=1) or 1
    normalised = [r + [""] * (col_count - len(r)) for r in rows]
    header, body = normalised[0], normalised[1:]
    lines = [_pipe_row(header), _pipe_row(["---"] * col_count)]
    lines.extend(_pipe_row(row) for row in body)
    return "\n".join(lines), lossy


def _events_to_markdown(
    events: list[Event],
    paragraphs: list[ParagraphMeta],
    styles_by_id: dict[str, dict],
    numbering_index: dict[int, dict[int, str]] | None = None,
) -> tuple[str, list[dict[str, Any]]]:
    """Shared rendering loop behind markdown_from_projection and
    markdown_from_elements below — factored out so both a whole-part
    projection and an ad hoc slice of a few w:body children render through
    the exact same rules (see markdown_from_projection's docstring).
    Returns (markdown_text, lossy_elements) — the latter populated only by
    table merges/nesting (_table_to_markdown); everything else this
    renderer supports (headings, bold/italic, lists, fields, images) has a
    lossless markdown token, per GoogleDocs-MCP markdown.py's own
    lossy_elements contract (image/chip/footnote there; table_merge/
    nested_table here — this backend has no chip/footnote concept).

    Block separation matches GoogleDocs-MCP markdown.py's convert(): two
    consecutive list items join with a single newline (a "tight" list, so
    CommonMark keeps them one list); any other pair of blocks (a heading,
    an ordinary paragraph, a table) joins with a blank line, so re-parsing
    the output never merges two distinct blocks into one.
    """
    numbering_index = numbering_index or {}
    meta_by_ref = {m.para_ref: m for m in paragraphs}
    blocks: list[tuple[bool, str]] = []  # (is_list_item, block_text)
    lossy_elements: list[dict[str, Any]] = []
    list_stack: list[_ListFrame] = []

    current_para_ref: str | None = None
    current_parts: list[str] = []

    def flush() -> None:
        nonlocal current_para_ref, current_parts
        if current_para_ref is None:
            return
        text = "".join(current_parts)
        meta = meta_by_ref.get(current_para_ref)
        level = _resolve_outline_level(meta.style_id, meta.outline_lvl, styles_by_id) if meta else None
        if level is not None and text.strip():
            blocks.append((False, f"{'#' * (level + 1)} {text.strip()}"))
        elif meta is not None and meta.num_id is not None:
            fmt = numbering_index.get(meta.num_id, {}).get(meta.ilvl or 0)
            ordered = fmt in _ORDERED_NUM_FMTS
            indent, counter = _advance_list_item(list_stack, meta.num_id, meta.ilvl or 0, ordered)
            marker = f"{counter}. " if ordered else "- "
            blocks.append((True, f"{indent}{marker}{text.strip()}"))
        else:
            blocks.append((False, text))
        current_para_ref = None
        current_parts = []

    skip_table_depth = 0
    for event in events:
        if isinstance(event, TableBoundaryEvent):
            if event.kind == "table_start":
                if skip_table_depth == 0:
                    # The OUTERMOST table only: render the whole w:tbl
                    # structurally from its raw element (row/cell/gridSpan/
                    # vMerge/nested-tbl), then skip every flat event inside
                    # it below — those are the same cells' paragraphs
                    # walked a second time for the "runs"/"text" formats'
                    # sake, and re-emitting them here would duplicate every
                    # cell's text as bogus extra body lines (the WP-03
                    # bug this WP replaces).
                    flush()
                    table_md, table_lossy = _table_to_markdown(event.element, event.table_id, styles_by_id)
                    if table_md:
                        blocks.append((False, table_md))
                    lossy_elements.extend(table_lossy)
                skip_table_depth += 1
            else:  # table_end
                skip_table_depth = max(0, skip_table_depth - 1)
            continue
        if skip_table_depth > 0:
            continue
        if isinstance(event, RunEvent):
            if event.run_ref.endswith("/break"):
                flush()
                continue
            if current_para_ref is None:
                current_para_ref = event.para_ref
            piece = event.text
            if event.rpr.get("bold") and event.rpr.get("italic"):
                piece = f"***{piece}***" if piece.strip() else piece
            elif event.rpr.get("bold"):
                piece = f"**{piece}**" if piece.strip() else piece
            elif event.rpr.get("italic"):
                piece = f"*{piece}*" if piece.strip() else piece
            current_parts.append(piece)
        elif isinstance(event, DrawingEvent):
            if current_para_ref is None:
                current_para_ref = event.para_ref
            current_parts.append(f"[image:{event.blip_rid}]" if event.blip_rid else "[image:unknown]")
        elif isinstance(event, FieldEvent):
            # A field's RESULT runs already flowed into current_parts as
            # ordinary RunEvents (field_result=True) before this FieldEvent
            # fires — at fldChar "end", or at the close of a w:fldSimple —
            # so appending event.result_text here would double it. Only
            # when there is NO result at all does this event contribute
            # anything: a lowercase "[field:<instr>]" placeholder, so a
            # field with an instr but no resolved result isn't silently
            # invisible in the rendered markdown.
            if not event.result_text and event.instr:
                if current_para_ref is None:
                    current_para_ref = event.para_ref
                current_parts.append(f"[field:{event.instr}]")

    flush()

    if not blocks:
        return "", lossy_elements
    out = [blocks[0][1]]
    for i in range(1, len(blocks)):
        prev_is_list = blocks[i - 1][0]
        cur_is_list, chunk = blocks[i]
        out.append("\n" if (prev_is_list and cur_is_list) else "\n\n")
        out.append(chunk)
    return "".join(out), lossy_elements


def markdown_from_projection(docx_path: Path, proj: Projection) -> tuple[str, list[str], list[dict[str, Any]]]:
    """The ``read_document(format="markdown")`` rendering, built from an
    already-resolved Projection — shared by read_document_markdown (a
    whole package part) and server.py's textbox-scoped read
    (project_textbox_scope). *docx_path* is still needed for
    list_styles_impl/load_numbering_index (a heading's outline level and a
    list's bullet-vs-numbered form resolve through the package's
    styles.xml/numbering.xml, not through anything the Projection itself
    carries).

    Paragraph text, headings (from find_sections' same outline-level
    resolution) as ``#`` runs, **bold**/*italic* run markers, bulleted/
    numbered lists, GFM pipe tables, ``[image:rId]``/``[field:instr]``
    placeholders for a drawing / a field with no result — see
    _events_to_markdown and _table_to_markdown for the exact conventions
    (matched to GoogleDocs-MCP markdown.py's; issue #28 WP-03b-a). Returns
    (markdown_text, warnings, lossy_elements) — lossy_elements is the
    table_merge/nested_table record described on _events_to_markdown,
    empty when the content has neither.
    """
    styles = list_styles_impl(docx_path)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    numbering_index = load_numbering_index(docx_path)
    markdown, lossy_elements = _events_to_markdown(proj.events, proj.paragraphs, styles_by_id, numbering_index)
    return markdown, proj.warnings, lossy_elements


def read_document_markdown(docx_path: Path, part_name: str = DEFAULT_PART) -> tuple[str, list[str], list[dict[str, Any]]]:
    """``read_document(format="markdown")`` for a whole package part — see
    markdown_from_projection for the rendering rules."""
    return markdown_from_projection(docx_path, project_part(docx_path, part_name))


def markdown_from_elements(
    elements: list[Any],
    styles_by_id: dict[str, dict],
    numbering_index: dict[int, dict[int, str]] | None = None,
) -> str:
    """Render an arbitrary list of block-level elements (e.g. a slice of a
    w:body's direct children, or one table cell's own paragraphs) through
    the exact same markdown rules as read_document_markdown, without
    requiring them to already be a part read from a docx on disk.

    Used by mutations.py (WP-04) to compute the ``before``/``after``
    evidence text for replace_range_markdown and replace_body_markdown (a
    plain Python list of Element objects is a valid "container" for
    _PartWalker.walk_block_container, so the identical walker + rendering
    pipeline applies to a range that was never itself written to a temp
    part) and by _table_cell_markdown above for one cell's paragraphs.
    Discards the lossy_elements half of _events_to_markdown's return — the
    callers here use this for plain evidence/cell text, not the top-level
    read_document response.
    """
    walker = _PartWalker()
    walker.walk_block_container(elements, [])
    markdown, _lossy = _events_to_markdown(walker.events, walker.paragraphs, styles_by_id, numbering_index)
    return markdown
