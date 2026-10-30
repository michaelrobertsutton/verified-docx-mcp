"""Insert styled paragraphs before/after one anchored body paragraph (issue #77).

Live mode sends the pane's ``paragraph_insert`` op (``live/paragraphs_live.py``);
file mode edits the OOXML body directly with the same guard -> write -> verify ->
audit pipeline ``apply_style`` and ``append_markdown`` use. Request validation is
shared, so both modes refuse the same malformed requests before anything is sent
or read.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection, text_edit, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .live import write_mode as live_write_mode
from .locate import locate

_HEX_COLOR_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")
_FORBIDDEN_TEXT_CHARS = ("\r", "\n", "\v", "\t")
_MAX_ANCHOR_LEN = 255  # Word's body.search limit
_MAX_PARAGRAPHS = 50
_SPEC_KEYS = frozenset({"text", "style", "color"})
# Issue #107: structural anchors that address a gap without searching for text.
POSITIONAL_ANCHOR_KEYS = ("after_table_id", "empty_paragraph_after_table_id")
# ONE policy for "empty paragraph": filling a paragraph holding any of these would lose or
# move something. File mode applies it to the parsed body; it is sent to the pane, which
# applies the same list to the OOXML it reads, so the two modes refuse the same paragraphs.
EMPTY_PARAGRAPH_FORBIDDEN_TAGS = (
    "drawing", "pict", "object", "sectPr", "tab", "br", "cr", "noBreakHyphen", "softHyphen", "sym",
    "bookmarkStart", "bookmarkEnd",
    "fldChar", "fldSimple", "instrText", "hyperlink", "footnoteReference", "endnoteReference",
    "commentRangeStart", "commentRangeEnd", "commentReference", "ins", "del", "moveFrom", "moveTo",
)


def _invalid(message: str, **diagnostics: Any) -> Exception:
    return _make_error(ErrorCode.INVALID_INPUT, message, diagnostics)


def positional_anchor(anchor: Any) -> tuple[str, int] | None:
    """``(kind, table_id)`` for a structural anchor, ``None`` for a text anchor;
    ``INVALID_INPUT`` for a malformed object."""
    if not isinstance(anchor, dict):
        return None
    if len(anchor) != 1 or next(iter(anchor)) not in POSITIONAL_ANCHOR_KEYS:
        raise _invalid(f"a structural anchor holds exactly one of {list(POSITIONAL_ANCHOR_KEYS)}")
    (kind, table_id), = anchor.items()
    if type(table_id) is not int or table_id < 1:
        raise _invalid(f"{kind} must be a positive integer table_id")
    return kind, table_id


def validate_insert_request(
    anchor: str | dict[str, Any], position: str, paragraphs: list[dict[str, Any]], expected_matches: int
) -> list[dict[str, Any]]:
    """Return normalized paragraph specs (``text``, ``style``, ``color``: uppercase
    6-hex without ``#``, or ``None``); raise ``INVALID_INPUT`` otherwise."""
    if positional_anchor(anchor) is not None:
        if position != "after":
            raise _invalid("a structural table anchor inserts position='after' only")
    elif not isinstance(anchor, str) or not anchor:
        raise _invalid("anchor must be a non-empty string (or a structural table anchor object)")
    elif any(ch in anchor for ch in _FORBIDDEN_TEXT_CHARS):
        raise _invalid("anchor must be one line of text (no newline, vertical tab or tab characters)")
    elif len(anchor) > _MAX_ANCHOR_LEN:
        raise _invalid(f"anchor is {len(anchor)} characters; Word searches at most {_MAX_ANCHOR_LEN}")
    if position not in ("after", "before"):
        raise _invalid(f"position must be 'after' or 'before', got {position!r}")
    if isinstance(expected_matches, bool) or expected_matches != 1:
        raise _invalid("insert_paragraphs inserts at exactly one anchor; expected_matches must be 1")
    if not isinstance(paragraphs, list) or not 1 <= len(paragraphs) <= _MAX_PARAGRAPHS:
        raise _invalid(f"paragraphs must be a list of 1 to {_MAX_PARAGRAPHS} paragraph objects")
    specs: list[dict[str, Any]] = []
    for i, raw in enumerate(paragraphs, start=1):
        if not isinstance(raw, dict):
            raise _invalid(f"paragraph {i} must be an object with text, style and color keys")
        unknown = sorted(set(raw) - _SPEC_KEYS)
        if unknown:
            raise _invalid(f"paragraph {i} has unknown key(s) {unknown}; allowed: text, style, color")
        text = raw.get("text")
        if not isinstance(text, str):
            raise _invalid(f"paragraph {i} needs a string 'text'")
        if any(ch in text for ch in _FORBIDDEN_TEXT_CHARS):
            raise _invalid(f"paragraph {i} text must be one line (no newline, vertical tab or tab characters); "
                           "send one paragraph object per paragraph")
        style = raw.get("style")
        if style is not None and (not isinstance(style, str) or not style.strip()):
            raise _invalid(f"paragraph {i} style must be a non-empty paragraph style name")
        color = raw.get("color")
        if color is not None:
            match = _HEX_COLOR_RE.match(color) if isinstance(color, str) else None
            if match is None:
                raise _invalid(f"paragraph {i} color must be 6 hex digits like '7030A0', got {color!r}")
            color = match.group(1).upper()
        specs.append({"text": text, "style": style, "color": color})
    return specs


def execute_insert_paragraphs(
    path: str,
    anchor: str | dict[str, Any],
    paragraphs: list[dict[str, Any]],
    expected_matches: int,
    *,
    position: str = "after",
    revision_before: str | None = None,
    track_changes: bool = False,
    write_mode: str = "auto",
    allow_concurrent_editor: bool = False,
) -> dict[str, Any]:
    specs = validate_insert_request(anchor, position, paragraphs, expected_matches)
    mode = live_write_mode.resolve_write_mode(path, write_mode)
    if mode == "live":
        from .live import paragraphs_live

        return paragraphs_live.insert_paragraphs(
            path, anchor, position, specs, expected_matches,
            revision_before=revision_before, track_changes=track_changes)
    return _insert_paragraphs_file(
        path, anchor, specs, expected_matches, position=position, revision_before=revision_before,
        track_changes=track_changes, allow_concurrent_editor=allow_concurrent_editor)


def _resolve_style_ids(resolved: Path, specs: list[dict[str, Any]]) -> list[str | None]:
    records = projection.list_styles_impl(resolved)
    # styles.xml stores Word's internal names ("heading 2"), while Word's UI and
    # callers say "Heading 2": match names case-insensitively, ids exactly.
    by_name = {s["name"].lower(): s for s in records if s.get("name")}
    by_id = {s["style_id"]: s for s in records if s.get("style_id")}
    out: list[str | None] = []
    for spec in specs:
        wanted = spec["style"]
        if wanted is None:
            out.append(None)
            continue
        record = by_id.get(wanted) or by_name.get(wanted.lower())
        if record is None:
            raise _make_error(
                ErrorCode.STYLE_NOT_FOUND,
                f"paragraph style {wanted!r} is not a style in this document's styles.xml.",
                {"style": wanted, "available_paragraph_styles": sorted(
                    s["name"] for s in records if s.get("type") == "paragraph" and s.get("name"))})
        if record.get("type") != "paragraph":
            raise _make_error(
                ErrorCode.UNSUPPORTED_STYLE_TYPE,
                f"style {wanted!r} is a {record.get('type')!r} style; insert_paragraphs takes paragraph styles only.",
                {"style": wanted, "type": record.get("type")})
        out.append(record["style_id"])
    return out


def _build_paragraph(spec: dict[str, Any], style_id: str | None) -> Any:
    p = ET.Element(text_edit._w("p"))
    if style_id is not None:
        ppr = ET.SubElement(p, text_edit._w("pPr"))
        ET.SubElement(ppr, text_edit._w("pStyle"), {text_edit._w("val"): style_id})
    if spec["text"]:
        rpr = None
        if spec["color"]:
            rpr = ET.Element(text_edit._w("rPr"))
            ET.SubElement(rpr, text_edit._w("color"), {text_edit._w("val"): spec["color"]})
        p.append(text_edit._build_run(spec["text"], rpr))
    return p


def _paragraph_style_id(p_elem: Any) -> str | None:
    for child in p_elem:
        if projection._ln(child) == "pPr":
            for grandchild in child:
                if projection._ln(grandchild) == "pStyle":
                    return projection._attr(grandchild, "val")
    return None


def empty_paragraph_problem(p_elem: Any) -> str | None:
    """Why *p_elem* cannot be filled as an "empty" paragraph, or ``None`` when it is empty
    under the shared policy (no text, none of ``EMPTY_PARAGRAPH_FORBIDDEN_TAGS``)."""
    for node in p_elem.iter():
        name = projection._ln(node)
        if name == "t" and node.text:
            return "it holds text"
        if name in EMPTY_PARAGRAPH_FORBIDDEN_TAGS:
            return f"it holds a {name} element"
    return None


def _fill_paragraph(target: Any, built: Any, style_id: str | None) -> None:
    """Put *built*'s content into the existing *target* paragraph, keeping its own
    properties; only a given style replaces the paragraph style."""
    if style_id is not None:
        ppr = next((c for c in target if projection._ln(c) == "pPr"), None)
        if ppr is None:
            ppr = ET.Element(text_edit._w("pPr"))
            target.insert(0, ppr)
        for old in [c for c in ppr if projection._ln(c) == "pStyle"]:
            ppr.remove(old)
        ppr.insert(0, ET.Element(text_edit._w("pStyle"), {text_edit._w("val"): style_id}))
    for child in list(built):
        if projection._ln(child) != "pPr":
            target.append(child)


def _insert_paragraphs_file(
    path: str,
    anchor: str | dict[str, Any],
    specs: list[dict[str, Any]],
    expected_matches: int,
    *,
    position: str,
    revision_before: str | None,
    track_changes: bool,
    allow_concurrent_editor: bool,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(
        resolved, revision_before, allow_concurrent_editor=allow_concurrent_editor, live_capable=True)
    style_ids = _resolve_style_ids(resolved, specs)

    document_root, raw_xml = mutations._load_document(resolved)
    parent_map = text_edit._build_parent_map(document_root)
    body = mutations._find_body(document_root)
    structural = positional_anchor(anchor)
    fill = structural is not None and structural[0] == "empty_paragraph_after_table_id"
    locate_result = None
    if structural is not None:
        from . import tables

        table = tables._find_table_element(document_root, structural[1])
        if parent_map.get(id(table)) is not body:
            raise _make_error(
                ErrorCode.STRUCTURAL_BOUNDARY, "the anchor table is not a top-level table", {"anchor": anchor})
        p_elem = table
        start = end = 0
        before_excerpt = ""
    else:
        proj = projection.project_document_root(document_root)
        locate_result = locate(anchor, proj, expected_matches)
        start, end = locate_result.spans[0]
        atoms = text_edit._atoms_for_span(proj, start, end)
        p_elem = text_edit._enclosing_paragraph(parent_map, atoms[0][2].r_elem) if atoms else None
        if p_elem is None or parent_map.get(id(p_elem)) is not body:
            raise _make_error(
                ErrorCode.STRUCTURAL_BOUNDARY,
                "the anchor is not in a top-level body paragraph (a table cell or other container); "
                "anchor on text in a body paragraph.",
                {"anchor": anchor})
        before_excerpt = text_edit._excerpt(proj.text, start, end)

    target = None
    if fill:
        children = list(body)
        following = children[children.index(p_elem) + 1] if children.index(p_elem) + 1 < len(children) else None
        if following is None or following.tag != text_edit._w("p"):
            raise _make_error(
                ErrorCode.STRUCTURAL_BOUNDARY,
                "there is no paragraph immediately after the table to fill", {"anchor": anchor})
        problem = empty_paragraph_problem(following)
        if problem:
            raise _make_error(
                ErrorCode.STRUCTURAL_BOUNDARY,
                f"the paragraph after the table is not empty: {problem}", {"anchor": anchor})
        mutations._check_hazards_or_raise(mutations._scan_range_hazards([following]), False, resolve_author_name())
        target = following
        # Filling keeps the paragraph's own style unless the caller names one.
        if style_ids[0] is None:
            style_ids = [_paragraph_style_id(target), *style_ids[1:]]

    new_ps = [_build_paragraph(spec, style_id) for spec, style_id in zip(specs, style_ids, strict=True)]
    track = None
    if track_changes:
        track = tracked_changes.TrackContext(document_root, author=resolve_author_name())
        tracked_changes.wrap_all_runs(
            new_ps,
            lambda r: tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date))
    at = list(body).index(p_elem) + (0 if position == "before" and structural is None else 1)
    body_count_before = len(list(body))
    next_sibling = list(body)[at + (1 if fill else 0)] if at + (1 if fill else 0) < body_count_before else None
    next_text = "".join(next_sibling.itertext()) if next_sibling is not None else None
    if fill:
        _fill_paragraph(target, new_ps[0], style_ids[0])
        remaining = new_ps[1:]
        at += 1
    else:
        remaining = new_ps
    for offset, new_p in enumerate(remaining):
        body.insert(at + offset, new_p)
    first_written = at - 1 if fill else at

    intended_text = projection.project_document_root(document_root).text

    def _post_verify(written_path: Path) -> None:
        diff = mutations._diff_modulo_whitespace(intended_text, projection.read_document_text(written_path))
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")
        new_root, _ = mutations._load_document(written_path)
        children = list(mutations._find_body(new_root))
        written = children[first_written : first_written + len(specs)]
        if len(written) != len(specs):
            raise ValueError("re-read body does not hold the inserted paragraphs where expected")
        for i, (p, style_id) in enumerate(zip(written, style_ids, strict=True), start=1):
            if _paragraph_style_id(p) != style_id:
                raise ValueError(f"re-read inserted paragraph {i} does not carry w:pStyle {style_id!r}")
        if structural is not None:
            if len(children) != body_count_before + len(specs) - (1 if fill else 0):
                raise ValueError("re-read body paragraph count differs from the intended count")
            after_index = first_written + len(specs)
            following_text = "".join(children[after_index].itertext()) if after_index < len(children) else None
            if following_text != next_text:
                raise ValueError("the content after the inserted paragraphs changed")
            if children[first_written - 1].tag != text_edit._w("tbl"):
                raise ValueError("the inserted paragraphs are not directly after the anchor table")

    conflict_sweep = text_edit._serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)
    evidence = text_edit._evidence(
        applied=True,
        match_count=1,
        rung=locate_result.rung if locate_result is not None else 1,
        before=before_excerpt,
        after="\n".join(spec["text"] for spec in specs),
        revision_before=pre_revision["token"],
        revision_after=conflict_sweep["revision_after"],  # the staged token, never a re-read of the file
        audit_logged=False,
        runs_before=[],
        runs_after=[],
        warnings=locate_result.warnings if locate_result is not None else [],
        track=track,
        conflict_sweep=conflict_sweep,
    )
    evidence["position"] = position
    evidence["inserted"] = [
        {"text": spec["text"], "style_id": style_id, "color": spec["color"]}
        for spec, style_id in zip(specs, style_ids, strict=True)
    ]
    logged, _ = audit.append_audit(path=str(resolved), tool="insert_paragraphs", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
