"""diff_table_vs_markdown (issue #109): report the cells that differ between a
Word table and a markdown pipe table, instead of making the caller eyeball two
renderings of a large table.

READ-ONLY. Nothing here writes the docx or the markdown, and nothing is
audited. The one rule that shapes the module: a comparison that could not be
made (unsupported table structure, an unmatched table, an ambiguous selector)
is never reported as "equal". ``identical`` is True only when every table
involved was fully compared and nothing differed.

Both sides are reduced to *visible cell text* structurally -- the docx side
from the cell's OOXML, the markdown side from markdown-it's token stream --
and then compared after ``locate.normalize_for_compare`` and whitespace
collapsing. The docx side deliberately does not reparse
``projection._table_cell_markdown`` output: that renderer does not escape
markdown metacharacters, so a literal ``*x*`` typed in Word would be lost.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from . import locate, paths, projection, tables
from .errors import ErrorCode, _make_error
from .live import reads_live
from .markdown_to_ooxml import _MD
from .projection import DEFAULT_PART, W_NS, TableBoundaryEvent

# Cell content the comparison cannot see faithfully: the table is reported
# under `unsupported` rather than compared.
_LOSSY_CELL_TAGS = frozenset({"fldChar", "fldSimple", "drawing", "pict", "txbxContent"})
_OFF_VALUES = frozenset({"0", "false", "off"})
_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)
_HEADING_MARK_RE = re.compile(r"^\s*#+\s+")


def _norm(text: str) -> str:
    return " ".join(locate.normalize_for_compare(text).split())


def _invalid(message: str, **diagnostics: Any) -> Exception:
    return _make_error(ErrorCode.INVALID_INPUT, message, diagnostics)


# ---------------------------------------------------------------------------
# Markdown side
# ---------------------------------------------------------------------------


class _MdTable:
    def __init__(self, index: int, heading: int | None) -> None:
        self.index = index  # 0-based position among every table in the markdown
        self.heading = heading  # index into the heading list, or None before any heading
        self.rows: list[list[str]] = []  # row 0 is the header row

    @property
    def header(self) -> list[str]:
        return [_norm(c) for c in self.rows[0]] if self.rows else []


def _inline_text(children: list[Any], strip_emphasis: bool, warnings: list[str]) -> str:
    out: list[str] = []
    for ch in children or []:
        kind = ch.type
        if kind in ("text", "text_special"):
            out.append(ch.content if strip_emphasis else ch.content.replace("*", "\\*"))
        elif kind == "code_inline":
            out.append(ch.content)
        elif kind in ("softbreak", "hardbreak"):
            out.append(" ")
        elif kind == "strong_open" or kind == "strong_close":
            if not strip_emphasis:
                out.append("**")
        elif kind == "em_open" or kind == "em_close":
            if not strip_emphasis:
                out.append("*")
        elif kind == "image":
            out.append(ch.content)
        elif kind == "html_inline":
            if _BR_RE.fullmatch(ch.content.strip()):
                out.append(" ")
            else:
                out.append(ch.content)
                warning = "markdown cell contains inline HTML that was compared as literal text"
                if warning not in warnings:
                    warnings.append(warning)
        # link_open / link_close: the link text between them is walked as normal text.
    return "".join(out)


def parse_markdown_tables(
    markdown: str, strip_emphasis: bool, warnings: list[str]
) -> tuple[list[_MdTable], list[str]]:
    """Every pipe table in *markdown* plus the normalized heading texts, from
    the token stream. A table belongs to the nearest preceding heading (of any
    level); a ``#`` line inside a fenced code block is a ``fence`` token, never
    a heading."""
    tokens = _MD.parse(markdown)
    headings: list[str] = []
    found: list[_MdTable] = []
    current: _MdTable | None = None
    for pos, tok in enumerate(tokens):
        if tok.type == "heading_open":
            inline = tokens[pos + 1] if pos + 1 < len(tokens) else None
            text = _inline_text(inline.children, True, []) if inline is not None else ""
            headings.append(_norm(text))
        elif tok.type == "table_open":
            current = _MdTable(len(found), len(headings) - 1 if headings else None)
            found.append(current)
        elif tok.type == "table_close":
            current = None
        elif current is not None:
            if tok.type == "tr_open":
                current.rows.append([])
            elif tok.type == "inline" and current.rows:
                current.rows[-1].append(_inline_text(tok.children, strip_emphasis, warnings))
    return found, headings


def _split_header_row(text: str) -> list[str]:
    """Cells of a pipe-row string (`| A | B |`), honoring escaped pipes."""
    body = text.strip()
    if body.startswith("|"):
        body = body[1:]
    if body.endswith("|") and not body.endswith("\\|"):
        body = body[:-1]
    cells = re.split(r"(?<!\\)\|", body)
    return [_norm(c.strip().replace("\\|", "|")) for c in cells]


def _heading_matches(headings: list[str], wanted: str) -> list[int]:
    target = _norm(_HEADING_MARK_RE.sub("", wanted))
    return [i for i, h in enumerate(headings) if h == target]


# ---------------------------------------------------------------------------
# Docx side
# ---------------------------------------------------------------------------


def _run_on(rpr: Any, tag: str) -> bool:
    if rpr is None:
        return False
    for child in rpr:
        if projection._ln(child) == tag:
            return child.get(f"{{{W_NS}}}val", "1").lower() not in _OFF_VALUES
    return False


def _paragraph_text(p: Any, strip_emphasis: bool) -> str:
    segments: list[list[Any]] = []  # [text, bold, italic], adjacent equal formats merged
    for r in p.iter(f"{{{W_NS}}}r"):
        text = ""
        for child in r:
            name = projection._ln(child)
            if name == "t":
                text += child.text or ""
            elif name == "tab":
                text += "\t"
            elif name in ("br", "cr"):
                text += " "
            elif name == "noBreakHyphen":
                text += "-"
        if not text:
            continue
        rpr = next((c for c in r if projection._ln(c) == "rPr"), None)
        bold, italic = (False, False) if strip_emphasis else (_run_on(rpr, "b"), _run_on(rpr, "i"))
        if segments and segments[-1][1] == bold and segments[-1][2] == italic:
            segments[-1][0] += text
        else:
            segments.append([text, bold, italic])
    out: list[str] = []
    for text, bold, italic in segments:
        if strip_emphasis:
            out.append(text)
            continue
        text = text.replace("*", "\\*")
        core = text.strip()
        if not core or not (bold or italic):
            out.append(text)
            continue
        marker = ("**" if bold else "") + ("*" if italic else "")
        lead = text[: len(text) - len(text.lstrip())]
        trail = text[len(text.rstrip()) :]
        out.append(f"{lead}{marker}{core}{marker[::-1]}{trail}")
    return "".join(out)


def _cell_text(tc: Any, strip_emphasis: bool) -> str:
    paragraphs = [_paragraph_text(p, strip_emphasis) for p in tc.iter(f"{{{W_NS}}}p")]
    return " ".join(paragraphs)


def _extract_docx_table(tbl: Any, strip_emphasis: bool) -> tuple[list[list[str]] | None, str | None]:
    """(rows of raw cell text, None), or (None, reason) when the table cannot
    be compared faithfully."""
    _rows, _cols, has_merged, has_nested = tables._table_stats(tbl)
    if has_merged:
        return None, "table has merged cells"
    if has_nested:
        return None, "table contains a nested table"
    rows: list[list[str]] = []
    for tr in (r for r in tbl if projection._ln(r) == "tr"):
        row: list[str] = []
        for tc in tables._row_cells(tr):
            lossy = {projection._ln(el) for el in tc.iter()} & _LOSSY_CELL_TAGS
            if lossy:
                return None, f"a cell contains content the comparison cannot read ({', '.join(sorted(lossy))})"
            row.append(_cell_text(tc, strip_emphasis))
        rows.append(row)
    return rows, None


def _docx_tables(
    docx_path: Path, table_id: int | None, section_key: str | None
) -> list[tuple[int, Any, str | None]]:
    """[(table_id, w:tbl, unsupported reason)] in document order: the one
    table *table_id*, or every top-level table of the *section_key* section."""
    whole = projection.project_part(docx_path, DEFAULT_PART)
    if table_id is not None:
        for event in whole.events:
            if isinstance(event, TableBoundaryEvent) and event.kind == "table_start" and event.table_id == table_id:
                return [(table_id, event.element, None)]
        raise _make_error(
            ErrorCode.TABLE_NOT_FOUND,
            f"No table with table_id {table_id} (call list_tables to enumerate the current ones).",
            {"table_id": table_id},
        )
    entry = next(
        (
            sec
            for sec in projection.find_sections_impl(docx_path)
            if sec["kind"] == "heading" and sec["section_key"] == section_key
        ),
        None,
    )
    if entry is None:
        raise _make_error(
            ErrorCode.SECTION_NOT_FOUND,
            f"section_key {section_key!r} does not name a heading section (call find_sections).",
            {"section_key": section_key},
        )
    sliced, partial = projection.slice_projection(
        whole, projection.para_index(entry["start_para_ref"]), projection.para_index(entry["end_para_ref"])
    )
    partial_ids = {rec["table_id"] for rec in partial if rec.get("kind") == "partial_table"}
    found: list[tuple[int, Any, str | None]] = []
    depth = 0
    for event in sliced.events:
        if not isinstance(event, TableBoundaryEvent):
            continue
        if event.kind == "table_start":
            if depth == 0:
                reason = "table lies only partly inside the section" if event.table_id in partial_ids else None
                found.append((event.table_id, event.element, reason))
            depth += 1
        else:
            depth -= 1
    return found


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _compare(docx_rows: list[list[str]], md_rows: list[list[str]]) -> tuple[list[dict[str, Any]], dict[str, int] | None]:
    differences: list[dict[str, Any]] = []
    n_rows = max(len(docx_rows), len(md_rows))
    for r in range(n_rows):
        drow = docx_rows[r] if r < len(docx_rows) else None
        mrow = md_rows[r] if r < len(md_rows) else None
        n_cols = max(len(drow or []), len(mrow or []))
        for c in range(n_cols):
            d = drow[c] if drow is not None and c < len(drow) else None
            m = mrow[c] if mrow is not None and c < len(mrow) else None
            dn = None if d is None else _norm(d)
            mn = None if m is None else _norm(m)
            if dn != mn:
                differences.append({"row": r + 1, "col": c + 1, "docx_text": dn, "md_text": mn})
    dims = {
        "docx_rows": len(docx_rows),
        "docx_cols": max((len(r) for r in docx_rows), default=0),
        "md_rows": len(md_rows),
        "md_cols": max((len(r) for r in md_rows), default=0),
    }
    differs = dims["docx_rows"] != dims["md_rows"] or dims["docx_cols"] != dims["md_cols"]
    return differences, (dims if differs else None)


def _read_markdown(markdown: str | None, markdown_path: str | None) -> str:
    if (markdown is None) == (markdown_path is None):
        raise _invalid(
            "pass exactly one of markdown (inline text) or markdown_path",
            markdown_given=markdown is not None,
            markdown_path_given=markdown_path is not None,
        )
    if markdown is not None:
        return markdown
    assert markdown_path is not None
    resolved = paths.resolve_allowed_docx_path(markdown_path, must_exist=True)
    if not resolved.is_file():
        raise _invalid(
            f"Path is not a regular file: {markdown_path!r}",
            markdown_path=markdown_path,
            resolved_path=str(resolved),
        )
    return resolved.read_text(encoding="utf-8")


def _select_single(
    md_tables: list[_MdTable], headings: list[str], md_heading: str | None, md_header_row: str | None
) -> _MdTable:
    candidates = md_tables
    if md_heading is not None:
        hits = _heading_matches(headings, md_heading)
        if not hits:
            raise _invalid(f"no markdown heading matches md_heading {md_heading!r}", md_heading=md_heading)
        if len(hits) > 1:
            raise _invalid(
                f"markdown heading {md_heading!r} appears {len(hits)} times; refusing to guess which one is meant",
                md_heading=md_heading,
            )
        candidates = [t for t in md_tables if t.heading == hits[0]]
        if not candidates:
            raise _invalid(f"no markdown table under heading {md_heading!r}", md_heading=md_heading)
    if md_header_row is not None:
        wanted = _split_header_row(md_header_row)
        candidates = [t for t in candidates if t.header == wanted]
        if not candidates:
            raise _invalid(
                "no markdown table has the header row given in md_header_row"
                + (" under that heading" if md_heading is not None else ""),
                md_header_row=md_header_row,
            )
        if len(candidates) > 1:
            raise _invalid(
                f"md_header_row matches {len(candidates)} markdown tables; add md_heading to choose",
                md_header_row=md_header_row,
            )
    elif md_heading is None:
        if not md_tables:
            raise _invalid("the markdown contains no pipe table")
        if len(md_tables) > 1:
            raise _invalid(
                f"the markdown contains {len(md_tables)} tables; pass md_heading and/or md_header_row to choose one",
                table_count=len(md_tables),
            )
    return candidates[0]


def diff_table_vs_markdown_impl(
    path: str,
    table_id: int | None = None,
    markdown: str | None = None,
    markdown_path: str | None = None,
    md_heading: str | None = None,
    md_header_row: str | None = None,
    section_key: str | None = None,
    source: str = "auto",
    strip_emphasis: bool = True,
) -> dict[str, Any]:
    if (table_id is None) == (section_key is None):
        raise _invalid(
            "pass exactly one of table_id (compare one table) or section_key (compare every table in a section)",
            table_id=table_id,
            section_key=section_key,
        )
    if section_key is not None and md_header_row is not None:
        raise _invalid("md_header_row applies only with table_id; a section_key comparison pairs tables by header")
    md_text = _read_markdown(markdown, markdown_path)
    warnings: list[str] = []
    md_tables, headings = parse_markdown_tables(md_text, strip_emphasis, warnings)

    # Single mode refuses a bad markdown selector before any docx (or live pane) read.
    md_table = _select_single(md_tables, headings, md_heading, md_header_row) if table_id is not None else None
    with reads_live.read_source(path, source) as rs:
        found = _docx_tables(rs.local_path, table_id, section_key)
        if md_table is not None:
            result = _single(found[0], md_table, strip_emphasis)
        else:
            result = _multi(rs.local_path, section_key, found, md_tables, headings, md_heading, strip_emphasis)
        result["path"] = str(rs.resolved)
        result["warnings"] = warnings
        return rs.annotate(result)


def _single(
    docx_entry: tuple[int, Any, str | None], md_table: _MdTable, strip_emphasis: bool
) -> dict[str, Any]:
    tid, tbl, forced_reason = docx_entry
    rows, reason = (None, forced_reason) if forced_reason else _extract_docx_table(tbl, strip_emphasis)
    result: dict[str, Any] = {"table_id": tid, "md_table_index": md_table.index}
    if rows is None:
        result.update(identical=False, differences=[], unsupported=[{"table_id": tid, "reason": reason}])
        return result
    differences, dims = _compare(rows, md_table.rows)
    result.update(identical=not differences and dims is None, differences=differences)
    if dims is not None:
        result["dimensions"] = dims
    return result


def _multi(
    docx_path: Path,
    section_key: str | None,
    found: list[tuple[int, Any, str | None]],
    md_tables: list[_MdTable],
    headings: list[str],
    md_heading: str | None,
    strip_emphasis: bool,
) -> dict[str, Any]:
    if md_heading is None:
        entry = next(s for s in projection.find_sections_impl(docx_path) if s["section_key"] == section_key)
        md_heading = entry["heading_text"]
    hits = _heading_matches(headings, md_heading or "")
    if not hits:
        raise _invalid(f"no markdown heading matches {md_heading!r}", md_heading=md_heading)
    if len(hits) > 1:
        raise _invalid(
            f"markdown heading {md_heading!r} appears {len(hits)} times; refusing to guess which one is meant",
            md_heading=md_heading,
        )
    under = [t for t in md_tables if t.heading == hits[0]]
    if not under:
        raise _invalid(f"no markdown table under heading {md_heading!r}", md_heading=md_heading)
    if not found:
        raise _invalid(f"no Word table in section {section_key!r}", section_key=section_key)

    md_by_header: dict[tuple[str, ...], _MdTable] = {}
    for t in under:
        key = tuple(t.header)
        if key in md_by_header:
            raise _invalid(
                f"two markdown tables under {md_heading!r} share the header row {list(key)!r}; refusing to pair them",
                header=list(key),
            )
        md_by_header[key] = t

    unsupported: list[dict[str, Any]] = []
    extracted: list[tuple[int, tuple[str, ...], list[list[str]]]] = []
    seen_headers: set[tuple[str, ...]] = set()
    for tid, tbl, forced in found:
        rows, reason = (None, forced) if forced else _extract_docx_table(tbl, strip_emphasis)
        if rows is None:
            unsupported.append({"table_id": tid, "reason": reason})
            continue
        key = tuple(_norm(c) for c in rows[0]) if rows else ()
        if key in seen_headers:
            raise _invalid(
                f"two Word tables in section {section_key!r} share the header row {list(key)!r}; refusing to pair them",
                header=list(key),
            )
        seen_headers.add(key)
        extracted.append((tid, key, rows))

    compared: list[dict[str, Any]] = []
    unmatched_docx: list[int] = []
    matched_md: set[int] = set()
    for tid, key, rows in extracted:
        md_table = md_by_header.get(key)
        if md_table is None:
            unmatched_docx.append(tid)
            continue
        matched_md.add(md_table.index)
        differences, dims = _compare(rows, md_table.rows)
        entry_out: dict[str, Any] = {
            "table_id": tid,
            "md_table_index": md_table.index,
            "header": list(key),
            "differences": differences,
        }
        if dims is not None:
            entry_out["dimensions"] = dims
        compared.append(entry_out)
    unmatched_md = [t.index for t in under if t.index not in matched_md]
    clean = all(not t["differences"] and "dimensions" not in t for t in compared)
    return {
        "identical": clean and not unmatched_docx and not unmatched_md and not unsupported,
        "tables": compared,
        "unmatched_docx": unmatched_docx,
        "unmatched_md": unmatched_md,
        "unsupported": unsupported,
    }
