"""Shared shaping and scoping for ``list_open_items`` (issue #111).

``fields="summary"`` rows, the ``section_key``/``within_text`` paragraph
scope, and the live-mode paragraph-text index live here so file mode
(``tracked_changes``) and live mode (``live/comments_live``) apply one rule.
Everything is pure: callers hand in a projection and get indexes or
trimmed rows back.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from . import projection
from .errors import ErrorCode, _make_error
from .locate import locate

ANCHOR_CHARS = 60
FIELDS = ("summary", "full")


def validate_fields(fields: str) -> None:
    if fields not in FIELDS:
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"fields must be one of {FIELDS}, got {fields!r}", {"fields": fields}
        )


def clip(text: str | None) -> str:
    return (text or "")[:ANCHOR_CHARS]


def summary_comment(record: dict[str, Any], anchor: str | None) -> dict[str, Any]:
    """One ``fields="summary"`` comment row. *anchor* is the record's anchor text
    (file ``quoted_text``, live ``anchor_text`` -- the per-mode adapters differ)."""
    row = {
        "comment_id": record["comment_id"],
        "author": record.get("author"),
        "resolved": bool(record.get("resolved", False)),
        "anchor": clip(anchor),
        "reply_count": len(record.get("replies") or []),
    }
    if record.get("scope_uncertain"):
        row["scope_uncertain"] = True
    return row


COMMENT_COLUMNS = ["comment_id", "author", "resolved", "anchor", "reply_count"]
FILE_SUGGESTION_COLUMNS = ["suggestion_id", "kind", "author", "anchor"]
LIVE_SUGGESTION_COLUMNS = ["revision_id", "type", "author", "anchor"]


def columnar(rows: list[dict[str, Any]], columns: list[str]) -> tuple[list[str], list[list[Any]]]:
    """(columns, rows-as-lists): keys named once instead of per row. A ``scope_uncertain`` column is
    appended only when some row carries the flag."""
    cols = list(columns)
    if any(r.get("scope_uncertain") for r in rows):
        cols.append("scope_uncertain")
    return cols, [[bool(r.get(c)) if c == "scope_uncertain" else r.get(c) for c in cols] for r in rows]


# Top-level keys kept even when null/empty: JennyStack reads comments/source/path, and a null or []
# pending_suggestions / suggestions / complete carries meaning (unknown vs none).
_ALWAYS_KEEP = frozenset({"comments", "source", "path", "pending_suggestions", "suggestions", "complete", "scope",
                          "correlation", "comment_rows", "suggestion_rows"})
_SUMMARY_DROPS = frozenset({"pane", "session_epoch", "revision_coverage", "pending_suggestions_reason",
                            "file_side_available"})


def lean_top(result: dict[str, Any], *, summary: bool) -> dict[str, Any]:
    """Drop top-level keys whose value is None or an empty list/dict (False and 0 stay). In summary
    mode the diagnostics go too: `pane`, `session_epoch` (every live id already carries it),
    `revision_coverage`, `pending_suggestions_reason` (`suggestion_rows: null` already says the
    suggestions are unknown) and `file_side_available`; `path` stays (consumers check it names the
    right document)."""
    drop = _SUMMARY_DROPS if summary else set()
    return {
        k: v
        for k, v in result.items()
        if k not in drop and (k in _ALWAYS_KEEP or not (v is None or (isinstance(v, (list, dict)) and not v)))
    }


def paragraph_scope(
    docx_path: Path, proj: projection.Projection, section_key: str | None, within_text: str | None
) -> set[int] | None:
    """Indexes (into ``proj.paragraphs``) of the paragraphs selected by *section_key* and/or
    *within_text* (both given: their intersection), or None when neither is given.

    Unknown *section_key* raises INVALID_INPUT listing the heading keys; *within_text* goes
    through ``locate`` and must match exactly once (ZERO_MATCH / MATCH_COUNT_MISMATCH otherwise).
    """
    if section_key is None and within_text is None:
        return None
    index_by_ref = {m.para_ref: i for i, m in enumerate(proj.paragraphs)}
    scope: set[int] | None = None
    if section_key is not None:
        headings = [s for s in projection.find_sections_impl(docx_path) if s["kind"] == "heading"]
        match = next((s for s in headings if s["section_key"] == section_key), None)
        if match is None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"section_key {section_key!r} is not a heading section of this document; "
                "call find_sections for the available keys.",
                {"section_key": section_key, "available_section_keys": [s["section_key"] for s in headings]},
            )
        lo, hi = index_by_ref[match["start_para_ref"]], index_by_ref[match["end_para_ref"]]
        scope = set(range(lo, hi + 1))
    if within_text is not None:
        start, end = locate(within_text, proj, 1).spans[0]
        hit = {
            index_by_ref[para_ref]
            for s, e, para_ref, _run in proj.offset_map
            if s < end and e > start and para_ref in index_by_ref
        }
        scope = hit if scope is None else scope & hit
    return scope


def _norm(text: str) -> str:
    return " ".join(text.split())


_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _squash(text: str) -> str:
    """All whitespace removed (tabs and breaks have no w:t text in the live revision markup)."""
    return "".join(text.split())


def paragraph_markup_texts(proj: projection.Projection) -> list[str]:
    """Each paragraph's w:t + w:delText text (deleted text included), aligned with
    ``proj.paragraphs`` -- exactly how revisions_live builds a revision's ``paragraph_context``."""
    return [
        "".join(n.text or "" for n in m.element.iter() if n.tag in (_W_NS + "t", _W_NS + "delText"))
        if m.element is not None
        else ""
        for m in proj.paragraphs
    ]


def paragraph_text_index(
    proj: projection.Projection, scope: set[int], *, markup: bool = False
) -> dict[str, tuple[int, int]]:
    """normalized paragraph text -> (paragraphs with that text, how many are in *scope*).
    Live comments only know their anchor paragraph's text, never its position. ``markup=True``
    keys on the revision-markup text (deleted text included, whitespace squashed) that live
    suggestions carry; look those up with ``classify_text(..., squash=True)``."""
    counts: dict[str, list[int]] = {}
    texts = paragraph_markup_texts(proj) if markup else projection.paragraph_texts(proj)
    for i, text in enumerate(texts):
        entry = counts.setdefault(_squash(text) if markup else _norm(text), [0, 0])
        entry[0] += 1
        entry[1] += i in scope
    return {text: (total, inside) for text, (total, inside) in counts.items()}


def classify_text(index: dict[str, tuple[int, int]], text: str | None, *, squash: bool = False) -> str:
    """``"in"`` | ``"out"`` | ``"uncertain"`` for a live item whose paragraph text is *text*.

    Uncertain (kept in the result, flagged ``scope_uncertain``) when the text is unknown or not
    found, or when identical paragraphs straddle the scope boundary so the item could be either.
    """
    key = _squash(text or "") if squash else _norm(text or "")
    entry = index.get(key) if key else None
    if entry is None:
        return "uncertain"
    total, inside = entry
    if inside == 0:
        return "out"
    return "in" if inside == total else "uncertain"
