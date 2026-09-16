"""geometry.py — pure, Word-free functions for export_pdf's per-section page
spans (issue #102): mapping projection.py's para_ref order onto Word's own
1-based paragraph ordinals, verifying a probed heading's text against
find_sections_impl's heading_text, and turning page/vertical-position probes
into a page-span per section. Every function here takes plain data (section
dicts from projection.find_sections_impl, a paragraph_geometry probe result,
a page height) and returns plain data — no Word automation, no file I/O — so
the whole module is unit-testable without Word or a real .docx on disk.

See server.py's execute_export_pdf for how this is wired to
projection.find_sections_impl / projection.project_part and
render.render_word(..., probe_paragraphs=...).

Design (issue #102's spec, "Server layer"; end-geometry source amended
2026-09-16 per a follow-up from the issue's lead):
  - Ordinal mapping: ordinal(para_ref) = index of that para_ref in
    [m.para_ref for m in proj.paragraphs] + 1 — the same main-story
    paragraph order find_sections_impl indexes as its own all_para_refs,
    so start_para_ref/end_para_ref map directly (ordinal_of below).
  - Probing is minimized: the start ordinal of every section, plus one
    more ordinal per section for its END geometry (build_probe_ordinals).
    A section's end is the FULL DOCUMENT's next heading's start — NOT
    just "the next entry in the *sections* list passed in" — whenever
    the document has one (each section dict carries its own
    "next_start_para_ref", precomputed by the caller from the unfiltered
    heading list; see server.py's _plan_section_probes). This matters
    specifically for an explicit, single-section section_keys=[...] call
    (the skills' primary use: a page-budget check for one section): even
    though only that one section is requested and returned, its end
    still borrows the real next heading's start when the document has
    one, rather than falling back to an approximation just because nothing
    else was requested. ONLY a section with no next heading AT ALL — the
    true last heading in the whole document — uses its own end_para_ref
    probe, nudged by +0.03 of a page (capped at 1.0) to approximate that
    last line's own height (assemble_sections below). The borrowed next
    heading's own text is never verified against anything (only a
    section's OWN start-probe text is checked against its heading_text);
    it exists purely as a page/vertical-position boundary.
  - Verification is not optional: a probed start paragraph's text must
    equal find_sections_impl's heading_text for that section (both run
    through normalize_probe_text), and every probed ordinal this module
    needs must have come back without an "error" key. assemble_sections
    never raises; it returns (None, "<detail>") on failure and lets the
    caller decide default-mode degrade vs explicit-mode raise (server.py).
"""

from __future__ import annotations

import re
from typing import Any

# ---------------------------------------------------------------------------
# Text normalization for probe-text verification
# ---------------------------------------------------------------------------

# Curly -> straight quote equivalence, the same mapping locate.py's
# _QUOTE_MAP uses. Kept local (not imported) rather than shared: this
# module has no other dependency on locate.py, and the two normalizers
# serve different callers with different rungs (locate.py's is a 4-rung
# search ladder; this is a single flat normalization for an equality
# check).
_QUOTE_MAP = str.maketrans(
    {
        "‘": "'",  # LEFT SINGLE QUOTATION MARK
        "’": "'",  # RIGHT SINGLE QUOTATION MARK
        "‚": "'",  # SINGLE LOW-9 QUOTATION MARK
        "‛": "'",  # SINGLE HIGH-REVERSED-9 QUOTATION MARK
        "′": "'",  # PRIME
        "‵": "'",  # REVERSED PRIME
        "“": '"',  # LEFT DOUBLE QUOTATION MARK
        "”": '"',  # RIGHT DOUBLE QUOTATION MARK
        "„": '"',  # DOUBLE LOW-9 QUOTATION MARK
        "‟": '"',  # DOUBLE HIGH-REVERSED-9 QUOTATION MARK
        "″": '"',  # DOUBLE PRIME
        "‶": '"',  # REVERSED DOUBLE PRIME
        "«": '"',  # LEFT-POINTING DOUBLE ANGLE QUOTATION MARK
        "»": '"',  # RIGHT-POINTING DOUBLE ANGLE QUOTATION MARK
    }
)

_SOFT_HYPHEN = "­"

# Word control/field characters a paragraph-geometry probe's `content of
# r` can carry that never appear in projection.py's own run-text view:
# the C0 range's field/tab-stop markers (U+0001-U+0008, U+000B, U+000C,
# U+000E-U+001F) and the C1 range (U+007F-U+009F), plus a literal NUL.
# TAB (U+0009), LF (U+000A), and CR (U+000D) are deliberately excluded
# from this drop-set — they are folded into a single space by the
# whitespace-collapse step below instead of being deleted outright, so
# e.g. "Factor 1<TAB>Overview" still normalizes with a space between the
# words rather than silently merging them.
_CONTROL_CODEPOINTS = (
    list(range(0x09)) + [0x0B, 0x0C] + list(range(0x0E, 0x20)) + list(range(0x7F, 0xA0))
)
_CONTROL_RE = re.compile("[" + "".join(chr(c) for c in _CONTROL_CODEPOINTS) + "]")
_WS_RE = re.compile(r"\s+")


def normalize_probe_text(s: str | None) -> str:
    """Normalize a paragraph-geometry probe's ``text`` (or a
    find_sections_impl ``heading_text``) for verification comparison:
    strip, collapse whitespace runs to a single space, straighten curly
    quotes, and drop soft hyphens (U+00AD) and Word's own control/field
    characters (see _CONTROL_RE above). Comparison is case-sensitive — a
    heading probed as "background" when find_sections_impl expects
    "Background" is exactly the kind of mismatch this exists to catch,
    not paper over."""
    if not s:
        return ""
    out = s.translate(_QUOTE_MAP)
    out = out.replace(_SOFT_HYPHEN, "")
    out = _CONTROL_RE.sub("", out)
    out = _WS_RE.sub(" ", out).strip()
    return out


# ---------------------------------------------------------------------------
# Ordinal mapping
# ---------------------------------------------------------------------------


def ordinal_of(para_ref: str, all_para_refs: list[str]) -> int:
    """The 1-based Word paragraph ordinal for *para_ref*.

    all_para_refs.index(para_ref) + 1 — projection.project_part()'s
    .paragraphs list is the same main-story paragraph order
    find_sections_impl indexes as its own all_para_refs (projection.py),
    which is intended to match Word's `paragraphs of document` order
    (render.py's module docstring, PARAGRAPH GEOMETRY PROBE section)."""
    return all_para_refs.index(para_ref) + 1


def build_probe_ordinals(sections: list[dict[str, Any]], all_para_refs: list[str]) -> list[int]:
    """Distinct 1-based ordinals export_pdf must pass to render_word() as
    probe_paragraphs for *sections* (a document-order list of
    find_sections_impl "heading" entries, already filtered to any
    requested section_keys, each carrying its own "next_start_para_ref"
    -- see server.py's _plan_section_probes): the start ordinal of every
    section, plus ONE more ordinal per section for its end geometry --
    the FULL document's next heading's start ordinal
    (ordinal_of(section["next_start_para_ref"], ...)) when that section
    has one, or its own end_para_ref ordinal (the last-paragraph
    approximation case) when it does not, i.e. when it is the document's
    own last heading. Returns [] for an empty *sections* list (no probes
    needed — the caller should then pass probe_paragraphs=None so
    render_word() adds no PROBE/PAGEH overhead at all)."""
    if not sections:
        return []
    ordinals: set[int] = set()
    for section in sections:
        ordinals.add(ordinal_of(section["start_para_ref"], all_para_refs))
        next_start_para_ref = section["next_start_para_ref"]
        if next_start_para_ref is not None:
            ordinals.add(ordinal_of(next_start_para_ref, all_para_refs))
        else:
            ordinals.add(ordinal_of(section["end_para_ref"], all_para_refs))
    return sorted(ordinals)


# ---------------------------------------------------------------------------
# Geometry assembly
# ---------------------------------------------------------------------------


def assemble_sections(
    sections: list[dict[str, Any]],
    all_para_refs: list[str],
    paragraph_geometry: dict[int, dict[str, Any]],
    page_height_pt: float | None,
) -> tuple[list[dict[str, Any]] | None, str | None]:
    """Turn a render_word() probe result into export_pdf's "sections" list.

    *sections* is find_sections_impl's "heading"-kind entries, in document
    order, already filtered to any requested section_keys, each carrying
    its own "next_start_para_ref" (the FULL document's next heading's
    start_para_ref, or None when the section is the document's own last
    heading — see server.py's _plan_section_probes; NOT relative to the
    filtered *sections* list itself, which may omit that next heading
    entirely on an explicit, single-section section_keys=[...] call).
    *paragraph_geometry* and *page_height_pt* are render_word()'s own
    probe results (render.py, issue #102's paragraph geometry probe).
    Returns (sections_result, None) on success, or (None, "<detail>")
    when verification fails — the caller (server.py's execute_export_pdf)
    decides whether that means degrading to sections: null (default mode,
    section_keys is None) or raising SECTION_GEOMETRY_UNAVAILABLE
    (explicit mode); this function never raises.

    An empty *sections* list (no headings, or none requested) returns
    ([], None) immediately without inspecting paragraph_geometry /
    page_height_pt at all — no probes were requested for it either (see
    build_probe_ordinals), so neither is expected to carry anything
    meaningful.

    Failure conditions, checked in this order:
      1. page_height_pt is None (render_word() itself already raises
         RENDER_FAILED when probes were requested but no PAGEH line came
         back, so this should not normally be reachable — checked here
         too, defensively, rather than assumed).
      2. any ordinal this function needs (every section's start, and per
         section either the full document's next heading's start or —
         only for the document's own last heading — that section's own
         end_para_ref) is missing from paragraph_geometry, or its entry
         is a probe error ({"error": ...}).
      3. a section's start-probe text, after normalize_probe_text(), does
         not equal its find_sections_impl heading_text, also normalized
         (the borrowed next-heading probe used for END geometry is never
         text-verified against anything — it is a boundary, not a claim
         about what that next section is).

    Geometry per section: start_page/start_fraction always come from the
    section's own start probe. end_page/end_fraction come from the FULL
    document's next heading's start probe when "next_start_para_ref" is
    not None — regardless of whether that next heading is itself part of
    *sections* — or, only when it IS None (this section is the document's
    last heading), from that section's own end_para_ref probe with
    end_fraction nudged by +0.03, capped at 1.0, to approximate that last
    line's own height. start_fraction/end_fraction are vpos_pt /
    page_height_pt, and pages = (end_page + end_fraction) - (start_page +
    start_fraction). Fractions and pages are all rounded to 2 places.
    start_paragraph/end_paragraph are always the section's OWN structural
    ordinals (start_para_ref/end_para_ref), never the borrowed next
    heading's — they describe this section's own extent in the document,
    independent of which probe supplied its end geometry."""
    if not sections:
        return [], None

    if page_height_pt is None:
        return None, "page_height_pt was not returned by the render (no PAGEH line)"

    starts = [ordinal_of(s["start_para_ref"], all_para_refs) for s in sections]

    # Per section, the ordinal that supplies its END geometry, and
    # whether that ordinal is the last-heading approximation (True) or a
    # borrowed next-heading start (False, never text-verified).
    end_ordinals: list[int] = []
    end_is_approximation: list[bool] = []
    for section in sections:
        next_start_para_ref = section["next_start_para_ref"]
        if next_start_para_ref is not None:
            end_ordinals.append(ordinal_of(next_start_para_ref, all_para_refs))
            end_is_approximation.append(False)
        else:
            end_ordinals.append(ordinal_of(section["end_para_ref"], all_para_refs))
            end_is_approximation.append(True)

    needed_ordinals = set(starts) | set(end_ordinals)
    for ordinal in sorted(needed_ordinals):
        entry = paragraph_geometry.get(ordinal)
        if entry is None:
            return None, f"paragraph ordinal {ordinal}: missing from paragraph_geometry"
        if "error" in entry:
            return None, f"paragraph ordinal {ordinal}: probe error: {entry['error']}"

    result: list[dict[str, Any]] = []
    for idx, section in enumerate(sections):
        start_ordinal = starts[idx]
        start_probe = paragraph_geometry[start_ordinal]

        expected = normalize_probe_text(section["heading_text"])
        got = normalize_probe_text(start_probe.get("text"))
        if got != expected:
            return None, (
                f"section {section['section_key']!r} (paragraph {start_ordinal}): "
                f"heading text mismatch, expected {expected!r}, got {got!r}"
            )

        start_page = start_probe["page"]
        start_fraction = round(start_probe["vpos_pt"] / page_height_pt, 2)

        end_probe = paragraph_geometry[end_ordinals[idx]]
        end_page = end_probe["page"]
        if end_is_approximation[idx]:
            end_fraction = round(min(1.0, end_probe["vpos_pt"] / page_height_pt + 0.03), 2)
        else:
            end_fraction = round(end_probe["vpos_pt"] / page_height_pt, 2)

        pages = round((end_page + end_fraction) - (start_page + start_fraction), 2)

        result.append(
            {
                "section_key": section["section_key"],
                "heading_text": section["heading_text"],
                "start_page": start_page,
                "start_fraction": start_fraction,
                "end_page": end_page,
                "end_fraction": end_fraction,
                "pages": pages,
                "start_paragraph": start_ordinal,
                "end_paragraph": ordinal_of(section["end_para_ref"], all_para_refs),
                "text_verified": True,
            }
        )

    return result, None
