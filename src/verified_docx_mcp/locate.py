# Adapted from GoogleDocs-MCP src/verified_googledocs_mcp/verify.py's
# normalization-ladder section (commit 374bcf8: rung constants + quote/
# whitespace/soft-hyphen normalizers ~lines 379-437, locate() ~line 557,
# orig_pos composition ~lines 592-608, the near-miss scan). Lifted per issue
# #28 WP-06 ("Lift, do not rewrite" -- "verify.py (rungs :379-382, locate()
# :557, orig_pos composition :592-608, ...)"). The normalization MECHANISM
# (exact -> curly/straight quotes -> NBSP/whitespace-run collapse ->
# soft-hyphen strip, stopping at the first rung with >=1 match, orig_pos
# composed rung-over-rung) and the near-miss scan are copied as-is.
#
# What does NOT carry over: verify.py's haystack is a Docs API tab-JSON tree
# (_flatten_tab), addressed in UTF-16 code units because that is the Docs API
# batchUpdate index space. This server's haystack is projection.py's flat
# Python str + offset_map over a live OOXML part -- a Python string index
# *is* this server's own address space (no surrogate-pair accounting is
# needed; text_edit.py splices runs directly, it never calls a remote API
# that counts in UTF-16), so every UTF-16-specific helper in the Google
# module (_utf16_width, u16_map, _u16_to_codepoint) has no analogue here.
#
# STRUCTURAL_BOUNDARY also differs in scope: the Google version raises on
# "crosses a paragraph" (its own block_map has no narrower unit). Issue #28's
# WP-06 text specifically names w:p/w:tbl/w:tc as the things a match must not
# cross, so _boundary_kinds below reports which of those a paragraph-crossing
# match actually crossed (a table only ever changes at a paragraph boundary
# too, so "crosses w:p" remains the trigger condition; w:tbl/w:tc are reported
# as refinements of *which* paragraph-boundary it was, from
# projection.ParagraphMeta.container_chain).
#
# RUNG_FIELD (issue #28 WP-06 plan text: "add RUNG_FIELD") is a DELIBERATE
# DEVIATION from the plan, decided after the base this branch rebased onto
# (PR #3, feat/28-wp-04-markdown-mutations @ 7e5aecf) fixed a real WP-03
# defect this module was originally built against: a field's markdown
# rendering used to emit BOTH its live result text AND a "[FIELD:instr]"
# placeholder back to back (projection.py's pre-fix _events_to_markdown),
# so the ORIGINAL RUNG_FIELD design substituted a field's live span with
# that uppercase placeholder and re-searched at a fifth rung -- letting a
# `find` copied from a markdown read match the field underneath.
#
# On the corrected projection (projection._events_to_markdown post-fix), a
# field's markdown rendering is exactly one of two things: a field WITH a
# resolved result renders as that result text directly -- no placeholder at
# all, so a `find` copied from a markdown read already contains ordinary
# live text and already matches at RUNG_EXACT (or 2-4); a field WITH NO
# result renders as a lowercase "[field:instr]" placeholder, but contributes
# ZERO RunEvents to the projection (no field_result=True runs exist for it
# at all) -- it is a ZERO-WIDTH point between two ordinary runs, not a span.
#
# That second fact is why a literal RUNG_FIELD is not just unneeded but
# actively unsafe to build: _atoms_for_span's overlap test (text_edit.py,
# `s < end and e > start`) can never select a zero-width point, so "locating"
# a no-result field via a substituted placeholder would produce a match
# whose span cannot be resolved to any real run atom -- and the field's own
# skeleton (w:fldChar begin/separate/end, or w:fldSimple) is not a RunEvent
# at all, so nothing in text_edit.py's run-splitter would touch it. A
# replace_text/format_text call whose matched span happened to straddle that
# skeleton (e.g. "before [field:PAGE] after" matching across the gap) would
# splice through it and leave an orphaned, unbalanced field behind --
# corrupting the document to buy a rung with nothing safe to locate.
#
# So RUNG_FIELD is NOT a haystack-substitution rung here (locate()'s own
# match ladder is back to the 4 Google-derived rungs below -- a completely
# different "rung" from core/document-backend-protocol.md's unrelated
# 4-rung EDIT ladder, replace_text/format_text/replace_range_markdown/
# replace_body_markdown; the two concepts just share the word). Instead:
# a match that overlaps a field's cached result text (field_result=True
# RunEvents) surfaces "touches_field_result" as a non-fatal WARNING --
# same shape as crosses_comment_range/crosses_revision below -- telling the
# caller they just matched/edited a value Word may silently recompute on
# its own (F9), without inventing a matching mechanism for it. The no-result
# case is an explicit, named SCOPE LIMIT, not a silent gap: a `find` copied
# verbatim as "[field:instr]" legitimately ZERO_MATCHes (real near-miss
# diagnostics), because there is genuinely no live text there and no safe
# way for a general-purpose text tool to splice across an unmodeled field
# skeleton. The lead may still want the plan text amended to match; that
# call is theirs, which is why this reasoning is spelled out here rather
# than only living in a PR description.
#
# "crosses_comment_range"/"crosses_revision"/"touches_field_result" are
# WARNINGS here (surfaced in LocateResult.warnings), not refusals -- WP-06
# only detects; WP-07 (which depends on this module) adds the actual
# TRACKED_CHANGES_PRESENT refusal for a revision-crossing replace_text/
# format_text on top of the crosses_revision detection (see text_edit.py
# and tracked_changes.py) -- touches_field_result stays warning-only; no WP
# in this plan refuses on it.
"""locate(): find `needle` in a projection.Projection's flat text via a
4-rung normalization ladder (exact -> curly/straight quotes -> NBSP/
whitespace-run collapse -> soft-hyphen strip), returning spans in
Python-string codepoint offsets (this server's own address space -- see
module docstring for why no UTF-16 accounting is needed, unlike the Google
server this is adapted from). NOT to be confused with the unrelated 4-rung
EDIT ladder in core/document-backend-protocol.md (replace_text/format_text/
replace_range_markdown/replace_body_markdown) -- see this module's header
comment for the full RUNG_FIELD deviation this module carries instead of a
fifth match rung.
"""

from __future__ import annotations

import dataclasses
import difflib
import re
from typing import Any

from .errors import ErrorCode, _make_error
from .projection import FieldEvent, Projection, RunEvent

# ---------------------------------------------------------------------------
# Rung labels
# ---------------------------------------------------------------------------

RUNG_EXACT = "exact"
RUNG_QUOTES = "curly_straight_quotes"
RUNG_WHITESPACE = "nbsp_whitespace_runs"
RUNG_SOFTHYPHEN = "soft_hyphen_strip"

# Warning key (LocateResult.warnings), NOT a rung -- see this module's
# header comment for why a fifth RUNG_FIELD match rung was not built.
WARNING_TOUCHES_FIELD_RESULT = "touches_field_result"

# Curly <-> straight quote mapping (both single and double) -- verbatim from
# verify.py's _QUOTE_MAP.
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

# NBSP + other Unicode space separators (verify.py's _NBSP_RE, spelled out in
# escapes here rather than as literal invisible characters in source).
_NBSP_CHARS = (
    "              　"
)
_NBSP_RE = re.compile("[" + _NBSP_CHARS + "]")
_WS_RUN_RE = re.compile(r"\s+")
_SOFTHYPHEN = "­"


def _norm_quotes(s: str) -> tuple[str, list[int]]:
    """Apply curly->straight quote equivalence. 1:1 so orig_pos is trivial."""
    n = s.translate(_QUOTE_MAP)
    return n, list(range(len(n) + 1))


def _norm_whitespace(s: str) -> tuple[str, list[int]]:
    """Collapse NBSP + whitespace runs to a single space, tracking orig_pos."""
    replaced = _NBSP_RE.sub(" ", s)
    parts: list[str] = []
    orig_pos: list[int] = []
    i = 0
    for m in _WS_RUN_RE.finditer(replaced):
        for j in range(i, m.start()):
            parts.append(replaced[j])
            orig_pos.append(j)
        parts.append(" ")
        orig_pos.append(m.start())
        i = m.end()
    for j in range(i, len(replaced)):
        parts.append(replaced[j])
        orig_pos.append(j)
    orig_pos.append(len(s))  # sentinel
    return "".join(parts), orig_pos


def _norm_softhyphen(s: str) -> tuple[str, list[int]]:
    """Strip soft hyphens (U+00AD), tracking orig_pos."""
    parts: list[str] = []
    orig_pos: list[int] = []
    for i, ch in enumerate(s):
        if ch == _SOFTHYPHEN:
            continue
        parts.append(ch)
        orig_pos.append(i)
    orig_pos.append(len(s))  # sentinel
    return "".join(parts), orig_pos


def _compose_orig_pos(inner: list[int], outer: list[int]) -> list[int]:
    """Compose two orig_pos maps: inner is applied first, then outer."""
    return [outer[p] for p in inner]


def _find_all(needle: str, haystack: str) -> list[int]:
    """Return all start positions of non-overlapping needle in haystack."""
    positions: list[int] = []
    start = 0
    nlen = len(needle)
    if nlen == 0:
        return positions
    while True:
        pos = haystack.find(needle, start)
        if pos == -1:
            break
        positions.append(pos)
        start = pos + nlen
    return positions


# ---------------------------------------------------------------------------
# touches_field_result: field_result spans, kept for the WARNING below --
# see this module's header comment for why there is no haystack-substitution
# rung here instead.
# ---------------------------------------------------------------------------


def _field_spans(proj: Projection) -> list[tuple[int, int, str]]:
    """(start, end, instr) for each RESULT-bearing field in *proj*, in
    projection.text codepoint offsets -- the same [start, end) a live
    field's constituent field_result=True RunEvents occupy, paired with the
    FieldEvent that closes them. A field with NO result contributes no
    field_result RunEvents at all (see this module's header comment) and so
    never appears here -- there is no span to report for one. Every
    RunEvent has non-empty text (projection._emit_text returns early on
    empty text), so proj.offset_map has exactly one entry per RunEvent, in
    the same order -- zipping the two gives each RunEvent's own [start, end).
    """
    run_events = [e for e in proj.events if isinstance(e, RunEvent)]
    offsets_by_id = {id(e): (s, en) for e, (s, en, _pr, _rr) in zip(run_events, proj.offset_map)}

    spans: list[tuple[int, int, str]] = []
    current_start: int | None = None
    current_end: int | None = None
    for event in proj.events:
        if isinstance(event, RunEvent):
            if event.field_result:
                s, en = offsets_by_id[id(event)]
                if current_start is None:
                    current_start = s
                current_end = en
            else:
                current_start = None
                current_end = None
        elif isinstance(event, FieldEvent):
            if current_start is not None and current_end is not None:
                spans.append((current_start, current_end, event.instr))
            current_start = None
            current_end = None
    return spans


# ---------------------------------------------------------------------------
# Structural boundary + hazard-warning checks
# ---------------------------------------------------------------------------


def _touched_para_refs(proj: Projection, start: int, end: int) -> list[str]:
    refs: list[str] = []
    for s, e, para_ref, _run_ref in proj.offset_map:
        if s < end and e > start and (not refs or refs[-1] != para_ref):
            refs.append(para_ref)
    return refs


def _boundary_kinds(proj: Projection, start: int, end: int) -> set[str]:
    """Subset of {"w:p", "w:tbl", "w:tc"} describing which structural
    boundaries a [start, end) span crosses (issue #28 plan WP-06: "when a
    span crosses w:p, w:tbl, w:tc"). A table/cell boundary can only be
    crossed AT a paragraph boundary (a single w:p cannot itself straddle two
    cells), so "more than one paragraph touched" is the trigger condition;
    w:tbl/w:tc are reported as refinements of *which* paragraphs those were,
    from each one's own container_chain.
    """
    para_refs = _touched_para_refs(proj, start, end)
    if len(para_refs) <= 1:
        return set()

    kinds = {"w:p"}
    meta_by_ref = {m.para_ref: m for m in proj.paragraphs}
    chains = [meta_by_ref[r].container_chain for r in para_refs if r in meta_by_ref]
    in_table = [bool(c) for c in chains]
    if len(set(in_table)) > 1 or any(in_table):
        # Either the span moves between "in a table" and "not in a table",
        # or every touched paragraph is inside one -- either way at least
        # one table boundary (table start/end, or a row/cell change) sits
        # between them.
        kinds.add("w:tbl")
    cells = {(c[-1]["table_id"], c[-1]["row"], c[-1]["cell"]) for c in chains if c}
    if len(cells) > 1:
        kinds.add("w:tc")
    return kinds


def _overlapping_run_events(proj: Projection, start: int, end: int) -> list[RunEvent]:
    run_events = [e for e in proj.events if isinstance(e, RunEvent)]
    out: list[RunEvent] = []
    for event, (s, e, _pr, _rr) in zip(run_events, proj.offset_map):
        if s < end and e > start:
            out.append(event)
    return out


def _crosses_comment_range(proj: Projection, start: int, end: int) -> bool:
    return any(e.comment_ids for e in _overlapping_run_events(proj, start, end))


def _crosses_revision(proj: Projection, start: int, end: int) -> bool:
    return any(e.in_revision for e in _overlapping_run_events(proj, start, end))


def _touches_field_result(proj: Projection, start: int, end: int, field_spans: list[tuple[int, int, str]]) -> bool:
    """True if [start, end) overlaps any RESULT-bearing field's own span --
    the WARNING_TOUCHES_FIELD_RESULT signal (see this module's header
    comment for why this is a warning rather than a rung)."""
    return any(s < end and e > start for s, e, _instr in field_spans)


# ---------------------------------------------------------------------------
# Near-miss scan (bounded) -- verbatim algorithm from verify.py.
# ---------------------------------------------------------------------------

_NEAR_MISS_THRESHOLD = 0.6
_NEAR_MISS_STRONG_EXIT = 0.92
_NEAR_MISS_MAX_WINDOWS = 5000


def _near_miss_scan(needle: str, haystack: str) -> dict[str, Any] | None:
    nl = len(needle)
    if nl == 0 or len(haystack) == 0:
        return None

    hl = len(haystack)
    best_ratio = 0.0
    best_span: tuple[int, int] | None = None
    windows_checked = 0

    for wlen in (nl, max(1, nl - nl // 4), nl + nl // 4):
        step = max(1, wlen // 3)
        pos = 0
        while pos + wlen <= hl:
            if windows_checked >= _NEAR_MISS_MAX_WINDOWS:
                break
            window = haystack[pos : pos + wlen]
            ratio = difflib.SequenceMatcher(None, needle, window, autojunk=False).ratio()
            windows_checked += 1
            if ratio > best_ratio:
                best_ratio = ratio
                best_span = (pos, pos + wlen)
            if ratio >= _NEAR_MISS_STRONG_EXIT:
                break
            pos += step
        if best_ratio >= _NEAR_MISS_STRONG_EXIT:
            break

    if best_ratio < _NEAR_MISS_THRESHOLD or best_span is None:
        return None
    return {
        "ratio": round(best_ratio, 3),
        "span_start": best_span[0],
        "span_end": best_span[1],
        "text": haystack[best_span[0] : best_span[1]],
    }


# ---------------------------------------------------------------------------
# locate()
# ---------------------------------------------------------------------------


@dataclasses.dataclass(frozen=True)
class LocateResult:
    """Successful location of needle in a projection's flat text."""

    spans: list[tuple[int, int]]  # (start, end) codepoint offsets, per match
    rung: str
    match_count: int
    warnings: list[str]  # "crosses_comment_range" / "crosses_revision" / "touches_field_result"


def locate(needle: str, proj: Projection, expected_matches: int) -> LocateResult:
    """Find every occurrence of *needle* in *proj*, returning codepoint spans.

    Normalization ladder (stops at first rung with >=1 match; this is
    locate()'s own MATCH ladder, distinct from core/document-backend-
    protocol.md's unrelated 4-rung EDIT ladder -- see this module's header
    comment):
        1. exact
        2. curly/straight quote equivalence
        3. NBSP and whitespace-run collapse (stacked on rung 2)
        4. soft-hyphen (U+00AD) strip (stacked on rung 3)

    A match overlapping a RESULT-bearing field's own span (a value Word
    computed and may silently recompute later) is not a fifth rung -- it
    surfaces WARNING_TOUCHES_FIELD_RESULT in the returned warnings instead
    (see this module's header comment for why a literal RUNG_FIELD rung, as
    the issue #28 plan text names it, was not built).

    Raises VerifyError on:
        INVALID_INPUT          - empty needle
        ZERO_MATCH              - all rungs exhausted; near-miss in diagnostics
        MATCH_COUNT_MISMATCH    - match count != expected_matches; every span listed
        STRUCTURAL_BOUNDARY     - any match crosses a w:p/w:tbl/w:tc boundary

    D4 (issue #28 plan, this server only): expected_matches is REQUIRED, no
    default -- every caller must say how many matches it expects.
    """
    if not needle:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty")

    haystack = proj.text
    rungs_data: list[tuple[str, list[int], str, str]] = [
        (haystack, list(range(len(haystack) + 1)), needle, RUNG_EXACT),
    ]

    nh2 = haystack.translate(_QUOTE_MAP)
    op2 = list(range(len(nh2) + 1))
    nn2 = needle.translate(_QUOTE_MAP)
    rungs_data.append((nh2, op2, nn2, RUNG_QUOTES))

    nh3, op3_inner = _norm_whitespace(nh2)
    op3 = _compose_orig_pos(op3_inner, op2)
    nn3, _ = _norm_whitespace(nn2)
    rungs_data.append((nh3, op3, nn3, RUNG_WHITESPACE))

    nh4, op4_inner = _norm_softhyphen(nh3)
    op4 = _compose_orig_pos(op4_inner, op3)
    nn4, _ = _norm_softhyphen(nn3)
    rungs_data.append((nh4, op4, nn4, RUNG_SOFTHYPHEN))

    field_spans = _field_spans(proj)
    ladder_report: list[dict[str, Any]] = []

    for norm_haystack, orig_pos, norm_needle, rung_label in rungs_data:
        if not norm_needle:
            ladder_report.append({"rung": rung_label, "matches": 0})
            continue

        positions = _find_all(norm_needle, norm_haystack)
        if not positions:
            ladder_report.append({"rung": rung_label, "matches": 0})
            continue

        spans: list[tuple[int, int]] = []
        for npos in positions:
            orig_start = orig_pos[npos]
            orig_end = orig_pos[npos + len(norm_needle)]
            spans.append((orig_start, orig_end))

        for s, e in spans:
            kinds = _boundary_kinds(proj, s, e)
            if kinds:
                raise _make_error(
                    ErrorCode.STRUCTURAL_BOUNDARY,
                    f"match crosses a structural boundary: {sorted(kinds)}",
                    {
                        "rung": rung_label,
                        "spans": [{"start": a, "end": b} for a, b in spans],
                        "boundary_kinds": sorted(kinds),
                    },
                )

        actual_count = len(spans)
        if actual_count != expected_matches:
            raise _make_error(
                ErrorCode.MATCH_COUNT_MISMATCH,
                f"expected {expected_matches} match(es) but found {actual_count} at rung {rung_label!r}",
                {
                    "rung": rung_label,
                    "expected": expected_matches,
                    "actual": actual_count,
                    "spans": [{"start": a, "end": b} for a, b in spans],
                },
            )

        warnings: set[str] = set()
        for s, e in spans:
            if _crosses_comment_range(proj, s, e):
                warnings.add("crosses_comment_range")
            if _crosses_revision(proj, s, e):
                warnings.add("crosses_revision")
            if _touches_field_result(proj, s, e, field_spans):
                warnings.add(WARNING_TOUCHES_FIELD_RESULT)

        return LocateResult(spans=spans, rung=rung_label, match_count=actual_count, warnings=sorted(warnings))

    near_miss = _near_miss_scan(needle, nh4)
    raise _make_error(
        ErrorCode.ZERO_MATCH,
        "needle not found after full normalization ladder",
        {"ladder_report": ladder_report, "near_miss": near_miss},
    )
