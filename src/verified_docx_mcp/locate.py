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
import hashlib
import re
import unicodedata
from collections.abc import Callable
from typing import Any

from .errors import ErrorCode, VerifyError, _make_error
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


_DASH_MAP = str.maketrans(
    {
        "‐": "-",  # HYPHEN
        "‑": "-",  # NON-BREAKING HYPHEN
        "‒": "-",  # FIGURE DASH
        "–": "-",  # EN DASH
        "—": "-",  # EM DASH
        "―": "-",  # HORIZONTAL BAR
        "−": "-",  # MINUS SIGN
        "﹘": "-",  # SMALL EM DASH
        "﹣": "-",  # SMALL HYPHEN-MINUS
        "－": "-",  # FULLWIDTH HYPHEN-MINUS
    }
)


def normalize_for_compare(text: str) -> str:
    """Fold *text* with the same equivalences ``locate`` uses to find text
    (curly->straight quotes, NBSP-class spaces, soft-hyphen strip) plus dash
    variants and NFC, for comparing two renderings of one document
    (``diff_body_vs_file``) and for matching headings across them.

    Whitespace runs are NOT collapsed -- callers diff line by line -- but each
    line is right-stripped, because trailing spaces (and a CR before the LF)
    are the commonest phantom difference between pandoc, markitdown and
    Word projections. Newlines are kept, so line numbers stay valid.
    """
    s = unicodedata.normalize("NFC", text)
    s = s.translate(_QUOTE_MAP).translate(_DASH_MAP)
    s = _NBSP_RE.sub(" ", s).replace(_SOFTHYPHEN, "")
    return "\n".join(line.rstrip() for line in s.split("\n"))


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


_MATCH_DIAGNOSTIC_MAX = 5
_MATCH_CONTEXT_CHARS = 40


def _match_diagnostics(proj: Projection, spans: list[tuple[int, int]]) -> dict[str, Any]:
    """``matches`` (at most five, each ``{start, end, context}``) and ``total_matches``.

    Bare offsets told a caller nothing: with 40 hits it could not tell them apart. The
    context is 40 characters starting 20 before the match, so two hits read differently
    (issue #112)."""
    text = proj.text
    shown = []
    for start, end in spans[:_MATCH_DIAGNOSTIC_MAX]:
        lo = max(0, start - _MATCH_CONTEXT_CHARS // 2)
        context = text[lo : lo + _MATCH_CONTEXT_CHARS].replace("\n", " ")
        shown.append({"start": start, "end": end, "context": context})
    return {"matches": shown, "total_matches": len(spans)}


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
    # Anchored find only (issue #113): where the prefix match ends and the
    # suffix match begins, in the same offsets as ``spans``.
    anchors: tuple[int, int] | None = None


def locate(
    needle: str,
    proj: Projection,
    expected_matches: int,
    *,
    span_filter: Callable[[int, int], bool] | None = None,
) -> LocateResult:
    """Find every occurrence of *needle* in *proj*, returning codepoint spans.

    Normalization ladder (stops at first rung with >=1 IN-SCOPE match --
    see *span_filter* below; this is locate()'s own MATCH ladder, distinct
    from core/document-backend-protocol.md's unrelated 4-rung EDIT ladder
    -- see this module's header comment):
        1. exact
        2. curly/straight quote equivalence
        3. NBSP and whitespace-run collapse (stacked on rung 2)
        4. soft-hyphen (U+00AD) strip (stacked on rung 3)

    A match overlapping a RESULT-bearing field's own span (a value Word
    computed and may silently recompute later) is not a fifth rung -- it
    surfaces WARNING_TOUCHES_FIELD_RESULT in the returned warnings instead
    (see this module's header comment for why a literal RUNG_FIELD rung, as
    the issue #28 plan text names it, was not built).

    span_filter (issue #22, optional): a predicate over a raw match's own
    (start, end) codepoint span, applied to each rung's matches BEFORE
    everything downstream -- rung selection ("at least one match" now
    means "at least one IN-SCOPE match"), the STRUCTURAL_BOUNDARY check,
    and the expected_matches count. This is NOT a post-hoc filter bolted
    onto the return value: applying it any later would let an
    OUT-OF-SCOPE match raise STRUCTURAL_BOUNDARY for a boundary crossing
    the caller never asked about, or let the ladder settle on a rung
    whose only matches are out of scope while a later rung would have had
    an in-scope one. text_edit.py's row-scoped replace_text/format_text
    (`within_row_containing`) is the only caller today, filtering to
    spans inside one located table row.

    Raises VerifyError on:
        INVALID_INPUT          - empty needle
        ZERO_MATCH              - all rungs exhausted (after span_filter); near-miss in diagnostics
        MATCH_COUNT_MISMATCH    - in-scope match count != expected_matches; every in-scope span listed
        STRUCTURAL_BOUNDARY     - an in-scope match crosses a w:p/w:tbl/w:tc boundary

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

        raw_match_count = len(spans)
        if span_filter is not None:
            spans = [(s, e) for s, e in spans if span_filter(s, e)]
            if not spans:
                ladder_report.append(
                    {"rung": rung_label, "matches": 0, "matches_outside_span_filter": raw_match_count}
                )
                continue

        for s, e in spans:
            kinds = _boundary_kinds(proj, s, e)
            if kinds:
                raise _make_error(
                    ErrorCode.STRUCTURAL_BOUNDARY,
                    f"match crosses a structural boundary: {sorted(kinds)}",
                    {
                        "rung": rung_label,
                        **_match_diagnostics(proj, spans),
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
                    **_match_diagnostics(proj, spans),
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


# ---------------------------------------------------------------------------
# Anchored ("ellipsis") find and paragraph targets -- issue #113.
#
# A long span is named by its first and last words: "<prefix>…<suffix>" (or
# "[...]" as the delimiter). The prefix and suffix are located on the full
# normalization ladder, each prefix pairs with the nearest suffix that
# follows it, and the caller's expected_length (+-2 code points of the
# document's own text) or find_sha256 (of the ladder-normalized span) must
# pick out exactly one pair, so a span the caller did not mean can never
# match silently.
# ---------------------------------------------------------------------------

ELLIPSIS = "…"
ELLIPSIS_ASCII = "[...]"
ANCHOR_MIN_CHARS = 12
ANCHORED_LENGTH_TOLERANCE = 2
ANCHORED_CANDIDATE_LIMIT = 50
RUNG_ANCHORED = "anchored"
# Word's Office.js body.search() rejects needles over 255 characters; the
# server keeps every live search needle at or under this.
PANE_SEARCH_MAX = 230
_SHA256_RE = re.compile(r"[0-9a-f]{64}")


def normalize_span_text(text: str) -> str:
    """*text* with the whole match ladder applied (curly -> straight quotes,
    NBSP and whitespace runs -> one space, soft hyphens removed). This is
    the text ``find_sha256`` / ``expected_sha256`` hash, and the pane's
    ``normalizeSpanText`` (addin/taskpane.js) must stay identical to it."""
    collapsed, _ = _norm_whitespace(text.translate(_QUOTE_MAP))
    stripped, _ = _norm_softhyphen(collapsed)
    return stripped


def span_sha256(text: str) -> str:
    """sha256 hex of ``normalize_span_text(text)`` (UTF-8)."""
    return hashlib.sha256(normalize_span_text(text).encode("utf-8")).hexdigest()


def split_anchored(find: str) -> tuple[str, str] | None:
    """``(prefix, suffix)`` for an anchored *find*, ``None`` for a plain one.

    Exactly one delimiter (``…`` or ``[...]``) makes a find anchored; more
    than one, or a prefix/suffix under 12 characters after normalization,
    is INVALID_INPUT. There is no escape for a literal ellipsis.
    """
    count = find.count(ELLIPSIS) + find.count(ELLIPSIS_ASCII)
    if count == 0:
        return None
    if count > 1:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "an anchored find takes exactly one '…' or '[...]' between its prefix and suffix; "
            f"found {count}. A literal ellipsis cannot be matched by an anchored find.",
            {"delimiters": count},
        )
    delimiter = ELLIPSIS if ELLIPSIS in find else ELLIPSIS_ASCII
    prefix, suffix = find.split(delimiter)
    for name, part in (("prefix", prefix), ("suffix", suffix)):
        if len(normalize_span_text(part).strip()) < ANCHOR_MIN_CHARS:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"the anchored find's {name} must be at least {ANCHOR_MIN_CHARS} characters; got {part!r}",
                {name: part},
            )
    return prefix, suffix


def check_anchored_args(expected_length: Any, find_sha256: Any) -> str | None:
    """Validate an anchored find's check arguments; returns the lowercased digest."""
    if expected_length is None and find_sha256 is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "an anchored find ('<prefix>…<suffix>') requires expected_length (characters in the full "
            "span, +-2) or find_sha256 (sha256 of the normalized span), or both",
        )
    if expected_length is not None and (
        isinstance(expected_length, bool) or not isinstance(expected_length, int) or expected_length < 1
    ):
        raise _make_error(ErrorCode.INVALID_INPUT, f"expected_length must be a positive integer; got {expected_length!r}")
    if find_sha256 is None:
        return None
    digest = find_sha256.lower() if isinstance(find_sha256, str) else None
    if digest is None or not _SHA256_RE.fullmatch(digest):
        raise _make_error(ErrorCode.INVALID_INPUT, "find_sha256 must be 64 hex characters")
    return digest


@dataclasses.dataclass(frozen=True)
class AnchoredSpan:
    """One prefix->suffix pairing, in original-text offsets: the span
    ``[start, end)``, where its prefix match ends and its suffix match begins."""

    start: int
    end: int
    prefix_end: int
    suffix_start: int


def _ladder_map(text: str) -> tuple[str, list[int]]:
    quoted = text.translate(_QUOTE_MAP)
    collapsed, ws_pos = _norm_whitespace(quoted)
    stripped, sh_pos = _norm_softhyphen(collapsed)
    return stripped, _compose_orig_pos(sh_pos, ws_pos)


def anchored_pairs(text: str, prefix: str, suffix: str) -> tuple[list[AnchoredSpan], int, int]:
    """Every prefix match in *text* paired with the nearest suffix match that
    starts at or after the prefix's end, both located on the fully
    normalized ladder. Returns ``(pairs, prefix_count, suffix_count)``;
    identical spans are listed once and enumeration stops at 50 pairs."""
    norm, orig = _ladder_map(text)
    np_, _ = _ladder_map(prefix)
    ns, _ = _ladder_map(suffix)
    prefix_count = 0
    suffix_count = len(_find_all(ns, norm)) if ns else 0
    pairs: list[AnchoredSpan] = []
    seen: set[tuple[int, int]] = set()
    i = norm.find(np_) if np_ else -1
    while i != -1 and len(pairs) < ANCHORED_CANDIDATE_LIMIT:
        prefix_count += 1
        j = norm.find(ns, i + len(np_)) if ns else -1
        if j != -1:
            span = AnchoredSpan(orig[i], orig[j + len(ns)], orig[i + len(np_)], orig[j])
            if (span.start, span.end) not in seen:
                seen.add((span.start, span.end))
                pairs.append(span)
        i = norm.find(np_, i + 1)
    return pairs, prefix_count, suffix_count


def _span_context(text: str, start: int, end: int) -> str:
    """The span's first and last words with a little surrounding text, so two
    candidates that share a prefix still read differently."""
    head = text[max(0, start - 10) : min(end, start + 30)]
    tail = text[max(start + 30, end - 30) : end + 10]
    joined = head if not tail else f"{head}…{tail}"
    return joined.replace("\n", " ")


def _anchored_diag(text: str, spans: list[AnchoredSpan]) -> list[dict[str, Any]]:
    return [
        {"start": s.start, "end": s.end, "length": s.end - s.start, "context": _span_context(text, s.start, s.end)}
        for s in spans[:_MATCH_DIAGNOSTIC_MAX]
    ]


def _passes(text: str, span: AnchoredSpan, expected_length: int | None, digest: str | None) -> bool:
    if expected_length is not None and abs((span.end - span.start) - expected_length) > ANCHORED_LENGTH_TOLERANCE:
        return False
    return digest is None or span_sha256(text[span.start : span.end]) == digest


def _rank(candidates: list[Any], expected_length: int | None, span_of: Callable[[Any], AnchoredSpan]) -> list[Any]:
    if expected_length is None:
        return candidates
    return sorted(candidates, key=lambda c: abs((span_of(c).end - span_of(c).start) - expected_length))


def _no_anchored_match(
    text: str,
    prefix: str,
    suffix: str,
    candidates: list[AnchoredSpan],
    counts: tuple[int, int],
    expected_length: int | None,
) -> VerifyError:
    diagnostics: dict[str, Any] = {
        "prefix_matches": counts[0],
        "suffix_matches": counts[1],
        "candidates": _anchored_diag(text, _rank(candidates, expected_length, lambda s: s)),
        "total_candidates": len(candidates),
    }
    if expected_length is not None:
        diagnostics["expected_length"] = expected_length
    if counts[0] == 0 or counts[1] == 0:
        norm, _ = _ladder_map(text)
        missing = prefix if counts[0] == 0 else suffix
        diagnostics["near_miss"] = _near_miss_scan(normalize_span_text(missing), norm)
    if not candidates:
        message = "no prefix -> suffix span found for the anchored find"
    else:
        message = (
            f"none of the {len(candidates)} prefix -> suffix span(s) has the expected length/sha256; "
            "the candidates' lengths are in diagnostics"
        )
    return _make_error(ErrorCode.ZERO_MATCH, message, diagnostics)


def locate_anchored(
    find: str,
    proj: Projection,
    *,
    expected_length: int | None = None,
    find_sha256: str | None = None,
    span_filter: Callable[[int, int], bool] | None = None,
) -> LocateResult:
    """Resolve an anchored *find* (``"<prefix>…<suffix>"``) to one full span.

    Raises INVALID_INPUT (not anchored, bad delimiter count, a prefix/suffix
    under 12 characters, neither expected_length nor find_sha256),
    STRUCTURAL_BOUNDARY (the only span(s) cross a paragraph, table cell or
    text-box boundary), AMBIGUOUS_SPAN (more than one span passes the check;
    up to five ``{start, end, context}`` plus ``total``) or ZERO_MATCH (no
    span passes; the nearest candidates' lengths are in diagnostics).
    *span_filter* applies to each candidate before the check, as in locate().
    """
    parts = split_anchored(find)
    if parts is None:
        raise _make_error(ErrorCode.INVALID_INPUT, "not an anchored find: it has no '…' or '[...]' delimiter")
    digest = check_anchored_args(expected_length, find_sha256)
    prefix, suffix = parts
    text = proj.text
    pairs, prefix_count, suffix_count = anchored_pairs(text, prefix, suffix)
    if span_filter is not None:
        pairs = [p for p in pairs if span_filter(p.start, p.end)]
    inside: list[AnchoredSpan] = []
    crossing: list[tuple[AnchoredSpan, set[str]]] = []
    for pair in pairs:
        kinds = _boundary_kinds(proj, pair.start, pair.end)
        if kinds:
            crossing.append((pair, kinds))
        else:
            inside.append(pair)
    passing = [p for p in inside if _passes(text, p, expected_length, digest)]
    if len(passing) > 1:
        raise _make_error(
            ErrorCode.AMBIGUOUS_SPAN,
            f"the anchored find matches {len(passing)} spans; lengthen the prefix or suffix, or pass find_sha256",
            {"candidates": _anchored_diag(text, passing), "total": len(passing)},
        )
    if not passing:
        crossing_passing = [(p, k) for p, k in crossing if _passes(text, p, expected_length, digest)]
        reported = crossing_passing or (crossing if not inside else [])
        if reported:
            kinds_union = sorted(set().union(*(k for _p, k in reported)))
            raise _make_error(
                ErrorCode.STRUCTURAL_BOUNDARY,
                f"the anchored span crosses a structural boundary: {kinds_union}",
                {
                    "rung": RUNG_ANCHORED,
                    "candidates": _anchored_diag(text, [p for p, _k in reported]),
                    "boundary_kinds": kinds_union,
                },
            )
        raise _no_anchored_match(text, prefix, suffix, inside, (prefix_count, suffix_count), expected_length)

    span = passing[0]
    return LocateResult(
        spans=[(span.start, span.end)],
        rung=RUNG_ANCHORED,
        match_count=1,
        warnings=_span_warnings(proj, span.start, span.end),
        anchors=(span.prefix_end, span.suffix_start),
    )


def _span_warnings(proj: Projection, start: int, end: int) -> list[str]:
    warnings: set[str] = set()
    if _crosses_comment_range(proj, start, end):
        warnings.add("crosses_comment_range")
    if _crosses_revision(proj, start, end):
        warnings.add("crosses_revision")
    if _touches_field_result(proj, start, end, _field_spans(proj)):
        warnings.add(WARNING_TOUCHES_FIELD_RESULT)
    return sorted(warnings)


def check_plain_find_args(expected_matches: Any, expected_length: Any, find_sha256: Any) -> None:
    """A plain (non-anchored) find: expected_matches required, no anchored-only arguments."""
    if expected_length is not None or find_sha256 is not None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "expected_length and find_sha256 apply only to an anchored find ('<prefix>…<suffix>')",
        )
    if expected_matches is None:
        raise _make_error(ErrorCode.INVALID_INPUT, "expected_matches is required for a plain find (D4)")


def check_anchored_expected_matches(expected_matches: Any) -> None:
    if expected_matches not in (None, 1) or isinstance(expected_matches, bool):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "an anchored find always names exactly one span; omit expected_matches (or pass 1)",
        )


def locate_find(
    find: str,
    proj: Projection,
    expected_matches: int | None,
    *,
    expected_length: int | None = None,
    find_sha256: str | None = None,
    span_filter: Callable[[int, int], bool] | None = None,
) -> LocateResult:
    """``locate`` for a plain *find*, ``locate_anchored`` for an anchored one.

    A plain find requires *expected_matches* and refuses expected_length /
    find_sha256; an anchored find requires one of those two instead and
    accepts *expected_matches* only as 1.
    """
    if not find:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty")
    if split_anchored(find) is not None:
        check_anchored_expected_matches(expected_matches)
        return locate_anchored(
            find, proj, expected_length=expected_length, find_sha256=find_sha256, span_filter=span_filter
        )
    check_plain_find_args(expected_matches, expected_length, find_sha256)
    assert expected_matches is not None
    return locate(find, proj, expected_matches, span_filter=span_filter)


def paragraph_spans(proj: Projection) -> dict[str, tuple[int, int]]:
    """``para_ref -> [start, end)`` of each non-empty paragraph's own text in
    ``proj.text`` (the paragraph-separator newline excluded)."""
    spans: dict[str, tuple[int, int]] = {}
    for s, e, para_ref, run_ref in proj.offset_map:
        if run_ref.endswith("/break"):
            continue
        spans[para_ref] = (spans[para_ref][0], e) if para_ref in spans else (s, e)
    return spans


def locate_paragraph(proj: Projection, para_ref: str, expected_sha256: str) -> LocateResult:
    """The whole text of paragraph *para_ref* (``pN``, as find_sections and
    the runs/markdown reads report it), verified against *expected_sha256*:
    the sha256 of the paragraph's normalized text (``span_sha256``).

    INVALID_INPUT for a malformed/out-of-range ref, a malformed hash or an
    empty paragraph; PARA_HASH_MISMATCH when the paragraph's text changed.
    """
    from .projection import para_index

    digest = expected_sha256.lower() if isinstance(expected_sha256, str) else None
    if digest is None or not _SHA256_RE.fullmatch(digest):
        raise _make_error(ErrorCode.INVALID_INPUT, "expected_sha256 must be 64 hex characters")
    try:
        index = para_index(para_ref) if isinstance(para_ref, str) else -1
    except ValueError:
        index = -1
    if not 0 <= index < len(proj.paragraphs):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"para_ref must name a paragraph p0..p{max(len(proj.paragraphs) - 1, 0)}; got {para_ref!r}",
            {"para_ref": para_ref, "paragraph_count": len(proj.paragraphs)},
        )
    bounds = paragraph_spans(proj).get(proj.paragraphs[index].para_ref)
    if bounds is None:
        raise _make_error(ErrorCode.INVALID_INPUT, f"paragraph {para_ref} is empty; there is no text to replace")
    start, end = bounds
    text = proj.text[start:end]
    if span_sha256(text) != digest:
        raise _make_error(
            ErrorCode.PARA_HASH_MISMATCH,
            f"paragraph {para_ref}'s text no longer matches expected_sha256; re-read it and retry",
            {"para_ref": para_ref, "excerpt": text[:80]},
        )
    return LocateResult(spans=[(start, end)], rung="para_ref", match_count=1, warnings=_span_warnings(proj, start, end))


# ---------------------------------------------------------------------------
# What the pane is sent (issue #113). The pane searches with Word's own
# body.search(), which knows nothing of the ladder, so it is given literal
# document text: the full span when it fits one search (<= 230 characters),
# otherwise a literal prefix and suffix plus the span's normalized sha256.
# exact_anchored_spans() replays the pane's own pairing over the server's
# projection so a target the pane could not single out is refused here,
# before anything is sent.
# ---------------------------------------------------------------------------


def pane_anchors(span_text: str, prefix_len: int | None = None, suffix_len: int | None = None) -> tuple[str, str]:
    """Literal prefix and suffix of *span_text* for the pane's search, each at
    most 230 characters. *prefix_len*/*suffix_len* are the caller's own
    anchors' lengths in the document (an anchored find); without them
    (a paragraph target) the span is split in two."""
    length = len(span_text)
    if prefix_len is None or suffix_len is None:
        prefix_len = min(PANE_SEARCH_MAX, length // 2)
        suffix_len = min(PANE_SEARCH_MAX, length - prefix_len)
    prefix = span_text[: min(prefix_len, PANE_SEARCH_MAX)]
    suffix = span_text[length - min(suffix_len, PANE_SEARCH_MAX) :]
    if not prefix or not suffix or any(c in prefix + suffix for c in "\r\n\v"):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "the span's prefix or suffix would include a line break, which Word cannot search for; "
            "choose anchors within one line",
        )
    return prefix, suffix


def exact_anchored_spans(units: list[tuple[int, str]], prefix: str, suffix: str, digest: str) -> list[tuple[int, int]]:
    """The pane's pairing, replayed: in each ``(offset, text)`` unit (one
    paragraph or cell), every literal *prefix* match pairs with the nearest
    literal *suffix* match after it; spans whose normalized sha256 equals
    *digest* are returned as absolute ``(start, end)``."""
    hits: list[tuple[int, int]] = []
    for offset, text in units:
        for p_start in _find_all(prefix, text):
            s_start = text.find(suffix, p_start + len(prefix))
            if s_start == -1:
                continue
            end = s_start + len(suffix)
            if span_sha256(text[p_start:end]) == digest:
                hits.append((offset + p_start, offset + end))
    return hits


def paragraph_units(proj: Projection) -> list[tuple[int, str]]:
    """``(start, text)`` of every non-empty paragraph in *proj* (cells included)."""
    return [(s, proj.text[s:e]) for s, e in paragraph_spans(proj).values()]


def resolve_anchored_in_units(
    find: str, units: list[str], *, expected_length: int | None, find_sha256: str | None
) -> tuple[int, AnchoredSpan]:
    """``locate_anchored`` over independent strings (an approved run's section
    paragraphs and cells) instead of a projection: ``(unit index, span)``.
    A span never crosses a unit, so there is no STRUCTURAL_BOUNDARY here."""
    parts = split_anchored(find)
    if parts is None:
        raise _make_error(ErrorCode.INVALID_INPUT, "not an anchored find: it has no '…' or '[...]' delimiter")
    digest = check_anchored_args(expected_length, find_sha256)
    prefix, suffix = parts
    found: list[tuple[int, AnchoredSpan]] = []
    candidates: list[tuple[int, AnchoredSpan]] = []
    prefix_count = suffix_count = 0
    for index, text in enumerate(units):
        pairs, pc, sc = anchored_pairs(text, prefix, suffix)
        prefix_count += pc
        suffix_count += sc
        candidates.extend((index, p) for p in pairs)
        found.extend((index, p) for p in pairs if _passes(text, p, expected_length, digest))
    if len(found) == 1:
        return found[0]

    def describe(entries: list[tuple[int, AnchoredSpan]]) -> list[dict[str, Any]]:
        return [
            {"unit": i, "start": s.start, "end": s.end, "length": s.end - s.start,
             "context": _span_context(units[i], s.start, s.end)}
            for i, s in entries[:_MATCH_DIAGNOSTIC_MAX]
        ]

    if found:
        raise _make_error(
            ErrorCode.AMBIGUOUS_SPAN,
            f"the anchored find matches {len(found)} spans; lengthen the prefix or suffix, or pass find_sha256",
            {"candidates": describe(found), "total": len(found)},
        )
    raise _make_error(
        ErrorCode.ZERO_MATCH,
        "no prefix -> suffix span with the expected length/sha256 in the section",
        {
            "prefix_matches": prefix_count,
            "suffix_matches": suffix_count,
            "candidates": describe(_rank(candidates, expected_length, lambda c: c[1])),
            "total_candidates": len(candidates),
        },
    )
