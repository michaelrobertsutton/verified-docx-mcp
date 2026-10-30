# New for issue #28 WP-06. No GoogleDocs-MCP module maps onto this one
# wholesale: that server's replace_text/format_text (mutations.py,
# formatting.py) compile Docs API batchUpdate requests (deleteContentRange +
# insertText, or updateTextStyle) against a server-assigned index space --
# there is no run-splitting to do, because the Docs API itself owns run
# boundaries. This backend edits the OOXML w:r/w:t tree directly, so THIS
# module's own job -- splitting a run that a match's boundary falls in the
# middle of, cloning its w:rPr verbatim onto the surviving pieces, and
# inheriting the first run's w:rPr onto a freshly-inserted replacement run --
# has no Google-side analogue to lift. What IS lifted: locate.py's
# normalization ladder (from verify.py, see that module's own header), and
# the guard/atomic-write/audit SHAPE of mutations.py's markdown-mutation
# pipeline (guard -> compile -> write -> post-read -> assemble evidence ->
# audit), applied here to a per-match edit instead of a whole-body/section
# rewrite.
"""Guarded pipeline for the two text-matching mutating tools: replace_text
(delete + reinsert, rung-agnostic to the caller) and format_text (style-only,
no content change).

Both share:
  1. locate.locate() over projection.project_document_root(document_root) --
     the normalization ladder, STRUCTURAL_BOUNDARY refusal, and
     crosses_comment_range/crosses_revision warnings (this WP only WARNS;
     WP-07 adds the actual TRACKED_CHANGES_PRESENT refusal on top of this
     same detection -- see tracked_changes.py).
  2. Run splitting (_atoms_for_span + the per-tool apply function below):
     a run whose text-bearing child (almost always w:t; a w:tab/w:br/
     w:noBreakHyphen/w:softHyphen/w:sym node is exactly one character wide,
     so a match boundary can only ever land AT its edges, never inside it)
     straddles a match boundary splits into up to three pieces -- unmatched
     prefix (the original element, shortened in place), the matched middle,
     and unmatched suffix (a NEW sibling element, its w:rPr CLONED VERBATIM
     from the original) -- never touching a run entirely outside the match.
  3. The same guard/atomic-write/audit machinery mutations.py already built
     (lock/sync/revision guard, atomic_replace_docx_parts with its .jsbak
     rollback, append_audit) -- reused directly, not reimplemented.
  4. The eight standard evidence keys, plus runs_before/runs_after (those
     exact names -- the lifted audit redaction set, audit._AUDIT_REDACTED_KEYS,
     already covers them) and, when present, a non-fatal `warnings` list.

Scope limit, documented rather than silently mishandled: a run whose parent
holds it via anything other than direct list membership one level up (the
common case for every fixture in this repo) is handled generically via
element identity (parent_elem.insert/.remove), so nesting inside a
hyperlink/w:ins/w:smartTag/inline w:sdt works the same as a bare paragraph
child. What is NOT specially handled: two matches (expected_matches > 1)
whose spans are directly adjacent with no unmatched character between them,
sharing one boundary run atom -- mutating that shared atom for the first
match could invalidate the second match's own element reference. Untested
and undocumented as supported; every fixture and acceptance test in this WP
uses either a single match or matches separated by ordinary unmatched text.

Named scope limit (WP-07b-a, track_changes=True): when a match's matched
text is ITSELF already inside a still-pending (not yet accepted/rejected)
w:ins from an EARLIER track_changes=True call -- e.g. replace_text("wolf",
"coyote", track_changes=True) run right after replace_text("fox", "wolf",
track_changes=True), before "wolf" is ever accepted or rejected -- this
module wraps the matched text in a NEW w:del nested directly inside the
existing w:ins, rather than cancelling the still-pending insertion
outright (which is what mutations.py's markdown tools do for the
identical scenario at the whole-body/section level via
tracked_changes.mark_elements_deleted, and what reject_tracked_changes
already does to a bare w:ins). Read-back correctness holds regardless
(projection.py excludes w:del content unconditionally, so the superseded
text is correctly invisible either way -- verified directly, not assumed),
but the resulting <w:ins><w:del>...</w:del><w:ins>...</w:ins></w:ins>
nesting is unverified against a real Word Review pane. Consider this if a
caller chains multiple track_changes=True replace_text/format_text calls
over the same span before any accept/reject in between.
"""

from __future__ import annotations

import copy
import functools
import re
from collections.abc import Callable
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection, tables, tracked_changes
from . import locate as locate_mod
from .author import resolve_author_name
from .errors import ErrorCode, VerifyError, _make_error, short_hash
from .live import write_mode as live_write_mode
from .live.session import LiveDisconnected, LiveOpFailed, LiveStale, LiveStaleAfterWrite
from .locate import (
    LocateResult,
    locate,
    locate_anchored,
    locate_find,
    locate_paragraph,
)
from .projection import W_NS, RunEvent

_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Style allowlist (format_text) -- our own rPr model carries strike in
# addition to Google's bold/italic/underline (projection._run_properties
# already reads/writes all four), so this server's format_text supports one
# more field than the Google original. Issue #22 adds "color" (a 6-hex
# string, not a bool) -- the proposal-lead incident that opened that issue
# had no way to mark inserted text in a font color.
# ---------------------------------------------------------------------------

_STYLE_ALLOWLIST = ("bold", "italic", "underline", "strike", "color")
_BOOL_STYLE_KEYS = ("bold", "italic", "underline", "strike")
_BOOL_TOGGLE_TAGS = {"bold": "b", "italic": "i", "strike": "strike"}

_EXCERPT_RADIUS = 200

_HEX_COLOR_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")

# Full CT_RPrBase child order (ECMA-376 5th ed. Part 1 §17.3.2.1), scoped
# to THIS module rather than reusing tables._RPR_CHILD_ORDER -- that list
# was built only for insert_table's cells, which are always FRESHLY BUILT
# runs that never carry a child outside its own narrow subset (rStyle/b/
# bCs/i/iCs/color). format_text/apply_style operate on ARBITRARY
# Word-authored runs, which routinely carry w:rFonts/w:sz/w:lang/etc. that
# tables.py's own list has never had to account for -- reusing it as-is
# here would silently misplace a new w:b/w:color relative to those
# unlisted siblings (verified: this is exactly the ordering bug a plain
# `ET.SubElement(rpr, tag)` append, THIS module's own pre-issue-#22 code,
# already had for any run whose rPr carried a trailing child -- just never
# exercised, since no caller combined color/sz with a boolean toggle
# before now). w:rPrChange is pinned LAST -- it must always stay rPr's
# final child (tracked_changes.apply_rpr_change appends to it directly,
# never through this ordered-insert path, but a NEW toggle/color on a run
# that already carries one must still land BEFORE it, never after).
_RPR_CHILD_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath", "rPrChange",
]


def _validate_style(style: Any) -> dict[str, bool | str]:
    if not isinstance(style, dict):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "style must be an object mapping bold/italic/underline/strike (true/false) and/or "
            "color (a 6-hex-digit string) to a value",
            {"style": repr(style)},
        )
    if not style:
        raise _make_error(ErrorCode.INVALID_INPUT, "style must not be empty")
    unknown = sorted(set(style) - set(_STYLE_ALLOWLIST))
    if unknown:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"unknown style key(s): {unknown}; allowed: {list(_STYLE_ALLOWLIST)}",
            {"unknown_keys": unknown},
        )
    non_bool = {k: repr(v) for k in _BOOL_STYLE_KEYS if k in style for v in [style[k]] if type(v) is not bool}
    if non_bool:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"style values must be true/false booleans, got: {non_bool}",
            {"invalid_values": non_bool},
        )
    result: dict[str, bool | str] = dict(style)
    if "color" in result:
        raw = result["color"]
        match = _HEX_COLOR_RE.match(raw) if isinstance(raw, str) else None
        if match is None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f'style["color"] must be a 6-hex-digit color string (e.g. "3B3838" or "#3B3838"), got {raw!r}',
                {"color": raw},
            )
        result["color"] = match.group(1).upper()
    return result


# ---------------------------------------------------------------------------
# Run atoms: RunEvents overlapping a span, in document order.
# ---------------------------------------------------------------------------


def _atoms_for_span(proj: projection.Projection, start: int, end: int) -> list[tuple[int, int, RunEvent]]:
    run_events = [e for e in proj.events if isinstance(e, RunEvent)]
    atoms = [
        (s, e, event)
        for event, (s, e, _pr, _rr) in zip(run_events, proj.offset_map)
        if s < end and e > start
    ]
    atoms.sort(key=lambda a: a[0])
    return atoms


def _excerpt(text: str, start: int, end: int, radius: int = _EXCERPT_RADIUS) -> str:
    lo = max(0, start - radius)
    hi = min(len(text), end + radius)
    return text[lo:hi]


def _run_dict(event: RunEvent, clip_start: int, clip_end: int, node_start: int) -> dict[str, Any]:
    return {
        "text": event.text[clip_start - node_start : clip_end - node_start],
        "start": clip_start,
        "end": clip_end,
        "bold": bool(event.rpr.get("bold", False)),
        "italic": bool(event.rpr.get("italic", False)),
        "underline": bool(event.rpr.get("underline")),
        "strike": bool(event.rpr.get("strike", False)),
        "color": event.rpr.get("color"),
    }


def _collect_style_runs(proj: projection.Projection, spans: list[tuple[int, int]]) -> list[list[dict[str, Any]]]:
    """Per span, the run(s) overlapping it, clipped to the span -- the
    docx analogue of GoogleDocs-MCP verify.py's _collect_style_runs. A
    `find` straddling two runs (e.g. half already bold) surfaces both, so a
    caller sees the prior run boundaries directly rather than one flattened
    style (same rationale as the Google original)."""
    result: list[list[dict[str, Any]]] = []
    for start, end in spans:
        span_runs = [_run_dict(event, max(s, start), min(e, end), s) for s, e, event in _atoms_for_span(proj, start, end)]
        result.append(span_runs)
    return result


def _style_matches(runs_before: list[list[dict[str, Any]]], style: dict[str, bool | str]) -> bool:
    return all(
        run.get(field) == value
        for span_runs in runs_before
        for run in span_runs
        for field, value in style.items()
    )


# ---------------------------------------------------------------------------
# Element construction / style toggling
# ---------------------------------------------------------------------------


def _build_run(text_value: str, rpr_elem: Any | None) -> Any:
    r = ET.Element(_w("r"))
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = ET.SubElement(r, _w("t"))
    t.text = text_value
    if text_value == "" or text_value != text_value.strip():
        t.set(f"{{{_XML_NS}}}space", "preserve")
    return r


def _set_node_text(event: RunEvent, new_text: str) -> None:
    """Mutate the live text-bearing element's own text to *new_text*.

    Only a <w:t> node's .text is ever actually reassigned: every other
    kind this projects (w:tab/w:br/w:noBreakHyphen/w:softHyphen/w:sym) is
    exactly one character wide, so a caller's own before/after slicing can
    only ever produce "" (route to removal instead, never here) or that
    node's own full original text (a no-op) for one of those -- there is
    no partial-character case to handle.
    """
    if event.text_elem is None:
        return
    if projection._ln(event.text_elem) == "t":
        event.text_elem.text = new_text
        if new_text == "" or new_text != new_text.strip():
            event.text_elem.set(f"{{{_XML_NS}}}space", "preserve")


def _own_rpr(r_elem: Any | None) -> Any | None:
    """The <w:rPr> child directly under *r_elem*, or None if absent/r_elem
    is None. Recomputed fresh wherever needed (never cached across a
    mutation) since _toggle_style/apply_rpr_change can add one where none
    existed."""
    if r_elem is None:
        return None
    for child in r_elem:
        if projection._ln(child) == "rPr":
            return child
    return None


def _index_of(parent: Any, elem: Any) -> int | None:
    for i, child in enumerate(parent):
        if child is elem:
            return i
    return None


def _toggle_style(r_elem: Any, style: dict[str, bool | str]) -> None:
    """Apply every field in *style* to r_elem's own w:rPr (creating one if
    absent). Issue #22: every insertion now goes through
    tables._insert_ordered against this module's own full _RPR_CHILD_ORDER
    (see that list's own comment) -- fixes a latent ordering bug this
    function already had for b/i/strike/u against a run whose rPr carries
    a trailing child outside the four booleans (e.g. sz, or now color),
    not just a new bug color would have introduced on its own."""
    rpr = None
    for child in r_elem:
        if projection._ln(child) == "rPr":
            rpr = child
            break
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        r_elem.insert(0, rpr)

    for field, value in style.items():
        if field == "color":
            existing = None
            for child in rpr:
                if projection._ln(child) == "color":
                    existing = child
                    break
            if existing is None:
                existing = ET.Element(_w("color"))
                tables._insert_ordered(rpr, existing, _RPR_CHILD_ORDER)
            existing.set(_w("val"), value)
            # Issue #22 (Codex review): a themeColor-based color takes
            # visual precedence over w:val in Word's own rendering -- an
            # explicit RGB set here without clearing the theme attributes
            # would silently keep displaying the OLD theme color, making
            # this write look applied (val is correct) while Word shows
            # something else. Clear all three whenever an explicit val is
            # set.
            for theme_attr in ("themeColor", "themeTint", "themeShade"):
                attr_name = _w(theme_attr)
                if attr_name in existing.attrib:
                    del existing.attrib[attr_name]
            continue

        if field == "underline":
            existing = None
            for child in rpr:
                if projection._ln(child) == "u":
                    existing = child
                    break
            if value:
                if existing is None:
                    new_u = ET.Element(_w("u"), {_w("val"): "single"})
                    tables._insert_ordered(rpr, new_u, _RPR_CHILD_ORDER)
                else:
                    existing.set(_w("val"), "single")
            elif existing is not None:
                rpr.remove(existing)
            continue

        tag = _BOOL_TOGGLE_TAGS[field]
        existing = None
        for child in rpr:
            if projection._ln(child) == tag:
                existing = child
                break
        if value:
            if existing is None:
                new_child = ET.Element(_w(tag))
                tables._insert_ordered(rpr, new_child, _RPR_CHILD_ORDER)
            else:
                for key in list(existing.attrib):
                    del existing.attrib[key]
        elif existing is not None:
            rpr.remove(existing)


def _apply_style_to_run(r_elem: Any, style: dict[str, bool | str], *, track: _TrackContext | None) -> None:
    """_toggle_style, plus (when tracking) recording the PRE-change rPr in
    a w:rPrChange (issue #28 plan WP-07b-a: "formatting changes use
    w:rPrChange"). The snapshot is taken BEFORE _toggle_style mutates
    anything, per apply_rpr_change's own contract.

    Issue #7: if this run already carries an rPrChange from *track.author*
    (a still-pending change from an earlier track_changes=True call on
    this same run, before any accept/reject), this is a no-op on the
    change record itself -- only the live style is updated, exactly as
    real Word does (verified: tests/fixtures/revision/
    pstyle-tracked-twice.docx, the paragraph-style analogue of this same
    behavior). A FOREIGN-authored existing rPrChange is refused before
    this function is ever called (_check_foreign_rpr_change), UNLESS
    force=True let it through -- in that case apply_rpr_change's own
    remove-existing-record safety net replaces it rather than producing
    invalid, doubly-nested XML."""
    old_rpr = _own_rpr(r_elem)
    existing_author = tracked_changes.existing_change_author(old_rpr, "rPrChange") if track else None
    reuse = bool(track) and existing_author == track.author
    old_rpr_snapshot = copy.deepcopy(old_rpr) if (track and not reuse and old_rpr is not None) else None
    _toggle_style(r_elem, style)
    if track and not reuse:
        new_rpr = _own_rpr(r_elem)
        rid = track.next_id()
        tracked_changes.apply_rpr_change(new_rpr, old_rpr_elem=old_rpr_snapshot, rid=rid, author=track.author, date=track.date)


# Tracking context (tracked_changes.TrackContext) -- bundles the pieces
# track_changes=True needs (author, date, id allocator, ids created so
# far) so every call site below takes one optional argument instead of
# four. Lives in tracked_changes.py, not here, so mutations.py's own
# markdown-mutation tools (also WP-07b-a) can share the identical class
# without this module and mutations.py importing each other at module
# level (mutations.py already imports tracked_changes only inside a
# function body, the same deferred-import pattern _guard_before_write
# uses for server.py, to avoid that exact cycle).
_TrackContext = tracked_changes.TrackContext


# ---------------------------------------------------------------------------
# replace_text: run splitting + single-replacement-run insertion
# ---------------------------------------------------------------------------


def _apply_replace_span(
    start: int,
    end: int,
    replace_value: str,
    atoms: list[tuple[int, int, RunEvent]],
    *,
    track: _TrackContext | None = None,
) -> None:
    """Delete [start, end)'s text from the live tree and insert exactly one
    new <w:r> carrying *replace_value*, its w:rPr cloned verbatim from the
    FIRST overlapping run's ORIGINAL (pre-mutation) w:rPr (issue #28 plan
    WP-06: "replacement text inherits the first run's w:rPr").

    When *track* is given (WP-07b-a, track_changes=True): every run's own
    ORIGINAL characters this call would otherwise discard survive instead,
    wrapped in a fresh w:del (w:t renamed to w:delText, per OOXML) at
    exactly the position they used to occupy; the replacement text is
    wrapped in a fresh w:ins instead of a bare w:r. Untouched (before/
    after) text is not wrapped in anything -- only what actually changed
    is. Each wrapper gets its own newly-allocated w:id/author/date from
    *track*; reading the projection continues to exclude w:del text and
    include w:ins text (unchanged from WP-03/WP-06), so a tracked write is
    visible to the next read as current text, per the plan text.
    """
    if not atoms:
        raise _make_error(ErrorCode.INVALID_INPUT, "match span has no overlapping run content", {"start": start, "end": end})

    _first_s, _first_e, first_event = atoms[0]
    first_rpr = _own_rpr(first_event.r_elem)
    cloned_rpr_for_replacement = copy.deepcopy(first_rpr) if first_rpr is not None else None

    if track:
        rid = track.next_id()
        new_run = tracked_changes.wrap_insertion(
            _build_run(replace_value, cloned_rpr_for_replacement), rid=rid, author=track.author, date=track.date
        )
    else:
        new_run = _build_run(replace_value, cloned_rpr_for_replacement)

    replacement_inserted = False

    # Every atom overlapping [start, end) is visited once, in document
    # order (first -> ... -> last); a MIDDLE atom (not first, not last) is
    # always fully inside [start, end) -- see this module's docstring /
    # locate.py's structural-boundary guarantee (every atom here shares
    # one paragraph) -- so before_text/after_text are both empty for one.
    # The replacement run is inserted exactly once, at the FIRST atom's
    # own split point (issue #28 plan: "replacement text inherits the
    # first run's w:rPr" -- this is also WHERE it lands).
    for s, e, event in atoms:
        r_elem = event.r_elem
        parent = event.parent_elem
        if r_elem is None or parent is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to edit", {"start": start, "end": end})

        local_start = max(start, s) - s
        local_end = min(end, e) - s
        full_text = event.text
        before_text = full_text[:local_start]
        middle_text = full_text[local_start:local_end]
        after_text = full_text[local_end:]
        own_rpr = _own_rpr(r_elem)

        idx = _index_of(parent, r_elem)
        if idx is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run is no longer attached to its parent", {"start": start, "end": end})

        if before_text:
            # r_elem stays in place, shortened to its own unmatched prefix;
            # everything else is inserted AFTER it.
            _set_node_text(event, before_text)
            insert_at = idx + 1
        else:
            # No unmatched prefix survives on this atom -- it is either
            # fully removed (nothing here to insert instead of it, other
            # than what the middle-handling below produces) or, if there
            # is no middle either, defensively skipped (see this
            # function's own docstring: cannot occur given the overlap
            # guarantee, but insert_at must still be defined).
            parent.remove(r_elem)
            insert_at = idx

        if middle_text and track:
            rid = track.next_id()
            middle_run = _build_run(middle_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            tracked_changes.convert_t_to_deltext(middle_run)
            del_elem = tracked_changes.wrap_deletion(middle_run, rid=rid, author=track.author, date=track.date)
            parent.insert(insert_at, del_elem)
            insert_at += 1
        # not tracked: middle_text is simply omitted -- nothing survives.

        if not replacement_inserted:
            parent.insert(insert_at, new_run)
            replacement_inserted = True
            insert_at += 1

        if after_text:
            after_run = _build_run(after_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            parent.insert(insert_at, after_run)

    if not replacement_inserted:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "internal error: the replacement run was never inserted",
            {"start": start, "end": end},
        )


def _remove_if_present(parent: Any, elem: Any) -> None:
    for child in list(parent):
        if child is elem:
            parent.remove(elem)
            return


# ---------------------------------------------------------------------------
# format_text: run splitting + rPr toggling (no content change)
# ---------------------------------------------------------------------------


def _apply_format_span(
    start: int,
    end: int,
    style: dict[str, bool | str],
    atoms: list[tuple[int, int, RunEvent]],
    *,
    track: _TrackContext | None = None,
) -> None:
    """Same run-splitting rule as _apply_replace_span for a boundary run,
    except the matched middle SURVIVES here (as a new run carrying the
    requested style) instead of being removed -- format_text never
    changes character counts. When *track* is given, every run whose
    style actually changes (whole-atom or split middle) gets a
    w:rPrChange recording its PRE-change rPr (_apply_style_to_run)."""
    for s, e, event in atoms:
        local_start = max(start, s) - s
        local_end = min(end, e) - s
        full_text = event.text
        before_text = full_text[:local_start]
        after_text = full_text[local_end:]

        if not before_text and not after_text:
            # Entirely inside the span -- toggle in place, no split.
            if event.r_elem is not None:
                _apply_style_to_run(event.r_elem, style, track=track)
            continue

        if event.r_elem is None or event.parent_elem is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to edit", {"start": start, "end": end})

        original_rpr = _own_rpr(event.r_elem)
        middle_text = full_text[local_start:local_end]
        middle_run = _build_run(middle_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
        _apply_style_to_run(middle_run, style, track=track)

        parent = event.parent_elem
        idx = list(parent).index(event.r_elem)
        if before_text and after_text:
            _set_node_text(event, before_text)
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx + 1, middle_run)
            parent.insert(idx + 2, after_run)
        elif before_text:
            _set_node_text(event, before_text)
            parent.insert(idx + 1, middle_run)
        else:  # after_text only
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx, middle_run)
            parent.insert(idx + 1, after_run)
            _remove_if_present(parent, event.r_elem)


# ---------------------------------------------------------------------------
# Shared write core
# ---------------------------------------------------------------------------


def _serialize_and_write(resolved: Path, document_root: Any, raw_xml: bytes, *, post_verify) -> dict[str, Any]:
    document_decls = mutations._capture_source_namespaces(raw_xml)
    new_xml_bytes = mutations._serialize_xml(document_root, document_decls)
    overrides = {projection.DEFAULT_PART: new_xml_bytes}
    return mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=post_verify)


def _evidence(
    *,
    applied: bool,
    match_count: int,
    rung: str,
    before: str,
    after: str,
    revision_before: str,
    revision_after: str,
    audit_logged: bool,
    runs_before: list[list[dict[str, Any]]],
    runs_after: list[list[dict[str, Any]]],
    warnings: list[str],
    track: _TrackContext | None = None,
    conflict_sweep: dict[str, Any] | None = None,
) -> dict[str, Any]:
    evidence: dict[str, Any] = {
        "applied": applied,
        "match_count": match_count,
        "rung": rung,
        "before": before,
        "after": after,
        "revision_before": revision_before,
        "revision_after": revision_after,
        "audit_logged": audit_logged,
        "runs_before": runs_before,
        "runs_after": runs_after,
        # issue #154: this builder is only ever used for the file-mode
        # path (live goes through live_write_mode.live_evidence, which
        # sets this same key to "live") -- see that module's own
        # write_mode key.
        "write_mode": "file",
    }
    if warnings:
        evidence["warnings"] = warnings
    if track is not None:
        # issue #28 plan WP-07b-a: "The evidence envelope gains
        # revision_ids (the ids created) and track_changes: true."
        evidence["track_changes"] = True
        evidence["revision_ids"] = track.revision_ids
    if conflict_sweep is not None:
        # issue #28 WP-10: layer 3's post-write conflict-copy sweep,
        # merged the same way mutations.py's own _evidence does.
        mutations._merge_conflict_sweep(evidence, conflict_sweep)
    return evidence


def _check_tracked_changes_guard(
    proj: projection.Projection, locate_result: LocateResult, force: bool, own_author: str
) -> None:
    """New for issue #28 WP-07: a match crossing a tracked change (w:ins/
    w:del) refuses unless force=True. WP-06 only detected and warned; this
    is the actual refusal, layered on top of that same detection
    (tracked_changes.py's accept_tracked_changes/reject_tracked_changes
    clear the way for a subsequent write). crosses_comment_range stays a
    warning-only signal at this WP -- WP-07's own text scopes the refusal
    to tracked changes.

    WP-07b-a's own-author exclusion (load-bearing per that WP's notes):
    a revision authored by *own_author* (the same identity this server
    itself stamps on a track_changes=True write, resolved via
    author.resolve_author_name()) does NOT trigger this refusal, or a
    second proposed edit over the server's own prior tracked change would
    deadlock -- accept/reject would be the only way forward even for
    edits nobody but this server ever proposed. A revision authored by
    anyone else still refuses normally. This checks EACH matched span
    directly via tracked_changes.foreign_crosses_revision (not the
    coarser "crosses_revision" warning already on locate_result, which
    does not distinguish authors) -- both remain available in the
    evidence: warnings still reports crosses_revision for ANY author, this
    guard only ever refuses for a FOREIGN one.
    """
    if force:
        return
    if "crosses_revision" not in locate_result.warnings:
        return
    for start, end in locate_result.spans:
        if tracked_changes.foreign_crosses_revision(proj, start, end, own_author):
            raise _make_error(
                ErrorCode.TRACKED_CHANGES_PRESENT,
                "The matched text overlaps a tracked change (w:ins/w:del) authored by someone other than "
                f"{own_author!r}; accept or reject it (tracked_changes.accept_tracked_changes / "
                "reject_tracked_changes) first, or pass force=True.",
                {
                    "spans": [{"start": s, "end": e} for s, e in locate_result.spans],
                    "warnings": locate_result.warnings,
                    "own_author": own_author,
                },
            )


def _check_foreign_rpr_change(proj: projection.Projection, spans: list[tuple[int, int]], force: bool, own_author: str) -> None:
    """New for issue #7: a match whose run already carries a w:rPrChange
    (a pending FORMATTING change, distinct from _check_tracked_changes_guard's
    w:ins/w:del check above) authored by someone other than *own_author*
    refuses TRACKED_CHANGES_PRESENT unless force=True.

    Unlike the w:ins/w:del case, this is NOT mirroring a rule Word itself
    enforces: verified against a real Word-authored fixture
    (tests/fixtures/revision/pstyle-tracked-foreign.docx -- the paragraph-
    style analogue, same underlying rPrChange/pPrChange model), Word's own
    desktop UI silently updates the live property and leaves the FIRST
    author's change record completely untouched even when a second Word
    session under a different user name edits the same run/paragraph
    before any accept/reject. This check is a deliberate safety policy
    this tool adds on top of that -- an automated caller should not
    silently touch another author's still-pending formatting change
    without at least an explicit force=True acknowledgment, even though
    Word's own UI would let it."""
    if force:
        return
    for start, end in spans:
        for _s, _e, event in _atoms_for_span(proj, start, end):
            existing_author = tracked_changes.existing_change_author(_own_rpr(event.r_elem), "rPrChange")
            if existing_author is not None and existing_author != own_author:
                raise _make_error(
                    ErrorCode.TRACKED_CHANGES_PRESENT,
                    "The matched text's run already carries a w:rPrChange (a pending formatting "
                    f"change) authored by someone other than {own_author!r}; accept or reject it "
                    "first, or pass force=True.",
                    {"start": start, "end": end, "existing_author": existing_author, "own_author": own_author},
                )


def _check_foreign_ppr_change(paragraphs: list[Any], force: bool, own_author: str) -> None:
    """New for issue #7: the paragraph-level analogue of
    _check_foreign_rpr_change (see that function's docstring for the
    verified Word behavior this policy is deliberately stricter than). A
    paragraph whose w:pPr already carries a w:pPrChange authored by
    someone other than *own_author* refuses TRACKED_CHANGES_PRESENT
    unless force=True."""
    if force:
        return
    for p_elem in paragraphs:
        ppr = None
        for child in p_elem:
            if projection._ln(child) == "pPr":
                ppr = child
                break
        existing_author = tracked_changes.existing_change_author(ppr, "pPrChange")
        if existing_author is not None and existing_author != own_author:
            raise _make_error(
                ErrorCode.TRACKED_CHANGES_PRESENT,
                "A paragraph in the matched range already carries a w:pPrChange (a pending "
                f"paragraph-style change) authored by someone other than {own_author!r}; accept or "
                "reject it first, or pass force=True.",
                {"existing_author": existing_author, "own_author": own_author},
            )


# ---------------------------------------------------------------------------
# Live mode (issue #106 WP-3): replace_text/format_text over the WSS ops
# channel instead of the OOXML file, when write_mode resolves to "live"
# (live/write_mode.py's resolve_write_mode -- the shared rule WP-4's
# comment tools also use). Structural verification against the SAVED
# file (rungs 3/4 never go live; a re-read via the existing read tools
# after live_save is the caller's own job -- these two functions never
# touch the .docx on disk at all).
# ---------------------------------------------------------------------------


def _live_describe(session, path: str, revision_before: str | None) -> str:
    """describe() first (record the pane's pre-op body hash), then the
    LIVE_STALE pre-flight check against *revision_before* -- refuses
    BEFORE any mutating op is sent, per the plan."""
    describe_result = session.request_threadsafe("describe")
    pre_hash = describe_result["bodySha256"]
    live_write_mode.check_not_stale(revision_before, pre_hash)
    return pre_hash


_ROW_SCOPE_CAPABILITY = "row_scope"


_COMMENT_LOSS_IDS_CAPABILITY = "comment_loss_ids"


def validate_comment_loss_ids(value: Any) -> list[dict[str, Any]] | None:
    """Validate ``allow_comment_loss_ids`` (live-only, per-thread comment-loss consent).

    ``None`` means "not used"; a list (even empty) switches the pane's comment guard to
    ids mode: each entry ``{"comment_id": str, "reply_count": int}`` names one RESOLVED
    thread the caller accepts losing, with the reply count it saw. Returns a normalized copy.
    """
    if value is None:
        return None
    if not isinstance(value, list):
        raise _make_error(ErrorCode.INVALID_INPUT, "allow_comment_loss_ids must be a list of {comment_id, reply_count}")
    out: list[dict[str, Any]] = []
    for i, entry in enumerate(value):
        if not isinstance(entry, dict) or set(entry) != {"comment_id", "reply_count"}:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"allow_comment_loss_ids[{i}] must be exactly {{comment_id, reply_count}}",
            )
        cid, count = entry["comment_id"], entry["reply_count"]
        if not isinstance(cid, str) or not cid:
            raise _make_error(ErrorCode.INVALID_INPUT, f"allow_comment_loss_ids[{i}].comment_id must be a non-empty string")
        if isinstance(count, bool) or not isinstance(count, int) or count < 0:
            raise _make_error(
                ErrorCode.INVALID_INPUT, f"allow_comment_loss_ids[{i}].reply_count must be a non-negative integer"
            )
        out.append({"comment_id": cid, "reply_count": count})
    return out


def _require_live_replace_capabilities(
    session, *, within_row_containing: str | None, allow_comment_loss_ids: list[dict[str, Any]] | None
) -> None:
    live_write_mode.require_capability(session, "replacement_formatting", feature_description="replacement formatting")
    live_write_mode.require_capability(session, "comment_loss_guard", feature_description="comment loss protection")
    if within_row_containing:
        live_write_mode.require_capability(
            session, _ROW_SCOPE_CAPABILITY, feature_description="within_row_containing"
        )
    if allow_comment_loss_ids is not None:
        live_write_mode.require_capability(
            session, _COMMENT_LOSS_IDS_CAPABILITY, feature_description="allow_comment_loss_ids"
        )


# ---------------------------------------------------------------------------
# Write targets (issue #113): a plain find, an anchored find
# ("<prefix>…<suffix>" + expected_length / find_sha256), or a whole
# paragraph (para_ref + expected_sha256).
#
# File mode resolves the target on the file's own projection. Live mode
# resolves it on a projection of the live body read through the
# NON-OBSERVING session live_session_for returns, so the read never re-arms
# the caller's shared-bridge baseline: revision_before, the broker's
# baseline and section-lock checks still decide at the mutation. The pane
# then gets the resolved text itself (<= 230 characters, one Word search,
# expected_matches=1) or a literal prefix + suffix + the span's normalized
# sha256 (anchored_find_v1), which it re-checks at the mutation boundary.
# ---------------------------------------------------------------------------

ANCHORED_FIND_CAPABILITY = "anchored_find_v1"


def check_target_args(
    find: str | None,
    expected_matches: Any,
    *,
    expected_length: Any = None,
    find_sha256: Any = None,
    para_ref: Any = None,
    expected_sha256: Any = None,
    allow_para_ref: bool = True,
) -> str:
    """Validate which target a call names, before any I/O: ``"plain"``,
    ``"anchored"`` or ``"para_ref"``. INVALID_INPUT for any mix the
    contract does not allow."""
    if para_ref is not None or expected_sha256 is not None:
        if not allow_para_ref:
            raise _make_error(ErrorCode.INVALID_INPUT, "para_ref is accepted by replace_text only")
        if find:
            raise _make_error(ErrorCode.INVALID_INPUT, "pass find or para_ref, not both")
        if para_ref is None or expected_sha256 is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "para_ref and expected_sha256 go together")
        if expected_length is not None or find_sha256 is not None:
            raise _make_error(ErrorCode.INVALID_INPUT, "expected_length/find_sha256 apply to an anchored find, not para_ref")
        if expected_matches not in (None, 1) or isinstance(expected_matches, bool):
            raise _make_error(ErrorCode.INVALID_INPUT, "a para_ref target is one paragraph; omit expected_matches")
        return "para_ref"
    if not isinstance(find, str) or not find:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty (or pass para_ref + expected_sha256)")
    if locate_mod.anchored_requested(expected_length, find_sha256):
        # Only expected_length/find_sha256 make a find anchored; a literal "…" is plain text otherwise.
        locate_mod.require_anchored(find)
        locate_mod.check_anchored_expected_matches(expected_matches)
        locate_mod.check_anchored_args(expected_length, find_sha256)
        return "anchored"
    locate_mod.check_plain_find_args(expected_matches, expected_length, find_sha256)
    return "plain"


def _locate_file_target(
    proj: projection.Projection,
    kind: str,
    find: str | None,
    expected_matches: int | None,
    *,
    expected_length: int | None,
    find_sha256: str | None,
    para_ref: str | None,
    expected_sha256: str | None,
    span_filter: Callable[[int, int], bool] | None = None,
) -> LocateResult:
    if kind == "para_ref":
        assert para_ref is not None and expected_sha256 is not None
        return locate_paragraph(proj, para_ref, expected_sha256)
    assert find is not None
    return locate_find(
        find, proj, expected_matches, expected_length=expected_length, find_sha256=find_sha256, span_filter=span_filter
    )


def _check_live_target_scope(kind: str, *, scope: str, within_row_containing: str | None) -> None:
    if kind == "plain":
        return
    if scope != "body" or within_row_containing:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "in live mode an anchored find or para_ref target works on the document body only "
            "(no text-box scope, no within_row_containing)",
        )


def _ambiguous_text(text: str, positions: list[tuple[int, int]], what: str) -> VerifyError:
    return _make_error(
        ErrorCode.AMBIGUOUS_SPAN,
        f"{what} occurs {len(positions)} times in the live document; Word cannot tell the copies apart",
        {
            "candidates": [
                {"start": s, "end": e, "context": text[max(0, s - 20) : s + 20].replace("\n", " ")}
                for s, e in positions[:5]
            ],
            "total": len(positions),
        },
    )


def pane_target(session, proj: projection.Projection, result: LocateResult) -> dict[str, Any]:
    """What the pane is sent for a resolved span: ``{"text", "find"}`` when the
    span fits one Word search and occurs once, else ``{"text", "prefix",
    "suffix", "expected_text_sha256"}`` (requires ``anchored_find_v1``).
    AMBIGUOUS_SPAN when Word could not single the span out."""
    start, end = result.spans[0]
    text = proj.text[start:end]
    if locate_mod.utf16_len(text) <= locate_mod.PANE_SEARCH_MAX and not any(c in text for c in "\r\n\v"):
        positions = [(p, p + len(text)) for p in locate_mod._find_all(text, proj.text)]
        if len(positions) != 1:
            raise _ambiguous_text(proj.text, positions, "the target text")
        return {"text": text, "find": text}
    if session is not None:
        live_write_mode.require_capability(
            session,
            ANCHORED_FIND_CAPABILITY,
            feature_description=f"a {len(text)}-character target (longer than one Word search)",
        )
    prefix, suffix = locate_mod.pane_anchors(text, result.anchors[1] - start if result.anchors else None)
    digest = locate_mod.span_sha256(text)
    locate_mod.check_pane_target(locate_mod.paragraph_units(proj), prefix, suffix, digest, (start, end))
    return {"text": text, "prefix": prefix, "suffix": suffix, "expected_text_sha256": digest}


def resolve_live_target(
    session,
    pre_hash: str,
    kind: str,
    *,
    find: str | None,
    expected_length: int | None = None,
    find_sha256: str | None = None,
    para_ref: str | None = None,
    expected_sha256: str | None = None,
) -> dict[str, Any]:
    """Resolve an anchored find or para_ref against the live body (a
    non-observing read, checked against *pre_hash*) and return ``pane_target``."""
    from .live.reads_live import live_projection

    proj = live_projection(session, expected_body_sha256=pre_hash)
    if kind == "para_ref":
        assert para_ref is not None and expected_sha256 is not None
        result = locate_paragraph(proj, para_ref, expected_sha256)
    else:
        assert find is not None
        result = locate_anchored(find, proj, expected_length=expected_length, find_sha256=find_sha256)
    return pane_target(session, proj, result)


def _apply_pane_target(payload: dict[str, Any], target: dict[str, Any] | None) -> None:
    if target is None or "find" in target:
        return
    payload.pop("find", None)
    payload["expected_matches"] = 1
    for key in ("prefix", "suffix", "expected_text_sha256"):
        payload[key] = target[key]


def _live_replace_once(
    session,
    *,
    find: str,
    replace: str,
    expected_matches: int,
    pre_hash: str,
    target: dict[str, Any] | None = None,
    track_changes: bool = False,
    allow_comment_loss: bool = False,
    allow_comment_loss_ids: list[dict[str, Any]] | None = None,
    within_row_containing: str | None = None,
    inherit_format: str = "replaced",
    replacement_segments: list[dict[str, Any]] | None = None,
    scope: str = "body",
    scope_info: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Send ONE pane ``replace`` guarded by ``expectedBodySha256=pre_hash`` and verify it.

    Shared by ``execute_replace_text_live`` and ``execute_replace_text_batch_live``.
    Returns the pane's verified result (``post`` is the new body hash). Raises a
    ``VerifyError``: a pane-side refusal keeps its own code (nothing written), while
    ``LIVE_STALE_AFTER_WRITE`` (the pane applied the op against a body that had moved) and
    ``VERIFICATION_FAILED`` mean the write may have landed. Never retries and never audits;
    the caller owns both.
    """
    payload: dict[str, Any] = {
        "find": find,
        "expected_matches": expected_matches,
        "replace": replace,
        "track_changes": track_changes,
        "allow_comment_loss": allow_comment_loss,
        "inherit_format": inherit_format,
    }
    if replacement_segments is not None:
        payload["replacement_segments"] = replacement_segments
    if allow_comment_loss_ids is not None:
        payload["allow_comment_loss_ids"] = allow_comment_loss_ids
    payload["scope"] = scope
    payload["expectedBodySha256"] = pre_hash
    if scope_info:
        payload["expectedScopeSha256"] = scope_info["scopeSha256"]
    if within_row_containing:
        payload["rowAnchor"] = within_row_containing
    _apply_pane_target(payload, target)
    try:
        result = session.request_threadsafe("replace", payload, expected_body_sha256=pre_hash)
    except LiveStaleAfterWrite as exc:
        raise _make_error(
            ErrorCode.LIVE_STALE_AFTER_WRITE,
            "the pane applied the replace, but its body hash before the op differed from the one "
            "this call expected: the document moved between the read and the write. The write "
            "ALREADY LANDED -- do not retry; re-read the live document and verify it.",
            {"expected": short_hash(exc.expected), "actual": short_hash(exc.actual)},
        ) from exc
    except LiveStale as exc:
        raise _make_error(
            ErrorCode.LIVE_STALE,
            "The live document changed since it was read; nothing was written. Call live_revision or "
            "re-read, then retry.",
            {"expected": short_hash(exc.expected), "actual": short_hash(exc.actual)},
        ) from exc
    except LiveOpFailed as exc:
        raise _make_error(
            live_write_mode.classify_op_failed(exc), exc.message, {"pane_code": exc.code}
        ) from exc
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    matches = result.get("matches") or []
    post_hash = result.get("post")

    if not result.get("applied") or (replacement_segments is None and (result.get("scope_post") == scope_info["scopeSha256"] if scope_info else post_hash == pre_hash)):
        raise _make_error(
            ErrorCode.VERIFICATION_FAILED,
            "live replace did not verify: the pane's post-op body hash did not change from its "
            "pre-op hash (or the pane did not report applied=true). Nothing to roll back in live "
            "mode -- Word, not this server, owns the document; re-read via describe/read_document "
            "and retry.",
            {"pre": pre_hash, "post": post_hash, "applied": result.get("applied")},
        )
    for m in matches:
        if m.get("after") != replace:
            raise _make_error(
                ErrorCode.VERIFICATION_FAILED,
                f"live replace did not verify: a matched range's after-text {m.get('after')!r} "
                f"does not equal the requested replacement {replace!r}. Nothing to roll back in "
                "live mode -- Word, not this server, owns the document.",
                {"matches": matches},
            )
    return result


def _audit_stale_after_write(exc: VerifyError, *, tool: str, path: str, document_name: str, pre_hash: str) -> None:
    """The write landed but its read-back could not complete: leave an audit trail."""
    live_write_mode.audit_live_failure(
        tool=tool,
        path=path,
        document_name=document_name,
        pre_body_sha256=pre_hash,
        post_body_sha256=None,
        reason=exc.envelope.message,
        detail=exc.envelope.diagnostics,
    )


def execute_replace_text_live(
    path: str,
    find: str | None,
    replace: str,
    expected_matches: int | None,
    *,
    revision_before: str | None = None,
    track_changes: bool = False,
    within_row_containing: str | None = None,
    allow_comment_loss: bool = False,
    allow_comment_loss_ids: list[dict[str, Any]] | None = None,
    inherit_format: str = "replaced",
    replacement_segments: list[dict[str, Any]] | None = None,
    scope: str = "body",
    expected_length: int | None = None,
    find_sha256: str | None = None,
    para_ref: str | None = None,
    expected_sha256: str | None = None,
) -> dict[str, Any]:
    """Live-mode ``replace_text`` (issue #106 WP-3): sends a ``replace``
    op over the pane's WSS ops channel instead of editing the .docx file.

    Flow: ``describe`` (record ``pre`` body hash + ``LIVE_STALE`` check)
    -> ``replace`` (with ``expected_body_sha256=pre`` so the session layer
    itself also re-checks staleness against the reply; see
    ``_live_replace_once``) -> verify (``post`` differs from ``pre``, and
    every match's ``after`` equals *replace* exactly) -> build the live
    evidence envelope (``live/write_mode.py``'s ``live_evidence``).

    ``force`` has no live-mode analogue (there is no tracked-change/
    comment-anchor override to force here -- Word owns the document, not
    this server) and is not accepted; a caller that wants a forced
    override uses ``write_mode="file"`` with the file's own ``force``
    instead.

    ``within_row_containing`` (issue #22 B2): sent as the ``replace`` op's
    ``rowAnchor`` field -- requires the connected pane to report the
    ``"row_scope"`` capability in its own ``hello``
    (``live_write_mode.require_capability``, checked BEFORE this op is
    sent), or ``LIVE_CAPABILITY_MISSING``. See ``text_edit.py``'s own
    ``_resolve_row_span_filter`` docstring (the file-mode equivalent) for
    the anchor-uniqueness contract; the pane resolves the SAME contract
    against the live document instead of a projection.

    ``allow_comment_loss_ids`` (issue #82): per-thread comment-loss consent,
    sent to the pane (``"comment_loss_ids"`` capability); see
    ``validate_comment_loss_ids``.

    Raises ``LIVE_UNAVAILABLE``/``LIVE_DISCONNECTED``/``LIVE_STALE`` (see
    ``live/session.py``'s exceptions of the same names), ``ZERO_MATCH``/
    ``MATCH_COUNT_MISMATCH``/``LIVE_OP_FAILED`` (the pane refused the
    ``expected_matches`` gate -- see ``live/write_mode.py``'s
    ``classify_op_failed``), ``LIVE_CAPABILITY_MISSING`` (see above),
    ``LIVE_STALE_AFTER_WRITE`` (the write landed against a body that had
    moved; audited), or ``VERIFICATION_FAILED`` (the pane's own read-back
    after the op did not confirm the intended change -- nothing to roll
    back in live mode, since Word, not this server, owns the file; the
    diagnostics say so explicitly).
    """
    if inherit_format not in ("replaced", "previous", "none"):
        raise _make_error(ErrorCode.INVALID_INPUT, "Invalid inherit_format")
    kind = check_target_args(
        find, expected_matches, expected_length=expected_length, find_sha256=find_sha256,
        para_ref=para_ref, expected_sha256=expected_sha256,
    )
    _check_live_target_scope(kind, scope=scope, within_row_containing=within_row_containing)
    allow_comment_loss_ids = validate_comment_loss_ids(allow_comment_loss_ids)
    if allow_comment_loss and allow_comment_loss_ids is not None:
        raise _make_error(
            ErrorCode.INVALID_INPUT, "allow_comment_loss and allow_comment_loss_ids are mutually exclusive"
        )

    from .formatted_text import validate_segments
    replacement_segments = validate_segments(replacement_segments, replace)
    session = live_write_mode.live_session_for(path)
    if replacement_segments is not None:
        live_write_mode.require_capability(session, "formatted_replacements_v1", feature_description="explicit formatted replacements")
    document_name = session.document_name
    _require_live_replace_capabilities(
        session, within_row_containing=within_row_containing, allow_comment_loss_ids=allow_comment_loss_ids
    )

    try:
        if scope != "body":
            from .live.scopes_live import scope_state
            scope_info = scope_state(session, scope, revision_before)
            pre_hash = scope_info["bodySha256"]
        else:
            scope_info = None
            pre_hash = _live_describe(session, path, revision_before)
        target = None if kind == "plain" else resolve_live_target(
            session, pre_hash, kind, find=find, expected_length=expected_length,
            find_sha256=find_sha256, para_ref=para_ref, expected_sha256=expected_sha256,
        )
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    try:
        result = _live_replace_once(
            session,
            find=target["text"] if target else (find or ""),
            replace=replace,
            expected_matches=1 if target else (expected_matches or 0),
            pre_hash=pre_hash,
            target=target,
            track_changes=track_changes,
            allow_comment_loss=allow_comment_loss,
            allow_comment_loss_ids=allow_comment_loss_ids,
            within_row_containing=within_row_containing,
            inherit_format=inherit_format,
            replacement_segments=replacement_segments,
            scope=scope,
            scope_info=scope_info,
        )
    except VerifyError as exc:
        if exc.envelope.error_code == ErrorCode.LIVE_STALE_AFTER_WRITE:
            _audit_stale_after_write(
                exc, tool="replace_text", path=path, document_name=document_name, pre_hash=pre_hash
            )
        raise

    matches = result.get("matches") or []
    match_count = result.get("match_count", len(matches))
    post_hash = result.get("post")
    before_text = "\n".join(m.get("before", "") for m in matches)
    after_text = "\n".join(m.get("after", "") for m in matches)

    evidence = live_write_mode.live_evidence(
        shape_result=result,
        applied=True,
        match_count=match_count,
        rung=2,
        before=before_text,
        after=after_text,
        pre_body_sha256=pre_hash,
        post_body_sha256=post_hash,
        document_name=document_name,
        tool="replace_text",
        path=path,
    )
    if scope_info:
        evidence["revision_before"] = f"live:scope:{scope_info['scopeSha256']}"
        evidence["revision_after"] = f"live:scope:{result['scope_post']}"
    evidence["matches"] = matches
    evidence["comments_removed"] = result.get("comments_removed", [])
    return evidence


# ---------------------------------------------------------------------------
# replace_text_batch (issue #82): N live replaces, one hash chain, one audit record.
#
# The edits run strictly in order; edit k is sent with expectedBodySha256 = the body hash
# edit k-1 produced, so the pane itself refuses ("stale", nothing written) if anything moved
# between two edits. The first failure stops the batch and nothing is ever retried: after a
# failure the caller cannot tell whether an op landed unless the failure is a KNOWN
# pre-write refusal, so only an explicit allowlist counts as "refused_before_write" -- every
# other failure is "unknown" and the caller must re-read the document before doing anything.
# Section locks: the shared broker drops body-hash checks for a client holding section locks
# (live/shared.py), so the hash chain may not be enforced there (docs/live-mode.md).
# ---------------------------------------------------------------------------

LIVE_BATCH_MAX_EDITS = 50

# Codes (local or pane-side) that prove the pane refused BEFORE touching the document.
_BATCH_REFUSED_CODES = frozenset(
    {
        ErrorCode.LIVE_STALE,
        ErrorCode.MATCH_COUNT_MISMATCH,
        ErrorCode.ZERO_MATCH,
        ErrorCode.ANCHORED_SHAPES,
        ErrorCode.WOULD_DELETE_COMMENTS,
        ErrorCode.LIVE_CAPABILITY_MISSING,
        ErrorCode.LOCKED_BY_OTHER_CLIENT,
        ErrorCode.OUTSIDE_LOCKED_SECTION,
        ErrorCode.INVALID_INPUT,
        ErrorCode.AMBIGUOUS_SPAN,
        ErrorCode.PARA_HASH_MISMATCH,
    }
)
_BATCH_REFUSED_PANE_CODES = frozenset(
    {"LOCK_SCOPE_UNRESOLVED", "WOULD_DELETE_COMMENTS", "ANCHORED_SHAPES", "zero_match", "match_count_mismatch", "stale",
     "AMBIGUOUS_SPAN"}
)
_BATCH_EDIT_KEYS = frozenset(
    {"find", "replace", "expected_matches", "allow_comment_loss_ids",
     "expected_length", "find_sha256", "para_ref", "expected_sha256"}
)


def _batch_refused_before_write(exc: VerifyError) -> bool:
    env = exc.envelope
    if env.error_code in _BATCH_REFUSED_CODES:
        return True
    return env.error_code == ErrorCode.LIVE_OP_FAILED and env.diagnostics.get("pane_code") in _BATCH_REFUSED_PANE_CODES


def _validate_batch_edits(edits: Any) -> list[dict[str, Any]]:
    if not isinstance(edits, list) or not 1 <= len(edits) <= LIVE_BATCH_MAX_EDITS:
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"edits must be a list of 1..{LIVE_BATCH_MAX_EDITS} edit objects"
        )
    out: list[dict[str, Any]] = []
    for i, edit in enumerate(edits):
        if not isinstance(edit, dict):
            raise _make_error(ErrorCode.INVALID_INPUT, f"edits[{i}] must be an object")
        if "allow_comment_loss" in edit:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"edits[{i}]: allow_comment_loss is not accepted in a batch; use allow_comment_loss_ids",
            )
        unknown = set(edit) - _BATCH_EDIT_KEYS
        if unknown:
            raise _make_error(ErrorCode.INVALID_INPUT, f"edits[{i}]: unknown key(s) {sorted(unknown)}")
        find, replace, expected = edit.get("find"), edit.get("replace"), edit.get("expected_matches")
        if find is not None and (not isinstance(find, str) or not find):
            raise _make_error(ErrorCode.INVALID_INPUT, f"edits[{i}].find must be a non-empty string")
        if not isinstance(replace, str):
            raise _make_error(ErrorCode.INVALID_INPUT, f"edits[{i}].replace must be a string")
        target_args: dict[str, Any] = {k: edit.get(k) for k in ("expected_length", "find_sha256", "para_ref", "expected_sha256")}
        try:
            kind = check_target_args(find, expected, **target_args)
            if kind == "plain" and (isinstance(expected, bool) or not isinstance(expected, int) or expected < 1):
                raise _make_error(ErrorCode.INVALID_INPUT, "expected_matches must be a positive integer")
            ids = validate_comment_loss_ids(edit.get("allow_comment_loss_ids"))
        except VerifyError as exc:
            raise _make_error(
                ErrorCode.INVALID_INPUT, f"edits[{i}]: {exc.envelope.message}", exc.envelope.diagnostics
            ) from exc
        out.append({"find": find, "replace": replace, "expected_matches": expected, "allow_comment_loss_ids": ids,
                    "kind": kind, **target_args})
    return out


def execute_replace_text_batch_live(
    path: str,
    edits: list[dict[str, Any]],
    revision_before: str,
    *,
    write_mode: str = "live",
    track_changes: bool = False,
    inherit_format: str = "replaced",
    verbose: bool = False,
) -> dict[str, Any]:
    """Live-only ``replace_text_batch`` (issue #82); see the block comment above.

    The report is compact (issue #112): top-level ``applied``, ``batch_status``,
    ``applied_count``, ``edit_count``, ``revision_before``, ``revision_after`` and one
    result per edit. An applied edit carries ``evidence: {match_count, revision_after}``
    (full tokens: callers chain them); a failed one carries ``code`` and ``message``.
    ``verbose=true`` restores each edit's before/after text, comments removed and shape
    counts.

    Returns ``{batch_status, results, revision_before, revision_after, applied, ...}``.
    ``batch_status`` is ``complete`` / ``partial`` / ``none_applied`` / ``unknown``;
    each result's ``status`` is ``applied`` / ``refused_before_write`` / ``unknown`` /
    ``not_attempted``. ``revision_after`` is the last applied edit's body hash, the original
    one when nothing applied, and ``None`` when the outcome is unknown. Argument errors, a
    missing pane, or a disconnect before the first op raise instead (nothing was sent).
    """
    if write_mode not in ("auto", "live"):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "replace_text_batch is live-only: write_mode must be 'live' (or 'auto' with a connected pane)",
        )
    if inherit_format not in ("replaced", "previous", "none"):
        raise _make_error(ErrorCode.INVALID_INPUT, "Invalid inherit_format")
    if not isinstance(revision_before, str) or not revision_before.startswith("live:sha256:"):
        raise _make_error(
            ErrorCode.INVALID_INPUT, "revision_before must be a 'live:sha256:<hex>' revision from a live read"
        )
    checked = _validate_batch_edits(edits)

    session = live_write_mode.live_session_for(path)
    document_name = session.document_name
    _require_live_replace_capabilities(
        session,
        within_row_containing=None,
        allow_comment_loss_ids=[] if any(e["allow_comment_loss_ids"] is not None for e in checked) else None,
    )

    try:
        pre_hash = session.request_threadsafe("describe")["bodySha256"]
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    results: list[dict[str, Any]] = []  # public: evidence per edit
    audit_results: list[dict[str, Any]] = []  # audit: no document text (audit redaction is top-level only)
    current = pre_hash
    applied_count = 0
    unknown = False

    def _record(
        i: int,
        status: str,
        *,
        code: str | None = None,
        message: str | None = None,
        evidence: dict[str, Any] | None = None,
    ) -> None:
        entry: dict[str, Any] = {"index": i, "status": status}
        if code is not None:
            entry["code"] = code
        if message:
            entry["message"] = message[:_BATCH_MESSAGE_CHARS]
        if evidence is not None:
            entry["evidence"] = evidence
        results.append(entry)
        # No document text: an error message can embed the pane's observed text, so only the
        # code is audited (the message goes to the caller).
        audit_entry = {k: v for k, v in entry.items() if k not in ("evidence", "message")}
        if evidence is not None:
            audit_entry["evidence"] = {k: v for k, v in evidence.items() if k not in ("before", "after")}
        audit_results.append(audit_entry)

    try:
        live_write_mode.check_not_stale(revision_before, pre_hash)
    except VerifyError as exc:
        _record(0, "refused_before_write", code=exc.envelope.error_code.value, message=exc.envelope.message)
        for i in range(1, len(checked)):
            _record(i, "not_attempted")
        return _finish_batch(
            path, document_name, results, audit_results, pre_hash, pre_hash, 0,
            unknown=False, total=len(checked), revision_before=revision_before, verbose=verbose,
        )

    for i, edit in enumerate(checked):
        if unknown or (results and results[-1]["status"] in ("refused_before_write", "not_attempted")):
            _record(i, "not_attempted")
            continue
        target = None
        if edit["kind"] != "plain":
            # Resolved against the body as edit i-1 left it; nothing is sent on a failure.
            try:
                target = resolve_live_target(
                    session, current, edit["kind"], find=edit["find"], expected_length=edit["expected_length"],
                    find_sha256=edit["find_sha256"], para_ref=edit["para_ref"],
                    expected_sha256=edit["expected_sha256"],
                )
            except VerifyError as exc:
                _record(i, "refused_before_write", code=exc.envelope.error_code.value, message=exc.envelope.message)
                continue
            except LiveDisconnected as exc:
                _record(i, "refused_before_write", code=ErrorCode.LIVE_DISCONNECTED.value, message=str(exc))
                continue
            except Exception as exc:  # noqa: BLE001 - a read/parse failure before dispatch: nothing was sent
                _record(i, "refused_before_write", code=ErrorCode.LIVE_OP_FAILED.value,
                        message=f"could not resolve the edit's target: {type(exc).__name__}: {exc}")
                continue
        try:
            result = _live_replace_once(
                session,
                find=target["text"] if target else edit["find"],
                replace=edit["replace"],
                expected_matches=1 if target else edit["expected_matches"],
                pre_hash=current,
                target=target,
                track_changes=track_changes,
                allow_comment_loss_ids=edit["allow_comment_loss_ids"],
                inherit_format=inherit_format,
            )
        except VerifyError as exc:
            code = exc.envelope.error_code.value
            if _batch_refused_before_write(exc):
                _record(i, "refused_before_write", code=code, message=exc.envelope.message)
            else:
                unknown = True
                _record(i, "unknown", code=code, message=exc.envelope.message)
            continue
        except Exception as exc:  # noqa: BLE001 - after dispatch, any other failure leaves the outcome unknown
            unknown = True
            _record(i, "unknown", code=ErrorCode.LIVE_OP_FAILED.value, message=str(exc))
            continue
        matches = result.get("matches") or []
        applied_count += 1
        current = result["post"]
        ev: dict[str, Any] = {
            "match_count": result.get("match_count", len(matches)),
            "before": "\n".join(m.get("before", "") for m in matches),
            "after": "\n".join(m.get("after", "") for m in matches),
            "revision_after": f"live:sha256:{current}",
            "comments_removed": result.get("comments_removed", []),
        }
        ev.update(live_write_mode.shape_counts(result))
        _record(i, "applied", evidence=ev)

    return _finish_batch(
        path, document_name, results, audit_results, pre_hash, current, applied_count,
        unknown=unknown, total=len(checked), revision_before=revision_before, verbose=verbose,
    )


_BATCH_MESSAGE_CHARS = 240


def _compact_batch_results(results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """One line per edit. An applied edit keeps ``evidence: {match_count, revision_after}``
    (the pieces callers read and chain from) and, only when non-empty, ``comments_removed``:
    losing a comment must never be silent."""
    compact: list[dict[str, Any]] = []
    for entry in results:
        row = {k: v for k, v in entry.items() if k != "evidence"}
        ev = entry.get("evidence")
        if ev is not None:
            row["evidence"] = {"match_count": ev["match_count"], "revision_after": ev["revision_after"]}
            if ev.get("comments_removed"):
                row["evidence"]["comments_removed"] = ev["comments_removed"]
        compact.append(row)
    return compact


def _finish_batch(
    path: str,
    document_name: str,
    results: list[dict[str, Any]],
    audit_results: list[dict[str, Any]],
    pre_hash: str,
    current_hash: str,
    applied_count: int,
    *,
    unknown: bool,
    total: int,
    revision_before: str,
    verbose: bool = False,
) -> dict[str, Any]:
    if unknown:
        batch_status = "unknown"
    elif applied_count == total:
        batch_status = "complete"
    elif applied_count == 0:
        batch_status = "none_applied"
    else:
        batch_status = "partial"
    evidence = live_write_mode.live_evidence(
        applied=batch_status == "complete",
        match_count=sum(r["evidence"]["match_count"] for r in results if r["status"] == "applied"),
        rung=2,
        before="",
        after="",
        pre_body_sha256=pre_hash,
        post_body_sha256=None if unknown else current_hash,
        document_name=document_name,
        tool="replace_text_batch",
        path=path,
        extra={
            "batch_status": batch_status,
            "applied_count": applied_count,
            "edit_count": total,
            "results": audit_results,
        },
    )
    if verbose:
        evidence["batch_status"] = batch_status
        evidence["applied_count"] = applied_count
        evidence["edit_count"] = total
        evidence["results"] = results
        # Echo the caller's own token (it may name an older body than the pane's describe on a stale start).
        evidence["revision_before"] = revision_before
        return evidence
    # Compact report (issue #112): none of the placeholder before/after/rung/match_count and
    # none of the constant live envelope. revision_before is the caller's own token (it may
    # name an older body than the pane's describe on a stale start) and is always present.
    report: dict[str, Any] = {
        "applied": evidence["applied"],
        "batch_status": batch_status,
        "applied_count": applied_count,
        "edit_count": total,
        "revision_before": revision_before,
        "revision_after": evidence["revision_after"],
        "results": _compact_batch_results(results),
    }
    if not evidence.get("audit_logged"):
        report["audit_logged"] = False
    if evidence.get("evidence_ref"):
        report["evidence_ref"] = evidence["evidence_ref"]
    return report


# ---------------------------------------------------------------------------
# Row-scoped edit (issue #22 B2): within_row_containing on replace_text/
# format_text (file mode). Chosen over a bare nth-occurrence index because
# a caller that already found ONE cell's own unique text (the real
# incident's own workaround: "put the program name at the front of the
# unique Evidence cell") already has a usable anchor -- an index would
# force counting matches blind instead. Contract, precisely: the anchor
# must be unique in the WHOLE DOCUMENT (the same guarantee locate() gives
# any needle already; no new per-row uniqueness logic invented) and must
# resolve to a table cell. Two cells in the SAME target row with
# identical `find` text remain inseparable -- a named, documented
# limitation (see server.py's docstrings), not solved here.
# ---------------------------------------------------------------------------


def _paragraph_by_ref(proj: projection.Projection) -> dict[str, projection.ParagraphMeta]:
    return {p.para_ref: p for p in proj.paragraphs}


def _innermost_table_row(para_meta: projection.ParagraphMeta | None) -> tuple[int, int] | None:
    """(table_id, row) from *para_meta*'s own container_chain -- the
    INNERMOST table a paragraph sits in (a nested table's row, not its
    host table's), or None if *para_meta* is None or not inside any
    table at all."""
    if para_meta is None or not para_meta.container_chain:
        return None
    last = para_meta.container_chain[-1]
    table_id = last.get("table_id")
    row = last.get("row")
    if table_id is None or row is None:
        return None
    return table_id, row


def _row_key_for_span(
    proj: projection.Projection, para_map: dict[str, projection.ParagraphMeta], start: int
) -> tuple[int, int] | None:
    located = proj.locate_offset(start)
    if located is None:
        return None
    para_ref, _run_ref, _run_offset = located
    return _innermost_table_row(para_map.get(para_ref))


def _resolve_row_span_filter(proj: projection.Projection, within_row_containing: str) -> Callable[[int, int], bool]:
    """Locate *within_row_containing* (globally unique, per locate()'s own
    contract -- no filter of its own) and return a span_filter scoped to
    that match's own table row, for a second locate() call over `find`.

    Raises ZERO_MATCH/MATCH_COUNT_MISMATCH (re-labeled to name
    within_row_containing as the failing needle, not `find`) if the
    anchor itself doesn't resolve uniquely, or INVALID_INPUT if it
    resolves but not inside any table cell.
    """
    try:
        anchor_result = locate(within_row_containing, proj, 1)
    except VerifyError as exc:
        if exc.envelope.error_code in (ErrorCode.ZERO_MATCH, ErrorCode.MATCH_COUNT_MISMATCH):
            raise _make_error(
                exc.envelope.error_code,
                f"within_row_containing={within_row_containing!r} must be unique in the whole "
                f"document (it is the anchor locate() call that failed, not the main `find`): "
                f"{exc.envelope.message}",
                {**exc.envelope.diagnostics, "within_row_containing": within_row_containing},
            ) from exc
        raise

    para_map = _paragraph_by_ref(proj)
    anchor_start, _anchor_end = anchor_result.spans[0]
    row_key = _row_key_for_span(proj, para_map, anchor_start)
    if row_key is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"within_row_containing={within_row_containing!r} does not resolve to a table cell "
            "(it must be unique text located inside a table row).",
            {"within_row_containing": within_row_containing},
        )

    def _span_filter(start: int, _end: int) -> bool:
        return _row_key_for_span(proj, para_map, start) == row_key

    return _span_filter


# ---------------------------------------------------------------------------
# Tool 1: replace_text
# ---------------------------------------------------------------------------


def execute_replace_text(
    path: str,
    find: str | None,
    replace: str,
    expected_matches: int | None,
    *,
    revision_before: str | None = None,
    force: bool = False,
    allow_concurrent_editor: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
    within_row_containing: str | None = None,
    allow_comment_loss: bool = False,
    inherit_format: str = "replaced",
    replacement_segments: list[dict[str, Any]] | None = None,
    scope: str = "body",
    allow_comment_loss_ids: list[dict[str, Any]] | None = None,
    expected_length: int | None = None,
    find_sha256: str | None = None,
    para_ref: str | None = None,
    expected_sha256: str | None = None,
) -> dict[str, Any]:
    """replace_text: *find* is a plain needle (with *expected_matches*), an
    anchored ``"<prefix>…<suffix>"`` (with *expected_length* and/or
    *find_sha256*), or None with *para_ref* + *expected_sha256* (the whole
    paragraph). See server.replace_text."""
    target_args: dict[str, Any] = {"expected_length": expected_length, "find_sha256": find_sha256,
                   "para_ref": para_ref, "expected_sha256": expected_sha256}
    kind = check_target_args(find, expected_matches, **target_args)
    if kind == "para_ref" and within_row_containing:
        raise _make_error(ErrorCode.INVALID_INPUT, "within_row_containing does not apply to a para_ref target")
    allow_comment_loss_ids = validate_comment_loss_ids(allow_comment_loss_ids)
    if allow_comment_loss and allow_comment_loss_ids is not None:
        raise _make_error(
            ErrorCode.INVALID_INPUT, "allow_comment_loss and allow_comment_loss_ids are mutually exclusive"
        )
    from .formatted_text import validate_segments
    replacement_segments = validate_segments(replacement_segments, replace)
    mode = live_write_mode.resolve_write_mode(path, write_mode)
    if replacement_segments is not None and mode != "live":
        raise _make_error(ErrorCode.LIVE_UNAVAILABLE, "Explicit segments require live Word")
    if allow_comment_loss_ids is not None and mode != "live":
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "allow_comment_loss_ids is live-only (it needs the Word pane's comment guard); in file mode "
            "use allow_comment_loss=True or write_mode=\"live\"",
        )
    if mode == "live":
        return execute_replace_text_live(
            path,
            find,
            replace,
            expected_matches,
            revision_before=revision_before,
            track_changes=track_changes,
            within_row_containing=within_row_containing,
            scope=scope,
            allow_comment_loss=allow_comment_loss,
            allow_comment_loss_ids=allow_comment_loss_ids,
            inherit_format=inherit_format,
            replacement_segments=replacement_segments,
            **target_args,
        )

    if scope != "body":
        raise _make_error(ErrorCode.INVALID_INPUT, "scope requires live mode")
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(
        resolved,
        revision_before,
        allow_concurrent_editor=allow_concurrent_editor,
        live_capable=True,
    )

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    span_filter = _resolve_row_span_filter(proj, within_row_containing) if within_row_containing else None
    locate_result: LocateResult = _locate_file_target(
        proj, kind, find, expected_matches, span_filter=span_filter, **target_args
    )
    own_author = resolve_author_name()
    _check_tracked_changes_guard(proj, locate_result, force, own_author)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    track = _TrackContext(document_root, author=own_author) if track_changes else None
    for start, end in locate_result.spans:
        atoms = _atoms_for_span(proj, start, end)
        _apply_replace_span(start, end, replace, atoms, track=track)

    # Re-walk the SAME live tree (not a re-parse) to find where the
    # inserted replacement text landed, for the after excerpt/runs_after,
    # and to build the "intended after" text the post-write re-read is
    # checked against below.
    mutated_proj = projection.project_document_root(document_root)
    after_excerpt = _excerpt(mutated_proj.text, before_first[0], before_first[0] + len(replace))
    runs_after = _collect_style_runs(mutated_proj, [(before_first[0], before_first[0] + len(replace))])
    intended_after_text = mutated_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")

    conflict_sweep = _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = {"token": conflict_sweep["revision_after"]}  # issue #27: the STAGED token, never a re-read of the file
    evidence = _evidence(
        applied=True,
        match_count=locate_result.match_count,
        rung=locate_result.rung,
        before=before_excerpt,
        after=after_excerpt,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        runs_before=runs_before,
        runs_after=runs_after,
        warnings=locate_result.warnings,
        track=track,
        conflict_sweep=conflict_sweep,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_text", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


def execute_format_text_live(
    path: str,
    find: str,
    style: dict[str, bool | str],
    expected_matches: int | None,
    *,
    revision_before: str | None = None,
    track_changes: bool = False,
    within_row_containing: str | None = None,
    scope: str = "body",
    expected_length: int | None = None,
    find_sha256: str | None = None,
) -> dict[str, Any]:
    """Live-mode ``format_text`` (issue #106 WP-3): sends a ``format`` op
    over the pane's WSS ops channel instead of editing the .docx file.

    Same describe -> op -> verify shape as ``execute_replace_text_live``,
    with a verification predicate suited to what ``format`` actually
    changes: unlike ``replace``, a format-only op never changes
    ``body.text`` (per ``live/protocol.py``'s own documented
    ``FormatResult`` shape and ``tests/unit/fake_pane.py``'s
    ``FakeDocument.format`` -- both report ``pre == post`` for a format
    op even when the style genuinely changed), so a ``post != pre``
    check would be wrong here -- it is REQUIRED to differ for replace,
    and EXPECTED to be equal for format. What IS checked: the pane
    reported ``applied: true``, and every matched range's ``after`` text
    still equals its own ``before`` text (format changes styling, never
    content -- a match whose text changed anyway is a real verification
    failure, not a false alarm).

    Issue #22: ``style`` may also carry ``strike``/``color`` (a 6-hex
    string). The pane's reply carries a read-back ``colorAfter``/
    ``strikeAfter`` per match (the pane's own ``font.color``/
    ``.strikeThrough`` re-loaded AFTER ``context.sync()``, not an echo of
    the request) -- when either was requested, this function checks the
    read-back value against what was asked for, so a write that silently
    didn't take on the real Word object model is a verification failure,
    not a false "applied: true".

    ``within_row_containing`` (issue #22 B2): same ``rowAnchor``/
    ``"row_scope"``-capability contract as
    ``execute_replace_text_live`` -- see that function's own docstring.
    """
    style = _validate_style(style)
    kind = check_target_args(
        find, expected_matches, expected_length=expected_length, find_sha256=find_sha256, allow_para_ref=False
    )
    _check_live_target_scope(kind, scope=scope, within_row_containing=within_row_containing)

    session = live_write_mode.live_session_for(path)
    document_name = session.document_name
    live_write_mode.require_capability(session, "format_readback", feature_description="all-property formatting readback")
    if within_row_containing:
        live_write_mode.require_capability(
            session, _ROW_SCOPE_CAPABILITY, feature_description="within_row_containing"
        )

    try:
        if scope != "body":
            from .live.scopes_live import scope_state
            scope_info = scope_state(session, scope, revision_before)
            pre_hash = scope_info["bodySha256"]
        else:
            scope_info = None
            pre_hash = _live_describe(session, path, revision_before)
        target = None if kind == "plain" else resolve_live_target(
            session, pre_hash, kind, find=find, expected_length=expected_length, find_sha256=find_sha256
        )
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    payload = {
        "find": target["text"] if target else find,
        "expected_matches": 1 if target else expected_matches,
        "bold": style.get("bold"),
        "italic": style.get("italic"),
        "underline": style.get("underline"),
        "strike": style.get("strike"),
        "color": style.get("color"),
        "track_changes": track_changes,
    }
    payload["scope"] = scope
    payload["expectedBodySha256"] = pre_hash
    if scope_info:
        payload["expectedScopeSha256"] = scope_info["scopeSha256"]
    if within_row_containing:
        payload["rowAnchor"] = within_row_containing
    _apply_pane_target(payload, target)
    try:
        result = session.request_threadsafe("format", payload, expected_body_sha256=pre_hash)
    except LiveStaleAfterWrite as exc:
        error = _make_error(
            ErrorCode.LIVE_STALE_AFTER_WRITE,
            "the pane applied the format, but its body hash before the op differed from the one "
            "this call expected: the write ALREADY LANDED -- do not retry; re-read the live document.",
            {"expected": exc.expected, "actual": exc.actual},
        )
        _audit_stale_after_write(
            error, tool="format_text", path=path, document_name=session.document_name, pre_hash=pre_hash
        )
        raise error from exc
    except LiveStale as exc:
        raise _make_error(
            ErrorCode.LIVE_STALE, str(exc), {"expected": exc.expected, "actual": exc.actual}
        ) from exc
    except LiveOpFailed as exc:
        raise _make_error(
            live_write_mode.classify_op_failed(exc), exc.message, {"pane_code": exc.code}
        ) from exc
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    matches = result.get("matches") or []
    match_count = result.get("match_count", len(matches))
    post_hash = result.get("post")

    if not result.get("applied"):
        raise _make_error(
            ErrorCode.VERIFICATION_FAILED,
            "live format did not verify: the pane did not report applied=true. Nothing to roll "
            "back in live mode -- Word, not this server, owns the document; re-read via describe/"
            "read_document and retry.",
            {"pre": pre_hash, "post": post_hash, "applied": result.get("applied")},
        )
    for m in matches:
        if m.get("after") != m.get("before"):
            raise _make_error(
                ErrorCode.VERIFICATION_FAILED,
                "live format did not verify: a matched range's text changed, but format_text must "
                f"never change content ({m.get('before')!r} -> {m.get('after')!r}). Nothing to "
                "roll back in live mode -- Word, not this server, owns the document.",
                {"matches": matches},
            )
        for key in ("bold", "italic", "underline"):
            if key in style and m.get(f"{key}After") is not style[key]:
                raise _make_error(ErrorCode.VERIFICATION_FAILED,
                                  f"Live {key} read-back differs from request", {"matches": matches})
        requested_color = style.get("color")
        if requested_color is not None:
            # Word's own Office.js font.color GETTER returns a "#"-prefixed
            # CSS hex string (e.g. "#3B3838") even though the SETTER
            # tolerates a bare one -- verified against a real Word
            # sideload (issue #22), not assumed: the first version of this
            # check compared the read-back value directly against
            # style["color"] (never "#"-prefixed, since _validate_style
            # strips a leading "#" on the way in), so a color that WAS
            # correctly applied still raised VERIFICATION_FAILED on every
            # single live color call. fake_pane.py's own echo-the-request
            # model never had a "#" to strip, so this never surfaced
            # against the fake pane -- only a real sideload caught it.
            color_after = (m.get("colorAfter") or "").upper().lstrip("#")
            if color_after != requested_color:
                raise _make_error(
                    ErrorCode.VERIFICATION_FAILED,
                    "live format did not verify: the pane's read-back font.color "
                    f"({m.get('colorAfter')!r}) does not equal the requested color "
                    f"({requested_color!r}) after context.sync(). Nothing to roll back in live mode "
                    "-- Word, not this server, owns the document.",
                    {"matches": matches, "requested_color": requested_color},
                )
        requested_strike = style.get("strike")
        if requested_strike is not None and bool(m.get("strikeAfter")) != bool(requested_strike):
            raise _make_error(
                ErrorCode.VERIFICATION_FAILED,
                "live format did not verify: the pane's read-back font.strikeThrough "
                f"({m.get('strikeAfter')!r}) does not equal the requested strike "
                f"({requested_strike!r}) after context.sync(). Nothing to roll back in live mode "
                "-- Word, not this server, owns the document.",
                {"matches": matches, "requested_strike": requested_strike},
            )

    before_text = "\n".join(m.get("before", "") for m in matches)
    after_text = "\n".join(m.get("after", "") for m in matches)

    evidence = live_write_mode.live_evidence(
        shape_result=result,
        applied=True,
        match_count=match_count,
        rung=1,
        before=before_text,
        after=after_text,
        pre_body_sha256=pre_hash,
        post_body_sha256=post_hash,
        document_name=document_name,
        tool="format_text",
        path=path,
    )


# ---------------------------------------------------------------------------
# Tool 2: format_text
# ---------------------------------------------------------------------------

    if scope_info:
        evidence["revision_before"] = f"live:scope:{scope_info['scopeSha256']}"
        evidence["revision_after"] = f"live:scope:{result['scope_post']}"
    evidence["matches"] = matches
    return evidence


def _execute_format_text(
    path: str,
    find: str,
    style: dict[str, bool | str],
    expected_matches: int | None,
    *,
    revision_before: str | None = None,
    force: bool = False,
    allow_concurrent_editor: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
    within_row_containing: str | None = None,
    scope: str = "body",
    expected_length: int | None = None,
    find_sha256: str | None = None,
) -> dict[str, Any]:
    style = _validate_style(style)
    target_args: dict[str, Any] = {"expected_length": expected_length, "find_sha256": find_sha256}
    kind = check_target_args(find, expected_matches, allow_para_ref=False, **target_args)
    mode = live_write_mode.resolve_write_mode(path, write_mode)
    if mode == "live":
        return execute_format_text_live(
            path,
            find,
            style,
            expected_matches,
            revision_before=revision_before,
            track_changes=track_changes,
            within_row_containing=within_row_containing,
            scope=scope,
            **target_args,
        )

    if scope != "body":
        raise _make_error(ErrorCode.INVALID_INPUT, "scope requires live mode")
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(
        resolved,
        revision_before,
        allow_concurrent_editor=allow_concurrent_editor,
        live_capable=True,
    )

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    span_filter = _resolve_row_span_filter(proj, within_row_containing) if within_row_containing else None
    locate_result: LocateResult = _locate_file_target(
        proj, kind, find, expected_matches, span_filter=span_filter, para_ref=None, expected_sha256=None, **target_args
    )
    own_author = resolve_author_name()
    _check_tracked_changes_guard(proj, locate_result, force, own_author)
    _check_foreign_rpr_change(proj, locate_result.spans, force, own_author)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    # No-op check: an idempotent re-run must not create a new revision (the
    # same rationale as GoogleDocs-MCP's format_text -- see formatting.py).
    # Nothing to record in a w:rPrChange either -- no rPr is actually
    # changing, tracked or not.
    if _style_matches(runs_before, style):
        after_excerpt = before_excerpt
        evidence = _evidence(
            applied=True,
            match_count=locate_result.match_count,
            rung=locate_result.rung,
            before=before_excerpt,
            after=after_excerpt,
            revision_before=pre_revision["token"],
            revision_after=pre_revision["token"],
            audit_logged=False,
            runs_before=runs_before,
            runs_after=runs_before,
            warnings=locate_result.warnings,
            track=_TrackContext(document_root, author=own_author) if track_changes else None,
        )
        logged, _ = audit.append_audit(path=str(resolved), tool="format_text", evidence=evidence)
        evidence["audit_logged"] = logged
        return evidence

    track = _TrackContext(document_root, author=own_author) if track_changes else None
    for start, end in locate_result.spans:
        atoms = _atoms_for_span(proj, start, end)
        _apply_format_span(start, end, style, atoms, track=track)

    # format_text never changes character counts, so the spans are stable
    # across the mutation -- no need to re-locate to find them again.
    mutated_proj = projection.project_document_root(document_root)
    after_excerpt = _excerpt(mutated_proj.text, before_first[0], before_first[1])
    runs_after = _collect_style_runs(mutated_proj, locate_result.spans)
    intended_after_text = mutated_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")
        post_proj = projection.project_part(written_path)
        post_runs = _collect_style_runs(post_proj, locate_result.spans)
        if not _style_matches(post_runs, style):
            raise ValueError(f"re-read style does not match the requested style: {post_runs}")

    conflict_sweep = _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = {"token": conflict_sweep["revision_after"]}  # issue #27: the STAGED token, never a re-read of the file
    evidence = _evidence(
        applied=True,
        match_count=locate_result.match_count,
        rung=locate_result.rung,
        before=before_excerpt,
        after=after_excerpt,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        runs_before=runs_before,
        runs_after=runs_after,
        warnings=locate_result.warnings,
        track=track,
        conflict_sweep=conflict_sweep,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="format_text", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool 3: apply_style -- issue #28 WP-15a. Applies a NAMED style (from
# list_styles) rather than format_text's four boolean toggles. Modeled on
# format_text's exact pipeline (locate -> _atoms_for_span -> per-atom
# apply -> _serialize_and_write -> 8-key evidence), which is also why it
# lives here rather than in a new module: it reuses _atoms_for_span,
# _own_rpr, _build_run, _index_of, _TrackContext, and
# _check_tracked_changes_guard directly.
#
# A CHARACTER style (w:type="character") applies to the matched run(s)
# exactly like format_text, including track_changes=True support
# (w:rPrChange, per WP-07b-a's own contract, which names rPrChange for
# "formatting changes"). A PARAGRAPH style (w:type="paragraph") applies to
# every paragraph CONTAINING a matched run (the whole paragraph, not just
# the matched substring -- w:pStyle has no sub-paragraph granularity).
#
# Issue #7: track_changes=True on a paragraph style now writes a real
# w:pPrChange (tracked_changes.apply_ppr_change), built from a Word-
# authored fixture (tests/fixtures/revision/pstyle-tracked.docx --
# Track Changes on, one paragraph style change, saved) rather than
# guessed -- this used to raise INVALID_INPUT naming the gap; see
# _apply_paragraph_style's own docstring for the exact shape and the
# own-author-reuse rule the twice/foreign fixtures informed.
# ---------------------------------------------------------------------------

_SUPPORTED_STYLE_TYPES = frozenset({"paragraph", "character"})


def _set_run_style(r_elem: Any, style_id: str, *, track: _TrackContext | None) -> None:
    """Set/replace <w:rStyle w:val=style_id> on r_elem's own w:rPr
    (creating one if absent; w:rStyle must be rPr's FIRST child per the
    OOXML schema's fixed element order). Tracked the same way
    _apply_style_to_run tracks a boolean toggle: a w:rPrChange recording
    the PRE-change rPr, snapshotted before anything is mutated -- same
    own-author-reuse rule as _apply_style_to_run (issue #7; see that
    function's docstring)."""
    old_rpr = _own_rpr(r_elem)
    existing_author = tracked_changes.existing_change_author(old_rpr, "rPrChange") if track else None
    reuse = bool(track) and existing_author == track.author
    old_rpr_snapshot = copy.deepcopy(old_rpr) if (track and not reuse and old_rpr is not None) else None
    rpr = old_rpr
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        r_elem.insert(0, rpr)
    existing = None
    for child in rpr:
        if projection._ln(child) == "rStyle":
            existing = child
            break
    if existing is None:
        rstyle = ET.Element(_w("rStyle"), {_w("val"): style_id})
        rpr.insert(0, rstyle)
    else:
        existing.set(_w("val"), style_id)
    if track and not reuse:
        new_rpr = _own_rpr(r_elem)
        rid = track.next_id()
        tracked_changes.apply_rpr_change(new_rpr, old_rpr_elem=old_rpr_snapshot, rid=rid, author=track.author, date=track.date)


def _apply_named_style_span(
    start: int,
    end: int,
    style_id: str,
    atoms: list[tuple[int, int, RunEvent]],
    *,
    track: _TrackContext | None = None,
) -> None:
    """Same run-splitting rule as _apply_format_span: a boundary run
    splits, the matched middle survives as a new run carrying the named
    style, character count never changes."""
    for s, e, event in atoms:
        local_start = max(start, s) - s
        local_end = min(end, e) - s
        full_text = event.text
        before_text = full_text[:local_start]
        after_text = full_text[local_end:]

        if not before_text and not after_text:
            if event.r_elem is not None:
                _set_run_style(event.r_elem, style_id, track=track)
            continue

        if event.r_elem is None or event.parent_elem is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to edit", {"start": start, "end": end})

        original_rpr = _own_rpr(event.r_elem)
        middle_text = full_text[local_start:local_end]
        middle_run = _build_run(middle_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
        _set_run_style(middle_run, style_id, track=track)

        parent = event.parent_elem
        idx = list(parent).index(event.r_elem)
        if before_text and after_text:
            _set_node_text(event, before_text)
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx + 1, middle_run)
            parent.insert(idx + 2, after_run)
        elif before_text:
            _set_node_text(event, before_text)
            parent.insert(idx + 1, middle_run)
        else:
            after_run = _build_run(after_text, copy.deepcopy(original_rpr) if original_rpr is not None else None)
            parent.insert(idx, middle_run)
            parent.insert(idx + 1, after_run)
            _remove_if_present(parent, event.r_elem)


def _build_parent_map(root: Any) -> dict[int, Any]:
    parent_map: dict[int, Any] = {}
    for parent in root.iter():
        for child in parent:
            parent_map[id(child)] = parent
    return parent_map


def _enclosing_paragraph(parent_map: dict[int, Any], elem: Any | None) -> Any | None:
    current = elem
    while current is not None:
        if projection._ln(current) == "p":
            return current
        current = parent_map.get(id(current))
    return None


def _apply_paragraph_style(p_elem: Any, style_id: str, *, track: _TrackContext | None) -> bool:
    """Set/replace <w:pStyle w:val=style_id> on p_elem's own w:pPr
    (creating one if absent; w:pStyle must be pPr's FIRST child).
    Returns True if the paragraph's own style actually changed (False for
    a no-op re-application of the style it already carries) -- the
    caller uses this to build the "paragraphs to verify carry a
    w:pPrChange" list precisely, rather than every touched paragraph
    (issue #7: a no-op must not require one).

    Issue #7, track_changes=True: records a w:pPrChange with the
    PRE-change pPr snapshot -- shape taken directly from a real
    Word-authored fixture (tests/fixtures/revision/pstyle-tracked.docx:
    Track Changes on, one paragraph style change, saved), not guessed.
    w:pPrChange lands as pPr's LAST child, right after w:pStyle; when the
    paragraph had no w:pPr at all before, the snapshot is an empty
    <w:pPr/> (mirrors apply_rpr_change's own "no explicit properties"
    convention for a run with no prior w:rPr).

    Own-author reuse (verified: tests/fixtures/revision/
    pstyle-tracked-twice.docx -- a SECOND tracked style change on the
    same paragraph, same Word session, before any accept/reject): Word
    updates pStyle in place and leaves the FIRST pPrChange's id/date/
    snapshot completely untouched, rather than stacking a second one.
    Mirrored here: if the paragraph already carries a pPrChange from
    *track.author*, this call only updates the live pStyle: it does not
    consume a new id and does not touch the existing change record. A
    FOREIGN-authored existing pPrChange is refused before this function
    is ever called (_check_foreign_ppr_change), unless force=True let it
    through -- in that case apply_ppr_change's own remove-existing-record
    safety net replaces it rather than producing invalid XML.
    """
    ppr = None
    for child in p_elem:
        if projection._ln(child) == "pPr":
            ppr = child
            break
    if ppr is None:
        ppr = ET.Element(_w("pPr"))
        p_elem.insert(0, ppr)
    existing = None
    for child in ppr:
        if projection._ln(child) == "pStyle":
            existing = child
            break
    current_val = projection._attr(existing, "val") if existing is not None else None
    if current_val == style_id:
        return False

    existing_author = tracked_changes.existing_change_author(ppr, "pPrChange") if track else None
    reuse = bool(track) and existing_author == track.author
    old_ppr_snapshot = copy.deepcopy(ppr) if (track and not reuse) else None
    if old_ppr_snapshot is not None:
        tracked_changes._remove_change_record(old_ppr_snapshot, "pPrChange")

    if existing is None:
        pstyle = ET.Element(_w("pStyle"), {_w("val"): style_id})
        ppr.insert(0, pstyle)
    else:
        existing.set(_w("val"), style_id)

    if track and not reuse:
        rid = track.next_id()
        tracked_changes.apply_ppr_change(ppr, old_ppr_elem=old_ppr_snapshot, rid=rid, author=track.author, date=track.date)
    return True


def _execute_apply_style(
    path: str,
    find: str,
    style_id: str,
    expected_matches: int | None,
    *,
    revision_before: str | None = None,
    force: bool = False,
    allow_concurrent_editor: bool = False,
    track_changes: bool = False,
    expected_length: int | None = None,
    find_sha256: str | None = None,
) -> dict[str, Any]:
    target_args: dict[str, Any] = {"expected_length": expected_length, "find_sha256": find_sha256}
    kind = check_target_args(find, expected_matches, allow_para_ref=False, **target_args)
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(
        resolved,
        revision_before,
        allow_concurrent_editor=allow_concurrent_editor,
        live_capable=False,
    )

    style_records = projection.list_styles_impl(resolved)
    styles_by_style_id = {s["style_id"]: s for s in style_records if s["style_id"]}
    style_info = styles_by_style_id.get(style_id)
    if style_info is None:
        raise _make_error(
            ErrorCode.STYLE_NOT_FOUND,
            f"style_id {style_id!r} is not a style in this document's styles.xml.",
            {"style_id": style_id, "available_style_ids": sorted(styles_by_style_id)},
        )
    style_type = style_info.get("type")
    if style_type not in _SUPPORTED_STYLE_TYPES:
        raise _make_error(
            ErrorCode.UNSUPPORTED_STYLE_TYPE,
            f"style_id {style_id!r} is a {style_type!r} style; apply_style supports paragraph and character "
            "styles only.",
            {"style_id": style_id, "type": style_type},
        )

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = _locate_file_target(
        proj, kind, find, expected_matches, para_ref=None, expected_sha256=None, **target_args
    )
    own_author = resolve_author_name()
    _check_tracked_changes_guard(proj, locate_result, force, own_author)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    track = _TrackContext(document_root, author=own_author) if track_changes else None

    if style_type == "character":
        _check_foreign_rpr_change(proj, locate_result.spans, force, own_author)
        for start, end in locate_result.spans:
            atoms = _atoms_for_span(proj, start, end)
            _apply_named_style_span(start, end, style_id, atoms, track=track)
    else:
        parent_map = _build_parent_map(document_root)
        touched_ids: set[int] = set()
        touched_paragraphs: list[Any] = []
        span_paragraph_id: dict[tuple[int, int], int] = {}
        for start, end in locate_result.spans:
            atoms = _atoms_for_span(proj, start, end)
            p_elem = _enclosing_paragraph(parent_map, atoms[0][2].r_elem) if atoms else None
            if p_elem is not None:
                span_paragraph_id[(start, end)] = id(p_elem)
                if id(p_elem) not in touched_ids:
                    touched_ids.add(id(p_elem))
                    touched_paragraphs.append(p_elem)
        _check_foreign_ppr_change(touched_paragraphs, force, own_author)
        changed_by_pid: dict[int, bool] = {}
        for p_elem in touched_paragraphs:
            changed_by_pid[id(p_elem)] = _apply_paragraph_style(p_elem, style_id, track=track)
        # span_changed (issue #7): per-span "did this span's enclosing
        # paragraph's style actually change" -- used by _post_verify below
        # to check for a w:pPrChange ONLY on a paragraph that changed, per
        # the no-op rule (a paragraph already at style_id must not gain
        # one). Keyed by span, not by paragraph object identity, because
        # _post_verify re-parses the written XML into a BRAND NEW tree
        # (fresh Element objects, same spans/positions) to verify against.
        span_changed = {span: changed_by_pid.get(pid, False) for span, pid in span_paragraph_id.items()}

        # No-op check (issue #7): when EVERY touched paragraph already
        # carried style_id, _apply_paragraph_style mutated nothing at all
        # -- skip the write entirely rather than re-serializing an
        # unchanged tree, which (like format_text's own identical no-op
        # rationale) would otherwise still shift the revision token on
        # ElementTree round-trip formatting alone (e.g. self-closing tag
        # spacing) despite no semantic change.
        if touched_paragraphs and not any(changed_by_pid.values()):
            evidence = _evidence(
                applied=True,
                match_count=locate_result.match_count,
                rung=locate_result.rung,
                before=before_excerpt,
                after=before_excerpt,
                revision_before=pre_revision["token"],
                revision_after=pre_revision["token"],
                audit_logged=False,
                runs_before=runs_before,
                runs_after=runs_before,
                warnings=locate_result.warnings,
                track=track,
            )
            evidence["style_id"] = style_id
            evidence["style_type"] = style_type
            logged, _ = audit.append_audit(path=str(resolved), tool="apply_style", evidence=evidence)
            evidence["audit_logged"] = logged
            return evidence

    # format_text never changes character counts, and neither does
    # apply_style (a style id is metadata, not content) -- spans are
    # stable across the mutation.
    mutated_proj = projection.project_document_root(document_root)
    after_excerpt = _excerpt(mutated_proj.text, before_first[0], before_first[1])
    runs_after = _collect_style_runs(mutated_proj, locate_result.spans)
    intended_after_text = mutated_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")
        new_document_root, _ = mutations._load_document(written_path)
        new_proj = projection.project_document_root(new_document_root)
        if style_type == "character":
            for start, end in locate_result.spans:
                for _s, _e, event in _atoms_for_span(new_proj, start, end):
                    if event.r_elem is None:
                        continue
                    rpr = _own_rpr(event.r_elem)
                    rstyle_val = None
                    if rpr is not None:
                        for child in rpr:
                            if projection._ln(child) == "rStyle":
                                rstyle_val = projection._attr(child, "val")
                    if rstyle_val != style_id:
                        raise ValueError(f"re-read run does not carry w:rStyle val={style_id!r} (got {rstyle_val!r})")
        else:
            new_parent_map = _build_parent_map(new_document_root)
            for start, end in locate_result.spans:
                for _s, _e, event in _atoms_for_span(new_proj, start, end):
                    p_elem = _enclosing_paragraph(new_parent_map, event.r_elem)
                    if p_elem is None:
                        raise ValueError("re-read match has no enclosing paragraph")
                    ppr_elem = None
                    pstyle_val = None
                    for child in p_elem:
                        if projection._ln(child) == "pPr":
                            ppr_elem = child
                            for gc in child:
                                if projection._ln(gc) == "pStyle":
                                    pstyle_val = projection._attr(gc, "val")
                    if pstyle_val != style_id:
                        raise ValueError(f"re-read paragraph does not carry w:pStyle val={style_id!r} (got {pstyle_val!r})")
                    # Issue #7: only a paragraph whose style ACTUALLY
                    # changed (span_changed, computed before mutation) must
                    # carry a w:pPrChange -- a no-op re-application of the
                    # same style must not gain one, matching the no-op rule
                    # above (and format_text's own identical rule).
                    if track_changes and span_changed.get((start, end), False):
                        change_author = tracked_changes.existing_change_author(ppr_elem, "pPrChange")
                        if change_author != own_author:
                            raise ValueError(
                                f"re-read paragraph does not carry a w:pPrChange authored by {own_author!r} "
                                f"despite track_changes=True (got author={change_author!r})"
                            )
                    break  # one enclosing paragraph per span (structural boundary guarantee)

    conflict_sweep = _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = {"token": conflict_sweep["revision_after"]}  # issue #27: the STAGED token, never a re-read of the file
    evidence = _evidence(
        applied=True,
        match_count=locate_result.match_count,
        rung=locate_result.rung,
        before=before_excerpt,
        after=after_excerpt,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        runs_before=runs_before,
        runs_after=runs_after,
        warnings=locate_result.warnings,
        track=track,
        conflict_sweep=conflict_sweep,
    )
    evidence["style_id"] = style_id
    evidence["style_type"] = style_type
    logged, _ = audit.append_audit(path=str(resolved), tool="apply_style", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


_FORMAT_DELTA_ATTRS = ("bold", "italic", "underline", "strike", "color")


def _compact_format_evidence(evidence: dict[str, Any], *, verbose: bool) -> dict[str, Any]:
    """Shape a formatting-only result (issue #112).

    format_text and apply_style never change text, so echoing a before and an after
    excerpt (and every run, both times) says nothing the caller needs. When the two
    excerpts are equal the result says ``text_unchanged: true`` with one ``excerpt``,
    and ``format_delta`` lists only the attributes that moved. ``runs_before`` /
    ``runs_after`` come back only with ``verbose=true``.
    """
    runs_before, runs_after = evidence.get("runs_before"), evidence.get("runs_after")
    if runs_before is not None and runs_after is not None:
        delta: dict[str, Any] = {}
        for attr in _FORMAT_DELTA_ATTRS:
            old = sorted({str(r.get(attr)) for span in runs_before for r in span})
            new = sorted({str(r.get(attr)) for span in runs_after for r in span})
            if old != new:
                delta[attr] = {"before": old, "after": new}
        if delta:
            evidence["format_delta"] = delta
    if not verbose:
        evidence.pop("runs_before", None)
        evidence.pop("runs_after", None)
    if "before" in evidence and evidence.get("before") == evidence.get("after"):
        evidence["text_unchanged"] = True
        evidence["excerpt"] = evidence.pop("before")
        evidence.pop("after", None)
    return evidence


@functools.wraps(_execute_format_text)
def execute_format_text(*args: Any, verbose: bool = False, **kwargs: Any) -> dict[str, Any]:
    """format_text: the parameters are ``_execute_format_text``'s, plus ``verbose``. The
    result is shaped by ``_compact_format_evidence``."""
    return _compact_format_evidence(_execute_format_text(*args, **kwargs), verbose=verbose)


@functools.wraps(_execute_apply_style)
def execute_apply_style(*args: Any, verbose: bool = False, **kwargs: Any) -> dict[str, Any]:
    """apply_style: the parameters are ``_execute_apply_style``'s, plus ``verbose``. The
    result is shaped by ``_compact_format_evidence``."""
    return _compact_format_evidence(_execute_apply_style(*args, **kwargs), verbose=verbose)


# ---------------------------------------------------------------------------
# Tool: live_save (issue #106 WP-3: https://github.com/michaelrobertsutton/JennyStack/issues/106)
# ---------------------------------------------------------------------------


def execute_live_save(path: str) -> dict[str, Any]:
    """Ask the connected pane to save the live document (``document.save()``
    via the ``save`` op), then bridge back to a file-mode revision token.

    Word owns the file for the whole live session -- nothing this server
    writes directly -- so ``live_save`` is the one point where a live
    edit becomes visible on disk again. After the pane confirms
    ``saved: true``, this re-describes the pane (for its own
    ``"live:sha256:..."`` ``revision_after``) and separately computes
    ``projection.compute_revision(path)["token"]`` (the SAME token
    ``replace_body_markdown``/``replace_text``/etc. use in file mode) as
    ``file_revision``, so a caller that wants to keep working against the
    file-mode revision contract after a live session has one to pass as
    the next file-mode call's ``revision_before``.

    Structural verification (tables, whole-section rewrites -- rungs 3/4,
    which never go live per the plan) against the now-saved file is the
    caller's own job afterward, via the existing read tools
    (``read_document``/``find_sections``/``diff_body_vs_file``/etc.) --
    this tool only confirms the save itself, not the file's contents.
    Those read tools remain FILE-ONLY (issue #22 B3 does not add a live
    read path for them) -- when there is no local file at all (see
    below), that follow-up verification simply isn't available.

    Issue #22 B3: ``file_revision`` is ``None`` when *path* names no
    local file -- e.g. a SharePoint/OneDrive document with no local sync
    (Word's own save just went there, not to anything this server could
    read). This call no longer requires *path* to exist locally at all.

    Errors:
      LIVE_UNAVAILABLE   - no connected pane session for this document
      LIVE_DISCONNECTED  - the pane's socket closed, or the save timed out
      LIVE_OP_FAILED     - the pane replied ok=false to 'save'
      VERIFICATION_FAILED - the pane replied ok=true but did not report saved=true
    """
    session = live_write_mode.live_session_for(path)
    document_name = session.document_name

    try:
        result = session.request_threadsafe("save")
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc
    except LiveOpFailed as exc:
        raise _make_error(
            live_write_mode.classify_op_failed(exc), exc.message, {"pane_code": exc.code}
        ) from exc

    if not result.get("saved"):
        raise _make_error(
            ErrorCode.VERIFICATION_FAILED,
            "live_save did not verify: the pane replied without saved=true.",
            {"result": result},
        )

    try:
        describe_result = session.request_threadsafe("describe")
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc
    post_hash = describe_result["bodySha256"]

    # Issue #22 B3: must_exist=False -- Word's own save just went to
    # wherever the live document actually lives (a SharePoint/OneDrive
    # document with no local sync saves there, never to a path this
    # server could read). file_revision is None when there is nothing
    # local to compute it from, rather than this call raising right
    # after a save that just succeeded.
    resolved = paths.resolve_allowed_docx_path(path, must_exist=False)
    file_revision = projection.compute_revision(resolved)["token"] if resolved.is_file() else None

    evidence: dict[str, Any] = {
        "applied": True,
        "saved": True,
        "document_name": document_name,
        # issue #154: live_save is a LIVE operation (it goes through the
        # pane, not atomic_replace_docx_parts) -- "live", matching every
        # other live-mode evidence's write_mode key, never "file".
        "write_mode": "live",
        "revision_after": f"live:sha256:{post_hash}",
        "file_revision": file_revision,
    }
    evidence.update(live_write_mode.shape_counts(result))
    logged, _ = audit.append_audit(path=str(resolved), tool="live_save", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
