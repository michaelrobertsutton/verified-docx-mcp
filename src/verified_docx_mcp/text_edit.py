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
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .live import write_mode as live_write_mode
from .live.session import LiveDisconnected, LiveOpFailed, LiveStale
from .locate import LocateResult, locate
from .projection import W_NS, RunEvent

_XML_NS = "http://www.w3.org/XML/1998/namespace"


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Style allowlist (format_text) -- our own rPr model carries strike in
# addition to Google's bold/italic/underline (projection._run_properties
# already reads/writes all four), so this server's format_text supports one
# more field than the Google original.
# ---------------------------------------------------------------------------

_STYLE_ALLOWLIST = ("bold", "italic", "underline", "strike")
_BOOL_TOGGLE_TAGS = {"bold": "b", "italic": "i", "strike": "strike"}

_EXCERPT_RADIUS = 200


def _validate_style(style: Any) -> dict[str, bool]:
    if not isinstance(style, dict):
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "style must be an object mapping bold/italic/underline/strike to true/false",
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
    non_bool = {k: repr(v) for k, v in style.items() if type(v) is not bool}
    if non_bool:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"style values must be true/false booleans, got: {non_bool}",
            {"invalid_values": non_bool},
        )
    return style


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


def _style_matches(runs_before: list[list[dict[str, Any]]], style: dict[str, bool]) -> bool:
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


def _toggle_style(r_elem: Any, style: dict[str, bool]) -> None:
    rpr = None
    for child in r_elem:
        if projection._ln(child) == "rPr":
            rpr = child
            break
    if rpr is None:
        rpr = ET.Element(_w("rPr"))
        r_elem.insert(0, rpr)

    for field, value in style.items():
        if field == "underline":
            existing = None
            for child in rpr:
                if projection._ln(child) == "u":
                    existing = child
                    break
            if value:
                if existing is None:
                    ET.SubElement(rpr, _w("u"), {_w("val"): "single"})
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
                ET.SubElement(rpr, _w(tag))
            else:
                for key in list(existing.attrib):
                    del existing.attrib[key]
        elif existing is not None:
            rpr.remove(existing)


def _apply_style_to_run(r_elem: Any, style: dict[str, bool], *, track: _TrackContext | None) -> None:
    """_toggle_style, plus (when tracking) recording the PRE-change rPr in
    a w:rPrChange (issue #28 plan WP-07b-a: "formatting changes use
    w:rPrChange"). The snapshot is taken BEFORE _toggle_style mutates
    anything, per apply_rpr_change's own contract."""
    old_rpr = _own_rpr(r_elem)
    old_rpr_snapshot = copy.deepcopy(old_rpr) if (track and old_rpr is not None) else None
    _toggle_style(r_elem, style)
    if track:
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
    style: dict[str, bool],
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


def execute_replace_text_live(
    path: str,
    find: str,
    replace: str,
    expected_matches: int,
    *,
    revision_before: str | None = None,
    track_changes: bool = False,
) -> dict[str, Any]:
    """Live-mode ``replace_text`` (issue #106 WP-3): sends a ``replace``
    op over the pane's WSS ops channel instead of editing the .docx file.

    Flow: ``describe`` (record ``pre`` body hash + ``LIVE_STALE`` check)
    -> ``replace`` (with ``expected_body_sha256=pre`` so the session layer
    itself also re-checks staleness against the reply) -> verify
    (``post`` differs from ``pre``, and every match's ``after`` equals
    *replace* exactly) -> build the live evidence envelope
    (``live/write_mode.py``'s ``live_evidence``).

    ``force`` has no live-mode analogue (there is no tracked-change/
    comment-anchor override to force here -- Word owns the document, not
    this server) and is not accepted; a caller that wants a forced
    override uses ``write_mode="file"`` with the file's own ``force``
    instead.

    Raises ``LIVE_UNAVAILABLE``/``LIVE_DISCONNECTED``/``LIVE_STALE`` (see
    ``live/session.py``'s exceptions of the same names), ``ZERO_MATCH``/
    ``MATCH_COUNT_MISMATCH``/``LIVE_OP_FAILED`` (the pane refused the
    ``expected_matches`` gate -- see ``live/write_mode.py``'s
    ``classify_op_failed``), or ``VERIFICATION_FAILED`` (the pane's own
    read-back after the op did not confirm the intended change -- nothing
    to roll back in live mode, since Word, not this server, owns the
    file; the diagnostics say so explicitly).
    """
    if not find:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty")

    session = live_write_mode.live_session_for(path)
    document_name = session.document_name

    try:
        pre_hash = _live_describe(session, path, revision_before)
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    payload = {
        "find": find,
        "expected_matches": expected_matches,
        "replace": replace,
        "track_changes": track_changes,
    }
    try:
        result = session.request_threadsafe("replace", payload, expected_body_sha256=pre_hash)
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

    if not result.get("applied") or post_hash == pre_hash:
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

    before_text = "\n".join(m.get("before", "") for m in matches)
    after_text = "\n".join(m.get("after", "") for m in matches)

    return live_write_mode.live_evidence(
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


# ---------------------------------------------------------------------------
# Tool 1: replace_text
# ---------------------------------------------------------------------------


def execute_replace_text(
    path: str,
    find: str,
    replace: str,
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
) -> dict[str, Any]:
    mode = live_write_mode.resolve_write_mode(path, write_mode)
    if mode == "live":
        return execute_replace_text_live(
            path,
            find,
            replace,
            expected_matches,
            revision_before=revision_before,
            track_changes=track_changes,
        )

    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = locate(find, proj, expected_matches)
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

    post_revision = projection.compute_revision(resolved)
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
    style: dict[str, bool],
    expected_matches: int,
    *,
    revision_before: str | None = None,
    track_changes: bool = False,
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
    """
    style = _validate_style(style)
    if not find:
        raise _make_error(ErrorCode.INVALID_INPUT, "find must not be empty")

    session = live_write_mode.live_session_for(path)
    document_name = session.document_name

    try:
        pre_hash = _live_describe(session, path, revision_before)
    except LiveDisconnected as exc:
        raise _make_error(ErrorCode.LIVE_DISCONNECTED, str(exc)) from exc

    payload = {
        "find": find,
        "expected_matches": expected_matches,
        "bold": style.get("bold"),
        "italic": style.get("italic"),
        "underline": style.get("underline"),
        "track_changes": track_changes,
    }
    try:
        result = session.request_threadsafe("format", payload, expected_body_sha256=pre_hash)
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

    before_text = "\n".join(m.get("before", "") for m in matches)
    after_text = "\n".join(m.get("after", "") for m in matches)

    return live_write_mode.live_evidence(
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


def execute_format_text(
    path: str,
    find: str,
    style: dict[str, bool],
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
    write_mode: str = "auto",
) -> dict[str, Any]:
    style = _validate_style(style)
    mode = live_write_mode.resolve_write_mode(path, write_mode)
    if mode == "live":
        return execute_format_text_live(
            path,
            find,
            style,
            expected_matches,
            revision_before=revision_before,
            track_changes=track_changes,
        )

    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = locate(find, proj, expected_matches)
    own_author = resolve_author_name()
    _check_tracked_changes_guard(proj, locate_result, force, own_author)

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

    post_revision = projection.compute_revision(resolved)
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
# Scope limit, documented rather than silently mishandled: a CHARACTER
# style (w:type="character") applies to the matched run(s) exactly like
# format_text, including track_changes=True support (w:rPrChange, per
# WP-07b-a's own contract, which names rPrChange for "formatting
# changes"). A PARAGRAPH style (w:type="paragraph") applies to every
# paragraph CONTAINING a matched run (the whole paragraph, not just the
# matched substring -- w:pStyle has no sub-paragraph granularity) but
# does NOT support track_changes=True: WP-07b-a's contract defines
# w:rPrChange only, never a w:pPrChange-style paragraph-formatting
# tracked-change shape, and inventing one untested is out of this WP's
# scope -- apply_style(..., track_changes=True) on a paragraph style
# raises INVALID_INPUT naming this gap explicitly rather than silently
# applying it untracked or guessing at an unspecified w:pPrChange shape.
# ---------------------------------------------------------------------------

_SUPPORTED_STYLE_TYPES = frozenset({"paragraph", "character"})


def _set_run_style(r_elem: Any, style_id: str, *, track: _TrackContext | None) -> None:
    """Set/replace <w:rStyle w:val=style_id> on r_elem's own w:rPr
    (creating one if absent; w:rStyle must be rPr's FIRST child per the
    OOXML schema's fixed element order). Tracked the same way
    _apply_style_to_run tracks a boolean toggle: a w:rPrChange recording
    the PRE-change rPr, snapshotted before anything is mutated."""
    old_rpr = _own_rpr(r_elem)
    old_rpr_snapshot = copy.deepcopy(old_rpr) if (track and old_rpr is not None) else None
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
    if track:
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


def _apply_paragraph_style(p_elem: Any, style_id: str) -> None:
    """Set/replace <w:pStyle w:val=style_id> on p_elem's own w:pPr
    (creating one if absent; w:pStyle must be pPr's FIRST child). Never
    tracked -- see this section's own module-level scope-limit comment."""
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
    if existing is None:
        pstyle = ET.Element(_w("pStyle"), {_w("val"): style_id})
        ppr.insert(0, pstyle)
    else:
        existing.set(_w("val"), style_id)


def execute_apply_style(
    path: str,
    find: str,
    style_id: str,
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

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
    if style_type == "paragraph" and track_changes:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            "apply_style(track_changes=True) is not supported for a paragraph style: WP-07b-a's track_changes "
            "contract defines w:rPrChange for run/character formatting only, not a paragraph-level "
            "w:pPrChange-style tracked change. Apply the paragraph style directly (track_changes=False), or use "
            "a character style if a tracked change is required.",
            {"style_id": style_id, "type": style_type},
        )

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    locate_result: LocateResult = locate(find, proj, expected_matches)
    own_author = resolve_author_name()
    _check_tracked_changes_guard(proj, locate_result, force, own_author)

    before_first = locate_result.spans[0]
    before_excerpt = _excerpt(proj.text, before_first[0], before_first[1])
    runs_before = _collect_style_runs(proj, locate_result.spans)

    track = _TrackContext(document_root, author=own_author) if track_changes else None

    if style_type == "character":
        for start, end in locate_result.spans:
            atoms = _atoms_for_span(proj, start, end)
            _apply_named_style_span(start, end, style_id, atoms, track=track)
    else:
        parent_map = _build_parent_map(document_root)
        touched_ids: set[int] = set()
        touched_paragraphs: list[Any] = []
        for start, end in locate_result.spans:
            for _s, _e, event in _atoms_for_span(proj, start, end):
                p_elem = _enclosing_paragraph(parent_map, event.r_elem)
                if p_elem is not None and id(p_elem) not in touched_ids:
                    touched_ids.add(id(p_elem))
                    touched_paragraphs.append(p_elem)
        for p_elem in touched_paragraphs:
            _apply_paragraph_style(p_elem, style_id)

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
                    pstyle_val = None
                    for child in p_elem:
                        if projection._ln(child) == "pPr":
                            for gc in child:
                                if projection._ln(gc) == "pStyle":
                                    pstyle_val = projection._attr(gc, "val")
                    if pstyle_val != style_id:
                        raise ValueError(f"re-read paragraph does not carry w:pStyle val={style_id!r} (got {pstyle_val!r})")

    conflict_sweep = _serialize_and_write(resolved, document_root, raw_xml, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
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
