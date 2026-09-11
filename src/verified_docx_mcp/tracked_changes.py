# New for issue #28 WP-07. Partial GoogleDocs-MCP analogue: that server's
# list_open_items (server.py) and suggestions.py/comments.py define the
# RESPONSE SHAPE this module's list_open_items mirrors ("Google's shape" per
# the plan text) -- comments from Drive, pending_suggestions with
# {suggestion_id, kind, text, anchor_context} -- but the underlying source is
# completely different (Drive API comments + Docs API suggestedInsertionIds/
# suggestedDeletionIds there; word/comments.xml + w:ins/w:del here), and
# accept_tracked_changes/reject_tracked_changes have no Google-side analogue
# at all: a Docs API suggestion is resolved via batchUpdate's
# acceptSuggestion/... no such call exists in that API surface either --
# Google Docs suggestions are resolved through the Docs UI, not this MCP
# server. Both mutating tools here are new OOXML-tree surgery (unwrap a
# w:ins / remove a w:del to accept; remove a w:ins / unwrap a w:del back to
# w:t to reject), built the same way mutations.py/text_edit.py already do:
# guard -> compile -> atomic write -> post-read -> evidence -> audit.
"""list_open_items (read-only) and accept_tracked_changes/
reject_tracked_changes (mutating) -- issue #28 WP-07.

list_open_items returns `comments` (from word/comments.xml +
commentsExtended.xml's resolved flag, anchor text taken from the live
projection's own comment_ids tracking -- see projection.RunEvent) and
`pending_suggestions` (every w:ins/w:del in word/document.xml, with author/
date/text, in Google's {suggestion_id, kind, text, anchor_context} shape).
Scope limit, documented rather than silently mishandled: comment REPLY
threading (commentsExtended's parent/child linking) is not resolved here --
every comment reports reply_count=0/replies=[]; WP-08 owns the fixture
(tests/fixtures/comments/golden-comment.docx) built specifically to test
threading, and this WP does not touch it.

accept_tracked_changes/reject_tracked_changes both take an optional
`revision_ids` (a w:ins/w:del `w:id` list) -- omitted or None means every
tracked change in the document; a named id not present raises
REVISION_ID_NOT_FOUND with the ids that ARE available. Accept: a w:ins
unwraps (its content becomes ordinary live text); a w:del is removed
outright (the deletion becomes permanent). Reject: a w:ins is removed
outright (the insertion is undone); a w:del unwraps with every w:delText
renamed back to w:t (the deletion is undone, the text becomes live again).
Same guard/atomic-write/audit machinery as text_edit.py's tools.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection
from .errors import ErrorCode, _make_error
from .projection import W_NS


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


# ---------------------------------------------------------------------------
# Read-side: comments + pending suggestions
# ---------------------------------------------------------------------------


def _collect_text(elem: Any, tag_names: set[str]) -> str:
    parts: list[str] = []
    for node in elem.iter():
        if projection._ln(node) in tag_names:
            parts.append(node.text or "")
    return "".join(parts)


def _find_pending_suggestions(root: Any) -> list[dict[str, Any]]:
    """Every w:ins/w:del in *root* (any nesting depth -- a body paragraph,
    or one inside a table cell), each with the owning paragraph's own live
    text as `anchor_context` (Google's suggestions.py: "the full text of
    the paragraph that contains the suggestion, assembled from all runs, so
    the change is locatable")."""
    suggestions: list[dict[str, Any]] = []

    def walk(elem: Any, current_para: Any | None) -> None:
        for child in elem:
            tag = projection._ln(child)
            next_para = child if tag == "p" else current_para
            if tag == "ins":
                suggestions.append(
                    {
                        "suggestion_id": projection._attr(child, "id") or "",
                        "kind": "insertion",
                        "text": _collect_text(child, {"t"}),
                        "author": projection._attr(child, "author"),
                        "date": projection._attr(child, "date"),
                        "anchor_context": _collect_text(next_para, {"t", "delText"}) if next_para is not None else "",
                    }
                )
                continue  # do not descend further for suggestion-collection
            if tag == "del":
                suggestions.append(
                    {
                        "suggestion_id": projection._attr(child, "id") or "",
                        "kind": "deletion",
                        "text": _collect_text(child, {"delText"}),
                        "author": projection._attr(child, "author"),
                        "date": projection._attr(child, "date"),
                        "anchor_context": _collect_text(next_para, {"t", "delText"}) if next_para is not None else "",
                    }
                )
                continue
            walk(child, next_para)

    walk(root, None)
    return suggestions


def _parse_comments(local_path: Path, proj: projection.Projection) -> list[dict[str, Any]]:
    import zipfile

    with zipfile.ZipFile(local_path) as zf:
        names = set(zf.namelist())
        if "word/comments.xml" not in names:
            return []
        comments_root = ET.fromstring(zf.read("word/comments.xml"))
        ext_by_paraid: dict[str, dict[str, Any]] = {}
        if "word/commentsExtended.xml" in names:
            ext_root = ET.fromstring(zf.read("word/commentsExtended.xml"))
            for child in ext_root:
                if projection._ln(child) != "commentEx":
                    continue
                pid = projection._attr(child, "paraId")
                if pid:
                    ext_by_paraid[pid] = {
                        "done": projection._attr(child, "done"),
                        "parent": projection._attr(child, "parent"),
                    }
        # New for the WP-08 interop fix: durableId per paraId, from
        # word/commentsIds.xml -- comments.py's add_anchored_comment/
        # get_comment_thread already key everything on durableId (the plan
        # text: "expose durableId as comment_id"), but this function used to
        # expose the raw w:id instead, so "list_open_items, then
        # get_comment_thread(that id)" -- the exact chain WP-11b's
        # resolve-gdoc-comments-on-docx will do -- failed outright. Fixed
        # here (where the missing lookup lives), not in comments.py.
        durable_by_paraid: dict[str, str] = {}
        if "word/commentsIds.xml" in names:
            ids_root = ET.fromstring(zf.read("word/commentsIds.xml"))
            for child in ids_root:
                if projection._ln(child) != "commentId":
                    continue
                pid = projection._attr(child, "paraId")
                did = projection._attr(child, "durableId")
                if pid and did:
                    durable_by_paraid[pid] = did

    # Anchor ("quoted") text per comment w:id, from the live projection's
    # own comment_ids tracking (projection.RunEvent.comment_ids -- see
    # that module's WP-06 addition; commentRangeStart/End/commentReference
    # in document.xml carry the w:id, not the durableId, so this lookup
    # stays keyed by w:id even though the RETURNED comment_id below is the
    # durableId).
    anchor_by_id: dict[str, list[str]] = {}
    for event in proj.events:
        if isinstance(event, projection.RunEvent):
            for cid in event.comment_ids:
                anchor_by_id.setdefault(cid, []).append(event.text)

    comments: list[dict[str, Any]] = []
    for c in comments_root:
        if projection._ln(c) != "comment":
            continue
        w_id = projection._attr(c, "id") or ""
        author = projection._attr(c, "author")
        date = projection._attr(c, "date")
        content_parts: list[str] = []
        own_para_id: str | None = None
        for p in c:
            if projection._ln(p) != "p":
                continue
            if own_para_id is None:
                own_para_id = projection._attr(p, "paraId")
            content_parts.append(_collect_text(p, {"t"}))
        ext = ext_by_paraid.get(own_para_id or "", {})
        # durableId is the durable, opaque identifier (Google's own shape:
        # a comment id independent of the document's own w:id numbering,
        # which is only ever max-existing-plus-1 and therefore not
        # stable/round-trippable across edits). Falls back to the raw
        # w:id only if commentsIds.xml is missing despite comments.xml
        # existing (malformed input, not a shape any tool in this repo
        # produces).
        durable_id = durable_by_paraid.get(own_para_id or "", w_id)
        comments.append(
            {
                "comment_id": durable_id,
                "w_id": w_id,
                "content": "".join(content_parts),
                "resolved": ext.get("done") == "1",
                "reply_count": 0,  # see module docstring: threading is WP-08 scope
                "replies": [],
                "quoted_text": "".join(anchor_by_id.get(w_id, [])),
                "author": author,
                "created_time": date,
                "modified_time": "",
                "scope": "document",
            }
        )
    return comments


def execute_list_open_items(path: str) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    # Read-only tool: never refuses for a lock (core/document-backend-
    # protocol.md §4) -- reads a validated snapshot instead, same as every
    # WP-03 read tool. Deferred import: server.py imports this module to
    # register the tools below, so a module-level import would be circular.
    from . import server as _server

    local_path, is_temp = _server._read_local_copy(resolved)
    try:
        proj = projection.project_part(local_path)
        document_root, _raw = mutations._load_document(local_path)
        suggestions = _find_pending_suggestions(document_root)
        comments = _parse_comments(local_path, proj)
        return {"path": str(resolved), "comments": comments, "pending_suggestions": suggestions}
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Write-side: accept / reject
# ---------------------------------------------------------------------------


def _remove_if_present(parent: Any, elem: Any) -> None:
    for child in list(parent):
        if child is elem:
            parent.remove(elem)
            return


def _unwrap(parent: Any, elem: Any) -> None:
    """Replace *elem* with its own children, in place, at *elem*'s current
    index in *parent* -- used both to accept a w:ins (its content becomes
    ordinary live text) and to reject a w:del (see _reject_one)."""
    children = list(elem)
    idx = None
    for i, child in enumerate(parent):
        if child is elem:
            idx = i
            break
    if idx is None:
        return
    parent.remove(elem)
    for offset, child in enumerate(children):
        parent.insert(idx + offset, child)


def _collect_ins_del(root: Any) -> list[tuple[Any, Any, str, str]]:
    """(parent, elem, kind, id) for every w:ins/w:del in *root*, any nesting
    depth, in document order. A read-only pass -- callers apply mutations in
    a SEPARATE pass afterward so the parent/child structure this collects
    against is not disturbed mid-walk."""
    out: list[tuple[Any, Any, str, str]] = []

    def walk(elem: Any) -> None:
        for child in list(elem):
            tag = projection._ln(child)
            if tag in ("ins", "del"):
                rid = projection._attr(child, "id") or ""
                out.append((elem, child, tag, rid))
            walk(child)

    walk(root)
    return out


# ---------------------------------------------------------------------------
# Write-side helpers shared with text_edit.py/mutations.py (WP-07b-a): w:id
# allocation, w:ins/w:del wrapping, w:rPrChange, and the author-aware
# "foreign revision" check the TRACKED_CHANGES_PRESENT guard needs.
# ---------------------------------------------------------------------------


class RevisionIdAllocator:
    """Hands out w:id values for new w:ins/w:del/w:rPrChange elements,
    each guaranteed above every id already in *document_root* (issue #28
    plan WP-07b-a: "w:id values allocated above the package maximum") and
    never repeated within one guarded call, without re-scanning the tree
    per id."""

    def __init__(self, document_root: Any) -> None:
        body = mutations._find_body(document_root)
        self._next = 1
        for _parent, _elem, _kind, rid in _collect_ins_del(body):
            if rid.isdigit():
                self._next = max(self._next, int(rid) + 1)

    def allocate(self) -> str:
        rid = str(self._next)
        self._next += 1
        return rid


def current_revision_date() -> str:
    """UTC timestamp in the w:date attribute's own format (the same shape
    every fixture's real Word-authored w:ins/w:del already carries, e.g.
    tracked.docx's "2026-09-10T17:17:00Z")."""
    from datetime import UTC, datetime

    return datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


class TrackContext:
    """Bundles what a track_changes=True call needs -- author, date, an
    id allocator, and the ids created so far -- so every call site in
    text_edit.py/mutations.py takes one optional argument instead of
    four. Shared here (rather than defined separately in each) so both
    modules use the identical author/date/id-allocation logic; each
    imports this module directly (mutations.py only inside a function
    body, the same deferred-import pattern _guard_before_write already
    uses for server.py, since this module itself imports mutations.py at
    the top level -- importing back would cycle)."""

    def __init__(self, document_root: Any, *, author: str | None = None) -> None:
        # *author* SHOULD always be passed explicitly by the caller (the
        # same value it already resolved for the TRACKED_CHANGES_PRESENT
        # guard's own-author exclusion, e.g. text_edit.py's own_author) --
        # resolving it independently here would risk the two diverging
        # (a config file changing between the two calls; a test patching
        # only one of the two import sites) and stamping a DIFFERENT
        # author on the write than the guard just excluded. Falling back
        # to a fresh resolve_author_name() call only covers a caller that
        # does not have one on hand yet.
        if author is None:
            from .author import resolve_author_name

            author = resolve_author_name()
        self.author = author
        self.date = current_revision_date()
        self.allocator = RevisionIdAllocator(document_root)
        self.revision_ids: list[str] = []

    def __bool__(self) -> bool:
        # A TrackContext is only ever constructed when track_changes=True;
        # always truthy -- exists so `if track:` reads naturally at call
        # sites without a separate track_changes bool threaded everywhere.
        return True

    def next_id(self) -> str:
        """Allocate one new w:id AND record it in self.revision_ids in the
        same call -- the only entry point call sites should use (never
        self.allocator.allocate() directly), so an id can never be
        allocated without also landing in the evidence's revision_ids."""
        rid = self.allocator.allocate()
        self.revision_ids.append(rid)
        return rid


def wrap_insertion(r_elem: Any, *, rid: str, author: str, date: str) -> Any:
    """Wrap *r_elem* in a new <w:ins id author date> element and return it
    -- the caller is responsible for putting the returned element where
    r_elem used to live."""
    ins_elem = ET.Element(_w("ins"), {_w("id"): rid, _w("author"): author, _w("date"): date})
    ins_elem.append(r_elem)
    return ins_elem


def wrap_deletion(r_elem: Any, *, rid: str, author: str, date: str) -> Any:
    """Like wrap_insertion, but for w:del -- *r_elem* must already have had
    convert_t_to_deltext applied (OOXML requires deleted text to live in
    w:delText, never w:t, inside a w:del)."""
    del_elem = ET.Element(_w("del"), {_w("id"): rid, _w("author"): author, _w("date"): date})
    del_elem.append(r_elem)
    return del_elem


def convert_t_to_deltext(r_elem: Any) -> None:
    """Rename every w:t inside *r_elem* to w:delText, in place -- the
    inverse of reject_tracked_changes' own delText->t rename (_reject_one
    above)."""
    for node in r_elem.iter():
        if projection._ln(node) == "t":
            node.tag = _w("delText")


def apply_rpr_change(rpr_elem: Any, *, old_rpr_elem: Any | None, rid: str, author: str, date: str) -> None:
    """Append <w:rPrChange id author date><w:rPr>...</w:rPr></w:rPrChange>
    to *rpr_elem*, recording the run's PRE-change formatting (issue #28
    plan WP-07b-a: "formatting changes use w:rPrChange"). *old_rpr_elem*
    must be a snapshot the caller took BEFORE applying the new style (a
    deep copy, or None if the run previously had no w:rPr at all, in which
    case the previous state was "no explicit properties" and an empty
    <w:rPr/> is recorded) -- this function does not itself snapshot
    anything; it only records what it is given.
    """
    change = ET.SubElement(rpr_elem, _w("rPrChange"), {_w("id"): rid, _w("author"): author, _w("date"): date})
    if old_rpr_elem is not None:
        import copy

        change.append(copy.deepcopy(old_rpr_elem))
    else:
        ET.SubElement(change, _w("rPr"))


def foreign_crosses_revision(proj: projection.Projection, start: int, end: int, own_author: str) -> bool:
    """True if [start, end) overlaps a w:ins authored by anyone OTHER than
    *own_author* -- the TRACKED_CHANGES_PRESENT guard's own-author
    exclusion (issue #28 plan WP-07b-a: "refusals must exclude revisions
    the server itself authored under the configured author, or a second
    proposed edit deadlocks on the first one's tracked change"). A w:del's
    own content never reaches here at all (excluded from the projection
    entirely -- see projection.py's module docstring), so this only ever
    concerns a live w:ins the match's span overlaps.
    """
    run_events = [e for e in proj.events if isinstance(e, projection.RunEvent)]
    for event, (s, e, _pr, _rr) in zip(run_events, proj.offset_map):
        if s < end and e > start and event.in_revision and event.revision_author != own_author:
            return True
    return False


def wrap_all_runs(elements: list[Any], wrap_fn) -> None:
    """For every <w:r> found anywhere inside *elements* (any nesting depth
    -- a table cell's own paragraphs included), replace it in its own
    immediate parent with wrap_fn(r_elem), in place. Used by mutations.py's
    markdown-mutation tools (WP-07b-a) to block-track a whole-body/section
    replace or an append: wrap_fn is wrap_deletion (after
    convert_t_to_deltext) for OLD content being marked deleted, or
    wrap_insertion for NEW content being marked inserted.

    Does NOT descend into an existing w:ins/w:del -- content already
    tracked (or a force=True-accepted pre-existing hazard) is left exactly
    as it is, never double-wrapped.
    """

    def walk(parent: Any) -> None:
        for child in list(parent):
            tag = projection._ln(child)
            if tag == "r":
                idx = None
                for i, c in enumerate(parent):
                    if c is child:
                        idx = i
                        break
                if idx is None:
                    continue
                parent.remove(child)
                parent.insert(idx, wrap_fn(child))
            elif tag in ("ins", "del"):
                continue
            else:
                walk(child)

    for el in elements:
        walk(el)


def mark_elements_deleted(elements: list[Any], track: Any) -> None:
    """Mark every LIVE run inside *elements* as deleted for a tracked
    write, handling a run that is ALREADY inside a pending w:ins or w:del
    correctly rather than via wrap_all_runs' generic "leave existing
    tracking alone" rule:

    - A run already inside a w:del is already gone from the live
      projection; left exactly as it is (re-deleting an already-deleted
      run is a no-op, not a new revision).
    - A run already inside a w:ins (a still-pending insertion nobody has
      accepted yet -- possibly from an EARLIER track_changes=True call
      this same server made) is REMOVED OUTRIGHT along with its w:ins
      wrapper, not wrapped in a second, redundant w:del: cancelling an
      insertion that was never accepted needs no deletion record (this is
      exactly what reject_tracked_changes already does to a w:ins, for
      the identical reason). Skipping this case (as a naive "never touch
      existing ins/del" rule would) leaves the earlier insertion visible
      in the live projection forever, silently accumulating alongside
      each new write -- found via this WP's own
      test_second_tracked_write_over_own_prior_tracked_change_does_not_deadlock,
      not merely theorized.
    - Every other run (ordinary, untracked live content) is wrapped in a
      NEW w:del, exactly like wrap_all_runs' own del path.
    """

    def walk(parent: Any) -> None:
        for child in list(parent):
            tag = projection._ln(child)
            if tag == "del":
                continue
            if tag == "ins":
                _remove_if_present(parent, child)
                continue
            if tag == "r":
                idx = None
                for i, c in enumerate(parent):
                    if c is child:
                        idx = i
                        break
                if idx is None:
                    continue
                parent.remove(child)
                convert_t_to_deltext(child)
                parent.insert(idx, wrap_deletion(child, rid=track.next_id(), author=track.author, date=track.date))
            else:
                walk(child)

    for el in elements:
        walk(el)


def _resolve_targets(
    document_root: Any, revision_ids: list[str] | None
) -> list[tuple[Any, Any, str, str]]:
    body = mutations._find_body(document_root)
    all_targets = _collect_ins_del(body)
    if revision_ids is None:
        return all_targets
    wanted = set(revision_ids)
    available = {rid for _p, _e, _k, rid in all_targets}
    missing = sorted(wanted - available)
    if missing:
        raise _make_error(
            ErrorCode.REVISION_ID_NOT_FOUND,
            f"revision id(s) not found: {missing}",
            {"missing_ids": missing, "available_ids": sorted(available)},
        )
    return [(p, e, k, rid) for p, e, k, rid in all_targets if rid in wanted]


def _accept_one(parent: Any, elem: Any, kind: str) -> None:
    if kind == "ins":
        _unwrap(parent, elem)
    else:  # "del"
        _remove_if_present(parent, elem)


def _reject_one(parent: Any, elem: Any, kind: str) -> None:
    if kind == "ins":
        _remove_if_present(parent, elem)
    else:  # "del"
        for node in elem.iter():
            if projection._ln(node) == "delText":
                node.tag = _w("t")
        _unwrap(parent, elem)


def _run_accept_or_reject(
    path: str,
    revision_ids: list[str] | None,
    *,
    revision_before: str | None,
    force: bool,
    tool_name: str,
    apply_one,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    before_proj = projection.project_document_root(document_root)
    before_text = before_proj.text

    targets = _resolve_targets(document_root, revision_ids)
    processed_ids = [rid for _p, _e, _k, rid in targets]
    for parent, elem, kind, _rid in targets:
        apply_one(parent, elem, kind)

    after_proj = projection.project_document_root(document_root)
    intended_after_text = after_proj.text

    def _post_verify(written_path: Path) -> None:
        actual_text = projection.read_document_text(written_path)
        diff = mutations._diff_modulo_whitespace(intended_after_text, actual_text)
        if diff:
            raise ValueError(f"re-read document does not match the intended text modulo whitespace: {diff}")

    document_decls = mutations._capture_source_namespaces(raw_xml)
    new_xml_bytes = mutations._serialize_xml(document_root, document_decls)
    mutations.atomic_replace_docx_parts(resolved, {projection.DEFAULT_PART: new_xml_bytes}, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": len(processed_ids),
        "rung": "all" if revision_ids is None else "by_id",
        "before": before_text,
        "after": intended_after_text,
        "revision_before": pre_revision["token"],
        "revision_after": post_revision["token"],
        "audit_logged": False,
        "revision_ids": processed_ids,
    }
    logged, _ = audit.append_audit(path=str(resolved), tool=tool_name, evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


def execute_accept_tracked_changes(
    path: str, revision_ids: list[str] | None = None, *, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    return _run_accept_or_reject(
        path, revision_ids, revision_before=revision_before, force=force,
        tool_name="accept_tracked_changes", apply_one=_accept_one,
    )


def execute_reject_tracked_changes(
    path: str, revision_ids: list[str] | None = None, *, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    return _run_accept_or_reject(
        path, revision_ids, revision_before=revision_before, force=force,
        tool_name="reject_tracked_changes", apply_one=_reject_one,
    )
