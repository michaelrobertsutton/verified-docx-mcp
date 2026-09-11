# New for issue #28 WP-08. No GoogleDocs-MCP analogue to lift wholesale:
# that server's add_anchored_comment (comments.py, server.py) creates a
# Drive-API comment via create_comment's quotedFileContent -- a single API
# call, no local package parts to construct. This backend has no comment
# API at all: a docx comment is five interlocking OOXML parts
# (comments.xml, commentsExtended.xml, commentsIds.xml,
# commentsExtensible.xml, people.xml) plus three anchor elements in
# word/document.xml (commentRangeStart/End, commentReference) plus a
# [Content_Types].xml override and a word/_rels/document.xml.rels
# relationship PER NEW PART -- all of which this module builds directly.
# What IS lifted, in spirit: the "locate the quote, verify it exists,
# refuse with a typed error and near-miss diagnostics if it doesn't"
# shape of Google's own add_anchored_comment (its own docstring: "quote
# must exist in the tab ... locates it via the same normalization ladder
# as replace_text") -- this module reuses locate.py/text_edit.py's own
# run-splitting machinery for exactly that reason, rather than building a
# second one.
"""add_anchored_comment (mutating) and get_comment_thread (read-only) --
issue #28 WP-08.

Namespace declarations: every new part below (and word/document.xml,
already-loaded) reuses the TARGET DOCUMENT's own captured root namespace
map (mutations._capture_source_namespaces on its document.xml) via
_serialize_xml's original_decls parameter -- the same 30+-namespace
boilerplate a real Word install stamps on every part it creates (verified
directly against tests/fixtures/comments/golden-comment.docx: its five
comment-related parts declare the IDENTICAL namespace set document.xml
itself does, plus one extra -- "cr" -- on commentsExtensible.xml only).
This is deliberately not a hand-picked minimal namespace set: reusing the
one real namespace boilerplate already sitting in the package (a) matches
what genuine Word output looks like without inventing anything, and (b)
automatically keeps mc:Ignorable's named prefixes declared, so opc_valid
never has anything to catch here -- this is the exact class of bug PR #3
found (a namespace declared-but-then-silently-dropped), approached from
the opposite direction: never CONSTRUCT a part whose mc:Ignorable outruns
its own declarations in the first place.

Id allocation (issue #28 plan WP-08): w:id (commentRangeStart/End,
commentReference, w:comment) = one past the highest existing w:comment
w:id in word/comments.xml (0 if the document has no comments yet, so ids
start at 0 like a fresh Word document's own first comment does -- verified
against golden's own id=0/1/2 sequence). paraId/durableId = a random
8-hex-uppercase string below 0x80000000, retried until it collides with
neither an existing paraId/durableId in the package nor one already
allocated earlier in this same call.
"""

from __future__ import annotations

import copy
import secrets
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, mutations, paths, projection, text_edit, tracked_changes
from .author import resolve_author_name
from .errors import ErrorCode, _make_error
from .locate import LocateResult, locate
from .projection import DEFAULT_PART, W_NS

_XML_NS = "http://www.w3.org/XML/1998/namespace"
_PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

# ---------------------------------------------------------------------------
# Namespaces the five comment parts use.
# ---------------------------------------------------------------------------

NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"

_MC_IGNORABLE_TOKENS = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"
_MC_IGNORABLE_TOKENS_CEX = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl cr w16du wp14"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _w15(name: str) -> str:
    return f"{{{NS_W15}}}{name}"


def _w16cid(name: str) -> str:
    return f"{{{NS_W16CID}}}{name}"


def _w16cex(name: str) -> str:
    return f"{{{NS_W16CEX}}}{name}"


def _mc(name: str) -> str:
    return f"{{{_MC_NS}}}{name}"


# Content types / relationship types for the five parts, plus their part
# names and default (empty-package) filenames -- one shared table so
# every "does this part exist yet / does it need a fresh Relationship +
# Override" check below reads from one place.
_COMMENTS_PART = "word/comments.xml"
_COMMENTS_EXT_PART = "word/commentsExtended.xml"
_COMMENTS_IDS_PART = "word/commentsIds.xml"
_COMMENTS_CEX_PART = "word/commentsExtensible.xml"
_PEOPLE_PART = "word/people.xml"

_PART_CONTENT_TYPES = {
    _COMMENTS_PART: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    _COMMENTS_EXT_PART: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
    _COMMENTS_IDS_PART: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
    _COMMENTS_CEX_PART: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
    _PEOPLE_PART: "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
}

_PART_REL_TYPES = {
    _COMMENTS_PART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    _COMMENTS_EXT_PART: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    _COMMENTS_IDS_PART: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    _COMMENTS_CEX_PART: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
    _PEOPLE_PART: "http://schemas.microsoft.com/office/2011/relationships/people",
}


# ---------------------------------------------------------------------------
# Id allocation
# ---------------------------------------------------------------------------


def _next_comment_id(comments_root: Any | None) -> int:
    if comments_root is None:
        return 0
    best = -1
    for child in comments_root:
        if projection._ln(child) != "comment":
            continue
        raw = projection._attr(child, "id")
        if raw is not None and raw.isdigit():
            best = max(best, int(raw))
    return best + 1


def _existing_hex_ids(*roots: Any | None) -> set[str]:
    found: set[str] = set()
    for root in roots:
        if root is None:
            continue
        for el in root.iter():
            for attr_name in ("paraId", "durableId"):
                val = projection._attr(el, attr_name)
                if val:
                    found.add(val.upper())
    return found


def _new_hex_id(taken: set[str]) -> str:
    for _ in range(1000):
        candidate = f"{secrets.randbelow(0x80000000):08X}"
        if candidate not in taken:
            taken.add(candidate)
            return candidate
    raise _make_error(ErrorCode.INVALID_INPUT, "could not allocate a unique paraId/durableId after 1000 attempts", {})


# ---------------------------------------------------------------------------
# Part loading (parse if present, else build a fresh, empty root)
# ---------------------------------------------------------------------------


def _read_part_or_none(zf: zipfile.ZipFile, part_name: str) -> bytes | None:
    return zf.read(part_name) if part_name in zf.namelist() else None


def _root_or_new(data: bytes | None, tag: str) -> tuple[Any, bool]:
    """(root, created) -- parses *data* if present, else builds a fresh
    <tag mc:Ignorable="..."/> root (created=True)."""
    if data is not None:
        return ET.fromstring(data), False
    root = ET.Element(tag)
    return root, True


class _CommentParts:
    """The five comment-related parts' (root, created) pairs plus the raw
    rels/[Content_Types].xml bytes -- loaded once, shared by
    add_anchored_comment (WP-08) and reply_to_comment/resolve_comment
    (WP-09), all of which need the same five parts open for editing."""

    def __init__(self, resolved: Path) -> None:
        with zipfile.ZipFile(resolved) as zf:
            names = set(zf.namelist())
            comments_bytes = _read_part_or_none(zf, _COMMENTS_PART)
            ext_bytes = _read_part_or_none(zf, _COMMENTS_EXT_PART)
            ids_bytes = _read_part_or_none(zf, _COMMENTS_IDS_PART)
            cex_bytes = _read_part_or_none(zf, _COMMENTS_CEX_PART)
            people_bytes = _read_part_or_none(zf, _PEOPLE_PART)
            rels_path = projection._rels_path_for(DEFAULT_PART)
            self.rels_bytes = zf.read(rels_path) if rels_path in names else None
            self.ct_bytes = zf.read("[Content_Types].xml")

        self.comments_root, self.comments_created = _root_or_new(comments_bytes, _w("comments"))
        self.ext_root, self.ext_created = _root_or_new(ext_bytes, _w15("commentsEx"))
        self.ids_root, self.ids_created = _root_or_new(ids_bytes, _w16cid("commentsIds"))
        self.cex_root, self.cex_created = _root_or_new(cex_bytes, _w16cex("commentsExtensible"))
        self.people_root, self.people_created = _root_or_new(people_bytes, _w15("people"))

    def build_overrides(self, *, document_root: Any, raw_xml: bytes) -> dict[str, bytes]:
        return _build_overrides(
            document_root=document_root,
            raw_xml=raw_xml,
            comments_root=self.comments_root,
            comments_created=self.comments_created,
            ext_root=self.ext_root,
            ext_created=self.ext_created,
            ids_root=self.ids_root,
            ids_created=self.ids_created,
            cex_root=self.cex_root,
            cex_created=self.cex_created,
            people_root=self.people_root,
            people_created=self.people_created,
            rels_bytes=self.rels_bytes,
            ct_bytes=self.ct_bytes,
        )

    def para_id_for_w_id(self, w_id: str) -> str | None:
        for c in self.comments_root:
            if projection._ln(c) != "comment" or projection._attr(c, "id") != w_id:
                continue
            for p in c:
                if projection._ln(p) == "p":
                    return projection._attr(p, "paraId")
        return None

    def w_id_for_durable_id(self, durable_id: str) -> str | None:
        para_id = None
        for child in self.ids_root:
            if projection._ln(child) == "commentId" and projection._attr(child, "durableId") == durable_id:
                para_id = projection._attr(child, "paraId")
                break
        if para_id is None:
            return None
        for c in self.comments_root:
            if projection._ln(c) != "comment":
                continue
            for p in c:
                if projection._ln(p) == "p" and projection._attr(p, "paraId") == para_id:
                    return projection._attr(c, "id")
        return None


# ---------------------------------------------------------------------------
# add_anchored_comment
# ---------------------------------------------------------------------------


def execute_add_anchored_comment(
    path: str,
    quote: str,
    text: str,
    expected_matches: int,
    *,
    revision_before: str | None = None,
    force: bool = False,
) -> dict[str, Any]:
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)
    locate_result: LocateResult = locate(quote, proj, expected_matches)

    author = resolve_author_name()
    date = tracked_changes.current_revision_date()

    parts = _CommentParts(resolved)
    next_comment_id = _next_comment_id(parts.comments_root if not parts.comments_created else None)
    taken_hex = _existing_hex_ids(parts.ext_root, parts.ids_root)

    before_excerpt = text_edit._excerpt(proj.text, locate_result.spans[0][0], locate_result.spans[0][1])

    comment_ids_created: list[str] = []
    for start, end in locate_result.spans:
        cid = str(next_comment_id)
        next_comment_id += 1
        para_id = _new_hex_id(taken_hex)
        durable_id = _new_hex_id(taken_hex)
        comment_ids_created.append(durable_id)

        _insert_anchor(proj, start, end, comment_id=cid)
        _append_comment(parts.comments_root, comment_id=cid, para_id=para_id, author=author, date=date, text=text)
        _append_comment_ex(parts.ext_root, para_id=para_id)
        _append_comment_id(parts.ids_root, para_id=para_id, durable_id=durable_id)
        _append_comment_extensible(parts.cex_root, durable_id=durable_id, date_utc=date)

    _ensure_person(parts.people_root, author=author)

    # Post-write assertion (issue #28 plan WP-08: "Verify after write that
    # the range brackets the quoted span"): re-project the MUTATED live
    # tree (not yet written to disk) and confirm every comment_ids-tagged
    # run for the newly created ids reconstructs back to the quote,
    # modulo whitespace -- the same tolerance _diff_modulo_whitespace uses
    # elsewhere in this repo.
    mutated_proj = projection.project_document_root(document_root)
    _assert_anchors_bracket_quote(mutated_proj, quote)

    def _post_verify(written_path: Path) -> None:
        post_proj = projection.project_part(written_path)
        _assert_anchors_bracket_quote(post_proj, quote)

    overrides = parts.build_overrides(document_root=document_root, raw_xml=raw_xml)
    mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_excerpt = before_excerpt  # a comment never changes document text

    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": locate_result.match_count,
        "rung": locate_result.rung,
        "before": before_excerpt,
        "after": after_excerpt,
        "revision_before": pre_revision["token"],
        "revision_after": post_revision["token"],
        "audit_logged": False,
        "comment_ids": comment_ids_created,
    }
    if len(comment_ids_created) == 1:
        # issue #28 plan WP-08: "expose durableId as comment_id" -- the
        # singular convenience key for the common (single-match) case.
        evidence["comment_id"] = comment_ids_created[0]
    logged, _ = audit.append_audit(path=str(resolved), tool="add_anchored_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# document.xml anchor insertion: commentRangeStart/End + a commentReference
# run, positioned via the SAME run-splitting primitives text_edit.py's
# replace_text/format_text already use (_atoms_for_span, _own_rpr,
# _index_of, _set_node_text, _build_run) -- unlike w:ins/w:del, these
# markers are self-closing (they never WRAP content), so no run needs to
# be removed or replaced; a boundary run only ever needs splitting to
# expose a clean insertion point at the quote's own start/end offset.
# ---------------------------------------------------------------------------


def _build_comment_range_marker(tag: str, comment_id: str) -> Any:
    return ET.Element(_w(tag), {_w("id"): comment_id})


def _build_comment_reference_run(comment_id: str) -> Any:
    run = ET.Element(_w("r"))
    rpr = ET.SubElement(run, _w("rPr"))
    ET.SubElement(rpr, _w("rStyle"), {_w("val"): "CommentReference"})
    # Fixed convention verified against the golden fixture's own
    # commentReference runs (both id=0 and id=2): 24 half-points (12pt),
    # not the CommentReference style's own 16-half-point (8pt) default --
    # Word overrides it to match the surrounding body text's size. Not
    # derived from the target paragraph's own inherited size (a
    # documented simplification, matching the golden's own literal
    # values rather than computing them).
    ET.SubElement(rpr, _w("sz"), {_w("val"): "24"})
    ET.SubElement(rpr, _w("szCs"), {_w("val"): "24"})
    ET.SubElement(run, _w("commentReference"), {_w("id"): comment_id})
    return run


def _insert_anchor(proj: projection.Projection, start: int, end: int, *, comment_id: str) -> None:
    atoms = text_edit._atoms_for_span(proj, start, end)
    if not atoms:
        raise _make_error(ErrorCode.INVALID_INPUT, "matched quote has no overlapping run content", {"start": start, "end": end})

    start_marker = _build_comment_range_marker("commentRangeStart", comment_id)
    end_marker = _build_comment_range_marker("commentRangeEnd", comment_id)
    ref_run = _build_comment_reference_run(comment_id)

    first_s, first_e, first_event = atoms[0]
    last_s, last_e, last_event = atoms[-1]

    def _split_and_anchor(event: Any, s: int, e: int, *, want_start: bool, want_end: bool) -> None:
        r_elem = event.r_elem
        parent = event.parent_elem
        if r_elem is None or parent is None:
            raise _make_error(ErrorCode.INVALID_INPUT, "matched run has no live element to anchor", {"start": start, "end": end})

        local_start = max(start, s) - s
        local_end = min(end, e) - s
        full_text = event.text
        before_text = full_text[:local_start]
        after_text = full_text[local_end:]
        matched_text = full_text[local_start:local_end]
        own_rpr = text_edit._own_rpr(r_elem)
        idx = text_edit._index_of(parent, r_elem)

        if not before_text and not after_text:
            start_idx, end_idx = idx, idx + 1
        elif before_text and not after_text:
            text_edit._set_node_text(event, before_text)
            matched_run = text_edit._build_run(matched_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            parent.insert(idx + 1, matched_run)
            start_idx, end_idx = idx + 1, idx + 2
        elif not before_text and after_text:
            text_edit._set_node_text(event, matched_text)
            after_run = text_edit._build_run(after_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            parent.insert(idx + 1, after_run)
            start_idx, end_idx = idx, idx + 1
        else:
            text_edit._set_node_text(event, before_text)
            matched_run = text_edit._build_run(matched_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            after_run = text_edit._build_run(after_text, copy.deepcopy(own_rpr) if own_rpr is not None else None)
            parent.insert(idx + 1, matched_run)
            parent.insert(idx + 2, after_run)
            start_idx, end_idx = idx + 1, idx + 2

        if want_start:
            parent.insert(start_idx, start_marker)
            end_idx += 1  # only matters when want_end also fires on this same atom (a single-atom quote)
        if want_end:
            parent.insert(end_idx, end_marker)
            parent.insert(end_idx + 1, ref_run)

    if first_event is last_event:
        _split_and_anchor(first_event, first_s, first_e, want_start=True, want_end=True)
    else:
        _split_and_anchor(first_event, first_s, first_e, want_start=True, want_end=False)
        _split_and_anchor(last_event, last_s, last_e, want_start=False, want_end=True)


def _assert_anchors_bracket_quote(proj: projection.Projection, quote: str) -> None:
    """Issue #28 plan WP-08: "Verify after write that the range brackets
    the quoted span." Reconstructs each comment id's own anchored text
    (via RunEvent.comment_ids -- projection.py's WP-06 addition) and
    confirms at least one matches *quote* modulo whitespace."""
    by_id: dict[str, list[str]] = {}
    for event in proj.events:
        if isinstance(event, projection.RunEvent):
            for cid in event.comment_ids:
                by_id.setdefault(cid, []).append(event.text)
    normalized_quote = " ".join(quote.split())
    for cid, parts in by_id.items():
        if " ".join("".join(parts).split()) == normalized_quote:
            return
    raise ValueError(
        f"no comment anchor's bracketed text matches the quote {quote!r} (modulo whitespace); anchored text found: "
        f"{ {cid: ''.join(parts) for cid, parts in by_id.items()} }"
    )


def _anchor_span_for_w_id(proj: projection.Projection, w_id: str) -> tuple[int, int] | None:
    """The [start, end) span in proj.text the given comment w:id currently
    brackets (min start, max end across every run RunEvent.comment_ids
    tags with it) -- used by reply_to_comment (WP-09) to give the reply
    the SAME anchor range as the parent it replies to, rather than
    re-locating any text. None if w_id anchors nothing found in the live
    projection (a malformed/stale id)."""
    starts: list[int] = []
    ends: list[int] = []
    run_events = [e for e in proj.events if isinstance(e, projection.RunEvent)]
    for event, (s, e, _pr, _rr) in zip(run_events, proj.offset_map):
        if w_id in event.comment_ids:
            starts.append(s)
            ends.append(e)
    if not starts:
        return None
    return min(starts), max(ends)


# ---------------------------------------------------------------------------
# The five comment parts: build fresh content into each root.
# ---------------------------------------------------------------------------

_INITIALS_MAX = 3


def _initials(author: str) -> str:
    parts = [p for p in author.split() if p]
    letters = "".join(p[0].upper() for p in parts[:_INITIALS_MAX])
    return letters or "?"


def _append_comment(comments_root: Any, *, comment_id: str, para_id: str, author: str, date: str, text: str) -> None:
    comment = ET.SubElement(
        comments_root,
        _w("comment"),
        {_w("id"): comment_id, _w("author"): author, _w("date"): date, _w("initials"): _initials(author)},
    )
    p = ET.SubElement(comment, _w("p"), {f"{{{NS_W14}}}paraId": para_id, f"{{{NS_W14}}}textId": "77777777"})
    ref = ET.SubElement(p, _w("r"))
    ref_rpr = ET.SubElement(ref, _w("rPr"))
    ET.SubElement(ref_rpr, _w("rStyle"), {_w("val"): "CommentReference"})
    ET.SubElement(ref, _w("annotationRef"))
    body = ET.SubElement(p, _w("r"))
    body_rpr = ET.SubElement(body, _w("rPr"))
    # Fixed convention verified against the golden fixture: a comment's
    # own body text is always 20 half-points (10pt), independent of the
    # CommentText style's own default and of the anchored document text's
    # own size.
    ET.SubElement(body_rpr, _w("sz"), {_w("val"): "20"})
    ET.SubElement(body_rpr, _w("szCs"), {_w("val"): "20"})
    t = ET.SubElement(body, _w("t"))
    t.text = text
    if text == "" or text != text.strip():
        t.set(f"{{{_XML_NS}}}space", "preserve")


def _append_comment_ex(ext_root: Any, *, para_id: str, parent_para_id: str | None = None) -> None:
    attrs = {_w15("paraId"): para_id, _w15("done"): "0"}
    if parent_para_id is not None:
        # issue #28 plan WP-09: a reply's own commentEx carries
        # w15:paraIdParent pointing at the PARENT's own paraId -- verified
        # against the golden fixture's real reply (paraId=5FA1F0F0,
        # paraIdParent=4F84A13A, the root's own paraId).
        attrs[_w15("paraIdParent")] = parent_para_id
    ET.SubElement(ext_root, _w15("commentEx"), attrs)


def _append_comment_id(ids_root: Any, *, para_id: str, durable_id: str) -> None:
    ET.SubElement(ids_root, _w16cid("commentId"), {_w16cid("paraId"): para_id, _w16cid("durableId"): durable_id})


def _append_comment_extensible(cex_root: Any, *, durable_id: str, date_utc: str) -> None:
    ET.SubElement(cex_root, _w16cex("commentExtensible"), {_w16cex("durableId"): durable_id, _w16cex("dateUtc"): date_utc})


def _ensure_person(people_root: Any, *, author: str) -> None:
    for child in people_root:
        if projection._ln(child) == "person" and projection._attr(child, "author") == author:
            return
    person = ET.SubElement(people_root, _w15("person"), {_w15("author"): author})
    # providerId/userId are AD-account-specific info a real signed-in Word
    # session fills in (golden: providerId="AD", userId="S::<email>::<guid>")
    # -- unreproducible authentically without a real signed-in account.
    # "None" is a valid ST_PresenceProvider value (per the OOXML people.xml
    # schema: "AD" | "Windows Live" | "None") for exactly this case, and
    # userId falls back to the author string itself, a stable placeholder
    # rather than a fabricated AD identity.
    ET.SubElement(person, _w15("presenceInfo"), {_w15("providerId"): "None", _w15("userId"): author})


# ---------------------------------------------------------------------------
# Serialization: namespace declarations, [Content_Types].xml, rels.
# ---------------------------------------------------------------------------


def _serialize_comment_part(root: Any, namespace_decls: dict[str, str], mc_ignorable: str) -> bytes:
    root.set(_mc("Ignorable"), mc_ignorable)
    return mutations._serialize_xml(root, namespace_decls)


def _build_overrides(
    *,
    document_root: Any,
    raw_xml: bytes,
    comments_root: Any,
    comments_created: bool,
    ext_root: Any,
    ext_created: bool,
    ids_root: Any,
    ids_created: bool,
    cex_root: Any,
    cex_created: bool,
    people_root: Any,
    people_created: bool,
    rels_bytes: bytes | None,
    ct_bytes: bytes,
) -> dict[str, bytes]:
    document_decls = mutations._capture_source_namespaces(raw_xml)
    cex_decls = dict(document_decls)
    cex_decls.setdefault("cr", "http://schemas.microsoft.com/office/comments/2020/reactions")

    overrides: dict[str, bytes] = {
        DEFAULT_PART: mutations._serialize_xml(document_root, document_decls),
        _COMMENTS_PART: _serialize_comment_part(comments_root, document_decls, _MC_IGNORABLE_TOKENS),
        _COMMENTS_EXT_PART: _serialize_comment_part(ext_root, document_decls, _MC_IGNORABLE_TOKENS),
        _COMMENTS_IDS_PART: _serialize_comment_part(ids_root, document_decls, _MC_IGNORABLE_TOKENS),
        _COMMENTS_CEX_PART: _serialize_comment_part(cex_root, cex_decls, _MC_IGNORABLE_TOKENS_CEX),
        _PEOPLE_PART: _serialize_comment_part(people_root, document_decls, _MC_IGNORABLE_TOKENS),
    }

    newly_created = {
        _COMMENTS_PART: comments_created,
        _COMMENTS_EXT_PART: ext_created,
        _COMMENTS_IDS_PART: ids_created,
        _COMMENTS_CEX_PART: cex_created,
        _PEOPLE_PART: people_created,
    }
    if any(newly_created.values()):
        rels_decls: dict[str, str] = {}
        if rels_bytes is not None:
            rels_decls = mutations._capture_source_namespaces(rels_bytes)
            rels_root = ET.fromstring(rels_bytes)
        else:
            rels_root = ET.Element("Relationships", {"xmlns": _PKG_RELS_NS})
        next_rid = mutations._max_rid_in(rels_root) + 1
        for part_name, created in newly_created.items():
            if not created:
                continue
            ET.SubElement(
                rels_root,
                "Relationship",
                {"Id": f"rId{next_rid}", "Type": _PART_REL_TYPES[part_name], "Target": part_name.split("/", 1)[1]},
            )
            next_rid += 1
        overrides[projection._rels_path_for(DEFAULT_PART)] = mutations._serialize_xml(rels_root, rels_decls)

        ct_decls = mutations._capture_source_namespaces(ct_bytes)
        ct_root = ET.fromstring(ct_bytes)
        existing_overrides = {
            child.get("PartName") for child in ct_root if child.tag.rsplit("}", 1)[-1] == "Override"
        }
        for part_name, created in newly_created.items():
            if not created:
                continue
            part_name_slash = f"/{part_name}"
            if part_name_slash in existing_overrides:
                continue
            ET.SubElement(
                ct_root, "Override", {"PartName": part_name_slash, "ContentType": _PART_CONTENT_TYPES[part_name]}
            )
        overrides["[Content_Types].xml"] = mutations._serialize_xml(ct_root, ct_decls)

    return overrides


# ---------------------------------------------------------------------------
# get_comment_thread (read-only)
# ---------------------------------------------------------------------------


def _comment_record(comment_elem: Any, *, durable_id: str, resolved: bool, quoted_by_para: dict[str, str]) -> dict[str, Any]:
    author = projection._attr(comment_elem, "author")
    date = projection._attr(comment_elem, "date")
    content_parts: list[str] = []
    para_id: str | None = None
    for p in comment_elem:
        if projection._ln(p) != "p":
            continue
        if para_id is None:
            para_id = projection._attr(p, "paraId")
        content_parts.append(tracked_changes._collect_text(p, {"t"}))
    return {
        "comment_id": durable_id,
        "content": "".join(content_parts),
        "author": author,
        "created_time": date,
        "resolved": resolved,
        "quoted_text": quoted_by_para.get(para_id or "", ""),
    }


def execute_get_comment_thread(path: str, comment_id: str) -> dict[str, Any]:
    """Read-only: the comment named by *comment_id* (a durableId, from
    add_anchored_comment's own evidence or commentsIds.xml's
    w16cid:durableId), plus its direct replies (commentsExtended.xml's
    w15:paraIdParent linking a reply back to this comment's own paraId).

    Scope limit: if *comment_id* itself names a REPLY (it has its own
    paraIdParent), this returns that reply alone with replies=[] -- it
    does not walk upward to find and return the whole thread's root.
    Reply CREATION and resolve are WP-09, not this WP; this only reads
    threading structure a document (e.g. one authored in Word desktop)
    already carries.
    """
    resolved_path = paths.resolve_allowed_docx_path(path, must_exist=True)
    from . import server as _server

    local_path, is_temp = _server._read_local_copy(resolved_path)
    try:
        with zipfile.ZipFile(local_path) as zf:
            names = set(zf.namelist())
            if _COMMENTS_PART not in names or _COMMENTS_IDS_PART not in names:
                raise _make_error(
                    ErrorCode.INVALID_INPUT,
                    f"document has no comments (comment_id {comment_id!r} cannot exist)",
                    {"comment_id": comment_id},
                )
            comments_root = ET.fromstring(zf.read(_COMMENTS_PART))
            ids_root = ET.fromstring(zf.read(_COMMENTS_IDS_PART))
            ext_root = ET.fromstring(zf.read(_COMMENTS_EXT_PART)) if _COMMENTS_EXT_PART in names else None

        proj = projection.project_part(local_path)
        anchor_by_id: dict[str, list[str]] = {}
        for event in proj.events:
            if isinstance(event, projection.RunEvent):
                for cid in event.comment_ids:
                    anchor_by_id.setdefault(cid, []).append(event.text)
        quoted_by_para: dict[str, str] = {}
        para_id_by_w_id: dict[str, str] = {}
        for c in comments_root:
            if projection._ln(c) != "comment":
                continue
            w_id = projection._attr(c, "id") or ""
            for p in c:
                if projection._ln(p) == "p":
                    pid = projection._attr(p, "paraId")
                    if pid:
                        para_id_by_w_id[w_id] = pid
                        quoted_by_para[pid] = "".join(anchor_by_id.get(w_id, []))
                    break

        para_id_by_durable: dict[str, str] = {}
        durable_by_para: dict[str, str] = {}
        for child in ids_root:
            if projection._ln(child) != "commentId":
                continue
            pid = projection._attr(child, "paraId")
            did = projection._attr(child, "durableId")
            if pid and did:
                para_id_by_durable[did] = pid
                durable_by_para[pid] = did

        target_para_id = para_id_by_durable.get(comment_id)
        if target_para_id is None:
            raise _make_error(
                ErrorCode.INVALID_INPUT,
                f"comment_id {comment_id!r} not found in commentsIds.xml",
                {"comment_id": comment_id, "available_comment_ids": sorted(para_id_by_durable)},
            )

        done_by_para: dict[str, str | None] = {}
        parent_by_para: dict[str, str | None] = {}
        children_by_parent: dict[str, list[str]] = {}
        if ext_root is not None:
            for child in ext_root:
                if projection._ln(child) != "commentEx":
                    continue
                pid = projection._attr(child, "paraId")
                if not pid:
                    continue
                done_by_para[pid] = projection._attr(child, "done")
                parent_pid = projection._attr(child, "paraIdParent")
                parent_by_para[pid] = parent_pid
                if parent_pid:
                    children_by_parent.setdefault(parent_pid, []).append(pid)

        w_id_by_para = {v: k for k, v in para_id_by_w_id.items()}

        def _record_for_para(pid: str) -> dict[str, Any]:
            w_id = w_id_by_para.get(pid)
            comment_elem = next(
                (c for c in comments_root if projection._ln(c) == "comment" and projection._attr(c, "id") == w_id),
                None,
            )
            durable = durable_by_para.get(pid, "")
            return _comment_record(
                comment_elem, durable_id=durable, resolved=done_by_para.get(pid) == "1", quoted_by_para=quoted_by_para
            )

        thread = _record_for_para(target_para_id)
        replies = [_record_for_para(child_pid) for child_pid in children_by_parent.get(target_para_id, [])]
        thread["reply_count"] = len(replies)
        thread["replies"] = replies
        thread["path"] = str(resolved_path)
        return thread
    finally:
        if is_temp:
            local_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# reply_to_comment / resolve_comment -- issue #28 WP-09.
# ---------------------------------------------------------------------------


def execute_reply_to_comment(
    path: str, comment_id: str, text: str, *, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    """Reply to an existing comment (a NEW, independent w:comment whose
    commentEx carries w15:paraIdParent pointing at the parent's own
    paraId -- see _append_comment_ex's docstring, verified against the
    golden fixture's own real reply).

    Reuses the PARENT comment's own EXISTING anchor range in
    word/document.xml (via _anchor_span_for_w_id) rather than locating any
    text of its own -- a reply comment gets its own full commentRangeStart/
    End/commentReference triplet (its own w:id), but brackets the SAME
    live span the parent already does; this matches the golden fixture's
    own shape exactly (its reply, id=1, has its own anchor triplet
    alongside the parent's, id=0, both around the same quoted text).
    """
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)
    proj = projection.project_document_root(document_root)

    parts = _CommentParts(resolved)
    parent_w_id = parts.w_id_for_durable_id(comment_id)
    if parent_w_id is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"comment_id {comment_id!r} not found", {"comment_id": comment_id}
        )
    parent_para_id = parts.para_id_for_w_id(parent_w_id)
    if parent_para_id is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"comment_id {comment_id!r} has no paraId in word/comments.xml",
            {"comment_id": comment_id},
        )

    span = _anchor_span_for_w_id(proj, parent_w_id)
    if span is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"comment_id {comment_id!r} has no live anchor in word/document.xml to reply against",
            {"comment_id": comment_id},
        )
    start, end = span
    quoted_text = proj.text[start:end]

    author = resolve_author_name()
    date = tracked_changes.current_revision_date()

    next_comment_id = _next_comment_id(parts.comments_root if not parts.comments_created else None)
    taken_hex = _existing_hex_ids(parts.ext_root, parts.ids_root)
    new_w_id = str(next_comment_id)
    new_para_id = _new_hex_id(taken_hex)
    new_durable_id = _new_hex_id(taken_hex)

    _insert_anchor(proj, start, end, comment_id=new_w_id)
    _append_comment(parts.comments_root, comment_id=new_w_id, para_id=new_para_id, author=author, date=date, text=text)
    _append_comment_ex(parts.ext_root, para_id=new_para_id, parent_para_id=parent_para_id)
    _append_comment_id(parts.ids_root, para_id=new_para_id, durable_id=new_durable_id)
    _append_comment_extensible(parts.cex_root, durable_id=new_durable_id, date_utc=date)
    _ensure_person(parts.people_root, author=author)

    def _post_verify(written_path: Path) -> None:
        with zipfile.ZipFile(written_path) as zf:
            ids_root = ET.fromstring(zf.read(_COMMENTS_IDS_PART))
            ext_root = ET.fromstring(zf.read(_COMMENTS_EXT_PART))
        found_para = None
        for child in ids_root:
            if projection._ln(child) == "commentId" and projection._attr(child, "durableId") == new_durable_id:
                found_para = projection._attr(child, "paraId")
                break
        if found_para is None:
            raise ValueError(f"reply durableId {new_durable_id!r} not found in commentsIds.xml after write")
        linked = any(
            projection._ln(child) == "commentEx"
            and projection._attr(child, "paraId") == found_para
            and projection._attr(child, "paraIdParent") == parent_para_id
            for child in ext_root
        )
        if not linked:
            raise ValueError(f"reply {new_durable_id!r} does not link back to parent paraId {parent_para_id!r} after write")

    overrides = parts.build_overrides(document_root=document_root, raw_xml=raw_xml)
    mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": 1,
        # "rung" has no locate()-ladder meaning here (no text search is
        # performed -- the reply reuses the parent's own existing anchor);
        # a fixed descriptive label, matching the same "adapt the shared
        # evidence shape to what this tool actually does" precedent
        # mutations.py's markdown tools set for a non-search rung.
        "rung": "reply",
        "before": quoted_text,
        "after": quoted_text,  # a reply never changes document text
        "revision_before": pre_revision["token"],
        "revision_after": post_revision["token"],
        "audit_logged": False,
        "comment_id": new_durable_id,
        "parent_comment_id": comment_id,
    }
    logged, _ = audit.append_audit(path=str(resolved), tool="reply_to_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


def execute_resolve_comment(
    path: str, comment_id: str, *, revision_before: str | None = None, force: bool = False
) -> dict[str, Any]:
    """Resolve a comment thread (w15:done="1" on its own commentEx entry).

    COMMENT_STILL_OPEN, adapted from GoogleDocs-MCP's own member of the
    same name (verify.py/comments.py): that server's resolve_comment
    issues a Drive API call, then RE-QUERIES the comment from the API
    (server-side, independently of the write) and raises
    COMMENT_STILL_OPEN if the re-query still shows it open -- guarding
    against genuine Drive-side eventual consistency (the resolve action
    not durably sticking server-side). This backend has no analogue of
    that hazard: word/commentsExtended.xml is a local file this server
    writes atomically and re-reads synchronously, so "still open after a
    successful local write" can only happen here via a bug in this
    server's own code, never an external service's consistency window.
    The check is kept anyway -- re-reading the FRESH file from disk after
    the atomic write and raising COMMENT_STILL_OPEN (not the generic
    VERIFICATION_FAILED) if it does not show resolved -- so the error
    VOCABULARY still matches Google's for this exact failure mode, giving
    a caller a more specific signal than VERIFICATION_FAILED would.
    """
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = mutations._guard_before_write(resolved, revision_before)

    document_root, raw_xml = mutations._load_document(resolved)

    parts = _CommentParts(resolved)
    w_id = parts.w_id_for_durable_id(comment_id)
    if w_id is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT, f"comment_id {comment_id!r} not found", {"comment_id": comment_id}
        )
    para_id = parts.para_id_for_w_id(w_id)
    if para_id is None:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"comment_id {comment_id!r} has no paraId in word/comments.xml",
            {"comment_id": comment_id},
        )

    target_ex = None
    for child in parts.ext_root:
        if projection._ln(child) == "commentEx" and projection._attr(child, "paraId") == para_id:
            target_ex = child
            break
    if target_ex is None:
        # Defensive: every comment THIS server creates always gets a
        # commentEx entry (_append_comment_ex); a comment authored
        # elsewhere that somehow lacks one gets one created fresh here
        # rather than failing outright.
        target_ex = ET.SubElement(parts.ext_root, _w15("commentEx"), {_w15("paraId"): para_id})
    target_ex.set(_w15("done"), "1")

    def _post_verify(written_path: Path) -> None:
        with zipfile.ZipFile(written_path) as zf:
            ET.fromstring(zf.read(_COMMENTS_EXT_PART))  # basic well-formedness re-check

    overrides = parts.build_overrides(document_root=document_root, raw_xml=raw_xml)
    mutations.atomic_replace_docx_parts(resolved, overrides, post_verify=_post_verify)

    # Independent post-write re-read (see this function's own docstring):
    # COMMENT_STILL_OPEN, not VERIFICATION_FAILED, if the FRESH file does
    # not show this comment resolved.
    with zipfile.ZipFile(resolved) as zf:
        final_ext_root = ET.fromstring(zf.read(_COMMENTS_EXT_PART))
    still_open = True
    for child in final_ext_root:
        if projection._ln(child) == "commentEx" and projection._attr(child, "paraId") == para_id:
            still_open = projection._attr(child, "done") != "1"
            break
    if still_open:
        raise _make_error(
            ErrorCode.COMMENT_STILL_OPEN,
            f"comment_id {comment_id!r} is still open after the resolve attempt",
            {"comment_id": comment_id},
        )

    post_revision = projection.compute_revision(resolved)
    evidence: dict[str, Any] = {
        "applied": True,
        "match_count": 1,
        "rung": "resolve",
        "before": "open",
        "after": "resolved",
        "revision_before": pre_revision["token"],
        "revision_after": post_revision["token"],
        "audit_logged": False,
        "comment_id": comment_id,
    }
    logged, _ = audit.append_audit(path=str(resolved), tool="resolve_comment", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
