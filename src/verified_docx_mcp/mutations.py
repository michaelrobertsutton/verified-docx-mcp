# New for issue #28 WP-04. No GoogleDocs-MCP analogue exists to lift
# wholesale: that server's markdown_mutations.py drives the Docs API's
# batchUpdate with a server-assigned requiredRevisionId precondition and a
# structural-inventory guardrail computed from Docs JSON; this backend has
# no such API, so the equivalent guard (lock/sync/revision precondition,
# comment-anchor and tracked-change hazard scan, atomic write with OPC
# validation and a .jsbak rollback) is implemented directly against the
# OOXML package. The SHAPE of the pipeline (guard -> compile -> write ->
# post-read -> evidence -> audit) mirrors that file's docstring
# ("validate -> pre-read -> [guardrail] -> compile -> ... -> post-read ->
# assemble evidence -> audit"); the mechanics differ because the target
# format does.
"""Guarded write pipeline for the three markdown-mutating tools:
``replace_body_markdown``, ``replace_range_markdown``, ``append_markdown``.

Pipeline, in order (issue #28 plan WP-04; "Codex defect: writes existed
before the guard" — the ordering below is the fix, not incidental):

1. Resolve + guard (``_guard_before_write``): ``lock_status`` first, before
   any temp file is written. An owner file present -> ``DOCX_LOCKED``. Not
   sync-quiesced -> one bounded wait (core/document-backend-protocol.md
   §4: "wait once, up to 10 seconds, then stop"), then ``SYNC_IN_FLIGHT``
   if still not quiesced. A ``revision_before`` that no longer matches ->
   ``REVISION_CONFLICT``.
2. Locate the target range (whole body / one section by ``section_key`` /
   the append point) and scan it for comment anchors
   (``w:commentRangeStart``/``End``/``commentReference``) and tracked
   changes (``w:ins``/``w:del``) — refuse (``COMMENT_ANCHORS_IN_RANGE`` /
   ``TRACKED_CHANGES_PRESENT``) unless ``force=True``; with ``force``, the
   range is replaced anyway and ``orphaned_comment_ids`` records what was
   removed so nothing disappears silently ahead of WP-08's comment tools.
3. Render the intended markdown via ``markdown_to_ooxml.render_blocks``
   against a ``StyleContext`` built from the TARGET document's own styles/
   numbering/rels (never a hardcoded style).
4. Atomic write (``atomic_replace_docx_parts``): build the full new zip in
   a temp file in the same directory, validate it (``opc_valid``) BEFORE
   ever touching the original, then ``os.replace()``, keeping a ``.jsbak``
   copy of the pre-write original until a post-write re-read/re-project
   confirms the change; on any post-check failure, restore from ``.jsbak``
   and raise ``VERIFICATION_FAILED``.
5. Assemble the eight evidence keys from a genuine re-read of the file on
   disk (never from what the tool believes it wrote) and log the call via
   ``audit.append_audit``.
"""

from __future__ import annotations

import difflib
import os
import re
import shutil
import tempfile
import time
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from . import audit, markdown_to_ooxml, paths, projection
from .errors import ErrorCode, _make_error
from .projection import DEFAULT_PART, R_NS, W_NS

# ---------------------------------------------------------------------------
# Namespace preservation
# ---------------------------------------------------------------------------

# xml.etree.ElementTree does not retain a parsed document's own xmlns prefix
# assignments (parsing collapses them into Clark-notation {uri}tag); left
# alone, re-serializing would emit auto-generated "ns0:"/"ns1:" prefixes for
# EVERY namespace on the root — which breaks a real Word document, because
# mc:Ignorable (and w:rsid tracking, mc:AlternateContent Requires=, etc.)
# names PREFIXES literally, not the namespace URIs those prefixes resolve
# to.
#
# Registering a source part's (prefix, uri) pairs via ET.register_namespace
# (a process-global table) is NECESSARY but NOT SUFFICIENT: it only fixes
# the prefix ElementTree's serializer chooses for a namespace that some
# element or attribute in the MODIFIED tree still actually uses. A
# namespace declared on the source root but referenced only by content this
# WP's tools removed (a whole w:body replace drops every w:ins/w:del,
# every w:tbl, every drawing — and their namespaces along with them) is, as
# far as ElementTree's serializer can tell, simply not present in the tree
# any more, so it emits NO declaration for it at all, regardless of what is
# registered. mc:Ignorable then survives on the root naming prefixes with
# no declaration anywhere in the file — Word refuses to open it. (Found in
# PR #3 review: reproduced against tracked.docx and sections.docx, 35
# declarations collapsing to 2-4, and — the second-order hazard the first
# pass missed — "r" (the relationships namespace every hyperlink/image
# r:id depends on) is in the lost set too, independent of whether
# mc:Ignorable happens to name it.)
#
# The fix re-emits, on the ROOT element's own opening tag, EVERY namespace
# declaration the source root carried — not just the ones ElementTree's
# serializer decides are still "in use" — via _ensure_namespace_declarations
# below. This is deliberately a post-serialization byte-level patch rather
# than adding xmlns:* as literal Element attributes before calling
# ET.tostring: a namespace that IS still in use would then get declared
# TWICE (once by ElementTree's own automatic declaration, once by our
# literal attribute), which is a well-formedness error. Patching the
# ALREADY-SERIALIZED bytes lets us see exactly what ElementTree did and
# only add what it left out.


def _capture_source_namespaces(xml_bytes: bytes) -> dict[str, str]:
    """Return {prefix: uri} for every namespace declared on *xml_bytes*'
    root — regardless of whether anything in a later-mutated tree still
    uses it — and register each with ElementTree's global prefix table
    (so a namespace that IS still in use keeps its original prefix
    spelling rather than an auto-generated one). The returned map is what
    _ensure_namespace_declarations re-injects for the ones ElementTree's
    serializer leaves out; see this module's namespace-preservation
    comment above for why both steps are required.
    """
    import io

    captured: dict[str, str] = {}
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        captured[prefix] = uri
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            # ElementTree reserves the "nsN" prefix shape for its own
            # auto-numbering and refuses to register it explicitly. A part
            # carrying that shape can only get there via a PRIOR
            # ElementTree round trip that itself did not preserve prefixes
            # (never a real Word/OOXML producer) — falling back to
            # ElementTree's own auto-numbering for just this one namespace
            # changes its prefix spelling, never its meaning, and no part
            # this server touches names a package-relationships-namespace
            # prefix inside an mc:Ignorable-style attribute value (only
            # real Word namespaces like w14/w15/... ever are).
            pass
    return captured


_XMLNS_DECL_RE = re.compile(rb'xmlns(?::([A-Za-z_][\w.-]*))?=')
_ROOT_TAG_NAME_RE = re.compile(rb"<([^\s/>]+)")


def _ensure_namespace_declarations(xml_bytes: bytes, original_decls: dict[str, str]) -> bytes:
    """Re-declare, on the ROOT element's own opening tag, every namespace
    prefix in *original_decls* that ElementTree's serializer did not
    already emit — see this module's namespace-preservation comment for
    why this is necessary in addition to ET.register_namespace.

    The root tag's boundary is found via the first raw b">" byte. That is
    not a general XML-safe assumption (an attribute value could in
    principle contain a literal ">"), but it holds for every root element
    this function is ever called on (w:document, w:numbering,
    Relationships, Types): none of their own attribute values — namespace
    URIs, mc:Ignorable's token list, a style id — contain one.
    """
    if not original_decls:
        return xml_bytes
    end = xml_bytes.find(b">")
    if end == -1:
        return xml_bytes
    open_tag = xml_bytes[: end + 1]

    declared: set[bytes] = set()
    for m in _XMLNS_DECL_RE.finditer(open_tag):
        declared.add(m.group(1) or b"")

    additions = bytearray()
    for prefix, uri in original_decls.items():
        key = prefix.encode("utf-8")
        if key in declared:
            continue
        declared.add(key)
        escaped_uri = uri.replace("&", "&amp;").replace('"', "&quot;")
        if prefix:
            additions += f' xmlns:{prefix}="{escaped_uri}"'.encode()
        else:
            additions += f' xmlns="{escaped_uri}"'.encode()

    if not additions:
        return xml_bytes

    tag_name_match = _ROOT_TAG_NAME_RE.match(xml_bytes)
    insert_at = tag_name_match.end() if tag_name_match else 1
    return xml_bytes[:insert_at] + bytes(additions) + xml_bytes[insert_at:]


def _serialize_xml(root: Any, original_decls: dict[str, str] | None = None) -> bytes:
    body = ET.tostring(root, encoding="utf-8")
    if original_decls:
        body = _ensure_namespace_declarations(body, original_decls)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body


# ---------------------------------------------------------------------------
# OPC validation
# ---------------------------------------------------------------------------

_RELATIONSHIP_ATTR_LOCALS = ("id", "embed")

# markup-compatibility namespace: mc:Ignorable lists PREFIXES (not URIs) of
# extension namespaces a non-supporting consumer may ignore. PR #3 review
# found this was the visible symptom of the namespace-declaration bug
# above — a part could pass every check opc_valid had before this rule and
# still be a file Word refuses to open, because mc:Ignorable survived
# naming prefixes nothing declared any more.
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_IGNORABLE_ATTR = f"{{{_MC_NS}}}Ignorable"


def _declared_prefixes(xml_bytes: bytes) -> set[str]:
    """Every namespace prefix actually declared somewhere in *xml_bytes*
    (via ET.iterparse's "start-ns" event) — the empty string for a bare
    default ``xmlns="..."``. This is independent of _capture_source_namespaces
    above: that function's caller wants the (prefix, uri) map to re-inject;
    this one is opc_valid's own read-only check of what ended up on disk,
    over the bytes actually written, not the pre-write source."""
    import io

    return {prefix for _, (prefix, _uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",))}


def opc_valid(docx_path: Path) -> tuple[bool, list[str]]:
    """Structural OPC validity, per the WP-04 acceptance contract: every
    XML/.rels part well-formed, every r:id/r:embed resolves against its
    part's own .rels, [Content_Types].xml covers every part present (a
    Default by extension, or an Override by part name), and every prefix
    named in a part's own mc:Ignorable is actually declared somewhere in
    that same part (PR #3 review: the namespace-declaration bug that made
    Word refuse a written file passed every earlier rule here, because
    none of them checked namespace-declaration consistency at all)."""
    problems: list[str] = []
    try:
        with zipfile.ZipFile(docx_path) as zf:
            bad = zf.testzip()
            if bad is not None:
                return False, [f"corrupt zip entry: {bad}"]
            names = set(zf.namelist())

            parsed: dict[str, Any] = {}
            raw_bytes: dict[str, bytes] = {}
            for name in sorted(names):
                if name.endswith((".xml", ".rels")):
                    data = zf.read(name)
                    raw_bytes[name] = data
                    try:
                        parsed[name] = ET.fromstring(data)
                    except ET.ParseError as exc:
                        problems.append(f"{name}: not well-formed XML: {exc}")
            if problems:
                return False, problems

            for name, root in parsed.items():
                ignorable = root.get(_MC_IGNORABLE_ATTR)
                if ignorable:
                    tokens = ignorable.split()
                    declared = _declared_prefixes(raw_bytes[name])
                    missing_prefixes = [t for t in tokens if t not in declared]
                    if missing_prefixes:
                        problems.append(
                            f"{name}: mc:Ignorable names undeclared prefix(es) {missing_prefixes} "
                            f"(declared: {sorted(declared)})"
                        )

            for name, root in parsed.items():
                if not name.endswith(".xml") or name == "[Content_Types].xml":
                    continue
                rids: set[str] = set()
                for el in root.iter():
                    for attr_name, attr_val in el.attrib.items():
                        if not attr_name.startswith(f"{{{R_NS}}}"):
                            continue
                        local = attr_name.rsplit("}", 1)[-1]
                        if local in _RELATIONSHIP_ATTR_LOCALS and attr_val:
                            rids.add(attr_val)
                if not rids:
                    continue
                rels_name = projection._rels_path_for(name)
                available: set[str] = set()
                rels_root = parsed.get(rels_name)
                if rels_root is not None:
                    for rel in rels_root:
                        rid = rel.get("Id")
                        if rid:
                            available.add(rid)
                missing = rids - available
                if missing:
                    problems.append(f"{name}: unresolved relationship id(s) {sorted(missing)} (rels part {rels_name!r})")

            ct_root = parsed.get("[Content_Types].xml")
            if ct_root is None:
                problems.append("[Content_Types].xml missing")
            else:
                default_exts: set[str] = set()
                overrides: set[str] = set()
                for child in ct_root:
                    tag = child.tag.rsplit("}", 1)[-1]
                    if tag == "Default":
                        ext = child.get("Extension", "")
                        if ext:
                            default_exts.add(ext.lower())
                    elif tag == "Override":
                        part_name = child.get("PartName", "")
                        if part_name:
                            overrides.add(part_name)
                for name in names:
                    if name == "[Content_Types].xml":
                        continue
                    part_name = name if name.startswith("/") else f"/{name}"
                    if part_name in overrides:
                        continue
                    ext = name.rsplit(".", 1)[-1].lower() if "." in name.rsplit("/", 1)[-1] else ""
                    if ext in default_exts:
                        continue
                    problems.append(f"{name}: not covered by [Content_Types].xml (no Default or Override)")
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        problems.append(str(exc))

    return (len(problems) == 0, problems)


# ---------------------------------------------------------------------------
# Atomic write with .jsbak rollback
# ---------------------------------------------------------------------------


def _rebuild_zip(source_path: Path, overrides: dict[str, bytes], temp_path: Path) -> None:
    with zipfile.ZipFile(source_path) as src, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as dst:
        written: set[str] = set()
        for item in src.infolist():
            data = overrides.get(item.filename, src.read(item))
            dst.writestr(item, data)
            written.add(item.filename)
        for name, data in overrides.items():
            if name not in written:
                dst.writestr(name, data)


def atomic_replace_docx_parts(
    original_path: Path,
    overrides: dict[str, bytes],
    *,
    post_verify: Callable[[Path], None],
    _corrupt_temp_for_test: bool = False,
) -> None:
    """Write *overrides* (part name -> new bytes; any name not already in
    the package is added) into a fresh copy of *original_path*, atomically.

    Order (issue #28 plan WP-04's "Atomic write" bullet, exactly):
    build a temp file in the SAME directory -> opc_valid (before anything
    is replaced; a failure here leaves the original completely untouched
    and raises OPC_INVALID) -> back up the original to ``.jsbak`` ->
    ``os.replace()`` -> *post_verify(original_path)* (re-open, re-project,
    diff — mutations.py's callers do this) -> on success, delete
    ``.jsbak``; on any exception from *post_verify*, restore the original
    from ``.jsbak`` and raise VERIFICATION_FAILED.

    ``_corrupt_temp_for_test`` is a test-only seam (never set by a tool):
    it corrupts the temp file's bytes AFTER it is built but BEFORE
    opc_valid runs, so a test can exercise the "corrupted temp write"
    acceptance case deterministically without depending on a real
    filesystem fault.
    """
    tmp_fd, tmp_name = tempfile.mkstemp(
        prefix=f"{original_path.stem}.tmp-", suffix=".docx", dir=str(original_path.parent)
    )
    os.close(tmp_fd)
    tmp_path = Path(tmp_name)
    tmp_consumed = False  # True once os.replace() has moved tmp_path onto original_path
    jsbak_path = original_path.with_name(original_path.name + ".jsbak")
    replaced = False
    try:
        _rebuild_zip(original_path, overrides, tmp_path)

        if _corrupt_temp_for_test:
            with open(tmp_path, "r+b") as fh:
                fh.seek(0)
                fh.write(b"\x00" * min(64, tmp_path.stat().st_size))

        valid, problems = opc_valid(tmp_path)
        if not valid:
            raise _make_error(
                ErrorCode.OPC_INVALID,
                "Rendered .docx failed OPC validation; the original file was left untouched.",
                {"problems": problems},
            )

        shutil.copyfile(original_path, jsbak_path)
        os.replace(str(tmp_path), str(original_path))
        replaced = True
        tmp_consumed = True

        try:
            post_verify(original_path)
        except Exception as exc:
            os.replace(str(jsbak_path), str(original_path))
            raise _make_error(
                ErrorCode.VERIFICATION_FAILED,
                f"Post-write verification failed; restored the original from .jsbak. Detail: {exc}",
                {"detail": str(exc)},
            ) from exc
        else:
            jsbak_path.unlink(missing_ok=True)
    finally:
        if not tmp_consumed and tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        if not replaced:
            jsbak_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Guard: lock/sync/revision, run BEFORE any temp file is written.
# ---------------------------------------------------------------------------

# Module-level so a test can shrink it (mutations._QUIESCE_INTERVAL_SECONDS =
# 0.05) rather than eating lock_status's real ~1.5s default on every single
# guarded call in the unit suite. Not exposed on any tool signature — the
# public contract (core/document-backend-protocol.md §4/§9) only promises
# "two (size, mtime_ns) samples ~1.5s apart", not this exact constant.
_QUIESCE_INTERVAL_SECONDS = 1.5


def _guard_before_write(resolved: Path, revision_before: str | None) -> dict[str, Any]:
    # Lazy import: server.py imports this module at load time to register
    # the tools below, so a module-level "from . import server" here would
    # be circular. By the time any tool is actually CALLED both modules are
    # fully initialized, so a deferred import inside the function body is
    # safe and avoids restructuring server.py's existing WP-02/WP-03 code.
    from . import server as _server

    status = _server.execute_lock_status(str(resolved), quiesce_interval=_QUIESCE_INTERVAL_SECONDS)
    if status["owner_file"]["present"]:
        raise _make_error(
            ErrorCode.DOCX_LOCKED,
            f"{resolved} is owned by another Word/LibreOffice session; report the owner and stop (no retry loop).",
            {"owner_file": status["owner_file"]},
        )
    if not status["sync_quiesced"]:
        already_waited = status["sync_detail"]["interval_seconds"]
        remaining = max(0.0, 10.0 - already_waited)
        time.sleep(remaining)
        status2 = _server.execute_lock_status(str(resolved), quiesce_interval=_QUIESCE_INTERVAL_SECONDS)
        if not status2["sync_quiesced"]:
            raise _make_error(
                ErrorCode.SYNC_IN_FLIGHT,
                "The sync client still appears to be writing this file after one bounded wait (<=10s total); not retried again.",
                {"sync_detail": status2["sync_detail"]},
            )

    current = projection.compute_revision(resolved)
    if revision_before is not None and current["token"] != revision_before:
        raise _make_error(
            ErrorCode.REVISION_CONFLICT,
            "revision_before no longer matches the file's current revision.",
            {"revision_before": revision_before, "current_revision": current["token"]},
        )
    return current


# ---------------------------------------------------------------------------
# Hazard scan: comment anchors and tracked changes inside a target range.
# ---------------------------------------------------------------------------

_COMMENT_ANCHOR_TAGS = frozenset({"commentRangeStart", "commentRangeEnd", "commentReference"})
_TRACKED_CHANGE_TAGS = frozenset({"ins", "del"})


def _scan_range_hazards(elements: list[Any]) -> dict[str, Any]:
    comment_ids: set[str] = set()
    tracked_change_authors: set[str | None] = set()
    for element in elements:
        for node in element.iter():
            tag = projection._ln(node)
            if tag in _COMMENT_ANCHOR_TAGS:
                cid = projection._attr(node, "id")
                if cid is not None:
                    comment_ids.add(cid)
            elif tag in _TRACKED_CHANGE_TAGS:
                tracked_change_authors.add(projection._attr(node, "author"))
    return {
        "comment_ids": sorted(comment_ids, key=lambda v: (len(v), v)),
        "tracked_change_authors": tracked_change_authors,
    }


def _check_hazards_or_raise(hazards: dict[str, Any], force: bool, own_author: str) -> list[str]:
    """Returns orphaned_comment_ids (empty unless force=True and comment
    anchors were present). Raises COMMENT_ANCHORS_IN_RANGE /
    TRACKED_CHANGES_PRESENT when hazards are present and force is False.

    Issue #28 plan WP-07b-a own-author exclusion (load-bearing, per that
    WP's own notes): a w:ins/w:del authored by *own_author* (this
    server's own configured identity, author.resolve_author_name()) does
    NOT trigger TRACKED_CHANGES_PRESENT, or a second track_changes=True
    write over the server's own prior tracked edit would deadlock. A
    tracked change authored by anyone else still refuses normally.
    """
    if hazards["comment_ids"] and not force:
        raise _make_error(
            ErrorCode.COMMENT_ANCHORS_IN_RANGE,
            "The target range contains comment anchors; pass force=True to proceed (the anchors will be removed and their ids reported as orphaned_comment_ids).",
            {"comment_ids": hazards["comment_ids"]},
        )
    foreign_authors = {a for a in hazards["tracked_change_authors"] if a != own_author}
    if foreign_authors and not force:
        raise _make_error(
            ErrorCode.TRACKED_CHANGES_PRESENT,
            f"The target range contains tracked changes (w:ins/w:del) authored by someone other than {own_author!r}; "
            "pass force=True to proceed.",
            {"foreign_authors": sorted(a for a in foreign_authors if a is not None)},
        )
    return hazards["comment_ids"] if force else []


def _strip_comment_anchors_by_id(elements: list[Any], ids: set[str]) -> None:
    """Remove every w:commentRangeStart/End/commentReference whose id is
    in *ids*, wherever it occurs inside *elements* (searched recursively —
    a comment anchor can sit nested inside a table cell, not just directly
    under a paragraph).

    PR #3 review, should-fix #2 ("half-anchor pairs"): when
    replace_range_markdown removes a comment anchor from inside the
    target range under force=True, that same comment's OTHER half (the
    matching commentRangeStart/End, or the commentReference) can live
    OUTSIDE the range — a comment whose span crosses the section
    boundary. Reporting the id in orphaned_comment_ids is not enough on
    its own if the other half is left behind: that half would then be an
    unpaired anchor in the surviving document, which is exactly what this
    removes. Called on the SURVIVING elements (never on the ones already
    being replaced) with the ids already known to be orphaned; a no-op
    when a given id's anchors were already fully inside the removed range
    (nothing in the surviving elements will match)."""
    if not ids:
        return
    for root_el in elements:
        _strip_comment_anchors_recursive(root_el, ids)


def _strip_comment_anchors_recursive(parent: Any, ids: set[str]) -> None:
    for child in list(parent):
        tag = projection._ln(child)
        if tag in _COMMENT_ANCHOR_TAGS:
            cid = projection._attr(child, "id")
            if cid in ids:
                parent.remove(child)
                continue
        _strip_comment_anchors_recursive(child, ids)


# ---------------------------------------------------------------------------
# Body/section access helpers
# ---------------------------------------------------------------------------


def _find_body(document_root: Any) -> Any:
    for child in document_root:
        if projection._ln(child) == "body":
            return child
    raise _make_error(ErrorCode.PART_NOT_FOUND, "word/document.xml has no w:body element.", {})


def _split_body(body: Any) -> tuple[list[Any], Any | None]:
    """(content_children, sect_pr_or_None) — the trailing w:sectPr (the
    body's own page-layout section, not a document SECTION per
    find_sections) is never part of a replaceable range."""
    children = list(body)
    if children and projection._ln(children[-1]) == "sectPr":
        return children[:-1], children[-1]
    return children, None


def _paragraph_style_and_outline(p: Any) -> tuple[str | None, int | None]:
    style_id = None
    outline_lvl = None
    for child in p:
        if projection._ln(child) == "pPr":
            for grandchild in child:
                if projection._ln(grandchild) == "pStyle":
                    style_id = projection._attr(grandchild, "val")
                elif projection._ln(grandchild) == "outlineLvl":
                    raw = projection._attr(grandchild, "val")
                    if raw is not None and raw.isdigit():
                        outline_lvl = int(raw)
    return style_id, outline_lvl


def _paragraph_text(p: Any) -> str:
    parts: list[str] = []
    for t in p.iter(f"{{{W_NS}}}t"):
        parts.append(t.text or "")
    return "".join(parts)


def locate_section_range(
    body_children: list[Any], section_key: str, styles_by_id: dict[str, dict]
) -> tuple[int, int]:
    """(start_index, end_index) into *body_children* (exclusive end) for
    the top-level children belonging to the section named *section_key* —
    from its own heading (inclusive) up to, but not including, the next
    TOP-LEVEL heading paragraph (any level), or the end of the body.

    Recomputes headings directly over *body_children* (rather than reusing
    find_sections_impl's full nested-paragraph walk) because this needs
    body-level CHILD INDICES to splice XML, not the para_ref-indexed view
    find_sections returns; the outline-level resolution rule itself
    (list_styles, falling back to a paragraph's own w:outlineLvl) is
    identical — see projection._resolve_outline_level. Raises
    SECTION_NOT_FOUND if no heading in *body_children* slugifies +
    disambiguates to *section_key* (mirrors find_sections_impl's own
    section_key construction: slug(heading text) + a 1-based ordinal).
    """
    heading_positions: list[int] = []
    for idx, child in enumerate(body_children):
        if projection._ln(child) != "p":
            continue
        style_id, direct_outline = _paragraph_style_and_outline(child)
        level = projection._resolve_outline_level(style_id, direct_outline, styles_by_id)
        if level is None:
            continue
        heading_positions.append(idx)

    slug_counts: dict[str, int] = {}
    for pos_idx, body_idx in enumerate(heading_positions):
        heading_text = _paragraph_text(body_children[body_idx]).strip()
        base_slug = projection._slugify(heading_text)
        slug_counts[base_slug] = slug_counts.get(base_slug, 0) + 1
        key = f"{base_slug}-{slug_counts[base_slug]}"
        if key == section_key:
            start = body_idx
            end = heading_positions[pos_idx + 1] if pos_idx + 1 < len(heading_positions) else len(body_children)
            return start, end

    raise _make_error(
        ErrorCode.SECTION_NOT_FOUND,
        f"No section with section_key {section_key!r} was found (call find_sections to enumerate the current ones).",
        {"section_key": section_key},
    )


# ---------------------------------------------------------------------------
# Part serialization: build the {part_name: bytes} overrides dict for one
# write, including any new numbering.xml / document.xml.rels /
# [Content_Types].xml content the markdown_to_ooxml.StyleContext needed.
# ---------------------------------------------------------------------------

_NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def _build_overrides(
    resolved: Path,
    document_root: Any,
    document_xml_bytes: bytes,
    ctx: markdown_to_ooxml.StyleContext,
) -> dict[str, bytes]:
    overrides: dict[str, bytes] = {}

    document_decls = _capture_source_namespaces(document_xml_bytes)
    overrides[DEFAULT_PART] = _serialize_xml(document_root, document_decls)

    with zipfile.ZipFile(resolved) as zf:
        names = set(zf.namelist())
        rels_bytes = zf.read(projection._rels_path_for(DEFAULT_PART)) if projection._rels_path_for(DEFAULT_PART) in names else None
        numbering_bytes = zf.read("word/numbering.xml") if "word/numbering.xml" in names else None
        content_types_bytes = zf.read("[Content_Types].xml")

    numbering_created = False
    if ctx.new_abstract_nums or ctx.new_nums:
        numbering_decls: dict[str, str] = {}
        if numbering_bytes is not None:
            numbering_decls = _capture_source_namespaces(numbering_bytes)
            numbering_root = ET.fromstring(numbering_bytes)
        else:
            numbering_created = True
            ET.register_namespace("w", W_NS)
            numbering_root = ET.Element(f"{{{W_NS}}}numbering")
        # abstractNum entries must precede num entries (OOXML numbering.xml
        # schema order) — insert new abstracts before the first existing
        # w:num, and append new nums at the end (order among nums is free).
        first_num_idx = len(list(numbering_root))
        for i, child in enumerate(numbering_root):
            if projection._ln(child) == "num":
                first_num_idx = i
                break
        for offset, el in enumerate(ctx.new_abstract_nums):
            numbering_root.insert(first_num_idx + offset, el)
        for el in ctx.new_nums:
            numbering_root.append(el)
        overrides["word/numbering.xml"] = _serialize_xml(numbering_root, numbering_decls)

    if ctx.new_relationships or numbering_created:
        rels_decls: dict[str, str] = {}
        if rels_bytes is not None:
            rels_decls = _capture_source_namespaces(rels_bytes)
            rels_root = ET.fromstring(rels_bytes)
        else:
            rels_root = ET.Element("Relationships", {"xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"})
        for rid, rel_type, target in ctx.new_relationships:
            ET.SubElement(rels_root, "Relationship", {"Id": rid, "Type": rel_type, "Target": target, "TargetMode": "External"})
        if numbering_created:
            ET.SubElement(rels_root, "Relationship", {"Id": f"rId{_max_rid_in(rels_root) + 1}", "Type": _NUMBERING_REL_TYPE, "Target": "numbering.xml"})
        overrides[projection._rels_path_for(DEFAULT_PART)] = _serialize_xml(rels_root, rels_decls)

    if numbering_created:
        content_types_decls = _capture_source_namespaces(content_types_bytes)
        ct_root = ET.fromstring(content_types_bytes)
        already_covered = any(
            child.get("PartName") == "/word/numbering.xml"
            for child in ct_root
            if child.tag.rsplit("}", 1)[-1] == "Override"
        )
        if not already_covered:
            ET.SubElement(ct_root, "Override", {"PartName": "/word/numbering.xml", "ContentType": _NUMBERING_CONTENT_TYPE})
        overrides["[Content_Types].xml"] = _serialize_xml(ct_root, content_types_decls)

    return overrides


def _max_rid_in(rels_root: Any) -> int:
    best = 0
    for rel in rels_root:
        m = markdown_to_ooxml._REL_ID_RE.match(rel.get("Id", ""))
        if m:
            best = max(best, int(m.group(1)))
    return best


# ---------------------------------------------------------------------------
# Shared write core
# ---------------------------------------------------------------------------


def _load_document(resolved: Path) -> tuple[Any, bytes]:
    with zipfile.ZipFile(resolved) as zf:
        raw = zf.read(DEFAULT_PART)
    # Registers prefixes now so any NEW element built against this parse
    # (W_NS/R_NS) picks up the source's own prefix spelling; the returned
    # map itself is recomputed (redundantly, but cheaply) by
    # _build_overrides from the same bytes when it actually serializes.
    _capture_source_namespaces(raw)
    return ET.fromstring(raw), raw


def _diff_modulo_whitespace(intended: str, actual: str) -> list[str]:
    norm_intended = " ".join(intended.split())
    norm_actual = " ".join(actual.split())
    if norm_intended == norm_actual:
        return []
    return list(difflib.unified_diff(norm_intended.splitlines(), norm_actual.splitlines(), lineterm=""))


def _write_and_verify(
    resolved: Path,
    document_root: Any,
    document_xml_bytes: bytes,
    ctx: markdown_to_ooxml.StyleContext,
    *,
    post_verify: Callable[[Path], None],
) -> None:
    overrides = _build_overrides(resolved, document_root, document_xml_bytes, ctx)
    atomic_replace_docx_parts(resolved, overrides, post_verify=post_verify)


def _evidence(
    *,
    applied: bool,
    match_count: int,
    rung: int,
    before: str,
    after: str,
    revision_before: str,
    revision_after: str,
    audit_logged: bool,
    orphaned_comment_ids: list[str],
    track: Any | None = None,
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
    }
    if orphaned_comment_ids:
        evidence["orphaned_comment_ids"] = orphaned_comment_ids
    if track is not None:
        # issue #28 plan WP-07b-a: "The evidence envelope gains
        # revision_ids (the ids created) and track_changes: true."
        evidence["track_changes"] = True
        evidence["revision_ids"] = track.revision_ids
    return evidence


# ---------------------------------------------------------------------------
# Tool 1: replace_body_markdown
# ---------------------------------------------------------------------------


def execute_replace_body_markdown(
    path: str, markdown: str, *, revision_before: str | None = None, force: bool = False, track_changes: bool = False
) -> dict[str, Any]:
    from . import tracked_changes
    from .author import resolve_author_name

    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = _guard_before_write(resolved, revision_before)

    document_root, raw_xml = _load_document(resolved)
    body = _find_body(document_root)
    target_elements, sect_pr = _split_body(body)

    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    numbering_index = projection.load_numbering_index(resolved)
    before_text = projection.markdown_from_elements(target_elements, styles_by_id, numbering_index)

    own_author = resolve_author_name()
    hazards = _scan_range_hazards(target_elements)
    orphaned_comment_ids = _check_hazards_or_raise(hazards, force, own_author)

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    new_elements = markdown_to_ooxml.render_blocks(markdown, ctx)
    # The verification target is what read_document_markdown would render
    # from THESE elements (run through the identical rendering rules
    # read_document_markdown itself uses; see projection.py's module
    # docstring) — NOT the raw markdown text. Diffing against the raw
    # input would spuriously fail for any table (GFM column-count
    # normalisation, cell-paragraph joining) or list whose markdown the
    # renderer reproduces in a different but equivalent form, even on a
    # perfectly correct write. new_elements' own paragraphs only ever
    # reference numIds markdown_to_ooxml just allocated (ctx.new_*), which
    # is not yet in this document's numbering.xml on disk — the index for
    # THIS preview comes from those in-memory elements, never
    # load_numbering_index(resolved).
    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_preview = projection.markdown_from_elements(new_elements, styles_by_id, new_numbering_index)

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # issue #28 plan WP-07b-a: track_changes=True marks the OLD content
        # deleted (w:del, its w:t renamed to w:delText) rather than
        # removing it, and the NEW content inserted (w:ins) rather than a
        # bare rewrite -- reading the projection continues to exclude
        # w:del text and include w:ins text (unchanged since WP-03/06), so
        # after_text below already reads as the new content alone, same
        # as the non-tracked path. Scope limit, named rather than
        # silently risked: this wraps RUNS, not whole table/list
        # structures -- a target range or replacement containing a table
        # is untested here (real Word's own table-level tracked-deletion
        # modeling is a much larger, separate problem; every fixture this
        # WP tests against is plain paragraph prose).
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        # mark_elements_deleted (not the generic wrap_all_runs): a run
        # already inside a PENDING w:ins (e.g. from an earlier
        # track_changes=True call over the same document) is removed
        # outright rather than wrapped in a second, redundant w:del --
        # see that function's own docstring for why.
        tracked_changes.mark_elements_deleted(target_elements, track)
        tracked_changes.wrap_all_runs(new_elements, _ins_wrap)
        insert_at = len(target_elements)
        for offset, el in enumerate(new_elements):
            body.insert(insert_at + offset, el)
        # target_elements (now del-wrapped) are left exactly where they
        # already were -- body's own children were never touched, so
        # sect_pr (still body's last child) needs no re-append.
    else:
        for child in list(body):
            body.remove(child)
        for el in new_elements:
            body.append(el)
        if sect_pr is not None:
            body.append(sect_pr)

    def _post_verify(written_path: Path) -> None:
        markdown_after, _, _ = projection.read_document_markdown(written_path)
        diff = _diff_modulo_whitespace(intended_preview, markdown_after)
        if diff:
            raise ValueError(f"re-read body does not match the intended rendering modulo whitespace: {diff}")

    _write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_text, _, _ = projection.read_document_markdown(resolved)
    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=4,
        before=before_text,
        after=after_text,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=orphaned_comment_ids,
        track=track,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_body_markdown", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool 2: replace_range_markdown(section_key)
# ---------------------------------------------------------------------------


def execute_replace_range_markdown(
    path: str,
    section_key: str,
    markdown: str,
    *,
    revision_before: str | None = None,
    force: bool = False,
    track_changes: bool = False,
) -> dict[str, Any]:
    from . import tracked_changes
    from .author import resolve_author_name

    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    # find_sections (WP-03's PR #2 fix) now lists text-box sub-scopes
    # alongside heading sections, both keyed by section_key, so a caller
    # could hand this tool a "textbox-<n>" key straight from that listing.
    # locate_section_range below only ever recomputes HEADING positions
    # over body_children — it does not walk into w:txbxContent — so a
    # genuine textbox-<n> key would almost always miss and raise
    # SECTION_NOT_FOUND on its own. But a heading whose own text slugifies
    # to "textbox-<n>" (e.g. a heading literally titled "Textbox") would
    # collide with that reserved form, and locate_section_range has no way
    # to tell the two apart. Rather than depend on that near-miss, refuse
    # explicitly: replacing a text box's content through a section-RANGE
    # write is not what this tool means (a text box is not part of the
    # body's top-level flow read_document(format=markdown) or
    # find_sections' heading ranges address) — read_document(section_key=
    # ...) is the (read-only, for now) tool that scopes to a text box.
    textbox_keys = {t["section_key"] for t in projection.iter_textbox_scopes(resolved, DEFAULT_PART)}
    if section_key in textbox_keys:
        raise _make_error(
            ErrorCode.INVALID_INPUT,
            f"section_key {section_key!r} names a text box, not a heading-delimited body "
            "section. replace_range_markdown only rewrites body ranges found by "
            "find_sections' heading scan; a text box's content is out of scope for this "
            "tool (there is no write path for it yet).",
            {"section_key": section_key, "available_textbox_keys": sorted(textbox_keys)},
        )
    pre_revision = _guard_before_write(resolved, revision_before)

    document_root, raw_xml = _load_document(resolved)
    body = _find_body(document_root)
    # A section (find_sections' sense) never includes the body's trailing
    # sectPr, so it is dropped from body_children and never re-appended —
    # unlike replace_body_markdown/append_markdown, which must preserve it.
    body_children, _sect_pr = _split_body(body)

    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    numbering_index = projection.load_numbering_index(resolved)

    start, end = locate_section_range(body_children, section_key, styles_by_id)
    target_elements = body_children[start:end]
    before_text = projection.markdown_from_elements(target_elements, styles_by_id, numbering_index)

    own_author = resolve_author_name()
    hazards = _scan_range_hazards(target_elements)
    orphaned_comment_ids = _check_hazards_or_raise(hazards, force, own_author)
    # PR #3 review, should-fix #2: a comment whose span CROSSES the section
    # boundary has one half inside target_elements (already accounted for
    # in orphaned_comment_ids and about to be removed below) and the other
    # half in the surviving document — left alone, that would be an
    # unpaired commentRangeStart/End/commentReference. Strip both halves
    # of every orphaned id from the surviving elements too; a no-op for
    # any id that was already fully inside the removed range.
    surviving_elements = body_children[:start] + body_children[end:]
    _strip_comment_anchors_by_id(surviving_elements, set(orphaned_comment_ids))

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    new_elements = markdown_to_ooxml.render_blocks(markdown, ctx)
    # See replace_body_markdown's identical comment: verify against the
    # rendered PREVIEW of these elements, not the raw markdown text — and
    # (also as there) new_elements' own paragraphs only reference numIds
    # markdown_to_ooxml just allocated, not yet in this document's
    # numbering.xml on disk.
    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_preview = projection.markdown_from_elements(new_elements, styles_by_id, new_numbering_index)

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # issue #28 plan WP-07b-a: mark the section's OLD content deleted
        # (kept in place, runs wrapped in w:del) instead of removing it,
        # and insert the NEW content (wrapped in w:ins) right after it --
        # see execute_replace_body_markdown's identical comment for the
        # table/list scope limit this also carries.
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.mark_elements_deleted(target_elements, track)
        tracked_changes.wrap_all_runs(new_elements, _ins_wrap)
        insert_at = end  # new content lands AFTER the kept, del-wrapped old range
        for offset, el in enumerate(new_elements):
            body.insert(insert_at + offset, el)
    else:
        for el in target_elements:
            body.remove(el)
        insert_at = start
        for offset, el in enumerate(new_elements):
            body.insert(insert_at + offset, el)

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = _load_document(written_path)
        new_body = _find_body(new_document_root)
        new_body_children, _ = _split_body(new_body)
        rewritten_range = new_body_children[insert_at : insert_at + len(new_elements)]
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        # written_path is the ALREADY-REPLACED original (atomic_replace_docx_parts'
        # docstring), so its numbering.xml on disk now includes ctx's new
        # lists — a fresh disk read resolves them correctly here.
        written_numbering_index = projection.load_numbering_index(written_path)
        actual = projection.markdown_from_elements(rewritten_range, new_styles_by_id, written_numbering_index)
        diff = _diff_modulo_whitespace(intended_preview, actual)
        if diff:
            raise ValueError(f"re-read section does not match the intended rendering modulo whitespace: {diff}")

    _write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    final_document_root, _ = _load_document(resolved)
    final_body = _find_body(final_document_root)
    final_body_children, _ = _split_body(final_body)
    after_text = projection.markdown_from_elements(
        final_body_children[insert_at : insert_at + len(new_elements)],
        {s["style_id"]: s for s in projection.list_styles_impl(resolved) if s["style_id"]},
        projection.load_numbering_index(resolved),
    )

    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=3,
        before=before_text,
        after=after_text,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=orphaned_comment_ids,
        track=track,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="replace_range_markdown", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence


# ---------------------------------------------------------------------------
# Tool 3: append_markdown
# ---------------------------------------------------------------------------


def execute_append_markdown(
    path: str, markdown: str, *, revision_before: str | None = None, force: bool = False, track_changes: bool = False
) -> dict[str, Any]:
    # force/hazard-scanning is a no-op for append (nothing existing is
    # removed) but the parameter is kept for signature symmetry with the
    # other two mutating tools and so a caller's generic retry code can
    # pass force uniformly.
    from . import tracked_changes
    from .author import resolve_author_name

    own_author = resolve_author_name()
    resolved = paths.resolve_allowed_docx_path(path, must_exist=True)
    pre_revision = _guard_before_write(resolved, revision_before)

    document_root, raw_xml = _load_document(resolved)
    body = _find_body(document_root)
    body_children, sect_pr = _split_body(body)

    styles = projection.list_styles_impl(resolved)
    styles_by_id = {s["style_id"]: s for s in styles if s["style_id"]}
    numbering_index = projection.load_numbering_index(resolved)
    before_text = projection.markdown_from_elements(body_children, styles_by_id, numbering_index)

    ctx = markdown_to_ooxml.StyleContext.build(resolved)
    new_elements = markdown_to_ooxml.render_blocks(markdown, ctx)
    # See replace_body_markdown's identical comment: verify against the
    # rendered PREVIEW of these elements, not the raw markdown text — and
    # (also as there) new_elements' own paragraphs only reference numIds
    # markdown_to_ooxml just allocated, not yet in this document's
    # numbering.xml on disk.
    new_numbering_index = projection.numbering_index_from_elements(ctx.new_abstract_nums, ctx.new_nums)
    intended_preview = projection.markdown_from_elements(new_elements, styles_by_id, new_numbering_index)

    track = tracked_changes.TrackContext(document_root, author=own_author) if track_changes else None
    if track:
        # issue #28 plan WP-07b-a: nothing existing is removed by an
        # append, so there is no w:del side here -- only the newly
        # appended content is wrapped in w:ins.
        def _ins_wrap(r: Any) -> Any:
            return tracked_changes.wrap_insertion(r, rid=track.next_id(), author=track.author, date=track.date)

        tracked_changes.wrap_all_runs(new_elements, _ins_wrap)

    insert_at = len(body_children)
    for offset, el in enumerate(new_elements):
        body.insert(insert_at + offset, el)
    if sect_pr is not None:
        body.remove(sect_pr)
        body.append(sect_pr)

    def _post_verify(written_path: Path) -> None:
        new_document_root, _ = _load_document(written_path)
        new_body = _find_body(new_document_root)
        new_body_children, _ = _split_body(new_body)
        appended = new_body_children[insert_at : insert_at + len(new_elements)]
        new_styles = projection.list_styles_impl(written_path)
        new_styles_by_id = {s["style_id"]: s for s in new_styles if s["style_id"]}
        # written_path is the ALREADY-REPLACED original; its numbering.xml
        # on disk now includes ctx's new lists.
        written_numbering_index = projection.load_numbering_index(written_path)
        actual = projection.markdown_from_elements(appended, new_styles_by_id, written_numbering_index)
        diff = _diff_modulo_whitespace(intended_preview, actual)
        if diff:
            raise ValueError(f"re-read appended range does not match the intended rendering modulo whitespace: {diff}")

    _write_and_verify(resolved, document_root, raw_xml, ctx, post_verify=_post_verify)

    post_revision = projection.compute_revision(resolved)
    after_text, _, _ = projection.read_document_markdown(resolved)

    evidence = _evidence(
        applied=True,
        match_count=1,
        rung=4,
        before=before_text,
        after=after_text,
        revision_before=pre_revision["token"],
        revision_after=post_revision["token"],
        audit_logged=False,
        orphaned_comment_ids=[],
        track=track,
    )
    logged, _ = audit.append_audit(path=str(resolved), tool="append_markdown", evidence=evidence)
    evidence["audit_logged"] = logged
    return evidence
