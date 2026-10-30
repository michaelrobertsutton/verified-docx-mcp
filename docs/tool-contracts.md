# Tool contracts

Full contracts for the verified-docx-mcp tools whose MCP descriptions were trimmed (issue #108).
The description each client sees is the short form in `server.py`; each `## tool_name` section
below holds the complete text it was trimmed from, kept verbatim (so it still carries issue
numbers and implementation notes). The first section states what the write tools share.

## Common write parameters

Most write tools share these parameters and behaviours. Each tool's description still names the
parameters and error codes that apply to it.

- `revision_before`: the revision token from a prior read (`<doc8>:<cmt8>` for file mode,
  `live:sha256:<hex>` for live mode). A stale token refuses with REVISION_CONFLICT (file) or
  LIVE_STALE (live) and nothing is written.
- `write_mode`: `auto` (default) | `file` | `live`. `live` edits the document open in Word through
  the Live pane (LIVE_UNAVAILABLE if no pane is connected for this document). `file` writes the
  .docx but refuses with LIVE_SESSION_ACTIVE while a pane session is connected for it, rather
  than writing under Word's autosaving buffer. `auto` goes live whenever a pane session is
  connected for the file name. LIVE_SESSION_MISMATCH: a session shares the file name but names a
  different local file. LIVE_DISCONNECTED: the socket closed or an op timed out.
  LIVE_CAPABILITY_MISSING: the pane predates the needed capability; reload the pane.
- Live evidence differs from file evidence: revisions are `live:sha256:<hex>` body hashes, it adds
  `write_mode`, `verified_via: "word-addin"`, `document_name` and the signed-in Word user as author,
  and it has no `conflict_copy_detected` / `conflict_copies` / `sibling_files_changed`. `force` and
  `revision_ids` have no live meaning.
- Write guard (file mode, in order): owner file -> DOCX_LOCKED (no retry); not sync-quiesced ->
  bounded wait, then SYNC_IN_FLIGHT; stale `revision_before` -> REVISION_CONFLICT (never
  bypassed); file changed since this server's last write within the observation window (default
  600 s, `VERIFIED_DOCX_MCP_EXTERNAL_EDIT_WINDOW_S`) or an unreadable write ledger ->
  EXTERNAL_EDITOR_ACTIVE.
- `allow_concurrent_editor=True` bypasses ONLY EXTERNAL_EDITOR_ACTIVE (not REVISION_CONFLICT,
  DOCX_LOCKED or LIVE_SESSION_ACTIVE) and records `concurrent_editor_override` in the evidence.
- `force=True`: proceed past a comment-anchor / tracked-change hazard
  (COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT); the removed comment ids are reported as
  `orphaned_comment_ids`.
- `track_changes=True`: wrap deletions in `w:del` and insertions in `w:ins` under the configured
  author (`author_name` in `~/.jennystack/config.json`, else the macOS full name); live mode
  records as the signed-in Word user. The evidence adds `revision_ids` and `track_changes`.
- Atomic write: a temp file is OPC-validated before the original is touched (OPC_INVALID leaves
  the original untouched); the original is swapped in with a `.jsbak` kept until a post-write
  re-read confirms the change; a failure restores it (VERIFICATION_FAILED).
- Evidence: the eight keys `applied`, `match_count`, `rung`, `before`, `after`, `revision_before`,
  `revision_after`, `audit_logged`, plus `conflict_copy_detected` in file mode (an evidence flag
  from the post-write conflict-copy sweep, never raised; `conflict_copies` and
  `sibling_files_changed` only when non-empty).

## export_pdf

Render a local .docx to PDF via Microsoft Word and report its page count.

output_path must fall inside VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS
(defaults to the user's home directory plus the Claude Code scratch
root, /private/tmp/claude-<uid>, when that directory exists), must not
resolve to a credential path, and its parent directory must already
exist. This is a read/render tool: the source .docx is never modified
(Word opens a private staged copy — see render.py), so the return value
has no "applied" key.

The staged copy Word renders from is a private file this server created
inside Word's own sandbox container; it is never the source document and
is deleted from disk in a `finally` block regardless of what happens to
its Word window. close_after (default True) closes that window's copy
after the PDF is written and the page count is read, so nothing is
discarded by closing: the PDF is already on disk first. A close failure
is reported in close_error rather than failing the render — the PDF and
page_count are still returned. Pass close_after=False to leave the
window open on purpose (e.g. to inspect it).

section_keys (issue #102, optional): restrict per-section page-span
reporting to these find_sections() section_key values (find_sections'
"heading" entries only — a "textbox" key is never probed and is never
a valid section_keys entry). Omit it (the default, None) to report
every heading section the document has. Every probe this needs runs
inside the SAME osascript call as the render, right after the page
count is read and before the document is closed, so the pagination
the probes see is identical to page_count's — never a second Word
round trip with its own, possibly different, layout.

Each entry of the returned "sections" list is one heading section's
page span: section_key (matches find_sections), heading_text,
start_page/start_fraction and end_page/end_fraction (the vertical
position on those pages as a fraction of page_height_pt — a partial
page is reported as a fraction, never rounded away), pages (the
derived span, end - start), start_paragraph/end_paragraph (that
section's own 1-based Word paragraph ordinals, from
projection.find_sections_impl's start_para_ref/end_para_ref — see
geometry.ordinal_of), and text_verified (always true on an entry that
made it into this list — see below). A section's end is the FULL
DOCUMENT's next heading's own start — not merely the next entry in
this call's own "sections" list — whenever the document has one, even
when section_keys narrowed the request down to just that one section
and its true next heading is not itself part of the returned list (so
a single-section, page-budget-style section_keys=[...] call still
gets an exact boundary, not an approximation, whenever a following
heading exists). Only a section with no next heading AT ALL — the
document's own last heading — uses its own last paragraph's probe,
with end_fraction nudged by +0.03 of a page (capped at 1.0) to
approximate that last line's own height. The borrowed next heading's
probed text is never verified against anything — only the start
position it supplies is used, purely as boundary geometry.

Every probed section's start paragraph must verify: its probed text,
after geometry.normalize_probe_text(), must equal that section's
heading_text, also normalized. Default mode (section_keys=None)
degrades on any verification failure or probe error rather than
failing the whole call: "sections" comes back null and
"sections_error" names which ordinal and what mismatched, while
pdf_path, page_count, and every other field are unaffected. Explicit
mode (section_keys given — the caller asked for these sections by
name) raises SECTION_GEOMETRY_UNAVAILABLE instead, carrying the same
detail in diagnostics, since a number the caller asked for by name
that this server cannot stand behind should not come back silently
as null. A document with zero headings (or an empty section_keys
filter after SECTION_NOT_FOUND validation) returns sections: [] and
sections_error: null — nothing to verify, nothing probed. Top-level
page_height_pt is the document's own page height in points (None
when nothing was probed).

Returns pdf_path, sha256, page_count (best-effort; None — never a
guessed 0 — when it cannot be determined), page_count_source
("word"|"pdfinfo"|"regex"|None), engine ("word"; the only engine, D2:
no LibreOffice), left_open_document (null when the window was closed,
else its title), closed_after (bool — whether the close succeeded),
close_error (the close attempt's error text, or null), page_height_pt,
sections, and sections_error (see above).

Errors:
  INVALID_INPUT              - a bad path or output_path
  RENDER_ENGINE_UNAVAILABLE  - no render engine available (Word only)
  AUTOMATION_NOT_GRANTED     - macOS declined Automation control of Word
                                for this app; run `verified-docx-mcp
                                doctor` for the fix, scoped to the app
                                hosting this MCP server's own process
  WORD_SANDBOX_UNAVAILABLE   - Word has never been launched on this
                                machine (its sandbox container does not
                                exist yet)
  RENDER_FAILED              - any other Word automation failure
  SECTION_NOT_FOUND          - a section_keys entry does not match any
                                current heading section (checked before
                                any render happens)
  SECTION_GEOMETRY_UNAVAILABLE - section_keys was given and a requested
                                section's page-span could not be
                                verified (see above)

## lock_status

Report Word/LibreOffice owner-file presence and sync-quiesce state.

Data only — never refuses. Owner-file detection recognizes Word's
"~$*" owner file (matched by shared filename suffix, not by
constructing one exact expected name — see
_word_owner_file_matches's docstring) and LibreOffice's
".~lock.<name>#". Sync-quiesce takes two (size, mtime_ns) samples
~1.5s apart and also checks for "*.tmp"/".~*" siblings in the same
directory (OneDrive/iCloud staging artifacts); "sync_quiesced" is
false if either signal suggests the file is still being written.

Returns (default) owner_file ({present, path, format, owner_name}),
desktop_word, sync_quiesced, and external_activity reduced to
{divergent, still_current, recent, ledger} (what the write guard decides on). verbose=True also returns path, sync_detail
({sample_1, sample_2, interval_seconds, tmp_siblings}) and the full
external_activity (issue #27): how the file on
disk compares with what THIS server last wrote --
{ledger: "ok"|"none"|"unreadable", ledger_reason, last_write
({token, written_at}), still_current (null with no ledger),
divergent, divergence_first_observed_at, divergence_age_s, recent,
unattributed_recent_mtime, window_s, note}. This is the after-the-fact
durability check for a co-authoring editor such as Office Online, which
writes no owner file and registers no Live session: call it once a
write has had time to sync -- still_current false means another
writer changed the package since, i.e. this server's edits may have
been overwritten. It shows another writer changed the bytes, not that
an editor is still open; no divergence is not proof of safety, and a
file this server never wrote reports ledger "none" (with
unattributed_recent_mtime as an informational hint only). Reading it
persists the divergence's first-seen time (a small state write).

Errors:
  INVALID_INPUT - path does not exist or is outside the allowed roots

## live_list_sections

List the live document's heading sections as the Word pane sees them.

Each entry has ``section_key`` (the find_sections convention: heading slug plus
ordinal), ``heading_text``, ``level``, ``paragraph_count``, ``slug_unique`` and
``sectionSha256``. Only sections with ``slug_unique`` true can be locked with
live_section_lock. Needs a pane that reports ``section_locks``; works with or
without the shared bridge.

``compact=true`` returns ``{sections: [[section_key, heading_text, sha12], ...],
body_sha12}`` instead (no 64-hex hashes, slugs or counts, so much smaller): sha12 is
the first 12 hex of ``sectionSha256``, which find_sections(digests=true) also
reports for unchanged text. Section locks still key on the full hash internally.

``content_digests=true`` adds ``content_sections: [[section_key, content_sha12],
...]`` and ``preamble_content_sha12`` (see find_sections) beside ``sections``,
computed server-side from one body_ooxml export taken after the section list
(LIVE_OP_FAILED if the body changed between the two and a re-list does not
match). content_sections is keyed by the server's own section keys; if they
differ from the pane's, ``warnings`` carries ``content_sections_key_mismatch``
and ``content_sections_key_mismatch`` lists them (never silently paired).
Export fidelity warnings (``live_styles_missing``, ...) also go in
``warnings``. Like the list itself, this never re-arms a shared-bridge write
baseline.

## live_revision

The live document's current revision and per-section digests, without
reading the body.

Returns ``{revision: "live:sha256:<64 hex>", body_sha12, sections:
[[section_key, sha12], ...], observed_at, via}``. ``revision`` is exactly the
token read_document(source="live") reports and live writes accept as
``revision_before``. Calling this re-arms your write baseline the same way a
live read_document does (shared-bridge mode), so use it to resume after
LIVE_STALE when you have no reason to re-read the text -- and read the text
instead if ``sections`` shows a section you care about changed.

One pane round trip (``via: "body_hash"``, pane capability ``body_hash_v1``).
An older pane still works but transfers the whole body to this server and
discards it (``via: "body_ooxml_fallback"``): correct, not cheap. Reload the
pane to get the fast path.

``content_digests=true`` also returns ``content_sections: [[section_key,
content_sha12], ...]`` and ``preamble_content_sha12`` (see find_sections):
one body_ooxml export (``via: "body_ooxml"``) from which every value in the
reply is computed, so they describe the same snapshot; ``sections`` then
comes from this server's projection of the export, not from the pane. Costs a body transfer;
use it to decide which sections to sync, and sha12 for "did the prose change".

Errors: LIVE_UNAVAILABLE (no pane), LIVE_SESSION_MISMATCH, LIVE_OP_FAILED,
LIVE_CAPABILITY_MISSING (the pane offers neither body_hash_v1 nor body_ooxml,
or no body_ooxml with content_digests=true), INVALID_INPUT (comment threading
that cannot be interpreted, content_digests=true only).

## live_status

Report live-bridge state: whether it is running, its ports, and
the connected Word task-pane sessions (all of them, or just *path*'s).

Starts the bridge lazily (idempotent -- a second call, or any future
``write_mode="live"`` tool call from WP-3/WP-4, reuses the same
running bridge) so the lead's pane has something to connect to the
first time this is called; does NOT itself require a pane to already
be connected -- see ``sessions: []`` below.

Read-only: never modifies a .docx, so it is not in ``MUTATING_TOOLS``
(middleware.py).

path (optional) scopes ``sessions`` to that document's file name.
verbose=True adds the per-session ``heartbeat_count``,
``hash_sample_count`` and ``requirement_sets`` plus the top-level
``recent_boots`` and broker ``shared`` blocks. In shared-bridge mode the
result always carries ``locks`` (document leases and section locks:
``{document_url, client_id, scope, remaining_s, ...}``, filtered to
*path*'s document when given) plus this client's ``client_id``: read
them to see who holds a lock after LOCKED_BY_OTHER_CLIENT.

Returns ``bridge_running``, ``pid``, ``owner_pid``, ``is_owner``,
``mode`` ("local"|"shared"; stated once, never repeated in nested
blocks), ``port`` (static HTTPS pane server), ``ops_port`` (WSS
``/ops``), ``sessions`` (one per connected pane, ``[]`` before any pane
connects; each ``{document_name, document_url, connected_since,
last_heartbeat_age_s, body_sha256, body_hash_state, body_hash_age_s,
instance_id, platform, revision_count, revision_coverage}``),
``session_collisions`` (two DIFFERENT documents (different
``document_url``) that connected under the same file name, which the
registry's basename-only keying cannot tell apart; each entry is
``{document_name, displaced_document_url, new_document_url, at}``;
``[]`` when none; never itself a refusal, but check it before trusting
a live write's routing), and ``warnings`` (e.g. ``pane_reload_loop``)
only when there is something to say.

Errors:
  LIVE_UNAVAILABLE - the bridge itself failed to start (e.g. no
    `openssl` on PATH for a first-run cert, or its ports are already
    bound by something else) -- distinct from "no pane connected
    yet", which is not an error.

## read_document

Read a .docx part's content as markdown, flat text, or a run-level
structural list.

source="auto"|"file"|"live" (issue #33) picks WHERE the content comes
from. "live" reads the body the connected Word pane currently holds
(body only: part must be the default word/document.xml, and the pane must
report the body_ooxml capability); "file" reads the local .docx; "auto"
(default) reads live when a pane session exists for this file name and
the read can be served live, otherwise the file. The response always
carries "source". A file read while a pane session exists for the name
also carries the "live_session_ignored" warning and a live_session
{document_url, reason, action} object, so a same-named placeholder file
can never pass as the live body. live_session.reason says why the live body
was not read: "requested_file" (pass source="live" or "auto" to read the live
document), "part_not_in_live" (live reads cover the document body only),
"pane_missing_capability" (older pane build: close and reopen the Live pane in
Word, then retry), "basename_collision" (another document connected under this
file name: confirm which is open, then pass source="live" explicitly) or
"session_mismatch" (the pane's document_url names a different local file; this
read is of the file you named). A live read never falls back to the
file: a pane failure is an error. In a live read revision is
"live:sha256:<body hash>" (the token live writes accept as
revision_before), revision_detail is absent, and "live" carries
{document_url, stripped_parts} (the body hash is the revision);
images are not included.
A live read may also warn "live_styles_missing" or
"live_numbering_missing" when Word's export omitted a part the body
references. para_ref and table_id values are positional: use them only
with the same source and revision they were read at.

part scopes the read to one package part — word/document.xml's body
(the default) is a completely separate scope from a header, footer,
footnotes, or endnotes part. Call list_parts first to enumerate the
parts actually present in this .docx.

A text box's own w:txbxContent is a further, separate sub-scope never
merged into its host part's text (see projection.py's module
docstring) — call find_sections first to get its section_key
("textbox-<n>"), then pass that section_key here to read JUST that
text box's content, scoped to it exactly as if it were its own part.
Omitting section_key (the default) reads the whole part named by
part, in which a text box's own text is absent (by design — it would
otherwise be double-counted between the host part and the text box's
own scope). A heading section_key (also from find_sections) reads that
section's paragraph range only, in any format, file or live, with
original para_ref/table_id values; the response adds start_para_ref and
end_para_ref. A heading read is a content read of that range only, not
proof of full-document or comment freshness. An unknown key is
INVALID_INPUT listing the available heading and text-box keys.

format="text" returns the flat projected string (runs concatenated in
document order; a paragraph break is "\n"). format="runs" returns, in
document order, run records {text, rPr, para_ref} interleaved with
structural records: {"type":"table_start"|"table_end", table_id},
{"type":"drawing", blip_rid, media_part, extent_in, para_ref} for every
w:drawing/a:blip, and {"type":"field", instr, result_text}.
format="markdown" (default) renders headings, bold/italic, bulleted/
numbered lists (nested by indent), and GFM pipe tables — matching
GoogleDocs-MCP's markdown.py conventions exactly (issue #28 WP-03b-a)
— plus stable placeholder tokens ("[image:rId]", "[field:instr]") for
a drawing / a field with no result. WP-04's inverse (markdown ->
OOXML) is not implemented here. A merged table cell (w:gridSpan/
w:vMerge) or a nested w:tbl has no pipe-table representation and is
reported in lossy_elements ({"kind": "table_merge"|"nested_table",
"table_id"}) instead of silently reproduced as an ordinary grid cell.

A deleted span (w:del/w:delText) and a field's own instruction text
(w:instrText) are excluded from every format; a field's RESULT text
(between fldChar "separate" and "end", or all of a w:fldSimple's
nested runs) IS included — in format="markdown" it flows in as
ordinary rendered text, never duplicated by a "[field:...]" token
(that placeholder appears only when a field carries NO result at all).

start_para_ref / end_para_ref (together, inclusive, positional "pN"
exactly as find_sections reports them) return only that paragraph
range, in any format, file or live. A malformed, out-of-range or
inverted range is INVALID_INPUT. A table cut by the range keeps only
the rows that hold in-range paragraphs and is reported in
lossy_elements as {"kind": "partial_table", "table_id"} (a table nested
inside a kept row is shown whole in markdown); an ordered
list's numbering restarts at the slice. Cannot be combined with
section_key.

to_file writes the rendered text/markdown/runs (runs as JSON) to that
path (allowed roots only; the parent directory must exist; the name must
end in .md, .txt or .json; an existing file is refused unless
overwrite=true, so a read can never replace a .docx or source file) and returns
NO body: {path, part, format, revision, written_to, sha256 (of the
written file), chars, paragraph_count, sections: [[section_key,
start_para_ref, end_para_ref, chars], ...] (the first 25 heading
sections only), section_count, sections_truncated (present when more
exist: call find_sections for the rest), warnings}. Use it for a large
document you only need to grep, hash or hand to another tool. A live
to_file or ranged read still counts as this client's baseline read
(the server saw the whole body).

Returns path, part, format, revision (the "<doc8>:<cmt8>" token), plus
text|runs|markdown per format. section_key (echoed back), warnings and
lossy_elements (format="markdown" only) appear only when non-null /
non-empty; revision_detail (the full {document_sha256, comments_sha256,
size, mtime_ns} tuple, file reads only) only with verbose=True; in a live
read "live" carries {document_url, stripped_parts} (its body hash is the
revision).

Not gated by DOCX_LOCKED — see list_parts' docstring.

Errors:
  INVALID_INPUT   - a bad path, an unrecognized format, or a
                     section_key that names no text box in this part
  PART_NOT_FOUND  - part names a package part absent from this .docx
  SNAPSHOT_FAILED - the read-path snapshot could not be validated

## find_sections

List heading-delimited section ranges AND text-box sub-scopes in a
.docx part — this is the one call that discovers every section_key
read_document(section_key=...) can then address.

source="auto"|"file"|"live" works as in read_document (issue #33): the
response carries "source", and a file read while a live pane session
exists for this file name carries the "live_session_ignored" warning.

DISAMBIGUATION: this is about DOCUMENT SECTIONS (heading ranges, by
style/outline level) — see list_page_sections for PAGE-LAYOUT sections
(w:sectPr page size/margins/columns). The two are unrelated OOXML
concepts that happen to share the English word "section".

A heading paragraph's w:pStyle is mapped through list_styles to an
outline level, falling back to the paragraph's own direct
w:outlineLvl. A section runs from one heading (inclusive) to the next
heading at any level, or the end of the document. section_key =
slug(heading text) + a 1-based ordinal disambiguating duplicate
headings. Headings inside a table cell are out of scope (not a real
document section boundary).

Every w:txbxContent in the part (see read_document's docstring) is
also listed, each as its own entry with kind="textbox" and
section_key="textbox-<n>" (heading_text/outline_level/start_para_ref/
end_para_ref are None on these entries; paragraph_count is still the
text box's own paragraph count). Pass one of these section_keys to
read_document to read that text box's content.

Returns path, part, sections (list of {section_key, kind
("heading"|"textbox"), heading_text, outline_level, start_para_ref,
end_para_ref, paragraph_count, chars}). chars is the section's size in
plain-text characters (paragraph texts joined with "\n"), so a caller
can judge whether reading it is affordable. With digests=true each
heading entry also carries sha12, the first 12 hex of the sha256 of the
section's paragraph texts joined with "\n" (table-cell paragraphs
included, no normalization; a soft line break counts as Word's U+000B) --
defined to agree with the live pane's section digests (live_list_sections
compact=true, live_revision) for unchanged text. Checked against real Word for
plain text, soft line breaks and <w:tab/> tabs; field results, tracked deletions
and a literal tab character typed inside a w:t are NOT verified and may differ.
sha12 is what section locks compare; use it for "did the prose change".

digests=true also adds content_sha12 to each heading entry and a top-level
preamble_content_sha12 (content before the first heading; the whole part if
it has none): a digest of what sync acts on -- paragraph texts, tracked
deletions (text and position), table structure (rows, cells, gridSpan,
vMerge, nested tables) and the root comments anchored in the section (author,
text, resolved, replies in order, orphaned replies). Formatting, images, text
boxes, footnotes/endnotes, rsids/paraIds, bookmarks, proofing marks and
comment dates/ids are excluded; so are accepting a tracked insertion (it is
plain text), anchor moves within the same paragraph, tblGrid column count,
gridBefore/gridAfter, legacy hMerge and tracked paragraph-mark deletions.
Computed only by this server from OOXML (a live read uses the pane's
body_ooxml export); file and live agreed on Word-saved documents in the
real-Word probe (docs/acceptance/117-content-digest.md), though Word's
"remove personal information on save" rewrites saved comment authors.
live_revision / live_list_sections return the same value with
content_digests=true. When the
comments or tables cannot be canonicalized (duplicate comment ids, a parent
cycle, a non-numeric gridSpan, ...), content_sha12 and preamble_content_sha12
are null, sha12 is still returned, and warnings carries
"content_digest_unavailable" with the reason in content_digest_unavailable.

Not gated by DOCX_LOCKED — see list_parts' docstring.

Errors:
  INVALID_INPUT   - a bad path
  PART_NOT_FOUND  - part names a package part absent from this .docx
  SNAPSHOT_FAILED - the read-path snapshot could not be validated

## replace_body_markdown

Replace an entire document body's content with markdown, atomically.

markdown_file (issue #113): a path to a .md/.txt file (UTF-8, <= 2 MB, under the
allowed roots; outside them -> DOCX_PATH_ESCAPE) used INSTEAD of `markdown`, so text
already on disk need not be re-emitted. Exactly one of the two (else INVALID_INPUT).
The text is read once, server-side, before any routing; the response adds
markdown_sha256 (hex of its UTF-8 bytes) and markdown_chars.

Rung 4 (last resort) of core/document-backend-protocol.md's backend-
neutral edit ladder — the whole-document rewrite. Prefer
replace_range_markdown when find_sections can locate the target
section.

Markdown is rendered against THIS document's own styles: headings
resolve through list_styles (STYLE_NOT_FOUND if a level has no
matching style — never a hardcoded "Heading2"), tables pick up the
document's own table style if one exists, and every bulleted/ordered
list gets a freshly allocated numbering definition. Supports
bold/italic/***both***, links, lists (nested), and pipe tables; a
thematic break, blockquote, or code fence degrades gracefully (see
markdown_to_ooxml's module docstring) rather than failing the call.

Guard order (this WP's whole point — Codex defect: an earlier draft
let writes ship before the guard): lock_status runs FIRST, before any
temp file exists. An owner file present -> DOCX_LOCKED (report the
owner, no retry). Not sync-quiesced -> one bounded wait (<=10s total)
then SYNC_IN_FLIGHT. A revision_before that no longer matches the
file's current revision -> REVISION_CONFLICT (never bypassed by
anything below). Then (issue #27) a file that no longer matches what
this server last wrote -- another writer, plausibly Office Online
co-authoring, changed it within the observation window (default 600s,
VERIFIED_DOCX_MCP_EXTERNAL_EDIT_WINDOW_S), or this server's own
write-ledger record for it is unreadable -> EXTERNAL_EDITOR_ACTIVE,
because that editor's autosave would likely silently revert this
write. `allow_concurrent_editor=True` bypasses ONLY that check (not
REVISION_CONFLICT, DOCX_LOCKED, or LIVE_SESSION_ACTIVE) and records
the override in the evidence (concurrent_editor_override). The window
is measured from when this server first observed the divergence;
waiting it out is a heuristic, not proof the editor has gone, and a
file this server never wrote has no ledger entry to compare against
(lock_status reports it as ledger "none"). Also, after a write
returns, call lock_status once edits have had time to sync:
external_activity.still_current false means something overwrote them.
If the CURRENT body
contains comment anchors or tracked changes, the call refuses
(COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT) unless
force=True; with force, they are removed and their comment ids are
reported as orphaned_comment_ids (nothing disappears silently ahead
of WP-08's comment tools).

The write itself is atomic: a temp file is built and OPC-validated
(XML well-formed, every r:id resolves, [Content_Types].xml covers
every part) BEFORE the original is ever touched (a failure here ->
OPC_INVALID, original untouched); the original is then swapped in via
os.replace() with a .jsbak kept until a post-write re-read/re-project
confirms the change (a failure here restores from .jsbak and raises
VERIFICATION_FAILED).

Residual risk (layers 3-4, issue #28 WP-10, wrap the atomic write
itself -- see the module-level comment above): a write that lands
while another client has the file open is not PREVENTED by anything
above, only detected after the fact, by the post-write conflict-copy
sweep on THIS SAME call (conflict_copy_detected in the evidence
below) or a later one.

track_changes=True (issue #28 WP-07b-a): the OLD body content is kept
(not removed) with its runs wrapped in w:del/w:delText, and the NEW
content is wrapped in w:ins -- both carrying w:author (author.
resolve_author_name())/w:date/an above-package-maximum w:id. Reading
the projection is unaffected either way, so after_text still reads as
the new content alone. Scope limit: this wraps RUNS, not whole table/
list structures -- untested against a target range containing a
table. The comment/tracked-change hazard scan's TRACKED_CHANGES_PRESENT
excludes a revision authored by the configured author (the server's
own prior tracked write), so a second track_changes call never
deadlocks on the first one's own tracked change.

Returns the eight evidence keys: applied, match_count (always 1),
rung (4), before, after, revision_before, revision_after,
audit_logged; plus orphaned_comment_ids when force removed anchors,
and (track_changes=True only) revision_ids/track_changes.

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (see above)
  REVISION_CONFLICT                 - revision_before is stale, or the file changed while
                                      this edit was being prepared (nothing written)
  EXTERNAL_EDITOR_ACTIVE            - issue #27: see above; override with
                                      allow_concurrent_editor=True
  COMMENT_ANCHORS_IN_RANGE          - comment anchors present, no force
  TRACKED_CHANGES_PRESENT           - w:ins/w:del present, no force
  STYLE_NOT_FOUND                   - a heading level has no style
  OPC_INVALID                       - the rendered .docx failed OPC validation
  VERIFICATION_FAILED               - post-write verification failed; rolled back (or, when
                                      another writer changed the file after this write
                                      landed, NOT rolled back: details.rollback_skipped)

## replace_range_markdown

Replace one heading-delimited section (by section_key, from
find_sections) with markdown, atomically.

markdown_file (issue #113): a path to a .md/.txt file (UTF-8, <= 2 MB, under the
allowed roots; outside them -> DOCX_PATH_ESCAPE) used INSTEAD of `markdown`, so text
already on disk need not be re-emitted. Exactly one of the two (else INVALID_INPUT).
The text is read once, server-side, before any routing; the response adds
markdown_sha256 (hex of its UTF-8 bytes) and markdown_chars.

Rung 3 of core/document-backend-protocol.md's edit ladder: use this
when find_sections can locate the target section, ahead of
replace_body_markdown's whole-document rewrite. section_key is keyed
by heading slug rather than a revision-stamped range — STALE_RANGE is
a gdoc-only hazard (the protocol doc's footnote on this rung).

Same guard, hazard scan, atomic-write, and markdown rendering rules as
replace_body_markdown — see that tool's docstring for the full guard
order, hazard-refusal/force semantics, and atomic-write mechanics; the
only difference is the target range (one section's top-level body
children, from its own heading up to but not including the next
top-level heading of any level, or the end of the document) instead
of the whole body.

track_changes=True: same as replace_body_markdown (WP-07b-a) -- the
section's OLD content is kept (w:del-wrapped), the NEW content is
wrapped in w:ins and inserted right after it.

Returns the eight evidence keys (rung is always 3 here), plus
orphaned_comment_ids when force removed anchors, and (track_changes=
True only) revision_ids/track_changes.

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors: as replace_body_markdown, plus:
  SECTION_NOT_FOUND - section_key does not match any current section
                       (call find_sections to see the current ones;
                       not named in the issue #28 plan text for this
                       WP, added because replace_range_markdown needs
                       some code for this case)

## append_markdown

Append markdown to the end of a document body (before its trailing
w:sectPr, if any), atomically.

markdown_file (issue #113): a path to a .md/.txt file (UTF-8, <= 2 MB, under the
allowed roots; outside them -> DOCX_PATH_ESCAPE) used INSTEAD of `markdown`, so text
already on disk need not be re-emitted. Exactly one of the two (else INVALID_INPUT).
The text is read once, server-side, before any routing; the response adds
markdown_sha256 (hex of its UTF-8 bytes) and markdown_chars.

Nothing existing is removed, so there is no comment-anchor/tracked-
change hazard to scan and force has no effect here — it is accepted
only so a caller's generic retry code can pass it uniformly across
all three mutating tools. Same guard order (lock_status first, then
revision_before) and atomic-write mechanics as replace_body_markdown
— see that tool's docstring.

track_changes=True (WP-07b-a): nothing existing is removed by an
append, so there is no w:del side here -- only the newly appended
content is wrapped in w:ins (w:author/w:date/w:id, same source as
replace_text's).

Returns the eight evidence keys: before/after are the whole body's
markdown immediately before/after the append; rung is reported as 4
(append_markdown is not itself a rung on
core/document-backend-protocol.md's 4-rung table, which only names
replace_body_markdown at rung 4 — grouped with it here as the other
whole-document-scoped write). Plus (track_changes=True only)
revision_ids/track_changes.

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
  REVISION_CONFLICT                 - revision_before is stale
  EXTERNAL_EDITOR_ACTIVE            - see replace_body_markdown (issue #27)
  STYLE_NOT_FOUND                   - a heading level has no style
  OPC_INVALID                       - the rendered .docx failed OPC validation
  VERIFICATION_FAILED               - post-write verification failed; rolled back

## replace_text

Replace every occurrence of `find` with `replace`, atomically. `replace` is required.

Three ways to name the target (issue #113), exactly one per call:
  - plain `find` + `expected_matches` (required for a plain find).
  - anchored `find` = "<prefix>…<suffix>" (one U+2026 or "[...]"; prefix
    and suffix >= 12 characters each) + `expected_length` and/or
    `find_sha256`. Only these two make a find anchored: without them a
    literal "…" is ordinary text in a plain find. The span is normalized
    (straight quotes, whitespace runs -> one space, soft hyphens removed);
    `expected_length` counts its characters (+-2) and `find_sha256` is the
    sha256 hex of it (`locate.span_sha256`). Each prefix pairs with the nearest following suffix in the
    same paragraph; exactly one span must pass the check, else
    AMBIGUOUS_SPAN (up to 5 candidates) or ZERO_MATCH (candidate lengths).
    No retyping a long span; no 230-character live limit.
  - `para_ref` ("pN") + `expected_sha256` (same hash of that paragraph's
    text): replaces the whole paragraph's text; PARA_HASH_MISMATCH if it changed.
    No read returns this hash yet: compute it as `locate.span_sha256` does
    (a read-side paragraph hash is a follow-up).
Live mode resolves anchored/para_ref targets on a read of the live body that
never re-arms a stale caller; text-box scopes and within_row_containing take a
plain find only there. Spans over 230 characters need the pane's
"anchored_find_v1" capability (LIVE_CAPABILITY_MISSING otherwise).

`verbose=true` keeps a ZERO_MATCH error's `ladder_report` (a match count of zero per
normalization rung); without it the error carries `near_miss` only, which is the useful
part. MATCH_COUNT_MISMATCH and STRUCTURAL_BOUNDARY errors list at most five matches as
{start, end, context} plus `total_matches`.

Live-only replacement_segments specifies explicit text/format/hyperlink segments.
Text concatenates exactly to replace; each format includes bold, italic, underline
(None/Single/Double), strikeThrough, doubleStrikeThrough, subscript, superscript,
color (#RRGGBB), name, and size (points). Null hyperlink explicitly means no link.
Mixed spans require explicit segments; no neighbour-format inference is used.
A format may name only some fields (pane capability partial_font_segments_v1): the
rest are inherited from the effective font of the first replaced run, and the write
refuses FORMAT_DRIFT (nothing written) if that run's formatting changed meanwhile.

`write_mode` (issue #106 WP-3, "auto" | "file" | "live", default
"auto"): whether this call edits the .docx file directly (today's
path, unchanged) or sends the edit to a connected Word task pane over
the live bridge instead. "file" always uses the file path -- issue
#154: BUT it now refuses with LIVE_SESSION_ACTIVE if a pane session
is connected for this document, rather than writing to disk
underneath Word's own open, autosaving buffer (see that error code
below). "live" always uses the live pane, raising LIVE_UNAVAILABLE if
none is connected for this document. "auto" (issue #154) uses live
whenever a pane session is connected for this file's name -- no
lock_status/owner-file check any more (that signal is unreliable: a
desktop Word owner file never appears in a SharePoint/OneDrive-synced
folder on Mac, so the old "both must hold" rule could never route
live there at all). This is why a document being open in Word
(`DOCX_LOCKED` today) is now writable instead of refused: the open
copy is the very thing "auto" routes the edit into, live, in front of
the lead, rather than treating "someone has it open" as a reason to
refuse.

Live-mode evidence differs from file mode in three ways: (1)
`revision_before`/`revision_after` are `"live:sha256:<hex>"` of the
pane's own `body.text` hash (there is no OOXML revision token until
`live_save` writes the file back to disk); (2) `write_mode: "live"`,
`verified_via: "word-addin"`, `document_name`, and
`track_changes_author: "word-signed-in-user"` are added (the pane
cannot set `w:author` -- live track-changes authorship is always
whoever is signed into that copy of Word); (3) no
`conflict_copy_detected`/`conflict_copies`/`sibling_files_changed` --
Word owns the file for the whole live session, so there is no
sync-conflict copy for a sweep to find. `force`/`revision_ids` have
no live-mode meaning and are not produced in that mode.

Rungs 3 and 4 (`replace_range_markdown`/`replace_body_markdown`, the
structural/whole-document rungs) never go live, in either mode --
`write_mode` on THIS tool only ever chooses between file and live for
rung 1/2 edits. After a live session, structural verification against
the file on disk happens by calling `live_save` first, then the
existing read tools (`read_document`/`find_sections`/
`diff_body_vs_file`) against the now-saved file, same as any other
file-mode read.

Locates `find` via a 4-rung normalization ladder (exact -> curly/
straight quote equivalence -> NBSP/whitespace-run collapse -> soft-
hyphen strip), stopping at the first rung with at least one match. A
match overlapping a RESULT-bearing field's cached text surfaces
"touches_field_result" in the evidence's `warnings` (see locate.py's
own header comment for why this is a warning, not a match rung, on
the corrected field-markdown projection).

expected_matches is REQUIRED (no default, D4): the call refuses with
MATCH_COUNT_MISMATCH if the actual count differs, listing every span
found at the matching rung.

A run whose text a match's boundary falls in the middle of splits into
up to three pieces: the unmatched prefix and suffix keep the run's
ORIGINAL w:rPr (cloned verbatim onto a fresh sibling run when a suffix
survives); exactly one new run is inserted for `replace`, its own
w:rPr inherited from the FIRST run the match touched. A run entirely
inside the match is removed outright. A match crossing a w:p/w:tbl/
w:tc boundary refuses with STRUCTURAL_BOUNDARY instead.

track_changes=True (issue #28 WP-07b-a): the deleted text is wrapped
in w:del/w:delText and the replacement in w:ins, each carrying
w:author (author.resolve_author_name() -- ~/.jennystack/config.json's
author_name, falling back to the macOS full name) and w:date, with
w:id values allocated above the package maximum. Reading the
projection is unaffected (w:del text stays excluded, w:ins text stays
included), so the write is visible to the next read as current text
either way. The evidence then also carries revision_ids (the ids just
created) and track_changes: true. A match crossing a tracked change
authored by someone OTHER than the configured author refuses with
TRACKED_CHANGES_PRESENT unless force=True; a match crossing the
server's OWN prior tracked change never refuses (accept/reject it via
tracked_changes.py, or just keep editing under track_changes).

Same atomic-write mechanics as replace_body_markdown (temp file, OPC-
validated before the original is touched, .jsbak-backed post-write
verification) -- see that tool's docstring.

within_row_containing (issue #22, optional): scope `find` to one
table row, identified by a SECOND, unique piece of text located
inside that row. Solves editing a single cell among many identical
ones (e.g. a table where 11 cells read "No approved proof on file")
without expected_matches=1 refusing (find matches every such cell)
or expected_matches=N rewriting all of them. Contract: the anchor
itself must be unique in the WHOLE document (the same guarantee
locate() already gives any needle -- no separate per-row uniqueness
rule) and must resolve to a table cell; `find` is then located (and
its own expected_matches enforced) only within that row. Named
limitation, not solved by this: two cells in the SAME anchored row
with identical `find` text remain inseparable. In live mode, requires
the connected pane report the "row_scope" capability -- see
LIVE_CAPABILITY_MISSING below.

allow_comment_loss_ids (issue #82, optional, LIVE ONLY): per-thread
consent to delete comment threads the match overlaps, instead of the
all-or-nothing allow_comment_loss=True. A list of
{"comment_id": str, "reply_count": int} (both required); each listed
thread must be RESOLVED and still have exactly that many replies, and
any overlapping thread NOT listed refuses the edit with
WOULD_DELETE_COMMENTS. Mutually exclusive with allow_comment_loss=True
(INVALID_INPUT), refused in file mode (INVALID_INPUT), and not supported
for text-box scopes. Requires the pane's "comment_loss_ids" capability.

Returns the eight evidence keys (before/after are ±200-character
excerpts around the first match, not the whole document), plus
runs_before/runs_after (each match's overlapping run(s), clipped to the
span, before and after), `warnings` when non-empty, and (track_changes=
True only) revision_ids/track_changes.

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or an empty find
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
  REVISION_CONFLICT                 - revision_before is stale
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27): the file diverged from this
                                       server's last write within the observation window, e.g.
                                       Office Online co-authoring. Prefer write_mode="live" with
                                       the Live pane open in that editor; or
                                       allow_concurrent_editor=True to write anyway. See
                                       replace_body_markdown.
  ZERO_MATCH                        - find not located after the full ladder
  MATCH_COUNT_MISMATCH              - the located count != expected_matches
  STRUCTURAL_BOUNDARY               - a match crosses a w:p/w:tbl/w:tc boundary
  TRACKED_CHANGES_PRESENT           - a match crosses a FOREIGN-authored tracked change; no force
  OPC_INVALID                       - the rendered .docx failed OPC validation
  VERIFICATION_FAILED               - post-write verification failed; rolled back (file mode);
                                       or the pane's read-back did not confirm the change, with
                                       nothing to roll back (live mode -- Word owns the file)
  LIVE_UNAVAILABLE                  - write_mode="live" (or "auto" routed to live) but no
                                       connected pane session for this document
  LIVE_SESSION_ACTIVE               - write_mode="file" (or "auto" routed to file for some
                                       other reason) but a pane session IS connected for this
                                       document -- use write_mode="live", or save and close the
                                       document in Word (not just the pane), then retry (issue #154)
  LIVE_SESSION_MISMATCH             - a session shares this document's file name but its own
                                       document_url names a different local file; refused rather
                                       than risk mutating the wrong document (issue #154)
  LIVE_DISCONNECTED                 - the pane's socket closed, or an op timed out, mid-call
  LIVE_STALE                        - a "live:sha256:..." revision_before no longer matches the
                                       pane's current body hash, or the pane refused because the
                                       document moved (nothing was written)
  LIVE_STALE_AFTER_WRITE            - the pane applied the edit but its pre-op body hash differed
                                       from the expected one: the write LANDED. Not retryable;
                                       re-read and verify (an audit record is written)
  LIVE_CAPABILITY_MISSING           - within_row_containing given, live mode, but the connected
                                       pane did not report the "row_scope" capability (an old
                                       pane build predating this feature) -- reload the pane

## replace_text_batch

Apply 1..50 replace_text edits to the LIVE Word document in order, in one call (issue #82).

LIVE ONLY: needs a connected Word pane (call `live_open_pane` first if
there is none); `write_mode` must be "live" (or "auto" with a connected
pane), never "file". `revision_before` is required and must be a
"live:sha256:<hex>" revision from a live read.

`edits` is a list of {"find", "replace", "expected_matches",
"allow_comment_loss_ids"?, "within_row_containing"?} objects; semantics per
edit are those of `replace_text` (expected_matches required).
`within_row_containing` is per edit (see replace_text): it scopes that edit
to the table row holding that text, so identical text elsewhere does not
stop the batch. An edit may instead use an anchored find with
"expected_length"/"find_sha256", or "para_ref" + "expected_sha256"
(see replace_text; no within_row_containing then); it is resolved against
the body as the previous edit left it, and a resolution failure is
refused_before_write. `allow_comment_loss_ids` is [{"comment_id", "reply_count"}]
(see replace_text); `allow_comment_loss` inside an edit is refused.
`track_changes` and `inherit_format` apply to every edit.

Edits run strictly in order. One describe + the LIVE_STALE check against
`revision_before` happens first; each edit is then sent guarded by the
body hash the previous edit produced, so the pane refuses (nothing
written) if the document moves mid-batch. The first failure stops the
batch; NOTHING IS EVER RETRIED. Under a section lock the shared broker
drops body-hash checks for the lock holder (see docs/live-mode.md), so
the hash chain may not apply there.

Returns `batch_status`: "complete" (every edit applied), "partial" (>=1
applied, then a known pre-write refusal), "none_applied" (the first edit,
or the start-of-batch staleness check, was refused; the document is
unchanged), or "unknown" (a failure after dispatch whose outcome cannot
be proven: disconnect, timeout, LIVE_OP_FAILED, VERIFICATION_FAILED,
LIVE_STALE_AFTER_WRITE -- re-read the document before doing anything).
`results` has one entry per edit: {index, status, evidence?, code?, message?} with
status "applied" | "refused_before_write" | "unknown" | "not_attempted". The report is
compact (issue #112): an applied edit's `evidence` is just {match_count,
revision_after} (plus comments_removed when a comment was lost); a failed edit has its
`code` and `message`. `verbose=true` returns each edit's before/after text and shape counts.
Only an explicit allowlist of codes counts as refused_before_write
(LIVE_STALE from the pane, MATCH_COUNT_MISMATCH/ZERO_MATCH,
ANCHORED_SHAPES, WOULD_DELETE_COMMENTS, capability and section-lock
refusals). `revision_before`, `revision_after` ("live:sha256:<hex>" of the
last applied edit; the original when none applied; null when unknown),
`applied` (true only for "complete"), `applied_count`, `edit_count`. One
audit record is written per batch, with per-edit statuses.

Errors (nothing was sent): INVALID_INPUT (bad edits, >50 edits, a
non-live write_mode, a malformed revision_before), LIVE_UNAVAILABLE,
LIVE_CAPABILITY_MISSING, LIVE_DISCONNECTED (before the first op).

## format_text

Apply character styling (bold/italic/underline/strike/color) to a
matched text span, without touching its content.

`find` may be anchored ("<prefix>…<suffix>" with expected_length and/or
find_sha256 instead of expected_matches), exactly as on replace_text.

The text never changes, so the result carries ``text_unchanged: true`` and one
``excerpt`` instead of a before and an after, plus ``format_delta`` (only the direct
character formatting that moved; a named style may change none of it, so ``style_id``
is what shows it applied). ``verbose=true`` also returns ``runs_before`` /
``runs_after`` (issue #112).

`write_mode` (issue #106 WP-3, "auto" | "file" | "live", default
"auto"): identical rule and evidence differences to `replace_text`'s
own `write_mode` -- see that tool's docstring. One live-mode
difference worth calling out here: a format-only op never changes
`body.text`, so the live evidence's `revision_before`/
`revision_after` (the pane's body-hash-derived tokens) are typically
EQUAL for a live format call -- that is expected, not a sign the
style failed to apply (unlike `replace_text`, where an unchanged
hash after a live call IS a verification failure). This tool never
goes live at rung 3/4 either; see `replace_text`'s docstring for the
structural-verification-after-`live_save` note, which applies here
identically.

style maps any of "bold"/"italic"/"underline"/"strike" to true/false,
and/or "color" to a 6-hex-digit string (an optional leading "#" is
accepted and stripped; normalized to uppercase) -- issue #22: a
proposal lead marking edits in a font color, with no way to do so
before this. Every requested field's value is applied verbatim
(including false, so {"bold": false} actually clears bold). Setting
an explicit color clears any w:themeColor/themeTint/themeShade on the
run -- Word renders the THEME color over a stale val otherwise, which
would leave this write looking applied while Word displays something
else. Same locate()/expected_matches contract as replace_text (see
that tool's docstring for the normalization ladder,
STRUCTURAL_BOUNDARY refusal, and within_row_containing).

Idempotent: if every located run already carries every requested
field's value, the call skips the write entirely (revision_before ==
revision_after in the returned evidence) rather than creating a new,
no-op revision -- with track_changes=True too: nothing to record in a
w:rPrChange when no style is actually changing.

Same run-splitting rule as replace_text for a boundary run (up to
three pieces; unmatched prefix/suffix keep the ORIGINAL w:rPr cloned
verbatim), except the matched middle SURVIVES here (as a new run
carrying the requested style) rather than being deleted -- this tool
never changes character counts.

track_changes=True (issue #28 WP-07b-a): the changed run's w:rPr gains
a w:rPrChange recording its PRE-change formatting (w:author/w:date/
w:id, same source and allocation as replace_text's track_changes).
Same own-author-vs-foreign-author TRACKED_CHANGES_PRESENT refusal
rule as replace_text.

Returns the eight evidence keys, plus runs_before/runs_after (each
match's overlapping run(s) and their style flags, before and after),
`warnings` when non-empty, and (track_changes=True only) revision_ids/
track_changes.

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors: as replace_text, plus:
  INVALID_INPUT - style is empty, not an object, names an unknown key,
                   a boolean value is not a literal true/false, or
                   "color" is not a 6-hex-digit string

## live_save

Ask the connected Word task pane to save the live document.

Word owns the open document for the whole live session -- nothing a
live `replace_text`/`format_text` call does writes to the .docx file
directly. `live_save` is the one point where a live session's edits
become visible on disk again: it sends the pane's `save` op
(`document.save()`), then reports both the pane's own live revision
token AND a bridge back to the file-mode revision contract, so a
caller that wants to keep working against the file after a live
session has a `file_revision` to pass as the next file-mode call's
`revision_before`.

Structural verification (tables, whole-section rewrites -- rungs 3
and 4, which never go live) against the now-saved file is the
caller's own job afterward, via the existing read tools
(`read_document`/`find_sections`/`diff_body_vs_file`) -- this tool
only confirms the save itself, not the file's contents.

Mutating (added to MUTATING_TOOLS): the middleware requires an
`applied` key, which this tool always returns on success.

Returns `applied`, `saved` (bool), `document_name`, `revision_after`
(`"live:sha256:<hex>"` of the pane's body hash right after the save),
`file_revision` (the plain revision TOKEN STRING --
`projection.compute_revision(path)["token"]` -- the same shape
`replace_text`/`replace_body_markdown`/etc. use as `revision_before`/
`revision_after` in file mode, computed from the now-saved file on
disk; `null` (issue #22 B3) when `path` names no local file at all --
e.g. a SharePoint/OneDrive document with no local sync, where Word's
own save just went there instead), and `audit_logged`.

Errors:
  LIVE_UNAVAILABLE   - no connected pane session for this document
  LIVE_DISCONNECTED  - the pane's socket closed, or the save timed out
  LIVE_OP_FAILED     - the pane replied ok=false to 'save'
  VERIFICATION_FAILED - the pane replied ok=true but did not report saved=true

## list_open_items

List every open comment and pending tracked change (w:ins/w:del) in
a .docx.

fields (issue #111): "full" (default) is the whole record; "summary"
is columnar -- `counts`, `comment_columns` ["comment_id", "author",
"resolved", "anchor" (first 60 chars), "reply_count"] with `comment_rows`
as a list of row lists, and `suggestion_columns` (file: suggestion_id,
kind, author, anchor; live: revision_id, type, author, anchor) with
`suggestion_rows` as row lists (a trailing "scope_uncertain" column
appears only when scoping flagged a row) -- and keeps `path`,
`complete`, `scope`, `warnings` and `source`. Ids are complete. Use it
for guard reads and to pick ids. A summary NEVER has `comments`,
`suggestions` or `pending_suggestions` keys: JennyStack's recorder and
comment loader treat a `comments` list as full objects, so the rows
live under `comment_rows`/`suggestion_rows`. In full, `reply_count`
(== len(replies)) and an empty `replies` are omitted.

section_key / within_text scope the result: a find_sections key (the
heading section), and/or a text span (must match exactly once, like
replace_text's find; both given = their intersection). `counts` then
adds `in_scope` next to the document-wide totals, and the response
carries top-level `scoped: true` (absent when unscoped) so a consumer
can tell a partial `comments` list from a whole-document one. `counts`
keys differ by source: file mode has `comments`/`suggestions`; live
mode has `live_total`/`live_open`/`returned` (plus `in_scope`, and
`open_comments`/`suggestions`, when scoped; an unknown live suggestion
inventory is null, not 0). File mode places a
comment by its anchored runs; a suggestion with no paragraph (text
box, table-row revision) is kept and flagged `scope_uncertain`. Live mode only knows a comment's anchor
PARAGRAPH TEXT, so it matches that text against the live body (read
without re-arming your write baseline); an item whose text cannot be
placed, or that occurs both inside and outside the scope, is kept and
flagged `scope_uncertain: true`; so is a comment anchored across
several paragraphs (live mode sees only one anchor paragraph). An unknown section_key is
INVALID_INPUT listing the available keys.

Returns `comments` (from word/comments.xml + commentsExtended.xml's
resolved flag; each with comment_id, w_id, content, resolved,
replies (omitted when empty), quoted_text, author, created_time -- the
constant modified_time/scope of the GoogleDocs-MCP shape are not
sent) and `pending_suggestions` (every w:ins/w:del, with
suggestion_id, kind ("insertion"|"deletion"), text, author, date,
para_ref; the owning paragraph's own live text is sent ONCE per
paragraph in the top-level `anchor_contexts` {para_ref: text}, not
per suggestion).

Comment replies are threaded (commentsExtended's w15:paraIdParent):
`comments` lists ROOT comments only, each with reply_count ==
len(replies) and replies = [{id, author, date, text}] in comment order;
a reply never appears as its own top-level entry. The root comment's
resolved state governs the thread (a resolved root is omitted together
with its replies; a reply's own resolved flag is ignored).

Each comment's identity is keyed on its LAST paragraph (Word itself
keys commentsIds.xml/commentsExtended.xml this way for a
multi-paragraph comment, not its first -- issue #108); comment_id is
that paragraph's durableId when commentsIds.xml has one, falling back
to the raw w:comment/@w:id otherwise (e.g. no commentsIds.xml at all).
Every comment_id list_open_items emits is accepted by get_comment_thread/
reply_to_comment/resolve_comment, including that raw w:id fallback.

source: "auto" (default) | "file" | "live" -- issue #106 WP-4. "file"
is today's OOXML read (unchanged), and the response gets an added
`source: "file"` key. "live" reads through the connected Word task
pane instead (`comments_list`, live/protocol.py), raising
LIVE_UNAVAILABLE if no pane session for this document's file name is
connected. "auto" picks live whenever such a session exists (a read
never refuses, so this is simpler than write_mode's own "auto" rule),
else file.

In live mode: `source: "live"`; each comment's `comment_id` is a live
handle `live:<Comment.id>` (Office.js's own id, unrelated to the OOXML
durableId/w:id -- docs/live-mode.md's WP-1 finding; it does not
survive a Word restart), and each comment carries
`anchor_text`/`anchor_paragraph_text`/`author`/`created_time`/
`resolved`/`replies` read straight from the pane (`session_epoch` is
sent once at the top level; there is no `w_id`/`quoted_text`, which
was always null / equal to `anchor_text`). Live mode already returns
the FULL reply threads, so get_comment_thread is only needed in file
mode. Resolved comments are filtered out the same way file mode does,
even though the pane reports them too.

Issue #39 -- READ THESE BEFORE TELLING A USER A COMMENT DOES NOT EXIST.
The live response leads with `counts` (`live_total`/`live_open`: the
pane's own count of the whole collection it read; `returned`: how many
open comments are in `comments`), `scope` ("body": only the main
document body is read, so comments anchored in headers, footers or text
boxes are NOT visible), `pane` (`instance_id`/`platform`/`host`/
`observed_at`: WHICH pane answered and when it read), and `warnings`
(`pane_collection_incomplete`, `same_document_duplicate` when a second pane for
this document was connected, `file_has_unlisted_comments` when the saved
file has open comments the pane did not report). `file_only_open_comments`
lists those. All of it is advisory: a pane that returns a stale
collection still reports matching counts, so the tool cannot certify
completeness. When a user can see a comment this does not list, say it
is "not visible to the pane" and quote `counts`/`scope`/`observed_at`;
never say it does not exist. A pane that predates `comment_counts`
gets computed counts plus a `pane_note`. A malformed pane reply (no
`comments` list, or `counts` that disagree with the comments returned)
raises LIVE_OP_FAILED rather than reporting an empty result.

With include_correlation=true the response
also carries a top-level `correlation` list: for each listed live
comment, the best-matching file-mode comment_id (durableId, or a raw
w:id fallback) from this same document's on-disk snapshot, matched by
normalized anchor text + normalized content (+ author, when both
sides have one) -- `confidence` is `"exact"` | `"content-only"` |
`"none"`. Date is NEVER part of the `confidence` match (issue #106
WP-6 real-pane finding: Word for Mac writes `w:date` as local
wall-clock time with a `Z` suffix while Office.js's `creationDate` is
true UTC, so the two routinely differ by several hours for the SAME
comment); each entry instead carries `date_skew_s`, the signed
whole-second gap between the two sides' timestamps (or `null`),
informational only. Each entry also carries `w_id_stable: false` --
`w_id` is only a meaningful correlation key within one open Word
session (a real save has been observed to renumber it); `comment_id`
(the durableId) is the durable key across saves/sessions (issue
#108). This is ADVISORY, never authoritative: a caller already
holding a `live:<id>` handle should use it directly; correlation only
exists so a durableId/w:id (from an earlier file-mode call, or pasted
in by the lead) can still be used with
`reply_to_comment`/`resolve_comment(write_mode="live")`.

Issue #22 B3: live mode no longer requires `path` to name a file that
exists locally (a SharePoint/OneDrive document with no local sync is
the real incident this fixes -- previously the caller had to
fabricate a same-named local stand-in file just to get here). When
there is genuinely no local file, `pending_suggestions` is `null`
(never `[]` -- an empty list would silently claim "nothing pending,"
which isn't knowable without the file) and `correlation` is `[]`; the
added `file_side_available` (bool) names the reason explicitly. A
`live:<id>` comment handle from this same call's own `comments` list
still works fully in `reply_to_comment`/`resolve_comment` either way
-- only durableId/w:id-based resolution needs the file-side
correlation this degrades. `read_document`/`list_parts`/
`find_sections`/etc. remain FILE-ONLY -- this does not add a live
read path for them.

Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
Word's owner file is present, like every other read tool.

Errors:
  INVALID_INPUT    - path does not exist or is outside the allowed
                      roots, source is not "auto"/"file"/"live", fields
                      is not "summary"/"full", or section_key is unknown
  ZERO_MATCH / MATCH_COUNT_MISMATCH - within_text did not match exactly once
  SNAPSHOT_FAILED  - the read-path snapshot could not be validated
  LIVE_UNAVAILABLE - source="live" requested but no connected pane
                      session for this document

## accept_tracked_changes

Accept tracked changes (w:ins/w:del), atomically -- all of them, or
only the ids named in revision_ids (from list_open_items'
pending_suggestions[].suggestion_id).

Accepting a w:ins makes its inserted text ordinary, permanent content
(the wrapper is removed, the text stays). Accepting a w:del makes the
deletion permanent (the w:del element and its w:delText content are
removed outright).

table_id (live only, issue #79): accept only the tracked changes inside
that table (list_tables numbering), leaving every other tracked change in
the document alone. Mutually exclusive with revision_ids; file mode refuses
it (LIVE_UNAVAILABLE); an older pane without the `revisions_scoped`
capability refuses it (LIVE_CAPABILITY_MISSING) rather than touching the
whole document. Like the document-wide live path it is refused while a
section lock is held. The pane verifies from the body OOXML that no
revision remains in the table and that the revisions outside it are
unchanged.

Same guard/atomic-write mechanics as replace_text (lock_status first,
then revision_before; a temp file OPC-validated before the original is
touched; .jsbak-backed post-write verification).

Returns the eight evidence keys (before/after are the WHOLE document's
plain text, not an excerpt -- there is no single match span here;
rung is "all" when revision_ids is omitted, else "by_id"; match_count
is the number of w:ins/w:del elements processed), plus revision_ids
(the ids actually processed).

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard
  REVISION_CONFLICT                 - revision_before is stale
  EXTERNAL_EDITOR_ACTIVE            - see replace_body_markdown (issue #27)
  REVISION_ID_NOT_FOUND             - a named id is not present (available ids listed)
  OPC_INVALID                       - the rendered .docx failed OPC validation
  VERIFICATION_FAILED               - post-write verification failed; rolled back

## add_anchored_comment

Add a comment anchored to a quoted passage, atomically.

write_mode: "auto" (default) | "file" | "live" -- issue #106 WP-4.
"file" is today's OOXML path -- issue #154: BUT it now refuses with
LIVE_SESSION_ACTIVE if a pane session is connected for this document
(see that error code below), rather than writing to disk underneath
Word's own open, autosaving buffer. "live" sends `comment_add` to the
connected Word task pane instead (live/protocol.py), raising
LIVE_UNAVAILABLE if no pane session for this document's file name is
connected. "auto" (issue #154) picks live whenever such a session
exists -- no lock_status/owner-file check any more (see
`replace_text`'s own `write_mode` docstring for why that signal was
dropped). A live comment's id is a `live:<Comment.id>` handle --
Office.js's own Comment.id is unrelated to the OOXML durableId/w:id
(docs/live-mode.md's WP-1 finding) and does not survive a Word
restart. `write_mode="live"` currently creates exactly ONE comment
anchored to the FIRST match even when expected_matches > 1 (the pane's
own comment_add op, built in WP-2, inserts on results.items[0] only) --
unlike file mode's one-comment-per-match behavior; the count is still
verified before anything is inserted. Word's signed-in user is always
the comment's author in live mode; this server cannot spoof a
different one.

Locates `quote` via the same locate()/expected_matches contract as
replace_text (normalization ladder, STRUCTURAL_BOUNDARY refusal,
MATCH_COUNT_MISMATCH on the wrong count) -- see that tool's docstring.
Every matched span gets its own new comment (same `text`, separate
w:id/paraId/durableId each) when expected_matches > 1.

Builds all five comment-related package parts a real Word comment
needs (word/comments.xml, commentsExtended.xml, commentsIds.xml,
commentsExtensible.xml, people.xml -- created fresh on a document's
first-ever comment, otherwise appended to) plus the anchor itself in
word/document.xml (w:commentRangeStart/End and a w:commentReference
run) and the matching [Content_Types].xml override + word/_rels/
document.xml.rels relationship for each newly created part.
w:id = one past the highest existing comment id (0 for a document's
first comment); paraId/durableId = a random 8-hex-uppercase value
below 0x80000000, unique within the package. Author (w:author, and
word/people.xml's w:15:person) comes from author.resolve_author_name()
(~/.jennystack/config.json's author_name, falling back to the macOS
full name).

A run whose text the quote's boundary falls in the middle of splits
the same way replace_text/format_text's boundary runs do (ORIGINAL
w:rPr cloned verbatim onto every surviving piece) -- but nothing here
wraps content the way w:ins/w:del does: commentRangeStart/End are
self-closing position markers, so a boundary run only ever needs
splitting to expose a clean insertion point, never any content change.

After the write, the target file is re-read from disk and the
anchored range is confirmed to bracket the requested quote (modulo
whitespace) before the call returns -- a mismatch raises
VERIFICATION_FAILED via the same .jsbak-backed atomic-write path
every other mutating tool in this server uses.

Returns the eight evidence keys (before/after are unchanged -- a
comment never edits document text), plus comment_ids (every durableId
created, one per matched span) and comment_id (the singular durableId,
only when exactly one span matched).

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty. Not present in live mode (no conflict-
copy fields there); live evidence instead carries write_mode
("live"), verified_via ("word-addin"), document_name, author
("word-signed-in-user"), and orphaned_comment_ids ([]). In live mode,
`rung` is `locate.RUNG_EXACT` ("exact") -- the same value file mode's
own `add_anchored_comment` reports for an ordinary single-pass match
(not the string "live", and not the unrelated numeric edit-ladder
`rung` replace_text/format_text's live evidence reports).

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or an empty quote
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (file mode only)
  REVISION_CONFLICT                 - revision_before is stale (file mode only)
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27): the file diverged from this
                                       server's last write within the observation window, e.g.
                                       Office Online co-authoring; prefer write_mode="live" with
                                       the Live pane open, or allow_concurrent_editor=True (see
                                       replace_body_markdown)
  ZERO_MATCH                        - quote not located (file mode's full ladder; live mode's pane search)
  MATCH_COUNT_MISMATCH              - the located count != expected_matches
  STRUCTURAL_BOUNDARY               - a match crosses a w:p/w:tbl/w:tc boundary (file mode only)
  OPC_INVALID                       - the rendered .docx failed OPC validation (file mode only)
  VERIFICATION_FAILED               - post-write verification failed; rolled back (file mode only)
  LIVE_UNAVAILABLE                  - write_mode="live" requested but no connected pane session
  LIVE_SESSION_ACTIVE               - write_mode="file" (or "auto" routed to file for some other
                                       reason) but a pane session IS connected for this document
                                       (issue #154)
  LIVE_SESSION_MISMATCH             - a session shares this document's file name but its own
                                       document_url names a different local file (issue #154)
  LIVE_DISCONNECTED                 - the pane disconnected mid-op or an op reply timed out
  LIVE_STALE                        - a caller-supplied expected body hash disagreed with the pane's own pre-op hash

## add_anchored_comment_batch

Add 1-50 anchored comments in ONE call (issue #113), instead of one
add_anchored_comment call per finding.

comments: [{quote, content, expected_matches (default 1), section_key
(optional)}], validated up front (non-empty quote/content, expected_matches
an int >= 1). `quote` is located as in add_anchored_comment; `section_key`
(from find_sections) restricts the item to that heading section.

Same contract as replace_text_batch: the first failure stops the batch and
NOTHING IS EVER RETRIED. One audit record per call holds every item's evidence.

write_mode "auto" | "file" | "live" as in add_anchored_comment.
LIVE: one pane comment_add per item, chained on the body hash; earlier comments
stay when a later item fails. The pane comments only the FIRST match, so every
item needs expected_matches == 1 (INVALID_INPUT otherwise, nothing sent). With a
section_key the quote is checked against a non-observing read of the live body
(it never re-arms a stale shared-bridge client): it must occur inside that
section and nowhere else (ZERO_MATCH if it is not in the section, AMBIGUOUS_SPAN
if it also occurs outside). revision_before ("live:sha256:<hex>") is checked once
at the start (LIVE_STALE). Item status: "applied" | "refused_before_write" |
"unknown" (a failure after dispatch whose outcome cannot be proven: re-read the
document before doing anything) | "not_attempted".
FILE: every comment is staged on one in-memory document and saved ONCE, guarded
(revision_before, allow_concurrent_editor as in add_anchored_comment) and
verified. expected_matches comments every match; section_key scopes the search
to that section, so the quote may repeat elsewhere (SECTION_NOT_FOUND if the key
is unknown). All-or-nothing: if an item fails, nothing is saved: earlier items
report "not_persisted", the failing one its code, the rest "not_attempted",
batch_status "none_applied"; a failed save makes every item "not_persisted" (or
"unknown", batch_status "unknown", when the save may have landed).

Returns {applied (true only for "complete"), batch_status ("complete" | "partial"
| "none_applied" | "unknown"), applied_count, item_count, revision_before,
revision_after, results: [{index, status, comment_id? (comment_ids for a
multi-match file item), code?, message?}], evidence_ref}. verbose=true adds each
applied item's evidence (match_count, revision_after / before excerpt).

Errors (nothing was written): INVALID_INPUT (bad items, >50, live
expected_matches != 1), DOCX_LOCKED / SYNC_IN_FLIGHT / REVISION_CONFLICT /
EXTERNAL_EDITOR_ACTIVE / LIVE_SESSION_ACTIVE (file mode guard), LIVE_UNAVAILABLE,
LIVE_CAPABILITY_MISSING, LIVE_DISCONNECTED (before the first op).

## reply_to_comment

Reply to an existing comment (durableId), atomically -- issue #28
WP-09.

write_mode: "auto" (default) | "file" | "live" -- issue #106 WP-4.
"file" is today's OOXML path -- issue #154: BUT it now refuses with
LIVE_SESSION_ACTIVE if a pane session is connected for this document,
rather than writing to disk underneath Word's own open, autosaving
buffer. "live" sends `comment_reply` to the connected Word task pane,
accepting a `live:<Comment.id>` handle directly OR a durableId/w:id
resolved through `list_open_items(source="live")`'s own `correlation`
list (exact or content-only confidence; INVALID_INPUT, naming both id
spaces and suggesting that call, if neither resolves). "auto" (issue
#154) picks live whenever a pane session exists for this document --
no lock_status/owner-file check any more (see `replace_text`'s own
`write_mode` docstring for why). Verified by re-listing the
pane's comments and confirming a reply with the given text now
exists. Word's signed-in user is always the reply's author in live
mode. A live comment id does not survive a Word restart; correlation
is advisory only, never authoritative.

Creates a NEW, independent comment (its own w:comment / w:id / paraId
/ durableId, and its own full commentRangeStart/End/commentReference
anchor triplet in word/document.xml) whose commentsExtended.xml entry
carries w15:paraIdParent pointing at the PARENT's own paraId -- this
is exactly how a real Word-authored reply is shaped (verified against
tests/fixtures/comments/golden-comment.docx's own real reply). The
reply's anchor brackets the SAME live text the parent's own anchor
currently does; no new text is located or matched.

The PARENT'S own identity paraId is its LAST paragraph's, matching
how Word itself keys commentsIds.xml/commentsExtended.xml for a
multi-paragraph comment (not its first -- issue #108) -- so a reply
to a multi-paragraph comment links w15:paraIdParent to the correct
paragraph. *comment_id* resolves via commentsIds.xml's durableId
first, then falls back to a raw w:comment/@w:id match (same fallback
as get_comment_thread/resolve_comment).

Returns the eight evidence keys (before/after are the parent's own
quoted text, unchanged -- a reply never edits document text; rung is
the fixed label "reply", since no text search is performed), plus
comment_id (the new reply's own durableId), parent_comment_id, and
comment_id_resolved_via ("durableId" | "w_id", saying which lookup
path matched *comment_id*, the parent).

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty. Not present in live mode; live
evidence instead carries write_mode ("live"), verified_via
("word-addin"), document_name, and author ("word-signed-in-user"). In
live mode, revision_before/revision_after are `"live:sha256:<hex>"`
of a `describe` call's body hash taken immediately before/after the
op (same token shape replace_text/format_text/add_anchored_comment's
live evidence uses) -- typically equal, since a reply never edits
body text, mirroring format_text's own before==after case.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or
                     comment_id does not match any existing comment,
                     or that comment has no live anchor to reply against,
                     or (live mode) comment_id matches neither a
                     live:<id> handle nor a correlated durableId/w:id
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (file mode only)
  REVISION_CONFLICT                 - revision_before is stale (file mode only)
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27): the file diverged from this
                                       server's last write within the observation window, e.g.
                                       Office Online co-authoring; prefer write_mode="live" with
                                       the Live pane open, or allow_concurrent_editor=True (see
                                       replace_body_markdown)
  OPC_INVALID                       - the rendered .docx failed OPC validation (file mode only)
  VERIFICATION_FAILED               - post-write verification failed; rolled back (file mode only)
  LIVE_UNAVAILABLE                  - write_mode="live" requested but no connected pane session
  LIVE_SESSION_ACTIVE               - write_mode="file" (or "auto" routed to file for some other
                                       reason) but a pane session IS connected for this document
                                       (issue #154)
  LIVE_SESSION_MISMATCH             - a session shares this document's file name but its own
                                       document_url names a different local file (issue #154)
  LIVE_DISCONNECTED                 - the pane disconnected mid-op or an op reply timed out
  LIVE_OP_FAILED                    - the reply was not found after re-listing the pane's comments

## resolve_comment

Resolve a comment thread (durableId), atomically -- issue #28
WP-09. Sets w15:done="1" on the comment's own commentsExtended.xml
entry; list_open_items excludes it afterward (it is no longer an
"open" item), but get_comment_thread still fetches it by id --
resolved is not deleted, only marked.

write_mode: "auto" (default) | "file" | "live" -- issue #106 WP-4.
"file" is today's OOXML path -- issue #154: BUT it now refuses with
LIVE_SESSION_ACTIVE if a pane session is connected for this document,
rather than writing to disk underneath Word's own open, autosaving
buffer. "live" sends `comment_resolve` (resolved=true) to the
connected Word task pane, accepting a `live:<Comment.id>` handle
directly OR a durableId/w:id resolved through
`list_open_items(source="live")`'s own `correlation` list (exact or
content-only confidence; INVALID_INPUT, naming both id spaces and
suggesting that call, if neither resolves). "auto" (issue #154) picks
live whenever a pane session exists for this document -- no
lock_status/owner-file check any more (see `replace_text`'s own
`write_mode` docstring for why). Verified the
same way file mode is: re-listed independently after the op, not
trusted on the pane's own ok=true -- COMMENT_STILL_OPEN if the
re-list does not show it resolved. Word's signed-in user is always
the resolver's identity in live mode; a live comment id does not
survive a Word restart, and correlation is advisory only.

COMMENT_STILL_OPEN is adapted, not lifted verbatim, from
GoogleDocs-MCP's own member of the same name: that server's version
guards genuine Drive-API eventual consistency (a resolve action that
does not durably stick server-side) by re-querying the comment from
the API after the write. This backend's write is a local, synchronous,
atomically-verified file replace with no such external-consistency
hazard -- so this code can only ever fire here via a bug in this
server's own code, not a real runtime race. Kept anyway (re-reading
the fresh file from disk and checking w15:done="1" independently of
the write's own success) so the error VOCABULARY still matches
Google's for this exact failure mode.

Idempotent: resolving an already-resolved comment succeeds again
(not an error).

A multi-paragraph comment's identity is its LAST paragraph's paraId,
matching how Word itself keys commentsIds.xml/commentsExtended.xml
for one (not its first -- issue #108). *comment_id* resolves via
commentsIds.xml's durableId first, then falls back to a raw
w:comment/@w:id match (same fallback as get_comment_thread/
reply_to_comment) -- this is what makes every comment_id
list_open_items emits resolvable, including three real
multi-paragraph comments that were listed but unresolvable before
this fix.

Returns the eight evidence keys (before="open", after="resolved";
rung is the fixed label "resolve", since no text search is
performed), plus comment_id and comment_id_resolved_via
("durableId" | "w_id", saying which lookup path matched).

Always also carries conflict_copy_detected (issue #28 WP-10's
post-write conflict-copy sweep -- never raised, an evidence flag on
an already-successful write); conflict_copies/sibling_files_changed
are added only when non-empty. Not present in live mode; live
evidence instead carries write_mode ("live"), verified_via
("word-addin"), document_name, and author ("word-signed-in-user"). In
live mode, revision_before/revision_after are `"live:sha256:<hex>"`
of a `describe` call's body hash taken immediately before/after the
op (same token shape replace_text/format_text/add_anchored_comment's
live evidence uses) -- typically equal, since a resolve never edits
body text.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, or
                     comment_id does not match any existing comment, or
                     (live mode) comment_id matches neither a live:<id>
                     handle nor a correlated durableId/w:id
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (file mode only)
  REVISION_CONFLICT                 - revision_before is stale (file mode only)
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27): the file diverged from this
                                       server's last write within the observation window, e.g.
                                       Office Online co-authoring; prefer write_mode="live" with
                                       the Live pane open, or allow_concurrent_editor=True (see
                                       replace_body_markdown)
  COMMENT_STILL_OPEN                - the post-write re-read did not confirm the resolve (see this tool's own docstring)
  OPC_INVALID                       - the rendered .docx failed OPC validation (file mode only)
  VERIFICATION_FAILED               - post-write verification failed; rolled back (file mode only)
  LIVE_UNAVAILABLE                  - write_mode="live" requested but no connected pane session
  LIVE_SESSION_ACTIVE               - write_mode="file" (or "auto" routed to file for some other
                                       reason) but a pane session IS connected for this document
                                       (issue #154)
  LIVE_SESSION_MISMATCH             - a session shares this document's file name but its own
                                       document_url names a different local file (issue #154)
  LIVE_DISCONNECTED                 - the pane disconnected mid-op or an op reply timed out

## delete_comment

Delete a comment thread (the root comment and every reply), atomically
-- issue #80. resolve_comment only marks a thread done and Word keeps
resolved threads visible; this removes the thread.

Thread roots only: a reply id is refused with INVALID_INPUT naming its
root. *comment_id* takes a durableId, a raw w:id, or (live mode) a
`live:<id>` handle, with the same fresh-ID/epoch/identity rules as
resolve_comment (session_epoch and match_spec are live-mode addressing).
The comment need not be resolved.

write_mode "auto" | "file" | "live" as resolve_comment. File mode
removes the w:comment, its commentsExtended/commentsIds/
commentsExtensible entries, and the range markers/reference run in
word/document.xml; people.xml, relationships and content types are not
touched. The written package is re-read and rolled back unless the
thread is gone and every other comment and anchor is unchanged. Live
mode sends `comment_delete` to the Word pane and verifies with an
independent re-list; it cannot roll back.

Returns the eight evidence keys (rung "delete", before = the thread's
anchor text, after "deleted"), plus comment_id,
comment_id_resolved_via, resolutions, deleted_thread_count,
deleted_reply_count, deleted_comment_ids, comments_before,
comments_after and other_comments_unchanged. Live evidence also
carries write_mode, verified_via, document_name, author, scope
("body": comments outside the main body are not covered) and
session_epoch.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, an empty list, or an id that is
                     unknown, a reply (only thread roots can be deleted; the error names the
                     root) or, for delete_resolved_comments, still open. diagnostics.offenders
                     lists every one and nothing is deleted.
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (file mode only)
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27); see replace_body_markdown
  OPC_INVALID, VERIFICATION_FAILED  - file mode: the written package failed validation, a removed
                     id survived, or a surviving comment/anchor differed from the original;
                     rolled back. Live mode: another comment changed during the delete (NOT
                     rolled back; see diagnostics).
  COMMENT_STILL_PRESENT             - live mode: the pane's independent re-list still shows a
                     deleted thread
  LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED - as resolve_comment
  LIVE_CAPABILITY_MISSING           - the pane predates comment_delete: close and reopen the Live pane
  COMMENT_ID_STALE                  - live mode: the handle is from an earlier pane epoch, or the
                     comment's identity changed
  LIVE_OP_FAILED                    - live mode: the pane's re-list could not be read or trusted, or
                     the pane refused; diagnostics.deleted/remaining/unknown say what is known.
                     A delete that errored part-way is never rolled back.

## delete_resolved_comments

Delete several RESOLVED comment threads in one call -- issue #80. Use
it after resolve_comment to clear the threads a push just resolved.

Every id is resolved, deduplicated by thread root (a durableId and a
w:id naming the same comment count once) and validated BEFORE anything
is deleted: an empty list, an unknown id, a reply id or a still-open
thread refuses the whole call (INVALID_INPUT, diagnostics.offenders
lists every offender). File mode deletes all threads in one atomic
write. Live mode resolves every id against one pane snapshot, deletes
them in turn and verifies with one re-list; if a delete fails part-way,
the error's diagnostics list `deleted`, `remaining` and `unknown`
(deleted before the error but unconfirmed) live handles and nothing is
rolled back.

Evidence is delete_comment's, with match_count and
deleted_thread_count = threads deleted, `before` = the list of anchor
texts, and `resolutions` = one entry per input id.

Errors:
  INVALID_INPUT, DOCX_PATH_ESCAPE, DOCX_ROOT_NOT_FOUND - a bad path, an empty list, or an id that is
                     unknown, a reply (only thread roots can be deleted; the error names the
                     root) or, for delete_resolved_comments, still open. diagnostics.offenders
                     lists every one and nothing is deleted.
  DOCX_LOCKED, SYNC_IN_FLIGHT       - the write guard (file mode only)
  EXTERNAL_EDITOR_ACTIVE            - file mode only (issue #27); see replace_body_markdown
  OPC_INVALID, VERIFICATION_FAILED  - file mode: the written package failed validation, a removed
                     id survived, or a surviving comment/anchor differed from the original;
                     rolled back. Live mode: another comment changed during the delete (NOT
                     rolled back; see diagnostics).
  COMMENT_STILL_PRESENT             - live mode: the pane's independent re-list still shows a
                     deleted thread
  LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE, LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED - as resolve_comment
  LIVE_CAPABILITY_MISSING           - the pane predates comment_delete: close and reopen the Live pane
  COMMENT_ID_STALE                  - live mode: the handle is from an earlier pane epoch, or the
                     comment's identity changed
  LIVE_OP_FAILED                    - live mode: the pane's re-list could not be read or trusted, or
                     the pane refused; diagnostics.deleted/remaining/unknown say what is known.
                     A delete that errored part-way is never rolled back.

## get_table

One table's content, addressed by the table_id list_tables reports.

compact=true (default): rows is a list of rows of cell-text strings and
merged lists only merged cells ({row_index, cell_index, grid_span,
v_merge}). compact=false returns the detailed shape: rows of
{row_index, cell_index, grid_span, v_merge, text} for every cell, and no
merged key. Live compact reads of a table with merged cells cannot say
which cells: merged is null and warnings has "live_merge_metadata_unavailable".

`source` (issue #34, "auto" | "file" | "live", default "auto"): "auto"
reads the open document through the Live pane when one is connected for
this document, else the file. A live read is the only way to see a table
a co-author just added (the saved file lags). LIVE READS DIFFER: the main
document body only (a non-default `part` is INVALID_INPUT with
source="live", and "auto" falls back to the file); `text` is Word's plain
text, not markdown; per-cell grid_span/v_merge are not available (every
cell reports grid_span 1 / v_merge "none"), so use the table-level
has_merged_cells flag; a document containing a nested table is refused;
the pane must report "table_edit" (LIVE_CAPABILITY_MISSING otherwise).
The result adds source ("live"), style and header_row_count. Every
response carries "source"; a file read while a live pane session exists
for this file name carries the "live_session_ignored" warning plus a
live_session object (issue #33), so a same-named placeholder file can
not pass as the live table. This live read goes through the pane's
table_get op (the same table numbering live table edits use), unlike
list_tables(source="live"), which reads the body OOXML: the two agree on
table_id when the document has no nested table.

Returns path, part, table_id, row_count, col_count, has_merged_cells,
has_nested_table, rows (detailed mode: list of list of {row_index,
cell_index, grid_span, v_merge, text}). grid_span is w:tcPr/w:gridSpan (1 when
absent); v_merge is "none"/"restart"/"continue" from w:tcPr/w:vMerge
(a continuation cell's own text is always ""). text is the same
per-cell markdown rendering read_document(format="markdown") uses for
a pipe-table cell (a nested w:tbl inside a cell flattens to inline
text, same as there).

Not gated by DOCX_LOCKED -- reads a validated snapshot instead when
Word's owner file is present.

Errors:
  INVALID_INPUT   - path does not exist or is outside the allowed roots
  PART_NOT_FOUND  - part names a package part absent from this .docx
  TABLE_NOT_FOUND - table_id does not match any table (call list_tables)
  SNAPSHOT_FAILED - the read-path snapshot could not be validated
  LIVE_UNAVAILABLE / LIVE_CAPABILITY_MISSING / LIVE_DISCONNECTED / LIVE_OP_FAILED - source="live"

## replace_table_row

Replace one table row's cell content wholesale, one markdown string
per cell (cells must have exactly as many entries as the row has
cells). row_index is 1-based in both modes; row 1 is the table's first
row (usually the header).

`write_mode` (issue #34, "auto" | "file" | "live", default "auto"):
"auto" routes through the Live pane when one is connected, else writes
the file -- the supported way to edit a table in a document a co-author
has open. LIVE MODE IS A SUBSET: cells are paragraphs with bold/italic/
links only; a table with a merged cell is refused
(MERGED_OR_NESTED_TABLE, as in file mode); `force` does not apply;
before/after are plain text, not markdown; revision_before/
revision_after are "live:sha256:<hex>" body hashes; the pane must report
"table_edit". The row is written as ONE Word batch, guarded by a
body-hash check and a re-read of every cell just before the write, then
verified by re-reading the table -- but Word batches are not
transactions, so under heavy co-authoring a cell edited in the last
instant can still be overwritten. A write that lands and then fails its
read-back is recorded in the audit log before VERIFICATION_FAILED is
raised.

Refuses -- for the WHOLE table, not just this row -- the moment
table_id names a table containing a merged cell (w:gridSpan != 1 or a
w:vMerge) or a nested w:tbl anywhere in it, mirroring
GoogleDocs-MCP's own replace_table_row refusal on a merged cell.
replace_cell_markdown is the escape hatch for that case: a cell-scoped
write that never touches w:tcPr.

Same guard (lock/sync/revision, before any temp file), comment-anchor/
tracked-change hazard scan (force=True to proceed; COMMENT_ANCHORS_IN_RANGE
/ TRACKED_CHANGES_PRESENT otherwise), atomic write with OPC validation
and a .jsbak rollback, and audit log as every other mutating tool.

track_changes=True wraps the row's OLD cell content in w:del and the
NEW content in w:ins (per cell), under the configured author, rather
than rewriting directly.

Returns the eight evidence keys (before/after are the row's cells
tab-joined), plus (track_changes=True only) revision_ids/track_changes,
and conflict_copy_detected (+ conflict_copies/sibling_files_changed
when non-empty).

Errors:
  INVALID_INPUT          - a bad path, or cells' length != the row's own cell count
  DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
  REVISION_CONFLICT      - revision_before is stale
  TABLE_NOT_FOUND         - table_id does not match any table
  TABLE_ROW_NOT_FOUND     - row_index out of range for this table
  MERGED_OR_NESTED_TABLE  - this table has a merged/nested-table cell; use replace_cell_markdown
  COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT - a hazard in the row; force=True to proceed
  OPC_INVALID             - the rendered .docx failed OPC validation
  VERIFICATION_FAILED     - post-write verification failed; rolled back (file mode), or the
                            pane's read-back did not confirm the edit (live mode; audit-logged)
  LIVE_UNAVAILABLE / LIVE_SESSION_ACTIVE / LIVE_SESSION_MISMATCH / LIVE_DISCONNECTED /
  LIVE_STALE / LIVE_CAPABILITY_MISSING / LIVE_OP_FAILED - as replace_text (issue #154);
                            LIVE_OP_FAILED here also covers a cell edited between the read and
                            the write, and a document with nested tables

## replace_cell_markdown

Replace one table cell's content with rendered markdown, leaving
its own w:tcPr byte-identical -- the only write path safe on a merged
(w:gridSpan/w:vMerge) cell, and the intended path for an Appendix-A
style band-and-border table's own cell content.

markdown_file (issue #113): a path to a .md/.txt file (UTF-8, <= 2 MB, under the
allowed roots; outside them -> DOCX_PATH_ESCAPE) used INSTEAD of `markdown`, so text
already on disk need not be re-emitted. Exactly one of the two (else INVALID_INPUT).
The text is read once, server-side, before any routing; the response adds
markdown_sha256 (hex of its UTF-8 bytes) and markdown_chars.

`write_mode` (issue #27, "auto" | "file" | "live", default "auto"):
"auto" routes through the Live pane when one is connected for this
document, else writes the file. This is the supported way to edit a
cell in a document a co-author has open in Office Online: a live edit
is a co-authoring edit, where a file write races that editor's
autosave and is likely to be silently reverted. LIVE MODE IS A SUBSET:
one or more paragraphs and hard line breaks (two trailing spaces then
a newline), with bold/italic/links (lists, headings, tables and other
block structure -> INVALID_INPUT, nothing sent); more than one
paragraph or a line break also needs the pane's "cell_multiline"
capability (LIVE_CAPABILITY_MISSING before anything is sent -- reload
the pane); the cell is
edited compare-and-set against the text just read, so a co-author's
concurrent edit to the same cell is refused (LIVE_OP_FAILED) rather
than overwritten; the pane re-reads the cell's paragraphs after
writing and restores the previous content if they do not match
(VERIFICATION_FAILED, diagnostics.rolled_back=true; false means the
cell needs a manual look); a document containing a NESTED table is
refused (the pane cannot map table_id onto nested tables); `force`
does not apply; before/after are the cell's plain text, not markdown;
revision_before/revision_after are "live:sha256:<hex>" body hashes;
and the connected pane must report the "cell_edit" capability
(LIVE_CAPABILITY_MISSING otherwise -- reload the pane). Requires the
pane on the same machine as this server (docs/live-mode.md).

row_index and cell_index are 1-based in BOTH modes: row 1 is the
table's first row (usually the header), cell 1 its first cell. Pass
`expected_before` (the cell's current text, plain or markdown; compared
modulo markdown marks and whitespace) to guard against a wrong index or
a stale read: a mismatch is CELL_TEXT_MISMATCH and nothing is written.

`inherit_format` (issue #107, "first_run" | "none", default "first_run"):
markdown carries no font, so by default every rendered run takes the font
name, size and colour of the cell's first text run (else its first paragraph
mark), where the markdown does not set them. FILE mode copies that run's
DIRECT properties (so a cell whose font comes only from its style inherits
nothing), also the complex-script size and the paragraph marks, and verifies
every run on re-read. LIVE mode applies the font Word reports at the start of
the cell (name, size and colour; Word's API cannot set the complex-script
size) and reads it back (VERIFICATION_FAILED otherwise). A link keeps its own
colour. Without this a multi-paragraph write into an 8.5pt cell came back at
the style's size (reproduced on Word 16.114.3). "none" keeps the old
behaviour (the style's font). A live write with "first_run" needs the
pane's "cell_font_inheritance_v1" capability (LIVE_CAPABILITY_MISSING:
reload the pane, or pass "none" to accept the style's font).

Supports multi-level bulleted/numbered markdown inside the cell
(issue #28 WP-16a): rendered through the SAME markdown_to_ooxml
machinery replace_body_markdown/append_markdown use, one numId/
abstractNum tree per top-level list shared across nesting levels via
increasing w:ilvl.

Same guard/hazard-scan/atomic-write/audit machinery as
replace_table_row. track_changes=True wraps the OLD cell content in
w:del and the NEW content in w:ins, under the configured author.

Returns the eight evidence keys (before/after are the cell's own
markdown), plus (track_changes=True only) revision_ids/track_changes,
and conflict_copy_detected (+ conflict_copies/sibling_files_changed
when non-empty).

Errors:
  INVALID_INPUT           - a bad path
  DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
  REVISION_CONFLICT       - revision_before is stale
  EXTERNAL_EDITOR_ACTIVE  - issue #27: see replace_body_markdown
  TABLE_NOT_FOUND          - table_id does not match any table
  TABLE_ROW_NOT_FOUND      - row_index out of range for this table
  TABLE_CELL_NOT_FOUND     - cell_index out of range for this row
  CELL_TEXT_MISMATCH       - expected_before does not match the cell's current text; nothing written
  COMMENT_ANCHORS_IN_RANGE / TRACKED_CHANGES_PRESENT - a hazard in the cell; force=True to proceed
  OPC_INVALID              - the rendered .docx failed OPC validation
  VERIFICATION_FAILED      - post-write verification failed (including w:tcPr drift); rolled
                             back (file mode), or the pane's read-back did not confirm the
                             edit (live mode; diagnostics.rolled_back says whether the pane
                             restored the cell)
  LIVE_UNAVAILABLE / LIVE_SESSION_ACTIVE / LIVE_SESSION_MISMATCH / LIVE_DISCONNECTED /
  LIVE_STALE / LIVE_CAPABILITY_MISSING / LIVE_OP_FAILED - as replace_text (issue #154);
                             LIVE_OP_FAILED here also covers a cell whose text changed between
                             the read and the write, and a document with nested tables

## insert_table

Insert a new table, one markdown string OR cell-spec object per
cell (issue #100: https://github.com/michaelrobertsutton/JennyStack/issues/100).

markdown_file (issue #113) is an alternative to `rows`: a .md/.txt file (UTF-8,
<= 2 MB, under the allowed roots; outside -> DOCX_PATH_ESCAPE) holding a GFM pipe
table, parsed into the same `rows` (header row first, the |---| separator row
dropped; pass header_rows=1 to repeat the header). Exactly one of rows/markdown_file
(else INVALID_INPUT). Assumption: the issue calls this `rows_markdown`, which does
not exist, so the file is parsed to `rows` here. The response adds markdown_sha256
and markdown_chars.

`write_mode` (issue #34, "auto" | "file" | "live", default "auto"):
"auto" routes through the Live pane when one is connected for this
document, else writes the file. This is the supported way to add a table
to a document a co-author has open: a file write races that editor's
autosave. LIVE MODE IS A SUBSET:
  - cells are paragraphs with bold/italic/links only (lists, headings,
    nested tables -> INVALID_INPUT, nothing sent); the cell-spec keys
    fill/color/bold/align/valign work, but span and v_merge (merged cells)
    and cant_split are INVALID_INPUT;
  - the anchor is {"paragraph_text": <exact text of one top-level
    paragraph>, "position": "after"|"before" (default "after")} or
    {"after_table_id": <int>}, or omitted to append at the end of the
    body. section_key / after_paragraph_text anchors are file-mode only
    (they come from the saved file, which lags a co-authored document).
    0 matching paragraphs -> ZERO_MATCH, more than 1 ->
    MATCH_COUNT_MISMATCH;
  - the style is chosen BEFORE anything is inserted (STYLE_NOT_FOUND
    leaves the document untouched): style_id is treated as a Word table
    style NAME (e.g. "Grid Table 4 - Accent 1"; it must already be used
    by a table in the document, or named in Word's style list on WordApi
    1.5+), style_from_table_id copies an existing table's style, and
    style_builtin takes a Word built-in style name (e.g.
    "GridTable4_Accent1"; live only). All three are optional in live
    mode; omit them for Word's default table look;
  - the new table's table_id is read back from Word (list_tables numbers
    the saved file, which lags);
  - grid_dxa becomes column widths in points; font_size_pt sets the table
    font size;
  - `force` does not apply; revision_before/revision_after are
    "live:sha256:<hex>" body hashes; before is "", after is the table's
    plain text; the pane must report "table_edit"
    (LIVE_CAPABILITY_MISSING otherwise -- reload the pane);
  - the pane refuses (LIVE_STALE) if the document changed since the
    body hash was read, BEFORE inserting anything.
Every requested property (text, fill, color, bold, alignment, widths,
font size, header rows, style) is re-read from Word and compared; a table
that lands but fails that check is recorded in the audit log and raised as
VERIFICATION_FAILED (there is nothing to roll back in live mode -- the
table stays in the document for you to fix or delete).

rows is a list of rows, each a list of cells. A plain string cell
means {"markdown": that string} -- today's exact behavior: short rows
are padded with empty cells up to the widest row, columns split
evenly across the text-column width list_page_sections reports.  A
cell may instead be an OBJECT: {"markdown" (required), "span" (int
>=1, default 1, -> w:tcPr/w:gridSpan), "v_merge" ("restart"|
"continue", -> w:tcPr/w:vMerge -- a "continue" cell's markdown must be
"", and the same grid column in the row above must itself be
"restart"/"continue"), "fill" (6-hex color, -> w:tcPr/w:shd),
"color" (6-hex color, -> w:rPr/w:color on every run in the cell),
"bold" (bool, -> w:rPr/w:b+w:bCs on every run), "align" ("left"|
"center"|"right"|"both", -> w:pPr/w:jc on every paragraph), "valign"
("top"|"center"|"bottom", -> w:tcPr/w:vAlign)}. The MOMENT any cell in
the table uses the object form, grid validation is strict: every
row's cell spans must sum to the same column count (grid_dxa's length
when given, else the widest span-sum) -- no padding, a mismatch is
INVALID_INPUT naming the row and both numbers.

Formatting caveat: fill/valign/span/v_merge live in w:tcPr, which
replace_cell_markdown preserves byte-identical, so they survive a
later cell edit. bold/color/align live on the cell's own runs/
paragraphs and are REPLACED by the next replace_cell_markdown on that
cell -- put content emphasis in the markdown itself (**bold**) and
reserve bold/color/align for header/title rows you will not re-edit.

File mode: style_id (or style_from_table_id, which copies another
table's w:tblStyle -- STYLE_NOT_FOUND if that table has none) is
REQUIRED, and must name an existing w:type="table" style in this
document's styles.xml (list_styles reports style type) -- unlike
GoogleDocs-MCP's insert_table, which relies on the Docs API's own
default table style, raw OOXML has no sensible default to fall back
to. style_builtin is live only (INVALID_INPUT in file mode).
font_size_pt (either mode) sets every run's size: w:sz/w:szCs in
half-points, the point size doubled and rounded half up; like
bold/color/align it lives on the runs and is replaced by the next
replace_cell_markdown on that cell.

header_rows (default 0) sets w:tblHeader (repeating header row(s)) on
the first N rows; must be < len(rows). grid_dxa (default None) gives
explicit per-column widths in dxa -- when given, w:tblW becomes an
explicit dxa sum and each cell's own w:tcW is the sum of the grid
columns it spans; when omitted, columns split evenly across the
text-column width as before (w:tblW stays "auto"). cant_split
(default False) sets w:cantSplit (no page-break-inside-row) on every
row.

anchor (default None keeps today's append-at-the-end behavior) places
the table somewhere else in the body instead:
  {"section_key": "<from find_sections>", "position": "start"|"end"}
    - "start": right after that section's own heading paragraph.
    - "end": at the end of that section (before the next heading, or
      the end of the body).
  {"section_key": "...", "after_paragraph_text": "<exact text of one
   top-level paragraph inside that section>"} - right after that
   paragraph (0 matches -> ZERO_MATCH, >1 -> MATCH_COUNT_MISMATCH; a
   paragraph inside a table is never a candidate). Matched through the
   same exact/quotes/whitespace/soft-hyphen normalization ladder
   locate.py's own text search uses.
  {"after_table_id": <int from list_tables>} - right after that
   top-level body table (a nested table -> INVALID_INPUT; use
   after_paragraph_text on its host cell's section instead).

The new table's own table_id is computed by re-walking the document
tree after insertion (its position may no longer be last), NOT by
counting existing tables.

Same guard/atomic-write/audit machinery as append_markdown.
track_changes=True wraps the new table's own runs in w:ins (nothing
existing is removed by an insertion, so there is no w:del side).

Returns the eight evidence keys (before="", after is the new table's
rows tab/newline-joined), plus table_id (the new table's own id, for a
follow-up get_table/replace_table_row/replace_cell_markdown call),
merged_cells (count of cells with span>1 or a v_merge), and
anchor_resolved ({"body_index", "section_key", "after_table_id"}, or
null when appended), (track_changes=True only) revision_ids/
track_changes, and conflict_copy_detected (+ conflict_copies/
sibling_files_changed when non-empty).

Post-write verification re-reads the table and checks row count, grid
column count, every cell's grid_span/v_merge, every cell's text
modulo whitespace, w:shd fill per requested cell, and w:tblHeader on
the header rows -- any mismatch is VERIFICATION_FAILED (rolled back).

Errors:
  INVALID_INPUT       - a bad path; rows is empty or contains an empty
                         row; a bad cell-spec field (span/v_merge/fill/
                         color/bold/align/valign); a grid-sum mismatch
                         naming the row and both numbers; a bad
                         header_rows/grid_dxa/anchor
  DOCX_LOCKED, SYNC_IN_FLIGHT - the write guard
  REVISION_CONFLICT   - revision_before is stale
  EXTERNAL_EDITOR_ACTIVE - issue #27: see replace_body_markdown
  STYLE_NOT_FOUND     - style_id does not name a table style in this document
  SECTION_NOT_FOUND   - anchor.section_key does not match find_sections' output
  TABLE_NOT_FOUND     - anchor.after_table_id does not match any table
  ZERO_MATCH / MATCH_COUNT_MISMATCH - anchor.after_paragraph_text matched
                         0 or >1 top-level paragraphs in the section
  OPC_INVALID         - the rendered .docx failed OPC validation
  VERIFICATION_FAILED - post-write verification failed; rolled back (file mode), or the
                         pane's read-back did not confirm the table (live mode; audit-logged)
  LIVE_UNAVAILABLE / LIVE_SESSION_ACTIVE / LIVE_SESSION_MISMATCH / LIVE_DISCONNECTED /
  LIVE_STALE / LIVE_CAPABILITY_MISSING / LIVE_OP_FAILED - as replace_text (issue #154);
                         LIVE_OP_FAILED here also covers a document with nested tables

## apply_style

Apply a NAMED style (from list_styles) to text located via find --
the named-style counterpart to format_text's four boolean toggles.

`find` may be anchored ("<prefix>…<suffix>" with expected_length and/or
find_sha256 instead of expected_matches), exactly as on replace_text.

The text never changes, so the result carries ``text_unchanged: true`` and one
``excerpt`` instead of a before and an after, plus ``format_delta`` (only the direct
character formatting that moved; a named style may change none of it, so ``style_id``
is what shows it applied). ``verbose=true`` also returns ``runs_before`` /
``runs_after`` (issue #112).

This tool has no `write_mode` parameter -- it always writes the
.docx file directly. Issue #154: that means it now refuses with
LIVE_SESSION_ACTIVE if a pane session is connected for this document
(the exact case a real incident hit -- 14 calls to this tool
returned `applied: true` while a Live pane was open on the same
document, and every one was silently reverted by Word's own next
autosave). There is no live-mode equivalent of this tool today; use
`format_text`/`replace_text` with `write_mode="live"` for a live
character-formatting edit, or save and close the document in Word
(not just the pane) before calling this tool.

A CHARACTER style (w:type="character") applies to the matched run(s)
exactly like format_text (same locate()/normalization-ladder/
STRUCTURAL_BOUNDARY contract, same run-splitting rule, same
track_changes=True support via w:rPrChange). A PARAGRAPH style
(w:type="paragraph") applies w:pStyle to every paragraph CONTAINING a
matched run (the whole paragraph, not just the matched substring).

Issue #7: track_changes=True on a paragraph style now records a real
w:pPrChange -- built from a Word-authored fixture (Track Changes on,
one paragraph style change, saved: tests/fixtures/revision/
pstyle-tracked.docx), not guessed. Placement/shape match Word's own
output exactly: w:pPrChange lands as w:pPr's LAST child; a paragraph
with no prior w:pPr records an empty <w:pPr/> snapshot. A SECOND
tracked style change on the same paragraph (before any accept/
reject) reuses the existing change record -- updates the live
w:pStyle only, leaves the original id/date/snapshot untouched --
exactly matching real Word (verified: tests/fixtures/revision/
pstyle-tracked-twice.docx). A paragraph whose w:pPr already carries a
w:pPrChange authored by someone OTHER than the configured author
refuses TRACKED_CHANGES_PRESENT unless force=True -- a deliberate
safety policy this tool adds; Word's own desktop UI does NOT itself
enforce this (verified: tests/fixtures/revision/
pstyle-tracked-foreign.docx -- a second Word session under a
different user name silently overwrites the live style and leaves
the FIRST author's change record untouched). The identical foreign-
author gap already exists today for run-level w:rPrChange (never
checked before this issue, only w:ins/w:del ancestry was) -- fixed
here too, for both format_text and apply_style's character path
(see _check_foreign_rpr_change in text_edit.py).

Out of scope, same as w:rPrChange today: accept_tracked_changes/
reject_tracked_changes/list_open_items only handle w:ins/w:del, never
w:rPrChange/w:pPrChange -- accepting or rejecting a tracked style
change is not yet exposed by any tool.

Returns the eight evidence keys, plus runs_before/runs_after (as
format_text), style_id, style_type ("paragraph"/"character"),
`warnings` when non-empty, and (track_changes=True only) revision_ids/
track_changes -- revision_ids is empty when every touched
run/paragraph reused an existing own-authored change record rather
than allocating a new id (see above).

Errors: as format_text's file-mode errors ONLY (this tool has no
write_mode parameter, so it never itself goes live -- the
LIVE_UNAVAILABLE/LIVE_DISCONNECTED/LIVE_STALE codes format_text's own
docstring lists do not apply here), plus:
  STYLE_NOT_FOUND        - style_id is not a style in this document's styles.xml
  UNSUPPORTED_STYLE_TYPE - style_id names neither a paragraph nor a character style
  TRACKED_CHANGES_PRESENT - a matched run/paragraph already carries a w:rPrChange/
                           w:pPrChange authored by someone other than the configured
                           author; force=True to proceed (see above)
  LIVE_SESSION_ACTIVE    - a pane session IS connected for this document; use
                           format_text/replace_text(write_mode="live") for a live edit, or
                           save and close the document in Word, then retry (issue #154)

## diff_body_vs_file

Export the docx body as markdown and diff against a local file.

Use this tool when you need to compare a .docx's body against a local
markdown file -- e.g. to decide whether a push/sync is needed. The
docx side is rendered through the exact same projection
``read_document(format="markdown")`` uses (``part`` fixed to the body,
``word/document.xml``), so this compares against the same rendering
every other skill and tool sees, never a second, divergent one. The
server reads the local file directly (it runs locally); both ``path``
and ``file_path`` are resolved through the same allowlist/denylist
(VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS; see paths.py).

This is a read-only tool -- it makes no changes to the document or
file. It does not join MUTATING_TOOLS and does not return the eight
evidence keys.

Both sides are split with splitlines(keepends=True) before diffing --
matching GoogleDocs-MCP's diff_tab_vs_file exactly, on purpose, so a
file on disk that ends with a trailing newline (the common case) will
surface as a one-line difference against the body projection, which
does not end with one; a caller comparing against a projection should
account for that rather than read it as real drift.

AFFIRMATIVE-ONLY, same as GoogleDocs-MCP's diff_tab_vs_file: a
returned hunk, or identical=False, is solid evidence a difference
exists. identical=True is NOT proof no difference exists -- it only
means this particular markdown rendering of the docx body found none
against this particular reading of the file, which can under-report a
real difference for docx-specific reasons a plain text diff has no
way to see:
  - The docx side is a PROJECTION, not the OOXML itself. Content the
    markdown renderer cannot express as GFM (a merged/spanned table
    cell, a nested table -- see lossy_elements) is dropped from the
    comparison entirely, so two docx files that differ only in ways
    the renderer discards compare identical here even though the
    underlying .docx bytes are not. A non-empty lossy_elements in the
    result is a direct signal to distrust an identical=True verdict.
  - Only the BODY is compared. Headers, footers, footnotes/endnotes,
    and comments are out of scope by design (see diff_body_vs_file's
    own name) -- a real difference confined to one of those parts
    never surfaces here.
  - When Word's owner file is present, this reads a validated
    snapshot rather than the live file (core/document-backend-
    protocol.md §4's "reads never refuse" rule) -- the comparison
    is then only as fresh as the last flush/save the snapshot
    captured, not necessarily the very latest keystrokes.
A skill deciding whether a sync is needed should treat identical=True
as "no difference found by this projection," not as a guarantee the
two are in sync.

Response size is O(diff), not O(document):
  - normalize=true (default) folds both sides first -- curly->straight
    quotes, dash variants, NBSP-class spaces, soft hyphens, NFC, and
    trailing whitespace per line -- so those never surface as phantom
    edits. The response says normalized. identical is then judged on the
    folded text (a trailing newline difference does not count);
    identical_raw is always the byte-for-byte comparison, i.e. the old
    meaning of identical: a caller that needs "nothing differs, not even
    dashes or quotes" must read identical_raw (or pass normalize=false).
  - format="hunks" (default) returns hunks only; "unified" returns
    unified_diff only. Never both. An equal hunk carries no text, just
    {tag:"equal", body_lines_count, file_lines_count, body_range,
    file_range}; a fully identical pair returns hunks: [].
  - section_key restricts the body to one heading section (as
    find_sections reports it) and the file to the heading whose folded
    text matches (same duplicate-ordinal rule), up to the next heading
    of any level. body_range is then relative to that body section and
    file_line_offset is the file line the section starts at.
  - max_hunks (default 50) caps the changed hunks returned; truncated
    and remaining_hunks say what was cut.

Returns path, file_path, revision (the "<doc8>:<cmt8>" token),
identical, identical_raw, normalized, truncated, hunks (tagged
body_lines/file_lines/body_range/file_range per
difflib.SequenceMatcher.get_opcodes()) or unified_diff, warnings, and
lossy_elements (present only when non-empty).

Errors:
  INVALID_INPUT      - a bad path, file_path not found, file_path is
                       not a regular file, a bad format/max_hunks, or
                       the file has no heading matching section_key
  SECTION_NOT_FOUND  - section_key names no heading section in the body

## diff_table_vs_markdown

Report the cells that differ between a Word table and a markdown pipe
table (issue #109). Read-only: nothing is written, nothing is audited.

Give exactly one markdown input (`markdown` inline, or `markdown_path`, a
local file under the same allowlist/denylist as diff_body_vs_file) and
exactly one Word selector:

- `table_id` (single mode): one table, as list_tables reports it. The
  markdown table is chosen by `md_heading` (the first table under that
  heading) and/or `md_header_row` (a pipe row such as "| Name | Role |",
  matched on normalized text; with both, it must be one table under that
  heading). With neither, the markdown must hold exactly one table. Zero
  or several candidates, or a heading that appears twice in the markdown,
  is refused (INVALID_INPUT), never guessed.
- `section_key` (multi mode, from find_sections): every Word table in that
  section is paired with the markdown tables under `md_heading` (default:
  the section's own heading text) by their header-row cells. A repeated
  header on either side, a missing markdown heading, or no table on either
  side is refused.

Cells are compared as visible text. The Word side is read from the cell
XML (runs, tabs, line breaks, hyperlink text), the markdown side from the
parsed table (escaped pipes, links, inline code), then both are folded
(curly quotes, dashes, NBSP, NFC) and whitespace-collapsed. Bold/italic
are ignored unless `strip_emphasis=false`, which compares them as `**`/`*`
markers. A literal `*x*` typed in Word is therefore a difference from a
markdown `*x*` (which is emphasis, i.e. `x`).

A table with merged cells, a nested table, or a cell holding a field,
drawing or text box is NOT compared: it is listed under `unsupported`
with a reason and `identical` is false. `identical` is true only when
every table involved was fully compared with no cell, dimension, unmatched
or unsupported finding.

`source` ("auto" | "file" | "live") reads the Word table as read_document
does, and the result carries the same provenance (`source`, `live`,
`live_session`, warnings).

Single mode returns path, table_id, md_table_index, identical,
differences ([{row, col, docx_text, md_text}], 1-based, header row = row
1, only differing cells; a row/column present on one side only has null on
the other), dimensions (only when they differ), unsupported (only when
non-empty) and warnings. Multi mode returns path, identical, tables
([{table_id, md_table_index, header, differences, dimensions?}]),
unmatched_docx (table_ids), unmatched_md (md_table_index values) and
unsupported. md_table_index is the 0-based position of the table among all
tables in the markdown.

Errors:
  INVALID_INPUT     - selector/markdown violations above, a bad path or source
  TABLE_NOT_FOUND   - table_id does not match any table
  SECTION_NOT_FOUND - section_key names no heading section
  LIVE_UNAVAILABLE / LIVE_CAPABILITY_MISSING / LIVE_DISCONNECTED / LIVE_OP_FAILED - source="live"

## merge_paragraph_with_next

Merge two adjacent paragraphs of ONE table cell into one (issue #107).

The first paragraph's properties win; both paragraphs' runs and hyperlinks are kept
as they are, not flattened to text. LIVE MODE: Word keeps the SECOND paragraph's
properties when it joins two (checked on Word 16.114.3), so a live merge of two
paragraphs whose properties differ (style, alignment, indents, ...) is refused
before anything is written (STRUCTURAL_BOUNDARY): make them match, or use
write_mode="file". All indices are 1-based: `paragraph_index` is the
FIRST of the pair, so it must not be the cell's last paragraph (STRUCTURAL_BOUNDARY).
`expected_text` ([first, second]) guards both texts (CELL_TEXT_MISMATCH, nothing
written). `separator` (inline text, default "") is placed between the two texts.

Refuses (nothing written) comments, tracked changes, anchored shapes, fields,
bookmarks, footnote references and section breaks in either paragraph, and tables
with merged or nested cells. `track_changes=true` is refused in both modes
(INVALID_INPUT): how Word records a tracked paragraph merge is unproven.

`write_mode` ("auto" | "file" | "live"): the live write needs the pane's
"cell_paragraph_merge_v1" capability (LIVE_CAPABILITY_MISSING: reload the pane).
Both modes verify the result independently: the cell's paragraph texts and count,
the merged paragraph's properties, every hyperlink's target and text, and every other
cell of the table. Live reads around the write do not re-arm this client's baseline;
a document that moved after the write adds a `warnings` entry rather than hiding it.
VERIFICATION_FAILED means the write may have landed: re-read the table before retrying.
Audited. Returns the eight evidence keys.

## insert_paragraphs

Insert new paragraphs (and headings) before or after the body paragraph
containing `anchor`, optionally as tracked changes (issue #77).

`paragraphs` is a list of 1-50 objects `{"text": str, "style": str | None,
"color": str | None}`: one paragraph each (no newlines in `text`), `style`
a paragraph style name such as "Heading 2" or "Body Text", `color` six hex
digits such as "7030A0". Omitted `style` means the document's default
(Normal) paragraph style, never the anchor's. Omitted `color` is not set:
in live mode Word may carry the anchor's run color onto the new paragraph,
so pass a color to pin it. Every inserted paragraph's text, style and color
is read back; the evidence lists them under `inserted`.

`anchor` is located like `replace_text`'s `find` and must match exactly
once (`expected_matches` must be 1); the anchor must sit in a top-level
body paragraph, not a table cell. `position` is "after" (default) or
"before"; under live section locks only "after" is allowed.

STRUCTURAL ANCHORS (issue #107): instead of text, `anchor` may be one of
{"after_table_id": N} (insert after top-level table N, the id list_tables
reports) or {"empty_paragraph_after_table_id": N} (fill the empty body
paragraph directly after table N, keeping its own style unless a `style` is
given; STRUCTURAL_BOUNDARY if there is none, or it holds text, a field,
bookmark, tab, break, hyperlink, drawing, comment anchor or tracked change).
Both need position="after" and expected_matches=1. File mode verifies the
paragraph count and the content on either side; live mode verifies the inserted
texts and the whole-document paragraph count. Live mode refuses to fill when
Word's table numbering and the body's disagree (a table inside a content
control). Live mode needs the
pane's "paragraph_positions_v1" capability (LIVE_CAPABILITY_MISSING:
reload the pane).

`track_changes=True`: live mode records the insertions as Word's
signed-in user; file mode wraps the new runs in `w:ins` under the
configured author (the paragraph marks themselves are not tracked).

`write_mode` ("auto" | "file" | "live", default "auto") routes exactly as
in `replace_text`: see its docstring. A live write that fails the pane's
read-back is deleted again when untracked (the message says whether that
worked); a tracked one is left for Word's Review pane. Live
`revision_before`/`revision_after` are `"live:sha256:<hex>"` body hashes.

Errors:
  INVALID_INPUT, ZERO_MATCH, MATCH_COUNT_MISMATCH, STRUCTURAL_BOUNDARY,
  STYLE_NOT_FOUND, UNSUPPORTED_STYLE_TYPE, OUTSIDE_LOCKED_SECTION,
  VERIFICATION_FAILED, LIVE_UNAVAILABLE, LIVE_SESSION_ACTIVE,
  LIVE_SESSION_MISMATCH, LIVE_DISCONNECTED, LIVE_STALE,
  LIVE_CAPABILITY_MISSING, DOCX_LOCKED, SYNC_IN_FLIGHT, REVISION_CONFLICT,
  EXTERNAL_EDITOR_ACTIVE, OPC_INVALID

## live_changes_prepare

Prepare a scoped approved run without modifying Word. Operations are
{op, section, args}; section is a unique heading slug. Supports
replace_cell_markdown, replace_table_row, replace_text, delete_paragraph,
and a final live_save. Returns a preview, run_id and manifest_sha256. The preview
shows each operation's find/replace/markdown/cells/anchor_text (cut to 200
characters with an ellipsis), its table address and expected_matches, so the
lead can approve from the response. It is DERIVED from the manifest: the
manifest file and its digest remain the authority, and approving that digest
binds the exact operations, not the clipped preview.
replace_text args may also include replacement_segments with the
explicit font/hyperlink schema documented on replace_text; concatenated segment
text must equal replace. The approval preview includes the exact segments.
replace_text may name its target as an anchored find ("<prefix>…<suffix>" with
expected_length and/or find_sha256) or as {para_ref, expected_sha256} where pN counts
the section's paragraphs (p0 = its heading). Prepare resolves it against the section as
the earlier operations leave it; the manifest stores the resolved full find (no 230-
character limit then; over 230 needs anchored_find_v1) and the preview shows it.
replace_cell_markdown args may carry markdown_file (a .md/.txt path, issue #113)
instead of markdown: prepare reads it once and stores the CONTENT as markdown in
the manifest, so the digest binds the text, not the path.
Requires approved_changes_v1; explicit segments also require formatted_replacements_v1.
Formats naming only some fields (partial_font_segments_v1) are resolved here against
the first replaced run; the manifest binds the resolved segments and that
font_baseline, and execute refuses FORMAT_DRIFT if the run was reformatted since.
