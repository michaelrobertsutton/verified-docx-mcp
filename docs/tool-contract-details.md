# Extended tool contracts

Historical explanations retained outside the advertised MCP descriptions.

## replace_text

Replace every occurrence of `find` with `replace`, atomically.

Live-only replacement_segments specifies explicit text/format/hyperlink segments.
Text concatenates exactly to replace; each format includes bold, italic, underline
(None/Single/Double), strikeThrough, doubleStrikeThrough, subscript, superscript,
color (#RRGGBB), name, and size (points). Null hyperlink explicitly means no link.
Mixed spans require explicit segments; no neighbour-format inference is used.

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
"allow_comment_loss_ids"?} objects; semantics per edit are those of
`replace_text` (expected_matches required, row/scope options not
available). `allow_comment_loss_ids` is [{"comment_id", "reply_count"}]
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
`results` has one entry per edit: {index, status, evidence?, code?} with
status "applied" | "refused_before_write" | "unknown" | "not_attempted".
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

## get_table

Full row/cell detail for one table, addressed by the table_id
list_tables reports.

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
has_nested_table, rows (list of list of {row_index, cell_index,
grid_span, v_merge, text}). grid_span is w:tcPr/w:gridSpan (1 when
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
