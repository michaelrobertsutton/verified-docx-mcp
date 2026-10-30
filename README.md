# verified-docx-mcp

An MCP server for local `.docx` files with verified writes. It is the
`docx` counterpart to
[`verified-googledocs-mcp`](https://github.com/michaelrobertsutton/verified-googledocs-mcp):
the same evidence-first contract (every mutating tool re-reads the file
after writing and returns before/after evidence, never a bare "success"),
applied to a local Word document on a synced drive (OneDrive, iCloud Drive)
instead of a Google Doc.

## Status

Issue #28 of the build plan, through WP-16a (server half; WP-16b, the
JennyStack-side KP resume rendering contract, is a separate PR). Reading
and writing a `.docx` package needs no Word installation at all — this
server manipulates OOXML directly. Word is required only for
`export_pdf` (rendering a page-accurate PDF), and that additionally
needs a macOS Automation grant for the app hosting this server's
process.

**Platform note:** this server is macOS-only, and not only for
`export_pdf`. `insert_image`'s SVG path (issue #28 WP-15a) shells out to
macOS's built-in `sips` to rasterize the PNG fallback part Word requires
for an SVG — no `cairosvg`/`rsvg-convert`/Inkscape dependency is added,
but `sips` itself is not optional and is not available on Linux/Windows.
PNG-only `insert_image` calls do not need it. (`mutations.py`'s own
`scutil`-based conflict-copy detection is the other pre-existing
macOS-only dependency, for reference.)

Tools implemented so far:

- **`export_pdf(path, output_path, close_after=True, section_keys=None)`** —
  render a `.docx` to PDF via Microsoft Word automation and report its page
  count; the result also reports `closed_after` (whether the staged copy's
  window was closed), `close_error` (why not, when it wasn't), and
  `left_open_document` (null once closed). It also reports per-section page
  spans: `sections` (each heading section's `start_page`/`start_fraction`,
  `end_page`/`end_fraction`, and derived `pages`, verified against
  `find_sections`' `heading_text`) and `page_height_pt`, restrictable to
  specific sections via `section_keys` — see the tool's own docstring for
  the default-mode-degrades-vs-explicit-mode-raises verification contract.
- **`lock_status(path)`** — report Word/LibreOffice owner-file presence
  and sync-quiesce state, as data only. Never refuses. Also reports
  `external_activity` (issue #27): whether the file on disk still matches
  what this server last wrote — the after-the-fact check for a
  co-authoring editor (Office Online) that leaves no owner file and no
  Live session. See [Live mode](#live-mode).
- **`list_parts(path)`**, **`read_document(path, format, part)`**,
  **`find_sections(path, part)`**, **`list_page_sections(path, part)`**,
  **`list_styles(path)`** — read a `.docx` part as markdown/text/runs,
  and enumerate its heading-delimited sections, page-layout sections, and
  styles. Never refuse on a locked file (a validated snapshot is read
  instead). `read_document`, `find_sections`, `list_tables` and
  `get_table` also take `source="auto"|"file"|"live"` (issue #33): they
  read the body a connected Word pane holds, and every response says
  which `source` it came from. A file read while a pane session exists
  for that file name carries a `live_session_ignored` warning. See "Live
  reads" under [Live mode](#live-mode).
  Issue #110, so a caller never has to read a whole body to learn one thing:
  `read_document(start_para_ref, end_para_ref)` returns only an inclusive
  `pN`..`pM` paragraph range (a table the range cuts is reported as a
  `partial_table` in `lossy_elements`); `read_document(to_file=...)` writes
  the rendered body to an allowed path and returns only
  `{written_to, sha256, chars, paragraph_count, sections}`;
  `find_sections` entries carry `chars`, and `find_sections(digests=true)`
  adds `sha12`, a per-section digest that agrees with the live pane's.
- **`replace_body_markdown(path, markdown, ...)`**,
  **`replace_range_markdown(path, section_key, markdown, ...)`**,
  **`append_markdown(path, markdown, ...)`** — markdown -> OOXML writes,
  guarded by a lock/sync/revision check that runs before any temp file is
  written, an atomic write with OPC validation and a rollback on a
  failed post-write verification, and a full before/after/revision
  evidence envelope on every call.
- **`replace_text(path, find, replace, expected_matches, ...)`**,
  **`format_text(path, find, style, expected_matches, ...)`** —
  targeted text edits located via a normalization ladder (exact ->
  curly/straight quotes -> NBSP/whitespace collapse -> soft-hyphen
  strip), with run splitting that clones a boundary run's original
  `w:rPr` verbatim onto every surviving piece. `format_text`'s `style`
  accepts `bold`/`italic`/`underline`/`strike` (booleans) and `color` (a
  6-hex string) — an explicit color clears any theme-color attributes on
  the run, so Word doesn't keep rendering a stale theme color over it.
  Both take `within_row_containing` (scope `find` to one table row,
  identified by a second unique anchor in that row — for a table with
  several cells reading identical text) and a `write_mode: "auto" |
  "file" | "live"` parameter (default `"auto"`): `"auto"` edits
  through a connected Word task pane instead of the file whenever a Live
  pane is connected for that document, otherwise the file path — see
  [Live mode](#live-mode) below. `apply_style(path, find, style_id,
  expected_matches, ...)` has no `write_mode` parameter and always takes
  the file path — including refusing (`LIVE_SESSION_ACTIVE`) while a Live
  pane is connected for that document, same as every other write_mode-less
  mutating tool (issue #154 — see "Live mode" below).
- **`live_save(path)`** — ask the connected pane to save the live
  document, then report both the pane's own live revision token and a
  bridge back to the file-mode revision contract (`file_revision`) for a
  caller that wants to keep working against the file afterward.
- **`list_open_items(path, source="auto")`**, **`accept_tracked_changes(path,
  revision_ids?, ...)`**, **`reject_tracked_changes(path, revision_ids?,
  ...)`** — list open comments/pending tracked changes (Google's response
  shape), and accept/reject `w:ins`/`w:del` revisions by id or all at
  once. `source="live"` reads through a connected Word task pane instead
  (issue #106 WP-4 — see "Live mode" below).
- **`track_changes: bool`** on every mutating tool above — when true,
  the edit is wrapped in `w:ins`/`w:del` (or a `w:rPrChange` for a style
  change) with an author/date/id, rather than applied directly. A
  revision authored by the server's own configured identity never blocks
  a later tracked write; one authored by anyone else does, unless
  `force=True`.
- **`add_anchored_comment(path, quote, text, expected_matches,
  write_mode="auto", ...)`**, **`get_comment_thread(path, comment_id)`**,
  **`reply_to_comment(path, comment_id, text, write_mode="auto")`**,
  **`resolve_comment(path, comment_id, write_mode="auto")`** — create an
  anchored comment, read a comment and its replies (by durable id), reply
  to one, and resolve one. Builds all five interlocking comment parts a
  real Word comment needs; verified part-by-part against a Word-authored
  golden fixture (`tests/fixtures/comments/golden-comment.docx`).
  A multi-paragraph comment's identity is keyed on its LAST paragraph,
  matching Word's own commentsIds.xml/commentsExtended.xml convention, and
  every `comment_id` `list_open_items` reports (durableId or, lacking one,
  the raw `w:id`) is accepted by the three tools above.
  `write_mode="live"` sends the same operation through a connected Word
  task pane instead (issue #106 WP-4 — see "Live mode" below), accepting
  either a `live:<id>` handle or a correlated durableId/w:id.
- **`delete_comment(path, comment_id, write_mode="auto", ...)`**,
  **`delete_resolved_comments(path, comment_ids, write_mode="auto")`** —
  remove a comment thread (the root and all its replies) instead of only
  marking it resolved (issue #80). Thread roots only: a reply id is
  refused and names its root. The batch form takes resolved threads, resolves
  and de-duplicates every id first, and refuses the whole call (listing
  each offender) before deleting anything if one is unknown, a reply or
  still open. File mode removes the `w:comment`, its
  commentsExtended/commentsIds/commentsExtensible entries and the range
  markers in one atomic write that touches no other part, then re-reads
  the package and rolls back unless every other comment and anchor is
  unchanged. Live mode (`Comment.delete()` through the pane, capability
  `comment_delete`; reopen the Live pane to pick it up) verifies with an
  independent re-list and cannot roll back, so a failure part-way reports
  `deleted`/`remaining`/`unknown` threads. Evidence also carries
  `deleted_thread_count`, `deleted_reply_count`, `comments_before`/
  `comments_after` and `other_comments_unchanged`.
- **Lock guard layers 0/3/4** — every mutating tool above now runs inside
  a same-machine `.jsclaim` mutex (`O_EXCL`, released even on failure)
  alongside a no-op `remote_checkout` seam for a future Microsoft Graph
  checkout, and reports `conflict_copy_detected` (plus, when non-empty,
  `conflict_copies`/`sibling_files_changed`) after every successful
  write — a post-write sweep for sync-conflict sibling files, never
  raised as an error since the write it describes already succeeded.
  Detection only, not prevention (core/document-backend-protocol.md §9):
  the naming patterns it matches are client- and locale-dependent, so
  their absence is not proof no conflict occurred.
- **`diff_body_vs_file(path, file_path, normalize, format, section_key, max_hunks)`** —
  export the docx body's markdown projection (the same one
  `read_document(format="markdown")` uses) and diff it against a local
  markdown file with `difflib`, the same mechanism GoogleDocs-MCP's
  `diff_tab_vs_file` uses. Read-only and affirmative-only: `identical: true`
  means no difference was found by this particular projection, not a
  guarantee none exists (see the tool's own docstring for the docx-specific
  reasons why). Issue #110: both sides are normalized by default (curly
  quotes, dashes, NBSP, soft hyphens, NFC, trailing whitespace) so those
  never show as edits (`identical_raw` is the byte comparison); the response
  is hunks **or** a unified diff, never both, with no text for equal hunks;
  `section_key` scopes the diff to one heading section; `max_hunks` caps it.
  The response is O(diff), not O(document).
- **`diff_table_vs_markdown(path, table_id, markdown, markdown_path, md_heading, md_header_row, section_key, source, strip_emphasis)`** —
  read-only. Reports only the cells that differ between a Word table and a
  markdown pipe table (single mode, `table_id`), or between every table in a
  heading section and the markdown tables under a heading, paired by header
  row (multi mode, `section_key`). Unsupported tables (merged cells, nested
  tables, fields/drawings/text boxes in a cell) are listed, never compared,
  and `identical` is true only when everything was compared and matched.
  Issue #109.
- **`list_tables(path, part)`**, **`get_table(path, table_id, part)`** —
  enumerate every `w:tbl` (including one nested inside a cell, which gets
  its own `table_id`) and report one table's full row/cell detail,
  including `w:gridSpan`/`w:vMerge` per cell. Never refuse on a locked
  file.
- **`replace_table_row(path, table_id, row_index, cells, ...)`** —
  replace one row's cells wholesale, one markdown string per cell.
  Refuses (`MERGED_OR_NESTED_TABLE`) for the WHOLE table the moment any
  cell in it is merged (`w:gridSpan`/`w:vMerge`) or a `w:tbl` is nested
  inside a cell, mirroring GoogleDocs-MCP's own merged-cell refusal.
- **`replace_cell_markdown(path, table_id, row_index, cell_index,
  markdown, ...)`** — replace one cell's content, leaving its own
  `w:tcPr` byte-identical — the only write path safe on a merged cell,
  and the intended path for an Appendix-A style band-and-border table.
  Supports multi-level bulleted/numbered markdown inside the cell, one
  `numId`/`abstractNum` tree shared across nesting levels via increasing
  `w:ilvl`, the same machinery `replace_body_markdown`/`append_markdown`
  already use.
- **`insert_table(path, rows, style_id, header_rows=0, grid_dxa=None,
  cant_split=False, anchor=None, write_mode="auto", style_from_table_id=None,
  style_builtin=None, font_size_pt=None, ...)`** — insert a new table, one
  markdown string OR cell-spec object (`{"markdown", "span", "v_merge",
  "fill", "color", "bold", "align", "valign"}`) per cell — a spanning
  (`w:gridSpan`) or vertically merged (`w:vMerge`) title/header row,
  shading, and per-row `w:tblHeader` are all write paths now, not just a
  read-side report. `style_id` is REQUIRED and must name an existing
  `w:type="table"` style in the document; `grid_dxa` gives explicit
  per-column dxa widths, otherwise columns split evenly across the
  text-column width `list_page_sections` reports (unchanged default).
  `anchor` places the table by `section_key` (`position` or
  `after_paragraph_text`) or `after_table_id` instead of always appending
  at the end of the body. Cell-level `fill`/`valign`/`span`/`v_merge` live
  in `w:tcPr` and survive a later `replace_cell_markdown`; `bold`/`color`/
  `align` live on the cell's own runs/paragraphs and do not.
- **`insert_image(path, image_path, width_in, ...)`** — append a new
  inline picture at the end of the body, from a LOCAL `.png` or `.svg`
  file (read natively — no `IMAGE_SOURCE_UNSUPPORTED`, unlike
  GoogleDocs-MCP's URL-only tool). `width_in` defaults to the
  text-column width `list_page_sections` reports; an `.svg` embeds
  natively (Word 2016+'s own SVG extension) WITH a PNG fallback part
  Word requires, rasterized from the SVG's own native pixel size via
  macOS `sips`. Records `design_width_in`/`design_height_in`,
  `placed_width_in`/`placed_height_in`, and `effective_scale =
  placed_width_in / design_width_in` — a real measured ratio.
- **`apply_style(path, find, style_id, expected_matches, ...)`** — apply
  a NAMED style (from `list_styles`) to text located via `find`, the
  named-style counterpart to `format_text`'s boolean toggles. A
  character style applies to the matched run(s) exactly like
  `format_text` (including `track_changes=True`, `w:rPrChange`); a
  paragraph style applies `w:pStyle` to every paragraph containing a
  matched run, including `track_changes=True` (`w:pPrChange`, built from
  a real Word-authored fixture — see `text_edit.py`'s docstrings for the
  own-author-reuse and foreign-author-refusal rules Word's own behavior
  informed). `accept_tracked_changes`/`reject_tracked_changes` still only
  handle `w:ins`/`w:del`, never `w:rPrChange`/`w:pPrChange`.
- **`read_header_footer(path)`** — read every header/footer part's
  content as markdown in one call. Never refuses on a locked file.

More tools land in a later work package of the same plan.

## Interoperability

Every read/write path above is exercised by `tests/unit` against
Word-authored fixtures and, for comments specifically, verified part by
part against a golden `.docx` produced by the lead in Word desktop
(`tests/fixtures/comments/golden-comment.docx`; its provenance was
confirmed by the lead). What those tests cannot cover is behavior that
only shows up with a live human at a keyboard and a real sync client in
the loop: how a sync client actually names a conflict file, how Word's
own Review pane renders a tracked change, how Word Online renders a
comment this server wrote.

The lead has ruled that a live round trip through Word desktop and Word
Online is no longer a landing gate for this repo. Instead, each item
below is recorded as **verify at first real use**: the exact assumption
this server's code depends on, stated precisely enough that a future
reader knows exactly what is unproven and what would actually break if
the assumption turns out to be wrong.

1. **Conflict-copy filename (issue #28 WP-10).**
   Assumption: OneDrive names a conflict copy `<stem>-<Machine>.docx`
   using the resolved local machine name -- on this machine
   `Michaels-MacBook-Pro` (`_local_machine_names()` in `mutations.py`;
   confirmed against this machine's own `scutil --get ComputerName` /
   `LocalHostName` output, both of which already normalize to that
   string). `conflict_copy_sweep`'s `_matches_conflict_copy_pattern`
   matches that branch only on an exact, case-insensitive match against
   `_local_machine_names()`.
   If wrong: a real OneDrive client emitting a different machine-name
   form (a different normalization, a user-set device label, a non-Mac
   client) means a genuine conflict lands in `sibling_files_changed`
   instead of raising the `conflict_copy_detected` evidence flag --
   still visible to a caller that reads the evidence, just not flagged
   as unambiguously as the matched branch is.
   This is a deliberate, already-documented trade-off, not a gap
   discovered here: `_matches_conflict_copy_pattern`'s own docstring
   notes that a conflict copy from another, unenumerable machine falls
   through the same way, consistent with
   core/document-backend-protocol.md §4's "the absence of a match is not
   proof no conflict occurred."

2. **Word Online round trip (issue #28 WP-09).**
   Assumption: a comment created, replied to, and resolved by this
   server's tools shows correct threading and resolved state when the
   file is reopened in Word desktop AND Word Online.
   What is actually proven: every comment part this server writes
   (`comments.xml`, `commentsExtended.xml`, `commentsIds.xml`,
   `commentsExtensible.xml`, `people.xml`) is verified part by part
   against the golden, Word-desktop-authored fixture above.
   What is not proven: this server has never had one of its own comments
   opened in a live Word Online session. If wrong, Word Online's comment
   renderer disagrees with Word desktop's about some part-level detail
   the golden-fixture comparison did not catch (the fixture was authored
   in Word desktop, never in Word Online) -- threading or resolved state
   could render incorrectly specifically in the browser client this
   repo's fixtures never exercised.

3. **Review-pane author (issue #28 WP-07b-a).**
   Assumption: a tracked `replace_text` shows in Word's Review pane
   under the configured author name -- resolved by `author.py`'s
   `resolve_author_name()` from `~/.jennystack/config.json`'s
   `author_name` key, or, absent that, the current macOS account's full
   name (`pw_gecos`, falling back to `id -F`) -- currently
   `Michael Sutton` on this machine.
   What is actually proven: the OOXML this server writes is well-formed
   and matches documented Word conventions for a `w:ins`/`w:del`'s
   `w:author` attribute, and this server's own projection reads that
   author string back unchanged.
   What is not proven: nobody has opened one of these files in Word and
   looked at the Review pane. If wrong, the name Word actually displays
   could differ from the literal `w:author` string this server wrote
   (e.g. a Word-side display quirk that resolves a name against a
   signed-in account) -- the UI rendering is unobserved.

4. **Nested tracked edit (issue #28 WP-07b-a).**
   Assumption: Word renders `<w:ins><w:del>...</w:del><w:ins>...</w:ins></w:ins>`
   sensibly when a second `track_changes=True` edit lands on a still-
   pending own insertion -- the named scope limit in `text_edit.py`'s own
   module docstring.
   What is actually proven: read-back correctness holds regardless --
   `projection.py` excludes `w:del` content unconditionally, so the
   superseded text is correctly invisible to every read tool either way,
   directly exercised by
   `test_text_edit.py::test_own_author_tracked_edit_does_not_deadlock_a_second_tracked_edit`
   (chained `replace_text(..., track_changes=True)` calls over the same
   span), not merely assumed.
   What is not proven: the visual shape of that nesting in Word's own
   Review pane. If wrong, Word could render the nested insertion/deletion
   confusingly (e.g. a crossed-out "insertion" that reads as ambiguous or
   duplicated) even though every read tool in this server still reports
   the correct final text.

5. **SVG text extractability through export_pdf (issue #28 WP-15a).**
   Assumption: text inside an `insert_image`-embedded SVG survives
   `export_pdf` (Word's own PDF export of a document containing that SVG)
   as extractable text, not a rasterized/flattened image -- i.e. that
   Word's PDF exporter renders the native SVG branch (`a:blip`'s
   `a14:svgBlip` extension this server writes) as real vector text rather
   than falling back to the PNG fallback part or flattening the SVG to a
   bitmap. The plan named this a `LEAD:` live check; the lead has since
   ruled the remaining live checks are no longer a landing gate for this
   repo (see this section's own opening paragraph), so it is recorded
   here in the same form as items 1-4 instead of being run.
   What is actually proven: the OOXML this server writes is well-formed
   (`opc_valid` passes with both the SVG part and its PNG fallback part
   present and correctly related -- `tests/unit/test_images.py`), and the
   SVG's own text content is written byte-for-byte into the `.svg` media
   part (nothing about this server's own write path could corrupt or
   strip it).
   What is not proven: nobody has run `export_pdf` against a document
   this server inserted an SVG into and inspected the resulting PDF's own
   text layer. If wrong, Word's PDF exporter treats the SVG extension
   branch as non-authoritative for export purposes and falls back to
   rasterizing the PNG fallback part instead -- the PDF would still look
   correct (the fallback is a real rasterization of the same SVG, via
   macOS `sips`), but the SVG's own text would not be independently
   selectable/extractable in the PDF.

None of the five blocks this server from being used; each is a
live-Word-rendering question this repo's own test suite, which needs no
Word installation to read or write a `.docx`, cannot answer by itself.
The next real use of this server against a live synced folder is the
first opportunity to confirm or correct any of them.

## Install

Requires Python 3.12+ and [`uv`](https://docs.astral.sh/uv/).

```bash
uv sync
```

## Run

As an MCP server (stdio transport), typically registered by a client rather
than run directly:

```bash
uv run --directory "<path-to-this-clone>" verified-docx-mcp
```

Diagnose the Word render path on this machine (run from the SAME
application that will host the MCP server — macOS grants Automation, and
file access, per hosting app, not per terminal emulator in general):

```bash
uv run --directory "<path-to-this-clone>" verified-docx-mcp doctor
```

## Test

```bash
uv run pytest tests/unit
```

`tests/unit` is offline and runs against committed fixture `.docx`/`.pdf`
files — no Word installation required. `tests/live` needs a real Word
install and a macOS Automation grant; it is skipped unless `--run-live` is
passed.

## Auto-allowing the read-only tools

Every read tool carries the MCP annotations `readOnlyHint: true` and
`idempotentHint: true`; the delete/replace-all/accept/reject tools carry
`destructiveHint: true`. Clients that honor `readOnlyHint` (and Claude Code
permission rules) can therefore treat reads differently from writes. To let Claude
Code run this server's read-only tools without a prompt, add them to
`permissions.allow` in `.claude/settings.json` (or `settings.local.json`):

```json
{
  "permissions": {
    "allow": [
      "mcp__verified-docx-mcp__read_document",
      "mcp__verified-docx-mcp__find_sections",
      "mcp__verified-docx-mcp__list_page_sections",
      "mcp__verified-docx-mcp__list_tables",
      "mcp__verified-docx-mcp__get_table",
      "mcp__verified-docx-mcp__list_open_items",
      "mcp__verified-docx-mcp__get_comment_thread",
      "mcp__verified-docx-mcp__list_styles",
      "mcp__verified-docx-mcp__list_parts",
      "mcp__verified-docx-mcp__list_shapes",
      "mcp__verified-docx-mcp__list_textboxes",
      "mcp__verified-docx-mcp__read_textbox",
      "mcp__verified-docx-mcp__read_header_footer",
      "mcp__verified-docx-mcp__diff_body_vs_file",
      "mcp__verified-docx-mcp__live_status",
      "mcp__verified-docx-mcp__live_list_sections",
      "mcp__verified-docx-mcp__live_revision",
      "mcp__verified-docx-mcp__lock_status",
      "mcp__verified-docx-mcp__export_pdf",
      "mcp__verified-docx-mcp__read_audit_evidence"
    ]
  }
}
```

`export_pdf` writes a PDF next to the document and never modifies the `.docx`;
`lock_status` may persist a small first-seen timestamp for a divergence. Drop either
line if you want a prompt for it. Write tools are deliberately not listed.

## Live mode

A Word task-pane add-in (`addin/`) plus a local bridge
(`src/verified_docx_mcp/live/`) for live co-editing:
[issue #106](https://github.com/michaelrobertsutton/JennyStack/issues/106).
The bridge serves the pane over local HTTPS (port 53135) and runs a WSS
ops channel (port 53136) the pane connects to and applies
search/replace/format/comment operations from, always reading the
document back after a mutation. The MCP server starts the bridge lazily
(idempotent) the first time a live-aware tool is called.

`live_status` (read-only) reports bridge and connected-pane state:
`{bridge_running, pid, owner_pid, is_owner, mode, port, ops_port, sessions: [{document_name,
document_url, connected_since, last_heartbeat_age_s, body_sha256,
requirement_sets}], session_collisions}` — an empty `sessions` list is
normal before the lead opens the pane in Word. `session_collisions`
(empty unless it's happened) flags two *different* documents that share
a file name connecting at once — `SessionRegistry` keys purely on
basename, so it can't tell them apart on its own; check this before
trusting a live write's routing if that's possible.

Write responses are bounded (issue #112). A write that rewrites a whole body, section or
revision set returns hashes, sizes and a short excerpt instead of the whole before/after, plus an
`evidence_ref`; **`read_audit_evidence(evidence_ref, field?, to_file?)`** fetches the full record
from the audit log (`$XDG_STATE_HOME/verified-docx-mcp/audit.jsonl`, append-only and never
rotated by the server; with `VERIFIED_DOCX_MCP_AUDIT_EXCERPTS=0` the text is not stored).
`verbose=true` (on `replace_text_batch`, `format_text`, `apply_style`, `replace_text` errors and
the approved-run tools) restores only the named extras (per-edit batch text, `runs_before`/
`runs_after`, `ladder_report`, `performance`), never whole-document text. Constant envelope
fields (`verified_via`, `document_name`, `audit_logged: true`, ...) are omitted. See
`docs/acceptance/112-write-tokens.md`.

`live_revision(path)` (issue #110) returns the live body's current
`revision` token plus a `sha12` digest per heading section in one pane round
trip, with no body in the response. It re-arms your write baseline exactly
like a live `read_document`, so a `revision_before` taken from it is accepted.
It needs the pane capability `body_hash_v1` (reload the pane after upgrading);
an older pane still works through a `body_ooxml` fallback, reported as
`via: "body_ooxml_fallback"`. `live_list_sections(compact=true)` returns
`[[section_key, heading_text, sha12], ...]` instead of the 64-hex hashes.

On POSIX several MCP processes share one bridge by default (one owner, the
rest attach; automatic re-election if the owner exits). Set
`VERIFIED_DOCX_SHARED_BRIDGE=0` to opt out. See
[docs/live-mode.md](docs/live-mode.md#shared-bridge-several-agent-sessions-one-word-pane-47).

`write_mode: "auto" | "file" | "live"` (default `"auto"`) is wired up on
`replace_text`/`format_text` (WP-3) and on
`add_anchored_comment`/`reply_to_comment`/`resolve_comment` (WP-4), with a
parallel `source="live"` on `list_open_items` — each goes through the connected
pane instead of the on-disk `.docx` parts whenever a Live pane is connected for
that document, otherwise the file path. `live_save(path)` asks the pane to save.
`list_open_items(source="live")` carries a `correlation` list bridging
a live comment's own id back to the durableId/w:id file mode reports (Office.js's
`Comment.id` is unrelated to either).

Issue #22 B3: none of the live comment tools (or `live_save`) require
`path` to name a file that exists locally anymore — a SharePoint/OneDrive
document with no local sync now works directly, without fabricating a
same-named local stand-in file first. When there's genuinely no local
file, `list_open_items(source="live")`'s `pending_suggestions` is `null`
(never `[]`, which would look like "nothing pending") and it carries
`file_side_available: false`; `live_save`'s `file_revision` is `null`.
Issue #33: `read_document`, `find_sections`, `list_tables` and `get_table`
take `source="auto"|"file"|"live"` and can read the live body (body part
only; `list_parts` and the other structural read tools stay file-only, and
warn with `live_session_ignored` when a pane session exists). `auto`
reads live when a pane that reports the `body_ooxml` capability is
connected, and falls back to the file, with a warning saying why, when it
can't be sure the pane holds this document. (`get_table`'s live read is
issue #34's `table_get` op, not `body_ooxml`.) A live read never falls back
to the file after it fails. In a live read `revision` is
`live:sha256:<body hash>`; images are not included. Details:
[docs/live-mode.md](docs/live-mode.md#live-reads-issue-33).

Issue #27: `replace_cell_markdown` also takes `write_mode` — a live edit is
compare-and-set against the cell's current text and read back
independently, limited to paragraphs with bold/italic/links (anything
structural is refused before it is sent). File-mode writes also guard
against a co-author the server cannot otherwise see (Office Online): if
the file no longer matches what this server last wrote, the write refuses
with `EXTERNAL_EDITOR_ACTIVE` unless `allow_concurrent_editor=True`, and
`revision_after` is now the token of the bytes actually written, so an
overwrite landing right after a write can't hide inside it. This is
detection after the fact, not prevention — read the limits, and the
unverified Word-for-the-web section, in
[`docs/live-mode.md`](docs/live-mode.md#office-online--word-for-the-web-co-authoring-issue-27).

Issue #34: `insert_table` and `replace_table_row` also take `write_mode`, and
`get_table` takes `source`, so a table can be added to or edited in a document
a co-author has open. Live `insert_table` is a subset of file mode (no merged
cells; a paragraph-text or `after_table_id` anchor; style by name, copied from
another table, or a built-in) and every requested property is re-read from
Word and verified; a table that lands but fails that check is audit-logged
before the error is raised. The pane logic is tested against a mock only; see
[`docs/live-mode.md`](docs/live-mode.md#live-tables-issue-34) for the limits
and the manual runbook.

Issue #77: `insert_paragraphs` adds styled, optionally colored paragraphs or
headings before/after one anchored body paragraph, optionally as tracked
changes, in a document a co-author has open (`write_mode`, like `replace_text`).
Live mode sends the pane's `paragraph_insert` op; the style is resolved before
anything is written, every inserted paragraph's text, style and color is read
back, and an untracked write that fails the read-back is removed again. The pane
logic is tested against a mock only; see
[`docs/live-mode.md`](docs/live-mode.md#live-paragraph-insertion-issue-77).

Issue #107: cell edits keep their formatting. `replace_cell_markdown` inherits the
cell's first-run font name, size and colour (`inherit_format="none"` opts out);
`delete_paragraph` may remove a non-final paragraph inside a cell (Word cannot remove
a cell's final one); `merge_paragraph_with_next` joins two adjacent paragraphs of one
cell, keeping both paragraphs' runs and hyperlinks (tracked merges are refused; a live
merge of two paragraphs with different properties is refused too, because Word keeps
the second's, while file mode keeps the first's); `insert_paragraphs` takes
`{"after_table_id": N}` or `{"empty_paragraph_after_table_id": N}` as a structural
anchor; `replace_text_batch` honours `within_row_containing` per edit; and `live_save`
no longer fails `LIVE_STALE` while someone types. Checked on Word for Mac 16.114.3
(what was and was not covered is in
[`docs/acceptance/107-cell-editing.md`](docs/acceptance/107-cell-editing.md)).

Issue #154: every OTHER mutating tool — `apply_style`,
`insert_image`,
`replace_range_markdown`, `replace_body_markdown`, `append_markdown`,
`accept_tracked_changes`, `reject_tracked_changes` — has no `write_mode`
parameter and always takes the file path, but now **refuses**
(`LIVE_SESSION_ACTIVE`) while a Live pane is connected for that document,
rather than writing to disk underneath Word's own open, autosaving
buffer. The remedy is to save and close the document itself in Word, not
just the pane. Every mutating tool's evidence also now carries its own
`write_mode` key (`"file"` or `"live"`), so a caller never has to infer
routing from the shape of a `revision_*` token.

Full architecture, the `write_mode` rule and
live evidence shape, the correlation algorithm, and the lead's sideload runbook:
[`docs/live-mode.md`](docs/live-mode.md).

## Path safety

Every tool resolves its path argument through an allowlist
(`VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS`, defaulting to the user's home
directory plus the Claude Code scratch root, `/private/tmp/claude-<uid>`,
when that directory exists) and a denylist of well-known credential
locations (`~/.ssh`, `~/.aws`, etc.) that is never overridable. Setting
`VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS` explicitly replaces the default list
verbatim rather than widening it. See `src/verified_docx_mcp/paths.py`.

## License

MIT — see [LICENSE](LICENSE).

## Approved live changes

Use `live_changes_prepare`, then one digest-bound `live_changes_execute` request for a live change set. Code handles sequencing, internal verification and durable resume without per-edit model calls. See [the execution contract and acceptance evidence](docs/acceptance/approved-changes.md). Requires the `approved_changes_v1` pane capability.

### Table structure edits

`insert_table_row` uses 1-based `table_id` and exactly one of
`before_row_index` / `after_row_index`, plus `cells` (one string per column).
`template_row_index` defaults to the adjacent row. `delete_table_row` requires
`row_index` and tab-joined `expected_text`; the final row cannot be deleted.
Both support file/live routing and revision guards. Merged/nested structures
are refused. Native nonadjacent templates require matching row/cell layout.

`insert_table(like_table_id=...)` inherits direct table, row, cell, paragraph
and first-run properties; explicit grid, style and font arguments override
inherited values. Copying immediately after a table adds a separator paragraph
to keep the two tables distinct. Refresh table IDs after structural edits.
