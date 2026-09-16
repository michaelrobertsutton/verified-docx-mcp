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
  and sync-quiesce state, as data only. Never refuses.
- **`list_parts(path)`**, **`read_document(path, format, part)`**,
  **`find_sections(path, part)`**, **`list_page_sections(path, part)`**,
  **`list_styles(path)`** — read a `.docx` part as markdown/text/runs,
  and enumerate its heading-delimited sections, page-layout sections, and
  styles. Never refuse on a locked file (a validated snapshot is read
  instead).
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
  `w:rPr` verbatim onto every surviving piece.
- **`list_open_items(path)`**, **`accept_tracked_changes(path,
  revision_ids?, ...)`**, **`reject_tracked_changes(path, revision_ids?,
  ...)`** — list open comments/pending tracked changes (Google's response
  shape), and accept/reject `w:ins`/`w:del` revisions by id or all at
  once.
- **`track_changes: bool`** on every mutating tool above — when true,
  the edit is wrapped in `w:ins`/`w:del` (or a `w:rPrChange` for a style
  change) with an author/date/id, rather than applied directly. A
  revision authored by the server's own configured identity never blocks
  a later tracked write; one authored by anyone else does, unless
  `force=True`.
- **`add_anchored_comment(path, quote, text, expected_matches, ...)`**,
  **`get_comment_thread(path, comment_id)`**, **`reply_to_comment(path,
  comment_id, text)`**, **`resolve_comment(path, comment_id)`** — create
  an anchored comment, read a comment and its replies (by durable id),
  reply to one, and resolve one. Builds all five interlocking comment
  parts a real Word comment needs; verified part-by-part against a
  Word-authored golden fixture (`tests/fixtures/comments/golden-comment.docx`).
  A multi-paragraph comment's identity is keyed on its LAST paragraph,
  matching Word's own commentsIds.xml/commentsExtended.xml convention, and
  every `comment_id` `list_open_items` reports (durableId or, lacking one,
  the raw `w:id`) is accepted by the three tools above.
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
- **`diff_body_vs_file(path, file_path)`** — export the docx body's
  markdown projection (the same one `read_document(format="markdown")`
  uses) and diff it against a local markdown file with `difflib`, the
  same mechanism GoogleDocs-MCP's `diff_tab_vs_file` uses. Read-only and
  affirmative-only: `identical: true` means no difference was found by
  this particular projection, not a guarantee none exists (see the
  tool's own docstring for the docx-specific reasons why, including a
  trailing newline on the file side alone surfacing as a spurious
  one-line difference).
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
  cant_split=False, anchor=None, ...)`** — insert a new table, one
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
  `format_text` (including `track_changes=True`); a paragraph style
  applies `w:pStyle` to every paragraph containing a matched run, but
  does not support `track_changes=True` (no `w:pPrChange`-style tracked
  change exists yet — named explicitly rather than silently ignored).
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

## Live mode (WP-1 spike)

A hello-world Word task-pane add-in (`addin/`) plus a local HTTPS bridge
(`python -m verified_docx_mcp.live.bridge --serve-only`) that a lead can
sideload into Word for Mac to check `WordApi 1.4` support and correlate
Office.js comment ids against this server's OOXML `durableId`s. This is
the WP-1 spike for
[issue #106](https://github.com/michaelrobertsutton/JennyStack/issues/106)'s
live co-editing bridge — throwaway scaffolding, not yet wired into the
MCP server itself (importing `verified_docx_mcp.live` never starts
anything). Full step-by-step: [`docs/live-mode.md`](docs/live-mode.md).

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
