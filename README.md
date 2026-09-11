# verified-docx-mcp

An MCP server for local `.docx` files with verified writes. It is the
`docx` counterpart to
[`verified-googledocs-mcp`](https://github.com/michaelrobertsutton/verified-googledocs-mcp):
the same evidence-first contract (every mutating tool re-reads the file
after writing and returns before/after evidence, never a bare "success"),
applied to a local Word document on a synced drive (OneDrive, iCloud Drive)
instead of a Google Doc.

## Status

Issue #28 of the build plan, through WP-09 (WP-10's lock-guard layers 3-4
land in the same PR, alongside this). Reading and writing a `.docx`
package needs no Word installation at all — this server manipulates
OOXML directly. Word is required only for `export_pdf` (rendering a
page-accurate PDF), and that additionally needs a macOS Automation grant
for the app hosting this server's process.

Tools implemented so far:

- **`export_pdf(path, output_path)`** — render a `.docx` to PDF via
  Microsoft Word automation and report its page count.
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

More tools (tables, images) land in later work packages of the same plan.

## Interoperability

Every read/write path above is exercised by `tests/unit` against
Word-authored fixtures and, for comments specifically, compared part by
part against a golden `.docx` produced by the lead in Word desktop. What
those tests cannot cover is a **live round trip through a real synced
folder with more than one client editing concurrently** — that requires
a human at a keyboard, in both Word desktop and Word Online, on a file
actually going through OneDrive/iCloud sync. Two such checks are
definition-of-done items for this PR, not agent gates, and their results
belong here once the lead has run them:

- **Word desktop + Word Online round trip (issue #28 WP-09):** the lead
  leaves one comment in Word desktop and one in Word Online on a file in
  a real synced folder. This server's tools then read, reply to, and
  resolve both comments. The lead reopens the file in both clients and
  confirms the threading and resolved state show correctly.

  **Result:** _not yet run — pending the lead's live round trip. This
  placeholder is intentional: no result is recorded here until the lead
  has actually performed the check above and reports what they observed
  in both clients._

- **Forced conflict copy (issue #28 WP-10):** the lead edits the same
  file in Word Online and locally within the sync window, forcing a real
  conflict-copy sibling, and records the filename the sync client
  actually produced alongside whether this server's `conflict_copy_detected`
  evidence flag fired for it.

  **Result:** _not yet run — pending the lead's forced conflict. This
  placeholder is intentional for the same reason as above._

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

## Path safety

Every tool resolves its path argument through an allowlist
(`VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS`, defaulting to the user's home
directory) and a denylist of well-known credential locations
(`~/.ssh`, `~/.aws`, etc.) that is never overridable. See
`src/verified_docx_mcp/paths.py`.

## License

MIT — see [LICENSE](LICENSE).
