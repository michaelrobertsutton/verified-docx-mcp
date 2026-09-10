# verified-docx-mcp

An MCP server for local `.docx` files with verified writes. It is the
`docx` counterpart to
[`verified-googledocs-mcp`](https://github.com/michaelrobertsutton/verified-googledocs-mcp):
the same evidence-first contract (every mutating tool re-reads the file
after writing and returns before/after evidence, never a bare "success"),
applied to a local Word document on a synced drive (OneDrive, iCloud Drive)
instead of a Google Doc.

## Status

Early scaffold (issue #28 WP-02 of the build plan). Reading and writing a
`.docx` package needs no Word installation at all — this server manipulates
OOXML directly. Word is required only for `export_pdf` (rendering a page-
accurate PDF), and that additionally needs a macOS Automation grant for the
app hosting this server's process.

Tools implemented so far:

- **`export_pdf(path, output_path)`** — render a `.docx` to PDF via
  Microsoft Word automation and report its page count.
- **`lock_status(path)`** — report Word/LibreOffice owner-file presence
  and sync-quiesce state, as data only. Never refuses.

More tools (document reads, markdown-based writes, tables, comments,
tracked changes) land in later work packages of the same plan.

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
