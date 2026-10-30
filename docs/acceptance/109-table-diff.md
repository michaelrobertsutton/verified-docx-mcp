# diff_table_vs_markdown (#109)

Read-only tool. It writes neither the docx nor the markdown, appends nothing to the audit
log, and is not in `MUTATING_TOOLS`.

## What was run

`tests/unit/test_table_diff.py`, against docx files built in the test from `frag.docx`'s
package with generated tables (so each case controls the exact cell XML), and the FakePane
pattern from `test_live_reads.py` for the live path.

- **Acceptance pair**: a 4x3 Word table against a markdown table where exactly three cells
  differ (`Austin`/`Dallas`, `Bob`/`Rob`, `QA`/`SRE`). The result lists exactly those three
  `{row, col, docx_text, md_text}` entries, no `dimensions`, `identical: false`.
- Identical tables report `identical: true`.
- Normalization: curly quotes, dashes, NBSP and whitespace fold; emphasis is ignored by default
  and compared as `**`/`*` with `strip_emphasis=false`; escaped pipes, links, inline code,
  hyperlink runs, tabs and line breaks. A literal `*x*` in Word differs from markdown `*x*`
  (emphasis, so `x`) and matches `\*x\*`.
- Input: inline vs file, missing file, outside the allowlist, denylisted path, directory.
- Every selector refusal (INVALID_INPUT): both or neither markdown input, both or neither of
  `table_id`/`section_key`, `md_header_row` with `section_key`, zero or several markdown tables,
  unknown or repeated `md_heading`, a heading with no table, conflicting `md_heading` and
  `md_header_row`; a `#` line inside a fenced code block is not a heading.
- Unsupported (never equal): a cell with a field, drawing, picture, text box or simple field,
  merged cells, a nested table.
- Extra/missing rows and columns appear as differences with `null` on the missing side, plus
  `dimensions`.
- Multi mode: pairing by header-row arrays regardless of order, per-table differences, unmatched
  tables on both sides, unsupported tables block `identical`, duplicate headers refused on either
  side, zero tables on either side refused, other sections' tables out of scope.
- Zero writes: docx and markdown bytes and mtimes unchanged, no stray files.
- Live (FakePane): `source="live"` reads the pane body and carries `source: "live"` and
  `live.body_sha256`; `auto` reads live; a `file` read with a session carries
  `live_session_ignored`; `source="live"` with no session is `LIVE_UNAVAILABLE`.

Gate: `uv run pytest tests/unit` (see the PR for pass counts).

## Design notes

- `md_table_index` is the 0-based position of the table among all tables in the markdown.
- With both `md_heading` and `md_header_row`, the header row must match exactly one table under
  that heading; with only `md_heading` the first table under it is used.
- Heading text is matched on normalized text without case folding.
- Inline HTML in a markdown cell other than `<br>` is compared as literal text and adds a warning.
- Character-style bold/italic (`w:rStyle`) is not resolved; only direct `w:b`/`w:i` run properties.

## Real Word acceptance

PASS (Word for Mac 16.114.3, disposable scratch fixture, live source through the default shared
bridge, run on the #109 branch's own package and pane build, taskpane `?v=66`).

1. Markdown identical to a 3x3 table whose cells include a two-paragraph cell and a hyperlink
   paragraph: `identical: true`, no differences.
2. Three cells edited in the markdown: exactly those three reported (`row`/`col`/`docx_text`/
   `md_text`), nothing else, `source: "live"`.
3. A 2x2 table compared with a 1-row markdown table: the missing row's cells reported with
   `md_text: null`, plus a `dimensions` entry.
4. A person's edit made in Word was reflected in the next live diff.

Not covered in Word: multi-table mode (`section_key` plus `md_heading`) and heading-scoped
selection, merged or nested tables reported as unsupported, tables with fields or text boxes,
`source="auto"` behaviour with a pane connected, and documents with real styles or co-authoring.
