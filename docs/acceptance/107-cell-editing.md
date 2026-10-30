# Cell edits without format loss (#107)

**Real Word acceptance: PENDING.** Everything below marked "automated" is a fake-pane or
file test; none of it proves Word's own font, range or OOXML-import behaviour. #107 stays
open until the checklist at the end passes in real Word (record the Word build, pane
version and audit references in this file).

## What changed

| # | Change | Where |
|---|--------|-------|
| 1 | `replace_cell_markdown` inherits the cell's first-run font (name, size, colour; file mode also the complex-script size). `inherit_format="none"` opts out. Applies to ordinary and approved cell writes. | `tables.py`, `live/approved_changes.py`, pane `opCellSet` / `opApprovedCellSet` |
| 2 | `delete_paragraph` may delete a paragraph inside a (non-nested) table cell, never the cell's last one. | pane `opParagraphDelete` |
| 3 | New `merge_paragraph_with_next`: joins two adjacent paragraphs of one cell, keeping the first paragraph's properties and both paragraphs' runs and hyperlinks. `replace_text` still refuses a match across a paragraph boundary. | `paragraph_merge.py`, pane `opCellParagraphMerge` |
| 4 | `insert_paragraphs` anchors `{"after_table_id": N}` and `{"empty_paragraph_after_table_id": N}`. | `paragraph_insert.py`, pane `opParagraphInsertPositional` |
| 5 | `replace_text_batch` forwards each edit's `within_row_containing`. | `text_edit.py` |
| 6 | `live_save` needs no body baseline, does not acknowledge a read, and does not stale other clients. | `live/shared.py` (`BASELINE_EXEMPT_OPS`) |

## Design decisions worth reviewing

- **One definition of "first run".** The server derives it (`tables._inherited_cell_font`: the
  first run with text, looking through hyperlinks, else the first paragraph mark's run
  properties; direct `rFonts`/`sz`/`szCs`/`color` only). Live mode reads the cell OOXML
  through `cell_get(include_ooxml=true)` and sends explicit values to the pane, so file and
  live modes cannot disagree. Markdown carries no font, so every rendered run inherits; a
  property the markdown sets itself is never overridden.
- **Fail closed on old panes.** A live `first_run` write needs `cell_font_inheritance_v1`:
  without it the server cannot see the font and would silently reset it (the reported bug). The
  error names the opt-out, `inherit_format="none"`. This also holds inside approved runs.
- **Live font readback.** Office.js cannot set the complex-script size, so a live write
  applies and verifies name, size and colour only.
- **Save is a write for authorization, not for freshness.** `save` still honours the document
  lease, section locks (`forbiddenSections`) and capability gates. It is exempt only from the
  `seen != version` stale check, `expectedBodySha256` injection and baseline bookkeeping (the
  version bump, `seen`, `hashes`). The candidate's alternative (`save` in `READ_OPS`) would
  have skipped the lease and section-lock checks. `live_save`'s `revision_after` (a describe)
  and `file_revision` (the disk token) are read at separate instants: neither proves the
  saved bytes are that body.
- **Merge verification** compares the cell's paragraph texts and count, the merged paragraph's
  `pPr`, every hyperlink's resolved target and text, and every other cell of the table. Live
  reads on both sides of the write are non-observing (`read_source(observe=False)`), so a failed
  verification cannot have already re-armed the client. A document that moved after the write
  adds a `warnings` entry rather than failing a merge whose content verified.
- **Tracked merges are refused** in both modes: how Word records a paragraph-mark deletion
  through `insertOoxml` is unproven.
- **"Empty paragraph" has one policy** (`paragraph_insert.EMPTY_PARAGRAPH_FORBIDDEN_TAGS`):
  file mode applies it to the parsed body and sends it to the pane, which applies the same list.
  A fill keeps the paragraph (and its properties) in place in both modes.
- **Approved runs** pay one extra `cell_get` per cell write for the font read (the write
  itself stays a single dispatch).
- Not done here (other repo): JennyStack guard/permit/revision registration of
  `merge_paragraph_with_next` and the new `insert_paragraphs` anchors; canonical Word workflow
  instructions.

## Automated coverage

`tests/unit/test_cell_editing_107.py` (file mode, a real table with 8.5pt direct formatting;
asserts untouched runs and cells are byte-identical), `test_live_107.py` (fake pane),
`test_shared_bridge.py` (save semantics), `test_approved_changes.py` / `test_approved_stream.py`,
`test_live_cells.py`, `js/cell_font_harness.mjs` and `js/paragraph_harness.mjs`.

## Real-Word checklist (PENDING)

Use the 8.5pt/9pt directly formatted tables from the session that reported #107.

1. `replace_cell_markdown` in an 8.5pt cell and a 9pt cell: size, name and colour unchanged;
   neighbouring cells unchanged; `inherit_format="none"` gives the style font.
2. Empty cell and a hyperlink-first cell: behaviour as described above.
3. `delete_paragraph` on one of a cell's two paragraphs; on a cell's only paragraph (refused);
   in a nested table (refused).
4. `merge_paragraph_with_next` on two paragraphs with different `pPr`, a hyperlink and mixed
   run formatting: first `pPr` kept, hyperlink target intact, paragraph count down by one. Check
   whether Word's `getOoxml` for a cell paragraph canonicalizes the same way the server's
   `before_xml` does (a mismatch refuses with `LIVE_STALE`, harmlessly, but would make the tool
   unusable).
5. `insert_paragraphs` after a table, and into an empty paragraph after a table; a non-empty
   one is refused; a section-locked destination is refused.
6. `replace_text_batch` with the same text in two tables, one `within_row_containing` per edit.
7. `live_save` while a second person types in the document.
8. For each: confirm no untouched run changed size, nothing required closing Word, and
   record the audit lines.

Word build: _pending_. Pane version: _pending_. Audit references: _pending_.
