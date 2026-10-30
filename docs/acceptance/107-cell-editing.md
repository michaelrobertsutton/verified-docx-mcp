# Cell edits without format loss (#107)

Real-Word acceptance was run on **Word for Mac 16.114.3**, pane `taskpane.js?v=67e`
(this branch), through the default shared-mode bridge, on a disposable scratch fixture.
Results are below. Two of the original design's mechanisms did not survive real Word and
were replaced; the fake-pane tests could not have shown either.

## What changed

| # | Change | Where |
|---|--------|-------|
| 1 | `replace_cell_markdown` inherits the cell's first-run font (name, size, colour; file mode also the complex-script size and paragraph marks). `inherit_format="none"` opts out. Applies to ordinary and approved cell writes. | `tables.py`, `live/approved_changes.py`, pane `opCellSet` / `opApprovedCellSet` / `opCellGet` |
| 2 | `delete_paragraph` may delete a non-final paragraph inside a (non-nested) table cell. | pane `opParagraphDelete` |
| 3 | New `merge_paragraph_with_next`: joins two adjacent paragraphs of one cell, keeping runs and hyperlinks. `replace_text` still refuses a match across a paragraph boundary. | `paragraph_merge.py`, pane `opCellParagraphMerge` |
| 4 | `insert_paragraphs` anchors `{"after_table_id": N}` and `{"empty_paragraph_after_table_id": N}`. | `paragraph_insert.py`, pane `opParagraphInsertPositional` |
| 5 | `replace_text_batch` forwards each edit's `within_row_containing`. | `text_edit.py` |
| 6 | `live_save` needs no body baseline and does not acknowledge a read or stale other clients. | `live/shared.py` (`BASELINE_EXEMPT_OPS`), pane `dispatchOp` |

## Real-Word results

| Check | Result |
|-------|--------|
| Pane loads this build; capabilities `cell_font_inheritance_v1`, `cell_paragraph_merge_v1`, `paragraph_positions_v1` present | PASS |
| `replace_cell_markdown` in an 8.5pt Arial table and a 9pt Calibri table | PASS: size, name and colour kept, each read from its own cell |
| **The bug itself:** a two-paragraph write with `inherit_format="none"` | **Reproduced**: the runs came back with no size or font (style default). Same write with the default `first_run`: both paragraphs 8.5pt Arial |
| Cells no step edited | PASS: text and font unchanged; saved file has no run at a default size |
| `delete_paragraph`: a non-final cell paragraph | PASS |
| `delete_paragraph`: a cell's final paragraph, or its only one | PASS (refused before writing; see "Found in real Word") |
| `merge_paragraph_with_next` with a hyperlink and mixed run formats | PASS: one paragraph, hyperlink intact, every run 8.5pt Arial, verified by the server |
| Merge while Word's Track Changes is on; `track_changes=true` | PASS (refused, nothing written) |
| Merge of two paragraphs with different properties | PASS (refused before writing; see below) |
| `insert_paragraphs` `after_table_id`; `empty_paragraph_after_table_id`; fill of a non-empty paragraph | PASS (inserted; filled; refused) |
| `replace_text_batch` with identical text in two tables, one `within_row_containing` per edit | PASS: both edits applied (the original `MATCH_COUNT_MISMATCH` case) |
| `live_save` through the shared broker | PASS (see "Not covered") |

## Found in real Word (the fake pane could not show these)

- **`cell.body.getOoxml()` returns the whole table, not the cell.** The first design read a
  cell's first run from it; the "passes" were coincidence (it read the table's first cell).
  The font is now read from the cell itself: the font at the start of its first paragraph.
  This is the font Word *reports* (resolved with the style), so mixed values and black
  (Automatic) are left out, and a link keeps its own colour.
- **`range.insertOoxml` into a cell range throws `GeneralException`.** The merge now deletes
  the paragraph mark between the two paragraphs through the object model: no OOXML is
  imported, so runs and hyperlinks stay where they are.
- **Word keeps the SECOND paragraph's properties when it joins two** (a centred second
  paragraph made the merged paragraph centred). A live merge of a pair whose properties differ
  is therefore refused before writing; file mode keeps the first's.
- **A cell's final paragraph cannot be removed**: `delete()` blanks it and leaves an empty
  paragraph, so a delete would land and then fail read-back. It is refused before writing.
- A paragraph's own `getOoxml` does not canonicalize like the body export, so an XML-equality
  stale guard refused every merge; the guard compares text (the body-hash check covers the rest).
- Word caches `taskpane.js?v=N` in its web view: bump `?v=` for every pane change you want
  Word to load (this branch ended at `v=67e` for that reason).

## Design decisions worth reviewing

- **One definition of "first run" in file mode** (`tables._inherited_cell_font`: the first run
  with text, looking through hyperlinks; direct `rFonts`/`sz`/`szCs`/`color`; the paragraph
  mark only for a cell with no text). Live mode applies the font Word reports (above).
- **Fail closed on old panes.** A live `first_run` write needs `cell_font_inheritance_v1`: without
  it the server cannot see the font and would silently reset it. The error names the opt-out,
  `inherit_format="none"`. This also holds inside approved runs.
- **Save is a write for authorization, not for freshness.** `save` still honours the document
  lease and capability gates. It is exempt only from the `seen != version` stale check,
  `expectedBodySha256` injection and baseline bookkeeping, in the broker AND in the pane's
  `dispatchOp` (an explicitly supplied hash is still checked). An independent review reproduced
  that the broker change alone left every plain shared-mode save refused by the pane. A plain
  save does not run the section guard (it edits nothing); only approved saves do.
  `live_save`'s `revision_after` and `file_revision` are read at separate instants and are not
  proof that the saved bytes are that body.
- **Merge verification** compares the cell's paragraph texts and count, the merged paragraph's
  `pPr`, every hyperlink's resolved target and text, the cell's run formats (the set of
  size/font/weight combinations, since Word splits and merges runs) and every other cell of the
  table. Reads on both sides of a live write are non-observing. A document that moved after the
  write adds a `warnings` entry rather than failing a merge whose content verified.
- **Tracked merges are refused** in both modes.
- **"Empty paragraph" has one policy** (`paragraph_insert.EMPTY_PARAGRAPH_FORBIDDEN_TAGS`):
  file mode applies it to the parsed body and sends it to the pane. A fill keeps the paragraph
  (and its properties) in place in both modes.
- **Approved runs** pay one extra `cell_get` per `first_run` cell write (the write itself stays a
  single dispatch). A manifest prepared before this change now runs with `first_run` implicitly;
  `inherit_format` is bound into the digest only when given.
- Pre-write refusals the new ops raise (`STRUCTURAL_BOUNDARY`, `TRACKED_CHANGES_PRESENT`,
  `TABLE_NOT_FOUND`, `CELL_TEXT_MISMATCH`, `INVALID_INPUT`) no longer stale other clients.
- Not done here (other repo): JennyStack guard/permit/revision registration of
  `merge_paragraph_with_next` and the new `insert_paragraphs` anchors; canonical Word workflow
  instructions.

## Not covered (still to verify)

- `live_save` **while a second person is typing**: the shared-mode baseline refusal is
  fixed and exercised against the real pane code, but a concurrent human edit was not
  simulated against real Word.
- The approved-run font write (`approved_cell_set`) in real Word: its font read-back reads the
  just-written proxy, which Word may invalidate.
- Filling an empty paragraph that holds a bookmark, field or other forbidden content, a table
  inside a content control, and an empty or style-only fill text (live mode treats an
  unchanged body as a failed insert).
- Complex-script size (live mode cannot set it) and list marks in live cells.

## Automated coverage

`tests/unit/test_cell_editing_107.py` (file mode, a real table with 8.5pt direct formatting;
asserts untouched runs and cells are byte-identical), `test_live_107.py` (fake pane),
`test_shared_bridge.py` (save semantics), `test_approved_changes.py` / `test_approved_stream.py`,
`test_live_cells.py`, `js/cell_font_harness.mjs` (including the real pane's `dispatchOp` for a
shared-mode save) and `js/paragraph_harness.mjs`.
