# Issue 44: anchored shape acceptance

This PR is stacked on #54 (paragraph deletion), which is stacked on #53
(textbox scopes). Test and merge those dependencies before this PR. No Word
acceptance was run during implementation: the user was actively working in Word.

## Behavior

`list_shapes(path)` reads the connected live body and reports DrawingML floating
anchors and legacy VML shapes, their containing paragraph (text and 1-based
index over all body paragraphs, including tables/textboxes), and geometry.
DrawingML extents use EMUs and position records retain their relative origin;
VML geometry retains the original style string and units. AlternateContent
fallbacks are excluded when a choice exists. This is body coverage, excluding
headers, footers and inline pictures; it does not claim page-layout coordinates.

All live mutation dispatches inspect shapes before and after. Shape identities
and stored geometry must remain equal; counts are returned and included in audit
evidence. Paragraph indices may change after insertion/deletion. Read-back failure
reports VERIFICATION_FAILED with no rollback claim. Existing panes without the
shape_guard capability refuse writes until reloaded.

Text replacement conservatively refuses **any** touched paragraph with an anchored
shape, including partial replacements. Whole-cell and row replacements inspect
all target cell bodies before writing anything. Deletion reuses #54's paragraph
inspection, now reporting ANCHORED_SHAPES and covering additional VML shape types.
Formatting, insertion and comment edits remain permitted with shape read-back.
Revision acceptance/rejection is refused whenever body shapes exist because it
can remove anchors and safe re-anchoring is unsupported.

## Which ops refuse up front, which are verified afterwards

- **Refuse before writing** (`ANCHORED_SHAPES`): `replace_text` / `replace_text_batch`
  (any touched anchored paragraph), `replace_cell_markdown` (`cell_set`),
  `replace_table_row` (`cells_set`), `delete_paragraph` (`paragraph_delete`), and
  tracked-change accept/reject (`revisions_*`, whenever the body has shapes).
  `delete_table` refuses a table anchoring a shape.
- **Verified after the write** (shape identities and counts must be unchanged,
  otherwise `VERIFICATION_FAILED`): `format_text`, `insert_table`, `insert_paragraphs`,
  comment add/reply/resolve/delete, and `live_save`. Every guarded op also returns
  `shapes_before`/`shapes_after`.

Automated end-to-end coverage through the bridge (fake pane, not Word):
`tests/unit/test_live_shape_e2e.py` (`list_shapes`, counts in evidence and the audit
log, `ANCHORED_SHAPES` refusals with no write). The pane's JS is covered by
`tests/unit/js/shapes_harness.mjs`.

## Word checks required before merge

1. On a disposable document, create multiple floating textboxes and a geometric
   shape anchored to paragraphs; include a shape in a table cell and duplicate
   paragraph text. Compare list_shapes anchors, counts, sizes and position origins
   with Word. Check legacy VML/AlternateContent on a compatible fixture.
2. Attempt whole-paragraph and partial replace_text on an anchored paragraph:
   ANCHORED_SHAPES, no text/shape loss, tracking mode unchanged. An unanchored
   paragraph replacement succeeds with identical before/after shape counts.
3. Try replace_cell_markdown and replace_table_row where one target cell contains
   an anchored shape: refusal before any cell changes. Repeat with an unanchored
   target and independently inspect retained shape IDs, text and geometry.
4. delete_paragraph refuses anchored paragraphs and permits an eligible unanchored
   one; verify remaining anchors even though paragraph indices shift.
5. format_text, insert_table, comment creation/reply/resolve and live_save preserve
   shape identities/counts; verify counts in tool results and audit entries.
6. Tracked change acceptance/rejection refuses with body shapes; an old pane build
   refuses mutation with LIVE_CAPABILITY_MISSING before sending the operation.
7. Simulate missing shapes or a failed post-edit OOXML read: VERIFICATION_FAILED,
   no applied-success response, no claim that Word rolled back the edit.

## Word results (2026-10-13)

Word for Mac, branch `issue-44-39-shapes-comments`, pane `taskpane.js?v=58`, scratch copy of
`tests/fixtures/textbox.docx` (one DrawingML text box anchored to the second paragraph).

| Row | Result |
|---|---|
| 1. `list_shapes` reports the text box: id, name, `drawingml`, anchor paragraph 2 and text, size (2540000 x 1016000 EMU), horizontal/vertical position origins | PASS (single shape only; multiple shapes, table-cell shapes and VML not run) |
| 2. `replace_text` on the anchored paragraph: `ANCHORED_SHAPES`, nothing changed. `replace_text` on the unanchored paragraph applies with `shapes_before 1`, `shapes_after 1`, shape still listed afterwards | PASS |
| 3. Cell and row refusal with an anchored cell | NOT RUN |
| 4. `delete_paragraph` on the anchored paragraph refused with nothing deleted (it hit `STRUCTURAL_BOUNDARY` first because it is the last paragraph); unanchored deletion with shifted indices not run | PARTIAL |
| 5. Counts across `format_text`, `insert_table`, comment ops, `live_save`; audit entries | PARTIAL (counts verified on `replace_text`; audit counts covered by `test_live_shape_e2e.py`) |
| 6. Tracked accept/reject refusal with shapes; old-pane refusal | NOT RUN in Word (unit tests cover both) |
| 7. Simulated failed post-edit OOXML read | NOT RUN in Word (unit tests cover it) |

#44 stays open until rows 3 to 7 and the multi-shape row 1 variants are run.
