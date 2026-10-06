# Live cell paragraphs and line breaks (#75)

Status: implemented; **pending real-Word acceptance**. The JS harness
(`tests/unit/js/tables_harness.mjs`) models Word's two quirks (`insertBreak`
accepts only Before/After; the `insertParagraph` proxy in a table cell resolves
to the previous paragraph) and fails on the pre-fix pane code, but real-Word
behavior of `paragraphs.getLast()`, `Range.insertBreak(After)` and the
`insertOoxml` restore is not verified until the checklist below is run.

## Root causes

- `Paragraph.insertBreak(line, "End")` is `InvalidArgument` (Before/After only).
- Writing paragraph 2+ through `body.insertParagraph("", "End")`'s returned
  proxy merged its text into paragraph 1 and left the real new paragraph empty.
- The server's whitespace-collapsing read-back could not see paragraph
  boundaries, so the bad write was reported but never undone.

## Real-Word checklist

Word for Mac, pane connected (reload it so `live_status` shows `cell_multiline`),
a table with a header row, `table_id=1`. Record the Word build and outcome of
each step.

1. Two paragraphs: the issue's call, `replace_cell_markdown(row_index=4,
   cell_index=2, write_mode="live", markdown="**Session 2: ...** Build a long
   list of AI use cases.\n\n*Pre-read: SOO Obj 4.3, 5.3*")`. Expect two
   paragraphs, no trailing empty paragraph, success.
2. Hard line break: two trailing spaces then a newline. Expect one paragraph
   with a line break, success (no `InvalidArgument`).
3. Three paragraphs, with bold across a line break.
4. 0-based index slip: call with `expected_before` set to the text you meant to
   replace but the index one off. Expect `CELL_TEXT_MISMATCH`, nothing written.
5. Forced mismatch (to see the restore): temporarily make the pane's write
   disagree with the intent and confirm the cell returns to its old text with
   `diagnostics.rolled_back: true` and no extra empty paragraph.
6. Same with `track_changes=true`: confirm the pane does not restore and the
   message points at the Review pane.

If step 1 or 2 still fails in real Word, switch to the fallback: refuse
multi-paragraph/line-break markdown server-side with `INVALID_INPUT` and say so
in the docs rather than ship an unproven write path.

## Not in scope

A live `replace_range_markdown` (replacing a body paragraph with a bulleted
list) remains a follow-up.
