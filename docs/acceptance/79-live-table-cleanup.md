# #79 — live table clean-up: numbering, clear_numbering, delete_table, scoped accept/reject

## Automated

`tests/unit/test_live_table_cleanup.py` (server side, against the fake pane) and
the "#79" scenarios in `tests/unit/js/tables_harness.mjs` (pane ops under node,
mock Word object model that queues writes and enforces load-then-sync).

The fake pane and the node mock are not Word. They prove the server's decisions
(capability gates, verification, error mapping) and the pane's call sequencing;
they do not prove that `Paragraph.detachFromList`, `Table.delete`, or
`Range.getTrackedChanges().rejectAll()` behave as modelled in a real document.

## Real Word: PENDING

Needs a disposable scratch copy in Word for Mac with the Live pane reloaded
(`taskpane.js?v=57`). Confirm `live_status` lists `table_delete`,
`paragraphs_unlist` and `revisions_scoped`. Record the Word version, each
result and the audit record.

1. **Numbering on insert.** In a scratch document make a numbered-list
   paragraph, then `insert_table` anchored after it
   (`anchor={"paragraph_text": "<that paragraph>"}`, a 2x2 table). Every cell is
   un-numbered in Word, the result reports `list_items_in_table: 0`, and the
   numbered paragraph and the rest of the list are unchanged.
2. **clear_numbering on existing content.** Make a table whose cells are
   numbered (apply a list to its cell paragraphs by hand). `clear_numbering(table_id=N)`
   returns `unlisted` equal to the number of numbered cell paragraphs and the
   cells show no numbering. Then `clear_numbering(paragraph_text=...)` on one
   numbered body paragraph, and on a paragraph text that occurs twice
   (`MATCH_COUNT_MISMATCH`, nothing changed). Run it again on the cleaned table:
   `unlisted: 0`.
3. **delete_table, match.** `get_table` a small table, join its cells with tabs
   and rows with newlines, pass that as `expected_text`. The table disappears,
   `tables_before - tables_after == 1`, neighbouring paragraphs and other tables
   are unchanged.
4. **delete_table, mismatch.** Same call with one character of `expected_text`
   changed: refused `CELL_TEXT_MISMATCH`, table still present. Also try a table
   that holds a comment (`WOULD_DELETE_COMMENTS`) and a table with a tracked
   change inside (`TRACKED_CHANGES_PRESENT`); nothing deleted either way.
5. **delete_table, tracked.** `track_changes=true`: the table shows as deleted in
   the Review pane and stays until accepted; the call still reports success.
6. **Scoped reject keeps a colleague's change and the anchor formatting.**
   Give the anchor paragraph purple italic run formatting, and have a second
   author (or Word with a different user) leave a tracked insertion in another
   paragraph. `insert_table(track_changes=true)` after the anchor, then
   `reject_tracked_changes(table_id=<new table>)`. The table is gone
   (`table_removed: true`), the colleague's tracked insertion is still pending,
   and the anchor paragraph's paragraph and run formatting are identical
   (compare against a copy made before the insert).
7. **Scoped accept.** With tracked edits inside table A and a tracked insertion
   elsewhere, `accept_tracked_changes(table_id=A)` accepts A's changes only.
8. **Refusals.** `table_id` together with `revision_ids` is `INVALID_INPUT`;
   `table_id` with `write_mode="file"` is `LIVE_UNAVAILABLE`; with a section lock
   held by another client it is `LOCKED_BY_OTHER_CLIENT`; against a pane that
   has not been reloaded (no `revisions_scoped`) it is `LIVE_CAPABILITY_MISSING`
   and the document is unchanged.

### Results

Run 2026-10-13, Word for Mac, branch `issue-79-live-table-cleanup`, pane from the branch
bridge (`taskpane.js?v=57`), generated scratch docx with a 2-item numbered list.

| Check | Result |
|---|---|
| `insert_table` anchored after numbered list item "Item two": `applied`, `list_items_in_table: 0` | PASS |
| `clear_numbering(paragraph_text="Item one")`: `unlisted: 1`; "Item two" renumbered to "1." | PASS |
| `delete_table` with wrong `expected_text`: refused `CELL_TEXT_MISMATCH`, nothing deleted | PASS |
| `delete_table` with the right text (`A1\tB1\nA2\tB2`): table removed, neighbours intact | PASS |
| Tracked `replace_text` elsewhere, tracked `insert_table`, then `reject_tracked_changes(table_id=1)`: `table_removed: true`; the other tracked edit ("Third plain CHANGED": ins, del, rPrChange) still pending | PASS |
| After the scoped reject, no revision residue from the table | **PARTIAL**: two empty paragraph-mark revisions (one `ins`, one `del`, text `""`) remain at the anchor paragraph ("Plain paragraph"). The visible text matches the original. |
| Anchor paragraph with run formatting (purple italic) survives tracked insert then scoped reject | NOT RUN (the scratch anchor was unformatted); covered by harness assertions only |
| `clear_numbering(table_id=...)` on an existing numbered table | NOT RUN |

Known limitation: a scoped reject of a tracked table insertion can leave empty
paragraph-mark revisions at the anchor. They are not the table's content and can be
cleared with the document-wide reject once no one else's revisions are pending, or by
hand. #79 stays open for this and for the anchor-formatting row.
