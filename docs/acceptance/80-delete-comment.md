# #80 — delete_comment / delete_resolved_comments

## Automated (run 2026-10-12)

`uv run pytest tests/unit -q`: all pass. That covers the file-mode tests
(`tests/unit/test_delete_comments.py`, including the Word-authored golden
fixture), the live tests against the fake pane (`tests/unit/test_live_comments.py`),
and the pane op under node (`tests/unit/js/comment_delete_harness.mjs`).

The fake pane and the node stub are not Word. They prove the server's
decisions (reconciliation, drift detection, error mapping) and that the pane op
loads `resolved` before reading it; they do not prove `Comment.delete()`
behaves as documented in a real document.

## Real Word: PENDING

Not run. Needs a disposable scratch copy in Word for Mac with the Live pane
reloaded (`taskpane.js?v=55`). Record the Word version, `live_status`
capabilities, evidence and audit record for each step.

1. `live_status` lists `comment_delete`.
2. Resolve a comment, then `delete_comment` it: it leaves Word's margin;
   `list_open_items(source="live")` shows `live_total` down by one and every other
   thread (content, replies, resolved state) unchanged.
3. A thread with two replies: `delete_comment` removes root and replies.
4. `delete_resolved_comments` with three resolved threads and one open:
   refuses with the open one listed and deletes nothing; without the open one,
   deletes all three in one call.
5. A reply handle refuses with `INVALID_INPUT`; an open comment refuses in the batch.
6. Save the document (`live_save`) and confirm the saved file has no orphan
   range markers for the deleted ids (file-mode `list_open_items` agrees).
7. File mode on a closed copy of a Word-authored document: delete a resolved
   thread, open the result in Word, and confirm it opens without a repair prompt
   and the other comments are intact.
