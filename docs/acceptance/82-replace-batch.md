# #82 — replace_text_batch, allow_comment_loss_ids, LIVE_STALE_AFTER_WRITE, open-pane diagnostics

## Automated

`tests/unit/test_live_replace_batch.py` (server side, against the fake pane),
the `allow_comment_loss_ids` cases in `tests/unit/js/edit_harness.mjs` (pane op
under node), and the shared-mode serialization test in the same Python file.

The fake pane and the node stub are not Word. They prove the server's
decisions (hash chaining, outcome classification, error mapping, audit) and the
pane guard's logic; they do not prove the Office.js calls behave as modelled in
a real document.

## Real Word: PENDING

Needs a disposable scratch copy in Word for Mac with the Live pane reloaded
(`taskpane.js?v=56`). Record the Word version, `live_status` capabilities
(`comment_loss_ids` must be listed), each result and the audit record.

1. A 10-edit `replace_text_batch` (distinct `find` strings, `expected_matches=1`)
   lands: `batch_status: "complete"`, `applied: true`, every result `applied`,
   `revision_after` equals the next `live_status`/read revision, one
   `replace_text_batch` audit record.
2. Mid-batch refusal: make edit 5's `expected_matches` wrong. Result is
   `partial`, edits 1-4 applied in Word, edit 5 `refused_before_write`
   (`MATCH_COUNT_MISMATCH`), 6-10 `not_attempted`, `revision_after` is the
   state after edit 4.
3. Stale at start: type in Word after reading, then send the batch with the old
   `revision_before`. Result is `none_applied`, edit 1 `refused_before_write`
   (`LIVE_STALE`), the document is unchanged.
4. Comment-loss ids: a resolved thread with 2 replies overlapping `find`.
   `allow_comment_loss_ids=[{comment_id, reply_count: 2}]` replaces and reports
   `comments_removed`; the same call with `reply_count: 1`, with the thread
   unresolved, or with the id omitted refuses `WOULD_DELETE_COMMENTS` and
   changes nothing. Also confirm the `allow_comment_loss=True` +
   `allow_comment_loss_ids` combination refuses `INVALID_INPUT`.
5. Mid-batch pane loss: close the pane during a long batch. Result is `unknown`
   with `revision_after: null`; re-read and see how many edits landed.
6. `live_open_pane` with the screen locked or Word in the background: the error
   carries `reason`, `hint` and `retryable: true`; after unlocking, a retry
   opens the pane.
7. Timing: record wall-clock for the 10-edit batch against 10 sequential
   `replace_text` calls on the same document.

### Results

pending
