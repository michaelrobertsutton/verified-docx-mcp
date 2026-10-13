# #39 — list_open_items(source="live"): the OOXML completeness cross-check

## Automated

`tests/unit/test_live_comments_complete.py` (server side, against the fake pane's
configurable `ooxml` block: `complete` true / false with one retry / null, additive fields)
and the OOXML scenarios in `tests/unit/js/comments_list_harness.mjs` (wrapper:
`tests/unit/test_taskpane_js.py`; thread counting over a stub DOM).

Neither is Word. They do not prove that desktop Word's `body.getOoxml()` includes the
comments part and `commentsExtended.xml`, nor that `w14:paraId` / `w15:paraIdParent` / `w15:done`
reflect Word's own reply and resolved state as modelled.

## Real Word: PENDING

Needs a disposable document in desktop Word with the Live pane reloaded
(`taskpane.js?v=58`). Confirm `live_status` lists `comments_ooxml_crosscheck`.
Record the Word version and platform, and each result.

1. **Baseline.** With 3 comments (one resolved, one with a reply) in the body, note the
   pane's "Open comments N" counter, then call `list_open_items(source="live")`.
   Expect `complete: true`, `counts.ooxml_open` equal to the open threads (replies not
   counted, the resolved thread not counted), `counts.live_open` equal to the counter.
2. **Add while connected.** Add 5 more comments in Word by hand (body text, a table cell,
   a reply, one resolved after creation) without reloading the pane. Call
   `list_open_items` again and compare `counts.live_open`, `counts.ooxml_open`, the pane
   counter and `complete`.
3. **A gap, if you can make one.** Add a comment in a header, footer or text box. The body
   OOXML should not contain it: record whether `complete` stays `true` (expected, a known
   limit) and whether `file_has_unlisted_comments` fires after saving.
4. **Reproduce #39 if possible.** On the document that showed 22 vs 27, record the
   counter, `counts`, `complete` and any `pane_collection_incomplete` warning
   (`ooxml_open_ids`, `missing_comment_ids`, `retried`).
5. **Is the OOXML there at all?** If `complete` is `null` on a document with comments,
   the pane reply's `ooxml.reason` says why (comments part or `commentsExtended` part
   missing from `getOoxml()`); record it. A document with no comments may legitimately
   report `null`.
6. **Id spaces.** When a warning fires, compare `missing_comment_ids` / `ooxml_open_ids`
   (`w:id`) with the returned `live:<id>` handles. Record whether any ever overlap.

## Results

Run 2026-10-13, Word for Mac, branch `issue-44-39-shapes-comments`, pane `taskpane.js?v=58`,
scratch copy of `tests/fixtures/sections.docx`.

| Check | Result |
|---|---|
| Word's `body.getOoxml()` includes the comments parts (cross-check available) | PASS (`counts.ooxml_open` present) |
| Add 3 comments live: `live_total 3`, `live_open 3`, `ooxml_open 3`, `complete: true` | PASS |
| Reply to one and resolve another: `live_open 2`, `ooxml_open 2`, `complete: true` (replies not counted as threads, resolved excluded) | PASS |
| A gap between the OOXML count and the collection yields `complete: false` plus `pane_collection_incomplete` | NOT RUN in Word (the gap from #39 did not reproduce on a fresh document; covered by harness and fake-pane tests) |
| Comments added in desktop Word's own UI while the pane is connected | NOT RUN |

#39 stays open: the 22-versus-27 discrepancy needs the original co-authored document
(comments possibly outside the body collection) to confirm the cause.
