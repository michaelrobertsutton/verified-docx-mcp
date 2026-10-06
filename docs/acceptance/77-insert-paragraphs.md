# #77 — live paragraph insertion (real-Word acceptance)

Run against a disposable scratch copy in real Word (never a colleague's document),
with the Live pane reloaded (`taskpane.js?v=54`). Record the Word version, the
`live_status` capabilities, revision tokens and the audit evidence for each step.

1. `live_status` lists `paragraph_insert` in the pane's capabilities.
2. Tracked insert: `insert_paragraphs(anchor=<unique paragraph text>,
   paragraphs=[Heading 2 heading, purple 7030A0 body, green 00B050 body],
   expected_matches=1, track_changes=True)`. Word shows three tracked insertions
   with the right styles and colors in the Review pane; the evidence lists each
   inserted paragraph's text, style and color. Reject all: the document returns
   to its original text.
3. Untracked insert after an anchor that is itself a heading: the new body
   paragraph without a `style` is Normal, not the heading style.
4. `position="before"`: the new paragraphs land above the anchor, in order.
5. A duplicate anchor refuses with `MATCH_COUNT_MISMATCH` and an absent one with
   `ZERO_MATCH`, with the document unchanged.
6. An unknown style refuses with `STYLE_NOT_FOUND` before anything is written;
   an anchor inside a table cell refuses with `STRUCTURAL_BOUNDARY`.
7. A stale `revision_before` refuses with `LIVE_STALE` before any write.
8. The change-tracking mode is back to its previous value after every path above.

## Result (2026-10-06, Word for Mac 16.113.3)

Run by a driver script against disposable copies of `sections.docx` and
`tables.docx` (pane opened through `live_open_pane`'s own code; bridge started from
this branch, pane loaded `taskpane.js?v=54`). All steps PASS:

1. Pane capabilities include `paragraph_insert`.
2. Tracked insert (Heading 2 + purple `7030A0` + green `00B050`) after
   "Some overview text.": the pane read back `Heading 2` / `Normal` / `Normal` and
   `#7030A0` / `#00B050`; Word listed the tracked insertions; the change-tracking
   mode was `false` afterwards; reject-all returned the body hash to the original.
   In a saved copy (`live_save`) the markup had six `w:ins` elements, `Heading2`
   `w:pStyle` and both colors.
3. An unstyled paragraph inserted after the heading "Overview" came back `Normal`,
   not the heading's style.
4. `position="before"`: saved file order is "Background", the new paragraph,
   "Background text.".
5. A duplicate anchor (`MATCH_COUNT_MISMATCH`, found 3) and an absent anchor
   (`ZERO_MATCH`) left the body hash unchanged.
6. An unknown style (`STYLE_NOT_FOUND`) and a table-cell anchor
   (`STRUCTURAL_BOUNDARY`, on `tables.docx`) left the body hash unchanged.
7. A stale `revision_before` refused with `LIVE_STALE`, body unchanged.
8. The tracking mode was `false` after every path.

Not exercised in real Word: the rollback of an untracked write that fails read-back
(mock harness only) and section locks.
