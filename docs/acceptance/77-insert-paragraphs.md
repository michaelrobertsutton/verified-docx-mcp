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

Result: _not yet run._
