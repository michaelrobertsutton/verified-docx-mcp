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
