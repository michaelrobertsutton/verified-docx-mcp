# #106 — row operations and faithful table copies

Tested 2026-10-29 with Microsoft Word for Mac 16.114.3 (build
16.114.26102720), existing pane v67, using
`/private/tmp/106-final-acceptance/106-final.docx`. Only the original add-in was
installed. Word stayed open during row insertion, deletion and table copying.

## Results

- Native row insertion and deletion passed independent text and direct row,
  cell, paragraph and run property verification. The fixture defines Compact
  paragraphs, 8.5pt runs, fixed layout, direct borders and cell margins.
- Native copies passed independent `tblPr`, `trPr`, `tcPr`, `pPr` and `rPr`
  verification, including an explicit 9pt override.
- Copying immediately after another table keeps separate tables. Inserting a
  standalone separator paragraph before import prevents Word's merge/split
  behavior. All three pre-existing tables retained their visible text and
  canonical direct properties in the final boundary test.
- Native Word may resolve an automatic table width into a concrete grid width
  when inserting a row. This host normalization does not change the grid or
  other table formatting; row/cell/paragraph/run properties are verified exactly.
- File tests cover formatting inheritance, round-trip insertion/deletion,
  expected-text refusals, invalid positions, final-row refusal, stale live
  revisions before dispatch, tracked row markers, and preservation of original
  Flat OPC namespace declarations during table import.

## Audit references

- Insert: `9f601dcc5dd6432ab0a55864232c4d79`
- Delete: `d8497628b21b458ea6017b8c279702ee`
- 9pt copy: `ad95608c923e4662946bb0c39d4677ea`
- Final after-table copy: `102ea983d4f74dbe9aa418ba29736d65`

Local audit root: `/private/tmp/verified-docx-acceptance-state`.
The failed copy experiments were inspected before recovery; uncertain writes
were never automatically replayed.

## Bounded native support

Merged/nested tables are refused. A nonadjacent live template with different
row/cell layout is refused before mutation; templates with matching layout can
supply different paragraph/run formatting. File mode clones direct row/cell
layout. Structural writes require clients to refresh positional table IDs.
