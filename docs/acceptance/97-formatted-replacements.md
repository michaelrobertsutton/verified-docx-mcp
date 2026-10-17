# Explicit formatted replacements (#97)

Pane v61 advertises `formatted_replacements_v1`. `replace_text` and approved
change-set operations accept `replacement_segments`. Each segment has text,
complete format, and optional hyperlink (null means no link); all text concatenates
exactly to replace. An empty list means deletion. There is no implicit style
inheritance. Legacy plain inputs retain their behavior.

Format requires bold, italic, strikeThrough, doubleStrikeThrough, subscript,
superscript (booleans), underline (None/Single/Double), color (#RRGGBB), name,
and size (points in half-point increments). Links accept HTTP(S)/mailto.
The preview and immutable manifest contain the exact segments and an explicit
formatting policy. Font or destination changes require a new approval digest.

Word inserts native text ranges, then independently reloads fonts, hyperlinks and
text. A scoped paragraph OOXML character ledger verifies unchanged prefix/suffix
text, run properties and link destinations. Word's extra empty exported paragraph
is treated as scaffolding only after the actual paragraph collection verifies one
paragraph. Target identity, comments, revisions and shape guards remain active.
Postwrite failures are uncertain and never automatically replayed. Text alone is
insufficient for reconciling formatted outcomes.

## Live evidence — PASS, Word 16.114, macOS, 2026-10-16

The synthetic matrix completed eight replacements plus save (9/9), with no issues:
bold/plain, italic, new hyperlinks, Unicode, empty deletion, within/across-run
spans and a soft line break. Independent font/link/range readback and unchanged
prefix/suffix ledgers passed. Run `6e4fd9582f8842b38c981bb4aa9d6986`, digest
`5cae08ef5e6f34abe2a7028975a11d0da38eabcc5ddaaa3bb336bd9bd4f4a629`.
Full evidence is local at `/private/tmp/word-formatted-matrix-evidence.json` and
in the durable manifest/journal directory. Reproduce using
`tests/live/formatted_replacements_probe.py --create`, open the printed fixture
and its Live pane, then run without --create.

A preliminary failed link readback quarantined only its target while independent
operations completed. The failure exposed adjacent-range hyperlink inheritance;
links are now assigned after all text exists, then fonts are explicitly applied.


## Repeated targets, comments and tracked changes — PASS

`tests/live/formatted_semantics_probe.py` added an identical string in another
section and verified that only the scoped target changed. A comment-overlapping
replacement refused with WOULD_DELETE_COMMENTS while an independent formatted
replacement and save completed. A tracked formatted replacement completed and
preserved its pending revisions; a subsequent intersecting replacement refused
with TRACKED_CHANGES_PRESENT while an independent target and save completed.

Runs: `410aa74ece3c48ffb32878c67fca8280` (scoped repeat),
`00fded273e3f4f308e353092d7404b93` (comment refusal),
`d44f7add9d084332a28aa2ca5118a9ac` (tracked formatted replacement),
`fa27b534de8442319bf76371373a99cb` (revision refusal and independent completion).
Full evidence: `/private/tmp/word-formatted-semantics-evidence.json`.


An additional preservation probe edited the plain body next to the existing Link
hyperlink. The scoped unchanged-character ledger verified the original destination
and formatting stayed intact. Independent before/after comment and tracked-revision
inventories also matched, including the OOXML revision count. Run
`451da097aaf44dc6979730539e08c60d`, evidence
`/private/tmp/word-format-preservation-evidence.json`. Reproduce with
`tests/live/formatted_preservation_probe.py` after the semantics probe.
