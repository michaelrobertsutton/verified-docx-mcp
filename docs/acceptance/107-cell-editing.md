# #107 — cell editing without format loss

Real Word acceptance is PENDING. Use the #106 scratch fixture (8.5pt direct
formatting), and a second 9pt table with two differently formatted paragraphs.

Check markdown replacement with default first-run name/size/color inheritance
and explicit `inherit_format="none"`; cell paragraph deletion (including refusing
the final paragraph); merging adjacent cell paragraphs while preserving both
runs and the first paragraph properties; paragraph insertion after tables and
filling immediately following empty paragraphs; per-edit row-scoped batches
with identical text in another table; and live saving during native typing.
Confirm untouched runs retain their sizes and no operation requires closing Word.

The shared bridge treats save as persistence, not a content mutation: it neither
requires nor refreshes a content baseline. Approved saves retain their protected
section validation. A save does not settle uncertain write outcomes.

Automated coverage: property-level file tests, pane paragraph deletion harness,
shared save-baseline regression, batch validation and existing replacement tests.
A tracked paragraph merge is supported only in live mode; file mode refuses it.
