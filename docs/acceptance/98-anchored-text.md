# Text-only edits beside callout anchors (#98)

Pane v62 advertises `anchored_text_v1`. Replace now checks the actual target's
OOXML separately from its containing paragraph. A single-paragraph replacement
that excludes the anchor can proceed; paragraph marks, anchor-containing spans,
and unidentifiable shapes refuse before writing with ANCHORED_SHAPES.

The mutation uses native text ranges, never paragraph or shape OOXML rewrites.
The same paragraph proxy is read afterward and every complete shape definition
is compared, including shape identity/text, dimensions, position, wrapping and
anchor metadata. Existing whole-document/protected-section shape inventory is
also retained and strengthened to include complete definitions. Any discrepancy
is VERIFICATION_FAILED (uncertain), never a retryable prewrite refusal.

Paragraph/table deletion, cell reconstruction and revision operations retain
their conservative shape guards. Formatting segments and all comment/revision
and scoped concurrency protections remain active.

Automated harnesses cover allowed text-only edits, changed geometry after write,
anchor-containing target refusal and paragraph-boundary refusal. Real Word
acceptance is pending: create the disposable fixture with
`tests/live/anchored_text_probe.py --create`, open it and its Live pane, then run
without --create. Do not merge until allowed and refused scenarios both pass.
