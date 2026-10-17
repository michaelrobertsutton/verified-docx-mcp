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

## Acceptance — PASS, Word 16.114/macOS, 2026-10-16

The disposable floating-callout fixture completed a mixed-format replacement
before the anchor, a text replacement after it, and save (3/3, no issues).
Shape identity, content, geometry, wrapping and complete definitions remained
unchanged. Allowed run: `313ef626ad87432fa1da2144e039b5b0`.

A second approved run refused paragraph deletion with ANCHORED_SHAPES before
writing, completed an independent paragraph edit, and saved (2/3, with only the
intentionally unsafe deletion remaining). Shapes again matched exactly.
Refusal run: `4e37e61447e246e79e2cc707a21354a9`.
Full evidence: `/private/tmp/word-callout-evidence.json` and durable journals.

Automated harnesses cover allowed text-only edits, changed geometry after write,
anchor-containing target refusal and paragraph-boundary refusal. Scheduler tests
cover independent continuation after an anchored deletion refusal on both
legacy and stream paths. Reproduce with `tests/live/anchored_text_probe.py
--create`, open the printed fixture and its Live pane, then run without --create.
