# Approved live change sets (JennyStack #421 / #423)

`live_changes_prepare(path, operations, track_changes=false)` captures protected
sections, table coordinates/structure/content, and the initial comment list in
an owner-readable local manifest. Each operation has exactly `op`, `section`
(unique heading slug), and `args`. Supported argument sets:

| op | args |
|---|---|
| replace_cell_markdown | table_id, row_index, cell_index, markdown |
| replace_table_row | table_id, row_index, cells |
| replace_text | find, replace, expected_matches (1) |
| delete_paragraph | anchor_text (the entire paragraph) |
| live_save | empty object, last operation only |

Indices are 1-based. Text replacements are single-paragraph, exact, and at most
230 characters. Structural edits and heading changes need a separate plan.
Preparation returns a compact preview, manifest path, run_id and SHA-256 digest;
review exact before/after changes in the local manifest before approving.

`live_changes_execute(path, run_id, manifest_sha256, approved_by)` executes the
immutable manifest, persisting a journal before dispatch and after each verified
operation. `live_changes_resume(path, run_id, manifest_sha256)` uses its stored
approval, skips verified operations, and attempts independent remaining targets. One document lock serializes
runs; separate document paths have separate manifests and journals. Completion
returns counts, remaining indices and exceptions; detailed evidence stays local.
Runs are not transactions and never roll back their verified edits.

The new pane capability is `approved_changes_v1` (reload taskpane.js v60).
The shared broker permits scoped writes despite unrelated clients' revisions,
but keeps document/section ownership restrictions. The pane independently checks
all protected section text, table structure, positions and target membership
before mutation. Old whole-body revision checking is unchanged for individual
calls. Cell/row compare-and-set, independent readbacks, comment/revision checks
and shape guards still run. This is an optimistic guard, not a Word transaction.

A proven pre-write stale/lock refusal retries at most twice after scoped and
target validation, after other independent edits have been attempted. Unknown
outcomes never retry. They quarantine overlapping targets and section runs until `live_changes_reconcile` records an explicit
lead disposition (`applied`/`not-applied`, decision_ref) supported by exact live
target state. Crashes during dispatch are uncertain, not permission to replay.
Conflicted or uncertain targets and their dependencies are deferred while other
edits proceed. Save runs after all mutation attempts, including retries. A final
read disconnect records verification_pending; resume retries the check without
replaying verified writes. Resuming does not expand approval.

## Automated evidence

`tests/unit/test_approved_changes.py` runs the real server cell/row mutators
against a scoped pane double. The 36-cell-plus-save case asserts two scoped
snapshots (prepare/final), one comment list, zero body_ooxml reads, no model
intervention, and no replay on duplicate execution. Additional tests cover
repeated/empty/multiline cells, same-section conflicts, best-effort scheduling and resumable remaining edits,
manifest/document identity, crashes, uncertain results, bounded retries and
explicit reconciliation. The section-guard JavaScript harness executes the pane
scope checks, including unrelated edits, protected edits and wrong targets.

## Live Word acceptance

PASS — 2026-10-16, Word 16.114 (16.114.26101123), macOS, pane v60
with `approved_changes_v1`.

The disposable 12x4 fixture completed 36 cell edits plus save through **one MCP
execution request**, with six actual unrelated section edits and an outside
comment resolution interleaved by a second client. No model intervened between
operations. The report was `complete`, 37/37 applied, no issues or remaining
operations. Elapsed time was 342.33 seconds, including collaborator operations.

| Main-client calls | Count |
|---|---:|
| MCP prepare / execute | 1 / 1 |
| execution_snapshot (initial / final) | 2 total |
| comments_list | 1 |
| cell_set / independent cell_get | 36 / 36 |
| save | 1 |
| describe / body_ooxml | 0 / 0 |

Clean run: `517306f16f1f4a1c9843b948abe1b48e`, manifest SHA-256
`cf692eba841f575896529523f8a63abb88687ce2605468a1864fef21b27da6b8`.
Its manifest, journal and full evidence remain in the local
`verified-docx-mcp/approved-changes/<document-sha>/` state directory.

A second live test changed an approved target cell before execution. The executor
returned a focused `CELL_TEXT_MISMATCH` conflict, preserved the collaborator's
content, applied the independent cell and saved. Its report was correctly
`partial` (2/3 operations verified, only the conflicted cell remaining).
Run: `8dc11bf8ed714e9094fe94ebea0bdbf1`.

A preliminary run also exercised two late replies under the legacy 15-second
deadline: all independent writes proceeded; exact target readback reconciled
the late writes without repeating them. Scoped broker, socket and independent
readback deadlines now agree at 60 seconds by default, with explicit timeout
overrides retained.

Reproduce with `tests/fixtures/approved-changes.docx`, copied to
`/private/tmp/jennystack-word-acceptance.docx` and opened with the Live pane.
Run `tests/live/approved_changes_probe.py` with
`VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp`. For the conflict probe, copy
the fixture separately to `/private/tmp/jennystack-word-conflict-final.docx`,
open its pane, then run `tests/live/approved_conflict_probe.py`.

Native round-trip reduction, mixed-format text replacements, and safe edits in
callout-anchored paragraphs are tracked separately in backend issues #96–#98.
