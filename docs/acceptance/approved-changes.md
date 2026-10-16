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

PENDING until a disposable document has run the same 12x4 table scenario with
v60 connected, outside-section edits and comment resolutions interleaved. Also
confirm a same-section edit refuses the next operation without overwriting it.
Record the run report, journal path, Word version and pane capabilities here.
Do not substitute simulated results for this acceptance check.
