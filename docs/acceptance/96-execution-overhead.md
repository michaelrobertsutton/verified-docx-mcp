# Approved execution overhead (#96)

Pane v63 advertises `approved_cell_stream_v1`. The server streams one cell operation
at a time through `approved_cell_set`: each durable journal dispatch receives one
acknowledgment containing independently resolved target readback. There is no group
whose uncertain suffix could accidentally replay. The manifest remains immutable.

The pane performs comment/revision/anchor preflight, freshly resolves protected
section content/structure at the mutation boundary, writes exactly once, then
resolves a fresh scoped XML snapshot for per-target verification and complete shape
comparison. Section maps are reused only within that snapshot/context, never
across requests or human edits. Scoped XML is never a whole-body fallback.

Old approved panes retain cell_set + independent cell_get. Unsupported stream
structures can use that conservative path only after an explicit prewrite
LIVE_CAPABILITY_MISSING refusal. Every postwrite exception becomes uncertain
VERIFICATION_FAILED and cannot fall back or replay. Independent edits, bounded
prewrite retries, interrupted dispatch reconciliation and final save retain the
existing scheduler/journal semantics. A proven prewrite deletion refusal now
quarantines only its target; structural uncertainty still quarantines its section.

`metrics` in local per-operation evidence records elapsed milliseconds, native
sync counts and phase sync timings. Sync counts include incidental heartbeat syncs
while that pane operation is active. Backend evidence records wall time and bridge
request counts; compact reports aggregate them. There is no extra model read loop.

## Acceptance

PASS — Word 16.114 (16.114.26101123), macOS, 2026-10-16. Both disposable
12×4 runs completed 37/37 operations, with six actual outside-section edits and
an outside comment resolution interleaved through a second client. No issues,
no remaining operations, no model interventions between mutations.

| Measurement | Conservative, same host | Optimized |
|---|---:|---:|
| Elapsed seconds | 550.86 | 225.87 |
| Measured main-pane syncs, including prepare/final | 2604 | 832 |
| Cell mutation requests | 36 | 36 |
| Separate cell readback requests | 36 | 0 (independent readback in each acknowledgment) |
| Initial/final scoped snapshots | 2 | 2 |
| Initial comments lists | 1 | 1 |
| Save requests | 1 | 1 |
| describe/body_ooxml calls | 0 | 0 |

Elapsed reduction: **59.0%**. Measured sync reduction:
**68.0%**. Each run used one MCP execution request. Every mutation kept fresh
protected scope checks, comment/revision/shape protections and per-target
independent readback; save followed all mutation attempts.

Baseline run `c72f048818fc4e8487928a40566f0ade`; optimized run
`5168a488e5b84ec6899688e0f8cdc674`. These measurements use the same host and
instrumented pane version. The historical PR #95 baseline was 342.33 seconds;
225.87 is a 34.0% reduction against that older run. Host/pane load affects elapsed
time; the paired sync/request counts provide the structural comparison. This is
not a claim of 50% reduction against the historical measurement.

Of the optimized elapsed time, 142.51 seconds were measured inside the 36 native
cell operations. Roughly 79 seconds were collaborator operations/probe overhead;
those operations deliberately remain outside the optimized cell stream.

Reproduce using separate copies of tests/fixtures/approved-changes.docx opened
with the v63 pane at the paths in tests/live/approved_stream_probe.py. Run normally
and with --legacy in separate processes; keep the target fixture active in Word
and do not run another live test concurrently. Full evidence stays in
/private/tmp/word-execution-{optimized,baseline}-evidence.json and durable journals.

Automated parity covers both legacy and stream scheduling, unknown writes,
independent edits, retries, saved resumes and digest/document identity. Additional
harnesses verify the actual pane stream's fresh XML projection, target readback,
prewrite conflict and postwrite uncertainty without strategy retries.

## Follow-up priorities

1. Batch compatible preflight reads and use direct table-cell access to remove
   collection traversal syncs, retaining the fresh mutation-boundary checks.
2. Separate heartbeat liveness from expensive document hashing; never label a
   retained digest fresh or weaken the next mutation's preconditions.
3. Explore a bounded pane-side table stream with durable per-cell acknowledgments
   and explicit interruption reconciliation. Do not cache a protected section
   across collaborator edits or replay an uncertain suffix.

Track native executor timing separately from collaborator/probe timing. The latter
accounts for about 79 seconds of the optimized end-to-end fixture and cannot be
removed by counting only the approved cells.
