# Approved cell preflight and heartbeat overhead (#101 / #102)

## Implementation and compatibility

Pane v64 keeps `approved_cell_stream_v1` and the immutable manifest/tool contract.
Within one request/context it loads table items/nesting together, uses WordApi 1.3
`Table.getCell` after exact scoped table resolution, and batches the initial
comment/revision/anchor/tracking reads. The final exact-target text load shares
its sync with the fresh section-location check. Protected content/structure and
section identity are still freshly resolved after all asynchronous preflight.
There is no cross-request resolution cache. Native missing-target/proxy errors
from the read-only preflight are typed prewrite stale refusals; mutation failures
remain uncertain. One write attempt and independent fresh XML readback remain.

`heartbeat_liveness_v1` is negotiated in a `hello_ack` only when the pane advertises
it. Old panes retain their hash heartbeat; new panes connected to old servers keep
the legacy wire format. Negotiated liveness is sent every five seconds without a
Word request. An idle hash is coalesced through the same queue as native operations,
at most once per 30 seconds while idle. Hashing never overlaps or bypasses a queued
mutation; reconnect hello reads also wait for that queue. Busy periods defer samples
until idle rather than treating an old sample as current content.

A negotiated heartbeat carries `digestState: unavailable` without a hash, or
`digestState: sampled` with a diagnostic sample. `live_status` adds
`body_hash_state`, `body_hash_age_s`, `heartbeat_count`, and `hash_sample_count`.
Sampled means a past observation, never a fresh client read. A hashless heartbeat
updates liveness only; it leaves sample age unchanged and marks the digest retained
or unavailable. Writes invalidate sample state. Older brokers with no age return
an unavailable age, not an invented fresh timestamp. No heartbeat modifies broker
client read baselines, revision counters, section locks, or approval receipts.

The disconnect threshold remains three missed five-second heartbeats. Negotiated
expired sessions close their socket so suspended panes perform a fresh hello on
resume; pending operations become unknown and never replay. Legacy eviction
behavior remains compatible. Timing metrics use the pane's monotonic performance
clock, and body hashing is outside operation metrics. Heartbeat/hash request counts
are collected separately from native preflight/mutation/verification counts.

## Reproduction

Use copies of `tests/fixtures/approved-changes.docx` only, named
`/private/tmp/word-101-{baseline,baseline-three,optimized,optimized-three}-1.docx`.
For the three-pane cases also open two disposable helper copies. Use the same Word
host, keep the target fixture active, and do not run another live mutation probe
concurrently. `BENCH_LABEL=<label> python tests/live/overhead_probe.py` runs the
one-request 12×4 collaborative scenario. `python tests/live/heartbeat_overhead_probe.py
<label> 35` samples a separate idle/busy heartbeat window. Set
`VERIFIED_DOCX_MCP_ALLOWED_FILE_ROOTS=/private/tmp` for the mutation probe.

The baseline assets preserve the final #96 algorithms from commit d3236cb. Both
asset sets use the v64 cache key; the only baseline instrumentation adjustment
excludes `currentBodyHash` from request metrics, so background syncs are counted
separately on both sides. Detailed evidence is in `/private/tmp/word-101-*` and the
durable per-run journals. Historical #95/#96 elapsed times are not substituted for
these paired host measurements.

## Results

PASS — Word 16.114 (16.114.26101123), macOS. Baselines collected October 16,
2026; final monotonic-timed candidate runs October 17 on the same host. Both
variants use the same approved stream, fixture, measurement wrapper and v64
cache key; the baseline preserves #96's preflight/heartbeat behavior.

| Measurement | Baseline, 1 pane | Candidate, 1 pane | Baseline, 3 panes | Candidate, 3 panes |
|---|---:|---:|---:|---:|

| Full collaborative elapsed seconds | 139.76 | 102.55 | 153.89 | 103.71 |
| Native cell executor seconds | 88.74 | 54.92 | 97.77 | 55.08 |
| Measured request syncs, including prepare/final | 727 | 475 | 727 | 475 |
| Cell preflight syncs | 540 | 288 | 540 | 288 |
| Cell mutation syncs | 36 | 36 | 36 | 36 |
| Cell independent verification syncs | 108 | 108 | 108 | 108 |

| Native sync time across 36 cells (seconds) | Baseline, 1 pane | Candidate, 1 pane | Baseline, 3 panes | Candidate, 3 panes |
|---|---:|---:|---:|---:|
| preflight | 66.80 | 35.57 | 75.38 | 35.33 |
| mutation | 4.71 | 4.31 | 4.83 | 4.19 |
| verification | 15.49 | 13.51 | 15.86 | 13.99 |

All four runs completed 37/37 with six interleaved unrelated edits and outside
comment resolution, one MCP execute request, 36 approved cell requests, one final
save, one initial comment list and two initial/final scoped snapshots. There were
no separate cell_get, describe or body_ooxml requests and no model interventions.
Per-cell preflight fell **15 → 8 syncs (46.7%)**; mutation and independent
verification counts stayed **1 and 3**. Total measured request syncs fell
**727 → 475 (34.7%)**. End-to-end elapsed fell **26.6%** with one pane and
**32.6%** with three panes. Native cell time is measured separately above; the
remaining time includes collaborator/probe work and prepare/final/save overhead.
These measured results do not promise a fixed speedup under different host load.

| 35-second heartbeat window | Baseline hash requests/native hash syncs | Candidate | Candidate maximum observed heartbeat age |
|---|---:|---:|---:|
| one pane(s), idle | 7 | 1 | 4.92s |
| one pane(s), busy | 7 | 0 | 4.95s |
| three pane(s), idle | 21 | 3 | 5.00s |
| three pane(s), busy | 21 | 2 | 5.00s |

Idle background sync reduction is **85.7%** for both pane counts. In busy windows,
the active pane sampled no hashes; only the two idle helpers sampled (one each).
The tradeoff is diagnostic content samples up to 30 seconds old while idle,
or older during long busy periods. Their age/state is explicit; writes and
explicit reads never rely on this sampling schedule. Baseline maximum observed
heartbeat ages were 4.96–5.06 seconds; the candidate remained 4.92–5.00 seconds.

An initial candidate timing run was discarded after a wall-clock jump produced
an impossible native elapsed sample. Its raw evidence is retained locally as
`word-101-discarded-wallclock-evidence.json`. All final candidate phase durations
use `performance.now()`; a deterministic clock regression verifies the choice.

Run IDs: baseline `6f3b0e234d2e488faa088f0a70d8dd11`, baseline-three
`206b78eaf8894c2681abb8b1feb33df9`, candidate
`298e9ebcfbed4033a914913925d7a120`, candidate-three
`2951c4e11c9e4c8e9dc0f9d9eb472675`.

## Safety and JennyStack evidence

`overhead_safety_probe.py` passed on real Word: empty, repeated and multiline
cells; a tracked cell write; prewrite TRACKED_CHANGES_PRESENT and
WOULD_DELETE_COMMENTS refusals with independent eligible edits and final save.
Actual native Word edits bypassed broker counters during both idle and busy
periods. Both edits were preserved, the two independent targets finished and
save ran (3/5, remaining only the two deliberately conflicted targets). Evidence:
`/private/tmp/word-101-safety-evidence.json`; human-edit run
`ab7e9351d58a487aa09440ca9b0e9484`.

Actual pane harnesses and backend regressions cover fresh protected structure and
section identity, callout anchors, enclosing revisions, merged/nested structures,
legacy capability fallback, queued heartbeat work, late/displaced heartbeats,
postwrite uncertainty, no strategy retries/replay, saved resumes and process
ownership takeover. A broker regression proves background samples cannot update
client read baselines; an old-broker regression requires an unknown digest age.

JennyStack needs no production changes. Its 11 executor/approval tests and 55 Word
integration tests pass. An additional explicit metrics test confirmed extra
performance evidence does not change authorization or verified completion. The
existing unchanged-manifest resume and partial/unknown outcome tests pass.

The complete backend suite passed **1,135 tests and 297 subtests**. Subsequent
focused pane checks passed 43 tests/120 subtests, and the final stream/freshness
checks (including the added old-broker case) passed 16 tests. Changed production
Python and new probes pass Ruff; JavaScript parses; mypy retains the same 41
pre-existing errors in 8 files.

## Real pane lifecycle

The three disposable panes were suspended for 18.18 seconds by stopping their
Word WebKit processes (all other Word documents were excluded). The registry
removed them at heartbeat ages **15.07–15.09 seconds**, preserving the existing
15-second threshold. All three automatically reconnected with fresh connection
generations **1.43 seconds after resume**. No document mutation was dispatched
by this test. The probe resumes every stopped process in finally. Reproduction:
`tests/live/pane_lifecycle_probe.py`; evidence:
`/private/tmp/word-101-lifecycle-evidence.json`.

Final backend protocol/session/broker checks passed **133 tests**, including
stale baselines, negotiated/legacy peers, expiry, reconnect and ownership recovery.
