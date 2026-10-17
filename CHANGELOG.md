# Changelog

## 0.1.2 — 2026-10-16

- Batch approved cell preflight reads and use direct cell access, retaining fresh mutation-boundary checks and one-attempt independent verification.
- Negotiate transport-only heartbeats, queue idle hash samples separately, expose sample age/state, and reconnect expired panes without replaying unknown writes.
- Separate hash work from native operation sync metrics and use monotonic phase timers.

- Add explicit approved font/link segments for mixed inline replacements, independent formatting and unchanged-text verification.
- Allow text-only edits beside callout anchors while preserving complete shape definitions and refusing structural edits. Keep independent targets eligible after a proven prewrite deletion refusal.
- Add a capability-gated sequential cell stream with fresh scoped OOXML checks, one write attempt, independent per-target readback and durable per-operation acknowledgments.
- Record native sync counts, phase timings and bridge request counts in local execution evidence. Older panes and unsupported stream structures retain the conservative path after proven prewrite capability refusal.

## 0.1.1 — 2026-10-16

- Add digest-bound prepare/execute/resume for scoped live change sets, durable journals and explicit unknown-outcome reconciliation.
- Permit unrelated coauthor edits while preserving protected-section and target checks at the pane boundary.
- Attempt independent edits before retrying pre-write refusals; quarantine uncertain targets, save after all mutations, and resume final verification without replay.
- Include tables between adjacent top-level headings in protected section ranges; give scoped writes and readbacks a consistent verification deadline while retaining explicit timeout overrides.
- Keep full execution evidence local; return compact run reports.
