# Changelog

## Unreleased

- `diff_body_vs_file` is normalized (quotes, dashes, NBSP, soft hyphens, NFC, trailing whitespace), returns hunks or a unified diff but never both, drops the text of equal hunks, and can be scoped by `section_key` and capped by `max_hunks`. `identical_raw` keeps the byte comparison. On the ~20K-word fixture an identical-document diff falls from ~99K to ~100 tokens.
- `read_document` gains inclusive `start_para_ref`/`end_para_ref` ranges (any format, file or live) and `to_file`, which writes the body to an allowed path and returns a summary with the file's sha256. `find_sections` entries carry `chars`; `find_sections(digests=true)` adds a per-section `sha12`.
- New `live_revision` tool returns the live revision and per-section digests in one pane round trip (pane op `body_hash`, capability `body_hash_v1`; pane v65). Older panes fall back to `body_ooxml`. `live_list_sections(compact=true)` returns `[key, heading, sha12]` rows.

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
