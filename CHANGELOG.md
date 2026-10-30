# Changelog

## Unreleased

- Cell edits keep their formatting (#107, real-Word acceptance pending). `replace_cell_markdown` inherits the cell's first-run font name, size and colour (`inherit_format="none"` opts out; a live write on a pane without `cell_font_inheritance_v1` is refused rather than resetting the font). `delete_paragraph` may remove a paragraph inside a cell, never its last. New `merge_paragraph_with_next` joins two paragraphs of one cell keeping the first paragraph's properties and both paragraphs' runs and hyperlinks (tracked merges refused). `insert_paragraphs` accepts `{"after_table_id": N}` and `{"empty_paragraph_after_table_id": N}`. `replace_text_batch` honours `within_row_containing` per edit. `live_save` no longer fails `LIVE_STALE` while someone types: it needs no body baseline and does not acknowledge a read, but still honours the lease and section locks. Pane capabilities `cell_font_inheritance_v1`, `cell_paragraph_merge_v1`, `paragraph_positions_v1`.
- Bounded write evidence (#112): `append_markdown`, `replace_body_markdown`, `replace_range_markdown` and file-mode accept/reject return hashes, sizes and a 200-character excerpt instead of the whole before/after (`append_markdown` of one paragraph to a 20K-word document: ~98K to ~175 tokens). The full text stays in the audit log under `evidence_ref`; new `read_audit_evidence` fetches it. `verbose` never restores it.
- Live `delete_paragraph` returns counts and the two neighbouring paragraphs (pane v66, `compact_evidence_v1`). `format_text`/`apply_style` report `text_unchanged`, one excerpt and a `format_delta`; `runs_before`/`runs_after` need `verbose=true`.
- Compact `replace_text_batch` and approved-run reports (no echo of `run_id`/`manifest_sha256`, `performance` behind `verbose`, `section_diff` cut to 600 characters, `remaining_ranges` beside `remaining`); `live_changes_prepare` previews each operation's text. A lean envelope drops constants (`verified_via`, `document_name`, `audit_logged: true`, ...). Errors list at most five match contexts with `total_matches`, drop `ladder_report` unless `verbose`, use 12-hex stale hashes and omit `retryable` when false.

- `diff_body_vs_file` is normalized (quotes, dashes, NBSP, soft hyphens, NFC, trailing whitespace), returns hunks or a unified diff but never both, drops the text of equal hunks, and can be scoped by `section_key` and capped by `max_hunks`. `identical_raw` keeps the byte comparison. On the ~20K-word fixture an identical-document diff falls from ~99K to ~100 tokens.
- `read_document` gains inclusive `start_para_ref`/`end_para_ref` ranges (any format, file or live) and `to_file`, which writes the body to an allowed path and returns a summary with the file's sha256. `find_sections` entries carry `chars`; `find_sections(digests=true)` adds a per-section `sha12`.
- New `live_revision` tool returns the live revision and per-section digests in one pane round trip (pane op `body_hash`, capability `body_hash_v1`; pane v65). Older panes fall back to `body_ooxml`. `live_list_sections(compact=true)` returns `[key, heading, sha12]` rows.
- File-mode `list_open_items` now threads comment replies via `w15:paraIdParent`: replies are no longer listed as top-level comments, and each root reports a real `reply_count` and `replies` (`{id, author, date, text}`). The root comment's resolved state governs the thread. Shared threading walk lives in `projection.comment_thread_links`.

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
