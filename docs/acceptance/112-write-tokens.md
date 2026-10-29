# Bounded write responses (#112)

Only response shaping changed. No verification, readback, comment/shape/revision
protection or journal semantics moved; every existing refusal test passes unchanged
(the few tests that read a removed field now assert the new shape).

## What changed

- **Whole-document evidence is bounded.** `append_markdown`, `replace_body_markdown`,
  `replace_range_markdown` and file-mode `accept`/`reject_tracked_changes` return
  `before_sha256`, `after_sha256`, `before_chars`, `after_chars`, `paragraphs_before`,
  `paragraphs_after` and a 200-character `before_excerpt`/`after_excerpt` at the first
  difference, instead of the whole text. The full text stays in the audit log; each record
  now has an `id`, returned as `evidence_ref` (only when the record was written), and
  `read_audit_evidence(evidence_ref, field?, to_file?)` fetches it. The log is
  `$XDG_STATE_HOME/verified-docx-mcp/audit.jsonl`, append-only and never rotated by the
  server; with `VERIFIED_DOCX_MCP_AUDIT_EXCERPTS=0` the text is never stored. The audit
  write happens before the response is shaped.
- **Live `delete_paragraph`** returns counts and the two neighbouring paragraphs. The pane
  (v66) advertises `compact_evidence_v1`; the server also trims whatever an older pane sends
  and keeps those lists in the audit record.
- **`format_text` / `apply_style`**: `text_unchanged: true`, one `excerpt`, and `format_delta`
  (only attributes that moved). `runs_before`/`runs_after` need `verbose=true`.
- **`replace_text_batch`**: `{applied, batch_status, applied_count, edit_count,
  revision_before, revision_after, results}`. An applied edit keeps `evidence:
  {match_count, revision_after}` in full; a failed one has `code` and `message`.
  `verbose=true` restores the previous shape.
- **Approved-run reports**: no `run_id`/`manifest_sha256` echo, `performance` only with
  `verbose`, `section_diff` cut to 600 characters (`section_diff_truncated`, `journal_path`),
  `remaining_ranges` beside the unchanged `remaining` list. `live_changes_prepare`'s preview
  shows find/replace/markdown/cells/anchor_text (200 characters, ellipsis), the table address and
  `expected_matches`; it is derived from the manifest, which stays the authority.
- **Lean envelope** (a middleware; the audit log still has everything): dropped when constant
  or already known: `verified_via`, `track_changes_author`, `document_name`, empty
  `orphaned_comment_ids`, `audit_logged`/`ledger_logged` when true, `conflict_copy_detected`
  when false, `write_mode` when requested explicitly, `revision_before` when equal to the one
  passed. A single-match `add_anchored_comment` returns `comment_id` only. False and
  non-empty values stay.
- **Errors**: `MATCH_COUNT_MISMATCH`/`STRUCTURAL_BOUNDARY` list at most five matches as
  `{start, end, context}` plus `total_matches`; `ZERO_MATCH` keeps `near_miss` and drops
  `ladder_report` unless `verbose`; `LIVE_STALE*` carries 12-hex `expected`/`actual`;
  `retryable` is omitted when false; `LIVE_UNAVAILABLE` is the code, `document_name` and a
  one-clause hint.

`verbose` never restores whole-document before/after text: that is what `evidence_ref` is for.

## JennyStack

`live_doc_stamp.py` / `live_doc_guard.py` read, from write results: `applied`, `match_count`,
`revision_after`, `status`, and for `replace_text_batch` the top-level `batch_status`,
`applied_count`, `edit_count`, `applied`, `revision_before`, `revision_after` and each applied
result's `evidence.match_count` / `evidence.revision_after` (all full tokens). All of those are
unchanged. They read `write_mode` and `revision_before` from the tool *arguments*, not the result.
`approved_run_guard.py` takes `run_id`/`manifest_sha256` from arguments and the durable journal, and
the `remaining` list is untouched. The batch keeps per-edit `revision_after` only because the hook
chains it, which is the floor on its size; JennyStack #446 asks for a hook that reads the top-level
`revision_after` alone, after which the server can drop the per-edit evidence.

## Method

Same as #110: `scripts/make_large_fixture.py` (150 sections, 300 paragraphs, ~20K words) and
`scripts/measure_write_tokens.py`, which calls the real FastMCP server in-process so the shaped
result is what is measured. "Before" ran the same script against a checkout of `main` before
this change. Tokens are characters / 4 (an estimate).

## Results (file mode, 20K-word document)

| Scenario | Before (tokens) | After (tokens) |
|---|---:|---:|
| `append_markdown`, one paragraph | 98,286 | 175 |
| `replace_body_markdown`, small replacement | 49,177 | 152 |
| `replace_range_markdown`, one section | 394 | 156 |
| `format_text`, one phrase | 335 | 158 |
| `replace_text`, one edit | 341 | 318 |
| `MATCH_COUNT_MISMATCH`, a string with 150 hits | 1,127 | 144 (5 contexts, `total_matches: 150`) |
| `ZERO_MATCH` | 79 | 30 |

Fake-pane and unit measurements:

| Scenario | Before (tokens) | After (tokens) |
|---|---:|---:|
| `replace_text_batch`, 50 edits, success | 3,159 | 2,030 |
| `live_changes_execute`, 12 prewrite conflicts | about 12,000 (estimated from the old 4000-character cap, 12 diffs; not re-run) | 2,449 |
| live `delete_paragraph`, 304-paragraph section | every paragraph twice | under 500 (`test_live_delete_evidence.py`) |

The 50-edit batch misses the issue's 1.5K target. Each applied edit must keep a full 76-character
`revision_after` (157 characters, about 40 tokens, per edit) because the JennyStack hook validates
and chains it; the agreed fix is on the JennyStack side. `replace_text` still returns its text
before/after and `runs_before`/`runs_after`: those belong to #108's `verbose` work.

## Real Word

PASS, 13 of 13 checks, Word for Mac on 2026-10-29, `tests/live/write_tokens_probe.py` against a
pristine disposable copy of the large fixture (300 paragraphs), pane v66.

| Check | Result |
|---|---|
| live `delete_paragraph` of a heading in the 300-paragraph body | 264 tokens: `paragraphs_before/after` 300/299, both neighbouring paragraphs, `evidence_ref` resolves to the audit record, paragraph gone |
| 50-edit `replace_text_batch` | all 50 applied, every hook field intact (`applied`, `batch_status`, `applied_count`, `edit_count`, `revision_before`, full `revision_after`, each edit's `evidence`), 2,043 tokens, **134 s** (about 2.7 s per edit in Word) |
| a write chained from the batch's `revision_after` | accepted |
| mid-batch refusal | `partial`; failing index 1, `ZERO_MATCH`, its message; the next edit `not_attempted` |
| `LIVE_STALE` through the MCP layer | `expected`/`actual` 12-hex, no `retryable` |
| `MATCH_COUNT_MISMATCH` from the live pane | 35 tokens, diagnostics `pane_code` only |
| `ZERO_MATCH` | no `ladder_report` |
| a live `replace_text` result | none of `verified_via`, `track_changes_author`, `document_name`, `audit_logged`, `write_mode`, `orphaned_comment_ids`; 185 tokens |
| live `format_text` | `text_unchanged: true`, one `excerpt`, 104 tokens |

Two limits the run showed. A match-count error raised by the live pane carries only `pane_code`:
the pane reports counts, not offsets, so the five-context list applies to file-mode and `locate`
errors. And `compact_evidence_v1` was exercised against the current pane only; the older-pane path
(server trims the full lists) is covered by the unit test, not by real Word.

## Not done

`LIVE_STALE` also reporting `changed_sections`: #110's digests now behave on real Word, but the
error path would need a second pane round trip on every stale refusal, and nothing yet reads it.
Left out rather than shipped speculative.
