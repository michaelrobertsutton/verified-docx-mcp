# Trimmed read payloads (#111)

No write-tool behaviour changed. The trimming is response shaping only. A lean
read envelope in `LeanEnvelopeMiddleware` never touches error results, and the
`list_open_items`/`list_tables`/`list_styles`/status changes keep every field
JennyStack reads (`tests/unit/test_jennystack_contract.py`).

## What changed

- **`list_open_items`**
  - `fields="summary"` returns columnar rows: `comment_columns`/`comment_rows` and `suggestion_columns`/`suggestion_rows`, plus `counts`, `path`, `complete`, `scope` and warnings. It drops `pane`, and never emits `comments`, `suggestions` or `pending_suggestions` (see Review).
  - A scoped response carries `scoped: true`.
  - Scoping uses `section_key` and/or `within_text`. Live scoping reads the body without re-arming the caller's shared-bridge write baseline. A row it cannot place is kept and flagged `scope_uncertain`.
  - `full` drops fields that are constant or derivable:
    - `modified_time` and `scope`
    - live `w_id` and `quoted_text`
    - `reply_count` and empty `replies`
    - per-comment `session_epoch`
    - null or empty top-level keys

    It also sends the file-mode `anchor_context` once per paragraph, in `anchor_contexts`, and makes `correlation` opt-in.
- **`list_tables`**: each table carries `header_row` and `anchor_section_key`.
- **`list_styles`**
  - Gains `type` and `name_contains` filters.
  - `STYLE_NOT_FOUND` now lists up to five `near_matches` with an `available_count`, instead of every style; with no near match it adds `available_sample` (the first 20 names, sorted).
- **`live_status`**
  - `path` scopes the sessions to one document.
  - Heartbeat, boot and shared-broker diagnostics now need `verbose=true`. `locks` (and this client's `client_id`) stay in the default output in shared mode.
- **`lock_status`**: returns `sync_quiesced` plus `external_activity {divergent, still_current, recent, ledger}` by default; the raw samples need `verbose=true`.
- **Read envelope**: the middleware drops the following:
  - `path` when it equals the input. `read_document` and `list_open_items` keep it, because JennyStack reads it.
  - empty `section_key`, `lossy_elements` and `warnings`
  - `revision_detail` unless `verbose`
  - `live.body_sha256`

  It also shortens `live_session` to `{document_url, reason}`.
- **Annotations and instructions**
  - `readOnlyHint` and `idempotentHint` are set on the 21 read tools (including #109's `diff_table_vs_markdown`, merged meanwhile); `destructiveHint` is set on the 8 destructive ones (including #106's `delete_table_row`).
  - The server instructions gain five workflow lines.
  - The README gains a permission rule for the read-only tools.
- **Fix found on the way.** `resolve_comment` and `reply_to_comment` refused the `live:<epoch>:<id>` handle that `list_open_items` returns: they sent `<epoch>:<id>` to the pane, which has no such comment.
  - The epoch now travels as `session_epoch`, so the pane still refuses a handle from an earlier epoch.
  - The #117 real-Word probe found this (pane activity: `refused comment_resolve ... no comment with id "<epoch>:<id>"`).
  - Two fake-pane tests cover it.

## Method

All numbers are compact-JSON characters, with tokens estimated as characters / 4, the same rough measure #110 and #112 use.

- `scripts/measure_read_payloads.py` calls the real FastMCP server through `fastmcp.Client`, so the middleware is included. Its document is `tests/fixtures/approved-changes.docx`. "Before" ran the same script against 0471fda (`main` before this change).
- The `list_open_items` numbers come from a fake-pane fixture of 27 open comments (`test_live_comments.py`, `_twenty_seven_comment_document`). It has realistic authors, anchors from 13 to 95 characters, and three comments with replies. "Before" rebuilds the old full shape on the same data.

## Results

| Scenario | Before | After |
|---|---:|---:|
| `list_open_items` full, 27 live comments | 19,099 chars (~4,775 tokens) | 11,482 chars (~2,870 tokens), 39.9% smaller |
| `list_open_items(fields="summary")`, same | not available | 2,308 chars (~577 tokens) |
| `live_status(path=...)`, three panes connected | every session, about 1,870 chars | 525 chars (~131 tokens) |
| `lock_status` | 991 chars (247 tokens) | 220 chars (55 tokens); `verbose` 877 |
| `read_document` text | 1,126 chars (281 tokens) | 859 chars (214 tokens) |
| `read_document` markdown | 1,255 chars (313 tokens) | 988 chars (247 tokens) |
| `find_sections` | 492 chars (123 tokens) | 378 chars (94 tokens) |
| `list_tables` | 289 chars | 274 chars, now with `header_row` and `anchor_section_key` |
| `list_styles` | 15,411 chars (3,852 tokens) | 15,297 unfiltered; 796 chars (199 tokens) with `type="paragraph", name_contains="Heading"` |
| `read_header_footer` | 140 chars | 26 chars |

**Acceptance against the issue:**

- `summary` is about 577 tokens (2,308 chars): under 600. Met. It keeps `path` (JennyStack's "response is for another document" check) and drops summary-only diagnostics (`session_epoch`, which every live id carries; `revision_coverage`; `pending_suggestions_reason`, since `suggestion_rows: null` already says unknown; `file_side_available`).
- `full` is 39.9% smaller against a 40% target, 23 characters short on this fixture.
  - What remains is content that cannot be derived elsewhere: `anchor_paragraph_text` (3,451 characters), `content` (2,718), `anchor_text` (1,354), `created_time` (1,134), `comment_id` (828), `author` (648), `resolved` (459) and `replies` (350).
  - Reaching 40% would mean dropping real content, which the issue rules out ("no information lost").
- `live_status(path)` is under 150 tokens. Met.
- `list_tables` gives one read for "the table under heading X". Met.
- The `list_styles` filters and the `Headng 2` → `Heading 2` near match: met.
- `tools/list` shows the annotations. Met (`tests/unit/test_read_payloads.py`).

The issue suggests flipping the `list_open_items` default to `summary` once JennyStack switches. This PR leaves the default as `full`, because JennyStack's comment loader expects the full shape.

## JennyStack

These are the fields JennyStack's scripts read from the tools changed here. All of them are pinned by `tests/unit/test_jennystack_contract.py`.

- **`list_open_items`** (`comment_dispositions.py`, `live_doc_stamp.py`)
  - Top level: `comments`, `source`, `path`.
  - Per comment: `comment_id`, `author`, `resolved`, `replies` (each reply's `id`/`comment_id`), and the anchor as `quoted_text` / `anchor_text` / `anchor`.
  - Suggestions: rows are recognised by `suggestion_id`, `kind` or `type`.
  - `reply_count` is now omitted. JennyStack reads it as `_count(item.get("reply_count"))`, which gives 0, then takes `max(reply_count, len(replies))`, so the result is unchanged.
- **`read_document`** (`live_doc_stamp.py`, `live_guard_state.py`): `path`, `source`, `format`, `text`/`markdown`, `revision`.
- **Other tools:** no JennyStack script reads `find_sections`, `list_tables`, `get_table`, `read_header_footer`, `live_status`, `lock_status`, `live_revision` or `live_list_sections` results.

## Notes

- **`export_pdf` and `lock_status`** are annotated read-only, as the issue lists them.
  - `export_pdf` writes a PDF and never the .docx.
  - `lock_status` may persist a divergence first-seen timestamp in the server's own state.
  - The README says either can be left out of the auto-allow rule.
- **Not measured here:** the measurement script does not cover `live_status`, because it needs a connected pane. The 525-character figure is from the unit test with three fake panes.

## Review

An independent review of the implementation found these; each is fixed with a test.

| # | Severity | Finding | Resolution |
|---|---|---|---|
| 1 | High | `summary` put row lists under `comments`/`suggestions` and dropped `path`. JennyStack's recorder stores any response whose `comments` is a list, `load_items` skips non-dict rows (an empty table with `ok: true`), and the "another document" check needs `path`. | Rows moved to `comment_rows`/`suggestion_rows`; no `comments`/`suggestions`/`pending_suggestions` key in summary; `path` kept. Size 2,351 to 2,451 chars, then 2,308 after dropping summary-only diagnostics. Scoped responses add `scoped: true`. `test_jennystack_contract.py` simulates the recorder rule. |
| 2 | Medium | Live scoping could never place a suggestion whose paragraph holds a tracked deletion: `paragraph_context` includes `w:delText`, the projection index does not. | A second index is built from the same body XML with `w:t` + `w:delText` and whitespace squashed on both sides. Tested on `revision/tracked.docx`. |
| 3 | Medium | Scoped live `in_scope.suggestions` was 0 when the revision inventory is unknown. | It is `null` then. |
| 4 | Medium | Compact `lock_status` hid `recent` and `ledger`, which the write guard refuses on (`EXTERNAL_EDITOR_ACTIVE`). | Both are kept; the unreadable-ledger case is tested. |
| 5 | Medium | `live_status` default hid `shared.locks`, the only view of who holds a lock behind `LOCKED_BY_OTHER_CLIENT`. | `locks` (filtered to `path`'s document) and `client_id` are top-level by default; the rest of `shared` stays behind `verbose`. |
| 6 | Low-medium | File-mode scoping silently dropped suggestions with no `para_ref` (text boxes, table-row revisions). | Kept and flagged `scope_uncertain`, counted in `in_scope`. |
| 7 | Low | `STYLE_NOT_FOUND` with no near match gave no names. | `available_sample` (first 20 names, sorted) via the shared helper, so all three raise sites get it. |
| 8 | Low (tests) | The read-only annotation test only covered `MUTATING_TOOLS`. | Every tool outside `READ_TOOLS` is checked. New end-to-end test calls `reply_to_comment` and `resolve_comment` with a listed `live:<epoch>:<id>` handle. |
| 9 | Docs | Count keys differ by source; multi-paragraph live comments. | Documented in the `list_open_items` docstring: file `counts` has `comments`/`suggestions`; live has `live_total`/`live_open`/`returned` (+ `in_scope` when scoped); a comment anchored across several paragraphs is `scope_uncertain` in live mode. |

## JennyStack follow-ups

Not done here (separate repo).

- **`skills/jennystack-resolve-doc-comments/SKILL.md`**
  - Lines ~74/83 read `correlation` from live `list_open_items`. It is now opt-in (`include_correlation=true`, as the issue specifies), so the skill must pass it.
  - Line ~145 displays `quoted_text`, which live full mode no longer sends; `anchor_text` carries it.
- **`live_doc_stamp.record_list_items`** should skip responses with `scoped: true`: a scoped `comments` list is partial and must not be recorded as the whole-document read.
