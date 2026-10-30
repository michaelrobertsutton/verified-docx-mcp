# Trimmed read payloads (#111)

No write-tool behaviour changed. The trimming is response shaping only. A lean
read envelope in `LeanEnvelopeMiddleware` never touches error results, and the
`list_open_items`/`list_tables`/`list_styles`/status changes keep every field
JennyStack reads (`tests/unit/test_jennystack_contract.py`).

## What changed

- **`list_open_items`**
  - `fields="summary"` returns columnar rows: `comment_columns` and `suggestion_columns` with row lists, plus `counts`, `complete`, `scope` and warnings. It drops `pane` and `path`.
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
  - `STYLE_NOT_FOUND` now lists up to five `near_matches` with an `available_count`, instead of every style.
- **`live_status`**
  - `path` scopes the sessions to one document.
  - Heartbeat, boot and shared-broker diagnostics now need `verbose=true`.
- **`lock_status`**: returns `sync_quiesced` plus `external_activity {divergent, still_current}` by default; the raw samples need `verbose=true`.
- **Read envelope**: the middleware drops the following:
  - `path` when it equals the input. `read_document` and `list_open_items` keep it, because JennyStack reads it.
  - empty `section_key`, `lossy_elements` and `warnings`
  - `revision_detail` unless `verbose`
  - `live.body_sha256`

  It also shortens `live_session` to `{document_url, reason}`.
- **Annotations and instructions**
  - `readOnlyHint` and `idempotentHint` are set on the 20 read tools; `destructiveHint` is set on the 7 destructive ones.
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
| `list_open_items(fields="summary")`, same | not available | 2,351 chars (~588 tokens) |
| `live_status(path=...)`, three panes connected | every session, about 1,870 chars | 525 chars (~131 tokens) |
| `lock_status` | 991 chars (247 tokens) | 220 chars (55 tokens); `verbose` 877 |
| `read_document` text | 1,126 chars (281 tokens) | 859 chars (214 tokens) |
| `read_document` markdown | 1,255 chars (313 tokens) | 988 chars (247 tokens) |
| `find_sections` | 492 chars (123 tokens) | 378 chars (94 tokens) |
| `list_tables` | 289 chars | 274 chars, now with `header_row` and `anchor_section_key` |
| `list_styles` | 15,411 chars (3,852 tokens) | 15,297 unfiltered; 796 chars (199 tokens) with `type="paragraph", name_contains="Heading"` |
| `read_header_footer` | 140 chars | 26 chars |

**Acceptance against the issue:**

- `summary` is under 600 tokens. Met.
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
