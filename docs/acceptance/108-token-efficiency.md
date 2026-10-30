# #108 token efficiency remainder: acceptance record

## Real Word

`tests/live/token_efficiency_probe.py` on Word for Mac 16.114.3 (2026-10-30). The pane was reopened
so it loaded this branch's `?v=69` build with `partial_font_segments_v1`. **8/8 checks passed.**

| Check | Result |
|---|---|
| Pane connected with `partial_font_segments_v1` | pass |
| (a) Live `replace_text` whose segments set only `bold` applies | pass |
| (a) The replaced run kept the italic and the C00000 colour it inherited, and is bold | pass |
| (b) Approved run with partial segments: prepare binds the font baseline in the manifest | pass |
| (b) A reformat between prepare and execute is refused with `FORMAT_DRIFT`; nothing is written | pass |
| (b) Re-prepared, the same run completes | pass |
| (c) `read_document(source="live", section_key=<heading>)` returns only that section (49 chars) | pass |
| (d) Live `get_table` is compact by default (`rows` grid plus `merged`) | pass |

The "manual" reformat in (b) was a live `format_text` colour change on the target, made between
prepare and execute.

## Token measurement

`scripts/measure_table_session.py`: the issue's session, 3 `get_table` reads and 20
`replace_cell_markdown` edits on a copy of `tests/fixtures/approved-changes.docx` (12x4 table).
Every call uses the tools' defaults. It runs through the FastMCP client, so the lean envelope
applies. Tokens are characters / 4.

| | Baseline `0471fda` | This branch | Fewer |
|---|---:|---:|---:|
| 3 x `get_table` | 3,156 | 595 | 81% |
| 20 x `replace_cell_markdown` | 807 | 807 | 0% |
| **Session** | **3,963** | **1,402** | **65%** |

Both runs end with the same `word/document.xml` (sha256 `f5c186b3…`). The cell-edit responses are
the same size in both runs: #112 had already made write responses compact before the baseline.

Other measurements (`scripts/measure_write_tokens.py`):
- `get_table` on `approved-changes.docx` is 78% smaller.
- A heading-section read is 456 tokens, against 49,003 for the whole large fixture.

## Tool descriptions

Each tool description is trimmed to its purpose, parameters, errors and the safety rules a caller
needs (nothing-written guarantees, "never retried", stale-revision and comment-loss rules,
capability requirements). The full original text moves to `docs/tool-contracts.md`, verbatim, one
section per tool, plus a "Common write parameters" section. Code is unchanged: the parsed code with
docstrings blanked is identical before and after.

| | Before | After |
|---|---:|---:|
| Total description characters (56 tools) | 147,172 | 58,503 |
| `replace_text` | 12,200 | 1,811 |
| Largest description | 12,200 | 2,415 (`insert_table`) |
| Tools over 3,000 characters | 17 | 0 |
| `tools/list` JSON | 192,438 | 101,375 |

`tests/unit/test_tool_descriptions.py` caps each description at 3,000 characters, the total under
60,000 and `replace_text` under 2,000.

## Mutation-response audit

- `tests/unit/test_mutation_response_audit.py` walks every tool in `MUTATING_TOOLS`. Each tool is either covered or explicitly skipped as live-only.
- For each covered tool it checks that the compact response keeps `applied`, `revision_after`, the counts and ids, and any warning, orphaned comment, conflict copy or failed log write the audit evidence recorded.
- It also checks that the audit record holds the full evidence, written before compaction, including the audit-failure path.
- `add_anchored_comment_batch` (#113) is covered.

## Re-sync regressions

`tests/unit/test_shared_bridge.py` covers:
- a partial batch's `revision_after` accepted as the next `revision_before` without a body read
- the next write refused after an unknown outcome, until `live_revision` re-arms it
- another client's write still staling this client
- section locks and unknown-write barriers still refusing
- non-observing lookups never re-arming

The fix: the broker now treats the pane's count gate (zero matches or a count mismatch, under either the old or the new code spelling) and `FORMAT_DRIFT` as proven pre-mutation refusals. Before the fix, a refused edit inside a batch staled every client.

## Merge with #113

Live `replace_text` with partial segments on an anchored or `para_ref` target looks up the font of
the resolved span. A span over 230 characters with partial segments is refused (`INVALID_INPUT`),
because Word cannot search for its first run; full segments still work. Both cases are unit-tested.

## Review

An independent opus review of the code confirmed the re-sync classification. The count-gate
message is only ever raised before a write, in current and older panes. `FORMAT_DRIFT` is thrown
before any insert. It found four defects, all fixed and tested:

1. **High:** partial segments with a text-box or `all` scope sampled the first run with
   `Range.search` inside shape text. That call hangs Word for Mac. Partial segments are now
   body-only: the server refuses them with `INVALID_INPUT` before sending anything, and the pane's
   `replacement_font` and drift paths refuse a non-body scope.
2. **Low:** a short `para_ref` or anchored target containing a soft line break looked up its font
   with a raw `find` and was refused with `ZERO_MATCH`. The lookup now names the target the same
   way the write does (anchored form).
3. **Low:** a heading titled "Textbox" collides with the first text box's key `textbox-1`. A read
   by that key is now refused with `INVALID_INPUT` and gives the heading's paragraph range,
   instead of silently reading the text box.
4. **Trivial:** a font baseline missing a field is rejected even when it carries an extra key.

The pane fix ships as `?v=69b`. The real-Word run above used `?v=69`; v69b only adds the
non-body refusal.
