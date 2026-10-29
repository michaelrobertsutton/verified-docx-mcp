# Cheaper reads: normalized diff, ranged and to-file reads, live revision (#110)

## What changed

- `diff_body_vs_file`: normalized by default (`normalize`), hunks **or** unified
  (`format`), equal hunks carry counts not text, `section_key` scope, `max_hunks` cap,
  `identical_raw` for the byte comparison.
- `read_document`: `start_para_ref`/`end_para_ref` ranges and `to_file`.
- `find_sections`: `chars` on every entry, `digests=true` for per-section `sha12`.
- `live_revision` (pane op `body_hash`, capability `body_hash_v1`, pane v65) and
  `live_list_sections(compact=true)`.

Write refusals are untouched: LIVE_STALE, REVISION_CONFLICT, section-lock and
shared-bridge baseline rules pass their existing tests unchanged. A `body_hash` request
re-arms a client's baseline only when it is an observed read (the `live_revision` tool);
the non-observing status view never does, so a lookup made while refusing a write cannot
re-arm the client it just refused (`tests/unit/test_shared_bridge.py`).

## Method

`scripts/make_large_fixture.py` builds the document: 150 headed sections, 300 paragraphs,
about 20K words, from the Word-authored `sections.docx` through the server's own
`replace_body_markdown`. `scripts/measure_read_tokens.py` runs each scenario and prints
the compact JSON size of the tool result; tokens are estimated as characters / 4, the
same rough measure the issue uses (an estimate, not a tokenizer count). The "before"
column ran the same script against a checkout of `main` before this change.

## Results (file mode, 20K-word document)

| Scenario | Before (tokens) | After (tokens) |
|---|---:|---:|
| `diff_body_vs_file`, identical document | 99,234 | 101 |
| `diff_body_vs_file`, typographic-only differences (NBSP, trailing spaces, CRLF) | 198,154, reported as different | 102, `identical: true`, `identical_raw: false` |
| `diff_body_vs_file`, three real edits | 101,439 | 310 (exactly 3 non-equal hunks, no `unified_diff`) |
| `read_document` markdown, whole body | 49,249 | 49,249 (unchanged, by design) |
| `read_document(to_file=...)` | not available | 419 (under the 500 target) |
| `read_document`, 10-paragraph range | not available (whole body) | 1,786 |
| `find_sections` | 6,106 | 6,593 (adds `chars`) |
| `find_sections(digests=true)` | not available | 7,456 |

`to_file` lists at most 25 heading sections as compact rows, then `section_count` and
`sections_truncated`; on a 150-section document the full list alone would be thousands of
tokens and defeat the point. `find_sections` has the rest.

## Real Word

**Pending.** The pane acceptance below has not been run. The pane loads from the fixed ports
53135/53136 the manifest names, and a process from another agent (a Codex app-server probe,
pid 6363, started 14:50 on 2026-10-29) holds them with a pane connected. Taking them over
would disrupt that session, so the run waits for the ports to be free.

To run: stop the other owner, open a copy of the large fixture in Word, load the v65 pane
from this branch, then check

- `live_revision` returns in under 1 s on the 300-paragraph document with `via: "body_hash"`
  and a `replace_text_batch` using that `revision` as `revision_before` is accepted;
- with `body_hash_v1` removed from the pane's capabilities the fallback works and reports
  `via: "body_ooxml_fallback"` (the 1 s bound applies to the native path only);
- `find_sections(digests=true)` and `live_list_sections(compact=true)` give the same `sha12`
  for every unchanged section (field results, tabs and tracked deletions are where file and
  live paragraph text can differ; fix the file-side text until they agree);
- a live ranged read and `to_file` return the right paragraphs and a matching sha256.
