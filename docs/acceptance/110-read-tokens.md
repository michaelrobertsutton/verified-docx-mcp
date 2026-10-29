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

PASS, 13 of 13 checks, Word for Mac on 2026-10-29, `tests/live/read_tokens_probe.py` against a
disposable copy of the large fixture (150 sections, 300 paragraphs). The pane was v66 from the PR 2
branch, which includes this PR's pane op unchanged (`body_hash_v1` advertised; this branch's own
pane is v65).

| Check | Result |
|---|---|
| `live_revision` native (`via: "body_hash"`), 300 paragraphs | five calls in a row took 2.31 (cold), 0.66, 0.46, 0.35 and 0.35 s. The 1 s bound holds once warm, not on the first call |
| its `revision` equals a live `read_document`'s | yes |
| `find_sections(digests)`, `live_list_sections(compact)`, `live_revision` vs the file's `sha12` | identical for all 150 sections |
| live ranged read (10 paragraphs) in text, markdown and runs vs the file range | identical |
| live `read_document(to_file=...)` on the whole body | 382 tokens, sha256 matches the written file |
| `replace_text_batch` with `live_revision`'s revision as `revision_before` | accepted, `complete` |
| the same revision after that write | `LIVE_STALE`, nothing written |
| fallback to `body_ooxml` (capability hidden from the server) | `via: "body_ooxml_fallback"`, same revision and sections as native, 0.31 s |

Two things the run found. First, a bug: `Body.getOoxml()` appends one empty paragraph that
`Body.paragraphs` (what the pane digests) does not have, so the last section's digest from the
fallback and from `find_sections(source="live")` never matched the pane's. `find_sections_impl` now
takes `ignore_trailing_empty` for live projections; there is a unit test, and the rerun is the
passing one above. Second, an earlier run had 4 failures that were real edits, not bugs: text typed
into the Word window (a heading became "Can you please Section 001") and the probe's own rename of
one heading; the other 148 sections agreed. The passing rerun started from a pristine copy of the
fixture (the probe expects one and does not enforce it).

Field results, tabs and tracked deletions, where file and live paragraph text can differ, are not in
this fixture and remain untested.
