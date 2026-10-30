# Section content digest: `content_sha12` (#117)

## What changed

- `find_sections(digests=true)`: each heading entry gains `content_sha12` beside `sha12`, and the
  result gains `preamble_content_sha12` (content before the first heading; the whole body if
  there is none).
- `live_revision(content_digests=true)` and `live_list_sections(content_digests=true)`: one
  `body_ooxml` export each, adding `content_sections: [[section_key, content_sha12], ...]` and
  `preamble_content_sha12`. Row shapes are unchanged. With `content_digests=false` both tools
  behave as before.
- `content_sha12` covers paragraph texts, tracked deletions (text and position), table structure
  (rows, cells, `gridSpan`, `vMerge`, nested tables) and the root comments anchored in the section
  (author, text, resolved state, replies in order, orphaned replies). The canonical item list is
  documented in `projection.section_content_items`; the README's "Section digests" lists what it
  does not cover.

Design decision (owner, Option A): one definition, computed only by the server from OOXML. File
mode reads the `.docx`; live mode reads the pane's existing `body_ooxml` export. There is no pane
change and no new capability, so the issue's "old pane fallback" is moot: any pane with
`body_ooxml` works. `sha12`, the pane's `sectionSha256`, `sectionGuard`, the shared-bridge lock
paths and every section-lock and shared-bridge test are unchanged. `live_revision` stays an
observed read that re-arms the caller's shared-bridge baseline with `content_digests=true`;
`live_list_sections` stays non-observing (both unit-tested).

## Step 0: feasibility (real Word)

Word for Mac, 2026-10-29. `Body.getOoxml()` (the pane's `body_ooxml` export) carries
`word/comments.xml`, `word/commentsExtended.xml` (with `w15:paraIdParent` and `w15:done`),
`word/commentsIds.xml`, `word/commentsExtensible.xml`, `word/people.xml`, `w:delText`,
`w:gridSpan`, `w:vMerge` and `w:commentRangeStart`. The Word-saved file carries the same parts, so
one server-side definition can serve both modes.

Fixture: `tests/fixtures/content-digest.docx` (provenance in `tests/fixtures/README.md`): two root
comments, one with a reply and one with a reply and resolved, a merged and nested table, and a
tracked replacement ("steps" deleted, "phase" inserted). Every part is Word-written.

## Real Word

PASS, 15 of 15 checks, Word for Mac on 2026-10-29/30, `tests/live/content_digest_probe.py` against
a disposable copy of `content-digest.docx`.

| Check | Result |
|---|---|
| `live_revision(content_digests=true)` | one `body_ooxml` export (`via: "body_ooxml"`) |
| `live_save` | changed no digest |
| file-mode `find_sections(digests=true)` on the saved file vs live `content_sha12` | identical for all 3 sections |
| `preamble_content_sha12` | `4f53cda18c2b` on both sides |
| `sha12`, file vs live | identical |
| `live_list_sections(compact=true, content_digests=true)` | agrees |
| live `resolve_comment` on "Root comment one" | only `overview-1`'s `content_sha12` changed; no `sha12` changed |
| `live_revision`, 5 calls, `content_digests=false` | median 0.348 s (0.316, 0.353, 0.348, 0.348, 0.355) |
| `live_revision`, 5 calls, `content_digests=true` | median 0.249 s (0.28, 0.238, 0.249, 0.256, 0.244) |

Image-heavy document (timing only): the 150-section large fixture (`scripts/make_large_fixture.py`)
plus 12 incompressible 600×400 PNG figures, 8.67 MB as a `.docx`. The `body_ooxml` reply is 318,723
characters, because the pane strips binary parts before transfer. One export takes 0.677 s. Median
`live_revision`: 0.358 s without content digests, 0.59 s with them. Both are under the issue's 1 s
bound, so no per-section export optimization is needed.

File mode on the large fixture without images (150 sections, 300 paragraphs, measured in this
branch): `find_sections` digests take 3.5 ms (`sha12` only) and 4.1 ms (with `content_sha12`);
the server's share of a live content-digest call (Flat OPC to digests, 276K-character export) is
14 ms.

A bug the probe found, on `main` and not in this change: `resolve_comment` / `reply_to_comment`
strip only `live:` from a `list_open_items` handle `live:<epoch>:<id>` and send `<epoch>:<id>` to
the pane, which refuses it. That is fixed in #111, which merges before this. The passing run used
the `live:<id>` form; the probe as committed passes the listed handle unchanged, which is correct
once #111 lands.

JennyStack kept its unchanged-section skip off until this record existed. It now does.

## Review

An independent review of the diff (a fresh opus agent) found five code issues plus gaps in the
docs and tests; the reviewer reproduced the first four with scripts. Fixed:

1. A comment whose `commentRangeStart`/`End` or `commentReference` sat inside a `w:del` was never
   placed, so it dropped out of every digest. Those markers are now recorded (unit test: the
   comment is placed, and editing it changes only that section's digest).
2. `find_sections(digests=true)` failed outright when comment or table canonicalization raised
   `INVALID_INPUT` (duplicate comment `w:id`, ambiguous `paraIdParent`, parent cycle, non-numeric
   `gridSpan`), so a caller who only wanted `sha12` lost it. It now returns `sha12`, sets
   `content_sha12` and `preamble_content_sha12` to null, and warns `content_digest_unavailable`
   with the reason. An explicit `content_digests=true` on the live tools still raises.
3. A deletion split across runs or `w:del` elements (`st` + `eps` vs `steps`) hashed differently,
   so file and live could disagree across producers. Consecutive deletions at the same point are
   now merged before hashing.
4. Comment threading was O(n²): `identity_para_id` scanned both comment parts once per comment.
   The known-paraId set is built once (unit test: 3,000 comments thread in under 1 s).
5. `live_list_sections(content_digests=true)` paired the pane's section keys with the server's
   without checking them. `content_sections` is keyed by the server's own keys, and a
   `content_sections_key_mismatch` warning lists any difference. Export fidelity warnings
   (`live_styles_missing`, ...) are passed through as `warnings` there and in `live_revision`.
   Heading detection is unchanged (it also drives `sha12` and section locks, which are out of
   scope).
6. The docs overstated coverage. README, docstrings and `docs/live-mode.md` now state the limits:
   tracked insertions are plain text, so accepting one does not change `content_sha12`; anchor
   moves within the same paragraph(s) are not seen; the `tblGrid` column count,
   `gridBefore`/`gridAfter`, legacy `hMerge`, tracked paragraph-mark deletions, footnotes/endnotes
   and text boxes are excluded; Word's "remove personal information on save" makes saved authors
   differ from live ones. "File and live agree" now reads as a measured result on Word-saved
   documents, not a guarantee.
7. The comment-ordering test only reordered a reply, which cannot change the order. A new test
   orders two roots, and findings 1, 2, 3 and 5 have their own tests.

Fixes 1, 3 and 5 change canonical output only for inputs the probe document does not contain (a
comment inside a deletion, a split deletion) or add warnings, so the probe was not re-run.
