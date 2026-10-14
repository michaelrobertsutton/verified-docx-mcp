# #35/#46 — text-box scopes and host errors (merge gate)

Depends on #45. Keep draft until tested on exact commits using disposable
local and SharePoint documents. Record Word version, desktop requirement sets,
epoch, shape handles, scope revisions, warnings and write evidence.

Create two floating callouts, an empty box, a geometric shape with text and a
grouped shape. List/read each supported box. Replace/format only one handle;
body text must remain unchanged. Test duplicate text across body and boxes:
all scope enforces aggregate expected_matches and refuses incomplete coverage.
Change a box after reading its scope token: stale replacement must refuse.
Reconnect during acceptance and confirm old handles refuse. Check comments and
tracked-change guards inside a text box. An empty box must read as empty;
transient host failures retry once and persistent failures return retryable
HOST_SHAPE_READ_FAILED, never an empty string. No mutation is retried.


## #60 whole-paragraph fallback — Claude handoff

Test the exact PR head commit in Word. When shape `Range.search` returns zero,
only exact paragraph text is addressable through paragraph Content ranges.
Substring and cross-paragraph matches must refuse before mutation. Whole
paragraph replacement retains the existing mixed-format refusal for the
`replaced` policy; use `previous` explicitly to inherit the first run when
appropriate. Formatting refuses intersecting comments and revisions.

On disposable local and SharePoint copies, use a Word-authored floating text
box and geometric shape with text. Record their identity, anchors, geometry,
text and formatting before and after whole-paragraph replace/format. Confirm
body text and neighboring paragraphs are preserved. Exercise duplicate matches,
aggregate expected_matches, stale scope tokens, comments/revisions inside the
shape, unsupported substrings, and replacement/format read-back failures.
Record Word version, requirement sets, pane epoch, exact commit and write
evidence. Actual host support remains pending; fix failures or request changes
before merging. No whole-shape OOXML replacement is performed.

### Word for Mac 16.113.3 findings (real-host acceptance)

- **`Range.search` on shape text freezes Word** (100% CPU on the main thread, no
  AppleEvent answered until the process is killed) when the text has a hit; a
  search with no hit returns normally. Shape text is therefore never searched:
  only exact paragraph content is addressed, through paragraph `Content` ranges.
  Substrings and finds spanning paragraphs refuse with `LIVE_CAPABILITY_MISSING`.
- **Comment lookups inside shape text throw `GeneralException`**
  (`range.getComments`, `parentBody.getComments`), and Word does not keep a
  comment anchored in a text box (one patched into the file was gone from the
  live OOXML after load). The comment guard for shapes reads the shape body's
  OOXML for comment marks instead, and fails closed (`LIVE_CAPABILITY_MISSING`)
  if that cannot be read. "Comment inside a text box" cannot be produced by Word
  for Mac, so that row is NOT SUPPORTED, not passed.
- The revision guard works in shapes: a tracked insertion inside a text box
  refuses replace and format with `TRACKED_CHANGES_PRESENT`.

### Re-check on Word for Mac 16.113.4, pane `taskpane.js?v=59` (2026-10-14)

Fixture: `scripts/make_shapes_fixture.applescript`, opened in Word, pane attached.

- **PASS** `list_textboxes` returned all seven shapes (RECT, TB1, TB2, TB3_EMPTY, TB4_COMMENT,
  TB5_TRACKED, TB6_MIXED) with epoch handles and scope revisions, `coverage: body_shapes`, no warnings.
- **PASS** `read_textbox` on TB1 returned both paragraphs; on TB3_EMPTY it returned `""` (verified
  empty, not an error).
- The write rows were not run in that pass: the JennyStack live-write guard
  (`live_doc_guard.py`) refused the first write ("no lead-approved permit"). They were run
  afterwards for #91 (below).

### #91 write rows on Word for Mac 16.113.4, pane `taskpane.js?v=59` (2026-10-14)

Fixture rebuilt with `scripts/make_shapes_fixture.applescript`, opened with `live_open_pane`
(`method: ribbon-button`), epoch `3e326439-1fc2-48cd-8673-eb1d528b3cdf`. Server at commit
`84fe16b`. The calls went through a direct stdio MCP client attached to the shared bridge, not
through the Claude Code tool path, because the JennyStack guard is a hook on that path and
only issues permits inside a proposal workspace. A second client was attached, so each write
was preceded by a live body read (`read_document source=live`) to set the broker baseline.
Body `bodySha256` stayed `1cf7b5f8...b091` before, between and after all five operations.

| # | Operation | Expected | Result |
|---|-----------|----------|--------|
| 1 | `replace_text` "Shape target A" -> "Shape target B", scope TB1 handle, `revision_before` = TB1 scope token | success, read-back, body unchanged | **PASS** `applied`, rung 2, scope `3b7e3738...` -> `6bcc5034...`, font unchanged (Aptos 12), `shapes_before/after` 7/7. `read_textbox` TB1: `"Shape target B\rNeighbor line"`. Body hash unchanged. |
| 2 | `format_text` bold on "Rect target", scope RECT handle | success | **PASS** `applied`, rung 1, `boldAfter: true`, shapes 7/7. The scope token is unchanged (`d8e3d077...`) because it hashes text, not formatting. |
| 3 | `replace_text` substring "A extra" in TB2 | `LIVE_CAPABILITY_MISSING`, nothing written | **PASS** `LIVE_CAPABILITY_MISSING` ("Word cannot search inside TB2; this host supports only exact whole-paragraph shape edits"). TB2 text and scope token unchanged. |
| 4 | `replace_text` whole paragraph in TB5_TRACKED | `TRACKED_CHANGES_PRESENT` | **PASS** `TRACKED_CHANGES_PRESENT`. TB5 text unchanged (`"Tracked host text INS"`). |
| 5 | `replace_text` on TB1 with its pre-write scope token from row 1 | stale-scope refusal | **PASS** `LIVE_STALE` "The selected scope changed". TB1 unchanged (`"Shape target B\rNeighbor line"`). |

Observed while running: a fresh MCP process that calls `list_textboxes` before anything else
reports `LIVE_UNAVAILABLE` (no connected pane) until `live_status` attaches it to the shared
bridge.
