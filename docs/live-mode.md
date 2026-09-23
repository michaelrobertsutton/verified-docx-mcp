# Live mode

Live co-editing in Word via an Office Add-in bridge
([issue #106](https://github.com/michaelrobertsutton/JennyStack/issues/106)):
a Word task-pane add-in plus a local bridge inside `verified-docx-mcp`
that lets the MCP server send edit and comment operations to a pane
running inside the lead's own open document. Design plan:
`docs/plans/issue-106-word-addin-bridge.md` in the JennyStack repo.

## Architecture

```
Claude Code ──stdio──> verified-docx-mcp (FastMCP)
                          │  existing OOXML tools (unchanged; used when the doc is closed)
                          │
                          ├─ live/bridge.py   HTTPS static (53135): serves addin/ + /ping + /report
                          │                   WSS ops channel (53136): /ops, live/session.py's SessionRegistry
                          │        ▲ wss://localhost:53136/ops        ▲ https://localhost:53135/taskpane.html
                          │        │                                  │
Word for Mac ─── task pane (Office.js, WordApi 1.4) ──────────────────┘
      │  applies ops in the open document; reads back; reports evidence
      └─ co-authoring to SharePoint / OneDrive as usual
```

- **Pane** (`addin/`): classic XML add-in manifest, `taskpane.html`,
  `taskpane.js`. No framework, no build step. Connects to the bridge's WSS
  ops channel on load, sends `hello`, heartbeats every 5s, and dispatches
  every op the MCP server sends (`live/protocol.py`), always reading the
  document back after a mutation (`pre`/`post` body SHA-256) before
  replying. Reconnects with backoff (1s → 30s, doubling) if the socket
  drops. The WP-1 read-only report/ping UI is unchanged and still works.
- **Bridge** (`src/verified_docx_mcp/live/`):
  - `bridge.py` — the WP-1 static HTTPS server (`addin/`, `/ping`,
    `/report`) on port **53135**, plus WP-2's WSS `/ops` channel on a
    **separate port, 53136** (`DEFAULT_PORT + 1`). Two ports because
    `http.server.HTTPServer` (synchronous, blocking-accept) and
    `websockets.asyncio.server` (asyncio) cannot share one listening
    socket's `accept()` loop cleanly; two independent listener threads,
    each with its own loop, is the straightforward option and leaves
    WP-1's static server byte-for-byte as it was. Both listeners bind
    `127.0.0.1` only. `start_in_background()` is what the MCP process
    calls lazily (idempotent — a second call while already running just
    returns the existing `SessionRegistry`) on the first live-aware tool
    call (`live_status` today; a `write_mode="live"` tool call from
    WP-3/4 later); `stop()` tears both listeners down. `--serve-only`
    still works for the lead's manual runbook below, and now also starts
    the WSS ops channel.
  - `protocol.py` — typed message schemas for the wire protocol: pane→
    server `hello`/`heartbeat`/op `reply`, server→pane op `request`
    (`ping`, `describe`, `search`, `replace`, `format`, `comments_list`,
    `comment_add`, `comment_reply`, `comment_resolve`, `save`). Each op's
    docstring names the exact Word JS calls the pane makes to satisfy it.
  - `session.py` — `LiveSession` (one per connected pane, keyed by the
    document's file name parsed from `documentUrl`) and `SessionRegistry`
    (thread-safe; stale eviction after 3 missed heartbeats at 5s each —
    15s). `request()`/`request_threadsafe()` send an op and await the
    matching reply, safe to call from any thread/event loop (not just the
    bridge's own ops-thread loop) — see that module's own comments for
    why that split exists (a real cross-loop bug caught while building
    WP-2). Raises `LiveUnavailable` (no session for that document),
    `LiveDisconnected` (socket closed mid-request, or a reply timeout),
    `LiveOpFailed` (pane replied `ok:false`, e.g. an `expected_matches`
    refusal), `LiveStale` (the pane's body hash before the op didn't
    match a caller-supplied expected hash) — mapped to the matching
    `ErrorCode` (`errors.py`) once a tool calls through this layer
    (WP-3/4).
- **`live_status` tool** (`server.py`, read-only, not in
  `MUTATING_TOOLS`): starts the bridge lazily and reports
  `{bridge_running, port, ops_port, sessions: [{document_name,
  document_url, connected_since, last_heartbeat_age_s, body_sha256,
  requirement_sets}]}`. An empty `sessions` list is normal before the
  lead opens the pane, not an error.
- **Live write mode**: `replace_text`/`format_text` (WP-3, below) and
  `add_anchored_comment`/`reply_to_comment`/`resolve_comment` (WP-4) all
  gained `write_mode: "auto" | "file" | "live"`; `list_open_items` gained
  a parallel `source: "auto" | "file" | "live"`. Not implemented by WP-2
  itself — this WP-2 section describes the transport, the protocol, the
  pane dispatcher, and `live_status` only; see "Live comments (WP-4)"
  below for the comment tools' own design.
- **Testing**: `tests/unit/fake_pane.py` is a Python `websockets` client
  that answers every op deterministically against an in-memory document,
  mimicking the pane's semantics (including the `expected_matches`
  refusal and `pre`/`post` hashes) with no Word installation.
  `tests/unit/test_live_protocol.py` and `tests/unit/test_live_session.py`
  exercise the schemas and the bridge/session wiring through it, on
  ephemeral ports.

## Live edits (WP-3)

`replace_text` and `format_text` (rungs 1 and 2 of
`core/document-backend-protocol.md`'s edit ladder, in the JennyStack
repo) and the new `live_save` tool now have a live path, alongside the
unchanged file path. The shared resolver lives in
`src/verified_docx_mcp/live/write_mode.py` — a small module both this WP
and WP-4's live comment tools import, so the `auto` rule and the live
evidence shape are defined exactly once.

**`write_mode` rule** (`resolve_write_mode(path, requested)`):

| `write_mode` | Rule |
|---|---|
| `"file"` | Always the file path — today's behavior, byte-for-byte, EXCEPT (issue #154): refuses with `LIVE_SESSION_ACTIVE` if a pane session is connected for this document. |
| `"live"` | Always the live pane. Raises `LIVE_UNAVAILABLE` if no pane session is connected for that document. |
| `"auto"` (default) | Live whenever a pane session is connected for the file's name. Otherwise file. |

Matching is by file name exactly as a connected pane reports it
(`documentUrl`'s basename, same as `live/session.py`'s
`document_name_from_url`), never by full path. `resolve_write_mode` never
starts the bridge itself (`live_bridge.current_registry()`, not
`start_in_background()`) — a plain `write_mode="auto"` call (the default
on every `replace_text`/`format_text` call today) must never bind the
bridge's real ports as a side effect; only `live_status`, or an already-
connected pane, brings the bridge up.

### Issue #22 additions: font color, row-scoped edit, live capabilities

- **`format_text`'s `style`** now accepts `color` (a 6-hex string) in
  addition to `bold`/`italic`/`underline`/`strike`, in both file and live
  mode. Live mode also gained `strike` (it was silently dropped before —
  never sent to the pane at all). The pane's `format` reply carries
  read-back `colorAfter`/`strikeAfter` per match (re-loaded from
  `range.font` AFTER `context.sync()`, not an echo of the request), and
  `execute_format_text_live` checks these against what was asked for —
  a write that silently didn't take is now `VERIFICATION_FAILED`, not a
  false `applied: true`.
- **`within_row_containing`** (both tools, file and live): scopes `find`
  to one table row, identified by a second, unique piece of text in that
  same row (`within_row_containing`). Solves the actual incident this
  issue opened with — eleven table cells reading identical text,
  impossible to edit one at a time. The anchor must be unique in the
  WHOLE document (the same guarantee `locate()` already gives any
  needle) and must resolve to a table cell; two cells in the SAME
  anchored row with identical `find` text remain inseparable — a named
  limitation, not solved by this feature.
- **Pane capabilities** (`hello.capabilities`, e.g. `["row_scope"]`):
  distinct from `requirementSets` (WordApi version support) — a
  capability names an OP-LEVEL wire feature this exact pane BUILD
  implements. `within_row_containing`'s live path checks for
  `"row_scope"` before ever sending an op with a `rowAnchor` field,
  raising `LIVE_CAPABILITY_MISSING` otherwise — an already-connected pane
  from before this feature existed would otherwise silently ignore the
  unrecognized field and run an UNSCOPED op instead of refusing, which is
  the exact "wrote to all N identical cells" bug this feature exists to
  prevent, just moved onto the wire. An old pane reports no
  `capabilities` field at all; `HelloMessage.from_json` defaults it to
  empty rather than refusing the connection.
- **Live mode without a local file** (comment tools only — see
  `list_open_items`'s own docstring): `write_mode.py`'s document-name
  resolution and every `comments_live.py` session lookup now tolerate a
  `path` that does not exist locally, using its basename to find a
  connected session by name. `read_document`/`list_parts`/`find_sections`
  and the other structural read tools remain file-only — this does not
  add a live read path for them.

**Why a document open in Word is now writable.** Before this WP, a desktop
Word owner file on the target path only ever showed up as data
(`lock_status`) or as `DOCX_LOCKED` on a file-mode write guard. `auto`
routes that exact situation — the lead has the pursuit's document open,
with the Live pane loaded — into a live edit instead: the same call that
would have refused now applies immediately, in front of the lead.

### Why a file-mode write refuses under a live pane (issue #154)

`auto` originally required BOTH a connected pane session AND
`lock_status` reporting a desktop Word owner file for the path. That
second condition is gone. It was never a reliable signal in the first
place: Word for Mac editing a document opened from a SharePoint/OneDrive-
synced path never writes the `~$*` owner file `lock_status` looks for, so
on that platform/storage combination `auto` could **never** route to
live — it silently fell back to file mode every time, with no warning.
A real incident hit this directly: `apply_style` (which has no
`write_mode` parameter at all, so it only ever took the file path) made
14 calls that each returned `applied: true` while a Live pane was open on
the same document; minutes later Word's own autosave silently reverted
every one, along with several comment replies/resolutions made the same
way.

The fix has two parts:

1. **`auto`/`"live"`'s own routing** (`live/write_mode.py`) no longer
   consults the owner file at all — a connected pane session is
   sufficient on its own, matching the rule the read-side
   `list_open_items(source="auto")` already used
   (`comments_live._resolve_source`).
2. **Every file-mode mutating write** — including the ten tools that
   have no `write_mode` parameter at all (`apply_style`, `insert_table`,
   `insert_image`, `replace_table_row`, `replace_cell_markdown`,
   `replace_range_markdown`, `replace_body_markdown`, `append_markdown`,
   `accept_tracked_changes`, `reject_tracked_changes`) — now refuses with
   `LIVE_SESSION_ACTIVE` when a pane session is connected for that
   document, via `mutations._guard_before_write` (checked once, well
   before the write; and again, immediately before
   `atomic_replace_docx_parts` actually writes, narrowing the race
   between the two). There is **no escape hatch**: an explicit
   `write_mode="file"` refuses too. A write to disk cannot be made
   durable while Word owns the open document's in-memory buffer, so
   there is nothing a caller could safely override to.

**The remedy is to close the document, not just the pane.** Closing only
the Live task pane drops its WebSocket connection, but Word still holds
the document open with its own unsaved in-memory buffer — a file-mode
write immediately after would still race Word's next autosave, which is
the exact hazard this refusal exists to prevent. `LIVE_SESSION_ACTIVE`'s
message says so explicitly: use `write_mode="live"` on a tool that has
one, or save and **close the document itself** in Word, then retry.

**Session identity** (`live/write_mode.py`'s `_check_session_identity`,
issue #154 WP-1b): `SessionRegistry` keys purely on file **basename**
(deliberate — a synced folder's absolute path differs per machine), so a
session for `/a/Foo.docx` and a call against `/b/Foo.docx` would
otherwise match by name alone. `resolve_write_mode`'s `"auto"`/`"live"`
branches additionally compare the target path against the session's own
`document_url`: when that URL resolves to a local file path and it
differs from the target, the call refuses with `LIVE_SESSION_MISMATCH`
rather than risk mutating the wrong document. When `document_url` is a
SharePoint/OneDrive web URL (not a local path — the actual shape of the
incident above), there is no local path to compare, so this falls back
to the basename match already performed; it does not fail closed, since
doing so would disable live mode on exactly the platform this issue is
about.

**Every mutating tool's evidence now states its own `write_mode`**
(`"file"` or `"live"`) — issue #154 WP-3. Previously only live evidence
carried this key; a caller had to infer file-mode routing from the
absence of a `"live:sha256:"`-prefixed `revision_*` token. Every
file-mode evidence-building site now sets `write_mode: "file"`
explicitly (`live_save`'s evidence, which is a live operation despite
living in `text_edit.py`, correctly carries `write_mode: "live"`).

**Live op flow** (`replace`/`format`): `describe` (record the pane's
pre-op `body.text` SHA-256; refuse `LIVE_STALE` if a caller-supplied
`revision_before` starting with `live:sha256:` already disagrees with it)
→ send the op with `expected_body_sha256` set to that same hash (so the
session layer's own staleness check also covers the moment between
`describe` and the op landing) → verify. Verification differs by op:
`replace_text` requires the post-op hash to differ from the pre-op hash
AND every matched range's `after` text to equal the requested
replacement, since a real content change is expected; `format_text`
requires the pane's `applied: true` and every matched range's `after`
text to equal its own `before` text, since a format-only op is expected
to leave `body.text` — and therefore the body hash — unchanged. Either
verification failure raises `VERIFICATION_FAILED` with diagnostics
stating there is nothing to roll back in live mode (Word, not this
server, owns the document).

A pane `ok:false` reply (an `expected_matches` gate failure) is mapped to
`ZERO_MATCH` or `MATCH_COUNT_MISMATCH` by `write_mode.classify_op_failed`,
which parses the actual count out of the pane's own documented message
format (`"...found N"`) rather than requiring a new wire-level error code
— an existing WP-2 test pins today's fake pane (and, per `protocol.py`'s
own documented op contract, the real pane) to a single `LIVE_OP_FAILED`
wire code with the count folded into the message, so this is additive,
not a breaking change to that contract. A future pane/protocol version
that reports `zero_match`/`match_count_mismatch` directly is honored
immediately, with no code change needed here.

**Live evidence** carries the same eight keys file-mode evidence always
carries (`applied`, `match_count`, `rung`, `before`, `after`,
`revision_before`, `revision_after`, `audit_logged`), plus:
`write_mode: "live"`, `verified_via: "word-addin"`, `document_name`,
`track_changes_author: "word-signed-in-user"` (the pane cannot set
`w:author` — live track-changes authorship is always whoever is signed
into that copy of Word), and `orphaned_comment_ids: []`. `rung` is `2`
for `replace_text` and `1` for `format_text` (the edit-ladder rung
numbers, not the file-mode normalization-ladder rung label like
`"exact"`). `revision_before`/`revision_after` are
`"live:sha256:<hex>"` of the pane's own body hash — there is no OOXML
revision token until `live_save` writes the file back to disk. Never
present: `conflict_copy_detected` and its siblings (Word owns the file
for the whole live session, so there is no sync-conflict copy for a
sweep to find) or `revision_ids` (the pane has no id-level view of a
tracked change to report).

**`live_save(path)`** sends the pane's `save` op
(`document.save()`), then reports `{applied, saved, document_name,
revision_after: "live:sha256:...", file_revision, audit_logged}` —
`file_revision` is the plain file-mode revision TOKEN STRING
(`projection.compute_revision(path)["token"]`), computed from the
now-saved file, so a caller can pass it as the next file-mode call's
`revision_before`. `live_save` is in `MUTATING_TOOLS`; `replace_text`/
`format_text` were already there from WP-06 and are not added twice.

**Rungs 3 and 4 never go live.** `replace_range_markdown`/
`replace_body_markdown` have no `write_mode` parameter and always take
the file path, even during a live session — structural verification
against the file happens after `live_save`, via the existing read tools
(`read_document`/`find_sections`/`diff_body_vs_file`) against the
now-saved file, same as any other file-mode read.

**Testing**: `tests/unit/test_live_write_mode.py` covers the resolver
truth table (mocked, no bridge) and, over a real ephemeral-port bridge +
`FakePane`, the live happy paths for `replace_text`/`format_text`/
`live_save`, `MATCH_COUNT_MISMATCH`/`ZERO_MATCH`/`LIVE_STALE`/
`LIVE_UNAVAILABLE`/`LIVE_DISCONNECTED`; the issue #154 regression pin
(`auto` with a connected pane and no owner file goes live, sends the op,
and never calls `lock_status`); `LIVE_SESSION_ACTIVE` refusing
`apply_style` (the literal repro) and an explicit `write_mode="file"`
under a connected pane, plus one representative markdown-rung and one
table-rung tool to show the guard reaches tools with no `write_mode`
parameter at all; `LIVE_SESSION_MISMATCH` on a session naming a
different local file, and the SharePoint-web-URL fallback that still
matches by basename; and a no-bridge test confirming an ordinary
file-mode write never calls `start_in_background`.

## WP-1 result (2026-09-16, Word for Mac 16.112.4)

Recorded from the issue's WP-1 comment
([full text](https://github.com/michaelrobertsutton/JennyStack/issues/106)):

Pane sideloaded from `wef/`, loaded over the local HTTPS bridge
(self-signed cert trusted in the login keychain), read the fixture, and
POSTed its report to the bridge. No prompt or error from Word.

- **Requirement sets:** WordApi 1.4 **true**, 1.5 true, 1.6 true.
  `host: Word`, `platform: Mac`.
- **Document URL** reported as the local path; `body.text` length and
  SHA-256 read fine (100 chars for the fixture).
- **Comment id correlation:** Office.js `Comment.id` values
  (`1487140964`, `586643963`) are **unrelated** to the OOXML `durableId`
  (`58A3F864`, `22F779FB`) and to `w:id` (0, 1).
  `scripts/compare_comment_ids.py` matched both comments by anchor text +
  content and reported `VERDICT: Comment.id is unrelated to both
  durableId and w_id`.
- **Consequence for live comment ops:** the pane's dispatcher
  (`comment_reply`/`comment_resolve`) looks a comment up by the pane's
  own `Comment.id`, valid only within one live session — never by
  `durableId`/`w:id`. Anything that needs to correlate a live comment
  back to `list_open_items`' output (WP-4) does so by anchor text +
  content + author + creation date, the same method
  `scripts/compare_comment_ids.py` uses.
- Multi-paragraph comment content arrives with `\r` between paragraphs.

WP-1 gate passed. WP-2 (bridge sessions, op protocol, pane dispatcher,
fake-pane tests) is implemented above.

## WP-1 sideload runbook (still current)

The steps below still apply unchanged for a fresh sideload — cert
generation, trust, `--serve-only`, and copying the manifest into Word's
`wef/` folder are all the same regardless of WP. `--serve-only` now also
starts the WSS `/ops` channel (a startup line prints its port), which the
pane will connect to automatically once loaded; nothing below requires
that connection to succeed, since the WP-1 report/ping flow is
independent of it.

## 1. Generate the local HTTPS certificate

From a `verified-docx-mcp` checkout:

```
cd /path/to/verified-docx-mcp
uv run python -m verified_docx_mcp.live.bridge --make-cert
```

This writes `localhost.pem` and `localhost-key.pem` to
`~/.cache/verified-docx-mcp/live-cert/` and prints the trust command
(step 2). Pass a directory argument (`--make-cert /some/dir`) to write
somewhere else instead.

## 2. Trust the certificate (once)

Run the exact command step 1 printed, e.g.:

```
security add-trusted-cert -d -r trustRoot -k ~/Library/Keychains/login.keychain-db ~/.cache/verified-docx-mcp/live-cert/localhost.pem
```

This uses your LOGIN keychain, not System, so it needs no `sudo`. Word
requires HTTPS for add-in resources even on localhost — an untrusted
cert is the #1 reason the pane fails to load (see Troubleshooting).

## 3. Start the bridge

```
uv run python -m verified_docx_mcp.live.bridge --serve-only
```

Leave this running in a terminal. It serves `addin/` on
`https://127.0.0.1:53135/` and answers `GET /ping`. Confirm it works
before touching Word:

```
curl -sk https://localhost:53135/ping
```

should print `{"ok": true, "server": "verified-docx-mcp", "time": ...}`.

## 4. Sideload the manifest

Word for Mac loads unsigned add-ins from a fixed folder. Create it if it
does not exist yet, then copy the manifest in:

```
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp addin/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
```

**Restart Word** (quit fully, then reopen) — it only scans the `wef`
folder on launch.

## 5. Open the fixture (never in place)

Copy the WP-1 fixture to a scratch location first — never edit the
checked-in fixture:

```
cp tests/fixtures/comments/multipara-comment.docx ~/Desktop/multipara-scratch.docx
open ~/Desktop/multipara-scratch.docx
```

## 6. Open the pane

In Word: **Home tab → Add-ins group → "verified-docx-mcp" → Live pane**.
The pane loads, calls `Office.onReady`, and immediately reads the
document — you'll see a JSON block and a status line summarizing host,
platform, `WordApi 1.4` support, and comment count.

## 7. Copy the JSON and save it

Click **Copy JSON** in the pane, then paste the clipboard contents into
a file, e.g.:

```
pbpaste > ~/Desktop/pane-report.json
```

(If the clipboard write fails — some sideload contexts block it — select
the JSON block's text directly in the pane and copy it manually; the
status line will say so.)

## 8. Run the comparison script

```
uv run python scripts/compare_comment_ids.py ~/Desktop/pane-report.json ~/Desktop/multipara-scratch.docx
```

This prints a table correlating each pane comment's `Comment.id` to the
server's `comment_id` (durableId) and `w_id` for the multipara fixture's
two comments, plus a one-line verdict: `Comment.id == durableId`,
`Comment.id == w_id`, or `unrelated`.

## 9. Paste both outputs back into the issue

Paste (a) the pane's JSON block (or a representative excerpt — it can be
long) and (b) the full `compare_comment_ids.py` output into
https://github.com/michaelrobertsutton/JennyStack/issues/106 as a
comment.

## What WP-1 must record

Check off each of these in the issue comment:

- [ ] `WordApi 1.4` supported: yes / no (from the pane's
      `requirementSets["1.4"]` field)
- [ ] Comment id-correlation verdict: `== durableId` / `== w_id` /
      `unrelated` / mixed (from the compare script's VERDICT line)
- [ ] Any prompt, warning, or error Word or the pane showed (screenshot
      or exact text) — including anything the "Ping bridge" button
      reported

## Live comments (WP-4)

`write_mode="live"` for `add_anchored_comment`/`reply_to_comment`/
`resolve_comment`, and `source="live"` for `list_open_items` — issue #106
WP-4, built off WP-2's bridge/session/protocol and WP-1's id-correlation
finding above. Implemented in `src/verified_docx_mcp/live/comments_live.py`
(private helpers, `# TODO(WP-3 merge)`-marked for dedup against WP-3's own
`live/write_mode.py` once both land); `server.py`'s four comment tools
dispatch to it. Tested against `tests/unit/fake_pane.py` in
`tests/unit/test_live_comments.py` — no Word, no real pane.

**write_mode / source resolution.** `"file"` is today's OOXML path —
issue #154: BUT it now refuses with `LIVE_SESSION_ACTIVE` if a pane
session is connected for the target document, same as every other
file-mode mutating tool (see "Why a file-mode write refuses under a live
pane" above). `"live"` requires a connected pane session for the target
document's file name (`live/session.py`'s `document_name_from_url`
basename rule), else `LIVE_UNAVAILABLE`. `"auto"` on a WRITE tool (issue
#154) picks live whenever such a session exists — no `lock_status`/
owner-file check any more; that signal never appears on Word for Mac
against a SharePoint/OneDrive-synced path, so the old "both must hold"
rule could never route live there at all. `"auto"` on the READ tool
(`list_open_items`) was already this simple — live whenever a session
exists, since a read never refuses.

**The id-correlation problem.** A live comment's id is a `live:<Comment.id>`
handle, valid only for the one connected session that created it (it does
not survive a Word restart) — WP-1 already showed Office.js's `Comment.id`
has no relation to the OOXML `durableId`/`w:id` `list_open_items`'s file
mode reports. So `reply_to_comment`/`resolve_comment(write_mode="live")`
need to accept EITHER a `live:<id>` handle directly OR a durableId/w:id a
caller already has (from an earlier file-mode call, or pasted in by the
lead), and `list_open_items(source="live")` needs to tell a caller which
durableId a live comment probably corresponds to. Both directions go
through `comments_live.correlate_comments`.

**Correlation algorithm.** For each live comment (`comments_list`'s own
wire shape: `id`, `content`, `authorName`, `creationDate`, `anchorText`),
find the best-matching entry from this same document's on-disk snapshot
(`tracked_changes._parse_comments` — the exact function file-mode
`list_open_items` itself uses, read via a validated snapshot when Word's
owner file is present, same as every other read tool):

1. *Normalize content* on both sides by dropping `\r`/`\n` outright
   (never replacing with a space). The pane's own `Comment.content`/reply
   content joins a multi-paragraph comment's paragraphs with `\r` (WP-1's
   finding); file mode's own join (`tracked_changes.execute_list_open_items`,
   `comments._comment_record`) has no separator between paragraphs at
   all. Dropping `\r` on the live side (not replacing it with a space)
   is what makes the two comparable — a space would break the very
   "comment.Second" boundary the file side already has no gap in.
2. *Normalize anchor text* on both sides by collapsing whitespace runs
   to single spaces and stripping the ends (`" ".join(text.split())`).
3. A live comment whose normalized content matches NO file comment gets
   `confidence: "none"`, `comment_id: null`, `w_id: null`, `date_skew_s:
   null`.
4. Otherwise, the first file comment whose normalized content matches:
   `confidence: "exact"` if its normalized anchor ALSO matches, AND
   (author matches, when both sides report one); `confidence:
   "content-only"` otherwise. **Date is never part of this gate** — see
   below.

**WP-6 real-pane fix: date never gates confidence.** The first real-pane
acceptance run ([issue #106](https://github.com/michaelrobertsutton/JennyStack/issues/106)'s
WP-3+WP-4 acceptance comment) found every real match coming back
`content-only` instead of `exact`: Word for Mac writes
a comment's `w:date` as LOCAL wall-clock time with a `Z` suffix (fixture:
`2026-09-16T14:06:00Z`) while Office.js's `creationDate` reports TRUE UTC
for the same moment (`2026-09-16T18:06:00.000Z` — a 4-hour offset here),
so a fixed tolerance window can never bridge that gap in general (the
offset is whatever the pane's own local timezone is). Date is now purely
informational: every correlation entry carries `date_skew_s`, the signed
whole-second gap `round((live_dt - file_dt).total_seconds())` between
the live comment's `creationDate` and the correlated file comment's
`w:date` (positive means the live side is later), or `null` when either
side has no date or nothing correlated at all.

Verified against `tests/fixtures/comments/multipara-comment.docx` (also
`tests/unit/test_comments.py`'s own fixture) with the exact skewed
timestamps above (`ROOT_DATE = "2026-09-16T14:06:00Z"` on the file side,
`"2026-09-16T18:06:00.000Z"` on the live side): its two-paragraph root
comment (file mode's `comment_id="58A3F864"`, anchored to `"Fixture"`,
content `"First paragraph of a two-paragraph comment.Second paragraph of
the same comment."`) still correlates `"exact"` against a fake-pane
comment with the same anchor/author and content joined with `\r`
(`"First paragraph of a two-paragraph comment.\rSecond paragraph of the
same comment."`), with `date_skew_s == 14400`; a live comment with
unrelated content correlates `"none"`.

**This is advisory, never authoritative.** Confidence is data for the
caller (or the lead) to read, not a gate — `reply_to_comment`/
`resolve_comment(write_mode="live")` accept both `"exact"` and
`"content-only"` matches when resolving a durableId/w:id, and raise
`INVALID_INPUT` (naming both id spaces and pointing at
`list_open_items(source="live")`) only when nothing correlates at all. A
caller already holding a `live:<id>` handle should always use it
directly rather than round-tripping through correlation.

**`w:id` is not stable across saves (WP-6 real-pane finding).** A real
Word for Mac save renumbered a document's second comment's `w:id` from
`1` to `3`. Every correlation entry therefore carries `w_id_stable:
false` — `w_id` is only ever a meaningful correlation key within the one
Word session that has not yet saved since it was read; `comment_id` (the
OOXML durableId) is the durable key across saves and sessions (issue
#108). Because of this, `reply_to_comment`/`resolve_comment`'s
`w_id-correlation` path always resolves against a correlation table
built fresh — by re-reading the on-disk file (`_file_comments_and_
suggestions`) immediately before the op, inside the same call — never a
correlation list cached from an earlier `list_open_items` call, which
could otherwise match the wrong comment (or none) after a save
renumbers `w:id`.

**`list_open_items(source="live")` response shape.** Same top-level shape
as file mode (`path`, `comments`, `pending_suggestions`) plus `source:
"live"` and a top-level `correlation` list (one entry per listed live
comment: `{live_comment_id, comment_id, w_id, w_id_stable, confidence,
date_skew_s}`). Each listed comment's own `comment_id` is
`live:<Comment.id>`, `w_id` is `null`, and it carries
`anchor_text`/`author`/`created_time`/`resolved`/`replies` read straight
from the pane — including real reply threading, which the pane sees but
file mode's own `list_open_items` does not resolve (its
`reply_count`/`replies` are always `0`/`[]` by that tool's own documented
scope limit). Resolved comments are filtered out the same way file mode
does, even though the pane reports them too — `pending_suggestions` is
unaffected either way (read from the same on-disk snapshot regardless of
source, since the live protocol has no tracked-change read op of its
own).

**`add_anchored_comment(write_mode="live")`.** Sends `comment_add`
(`find`, `expected_matches`); the pane's own `zero_match`/
`match_count_mismatch` `OpError.code` values (`live/protocol.py`'s
`OP_ERROR_ZERO_MATCH`/`OP_ERROR_MATCH_COUNT_MISMATCH` — the same two
names WP-3 adds for `replace_text`/`format_text`, so the two branches'
additions to `protocol.py` merge without conflict) map onto `ZERO_MATCH`/
`MATCH_COUNT_MISMATCH`. Known limitation, inherited from WP-2's already-
built pane dispatcher (unchanged by WP-6): the pane's `comment_add` op
inserts on the FIRST match only, even when `expected_matches > 1` —
unlike file mode's one-comment-per-match behavior — though the count is
still verified before anything is inserted. Evidence carries the usual
eight keys (`before`/`after` are the quote, unchanged; `rung` is
`locate.RUNG_EXACT`, `"exact"` — the same value file mode's own
`add_anchored_comment` reports for an ordinary single-pass match, fixed
from an earlier `"live"` placeholder — not the unrelated numeric
edit-ladder `rung` `replace_text`/`format_text`'s live evidence reports)
plus `comment_id`/`comment_ids` (the `live:<id>` handle),
`revision_before`/`revision_after` as `"live:sha256:<pre/post>"`,
`write_mode: "live"`, `verified_via: "word-addin"`, `document_name`,
`author: "word-signed-in-user"` (Word's signed-in user; the pane cannot
be told to claim a different one), and `orphaned_comment_ids: []`. No
conflict-copy fields — those are a file-mode-only concept.

**`reply_to_comment`/`resolve_comment(write_mode="live")`.** Resolve the
handle as described above (`comment_id_resolved_via`: `"live-handle"` |
`"durableId-correlation"` | `"w_id-correlation"`), send `comment_reply`/
`comment_resolve`, then re-list independently to verify — never trusting
the pane's own `ok:true` alone, the same discipline file mode's own
`resolve_comment` already applies to its post-write re-read. Reply
verification looks for a reply whose content equals the text sent;
resolve verification raises the existing `COMMENT_STILL_OPEN` if the
re-list does not show `resolved: true`. `revision_before`/`revision_after`
are now (WP-6 fix) `"live:sha256:<hex>"` of a `describe` call's body hash
taken immediately before and after the op — fixed from an earlier `null`
placeholder — the same token shape `replace_text`/`format_text`/
`add_anchored_comment`'s live evidence already used; the two typically
read equal, since neither a reply nor a resolve edits body text
(mirrors `format_text`'s own `revision_before == revision_after` case).

## What could block this, and the fix

**Untrusted cert.** Word (or the OS WebView) refuses to load
`https://localhost:53135/...` — the pane area stays blank or shows a
certificate error.
*Fix:* re-run the `security add-trusted-cert` command from step 2 (check
it actually prints no error), then fully quit and reopen Word. If it
still fails, open `https://localhost:53135/taskpane.html` directly in
Safari first — Safari will prompt to trust the cert if it somehow was
not added, which is a faster signal than Word's silent blank pane.

**Pane not listed under Add-ins.** The "verified-docx-mcp" button does
not appear on the Home tab.
*Fix:* confirm `manifest.xml` is actually inside
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (not a
subfolder, not renamed), then fully quit and reopen Word — it only reads
that folder at launch, so a copy made while Word is running will not
appear until the next restart.
