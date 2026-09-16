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
- **Live write mode**: `replace_text`/`format_text`/`live_save` gained
  `write_mode: "auto" | "file" | "live"` in WP-3 (below). The comment
  tools (`add_anchored_comment`, `reply_to_comment`, `resolve_comment`)
  gain the same parameter in WP-4, a separate, parallel PR against this
  same WP-2 base.
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
| `"file"` | Always the file path — today's behavior, byte-for-byte. No bridge/registry lookup at all. |
| `"live"` | Always the live pane. Raises `LIVE_UNAVAILABLE` if no pane session is connected for that document. |
| `"auto"` (default) | Live only when BOTH hold: a pane session is connected for the file's name, AND `lock_status` reports a desktop Word owner file for the path. Otherwise file. |

Matching is by file name exactly as a connected pane reports it
(`documentUrl`'s basename, same as `live/session.py`'s
`document_name_from_url`), never by full path. `resolve_write_mode` never
starts the bridge itself (`live_bridge.current_registry()`, not
`start_in_background()`) — a plain `write_mode="auto"` call (the default
on every `replace_text`/`format_text` call today) must never bind the
bridge's real ports as a side effect; only `live_status`, or an already-
connected pane, brings the bridge up.

**Why a document open in Word is now writable.** Before this WP, a desktop
Word owner file on the target path only ever showed up as data
(`lock_status`) or as `DOCX_LOCKED` on a file-mode write guard. `auto`
routes that exact situation — the lead has the pursuit's document open,
with the Live pane loaded — into a live edit instead: the same call that
would have refused now applies immediately, in front of the lead.

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
`LIVE_UNAVAILABLE`/`LIVE_DISCONNECTED`, and that `write_mode="file"`
and `write_mode="auto"` with no owner file both still take the exact
file-mode path (the fake pane receives no op at all in the latter case).

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
