# Shared live bridge draft (#47)

Status: implemented foundation, pending real Word acceptance and section-lock design.
Base: main `14f8408`; retains merged #66 auto-open/on-demand behavior. Recorded
local acceptance baseline is Word for Mac 16.113.3, macOS 26.6.2. No new real
Word acceptance is claimed by this draft.

## Research and decisions

- The current bridge starts HTTPS/WSS in each MCP process. `SessionRegistry`
  and its live WebSocket objects cannot be shared by merely moving a Python
  singleton. A local broker must own those objects and forward operations.
- Use a private POSIX Unix-domain socket and JSON messages. Python documents
  [UnixStreamServer and threaded servers](https://docs.python.org/3.12/library/socketserver.html)
  and [advisory flock](https://docs.python.org/3.12/library/fcntl.html#fcntl.flock).
  Hold election ownership before generating a certificate/binding either TLS
  port. Only the elected owner removes a stale socket. Same-user access is the
  trust boundary; this does not isolate mutually hostile processes of one user.
- Microsoft documents that concurrent Office synchronization batches are
  [not guaranteed to finish in order](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern).
  Serialize whole operations in the pane, including shape guards/read-back,
  and keep the queue across socket reconnects. A bridge-only lock cannot stop
  an operation that continues in Word after its RPC times out.
- The existing body hash covers text, not comments, formatting, or structure.
  Broker generations invalidate other clients after every attempted write,
  including failure. This protects participating agents across those edits;
  it does not detect every outside human/coauthor formatting/comment change.
- Reviewed [office-word-mcp-server](https://github.com/gongrzhe/office-word-mcp-server):
  MIT license, `word_document_server/tools/document_tools.py` uses
  `python-docx` to read/write saved files; `main.py` offers stdio/SSE/HTTP MCP
  transports. Its tree has no Office.js task-pane bridge or shared live write
  coordinator. No code copied. File editing would bypass this project's live
  verification and would not solve #47.

## Implemented

- Opt-in `VERIFIED_DOCX_SHARED_BRIDGE=1` on every participating MCP server.
  The first explicit `live_status`/`live_open_pane` starts the owner; subsequent
  processes attach. All clients, including the owner, use the broker proxy.
- Unique process client IDs and pane activity labels. `live_status.shared`
  reports current client, attached clients, idle time, and leases.
- Per-document broker operation serialization plus a pane-wide promise queue.
  Reads need no lease but queue behind active operations for consistent state.
- All mutating op families require a read baseline, a current broker generation,
  and a pane advertising `shared_queue`. Caller-provided body hashes are
  forwarded; otherwise the client's last body-bearing read is used, never a
  new read silently performed on behalf of a stale writer.
- Pane-side body freshness preflight for every shared mutation, including
  comment replies, settings and save. Existing per-operation checks remain.
- `live_document_lock(path, lease_s=60)` acquires/renews a document lease;
  `release=true` releases it. Others may read, but writes/lock stealing refuse
  with `LOCKED_BY_OTHER_CLIENT`. Lease bounds: 1-300 seconds. Disconnect/expiry
  releases leases; clients send a heartbeat every 30 seconds.
- Requests include exact pane URL and connection timestamp. Changed panes and
  previously recorded basename collisions refuse. Reconnect clears baselines
  and leases. Transport failures never replay mutations.

## Explicit gaps for Claude to finish

1. **Heading/range locks are not implemented.** Document leases are conservative
   and cannot support parallel ownership of separate sections. Design stable
   heading/range identity and validate the actual mutation target in the pane.
   Never accept an unverified caller-supplied section label as proof of scope.
2. **Owner lifetime:** the first MCP process owns the listeners. Its exit drops
   other clients. Call `live_status` explicitly to elect/attach again and wait
   for the pane to reconnect, then re-read. There is no detached daemon or
   automatic retry of writes. Decide whether the product needs a standalone
   owner before enabling shared mode by default.
3. **Cross-client verification windows:** operation serialization does not make
   a multi-RPC tool (describe, edit, independent verify) atomic. Another client
   can write between those calls. Use a document lease for the entire tool
   workflow until transaction boundaries are designed and tested.
4. **Non-text outside edits:** existing text hashes cannot detect every change
   made outside this broker. Define stronger scope/OOXML revisions before
   claiming full optimistic concurrency against human/coauthor edits. Broker
   generations are not yet embedded in public revision tokens; tool-internal
   reads can acknowledge a newer generation. Carry generation/scope revisions
   through complete tool workflows before claiming caller-read isolation for
   non-text changes. Status-only probes do not acknowledge generations.
5. **Platform/configuration:** POSIX shared mode is opt-in, Windows unsupported.
   Each MCP process must call `live_status` first. All must run compatible
   checkout/pane builds. Startup doesn't automatically switch a legacy owner.
6. **RPC hardening:** version negotiation is v1 only. Frame size is bounded;
   review resource limits/thread counts and shutdown with long-running requests
   before production. The current socket uses the same-user trust boundary.

## Acceptance checklist (all real-client rows pending)

Use disposable local and hosted copies; record exact PR SHA, pane build,
Word/OS version, epoch, sanitized URL and before/after evidence.

- Start two independent Claude/MCP processes together with shared mode enabled.
  Exactly one owns ports 53135/53136; both see the same pane and different IDs.
- Cold-open through merged `live_open_pane`, then auto-open a saved tagged file.
  Reopen the pane to load this draft's JavaScript; reconnect alone is insufficient.
- Both read; A edits; B's stale write refuses without changing the document.
  Repeat with formatting, comment reply/resolve, cells, revisions and auto-open.
- A holds a document lease; B reads successfully and cannot write or release A's
  lease. Verify renewal, expiry, graceful disconnect and crashed-client expiry.
- Delay an edit beyond the RPC timeout. Verify a later op cannot interleave in
  Word, output says outcome unknown, and neither client silently replays it.
- Close/reopen pane, including duplicate-pane and same-filename/different-URL
  cases. Old proxies/baselines must refuse; identity collisions must fail closed.
- Kill owner, observe explicit disconnect, call `live_status` to recover, wait
  for new pane connection, and re-read before any subsequent edit.
- Exercise two different documents, large OOXML reads, and multi-RPC verification
  with and without leases. Record limits rather than treating mock tests as Word
  acceptance. Integrate findings with [#61 handoff](61-live-stack-handoff.md).

Keep #47 open. This draft intentionally uses `Refs #47`, not an auto-close keyword.

## Real-Word acceptance run (2026-10-05, Word for Mac 16.113.3, macOS 26.6.2)

Two separate Python processes (`VERIFIED_DOCX_SHARED_BRIDGE=1`) called the real
server tool functions (`live_status`, `live_open_pane`, `read_document`,
`replace_text`, `live_document_lock`) against disposable copies of
`tests/fixtures/sections.docx`. Branch tip at start: `51e95fe`.

| Row | Result |
| --- | --- |
| Two processes, one owner of 53135/53136, distinct client IDs | Pass |
| `live_open_pane` cold-open connects pane advertising `shared_queue` | Pass |
| A edits; B's stale `replace_text` refuses `LIVE_STALE`, document unchanged | **Failed on `51e95fe`** (B's write applied); fixed, now Pass |
| B re-reads, then edits successfully | Pass |
| Lease: B reads, B write and B unlock refuse `LOCKED_BY_OTHER_CLIENT`; expiry frees writes; holder lock/unlock | Pass |
| SIGKILL owner: B write refuses `LiveDisconnected` with no change; `live_status` re-elects; pane reconnects; stale write refuses until re-read; then succeeds | Pass |
| Env var unset: single process live status/open/edit unchanged, `shared: null` | Pass |

**Bug found and fixed.** Write tools read the live body internally before
editing. The broker counted that tool-internal read as the client observing the
latest generation, so a stale client's write silently re-baselined and applied.
`write_mode.live_session_for` now returns a non-observing view in shared mode;
only an explicit read (`read_document source="live"`) advances a client's
baseline. Consequence: in shared mode a client must read the live document
before its first write (documented behaviour), and reads made by non-`reads_live`
tools do not acknowledge another client's changes. Regression test:
`test_write_tool_session_does_not_acknowledge_other_writes`.

Not exercised here: hosted Word, delayed-edit/timeout interleaving, duplicate
panes, two different documents at once, large OOXML reads, Windows. These and
the design gaps above remain open under #47.

## Automated validation recorded for this draft

- Full unit suite: 836 passed, 23 failed, 217 subtests passed. The 23 failures
  are render/export tests unable to create/read staging directories inside
  Word's macOS container. Running `test_render.py` and `test_server.py` on
  unchanged main reproduces the identical 23 failures (25 pass).
- Final shared-bridge suite: 12 passed. Final live-status and task-pane
  regressions: 16 passed, 58 subtests passed.
- Ruff on new Python module/tests, mypy on the shared module with
  `--check-untyped-defs`, JavaScript syntax and diff whitespace checks passed.
- These results do not establish real Word/hosted multi-agent correctness.
