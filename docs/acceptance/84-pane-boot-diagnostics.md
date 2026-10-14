# #84 — pane boot diagnostics and non-fatal load

## Automated

`tests/unit/js/boot_diagnostics_harness.mjs` (wrapper: `tests/unit/test_taskpane_js.py`,
`BootDiagnosticsTests`): ring buffer cap and persistence, `previous_events` reported then
cleared, `boot_seq`, throwing storage, the ops socket starting before and regardless of the
report build, and the minimal report on a rejected or hung build.
`tests/unit/test_pane_boot_log.py` and `tests/unit/test_live_bridge.py`: the bridge boot log,
`recent_boots`, shared-mode status, and the `pane_reload_loop` warning.

None of this is Word. It does not show that Word fires `pagehide` / `beforeunload` in the
task pane before it reloads one, nor what makes Word reload the pane on a co-authored document.

## Real Word: DONE except step 5's root cause (see Results)

Run on desktop Word with the Live pane reloaded (`taskpane.js?v=59`). Record the Word version
and platform, and each result. Read the report with
`cat ~/.cache/verified-docx-mcp/live-report/latest.json`.

1. **Normal load.** Open the pane on a local scratch document. In `latest.json`, expect
   `boot_seq` one higher than on the previous load, `previous_events` empty (or only the prior
   session's events if the pane was reloaded), `instance_id` set, and `boot_ms` with numbers.
   "Open comments N" and the rest of the report are unchanged. `live_status` lists the session
   and `recent_boots` has one entry for the document.
2. **Close, then reopen the pane.** Close the pane, open it again. The new `latest.json` has
   `boot_seq` +1 and `previous_events` containing `pagehide` and/or `beforeunload`. Record which
   of the two (and `visibilitychange`) Word actually fires.
3. **Reload cause.** Reload the pane from the pane's own context menu if available. Record the
   events that end `previous_events`.
4. **Simulated reload loop.** Close and reopen the pane three times within 120 s. `live_status`
   returns a `warnings` entry with `code: "pane_reload_loop"` naming the document, 3 boots, the
   average interval and the `previous_event_kinds`. Two reopens, or three spread over more than
   120 s, give no warning.
5. **SharePoint-synced copy.** If you have the document from issue #84, open the pane on it and
   record whether it loops, and the `previous_events` of the loop.

Forcing an error with a test hook is not required.

## Results

Word for Mac 16.113.4, pane `taskpane.js?v=59`, shared bridge, 2026-10-14. Scratch document
`pane-test-84.docx` (local, in Word's container). The pane was closed through its own close
button and reloaded through the pane's right-click menu (the menu offers Reload, View Source,
Security Info).

1. **Normal load: PASS.** `boot_seq` rose by one over the previous load, `previous_events` was
   empty, `instance_id` was set, `boot_ms` held `office_ready`, `report_build` and `ws_connect`
   (165 / 698 / 468 ms), `load_error` was null. `live_status` listed the session, and
   `recent_boots` had one entry for the document.
2. **Close, then reopen: PASS, with a finding.** `boot_seq` +1, but `previous_events` was
   **empty**. Closing the pane with its close button fires none of `pagehide`, `beforeunload` or
   `visibilitychange` that the pane could record. A close is therefore invisible in
   `previous_events`; only the gap in `recent_boots` shows it.
3. **Reload cause: PASS.** The pane's Reload menu item ends `previous_events` with `pagehide`
   then `visibilitychange` (`hidden`), in the same millisecond. `beforeunload` did not fire.
   The pane's own `boot` event (`boot_seq=N`) leads the list.
4. **Loop warning: PASS.** Three boots at 21:44:41, 21:45:27 and 21:46:11 raised
   `pane_reload_loop` (3 boots, average interval 45.0 s) naming the document. Two boots gave no
   `warnings` key. Three boots spread over 168 s (21:46:39, 21:48:14, 21:49:27) gave none.
5. **SharePoint-synced Row01: loop NOT reproduced.** Opened normally in Word (the pane cannot
   attach to a read-only copy), then the pane was reloaded and left alone. Over a 5 minute watch
   of the boot-report folder there was one boot (`boot_seq` 10, `previous_events` empty,
   `report_build` 451 ms, `load_error` null), the same `instance_id` and steady heartbeats
   throughout. An earlier boot of the same document took `report_build` 3621 ms and
   `ws_connect` 2677 ms (`boot_seq` 9); no failure followed it. The cause of the 2026-10-13
   loop is still unknown. A local, unsynced copy was prepared but not run.

Side finding (not #84): cancelling a `live_open_pane` call from the client dropped the MCP stdio
connection (`STDIO connection dropped after 792s uptime`), which ended the shared-bridge owner
process and left ports 53135/53136 unbound until `/mcp` reconnected. No other server process
took the bridge over.
