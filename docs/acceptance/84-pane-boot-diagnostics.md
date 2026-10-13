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

## Real Word: PENDING

Needs desktop Word with the Live pane reloaded (`taskpane.js?v=59`). Record the Word version
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

## Results: pending
