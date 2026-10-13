# Issue 83: bridge ownership acceptance

Run 2026-10-13 on macOS with Word for Mac, branch `issue-83-bridge-ownership`,
scratch document `tests/fixtures/sections.docx` copy (never a user document).

| Check | Result |
|---|---|
| Two server processes, env unset: first is owner (`is_owner: true`, `mode: shared`), second attaches (`is_owner: false`, same `owner_pid`) | PASS |
| SIGKILL the owner; the surviving process re-elects on its next call (`current_registry()`), becomes owner of 53135/53136 | PASS |
| After all processes exit, ports 53135/53136 are free | PASS |
| `live_open_pane` on the scratch doc connects a pane session through the shared owner; session stays connected 20 s | PASS (`method: addins-popover`) |
| SIGTERM to a stdio server exits within 5 s (unit subprocess test) | PASS |
| Local-mode port conflict names the port and holder PID (unit test with a held socket) | PASS |

Not run: a pane reconnect across an owner kill (pane reconnect timing is the
pane's own retry loop, unchanged here), and the SharePoint-synced crash-loop
(tracked separately in #84).
