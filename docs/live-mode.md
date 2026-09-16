# Live mode: WP-1 spike runbook

This is the lead's step-by-step for issue #106 WP-1 — a hello-world Word
task-pane add-in, served over local HTTPS by
`verified_docx_mcp.live.bridge`, sideloaded once into Word for Mac
(16.112.4). It proves four things before anything else in the plan
(`docs/plans/issue-106-word-addin-bridge.md` in the JennyStack repo) gets
built:

1. the pane loads in Word,
2. `WordApi 1.4` is supported,
3. the pane can read the document URL and a body-text hash, and
4. how Office.js's `Comment.id` relates to the OOXML `durableId`
   `list_open_items` reports.

No document is edited at any step below. The agent cannot run Word or
sideload the manifest — every step here is something only the lead can
do at the keyboard.

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
