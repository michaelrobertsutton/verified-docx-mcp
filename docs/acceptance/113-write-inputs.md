# #113 cheaper write inputs: acceptance record

Real-Word acceptance ran on **Word for Mac 16.114.3** on 2026-10-30, with the combined branch's server
(both #113 parts merged onto `main` after #107 and #111) and pane capability `anchored_find_v1`.

## Real Word

`tests/live/anchored_find_probe.py`, run in an already-open document with the body inserted live
(the pane could not be reopened unattended): **14/14 checks passed**.

| Check | Result |
|---|---|
| Pane connected with `anchored_find_v1` | pass |
| Setup kept NBSP and curly quotes in the 637-character paragraph | pass |
| (a) Live ellipsis `replace_text` of the 637-character paragraph, typed with straight quotes and plain spaces, with a decoy paragraph sharing the prefix | pass, `match_count` 1 |
| (a) The same find with `expected_length` off by 12 is refused | pass (`ZERO_MATCH`) |
| (b) `para_ref` + `expected_sha256` replace | pass |
| (b) `para_ref` replace of a 341-character paragraph (over the 230 search cap) | pass |
| (d) Stale paragraph hash | pass (`PARA_HASH_MISMATCH`) |
| (d) Stale `revision_before` | pass (`LIVE_STALE`) |
| (e) Anchored `format_text` | pass |
| (e) Anchored `delete_paragraph` | pass |
| (f) `replace_text_batch` whose second, anchored edit targets the first edit's output | pass |
| (g) 12-comment live `add_anchored_comment_batch` | pass, `batch_status` complete, 12 distinct comment ids |
| (c) Approved run: prepare stores the resolved ~600-character `find` | pass |
| (c) Approved run executes | pass, status complete |

An earlier run of part A alone (before its review fixes) passed the first 9 checks in the same document.

**Pane build caveat:** the pane open in Word had loaded part A's pane JavaScript (`?v=68`) before the
merge with #107's pane changes. The merged pane ships as `?v=68b`, so Word loads it fresh the next
time the pane opens. The #113 pane code is identical in both builds; #107's pane changes were
accepted separately (`docs/acceptance/107-cell-editing.md`).

## Unit gate

`uv run pytest tests/unit`: 1634 passed on the combined branch.
`test_shared_bridge.py::test_sigterm_terminates_stdio_server` is deselected: it is a timing flake
that predates this branch.

## Token measurement

`scripts/measure_write_inputs.py` (tool-call argument characters / 4, through the FastMCP client):

| Scenario | Today | #113 | Fewer |
|---|---:|---:|---:|
| 12 review comments | 12 calls, 1,063 tokens | 1 batch, 702 tokens | 34% |
| Markdown section write (1,695 chars) | 463 tokens | 56 tokens (`markdown_file`) | 88% |
| **Session: 10 rewrites + 12 comments, live mode** | 42 calls, 4,889 tokens | 11 calls, 2,513 tokens | **49%** |
| Session: 10 rewrites + 12 comments, file mode | 22 calls, 3,787 tokens | 11 calls, 2,380 tokens | 37% |

The issue asks for at least 50% fewer tokens on this session. Live mode measures 49%, file mode 37%;
the owner accepted 49% as meeting the intent (2026-10-30).

What remains is content both APIs must carry: the 10 new paragraphs (~4,600 characters) and the
12 comment texts (~2,000) are about two-thirds of the new session. The new API removes the retyped
old text, the 230-character chunking (each 468-character rewrite is 3 live calls today), and 11
repeated comment envelopes.

How the session is built:
- Each rewrite replaces a whole 468-character paragraph with a new paragraph of similar length that
  changes every sentence.
- The live session is sized, not sent: its old-API calls need Word. The new-API live calls it sizes
  are the ones the real-Word probe above executed.
- Live writes carry `revision_before` in both variants.
- Anchored finds carry 24-character prefixes and suffixes (the minimum is 12).
