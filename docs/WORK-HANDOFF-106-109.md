# Resume #106–#109 here

User requested a pause for usage reset on 2026-10-29. Resume this same chat.
**No issue is complete. No PR has been merged or issue closed.** Do not imply otherwise.

## User intent and constraints

Implement/accept/merge/close #106, #107, #108, #109 separately, in that order.
Include JennyStack compatibility. Compact responses should be defaults. #109 includes
multi-table comparison under headings. Real Word acceptance gates merges. At the end,
clean up only task-created branches/worktrees and update both primary local mains.

The user is understandably frustrated about spending the session without a landed
issue. Prioritize finishing and landing #106; do not expand scope or redo exploration.
**Use one original Word add-in only. Never create a second testing add-in.** The
attempted duplicate was removed; Word was quit/reopened per the user's request.
No further manual menu hunting/hovering should be asked of the user.

## Repositories and saved branches

Backend working directory: `/Users/msutton/.codex/worktrees/67c3/verified-docx-mcp`.
- `codex/106-table-structure`: `ccd44af`, initial #106 candidate.
- `codex/107-cell-editing`: `ca87b82`, stacked on #106.
- `codex/108-token-efficiency`: current branch; checkpoint includes unfinished #108
  AND subsequent real-Word fixes for #106. Extract those fixes into #106 before landing
  and restack later branches. The checkpoint is NOT a merge-ready implementation.

JennyStack isolated worktree: `/private/tmp/jennystack-word-compat`, branch
`codex/106-word-compat`, based on `098640a`. Only guard registration and a regression
have been changed so far. Primary JennyStack checkout was clean when inspected.
Its CLAUDE.md requires verify.sh, version/changelog synchronization, rebase before
landing, pre-landing review, and core-divergence review for core/skills changes.

## Verification already completed

- #106 initial unit gate: **1,144 tests, 297 subtests passed**.
- #107 final unit gate: **1,149 tests, 297 subtests passed**. A subsequently added
  per-edit row-scope test passed with the focused five-test #107 regression suite.
- Focused #107 behavior/shared/approval tests: **95 passed**, seven deliberately
  deselected socket/process tests; full gate above covered the complete suite.
- Latest focused #108/formatted/read/table tests: **54 passed**. This is NOT a full
  #108 gate or acceptance. No token benchmark has run.
- JennyStack guard tests: **48 passed** before later additions to merge-tool lists.
- node --check passed at intermediate checkpoints; rerun after subsequent edits.

Full test commands need normal local access: sandboxed runs fail on localhost sockets
and macOS sips. Use the approved escalation `uv run pytest tests/unit --no-cov -q`.
Logs: `/private/tmp/verified-docx-107-final-unit.log` and
`/private/tmp/verified-docx-107-unit.log` (the latter contains obsolete failures that
were fixed). Avoid rerunning full suites while still making the changes under test.

## Word setup: fixed to ONE normal add-in

Only `manifest.xml` remains installed in
`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`.
The original manifest/installed Python environment/primary checkout were not edited.
The extra `verified-docx-acceptance.xml` was removed. Alternate-port test bridge stopped.
The old development bridge owner (PID 66154) was stopped while Word was closed.

A branch-backed shared bridge now owns the usual 53135/53136 ports, process PID 6363
at checkpoint (verify owner.pid before relying on it), via:
`/private/tmp/verified_docx_single_pane_probe.py`.
It runs in exec session **1733** and accepts JSON lines over stdin. It imports Python
modules at startup, so **latest on-disk namespace/audit/#108 changes are not loaded
in that running process**. Stop/restart this controlled probe to load them, once no
write is pending. Its current script startup simply starts/claims the normal bridge;
there is no longer any kill logic in the script. Use PYTHONPATH=src and .venv/bin/python.
The old alternate-port probe script is obsolete; do not use it.

Word has `/private/tmp/verified-docx-106/rows.docx` open, with the original live pane.
User handed over Word control. Only scratch documents have been changed.
Last verified text hash: `b36e7ff61c9cb8aafc1f1f64e896773726f3f9f7974127a53db8e5106854d29a`.
Table 1 currently has rows:
`R1C1\tR1C2`, `New A\tNew B`, `R2C1\tR2C2`.
No request was pending at checkpoint. Last failed copy attempt did not change body hash.
Always snapshot/read before further writes; do not trust this saved hash indefinitely.

Probe commands: status, snapshot(output filename), row(args forwarded to
 table_structure.execute_row), copy(args to tables.execute_insert_table), cell(args),
paragraphs(args), merge(args), delete_paragraph(args), table(args), save(args).
A snapshot writes full Flat OPC to the requested local file and returns only the hash.

CUA is available for native UI. Its persistent binding is `word`.
For pane reload, read a fresh full AX tree, click `Close verified-docx-mcp live`, read
again, and click the ordinary `Live pane` ribbon button. This worked and does not
close Word. Native ribbon menus sometimes needed raising the window first. Do not
repeat the earlier add-in setup experiment. Do not restart Codex for recording
permissions; coding does not require it. No audio recording was used.

## #106: what remains before landing

Code: table_structure.py, tables.py/server registration, pane ops, middleware,
protocol/session capability registration, tests/unit/test_table_structure.py.

Real Word insertion and deletion now **passed**, including exact direct formatting
comparison. Initial live failures were diagnosed and corrected on the current branch:
1. Bare w:document cell imports are InvalidArgument; use cloned full Flat OPC.
2. Native insertRows adds font properties on paragraph marks. Reset the new empty
   cell font before importing the template pPr/rPr; exact formatting then passed.
3. Paragraph-by-paragraph shape inspection of empty table cells raised GeneralException;
   inspect the already-loaded table OOXML for shape nodes instead.
4. Mutation-stage errors now map to VERIFICATION_FAILED with stage/location details.
   Server-side failure auditing was added later and still needs rerun/review.

Faithful table copy currently fails GeneralException with no text change. Suspected
namespace loss during Python Flat OPC repackaging. A candidate namespace-preservation
fix is ON DISK but not loaded/tested: ReadSource carries internal namespace_declarations,
flat_package serializes with the original package declarations. Verify the fix first.
Add stage/errorLocation diagnostics to table copy if needed; the row diagnostics helped.

The native row path still refuses a nonadjacent template with different trPr/tcPr.
This does not fully satisfy the optional arbitrary template requirement; resolve it
before claiming #106 complete. Do not silently waive this limitation.

The scratch fixture sets pStyle=Compact without defining that style, so Word drops
that style. Define a real non-default Compact paragraph style in the next scratch
fixture to satisfy acceptance. Existing live tests did exercise actual direct 8.5pt,
keepNext, spacing, borders, fixed layout and cantSplit. Also test explicit copy overrides,
refusals, tracking policy and untouched neighbors. Native Word build still needs recording.

Raw snapshots retained in /private/tmp/verified-docx-106-*.xml, including before,
after-first-attempt, after-delete-attempt, after-import, and copy-attempt.
Evidence/audits under `/private/tmp/verified-docx-acceptance-state`.
See docs/acceptance/106-table-structure.md (its initial status predates the live passes).

## #107: candidate implemented; real Word acceptance pending

Default first-run name/size/color inheritance, including approved cell stream;
cell paragraph deletion refusing the last paragraph; merge_paragraph_with_next;
positional table anchors/filling empty paragraphs; per-edit batch row anchors;
save classified separately from body writes/read baselines. Approved saves still
validate protected scopes. File tracked paragraph merges explicitly refuse.

Real Word tests for these operations remain to run. Review merge safety, first pPr
and hyperlink preservation, and font invariance; passing mock/file tests are insufficient.
See docs/acceptance/107-cell-editing.md and tests/unit/test_cell_editing_107.py.

## #108: unfinished work on current branch

- responses.py adds public_write decorator/signature verbose=false to public mutation
  tools, after internal verification/audit; compact batch entries retain statuses,
  match_count, revisions, warnings. Signature import was checked after fixing a
  Pydantic KeyError for verbose annotations. Review compaction for every tool.
- Public get_table defaults compact=true; internal helpers remain verbose.
- live_revision explicitly calls observing describe; needs shared-client regressions.
- Partial segment validation implemented in Python and JS. Approved prepare resolves
  font baselines through new read-only replacement_font op, binds resolved values and
  baseline in the manifest, checks baseline before mutation. Needs meaningful tests,
  fake pane support, capability/unknown-outcome review, and real Word acceptance.
- Heading-scoped projection/read path added, retaining table/paragraph IDs; new targeted
  heading tests and full file/live checks still needed.
- Six high-traffic tool descriptions shortened; old detail moved to
  docs/tool-contract-details.md (heading ordering needs cleanup).
- Need add table revision/provenance support where absent, complete JennyStack parser
  compatibility, benchmark 20 edits + three table reads with a fixed tokenizer and
  prove >=50% fewer response tokens, docs/version/changelog and full gate.
- No version bump yet; agreed next minor release for changed response defaults.

JennyStack remaining compatibility: accept compact + legacy verbose batch results
strictly; preserve uncertain-outcome barriers; ignore response-only verbose flag in
permit matching; avoid stamping heading reads as full-body evidence; keep live_revision
separate from drift/comment/unknown-outcome reconciliation; update canonical docs and
formatted segment instructions. Run required tests/reviews/release checks before landing.

## #109: not started

Implement read-only single/multi-table comparison using existing read_source and
MarkdownIt table parsing. Inline or allowed markdown file; heading/header selection;
normalized quote/whitespace/emphasis handling; differing cells only with 1-based
coordinates; duplicate headers refuse; unmatched/dimension/merged/nested cases explicit.
Acceptance requires exactly three changed cells and multi-table edge cases. Plan details
are in this chat. Preserve the requested landing order.

## Landing/cleanup still entirely outstanding

Create/attach reviewable PRs, pass acceptance for each, merge with closing references,
verify issue closure. No PRs created yet. Restack the issue branches correctly first,
including extracting the subsequent #106 fixes out of the #108 checkpoint.
Do not remove unrelated pre-existing `issue-35-live-textboxes` worktree.
After all four issues and JennyStack compatibility land, remove task branches/worktrees,
restore normal bridge startup against updated primary main, and fast-forward both mains.
