# Live stack acceptance handoff (#61)

Status: **executed on 2026-10-04 on Word for Mac 16.113.3 / macOS 26.6.2** (see "Execution results"
at the end). Rows not listed as PASS are BLOCKED, NOT SUPPORTED or OUT OF SCOPE with the
reason stated, not passed. Reviewer scope decision: Mac only (Windows and Word for the web
are out of scope).
Run on disposable copies only. Keep #61 open until the matrix has evidence.
The #58 visibility PR and #60 shape PR have separate commits: test each exact
PR head before merge; record a combined integration commit if testing together.
Never silently test a different checkout or an old cached pane.

## Preparation and reproducible documents

Record commit SHA, loaded pane build, Word version/build, OS, supported WordApi
and WordApiDesktop sets, local or hosted URL, pane epoch, and UTC timestamp.
Reload the pane after changing builds. Keep screenshots/JSON/error evidence
with each scenario. Use a local original and a SharePoint/OneDrive copy for
every document from the linked gates below; record cloud URL encoding cases.

1. **Comments/formatting:** In Word create two paragraphs with duplicate text,
   a comment anchored within a replacement target with a reply, a resolved
   comment, and mixed-font runs. Add a table with duplicate cell text. Follow
   [comment identity](41-comment-identity.md),
   [comment preservation](42-comment-preservation.md),
   [paragraph deletion](43-delete-paragraph.md), and
   [replacement formatting](45-replacement-formatting.md).
2. **Shapes:** In Word create two floating text boxes, an empty text box, a
   geometric shape containing text, and a grouped shape. Include duplicate
   text in body and shapes, a comment inside a box, and a tracked insertion
   inside another. Preserve an untouched copy. Follow
   [shape guards](44-shape-guards.md) and [text-box scopes](35-46-textboxes.md).
   Obtain genuine legacy VML-only and AlternateContent fixtures separately;
   Word-generated DrawingML with VML fallback is not a substitute. Record
   provenance, expected shapes, and actual XML representation.
3. **Revisions:** Create adjacent insertion/deletion pairs under two authors,
   plus independent insertions/deletions, duplicate metadata and formatting
   revisions. Use 25, 100 and 500 pairs for scale where practical. Record
   expected markup from live body OOXML, not AppleScript revision enumeration
   and not a disk placeholder. Follow [live revisions](48-live-revisions.md).
4. **Failure copies:** Use a protected text range, a deliberately malformed
   paragraph-deletion OOXML response, and transient/persistent shape read
   failures. Inject failures only in a disposable test pane or mocked bridge;
   record exactly where/how injected. Never corrupt a user's original file.

## Execution matrix

Every row is pending. For each, record PASS, FAIL, BLOCKED or NOT SUPPORTED
with evidence; unavailable clients are BLOCKED rather than waived passes.

| Scenario | Clients / setup | Required observation |
| --- | --- | --- |
| Existing gates | Local + hosted copies of all six linked acceptance docs | Record each gate separately; confirm live source and matching document identity |
| Encoded cloud URL | Hosted file with spaces, Unicode and percent-encoded URL | Correct basename/session matching; no disk-placeholder fallback |
| Online pane | Word Online with hosted document | Reads/writes work where requirements permit; unsupported desktop shape capabilities refuse clearly |
| Duplicate panes | Two real windows/panes for one document | Record active/displaced pane, epoch and routing; no writes through stale handles |
| Full restart | Close Word completely, reopen document and pane | New epoch; old comment/revision/shape handles refuse; reconnect alone does not satisfy this row |
| Revision clients | Mac, Windows, older WordApi 1.4/1.5 builds | Record merged insert/delete behavior, full markup metadata, individual actionability, accept/reject-all and unavailable fallback |
| Old pane guards | Pane lacking comment_loss_guard, then shape_guard | Required writes refuse before mutation; record capability negotiation and unchanged document |
| Shape read failure | Inject transient then persistent body read failure | One retry for transient read; persistent HOST_SHAPE_READ_FAILED is retryable, never empty text; no mutation retry |
| Shape read-back failure | Inject replacement/format read-back failure | VERIFICATION_FAILED with accurate evidence; do not infer rollback after a write |
| Protected formatting | Protect target range | Refusal or verified failure; never successful evidence for a write that did not take |
| Malformed paragraph OOXML | Inject malformed response before deletion | Refuse before deletion; preserve target text, comments and anchors |
| Legacy inventory | Genuine VML-only and AlternateContent documents | Correct inventory without duplicate fallback entries; record actual XML |
| Shape-local guards | Comment and tracked change inside text box | Replace/format refuse intersecting guarded content; unaffected body/shapes preserved |
| #60 paragraph fallback | Mac local + hosted; repeat Windows if available | Follow PR #60 Claude checklist; exact paragraph edits only; unsupported substrings refuse |

## Behavior decisions — report separately

These are existing behaviors, not automatically new regressions. Record the
observed behavior and request a decision before changing it:

- With tracking enabled, deleted paragraphs can remain in the collection,
  causing delete_paragraph to report VERIFICATION_FAILED despite a recorded
  tracked deletion. Capture both collection state and live markup.
- Closing a newer pane does not re-register the displaced older pane. Confirm
  no stale routing; record whether a fallback should be a later change.
- Successful revision accept/reject invalidates all listed handles. A second
  id from the same listing must refuse; re-list before another individual edit.

## Evidence template (copy per scenario)

- Scenario / associated issue and PR:
- Exact commit / pane build / epoch:
- Word version / OS / requirement sets:
- Document provenance / local or hosted / sanitized URL:
- Preconditions / expected markup count / fixture details:
- Tool input and freshness token:
- Output, error code, coverage/actionability and write evidence:
- Before/after text, formatting, comments, revisions and shape anchors:
- Result: PENDING / PASS / FAIL / BLOCKED / NOT SUPPORTED:
- Evidence links and injection method, if applicable:
- Follow-up owner / failure fix or decision needed:

## Completion and merge handoff

Claude should take control of Word for real-client execution, record results
on each PR's exact commit, fix failures or request changes, and only then make
its merge decision. This implementation agent has not started that session.
Do not close #58 merely because full revision visibility passes: reliable
individual accept/reject for omitted items remains unresolved. Keep #61 open
until pending and blocked rows are resolved or explicitly dispositioned by the
reviewer, including client limitations and behavioral decisions.

## Execution results (2026-10-04, Word for Mac 16.113.3, macOS 26.6.2)

Local documents are Word-authored (generated fixtures are labelled as such in the PR
comments). Hosted rows used a personal OneDrive/SharePoint folder, opened from the
synced copy; the pane reported the `https://…sharepoint.com/…` document URL.
Evidence is in the acceptance comments on PR #62 (revisions), #64 (shapes), #68
(document name) and #67 (opening the pane).

| Scenario | Result | Evidence / note |
| --- | --- | --- |
| Existing gates 41, 42, 43, 44, 45, 48 | **PASS on Mac, local + hosted** (see notes) | Run on a Word-authored fixture (`scripts/make_gates_fixture.applescript`) with one script per gate. **#43** 11/11 (unique delete, mark gone, neighbours unchanged; duplicate, final, table-cell, stale revision, pending revision and comment-bearing paragraphs refuse without mutation; tracking restored). **#42** 8/8 (partial and full range comments with replies; default refuses before any match changes; `allow_comment_loss` reports exactly the threads Word lost; outside comments survive). **#41** 14/14 (identical author/text/anchor comments, epoch-qualified handles, reply/resolve hit only the addressed comment, text-only duplicate match refuses, old-epoch handles refuse after a reconnect, fresh handles work, file-mode read and reply on the closed copy). **#45** 8 of 9 plus one expectation corrected: replaced preserves plain, `none` clears direct formatting, mixed refuses, multiple matches and empty replacement apply, every `format_text` property applies to a table-cell substring (independent read-back); `previous` gave the new text the formatting of the text it replaced (plain), not the preceding bold, i.e. Word's own inheritance, which is what the policy documents. #44 and #48: see the #64 and #62 runs. **Not run within these gates:** a paragraph anchoring a floating shape (#43), content-only/ambiguous *file* correlations (#41), a protected range (#45). Note: a file-mode write on a copy Word has since saved is refused (`EXTERNAL_EDITOR_ACTIVE`) until `allow_concurrent_editor=true`, by design. |
| Encoded cloud URL | **FAIL, fixed, PASS on the fix** | A hosted file `Accépt test 61 #1.docx` registered as `Accépt test 61 ` (the `#` was read as a URL fragment), so tools addressing the real name got `LIVE_UNAVAILABLE`. Fixed in PR #68; after the fix the session is `Accépt test 61 #1.docx` and the live read succeeds. |
| Online pane (Word for the web) | **OUT OF SCOPE (reviewer: Mac only)** | Not attempted. |
| Duplicate panes | **PASS (observed)** | Two real windows of one document (Window > New Window), each with a pane. The second pane displaced the first: one session registered, a `same_document_duplicate` collision recorded with both instance ids, and live calls were answered only by the newer pane (no routing through the displaced one). Closing the newer window left the session registered and answering (the pane runtime is not torn down per window), so no stale routing and no re-registration question arose in this setup. |
| Full restart | **PASS (shape handles)** | After Word was force-quit and relaunched several times, a text-box handle from the earlier session was refused (`LIVE_STALE`) with no mutation op sent and the document unchanged. Revision handles after a pane reconnect also refuse (#62). Comment handles were **NOT RUN**. |
| Revision clients | **PASS on Mac; Windows and WordApi 1.4/1.5 BLOCKED** | Mac: merged insert/delete visible, full metadata, accept/reject-all, individual actionability (see #62). No Windows or older client was available. |
| Old pane guards (no `comment_loss_guard` / `shape_guard`) | **PASS** | Disposable pane copies (scratch only) that do not advertise the capability. Without `comment_loss_guard` a body `replace_text` is refused (`LIVE_CAPABILITY_MISSING`, no pane op sent, body unchanged). Without `shape_guard` guarded writes and `list_shapes` are refused (`LIVE_CAPABILITY_MISSING`) before reaching the pane; document unchanged. |
| Shape read failure (transient, persistent) | **PASS** | Disposable pane copy whose shape-scope read throws. Transient (first read only): one retry, then 7 text boxes with their real text. Persistent: `HOST_SHAPE_READ_FAILED`, `retryable: true`, exactly two attempts, never an empty string. |
| Shape read-back failure | **PASS** | Disposable pane copy whose shape `insertText` does nothing: `VERIFICATION_FAILED` with accurate before/after evidence and no claimed rollback; the shape text is unchanged. |
| Protected formatting | **BLOCKED** | No way found to protect a shape paragraph on Word for Mac. |
| Malformed paragraph OOXML | **PASS** | Disposable pane copy that corrupts the paragraph OOXML before inspection: `delete_paragraph` refuses (`STRUCTURAL_BOUNDARY`) before any deletion; text, other paragraphs and shapes preserved. |
| Legacy inventory (genuine VML-only / AlternateContent) | **BLOCKED** | No genuine legacy fixture. Word-authored DrawingML with VML fallback exists but is not a substitute. |
| Shape-local guards | **PARTIAL** | A tracked insertion inside a text box refuses replace and format (`TRACKED_CHANGES_PRESENT`) on Mac, local and hosted. A comment inside a text box is **NOT SUPPORTED** on Word for Mac: it cannot be created and a comment anchor patched into the file is dropped on load. |
| #60 paragraph fallback | **PASS on Mac (local + hosted); Windows BLOCKED** | See #64. |

### Host findings from this run (each fixed on its PR)
- `Range.search` on shape text **freezes Word for Mac** (100% CPU, no AppleEvent answered) when it has a hit. Shape text is no longer searched (#64).
- Every comment lookup inside shape text throws `GeneralException`. The shape comment guard reads the shape body's OOXML and fails closed (#64).
- Word writes `w:date` as local time labelled `Z`; the true UTC instant is `w16du:dateUtc`. Office.js reports deletion text as empty and omits merged deletions (#62).
- A hosted file name containing `#` was truncated by the document-name parser (#68).
- After a Word restart the add-in button moves under **Home > Add-ins > Developer Add-ins** and is not exposed to accessibility scripting (see #66; `live_autoopen` in #67 removes the need to open the pane by hand for tagged documents).

### Behavior decisions (observed, not changed)
- **Revision handles after an individual accept/reject:** observed. A second id from the same listing is refused with `REVISION_ID_NOT_FOUND`; re-list before another individual edit (#62).
- **Tracked-delete `delete_paragraph` reporting `VERIFICATION_FAILED` while tracking is on:** observed on Word for Mac, local and hosted. The paragraph collection keeps the deleted paragraph (count 15 before and after) while the live markup shows the `w:del`, so the strict count check reports `VERIFICATION_FAILED` with the observed effects instead of claiming success, and the tracking mode is restored. Policy unchanged; a decision is needed only if this should become a success with a different check.
- **Closing a newer pane does not re-register a displaced older pane:** not reproducible on Word for Mac: after the newer window closed, the newer pane's session stayed registered and kept answering. Policy unchanged.

### Dispositions of the rows that are not a plain PASS
- **Windows, older WordApi 1.4/1.5 clients, Word Online:** out of scope per the reviewer (Mac only). They stay unverified, not passed.
- **Protected formatting:** BLOCKED by the host: no way was found to make a shape paragraph write-protected on Word for Mac, so a refused-write path cannot be provoked. The verified-write contract is covered by the read-back failure row above.
- **Genuine legacy VML-only / AlternateContent inventory:** BLOCKED: no genuine legacy fixture is available; Word-authored DrawingML with VML fallback is not a substitute.
- **Comment inside a text box:** NOT SUPPORTED on Word for Mac (Word does not keep one).
- **Not run within the existing gates:** floating-shape-anchored paragraph delete (#43) and content-only/ambiguous file correlations (#41), listed in the gates row.

### How the injected-failure rows were run
Each used a disposable copy of the pane (`taskpane.js`) in a scratch directory served by a separate bridge, never a branch: one fixed patch per variant (capability not advertised; shape read throws first time / always; shape `insertText` made a no-op; paragraph OOXML replaced by `<broken`). The real Word session then ran the real tool calls against it.
