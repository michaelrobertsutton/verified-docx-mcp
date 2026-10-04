# Live stack acceptance handoff (#61)

Status: **pending real Word execution**. This document records no new pass.
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
