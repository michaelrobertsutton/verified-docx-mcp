# #48 — live revisions (merge gate)

Prerequisites: none. Offline evidence does not prove Word interoperability.
Keep this PR draft until this runbook passes on its exact commit.

When Word is available, use disposable local and SharePoint copies. Record
commit, Word version, WordApi support, pane epoch, counts, and tool evidence.
Create insertions, deletions and formatting changes by two authors, including
duplicate text. Put a revision-free same-named placeholder on disk.

1. List live open items: revisions must match Word, not the placeholder.
2. Check live_status count; unsupported 1.6 must report null with a reason.
3. Accept selected handles, reject others, and exercise all-at-once operations.
   Check selected text and counts in Word; stale/deleted/reconnected handles refuse.
4. Replace text intersecting a pending revision: must refuse without changes,
   and preserve Word's tracking mode. Accept it first, then retry.
5. Confirm file-mode accept/reject on a closed disposable file still works.

No current Word session may be touched to run these checks during implementation.


## Safe revision oracle (#59)

Do not enumerate or count Word revisions through AppleScript on large redlined
documents. Use live body OOXML and compare its revision markup with the live
list's coverage; Office.js alone can undercount adjacent insertions/deletions.
The same-named disk placeholder is not an oracle. Record the live markup count,
listed count and coverage. For host responsiveness, use a bounded document-level
AppleEvent (`count of documents`), not the application name. A timeout does not
cancel Word's in-flight layout work: stop probing and report it. See the
AppleScript freeze guidance in `docs/live-mode.md`.
