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
