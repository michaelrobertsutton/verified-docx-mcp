# #43 — whole-paragraph deletion (merge gate)

Depends on #35/#46 and the safety PRs. Keep draft until tested on exact stack
commits using disposable local and SharePoint documents. Record Word version,
revision tokens, paragraph counts, neighboring content and audit evidence.

Delete one uniquely anchored complete paragraph: its paragraph mark must also
be gone and neighbors unchanged. Repeat with change tracking enabled, inspect
Word's Review pane and accept/reject the deletion; if the paragraph collection
retains deleted marks, the current strict count check must report verification
failure rather than claim success. Verify tracking mode restores on all paths.
Duplicate anchors, anchored floating/VML shapes, comments, pending revisions,
table-cell paragraphs, section boundaries, final paragraphs and malformed OOXML
must refuse without mutation. A stale body revision must refuse before deletion.
