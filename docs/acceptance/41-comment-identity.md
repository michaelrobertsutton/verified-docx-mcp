# #41 — comment identity (merge gate)

Depends on the live revisions PR. Keep draft until tested in disposable Word
local and SharePoint documents on the exact stack commits. Record Word version,
commit, pane epoch, returned handles and evidence.

Create two comments with identical author, text and anchor text in different
paragraphs, plus one uniquely correlated file comment. List live comments and
read, reply and resolve using epoch-qualified handles and paragraph match specs.
A text-only duplicate match must refuse. Content-only/ambiguous file correlations
must refuse. Restart Word or reconnect the pane only during the later acceptance
session: old epoch-qualified handles must refuse, and freshly listed handles work.
File-mode comment reads/replies/resolution must still work on closed copies.
