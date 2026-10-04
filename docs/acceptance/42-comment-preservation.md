# #42 — comment preservation (merge gate)

Depends on #41 and #48. On disposable local and SharePoint documents, record
commit, Word version, tool errors and evidence. Keep draft until checks pass.

Create full-range and partial-range comments with replies. Replace each range:
default must refuse before changing any match, and preserve tracking mode.
Repeat with allow_comment_loss=true: inspect Word and compare the returned
comments_removed threads against actual removals. Multiple matches must all be
preflighted before the first changes. Comments outside the range must survive.
An old pane lacking comment_loss_guard must refuse, not silently ignore the flag.
