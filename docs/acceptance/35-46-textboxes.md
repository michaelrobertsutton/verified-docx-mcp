# #35/#46 — text-box scopes and host errors (merge gate)

Depends on #45. Keep draft until tested on exact commits using disposable
local and SharePoint documents. Record Word version, desktop requirement sets,
epoch, shape handles, scope revisions, warnings and write evidence.

Create two floating callouts, an empty box, a geometric shape with text and a
grouped shape. List/read each supported box. Replace/format only one handle;
body text must remain unchanged. Test duplicate text across body and boxes:
all scope enforces aggregate expected_matches and refuses incomplete coverage.
Change a box after reading its scope token: stale replacement must refuse.
Reconnect during acceptance and confirm old handles refuse. Check comments and
tracked-change guards inside a text box. An empty box must read as empty;
transient host failures retry once and persistent failures return retryable
HOST_SHAPE_READ_FAILED, never an empty string. No mutation is retried.
