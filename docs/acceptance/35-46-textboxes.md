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


## #60 whole-paragraph fallback — Claude handoff

Test the exact PR head commit in Word. When shape `Range.search` returns zero,
only exact paragraph text is addressable through paragraph Content ranges.
Substring and cross-paragraph matches must refuse before mutation. Whole
paragraph replacement retains the existing mixed-format refusal for the
`replaced` policy; use `previous` explicitly to inherit the first run when
appropriate. Formatting refuses intersecting comments and revisions.

On disposable local and SharePoint copies, use a Word-authored floating text
box and geometric shape with text. Record their identity, anchors, geometry,
text and formatting before and after whole-paragraph replace/format. Confirm
body text and neighboring paragraphs are preserved. Exercise duplicate matches,
aggregate expected_matches, stale scope tokens, comments/revisions inside the
shape, unsupported substrings, and replacement/format read-back failures.
Record Word version, requirement sets, pane epoch, exact commit and write
evidence. Actual host support remains pending; fix failures or request changes
before merging. No whole-shape OOXML replacement is performed.
