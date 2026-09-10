# tests/fixtures/ — provenance

Per the fixture-integrity rule for issue #28 WP-03: a fixture that must catch
the gap between real Word output and this repo's assumptions about it is
worthless if it is a hand-built OOXML file constructed to match those same
assumptions. Every fixture below (existing `pdf/`, `word/` fixtures from
WP-02 aside) is **Word-authored** — produced by driving a real, licensed
Microsoft Word 16.x via AppleScript automation (`osascript`, classic
AppleScript — the same mechanism `render.py`'s module docstring documents
working end-to-end on this machine), never hand-written XML. Every file was
generated into a scratch directory inside Word's own sandbox container
(`~/Library/Containers/com.microsoft.Word/Data/Documents/`) — the same
staging approach `render.py` uses for the render path — then copied out to
this directory with a plain filesystem copy (never a Word "Save As" into
the iCloud-synced repo path itself, which would trigger a native
file-access consent dialog; see `render.py`'s module docstring for why).

No fixture in this WP required human GUI interaction that could not be
scripted — including the comment fixture (`revision/commented.docx`):
`make new Word comment at <range> with properties {comment text:"..."}`
worked via AppleScript and produced the full modern comment part set
(`comments.xml`, `commentsExtended.xml`, `commentsIds.xml`,
`commentsExtensible.xml`, `people.xml`) that a real Word-desktop-authored
comment carries. Nothing here needed a `LEAD:` step.

| File | Purpose | How it was made |
|---|---|---|
| `frag.docx` | A phrase ("The quick brown fox jumps over the lazy dog.") split across 3 runs: bolding the interior word "brown" (`create range ... start 10 end 15` + `set bold to true`) splits the paragraph into a before/bold/after run triple; a same-length typo correction ("jumsp" → "jumps", `create range ... start 20 end 25` + `set content to`) is applied inside the trailing run. Tests that the projection reassembles the fragmented phrase contiguously. | Word-authored (required verbatim by the WP). |
| `tables.docx` | "Before the table." / a 2×2 table (R1C1..R2C2, one paragraph per cell) / "After the table." Tests table-cell paragraphs carry a `container_chain` (table id, row, cell), and the `table_start`/`table_end` structural records in `read_document(format="runs")`. | Word-authored (`make new table at <range> with properties {number of rows:2, number of columns:2}`, then `content of text object of (cell R of row R)` per cell). |
| `fields.docx` | A bookmarked paragraph ("Target1") plus a paragraph containing a complex `PAGE` field (`fldChar` begin/instrText/separate/result/end) and a simple `REF Target1` cross-reference field (`w:fldSimple`). Tests that field instruction text (`w:instrText`, and a `w:fldSimple`'s `w:instr` attribute) is excluded from the projection while the field's result text is included. | Word-authored (`make new bookmark`, `create new field text range <point> field text "PAGE "` and `"REF Target1 "` — Word itself decided the complex-vs-simple field-code shape per type). |
| `textbox.docx` | Body text before/after a rectangle text box (`make new text box`) containing "Text inside the text box." — real Word wraps the box's own `w:drawing` in `mc:AlternateContent`/`mc:Choice` (DrawingML) with an `mc:Fallback` VML (`w:pict`/`v:shape`) branch carrying the identical text; this repo's projection walks only the `mc:Choice` branch (see `projection.py`'s module docstring) to avoid double-counting. Tests the `drawing` structural record and `w:txbxContent` sub-scope projection (`iter_textbox_scopes`). | Word-authored. The one non-obvious step: `text frame`/`text object` of a freshly created shape reads as `missing value` on this Word build's AppleScript surface (a real gap, parallel to the documented `close` command gap in `render.py`) — typing into the box instead required `select <shape>` then `type text (selection) text "..."`, which does work. |
| `headers.docx` | Body text plus a header and a footer, each with text distinct from the body and from each other. Tests `list_parts` resolves `header_footer_type` ("default"/"first"/"even") via `word/document.xml`'s own `headerReference`/`footerReference` elements, and that a header/footer part reads as an independent scope. | Word-authored (`get header`/`get footer` on `section 1`, `index header footer primary`, then `set content of text object`). |
| `sections.docx` | Two `Heading 1` paragraphs ("Overview", "Next Steps") and one `Heading 2` ("Background") in between, each followed by a plain paragraph. Tests `find_sections`' heading-range detection and outline-level resolution. Not explicitly named in the WP's fixture list, but `find_sections` had no other coverage without it. | Word-authored (`set style of (text object of paragraph N) to style heading1`/`heading2`). |
| `revision/base.docx` | A short single-paragraph document; the pin for the revision-token pair below. | Word-authored. |
| `revision/reopened.docx` | `revision/base.docx`, opened fresh in Word and immediately "Save As"-ed with **zero edits**. Tests that `compute_revision`'s token is unchanged after an open-and-resave cycle with no edits — confirmed byte-identical `word/document.xml` (`shasum -a 256` on both files matches). Note: Word's AppleScript surface has no `close` command (documented in `render.py`); this pair round-trips via `open` + `save as` to a different filename instead, which is the exact real-world case the protocol doc's §9 "the lead has the file open in Word for nearly the whole session" describes — a plain `open` with no save at all would trivially leave the file byte-identical, so this pair specifically exercises Word's own resave path, not just "never touched." | Word-authored. |
| `revision/typed.docx` | `revision/base.docx`, opened, with one character ("X") inserted 4 characters in, then saved. Tests the token changes after a real edit. | Word-authored. |
| `revision/commented.docx` | `revision/base.docx`, opened, with one real Word comment anchored to the first word, then saved. Tests the token changes after a comment — and specifically that the `comments` half of the token (not just the `document` half) changes, since `word/document.xml` also gained `commentRangeStart`/`End`/`commentReference` markers in this case. | Word-authored (`make new Word comment at <range> with properties {comment text:"..."}`). |

`tests/fixtures/pdf/` and `tests/fixtures/word/one-page.docx` are WP-02's
existing fixtures; unchanged here.
