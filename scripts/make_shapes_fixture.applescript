-- Build the Word-authored shape fixture used by the live text-box / shape acceptance (#35-46, #44, #60).
-- usage: osascript scripts/make_shapes_fixture.applescript [output.docx]
--   default output: ~/Library/Containers/com.microsoft.Word/Data/Documents/shapes-fixture.docx
--   (Word is sandboxed: it can only save inside its container without a permission prompt)
--
-- Creates a document with body text and these shapes (names are the shape names):
--   TB1          floating text box: "Shape target A" + "Neighbor line" (two paragraphs)
--   TB2          text box: "Shape target A extra" + "Shape target A" (substring + whole-paragraph duplicate)
--   TB3_EMPTY    empty text box
--   RECT         geometric rectangle: "Rect target"
--   TB4_COMMENT  text box (Word for Mac cannot hold a comment in a text box; this is the host text only)
--   TB5_TRACKED  text box with a tracked insertion " INS"
--   TB6_MIXED    text box "Mixed format line", first word bold (mixed formatting)
-- Notes learned on Word for Mac 16.113.3: text goes in via `text range of text frame` for text boxes,
-- a rectangle needs select-and-type, ranges inside a shape must be addressed through the shape's own
-- text range (`word 1 of ...`, `insert ... at end of ...`), never `create range` offsets (that is the body).
on run argv
	set outPath to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Word/Data/Documents/shapes-fixture.docx"
	if (count of argv) > 0 then set outPath to item 1 of argv
	tell application "Microsoft Word"
		set d to make new document
		set content of text object of d to "Body paragraph with Shape target A inside." & return & "Body paragraph two."
		set tb1 to make new text box at d with properties {left position:60, top:200, width:240, height:80}
		set name of tb1 to "TB1"
		set content of text range of text frame of tb1 to "Shape target A" & return & "Neighbor line"
		set tb2 to make new text box at d with properties {left position:60, top:300, width:240, height:80}
		set name of tb2 to "TB2"
		set content of text range of text frame of tb2 to "Shape target A extra" & return & "Shape target A"
		set tb3 to make new text box at d with properties {left position:340, top:200, width:150, height:50}
		set name of tb3 to "TB3_EMPTY"
		set rc to make new shape at d with properties {auto shape type:autoshape rectangle, left position:340, top:300, width:200, height:60}
		set name of rc to "RECT"
		insert text "Rect target" at beginning of (text range of text frame of (shape "RECT" of d))
		set tb4 to make new text box at d with properties {left position:60, top:400, width:240, height:60}
		set name of tb4 to "TB4_COMMENT"
		set content of text range of text frame of tb4 to "Comment host text"
		set tb5 to make new text box at d with properties {left position:340, top:400, width:240, height:60}
		set name of tb5 to "TB5_TRACKED"
		set content of text range of text frame of tb5 to "Tracked host text"
		set tb6 to make new text box at d with properties {left position:60, top:480, width:240, height:60}
		set name of tb6 to "TB6_MIXED"
		set content of text range of text frame of tb6 to "Mixed format line"
		set bold of font object of (word 1 of (text range of text frame of (shape "TB6_MIXED" of d))) to true
		set track revisions of d to true
		insert text " INS" at end of (text range of text frame of (shape "TB5_TRACKED" of d))
		set track revisions of d to false
		save as d file name outPath file format format document default
	end tell
	return outPath
end run
