-- Open the verified-docx-mcp Live pane for a document that is already open in Word for Mac.
-- usage: osascript scripts/open_live_pane.applescript "<document name, e.g. report.docx>"
--
-- Same logic as the `live_open_pane` tool (src/verified_docx_mcp/live/open_pane_mac.py), for use by hand
-- or from other scripts. Needs the calling app to have macOS Accessibility permission. Prints which
-- method was used; it cannot confirm the pane connected, so check `live_status` afterwards.
on run argv
	set docName to item 1 of argv
	tell application "Microsoft Word"
		activate
		activate object (first document whose name is docName)
		-- Wide enough that the ribbon's right-hand groups are not cut off (Word clamps to the screen).
		set bounds of front window to {0, 40, 4000, 1000}
	end tell
	delay 1
	-- Word's window titles drop the extension ("report", not "report.docx").
	set AppleScript's text item delimiters to "."
	set parts to text items of docName
	if (count of parts) > 1 then set docBase to (items 1 thru -2 of parts) as text
	if (count of parts) = 1 then set docBase to docName
	set AppleScript's text item delimiters to ""
	tell application "System Events" to tell process "Microsoft Word"
		set frontmost to true
		-- `activate object` changes Word's active document, but window 1 can still be another window: raise
		-- the target explicitly and refuse to continue on the wrong one (the pane would attach to it).
		set target to first window whose name contains docBase
		perform action "AXRaise" of target
		delay 0.5
		if (name of window 1) does not contain docBase then error "could not bring " & docName & " to the front"
		-- The Home tab click TOGGLES the ribbon: only click it when the ribbon contents are not exposed.
		if (count of scroll areas of tab group 1 of window 1) = 0 then
			click radio button "Home" of tab group 1 of window 1
			delay 2
		end if
		-- 1. The top-level "Live pane" button (present once the add-in has run in this Word session).
		repeat with g in (every group of scroll area 1 of tab group 1 of window 1)
			try
				click (first button of g whose name is "Live pane")
				return "ribbon-button"
			end try
		end repeat
		-- 2. The Add-ins popover. It is not in the accessibility tree, so it needs real mouse events,
		--    positioned relative to the Add-ins button.
		repeat with g in (every group of scroll area 1 of tab group 1 of window 1)
			try
				set b to first button of g whose name contains "ins" and name contains "Add"
				set p to position of b
				set s to size of b
				set cx to (((item 1 of p) + (item 1 of s) / 2) as integer)
				set cy to (((item 2 of p) + (item 2 of s) / 2) as integer)
				my clickAt(cx, cy)
				delay 2.5
				my clickAt(cx - 108, cy + 357) -- first click hovers the "verified-docx-mcp" entry...
				delay 3
				my clickAt(cx - 108, cy + 357) -- ...the second one activates it
				return "addins-popover"
			end try
		end repeat
	end tell
	error "Could not find the Live pane button or the Add-ins button on the ribbon"
end run

on clickAt(x, y)
	set js to "ObjC.import('CoreGraphics'); var pt = $.CGPointMake(" & x & "," & y & "); " & ¬
		"$.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent(null, $.kCGEventMouseMoved, pt, $.kCGMouseButtonLeft)); " & ¬
		"delay(0.2); " & ¬
		"$.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent(null, $.kCGEventLeftMouseDown, pt, $.kCGMouseButtonLeft)); " & ¬
		"delay(0.08); " & ¬
		"$.CGEventPost($.kCGHIDEventTap, $.CGEventCreateMouseEvent(null, $.kCGEventLeftMouseUp, pt, $.kCGMouseButtonLeft));"
	do shell script "osascript -l JavaScript -e " & quoted form of js
end clickAt
