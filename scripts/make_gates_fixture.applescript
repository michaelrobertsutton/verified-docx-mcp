-- Word-authored fixture for the #41/#42/#43/#45 gates.
on run argv
	set outPath to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Word/Data/Documents/gates.docx"
	tell application "Microsoft Word"
		set d to make new document
		set t to "Gate alpha paragraph." & return & "Gate beta paragraph." & return & "Lead-in bold. Plain sentence follows here." & return & "Twin sentence." & return & "Twin sentence." & return & "Comment target one." & return & "Comment target two." & return & "Delete me paragraph." & return & "Neighbor paragraph." & return & "Table anchor paragraph." & return & "Final paragraph."
		set content of text object of d to t
		-- bold lead-in "Lead-in bold." (paragraph 3)
		set r3 to create range d start 43 end 56
		set bold of font object of r3 to true
		-- a 2x2 table after "Table anchor paragraph." (ends at offset 220)
		try
			set tr to make new table at d with properties {text object:(create range d start 220 end 220), number of rows:2, number of columns:2}
			set content of text object of (cell 1 of row 1 of tr) to "Cell one"
			set content of text object of (cell 2 of row 1 of tr) to "Cell two"
			set content of text object of (cell 1 of row 2 of tr) to "Cell three"
			set content of text object of (cell 2 of row 2 of tr) to "Cell four"
		on error e
			log "table: " & e
		end try
		save as d file name outPath file format format document default
	end tell
	return outPath
end run
