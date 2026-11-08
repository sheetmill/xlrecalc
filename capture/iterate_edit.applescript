-- Open one workbook (argument 1, a path inside Excel's container) with alerts off, report what
-- Excel shows in the cells named by arguments 4 onward ("sheet!A1"), enter the number in argument 3
-- into the cell named by argument 2, report the cells again, and close without saving.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Opens only the
-- workbook it is given and refuses to run if a workbook of that name is already open.
--
-- Output, fields separated by ASCII 31: "open" or "error"; for each cell its value after
-- opening; "edited"; each value after the edit.
on readCells(xlWb, refs, US)
	set xlOut to ""
	tell application "Microsoft Excel"
		repeat with xlRef in refs
			set AppleScript's text item delimiters to "!"
			set xlSheet to text item 1 of xlRef
			set xlAddr to text item 2 of xlRef
			set AppleScript's text item delimiters to ""
			set xlOut to xlOut & US & (value of range xlAddr of worksheet xlSheet of xlWb as text)
		end repeat
	end tell
	return xlOut
end readCells

on run argv
	set xlPath to item 1 of argv
	set xlTarget to item 2 of argv
	set xlNumber to (item 3 of argv) as real
	set refs to items 4 thru -1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set US to ASCII character 31
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
		on error errMsg number errNum
			set display alerts to true
			return "error" & US & errNum & US & errMsg
		end try
		try
			set xlWb to workbook xlName
			set xlOut to "open" & my readCells(xlWb, refs, US)
			set AppleScript's text item delimiters to "!"
			set xlSheet to text item 1 of xlTarget
			set xlAddr to text item 2 of xlTarget
			set AppleScript's text item delimiters to ""
			set value of range xlAddr of worksheet xlSheet of xlWb to xlNumber
			set xlOut to xlOut & US & "edited" & my readCells(xlWb, refs, US)
			close xlWb saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return xlOut
end run
