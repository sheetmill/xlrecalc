-- Open one workbook (argument 1, a path inside Excel's container) with alerts off, report what
-- Excel shows in the cells named by the other arguments ("sheet!A1"), ask Excel to calculate once,
-- report them again, and close the workbook without saving. A file Excel would offer to repair
-- fails to open with alerts off ("Parameter error (-50)"), so the error is reported as the result.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Opens only the
-- workbook it is given and refuses to run if a workbook of that name is already open.
--
-- Output, fields separated by ASCII 31: "open" or "error"; for each cell its value and its
-- displayed text after opening; "calculated"; the same pairs after the calculation.
on readCells(xlWb, refs, US)
	set xlOut to ""
	tell application "Microsoft Excel"
		repeat with xlRef in refs
			set AppleScript's text item delimiters to "!"
			set xlSheet to text item 1 of xlRef
			set xlAddr to text item 2 of xlRef
			set AppleScript's text item delimiters to ""
			set xlCell to range xlAddr of worksheet xlSheet of xlWb
			set xlOut to xlOut & US & (value of xlCell as text) & US & (string value of xlCell)
		end repeat
	end tell
	return xlOut
end readCells

on run argv
	set xlPath to item 1 of argv
	set refs to items 2 thru -1 of argv
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
			calculate
			set xlOut to xlOut & US & "calculated" & my readCells(xlWb, refs, US)
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
