-- Open one workbook (argument 1, a path inside Excel's container) with alerts off and report
-- what Excel shows, then close it without saving. A file Excel would offer to repair fails to
-- open with alerts off ("Parameter error (-50)"), so the error is reported as the result.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Opens only the
-- workbook it is given and refuses to run if a workbook of that name is already open.
--
-- Output, fields separated by ASCII 31: "open" or "error", then for each of rows 1 to argument
-- 2 of column A: value, displayed text and formula; last, the workbook's saved flag (whether
-- Excel considers it unchanged since opening) and the application's calculation mode.
on run argv
	set xlPath to item 1 of argv
	set xlRows to (item 2 of argv) as integer
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
			set xlWs to worksheet 1 of xlWb
			set xlOut to "open"
			repeat with xlR from 1 to xlRows
				set xlCell to cell xlR of column 1 of xlWs
				set xlOut to xlOut & US & (value of xlCell as text) & US & (string value of xlCell) & US & (formula of xlCell)
			end repeat
			set xlOut to xlOut & US & "saved=" & (saved of xlWb as text) & US & "calc=" & (calculation as text)
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
