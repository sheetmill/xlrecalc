-- For a workbook opened by hand (by Launch Services, as a double-click does), named by argument
-- 1. With "ready": whether it is open and Excel answers a question about it within 5 seconds
-- ("yes", "no", or "blocked" and the error number). Otherwise:
-- read column A of its first sheet for rows 1 to argument 2, save a copy (argument 3) and close
-- the workbook without saving, alerts off for the save and restored on every path.
--
-- Output of a read: one line per row, fields separated by ASCII 31: row, formula, value, text.
on run argv
	set xlName to item 1 of argv
	if (item 2 of argv) is "ready" then
		try
			with timeout of 5 seconds
				tell application "Microsoft Excel"
					if not (exists workbook xlName) then return "no"
					set xlProbe to name of worksheet 1 of workbook xlName
				end tell
			end timeout
			return "yes"
		on error errMsg number errNum
			return "blocked " & errNum
		end try
	end if
	set xlRows to (item 2 of argv) as integer
	set xlCopy to item 3 of argv
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set xlAlerts to display alerts
		try
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			repeat with xlR from 1 to xlRows
				set xlCell to cell xlR of column 1 of xlWs
				set end of xlOut to (xlR as text) & US & (formula of xlCell) & US & (value of xlCell as text) & US & (string value of xlCell)
			end repeat
			set display alerts to false
			save workbook as xlWb filename xlCopy file format workbook normal file format
			close active workbook saving no
			set display alerts to xlAlerts
		on error errMsg number errNum
			set display alerts to xlAlerts
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
