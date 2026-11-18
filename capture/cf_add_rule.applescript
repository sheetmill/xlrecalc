-- Add one formula-driven conditional format to a range of a workbook the probe made and save it:
-- how Excel itself stores a rule (its range and formula), to read back from the file.
-- Arguments: 1 the workbook's path inside Excel's container, 2 the sheet, 3 the range ("D2:D3"),
-- 4 the formula as typed in Excel's dialog ("=D2>2"). The formula is relative to the range's
-- active cell, which is its top-left cell here. Non-interactive: alerts are off while it runs and
-- restored on every exit path; opens only the workbook it is given, refuses to run if a workbook
-- of that name is already open, and closes it after saving.
--
-- Output: "added" with the rule's formula 1 as Excel reports it, or "error", the number and the
-- message (AppleScript may not create a format condition: the answer this script exists to find).
on run argv
	set xlPath to item 1 of argv
	set xlSheet to item 2 of argv
	set xlRange to item 3 of argv
	set ruleFormula to item 4 of argv
	set US to ASCII character 31
	set xlName to do shell script "basename " & quoted form of xlPath
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
		on error errMsg number errNum
			set display alerts to true
			return "error" & US & errNum & US & "open: " & errMsg
		end try
		set xlWb to workbook xlName
		try
			set xlTarget to range xlRange of worksheet xlSheet of xlWb
			select xlTarget
			set xlRule to make new format condition at xlTarget with properties {format condition type:expression, formula 1:ruleFormula}
			set xlOut to "added" & US & (formula 1 of xlRule)
		on error errMsg number errNum
			set xlOut to "error" & US & errNum & US & errMsg
		end try
		try
			if xlOut starts with "added" then
				save xlWb
			end if
			close xlWb saving no
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
		set display alerts to true
	end tell
	return xlOut
end run
