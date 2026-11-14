-- Set Excel's calculation mode back to automatic, through a workbook the probe made (argument 1,
-- a path inside Excel's container), and report the mode before and after. Excel's mode belongs to
-- the application and a workbook opened first sets it, so a probe that opened a workbook saved in
-- manual mode ends with this. The workbook is closed without saving.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Opens only the
-- workbook it is given and refuses to run if a workbook of that name is already open.
--
-- Output, fields separated by ASCII 31: "before=<mode>", "after=<mode>".
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set US to ASCII character 31
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlBefore to (calculation as text)
			set calculation to calculation automatic
			set xlAfter to (calculation as text)
			close workbook xlName saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return "before=" & xlBefore & US & "after=" & xlAfter
end run
