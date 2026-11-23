-- Enter formulas as pasted text, as pasting a column into Excel by hand enters them.
-- Argument 1: a workbook (Excel's container) to open; argument 2: a UTF-8 file of lines, one cell
-- each from A1 down (blank lines leave cells empty); argument 3: the copy to save.
-- The text goes on the clipboard and Excel pastes it into A1 of the first sheet; then the
-- workbook is saved as the copy and closed without saving. Alerts off on every path. The
-- caller saves the clipboard's text before and restores it after.
on run argv
	set xlPath to item 1 of argv
	set xlText to read POSIX file (item 2 of argv) as «class utf8»
	set xlCopy to item 3 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set the clipboard to xlText
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set xlAlerts to display alerts
		set display alerts to false
		try
			with timeout of 900 seconds
				open workbook workbook file name xlPath
				set xlWs to worksheet 1 of workbook xlName
				activate object xlWs
				paste worksheet xlWs destination (range "A1" of xlWs)
				calculate
				save workbook as (workbook xlName) filename xlCopy file format workbook normal file format
				close active workbook saving no
			end timeout
			set display alerts to xlAlerts
		on error errMsg number errNum
			set display alerts to xlAlerts
			error errMsg number errNum
		end try
	end tell
	return "pasted"
end run
