-- Open a workbook (argument 1), save it as an .xlsx workbook under another name (argument 2),
-- close that copy. Excel recalculates formulas without cached values when it opens the file, so
-- the copy holds Excel's own cached values and XML. Non-interactive: alerts are off while it
-- runs and restored on every exit path. Touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlCopy to item 2 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			save workbook as xlWb filename xlCopy file format workbook normal file format
			close workbook xlCopyName saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			try
				close workbook xlCopyName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return "ok"
end run
