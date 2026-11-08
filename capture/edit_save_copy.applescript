-- Open a workbook (argument 1), enter the number in argument 4 into the cell named by argument 3
-- ("sheet!A1"), save it as an .xlsx workbook under another name (argument 2), close that copy.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Touches no other
-- open workbook, and refuses to run if a workbook of either name is open.
on run argv
	set xlPath to item 1 of argv
	set xlCopy to item 2 of argv
	set xlTarget to item 3 of argv
	set xlNumber to (item 4 of argv) as real
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set AppleScript's text item delimiters to "!"
	set xlSheet to text item 1 of xlTarget
	set xlAddr to text item 2 of xlTarget
	set AppleScript's text item delimiters to ""
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set value of range xlAddr of worksheet xlSheet of xlWb to xlNumber
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
