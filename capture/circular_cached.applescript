-- Open one workbook (argument 1) with alerts off, do one thing to it (argument 3: none, calculate,
-- full, rebuild, or an edit: set_input, set_formula, break, new_cycle, new_dependent), report
-- column A of its first sheet, rows 1 to argument 2, then save a copy under argument 4 if one is
-- named, and close without saving. Non-interactive: alerts are restored on every
-- exit path. Refuses a workbook of a name already open, and touches no other.
--
-- Output, fields separated by ASCII 31: "open" or "error", then value and displayed text per row.
on run argv
	set xlPath to item 1 of argv
	set xlRows to (item 2 of argv) as integer
	set xlAction to item 3 of argv
	set xlCopy to item 4 of argv
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
			if xlAction is "calculate" then
				calculate
			else if xlAction is "full" then
				calculate full
			else if xlAction is "rebuild" then
				calculate full rebuild
			else if xlAction is "set_input" then
				set value of range "B1" of xlWs to 5
			else if xlAction is "set_formula" then
				set formula of range "A2" of xlWs to "=A1+2"
			else if xlAction is "break" then
				set formula of range "A2" of xlWs to "=7"
			else if xlAction is "new_cycle" then
				set formula of range "A5" of xlWs to "=A5+1"
			else if xlAction is "new_dependent" then
				set formula of range "A6" of xlWs to "=A1*3"
			end if
			set xlOut to "open"
			repeat with xlR from 1 to xlRows
				set xlCell to cell xlR of column 1 of xlWs
				set xlOut to xlOut & US & (value of xlCell as text) & US & (string value of xlCell)
			end repeat
			if xlCopy is not "" then
				save workbook as xlWb filename xlCopy file format workbook normal file format
				set xlCopyName to do shell script "basename " & quoted form of xlCopy
				close workbook xlCopyName saving no
			else
				close xlWb saving no
			end if
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
