-- Enter formulas through a cell's "value" (argument 4 = "value") or "formula local" (any other
-- word), as typing does, and save a copy: which formulas Excel marks as dynamic arrays.
-- Argument 1: a workbook to open (created by capture/marker_probe.py); argument 2: a UTF-8 file
-- of formulas separated by ASCII 30; argument 3: the copy to save. Formula k goes into cell
-- A<2k-1> of the first sheet. Non-interactive: alerts are off while it runs and restored on
-- every exit path; opens only the workbook it is given and refuses to run if a workbook of
-- that name is open; closes it without saving.
on run argv
	set xlPath to item 1 of argv
	set xlList to item 2 of argv
	set xlCopy to item 3 of argv
	set xlWay to item 4 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set RS to ASCII character 30
	set xlText to read (POSIX file xlList) as «class utf8»
	set AppleScript's text item delimiters to RS
	set formulaTexts to text items of xlText
	set AppleScript's text item delimiters to ""
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			repeat with xlR from 1 to count of formulaTexts
				set xlCell to cell (xlR * 2 - 1) of column 1 of xlWs
				if xlWay is "value" then
					set value of xlCell to (item xlR of formulaTexts)
				else
					set formula local of xlCell to (item xlR of formulaTexts)
				end if
			end repeat
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
