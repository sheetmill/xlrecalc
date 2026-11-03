-- Enter formulas as the formula bar takes them and save the workbook: Excel's file form of each.
-- Argument 1: a workbook to open (created by capture/formula_bar.py); argument 2: a UTF-8 file
-- of formulas separated by ASCII 30; argument 3: the copy to save. Formula k goes into cell
-- K<2k-1> of the first sheet ("formula" of the cell, as typed; the rows between hold constants,
-- so Excel cannot join neighbors into shared formulas). Non-interactive: alerts are off
-- while it runs and restored on every exit path; opens only the workbook it is given and
-- refuses to run if a workbook of that name is open; closes it without saving.
--
-- Output: the rows Excel rejected, one per line, as "row<US>error number<US>message".
on run argv
	set xlPath to item 1 of argv
	set xlList to item 2 of argv
	set xlCopy to item 3 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set RS to ASCII character 30
	set US to ASCII character 31
	set xlText to read (POSIX file xlList) as «class utf8»
	set AppleScript's text item delimiters to RS
	set formulaTexts to text items of xlText
	set AppleScript's text item delimiters to ""
	set xlOut to {}
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			repeat with xlR from 1 to count of formulaTexts
				try
					set formula of cell (xlR * 2 - 1) of column 11 of xlWs to (item xlR of formulaTexts)
				on error errMsg number errNum
					set end of xlOut to (xlR as text) & US & errNum & US & errMsg
				end try
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
	set AppleScript's text item delimiters to linefeed
	set xlResult to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlResult
end run
