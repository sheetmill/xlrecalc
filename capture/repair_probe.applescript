-- Open a workbook Excel will want to repair (argument 1) WITH alerts on, so Excel shows its
-- repair dialogs and a person answers them; then read column A of the first sheet for rows 1 to
-- argument 2, save a copy of the repaired workbook (argument 3) and close it without saving.
-- capture/repair_probe.py runs it, announced, with the person at the keyboard. Interactive by
-- design: alerts stay on (as they were) on every path. Opens only the workbook it is given.
--
-- Output: the repaired workbook's name, then one line per row, fields separated by ASCII 31:
-- row, formula, value, displayed text.
on run argv
	set xlPath to item 1 of argv
	set xlRows to (item 2 of argv) as integer
	set xlCopy to item 3 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to true
		activate
		with timeout of 900 seconds
			open workbook workbook file name xlPath
		end timeout
		set xlWb to active workbook
		set end of xlOut to "workbook" & US & (name of xlWb)
		set xlWs to worksheet 1 of xlWb
		repeat with xlR from 1 to xlRows
			set xlCell to cell xlR of column 1 of xlWs
			set end of xlOut to (xlR as text) & US & (formula of xlCell) & US & (value of xlCell as text) & US & (string value of xlCell)
		end repeat
		save workbook as xlWb filename xlCopy file format workbook normal file format
		close active workbook saving no
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
