-- Set formulas into one open workbook as a person's formula bar would take them, and read each back.
-- Argument 1: the workbook (Excel's container); argument 2: a UTF-8 file, one formula a line, each in
-- the formula bar's form with its "=", set as AppleScript's legacy formula. Formula k goes into A<k>
-- of the first sheet. A formula Excel refuses leaves the cell unchanged without an error, so the
-- reply is the formula read back, empty for a refusal.
-- Alerts off on every path; the workbook is closed without saving. Opens only the workbook given.
--
-- Output: one line per formula: k, ASCII 31, the formula read back (or "error " and the number).
on run argv
	set xlPath to item 1 of argv
	set xlLines to paragraphs of (read POSIX file (item 2 of argv) as «class utf8»)
	set xlName to do shell script "basename " & quoted form of xlPath
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set xlAlerts to display alerts
		set display alerts to false
		try
			with timeout of 900 seconds
				open workbook workbook file name xlPath
				set xlWs to worksheet 1 of workbook xlName
				set xlK to 0
				repeat with xlF in xlLines
					set xlK to xlK + 1
					if (length of xlF) > 0 then
						set xlCell to cell xlK of column 1 of xlWs
						try
							set formula of xlCell to (xlF as text)
							set end of xlOut to (xlK as text) & US & (formula of xlCell)
						on error errMsg number errNum
							set end of xlOut to (xlK as text) & US & "error " & errNum
						end try
					end if
				end repeat
				close workbook xlName saving no
			end timeout
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
