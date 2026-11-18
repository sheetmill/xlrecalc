-- Open one workbook (argument 1, a path inside Excel's container) with alerts off and read, for
-- each cell named (argument 2: "sheet|A1" items separated by ASCII 30), its displayed text and
-- the fill color and boldness Excel shows, conditional formats included (display format), one
-- cell at a time (a bulk read of display format returns missing value). Then close it without
-- saving. A file Excel would offer to repair fails to open with alerts off ("Parameter error
-- (-50)"), so the error is reported as the result. Non-interactive: alerts are off while it runs
-- and restored on every exit path. Opens only the workbook it is given and refuses to run if a
-- workbook of that name is already open.
--
-- Output, fields separated by ASCII 31: "open" or "error"; then for each cell, its text, its
-- interior color as Excel reports it, and whether its font is bold.
on run argv
	set xlPath to item 1 of argv
	set RS to ASCII character 30
	set US to ASCII character 31
	set AppleScript's text item delimiters to RS
	set xlCells to text items of (item 2 of argv)
	set AppleScript's text item delimiters to ""
	set xlName to do shell script "basename " & quoted form of xlPath
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
			set xlOut to "open"
			repeat with xlItem in xlCells
				set AppleScript's text item delimiters to "|"
				set xlParts to text items of (xlItem as text)
				set AppleScript's text item delimiters to ""
				set xlCell to range (item 2 of xlParts) of worksheet (item 1 of xlParts) of xlWb
				set xlShown to display format of xlCell
				set xlColor to color of interior object of xlShown
				set xlBold to bold of font object of xlShown
				if class of xlColor is list then
					set AppleScript's text item delimiters to ","
					set xlColor to xlColor as text
					set AppleScript's text item delimiters to ""
				end if
				set xlOut to xlOut & US & (string value of xlCell) & US & (xlColor as text) & US & (xlBold as text)
			end repeat
			close xlWb saving no
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
