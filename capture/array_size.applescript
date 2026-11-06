-- One step of the array-size probe (capture/array_size_probe.py), so that the caller can time
-- each step and watch Excel's memory between them. Argument 1 is the step:
--   count             the number of open workbooks
--   open PATH         open a workbook the probe made, alerts off (they stay off until "close")
--   read NAME         the value and the formula of A1 on sheet S of the open workbook NAME
--   calc              a full recalculation of the open workbooks
--   close NAME        close NAME without saving, and restore alerts
--   restore           restore alerts (after a step was stopped)
-- Non-interactive; opens only the workbook it is given, refuses one of the same name already open.
on run argv
	set xlStep to item 1 of argv
	tell application "Microsoft Excel"
		if xlStep is "count" then
			return (count of workbooks) as text
		else if xlStep is "open" then
			set xlPath to item 2 of argv
			set xlName to do shell script "basename " & quoted form of xlPath
			if exists workbook xlName then error "already open in Excel, close it first: " & xlName
			set display alerts to false
			try
				open workbook workbook file name xlPath
			on error errMsg number errNum
				set display alerts to true
				return "error " & errNum
			end try
			return "opened"
		else if xlStep is "read" then
			set xlCell to range "A1" of worksheet "S" of workbook (item 2 of argv)
			return (value of xlCell as text) & " | " & (formula of xlCell)
		else if xlStep is "calc" then
			calculate full
			return "calculated"
		else if xlStep is "close" then
			try
				close workbook (item 2 of argv) saving no
			end try
			set display alerts to true
			return "closed"
		else if xlStep is "restore" then
			set display alerts to true
			return "restored"
		end if
	end tell
	error "unknown step: " & xlStep
end run
