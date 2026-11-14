-- Quit Excel, for a capture that needs a fresh session, only when no workbook is open: a person's
-- open workbook is never closed by it. Returns "quit", or an error naming the open workbooks.
on run
	tell application "Microsoft Excel"
		set xlCount to count of workbooks
		if xlCount > 0 then error "workbooks are open in Excel (" & xlCount & "): not quitting"
		quit saving no
	end tell
	return "quit"
end run
