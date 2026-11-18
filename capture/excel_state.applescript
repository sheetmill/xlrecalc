-- Report Excel's state after a run, changing nothing: open workbooks and whether alerts are on.
--
--     osascript capture/excel_state.applescript
--
-- Prints "workbooks N, display alerts true|false". Read only: it opens, closes and sets nothing.
tell application "Microsoft Excel"
	set n to count of workbooks
	set alerts to display alerts
end tell
return "workbooks " & n & ", display alerts " & alerts
