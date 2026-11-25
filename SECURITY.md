# Security policy

## Scope

xlrecalc evaluates the formulas of workbooks it is given, which may be hostile, and writes their
values back. It never opens a connection: functions that would reach outside the workbook are not
computed. Its own security problems are mostly about resources: a workbook or formula that makes
it hang, or use time or memory out of proportion to its size, or a limit that does not hold.

## Reporting a vulnerability

Please report security problems privately, through GitHub's private vulnerability reporting:
open the repository's **Security** tab and choose **Report a vulnerability**
(https://github.com/sheetmill/xlrecalc/security/advisories/new). Do not open a public issue.

Useful in a report: the version, a workbook or formula that reproduces the problem (a crafted
workbook is fine), and what happens (a crash, a hang, time or memory out of proportion).

A problem in reading the file itself (a zip bomb, a malformed part, a limit of the reader that
does not hold) belongs to xlio, which xlrecalc reads workbooks with: please report it there
(https://github.com/sheetmill/xlio/security/advisories/new). A problem in number formatting, which
xlrecalc's `TEXT` uses, belongs to xlformat
(https://github.com/sheetmill/xlformat/security/advisories/new), and one in rendering workbooks to
sheetpeek (https://github.com/sheetmill/sheetpeek/security/advisories/new). If you are not sure
which package is at fault, report it here.

## Supported versions

xlrecalc is in alpha. Fixes go into the next release; only the latest release receives fixes.

## What happens after a report

The report is acknowledged, reproduced and fixed in the next release, whose changelog states the fix in its Security
section. Reporters are credited there unless they prefer not to be.
