# xlrecalc

xlrecalc recalculates Excel workbooks in pure Python. It evaluates a workbook's formulas with
the results Excel computes and writes them back into the file as the cached values Excel
stores, so that openpyxl's data-only mode, pandas and other readers see values where a
workbook written by a program has only formulas. Results are tested against values captured
from Excel 16 for Mac under Microsoft 365, the default. With `platform="windows"`, what Excel
for Windows computes differently (the character set of `CHAR` and `CODE`, the order of text, the
names it knows, `ENCODEURL`) is tested against values captured on Excel for Windows; the rest is
not yet checked there. Excel in other languages and locales is not yet supported.

```bash
pip install xlrecalc
```

Python 3.11 or later. The source is at <https://github.com/sheetmill/xlrecalc>. The design,
which the Quirks and Capturing pages cite by section (§), is described in
[ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md).

xlrecalc is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats,
[xlio](https://github.com/sheetmill/xlio), which reads and surgically writes xlsx files, and
[sheetpeek](https://github.com/sheetmill/sheetpeek), which shows workbooks as Excel does.

```{toctree}
:maxdepth: 2

usage
iteration
coverage
quirks
glossary
capturing
api
changelog
```
