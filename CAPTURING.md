# Capturing Excel's results

Expected results are what Excel 16 for Mac computes. This document describes how they are
captured, the facts about Excel and openpyxl the method relies on, and how formulas from other
people's workbooks are used.

## 1. What exists

| suite | rows | what |
|---|---|---|
| `number_text` | 324 | a number as text through `&` and through `TEXT(x,"General")` |
| `text_number` | 516 | `VALUE`, `+0`, `DATEVALUE`, `TIMEVALUE` over numbers, dates and times written as text |
| `compare` | 1,045 | operators across every pair of kinds, 15-digit equality, the text `"#N/A"` against the error in operators, lookups and criteria, blank against empty text |
| `cancellation` | 703 | when an addition or subtraction gives 0, and subnormal results |
| `dates`, `dates_1904` | 57 each | `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `DATEVALUE` around the 1900 leap day and the limits |
| `sums` | 32 | summation order and precision: `SUM` over a range and over arguments, `+` chains, `AVERAGE` |
| `arrays` | 311 cells | legacy formulas inside and outside a range's rows, legacy array formulas over declared ranges, dynamic arrays and their spills, with Excel's metadata and rich-value parts ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.7) |
| `marker_typed` | 10 | which formulas typed in Excel's formula bar it marks as dynamic arrays (typed by a person; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.7) |
| `marker_value` | 14 | the same entered through AppleScript's `value` of a cell: legacy formulas, never marked |
| `dynrefs` | 90 | `OFFSET`, `INDIRECT`, `CELL`, `INDEX` as a reference, and dependencies only evaluation finds ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §6.3) |
| `cell_noref` | 36 | `CELL` without a reference in workbooks differing in the active sheet and cells |
| `n_lookup` | 56 | `N` and `LOOKUP` across argument kinds, both forms of `LOOKUP`, short result ranges |
| `finance` | 931 | `PV`, `FV`, `PMT`, `IPMT`, `PPMT`, `NPER`, `NPV`, `XNPV`, `SLN`, `VDB` over signs, sizes, fractional and negative periods, both payment types and argument kinds; cash flows and dates as fixed cells |
| `iterative` | 271 | `IRR`, `XIRR`, `RATE` over 18 cash-flow shapes (two roots, a double root, no sign change, huge and tiny rates), 11 guesses, date orders and argument kinds |
| `bonds` | 2,377 | `DAYS360` and `YEARFRAC` over every pair of 14 dates (month ends, February in leap and common years) in all bases; the six `COUP` functions, `PRICE`, `YIELD`, `DURATION`, `MDURATION` over six bonds, three frequencies and five bases; argument errors |
| `workdays` | 481 | `WORKDAY`, `NETWORKDAYS` and their `.INTL` forms over starts on weekdays, weekends and holidays, every weekend code and masks, holidays with duplicates, text and errors; `HOUR`, `MINUTE`, `SECOND` around rounding to the second |
| `statistics` | 800 | variances, standard deviations, the A forms, `MEDIAN`, `SUMSQ` over spread, mixed kinds, values near 1E9, 1E-300 and 1E300; `CORREL`, `LINEST` with every statistic, `RANK` with ties; the normal distribution over z from -40 to 40 and p from 1E-300 |
| `misc` | 611 | `TRANSPOSE` in legacy formulas, `ADDRESS` in every style and sheet quoting, `REPT`, `REPLACE`, `CHAR` 0 to 256, `CODE`, `RAND` and `RANDBETWEEN` by property, the five D functions over eleven criteria ranges and every kind of field |
| `textfmt` | 1,414 | `TEXT` over 24 values in 39 formats (sections, conditions, fills, `_x` spacing, fractions, dates and times, refused formats, the empty format) and argument kinds; `DOLLAR` and `FIXED` over signs, magnitudes, ties and decimals from -2.9 to 128 |
| `rounding` | 1,836 | `ROUND`, `ROUNDUP`, `ROUNDDOWN` over ties, near-ties, 15- to 17-digit values and magnitudes from 1E-300 to 1E21 in both signs, digits -2 to 10 |
| `literals` | 126 | number literals of 15 to 22 significant digits in formulas, in seven contexts |
| `functions` | 1,180 | every implemented function across argument kinds; lookup tables as fixed cells |
| `formula_forms` | 249 | how Excel writes formulas back to a file (§3) |
| `repair` | 17 | what Excel's repair does to formulas it refuses in a file (interactive; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §9.5) |
| `repair_uncached` | 17 | the same for refused formulas without a cached value (interactive) |
| `formula_bar` | 3,758 | Excel's file form of formulas typed in the formula bar: the corpus in display form, and calls of about 500 functions with 0 to 6 arguments ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §5.2) |
| `formula_bar_ranges` | 408 | the same for functions that need range arguments |
| `prefixes` | 32 | functions written into a file with and without `_xlfn.`: the form Excel computes |
| `whitespace` | 15 | whitespace at the ends of inline and shared strings without `xml:space="preserve"` ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §9.1) |
| `value_forms` | 26 | number forms Excel never writes in `<v>` (`nan`, `INF`, `1e400`, `1_000`, `+5`...) and rows and columns off its grid: which files Excel refuses and how it reads the others (`capture/value_probe.py`) |
| `array_size` | 6 | declared arrays over one, ten and every column, and `SUMPRODUCT` over another sheet's whole columns: whether Excel opens and recalculates them, how long it takes and what memory it uses (`capture/array_size_probe.py`; it ends Excel past 8 GB) |
| `array_refs` | 7,042 cells | `ROW`, `COLUMN`, `CELL`, ranges and arrays as a criterion or a lookup key, and arrays of references, each as a legacy formula, a legacy array formula and a dynamic array |
| `array_functions` | 41 | the `SUMIF` family, `DSUM`, `CORREL` and `ROW` times `COLUMN` over one column, ten and every column of another sheet, with Excel's memory: Excel's value, or the file's value kept |
| `circular`, `circular_alerts` | 16 each | circular references with iteration off: what Excel saves and shows, alerts off and on (the second interactive; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.2) |

All live in `tests/captures/` as TSV files with a header: suite, Excel build and locale, date
system, capture date, generator and method. `tests/test_captures.py` and
`tests/test_formula_forms.py` compare the package with them: each suite becomes one engine
sheet laid out as Excel's workbook was, and every formula cell must equal what Excel saved.

## 2. The pipeline

```
capture/pipeline.py    builds a suite's workbook, has Excel calculate and save it, reads it back
capture/suites.py      the suites: Case(tag, formula with {1}, {2}... for inputs, inputs),
                       and cells written once for a suite (lookup tables)
capture/save_copy.applescript   opens a workbook and saves a copy, alerts off
```

`uv run python -m capture.pipeline SUITE ...` (or `all`) builds a workbook with openpyxl inside
Excel's sandbox container (`~/Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc/`),
one case per row: the formula in column A, its inputs in B onward. Excel opens it, calculates
everything (openpyxl's files ask for that), and saves a copy; the script reads column A's cached
value and type from the copy's sheet XML and writes `tests/captures/<suite>.tsv`. Reading the
file Excel saved, not AppleScript's values, gives the exact double (Excel writes 17 significant
digits) and the real error code (AppleScript returns `missing value` for an error). New cases
go at the end of their suite so earlier case numbers stay stable; a recapture that changes any expected result is reviewed as a change of evidence.

**Crafted and extreme input.** A test or script that feeds the package crafted or extreme input
(hostile packages, huge ranges, extreme numbers) runs in a child process under the memory watchdog
of `tests/watchdog.py`, which kills it at 4 GB: `bounded` for a test, `uv run python -m
tests.watchdog -- COMMAND` for anything else. A time limit alone does not bound memory, and
`ulimit -v` is no limit on macOS; a sweep without the watchdog once used 82 GB on a 24 GB Mac,
and macOS paused every application, Excel included. A killed run (return code 137, or a missing
result) is a finding: find out why before running it again. No such sweep runs while Excel
captures, and a capture starts only once Excel answers (for example `osascript -e 'tell
application "Microsoft Excel" to get version'`).

**Adding a function or a family of functions.** Write its cases at the end of a suite in
`capture/suites.py`, or as a suite of its own (each case a formula with `{1}`, `{2}`... for its
inputs; fixed lookup tables through a suite's `fixed` cells), across argument kinds, edges and
errors; run `uv run python -m capture.pipeline SUITE`; implement it under
`src/xlrecalc/_functions/` with its registry kinds (§8.1 of [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md)); add a new suite
to `SUITES` in `tests/test_captures.py`. A row that cannot be matched needs its reason in
`KNOWN_DEVIATIONS` there and in `QUIRKS.md`. The list is strict: a listed row that matches Excel
fails the test, so an entry cannot outlive its cause; a random result is exempt
(`EXEMPT`), and a row listed because the platform's math library differs from Excel's in the
last bits (`MATH_LIBRARY`) must differ only on the platform `docs/coverage.md` records. A function whose results stay a few units in the
last place from Excel's after a time-boxed attempt at bit-exactness declares a tolerance in the
registry instead ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1), listed in `QUIRKS.md` too. Iterative functions (`IRR`, `XIRR`, `RATE`,
`YIELD`) get many inputs, so that a rule for their iteration can be fitted. With
`XLRECALC_CAPTURE_KEEP=1` the pipeline keeps the copy Excel saved, to read the formulas as
Excel wrote them (the file form of a function, for example).

Driving Excel: scripts are files under `capture/`, non-interactive, with `display alerts` off
and restored on every exit path (the setting otherwise persists until Excel quits, and Excel
then silently fails to open or repair files opened by hand); they open only workbooks they
create, refuse to run when a workbook of the same name is open, and never use the clipboard.
With alerts off, a file Excel would offer to repair fails to open with error -50 ("Parameter
error"), which is how probes detect a repair prompt without a person at the keyboard.
`capture/repair_probe.py` and `capture/circular_probe.py --interactive` are the interactive
captures: they open a file with alerts on and wait for a person to accept the repair or
dismiss the circular-reference warning. `capture/marker_probe.py --typed` reads a workbook a
person typed formulas into and saved (AppleScript can only enter legacy formulas).

**Facts the method relies on** (openpyxl 3.1.5, Excel 16.114.3):

1. openpyxl writes floats with `%.16g`, which names a neighboring double for about half of all
   floats. The pipeline writes `repr(value)` as the cell's text with type `n`; Excel parses it
   correctly rounded.
2. openpyxl writes text as inline strings and adds `xml:space="preserve"` only when the text
   has leading or trailing whitespace and other characters: a whitespace-only text such as
   `" "` is written without it, and Excel reads it as an empty text. An empty text is not
   written at all, so the cell is blank. The pipeline writes both as formulas (`=" "`,
   `=""`).
3. openpyxl writes a formula without a cached value as `<f>…</f><v />`, writes no
   `calcChain.xml`, and sets `<calcPr calcId="124519" fullCalcOnLoad="1"/>`, so Excel
   calculates every formula on open.
4. Excel saves numbers with up to 17 significant digits, text results as `t="str"`, logicals
   as `t="b"` with `0`/`1`, errors as `t="e"` with the code; it writes `calcChain.xml` again and
   drops `fullCalcOnLoad`.
5. Texts without a year (`"3/4"`) take the current year, so the header's capture date is part of
   the evidence for those rows.
6. Excel decodes OOXML's `_xHHHH_` escapes in text, cached text results and formula text
   (`capture/escape_probe.py`): `a_x000D_b` holds a carriage return, `_x005F_x0041_` is the text
   `_x0041_`. openpyxl writes a formula's text as given, so a string literal containing such a
   sequence (`="_x0041_"`) reaches Excel as `="A"`; capture formulas never contain one.
7. Excel drops the leading and trailing whitespace (spaces, tabs, line feeds) of a formula's
   text result in `<v>`, and of each `<t>` of a shared or inline string, unless the element
   says `xml:space="preserve"` (`tests/test_excel_live.py`, capture `whitespace`); it keeps
   whitespace inside, and an escaped carriage return (`_x000D_`) at either end.

Files the package writes are opened by Excel in `tests/test_excel_live.py`, which runs only with
`pytest --tier excel`, under the same rules as the captures.

The probes behind the design (`capture/writeback_probe.py`, `capture/semantics_probe.py`,
`capture/formula_forms.py`) use the same conventions; their results are in [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md)
§5.2 and §9.4.

## 3. Formulas as Excel writes them

`capture/formula_forms.py` writes every formula of `tests/formula_corpus.py` (external references
and hostile sheet names left out) into a workbook with the sheets, names and table they refer
to, a constant between formulas so Excel cannot join them into shared formulas, and compares
the text Excel saves with what was written. A workbook Excel refuses is split in halves until
each refused formula stands alone. The results, and what Excel normalizes, are in
[ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §5.2.

`capture/formula_bar.py` types formulas into cells instead, through AppleScript's `formula` of
a cell (the formula bar's form), and reads the file form from the copy Excel saves. A formula
the formula bar refuses leaves the cell as it was, without an error, so each cell starts with
a placeholder formula. AppleScript sets the legacy formula and at most 255 characters of it.
`capture/prefix_probe.py` settles the prefixes the formula bar could not, by what Excel
computes; `tools/prefix_table.py` turns the three captures into the printer's prefix table.

## 4. Formulas from other workbooks

Sources analyzed: the Enron spreadsheet corpus (Hermans and Murphy-Hill, MSR 2015, CC BY 4.0;
15,871 workbooks, 2.1 million distinct formulas) and the test workbooks of xlcalculator (MIT).
They are used locally, to find the formula forms the project's own corpus must cover and to
rank functions by use; nothing from them is committed except the aggregate ranking in
`tools/function_frequency.tsv`, with its attribution. Rules:

- Expected values never come from another program or from another file's cached values: a
  formula worth capturing is put into a fresh workbook and Excel computes it.
- A downloaded workbook is never opened in Excel. It is read only by this package's reader or
  openpyxl, in a directory of its own, and only `.xlsx` files are read (never `.xlsm`).
- Before a formula from such a workbook goes into a workbook for Excel, `capture/corpus_filter.py`
  drops and counts DDE links, references to other workbooks, and functions that reach outside
  the workbook (files, network, programs, the machine) or `HYPERLINK`.
- Third-party files and their contents are never committed. The committed formula corpus
  (`tests/formula_corpus.py`) is written for this project and covers every form the analysis
  found (`tools/formula_census.py` lists the forms a corpus has and ours lacks).
- `tools/fetch_enron.py` fetches the Enron archive, pinned to its URL and checksums, into
  `../corpora/enron/` outside the repository.

Local checks, which do nothing when the corpus is absent:
`tools/formula_roundtrip.py FILE_OR_DIRECTORY ...` reads the `<f>` elements of every worksheet
with the standard library and round-trips each distinct formula through the parser and printer
(on the Enron corpus: 2,126,806 of 2,126,809, the other three being text that is not a formula);
`tools/formula_census.py` compares forms; `tools/function_frequency.py` writes the ranking.
