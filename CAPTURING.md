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
| `arrays` | 310 cells | legacy formulas inside and outside a range's rows, legacy array formulas over declared ranges, dynamic arrays and their spills, with Excel's metadata and rich-value parts ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.7) |
| `dyn_sequence`, `dyn_mmult`, `dyn_unique`, `dyn_sort`, `dyn_sortby`, `dyn_filter`, `dyn_choose`, `dyn_droptake`, `dyn_stack`, `dyn_tocol`, `dyn_wrap`, `dyn_index`, `dyn_matrix`, `dyn_frequency`, `dyn_textsplit`, `dyn_randarray`, `dyn_let` | 644 cases | the functions that make arrays as dynamic arrays (`capture/dynamic_probe.py`, cases in `capture/dynamic_suites.py`): two sets of values each; mixed kinds, case, several keys, descending, `by_col`, equal keys, `exactly_once`, empty results and `if_empty`, text and blanks in a matrix, sizes that do not fit, arrays as arguments, a constant in a result's way (`#SPILL!`), `#` references, a plain formula without the marker, bare names; for the reshaping functions, positive, negative, zero and fractional counts and indexes, counts past the size, arrays and ranges of counts, padding left out, empty and given, errors given directly and inside an array, plain formulas over whole columns (which show that `TAKE` and `DROP` of a range give a range); for `TEXTSPLIT`, delimiters that overlap: two starting at one place in either order, a shorter one starting first, one inside another, column and row delimiters overlapping or the same, with and without regard to case (the text is read once from the start; of delimiters starting at one place a column one is taken before a row one, whatever their lengths, and the longest of one kind); for `LET`, with defined names in the workbook (`NAMES`): scoping, nesting and case, a defined name shadowed and a name without `_xlpm.` (the defined name), a name bound later (`#NAME?`), names shaped like a function, a cell or with a dot, a binding never used that reads the formula's own cell (not evaluated), one used twice (evaluated once), references bound, spilled and intersected, errors in a binding, no calculation after the bindings (Excel writes the last name as the calculation); the forms Excel refuses in a file (`REFUSED`; `capture/refusal_probe.py` opens each alone in a workbook) |
| `number_literals` | 25 | which number literals Excel accepts in a file formula, each alone in a workbook (`capture/literal_probe.py`): the largest, 9.99999999999999E+307, and past it, the 15-digit cut, the smallest double and below it, two sets of mantissas |
| `marker_typed` | 75 | which formulas typed in Excel's formula bar it marks as dynamic arrays: the functions that make arrays with a one-cell result, wrapped in another function, under an operator, a branch and a function taken element by element, and the functions of a reference with sizes in the text or not (typed by a person in four workbooks, kept as `marker_typed.xlsx` and `marker_typed2.xlsx` to `marker_typed4.xlsx`, without who saved them or where, and checked against the capture; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.7) |
| `let_typed` | 8 | `LET` typed in Excel's formula bar (`capture/lambda_probe.py --typed`): its file form (`_xlfn.LET`, `_xlpm.` on each name), which formulas Excel marks as dynamic arrays (only one whose result spills: not `LET(x,SEQUENCE(3),SUM(x))`, nor a single value bound to a range or a cell) and the values. The same workbook typed `LAMBDA` and its helpers, which Office Home & Business 2021 for Mac does not have (saved without prefixes, `#NAME?`); those rows are not part of the capture |
| `marker_value` | 14 | the same entered through AppleScript's `value` of a cell: legacy formulas, never marked |
| `dynrefs` | 90 | `OFFSET`, `INDIRECT`, `CELL`, `INDEX` as a reference, and dependencies only evaluation finds ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §6.3) |
| `cell_noref` | 36 | `CELL` without a reference in workbooks differing in the active sheet and cells |
| `n_lookup` | 56 | `N` and `LOOKUP` across argument kinds, both forms of `LOOKUP`, short result ranges |
| `finance` | 1,248 | `PV`, `FV`, `PMT`, `IPMT`, `PPMT`, `NPER`, `NPV`, `XNPV`, `SLN`, `VDB` over signs, sizes, fractional and negative periods, both payment types and argument kinds; cash flows and dates as fixed cells; at the end, `VDB` over fractional start and end periods around a corpus call, with `no_switch` both ways, on two sets of values; `VDB` from fractional starts around the switch to straight line on two assets, with each asset's whole periods |
| `iterative` | 402 | `IRR`, `XIRR`, `RATE` over 18 cash-flow shapes (two roots, a double root, no sign change, huge and tiny rates), 11 guesses, date orders and argument kinds; `IRR`, `NPV`, `MIRR` and `XIRR` over two sets of flows holding an error (four kinds, and two at once), a text, a number as text, a blank or a logical at the start, middle or end, and `XIRR` over dates holding them |
| `bonds` | 2,377 | `DAYS360` and `YEARFRAC` over every pair of 14 dates (month ends, February in leap and common years) in all bases; the six `COUP` functions, `PRICE`, `YIELD`, `DURATION`, `MDURATION` over six bonds, three frequencies and five bases; argument errors |
| `workdays` | 481 | `WORKDAY`, `NETWORKDAYS` and their `.INTL` forms over starts on weekdays, weekends and holidays, every weekend code and masks, holidays with duplicates, text and errors; `HOUR`, `MINUTE`, `SECOND` around rounding to the second |
| `statistics` | 800 | variances, standard deviations, the A forms, `MEDIAN`, `SUMSQ` over spread, mixed kinds, values near 1E9, 1E-300 and 1E300; `CORREL`, `LINEST` with every statistic, `RANK` with ties; the normal distribution over z from -40 to 40 and p from 1E-300 |
| `misc` | 611 | `TRANSPOSE` in legacy formulas, `ADDRESS` in every style and sheet quoting, `REPT`, `REPLACE`, `CHAR` 0 to 256, `CODE`, `RAND` and `RANDBETWEEN` by property, the five D functions over eleven criteria ranges and every kind of field |
| `textfmt` | 1,416 | `TEXT` over 24 values in 39 formats (sections, conditions, fills, `_x` spacing, fractions, dates and times, refused formats, the empty format) and argument kinds; `DOLLAR` and `FIXED` over signs, magnitudes, ties and decimals from -2.9 to 128 |
| `rounding` | 1,836 | `ROUND`, `ROUNDUP`, `ROUNDDOWN` over ties, near-ties, 15- to 17-digit values and magnitudes from 1E-300 to 1E21 in both signs, digits -2 to 10 |
| `literals` | 126 | number literals of 15 to 22 significant digits in formulas, in seven contexts |
| `functions` | 1,238 | every implemented function across argument kinds; lookup tables as fixed cells |
| `formula_forms` | 249 | how Excel writes formulas back to a file (§3) |
| `repair` | 17 | what Excel's repair does to formulas it refuses in a file (interactive; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §9.5) |
| `repair_uncached` | 17 | the same for refused formulas without a cached value (interactive) |
| `formula_bar` | 3,758 | Excel's file form of formulas typed in the formula bar: the corpus in display form, and calls of about 500 functions with 0 to 6 arguments ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §5.2) |
| `formula_bar_ranges` | 408 | the same for functions that need range arguments |
| `prefixes` | 62 | functions written into a file with and without `_xlfn.`: the form Excel computes |
| `whitespace` | 15 | whitespace at the ends of inline and shared strings without `xml:space="preserve"` ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §9.1) |
| `value_forms` | 26 | number forms Excel never writes in `<v>` (`nan`, `INF`, `1e400`, `1_000`, `+5`...) and rows and columns off its grid: which files Excel refuses and how it reads the others (`capture/value_probe.py`) |
| `array_size` | 6 | declared arrays over one, ten and every column, and `SUMPRODUCT` over another sheet's whole columns: whether Excel opens and recalculates them, how long it takes and what memory it uses (`capture/array_size_probe.py`; it ends Excel past its time limit or 8 GB, one of the two exceptions to never ending Excel in §2, and starts only with no workbook open) |
| `modern` | 2,230 | `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `LARGE`, `SMALL`, `PERCENTILE`, `QUARTILE`, `MODE`, `IFS`, `SWITCH`, `TEXTJOIN`, `CONCAT`, `XLOOKUP`, `XMATCH` across argument kinds, edges and modes; how text compares, every pair of 39 characters |
| `iteration` | 1,337 | `IRR`, `RATE`, `XIRR` from 20 guesses over flows of many shapes, and over a double root with guesses 2^k away, which shows the iteration's limit and stopping test; `YIELD` over a grid of bonds. Evidence for inferring Excel's iteration: not compared by `tests/test_captures.py` yet |
| `iterate_open` | 55 | what Excel does on opening a file with iterative calculation on: a counter, a loop that converges, a dependent, a control and a loop with no cached values, across settings, a file `recalculate` wrote and a copy Excel saved (`capture/iterate_open_probe.py`) |
| `iterate_calc` | 86 cells | how Excel iterates (49 workbooks and 2 edits, `capture/iterate_probe.py`): runs of `iterateCount` on opening, on `calculate`, on saving and after an edit; the order of cells within a sheet and across sheets, with sheet names in case, digits and accents; a precedent outside a loop, a dependent after it and above it; `iterateDelta` at, above and below a change, over large and small magnitudes, and with a change between kinds; one loop or two; loops through `INDIRECT` and `OFFSET`; oscillation, overflow, a loop already converged |
| `calc_chain` | 24 | whether `calcChain.xml` sets the order of iteration (`capture/calc_chain_probe.py`): a loop within a sheet and one across sheets, saved by Excel, then with its chain as Excel wrote it, reversed, or removed, each opened, calculated and saved, and the saved copy opened again. The chain does not: the order is the same in all, and Excel writes its own chain back |
| `iterate_uncached` | 84 | iteration in files whose formulas have no cached values (`capture/iterate_uncached_probe.py`): 17 loops, each with and without `fullCalcOnLoad` and with every formula cached 0; after opening, after `calculate`, saved |
| `iterate_opening` | 72 | what opening does to a loop with no cached values (`capture/iterate_opening_probe.py`, 32 files as openpyxl writes them): the same loops in other row and column orders, entered from a constant or a formula or not, IFERROR in each cell in turn, opening's counts 1 to 3, and 8 loops from an earlier probe of openpyxl files read through the public API |
| `iterate_cached` | 6 | which Excel actions run a loop with cached values (`capture/iterate_cached_probe.py`): Excel's saved copy `X` of an openpyxl loop (kept, its markers removed, as `iterate_cached.xlsx`, the input of the end-to-end tests); `X` opened and saved (two runs); `X` opened, a cell nothing reads edited, and saved (three runs) |
| `workbook_name` | 7 | what `[0]!Name` means (`capture/workbook_name_probe.py`): the workbook-level name, from either sheet, even one with a sheet-level name of the same name; a missing one is `#NAME?`; Excel writes `[0]!` back as it was. `[1]!Name` with no external link makes Excel refuse the file |
| `circular_cached` | 410 | a real loop, iteration off, in a file holding sentinel cached values: in a file openpyxl built or Excel saved, with and without `fullCalcOnLoad`, on opening, after a recalculation, a full one and a rebuild, in what Excel shows and what it saves; edits of the loop (`capture/circular_cached_probe.py`) |
| `names` | 92 | defined names as references: one cell, a range, built from other names or `OFFSET`, on another sheet, a value, `#REF!`; through `OFFSET`, `ROW`, `SUMIF` and the range operator; a name for a row, a column or a block taken where one cell is meant (`capture/name_probe.py`) |
| `criteria` | 144 | criteria and lookup keys that are blank: a blank cell, an empty text, `""`, `"="`, `"<>"`, 0, across the `COUNTIF` family; whole columns, computed blanks, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `XMATCH`; what a lookup gives when it finds a blank cell; the range operator given errors; intersection of single cells |
| `hidden_rows` | 98 | `SUBTOTAL` of every function number over rows a filter hid and rows hidden by hand, a filter's range without a filter, hidden columns, with controls (`capture/hidden_probe.py`) |
| `math_trig` | 2,076 | trigonometry, hyperbolics, logarithms, `SIGN`, `TRUNC`, `EXACT` and `TIME` over edge inputs: zero, tiny and huge numbers, text, blanks, logicals, errors |
| `math_integers` | 4,457 | rounding to a multiple, `QUOTIENT`, factorials, combinations, `GCD`, `LCM`, `MULTINOMIAL`, `ROMAN`, `ARABIC`, `BASE`, `DECIMAL`, `SERIESSUM`, `SUMX2MY2` and its family, over the same inputs |
| `info_family` | 599 | `XOR`, `ISEVEN`, `ISODD`, `ISNONTEXT`, `ISREF`, `ISFORMULA`, `FORMULATEXT`, `ERROR.TYPE`, `TYPE`, `SHEET`, `SHEETS`, `AREAS`, `HYPERLINK` across argument kinds, references, unions and errors. Excel refuses a file with `AREAS`, `ISFORMULA` or `FORMULATEXT` over a constant, an array, an error other than `#REF!`, or a function that gives no reference, such as `SUM` or `NA` |
| `matrix_family` | 360 | `MDETERM`, `MINVERSE` (each element through `INDEX`) and `MUNIT` over matrices stacked in cells: integers, permutations, singular and near-singular ones, Hilbert matrices up to 7 by 7, random ones up to 8 by 8, magnitudes that overflow or underflow, a subnormal element, a text, a blank, a logical, an error and a number as text, not square; scalars, written and computed 1 by 1 arrays; `MUNIT` sizes from fractional, text, logical, error and an array |
| `distributions_family` | 3,848 | the discrete and closed-form distributions (`BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `NEGBINOM.DIST`, `POISSON.DIST`, `HYPGEOM.DIST`, `EXPON.DIST`, `WEIBULL.DIST`, `LOGNORM.DIST`, `LOGNORM.INV` and their older forms), `GAMMA`, `GAMMALN`, `ERF`, `ERFC` and their `.PRECISE` forms: grids of counts with fractions, zeros, negatives and values past the domain, probabilities at 0, 1 and past them, 1,000 to 100,000 trials, gamma at its poles and overflow, erf in its tails, every kind in each argument |
| `continuous_family` | 5,059 | the beta, gamma, chi-squared, F and t distributions (`BETA.DIST`, `GAMMA.DIST`, `CHISQ.DIST` and `.RT`, `F.DIST` and `.RT`, `T.DIST`, `.2T` and `.RT`, their inverses and older forms), the tests (`T.TEST`, `F.TEST`, `CHISQ.TEST`, `Z.TEST` and older forms) over fixed samples, `CONFIDENCE.T`: probabilities from 1E-300 to 1 and past, degrees of freedom fractional, below 1 and up to 1E10, densities at 0, every kind in each argument; then a second set: the right tail's inverse at one degree of freedom at probabilities from 1E-17 to 0.3, `T.INV.2T` near 1, tails and densities at shapes from 2E5 to 1E8; and a third: extreme shapes and values, from 1E-307 to 1.7E308, but for seven formulas Excel does not finish (`slow_cases`); and a fourth: the tests given array constants beside ranges, with text, a logical or a blank, and arguments of other shapes holding as many elements |
| `slow_cases` | 62 | how long Excel takes over one formula alone in a workbook (`capture/slow_probe.py`): `BETA.DIST`, `GAMMA.DIST` and `BETA.INV` at shapes from 1E5 to 1E100, each within 3 seconds, and the third set of `continuous_family`, where seven formulas at a shape of 1.7E308 keep Excel calculating past 90 seconds and Excel was ended (§2). Not compared by `tests/test_captures.py`; `tests/test_quirks_figures.py` reads it |
| `complex_family` | 1,229 | complex numbers as text: `COMPLEX` over parts of 15 and 16 significant digits, 1E-300 to 1E300, 0.1+0.2 and 1/3, each suffix and kind; `IMREAL` and `IMAGINARY` over the forms a text can take (signs, spaces, `i` alone, exponents, malformed ones); the other 24 functions over complex numbers with both suffixes, zero parts and extreme magnitudes, `IMPOWER` at seven powers, mixed suffixes and arrays; then a second set: parts on either side of the switches to the exponent form near 1E21 and below 1E-5, `IMCOT` and `IMTAN` around 2^26 and 2^27, `IMPOWER` at 14 more powers (whole, negative, fractional, up to 1E15), the logarithms over 13 moduli, the suffix and trimming rules on other values |
| `mac_gaps` | 8 | whether this Excel computes `ENCODEURL`, `FILTERXML` and `EUROCONVERT`, with and without `_xlfn.`: `#NAME?` in every form, as for a name it does not know, so their behavior needs another Excel. Not compared by `tests/test_captures.py` |
| `text_family` | 10,032 | `PROPER`, `CLEAN`, `T`, `UNICHAR`, `UNICODE`, `SEARCH`, `SUBSTITUTE`, `NUMBERVALUE`, `USDOLLAR`, `VALUETOTEXT`, `ARRAYTOTEXT`, `TEXTBEFORE`, `TEXTAFTER` and the byte functions over argument kinds, wildcards, separators, instances and matching modes; how `LEN`, `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH` and `REPLACE` count a character outside the first plane; results at the limit of a text; `DOLLAR` and `USDOLLAR` of huge numbers |
| `dates_family` | 5,028 | `DAYS`, `DATEDIF` in every unit over month ends and the 1900 leap day, `ISOWEEKNUM`, `WEEKNUM` of every type, argument kinds and the limits of the calendar |
| `database_family` | 1,610 | `DCOUNTA`, `DGET`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR`, `DVARP` over the criteria of `misc`; all twelve D functions over a table with a blank row and every kind of value, over whole columns and with whole columns as criteria; errors and constants in each argument (the 192 forms Excel refuses in a file, a database that is no reference, are in `database_refused.tsv`, from `capture/database_probe.py`) |
| `engineering_family` | 3,814 | the base conversions over every width, two's complement, digits Excel refuses and `places`, and which problem comes first; `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT` to the limits of 2^48 and 53; `DELTA`, `GESTEP` across kinds |
| `statistics_family` | 4,475 | the descriptive and regression functions over spread, mixed kinds, errors, zeros, values near 1E150 and 1E-150; `TRIMMEAN`, `PERCENTRANK` and its forms with `significance`, `PERCENTILE.EXC`, `QUARTILE.EXC`, `PROB`, `CONFIDENCE`; an empty argument in every function taking many |
| `finance_family` | 15,513 | `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `SYD`, `DISC`, `INTRATE`, `RECEIVED`, `PRICEDISC`, `YIELDDISC`, the Treasury bills, `DOLLARDE`, `DOLLARFR`, `EFFECT`, `NOMINAL`, `FVSCHEDULE`, `ISPMT`, `MIRR`, `PDURATION`, `RRI` over rates, periods, dates in every basis, argument kinds and the order of errors |
| `empty_args` | generated | every implemented function from the registry, with real values: a baseline with nothing empty, each argument empty in turn, its optional ones empty together, and each position of a first and a later group of a repeating tail (`*IFS`, `IFS`, `SWITCH`). `capture/empty_probe.py` opens them in batches and halves a refused batch: what Excel refuses is in `empty_args_refused.tsv`, and the suite holds the rest |
| `cum_large` | 40 | `CUMIPMT`, `CUMPRINC`, `DB` over 1E6 to 1E8 periods: answered at once, with no bound |
| `combin_limit` | 164 | `COMBIN` and `PERMUT` at powers of 2 and of 10 up to 2^33: 2,147,483,646 is computed and 2,147,483,647 is `#NUM!` |
| `index_range` | 38 | ranges built with `INDEX`, lookups and aggregates over a formula's own column: which Excel calls circular |
| `static_cycle` | 37 | loops of references through a branch not taken, and real loops as controls: which Excel calls circular |
| `array_refs` | 8,602 cells | `ROW`, `COLUMN`, `CELL`, ranges and arrays as a criterion or a lookup key, and arrays of references, each as a legacy formula, a legacy array formula and a dynamic array |
| `array_functions` | 41 | the `SUMIF` family, `DSUM`, `CORREL` and `ROW` times `COLUMN` over one column, ten and every column of another sheet, with Excel's memory: Excel's value, or the file's value kept |
| `circular`, `circular_alerts` | 16 each | circular references with iteration off: what Excel saves and shows, alerts off and on (the second interactive; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.2) |

All live in `tests/captures/` as TSV files with a header: suite, Excel build and locale, edition
and update channel (§2), date system, capture date, generator and method. `tests/test_captures.py` and
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
`ulimit -v` is no limit on macOS; a sweep without the watchdog can use several times the
machine's memory, and macOS then pauses every application, Excel included. A killed run (return code 137, or a missing
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
last bits (`MATH_LIBRARY`) must differ only on the platform `docs/coverage.md` records; one that matches there and differs
on another, where a cancellation exposes the library's last bits (`OTHER_PLATFORMS`), must match there. A function whose results stay a few units in the
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

**Opens.** Excel refuses a whole file for one formula it does not take, so a probe that must name refused formulas
opens them in batches of 60 and halves a refused batch until each refusal is opened alone (`capture/refusals.py`).
Settled answers are not asked again: a refusal of the same build (its list's `# excel:` line) is kept without an open,
and the forms a captured suite holds are confirmed in full batches, where a refusal is news and is bisected; a new form,
or a refusal of another build, is probed. `--full` asks everything again. Measured on Excel 16.114.4: a batch of 60
opens in about 4.5 s, a workbook of one formula in about 1 s. With nothing new, `capture/empty_probe.py` makes 22
opens (about 100 s) and `capture/database_probe.py` 2 (8 s); with `--full`, 955 (about 15 minutes) and 514 (21
minutes). The suites
(`capture/pipeline.py`, `capture/dynamic_probe.py`) open one workbook a suite. One workbook a case remains where the
case is the workbook: the iteration probes and `calc_chain` (a workbook's calculation settings, its cached values,
its sheets' order; Excel takes iteration settings from the first workbook open), `circular_cached`, `cell_probe` (the
active sheet), `array_size_probe` (Excel's memory per file), and `value_probe`, `literal_probe` and
`workbook_name_probe`, whose cases mostly make Excel refuse the file and each record whether its file opens;
`capture/refusal_probe.py` opens each formula given alone by design, to name the one a suite's workbook fails on.

**The edition is part of the evidence.** The build number does not say which functions Excel has:
Office Home & Business 2021 for Mac at 16.114.4 lacks `LAMBDA` and its helpers, which Microsoft 365
and Excel 2024 at the same build have. Every capture header records the edition and the update
channel beside the build (`# edition: Office Home & Business 2021, update channel Current`;
`capture/edition.py`). No file on the Mac states the license plainly, so the person capturing
states the edition, as shown in Excel > About Excel, in `capture/build/excel_edition.txt` or the
variable `XLRECALC_EXCEL_EDITION`, and a capture refuses to run without it; the channel is
Microsoft AutoUpdate's. A capture without the line was made before the rule, under Office Home &
Business 2021 for Mac.

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
Each Excel step there has a time limit (`tests/excel_steps.py`); when Excel then does not answer, the
run ends with a message naming the test, and Excel is left for the person using it to look at. The pipeline
has a limit too, 600 s and 0.1 s per case, and so do the probes that open a workbook and save a copy.

**When Excel stops answering.** Excel can stop answering questions about its workbooks while it
still answers `version`: at 0% CPU, with no dialog and its menus working, its main thread idle in its
event loop, so the request was never dispatched. A Quit clears it, and the same work then passes; it
can also clear by itself within minutes. A question about workbooks therefore tells whether Excel is
alive (`tests/excel_steps.py`). Four such hangs have no known cause, and each came with the display
off; a locked screen, a display that is off and system sleep, each tried deliberately, did not cause
one. A display that the tools wake does not stay on for `caffeinate -d`: in the fourth, it went off
again 11 seconds after the wake while the assertion was held. A known cause looks the same: a window
of Excel's own (the About window, Preferences, any dialog) blocks every request about workbooks.

What the tools do about it. A process that drives Excel keeps the display awake (`caffeinate -d`,
from its first request until it ends, after declaring user activity once to wake a display that is
off), and a hang record says whether it did and how the display stood when the run began. Before
the first request, `capture/excel_watch.py` compares the windows Excel has on screen with the
document windows it reports; a window more, or no answer, stops the run with a message naming the
windows by size, before anything is sent. A request past its limit, or one AppleScript gives up on
first (`AppleEvent timed out`, -1712), gets a record under `capture/build/excel_hangs/` (Excel's
`sample`, uptime, CPU and memory, the display's last change, the request and its file), and Excel is
asked about workbooks at once and after 1 and 5 minutes. When it answers, the workbooks the request
named in Excel's container are closed without saving and alerts are turned back on, so a recovered
hang leaves no probe's file open with alerts off. `uv run python -m capture.excel_watch [RECORD]`
asks again and adds the answer to that record, or to the latest if it is from the last half hour
(else to a record of its own): run it before anyone quits Excel.

The tools do not end Excel, since the person using it may have work open in it: they report and
leave the decision to that person, with two exceptions, each only when nothing of anyone else's can
be in Excel. `capture/array_size_probe.py` refuses to start with a workbook open, and ends Excel
when a step passes its time limit or Excel passes 8 GB of memory; there, running past a limit is a
result, as it is for `capture/value_probe.py`'s own limit. And a probe whose case can keep Excel
calculating without end (a function at an extreme input; `capture/slow_probe.py`) may let the watch
end it when `may_end` holds (`capture/excel_watch.py`, tested): the run began with Excel not running
or holding no workbook, every Excel window on screen is the request's own (no more of them than its
workbooks, and each title macOS gives one of their names), and one process is named Microsoft
Excel. The watch then sends `TERM`, and `KILL` after 10 seconds, records it and goes on; the next
request checks Excel's windows again (a recovery window would stop the run). In any other state it
stops and asks.

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
`tools/function_frequency.tsv`, with its attribution. Those workbooks date from about 2001, so
they cannot contain a newer function, and a count of zero says nothing about one. Rules:

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
`tools/corpus_compare.py` evaluates every workbook against the values Excel cached in it, each
formula alone and in a full recalculation, so that a formula differing for itself is told from one
that inherits a difference. It finds questions and answers none: a cache may be stale, so a
difference found there is settled by a capture, and its output stays local.
