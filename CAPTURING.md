# Capturing Excel's results

Expected results are what Excel 16 for Mac computes, and for the suites of `capture/suites.py`
also what Excel 16 for Windows computes (§5). This document describes how they are captured, the
facts about Excel and openpyxl the method relies on, and how formulas from other people's
workbooks are used.

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
| `marker_typed` | 75 | which formulas typed in Excel's formula bar it marks as dynamic arrays: the functions that make arrays with a one-cell result, wrapped in another function, under an operator, a branch and a function taken element by element, and the functions of a reference with sizes in the text or not (entered as typed, kept as `marker_typed.xlsx` without who saved it or where, and checked against the capture; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.7) |
| `let_typed` | 8 | `LET` typed in Excel's formula bar (`capture/lambda_probe.py --typed`): its file form (`_xlfn.LET`, `_xlpm.` on each name), which formulas Excel marks as dynamic arrays (only one whose result spills: not `LET(x,SEQUENCE(3),SUM(x))`, nor a single value bound to a range or a cell) and the values. The same workbook types `LAMBDA` and its helpers; those rows stay local, since the `lambda_*` captures hold their typed forms and values |
| `lambda_calls` | 37 | `LAMBDA` calls pasted as typed (`capture/lambda_probe.py --paste`, every cell of each result): argument counts (too many, too few, none: `#VALUE!`), omitted and empty arguments, `ISOMITTED`, an optional parameter (`[y]`, read as 0 when omitted), errors as arguments, a lambda where a value is expected (`#VALUE!`) and uncalled as a result (`#CALC!`), `TYPE` (128), references returned, calls of `IF`, `CHOOSE` and `LET` results |
| `lambda_helpers` | 53 | `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` pasted as typed: sizes that differ (`#N/A`), omitted initial values, row-major order, references passed for a range's elements, a result that is not one value (`#CALC!`), sizes that are not positive integers (`#VALUE!`), a non-lambda |
| `lambda_names` | 23 | named lambdas (defined names written as Excel stores them) pasted as typed: calls, recursion (`FACTR(170)`; 171 is `#NUM!`), mutual recursion, a named lambda passed as an argument |
| `lambda_depth` | 108 | where recursion through a named lambda stops (`#NUM!`), by parameter count and body: 5,461 calls of one parameter, 4,095 of two, 3,276 of three, under 65,536 at 8 + 4 per parameter a call, as are bodies with `SUM` of two arguments or three nested calls around the recursive one; `SUM` of four or eight arguments stops one or two calls earlier, a body calling itself in tail position goes past 8,000 (not 10,000); a body binding one `LET` name stops at 4,095, as two parameters do |
| `lambda_more` | 79 | pasted as typed, with names: optional parameters in the budget (given or not, 4,095 calls of one required and one optional), two `LET` names (3,276), the `IS` functions, `N` and `T` given a lambda, a name holding 5 called (`#VALUE!`), `MAP` padding a one-element argument with `#N/A`, `REDUCE`/`SCAN` elements as references, recursion through `MAP` (20 nested compute, 21 `#VALUE!`), `LAMBDA(x,x)=1`, `(LAMBDA(x,x))(2)`, `MAKEARRAY` with an empty size (1), `LET` names in untaken branches (charged) |
| `marker_value` | 14 | the same entered through AppleScript's `value` of a cell: legacy formulas, never marked. Not compared by a capture test: evidence for the finding it states. |
| `dynrefs` | 90 | `OFFSET`, `INDIRECT`, `CELL`, `INDEX` as a reference, and dependencies only evaluation finds ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §6.3) |
| `cell_noref` | 36 | `CELL` without a reference in workbooks differing in the active sheet and cells |
| `n_lookup` | 56 | `N` and `LOOKUP` across argument kinds, both forms of `LOOKUP`, short result ranges |
| `finance` | 1,248 | `PV`, `FV`, `PMT`, `IPMT`, `PPMT`, `NPER`, `NPV`, `XNPV`, `SLN`, `VDB` over signs, sizes, fractional and negative periods, both payment types and argument kinds; cash flows and dates as fixed cells; at the end, `VDB` over fractional start and end periods around a corpus call, with `no_switch` both ways, on two sets of values; `VDB` from fractional starts around the switch to straight line on two assets, with each asset's whole periods |
| `iterative` | 402 | `IRR`, `XIRR`, `RATE` over 18 cash-flow shapes (two roots, a double root, no sign change, huge and tiny rates), 11 guesses, date orders and argument kinds; `IRR`, `NPV`, `MIRR` and `XIRR` over two sets of flows holding an error (four kinds, and two at once), a text, a number as text, a blank or a logical at the start, middle or end, and `XIRR` over dates holding them |
| `bonds` | 2,395 | `DAYS360` and `YEARFRAC` over every pair of 14 dates (month ends, February in leap and common years) in all bases; the six `COUP` functions, `PRICE`, `YIELD`, `DURATION`, `MDURATION` over six bonds, three frequencies and five bases; argument errors |
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
| `iteration` | 1,337 | `IRR`, `RATE`, `XIRR` from 20 guesses over flows of many shapes, and over a double root with guesses 2^k away, which shows the iteration's limit and stopping test; `YIELD` over a grid of bonds. Its `XIRR` rows are compared by `tests/test_captures.py`; the `IRR`, `RATE` and `YIELD` rows are evidence for inferring Excel's iteration, not compared |
| `iterate_open` | 55 | what Excel does on opening a file with iterative calculation on: a counter, a loop that converges, a dependent, a control and a loop with no cached values, across settings, a file `recalculate` wrote and a copy Excel saved (`capture/iterate_open_probe.py`) |
| `iterate_calc` | 86 cells | how Excel iterates (49 workbooks and 2 edits, `capture/iterate_probe.py`): runs of `iterateCount` on opening, on `calculate`, on saving and after an edit; the order of cells within a sheet and across sheets, with sheet names in case, digits and accents; a precedent outside a loop, a dependent after it and above it; `iterateDelta` at, above and below a change, over large and small magnitudes, and with a change between kinds; one loop or two; loops through `INDIRECT` and `OFFSET`; oscillation, overflow, a loop already converged |
| `calc_chain` | 24 | whether `calcChain.xml` sets the order of iteration (`capture/calc_chain_probe.py`): a loop within a sheet and one across sheets, saved by Excel, then with its chain as Excel wrote it, reversed, or removed, each opened, calculated and saved, and the saved copy opened again. The chain does not: the order is the same in all, and Excel writes its own chain back. Not compared by a capture test: evidence for the finding it states. |
| `iterate_uncached` | 84 | iteration in files whose formulas have no cached values (`capture/iterate_uncached_probe.py`): 17 loops, each with and without `fullCalcOnLoad` and with every formula cached 0; after opening, after `calculate`, saved |
| `iterate_opening` | 72 | what opening does to a loop with no cached values (`capture/iterate_opening_probe.py`, 32 files as openpyxl writes them): the same loops in other row and column orders, entered from a constant or a formula or not, IFERROR in each cell in turn, opening's counts 1 to 3, and 8 more loops in files as openpyxl writes them, read through the public API |
| `iterate_reads` | 9 | opening's one run over every loop (`capture/iterate_reads_probe.py`, every formula cached 0, `fullCalcOnLoad`, 100 iterations): two counters that never settle, loops of one and three cells that settle at different speeds, a loop read by another; after opening, after `calculate`, saved |
| `volatile_edit` | 156 | what an edit recomputes after a trusting opening (`capture/volatile_edit_probe.py`, Excel's saved copies planted, 30 files, each taken through its own steps: edits, `calculate`, `calculate full`, listed in `capture/volatile_edit_cases.py`): a cell volatile through a defined name without `ca="1"`, not computed since opening, keeps its value through an edit until a full calculation, whatever the volatile function in the name (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `CELL`, `OFFSET`, `INDIRECT`, or another volatile name), and nothing reaches it through the name's references, plain ones included; an edit of a cell its own text names, or a plain name beside the volatile one names, computes it, and from then on every edit does, one touching nothing it reads too; nothing reaches its dependents through it, and a dependent computed for another reason reads its cached value; Excel's save marks a dependent `ca="1"`; in manual calculation, `calculate` leaves it too; computed once (a full calculation, `fullCalcOnLoad`), it follows every edit; marked, or volatile in its own text, it is computed; a plain name's references, and plain dependents, are followed. Rows have one shown column per step (the header says how), so a variant's row length follows its steps, and the test checks it |
| `iterate_cached` | 6 | which Excel actions run a loop with cached values (`capture/iterate_cached_probe.py`): Excel's saved copy `X` of an openpyxl loop (kept, its markers removed, as `iterate_cached.xlsx`, the input of the end-to-end tests); `X` opened and saved (two runs); `X` opened, a cell nothing reads edited, and saved (three runs) |
| `workbook_name` | 7 | what `[0]!Name` means (`capture/workbook_name_probe.py`): the workbook-level name, from either sheet, even one with a sheet-level name of the same name; a missing one is `#NAME?`; Excel writes `[0]!` back as it was. `[1]!Name` with no external link makes Excel refuse the file |
| `circular_cached` | 410 | a real loop, iteration off, in a file holding sentinel cached values: in a file openpyxl built or Excel saved, with and without `fullCalcOnLoad`, on opening, after a recalculation, a full one and a rebuild, in what Excel shows and what it saves; edits of the loop (`capture/circular_cached_probe.py`) |
| `open_calc` | 38 | when Excel recalculates a workbook on opening (`capture/open_calc_probe.py`): a copy Excel saved with its cached values planted wrong, one change each (no `calcPr`; 16 `calcId`s besides the copy's own, from 124519 to 999999; `calcId` removed; `fullCalcOnLoad`; a volatile cell; manual calculation alone, with `fullCalcOnLoad`, with an old `calcId` and with a volatile cell; openpyxl's own files; and one difference each between openpyxl's file and Excel's: no calculation chain, no `docProps/app.xml`, no `<fileVersion>`, and each attribute of `<fileVersion>`, of which only `appName="xl"` makes Excel keep the cached values); what Excel shows on opening and what it saves after (its `volatile` row is a volatile cell without Excel's `ca="1"`, which only another program writes; `volatile_names` covers the marked case) |
| `volatile_names` | 18 | what Excel recalculates on opening a copy it saved, its cached values planted wrong (`capture/volatile_name_probe.py`; A2 = A1*3, A3 the formula under test, A4 = A3+1): a volatile function in the cell, through a defined name (`OFFSET`, `INDIRECT`, `RAND`, `NOW`, a chain of names, a `LAMBDA` name calling `NOW()`), and controls (a plain name or `LAMBDA`, an unused volatile name, `"RAND()"` in a literal); Excel's `ca="1"` mark removed from a volatile cell or added to a plain one; manual calculation. Excel writes `ca="1"` on every volatile cell and its dependents, recalculates the marked cells on opening and keeps the rest; an unmarked volatile function in a cell's own text recalculates every formula; manual calculation does the same when saving |
| `cf_rules` | 252 | how Excel evaluates a conditional format's formula (`capture/cf_probe.py`: 76 cases, each a workbook with one rule, two for `stopIfTrue`, read through `display format` cell by cell against a control rule): the anchor is the top-left corner of the box around all the rule's areas, whatever their order; a rule is evaluated as a formula is today, not as a legacy one, and a range or array result counts by its top-left element; the format applies where `IF` takes its true branch on it (TRUE, a nonzero number, the text `"TRUE"`), not for an error, other text or a blank; another sheet works directly, through a name and in the 2010 extension; an array constant or a structured reference in a rule makes Excel offer to repair the file; `cellIs` compares as Excel's operators do, for every operator and for text, logicals and blanks against number and text bounds, `between` and `notBetween` with their bounds in either order; a volatile function in a rule recalculates no cell on opening |
| `bessel_family` | 4,359 | `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY` over x from 0 to 1E100 (negative for I and J) and n from 0 to 100, dense sweeps of x across 3 to 4.5 (I), 1.5 to 2.5 (K) and 7.5 to 8.5 (J, Y), n's and x's edges and the largest n: Excel's values are approximations, about 1E-8 to 1E-7 relative from the exact functions (`BESSELJ` with n = 1 from x = 8 on about 3E-5), with a change of approximation at x = 8 for J and Y; n is truncated, negative n and TRUE or FALSE are errors (`#NUM!`, `#VALUE!`), a blank n is 0. The limits (79 rows at the end): every function takes n up to 1E7 (10,000,001 is `#NUM!`, `BESSELK(1E6,n)` too, where nothing overflows); K and Y overflow to `#NUM!` from n = 180 at x = 2.5; J and Y are `#NUM!` once x less 0.785398164 (order 0) or 2.356194491 (order 1) reaches 2^27 - 1; I is computed at x = 709.7 and `#NUM!` from 709.71. Every row matches the Numerical Recipes approximations, with their published coefficients and four changes (QUIRKS.md) |
| `convert_family` | 2,558 | `CONVERT` over every unit Excel documents, to and from its category's base, pairs within categories, decimal and binary prefixes, temperatures, spelling and the number's coercion: every documented name is taken, case counts, 52 units take a decimal prefix (`mmHg`, `mph`, `cal` and `ga` among them, not `in` or `lbm`), only `bit` and `byte` a binary one, and `e` is a prefix of 10 as `da` is; a prefix on an area or volume applies once per power. Temperatures convert by a formula for each pair of scales, not through kelvin (`CONVERT(0,"F","C")` is -17.77777777777778, where through kelvin it would be -17.77777777777777); a prefix on kelvin is ignored from Réaumur and to Rankine or Réaumur. Conversions between two units neither of which is the base differ from a product of the captured factors in the last bit or two (QUIRKS.md) |
| `convert_pairs` | 8,928 | `CONVERT` at 1, 2.5 and 3 over every ordered pair of units within each category (2,976 pairs, temperature included): the ratio Excel uses for each pair and whether its result is that ratio times the number: the evidence for each unit's constant in its category's internal unit (`src/xlrecalc/_functions/_convert.py`) |
| `convert_odd` | 741 | `CONVERT` at odd numbers (0.1, 7.3, 1/3, 123456.789, 1E-7, e, 0.7, 19.99): each unit to the next in its category and back, each unit to itself, prefixes on both sides, prefixed units to other units, binary and decimal prefixes mixed, every pair of temperature scales |
| `iteration_steps` | 36 | `IRR` of -1 and X from 3E6 to 1E9, `RATE(n,-1,n/2)` from n = 100 to 1500, and `RATE` from far guesses: where the number of Newton steps decides the outcome |
| `rate_method` | 88 | `RATE` where its limits decide the outcome: `RATE(n,-1,n/2)` for n from 100 to 2000, loans from far guesses whose Newton paths take 128 to 152 steps, and two families of negative guesses whose steps multiply the equation's value by from about 20 to about 60,000 before converging |
| `xirr_method` | 276 | `XIRR` where its scan and bisection decide the outcome: roots from 1E14 to 7E65 a day apart from four guesses (the top of the upward scan), negative guesses against negative and positive roots, guesses of 1 and more, single roots on random dates (six fractional), and flows with two and three roots, a root at 0 and roots 0.13 and 23 |
| `xirr_sign_rules` | 22 | `XIRR` where a change below 0 runs away from the first flow's sign or XNPV at -1 is -inf + inf: two roots either side of 0, two positive and two negative roots, and flows whose latest has the first flow's sign; written to separate two rules before Excel computed it |
| `xirr_nan` | 24 | `XIRR` of flows (-, -, +) and their mirror image, whose XNPV at -1 is -inf + inf, with roots -0.2 to -0.9 from no guess and four guesses |
| `whole_columns` | 86 | whole columns and rows where the cells past the data count: sizes, blanks counted, reads past the last value (`INDEX`, `MATCH`, `XMATCH` from the end, `LOOKUP(2,1/(...))`), extremes over the blanks, `OFFSET` from a whole column, an empty column, whole rows; then `MATCH` given one value as its lookup array, under each match type. Not compared by `tests/test_captures.py`, which runs each formula more than once; `tests/test_whole_columns.py` compares every row on its captured row, the four over column AB (whose only value is in its last row, so their array is the whole column) in the slow tier. |
| `dyn_whole_columns` | 20 cases | whole columns as dynamic arrays (`capture/dynamic_probe.py`): `IF` over a column, `TAKE`, `UNIQUE`, `FILTER` and `SORT` of a whole column, a search from its end, and a bare whole-column array, which cannot spill. Not compared by `tests/test_dynamic_arrays.py`; `tests/test_whole_columns.py` compares every cell and spill range in the slow tier, since `SORT`, `UNIQUE`, `FILTER` and `TAKE` of whole columns make arrays of a million rows. |
| `dyn_cancel` | 21 cases | cancellation snapping inside arrays (`capture/dynamic_probe.py`): a subtraction or an addition of two columns as the formula's last operation, in a dynamic array and in a legacy array formula (`cse`), each pair of values 1, 4, 7, 8 or 16 units in the last place of the left one apart; the same nested (`(...)*1`, `+0`, `IF`, `LET`, `TRANSPOSE`), an array less one cell and one cell less an array; `SUM` and `SUMPRODUCT` of an array expression and their own additions. Each element snaps as one value does: below 8 units of its left operand, to 0, and only as the last operation; `SUM` snaps its additions over an array as over a range, `SUMPRODUCT` does not. Compared by `tests/test_dynamic_arrays.py`, as is Excel for Windows' capture of it (§5) |
| `names` | 92 | defined names as references: one cell, a range, built from other names or `OFFSET`, on another sheet, a value, `#REF!`; through `OFFSET`, `ROW`, `SUMIF` and the range operator; a name for a row, a column or a block taken where one cell is meant (`capture/name_probe.py`) |
| `criteria` | 144 | criteria and lookup keys that are blank: a blank cell, an empty text, `""`, `"="`, `"<>"`, 0, across the `COUNTIF` family; whole columns, computed blanks, `MATCH`, `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `XMATCH`; what a lookup gives when it finds a blank cell; the range operator given errors; intersection of single cells |
| `hidden_rows` | 98 | `SUBTOTAL` of every function number over rows a filter hid and rows hidden by hand, a filter's range without a filter, hidden columns, with controls (`capture/hidden_probe.py`) |
| `math_trig` | 2,076 | trigonometry, hyperbolics, logarithms, `SIGN`, `TRUNC`, `EXACT` and `TIME` over edge inputs: zero, tiny and huge numbers, text, blanks, logicals, errors |
| `math_integers` | 4,457 | rounding to a multiple, `QUOTIENT`, factorials, combinations, `GCD`, `LCM`, `MULTINOMIAL`, `ROMAN`, `ARABIC`, `BASE`, `DECIMAL`, `SERIESSUM`, `SUMX2MY2` and its family, over the same inputs |
| `info_family` | 599 | `XOR`, `ISEVEN`, `ISODD`, `ISNONTEXT`, `ISREF`, `ISFORMULA`, `FORMULATEXT`, `ERROR.TYPE`, `TYPE`, `SHEET`, `SHEETS`, `AREAS`, `HYPERLINK` across argument kinds, references, unions and errors. Excel refuses a file with `AREAS`, `ISFORMULA` or `FORMULATEXT` over a constant, an array, an error other than `#REF!`, or a function that gives no reference, such as `SUM` or `NA` |
| `matrix_family` | 360 | `MDETERM`, `MINVERSE` (each element through `INDEX`) and `MUNIT` over matrices stacked in cells: integers, permutations, singular and near-singular ones, Hilbert matrices up to 7 by 7, random ones up to 8 by 8, magnitudes that overflow or underflow, a subnormal element, a text, a blank, a logical, an error and a number as text, not square; scalars, written and computed 1 by 1 arrays; `MUNIT` sizes from fractional, text, logical, error and an array |
| `distributions_family` | 3,848 | the discrete and closed-form distributions (`BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `NEGBINOM.DIST`, `POISSON.DIST`, `HYPGEOM.DIST`, `EXPON.DIST`, `WEIBULL.DIST`, `LOGNORM.DIST`, `LOGNORM.INV` and their older forms), `GAMMA`, `GAMMALN`, `ERF`, `ERFC` and their `.PRECISE` forms: grids of counts with fractions, zeros, negatives and values past the domain, probabilities at 0, 1 and past them, 1,000 to 100,000 trials, gamma at its poles and overflow, erf in its tails, every kind in each argument |
| `continuous_family` | 5,059 | the beta, gamma, chi-squared, F and t distributions (`BETA.DIST`, `GAMMA.DIST`, `CHISQ.DIST` and `.RT`, `F.DIST` and `.RT`, `T.DIST`, `.2T` and `.RT`, their inverses and older forms), the tests (`T.TEST`, `F.TEST`, `CHISQ.TEST`, `Z.TEST` and older forms) over fixed samples, `CONFIDENCE.T`: probabilities from 1E-300 to 1 and past, degrees of freedom fractional, below 1 and up to 1E10, densities at 0, every kind in each argument; then a second set: the right tail's inverse at one degree of freedom at probabilities from 1E-17 to 0.3, `T.INV.2T` near 1, tails and densities at shapes from 2E5 to 1E8; and a third: extreme shapes and values, from 1E-307 to 1.7E308, except seven formulas Excel does not finish (`slow_cases`); and a fourth: the tests given array constants beside ranges, with text, a logical or a blank, and arguments of other shapes holding as many elements |
| `slow_cases` | 62 | how long Excel takes over one formula alone in a workbook (`capture/slow_probe.py`): `BETA.DIST`, `GAMMA.DIST` and `BETA.INV` at shapes from 1E5 to 1E100, each within 5 seconds, opening the workbook included, and the third set of `continuous_family`, where seven formulas at a shape of 1.7E308 keep Excel calculating past 90 seconds and Excel was ended (§2). Not compared by `tests/test_captures.py`; `tests/test_quirks_figures.py` reads it |
| `complex_family` | 1,229 | complex numbers as text: `COMPLEX` over parts of 15 and 16 significant digits, 1E-300 to 1E300, 0.1+0.2 and 1/3, each suffix and kind; `IMREAL` and `IMAGINARY` over the forms a text can take (signs, spaces, `i` alone, exponents, malformed ones); the other 24 functions over complex numbers with both suffixes, zero parts and extreme magnitudes, `IMPOWER` at seven powers, mixed suffixes and arrays; then a second set: parts on either side of the switches to the exponent form near 1E21 and below 1E-5, `IMCOT` and `IMTAN` around 2^26 and 2^27, `IMPOWER` at 14 more powers (whole, negative, fractional, up to 1E15), the logarithms over 13 moduli, the suffix and trimming rules on other values |
| `mac_gaps` | 8 | whether this Excel computes `ENCODEURL`, `FILTERXML` and `EUROCONVERT`, with and without `_xlfn.`: `#NAME?` in every form, as for a name it does not know, so their behavior needs another Excel (Excel for Windows: `windows/function_names`, `windows/encodeurl`). Not compared by `tests/test_captures.py` |
| `function_names` | 511 | whether this Excel knows each name of `capture/functions_list.py`: TRUE when a call gives `#NAME?` (`IFERROR(ERROR.TYPE(...)=5,FALSE)`), each called in the first form the `formula_bar` capture records Excel taking, or a form of its own (`_NAME_CALLS`); the names Excel may not know with and without `_xlfn.`. `ENCODEURL`, `FILTERXML`, `EUROCONVERT` and `JIS` give `#NAME?` in both forms (`JIS`: cases 265 and 266, and on Windows too, so no en-US Excel knows it and `tools/excel_names.py` leaves it out), every other name is known. The 14 names that reach outside the workbook when computed (`NAMES_LEFT_OUT`: code libraries, real-time data, the network, cube connections) are not called |
| `encodeurl` | 177 | `ENCODEURL` (`_xlfn.`) over every code from 1 to 127, characters beyond ASCII, texts, other values and long text, and short arguments with results of 32,766 and 32,767 characters: `#NAME?` in every row, as `function_names` records; the values are Excel for Windows' (`windows/encodeurl.tsv`). Not compared by `tests/test_captures.py` |
| `bare_names` | 153 | each name Excel writes with a file prefix (`PREFIXES`, from `formula_bar` and `prefixes`), called without it in the form `function_names` calls it with the prefix, wrapped as there (`IFERROR(ERROR.TYPE(...)=5,FALSE)`), in a file openpyxl writes (no cached values, `fullCalcOnLoad`): `#NAME?` for every one. `SORT` and `FILTER`, which make Excel refuse the file written so, and `LET`, are not called |
| `text_family` | 10,183 | `PROPER`, `CLEAN`, `T`, `UNICHAR`, `UNICODE`, `SEARCH`, `SUBSTITUTE`, `NUMBERVALUE`, `USDOLLAR`, `VALUETOTEXT`, `ARRAYTOTEXT`, `TEXTBEFORE`, `TEXTAFTER` and the byte functions over argument kinds, wildcards, separators, instances and matching modes; how `LEN`, `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, wildcards in criteria and the limit of a text count a character outside the first plane; results at the limit of a text; `DOLLAR` and `USDOLLAR` of huge numbers |
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
| `circular`, `circular_alerts` | 16 each | circular references with iteration off: what Excel saves and shows, alerts off, and opened as a double-click opens it, with the warning dismissed by hand (the second interactive; [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §7.2) |
| `codepage` | 912 | `CHAR`, `UNICODE(CHAR(n))` and `CODE(CHAR(n))` for 0 to 256, and `CODE(UNICHAR(n))` for 127 to 255 and twelve characters outside an 8-bit code page: Mac Roman throughout, and `CODE` of U+0080 to U+009F gives the code itself. Not captured, so the package's choice: `CODE` of any other character outside Mac Roman is 95, as for the twelve captured |
| `clean_controls` | 510 | `CLEAN` over `UNICHAR(n)`, alone and between two letters, for 1 to 255: it removes 1 to 31 and U+0080 to U+009F, and keeps 127. Not captured, so the package's choice: it keeps every character past U+00FF (U+200B, U+FEFF and other format characters included) |
| `collate_chars` | 23,375 | `<` over every ordered pair of 125 single characters (printable ASCII, 28 Latin-1 letters and signs, and `œ`, `Œ`) and `=` over every unordered pair: one order after folding case, marks before digits before letters, `æ ð ø œ þ µ` letters of their own, `ª` and `º` equal to `a` and `o`, an accent counting only on a tie (`_compare.py`, `MAC_ORDER`). Not captured, so the package's choice: any other character sorts after all of these by code point, a control character before them |
| `collate_words` | 2,380 | the same over 40 words that separate the rules of a word order: hyphens and apostrophes inside, leading and trailing, case, `æ`, `ß`, `œ`, digits and accents: marks are not ignored (`co-op` < `coop`), `ß` equals `ss`, `æ` and `œ` follow `a` and `o` words |
| `collate_functions` | 169 | whether `SORT`, `COUNTIF` criteria, `MATCH`, `XMATCH`'s binary search, `VLOOKUP` and `UNIQUE` order and match text as `<` and `=` do, over a table of those words: they follow `<`; `COUNTIF`'s equality and `XMATCH`'s binary search follow `=` (`ß` counts and finds `ss`), while `MATCH(..., 0)` ignores only case (`ß` finds `ß`, not `ss`); a binary search finds the first of a run of equal entries (descending, `-2`, not captured: the first too, our choice) |
| `spill_edit` | 33 | what Excel shows after an edit makes a dynamic array spill further or less (`capture/spill_edit_probe.py`, six of Excel's saved copies, each taken through one or two edits, listed in `capture/spill_edit_cases.py`): a spill growing, shrinking, growing then shrinking, a spill reference growing with it, a spill blocked by a constant and free again, a reader on another sheet; a cell the spill now reaches, and a formula reading it singly or through a range, show the new values, and a cell it left is blank. It records what Excel shows after each edit, not an order in which Excel computes |

All live in `tests/captures/` as TSV files with a header: suite, Excel build and locale, edition
and update channel (§2), date system, capture date, generator and method. `tests/test_captures.py` and
`tests/test_formula_forms.py` compare the package with them, except where a row above says it is not
compared: each suite becomes one engine sheet laid out as Excel's workbook was, and every formula
cell must equal what Excel saved.

Where a fact belongs to [xlio](https://github.com/sheetmill/xlio), which reads and writes the
files, its capture is cited here rather than repeated: [`errors`](https://github.com/sheetmill/xlio/blob/main/tests/captures/errors.tsv)
(Excel 16.115) gives the rich-value error types and their `ERROR.TYPE` numbers in
`tests/test_error_types.py`, and `#GETTING_DATA` written back as `#N/A`.

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

**Crafted and extreme input** runs under a memory watchdog (CONTRIBUTING.md), never while Excel
captures: when memory runs out, macOS pauses every application, Excel included. A capture starts
once Excel answers (for example `osascript -e 'tell application "Microsoft Excel" to get
version'`).

**Adding a function or a family of functions.** Write its cases at the end of a suite in
`capture/suites.py`, or as a suite of its own (each case a formula with `{1}`, `{2}`... for its
inputs; fixed lookup tables through a suite's `fixed` cells), across argument kinds, edges and
errors; run `uv run python -m capture.pipeline SUITE`; implement it under
`src/xlrecalc/_functions/` with its registry kinds (§8.1 of [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md)); add a new suite
to `SUITES` in `tests/test_captures.py`. A row that cannot be matched needs its reason in
`KNOWN_DEVIATIONS` there and in `QUIRKS.md`. The list is strict: a listed row that matches Excel
fails the test, so an entry cannot outlive its cause; a random result is exempt
(`EXEMPT`), and a row listed because the platform's math library differs from Excel's in the
last bits (`MATH_LIBRARY`) must differ only on the reference platform (`REFERENCE_PLATFORM` in `tests/test_captures.py`); one that matches there and differs
on another, where a cancellation exposes the library's last bits (`OTHER_PLATFORMS`), must match there. A function whose results stay a few units in the
last place from Excel's after a fixed time spent fitting its last bits declares a tolerance in the
registry instead ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1), listed in `QUIRKS.md` too. Iterative functions (`IRR`, `XIRR`, `RATE`,
`YIELD`) get many inputs, so that a rule for their iteration can be fitted. With
`XLRECALC_CAPTURE_KEEP=1` the pipeline keeps the copy Excel saved, to read the formulas as
Excel wrote them (the file form of a function, for example).

Driving Excel: scripts are files under `capture/`, non-interactive, with `display alerts` off and
restored on every exit path (the setting otherwise persists until Excel quits, and Excel then
silently fails to open or repair files opened by hand); they open only workbooks they create, and
refuse to run when a workbook of the same name is open. One script uses the clipboard:
`paste_formulas.applescript` (below), which saves the clipboard's text first and restores it after
(anything on the clipboard that is not text is lost). With alerts off, a file Excel would offer to
repair fails to open with error -50 ("Parameter error"), which is how probes detect a repair prompt
unattended. `capture/repair_probe.py` and `capture/circular_probe.py
--interactive` are the interactive captures, and whether a dialog appears depends on how the file is
opened, dialog by dialog. Excel shows its repair dialogs for a file AppleScript opens with alerts
on, so `repair_probe` opens it that way and waits for you to accept them. It does not show its
circular-reference warning for a file AppleScript opens, with alerts on or off; it does for one
opened by Launch Services, as a double-click in the Finder opens it (Excel 16.115: a small loop with
`fullCalcOnLoad` showed the warning when double-clicked and none under AppleScript). So
`circular_probe --interactive` opens the file with `open -a "Microsoft Excel"` and, since the
warning does not block requests about the workbook, asks you in a dialog of its own to say when
it is answered (the header records the reply); then it reads and saves the workbook through
AppleScript (`read_save.applescript`). `capture/marker_probe.py --typed` reads a workbook whose
formulas were typed by hand and saved (AppleScript can only enter legacy formulas); that workbook is
kept as `tests/captures/marker_typed.xlsx`.

**Formulas as typed, unattended.** Pasting a column of formulas as text enters them as typing
them does: `capture/paste_probe.py` (`paste_formulas.applescript`: the text on the clipboard,
Excel's `paste worksheet` into A1) built a copy of the typed workbook of `marker_typed` (75
formulas) and of that of `lambda_probe --typed` (6 `LET`, 23 `LAMBDA` formulas and two named
lambdas, the names written as Excel stored them), and every anchor cell matched the typed one on
Excel 16.115: the formula's text with its prefixes (`_xlfn.`, `_xlpm.`, `_xlop.`), its `t` and
`ref`, the dynamic-array marker (`cm`) and the cached value, except `RANDARRAY()`'s random value.
The `lambda_*` captures are made this way (`lambda_probe --paste`); a macro setting `Range.Formula2`
and keystrokes were not needed (and not tried).

**A fresh Excel after LAMBDA captures.** Quit Excel before a capture that follows a LAMBDA capture
(`quit_excel.applescript`, which quits only when no workbook is open): after `lambda_probe`'s recursion cases (deep
recursion, recursion through `MAP` past its limit), the same Excel session computes `#NUM!` for ten unrelated rows of
`empty_args` (`ADDRESS` with a sheet's text, `INDIRECT`, `FLOOR.PRECISE`, `FORMULATEXT`), reproducibly, where a fresh
session gives the captured values.

**Opens.** Excel refuses a whole file for one formula it does not take, so a probe that must name refused formulas
opens them in batches of 60 and halves a refused batch until each refusal is opened alone (`capture/refusals.py`).
Settled answers are not asked again: a refusal of the same build (its list's `# excel:` line) is kept without an open,
and the forms a captured suite holds are confirmed in batches of 1,500, where a refusal is news and is bisected (1,496
forms of `empty_args` open in one workbook in 0.7 s on 16.115); a new form is probed in batches of 60. A refusal of
another build is first set in one open workbook (`capture/entry_probe.py`, `set_formulas.applescript`): a formula Excel
refuses, set through AppleScript's `formula`, leaves the cell unchanged without an error, so the cell read back gives
the verdict. Those Excel still refuses there are kept, and only those it takes now are opened in a file. On 16.115 the
two verdicts agreed for 3,448 of the 3,449 forms `empty_probe` and `database_probe` had settled, set in 234 s; the
one that differs is file-only: `_xlfn.LET(_xlpm.x,1,_xlpm.y,)` opens in a file, and `=LET(x,1,y,)` is refused set in
a cell. So the entry check settles refusals, never acceptances, and a form longer than 255 characters (AppleScript's
limit) is opened in a file. After a build change, `empty_probe` then takes 12 s (one open) and `database_probe` 13 s,
where the file opens took 1,322 s and 231 s. `--full` asks everything again through files. Measured on Excel 16.114.4: a batch of 60
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
`capture/edition.py`). No file on the Mac states the license plainly, so you state the edition
when capturing, as shown in Excel > About Excel, in `capture/build/excel_edition.txt` or the
variable `XLRECALC_EXCEL_EDITION`, and a capture refuses to run without it; the channel is
Microsoft AutoUpdate's.

**After every Excel update**, rerun `capture/open_calc_probe.py` (it rewrites
`tests/captures/open_calc.tsv`): which `calcId` Excel trusts on opening moves with its build, and
`VERIFIED_CALC_ID` in `_opening.py` is set from the lowest one it trusts (`tests/test_open_calc.py`
fails until they agree).

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
Each Excel step there has a time limit (`tests/excel_steps.py`); when Excel then does not answer,
the run ends with a message naming the test, and Excel is left as it is for you to look at. The
pipeline has a limit too, 600 s and 0.1 s per case, and so do the probes that open a workbook and
save a copy.

**When Excel stops answering.** Excel can stop answering questions about its workbooks while it
still answers `version`: at 0% CPU, with no dialog and its menus working, its main thread idle in
its event loop, so the request was never dispatched. A Quit clears it, and the same work then
passes; it can also clear by itself within minutes. A question about workbooks therefore tells
whether Excel is alive (`tests/excel_steps.py`). Some such hangs have no known cause, and each came
with the display off; a locked screen, a display that is off and system sleep, each tried
deliberately, did not cause one. A known cause looks the same: a window of Excel's own (the About
window, Preferences, any dialog) blocks every request about workbooks.

The tools guard against it. A process that drives Excel keeps the display awake (`caffeinate -d`,
from its first request until it ends, after declaring user activity once to wake a display that is
off), and a hang record says whether it did and how the display stood when the run began. Before the
first request, `capture/excel_watch.py` compares the windows Excel has on screen with the document
windows it reports; a window more, or no answer, stops the run with a message naming the windows by
size, before anything is sent. A request past its limit, or one AppleScript gives up on first
(`AppleEvent timed out`, -1712), gets a record under `capture/build/excel_hangs/` (Excel's `sample`,
uptime, CPU and memory, the display's last change, the request and its file), and Excel is asked
about workbooks at once and after 1 and 5 minutes. When it answers, the workbooks the request named
in Excel's container are closed without saving and alerts are turned back on, so a recovered hang
leaves no probe's file open with alerts off. `uv run python -m capture.excel_watch [RECORD]` asks
again and adds the answer to that record, or to the latest if it is from the last half hour (else to
a record of its own): run it before you quit Excel.

The tools do not end Excel, since you may have work open in it: they report and leave the decision
to you, with two exceptions, each only when no workbook of yours can be open in Excel.
`capture/array_size_probe.py` refuses to start with a workbook open, and ends Excel when a step
passes its time limit or Excel passes 8 GB of memory; there, running past a limit is a result, as it
is for `capture/value_probe.py`'s own limit. And a probe whose case can keep Excel calculating
without end (a function at an extreme input; `capture/slow_probe.py`) may let the watch end it when
`may_end` holds (`capture/excel_watch.py`, tested): the run began with Excel not running or holding
no workbook, every Excel window on screen is the request's own (no more of them than its workbooks,
and each title macOS gives one of their names), and one process is named Microsoft Excel. The watch
then sends `TERM`, and `KILL` after 10 seconds, records it and goes on; the next request checks
Excel's windows again (a recovery window would stop the run). In any other state it stops and asks.

The probes behind the design (`capture/writeback_probe.py`, `capture/semantics_probe.py`,
`capture/formula_forms.py`) use the same conventions; their results are in [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md)
§5.2 and §9.4. `capture/vm_probe.py` is such a probe; its output (`capture/build/vm_probe.out`) is
not kept, and its result is in ARCHITECTURE.md §7.7: Excel 16.115 shows and saves the rich value's
error, not the new `<v>`, for a cell given a new value that keeps the old value's `vm` (a number, a
logical, `#DIV/0!` and a plain `#VALUE!` alike), so a new value drops `vm`
(`tests/test_error_types.py`).

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
- A downloaded workbook is never opened in Excel. It is read only by this repository's tools (with xlio, the
  reader this package uses, openpyxl or the standard library's `zipfile`), in a directory of its own, and only
  `.xlsx` files are read (never `.xlsm`).
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

## 5. Captures made by a job on Windows

A capture can also be made by an Excel capture job, in the format sheetmill describes
([docs/capture-jobs.md at 3adb865](https://github.com/sheetmill/sheetmill/blob/3adb865500f4d2223e5c1a758b466e081d3fdc79/docs/capture-jobs.md),
version 1): a manifest of operations beside the input files, run by a program that drives Excel
and records the machine. `capture/job.py` writes the job for a suite of `capture/suites.py` (the
manifest opens the suite's workbook and saves a copy; the workbook is built as the pipeline
builds it, with fixed zip times and document properties, so that writing it again gives the same
bytes) and turns the result into the capture: each formula's cached value and type, read from the
saved copy as the pipeline reads them (for a suite of `capture/dynamic_suites.py`, named
`dyn_<suite>`, the workbook `capture/dynamic_probe.py` builds and every cell of the copy, as that
probe records them), under the result's `header` copied whole (the format
version, the runner's version, the job, the SHA-256 of the manifest and of each input, and the
machine's record). The manifest is committed beside the capture, byte for byte. A job states
what it requires (`requires`): the edition, Microsoft 365, since it decides which functions Excel
has, and the first Windows build these captures were made with, 16.0.20527.0; the first job,
`collate_words`, stated neither. `tests/test_windows_captures.py` writes each job again from its
header's `job_id` and its manifest's `requires`, and requires the same manifest and inputs,
checksum for checksum, and the generator's cases in order. It also computes each suite on an
engine for Excel for Windows and classifies every row against the Windows capture as
`tests/test_captures.py` does against the Mac's (`mac_gaps` is compared on neither platform, and
`whole_columns` in `tests/test_whole_columns.py` and `dyn_cancel` in
`tests/test_dynamic_arrays.py`, on both).

| capture | rows | what it records |
|---|---|---|
| `tests/captures/windows/clean_controls.tsv` (job `windows/clean_controls.job.json`) | 510 | the suite `clean_controls` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale en-US: `CLEAN` removes codes 1 to 31 and U+0080 to U+009F and keeps 127, as on the Mac; no row differs from the Mac capture. The package, computing for Windows, gives every row |
| `tests/captures/windows/codepage.tsv` (job `windows/codepage.job.json`) | 912 | the suite `codepage` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale en-US, ANSI code page 1252: `CHAR` is Windows-1252, its five undefined codes the C1 controls of those codes; `CODE` of a character outside it is 63, of a C1 control its own code; 349 rows differ from the Mac capture. The package, computing for Windows, gives every row (`tests/test_windows_captures.py`) |
| `tests/captures/windows/collate_chars.tsv` (job `windows/collate_chars.job.json`) | 23,375 | the suite `collate_chars` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US: one order over the 125 characters, `'` and `-` weighing almost nothing, a marked letter right after its base, case equal, `ª` and `º` not equal to `a` and `o`; 806 rows differ from the Mac capture. The package, computing for Windows, gives every row (`tests/test_windows_captures.py`) |
| `tests/captures/windows/collate_functions.tsv` (job `windows/collate_functions.job.json`) | 169 | the suite `collate_functions` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US: `SORT`, ordered criteria, approximate and binary lookups follow Windows' `<`; criterion equality, the binary search and `UNIQUE` follow its `=` (`æ`=`ae`, `œ`=`oe`, `ß`=`ss`: 20 distinct entries of 24, the Mac 22); exact `MATCH` ignores case only; 53 rows differ from the Mac capture. The package, computing for Windows, gives every row (`tests/test_windows_captures.py`) |
| `tests/captures/windows/collate_words.tsv` (job `windows/collate_words.job.json`) | 2,380 | the suite `collate_words` computed by Excel 16.0.20527.20196 (Microsoft 365, Current Channel) on Windows 26H2, locale and sort locale en-US, ANSI code page 1252; 134 rows differ from the Mac capture `collate_words.tsv`. The package, computing for Windows, gives every row (`tests/test_windows_captures.py`) |
| `tests/captures/windows/modern.tsv` (job `windows/modern.job.json`) | 2,230 | the suite `modern` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US: 173 rows differ from the Mac capture, every one a comparison of text (`compare_text`: 172 `<`, and `"ae"="æ"` TRUE); the rest are the Mac's. The package, computing for Windows, gives every row but the four listed as its deviations on the Mac (unions of ranges in `LARGE` and `SMALL`; `XLOOKUP`'s and `XMATCH`'s regular-expression mode), where Excel for Windows gives the Mac's value (`tests/test_windows_captures.py`) |
| `tests/captures/windows/function_names.tsv` (job `windows/function_names.job.json`) | 511 | the suite `function_names` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US: 2 rows differ from the Mac capture, `_xlfn.ENCODEURL` and `_xlfn.FILTERXML`, which Excel for Windows knows; written without the prefix they give `#NAME?`, as `EUROCONVERT` and `JIS` do in both forms |
| `tests/captures/windows/dyn_cancel.tsv` (job `windows/dyn_cancel.job.json`) | 340 cells | the suite `dyn_cancel` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale en-US: every cell holds the value the Mac capture holds |
| `tests/captures/windows/bare_names.tsv` (job `windows/bare_names.job.json`) | 153 | the suite `bare_names` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale en-US: `#NAME?` for every bare name, as on the Mac |
| `tests/captures/windows/<suite>.tsv` for the other 47 suites of `capture/suites.py` (each with its job beside it) | 91,941 | each suite computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US. 38 suites give every value the Mac capture holds. The others differ where the platform's code page, names or math library do: `CHAR` and `CODE` (`misc` 128 rows, `text_family` 6), `ENCODEURL` and `FILTERXML` (`mac_gaps` 2), the last digits of `ATAN`, `ATAN2`, `TAN`, `COT`, `ACOT`, `LOG`, `LOG10` (`math_trig` 38, `empty_args` 2; 1 or 2 units in the last place), the t distribution, `GAMMA.DIST`, `CONFIDENCE.T` (1 or 2 units) and `BETA.INV`, `F.INV` near 1E-150 (2 rows each, up to 2,551 units; `continuous_family` 27), `BESSELJ` and `BESSELY` (`bessel_family` 45, up to 4 units), the 15th digit of one `IMCOSH` text (`complex_family`), `WEEKNUM` with the undocumented types 22 and -1 (`dates_family` 10, which vary between runs on the Mac too) and random draws (`misc` 1, `empty_args` 4) |
| `tests/captures/windows/encodeurl.tsv` (job `windows/encodeurl.job.json`) | 177 | the suite `encodeurl` computed by Excel 16.0.20527.20196 (Microsoft 365) on Windows, locale and sort locale en-US: the argument as text (a number as Excel shows it, a logical as `TRUE`/`FALSE`, a blank as empty, an error passed on, a range by implicit intersection), its UTF-8 bytes percent-encoded with upper-case hex but for `A-Z a-z 0-9 - _ .` (`~` is `%7E`, a space `%20`); a result of 32,767 characters or more is `#VALUE!` whatever the argument's length (10,923 characters encoding to 32,767 included), and one of 32,766 is computed (a 32,766-character argument encoding to itself included) |
