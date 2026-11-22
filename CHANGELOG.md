# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.8.0] - 2026-11-21

### Added
- `Workbook.status(cell)`: what `Workbook.report` says of one cell, as a `CellStatus` (`state`: `"refused"`,
  `"uncomputed"`, `"circular"`, `"kept"`, `"computed"` or `"constant"`; the cell's own `labels`, `error` or
  `refusal`; `inherited`, when an uncomputed or kept state comes from a precedent; `anchor`, for a cell an array
  formula fills). It reads only the sheets that cell needs (its own, and those its evaluation reaches, through every
  reference of the formulas evaluated, branches not taken included), where `report` reads every sheet a lazy `Loader`
  left, and it equals the report's entry for the cell. Reading a range's values first and then each cell's status
  runs the report's passes once.

### Changed
- Reading formulas filled down a column without being shared (each cell holding its own text, as some programs
  write them) computes their template keys about twice as fast: a formula whose text differs from one met before
  in the same column only by rows moved with it takes that formula's key without scanning its text. A `Loader` fed
  every cell of a sheet of such formulas (`bench/values_bench.py`, `unshared_form_N.xlsx`, three formulas a row) takes
  1.32 to 1.33 s for 100,000 rows, from 1.48 to 1.49 s, and 13.4 to 13.5 s for 1,000,000, from 15.2 to 15.3 s; reading
  the same cells through `xlio` alone takes 0.83 s and 8.2 s. Shared formulas were already read once a formula.

### Fixed
- A cell depending on a circular reference keeps the value the file holds, as Excel keeps it (capture
  `circular_cached`), when a `Loader` reads its sheet only after the cycle was evaluated. In 0.7.0, a workbook with
  iterative calculation off and a cycle on a sheet the caller fed, read through `Loader.workbook()` with the dependent's
  sheet not fed, computed that cell from the cycle's kept value: `one!A1 = two!A1+1`, read after `two`'s cycle
  (cached 7), gave 8. `load()`, `finish()` and a `Loader` fed every sheet first gave 7.
- With a `Loader`, a sheet read while a cycle is evaluated keeps the cells the cycle keeps on it: an array formula
  depending on the cycle there was evaluated again on the next read, and the report listed it and the cells it read
  as uncomputed where `load()` lists them as circular. The values were the same.
- `Report.unsupported` counts each cell once for each function or feature it uses that is not implemented, as its
  docstring says: a count grew with every recalculation that evaluated the cell, and a cell an edit held back, or one
  uncomputed through a precedent, was not counted unless it was evaluated.
- A data table, or an array formula's range with no formula text, keeps its cached values and is now in
  `Report.from_kept`, with the cells computed from it, as `from_kept`'s docstring says; it was in none of the report's
  sets, and `Workbook.status` gave it as `"constant"`.

## [0.7.0] - 2026-11-21

### Added
- `Loader.workbook()`: the workbook at any moment, also before every sheet is fed; a sheet not fed, or left part-way,
  is read from the book then or when a value needs it (below), so once the book is closed, a read of a sheet not read
  yet raises `ValueError`. `Loader.finish()` returns the same object after reading every sheet not read yet, so the
  book may be closed after it.
- A `Loader` whose file makes every formula computed (an older Excel's file, one another program wrote,
  `fullCalcOnLoad`, `recalc_on_load="always"`; not with iterative calculation) reads a sheet not fed only when a value
  needs it: through a reference, a range, a defined name, `INDIRECT`, `OFFSET` or `CELL`'s active cell. A sheet no
  value reaches is never read. `Workbook.report` and `Workbook.recalculate` read every sheet first.
- A `Loader` whose opening rule answers "none" (every cached value stands: no formula without one, none marked
  `ca="1"`, no volatile function unmarked in a formula's own text) reads a sheet not fed only when a value needs it, and
  compiles a sheet's formulas only when they are needed (`Workbook.formula`, `FORMULATEXT`, the report, a
  recalculation, an edit), fed sheets included: a file Excel saved, read for its values, compiles nothing. Reading
  every value of the 1M-row benchmark in the form Excel saves (`bench/values_bench.py`) through a `Loader` takes 12.6
  to 12.7 s and 824 to 832 MB at peak, from 16.8 s and 1,284 to 1,287 MB.
- `Loader.opening` and `Loader.workbook()` decide the opening rule from `xlio`'s formulas-only reading
  (`Sheet.formulas`) of the sheets not fed, before any of their values is read.
- `Loader.opening`: what the load computes on opening, `"all"`, `"some"` or `"none"`, from the same rule and facts as
  `RecalcCheck` (`None` while cells not read yet decide it), so a caller using a `Loader` needs no `RecalcCheck`.

### Changed
- `Loader.cells` takes the sheets in any order and at any time, also after `workbook()` or `finish()`: a sheet the
  loader has read already passes its cells through unrecorded instead of raising, and a loop the caller is still
  running when the loader reads its sheet hands on the rest of its cells, unrecorded, where it stopped before.
- Requires `xlio` 0.4 (`xlio==0.4.*`): the `Loader` asks `Workbook.closed` whether it can still read a sheet, and
  reads formulas alone through `Sheet.formulas`.
- Stale cells are evaluated in an order set by which cells are stale (store order past an eighth of the formulas, else
  sheet, row and column), not by the order they became stale, so step limits and `RAND`'s draws do not depend on it.
  Re-evaluating 1,000,005 cells after an edit on the 1M-row benchmark takes 5.1 to 5.2 s, from 5.2 to 5.3 s.
- `Report.refused` and `Report.exceptions` list their cells in sheet, row and column order, and `Report.unsupported`
  its labels in text order, whatever order the cells were compiled and evaluated in; they followed that order.
- Loading records a sheet's values as they are read and compiles its formulas afterwards, every sheet in the
  workbook's order, whatever order the sheets were fed in. `xlrecalc.load` of the 1M-row benchmark, then reading every
  value (`bench/values_bench.py`), takes 30.3 to 30.9 s, from 32.5 to 32.9 s, and 1,311 MB at peak, from 1,231 to
  1,235 MB; of the same table in the form Excel saves, 17.1 to 17.2 s, from 16.7 to 17.0 s, and 945 to 948 MB, from
  1,155 to 1,171 MB.
- `SUMPRODUCT` over whole columns, alone or under operators (`SUMPRODUCT((A:A="x")*B:B)`), reads each column down to
  its last value instead of a million cells: 3 ms where it took 2.6 to 3.3 s on a sheet of 1,000 rows, and 27 MB at
  peak where it took 366 to 368 MB (`bench/values_bench.py`, which times every form below).
  Results are the same; `SUMPRODUCT` over every column of a sheet is now computed, as Excel computes it (capture
  `array_size`), where it passed the array limit and kept the file's value.
- `SUM`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `AND` and `OR` over whole columns under operators in an array formula
  (`{=SUM(--(A:A<>""))}`, a dynamic array) read each column the same way: 1 to 116 ms where they took 1.0 to 2.6 s on
  a sheet of 1,000 rows, with the same results.
- `INDEX`, `LOOKUP`, `XMATCH`, `ROWS`, `COLUMNS` and, in an array formula, `MATCH` over whole columns under operators
  (`LOOKUP(2,1/(A:A<>""),A:A)` for the last value of a column, `MATCH(TRUE,A:A="",0)`) read each column the same way:
  at most 2 ms where they took 1.2 to 2.4 s on a sheet of 1,000 rows, with the same results.
- `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISNA`, `ISLOGICAL`, `LEN` and `ABS` over whole columns, read by one of
  these functions (`SUMPRODUCT(--ISBLANK(A:A))`, `MATCH(TRUE,ISBLANK(A:A),0)`), apply to each column's rows with data and
  once to the rest: 2 to 115 ms where they took 1.7 to 2.9 s on a sheet of 1,000 rows, with the same results.
- `IF` over whole columns in an array formula read by one of these functions (`{=SUM(IF(A:A="k",B:B))}`,
  `{=MAX(IF(A:A<>"",B:B))}`) does the same: 3 to 116 ms where it took 2.9 to 3.7 s on a sheet of 1,000 rows.

### Fixed
- An edit leaves a volatile cell whose cached value a trusting load kept, until the cell is computed, as Excel leaves
  one volatile through a defined name without `ca="1"` (capture `volatile_edit`); a marked cell, and one Excel has
  computed since opening, are computed after every edit. Whether an edit computed such a cell depended on whether a
  value had been read before it.
- Two formulas that differ only in spacing (`A1+1` in one cell, `A2 + 1` below it) each show their own text in
  `Workbook.formula` and `FORMULATEXT`, as Excel keeps spaces as written (capture `formula_bar`); both showed the spelling
  of whichever was compiled first.
- With iterative calculation, reading a value that needs a loop gives the run Excel makes on opening (captures
  `iterate_calc`, `iterate_reads`): every stale cell is evaluated first, then every loop runs once, together, with one
  stopping test for all of them. A read that met one loop ran
  it alone, and a later read meeting another loop ran the first again, so values depended on the order they were read
  (`B1=B1+1` and `C1=C1+1` from 0 at 100 iterations: 100 and 100 read together, `B1` 200 after reading `C1`).
- A volatile function in a cell's own text without `ca="1"` makes a load compute every formula (capture `volatile_names`)
  also where xlrecalc's parser cannot read the rest of the formula (`RAND()+`): such a formula was left out, so the
  load kept cached values that `RecalcCheck` answered would be computed. Both now decide from the text alone.

## [0.6.0] - 2026-11-20

### Added
- `Loader.cells(sheet, rich=True)`: each cell as `xlio`'s `Sheet.cells(rich=True)` gives it, with its text's runs and
  phonetic guides, for a caller that renders them; what the loader records is the same either way.

### Fixed
- `MATCH` given one value as its lookup array, not a reference or an array, converts it to a number as Excel does:
  `#VALUE!` for a text that is not one (`MATCH("a","a",0)`, or `MATCH("",A:A&"",0)` in a plain formula, which reads
  its own row of `A`), where it gave `#N/A`; a number, a logical or a numeric text still gives `#N/A` (capture
  `whole_columns`).

## [0.5.0] - 2026-11-19

### Added
- `ConditionalFormula`: where Excel shows a conditional format's rule, cell by cell. `parse(formula, sqref=, sheet=)`
  takes an expression rule's formula and `cell_is(operator, formulas, sqref=, sheet=)` a `cellIs` rule's operator and
  bounds, as a file stores them; `applies(cell, values, resolver=, names=)` evaluates the rule there. Excel's rules
  from capture `cf_rules` (Excel 16.115 for Mac, 76 cases, every readable cell matched): relative references are
  measured from the top-left corner of the box around all the range's areas, whatever their order; the formula is a
  dynamic array whose top-left element counts; the format applies where `IF` takes its true branch (TRUE, a nonzero
  number, the text `"TRUE"`), not for an error or other text; `between` and `notBetween` take their bounds in either order. A formula
  holding an array constant or a structured reference, which make Excel offer to repair the file, is refused with
  `FormulaError`.
- `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, with Excel's values: the approximations Excel uses, those of Numerical
  Recipes in C (2nd edition, sections 6.5 and 6.6), evaluated with their published coefficients in the package's own
  code, with the four places where Excel's results differ from the printed approximations (QUIRKS.md). Capture `bessel_family` (Excel
  16.115 for Mac, 4,359 rows: x from 0 to 1E308, n from 0 to 2^32, coercion, errors and the limits of x and n): every row matches bit for bit
  on macOS; `BESSELJ` and `BESSELY` allow 8 units in the last place for the platform's sine and cosine.
- `CONVERT`, every unit Excel documents with its aliases, decimal and binary prefixes, and temperatures by Excel's
  formula for each pair of scales. Capture `convert_family` (Excel 16.115 for Mac, 2,558 rows): every conversion to
  or from a category's base unit, every prefix and every temperature matches bit for bit; between two other units,
  839 of 1,056 do and the rest are within 2 units in the last place, as declared (QUIRKS.md).
- `ReadLimits.max_function_steps` (1,000,000,000 by default): the iterations that loops inside functions may take in
  one recalculation, an allowance shared by its cells and charged before each loop runs. A cell whose function would
  pass it keeps the value the file holds and is reported (`Report.uncomputed`; `Report.unsupported` names the limit).
  Its first users are the Bessel recurrences, about 20 million steps (0.7 to 1 s) for a cell at the largest order
  Excel computes, so a crafted file of such cells stops at the allowance instead of running for hours.
- `Loader`: a workbook read from cells the caller already streams through `xlio`, so the file is read once.
  `Loader(book, recalc_on_load=, limits=)` takes an open `xlio` workbook; `cells(sheet)` passes each cell through
  while recording it, as `csv.reader` leaves the loop to its caller; `finish()` gives the `Workbook` that `load` gives,
  reading again any sheet not fed or left part-way. `load` is a `Loader` with every sheet fed in order. Such a
  workbook has no file to copy, so `save` raises.
- `ReadLimits.UNTRUSTED`, a preset for files from strangers, stricter than the defaults on every field: `xlio`'s
  archive limits (10,000 parts, 256 MiB a part, 1 GiB in all), arrays of one full column (1,048,576 cells),
  1,000,000 iteration evaluations and 20,000,000 function steps, about a second each.

### Fixed
- Reading cells one by one with `Workbook.value` was quadratic when formulas rest on a value the package keeps (a
  function it does not compute): each read of a cell that needed computing finished the report again, walking every
  cell resting on the kept value (8,000 such cells took 21 s; a corpus workbook over an hour). The report's
  whole-workbook passes now run once, when `Workbook.report` is read or the workbook is saved; the reads cost what
  one recalculation costs (8,000 cells: 0.05 s). The report is the same either way, and `recalc_on_load` is unchanged.
- The README and the glossary said that in manual calculation every cached value is kept, for `load` and
  `recalculate` alike. That holds for `load`; `recalculate`, which saves, computes what Excel recalculates when it
  saves such a file (the cells marked `ca="1"`, or every formula), as the 0.4.0 entry says. They now state the two
  cases apart, and a test pins what `recalculate` writes for a manual-calculation workbook with marked cells.

## [0.4.0] - 2026-11-18

### Added
- `RecalcCheck`: whether `load` would compute any formula of a workbook, `"all"`, `"some"` or `"none"`, without
  building an engine, for a program that reads the cells itself. Created from the file (`source`, with
  `recalc_on_load`, `saving` and `limits` as `load` takes them), it reads only the workbook part, which alone decides
  most files (14,067 of 15,871 workbooks of the Enron spreadsheet corpus, by an older `calcId` or `fullCalcOnLoad`);
  it is then fed the formula cells the caller reads (`add(formula, cached=, ca=)`, then `finish()`), as `hashlib`'s
  objects are fed, and `result` is `"all"` as soon as that is certain. Its answer comes from the same rule as
  `load`'s.

### Changed
- `recalc_on_load="excel"` (the default of `load` and `recalculate`) on a file Excel saved with a volatile function
  computes what Excel recalculates on opening, the cells whose formula Excel marked `ca="1"` (every volatile cell and
  every cell depending on one), and keeps every other cached value; it computed every formula. A volatile function in
  a cell without the mark, as another program may write it, still computes every formula, as in Excel; in manual
  calculation the same holds for the copy a save writes. Capture `volatile_names` (18 variants, Excel 16.115 for
  Mac); the earlier rule came from capture `open_calc`, whose volatile cell had lost Excel's mark when its value was
  planted.
- `recalc_on_load="never"` computes a formula volatile through a defined name, as it computes `=OFFSET(...)`.

### Fixed
- A formula volatile only through a defined name (`N` = `OFFSET(...)`, a chain of names, a `LAMBDA` name calling
  `NOW()`) counts as volatile: it is computed again after a change and at every `recalculate()`, with its dependents;
  it kept its cached value. A function name inside a string literal is not a call, and a volatile name no formula uses
  changes nothing.
- `CODE` of a C1 control character (U+0080 to U+009F) is its code, 128 to 159, as in Excel; it was 95, the code of
  a character outside Mac OS Roman. `CLEAN` removes those characters as well as the codes below 32, as Excel does; it
  kept them (captures `codepage`, `clean_controls`).
- Text compares, sorts and looks up in Excel for Mac's order for Latin-1 signs and letters and `œ` as well as ASCII:
  `¡ ¿ « » § ¶` among the marks, `æ`, `ð`, `ø`, `œ`, `þ` and `µ` as letters of their own, `ª` and `º` equal to `a` and
  `o`; they were ordered by code point after the letters (captures `collate_chars`, `collate_words`: 23,375 and 2,380
  rows, where 1,918 and 84 differed).
- A criterion's text equality is the `=` operator's, as in Excel: `COUNTIF(r,"ß")` counts `ss`; it counted only `ß`.
  An exact `MATCH`, `VLOOKUP` or `XLOOKUP` scan ignores case only, as in Excel: `MATCH("ß",r,0)` passes over `ss`;
  it found it. `XMATCH` and `XLOOKUP`'s binary search finds the first of a run of equal entries and compares as `=`
  does, as in Excel; it could find a later one (capture `collate_functions`: 169 rows, where 5 differed).

## [0.3.0] - 2026-11-17

### Added
- `xlrecalc.evaluate(formula, values, *, resolver=..., at=..., names=...)` and `xlrecalc.Formula.parse(...)`:
  one formula evaluated outside a workbook, compiled once and evaluated at any cell, relative references moving with
  it (as a conditional format's formula does over its range). Cell values come from a mapping (cells, ranges,
  sheet-qualified) and then a callback given a `Reference`; a cell neither gives is blank, as in Excel. A formula
  starting with `=` is read as typed, one without it in file form. Dynamic-array evaluation by default, `legacy=True`
  for implicit intersection; a dynamic result comes back as rows. A name Excel does not know gives `#NAME?`; a
  function or form Excel computes and the package does not (a union of ranges, a structured, 3-D or external
  reference, `CELL`'s formatting kinds) raises `UnsupportedError` (a `FormulaError`); `Formula.unsupported` names
  beforehand what the text shows. A captured formula evaluated alone gives the engine's value, in every
  suite whose inputs are values (not formulas or fixed cells), the unsupported and volatile ones left out.
- `LAMBDA` and `ISOMITTED`, captured from Excel (captures `lambda_calls`, `lambda_names`, `lambda_depth`,
  `lambda_more`): a function value called where it is made (`LAMBDA(x,x+1)(2)`, `(LAMBDA(x,x))(2)`), through a `LET`
  name or a parameter, or through a defined name that holds one (`ADDONE(1)`), recursion and mutual recursion
  included. It closes over the names around it; optional parameters (`[y]`) and arguments given empty read as a
  blank, and `ISOMITTED` tells them apart; a wrong argument count is `#VALUE!`; a call returns references as they are,
  and a call of anything but a function value is `#VALUE!`. A function value as a cell's value is `#CALC!`, given to an
  operator or a function `#VALUE!` (the `IS` functions FALSE, `ISNONTEXT` TRUE, `TYPE` 128). Recursion runs on an
  explicit stack and stops with `#NUM!` at Excel's budget: 65,536, each call in progress costing 8 plus 4 a parameter
  (optional ones included) and 4 a `LET` name its body binds (5,461 calls of one parameter); where Excel stops a
  call or two earlier, or runs a call in tail position past the budget, the package charges it as any other (5 of
  the 233 captured formulas, listed in QUIRKS.md). Excel refuses a file
  with an empty `LAMBDA` parameter or body, `ISOMITTED()` or `REDUCE` of two arguments, and so does the package.
- `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` and `MAKEARRAY`, captured from Excel (captures `lambda_helpers`,
  `lambda_more`, `empty_args`): `MAP` passes a range's elements as references and an array's as values, padding a
  smaller argument with `#N/A`; `REDUCE` and `SCAN` go row by row from their initial value; `BYROW` and `BYCOL` pass
  each row or column; `MAKEARRAY` takes positive whole sizes; an empty array argument is `#VALUE!`. A call whose
  result is not one value makes the result `#CALC!` (`REDUCE`'s last excepted). Twenty such functions may call one
  inside another through a recursion; the next call is `#VALUE!`, as in Excel. Written without `_xlfn.`, each is a
  name Excel does not know (`#NAME?`), and is reported with that cause (capture `prefixes`).
- A formula set through the API that calls a function value, or gives an uncalled `LAMBDA` as its result, is a
  dynamic array, as Excel marks it typed (every typed formula of the `lambda_*` captures).

### Changed
- xlsx files are read and written by [xlio](https://github.com/sheetmill/xlio), a new required dependency
  (`xlio==0.3.*`), which holds the reader and surgical writer this package had in `xlrecalc._xlsx`; that private
  subpackage is gone. `PackageError` and `LimitError` are xlio's, re-exported, so `except xlrecalc.PackageError` works as
  before; `ReadLimits` keeps its five fields and import path. Every value computed and every byte written is unchanged on
  the capture and benchmark workbooks. Two readings follow Excel where the old reader did not (xlio's captures): a cell
  whose format or metadata index is past its table makes the file refused (`PackageError`), as Excel refuses it, where
  it was read; the active cell is `activeCell` when it lies in the selection, else the selection's first cell.
- `save(..., full_calc_on_load=False)` and `recalculate(..., full_calc_on_load=False)` save the file without a
  request to recalculate on opening, removing one the file held, as Excel removes it when it saves; they left it.
- A load trusts a file's cached values where Excel does on opening it, and evaluates every formula where Excel
  recalculates (capture `open_calc`, Excel 16.115): when the file has no `calcPr`, has `fullCalcOnLoad`, was last
  calculated by an older Excel (`calcId` below 191028, `VERIFIED_CALC_ID`), was not saved by Excel (no
  `<fileVersion appName="xl">`: openpyxl writes none), or holds a volatile function, which
  recalculates every formula, not only its dependents; never in manual calculation. Before, only `fullCalcOnLoad` made
  a load evaluate everything. The new option `recalc_on_load` of `load` and `recalculate` chooses: `"excel"` (the
  default), `"always"` (evaluate every formula) or `"never"` (trust every cached value the file holds; a formula
  without one, or a volatile one, is evaluated all the same; in manual calculation `"excel"` keeps a volatile one's
  cached value too, as Excel shows it on opening). A `calcPr` without `calcId` counts as none. In manual
  calculation Excel shows the cached values on opening and saves them recalculated where a reason holds, so a load
  trusts them and `recalculate(path)` recalculates.
- Excel 16.115 for Mac under Microsoft 365 is the reference: every capture made on 16.114 (Office Home & Business 2021)
  was rerun on it. Its differences: `LEN`, `MID`, `REPLACE`, `FIND` and `SEARCH` count a character outside the first
  plane (an emoji) as one, as `LEFT` and `RIGHT` did already (`LEN` of an emoji is 1, was 2), and so does the package;
  `LENB`, `MIDB` and `REPLACEB` still count its two UTF-16 units, the limit of a text, 32,767, counts units too, and a
  wildcard's `?` matches one unit (capture `text_family`, 151 rows added). The `LAMBDA` helpers compute there, and
  `LAMBDA` typed in its formula bar is written with `_xlfn.` and `_xlpm.` (captures `prefixes`, `formula_bar`). Every
  other capture is unchanged except for random results.
- The coverage page (`docs/coverage.md`) counts a captured row that matches Excel exactly or within the function's
  declared tolerance as one figure, so it is the same on every platform; the split between them depends on the platform's
  math library. The rows that depend on that library are held strictly only on a named reference platform (macOS 26.5,
  arm64, Python 3.11): Python 3.14 on the same system brings four `BINOM.DIST` rows of `distributions_family` within 8
  units in the last place that 3.11 leaves past them.

### Deprecated
- `load(..., trust_cached=...)` and `recalculate(..., full=...)`, to be removed in 1.0: they warn (`DeprecationWarning`) and
  map to `recalc_on_load` (`trust_cached=False` and `full=True` are `"always"`; `trust_cached=True` and `full=False`
  the default, `"excel"`). `Workbook.recalculate(full=True)` stays.

### Fixed
- `ERROR.TYPE` gives Excel's numbers for `#CONNECT!` (10), `#BLOCKED!` (11), `#UNKNOWN!` (12), `#FIELD!` (13),
  `#BUSY!` (16) and `#TIMEOUT!` (20), where it gave `#N/A`, and a file's rich-value errors of those kinds are read as
  them, where they were read as `#VALUE!` (xlio's capture [`errors`](https://github.com/sheetmill/xlio/blob/main/tests/captures/errors.tsv): Excel 16.115 reading
  each from a file). `#GETTING_DATA` is written as `#N/A`, as Excel writes it back.
- A formula reading a defined name that refers to another name, or calls one holding a `LAMBDA`, is recalculated when
  a cell the other name reads changes; only the first name's own references were followed.
- A wildcard's `?` in criteria (`COUNTIF` and the other `IF` functions, the database functions), `MATCH`, `XMATCH`
  and `XLOOKUP` matches one UTF-16 unit, as in Excel, so a character outside the first plane takes `??`; it took `?`
  (capture `text_family`).
- The limit of a text in `&`, `CONCATENATE`, `CONCAT`, `TEXTJOIN` and `REPT` counts UTF-16 units, as Excel does; it
  counted characters (capture `text_family`).
- A formula cell whose cached value Excel had stored as a rich value (a newer error such as `#SPILL!` or `#CALC!`,
  cached as `#VALUE!` with `vm`) and that the package gave a new value kept its `vm`, in 0.2.0 and before: Excel,
  opening the file without recalculating (`full_calc_on_load=False`), showed and saved the old error in place of the
  new value, and a reader of rich values (this package's included) read it. A new value now drops `vm`; a cell still
  holding the error its rich value stands for keeps it (`capture/vm_probe.py`, Excel 16.115).
- A sheet listed in the workbook without a part (`r:id=""`), which is skipped, no longer moves sheet-local defined
  names and the active sheet onto the next sheet: `localSheetId` and `activeTab` count every listed sheet, and a name
  local to a skipped sheet is dropped. In 0.2.0 and before such a name was defined on the following sheet.
- `PRICE`, `YIELD`, `DURATION`, `MDURATION` and the `COUP` functions take a logical settlement or maturity as
  `#VALUE!`, as Excel does (capture `bonds`); `TRUE` was read as the date 1, and the result reached `#VALUE!` only
  through an internal error. A coupon period that would begin before 1900 (1904) is `#NUM!`, as in Excel, except
  in `COUPDAYS` outside basis 1, which needs no coupon date and gives the period's length: 365 days (basis 3) or 360
  (the others) divided by the frequency.
- A dynamic array that a trusted file caches as blocked (`#SPILL!`) spills when the cell in its way is cleared; in
  0.2.0 and before it stayed `#SPILL!` until a full recalculation.

## [0.2.0] - 2026-11-10

### Added
- Iterative calculation, for a workbook with iteration on (`<calcPr iterate="1">`), as Excel does it (captures
  `iterate_calc`, `calc_chain`, `iterate_open`): a run of up to `iterateCount` iterations from the values the cells hold, each
  computing the cells on and after circular references in place, in Excel's order (sheets by name, then row, then column),
  until the largest change of a number is below `iterateDelta`. A run happens where Excel runs one: an evaluating load
  (as an opening), each edit and each recalculation, and each save once the loop has run, so `Workbook.save` then changes
  the values in memory too; a trusting load with no edit runs nothing. `recalculate(path)` of a file it evaluates is a load and a save, two runs, which is what Excel
  saves after opening the file, whether or not the loop has cached values; a loop with none starts from the empty text, as
  in Excel (captures `iterate_uncached`, `iterate_opening`, `iterate_cached`). `ReadLimits.max_iteration_evaluations` (100,000,000 by default) bounds the cell evaluations of one run;
  past it, the cells keep their values and are reported.
- `LET`, captured from Excel (captures `dyn_let`, `let_typed`): names bound to values for a calculation, compared
  without regard to case, an inner `LET`'s shadowing an outer one's and a defined name; only names written with
  `_xlpm.` are `LET`'s, as Excel reads a file. A binding is evaluated when first read and once, so one nothing reads
  is never evaluated (a binding reading the formula's own cell is no circular reference); a name read before it is
  bound is `#NAME?`. A bound range stays a reference, spills, and is intersected in a plain formula. With no
  calculation after the bindings the last name is the calculation, as Excel writes it back. An empty value binds a
  blank (`LET(x,,ISBLANK(x))` is TRUE, `LET(x,,x+1)` 1) and an empty calculation is 0. A file with a name bound
  twice, a name without `_xlpm.`, or a number where a name goes does not open in Excel: such a formula keeps the value
  it holds and is in `Report.refused`. `set_formula` writes `LET` with the dynamic-array marker only where its
  calculation needs it (`LET(x,SEQUENCE(3),SUM(x))` is a plain formula, as Excel writes it typed).
- Functions that make arrays, captured from Excel as dynamic arrays (captures `dyn_*`, two sets of values each):
  - `SEQUENCE`: counts are truncated, an empty one is 1, 0 is `#CALC!` and a negative or too large one `#VALUE!`; each
    number is the one before plus the step; an array argument gives each element's first number, a range too in a
    dynamic array, where a plain formula intersects it.
  - `MMULT`: every element must be a number (`#VALUE!` for a text, a blank or a logical, the error for an error) and the
    sizes must fit; each element is summed in order, as Excel adds.
  - `UNIQUE`: rows, or columns with `by_col`, in the order they first appear, text compared without regard to case and
    kept as first written, numbers and numbers as text apart, `exactly_once`; none left is `#CALC!`.
  - `SORT`: numbers, then text in Excel's order without regard to case, then logicals, then errors by their code, blanks
    last in either direction; equal keys keep their order; several keys with their own orders, `by_col`.
  - `SORTBY`: the same order over other rows or columns, each with its own direction; a key of another length or shape is
    `#VALUE!`.
  - `FILTER`: the rows, or columns, whose include is TRUE or a number other than 0 (a text is `#VALUE!`, an include of
    another length or two-dimensional is `#VALUE!`); when none is left, `if_empty`, an array included, or `#CALC!`.
  - `CHOOSECOLS`, `CHOOSEROWS`: the columns or rows at the indexes given, negative ones from the end, repeats kept; a row
    or column of indexes picks several, a two-dimensional one, or an index of 0 or past the size, is `#VALUE!`.
  - `TAKE`, `DROP`: the first or last rows and columns, or all but them; a count past the size takes all, dropping all
    or taking none is `#CALC!`. Of a range they give a range, which a plain formula's cell intersects.
  - `EXPAND`: padded to a size with `pad_with`, `#N/A` when it is left out and 0 when it is empty; a size smaller than
    the array is `#VALUE!`, one past the sheet `#NUM!`.
  - `VSTACK`, `HSTACK`: one below another or side by side, padded with `#N/A`; an error given as an argument is the
    result, an empty argument `#VALUE!`.
  - `TOCOL`, `TOROW`: the values by rows or by columns, blanks, errors or both ignored; none left is `#CALC!`.
  - `MUNIT`: the identity matrix of a size truncated, read from text or a logical; below 1 `#VALUE!`.
  - `MINVERSE`: the inverse, through the factorization `MDETERM` uses, each column of the identity solved forward and
    back with sums in increasing order and a division by each pivot, as Excel computes it bit for bit; a zero pivot is
    `#NUM!`, and a subnormal result 0 (capture `matrix_family`, 285 elements).
  - `FREQUENCY`: how many numbers fall in each bin, in the bins' order, and above the largest in a last row; a number
    counts in the smallest bin at or above it, equal bins in the first; text, logicals and blanks are left out of both;
    with no number among the bins there is one bin, 0; an empty argument is `#VALUE!`.
  - `MODE.MULT`: every number that occurs most often, at least twice, in the order each first appears; logicals and text
    are left out, given directly too; none repeating is `#N/A`.
  - `TEXTSPLIT`: a text split into columns and rows at one or several delimiters, padded with `#N/A` or `pad_with`,
    empty pieces kept or dropped, matched with or without regard to case; empty text or an empty delimiter is `#VALUE!`;
    a range or array of texts is taken element by element.
    The text is read once from the start: where delimiters overlap the first to start is taken, and of several
    starting at one place a column delimiter before a row delimiter, whatever their lengths, and the longest of one
    kind (`TEXTSPLIT("a;;b",";",";;")` is one row, `a`, empty, `b`).
  - `RANDARRAY`: random numbers of a shape, between bounds, whole with `whole_number`, an array of counts or bounds
    taken element by element; new at every calculation, as Excel's are (its values are checked through their number,
    bounds and kind).
  - `WRAPROWS`, `WRAPCOLS`: a row or column of values in rows or columns of a count no wider than the values, the last
    padded (`#N/A` by default, empty included); a count below 1 is `#NUM!`, a two-dimensional array `#VALUE!`.

  Each takes its array argument whole in every formula; an error given as the array is the result, and an empty one
  `#VALUE!` (`EXPAND` takes it as a blank). A count given an array is taken element by element, each element giving
  the first value of its result (`#VALUE!` for a reference from `TAKE` or `DROP`); a range there is too in a dynamic
  array, where a plain formula intersects it.

  A formula set through `set_formula` whose text may give an array is written as a dynamic array whatever the size of
  its result, as Excel writes it typed, so that it spills when its data grows (capture `marker_typed`, 75 typed
  formulas): an array constant (`={1}`) or a call of one of them, or of `TRANSPOSE`, `LINEST` or `INDIRECT`, carried
  through operators, arguments taken element by element (`-SEQUENCE(1)`, `IF(TRUE,SEQUENCE(1))`, `ABS(SEQUENCE(1))`,
  `OFFSET(Z1,{0},0)`), `N`'s argument, `INDEX`'s array and defined names, but not through a function that takes an
  array (`SUM(SEQUENCE(3))`) nor an explicit `@`; `INDEX` with an index that is not a number other than 0, and `OFFSET`
  with a height or width that is not a number.
  A file holding `SORT` or `FILTER` without its `_xlfn._xlws.` prefix, or one of these functions with an argument count
  it does not take, does not open in Excel: such a formula keeps the value it holds and is reported in `Report.refused`.
  `UNIQUE`, `SORTBY` and `SEQUENCE` without their prefix are names Excel does not know, `#NAME?`. A `#CALC!` Excel
  caches (as `#VALUE!` with a rich value) is read as `#CALC!`.
- `Report.from_kept`: a cell holding the cached value of a function or feature not implemented, and every
  cell computed from one. After a change upstream of the formula they move to `uncomputed` and keep the
  values they hold; a change that removes the reliance takes them out.
- `ReadLimits.max_array_cells` (default 10,485,760, ten full columns): no formula makes an array
  with more cells (a range an array formula fills, a reference taken as an array, the result of
  an operation between arrays); one that would keeps the value the file holds and is reported.
  `ReadLimits` now bounds computing as well as reading.
- Functions, each captured from Excel over edge inputs: trigonometry (`SIN`, `COS`, `TAN`, `COT`, `SEC`, `CSC`,
  `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `ACOT`), hyperbolics (`SINH`, `COSH`, `TANH`, `COTH`, `SECH`, `CSCH`, `ASINH`,
  `ACOSH`, `ATANH`, `ACOTH`), `DEGREES`, `RADIANS`, `SIGN`, `TRUNC`, `LOG`, `LOG10`, `SQRTPI`, `EXACT`, `TIME`;
  `CEILING`, `FLOOR` and their `.MATH` and `.PRECISE` forms, `ISO.CEILING`, `MROUND`, `EVEN`, `ODD`, `QUOTIENT`,
  `FACT`, `FACTDOUBLE`, `GCD`, `LCM`, `COMBIN`, `COMBINA`, `PERMUT`, `PERMUTATIONA`, `MULTINOMIAL`, `SERIESSUM`,
  `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `ROMAN`, `ARABIC`, `BASE`, `DECIMAL`.
- Distributions, captured from Excel (capture `distributions_family`, 3,848 rows): `BINOM.DIST`, `BINOMDIST`,
  `BINOM.DIST.RANGE`, `BINOM.INV`, `CRITBINOM`, `NEGBINOM.DIST`, `NEGBINOMDIST`, `POISSON.DIST`, `POISSON`,
  `HYPGEOM.DIST`, `HYPGEOMDIST`, `EXPON.DIST`, `EXPONDIST`, `WEIBULL.DIST`, `WEIBULL`, `LOGNORM.DIST`, `LOGNORMDIST`,
  `LOGNORM.INV`, `LOGINV`, `GAMMA`, `GAMMALN`, `GAMMALN.PRECISE`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`. Counts are
  truncated and every domain error is Excel's, its exceptions included: a hypergeometric count outside what can occur
  is 0, not `#NUM!`; `POISSON.DIST` of 0 is `EXP(-mean)` for a negative mean too; `WEIBULL.DIST`'s density at 0 is 0;
  `BINOM.INV` needs a probability and an alpha strictly between 0 and 1; `ERF` and `ERFC` take no logical. A subnormal
  result is 0. The discrete distributions are within 8 units in the last place of Excel and the others within 4, but
  for 170 listed rows, most of the binomial at 1,000 trials and more (`QUIRKS.md`). A sum of more than a million
  terms (`BINOM.DIST` over 1E15 trials) is not computed: the formula keeps the value it holds and is reported.
- The beta, gamma, chi-squared, F and t distributions, captured from Excel (capture `continuous_family`, 5,059 rows):
  `BETA.DIST`, `BETADIST`, `BETA.INV`, `BETAINV`, `GAMMA.DIST`, `GAMMADIST`, `GAMMA.INV`, `GAMMAINV`, `CHISQ.DIST`,
  `CHISQ.DIST.RT`, `CHIDIST`, `CHISQ.INV`, `CHISQ.INV.RT`, `CHIINV`, `F.DIST`, `F.DIST.RT`, `FDIST`, `F.INV`,
  `F.INV.RT`, `FINV`, `T.DIST`, `T.DIST.2T`, `T.DIST.RT`, `TDIST`, `T.INV`, `T.INV.2T`, `TINV`, and the tests
  `T.TEST`, `TTEST`, `F.TEST`, `FTEST`, `CHISQ.TEST`, `CHITEST`, `Z.TEST`, `ZTEST`, `CONFIDENCE.T`. Degrees of freedom
  are truncated, from 1 to 1E10; each tail is computed directly where it is the smaller, and each inverse is the
  root to the last bit, so that far tails and probabilities near 1 keep their digits; at large shapes, where logarithms of
  the gamma function cancel, Loader's saddle point, DiDonato and Morris's expansion and Temme's uniform expansion
  take over. As in Excel, `CHISQ.INV.RT` and `CHIINV` at one degree of freedom invert the left tail at 1 - p (a p
  of at most 2^-54, about 5.55E-17, is `#NUM!`), an empty sample in a test is `#VALUE!` (`#DIV/0!` in `Z.TEST`), and an empty bound of
  the beta functions takes its default. The tests take ranges or arrays alike; `CHISQ.TEST` and a paired `T.TEST` pair
  their two arguments element by element in row-major order, whatever their shapes, when they hold as many elements
  (`#N/A` otherwise), and a paired `T.TEST` drops a pair with anything but a number on either side. Each function is within a relative tolerance of Excel from 1E-15 to 1E-12;
  110 rows are listed, 97 of them where Excel's result is the further from the true value (`QUIRKS.md`). Every one
  does bounded work on any input: a series or continued fraction stops at 20,000 terms (`#NUM!` past them) and an
  inverse is a bisection over the doubles' bit patterns, at most 64 evaluations; at extreme shapes, where Excel itself
  may calculate without end (seven formulas at a shape of 1.7E308), each gives its limit or `#NUM!`, never an error
  of its own.
- The complex-number functions, captured from Excel (capture `complex_family`, 1,229 rows, every one matching on
  the reference platform; 67 listed for others, where the math library can round a part's 15th digit the other way):
  `COMPLEX`, `IMREAL`, `IMAGINARY`, `IMABS`, `IMARGUMENT`, `IMCONJUGATE`, `IMSUM`, `IMSUB`, `IMPRODUCT`, `IMDIV`,
  `IMPOWER`, `IMSQRT`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMSIN`, `IMCOS`, `IMTAN`, `IMCOT`, `IMSEC`, `IMCSC`,
  `IMSINH`, `IMCOSH`, `IMSECH`, `IMCSCH`. A complex number is text such as `3+4i` or `-2.5-1.5j`, each part written
  with 15 significant digits, in digits below 1E21 and while they take at most 19 places after the point, else in
  the exponent form (`1E+21`, `1.23456789012345E-06`); `i` alone for 1, a part of 0 left out, spaces at either end
  dropped. Text that is no complex number is `#NUM!`; `i` and `j` in one calculation are `#NUM!` in `IMSUB` and
  `IMDIV` and `#VALUE!` in `IMSUM` and `IMPRODUCT`. The last digits are Excel's: `IMSQRT` and `IMPOWER` in polar form,
  a whole power by squaring and a negative one as a reciprocal (`IMPOWER(2,-0.5)` is 0.707106781186547), `IMLOG2`
  dividing by the double just below ln 2 (`IMLOG2(1E20)` is 66.4385618977473), `IMTAN` and `IMCOT` in the
  double-angle form. As in Excel, the trigonometric functions are `#NUM!` from a real part of 2^27 and the hyperbolic
  ones and `IMEXP` from an imaginary part of 2^27, but `IMCOT` of a real number from 2^27 is 0. An empty required
  argument is `#N/A`, an empty suffix is `i`, a later empty argument of `IMSUM` and `IMPRODUCT` is skipped, and
  Excel refuses a file holding a one-argument function called with its argument empty (`IMABS()`), as `set_formula`
  does (capture `empty_args`, 39 rows more).
- `MDETERM`, captured from Excel (capture `matrix_family`): elimination with partial pivoting, the first largest pivot,
  the pivots multiplied in order and signed by the swaps, bit for bit over integer, singular, Hilbert and random matrices
  up to 8 by 8; a zero pivot gives 0, an overflow `#NUM!`. Every element must be a number (`#VALUE!` for a text, a blank
  or a logical, the error for an error) and the matrix square.
- Closed-form finance, captured from Excel: `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `SYD`, `DISC`, `INTRATE`, `RECEIVED`, `PRICEDISC`, `YIELDDISC`,
  `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `DOLLARDE`, `DOLLARFR`, `EFFECT`, `NOMINAL`, `FVSCHEDULE`, `ISPMT`, `MIRR`, `PDURATION`, `RRI`.
- Descriptive statistics and regression, captured from Excel: `AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, `SKEW.P`, `PEARSON`, `RSQ`, `SLOPE`,
  `INTERCEPT`, `STEYX`, `FORECAST`, `FORECAST.LINEAR`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`, `TRIMMEAN`, `STANDARDIZE`,
  `FISHER`, `FISHERINV`, `PHI`, `GAUSS`, `PERCENTRANK` and its `.INC` and `.EXC` forms, `PERCENTILE.EXC`, `QUARTILE.EXC`,
  `PROB`, `CONFIDENCE`, `CONFIDENCE.NORM`.
- Engineering, bases and bits, captured from Excel: the twelve conversions between `BIN`, `OCT`, `DEC` and `HEX`, `BITAND`, `BITOR`, `BITXOR`,
  `BITLSHIFT`, `BITRSHIFT`, `DELTA`, `GESTEP`.
- Database, captured from Excel: `DCOUNTA`, `DGET`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DVAR`, `DVARP`.
- Text, captured from Excel: `SEARCH`, `SUBSTITUTE`, `PROPER`, `CLEAN`, `T`, `UNICHAR`, `UNICODE`, `NUMBERVALUE`,
  `USDOLLAR`, `VALUETOTEXT`, `ARRAYTOTEXT`, `TEXTBEFORE`, `TEXTAFTER`, and `FINDB`, `SEARCHB`, `LEFTB`, `RIGHTB`, `MIDB`,
  `LENB` and `REPLACEB`, as they behave in an English language setting.
- Date and time, captured from Excel: `DAYS`, `ISOWEEKNUM`, `WEEKNUM`, `DATEDIF`.
- Logical and information, captured from Excel: `XOR`, `ISEVEN`, `ISODD`, `ISNONTEXT`, `ISREF`, `ISFORMULA`, `ERROR.TYPE`, `TYPE`,
  `SHEET`, `SHEETS`, `AREAS` and `HYPERLINK`, whose result is the text it shows and which follows nothing.
  `FORMULATEXT` does not show the `@` the formula bar adds to a legacy formula.
- Functions: `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `LARGE`, `SMALL`, `PERCENTILE`,
  `PERCENTILE.INC`, `QUARTILE`, `QUARTILE.INC`, `MODE`, `MODE.SNGL`, `IFS`, `SWITCH`, `TEXTJOIN`, `CONCAT`,
  `XLOOKUP` and `XMATCH` (not the regular-expression match mode). Each is captured from Excel.

### Changed
- A function written in a file without the prefix Excel gives it (`XLOOKUP` for `_xlfn.XLOOKUP`, as openpyxl writes what it
  is given) is still a name Excel does not know, `#NAME?`; the report now says so and names the prefix:
  `XLOOKUP written without Excel's file prefix (_xlfn.): Excel shows #NAME?`. The README shows how to write these names
  with openpyxl.
- The database functions read a range's non-blank cells, not every cell, so they are computed over whole columns and
  sheets (one `DSUM` over ten whole columns of a small sheet took 1.9 s in 0.1.0, and now takes under 0.01 s), and
  `DCOUNT` passes over an error in its field, as Excel does (capture `database_family`).
- A cell naming a function not implemented keeps its cached value in a full recalculation too, as it
  did when cached values were trusted, and what depends on it is computed from that, not `#NAME?`.
  `Workbook.value` gives the kept value. Without a cached value it is `#NAME?`, as before.
- Evaluation lists fewer precedents (about 10% faster on the benchmark, and a workbook of whole-
  column references from 445 s to about 12 s), and writing values alone skips the code that inserts
  cells (about 7% faster).
- `MATCH`, `VLOOKUP`, `HLOOKUP` and `LOOKUP` over a whole column read its non-blank cells, not its
  million rows: a workbook of 31,912 such lookups went from more than 15 minutes to under 3 s.
- Requires xlformat 0.3 or later (`xlformat>=0.3,<1`): `TEXT` is xlformat's `text_function`,
  which does what `TEXT` does differently from a cell's display (a blank is 0, the empty format,
  a format Excel refuses, a value the format cannot show).

### Fixed
- `LINEST` with several regressors gave `#REF!`, and a computed criterion in a database function (a formula under a
  name the database lacks) matched no record, each with nothing reported. Neither is implemented: both are now
  reported (`Report.unsupported`, `uncomputed`) and keep the value the file holds.
- `INDEX` with a row index of 0 and a column index over a single column or row was `#VALUE!`; it is that column
  (`INDEX(Z1:Z4,0,1)` is Z1:Z4). A column index of 0 given over a single row or column is a row index and a column
  index as over any range (`INDEX(Z1:AB1,1,0)` is the row, `INDEX(Z1:AB1,2,0)` `#REF!`), where a column index left out
  counts along the line (`INDEX(Z1:AB1,2)` is AA1) (capture `dyn_index`).
- A number literal Excel refuses in a file formula, from `1E+308` up (the largest Excel takes is 9.99999999999999E+307) or
  below `1E-309`, was computed; the formula keeps the value it holds and is in `Report.refused`, as Excel refuses the file
  (capture `number_literals`).
- `A1#` of a dynamic array that cannot spill was the `#SPILL!` error; it is the anchor cell alone, which holds `#SPILL!`,
  so `ROWS(A1#)` is 1 and `SUM(A1#)` is `#SPILL!`, as in Excel (capture `dyn_sequence`).
- `VDB` with a fractional start or end period gave values far from Excel's (`VDB(78714.55552,0,15,14.43,15,1.5)` was 1,509.41
  where Excel gives 2,648.73). Without the switch to straight line, each period's depreciation is taken in proportion to the
  part of it the span covers; with it, the value left at a fractional start comes from periods aligned to that start, the
  fractional part first, and a straight-line period is capped after its part is taken, a declining-balance one before
  (capture `finance`: a sweep of fractional periods and 137 rows around the switch on two assets). 75 rows differ in their
  last bits.
- `XIRR` gave `#VALUE!` for values holding a blank or a number as text; a blank is 0 and the text is read as its number, as
  in Excel, while other text and a logical stay `#VALUE!` (capture `iterative`: `IRR`, `NPV`, `MIRR` and `XIRR` over flows
  holding an error, a text or a blank, where `IRR` gives `#VALUE!` for any error and the others give the error).
- A cell whose formula refers to another workbook (`[1]Sheet1!A1`, `[1]!Rate`) or across sheets (`SUM(S1:S3!A1)`), which
  is not computed, held `#NAME?` in a full recalculation, so `Workbook.value` and what was computed from it, an iterated
  loop included, disagreed with the written file, which kept the file's values. It keeps the value the file holds, as a
  cell calling a function not implemented does, and is named in `Report.from_kept` with what is computed from it, no longer
  in `Report.uncomputed`; it is still reported under `external or 3-D reference`, and the written file keeps the values it
  kept before. Without a cached value it is `#NAME?` and uncomputed, as before.
- A name qualified by the workbook itself, `[0]!Rate` as Excel writes it, gave `#REF!`: it is the workbook-level name, from any
  sheet, even one with a sheet-level name of its own (capture `workbook_name`). One qualified by another workbook (`[1]!Rate`)
  is reported as an external reference and keeps the value the file holds, as a cell reference into another workbook did.
- `INDEX` takes its fourth argument, the area: with one area, 1 is that area, a number below 1 is `#VALUE!` and above it
  `#REF!`, as in Excel (capture `functions`). It was `#VALUE!` for every area.
- A database function given an error, such as a name defined as `#REF!`, gave `#VALUE!`; the error is the result, the first
  in argument order (database, field, criteria), as in Excel. A file whose formula gives one a database that is no reference
  (`DSUM(1,"a",A1:A2)`, or an error other than `#REF!`) keeps the value it holds, in `Report.refused`, since Excel refuses
  the file (capture `database_family`).
- **An empty argument gave a wrong value with no report.** What it does differs by function, and each now follows a generated
  capture over every implemented function (`empty_args`): `AVERAGE(1,,3)` is 1.33, not 2, and `AND(1,,3)` is FALSE, not TRUE;
  `TEXTBEFORE("a-b","-",)` and `TEXTAFTER("abc","z",,,,"none")` take the default of what is empty, as `DEC2BIN(5,)`, `WEEKNUM(d,)` and
  `PERCENTRANK(r,3,)` do; `EDATE(d,)`, `MROUND(5,)` and other older functions are `#N/A`, and `LOOKUP(10,r,)` is `#VALUE!`. A file with a
  formula Excel refuses for an empty argument, such as `ABS()` or `SUMIF(r,10,)`, keeps the value it holds, in `Report.refused`.
- `RIGHT` with a count past the text's length but less than twice it gave too little: `RIGHT("ab",3)` was `b`.
- `LEN`, `MID`, `FIND` and `REPLACE` count a character outside the first plane, such as an emoji, as two, as Excel does:
  `LEN("😀")` is 2 and `FIND("b","a😀b")` is 4; `MID` can cut one in two (capture `text_family`).
- `DOLLAR` of a number whose result is longer than 255 characters is `#VALUE!`.
- **A workbook with iterative calculation on lost its values.** Version 0.1.0 treats a deliberate
  loop, such as circular interest in a financial model, as a circular reference with iteration
  off, so `recalculate` replaced the values Excel had iterated to, on the loop and on every cell
  depending on it, with the empty text, in the source file itself unless a destination was given,
  and reported them only as `circular`. This happened with `full=True`, in a file with
  `fullCalcOnLoad` set (openpyxl sets it on every file), and on the second `recalculate` of any such
  file, since the first writes that flag. These cells are now iterated as Excel iterates them (see
  Added) and are not in `Report.circular`. A file already run through 0.1.0 has lost those values
  until Excel calculates it again.
- **A circular reference with iteration off lost its values too.** A file holding values for cells on a
  loop, or depending on one, had them replaced with the empty text, whenever those cells were
  evaluated: the same paths as above. The rule came from captures of files with no cached values. Excel
  keeps the values, on opening, with `fullCalcOnLoad`, after a full recalculation and after saving
  (capture `circular_cached`), and so does the package; a cell with no value is still the empty text,
  and one set through the API is 0, as Excel gives. They stay in `Report.circular`.
- `VLOOKUP`, `HLOOKUP` and `LOOKUP` that find a blank cell give a blank, not 0, so `ISBLANK` of one is
  TRUE, `&""` is empty and `=""` is TRUE; as a cell's value it is still 0 (capture `criteria`).
- `SUBTOTAL` reads hidden rows: 101 to 111 leave out every hidden row, and 1 to 11 do too when a filter is
  hiding rows on a sheet the call names, as Excel does (capture `hidden_rows`). All counted every row.
- A name or reference to a block (several rows and columns) where one cell is meant takes the cell in the
  formula's row and column, on its sheet or another (capture `names`). It was `#VALUE!`.
- The intersection of two single cells works: `A1 A1` is that cell and `A1 A2` is `#NULL!`; both were
  `#VALUE!` (capture `criteria`).
- A blank cell as the criterion of `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`,
  `MAXIFS` and `MINIFS` matched blank cells; it is 0, as in Excel (capture `criteria`):
  `COUNTIF(T1:T9000,A5)` with `A5` empty gave the number of blank cells, and `0` is the answer when no
  cell holds 0.
- A defined name for one cell gave that cell's value, not a reference, so `OFFSET(Name,1,0)`,
  `SUMIF(OFFSET(Name,1,0),">0")`, `ROW(Name)`, `Name:OFFSET(...)` and ranges defined through such a name
  (`Sheet!$C$2:Name`) were `#VALUE!` or `#REF!` (capture `names`). An error in the range operator is now
  that error, not `#VALUE!`, and `ROW` of a name for a value is `#VALUE!`.
- A loop of references through a branch not taken was called circular, depending on the order of
  its cells: `A1` as `=IF(TRUE,1,B1)` and `B1` as `=A1+1` were both empty, and are 1 and 2 (capture
  `static_cycle`). Only references that evaluation follows make a cycle.
- An exact lookup over its own column was called circular wherever the formula sat. `MATCH`,
  `VLOOKUP`, `HLOOKUP`, `XLOOKUP` and `XMATCH` stop at their match, as Excel does (capture
  `index_range`): a formula below the match is computed, and only one that reads past itself is
  circular.
- A union of ranges as an argument gave a wrong value with no report: `SUM((A1,B2))` and
  `LARGE((A1,B2),1)` were `#VALUE!`, `COUNT((A1,B2))` was 0. Unions are not implemented; such a
  formula now keeps the value the file holds and is reported.
- Text no longer compares in code-point order, so `"a"<"~"` is FALSE, as in Excel, for the operators,
  criteria such as `">=a"`, and sorted lookups. Captured on Excel for Mac: every pair of 39 ASCII
  characters and 24 pairs of words. Marks and symbols sort before digits and letters, and accents
  after their base letter. Other characters are approximated (`QUIRKS.md`).
- A formula whose result is an array of references (`INDEX(r,{1;2})`, `OFFSET(A1,{0;1},0)`) stored a
  reference object as its value: the writer left the file's old cached value in place, and the
  report did not say. It now gives Excel's value. A function given such an array, which is not
  implemented, gave a wrong value (`N(OFFSET(...))` was 0); the formula now keeps the value the
  file holds and is reported.
- `ROW` and `COLUMN` of a range make an array in an array context, as Excel does: `ROW(A1:A4)` over
  a legacy array fills 1, 2, 3, 4, not 1, 1, 1, 1, and `SUMPRODUCT(ROW(A1:A4))` is 10, not 1.
- A range or array as a criterion (`COUNTIF`, `SUMIF`, `COUNTIFS`, `SUMIFS`) or a lookup key (`MATCH`,
  `VLOOKUP`, `HLOOKUP`) is taken element by element where Excel takes it so:
  `SUMPRODUCT(1/COUNTIF(A1:A4,A1:A4))` counts distinct values, and was `#DIV/0!`.
- `INDEX(A1:A4,0)` is the whole range, where it was `#VALUE!`, and `ROW` or `COLUMN` of a union is
  `#REF!`.
- `SUMIF` with a sum range that runs past the sheet's edge read cells from the next row; it stops
  at the edge, as Excel does.
- Writing an evaluated array formula took time quadratic in its rows and about 1.5 KB per cell;
  a declared array costs about 180 bytes a cell from reading to writing, and one over ten full
  columns takes about 18 s and 1.8 GiB.
- `TEXT` with a format holding `b` reads it as the Buddhist-era year (`TEXT(12.5,"abc")` is
  `a43c`), and `TEXT(x,"General")` with a three-digit exponent keeps what fits 11 characters
  (`2.2251E-308`), as Excel does (xlformat 0.3); both were known deviations.
- `TEXT(1,"""#""")`, a format of a quoted `#` alone, gave `#VALUE!` where Excel gives `#`
  (`-#` for -1): any result made only of `#` was taken for a run Excel shows when a value does
  not fit.
- A part that is not well-formed XML (a worksheet or the shared strings) raised expat's
  `ExpatError`; it is a `PackageError`, as documented. So do damaged data and false sizes
  (`zipfile.BadZipFile`, `zlib.error`), a row number off the sheet (`OverflowError`), a number,
  shared string or shared-formula index that is not one (`ValueError`, `IndexError`), and a
  sheet scope (`localSheetId`) that is not a number.
- Numbers in `<v>` are read as Excel reads them (capture `value_forms`): `nan`, `INF`,
  `Infinity`, `1_000`, `+5` and `0x10`, which Python reads and Excel refuses, make the file a
  `PackageError`; a number too large for a double is `#NULL!`; a negative shared-string index
  no longer reads the last string. Calculation settings that are not numbers are ignored.

### Security
- A small file could make the engine allocate without bound: a legacy array formula declared
  over the whole sheet, or `SUMPRODUCT(D!A:XFD*1)` over a sheet of two cells, made arrays of 17
  billion cells. See `max_array_cells` above.
- `SUMIFS` and `COUNTIFS` listed every position of their ranges before testing a criterion, so a
  small file with `SUMIFS(D!A:XFD,D!A:XFD,">0")` used memory without bound, past any limit. They
  now walk a range's non-blank cells; so does `SUMIF`, which refused such a range when blank cells matched. Over
  whole sheets, capture `array_functions` holds Excel's value for 9 formulas of these three and `COUNTIF`, and each
  matches.
- In 0.1.0 `MATCH`, `VLOOKUP`, `HLOOKUP` and `LOOKUP` read every row of a whole column, so a crafted workbook of such
  lookups could keep `recalculate` busy for hours; they now read its non-blank cells (figures under Changed).
- In 0.1.0 the database functions read every cell of their range, so a crafted workbook of them over whole columns
  could keep `recalculate` busy for hours, at about a second or more a formula; they now read its non-blank cells (figures under
  Changed).
- The refusal of a DTD in a package part could be bypassed: it matched the prolog's bytes, so a
  `<!DOCTYPE>` after a long comment (200 KB) or processing instruction (900 KB), in UTF-16, or in
  a worksheet part whose name does not end in `.xml` (`data.bin`) was accepted. expat now reads
  each part's prolog with a handler that refuses a doctype declaration and stops at the first
  element, for every part read, whatever its name. An entity expansion that got through was
  stopped only by expat's amplification limit, a ratio that depends on expat's version.
- A part stated in the archive as smaller than it is was read cut short, silently when its CRC
  matched the shorter part; bytes hidden past a part's stated size were ignored. A part must now
  inflate to exactly its stated size and CRC. Encrypted parts and compression methods other than
  stored and deflated (bzip2 and LZMA were read) are refused, as Excel refuses them.
- A cell reference to a column past XFD gave a key that collided with another cell; the cell is
  placed by position, as Excel places it. A row number off the sheet, which was stored at a negative row or raised
  `OverflowError`, and a row holding more cells than the sheet has columns are refused (`PackageError`).

## [0.1.0] - 2026-11-02

First release.

### Added
- **Recalculating a file.** `recalculate(source, destination=None, ...)` evaluates a
  workbook's formulas and writes the results into the file as its cached values, in place and
  atomically or to another file, so pandas, openpyxl's data-only mode and other readers see
  values. Every part of the file that does not change is copied byte for byte, compressed data
  included; a sheet is patched by a streaming scan that replaces only the `t` attribute and
  `<v>` element of formula cells whose value changed. By default the file asks Excel to
  recalculate when it opens it (`full_calc_on_load`). A million-row workbook written by
  openpyxl (3 million formulas) is read, computed and written in about 50 seconds on an Apple M4
  Pro.
- **Working with a workbook.** `load(source)` returns a `Workbook`: `value`, `values` and
  `formula` read cells, evaluating only what is stale; `set` stores a constant (always, unlike
  openpyxl's assignment) and `set_formula` a formula as the formula bar takes it
  (`FormulaError` for what Excel refuses); `recalculate` evaluates; `save` writes the changed
  cells and the current values surgically. Formulas set through the API are written as Excel
  writes them in a file (`_xlfn.` prefixes, the dynamic-array marker exactly when Excel sets
  it), constants as Excel reads them, new cells and rows in order; the calculation chain is
  removed when a formula is.
- **Report.** `Report` lists what a recalculation could not do: unsupported functions and
  features, uncomputed cells (which keep the values the file holds), circular references,
  formulas Excel refuses in a file, and internal errors, each contained to its cell. Cells are
  named as `CellRef(sheet, row, col)`, whose `str()` is the address as Excel writes it.
- **Excel's value model.** Error values (`ErrorValue`, not a `str`), blanks, coercion and
  comparison across kinds, number-to-text and text-to-number conversion with 15 significant
  digits, number literals read from their first 15 digits, the 1900 and 1904 date systems, and
  four quirks that can be turned off (`Quirks`: `leap_year_1900`, `fifteen_digit_precision`,
  `comparison_rounding`, `cancellation_snapping`).
- **Arrays and references.** Array constants and arithmetic with broadcasting, implicit
  intersection in legacy formulas, legacy array formulas, dynamic arrays that spill (`#SPILL!`
  when blocked, `A1#`, `@`); `OFFSET`, `INDIRECT`, `CELL` and `INDEX` returning a reference,
  with dependencies found during evaluation and volatility as in Excel; circular references
  left uncalculated, as Excel leaves them with iteration off; `IF`, `IFERROR`, `IFNA` and
  `CHOOSE` evaluate only the branch they take.
- **Functions.** 156 of Excel's functions, every one of the hundred most used in a census of
  real workbooks: math and rounding (`ROUNDUP` as Excel computes it), statistics, logic, text
  (`TEXT` through xlformat, `DOLLAR`, `FIXED`), dates and working days, lookup and reference,
  financial and bond functions, database functions. The documentation's coverage matrix lists
  every Excel function with its status and evidence.
- **Captures.** 12,731 formulas computed by Excel 16 for Mac and compared with the package's
  results: 12,608 match, bit for bit or within a tolerance declared per function (a few units in
  the last place; a relative 1E-9 for the iterative `IRR`, `XIRR`, `RATE` and `YIELD`), 121
  known deviations
  listed with their reasons in `QUIRKS.md`, and 2 not evaluated (they call `XLOOKUP`, not
  implemented yet).
- **Speed and memory.** Formulas of the same relative shape share one compiled function;
  dependents come from the templates' geometry, not per-cell edges, so changing one cell of a
  million-row model finds what it affects in under a millisecond; the cell store keeps rows and
  template ids in typed arrays. Three million formulas evaluate in about 14 seconds on an Apple
  M4 Pro.
- **Reading.** An xlsx reader in the standard library only: one streaming pass gives every
  cell's formula and cached value, with shared and array formulas, defined names, tables and
  calculation settings; archive limits (`ReadLimits`, `LimitError`) and a refusal of DTDs
  protect against hostile files; `PackageError` for a file that is not a readable package.

[Unreleased]: https://github.com/sheetmill/xlrecalc/compare/v0.8.0...HEAD
[0.8.0]: https://github.com/sheetmill/xlrecalc/compare/v0.7.0...v0.8.0
[0.7.0]: https://github.com/sheetmill/xlrecalc/compare/v0.6.0...v0.7.0
[0.6.0]: https://github.com/sheetmill/xlrecalc/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/sheetmill/xlrecalc/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/sheetmill/xlrecalc/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/sheetmill/xlrecalc/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/sheetmill/xlrecalc/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sheetmill/xlrecalc/releases/tag/v0.1.0
