# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Iterative calculation, for a workbook with iteration on (`<calcPr iterate="1">`), as Excel does it (captures
  `iterate_calc`, `calc_chain`, `iterate_open`): a run of up to `iterateCount` iterations from the values the cells hold, each
  computing the cells on and after circular references in place, in Excel's order (sheets by name, then row, then column),
  until the largest change of a number is below `iterateDelta`. A run happens where Excel runs one: an evaluating load
  (as an opening), each edit, each recalculation and each save, so `Workbook.save` changes the values in memory too; a
  trusting load runs nothing. `recalculate(path)` of a file it evaluates is a load and a save, two runs, which is what Excel
  saves after opening the file, whether or not the loop has cached values; a loop with none starts from the empty text, as
  in Excel (captures `iterate_uncached`, `iterate_opening`, `iterate_cached`). `ReadLimits.max_iteration_evaluations` (100,000,000 by default) bounds the cell evaluations of one run;
  past it, the cells keep their values and are reported.
- `Report.from_kept`: a cell holding the cached value of a function or feature not implemented, and every
  cell computed from one. After a change upstream of the formula they move to `uncomputed` and keep the
  values they hold; a change that removes the reliance takes them out.
- Functions, each captured from Excel over edge inputs: trigonometry (`SIN`, `COS`, `TAN`, `COT`, `SEC`, `CSC`,
  `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `ACOT`), hyperbolics (`SINH`, `COSH`, `TANH`, `COTH`, `SECH`, `CSCH`, `ASINH`,
  `ACOSH`, `ATANH`, `ACOTH`), `DEGREES`, `RADIANS`, `SIGN`, `TRUNC`, `LOG`, `LOG10`, `SQRTPI`, `EXACT`, `TIME`;
  `CEILING`, `FLOOR` and their `.MATH` and `.PRECISE` forms, `ISO.CEILING`, `MROUND`, `EVEN`, `ODD`, `QUOTIENT`,
  `FACT`, `FACTDOUBLE`, `GCD`, `LCM`, `COMBIN`, `COMBINA`, `PERMUT`, `PERMUTATIONA`, `MULTINOMIAL`, `SERIESSUM`,
  `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `ROMAN`, `ARABIC`, `BASE`, `DECIMAL`.
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
- `ReadLimits.max_array_cells` (default 10,485,760, ten full columns): no formula makes an array
  with more cells (a range an array formula fills, a reference taken as an array, the result of
  an operation between arrays); one that would keeps the value the file holds and is reported.
  `ReadLimits` now bounds computing as well as reading.
- `SUMIF`, `SUMIFS` and `COUNTIFS` visit only a range's non-blank cells, so they are computed over
  any range, whole sheets included (capture `array_functions`).

### Changed
- The database functions read a range's non-blank cells, not every cell, so they are computed over whole columns and
  sheets, and `DCOUNT` passes over an error in its field, as Excel does (capture `database_family`).
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
  now make no such list; nor does `SUMIF`, which refused such a range when blank cells matched.
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
  placed by position, as Excel places it.

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

[Unreleased]: https://github.com/sheetmill/xlrecalc/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sheetmill/xlrecalc/releases/tag/v0.1.0
