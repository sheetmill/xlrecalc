# Architecture

xlrecalc evaluates the formulas of an Excel workbook from Python with the results Excel
computes, and writes those results back into the file as the cached values Excel stores. This
document describes the design: what the engine is for, the evidence behind its shape, the value
model, how formulas become compiled templates, how dependencies and scheduling work, the
function library, the xlsx reader and surgical writer, the ways to use it with openpyxl, and how
Excel's results are captured for the tests. Measured numbers are in `bench/results/SUMMARY.md`.

## 1. Goals and non-goals

### 1.1 Goals

1. **Excel's results.** Excel's type coercion, error values, date serials, array semantics and
   precision quirks (§4.4), with Excel 16 for Mac as the reference.
2. **Recalculate and write back.** Fill in the cached values of a workbook written by openpyxl
   or another generator, so pandas, openpyxl's data-only mode and other readers see values.
   The writer changes only what it must (§9.2).
3. **Fast on large workbooks.** Targets, on the benchmark workbook in `bench/` on a laptop:
   - full recalculation of 1M simple formulas in under 15 s, excluding file parsing;
   - at least 100k simple formulas per second sustained;
   - a change with a small fan-out re-evaluated in under 1 ms plus the cost of the affected
     formulas;
   - a change with a large fan-out never slower than a full recalculation (§6.4);
   - no quadratic behavior on long chains (running totals), wide fan-outs (one cell feeding
     100k formulas) or large ranges.
4. **Low memory.** Under 200 bytes per formula cell and under 50 bytes per value cell in the
   default store, with a path to much less.
5. **Pure Python.** No office suite, no compiled extension. The one required dependency is
   xlformat (pure Python, no dependencies of its own), which formats values for `TEXT`. There
   are no optional extras; any added later (§13) are never needed for correctness.
6. **Complete function coverage** as the long-term goal (Excel has about 500 worksheet
   functions), with a coverage matrix generated from the function registry (§8.4).

### 1.2 Non-goals

- Rendering or formatting cells; the companion packages do that.
- VBA, macros, pivot tables, charts, data tables (`t="dataTable"`), what-if analysis, Power
  Query, cube functions (they keep their cached values).
- Reaching outside the workbook. `WEBSERVICE`, `STOCKHISTORY`, `TRANSLATE`, `IMAGE`, `RTD` and the `CUBE`
  functions need a provider: the engine never reaches out by itself, by default or from anything a
  workbook says, and they keep their cached values and are reported, as an absent function does. Whether
  a caller can supply one under that name is for the design of custom functions (§12).
- `CALL` and `REGISTER.ID` stay unsupported: a workbook must not be able to name native code for the
  package to run. This can be revisited if a legitimate use and a safe way appear, and someone's own
  extension, running on their own system, may implement them.
- Evaluating links to external workbooks: their cached values in the file are constants.
- Localized function names or separators: English names, `,` between arguments, `.` as the
  decimal point; locale-dependent parsing and formatting follow en-US.
- Structural editing (inserting rows or columns, moving ranges). The design keeps two things
  possible for a later editor: change tracking in the model (§9.2) and formulas that print back
  to their exact text (§5.2), so references can be rewritten.
- Concurrent mutation of one workbook from several threads.

## 2. Evidence that shaped the design

Measured with prototypes in `bench/` (`proto.py`, `proto_compact.py`) on a workbook of N rows
and 3N+3 formulas: a `=A*2` column, a running-total chain, a sliding-window `SUM` column and
three `SUM` cells over whole columns. Full tables in `bench/results/SUMMARY.md`.

| finding | number | consequence |
|---|---|---|
| Existing pure-Python engines on the same workbooks | an order of magnitude or more longer to evaluate and more memory; the measurements and scripts are in `bench/` | 10x to 50x is available |
| A plain Python evaluation loop is fast enough | about 400k trivial formulas per second | CPython is not the bottleneck; structure is |
| Per-cell Python objects are the memory problem | a dict of floats costs 148 B per cell; a per-cell dependency graph about 1.5 KB per formula | no per-cell objects, no per-cell edge lists |
| Formulas are massively repetitive | 16 distinct templates for 3M formulas | compile once per template, store a template id per cell |
| Template deduplication | load 7x faster and memory 9x lower than compiling per cell; 1M rows load in 5.4 s, evaluate in 8.4 s, 1.5 GB peak | the core idea |
| Unindexed reverse lookup | changing A1 on 1M rows re-evaluates in 9.9 s, slower than the full recalculation (8.4 s); computing the dirty set alone takes 7.4 s | the reverse probe must be indexed (§6.4) |
| Evaluation driven from SQL in phases | 107 s for 30k formulas, quadratic in chain depth | evaluation is Python-driven, in dependency order |
| SQLite as a value store | about 40% throughput cost; 20 B per cell against 148 | a possible optional backend, not the default |
| SQLite R-tree for range reverse lookup | 8.4 µs per lookup against 5.4 µs for a bucketed interval index in Python; 17x slower to build | a Python interval index suffices |
| Numbers stored as 15-digit text | lossy (`0.1+0.2` comes back as `0.3`) | values are stored natively typed |

## 3. Design principles

1. **Templates, not cells.** A formula is normalized into a position-independent template
   (relative references as offsets, absolute references kept absolute), parsed and compiled
   once. A formula cell is a value slot plus a template id.
2. **Dependencies are geometry.** A cell's precedents are derived from its template and its
   position, never stored per cell. Reverse lookups probe the few positions that could
   reference a changed cell, through an index keyed by reference offset (§6).
3. **Python drives evaluation**, iteratively, in dependency order, over the dirty set. No
   recursion: Excel chains are deeper than any recursion limit.
4. **Values are native and typed.** Floats stay floats. Errors are values, not exceptions.
5. **Storage behind an interface.** The evaluator talks to a small cell-store API (§7.5).
6. **Functions are data-described.** Each function declares its arguments and coercion; the
   dispatcher handles coercion, error propagation and array broadcasting (§8.1).
7. **Workbook text never becomes code.** Literals reach compiled templates only through a
   constants table (§5.4).
8. **Excel is the reference.** Correctness is what Excel computes, tested against captures (§10).

## 4. Values

### 4.1 Types

| Excel type | Python representation | notes |
|---|---|---|
| number | `float` | IEEE doubles. Dates and times are serials (1900 system, or 1904 from the workbook). The public API returns `float` for every number: an integral constant read from a file comes back as `2.0`, where openpyxl returns `2` |
| text | `str` | at most 32,767 characters, as in Excel |
| logical | `bool` | distinct from numbers in comparison and coercion |
| error | `ErrorValue` (§4.2) | not a `str` |
| blank | `BLANK` internally, `None` in the public API | distinct from `""` and `0`; coerces by context |
| array | a tuple of rows, each a tuple of values (§7.7) | array constants, array arithmetic, functions returning arrays |
| reference | `Ref` (sheet, box, or a union of boxes) | exists only during evaluation |

Type tests in the hot path use `type(v) is float` and friends; because `bool` is a subclass of
`int` and Excel's logicals are not numbers, no test uses `isinstance(v, int | float)`.

### 4.2 Error values

xlformat owns the shared error type, `xlformat.ExcelError`, a `StrEnum` whose members equal their
text. The engine does not use it: with a `str` subclass for errors, every text test in the hot
path would have to exclude errors first, and a forgotten exclusion would turn `#N/A` into text
silently. So xlrecalc has its own class:

- `ErrorValue`, a small immutable class with `__slots__ = ("code",)`. The 18 codes xlformat
  lists (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA`,
  `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`, `#BUSY!`, `#EXTERNAL!`,
  `#UNKNOWN!`, `#PYTHON!`, `#TIMEOUT!`) are interned singletons (`ErrorValue.NA`, ...), so the
  engine compares by identity. A code read from a file that is not in the list makes a new
  instance that propagates like any error.
- Interoperation: `ErrorValue` equals any `str` with the same text, and hashes
  as its text. Since `xlformat.ExcelError` members are `str`, `ErrorValue.NA ==
  xlformat.ExcelError.NA` is true in both directions and both are the same dictionary key,
  matching xlformat's own rule that an error equals its text.
- Conversion: `err.to_xlformat()` returns `xlformat.ExcelError(err.code)`. Values passed in (`wb.set(ref, value)`) that are `Enum`
  members, are `str` and are shaped like an error code become `ErrorValue`; a plain `str` stays
  text, as in Excel, where the text `"#N/A"` is not an error.
- Diagnostics (which exception, which function was unsupported) are not stored on the value;
  they go into the recalculation report, keyed by cell (§7.6).

The equality with `str` exists for interoperation only, and it makes `ErrorValue.NA == "#N/A"`
true and both the same dictionary key. In Excel the text `"#N/A"` and the error `#N/A` are
different values, so the engine never compares values of different kinds with Python's `==` and
never mixes kinds in one dictionary's keys: comparison operators, `MATCH`, `COUNTIF`, the
lookup functions and every lookup cache dispatch on the type first. Tests check that the text and
the error stay distinct through each of these. Callers who need the distinction check the type
(`isinstance(value, ErrorValue)`); the API documentation says so.

### 4.3 Coercion and comparison

Conversions and comparisons live in a few small modules, each tested against captures
(`tests/captures/*.tsv`, §10); everything stated here is captured from Excel 16.114.3 for Mac.

**Number to text** (`_numtext.py`; `&`, the text functions, text comparisons) is Excel's text
form of a number, not a cell format: up to 15 significant digits, rounded with ties down (as
xlformat's 15-digit rule), written out in full when that takes at most 20 characters without
the sign, and in scientific notation otherwise with at least two exponent digits:
`=1/3&""` is `0.333333333333333`, `=1234567890123456&""` is `1234567890123450`, `1E-18` and
`1E+19` are written out, `1E-19`, `1E+20` and `1.23456789012345E-05` are not. The largest double
gives `1.7976931348623E+308` (rounding up would leave the doubles). `TEXT(x,"General")` is
different: it is the `General` cell format (`0.333333333`, `1E-18`), which xlformat computes
(§8.3); suite `number_text` holds both.

**Text to number** (`VALUE`, `+0` and every implicit conversion; suite `text_number`, 516
rows) accepts surrounding spaces (not tabs or line breaks); a sign; `$` or `€` before the number
(not `£` or `¥`); `%` after or before it; parentheses for a negative (but not `(-5)`); grouping
commas when every group after the first has at least three digits (`1,2345` is 12345, `12,34`
is refused); an exponent; mixed fractions (`1 1/2`); Arabic-Indic digits (not fullwidth ones);
dates (`3/4/2025`, `3/4/25`, `2025-03-04`, `4-Mar-2025`, `Mar 4, 2025`, `4 March 2025`,
`Mar 2025`, and `3/4` or `1-2` in the current year, `1/0` as January 2000), times (`10:30`,
`10:30 PM`, `10:30:15.5`, hours past 24 and minutes past 59), and a date with a time in either
order. Two-digit years below 30 are 20xx. It keeps the first 15 significant digits and drops the
rest, refuses magnitudes from 1E308 up, and gives 0 below the smallest normal double. `VALUE`
and `+0` agree on every text and differ on logicals: `VALUE(TRUE)` is `#VALUE!`, `TRUE+0` is 1.
`DATEVALUE` and `TIMEVALUE` accept only dates and times, taking the whole and the fractional day.
Texts without a year depend on the date of evaluation; the capture header records the date.

**Comparison** (`_compare.py`; suite `compare`, 1,045 rows): numbers < text < logicals; text
without case, through `casefold` (`"straße"="STRASSE"` is TRUE); a blank is 0, `""` or FALSE by
what it meets, while an empty text is only `""`; an error operand gives that error. Values of
different kinds are never compared with Python's `==` (§4.2).

**Lookups and criteria** differ from the operators. An exact lookup (`MATCH` type 0,
`VLOOKUP` FALSE, `HLOOKUP`, `XLOOKUP`) compares numbers exactly, without `comparison_rounding`
(`MATCH(0.1+0.2,{0.3},0)` is `#N/A`), never matches a blank cell, and finds the text `"#N/A"`,
never the error. A `COUNTIF`-family criterion applies `comparison_rounding`, and a criterion
text spelled like an error code (`"#N/A"`, `"#n/a"`) matches that error and never the text;
`"="` matches blank cells only, `""` blank cells and empty texts, `"<>"` everything not blank,
errors included. Lookup caches key values by kind first (`lookup_key`), so the text and the
error can never share a key.

**Arithmetic** (`_arith.py`): any result smaller in magnitude than the smallest normal double
is 0; Excel has no subnormal numbers, and reads a subnormal constant from a file as 0. Sums are
taken left to right in plain double arithmetic, no compensated summation, in `SUM` over a range
or over arguments and in a chain of `+` (suite `sums`). `MOD(n,d)` is `n - d·⌊n/d⌋` computed
exactly and rounded once (`MOD(-8,1/3)` is 0.33333333333333287); a negative number to a
fractional power is its odd root when `1/y` is an odd integer (`(-8)^(1/3)` is
-1.9999999999999998, `-exp(y·ln|x|)`), `#NUM!` otherwise. `ROUND` and `ROUNDDOWN`
work on the 15-digit decimal value (`ROUND(2.675,2)` is 2.68); `ROUNDUP` is `ROUNDDOWN` plus
one unit at the last kept place, added in doubles, when digits are dropped
(`ROUNDUP(2.675,2)` is `2.67 + 0.01`, 2.6799999999999997). Capture `rounding` (ties,
near-ties, 15- to 17-digit values, magnitudes from 1E-300 to 1E21, both signs, digits -2 to
10) matches on every row.

**Criteria and dates in functions**: a `COUNTIF`-family criterion that reads as a number also
matches a text cell spelling it the same way (`COUNTIF(r,"3")` counts the text `"3"`). `EOMONTH`
and `EDATE` use the real calendar for month lengths: the end of February 1900 is serial 59, not
the fictitious 60. `TEXT` of a value its date format cannot show (a negative date) is `#VALUE!`
where a cell shows a run of `#`.

**Dates** (`_dates.py`; suites `dates` and `dates_1904`): a year from 0 to 1899 given to `DATE`
has 1900 added (`DATE(1899,12,31)` is December 31, 3799); a month or day out of range moves into
the next or previous one; the last day is December 31, 9999 in both systems; serial 0 is
"January 0, 1900" in the 1900 system (`DAY(0)` is 0).

### 4.4 Quirks

Settings object `xlrecalc.Quirks`, shaped like xlformat's: a frozen dataclass of booleans, all on
by default, with presets `Quirks.EXCEL` (all on; what `quirks=None` means) and `Quirks.STRICT`
(all off). Every quirk has a `QUIRKS.md` entry with the Excel evidence. Names shared with
xlformat mean the same thing in both packages, and when xlrecalc calls xlformat it passes the
shared fields through (`Quirks.to_xlformat`).

| quirk | Excel | strict | shared |
|---|---|---|---|
| `leap_year_1900` | serial 60 is February 29, 1900; `DATE(1900,2,29)` is 60; `WEEKDAY(1)` is 1 (Sunday); date parsing and `YEAR`, `MONTH`, `DAY` before March 1900 follow it | serial `d` is `1899-12-30 + d`; no February 29, 1900 | xlformat |
| `fifteen_digit_precision` | number to text uses at most 15 significant digits (`=0.1+0.2&""` is `0.3`); text to number keeps the first 15 significant digits, truncating (`VALUE("0.12345678901234567")` is 0.123456789012345) | number to text is the shortest text that round-trips the double (`0.30000000000000004`); text to number is correctly rounded | xlformat |
| `comparison_rounding` | the comparison operators and the `COUNTIF`-family criteria treat numbers that agree to 15 significant digits (ties rounded down) as equal (`=0.1+0.2=0.3` is TRUE, `<>` and `>` FALSE); exact lookups do not | IEEE comparison | — |
| `cancellation_snapping` | a formula whose last operation, outside any parentheses, is an addition or subtraction gives 0 when the result is less than 8 units in the last place of its left operand: `=0.1+0.2-0.3` and `=(0.1+0.2)-0.3` are 0; `=(0.1+0.2-0.3)`, `=0.1+0.2-0.3+0`, `=(0.1+0.2-0.3)*1` and `=IF(TRUE,0.1+0.2-0.3)` are 5.55E-17, and `=(0.1+0.2-0.3)=0` is FALSE. `SUM` does the same over its own additions (`=SUM(0.1,0.2,-0.3)` is 0) | IEEE result | — |

Matched unconditionally as Excel semantics: `ROUND` and friends rounding the decimal value
(`ROUND(2.675,2)` is 2.68), `=0^0` and `POWER(0,0)` being `#NUM!`, no subnormal numbers,
case-insensitive text comparison, blank as 0, `""` or FALSE by context, truncation of
non-integer arguments where Excel truncates.

The date system is a workbook property (`<workbookPr date1904>`), not a quirk.

## 5. Formulas

### 5.1 Pipeline

Source text → tokens → syntax tree → normalized template → compiled callable.

### 5.2 Parsing and printing

The tokenizer keeps every token's leading whitespace, so joining the tokens gives the source
back exactly, and every node of the syntax tree keeps its tokens, so printing the tree is
joining them: `parse(text).text() == text` for every formula (tests over a corpus written for
this project, over formulas as Excel saved them, and over every formula of the benchmark
workbooks). A reference rewrite changes only the reference token and leaves the rest of the
text as it was, which is what a later editor needs to move ranges.

The parser is a precedence climber with Excel's precedence, tightest first: the spill operator
`#`, range `:`, `@`, intersection (a space between two references), negation, `%`, `^`, `*` and
`/`, `+` and `-`, `&`, comparisons. Binary operators are left-associative (`=2^3^2` is 64);
negation binds tighter than `^` (`=-2^2` is 4); `,` inside parentheses that are not a call's is
union. Long operator chains are parsed and compiled without recursion; nesting (parentheses,
calls, prefix operators) is limited to 200 levels, well above the 64 levels of nested
functions Excel allows, and deeper text is a syntax error rather than a `RecursionError`.

The grammar covers numbers, strings, logicals, error literals and array constants; A1
references with `$`, ranges, whole columns and rows, sheet-qualified, quoted and 3-D references
(`Sheet1:Sheet3!A1`), external references (`[1]Sheet1!A1`, `[1]!Name`, constants at run time)
and `Sheet1!#REF!`; defined names, including names with `.`, `?` and `\`; structured
references (`Table1[[#This Row],[Price]]`, `[@Price]`); function names with their prefixes
(`_xlfn.`, `_xlfn._xlws.`, `_xludf.`, and `_xlpm.` for `LET` and `LAMBDA` parameters); calls of
a call's result (`_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)`); union, intersection, range, `@` and
`#`. A name is a cell reference only when its column is at most `XFD` and its row at most
1,048,576 (`XFE1` and `A1048577` are names); a name directly followed by `(` is a function
(`LOG10(` is a call, `LOG10` a cell).

**The file form is Excel's own.** Formulas in a file have no leading `=` and carry the prefixes
above. Capture `formula_forms` (`capture/formula_forms.py`, `tests/captures/formula_forms.tsv`)
wrote the corpus (external references left out) into a workbook and had Excel save it. Excel
opened 241 of 249 formulas and wrote 208 back unchanged, among them 64 levels of nested
functions, 255 arguments and an 8,000-character formula. It rewrites number literals in its own
form, keeping the first 15 significant digits and dropping the rest (`.5` as `0.5`, `1E3` as
`1000`, `1.5E-10` as `0.00000000015`, `0.123456789012345678` as `0.123456789012345`, a 30-digit
number as `1.23456789012345E+29`), and computes them from those digits too (capture
`literals`: `=0.30000000000000004` is 0.3 in any context, so the compiled formula holds the
truncated value under `fifteen_digit_precision`), drops a leading `+`, whitespace before and after the
formula and before a comma, turns CR LF into LF, upper-cases cell references, function names
and logicals, drops quotes a sheet name does not need (`'Résumé'!A1` as `Résumé!A1`), and
drops `@` and `_xlfn.SINGLE` in a formula that is not a dynamic array (implicit intersection is
already the rule there). It refuses a file holding the spill operator (its file form is
`_xlfn.ANCHORARRAY(A1)`), the error literals `#SPILL!` and `#CALC!`, a tab, `SUM( )`, and
`OFFSET` with a constant as its reference (`OFFSET(5,0,0)`, capture `dynrefs`).
**Strictness follows the input path, as in Excel.**
- *Reading a file*: a formula Excel refuses in a file is treated the way Excel's repair treats
  it (§9.5): the formula is dropped and the cell keeps its cached value as a constant. Each one
  is listed in the recalculation report. A lenient reading mode that evaluates such formulas
  anyway is an option, not the default.
- *Formulas set through the API* (`set_formula`): what Excel's formula bar takes, in its form
  (`=`, `@`, `A1#`, no `_xlfn.` prefixes). Capture `formula_bar` (`capture/formula_bar.py`)
  typed every corpus formula, in that form, into a cell through AppleScript and had Excel save
  the file. The formula bar refuses what a file may not hold, the error literals `#SPILL!`
  and `#CALC!`, a tab between tokens and `SUM( )`, so `set_formula` refuses them too
  (`FormulaError`); it keeps leading whitespace, a tab inside a string literal, and the rest.
  (AppleScript sets the legacy formula and at most 255 characters of it, and that path does not
  know `LAMBDA` calls; a file holding `_xlfn.LAMBDA(_xlpm.x,...)` opens and is kept.)
- *Writing*: always Excel's file form. The **file-form printer** (`_formula/_fileform.py`)
  turns the formula bar's form into the text Excel writes: function names in upper case with
  the prefix their file form needs (`_xlfn.IFNA`, `_xlfn._xlws.SORT`), `_xlpm.` on `LET` and
  `LAMBDA` parameters, `A1#` as `_xlfn.ANCHORARRAY(A1)`, `@x` as `_xlfn.SINGLE(x)` in a
  dynamic-array formula and dropped in a legacy one (where implicit intersection is already
  the rule, §7.7),
  `Table1[@Price]` as `Table1[[#This Row],[Price]]` (Excel refuses the short form in a file),
  number literals in Excel's form (the first 15 significant digits: `.5` as `0.5`, `1E3` as
  `1000`), a leading `+` and the whitespace in array constants and before commas dropped,
  sheet names unquoted where Excel does not need the quotes, trailing whitespace dropped and
  line breaks as LF. The prefixes come from the captures, through `tools/prefix_table.py`:
  `formula_bar` (calls of about 500 function names with 0 to 6 arguments, and with range
  arguments where the formula bar wants them) and `prefixes` (`capture/prefix_probe.py`: a
  call written with and without `_xlfn.`, the form Excel computes being the file form), so 146
  functions; `LAMBDA`'s comes from `formula_forms`. On the corpus the printer gives Excel's own
  file form for 229 of the 237 formulas the formula bar takes, the others differing only where
  `tests/test_fileform.py` lists why; Excel opens a file holding all 245 the API takes
  (`tests/test_excel_live.py`). `display_form`, the inverse, shows a file's formula in the
  formula bar's form (`Workbook.formula`). Formulas that came from the file and were not
  changed are written back byte for byte.

### 5.3 Templates

Normalization rewrites each reference relative to the formula's own cell: a relative row or
column becomes an offset, an absolute one stays absolute, sheet-qualified references keep the
sheet. The canonical template text (R1C1-like) is the deduplication key. xlsx shared formulas
(`<f t="shared" si="0" ref="B2:B100">`) map onto one template directly; array formulas
(`t="array"`) carry their range.

**Load path.** Parsing costs about 9 µs per formula (5.6 µs of it tokenizing) and a template
key 3.6 µs more, measured on the 300,003 formulas of the 100k-row benchmark. A file Excel saved
stores a filled-down range as one shared formula, so it is parsed once. openpyxl writes every
formula as its own text, so the loader derives a template key straight from the raw text
(`_formula/_fastkey.py`, about 1 µs per formula): every reference the tokenizer reads as one is
rewritten relative to the cell (`R[-1]C[0]`, `R$5C$2`), and strings, quoted sheet names and
items, and bracketed parts are left alone. It parses a formula only when its raw key is new.
Because the raw key can be inverted at the cell, two formulas with the same key are the same
formula moved and share the parsed template; a reference the tokenizer reads as a name (a
column past `XFD`, a row past 1,048,576) stays text, so it cannot collide with a reference.
`tests/test_fastkey.py` moves every corpus formula to other cells and checks that equal raw keys
always mean equal parsed keys. Whether Excel refuses a formula in a file (§9.5) depends only on
its text outside the references, so it is decided once per raw key too. Loading
openpyxl's 1M-row benchmark file (3M formulas, 4M cells) takes 23.6 s with the raw key and
58.8 s parsing every formula; from the prototypes' pickled table, which skips the xlsx reader,
8.5 s.

A template records its canonical text, the compiled callable, its constants table, its reference
list (offsets, anchors, sheets, cell or range), whether it is volatile (`NOW`, `TODAY`, `RAND`,
`RANDBETWEEN`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`), whether it has dynamic references, whether
it can spill, and which functions it uses (for the unsupported-function report). Templates are
interned per workbook.

### 5.4 Code generation

A template compiles to a Python function `fn(ctx, r, c)` through generated source and
`compile()`: simple, and fast enough (about 400k evaluations per second measured). Nothing else
depends on it; a different backend can replace it per template. The function is straight-line
code, one statement per operation assigning a temporary (`t3 = ops.add(t1, t2)`), so the source
nests at most two levels however deep or long the formula: CPython refuses source nested more
than 200 parentheses deep, and `=A1+A2+...+A300`, compiled as one nested expression, would
already exceed that. Lazy functions (`IF`, `IFERROR`, `IFNA`, `CHOOSE`) compile each later argument as a
function of its own in the same form (`b0(ctx, r, c)`), which `ops.lazy_if` and its siblings
call only for the branch taken (§6.3).

**Workbook text never enters generated source.** Strings, numbers, error literals, sheet names,
defined names and array constants from the formula go into the template's constants table `K`;
functions are looked up in the registry table `F` by index, never by the name in the file. The
generated source therefore contains only names chosen by the generator (`ctx`, `r`, `c`, `K`,
`F` and helper names), integer literals the generator formats from Python `int`s, operators and
punctuation. Example: `=A1*2&" units"` at B1 becomes

```python
def _formula(ctx, r, c):
    t0 = ctx.cell(None, r, (c - 1))
    t1 = ops.mul(t0, K[0])
    t2 = ops.concat(t1, K[1])
    return t2
```

with `K = (2.0, " units")`; generated code runs with no builtins.

The test: generated source for every formula of the corpus, including hostile ones (quotes,
`__import__`, `exec`, newlines, `\N{...}` escapes, sheet names such as `'); import os; ('`), is
run through Python's `tokenize`; it must contain no string tokens, only names from the
generator's allowlist and its temporaries, only plain integers, and no nesting deeper than two
levels; string literals and sheet names must reach the compiled function unchanged.

## 6. Dependencies

No per-cell dependency lists.

### 6.1 The reverse-dependency index

`_deps.py`. A template's reference covers, for a formula at `(row, col)`, a box whose corners are
offsets from `(row, col)` or absolute. For a changed cell `(r, c)`, the formula positions whose
box contains it form a rectangle, found per axis:

- both ends relative (offsets `d1 <= d2`): positions `[r - d2, r - d1]`;
- both ends absolute (`a <= b`): every position when `a <= r <= b`, none otherwise;
- one end absolute `a`, one relative `d` (`$A$1:A5`, a range that grows down a column):
  `[r - d, MAX]` when `r > a`, `[0, r - d]` when `r < a`, every position when `r = a`.

A changed run `[r1, r2]` in one column works the same way with interval ends. The dependents are
the template's formula cells inside the rectangle; each template's cells are kept as row runs per
column, so a rectangle yields runs, found by bisection, not cells. Single-cell references with
both axes relative, the common case, are grouped by offset, so a change is probed once per
distinct offset in the workbook, not once per template. References to another sheet apply on
the sheets that use the template; a defined name contributes its own references. The index is
built from the templates and the cells' runs (0.4 s for the 3M formulas of the 1M-row
benchmark) and rebuilt when formulas change, not when values do.

### 6.2 Ranges

Ranges need no index of their own: the rectangle rule of §6.1 covers relative, absolute, mixed,
whole-column and whole-row ranges alike, and the template runs replace the per-instance interval
index the prototypes measured. `tests/test_deps.py` builds random workbooks mixing every shape
(relative cells and boxes, running totals, absolute cells, growing ranges, whole columns and
rows, other sheets, defined names) and checks the index's dependents against brute force.

### 6.3 Dynamic references

`OFFSET`, `INDIRECT` and `INDEX` make references while a formula is evaluated, so their
precedents are not in the template's reference list. Capture `dynrefs`
(`capture/dynref_probe.py`, `tests/test_dynrefs.py`) settles their behavior: `OFFSET`
truncates its offsets, takes an omitted size from the reference, extends up or left for a
negative size, and is `#REF!` for a size of 0 or off the sheet (Excel refuses a file holding
`OFFSET(5,0,0)`, a constant where the reference goes); `INDIRECT` reads A1 text in any case, with
a sheet prefix (quoted where the sheet needs it), a name, or R1C1 text with relative parts taken
from the formula's cell, and is `#REF!` for a leading space, another workbook or anything else;
`CELL("address")` of another sheet names the book (`[book.xlsx]Sheet2!$A$1`), and
`CELL("filename")` is the file's directory, `/[`, its name, `]` and the sheet. Without a
reference `CELL` reads the active cell of the sheet the workbook opens on (`workbookView
activeTab`, that sheet's `selection activeCell`), from every sheet, as Excel does when it
recalculates a file it opens (capture `cell_noref`, `capture/cell_probe.py`). `INDEX` into a
range returns a reference, so `Z1:INDEX(Z1:Z5,3)` is a range.

**Dependencies found during evaluation.** The scheduler orders by the static precedents
(§7.1); a formula that then reads a formula cell not evaluated yet stops (`_NotEvaluatedError`):
when that cell is on the current path, it is a circular reference (§7.2); otherwise it is
evaluated first and the formula again. Reads are checked only where the static order cannot be
trusted: through references `OFFSET` and `INDIRECT` make (`Engine.dynamic_ref`), and, while a
formula on a static cycle is evaluated, every read. So `INDIRECT("C8")+1` waits for C8,
`INDIRECT("C1")` in C1 is circular, and `IF(FALSE,C5,5)` in C5, a static cycle that evaluation
does not follow, is 5, as in Excel. `IF`, `IFERROR`, `IFNA` and `CHOOSE` evaluate an argument
only when they take it (§5.4).

**Volatility.** `OFFSET`, `INDIRECT` and `CELL` are volatile, as in Excel, with `NOW`, `TODAY`,
`RAND` and `RANDBETWEEN`: after any change, and at every `recalculate()`, they are evaluated
again with their dependents, so a reference that moves follows. `INDEX` is not volatile; its
references stay within its static range.

### 6.4 Dirty sets as runs, and the acceptance test

A change to a cell high up a long chain dirties a million cells, and the prototype spent 7.4 s
just finding them, cell by cell. A dirty set travels as row runs per `(sheet, column)`, and a
template that refers to its own column above (`C2 = C1 + B2`, a running total) makes every later
cell of its run depend on any earlier one, so a dirty cell of it dirties the rest of its run in
one step. Changing A1 in the 1M-row benchmark dirties 1,000,005 cells as 4 runs.

**Acceptance test** (`bench/engine_bench.py`, 1M rows, 3,000,003 formulas, one process; full
table in `bench/results/SUMMARY.md`): changing A1 must re-evaluate faster than a full
recalculation, and computing what it dirties must take under 1 s.

| | full recalculation | change A1, re-evaluated | its dirty set | change last A |
|---|---|---|---|---|
| prototype, unindexed | 8.2 s | 11.3 s | 8.8 s | 0.92 s |
| the engine | 13.6 s | 5.0 s | 0.14 ms | 0.95 s |

Both conditions hold (both rows on Python 3.11). The engine's full recalculation runs at about
60% of the prototype's speed (220,000 formulas a second): it carries Excel's coercion, error
values, references and snapping where the prototype added floats. Changing the last row re-evaluates its 15 dependents, among
them three `SUM`s over a million cells each, which is where the 0.95 s goes; incremental
aggregates are a later optimization.

## 7. Scheduling and evaluation

### 7.1 Ordering

Two levels, of which the second is built. **Template-level order** when possible: a graph over templates (`T2` depends on `T1`
if a reference of `T2` can land on a cell of `T1`); if it is acyclic, whole template columns
evaluate in order, and a self-referencing template (a running total) in row order. This enables
block and later vectorized evaluation. **Cell-level order** as the fallback: an iterative
depth-first post-order over the dirty set with an explicit stack, as in `bench/proto_compact.py`;
correct for any graph, and what `_engine.py` does.

### 7.2 Circular references

Detected during evaluation: a chain of cells, each stopped by reading the next, that closes on
itself (capture `static_cycle`). A loop of references through a branch not taken is not a cycle,
wherever its cells sit: Excel finds cycles through the references a formula evaluates, so `=IF(FALSE,A13,5)`
in A13 is 5, and so does the engine. A static cycle is a suspect, evaluated with every read checked,
and only a read of a cell on the current path is circular (§6.3).

With iterative calculation off, as Excel's default, Excel calculates no cell on a cycle and no cell
depending on one. It keeps the value the file holds, on opening, with `fullCalcOnLoad`, after a full
recalculation and after saving (capture `circular_cached`), and a cell with no value is the empty text
(`t="str"` with an empty `<v/>`; captures `circular` and `circular_alerts`). So does the engine, and
the report lists them as circular. After an edit, Excel keeps the values of a loop when a constant
it reads changes, and calculates its cells when the loop is broken; a formula entered on a loop or
reading one is 0, not the empty text (capture `circular_cached`).

When a cycle closes, its cells and every cell depending on them leave the stale set at once, so
evaluation replaces none of their values and `Workbook.value` and the written file agree. They share
one walk (`DependencyIndex.spread`), so that many cycles reading into one chain visit it once. A full
evaluation holds a file's cached values in the store for this. After a change, these cells are
evaluated again, since an edit may have broken the loop. A run that trusts cached values evaluates
nothing, so it changes nothing and reports nothing, whether or not the workbook holds a cycle.

**Iterative calculation** (`<calcPr iterate="1" iterateCount iterateDelta>`; captures `iterate_open`,
`iterate_calc`, `calc_chain`, checked cell by cell in `tests/test_iteration.py`). With iteration on,
the cells kept when a cycle closes (the cycle's and every cell depending on it) are iterated once the
evaluation pass ends, as Excel does:

- A **run** is up to `iterateCount` iterations (100 by default, at most 32,767) and at least one. It
  starts from the values the cells held before the recalculation. A cell with none holds the empty
  text, as an uncalculated circular cell does in Excel: arithmetic on it is `#VALUE!`, and
  `IFERROR` and `ISERROR` see no error in it, while `ISNUMBER` says FALSE (capture `iterate_opening`).
  So `A1+1` alone ends at `#VALUE!`, and a loop whose first cell is `IFERROR(A2,0)+1` ends its first
  iteration with every cell `#VALUE!`, unchanged, and stops there.
- Each iteration computes every one of those cells **in place**, in Excel's order: sheets by name in
  Excel's text order (`a` before `B`, `10` before `9`, `é` before `f`; not tab order or sheet id),
  then by row, then by column. A cell reads what the others hold at that moment, the ones before it
  already updated. Excel's `calcChain.xml` does not change this order (capture `calc_chain`), so the
  engine does not read it, and the writer may drop it. A precedent outside the cycle is computed
  before the run; a cell depending on the cycle is one of the iterated cells, and its change counts.
- A run **stops** after the first iteration in which no number changed by `iterateDelta` (0.001 by
  default) or more: the largest absolute change, over every iterated cell of the workbook, strictly
  below it. A cell's change is `|new - old|`, where an old value that is not a number (an error,
  text, the empty text) counts as 0, and a new value that is not a number is no change: so
  `IFERROR(A1,0)+3` from an error stops after one iteration at a delta of 10 and goes on at 0.001.
- **When a run happens.** Excel runs a loop once on opening (it marks a loop's formulas "calculate
  always", `<f ca="1">`), once for every edit, whether or not the loop reads the edited cell, once for
  every `calculate`, and once on saving (captures `iterate_calc`, `iterate_cached`). The engine runs
  where Excel does, with one departure: a trusting load runs nothing, so the file's values are the
  ones Excel saved, though Excel would run once more on opening it. An evaluating load (`full=True`,
  `trust_cached=False`, `fullCalcOnLoad`, or a loop without cached values) runs once, as an opening;
  each edit runs once (the file's `ca="1"` formulas become stale, so a loop a trusting load never
  met runs too); each recalculation runs once; and `Workbook.save` runs once more once a loop has
  run, so saving changes the values in memory as well as the file's, as Excel's saving changes
  what it shows. `recalculate(path)` is a load and a save: for a file it evaluates, two runs, which
  is Excel's saved copy after opening the file, what pandas and openpyxl readers then see (checked
  end to end in `tests/test_iteration.py`, from openpyxl files and from a file Excel saved).
  A limitation: after a trusting load, an edit finds a loop only through the file's `ca="1"`
  marking. A file with cached loop values but no marking, as another tool might write it, does not
  run its loop on an edit unless the edit reaches a cell the loop reads, or until a full
  recalculation (a test pins this, named for the limitation). openpyxl writes no cached values, so
  its files are always evaluated and are not affected.
- A run may make at most `ReadLimits.max_iteration_evaluations` cell evaluations (100,000,000 by
  default; the corpora's largest need is 8,256 cells at 1,000 iterations). Past it, the cells keep
  the values they held and are reported as uncomputed under a label naming the limit. A run that
  reaches, through `INDIRECT` or `OFFSET`, a formula its pass did not evaluate (a lazy read
  evaluates only what it needs) is undone the same way, under a label saying so.


What matches Excel exactly: each cell's arithmetic is the engine's, so a run gives Excel's doubles
wherever every function on the loop does; where a function in it is within a declared tolerance,
the result is within that tolerance, and for a loop that converges the stopping iteration can move
when a change lands within a few units in the last place of `iterateDelta`.

Excel recalculates a workbook with iteration on every time it opens it, whatever
`fullCalcOnLoad` says, starting from the cached values (capture `iterate_open`): a cell adding 1
to itself goes from 10 to 110 at the default 100 iterations, a copy saved after that holds 210, and
reopening that copy shows 310. A file the writer has passed through gives the same numbers as the
file it came from, so the writer's `fullCalcOnLoad` changes nothing there. A loop whose cells
have no cached value (`<v />`, as openpyxl writes) ends at `#VALUE!` in Excel, where a loop with
cached values converges.

### 7.3 Evaluation context

The engine (`_engine.py`) is itself the context the generated code calls (§5.4): `cell` gives a
cell's value, `cref` and `box` a `Ref` (a single cell passed as a function argument is a
reference, so `ROW(A1)` and `SUMIF(A1,...)` see one; in operators it is the value), `columns`
and `rows` whole-column and whole-row references, `name` a defined name (compiled once, with its
sheet scope first; a name for one cell is a reference, as a name for a range is, capture `names`),
and `unsupported` the `#NAME?` of a function or feature not implemented,
recorded in the report. It also holds the quirks, the date system, the date that texts without a
year take, and the cell being evaluated, for implicit intersection, `ROW()` and `COLUMN()`. The
operators are one `Ops` object per engine (`_ops.py`). A `Ref` reads through the store without
evaluating anything: the scheduler has already computed what a formula reads. Its non-blank
cells come row-major (the order sums are taken in), and only those are visited, so a
whole-column reference costs what the column holds, not a million positions.

### 7.4 When cached values are trusted

A workbook loaded from a file starts with the cached values the file holds. The rule:

1. A formula cell's cached value is **trusted** when the cell has one (a `<v>` with content, or
   `t="str"` with an empty `<v/>`, which is the empty string) and the workbook does not ask for a
   full calculation on load (`<calcPr fullCalcOnLoad="1">`).
2. Otherwise the cell is **stale**: no cached value (what openpyxl writes, `<v />`), or the flag
   is set (openpyxl sets it on every file it saves), or the template is volatile.
3. A cell that depends on a stale cell is stale.
4. A change made through the API makes the changed cell's dependents stale.

`_load.py` applies the rule when it reads a file: shared formulas' dependent cells take
their master's template without parsing, array formulas become anchors whose ranges' cells are
their members (§7.7), data tables keep their cached values (reported), and a formula Excel
refuses in a file becomes a constant (§9.5). Reading a
value evaluates only what is stale among its precedents. `recalculate(..., full=True)`
and `load(..., trust_cached=False)` treat every formula as stale. The rule matches Excel's own
behavior in the probe of §9.4: Excel shows a file's cached values without recalculating unless
`fullCalcOnLoad` is set or the file was last calculated by an older calculation engine
(`calcId`). xlrecalc ignores `calcId`: values from an older Excel are still Excel's values. A
file xlrecalc wrote with `fullCalcOnLoad` set (the default, §9.2) is therefore loaded with every
formula stale: reading it again costs a full recalculation, never a wrong value;
`full_calc_on_load=False` writes a file whose values a later load trusts.

### 7.5 Cell store

`_store.py`, per sheet: values in a dict keyed `row * 16384 + col`; template ids per column in a
32-bit typed array over the span of the column's formula rows (a dict of row to id for a column
whose formulas are sparse); and per column the sorted rows holding anything and the sorted rows
holding a formula, as 32-bit typed arrays, so a box yields its non-blank cells and its formula
cells by bisection. A cell of the workbook is one integer,
`sheet << 34 | key`, in the scheduler's sets. Array-backed column blocks (typed arrays for
numbers and template ids, side dicts for text and errors) are planned for a later minor release
(§13), and an SQLite backend for workbooks larger than memory after that.

### 7.6 Errors and the report

Errors propagate through operators and most functions; `IFERROR`, `ISERROR`, `IFNA`
and the `IS` functions inspect them. An internal exception in a function is caught per cell,
becomes `#VALUE!`, and is recorded in the report with the exception, so one bug never aborts a
recalculation. An unsupported function is listed in the report. A cell using one keeps its cached value when the file
holds one, since Excel computed it, in a full load too, and what depends on it is computed from that.
`Report.from_kept` names them all, so "which values rest on something not computed here?" has one
answer. A change upstream of the formula puts its kept value out of date: it and what depends on it
are not evaluated, keep the values they hold, and move to `uncomputed`. A function not implemented whose
formula reads through `OFFSET` or `INDIRECT` is never evaluated, so what it reads is not known, and any
change counts as upstream of it. A dependent that reads through them is evaluated, and its reads are recorded.
A change that removes the reliance (the formula replaced) takes the cells out of both. A read through
`OFFSET` or `INDIRECT` of such a cell has no static edge, so each evaluation records its boxes, in a
workbook holding such a formula only. A run that trusts cached values and evaluates nothing names
nothing; one that evaluates a cell names it and the formula it rests on.

A formula set through `set_formula` is taken as the formula bar takes it: without its `_xlfn.` prefix it computes and is written with
it. The same text read from a file is what Excel reads there, a name it does not know: `#NAME?`, reported under the name as the file
holds it.

Cells without a computable value (an unsupported function and no cached value, a caught exception),
and every cell depending on them, are **uncomputed**: the report lists them, and the writer
leaves their cached values as they were (§9.2). The report (`Report`) also lists circular
references, formulas Excel refuses in a file, and unsupported features by name and count.

### 7.7 Arrays, spill and implicit intersection

Capture `arrays` (`capture/array_probe.py`, `tests/test_arrays.py`) had Excel compute three
kinds of formula over a block of data: legacy formulas (as openpyxl writes them), each inside the
data's rows and outside them; legacy array formulas (`t="array"` over a declared range, entered
with Ctrl+Shift+Enter); and dynamic-array formulas (`cm` pointing at the dynamic-array cell
metadata, `XLDAPR`), as Excel 365 writes a typed formula. The package computes every one of the
311 captured cells as Excel does.

An array is a tuple of rows (`_arrays.py`). An operation between arrays takes the larger size in
each dimension; a dimension of 1 repeats, and a position past the end of the other operand is
`#N/A` (`{1,2,3}+{1;2}` is 2 by 3, `{1,2,3}+{1,2}` is `{2,4,#N/A}`). A function taking one value
is applied to each element of an array (`SUM(ABS({-1,-2}))` is 3). A cell shows an array's
first element.

**Legacy formulas.** An operator given a range, or a function argument that takes one value,
takes the range's *implicit intersection*: the cell in the formula's row (for a column range)
or column (for a row range), `#VALUE!` outside it. This holds inside the arguments of `SUM`,
`MAX`, `AND`, `IF`, `CONCATENATE` alike: `=SUM(Z1:Z5*2)` in row 3 is `Z3*2`, and `#VALUE!` in row
20. An argument Excel takes as an array is evaluated as one, ranges included: `SUMPRODUCT`'s
arguments, `INDEX`'s array, `ROWS` and `COLUMNS` (registry kind `array`, §8.1), so
`=SUMPRODUCT(--(Z1:Z5>2))` counts. A reference argument given a value instead is `#N/A`
(`MATCH(6,Z1:Z5*2,0)` is `#N/A` in row 3, and `#VALUE!`, the intersection's error, in row 20).
Array constants are arrays everywhere. The code generator compiles these rules: in an array
context (an `array` argument, or an array formula) it turns ranges into arrays (`ctx.arr`); and it
records whether a legacy formula intersects anywhere (`Template.intersects`).

**Functions of a reference.** In an array context `ROW` of a range is one column, a number for
each row, and `COLUMN` one row (capture `array_refs`): `ROW(Z1:AA3)` is `{1;2;3}`, and
`ROW(1:1048576)` a million numbers, not a whole sheet. Outside one each gives its first number:
`=SUM(ROW(Z1:Z4))` is 1 in a legacy formula and 10 in an array formula, and
`=SUMPRODUCT(ROW(Z1:Z4))` is 10 in both. The registry gives such a function another core for an
array context (`in_array`). `CELL` and `N` take a range's top-left cell in every form. A criterion
of the `SUMIF` family and `MATCH`'s key are one-value arguments, intersected in a legacy formula
and taken element by element in an array context, so `=SUMPRODUCT(1/COUNTIF(Z1:Z4,Z1:Z4))` counts
distinct values; `VLOOKUP`'s and `HLOOKUP`'s key are one value in a legacy formula, inside
`SUMPRODUCT` too, and element by element only in an array formula (registry kind `key`, §8.1).
`OFFSET`, `INDIRECT` and `INDEX` given an array make an array of references. As a whole formula
it gives Excel's value (the first reference's cell in a legacy formula; in an array formula
`#VALUE!` from `OFFSET` and `INDIRECT`, the cells from `INDEX`). Passed to another function it is
not implemented: the formula keeps its cached value and is reported.

After a full recalculation every formula cell either has a newly computed value or is reported:
`tests/test_invariant.py` plants a sentinel in the cached value of every captured formula.

**Legacy array formulas** are evaluated as arrays and fill their declared range: an array of one
row or column repeats across the range, a position past the end of the result is `#N/A`. Their
cells cannot be changed one by one.

**Dynamic arrays** are evaluated as arrays and spill: a result of `r` by `c` fills the `r` by `c`
range from the formula's cell, and the file records it as the formula's `ref`. A non-empty cell
in that range (a constant or a formula) blocks the spill: the formula is `#SPILL!` and fills
nothing, and clearing the cell lets it spill again. `A1#` (`_xlfn.ANCHORARRAY(A1)`) is the range
the dynamic array at A1 spills into; `@x` (`_xlfn.SINGLE(x)`) is the implicit intersection of
`x`. Excel caches `#SPILL!` in a file as `#VALUE!` with a rich value behind it (`vm`, error type
8); the reader decodes it, and the writer writes `#VALUE!` without the rich value, as other
readers see such a cell.

**In the engine** a formula that fills more than its cell is an *anchor* (`Engine.anchors`) with
its range; the cells it fills are *members*, found from the anchors' ranges (an index of row
spans per column, `Engine.member_of`) rather than stored per cell. A formula reading a member
orders after the anchor, and a member's dependents become stale with the anchor. A spill that
changes size makes the dependents of the cells it reached or left stale, and evaluation repeats
until nothing moves (at most 100 rounds). An anchor writes its result into the store element by
element, and the writer generates the cells of its range as it writes them (`Fill`), so a
legacy array over ten full columns (10,485,760 cells) is computed and written in about 1.8 GiB.

**Size.** Every array a formula makes is checked against `ReadLimits.max_array_cells` (default
10,485,760, ten full columns) before anything is allocated: a range an array formula fills, a
reference taken as an array, the result of an operation between arrays (a column times a row
makes every pair), and the ranges `SUMPRODUCT` walks position by position. A formula that would
pass it keeps its cached value and is reported (§7.6). Excel itself refuses a file declaring an
array over the whole sheet, opens one over ten full columns at a cost of about 3 GB of its own
memory, and computes `SUMPRODUCT` over whole columns without making the array (capture
`array_size`).

`SUMIF`, `SUMIFS` and `COUNTIFS` make no array and need no limit: they walk non-blank cells, as
`COUNTIF` does, and so do `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and `MINIFS`. A blank adds nothing to a sum. In a count, a criterion that rejects a blank
limits the matches to that range's non-blank cells; otherwise a position fails only where a
non-blank cell fails its criterion, and every other position counts. Their cost follows the cells
a sheet holds, and they give Excel's value over a whole sheet (capture `array_functions`).
`MATCH`, `VLOOKUP`, `HLOOKUP` and `LOOKUP` read a range's non-blank cells likewise, since neither an
exact match nor the bisection of cells of the key's type can land on a blank.
`XLOOKUP` and `XMATCH` do too; a blank key, which finds a blank cell, looks for the first gap.
An exact search reads only as far as its match, so it is circular only if it reaches its own cell
(capture `index_range`); an approximate one reads the whole line. `INDEX` reads nothing, so
`INDEX(C:C,1):INDEX(C:C,3)` in column C is circular only if the range it makes holds the formula.
`tests/test_function_sweep.py`, generated from the registry, runs every function over a whole
sheet of two cells under the memory watchdog, so one that walks a range by position fails there.

**Formulas set through the API** are dynamic arrays, as Excel 365 takes a typed formula: `=A1:A5*2`
spills. Capture `marker_typed` (ten formulas a person typed into Excel's formula bar and saved;
`capture/marker_probe.py --typed`) shows when Excel marks one in the file: when it spills, or
when its legacy reading would take an implicit intersection (`=SUM(Z1:Z3*2)`, a single value,
is marked); not otherwise (`=SUM(Z1:Z3)`, `=INDEX(Z1:Z3,2)`), and not for an explicit `@`, which
is the legacy meaning: `=@Z1:Z3` is saved as the plain formula `Z1:Z3`. The writer does the
same: the intersection test reads the dynamic-array text (where `@x` is `_xlfn.SINGLE(x)`) as a
legacy formula, and an unmarked formula is written in its legacy form, `@` dropped.
AppleScript cannot enter such a formula itself: a cell's `formula`, `formula local` and
`value` all enter legacy formulas (capture `marker_value`).

## 8. Functions

### 8.1 Registry

`@function("ROUND", "number", "number")` (`_functions/_registry.py`) registers a function under
its file-form name (`_XLFN.IFNA` for functions newer than Excel 2007) with the kinds of its
arguments: `number`, `text` and `logical` convert one value (§4.3), `value` takes one value as
it is, `raw` passes references and arrays untouched, `array` too and has its argument evaluated
as an array even in a legacy formula, and `branch` (the results of `IF`, `IFERROR`, `CHOOSE`)
passes it untouched but element by element when the call is; the last kind repeats for
variadic functions. When an argument of a one-value kind is an array, the function is applied
to each element (§7.7). The wrapper checks the argument count, converts, returns the first error among
converted arguments unless the function inspects errors itself, and calls the core function,
which contains only its own logic. One module per Excel category. The first functions were chosen
from `tools/function_frequency.tsv` (§10), whose counts come from workbooks of about 2001: it says
nothing about how much functions newer than that are used.

Results are compared with Excel's bit for bit. A function that stays a few units in the last
place from Excel after a time-boxed attempt at bit-exactness declares a tolerance
(`ulps=4, ulps_reason=...`): the capture tests accept a row within it, count it apart from
the exact matches, and use the largest tolerance among the functions a formula calls, except
in a formula that also rounds, compares or formats (`ROUND`, a comparison, `&`, `TEXT`, a
lookup or criterion), where a few units can change a visible digit and the row must match
exactly or be a listed deviation. A tolerance is a test concept only; the engine computes one
value. `QUIRKS.md` lists each one as a known deviation. Today `PRICE` declares 8 units and
`PMT`, `IPMT`, `PPMT`, `NPV`, `VDB`, `DURATION`, `MDURATION`, the variances and standard deviations,
`CORREL`, `LINEST` and the normal-distribution functions 4, for Excel's `exp`, `log1p`,
`expm1` and `erfc` and its summations; trigonometry, hyperbolics, logarithms and `SERIESSUM` declare 4 for
Excel's math library, `FACT` and `FACTDOUBLE` 8 and `COMBIN`, `COMBINA` and `PERMUT` 4 for the order it
multiplies in. The iterative functions (`IRR`, `XIRR`, `RATE`,
`YIELD`) declare a relative tolerance instead (`rel=1e-9`), since Excel itself guarantees
their results only to its documented stopping tolerance; rows within it are counted apart
from exact and ulp matches, under the same guard.

### 8.2 Status

Implemented, each against captures (suite `functions`, 1,238 rows): `SUM`, `IF`, `NOW`, `AVERAGE`,
`VLOOKUP`, `ROUND`, `TODAY`, `SUBTOTAL`, `MIN`, `MAX`, `SUMIF`, `MONTH`, `YEAR`, `AND`,
`ISNUMBER`, `OR`, `CONCATENATE`, `MATCH`, `ABS`, `EOMONTH`, `WEEKDAY`, `INDEX`, `TEXT`,
`HLOOKUP`, `NA`, `COUNTA`, `DATE`, `COUNTIF`, `SUMPRODUCT`, `COUNT`, `LEFT`, `ISNA`, `VALUE`,
`SQRT`, `ISERROR`, `ROWS`, `DAY`, `CHOOSE`, `ISTEXT`, `MID`, `FIND`, `LN`, `DATEVALUE`, `INT`,
`EDATE`, `EXP`, `LEN`, `ROW`, `TRIM`,
then `IFERROR`, `SUMIFS`, `COUNTIFS` and `_xlfn.IFNA` from the newer functions, and `NOT`, `TRUE`,
`FALSE`, `ISBLANK`, `ISERR`, `ISLOGICAL`, `ROUNDUP`, `ROUNDDOWN`, `MOD`, `POWER`, `PI`,
`PRODUCT`, `RIGHT`, `UPPER`, `LOWER`, `COLUMN`, `COLUMNS`, `TIMEVALUE`, `COUNTBLANK`, and
`_xlfn.ANCHORARRAY` and `_xlfn.SINGLE` (§7.7), `OFFSET`, `INDIRECT` and `CELL` (§6.3), `N`
(a range's top-left cell, an array element by element) and `LOOKUP` (vector and array forms;
a result range extended past its end; capture `n_lookup`), and the financial functions `PV`,
`FV`, `PMT`, `IPMT`, `PPMT`, `NPER`, `NPV`, `XNPV` and `SLN` (capture `finance`): Excel computes
the growth factor `(1+rate)^nper` by repeated squaring for a whole number of periods and as
`exp(nper*ln(1+rate))` otherwise, takes any nonzero type as payments at the start, and gives
`PPMT` from the closed form `-rate*(pv+fv)*(1+rate)^(per-1)/((1+rate)^nper-1)` rather than as
`PMT` minus `IPMT`, so late periods at high rates keep their precision; and `IRR`, `XIRR` and
`RATE` (capture `iterative`) by Newton's method from the guess (0.1 by default) on `1+rate`,
halving rather than stepping to `1+rate <= 0`, until a step is below 1E-7 (20 iterations, 100
for `XIRR`), `#NUM!` otherwise; `RATE` iterates on the present-value form of the annuity
equation. Excel's iterates rest on its own NPV and are not reproduced bit for bit: these
functions and `YIELD` declare a relative tolerance of 1E-9 (§8.1), and the rows beyond it are
listed with their reasons;
`XIRR` reads text dates and dates in any order after the first, unlike `XNPV`; the day counts
`DAYS360` and `YEARFRAC` (bases 0 to 4; capture `bonds`), the coupon functions `COUPPCD`,
`COUPNCD`, `COUPNUM`, `COUPDAYBS`, `COUPDAYS` and `COUPDAYSNC`, and `PRICE`, `YIELD`,
`DURATION` and `MDURATION`: `DAYS360`'s US method and `YEARFRAC` basis 0 differ at the end of
February, actual/actual takes year lengths from the calendar (1900 has 365 days there),
coupon dates step back from maturity keeping its day (the month's last day when maturity is
one), basis-1 `COUPDAYS` is `EDATE` from the previous coupon date, `PRICE` and the
durations take the days to the next coupon as `COUPDAYS - COUPDAYBS` in every basis, and
`YIELD` with one period left uses Excel's closed form with A as `COUPDAYBS`, E as the actual
period in bases 1 to 3, and the days to redemption in the basis's own count;
`WORKDAY`, `NETWORKDAYS` and their `.INTL` forms (capture `workdays`; weekend codes 1 to 7 and
11 to 17 or a seven-character mask, holidays as a range or array, logical arguments
`#VALUE!`, a day count floored; Excel also computes `_xlfn.WORKDAY.INTL`, so both spellings
are registered, while the file form has no prefix), and `HOUR`, `MINUTE`, `SECOND` (rounded to
the nearest second);
the statistics `STDEV`, `STDEVP`, `VAR`, `VARP` and their `_xlfn.` forms `STDEV.S`, `STDEV.P`,
`VAR.S`, `VAR.P`, the A forms `AVERAGEA`, `MINA`, `MAXA`, `STDEVA`, `STDEVPA`, `VARA`,
`VARPA` (text in references 0 and logicals 1 or 0; in an array constant only `AVERAGEA` counts
them), `MEDIAN`, `SUMSQ`, `CORREL`, `LINEST` (one regressor, with its statistics), `RANK`,
`RANK.EQ`, `RANK.AVG`, and the normal distribution `NORMSDIST`, `NORMDIST`, `NORMSINV`,
`NORMINV` with `NORM.S.DIST`, `NORM.DIST`, `NORM.S.INV`, `NORM.INV` (capture `statistics`);
`VDB` (the asset left at the start depreciated period by period from there; capture
`finance`); and `TRANSPOSE` (registry kind `matrix`: a range in a legacy formula is
intersected first, even inside `INDEX` or `SUMPRODUCT`, as in Excel), `ADDRESS`, `REPT`,
`REPLACE`, `CHAR` and `CODE` (the Mac OS Roman set, as Excel for Mac numbers characters;
Excel for Windows uses Windows-1252; there is no setting for it yet),
`RAND` and `RANDBETWEEN` (volatile; both bounds rounded up), and `DSUM`, `DAVERAGE`,
`DCOUNT`, `DMAX`, `DMIN` (a text condition matches by prefix; computed criteria are not
supported) (capture `misc`);
`INDEX` with a 0 index gives a whole row or column, and into a range returns a reference; its area
number is 1, since a union is not computed (below 1 `#VALUE!`, above `#REF!`).
`SUBTOTAL` reads the rows the file marks hidden: codes 101 to 111 leave out every hidden row, and 1 to
11 do too when a filter is hiding rows on a sheet the call names (capture `hidden_rows`).
`CELL`'s kinds that depend on formatting (`format`, `width`, `prefix`, `protect`, `color`,
`parentheses`) are reported unsupported, since styles are not read. Not yet: structured
references, 3-D references.

Each registered function declares a status: `implemented`, or `partial` with a note on what
is missing (`CELL`, `FORMULATEXT`, `LINEST`, `XLOOKUP`, `XMATCH`, the database functions); a function Excel has that is
not registered is absent. Excel's function list (`capture/functions_list.py`, the names the
`formula_bar` capture typed) lets the matrix list the absent ones too.

Newer and criteria functions (capture `modern`, 2,230 rows, and the end of `array_refs`):
`AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` (walking non-blank cells, §7.7); `LARGE` (`k` rounded
up), `SMALL` (`k` rounded down), `PERCENTILE`, `QUARTILE`, `MODE` and their names `PERCENTILE.INC`,
`QUARTILE.INC`, `MODE.SNGL`; `IFS` and `SWITCH`, which evaluate only what they take; `TEXTJOIN` and
`CONCAT`; `XLOOKUP` and `XMATCH`, with every search mode and the match modes but regular expressions
(mode 3, reported). `XLOOKUP`'s result is a reference, so `XLOOKUP(...):A9` is a range. Their
next-smaller and next-larger order numbers, text, logicals and a blank key last; a blank key finds
a blank cell. `MODE.MULT` is not implemented.

### 8.3 The xlformat seam

xlformat is a required dependency, imported normally.

| function | how |
|---|---|
| `TEXT(value, format)` | `xlformat.text_function` with the shared quirks and the workbook's date system; the `ExcelError` it returns (a format Excel refuses, a value the format cannot show) becomes the `ErrorValue` |
| `DOLLAR`, `FIXED` | built in |
| `&`, text coercion of numbers | built in: Excel's text form of a number is not a cell format (§4.3) |
| `VALUE`, `DATEVALUE`, `TIMEVALUE`, implicit text to number | built in: parsing is not formatting, and xlformat has no parser (§4.3) |

`DOLLAR` and `FIXED` round the 15-significant-digit decimal a number shows, half away from
zero as `ROUND` does, and write it with grouping (unless `FIXED`'s third argument says not),
up to 127 decimals, digits past the fifteenth as zeros (`FIXED(1.5E20)` is
`150,000,000,000,000,000,000.00`); `DOLLAR` puts a negative in parentheses, and keeps them on
a negative that rounds to zero when its decimals are 0 or more (`DOLLAR(-0.004)` is
`($0.00)`), while `FIXED` drops the sign; an empty decimals argument means 0 in `DOLLAR` and
2 in `FIXED` (capture `textfmt`). `TEXT` differs from cell display: `TEXT(1234.5,"*-0")` is
`1235` (the fill is dropped), `TEXT(1234.5,"0_)")` is `1235 ` (the `_x` space is a real
space), and a format Excel refuses in a file (`0.00e+00`) gives `#VALUE!`; text that reads as
a number or a date is formatted as that number (`TEXT("12","0.00")` is `12.00`), an empty
format gives the empty text for a number (`-` for a negative one) and leaves text and logicals
as text, a logical format is `#VALUE!`, and a format of a quoted `#` alone is that text
(`TEXT(1,"""#""")` is `#`, `-#` for -1). The seam converts text that reads as a number
and refuses a logical format; xlformat's `text_function` does the rest as `TEXT` does it.
Capture `textfmt` (24 values in 39 formats and argument kinds) covers them, every row
matching, `b` as the Buddhist-era year included (`TEXT(12.5,"abc")` is `a43c`); and
`TEXT(x,"General")` matches Excel's `General` on every row of `number_text`, a three-digit
exponent included (`2.2251E-308`).

### 8.4 Coverage matrix

`tools/coverage.py` joins the registry, Excel's function list and the capture suites: per
function its category, status, note, the captured rows that call it (a row counts for every
function in its formula), how many match exactly, are within its declared tolerance or are
listed deviations, and the tolerance. It writes `docs/coverage.md`, a page of the
documentation; `tests/test_coverage.py` fails when the committed page is out of date. The
tool runs in the repository, since it needs the captures and the capture tests.

## 9. Workbook I/O

### 9.1 Reader

xlrecalc reads xlsx files with its own reader in the private subpackage `xlrecalc._xlsx`
(standard library only, and it imports nothing from the rest of xlrecalc; a test enforces both).
`Book` opens a package and reads:
- the package: relationships, and the workbook part found through the package's own
  relationships;
- the workbook: sheets with their names, states and kinds (worksheet, chartsheet, dialog or
  macro sheet; only worksheets have cells), defined names with their scope, `<workbookPr
  date1904>`, `<calcPr>` (calculation id, `fullCalcOnLoad`, mode, iteration settings), and the
  external-link parts;
- shared strings, once, with phonetic runs (`<rPh>`) left out, as Excel shows text;
- each worksheet in one streaming pass: every `<c>` with its position (from `r`, or following
  the previous cell or row when `r` is absent), its cached value by type (`n`, `s`, `str`,
  `inlineStr`, `b`, `e`, `d`, or none for an uncached formula), its formula with the shared and
  array metadata (`t`, `ref`, `si`, all attributes as written), and its `cm` and `vm` (the
  dynamic-array and value-metadata markers), with the workbook's metadata part read so that a
  `cm` marking a dynamic array and a `vm` holding a rich-value error (`#SPILL!`) are known
  (§7.7);
- the tables on a worksheet (name, display name, range, columns, header and totals rows), for
  structured references.

Element names are matched by their local part, so files that prefix the namespace and Strict
OOXML files read like the usual ones. Numbers stay the doubles in the file; nothing converts
dates, so serial 60 and fractional seconds survive. Text, cached text results and formula text
are decoded from OOXML's `_xHHHH_` escapes, as Excel decodes them (`capture/escape_probe.py`:
`a_x000D_b` is three characters with a carriage return, `_x005F_x0041_` is the text `_x0041_`,
an incomplete `a_x000Db` stays as written, and the formula `="_x0041_"` written into a file
reads as `="A"`). openpyxl is not used at run time.

The reader is built on expat callbacks rather than ElementTree's `iterparse` (1.8 times faster on
the same sheet) and keeps its state in closure variables. Measured (`bench/read_bench.py`,
reading every cell of every worksheet in a fresh process, Apple M4 Pro):

| files | xlrecalc, one pass (formulas and values) | openpyxl read-only, one pass | openpyxl normal |
|---|---|---|---|
| benchmark, 10k rows | 0.07 s, 25 MB | 0.12 s (values 0.10 s), 32 MB | 0.16 s, 56 MB |
| benchmark, 100k rows | 0.59 s, 25 MB | 0.78 s (values 0.71 s), 40 MB | 1.60 s, 279 MB |
| 300 real-world workbooks, 35 MB | 5.9 s, 28 MB | 10.5 s (values 10.5 s), 118 MB | 149 s, 1,796 MB |

openpyxl gives formulas or cached values in a pass, never both, so loading a workbook for the
engine costs it two passes: the reader is 2.5 to 3.6 times faster than that, and against a
single openpyxl pass 1.3 to 1.8 times faster. Its memory stays flat (25 to 28 MB here, the
interpreter included); openpyxl's read-only mode used 118 MB on the real-world sample.

**What the reader reads differently from openpyxl** (`bench/read_compare.py` on 300 real-world
workbooks, 5.2 million cells, 363,000 formulas; every number, logical and error agreed):
- a formula whose cached result is the empty text (`<c t="str"><f>…</f><v/></c>`, as Excel
  writes `=""`) is the empty text; openpyxl reads an empty cell (11,131 cells). The reader
  matches Excel;
- a `<v>` holding a formula's text result, and each `<t>` of a shared or inline string (each
  run of a rich text separately), loses its leading and trailing whitespace when it is not
  marked `xml:space="preserve"`, as Excel reads it: `<t> </t>` is the empty text, the runs
  `a ` and ` b` read as `ab`, and an escaped `_x000D_` survives, as Excel strips before
  decoding (capture `whitespace`, `capture/space_probe.py`; §9.4); openpyxl keeps it;
- `_xHHHH_` escapes are decoded; openpyxl keeps them as text (`a_x000D_b`). The reader matches
  Excel (above);
- an array formula is a formula with its range; openpyxl returns an `ArrayFormula` object;
- a shared formula's dependent cells hold no text in the file, and the reader reports them so,
  with the master's index; openpyxl translates the master's text for each;
- an integral number is a float; openpyxl returns `int` (840,845 cells).

**Every file is untrusted.** The reader reads parts and evaluates formulas; it never runs or
fetches anything a file names:
- *Archive limits*, protection against zip bombs: at most 100,000 parts, 16 GiB decompressed
  per part and 64 GiB in total by default, above anything Excel can practically open, so the
  defaults never refuse a workbook Excel opens; parts are streamed, so memory does not grow with
  them. The real protection is streaming and the DTD refusal below; the limits bound time.
  `load` and `recalculate` take `limits=ReadLimits(...)` (`max_parts`, `max_part_bytes`,
  `max_total_bytes`, and two that bound computing: `max_array_cells`, §7.7, and
  `max_iteration_evaluations`, §7.2); a service that accepts files from untrusted users should lower them to
  what it is willing to spend on one file. Sizes
  are counted while decompressing, and a part must inflate to exactly the size and CRC the
  archive states, or it is refused (`zipfile` stopped at the stated size, so bytes past it were
  dropped silently). A compression ratio is not a usable test: real sheet parts
  compress about 5:1, but a legitimately repetitive one reaches 400:1, close to deflate's
  limit. Duplicate part names, encrypted parts and compression methods other than stored and
  deflated are refused, as Excel refuses them (sheetpeek's capture [`package`](https://github.com/sheetmill/sheetpeek/blob/main/tests/captures/package.tsv)).
- *Content Excel refuses*: a number in `<v>` that Excel does not read (`nan`, `INF`, `1_000`,
  `+5`, `0x10`; a value too large for a double is `#NULL!`, as Excel reads it), a row off the
  grid, more cells in a row than columns, a shared string or shared-formula index that is not
  one, are refused (capture `value_forms`); a column reference off the grid places the cell by
  position, as Excel does. Calculation settings, styles and metadata indexes that lead nowhere
  are ignored.
- *Array sizes*: no formula makes an array past `ReadLimits.max_array_cells`, and the functions
  that walk a range over a whole sheet visit only its non-blank cells (§7.7), so a small file
  cannot ask for an array, or a list of positions, of 17 billion cells.
- *No DTDs*, protection against XML entity attacks: Open Packaging Conventions parts never
  declare a DTD, so a part that declares one is refused before it is parsed. The check is expat
  reading the part's prolog, with a handler that refuses a doctype declaration and stops at the
  first element, so no encoding (UTF-16 included), comment or processing instruction hides one,
  and it applies to every part read, whatever its name; a prolog longer than 1 MiB is refused.
  This makes entity expansion impossible whatever expat the Python build links. A part that is
  not well-formed is a `PackageError`. The bundled expat (2.7 and 2.9 in Python 3.11 to 3.14) already
  refuses the billion-laughs and quadratic-blowup documents ("limit on input amplification
  factor") and undefined external entities, but ignores an external parameter entity rather
  than refusing it, and distribution builds may link an older system expat. Tests feed each of
  these documents to the reader.
- *No external content.* External links, DDE formulas (`=app|'topic'!item`) and OLE links are
  constants (their cached values); `INDIRECT` into another workbook is `#REF!`; functions that
  would do I/O (`WEBSERVICE`, `RTD`, `CALL`, `REGISTER.ID`) are unsupported and never perform it.
  Macros (`.xlsm`, `vbaProject.bin`) are never read.

### 9.2 Surgical writer

`xlrecalc._xlsx._write` produces a new package from the source package and new cached values,
given per sheet as a lookup from a formula cell to its value in the file's terms (`n`, `str`,
`b`, `e`, as the reader returns them); `_save.py`, outside `_xlsx`, builds the lookups from the
engine:
- every part it does not change is copied byte for byte, compressed data included: the
  package's own zip writer (`_xlsx/_zip.py`) copies an entry's local header, data and data
  descriptor as stored and writes the central directory, with Zip64 records where needed;
- a sheet part is rewritten by a streaming scan of its bytes: expat reports where each `<c>`
  starts and ends, and for a formula cell whose new value differs from the one the file holds
  only the `t` attribute and the `<v>` element are replaced; everything else, the `<f>`
  element and its attributes included, is copied as written, so namespace prefixes, comments
  and extension elements survive. A sheet in which nothing changed is copied raw;
- whether a value changed is decided while streaming, against the value as Excel reads it
  (numbers by value, so `0.33333333333333331` and `0.3333333333333333` are the same), so no
  change set is kept in memory;
- a formula result is written as `<v>` plus `t` (absent for numbers, `str`, `b`, `e`): numbers
  as the shortest exact text (integral values without `.0`, `1E+16`), text XML-escaped, with
  `_xHHHH_` for a carriage return and the characters XML cannot hold, `_x005F_` before a
  literal `_xHHHH_`, and `xml:space="preserve"` on a `<v>` or `<t>` that starts or ends with
  whitespace (Excel drops it otherwise, §9.1, §9.4);
- `<calcPr fullCalcOnLoad="1">` is kept where present and set where absent (inserted where the
  schema puts `calcPr`), so Excel recalculates when it opens the file and a value xlrecalc got
  wrong never persists in what Excel shows; `full_calc_on_load=False` leaves the file's setting
  as it was, for files Excel should show as written. A file with the setting makes a later
  load treat every formula as stale (§7.4): a full recalculation, a cost and never an error;
- cells on or after a circular reference are written as Excel saves them, the empty text
  (§7.2); uncomputed cells (§7.6) keep the cached value they had; refused formulas (§9.5) are
  left as written and reported: Excel's repair turns each into its cached value, the value the
  engine already used for the cells depending on it;
- cells changed through the API are written in Excel's forms: a number as `<v>`, a text as an
  inline string (`t="inlineStr"`, `<is><t>`, with `xml:space="preserve"` where it needs it),
  which Excel reads and turns into a shared string when it saves, so the shared-strings part is
  never rewritten; a logical or an error with its `t`; a formula as `<f>` in file form with its
  value. A cleared cell keeps its style (`<c r="A3" s="1"/>`) or goes. A cell or row the sheet
  does not hold is inserted in order, and a row that gains cells loses its `spans` hint. The
  cell keeps its `r` and `s`, and loses `t`, `cm` and `vm`;
- an array formula evaluated here is written with its range: the formula (`<f t="array"
  ref="B1:B5">`, with `cm` for a dynamic array), the cells it fills, and the cells a smaller
  spill no longer reaches, cleared (§7.7). A dynamic array needs the dynamic-array cell
  metadata: the file's own when it has it, else a metadata part is added (with its
  relationship and content type); a file whose metadata lacks it gets a legacy array formula
  over the same range. `#SPILL!` and the other newer errors are written as `#VALUE!`, as Excel
  caches them;
- when the master of a shared formula is changed, every other cell of it gets its own formula,
  the master's text moved to it (`_formula/_fastkey.py` `translate`: relative parts shift, a
  reference moved off the sheet is `#REF!`). Cells of an array formula or data table cannot be
  changed one by one (`ValueError`), as in Excel;
- `calcChain.xml` is kept while no formula is removed: a new or changed formula is fine with
  the chain Excel saved (§9.4). When a formula cell is made a constant or cleared, the chain is
  removed with its relationship and content-type entry, as Excel needs (§9.4);
- the output is written to a temporary file next to the target, synced, given the target's
  permissions (the source's for a new file) and moved into place, so a failure leaves the
  target as it was. Only UTF-8 parts are rewritten.

On the 1M-row benchmark (3,000,003 formulas, openpyxl's file) the write takes 20.2 s, the
reader's load 17.2 s and the evaluation 13.2 s; a write that changes nothing, 13.8 s
(`bench/recalc_bench.py`, `bench/results/SUMMARY.md`).

Changes made through the API (constants, formulas) are recorded per sheet in a change set the
writer consumes; formula results need none, as above. A file Excel saved, with shared formulas
and a calculation chain, given edits of every kind, opens in Excel and recalculates to the
values written (`tests/test_excel_live.py`). Inserting or deleting rows and columns, styles and
anything structural is not in scope.

### 9.3 Working with openpyxl

openpyxl cannot hold a cached value for a formula cell in memory (`cell.value` is the formula),
and its load and save drop every cached value (§9.4). xlrecalc works with it on files: save with
openpyxl, then `xlrecalc.recalculate(path)` loads the file with the reader, evaluates everything
stale (with openpyxl's files, every formula) and writes the values back with the surgical
writer. An adapter that evaluates an in-memory openpyxl workbook is planned (§13).

### 9.4 How Excel treats a written-back file

Probe `capture/writeback_probe.py` (Excel 16.114.3 for Mac, openpyxl 3.1.5). A workbook with
formulas returning a number, text, a logical, an error and a sum was written by openpyxl; Excel
saved a copy; variants of both had their cached values replaced by deliberately wrong values
and were opened with alerts off, where a file Excel would offer to repair fails to open
(error -50).

| file | Excel opens it | Excel shows |
|---|---|---|
| openpyxl's file, values patched (`calcId="124519" fullCalcOnLoad="1"`) | yes | its own results |
| the same without `fullCalcOnLoad` | yes | its own results (the older `calcId` makes it recalculate) |
| Excel's file, values patched, `calcChain.xml` kept | yes | **the patched values** |
| Excel's file, values patched, `calcChain.xml` removed (part, relationship, content type) | yes | the patched values |
| Excel's file, values patched, only the part removed (relationship left) | yes | the patched values |
| Excel's file, values patched, `fullCalcOnLoad="1"` (chain kept or removed) | yes | its own results |
| Excel's file, a formula cell made a constant, chain kept | **no: repair** | — |
| Excel's file, a formula cell added, chain kept | yes | the patched value of the new cell |
| control: a number cell holding `abc` | no: repair | — |

When Excel saves, it writes `calcChain.xml` again: it rebuilds one that was removed, adds a new
formula cell to a kept one, and drops `fullCalcOnLoad`. It keeps patched values it trusted
(saving the file with the wrong values patched writes them again). openpyxl writes a formula
without a cached value as `<c r="A2"><f>A1*3</f><v /></c>` (no `t`), writes no `calcChain.xml`,
and sets `<calcPr calcId="124519" fullCalcOnLoad="1"/>`; loading an Excel file and saving it with
openpyxl drops every cached value and `calcChain.xml` and adds `fullCalcOnLoad="1"`.

Written by the package's own writer (`tests/test_excel_live.py`, run with `--tier excel`):
Excel opens every written file without a repair (each capture suite's workbook, the benchmark
workbooks up to a million rows), and a file Excel saved, given cached values of every kind with
`fullCalcOnLoad` left off, shows and saves again exactly the values written: 17-digit numbers,
line breaks, tabs, carriage returns, `_xHHHH_` escapes, control characters, every error code.
Leading and trailing whitespace of a `<v>` survives only with `xml:space="preserve"`: without
it Excel drops it (`" leading"` reads as `leading`, `"  "` as the empty text), and the reader
does the same.

Consequences: values written into a file Excel saved are what Excel users see until something
recalculates, so they must be right or the file must ask for a recalculation; a removed formula
needs the calculation chain removed; patching values alone needs nothing else.

### 9.5 What Excel's repair does to a refused formula

Capture `repair` (`capture/repair_probe.py`, `tests/captures/repair.tsv`; Excel 16.114.3 for
Mac). One workbook held each formula Excel refuses in a file (§5.2: the error literals `#SPILL!`
and `#CALC!`, the spill operator `A1#` and `Sheet2!A1#`, `Table1[@[Unit Price]]` and
`[@[Unit Price]]*2` outside the table, a tab, `SUM( )`) between valid control formulas, every
formula cell with a deliberately wrong cached value and the calculation settings of a file
Excel saved, so that Excel trusts cached values. Opened with alerts on, Excel offered to recover
the content; after the repair was accepted:

- every refused formula was gone, and its cell held its cached value as a constant (the cell
  of `#SPILL!`, cached 102, became the number 102, saved as `<c r="A2"><v>102</v></c>`);
- every control formula was kept and recalculated (cached 101, it showed and saved 2): after a
  repair, Excel computes the remaining formulas.

A refused formula without a cached value (openpyxl's `<v />`) leaves a blank cell: the
repaired file holds no `<c>` for it and `ISBLANK` of it is TRUE (capture `repair_uncached`,
`repair_probe.py --uncached`).

The reader does the same by default: a refused formula becomes a constant holding the cell's
cached value (a blank cell when the file has none), it is listed in the report, and the cells
that depend on it are computed from that constant.

## 10. Captures and tests

**Function results** (`capture/pipeline.py`, `capture/suites.py`; method and Excel facts in
`CAPTURING.md`). A suite lists cases, each a formula with placeholders for its input cells and
the input values; openpyxl writes a workbook with one case per row (the formula in column A,
inputs after it, floats as `repr` text); Excel opens it, calculates, and saves a copy; the
cached values are read from the copy's XML, which gives exact doubles, error codes and types
(AppleScript returns `missing value` for an error). Each suite is a TSV under
`tests/captures/` with a header recording the Excel build, date system, locale, capture date,
generator and method; columns `suite`, `tag`, `case`, `formula`, `inputs`, `expected_type`,
`expected_repr`. Comparison is exact: same kind, same value, numbers bit for bit, except for a function that declares a tolerance (§8.1). `tests/test_captures.py` computes
every row it can, reports per suite how many match, are within a declared tolerance, differ, are known deviations (each with its
reason) or are not evaluated yet, and fails on any difference.

**Whole-model behavior.** Workbooks built for a behavior (coercion, implicit intersection, spill,
circular references, the write-back rules of §9.4) captured the same way.

**Formula corpus.** The parser's round trip and the order in which functions are implemented
need real formulas. Expected values always come from Excel recomputing formulas put into
workbooks this project generates, never from a third party's cached values. Real-world
collections of workbooks hold personal data: they are used locally, for formula text and
function frequencies, and never committed. Formula text from openly licensed test suites is
committed only where the license allows, with attribution. Downloaded workbooks are untrusted:
they are opened only by this package's reader or openpyxl, in a directory of their own, never in
Excel, and only `.xlsx` files are read. To capture Excel's result for a formula found in one,
the capture tooling writes a fresh workbook with openpyxl holding only that formula and the
constant inputs it refers to; before writing, `capture/corpus_filter.py` drops and counts DDE
formulas (`app|'topic'!item`), references to other workbooks, and calls of functions with side
effects or file or network access (`WEBSERVICE`, `FILTERXML`, `RTD`, `CALL`, `REGISTER.ID`,
`INFO`, `CELL`) or `HYPERLINK` (harmless, but not worth capturing from such a formula). The
exclusion covers formulas taken from those workbooks. Capture suites written for this package call
`CELL`, `HYPERLINK`, `FILTERXML` and the like directly, with fixed local input.

**Other tests.** Tokenizer and printer round trip over the formula corpus; normalizer (template keys,
anchors, sheets); code generation safety (§5.4); dependency index (property tests: reverse
lookup agrees with brute force); scheduler (chains, fan-out, cycles, volatility); reader
(against openpyxl on the same files); writer (untouched parts byte-identical, text and numbers read back as written, atomic
replacement; `tests/test_write.py`); Excel opening what the writer wrote, with `--tier excel`
(§9.4). The default `pytest` run needs neither Excel nor openpyxl. `bench/` reruns the measurements of §2 and the acceptance test of §6.4.

## 11. Package layout and API

### 11.1 Layout

```
src/xlrecalc/
  __init__.py      public names
  _values.py       ErrorValue, BLANK, Ref
  _quirks.py       Quirks
  _numtext.py      number to text, text to number, dates and times in text (§4.3)
  _coerce.py       conversions shared by operators and functions (§4.3)
  _compare.py      operators, exact lookups, criteria, lookup keys (§4.3)
  _arith.py        no subnormals, cancellation snapping (§4.3, §4.4)
  _dates.py        the 1900 and 1904 calendars
  _arrays.py       arrays: broadcasting, filling a declared range (§7.7)
  _ops.py          the operators compiled formulas call (§5.4)
  _formula/        tokenizer, parser, printer, template keys, raw-text key and translation,
                   code generator, the file-form printer and its prefix table (§5.2)
  _store.py        cell store, with per-column row indexes (§7.5)
  _engine.py       templates, scheduling, evaluation context, report (§7)
  _deps.py         the reverse-dependency index and dirty runs (§6)
  _load.py         an engine from an xlsx file: the trust rule, refused formulas (§7.4, §9.5)
  _report.py       Report, CellRef (§7.6)
  _workbook.py     load(), Workbook: values, changes, saving (§11.2)
  _save.py         the engine's values in the writer's file terms (§9.2)
  _recalc.py       recalculate() (§9.2)
  _functions/      registry and one module per category (§8)
  _xlsx/           reader (Book, ReadLimits), surgical writer and zip writer (§9.1, §9.2);
                   standard library only, no imports from the rest of xlrecalc
capture/           capture tooling (not in the wheel)
tools/             developer tools: the formula round trip over a set of workbooks, the
                   prefix table from the captures, the coverage matrix
bench/             benchmarks, a separate uv project with the other engines it measures
```

### 11.2 Public API

```python
import xlrecalc

report = xlrecalc.recalculate("model.xlsx")              # in place, atomically (§9.2)
report = xlrecalc.recalculate("model.xlsx", "out.xlsx", full=True)
# recalculate(source, destination=None, *, full=False, full_calc_on_load=True,
#             limits=None, quirks=None) -> Report
report.uncomputed                      # {CellRef("Sheet 1", 7, 2), ...}; str(): 'Sheet 1'!B7

wb = xlrecalc.load("model.xlsx")       # cached values trusted per §7.4; or a binary file
wb.sheet_names                         # ["Sheet1", ...]
wb.value("Sheet1!B7")                  # stale precedents evaluated; float|str|bool|ErrorValue|None
wb.values("Sheet1!A1:D100")            # list of rows
wb.formula("Sheet1!B8")                # "=B7*2", the formula bar's form; None for a constant
wb.set("Sheet1!A1", 42)                # a constant (text stays text); None clears
wb.set_formula("Sheet1!B8", "=B7*2")   # as typed in the formula bar; FormulaError if refused
wb.recalculate()                       # everything stale; full=True for everything -> Report
wb.save("out.xlsx")                    # surgical write (§9.2); in place by default
wb.report                              # unsupported functions, uncomputed cells, cycles
```

One workbook is used from one thread at a time.

## 12. Open decisions

| question | leaning |
|---|---|
| custom functions | a design of their own: registration per workbook or engine, not process-wide; a rule for a file naming a function not installed; they take references too, so unions and arrays of references are designed with them |
| the `_xlsx` subpackage's future | extracted into its own package once a second user (sheetpeek) needs its API |
| numpy | optional extra, never required |

## 13. Roadmap

1. 0.2: Excel's iteration for `IRR`, `XIRR`, `RATE` and `YIELD` (its starting step, update
   rule and stopping test, inferred from capture sweeps), aiming to match the rows QUIRKS.md
   lists.
2. 0.2: a live openpyxl adapter: `xlrecalc.from_openpyxl(workbook)` would build an engine from an
   in-memory openpyxl `Workbook` through its public API only (`iter_rows`, `cell.value` and
   `cell.data_type`, `ArrayFormula`, `wb.defined_names`, `ws.tables`, `wb.epoch`), evaluate
   values on demand, and after the caller edits the workbook rebuild only the cells whose value
   or formula changed (`engine.refresh()`, or `engine.changed("Sheet1!A1", ...)` naming them).
   It would answer code, not persist values: saving with values stays `recalculate`. openpyxl
   would be an optional extra, imported when the adapter is used.
3. 0.2: the functions that make arrays: `FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE` and
   `MMULT`, with the `#CALC!` error they produce. Spill, `#SPILL!`, `A1#` and `@` are in place
   (§7.7).
4. 0.2: unions of ranges together with arrays of references, which QUIRKS.md lists as not
   computed.
5. 0.2: the functions still absent, by family, each family's captures first. An order, not a promise
   or a date: the functions of item 3 are required; the other families (distributions and their
   inverses, complex numbers, Bessel functions and `CONVERT`, bonds with odd periods, the array and
   matrix functions) follow as far as they get; those with large questions of their own (regular
   expressions, `GETPIVOTDATA`, `FORECAST.ETS`) only if tractable. `docs/coverage.md` lists what
   exists.
6. `xlrecalc.evaluate(text)` for one formula outside a workbook, custom functions, and `LET` and `LAMBDA` with its helpers are designed together,
   from a list of what an editor built on the package needs.
7. Then: a setting or capture for Excel for Windows' text order, the array-backed store and
   vectorized templates (numpy as an optional extra), a setting for
   Excel for Windows' character set in `CHAR` and `CODE` (Windows-1252), the SQLite backend,
   parallel evaluation of independent template columns on free-threaded Python, and a compiled
   core only if profiling demands it.
