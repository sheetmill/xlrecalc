# xlrecalc

xlrecalc recalculates Excel workbooks in pure Python. It evaluates a workbook's formulas with
the results Excel computes and writes them back into the file as the cached values Excel
stores, so that openpyxl's data-only mode, pandas and other readers see values where a
workbook written by a program has only formulas. Results are tested against values captured
on Excel for Mac 16.115 (Microsoft 365); they are not yet checked against Excel for Windows,
and Excel in other languages and locales is not yet supported.

Documentation: <https://xlrecalc.readthedocs.io>.

```bash
pip install xlrecalc
```

Python 3.11 or later. The two dependencies are [xlformat](https://github.com/sheetmill/xlformat),
which applies Excel's number formats for `TEXT`, and [xlio](https://github.com/sheetmill/xlio),
which reads and writes the xlsx files. The examples below make their workbooks with openpyxl,
installed separately (`pip install openpyxl`); xlrecalc itself does not need it.

It reads `.xlsx`, `.xlsm`, `.xltx` and `.xltm` files and writes each back as its own type, macros
kept as they are and never run; a binary (`.xlsb`) or legacy (`.xls`) workbook is refused with
`xlrecalc.PackageError`.

<!-- usage-start -->
## Recalculating a file

openpyxl writes formulas without their values, so a reader of the file finds nothing in those
cells. `recalculate` computes them and writes the values into the file:

```python
>>> import pathlib, tempfile, openpyxl, xlrecalc
>>> path = pathlib.Path(tempfile.mkdtemp()) / "model.xlsx"
>>> wb = openpyxl.Workbook()
>>> ws = wb.active
>>> ws.append(["Units", "Price", "Revenue"])
>>> ws.append([120, 9.5, "=A2*B2"])
>>> ws.append([80, 12.25, "=A3*B3"])
>>> ws["C4"] = "=SUM(C2:C3)"
>>> wb.save(path)
>>> print(openpyxl.load_workbook(path, data_only=True).active["C4"].value)
None
>>> report = xlrecalc.recalculate(path)  # in place; or recalculate(source, destination)
>>> openpyxl.load_workbook(path, data_only=True).active["C4"].value
2120
>>> report.uncomputed  # cells it could not compute, left as they were
set()
```

A cell xlrecalc cannot compute keeps the value the file holds for it; a file openpyxl wrote holds
none, so such a cell, and every cell depending on it, is written with no value. Nothing is
raised: check `report.uncomputed` after each `recalculate`. It is not the only place a cell left
as it was can be: the report also names cells that keep the file's value because they rely on a
function or feature xlrecalc does not implement (`from_kept`), cells on a circular reference
with iteration off (`circular`), and formulas Excel itself refuses in a file (`refused`);
`unsupported` counts what was missing, and `exceptions` names cells where an internal error was
caught, each written as `#VALUE!`. `workbook.status(cell)` gives one cell's entry (a
`CellStatus`: its state, the cell's own reason, and whether the state comes from a precedent),
reading only the sheets that cell needs, where `workbook.report` reads every sheet.

Only the cached values and the request to recalculate change: every other part of the file is
copied as it was, byte for byte. By default the file asks Excel to recalculate when it opens it,
so a value xlrecalc got wrong never persists in what Excel shows (`full_calc_on_load=False`
saves the file without that request, as Excel saves it).

Which formulas it computes follows what Excel does on opening the file (`recalc_on_load="excel"`,
the default). It computes every formula when Excel would recalculate the file in full: a file
openpyxl wrote, one an older Excel saved, one asking for a full calculation, or one with a
volatile function in a cell Excel did not mark `ca="1"`. Otherwise, in a file Excel saved, it
computes the cells Excel marked `ca="1"` (every volatile cell and every cell depending on one) and
keeps the cached values of the rest.
In manual calculation Excel keeps every cached value on opening and recalculates when it saves:
`load` keeps a manual-calculation file's cached values, and `recalculate`, which saves, computes
what Excel recalculates when it saves one. `recalc_on_load="always"` computes every
formula; `"never"` keeps every cached value but the volatile formulas'. In every mode a formula
without a cached value is computed, and so is what depends on a cell whose value changes.
`RecalcCheck` says whether a load would compute any formula, from the formulas a caller reads,
without building an engine.

### Functions newer than Excel 2007, written with openpyxl

In a file, Excel writes a function added since Excel 2007 with a prefix, and openpyxl writes a
formula as it is given. Without the prefix, Excel does not know the name: `=UNIQUE(A1:A9)`
shows `#NAME?`, and `=SORT(A1:A9)` or `=FILTER(...)` make Excel refuse the file, as does a `LET`
whose names lack their own prefix, `_xlpm.`. Write the prefixes with openpyxl:

```
ws["B1"] = "=_xlfn.XLOOKUP(5,A1:A9,C1:C9)"
ws["D1"] = "=_xlfn.UNIQUE(A1:A9)"
ws["E1"] = "=_xlfn._xlws.SORT(A1:A9)"      # SORT and FILTER take _xlfn._xlws.
ws["F1"] = "=_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)"  # LET's names take _xlpm.
```

xlrecalc reads such a file as Excel does and names the cause in the report:
`Report.unsupported` holds `UNIQUE written without Excel's file prefix (_xlfn.): Excel shows
#NAME?`, and a bare `SORT` or `FILTER`, or a `LET` name without `_xlpm.`, is in
`Report.refused`. openpyxl writes no dynamic-array marker either, so its formulas are legacy
ones, computed with implicit intersection: `D1` above gets one value and does not spill.
`Workbook.set_formula` takes the names as typed in Excel and writes the prefixes itself, and for
a function that makes an array, the marker that lets the result spill.

## Working with a workbook

`load` reads a workbook into memory. Reading a value computes what it needs; changing a cell
makes its dependents stale; `save` writes the changed cells and the new values.

```python
>>> book = xlrecalc.load(path)
>>> book.value("Sheet!C4")
2120.0
>>> book.set("B2", 10)  # a constant
>>> book.value("C4")
2180.0
>>> book.set_formula("D4", '=TEXT(C4,"#,##0.00")')
>>> book.value("D4")
'2,180.00'
>>> book.values("C2:C4")
[[1200.0], [980.0], [2180.0]]
>>> book.save(path.with_name("scenario.xlsx"))
```

`set` always stores a constant, unlike openpyxl, where assigning `"=A1"` makes a formula:
`book.set("A2", "=A1")` stores the text `=A1`. Use `set_formula` for formulas. Numbers come back
as floats, as Excel holds them; errors as `xlrecalc.ErrorValue`. An error compares equal to its
text (`ErrorValue.NA == "#N/A"`, as xlformat's errors do), so a set or a dict key holds the
error `#N/A` and the text `"#N/A"` as one: tell them apart by type, `isinstance(v,
xlrecalc.ErrorValue)`.

`load` takes a path or a binary stream; `save` and `recalculate` take paths, writing a temporary
file beside the target and moving it into place. One workbook is used from one thread at a time.
Every file is read as untrusted: nothing it names is run or fetched, and
`limits=xlrecalc.ReadLimits(...)` bounds what one file may cost. The defaults admit any workbook
Excel opens, so a service that takes files from others should lower them, for instance to
`xlrecalc.ReadLimits.UNTRUSTED`.

A program that already reads a workbook's cells with `xlio` hands them to xlrecalc as they pass,
so the file is read once:

```python
>>> import xlio
>>> with xlio.open(path) as book:
...     loader = xlrecalc.Loader(book)
...     for sheet in book.sheets:
...         for cell in loader.cells(sheet):
...             pass  # the program's own use of the cell
...     workbook = loader.finish()  # the Workbook load would give
>>> workbook.value("Sheet!C4")
2120.0
```

`loader.cells(sheet, rich=True)` gives each cell as `sheet.cells(rich=True)` does, with its text's
runs and phonetic guides. Sheets are fed in any order. `loader.finish()` reads every sheet not fed,
or left part-way, so the book may be closed after it. `loader.workbook()` gives the workbook at any
moment and may leave such sheets to be read when a value needs them: the book must then stay open
while the workbook is used, and once it is closed, a read of a sheet not read yet raises
`ValueError`. A sheet fed after `workbook()` or `finish()` is recorded as it passes, unless the
loader has read it already: then its cells pass through unrecorded.
`loader.opening` is what the load computes on opening (`"all"`, `"some"` or `"none"`), as
`RecalcCheck` answers it. A workbook made this way computes and changes as any other, but has no
file to save over.

## Evaluating one formula

`evaluate` computes one formula without a workbook, from cell values given as a mapping or
by a callback; a cell given no value is blank, as in Excel. A formula starting with `=` is
read as typed in Excel's formula bar; without `=`, in the file form a file stores.

```python
>>> xlrecalc.evaluate("=SUM(A1:A3)*Rate", {"A1:A3": [[1], [2], [3]]}, names={"Rate": 2})
12.0
>>> xlrecalc.evaluate("=SEQUENCE(2,2)")  # a dynamic array: its rows
((1.0, 2.0), (3.0, 4.0))
>>> rule = xlrecalc.Formula.parse("=A1>5", origin="A1")  # compiled once
>>> [rule.evaluate({"A1:A3": [[3], [9], [6]]}, at=f"A{r}") for r in (1, 2, 3)]
[False, True, True]
>>> rule.references
(Reference(sheet=None, first_row=1, first_col=1, last_row=1, last_col=1),)
>>> xlrecalc.evaluate("=SUM(Data!B:B)", resolver=lambda ref: [[10], [20]])
30.0
```

Relative references move from `origin` to the cell `at`, as a conditional format's formula does
over its range. Errors are values (`#DIV/0!`, and `#NAME?` for a name Excel does not know); a
function or form Excel computes and xlrecalc does not (a union of ranges, a table reference,
`CELL("format")`) raises `UnsupportedError`, and `Formula.unsupported` names those the formula's
text shows before evaluating.

## What it computes, and what not yet

- **Functions**: 474 of Excel's functions (17 of them partly): math, statistics,
  logic, text, dates, lookup (`VLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`...), the
  financial and bond functions, the database functions, complex numbers, Bessel functions,
  unit conversion.
  The [coverage matrix](https://xlrecalc.readthedocs.io/en/latest/coverage.html) lists every
  Excel function, whether it is computed, and the captured evidence for it.
- **Evidence**: 119,132 formulas computed by Excel and compared with xlrecalc's results: 118,348
  match, bit for bit or within a small tolerance declared per function (a few units in the last
  place, or a relative one: 1E-9 for the iterative functions, 1E-15 to 1E-12 for the beta,
  gamma, chi-squared, F and t distributions), 784 known deviations, each listed with its reason
  in [QUIRKS.md](https://github.com/sheetmill/xlrecalc/blob/main/QUIRKS.md). Each listed
  deviation is still computed and compared with Excel on every test run, and one that starts to
  match fails the run, except a few that can match by chance (a random result, a result Excel
  varies between runs, a row that depends on the platform's math library). The largest group of deviations is `BINOM.DIST`,
  `BINOM.DIST.RANGE`, `NEGBINOM.DIST`, `POISSON.DIST` and `POISSON` past their tolerance, most at
  1,000 trials and more, where Excel's terms match none of the methods tried (121 rows). Among the
  others, `CUMIPMT` and `CUMPRINC` at tiny rates and in tiny principals carry Excel's noise, and
  from some starting guesses
  Excel's `IRR` and `RATE` reach another root
  than xlrecalc, or a rate where xlrecalc gives `#NUM!`, and `XIRR` stops about 1E-8 from the
  root; 79 of the 352 captured rows of `IRR`, `XIRR` and `RATE` differ that way.
- **Excel's behavior**: coercion between kinds, error values, the 15-digit rules, the 1900
  leap-year bug and the 1904 date system, arrays, dynamic arrays and spill, implicit
  intersection, circular references and iterative calculation, as Excel computes them
  ([circular references](https://xlrecalc.readthedocs.io/en/latest/iteration.html)).
  Four quirks can be turned off (`Quirks`).
- **Not yet**: the functions the
  [coverage matrix](https://xlrecalc.readthedocs.io/en/latest/coverage.html) lists as absent;
  unions of ranges; structured (table) references and 3-D references; `CELL`'s formatting
  kinds; computed criteria in the database functions; `LINEST` with several regressors;
  inserting or deleting rows and columns. A formula xlrecalc cannot compute keeps the value the
  file holds, or is written with none where the file holds none, and the report lists it.

<!-- usage-end -->
## Speed and memory

On a workbook of 1,000,000 rows and 3,000,003 formulas written by openpyxl (so no formula has
a cached value), `recalculate` reads the file, computes every formula and writes the values
back in 50.5 s, with a peak memory of 887 MB. Changing the first cell of a dependency chain
a million cells long and computing again takes 5.3 s. Measured with xlrecalc 0.3.0 and xlio
0.3.0 on an Apple M4 Pro with 24 GB, CPython 3.11.18. Formulas of the same relative shape share
one compiled function, and dependents come from their geometry rather than per-cell links, so
most of the memory holds the cells' values.

To reproduce, in a clone of the repository:

```bash
cd bench && uv run python gen.py 1000000 && cd ..   # the workbook (openpyxl)
uv run python bench/recalc_bench.py 1000000         # read, compute, write back
uv run python bench/engine_bench.py 1000000 --xlsx  # compute, then change A1
```

## Part of sheetmill

xlrecalc is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats,
[xlio](https://github.com/sheetmill/xlio), which reads and surgically writes xlsx files, and
[sheetpeek](https://github.com/sheetmill/sheetpeek), which shows workbooks as Excel does.

## Development

`uv sync --group docs` then `uv run pytest`; [`CONTRIBUTING.md`](https://github.com/sheetmill/xlrecalc/blob/main/CONTRIBUTING.md) has the pip
route, the full set of checks and the commit conventions.
[`ARCHITECTURE.md`](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) describes the design and the measurements behind it,
[`CAPTURING.md`](https://github.com/sheetmill/xlrecalc/blob/main/CAPTURING.md) how results are captured from Excel,
[`bench/`](https://github.com/sheetmill/xlrecalc/tree/main/bench) reruns the measurements.

Most of the code is written with Claude Code, reviewed before each commit and tested against
results captured from Excel; see [`CONTRIBUTING.md`](https://github.com/sheetmill/xlrecalc/blob/main/CONTRIBUTING.md).

## License

MIT.
