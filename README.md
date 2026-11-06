# xlrecalc

xlrecalc recalculates Excel workbooks in pure Python. It evaluates a workbook's formulas with
the results Excel computes and writes them back into the file as the cached values Excel
stores, so that openpyxl's data-only mode, pandas and other readers see values where a
workbook written by a program has only formulas. Results are tested against values captured
from Excel 16 for Mac.

Documentation: <https://xlrecalc.readthedocs.io>.

```bash
pip install xlrecalc
```

Python 3.11 or later. The one dependency is [xlformat](https://github.com/sheetmill/xlformat),
which applies Excel's number formats for `TEXT`.

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

Only the cached values change: every other part of the file is copied as it was, byte for
byte. By default the file also asks Excel to recalculate when it opens it, so a value
xlrecalc got wrong never persists in what Excel shows (`full_calc_on_load=False` leaves the
file's setting alone).

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
`book.set("A2", "=A1")` stores the text `=A1`. Use `set_formula` for formulas. Numbers come
back as floats, as Excel holds them; errors as `xlrecalc.ErrorValue`, which is not a string,
so the error `#N/A` and the text `"#N/A"` stay apart.

## What it computes, and what not yet

- **Functions**: 156 of Excel's functions (8 of them partly), the most used ones first: math, statistics,
  logic, text, dates, lookup (`VLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`...), the
  financial and bond functions, the database functions. The
  [coverage matrix](https://xlrecalc.readthedocs.io/en/latest/coverage.html) lists every
  Excel function, whether it is computed, and the captured evidence for it.
- **Evidence**: 12,733 formulas computed by Excel and compared with xlrecalc's results: 12,613
  match, bit for bit or within a small tolerance declared per function (a few units in the last
  place, or a relative 1E-9 for the iterative functions), 118 known deviations, each listed
  with its reason in
  [QUIRKS.md](https://github.com/sheetmill/xlrecalc/blob/main/QUIRKS.md), and 2 not
  evaluated (they call `XLOOKUP`, not implemented yet). The largest group of deviations is the
  iterative functions: from some starting guesses Excel's `IRR` and `RATE` reach another root
  than xlrecalc, or a rate where xlrecalc gives `#NUM!`, and `XIRR` stops about 1E-8 from the
  root; 73 of the 271 captured rows of `IRR`, `XIRR` and `RATE` differ that way.
- **Excel's behavior**: coercion between kinds, error values, the 15-digit rules, the 1900
  leap-year bug and the 1904 date system, arrays, dynamic arrays and spill, implicit
  intersection, circular references (left uncalculated, as Excel does with iteration off).
  Four quirks can be turned off (`Quirks`).
- **Not yet**: iterative calculation of circular references; the array-making functions
  (`SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`, `LAMBDA` and its helpers) and `XLOOKUP`; structured
  (table) references and 3-D references; `CELL`'s formatting kinds; computed criteria in the
  database functions; `LINEST` with several regressors; inserting or deleting rows and
  columns. A formula xlrecalc cannot compute keeps the value the file holds, and the report
  lists it.

<!-- usage-end -->
## Speed and memory

On a workbook of 1,000,000 rows and 3,000,003 formulas written by openpyxl (so no formula has
a cached value), `recalculate` reads the file, computes every formula and writes the values
back in 50.6 s, with a peak memory of 1,163 MB. Changing the first cell of a dependency chain
a million cells long and computing again takes 5.1 s. Measured on an Apple M4 Pro with 24 GB,
CPython 3.11.18. Formulas of the same relative shape share one compiled function, and
dependents come from their geometry rather than per-cell links, so most of the memory holds
the cells' values.

To reproduce, in a clone of the repository:

```bash
cd bench && uv run python gen.py 1000000 && cd ..   # the workbook (openpyxl)
uv run python bench/recalc_bench.py 1000000         # read, compute, write back
uv run python bench/engine_bench.py 1000000 --xlsx  # compute, then change A1
```

## Part of sheetmill

xlrecalc is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats, and
[sheetpeek](https://github.com/sheetmill/sheetpeek), which shows workbooks as Excel does.

## Development

`uv sync --group docs` then `uv run pytest`; `CONTRIBUTING.md` has the pip route, the full set of
checks and the commit conventions. `ARCHITECTURE.md` describes the design and the measurements
behind it, `CAPTURING.md` how results are captured from Excel, `bench/` reruns the
measurements.

Most of the code is written with Claude Code, reviewed before each commit and tested against
results captured from Excel; see `CONTRIBUTING.md`.

## License

MIT.
