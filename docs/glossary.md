# Glossary

Terms the other pages use, as they apply to xlrecalc.

Cached value
: The value a file stores beside a formula: what Excel showed when it last calculated. Readers
  that do not compute (openpyxl's data-only mode, pandas) read it. openpyxl writes formulas
  without one; `recalculate` writes them.

Data-only mode
: openpyxl's `load_workbook(path, data_only=True)`, which gives each formula cell's cached value
  instead of its formula.

Recalculation on opening; a trusting and an evaluating load
: Excel recalculates every formula of some files when it opens them (no calculation settings, a
  request to recalculate, a calculation by an older Excel, a file Excel did not save, a volatile
  function in a cell without Excel's `ca="1"` mark); of the others it recalculates the cells so
  marked (volatile cells and what depends on them) and shows the rest's cached values; in manual
  calculation it shows every cached value. `load` and `recalculate` follow the same rule by
  default (`recalc_on_load="excel"`): a trusting load keeps the file's cached values and computes
  what Excel recalculates and what changes, an evaluating load computes every formula.
  `recalc_on_load="always"` evaluates every file; `"never"` keeps every cached value but the
  volatile formulas'. In every mode a formula without a cached value is computed.

Volatile function
: A function whose result can change with nothing it reads changing: `RAND`, `RANDBETWEEN`,
  `RANDARRAY`, `NOW`, `TODAY`, `OFFSET`, `INDIRECT` and `CELL`. A formula calling one, itself or
  through a defined name, is volatile; Excel recalculates it and its dependents on opening.

Dynamic array; spill
: A formula Excel 365 marks in the file (the dynamic-array cell metadata) gives all of an array
  result, which fills ("spills into") the cells below and to the right of it; `#SPILL!` when a
  cell in the way holds something.

Legacy formula; implicit intersection
: A formula without that marker, computed as Excel before 365 computed it: where one value is
  expected and a range is given, it takes the range's cell in the formula's own row or column
  (implicit intersection, written `@` in Excel 365), and an array result shows its first value.
  openpyxl writes legacy formulas; `Workbook.set_formula` writes the marker.

Typed form; file form
: A formula as typed in Excel's formula bar (`=XLOOKUP(...)`), and as a file stores it, where
  functions added since Excel 2007 carry a prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`) and the
  names of `LET` and `LAMBDA` carry `_xlpm.`. A name written without its prefix is one Excel does
  not know.

Refused formula
: A formula Excel does not take in a file: opening the file, Excel offers to repair it, and the
  repair turns the formula into its cached value. xlrecalc lists it in `Report.refused` and uses
  that value.

Capture; suite; row
: The evidence: formulas Excel computed in a workbook built for the purpose, with the values it
  saved, read back exactly ([Capturing](capturing.md)). A suite is one such set
  (`tests/captures/<suite>.tsv`), a row one formula of it.

Known deviation
: A captured row where xlrecalc's value still differs from Excel's, listed with its reason in
  [Quirks](quirks.md) and checked on every test run.

Declared tolerance; units in the last place
: How far from Excel's value a result may be and still count as matching, declared per function:
  a number of units in the last place (the spacing between neighboring double-precision numbers
  at that value, so 2 units means at most two representable numbers apart) or a relative
  tolerance such as 1E-9.
