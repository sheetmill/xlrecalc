# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
