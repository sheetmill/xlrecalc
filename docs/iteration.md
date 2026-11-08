# Circular references and iterative calculation

With iterative calculation off, as Excel's default, the cells on a circular reference and those
depending on one keep the values the file holds, as Excel keeps them, and the report lists them in
`circular`. With it on (File > Options > Formulas, saved in the file), xlrecalc iterates as Excel
does: the same order of cells, the same `iterateCount` and `iterateDelta`, the same stopping test,
and the same starting values, checked against Excel's results for each case it was captured on.

- `recalculate(path)` gives what Excel would save after opening the file: Excel's opening runs the
  loop once and its saving once more, and so does `recalculate`, whether or not the file holds
  cached values (with `full=True`, or a file that asks for a full calculation, or formulas without
  values, as openpyxl writes them).
- A trusting load (`load(path)`, or `recalculate(path)` of a file whose values are trusted) changes
  nothing, though Excel would run the loop once more on opening the file.
- Each edit and each save runs the loop once, as in Excel, so `save()` changes the values in memory
  as well as in the file.
- A loop that never converges, such as a counter adding 1 to itself, moves on every run, as it does
  in Excel.
- `ReadLimits.max_iteration_evaluations` (100,000,000 by default, about two minutes) bounds the cell
  evaluations of one run; past it the cells keep their values and are reported. A service that
  accepts workbooks from others should lower it.
