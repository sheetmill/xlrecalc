# Quirks

Excel behaviors that xlrecalc reproduces on purpose and that a caller may switch off through
`Quirks` (see [ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §4.4). Every entry names the switch, states what Excel does,
what `strict` does instead, and the evidence: rows captured from Excel 16.114.3 for Mac under
`tests/captures/`, with the method in `CAPTURING.md`. `leap_year_1900` and
`fifteen_digit_precision` are shared with xlformat and mean the same there.

## `leap_year_1900`

**Excel**: in the 1900 date system, serial 60 is February 29, 1900, a day that did not exist.
`DATE(1900,2,29)` is 60, `DATEVALUE("2/29/1900")` is 60, `YEAR`, `MONTH`, `DAY` of 60 give
1900, 2, 29, and weekdays before March 1900 count the false day: `WEEKDAY(1)` is 1 (Sunday)
although January 1, 1900 was a Monday. Serial 0 is "January 0, 1900" (`DAY(0)` is 0,
`WEEKDAY(0)` is 7).

**Strict**: serial `d` is `1899-12-30 + d`, there is no February 29, 1900, and weekdays are the
true ones. Dates in January and February 1900 differ from Excel by one day.

**Evidence**: suites `dates` (57 rows) and `text_number` (`2/28/1900`, `2/29/1900`,
`3/1/1900`); suite `dates_1904` shows the 1904 system has no such day.

## `fifteen_digit_precision`

**Excel**: number to text keeps at most 15 significant digits, rounded with ties down
(`=0.1+0.2&""` is `0.3`, `=999999999999999.5&""` is `999999999999999`); text to number keeps
the first 15 significant digits and drops the rest (`VALUE("0.12345678901234567")` is
0.123456789012345, `VALUE("12345678901234.567")` is 12345678901234.5). A number literal in a
formula is read the same way when Excel computes it: `=0.30000000000000004` is 0.3,
`=9007199254740993` is 9007199254740990, in any context (an operand, a function argument, an
array constant). A text form longer than 20 characters is written in scientific form, and
with a three-digit exponent the mantissa loses digits to keep to 20, rounded half up
(`=1.2345678901234567E-300&""` is `1.2345678901235E-300`).

**Strict**: number to text is the shortest text that gives the double back
(`0.30000000000000004`); text to number and number literals are correctly rounded.

**Evidence**: suites `number_text` (324 rows), `text_number` (516 rows) and `literals` (126
rows: 18 literals of 15 to 22 significant digits in seven contexts).

## `comparison_rounding`

**Excel**: the comparison operators and the `COUNTIF`-family criteria treat two numbers as
equal when they agree after rounding to 15 significant digits with ties down: `=0.1+0.2=0.3` is
TRUE and `=0.1+0.2>0.3` FALSE; `COUNTIF` counts `0.1+0.2` as 0.3. Exact lookups do not:
`MATCH(0.1+0.2,{0.3},0)` and `VLOOKUP(0.1+0.2,…,FALSE)` give `#N/A`.

**Strict**: numbers compare as doubles.

**Evidence**: suite `compare`, tags `rounding=`, `rounding<`, `rounding>` (seven bases, each
against values 1 to 15 parts in 1E16 away, 1 to 8 units in the last place either way, and its
14-, 15- and 16-digit roundings) and `rounding_formula`, `rounding_match`, `rounding_countif`,
`rounding_vlookup`.

## `cancellation_snapping`

**Excel**: when the last operation of a formula, outside any parentheses, is an addition or a
subtraction, and its result is less than 8 units in the last place of its left operand, the
formula gives 0. `=0.1+0.2-0.3`, `=(0.1+0.2)-0.3` and `=-0.3+0.1+0.2` are 0. Anywhere else the
IEEE result stands: `=(0.1+0.2-0.3)`, `=0.1+0.2-0.3+0`, `=(0.1+0.2-0.3)*1`,
`=IF(TRUE,0.1+0.2-0.3)` and `=SUM(0.1+0.2-0.3)` are 5.55E-17, and `=0.1+0.2-0.3=0` is FALSE.
`SUM` applies the same rule to its own additions: `=SUM(0.1,0.2,-0.3)` is 0.

**Strict**: the IEEE result.

**Evidence**: suite `cancellation` (703 rows): nine bases from 5E-300 to 1E15, each minus a value
0 to 4,096 units in the last place away, as the last operation and inside `(…)*1`; then 5 to 10
units either way, in both operand orders, for bases at both ends of a binade. Measured in units
of the left operand's last place, every result below 7.5 units snapped and every result from 8
units on did not; measured relative to either operand's magnitude, or in units of the larger or
the smaller operand, the snapped and kept cases overlap.

## Excel semantics, not quirks

`CHAR` and `CODE` number characters in the Mac OS Roman set, as Excel for Mac does, and
`CODE` of a character outside the set is 95. Excel for Windows uses Windows-1252 instead;
there is no setting for it yet.

Matched unconditionally, because no caller would want otherwise or there is no meaningful
alternative: no subnormal numbers (a result or a constant below 2.2250738585072014E-308 in
magnitude is 0); `=0^0` and `POWER(0,0)` are `#NUM!`; text compares without case; a blank is 0,
`""` or FALSE by context; `DATE` adds 1900 to a year from 0 to 1899; `ROUND` and `ROUNDDOWN` work on the 15-digit
decimal value, and `ROUNDUP` is `ROUNDDOWN` plus one unit at the last kept place, added in
doubles (`ROUNDUP(2.675,2)` is 2.6799999999999997; capture `rounding`).

## Known deviations

- A row off Excel's grid in a file (`r="99999999999"`): Excel opens the file and drops the cell
  (capture `value_forms`), though it refuses rows 0 and 1048577; the reader refuses every row
  off the grid with a `PackageError`.
- `VDB` (`finance`) within 4 units in the last place but five rows: the last period of a
  long schedule (7 to 48 units), and two edge cases Excel answers oddly (`#DIV/0!` for start
  = end = life at a rate of 1 per period; cost - salvage when salvage exceeds cost).
- Database functions (`misc`): a computed criterion is not supported (five rows).
- Random numbers (`misc`): two rows draw them (`RANDBETWEEN(1.5,2.5)`, and two draws
  compared), so no result can be expected to repeat Excel's; both are listed whatever the draw.
- Declared tolerances ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1), statistics: the variances and standard
  deviations (five formulas tried), `CORREL`, `LINEST` and the normal distribution within 4
  units in the last place on every captured row of `statistics` but 21: `LINEST` near
  cancellation (an intercept of -0.02 from means of 7.05 and 3.5; R squared of values near
  1E9), 26 to 4,607 units away, and the normal distribution's lower tail (z of -2.3 and
  below), 5 to 80 units away, where Excel is closer to the true value than `math.erfc` and
  does not use Cody's algorithm either.
- Declared tolerances ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1), bonds: `PRICE` within 8 units in the last
  place (7 at quarterly coupons; two summation orders tried), `DURATION` and `MDURATION`
  within 4, on every captured row of `bonds` but those listed below.
- `YIELD` (`bonds`): with one coupon period left Excel's closed form takes A as `COUPDAYBS`,
  E as the actual period in bases 1 to 3 (360/frequency in 0 and 4) and the days to
  redemption in the basis's own count; every such row matches. With more periods its
  iteration on `PRICE` is within the relative tolerance (1E-9); one zero-coupon row is 1.3E-9
  away. A bond maturing on February 29 gets a 366-day basis-1 `COUPDAYS` from
  Excel (four rows), and `PRICE` at a 1000% yield is 64 units away (one row).
- Declared tolerances ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1): `PMT`, `IPMT` and `PPMT` work with `exp`,
  `log1p` and `expm1`, whose last bits differ between Excel and the platform's math library,
  and `NPV`'s summation of discounted flows could not be reproduced in the last bits by any
  tried order, power or compensated sum; each is within 4 units in the last place of Excel on
  every captured row of `finance`, except six `IPMT` rows at a fractional
  period or late in a long term, 13 to 109 units away, listed in `tests/test_captures.py`.
- `IRR`, `XIRR`, `RATE`, `YIELD` ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1, §8.2): a relative tolerance of
  1E-9, declared per function, since Excel's Newton iterates rest on its own NPV, whose last
  bits are not reproduced, and Excel guarantees its result only to its documented stopping
  tolerance. Of the 271 rows of `iterative`, 198 match, exactly or within the relative
  tolerance, and 73 are listed: from some guesses Excel's `IRR` and `RATE` reach another root,
  a rate where the package gives `#NUM!`, or the reverse; `XIRR` stops about 1E-8 from the
  root (sometimes after what looks like a bracketing search); and three `IRR` results at a
  rate of 0 differ by less than 1E-14, where a relative tolerance cannot apply. Tried, at most
  five hypotheses each: Newton on the rate, on 1 + rate and on 1 / (1 + rate), analytic and
  numerical derivatives, stops on the step or on the NPV, 20 to 100 iterations, damping. An
  investigation of Excel's iteration is planned. Each row is listed in
  `tests/test_captures.py`.
