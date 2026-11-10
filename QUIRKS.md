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
doubles (`ROUNDUP(2.675,2)` is 2.6799999999999997; capture `rounding`); `IMCOT` of a real number is 0 from
2^27, where `IMTAN` and `SIN` are `#NUM!`, and `IMLOG2` divides by the double just below ln 2, so `IMLOG2(1E20)` is
66.4385618977473, a unit above the true value's 15th digit (capture `complex_family`).

## Known deviations

- An array past `ReadLimits.max_array_cells` (ten full columns by default) is not made: the
  formula keeps the value the file holds and is reported. Over a whole sheet this stops
  `SUMPRODUCT`, `CORREL`, `LINEST`, `XIRR` and `XNPV`. Excel computes `SUMPRODUCT` there without making the
  array; given `CORREL` over the whole sheet it too keeps the value the file holds, though it
  computes `CORREL` over ten columns (captures `array_size`, `array_functions`). `LINEST`,
  `XIRR` and `XNPV` were not captured at that size. The database functions walk their non-blank
  cells, and are computed at any size.
- A function given an array of references (`N(OFFSET(Z1,{0;1;2;3},0))`, `SUBTOTAL(9,OFFSET(...))`,
  `COUNTIF(INDIRECT({"Z1";"Z2"}),">0")`) is not computed: the formula keeps the value the file
  holds and is reported (16 formulas of capture `array_refs`, listed in
  `tests/test_array_refs.py`). `OFFSET`, `INDIRECT` and `INDEX` given an array as the whole formula
  give Excel's value.
- A formula set through the API that reads a cycle in a workbook whose cached values are trusted, before
  any evaluation has reached the cycle, is computed from the cycle's cached value; Excel gives 0 (capture
  `circular_cached`). Once evaluation has met the cycle, the formula is 0, as in Excel.
- `XLOOKUP` and `XMATCH` with the regular-expression match mode (3) are not computed, and are
  reported. Nor is a union of ranges (`LARGE((A1:A3,C1:C3),1)`, `SUM((A1,B2))`): the formula keeps
  the value the file holds and is reported (capture `modern`). `ROW` and `COLUMN` of a union are
  `#REF!`, as in Excel.
- `FORMULATEXT` gives a legacy formula as the file holds it, without the `@` the formula bar shows
  (`=@{1,2;3,4}`; capture `info_family`). `AREAS` and `ISREF` of a union are computed, from the
  formula alone; the other functions of that family are not, as for any union.
- `WEEKNUM` with a type the documentation lacks gives `#NUM!` here. Excel's result for such a type can change between runs, and even between two cells holding the same formula in one calculation, `#NUM!` or a week number, so no result can be expected to repeat it; every such row of capture `dates_family` is listed, whatever it gives. `DATEDIF` from December 31, 1900 to March
  1, 2000 in `YD` is 60 here and 61 in Excel, where every neighbor fits the rule (one row).
- `FINDB`, `SEARCHB`, `LEFTB`, `RIGHTB`, `MIDB`, `LENB` and `REPLACEB` behave as this English Excel for Mac computes
  them, counting no text double (capture `text_family`). With a double-byte language as the default, Excel documents
  them counting such text as two; that belongs to localization, which the package does not do (ARCHITECTURE.md §1.2).
- `ENCODEURL`, `FILTERXML` and `EUROCONVERT` give `#NAME?` in Excel 16 for Mac in every file form (capture `mac_gaps`),
  so there is no reference for them here: they are not implemented, and a cell using one keeps its cached value and is
  reported.
- `NUMBERVALUE` with digits as separators reads a group separator after the decimal one as a digit: `NUMBERVALUE(" 1 2 ",1,2)`
  is 0.2 in Excel and `#VALUE!` here (one row of capture `text_family`).
- Declared tolerances, `statistics_family`: the sums of powers and regressions (`AVEDEV`, `DEVSQ`, `GEOMEAN`, `HARMEAN`,
  `KURT`, `SKEW`, `SKEW.P`, `TRIMMEAN`, `FISHER`, `FISHERINV`, the regression family, `PHI`, `GAUSS`, `CONFIDENCE`) within
  4 units in the last place; `GEOMEAN` near 1E150 and `KURT` and `FORECAST` on a few rows are further, and listed.
- Declared tolerances, `finance_family`: `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `MIRR` within 8 units in the last place,
  `EFFECT`, `NOMINAL`, `ISPMT`, `PDURATION`, `RRI`, the discounted securities and Treasury bills within 4. 84 rows of `CUMIPMT`
  and `CUMPRINC` are further, at tiny rates and in tiny principals, where Excel's result carries noise that neither a sum
  of periods nor a difference of balances reproduces (with one more in `empty_args`, 85 rows in all); one row of `TBILLEQ` too.
- Declared tolerances, `distributions_family`: the discrete distributions (`BINOM.DIST`, `BINOMDIST`, `BINOM.DIST.RANGE`,
  `NEGBINOM.DIST`, `NEGBINOMDIST`, `POISSON.DIST`, `POISSON`, `HYPGEOM.DIST`, `HYPGEOMDIST`) within 8 units in the last
  place; `GAMMA`, `GAMMALN`, `ERF`, `ERFC` and their `.PRECISE` forms, `EXPON.DIST`, `WEIBULL.DIST`, `LOGNORM.DIST` and
  `LOGNORM.INV` within 4 (`EXPON.DIST` matches bit for bit on the reference platform, but about ten of its rows lie within
  0.15 units of a rounding midpoint, where another platform's `exp` can round the other way). 121 rows of the discrete ones are further, most at 1,000 trials and more, where Excel's terms are none of the
  exact count of ways times the powers, `lgamma` or Loader's saddle point (the largest group of known deviations), and
  48 of the others, for Excel's own `exp`, `pow`, `erf` and `gamma`; both groups rest on the platform's math library.
  `BINOM.INV(TRUE,0.5,0.5)` is 1 in Excel, where a cumulative probability of 0.5 reaches alpha at 0, as documented and
  as `BINOM.INV("3",0.5,0.5)` gives at its own tie (one row).
- Declared tolerances, `continuous_family` (5,059 rows): the beta, gamma, chi-squared, F and t distributions, their
  inverses and tests declare a relative tolerance ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md)
  §8.1), each its largest difference from Excel on the capture (in parentheses, on the platform `docs/coverage.md`
  records) rounded up to a power of ten. 1E-12: `CHISQ.DIST` (3.1E-13), `F.DIST` (7.5E-13), `F.DIST.RT` (6.9E-13),
  `F.INV` (9.4E-13), `T.DIST` (1.4E-13), `T.INV` (7.8E-13), `Z.TEST` (1.6E-13), `ZTEST` (1.6E-13). 1E-13: `BETA.DIST`
  (4.7E-14), `BETA.INV` (2.8E-14), `CHISQ.DIST.RT` (6.2E-14), `CHISQ.INV` (3.3E-14), `CHISQ.INV.RT` (1.1E-14),
  `F.INV.RT` (2.8E-14), `GAMMA.DIST` (9.6E-14), `GAMMA.INV` (1.4E-14), `T.DIST.2T` (1.2E-14), `T.DIST.RT` (1.2E-14),
  `T.INV.2T` (3.7E-14), `TINV` (1.2E-14). 1E-14: `BETADIST` (1.5E-15), `BETAINV` (1.2E-15), `CHIDIST` (8.1E-15),
  `CHIINV` (1.0E-15), `CONFIDENCE.T` (5.4E-15), `F.TEST` (2.5E-15), `FDIST` (1.7E-15), `FINV` (4.0E-15), `FTEST`
  (2.5E-15), `GAMMADIST` (1.3E-15), `T.TEST` (2.7E-15), `TDIST` (6.0E-15), `TTEST` (2.7E-15). 1E-15: `CHISQ.TEST`
  (6.1E-16), `CHITEST` (6.1E-16), `GAMMAINV` (9.2E-16).
  110 rows are listed. In 97 of them Excel's result is the further from the true value (each group checked against
  40-digit arithmetic, a closed form or a limit): an inverse at a probability of 1E-300, where Excel's search stops
  short or gives `#NUM!` (19 rows); `T.INV`, `T.INV.2T`, `CHISQ.INV`, `CHISQ.INV.RT` and `CONFIDENCE.T` at 1E10 degrees
  of freedom, about 2E-7 from the root or `#NUM!` (22 rows); `T.INV`, `T.INV.2T` and `CONFIDENCE.T` at 1E6 degrees of
  freedom, 1E-12 to 3E-9 from it (19 rows); `T.INV.2T` near 1 at one degree of freedom, 6E-12 to 3E-5 from
  tan(π(1 − p)/2) (4 rows); the t distribution at 1E-10, exactly 0.5 in Excel, 1 for two tails, where it is 0.5 and
  the density times 1E-10 (21 rows); at an extreme shape or value, `#NUM!` where the result exists, 0, a quantile that
  underflows to 0 or a tail of 2.5E-154 (8 rows), or a search that stops short, 0.999999 or 0.5 for 1 (4 rows).
  `CHISQ.INV.RT` and `CHIINV` at one degree of freedom invert the left tail at 1 − p, as Excel does, so that a p of
  at most 2^-54 (about 5.55E-17) is `#NUM!`; Excel's p is a unit of 1.1E-16 from 1 − (1 − p) on some rows, which
  shows where p is small (12 rows). One row is ours: `BETA.DIST` with both shapes 1E300 is `#NUM!` here, where the
  continued fraction does not converge within its terms, and 0.5 in Excel (1 row).
  Seven formulas at a shape of 1.7E308 (`BETA.DIST` and `GAMMA.DIST` cumulative, `BETA.INV`) keep Excel calculating
  past 90 seconds, alone in a workbook, where the other 55 timed, densities at 1.7E308 among them, take at most 3
  seconds with the open (`tests/captures/slow_cases.tsv`);
  they are not captured, and here they give their limits, 0 or 1, or `#NUM!`. Every function here does bounded work on
  any input: each series and continued fraction stops at 20,000 terms, `#NUM!` past them, and each inverse is a
  bisection over the doubles' bit patterns, at most 64 evaluations.
- The platform's math library: three rows match Excel on the platform `docs/coverage.md` records (bit for bit, or within the
  function's declared tolerance) and differ on another, where a last-bit difference in `erfc`, `exp`, `log1p` or `expm1` survives a cancellation: `GAUSS(1E-5)`, which
  subtracts 0.5 from the normal distribution (as Excel's own result shows), and two rows of `IPMT` with payments due in
  advance (`empty_args`). They are listed on every platform, so the published counts do not depend on it.
- The platform's math library, in text: 67 rows of `complex_family` match Excel here and may differ elsewhere. A complex
  result is text, each part rounded to 15 significant digits, so a declared tolerance in units in the last place cannot
  apply; where a part computed through `sin`, `cos`, `exp`, `log`, `sinh`, `cosh` or `atan2` lies within one unit in the
  last place of a midpoint between two 15-digit decimals (63 rows), or an imaginary part is exactly 1, written `i`
  where the double below it is written `1i` (9 rows, 5 of them among the 63), another platform's last bit can round
  the text the other way: on Linux, `IMSEC("-2.5-1.5j")` gives `-0.364130793588459+...` where Excel gives
  `-0.36413079358846+...`. `tests/test_complex_boundary.py` derives the rows. They are listed on every platform.
- `VDB` from a fractional start, or late in a schedule, differs in its last bits on 75 of the 317 rows of `finance` that
  sweep fractional periods (5 to 1,131 units in the last place, Excel's rule otherwise reproduced): the book value carried
  period by period rounds differently from Excel's, and a straight-line amount computed once or a closed-form book value
  fit no more rows. Listed in `tests/test_captures.py`.
- An empty argument: `VDB` with an empty `factor` is `#DIV/0!` in Excel and 0 here (two rows of `empty_args`). Four others there carry the known noise of `MULTINOMIAL`, `CUMIPMT` and `XIRR`.
- How text orders is captured only in part: every pair of 39 ASCII characters and 24 pairs of words,
  on Excel 16 for Mac (capture `modern`). Those give one consistent order, marks and symbols before
  digits before letters, with accents after their base letter. The rest is approximated: other
  ASCII controls sort first and other characters after the letters, by code point, so another
  script or a symbol outside ASCII may not sort as Excel sorts it. Excel for Windows' order has not
  been captured and may differ.
- A row off Excel's grid in a file (`r="99999999999"`): Excel opens the file and drops the cell
  (capture `value_forms`), though it refuses rows 0 and 1048577; the reader refuses every row
  off the grid with a `PackageError`.
- `VDB` (`finance`) within 4 units in the last place but five rows: the last period of a
  long schedule (7 to 48 units), and two edge cases Excel answers oddly (`#DIV/0!` for start
  = end = life at a rate of 1 per period; cost - salvage when salvage exceeds cost).
- Database functions (`misc`, `database_family`): a computed criterion is not supported; the cell is reported and keeps
  the file's value. `DVAR`, `DVARP`, `DSTDEV` and
  `DSTDEVP` declare 4 units in the last place for Excel's sum of squares; two rows are 6 and 10 units away. Blank records
  are counted up to the sheet's last row holding a value, where Excel's used range also counts formatted rows.
- Random numbers: two rows of `misc` draw them (`RANDBETWEEN(1.5,2.5)`, and two draws compared), and four of
  `empty_args` (`RANDARRAY` with bounds that leave a range), so no result can be expected to repeat Excel's; all six are
  listed whatever the draw. `RANDARRAY`'s dynamic arrays (capture `dyn_randarray`) are checked through plain formulas that
  count their elements within the bounds, and whole ones with the integer flag.
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
- Declared tolerances ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1), `math_trig` and `math_integers`: trigonometry, hyperbolics,
  logarithms and `SERIESSUM` within 4 units in the last place, for Excel's own math library; `FACT` and `FACTDOUBLE`
  within 8 and `COMBIN`, `COMBINA` and `PERMUT` within 4, for the order Excel multiplies in, which a few tried orders
  did not reproduce. `MULTINOMIAL` is exact, where Excel's result carries noise from its logarithm of the gamma
  function (279.99999999999966 for 280); 12 rows of `math_integers` listed in `tests/test_captures.py`, with one of
  `statistics_family` and one of `empty_args`, 14 in all.
- `IRR`, `XIRR`, `RATE`, `YIELD` ([ARCHITECTURE.md](https://github.com/sheetmill/xlrecalc/blob/main/ARCHITECTURE.md) §8.1, §8.2): a relative tolerance of
  1E-9, declared per function, since Excel's Newton iterates rest on its own NPV, whose last
  bits are not reproduced, and Excel guarantees its result only to its documented stopping
  tolerance. Of the 402 rows of `iterative`, 323 match, exactly or within the relative
  tolerance, and 79 are listed: from some guesses Excel's `IRR` and `RATE` reach another root,
  a rate where the package gives `#NUM!`, or the reverse; `XIRR` stops about 1E-8 from the
  root (sometimes after what looks like a bracketing search); and three `IRR` results at a
  rate of 0 differ by less than 1E-14, where a relative tolerance cannot apply. Compared with
  the capture, each in at most five variants: Newton on the rate, on 1 + rate and on 1 / (1 + rate), analytic and
  numerical derivatives, stops on the step or on the NPV, 20 to 100 iterations, damping. Excel's
  iteration is on the Roadmap (ARCHITECTURE.md §13). Each row is listed in
  `tests/test_captures.py`.
