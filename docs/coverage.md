# Function coverage

Every worksheet function Excel has, whether this package computes it, and the
evidence: captured rows that call the function (a row counts for every function in
its formula), how many match Excel, exactly or within the function's declared
tolerance (in units in the last place, or relative for the iterative functions and
the beta, gamma, chi-squared, F and t distributions), and how many are listed
deviations with their reasons (`QUIRKS.md`, `tests/test_captures.py`,
`tests/test_dynamic_arrays.py`). A random draw counts as listed: it is tested by its
properties, not its value, so a draw that equals Excel's by chance is never counted
as matching. Whether a matching row is exact or within its tolerance depends on the
platform's math library, so the page counts them together and is the same on every
platform.

461 implemented, 17 partial, 45 absent.

| function | category | status | rows | match | listed | tolerance | note |
|---|---|---|---:|---:|---:|---|---|
| `ABS` | math | implemented | 18 | 18 | 0 |  |  |
| `ACCRINT` | | absent | | | | | |
| `ACCRINTM` | | absent | | | | | |
| `ACOS` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `ACOSH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `ACOT` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `ACOTH` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `ADDRESS` | lookup and reference | implemented | 152 | 152 | 0 |  |  |
| `AGGREGATE` | | absent | | | | | |
| `AMORDEGRC` | | absent | | | | | |
| `AMORLINC` | | absent | | | | | |
| `ANCHORARRAY` | lookup and reference | implemented | 0 | 0 | 0 |  |  |
| `AND` | logical | implemented | 28 | 28 | 0 |  |  |
| `ARABIC` | math | implemented | 45 | 45 | 0 |  |  |
| `AREAS` | lookup and reference | implemented | 30 | 30 | 0 |  |  |
| `ARRAYTOTEXT` | text | implemented | 200 | 200 | 0 |  |  |
| `ASC` | | absent | | | | | |
| `ASIN` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `ASINH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `ATAN` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `ATAN2` | math | implemented | 200 | 200 | 0 | 4 ulps |  |
| `ATANH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `AVEDEV` | statistical | implemented | 35 | 35 | 0 | 4 ulps |  |
| `AVERAGE` | statistical | implemented | 27 | 27 | 0 |  |  |
| `AVERAGEA` | statistical | implemented | 42 | 42 | 0 |  |  |
| `AVERAGEIF` | math | implemented | 74 | 74 | 0 |  |  |
| `AVERAGEIFS` | math | implemented | 87 | 87 | 0 |  |  |
| `BAHTTEXT` | | absent | | | | | |
| `BASE` | math | implemented | 306 | 306 | 0 |  |  |
| `BESSELI` | engineering | implemented | 1287 | 1287 | 0 |  |  |
| `BESSELJ` | engineering | implemented | 1366 | 1366 | 0 | 8 ulps |  |
| `BESSELK` | engineering | implemented | 759 | 759 | 0 |  |  |
| `BESSELY` | engineering | implemented | 963 | 963 | 0 | 8 ulps |  |
| `BETA.DIST` | statistical | implemented | 695 | 694 | 1 | 1e-13 relative |  |
| `BETA.INV` | statistical | implemented | 167 | 157 | 10 | 1e-13 relative |  |
| `BETADIST` | statistical | implemented | 33 | 33 | 0 | 1e-14 relative |  |
| `BETAINV` | statistical | implemented | 17 | 17 | 0 | 1e-14 relative |  |
| `BIN2DEC` | engineering | implemented | 21 | 21 | 0 |  |  |
| `BIN2HEX` | engineering | implemented | 304 | 304 | 0 |  |  |
| `BIN2OCT` | engineering | implemented | 205 | 205 | 0 |  |  |
| `BINOM.DIST` | statistical | implemented | 597 | 532 | 65 | 8 ulps |  |
| `BINOM.DIST.RANGE` | statistical | implemented | 171 | 145 | 26 | 8 ulps |  |
| `BINOM.INV` | statistical | implemented | 182 | 172 | 10 |  |  |
| `BINOMDIST` | statistical | implemented | 94 | 94 | 0 | 8 ulps |  |
| `BITAND` | engineering | implemented | 201 | 201 | 0 |  |  |
| `BITLSHIFT` | engineering | implemented | 229 | 229 | 0 |  |  |
| `BITOR` | engineering | implemented | 201 | 201 | 0 |  |  |
| `BITRSHIFT` | engineering | implemented | 229 | 229 | 0 |  |  |
| `BITXOR` | engineering | implemented | 201 | 201 | 0 |  |  |
| `BYCOL` | logical | implemented | 7 | 7 | 0 |  |  |
| `BYROW` | logical | implemented | 7 | 7 | 0 |  |  |
| `CALL` | | absent | | | | | |
| `CEILING` | math | implemented | 276 | 276 | 0 |  |  |
| `CEILING.MATH` | math | implemented | 247 | 247 | 0 |  |  |
| `CEILING.PRECISE` | math | implemented | 306 | 306 | 0 |  |  |
| `CELL` | information | partial | 6 | 5 | 1 |  | the kinds that depend on formatting are unsupported: styles are not read |
| `CHAR` | text | implemented | 1173 | 1173 | 0 |  |  |
| `CHIDIST` | statistical | implemented | 16 | 16 | 0 | 1e-14 relative |  |
| `CHIINV` | statistical | implemented | 34 | 30 | 4 | 1e-14 relative |  |
| `CHISQ.DIST` | statistical | implemented | 198 | 198 | 0 | 1e-12 relative |  |
| `CHISQ.DIST.RT` | statistical | implemented | 88 | 88 | 0 | 1e-13 relative |  |
| `CHISQ.INV` | statistical | implemented | 131 | 127 | 4 | 1e-13 relative |  |
| `CHISQ.INV.RT` | statistical | implemented | 178 | 164 | 14 | 1e-13 relative |  |
| `CHISQ.TEST` | statistical | implemented | 20 | 20 | 0 | 1e-15 relative |  |
| `CHITEST` | statistical | implemented | 19 | 19 | 0 | 1e-15 relative |  |
| `CHOOSE` | lookup and reference | implemented | 14 | 14 | 0 |  |  |
| `CHOOSECOLS` | lookup and reference | implemented | 44 | 44 | 0 |  |  |
| `CHOOSEROWS` | lookup and reference | implemented | 30 | 30 | 0 |  |  |
| `CLEAN` | text | implemented | 539 | 539 | 0 |  |  |
| `CODE` | text | implemented | 418 | 418 | 0 |  |  |
| `COLUMN` | lookup and reference | implemented | 8 | 8 | 0 |  |  |
| `COLUMNS` | lookup and reference | implemented | 14 | 14 | 0 |  |  |
| `COMBIN` | math | implemented | 290 | 290 | 0 | 4 ulps |  |
| `COMBINA` | math | implemented | 209 | 209 | 0 | 4 ulps |  |
| `COMPLEX` | engineering | implemented | 385 | 385 | 0 |  |  |
| `CONCAT` | text | implemented | 39 | 39 | 0 |  |  |
| `CONCATENATE` | text | implemented | 22 | 22 | 0 |  |  |
| `CONFIDENCE` | statistical | implemented | 805 | 805 | 0 | 4 ulps |  |
| `CONFIDENCE.NORM` | statistical | implemented | 806 | 806 | 0 | 4 ulps |  |
| `CONFIDENCE.T` | statistical | implemented | 99 | 94 | 5 | 1e-14 relative |  |
| `CONVERT` | engineering | implemented | 12232 | 12232 | 0 |  |  |
| `CORREL` | statistical | implemented | 16 | 16 | 0 | 4 ulps |  |
| `COS` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `COSH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `COT` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `COTH` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `COUNT` | statistical | implemented | 33 | 33 | 0 |  |  |
| `COUNTA` | statistical | implemented | 23 | 23 | 0 |  |  |
| `COUNTBLANK` | statistical | implemented | 9 | 9 | 0 |  |  |
| `COUNTIF` | math | implemented | 158 | 158 | 0 |  |  |
| `COUNTIFS` | math | implemented | 53 | 53 | 0 |  |  |
| `COUPDAYBS` | financial | implemented | 96 | 96 | 0 |  |  |
| `COUPDAYS` | financial | implemented | 110 | 109 | 1 |  |  |
| `COUPDAYSNC` | financial | implemented | 96 | 96 | 0 |  |  |
| `COUPNCD` | financial | implemented | 96 | 96 | 0 |  |  |
| `COUPNUM` | financial | implemented | 110 | 110 | 0 |  |  |
| `COUPPCD` | financial | implemented | 96 | 96 | 0 |  |  |
| `COVAR` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `COVARIANCE.P` | statistical | implemented | 26 | 26 | 0 | 4 ulps |  |
| `COVARIANCE.S` | statistical | implemented | 26 | 26 | 0 | 4 ulps |  |
| `CRITBINOM` | statistical | implemented | 67 | 61 | 6 |  |  |
| `CSC` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `CSCH` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `CUBEKPIMEMBER` | | absent | | | | | |
| `CUBEMEMBER` | | absent | | | | | |
| `CUBEMEMBERPROPERTY` | | absent | | | | | |
| `CUBERANKEDMEMBER` | | absent | | | | | |
| `CUBESET` | | absent | | | | | |
| `CUBESETCOUNT` | | absent | | | | | |
| `CUBEVALUE` | | absent | | | | | |
| `CUMIPMT` | financial | implemented | 2994 | 2932 | 62 | 8 ulps |  |
| `CUMPRINC` | financial | implemented | 2994 | 2969 | 25 | 8 ulps |  |
| `DATE` | date and time | implemented | 41 | 41 | 0 |  |  |
| `DATEDIF` | date and time | implemented | 3653 | 3652 | 1 |  |  |
| `DATEVALUE` | date and time | implemented | 138 | 138 | 0 |  |  |
| `DAVERAGE` | database | partial | 141 | 140 | 1 |  | computed criteria are not supported |
| `DAY` | date and time | implemented | 38 | 38 | 0 |  |  |
| `DAYS` | date and time | implemented | 395 | 395 | 0 |  |  |
| `DAYS360` | date and time | implemented | 402 | 402 | 0 |  |  |
| `DB` | financial | implemented | 1494 | 1494 | 0 | 8 ulps |  |
| `DBCS` | | absent | | | | | |
| `DCOUNT` | database | partial | 142 | 141 | 1 |  | computed criteria are not supported |
| `DCOUNTA` | database | partial | 154 | 152 | 2 |  | computed criteria are not supported |
| `DDB` | financial | implemented | 1197 | 1197 | 0 | 8 ulps |  |
| `DEC2BIN` | engineering | implemented | 313 | 313 | 0 |  |  |
| `DEC2HEX` | engineering | implemented | 313 | 313 | 0 |  |  |
| `DEC2OCT` | engineering | implemented | 313 | 313 | 0 |  |  |
| `DECIMAL` | math | implemented | 89 | 89 | 0 |  |  |
| `DEGREES` | math | implemented | 31 | 31 | 0 |  |  |
| `DELTA` | engineering | implemented | 136 | 136 | 0 |  |  |
| `DEVSQ` | statistical | implemented | 35 | 35 | 0 | 4 ulps |  |
| `DGET` | database | partial | 154 | 152 | 2 |  | computed criteria are not supported |
| `DISC` | financial | implemented | 381 | 381 | 0 | 4 ulps |  |
| `DMAX` | database | partial | 141 | 140 | 1 |  | computed criteria are not supported |
| `DMIN` | database | partial | 141 | 140 | 1 |  | computed criteria are not supported |
| `DOLLAR` | text | implemented | 238 | 238 | 0 |  |  |
| `DOLLARDE` | financial | implemented | 225 | 225 | 0 |  |  |
| `DOLLARFR` | financial | implemented | 225 | 225 | 0 |  |  |
| `DPRODUCT` | database | partial | 154 | 152 | 2 |  | computed criteria are not supported |
| `DROP` | lookup and reference | implemented | 29 | 29 | 0 |  |  |
| `DSTDEV` | database | partial | 154 | 152 | 2 | 4 ulps | computed criteria are not supported |
| `DSTDEVP` | database | partial | 154 | 152 | 2 | 4 ulps | computed criteria are not supported |
| `DSUM` | database | partial | 142 | 141 | 1 |  | computed criteria are not supported |
| `DURATION` | financial | implemented | 112 | 111 | 1 | 4 ulps |  |
| `DVAR` | database | partial | 154 | 150 | 4 | 4 ulps | computed criteria are not supported |
| `DVARP` | database | partial | 154 | 152 | 2 | 4 ulps | computed criteria are not supported |
| `EDATE` | date and time | implemented | 42 | 42 | 0 |  |  |
| `EFFECT` | financial | implemented | 148 | 148 | 0 | 4 ulps |  |
| `ENCODEURL` | text | partial | 179 | 179 | 0 |  | computed under platform="windows"; #NAME? under "mac", as Excel for Mac does not know it |
| `EOMONTH` | date and time | implemented | 41 | 41 | 0 |  |  |
| `ERF` | engineering | implemented | 108 | 103 | 5 | 4 ulps |  |
| `ERF.PRECISE` | engineering | implemented | 57 | 57 | 0 | 4 ulps |  |
| `ERFC` | engineering | implemented | 56 | 54 | 2 | 4 ulps |  |
| `ERFC.PRECISE` | engineering | implemented | 57 | 55 | 2 | 4 ulps |  |
| `ERROR.TYPE` | information | implemented | 722 | 692 | 0 |  |  |
| `EUROCONVERT` | | absent | | | | | |
| `EVEN` | math | implemented | 31 | 31 | 0 |  |  |
| `EXACT` | text | implemented | 85 | 85 | 0 |  |  |
| `EXP` | math | implemented | 18 | 18 | 0 |  |  |
| `EXPAND` | lookup and reference | implemented | 37 | 37 | 0 |  |  |
| `EXPON.DIST` | statistical | implemented | 144 | 144 | 0 | 4 ulps |  |
| `EXPONDIST` | statistical | implemented | 17 | 17 | 0 | 4 ulps |  |
| `F.DIST` | statistical | implemented | 546 | 545 | 1 | 1e-12 relative |  |
| `F.DIST.RT` | statistical | implemented | 254 | 254 | 0 | 1e-12 relative |  |
| `F.INV` | statistical | implemented | 204 | 200 | 4 | 1e-12 relative |  |
| `F.INV.RT` | statistical | implemented | 174 | 171 | 3 | 1e-13 relative |  |
| `F.TEST` | statistical | implemented | 24 | 24 | 0 | 1e-14 relative |  |
| `FACT` | math | implemented | 37 | 37 | 0 | 8 ulps |  |
| `FACTDOUBLE` | math | implemented | 37 | 37 | 0 | 8 ulps |  |
| `FALSE` | logical | implemented | 1 | 1 | 0 |  |  |
| `FDIST` | statistical | implemented | 17 | 17 | 0 | 1e-14 relative |  |
| `FILTER` | lookup and reference | implemented | 33 | 33 | 0 |  |  |
| `FILTERXML` | | absent | | | | | |
| `FIND` | text | implemented | 52 | 52 | 0 |  |  |
| `FINDB` | text | implemented | 353 | 353 | 0 |  |  |
| `FINV` | statistical | implemented | 17 | 17 | 0 | 1e-14 relative |  |
| `FISHER` | statistical | implemented | 19 | 19 | 0 | 4 ulps |  |
| `FISHERINV` | statistical | implemented | 19 | 19 | 0 | 4 ulps |  |
| `FIXED` | text | implemented | 259 | 259 | 0 |  |  |
| `FLOOR` | math | implemented | 276 | 276 | 0 |  |  |
| `FLOOR.MATH` | math | implemented | 247 | 247 | 0 |  |  |
| `FLOOR.PRECISE` | math | implemented | 306 | 306 | 0 |  |  |
| `FORECAST` | statistical | implemented | 173 | 172 | 1 | 4 ulps |  |
| `FORECAST.ETS` | | absent | | | | | |
| `FORECAST.ETS.CONFINT` | | absent | | | | | |
| `FORECAST.ETS.SEASONALITY` | | absent | | | | | |
| `FORECAST.ETS.STAT` | | absent | | | | | |
| `FORECAST.LINEAR` | statistical | implemented | 174 | 173 | 1 | 4 ulps |  |
| `FORMULATEXT` | information | partial | 28 | 27 | 1 |  | the @ the formula bar adds to a legacy formula is not shown |
| `FREQUENCY` | statistical | implemented | 24 | 24 | 0 |  |  |
| `FTEST` | statistical | implemented | 23 | 23 | 0 | 1e-14 relative |  |
| `FV` | financial | implemented | 129 | 129 | 0 |  |  |
| `FVSCHEDULE` | financial | implemented | 88 | 88 | 0 |  |  |
| `GAMMA` | statistical | implemented | 64 | 62 | 2 | 4 ulps |  |
| `GAMMA.DIST` | statistical | implemented | 378 | 376 | 2 | 1e-13 relative |  |
| `GAMMA.INV` | statistical | implemented | 210 | 207 | 3 | 1e-13 relative |  |
| `GAMMADIST` | statistical | implemented | 22 | 22 | 0 | 1e-14 relative |  |
| `GAMMAINV` | statistical | implemented | 11 | 11 | 0 | 1e-15 relative |  |
| `GAMMALN` | statistical | implemented | 63 | 59 | 4 | 4 ulps |  |
| `GAMMALN.PRECISE` | statistical | implemented | 64 | 60 | 4 | 4 ulps |  |
| `GAUSS` | statistical | implemented | 20 | 19 | 1 | 4 ulps |  |
| `GCD` | math | implemented | 214 | 214 | 0 |  |  |
| `GEOMEAN` | statistical | implemented | 35 | 33 | 2 | 4 ulps |  |
| `GESTEP` | engineering | implemented | 136 | 136 | 0 |  |  |
| `GETPIVOTDATA` | | absent | | | | | |
| `GROUPBY` | | absent | | | | | |
| `GROWTH` | statistical | implemented | 24 | 24 | 0 | 5e-14 relative |  |
| `HARMEAN` | statistical | implemented | 35 | 35 | 0 | 4 ulps |  |
| `HEX2BIN` | engineering | implemented | 311 | 311 | 0 |  |  |
| `HEX2DEC` | engineering | implemented | 28 | 28 | 0 |  |  |
| `HEX2OCT` | engineering | implemented | 212 | 212 | 0 |  |  |
| `HLOOKUP` | lookup and reference | implemented | 27 | 27 | 0 |  |  |
| `HOUR` | date and time | implemented | 27 | 27 | 0 |  |  |
| `HSTACK` | lookup and reference | implemented | 20 | 20 | 0 |  |  |
| `HYPERLINK` | lookup and reference | implemented | 64 | 64 | 0 |  |  |
| `HYPGEOM.DIST` | statistical | implemented | 634 | 634 | 0 | 8 ulps |  |
| `HYPGEOMDIST` | statistical | implemented | 70 | 70 | 0 | 8 ulps |  |
| `IF` | logical | implemented | 69 | 69 | 0 |  |  |
| `IFERROR` | logical | implemented | 678 | 648 | 0 |  |  |
| `IFNA` | logical | implemented | 16 | 16 | 0 |  |  |
| `IFS` | logical | implemented | 34 | 34 | 0 |  |  |
| `IMABS` | engineering | implemented | 25 | 25 | 0 |  |  |
| `IMAGE` | | absent | | | | | |
| `IMAGINARY` | engineering | implemented | 35 | 35 | 0 |  |  |
| `IMARGUMENT` | engineering | implemented | 16 | 16 | 0 | 4 ulps |  |
| `IMCONJUGATE` | engineering | implemented | 16 | 16 | 0 |  |  |
| `IMCOS` | engineering | implemented | 16 | 16 | 0 |  |  |
| `IMCOSH` | engineering | implemented | 17 | 16 | 1 |  |  |
| `IMCOT` | engineering | implemented | 36 | 33 | 3 |  |  |
| `IMCSC` | engineering | implemented | 17 | 15 | 2 |  |  |
| `IMCSCH` | engineering | implemented | 17 | 14 | 3 |  |  |
| `IMDIV` | engineering | implemented | 70 | 70 | 0 |  |  |
| `IMEXP` | engineering | implemented | 16 | 11 | 5 |  |  |
| `IMLN` | engineering | implemented | 35 | 35 | 0 |  |  |
| `IMLOG10` | engineering | implemented | 29 | 26 | 3 |  |  |
| `IMLOG2` | engineering | implemented | 29 | 23 | 6 |  |  |
| `IMPOWER` | engineering | implemented | 186 | 160 | 26 |  |  |
| `IMPRODUCT` | engineering | implemented | 72 | 72 | 0 |  |  |
| `IMREAL` | engineering | implemented | 42 | 42 | 0 |  |  |
| `IMSEC` | engineering | implemented | 17 | 14 | 3 |  |  |
| `IMSECH` | engineering | implemented | 17 | 14 | 3 |  |  |
| `IMSIN` | engineering | implemented | 16 | 14 | 2 |  |  |
| `IMSINH` | engineering | implemented | 17 | 15 | 2 |  |  |
| `IMSQRT` | engineering | implemented | 16 | 10 | 6 |  |  |
| `IMSUB` | engineering | implemented | 70 | 70 | 0 |  |  |
| `IMSUM` | engineering | implemented | 72 | 72 | 0 |  |  |
| `IMTAN` | engineering | implemented | 36 | 34 | 2 |  |  |
| `INDEX` | lookup and reference | implemented | 1377 | 1343 | 34 |  |  |
| `INDIRECT` | lookup and reference | implemented | 15 | 15 | 0 |  |  |
| `INFO` | | absent | | | | | |
| `INT` | math | implemented | 19 | 18 | 1 |  |  |
| `INTERCEPT` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `INTRATE` | financial | implemented | 381 | 381 | 0 | 4 ulps |  |
| `IPMT` | financial | implemented | 92 | 84 | 8 | 4 ulps |  |
| `IRR` | financial | implemented | 197 | 169 | 28 | 1e-09 relative |  |
| `ISBLANK` | information | implemented | 38 | 38 | 0 |  |  |
| `ISERR` | information | implemented | 19 | 19 | 0 |  |  |
| `ISERROR` | information | implemented | 25 | 25 | 0 |  |  |
| `ISEVEN` | information | implemented | 59 | 59 | 0 |  |  |
| `ISFORMULA` | information | implemented | 26 | 26 | 0 |  |  |
| `ISLOGICAL` | information | implemented | 20 | 20 | 0 |  |  |
| `ISNA` | information | implemented | 22 | 22 | 0 |  |  |
| `ISNONTEXT` | information | implemented | 60 | 60 | 0 |  |  |
| `ISNUMBER` | information | implemented | 51 | 51 | 0 |  |  |
| `ISO.CEILING` | math | implemented | 305 | 305 | 0 |  |  |
| `ISODD` | information | implemented | 59 | 59 | 0 |  |  |
| `ISOMITTED` | information | implemented | 7 | 7 | 0 |  |  |
| `ISOWEEKNUM` | date and time | implemented | 43 | 43 | 0 |  |  |
| `ISPMT` | financial | implemented | 1326 | 1326 | 0 | 4 ulps |  |
| `ISREF` | information | implemented | 44 | 44 | 0 |  |  |
| `ISTEXT` | information | implemented | 22 | 22 | 0 |  |  |
| `JIS` | | absent | | | | | |
| `KURT` | statistical | implemented | 35 | 32 | 3 | 4 ulps |  |
| `LAMBDA` | logical | implemented | 92 | 92 | 0 |  |  |
| `LARGE` | statistical | implemented | 36 | 36 | 0 |  |  |
| `LCM` | math | implemented | 214 | 214 | 0 |  |  |
| `LEFT` | text | implemented | 37 | 37 | 0 |  |  |
| `LEFTB` | text | implemented | 92 | 92 | 0 |  |  |
| `LEN` | text | implemented | 431 | 431 | 0 |  |  |
| `LENB` | text | implemented | 29 | 29 | 0 |  |  |
| `LET` | logical | implemented | 61 | 61 | 0 |  |  |
| `LINEST` | statistical | implemented | 795 | 774 | 21 | 4 ulps |  |
| `LN` | math | implemented | 18 | 18 | 0 |  |  |
| `LOG` | math | implemented | 229 | 229 | 0 | 4 ulps |  |
| `LOG10` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `LOGEST` | statistical | implemented | 119 | 106 | 13 | 5e-14 relative |  |
| `LOGINV` | statistical | implemented | 21 | 21 | 0 | 4 ulps |  |
| `LOGNORM.DIST` | statistical | implemented | 263 | 247 | 16 | 4 ulps |  |
| `LOGNORM.INV` | statistical | implemented | 168 | 168 | 0 | 4 ulps |  |
| `LOGNORMDIST` | statistical | implemented | 51 | 50 | 1 | 4 ulps |  |
| `LOOKUP` | lookup and reference | implemented | 48 | 48 | 0 |  |  |
| `LOWER` | text | implemented | 23 | 23 | 0 |  |  |
| `MAKEARRAY` | logical | implemented | 14 | 14 | 0 |  |  |
| `MAP` | logical | implemented | 18 | 18 | 0 |  |  |
| `MATCH` | lookup and reference | implemented | 145 | 145 | 0 |  |  |
| `MAX` | statistical | implemented | 24 | 24 | 0 |  |  |
| `MAXA` | statistical | implemented | 42 | 42 | 0 |  |  |
| `MAXIFS` | math | implemented | 90 | 90 | 0 |  |  |
| `MDETERM` | math | implemented | 58 | 58 | 0 |  |  |
| `MDURATION` | financial | implemented | 115 | 114 | 1 | 4 ulps |  |
| `MEDIAN` | statistical | implemented | 42 | 42 | 0 |  |  |
| `MID` | text | implemented | 114 | 114 | 0 |  |  |
| `MIDB` | text | implemented | 453 | 453 | 0 |  |  |
| `MIN` | statistical | implemented | 20 | 20 | 0 |  |  |
| `MINA` | statistical | implemented | 42 | 42 | 0 |  |  |
| `MINIFS` | math | implemented | 88 | 88 | 0 |  |  |
| `MINUTE` | date and time | implemented | 27 | 27 | 0 |  |  |
| `MINVERSE` | math | implemented | 303 | 303 | 0 |  |  |
| `MIRR` | financial | implemented | 597 | 597 | 0 | 8 ulps |  |
| `MMULT` | math | implemented | 35 | 35 | 0 |  |  |
| `MOD` | math | implemented | 15 | 15 | 0 |  |  |
| `MODE` | statistical | implemented | 38 | 38 | 0 |  |  |
| `MODE.MULT` | statistical | implemented | 20 | 20 | 0 |  |  |
| `MODE.SNGL` | statistical | implemented | 37 | 37 | 0 |  |  |
| `MONTH` | date and time | implemented | 38 | 38 | 0 |  |  |
| `MROUND` | math | implemented | 276 | 276 | 0 |  |  |
| `MULTINOMIAL` | math | implemented | 46 | 32 | 14 |  |  |
| `MUNIT` | math | implemented | 40 | 40 | 0 |  |  |
| `N` | lookup and reference | implemented | 23 | 23 | 0 |  |  |
| `NA` | information | implemented | 34 | 34 | 0 |  |  |
| `NEGBINOM.DIST` | statistical | implemented | 397 | 383 | 14 | 8 ulps |  |
| `NEGBINOMDIST` | statistical | implemented | 32 | 32 | 0 | 8 ulps |  |
| `NETWORKDAYS` | date and time | implemented | 140 | 140 | 0 |  |  |
| `NETWORKDAYS.INTL` | date and time | implemented | 77 | 77 | 0 |  |  |
| `NOMINAL` | financial | implemented | 148 | 148 | 0 | 4 ulps |  |
| `NORM.DIST` | statistical | implemented | 23 | 22 | 1 | 4 ulps |  |
| `NORM.INV` | statistical | implemented | 24 | 24 | 0 | 4 ulps |  |
| `NORM.S.DIST` | statistical | implemented | 45 | 41 | 4 | 4 ulps |  |
| `NORM.S.INV` | statistical | implemented | 21 | 21 | 0 | 4 ulps |  |
| `NORMDIST` | statistical | implemented | 32 | 31 | 1 | 4 ulps |  |
| `NORMINV` | statistical | implemented | 35 | 35 | 0 | 4 ulps |  |
| `NORMSDIST` | statistical | implemented | 34 | 29 | 5 | 4 ulps |  |
| `NORMSINV` | statistical | implemented | 20 | 20 | 0 | 4 ulps |  |
| `NOT` | logical | implemented | 18 | 18 | 0 |  |  |
| `NOW` | date and time | implemented | 1 | 1 | 0 |  |  |
| `NPER` | financial | implemented | 120 | 120 | 0 |  |  |
| `NPV` | financial | implemented | 88 | 88 | 0 | 4 ulps |  |
| `NUMBERVALUE` | text | implemented | 522 | 522 | 0 |  |  |
| `OCT2BIN` | engineering | implemented | 306 | 306 | 0 |  |  |
| `OCT2DEC` | engineering | implemented | 23 | 23 | 0 |  |  |
| `OCT2HEX` | engineering | implemented | 207 | 207 | 0 |  |  |
| `ODD` | math | implemented | 31 | 31 | 0 |  |  |
| `ODDFPRICE` | | absent | | | | | |
| `ODDFYIELD` | | absent | | | | | |
| `ODDLPRICE` | | absent | | | | | |
| `ODDLYIELD` | | absent | | | | | |
| `OFFSET` | lookup and reference | implemented | 33 | 33 | 0 |  |  |
| `OR` | logical | implemented | 26 | 26 | 0 |  |  |
| `PDURATION` | financial | implemented | 646 | 646 | 0 | 4 ulps |  |
| `PEARSON` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `PERCENTILE` | statistical | implemented | 56 | 56 | 0 |  |  |
| `PERCENTILE.EXC` | statistical | implemented | 125 | 125 | 0 |  |  |
| `PERCENTILE.INC` | statistical | implemented | 55 | 55 | 0 |  |  |
| `PERCENTOF` | | absent | | | | | |
| `PERCENTRANK` | statistical | implemented | 181 | 181 | 0 |  |  |
| `PERCENTRANK.EXC` | statistical | implemented | 182 | 182 | 0 |  |  |
| `PERCENTRANK.INC` | statistical | implemented | 182 | 182 | 0 |  |  |
| `PERMUT` | math | implemented | 290 | 290 | 0 | 4 ulps |  |
| `PERMUTATIONA` | math | implemented | 209 | 209 | 0 |  |  |
| `PHI` | statistical | implemented | 20 | 20 | 0 | 4 ulps |  |
| `PHONETIC` | | absent | | | | | |
| `PI` | math | implemented | 2 | 2 | 0 |  |  |
| `PIVOTBY` | | absent | | | | | |
| `PMT` | financial | implemented | 141 | 141 | 0 | 4 ulps |  |
| `POISSON` | statistical | implemented | 53 | 51 | 2 | 8 ulps |  |
| `POISSON.DIST` | statistical | implemented | 132 | 118 | 14 | 8 ulps |  |
| `POWER` | math | implemented | 15 | 15 | 0 |  |  |
| `PPMT` | financial | implemented | 82 | 82 | 0 | 4 ulps |  |
| `PRICE` | financial | implemented | 123 | 121 | 2 | 8 ulps |  |
| `PRICEDISC` | financial | implemented | 381 | 381 | 0 | 4 ulps |  |
| `PRICEMAT` | | absent | | | | | |
| `PROB` | statistical | implemented | 102 | 102 | 0 |  |  |
| `PRODUCT` | math | implemented | 20 | 20 | 0 |  |  |
| `PROPER` | text | implemented | 29 | 29 | 0 |  |  |
| `PV` | financial | implemented | 131 | 131 | 0 |  |  |
| `QUARTILE` | statistical | implemented | 27 | 27 | 0 |  |  |
| `QUARTILE.EXC` | statistical | implemented | 93 | 93 | 0 |  |  |
| `QUARTILE.INC` | statistical | implemented | 26 | 26 | 0 |  |  |
| `QUOTIENT` | math | implemented | 208 | 208 | 0 |  |  |
| `RADIANS` | math | implemented | 31 | 31 | 0 |  |  |
| `RAND` | math | implemented | 5 | 4 | 1 |  |  |
| `RANDARRAY` | math | implemented | 29 | 19 | 10 |  |  |
| `RANDBETWEEN` | math | implemented | 9 | 7 | 2 |  |  |
| `RANK` | statistical | implemented | 24 | 24 | 0 |  |  |
| `RANK.AVG` | statistical | implemented | 22 | 22 | 0 |  |  |
| `RANK.EQ` | statistical | implemented | 22 | 22 | 0 |  |  |
| `RATE` | financial | implemented | 208 | 202 | 6 | 1e-09 relative |  |
| `RECEIVED` | financial | implemented | 381 | 381 | 0 | 4 ulps |  |
| `REDUCE` | logical | implemented | 10 | 10 | 0 |  |  |
| `REGEXEXTRACT` | | absent | | | | | |
| `REGEXREPLACE` | | absent | | | | | |
| `REGEXTEST` | | absent | | | | | |
| `REGISTER.ID` | | absent | | | | | |
| `REPLACE` | text | implemented | 128 | 128 | 0 |  |  |
| `REPLACEB` | text | implemented | 1350 | 1350 | 0 |  |  |
| `REPT` | text | implemented | 107 | 107 | 0 |  |  |
| `RIGHT` | text | implemented | 64 | 64 | 0 |  |  |
| `RIGHTB` | text | implemented | 92 | 92 | 0 |  |  |
| `ROMAN` | math | implemented | 168 | 168 | 0 |  |  |
| `ROUND` | math | implemented | 678 | 678 | 0 |  |  |
| `ROUNDDOWN` | math | implemented | 678 | 678 | 0 |  |  |
| `ROUNDUP` | math | implemented | 678 | 678 | 0 |  |  |
| `ROW` | lookup and reference | implemented | 13 | 13 | 0 |  |  |
| `ROWS` | lookup and reference | implemented | 22 | 22 | 0 |  |  |
| `RRI` | financial | implemented | 646 | 646 | 0 | 4 ulps |  |
| `RSQ` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `RTD` | | absent | | | | | |
| `SCAN` | logical | implemented | 11 | 11 | 0 |  |  |
| `SEARCH` | text | implemented | 403 | 403 | 0 |  |  |
| `SEARCHB` | text | implemented | 353 | 353 | 0 |  |  |
| `SEC` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `SECH` | math | implemented | 32 | 32 | 0 | 4 ulps |  |
| `SECOND` | date and time | implemented | 27 | 27 | 0 |  |  |
| `SEQUENCE` | math | implemented | 45 | 45 | 0 |  |  |
| `SERIESSUM` | math | implemented | 69 | 69 | 0 | 4 ulps |  |
| `SHEET` | information | implemented | 28 | 28 | 0 |  |  |
| `SHEETS` | information | implemented | 28 | 28 | 0 |  |  |
| `SIGN` | math | implemented | 31 | 31 | 0 |  |  |
| `SIN` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `SINGLE` | lookup and reference | implemented | 1 | 1 | 0 |  |  |
| `SINH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `SKEW` | statistical | implemented | 35 | 35 | 0 | 4 ulps |  |
| `SKEW.P` | statistical | implemented | 34 | 34 | 0 | 4 ulps |  |
| `SLN` | financial | implemented | 41 | 41 | 0 |  |  |
| `SLOPE` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `SMALL` | statistical | implemented | 35 | 35 | 0 |  |  |
| `SORT` | lookup and reference | implemented | 137 | 137 | 0 |  |  |
| `SORTBY` | lookup and reference | implemented | 27 | 27 | 0 |  |  |
| `SQRT` | math | implemented | 18 | 18 | 0 |  |  |
| `SQRTPI` | math | implemented | 31 | 31 | 0 |  |  |
| `STANDARDIZE` | statistical | implemented | 821 | 821 | 0 |  |  |
| `STDEV` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `STDEV.P` | statistical | implemented | 41 | 41 | 0 | 4 ulps |  |
| `STDEV.S` | statistical | implemented | 41 | 41 | 0 | 4 ulps |  |
| `STDEVA` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `STDEVP` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `STDEVPA` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `STEYX` | statistical | implemented | 25 | 25 | 0 | 4 ulps |  |
| `STOCKHISTORY` | | absent | | | | | |
| `SUBSTITUTE` | text | implemented | 521 | 521 | 0 |  |  |
| `SUBTOTAL` | math | implemented | 21 | 21 | 0 |  |  |
| `SUM` | math | implemented | 179 | 179 | 0 |  |  |
| `SUMIF` | math | implemented | 90 | 90 | 0 |  |  |
| `SUMIFS` | math | implemented | 45 | 45 | 0 |  |  |
| `SUMPRODUCT` | math | implemented | 22 | 22 | 0 |  |  |
| `SUMSQ` | statistical | implemented | 40 | 40 | 0 |  |  |
| `SUMX2MY2` | math | implemented | 9 | 9 | 0 |  |  |
| `SUMX2PY2` | math | implemented | 9 | 9 | 0 |  |  |
| `SUMXMY2` | math | implemented | 9 | 9 | 0 |  |  |
| `SWITCH` | logical | implemented | 41 | 41 | 0 |  |  |
| `SYD` | financial | implemented | 261 | 261 | 0 |  |  |
| `T` | text | implemented | 35 | 35 | 0 |  |  |
| `T.DIST` | statistical | implemented | 234 | 227 | 7 | 1e-12 relative |  |
| `T.DIST.2T` | statistical | implemented | 104 | 97 | 7 | 1e-13 relative |  |
| `T.DIST.RT` | statistical | implemented | 108 | 101 | 7 | 1e-13 relative |  |
| `T.INV` | statistical | implemented | 131 | 114 | 17 | 1e-12 relative |  |
| `T.INV.2T` | statistical | implemented | 160 | 139 | 21 | 1e-13 relative |  |
| `T.TEST` | statistical | implemented | 336 | 336 | 0 | 1e-14 relative |  |
| `TAKE` | lookup and reference | implemented | 38 | 38 | 0 |  |  |
| `TAN` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `TANH` | math | implemented | 31 | 31 | 0 | 4 ulps |  |
| `TBILLEQ` | financial | implemented | 271 | 270 | 1 | 4 ulps |  |
| `TBILLPRICE` | financial | implemented | 271 | 271 | 0 | 4 ulps |  |
| `TBILLYIELD` | financial | implemented | 271 | 271 | 0 | 4 ulps |  |
| `TDIST` | statistical | implemented | 115 | 115 | 0 | 1e-14 relative |  |
| `TEXT` | text | implemented | 1173 | 1173 | 0 |  |  |
| `TEXTAFTER` | text | implemented | 2448 | 2448 | 0 |  |  |
| `TEXTBEFORE` | text | implemented | 2448 | 2448 | 0 |  |  |
| `TEXTJOIN` | text | implemented | 85 | 85 | 0 |  |  |
| `TEXTSPLIT` | text | implemented | 78 | 78 | 0 |  |  |
| `TIME` | date and time | implemented | 551 | 551 | 0 |  |  |
| `TIMEVALUE` | date and time | implemented | 128 | 128 | 0 |  |  |
| `TINV` | statistical | implemented | 24 | 24 | 0 | 1e-13 relative |  |
| `TOCOL` | lookup and reference | implemented | 35 | 35 | 0 |  |  |
| `TODAY` | date and time | implemented | 1 | 1 | 0 |  |  |
| `TOROW` | lookup and reference | implemented | 17 | 17 | 0 |  |  |
| `TRANSLATE` | | absent | | | | | |
| `TRANSPOSE` | lookup and reference | implemented | 16 | 16 | 0 |  |  |
| `TREND` | statistical | implemented | 26 | 26 | 0 | 5e-14 relative |  |
| `TRIM` | text | implemented | 24 | 24 | 0 |  |  |
| `TRIMMEAN` | statistical | implemented | 246 | 246 | 0 | 4 ulps |  |
| `TRIMRANGE` | | absent | | | | | |
| `TRUE` | logical | implemented | 2 | 2 | 0 |  |  |
| `TRUNC` | math | implemented | 307 | 307 | 0 |  |  |
| `TTEST` | statistical | implemented | 335 | 335 | 0 | 1e-14 relative |  |
| `TYPE` | information | implemented | 60 | 60 | 0 |  |  |
| `UNICHAR` | text | implemented | 706 | 706 | 0 |  |  |
| `UNICODE` | text | implemented | 287 | 287 | 0 |  |  |
| `UNIQUE` | lookup and reference | implemented | 37 | 37 | 0 |  |  |
| `UPPER` | text | implemented | 24 | 24 | 0 |  |  |
| `USDOLLAR` | text | implemented | 198 | 198 | 0 |  |  |
| `VALUE` | text | implemented | 150 | 150 | 0 |  |  |
| `VALUETOTEXT` | text | implemented | 194 | 194 | 0 |  |  |
| `VAR` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `VAR.P` | statistical | implemented | 41 | 41 | 0 | 4 ulps |  |
| `VAR.S` | statistical | implemented | 41 | 41 | 0 | 4 ulps |  |
| `VARA` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `VARP` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `VARPA` | statistical | implemented | 40 | 40 | 0 | 4 ulps |  |
| `VDB` | financial | implemented | 505 | 423 | 82 | 4 ulps |  |
| `VLOOKUP` | lookup and reference | implemented | 95 | 95 | 0 |  |  |
| `VSTACK` | lookup and reference | implemented | 26 | 26 | 0 |  |  |
| `WEBSERVICE` | | absent | | | | | |
| `WEEKDAY` | date and time | implemented | 76 | 76 | 0 |  |  |
| `WEEKNUM` | date and time | implemented | 954 | 874 | 80 |  |  |
| `WEIBULL` | statistical | implemented | 70 | 70 | 0 | 4 ulps |  |
| `WEIBULL.DIST` | statistical | implemented | 383 | 371 | 12 | 4 ulps |  |
| `WORKDAY` | date and time | implemented | 152 | 152 | 0 |  |  |
| `WORKDAY.INTL` | date and time | implemented | 77 | 77 | 0 |  |  |
| `WRAPCOLS` | lookup and reference | implemented | 20 | 20 | 0 |  |  |
| `WRAPROWS` | lookup and reference | implemented | 35 | 35 | 0 |  |  |
| `XIRR` | financial | implemented | 707 | 667 | 40 |  |  |
| `XLOOKUP` | lookup and reference | partial | 329 | 328 | 1 |  | the regular-expression match mode is not implemented |
| `XMATCH` | lookup and reference | partial | 489 | 488 | 1 |  | the regular-expression match mode is not implemented |
| `XNPV` | financial | implemented | 32 | 32 | 0 |  |  |
| `XOR` | logical | implemented | 130 | 130 | 0 |  |  |
| `YEAR` | date and time | implemented | 38 | 38 | 0 |  |  |
| `YEARFRAC` | date and time | implemented | 995 | 995 | 0 |  |  |
| `YIELD` | financial | implemented | 124 | 123 | 1 | 1e-09 relative |  |
| `YIELDDISC` | financial | implemented | 381 | 381 | 0 | 4 ulps |  |
| `YIELDMAT` | | absent | | | | | |
| `Z.TEST` | statistical | implemented | 74 | 74 | 0 | 1e-12 relative |  |
| `ZTEST` | statistical | implemented | 73 | 73 | 0 | 1e-12 relative |  |
