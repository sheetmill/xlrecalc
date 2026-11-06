"""The capture suites: cases whose results Excel computes (``capture/pipeline.py``).

Every case is written for this project. New cases go at the end of their suite so earlier case
numbers stay stable.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from dataclasses import dataclass

from capture.pipeline import EMPTY_TEXT, Case, FormulaInput, Input
from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._values import ErrorValue

NA = ErrorValue.NA


@dataclass(frozen=True)
class Suite:
    cases: Callable[[], list[Case]]
    date1904: bool = False
    fixed: dict[str, Input] | None = None  # cells written once (lookup tables)


def _ulps(x: float, k: int) -> float:
    """``x`` moved ``k`` units in the last place (down for positive k)."""
    for _ in range(abs(k)):
        x = math.nextafter(x, -math.inf if k > 0 else math.inf)
    return x


# Number to text ----------------------------------------------------------------------------


def number_values() -> list[float]:
    values = [0.0, -0.0, 1.0, -1.0, 0.5, -0.5, 0.1 + 0.2, 1 / 3, 2 / 3, -2 / 3, 44927.5]
    values += [float(f"1e{e}") for e in range(-30, 31)]
    values += [float(f"1.23456789012345e{e}") for e in range(-25, 26)]
    values += [-float(f"1.23456789012345e{e}") for e in (-10, -5, 0, 5, 14, 15, 19, 20)]
    values += [float(f"1.2345678901234567e{e}") for e in (-20, -5, 0, 5, 14, 15, 16, 19, 20, 21)]
    values += [
        0.1234567890123455,
        0.1234567890123465,
        1234567890123455.0,
        999999999999999.5,
        9999999999999995.0,
        99999999999999.99,
        999999999999999.0,
        1e15 - 1,
        9.99999999999999e19,
        99999999999999990000.0,
        1.5e19,
        12345678901234567890.0,
        0.000001,
        0.0000001,
        1.5e-9,
        0.00012345678901234,
        0.000012345678901234,
        123.456,
        -123.456,
        5e-324,
        2.2250738585072014e-308,
        1.7976931348623157e308,
    ]
    return list(dict.fromkeys(values))


def number_text() -> list[Case]:
    cases = []
    for v in number_values():
        cases.append(Case("concat", '{1}&""', (v,)))
        cases.append(Case("text_general", 'TEXT({1},"General")', (v,)))
    for v in (True, False, None, EMPTY_TEXT, "text", NA):
        cases.append(Case("concat_other", '{1}&""', (v,)))
    cases += [
        Case("concat_formula", '0.1+0.2&""'),
        Case("concat_formula", '1/3&""'),
        Case("concat_formula", '(0.1+0.2-0.3)&""'),
        Case("concat_formula", 'DATE(2025,3,4)&""'),
    ]
    return cases


# Text to number ----------------------------------------------------------------------------

TEXTS = [
    "0", "1", "-1", "+1", "1.5", ".5", "1.", "1e3", "1E3", "1E+3", "1e-3", "-1.5E-3",
    "1,234", "1,234.5", "-1,234", "12,34", "1,2345", "1,,234", ",123", "123,", "1,234,567",
    "1 234", "  42  ", "\t42", "42\n", "0005", "-0", "00.50",
    "$5", "-$5", "$-5", "($5)", "(5)", "(-5)", "( 5 )", "$1,234.50", "$ 5",
    "5%", "50%", "-50%", "5 %", "%5", "$5%", "1e2%", "(5%)",
    "1/2", "1 1/2", "-1 1/2", "1 1/3", "0 1/2", "1/0",
    "3/4/2025", "3/4/25", "03/04/2025", "2025-03-04", "2025/3/4", "4-Mar-2025", "4-Mar-25",
    "Mar 4, 2025", "March 4, 2025", "4 March 2025", "March 4 2025", "3/4", "Mar-2025",
    "Mar 2025", "13/1/2025", "2/30/2025", "2/29/2024", "2/29/2023",
    "2/29/1900", "2/28/1900", "3/1/1900", "1/1/1900", "12/31/1899", "1/0/1900", "1/1/10000",
    "10:30", "10:30 AM", "10:30 PM", "10:30AM", "10:30 am", "10:30:15", "10:30:15.5",
    "10:30:15.123", "25:00", "24:00", "1:00:00 AM", "12:00 AM", "12:00 PM", "0:0", "10:61",
    "3/4/2025 10:30", "10:30 3/4/2025", "2025-03-04 10:30", "2025-03-04T10:30",
    "0x10", "1e309", "1e308", "123456789012345678", "0.12345678901234567",
    "12345678901234.567", "1234567890123456789012345",
    "TRUE", "FALSE", "true", " ", "abc", "1a", "a1", "1-2", "--1", "1d", "1f",
    "€5", "£5", "¥5", "Infinity", "NaN", "\u0661\u0662\u0663", "\uff11\uff12\uff13", "\u22125",
    "5-", "5+", "1.2.3", "1e", "e3",
]  # fmt: skip


def text_number() -> list[Case]:
    cases = []
    for s in TEXTS:
        cases.append(Case("value", "VALUE({1})", (s,)))
        cases.append(Case("implicit", "{1}+0", (s,)))
        cases.append(Case("datevalue", "DATEVALUE({1})", (s,)))
        cases.append(Case("timevalue", "TIMEVALUE({1})", (s,)))
    for v in (True, False, None, 1.5, NA, EMPTY_TEXT):
        cases.append(Case("value_other", "VALUE({1})", (v,)))
        cases.append(Case("implicit_other", "{1}+0", (v,)))
    return cases


# Comparison --------------------------------------------------------------------------------

KINDS: list[Input] = [1.0, 0.0, -1.0, "1", "a", "A", "b", EMPTY_TEXT, True, False, None, NA, "#N/A"]
ROUNDING_BASES = [0.3, 1 / 3, 123456.789, 3e-10, 1e20 / 3, 12345678901234.5, -0.7]


def compare() -> list[Case]:
    cases = []
    for a in KINDS:
        for b in KINDS:
            for op in ("=", "<>", "<", ">"):
                cases.append(Case(f"kinds{op}", f"{{1}}{op}{{2}}", (a, b)))
    for base in ROUNDING_BASES:
        others = [base * (1 + k * 1e-16) for k in range(1, 16)]
        others += [_ulps(base, k) for k in (-8, -4, -2, -1, 1, 2, 4, 8)]
        others += [float(f"{base:.15g}"), float(f"{base:.14g}"), float(f"{base:.16g}")]
        for other in dict.fromkeys(others):
            cases.append(Case("rounding=", "{1}={2}", (base, other)))
            cases.append(Case("rounding<", "{1}<{2}", (base, other)))
            cases.append(Case("rounding>", "{1}>{2}", (base, other)))
    cases += [
        Case("rounding_formula", "0.1+0.2=0.3"),
        Case("rounding_formula", "0.1+0.2-0.3=0"),
        Case("rounding_formula", "1/3*3=1"),
        Case("rounding_match", "MATCH(0.1+0.2,{1}:{1},0)", (0.3,)),
        Case("rounding_countif", "COUNTIF({1}:{1},0.1+0.2)", (0.3,)),
        Case("rounding_vlookup", "VLOOKUP(0.1+0.2,{1}:{1},1,FALSE)", (0.3,)),
    ]
    # The text "#N/A" and the error #N/A, in both orders in the range.
    for row in (("#N/A", NA, 1.0, "x"), (NA, "#N/A", 1.0, "x")):
        for formula in (
            'COUNTIF({1}:{4},"#N/A")',
            'COUNTIF({1}:{4},"=#N/A")',
            'COUNTIF({1}:{4},"<>#N/A")',
            'COUNTIF({1}:{4},"#n/a")',
            "COUNTIF({1}:{4},{1})",
            "COUNTIF({1}:{4},{2})",
            "COUNTIF({1}:{4},NA())",
            'MATCH("#N/A",{1}:{4},0)',
            "MATCH({1},{1}:{4},0)",
            "MATCH({2},{1}:{4},0)",
            "MATCH(NA(),{1}:{4},0)",
            'HLOOKUP("#N/A",{1}:{4},1,FALSE)',
            '_xlfn.XLOOKUP("#N/A",{1}:{4},{1}:{4},"none")',
            '{1}="#N/A"',
            '{2}="#N/A"',
            "ISNA({1})",
            "ISTEXT({1})",
            "ISERROR({2})",
            'SUMPRODUCT(--({1}:{4}="#N/A"))',
            'COUNTIF({1}:{4},"*")',
            'COUNTIF({1}:{4},"#*")',
        ):
            cases.append(Case("text_vs_error", formula, row))
    # Ranges holding only the text or only the error, to tell which one a criterion matches.
    for row in (("#N/A", 1.0), (NA, 1.0)):
        for formula in (
            'COUNTIF({1}:{2},"#N/A")',
            'COUNTIF({1}:{2},"<>#N/A")',
            "COUNTIF({1}:{2},NA())",
            'SUMIF({1}:{2},"#N/A",{2}:{2})',
            'MATCH("#N/A",{1}:{2},0)',
            'COUNTIF({1}:{2},"=")',
            'COUNTIF({1}:{2},"<>")',
        ):
            cases.append(Case("text_or_error", formula, row))
    for row in ((None, 1.0), (EMPTY_TEXT, 1.0), (" ", 1.0)):
        for formula in (
            'COUNTIF({1}:{2},"")',
            'COUNTIF({1}:{2},"=")',
            'COUNTIF({1}:{2},"<>")',
            "COUNTBLANK({1}:{2})",
            'MATCH("",{1}:{2},0)',
        ):
            cases.append(Case("blank_or_empty", formula, row))
    cases += [
        Case("case", 'MATCH("A",{1}:{2},0)', ("a", "A")),
        Case("case", 'COUNTIF({1}:{2},"A")', ("a", "A")),
        Case("case", "{1}={2}", ("straße", "STRASSE")),
        Case("case", "{1}<{2}", ("a", "B")),
        Case("case", "{1}<{2}", ("Z", "a")),
        Case("case", "{1}<{2}", ("a", "ä")),
        Case("case", "{1}<{2}", ("10", "9")),
    ]
    return cases


# Cancellation ------------------------------------------------------------------------------

CANCEL_BASES = [1.0, 0.1 + 0.2, 0.3, 123.456, 1e-5, 1e10, 1e15, 5e-300, -2.5]
CANCEL_ULPS = [0, 1, 2, 3, 4, 6, 8, 12, 16, 24, 32, 48, 64, 96, 128, 256, 512, 1024, 4096]


def cancellation() -> list[Case]:
    cases = []
    for a in CANCEL_BASES:
        for k in CANCEL_ULPS:
            b = _ulps(a, k)
            cases.append(Case("sub_top", "{1}-{2}", (a, b)))
            cases.append(Case("sub_nested", "({1}-{2})*1", (a, b)))
            cases.append(Case("add_top", "{1}+{2}", (a, -b)))
    cases += [
        Case("chain", "{1}+{2}-{3}", (0.1, 0.2, 0.3)),
        Case("chain", "({1}+{2})-{3}", (0.1, 0.2, 0.3)),
        Case("chain", "({1}+{2}-{3})", (0.1, 0.2, 0.3)),
        Case("chain", "-{3}+{1}+{2}", (0.1, 0.2, 0.3)),
        Case("chain", "{1}+{2}-{3}+0", (0.1, 0.2, 0.3)),
        Case("chain", "0+{1}+{2}-{3}", (0.1, 0.2, 0.3)),
        Case("chain", "SUM({1},{2},-{3})", (0.1, 0.2, 0.3)),
        Case("chain", "SUM({1}+{2}-{3})", (0.1, 0.2, 0.3)),
        Case("chain", "{1}+{2}-{3}=0", (0.1, 0.2, 0.3)),
        Case("chain", "IF(TRUE,{1}+{2}-{3})", (0.1, 0.2, 0.3)),
        Case("chain", "({1}+{2}-{3})*1", (0.1, 0.2, 0.3)),
        Case("chain", "{1}*3-{2}", (0.1, 0.30000000000000004)),
        Case("chain", "{1}-{2}", (0.30000000000000004, 0.3)),
    ]
    # The threshold: k ulps above and below, both operand orders, bases at both ends of a binade.
    for base in (1.0, 1.5, 1.9999999999999998, 0.3, 123.456, 1e15, 2.0**-1000):
        for k in range(5, 11):
            for b in (_ulps(base, k), _ulps(base, -k)):
                cases.append(Case("threshold", "{1}-{2}", (base, b)))
                cases.append(Case("threshold_reversed", "{2}-{1}", (base, b)))
    for x in (2.0**-1021, 2.0**-1022, 2.0**-1022 * 1.5):
        cases.append(Case("tiny", "{1}*1", (x,)))
        cases.append(Case("tiny", "{1}/2", (x,)))
        cases.append(Case("tiny", "{1}/4*4", (x,)))
    return cases


# Dates -------------------------------------------------------------------------------------

DATE_PARTS = [
    (1900, 1, 0), (1900, 1, 1), (1900, 2, 28), (1900, 2, 29), (1900, 3, 1), (1899, 12, 31),
    (1900, 0, 1), (1904, 1, 1), (1904, 1, 2), (2025, 3, 4), (9999, 12, 31), (10000, 1, 1),
]  # fmt: skip
SERIALS = [-1.0, 0.0, 1.0, 59.0, 60.0, 61.0, 1461.0, 2958465.0, 2958466.0, 60.5]


def dates() -> list[Case]:
    cases = [Case("date", "DATE({1},{2},{3})", tuple(map(float, p))) for p in DATE_PARTS]
    for s in SERIALS:
        for fn in ("YEAR", "MONTH", "DAY", "WEEKDAY"):
            cases.append(Case(fn.lower(), f"{fn}({{1}})", (s,)))
    for text in ("2/29/1900", "1/1/1900", "1/1/1904", "1/2/1904", "12/31/1903"):
        cases.append(Case("datevalue", "DATEVALUE({1})", (text,)))
    return cases


# Summation order and precision ---------------------------------------------------------------

SUMS: list[tuple[float, ...]] = [
    (0.1,) * 10,
    (1e16, 1.0, -1e16),
    (1.0, 1e100, 1.0, -1e100),
    (0.1, 0.2, 0.3),
    (1e16, 1.0, 1.0, 1.0, 1.0),
    (0.7, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1),
    (123456.789, 0.001, 0.001, 0.001, 0.001, 0.001, 0.001, 0.001, 0.001, 0.001),
    (3.0, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, -3.7),
]


def sums() -> list[Case]:
    cases = []
    for values in SUMS:
        n = len(values)
        last = f"{{{n}}}"
        cases.append(Case("sum_range", f"SUM({{1}}:{last})", values))
        cases.append(
            Case("sum_args", "SUM(" + ",".join(f"{{{k}}}" for k in range(1, n + 1)) + ")", values)
        )
        cases.append(
            Case("plus_chain", "+".join(f"{{{k}}}" for k in range(1, n + 1)) + "+0", values)
        )
        cases.append(Case("average", f"AVERAGE({{1}}:{last})", values))
    return cases


# Functions -------------------------------------------------------------------------------------

KINDS_ARG: list[Input] = [2.5, -2.5, 0.0, "3", "abc", True, False, None, EMPTY_TEXT, NA]
NUMBERS = [0.0, 1.0, -1.0, 2.5, -2.5, 0.5, 1e-10, 123456.789, 2.675, -2.675, 1.005, 1e15 + 0.5]
MIXED: tuple[Input, ...] = (1.0, 2.0, "3", True, None, "abc", 4.5, EMPTY_TEXT, 0.0, -1.5)
CRITERIA: list[Input] = [
    ">2", "<=2", "<>1", "1", 1.0, "3", "a*", "?bc", "ABC", "", "=", "<>", True, "TRUE",
    ">=abc", "<a", "~*", 4.5, "4.5", "-1.5",
]  # fmt: skip
TABLE = {
    "Z1": "a", "AA1": 1.0, "Z2": "b", "AA2": 2.0, "Z3": "c", "AA3": 3.0,
    "Z4": 1.0, "AA4": "one", "Z5": 2.0, "AA5": "two", "Z6": "B", "AA6": "upper b",
    "AC1": 1.0, "AD1": "x1", "AC2": 3.0, "AD2": "x3", "AC3": 5.0, "AD3": "x5",
    "AC4": 7.0, "AD4": "x7", "AC5": 9.0, "AD5": "x9",
    "AF1": 9.0, "AF2": 7.0, "AF3": 5.0, "AF4": 3.0, "AF5": 1.0,
}  # fmt: skip


def functions() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for fn in ("ABS", "INT", "SQRT", "LN", "EXP", "ISNUMBER", "ISTEXT", "ISBLANK", "ISERROR",
               "ISERR", "ISNA", "ISLOGICAL", "NOT", "LEN", "TRIM", "UPPER", "LOWER", "VALUE",
               "YEAR", "MONTH", "DAY", "WEEKDAY"):  # fmt: skip
        for v in [*KINDS_ARG, *NUMBERS[3:8], 45720.75]:
            add(Case(fn.lower(), f"{fn}({{1}})", (v,)))
    for fn in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
        for x in NUMBERS:
            for d in (0.0, 1.0, 2.0, -1.0, -2.0):
                add(Case(fn.lower(), f"{fn}({{1}},{{2}})", (x, d)))
        add(Case(fn.lower(), f"{fn}({{1}},{{2}})", ("2.5", True)))
        add(Case(fn.lower(), f"{fn}({{1}},{{2}})", (NA, 1.0)))
    for a, b in ((10.0, 3.0), (-10.0, 3.0), (10.0, -3.0), (5.5, 2.0), (1.0, 0.0), (1e10, 7.0),
                 (0.0, 0.0), (2.0, 0.5), (-8.0, 1.0 / 3), (0.0, -1.0), (10.0, 400.0)):  # fmt: skip
        add(Case("mod", "MOD({1},{2})", (a, b)))
        add(Case("power", "POWER({1},{2})", (a, b)))
    add(Case("pi", "PI()"))
    for fn in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "PRODUCT", "COUNTBLANK"):
        add(Case(fn.lower(), f"{fn}({{1}}:{{10}})", MIXED))
        add(Case(fn.lower(), f"{fn}({{5}}:{{5}})", MIXED))
        if fn != "COUNTBLANK":
            add(Case(fn.lower(), f'{fn}({{1}},"3",TRUE)', MIXED))
            add(Case(fn.lower(), f'{fn}({{1}},"abc")', MIXED))
            add(Case(fn.lower(), f"{fn}({{1}}:{{3}},{{11}})", (*MIXED, NA)))
            add(Case(fn.lower(), f"{fn}({{4}},{{5}},{{8}})", MIXED))
    for c in CRITERIA:
        add(Case("countif", "COUNTIF({1}:{10},{11})", (*MIXED, c)))
        add(Case("sumif", "SUMIF({1}:{10},{11})", (*MIXED, c)))
        add(Case("sumif", "SUMIF({1}:{10},{11},{1}:{10})", (*MIXED, c)))
    add(Case("countifs", 'COUNTIFS({1}:{10},">0",{1}:{10},"<3")', MIXED))
    add(Case("sumifs", 'SUMIFS({1}:{10},{1}:{10},">0",{1}:{10},"<>2")', MIXED))
    add(Case("sumifs", 'SUMIFS({1}:{10},{1}:{9},">0")', MIXED))
    for cond in [*KINDS_ARG, "TRUE", "false", 1.0]:
        add(Case("if", 'IF({1},"yes","no")', (cond,)))
    add(Case("if", "IF(FALSE,1)"))
    add(Case("if", "IF(TRUE,,2)"))
    add(Case("if", "IF(FALSE,1,)"))
    add(Case("if", "IF(TRUE,{1})", (None,)))
    for fn in ("AND", "OR"):
        for args in ((True, True), (True, False), (1.0, 0.0), ("TRUE", True), ("x", True),
                     (None, True), (NA, True), (2.5, "abc")):  # fmt: skip
            add(Case(fn.lower(), f"{fn}({{1}},{{2}})", args))
            add(Case(fn.lower(), f"{fn}({{1}}:{{2}})", args))
    for v in KINDS_ARG:
        add(Case("iferror", 'IFERROR({1},"alt")', (v,)))
        add(Case("ifna", '_xlfn.IFNA({1},"alt")', (v,)))
    add(Case("iferror", 'IFERROR(1/0,"alt")'))
    add(Case("ifna", '_xlfn.IFNA(1/0,"alt")'))
    for v in KINDS_ARG:
        add(Case("concatenate", 'CONCATENATE({1},"|",{1})', (v,)))
    for text, n in (("hello", 2.0), ("hello", 0.0), ("hello", 10.0), ("hello", -1.0),
                    ("", 1.0), (12345.0, 2.0), (True, 2.0)):  # fmt: skip
        add(Case("left", "LEFT({1},{2})", (text, n)))
        add(Case("right", "RIGHT({1},{2})", (text, n)))
        add(Case("mid", "MID({1},2,{2})", (text, n)))
    add(Case("left", "LEFT({1})", ("hello",)))
    add(Case("mid", "MID({1},0,2)", ("hello",)))
    add(Case("mid", "MID({1},9,2)", ("hello",)))
    for needle, hay, start in (("l", "hello", 1.0), ("L", "hello", 1.0), ("l", "hello", 4.0),
                               ("", "hello", 3.0), ("z", "hello", 1.0), ("l", "hello", 9.0),
                               ("l", "hello", 0.0)):  # fmt: skip
        add(Case("find", "FIND({1},{2},{3})", (needle, hay, start)))
    add(Case("trim", "TRIM({1})", ("  a   b  ",)))
    for v in (1234.5, -1234.5, 0.5, 45720.75, "abc", True, None):
        for fmt in ("0.00", "#,##0", "0%", "yyyy-mm-dd", "General", "@", "0.0E+00", "mmm d"):
            add(Case("text", f'TEXT({{1}},"{fmt}")', (v,)))
    for y, m, d in ((2025.0, 3.0, 4.0), (2025.0, 14.0, 1.0), (2025.0, 0.0, 0.0),
                    (2024.0, 2.0, 30.0), (99.0, 1.0, 1.0), (2025.5, 3.7, 4.9),
                    (-1.0, 1.0, 1.0)):  # fmt: skip
        add(Case("date", "DATE({1},{2},{3})", (y, m, d)))
    for s in (45720.0, 45351.0, 1.0, 59.0, 60.0, 61.0):
        for k in (1.0, 2.0, 3.0, 11.0, 17.0, 4.0):
            add(Case("weekday", "WEEKDAY({1},{2})", (s, k)))
        for months in (0.0, 1.0, -1.0, 12.0, 1.5, -13.0):
            add(Case("eomonth", "EOMONTH({1},{2})", (s, months)))
            add(Case("edate", "EDATE({1},{2})", (s, months)))
    add(Case("edate", "EDATE(DATE(2024,1,31),1)"))
    add(Case("edate", "EDATE(DATE(1900,1,29),1)"))
    add(Case("eomonth", "EOMONTH(DATE(1900,1,15),1)"))
    # Lookups against the fixed tables: Z1:AA6 text and numbers, AC1:AD5 sorted ascending,
    # AF1:AF5 descending.
    for v in ("a", "B", "c", 1.0, 2.0, "zz", "b*", "?", "~*", True, None, NA, 4.0):
        add(Case("vlookup", "VLOOKUP({1},$Z$1:$AA$6,2,FALSE)", (v,)))
        add(Case("vlookup", "VLOOKUP({1},$Z$1:$AA$6,2,)", (v,)))
        add(Case("match", "MATCH({1},$Z$1:$Z$6,0)", (v,)))
        add(Case("hlookup", "HLOOKUP({1},{2}:{4},1,FALSE)", (v, "a", 1.0, "b")))
    for v in (0.0, 1.0, 2.0, 3.0, 6.0, 9.0, 10.0, "5", 5.5):
        add(Case("vlookup", "VLOOKUP({1},$AC$1:$AD$5,2)", (v,)))
        add(Case("vlookup", "VLOOKUP({1},$AC$1:$AD$5,2,TRUE)", (v,)))
        add(Case("match", "MATCH({1},$AC$1:$AC$5)", (v,)))
        add(Case("match", "MATCH({1},$AC$1:$AC$5,1)", (v,)))
        add(Case("match", "MATCH({1},$AF$1:$AF$5,-1)", (v,)))
    add(Case("vlookup", "VLOOKUP(1,$AC$1:$AD$5,3,FALSE)"))
    add(Case("vlookup", "VLOOKUP(1,$AC$1:$AD$5,0,FALSE)"))
    for r, c in ((1.0, 1.0), (2.0, 2.0), (6.0, 1.0), (7.0, 1.0), (0.0, 1.0), (1.0, 3.0)):
        add(Case("index", "INDEX($Z$1:$AA$6,{1},{2})", (r, c)))
    add(Case("index", "INDEX($Z$1:$Z$6,3)"))
    add(Case("index", "INDEX({1}:{4},3)", ("a", 1.0, "b", 2.0)))
    for i in (1.0, 2.0, 3.0, 0.0, 4.0, 2.9):
        add(Case("choose", 'CHOOSE({1},"x","y","z")', (i,)))
    add(Case("rows", "ROWS($Z$1:$AA$6)"))
    add(Case("columns", "COLUMNS($Z$1:$AA$6)"))
    add(Case("row", "ROW($Z$4)"))
    add(Case("column", "COLUMN($AA$4)"))
    add(Case("row", "ROW()"))
    add(Case("column", "COLUMN()"))
    add(Case("sumproduct", "SUMPRODUCT($AC$1:$AC$5,$AC$1:$AC$5)"))
    add(Case("sumproduct", "SUMPRODUCT({1}:{10},{1}:{10})", MIXED))
    add(Case("sumproduct", "SUMPRODUCT($AC$1:$AC$5,$Z$1:$Z$5)"))
    add(Case("sumproduct", "SUMPRODUCT($AC$1:$AC$5,$AC$1:$AC$4)"))
    for code in (1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 109.0, 12.0):
        add(Case("subtotal", "SUBTOTAL({11},{1}:{10})", (*MIXED, code)))
    add(Case("subtotal", "SUBTOTAL(9,{1}:{2})", (1.0, FormulaInput("=SUBTOTAL(9,C1:C1)"))))
    add(Case("na", "NA()"))
    add(Case("true", "TRUE()"))
    # The SUMIF family over blanks, errors and the sheet's edge: which cells a walk may skip.
    for c in ("", "=", "<>", "<>1", "<>abc", "<=2", ">0"):
        add(Case("countifs", "COUNTIFS({1}:{10},{11})", (*MIXED, c)))
        add(Case("countifs", "COUNTIFS({1}:{10},{11},{1}:{10},{11})", (*MIXED, c)))
        add(Case("countifs", 'COUNTIFS({1}:{10},{11},{1}:{10},"<>2")', (*MIXED, c)))
        add(Case("countifs", 'COUNTIFS({1}:{9},{11},{2}:{10},">0")', (*MIXED, c)))
        add(Case("sumifs", "SUMIFS({1}:{10},{1}:{10},{11})", (*MIXED, c)))
        add(Case("sumifs", "SUMIFS({2}:{10},{1}:{9},{11})", (*MIXED, c)))
        add(Case("sumifs", 'SUMIFS({2}:{10},{1}:{9},{11},{2}:{10},"<>2")', (*MIXED, c)))
        add(Case("sumif", "SUMIF({1}:{9},{11},{2}:{10})", (*MIXED, c)))
    errors: tuple[Input, ...] = (1.0, NA, None, FormulaInput("=1/0"), 2.0)
    for c in ("<>1", ">0", "", "#N/A", "<>"):
        add(Case("sumif", "SUMIF({1}:{5},{6})", (*errors, c)))
        add(Case("sumif", "SUMIF({1}:{5},{6},{1}:{5})", (*errors, c)))
        add(Case("sumifs", "SUMIFS({1}:{5},{1}:{5},{6})", (*errors, c)))
        add(Case("countifs", "COUNTIFS({1}:{5},{6})", (*errors, c)))
    add(Case("sumif", 'SUMIF($Z$1:$AA$6,">0",$XFD$1)'))
    add(Case("sumif", 'SUMIF($Z$1:$AA$6,"<>z",$XFD$1)'))
    add(Case("sumif", 'SUMIF($AC$1:$AD$5,"<>",$XFD$1048576)'))
    add(Case("sumif", 'SUMIF($AC$1:$AD$5,">0",$AD$1048576)'))
    # An error in the key and in a number argument: which one a lookup gives.
    add(Case("vlookup", "VLOOKUP(1/0,$AC$1:$AD$5,NA())"))
    add(Case("vlookup", "VLOOKUP(1,$AC$1:$AD$5,NA())"))
    add(Case("hlookup", "HLOOKUP(1/0,$AC$1:$AD$5,NA())"))
    add(Case("hlookup", "HLOOKUP(1,$AC$1:$AD$5,NA())"))
    add(Case("match", "MATCH(1/0,$AC$1:$AC$5,NA())"))
    add(Case("match", "MATCH(1,$AC$1:$AC$5,NA())"))
    return cases


# Number literals in formulas ------------------------------------------------------------------

LITERALS = [
    "0.1234567890123456",
    "0.12345678901234567",
    "0.123456789012345678",
    "0.99999999999999999",
    "0.9999999999999999",
    "1.0000000000000009",
    "1.00000000000000000001",
    "123456789012345678",
    "12345678901234567890",
    "9007199254740993",
    "1234567890123456.7",
    "2.6749999999999999",
    "1.234567890123456789E+20",
    "1.2345678901234567E-300",
    "0.000000000000000012345678901234567",
    "3.14159265358979323846",
    "0.123456789012345",
    "0.30000000000000004",
]
LITERAL_CONTEXTS = ["{x}", "-{x}", "{x}*1", "SUM({x})", "SUM({{{x},0}})", "IF(TRUE,{x})", '{x}&""']


def literals() -> list[Case]:
    return [
        Case(f"literal:{context}", context.format(x=x))
        for context in LITERAL_CONTEXTS
        for x in LITERALS
    ]


# N and LOOKUP -----------------------------------------------------------------------------------


def n_lookup() -> list[Case]:
    cases = [Case("n", "N({1})", (v,)) for v in KINDS_ARG]
    cases += [
        Case("n", "N(AA1:AA3)"),
        Case("n", "N(Z1:Z3)"),
        Case("n", "N({1,2})"),
        Case("n", "SUM(N({1,2}))"),
        Case("n", "N(1/0)"),
        Case("n", "N(DATE(2024,1,2))"),
    ]
    keys = [0.0, 1.0, 2.0, 4.0, 5.0, 9.0, 100.0, "3", True, "a", None, NA]
    cases += [Case("lookup", "LOOKUP({1},AC1:AC5,AD1:AD5)", (k,)) for k in keys]
    cases += [Case("lookup", "LOOKUP({1},AC1:AC5)", (k,)) for k in keys]
    cases += [
        Case("lookup", 'LOOKUP("b",Z1:Z3,AA1:AA3)'),
        Case("lookup", 'LOOKUP("bb",Z1:Z3,AA1:AA3)'),
        Case("lookup", 'LOOKUP("zz",Z1:Z3)'),
        Case("lookup", 'LOOKUP("A",Z1:Z3,AA1:AA3)'),
        Case("lookup", 'LOOKUP("0",Z1:Z3,AA1:AA3)'),
        Case("lookup", "LOOKUP(4,AC1:AD5)"),
        Case("lookup", "LOOKUP(2,{1,2,3;4,5,6})"),
        Case("lookup", "LOOKUP(2,{1,2;3,4})"),
        Case("lookup", 'LOOKUP(4,{1,3,5},{"a","b","c"})'),
        Case("lookup", "LOOKUP(7,AC1:AC5,AD1:AD3)"),
        Case("lookup", "LOOKUP(4,AC1:AC5,AE1:AI1)"),
        Case("lookup", "LOOKUP(4,AC1:AC5,{1,2,3,4,5})"),
        Case("lookup", "LOOKUP(1,Z4:Z6)"),
        Case("lookup", "LOOKUP(2,Z4:AA6)"),
        Case("lookup", "LOOKUP(4,AC1:AC5,AD1)"),
        Case("lookup", "LOOKUP(4,AC1)"),
    ]
    return cases


# Financial functions: annuities, cash flows, depreciation -------------------------------------

FIN_TABLE: dict[str, Input] = {
    # cash flows and their dates
    "AK1": -10000.0, "AK2": 3000.0, "AK3": 4200.0, "AK4": 6800.0, "AK5": 1500.0,
    "AL1": 39448.0, "AL2": 39508.0, "AL3": 39751.0, "AL4": 39859.0, "AL5": 39904.0,
    # flows with text, a logical, a blank and an empty text among them
    "AM1": -100.0, "AM2": "abc", "AM3": True, "AM5": 50.0, "AM6": EMPTY_TEXT, "AM7": 30.0,
    # dates out of order, with fractions, and as text
    "AN1": 39448.0, "AN2": 39400.0, "AN3": 39751.0, "AN4": 39859.0, "AN5": 39904.0,
    "AO1": 39448.75, "AO2": 39508.5, "AO3": 39751.25, "AO4": 39859.99, "AO5": 39904.1,
    "AP1": 39448.0, "AP2": "2/1/2008", "AP3": 39751.0, "AP4": 39859.0, "AP5": 39904.0,
    # three flows (a size mismatch), an error among flows, all positive, a date before 1900
    "AQ1": -500.0, "AQ2": 200.0, "AQ3": 400.0,
    "AR1": -500.0, "AR2": NA, "AR3": 400.0, "AR4": 100.0, "AR5": 100.0,
    "AS1": 500.0, "AS2": 200.0, "AS3": 400.0, "AS4": 100.0, "AS5": 100.0,
    "AT1": 39448.0, "AT2": -1.0, "AT3": 39751.0, "AT4": 39859.0, "AT5": 39904.0,
}  # fmt: skip

FIN_RATES = [0.05, 0.1 / 12, 0.0, -0.5, -1.0, -1.5, 1e-10, 2.5, 1e-300]
# (nper, pv or pmt, fv, type); None is an omitted argument
FIN_TERMS: list[tuple[float, float, float | None, float | None]] = [
    (10.0, 1000.0, None, None),
    (360.0, -200000.0, 0.0, 0.0),
    (12.0, -500.0, 5000.0, 1.0),
    (2.5, 100.0, -50.0, 0.0),
    (0.0, 1000.0, 0.0, 0.0),
    (-5.0, 1000.0, 0.0, 1.0),
    (10.0, 1000.0, 0.0, 2.0),
    (10.0, 1000.0, 0.0, 0.5),
    (10.0, 1000.0, 0.0, -1.0),
    (1000.0, 1.0, 0.0, 0.0),
]


def _call(name: str, args: list[str | None]) -> str:
    while args and args[-1] is None:
        args.pop()
    return f"{name}({','.join('' if a is None else a for a in args)})"


def finance() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for fn in ("PV", "FV", "PMT"):
        for rate in FIN_RATES:
            for nper, x, fv, kind in FIN_TERMS:
                args = [
                    "{1}",
                    "{2}",
                    "{3}",
                    None if fv is None else "{4}",
                    None if kind is None else "{5}",
                ]
                inputs = tuple(v for v in (rate, nper, x, fv, kind) if v is not None)
                add(Case(fn.lower(), _call(fn, args), inputs))
    for rate in FIN_RATES:
        for pmt, pv, fv, kind in [
            (-100.0, 1000.0, None, None), (-100.0, 1000.0, 0.0, 1.0), (100.0, 1000.0, 0.0, 0.0),
            (-1000.0, 1000.0, 0.0, 0.0), (0.0, 1000.0, -2000.0, 0.0), (0.0, 1000.0, 0.0, 0.0),
            (-50.0, -1000.0, 5000.0, 0.0), (-100.0, 1000.0, 0.0, 3.0), (-1e-3, 1000.0, 0.0, 0.0),
        ]:  # fmt: skip
            args = [
                "{1}",
                "{2}",
                "{3}",
                None if fv is None else "{4}",
                None if kind is None else "{5}",
            ]
            inputs = tuple(v for v in (rate, pmt, pv, fv, kind) if v is not None)
            add(Case("nper", _call("NPER", args), inputs))
    for fn in ("IPMT", "PPMT"):
        for rate in (0.05, 0.1 / 12, 0.0, -0.5, -1.0, 2.5):
            for per, nper, pv, fv, kind in [
                (1.0, 10.0, 1000.0, None, None), (1.0, 10.0, 1000.0, 0.0, 1.0),
                (2.0, 10.0, 1000.0, 0.0, 1.0), (10.0, 10.0, 1000.0, 500.0, 0.0),
                (5.5, 10.0, 1000.0, 0.0, 0.0), (0.0, 10.0, 1000.0, 0.0, 0.0),
                (11.0, 10.0, 1000.0, 0.0, 0.0), (10.5, 10.0, 1000.0, 0.0, 0.0),
                (0.5, 10.0, 1000.0, 0.0, 0.0), (180.0, 360.0, -200000.0, 0.0, 0.0),
                (1.0, 2.5, 100.0, -50.0, 0.0), (3.0, 10.0, 1000.0, 0.0, 2.0),
            ]:  # fmt: skip
                args = [
                    "{1}",
                    "{2}",
                    "{3}",
                    "{4}",
                    None if fv is None else "{5}",
                    None if kind is None else "{6}",
                ]
                inputs = tuple(v for v in (rate, per, nper, pv, fv, kind) if v is not None)
                add(Case(fn.lower(), _call(fn, args), inputs))
    # argument kinds, one position at a time
    base = {"PV": (0.05, 10.0, 100.0), "FV": (0.05, 10.0, 100.0), "PMT": (0.05, 10.0, 1000.0),
            "NPER": (0.05, -100.0, 1000.0), "SLN": (1000.0, 100.0, 5.0)}  # fmt: skip
    for fn, args3 in base.items():
        for pos in range(3):
            for v in KINDS_ARG:
                inputs = tuple(v if i == pos else a for i, a in enumerate(args3))
                add(Case(fn.lower(), f"{fn}({{1}},{{2}},{{3}})", inputs))
    for v in KINDS_ARG:
        add(Case("pmt", "PMT(0.05,10,1000,0,{1})", (v,)))
        add(Case("ipmt", "IPMT(0.05,{1},10,1000)", (v,)))
    add(Case("pmt", "PMT({0.05,0.1},10,1000)"))
    add(Case("pmt", "SUM(PMT({0.05,0.1},10,1000))"))
    add(Case("pv", "PV(0.05,10,,1000)"))
    add(Case("pv", "PV(0.05,10,100,,1)"))
    # NPV
    for rate in (0.1, 0.0, -0.5, -1.0, -1.5, 1e-10, 10.0):
        add(Case("npv", "NPV({1},AK1:AK5)", (rate,)))
        add(Case("npv", "NPV({1},AK1,AK2,AK3:AK5)", (rate,)))
        add(Case("npv", "NPV({1},-10000,3000,4200,6800,1500)", (rate,)))
    for v in KINDS_ARG:
        add(Case("npv", "NPV(0.1,{1},100)", (v,)))
        add(Case("npv", "NPV({1},100,200)", (v,)))
    cases += [
        Case("npv", "NPV(0.1,AM1:AM7)"),
        Case("npv", "NPV(0.1,AR1:AR5)"),
        Case("npv", "NPV(0.1,AU1:AU5)"),
        Case("npv", 'NPV(0.1,"5",TRUE,100)'),
        Case("npv", "NPV(0.1,{100,200;300,400})"),
        Case("npv", 'NPV(0.1,{100,"a",TRUE,200})'),
        Case("npv", "NPV(0.1,100,,200)"),
        Case("npv", "NPV(0.1,AK1:AK5,AQ1:AQ3)"),
        Case("npv", "NPV(AK1:AK2,100)"),
        Case("npv", "NPV(0.05,1E+300,1E+300)"),
        Case("npv", "NPV(0.05,AK1:AK5)+AK1"),
    ]
    # XNPV
    for rate in (0.09, 0.0, -0.5, -0.99, -1.0, -1.5, 1e-10, 5.0):
        add(Case("xnpv", "XNPV({1},AK1:AK5,AL1:AL5)", (rate,)))
    cases += [
        Case("xnpv", "XNPV(0.09,AK1:AK5,AN1:AN5)"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AO1:AO5)"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AP1:AP5)"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AT1:AT5)"),
        Case("xnpv", "XNPV(0.09,AQ1:AQ3,AL1:AL5)"),
        Case("xnpv", "XNPV(0.09,AR1:AR5,AL1:AL5)"),
        Case("xnpv", "XNPV(0.09,AS1:AS5,AL1:AL5)"),
        Case("xnpv", "XNPV(0.09,AM1:AM5,AL1:AL5)"),
        Case("xnpv", "XNPV(0.09,AK1,AL1)"),
        Case("xnpv", "XNPV(0.09,AK1:AK2,AL1:AL2)"),
        Case("xnpv", "XNPV(0.09,{-1000,500,700},{39448,39600,39800})"),
        Case("xnpv", "XNPV(0.09,{-1000,500,700},{39448,39300,39800})"),
        Case("xnpv", 'XNPV("0.09",AK1:AK5,AL1:AL5)'),
        Case("xnpv", 'XNPV("abc",AK1:AK5,AL1:AL5)'),
        Case("xnpv", "XNPV(TRUE,AK1:AK5,AL1:AL5)"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AL1:AL5)+0"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,{39448,39508,39751,39859,39904})"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AL1:AL4)"),
        Case("xnpv", "XNPV(0.09,AK1:AK5,AL1:AP1)"),
    ]
    # SLN and VDB
    lives = [(10000.0, 1000.0, 5.0), (1000.0, 0.0, 0.5), (1000.0, 2000.0, 3.0),
             (1000.0, 100.0, 0.0), (-1000.0, 100.0, 3.0), (1000.0, 100.0, -3.0)]  # fmt: skip
    for cost, salvage, life in lives:
        add(Case("sln", "SLN({1},{2},{3})", (cost, salvage, life)))
    for cost, salvage, life in [(2400.0, 300.0, 10.0), (10000.0, 1000.0, 5.0), (1000.0, 0.0, 3.0)]:
        spans = [(0.0, 1.0), (0.0, 0.5), (1.0, 2.0), (0.5, 1.5), (0.0, life), (2.0, 4.5),
                 (life - 1, life), (3.0, 3.0), (2.0, 1.0), (0.0, life + 1),
                 (-1.0, 1.0)]  # fmt: skip
        for start, end in spans:
            for factor, switch in [
                (None, None),
                (1.5, None),
                (2.0, True),
                (3.0, False),
                (1.0, None),
            ]:
                args = ["{1}", "{2}", "{3}", "{4}", "{5}",
                        None if factor is None and switch is None else "{6}",
                        None if switch is None else "{7}"]  # fmt: skip
                vals: list[Input] = [cost, salvage, life, start, end]
                if factor is not None or switch is not None:
                    vals.append(2.0 if factor is None else factor)
                if switch is not None:
                    vals.append(switch)
                add(Case("vdb", _call("VDB", args), tuple(vals)))
    cases += [
        Case("vdb", "VDB(1000,2000,5,0,1)"),
        Case("vdb", "VDB(1000,100,0,0,1)"),
        Case("vdb", "VDB(1000,100,5,0,1,0)"),
        Case("vdb", "VDB(1000,100,5,0,1,-1)"),
        Case("vdb", "VDB(-1000,100,5,0,1)"),
        Case("vdb", "VDB(1000,-100,5,0,1)"),
        Case("vdb", "VDB(1000,100,2.5,0,2.5)"),
        Case("vdb", "VDB(1000,100,5.5,1.25,3.75,2.5)"),
        Case("vdb", 'VDB(1000,100,5,0,1,2,"x")'),
        Case("vdb", "VDB(1000,100,5,0,1,2,2)"),
        Case("vdb", "VDB(1000,0,1,0,1)"),
        Case("vdb", "VDB(1000,100,5,1,1.0001)"),
    ]
    return cases


# Iterative financial functions: IRR, XIRR, RATE --------------------------------------------------

_FLOWS: dict[str, list[Input]] = {
    "BA": [-10000.0, 3000.0, 4200.0, 6800.0, 1500.0],
    "BB": [-70000.0, 12000.0, 15000.0, 18000.0, 21000.0, 26000.0],
    "BC": [-100.0, 230.0, -132.0],  # two rates: 10% and 20%
    "BD": [100.0, 200.0, 300.0],  # no sign change
    "BE": [-1000.0, 0.0, 0.0, 0.0, 1.0],
    "BF": [-1.0, 1e10],
    "BG": [-1000.0, 999.0],
    "BH": [-1000.0, 1100.0],
    "BI": [-100.0] + [0.0] * 10 + [200.0],
    "BJ": [1000.0, -500.0, -600.0],
    "BK": [-100.0, "abc", True, None, 50.0, 70.0],
    "BL": [-100.0, NA, 120.0],
    "BM": [-5000.0] + [1000.0] * 10,
    "BN": [-1000.0, 300.0, 300.0, 300.0, 300.0],
    "BO": [-1e-10, 1.5e-10],
    "BP": [-1000.0, 2000.0, -1000.0],  # a double root at 0
    "BQ": [-100.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 110.0],
    "BR": [-1000.0, 5000.0, -6000.0, 2000.0],
}
_DATES: dict[str, list[Input]] = {
    "BU": [39448.0, 39508.0, 39751.0, 39859.0, 39904.0],
    "BV": [39448.0, 39813.0, 40178.0, 40543.0, 40909.0, 41274.0],
    "BW": [39448.0, 39449.0, 39450.0],
    "BX": [39448.0, 39448.0, 39813.0],
    "BY": [39448.0, 39813.0, 39500.0, 40178.0, 40543.0],
    "BZ": [39448.0, 43000.0],
    "CA": [39448.0, 39448.5, 39449.7],
}
ITER_TABLE: dict[str, Input] = {
    f"{col}{i}": v
    for table in (_FLOWS, _DATES)
    for col, values in table.items()
    for i, v in enumerate(values, 1)
    if v is not None
}


def _rng(col: str, table: dict[str, list[Input]]) -> str:
    return f"{col}1:{col}{len(table[col])}"


IRR_GUESSES = [0.1, 0.0, -0.5, 0.5, 2.0, -0.99, -1.0, 10.0, 1e-10, -0.9, 0.15]


def iterative() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for col in _FLOWS:
        add(Case("irr", f"IRR({_rng(col, _FLOWS)})"))
    for col in ("BA", "BC", "BE", "BJ", "BP", "BR", "BF", "BG"):
        for g in IRR_GUESSES:
            add(Case("irr", f"IRR({_rng(col, _FLOWS)},{{1}})", (g,)))
    for v in KINDS_ARG:
        add(Case("irr", "IRR(BA1:BA5,{1})", (v,)))
    cases += [
        Case("irr", "IRR({-100,60,60})"),
        Case("irr", "IRR({-100;60;60})"),
        Case("irr", "IRR({-100,60;60,10})"),
        Case("irr", 'IRR({-100,"a",TRUE,60,60})'),
        Case("irr", "IRR(BA1)"),
        Case("irr", "IRR(-100)"),
        Case("irr", "IRR((BA1:BA2,BA3:BA5))"),
        Case("irr", "IRR(BA1:BB5)"),
        Case("irr", "IRR(BA1:BA5)*1"),
        Case("irr", "IRR(BA1:BA5,)"),
        Case("irr", "IRR({-1000,100,100,100,100,100,100,100,100,100,100,100,100})"),
        Case("irr", "IRR({-1000,1,1,1,1})"),
        Case("irr", "IRR({-1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000})"),
    ]
    pairs = [("BA", "BU"), ("BB", "BV"), ("BC", "BW"), ("BC", "BX"), ("BA", "BY"), ("BF", "BZ"),
             ("BC", "CA"), ("BG", "BZ"), ("BH", "BZ"), ("BD", "BW"), ("BJ", "BW")]  # fmt: skip
    for f, d in pairs:
        n = min(len(_FLOWS[f]), len(_DATES[d]))
        add(Case("xirr", f"XIRR({f}1:{f}{n},{d}1:{d}{n})"))
    for g in IRR_GUESSES:
        add(Case("xirr", "XIRR(BA1:BA5,BU1:BU5,{1})", (g,)))
        add(Case("xirr", "XIRR(BC1:BC3,BW1:BW3,{1})", (g,)))
    for v in KINDS_ARG:
        add(Case("xirr", "XIRR(BA1:BA5,BU1:BU5,{1})", (v,)))
    cases += [
        Case("xirr", "XIRR(BA1:BA5,BU1:BU4)"),
        Case("xirr", "XIRR(BK1:BK5,BU1:BU5)"),
        Case("xirr", "XIRR(BL1:BL3,BU1:BU3)"),
        Case("xirr", "XIRR(AK1:AK5,AP1:AP5)"),
        Case("xirr", "XIRR(AK1:AK5,AT1:AT5)"),
        Case("xirr", "XIRR(AK1:AK5,AO1:AO5)"),
        Case("xirr", "XIRR({-1000,600,600},{39448,39813,40178})"),
        Case("xirr", "XIRR({-1000,600,600},{39448,39300,40178})"),
        Case("xirr", "XIRR(BA1:BA5,BU1:BU5)*1"),
        Case("xirr", "XIRR(BN1:BN5,BU1:BU5)"),
        Case("xirr", "XIRR(BA1,BU1)"),
        Case("xirr", "XIRR(BA1:BA2,BU1:BU2)"),
    ]
    terms: list[tuple[float, float, float, float | None, float | None]] = [
        (48.0, -200.0, 8000.0, None, None),
        (10.0, -100.0, 1000.0, None, None),
        (360.0, -1000.0, 200000.0, 0.0, 0.0),
        (12.0, 0.0, -1000.0, 2000.0, 0.0),
        (5.0, 100.0, 1000.0, 0.0, 0.0),
        (2.5, -100.0, 200.0, 0.0, 0.0),
        (48.0, -200.0, 8000.0, 0.0, 1.0),
        (10.0, -100.0, 900.0, 100.0, 0.0),
        (1.0, -110.0, 100.0, 0.0, 0.0),
        (0.0, -100.0, 1000.0, 0.0, 0.0),
        (-10.0, -100.0, 1000.0, 0.0, 0.0),
        (600.0, -100.0, 1000.0, 0.0, 0.0),
        (10.0, -1000.0, 1000.0, 0.0, 0.0),
        (10.0, 0.0, 0.0, 0.0, 0.0),
        (36.0, -300.0, 10000.0, -500.0, 0.5),
        (10.0, -100.0, 1000.0, 0.0, 2.0),
        (4.0, 0.0, -1000.0, 1e9, 0.0),
    ]
    for n, pmt, pv, fv, kind in terms:
        args = ["{1}", "{2}", "{3}", None if fv is None else "{4}", None if kind is None else "{5}"]
        inputs = tuple(v for v in (n, pmt, pv, fv, kind) if v is not None)
        add(Case("rate", _call("RATE", args), inputs))
    for g in (0.1, 0.0, -0.5, 1.0, 10.0, -0.99, -1.0, 1e-10, 0.5):
        add(Case("rate", "RATE(48,-200,8000,0,0,{1})", (g,)))
        add(Case("rate", "RATE(12,0,-1000,2000,0,{1})", (g,)))
        add(Case("rate", "RATE(5,100,1000,0,0,{1})", (g,)))
    for pos in range(3):
        for v in KINDS_ARG:
            inputs = tuple(v if i == pos else a for i, a in enumerate((48.0, -200.0, 8000.0)))
            add(Case("rate", "RATE({1},{2},{3})", inputs))
    for v in KINDS_ARG:
        add(Case("rate", "RATE(48,-200,8000,0,0,{1})", (v,)))
    cases += [
        Case("rate", "RATE(48,-200,8000,,,0.2)"),
        Case("rate", "RATE({48,24},-200,8000)"),
        Case("rate", "RATE(48,-200,8000)*12"),
    ]
    return cases


# Bonds and day counts ---------------------------------------------------------------------------


def _serial(y: int, m: int, d: int) -> float:
    import datetime as dt

    return float((dt.date(y, m, d) - dt.date(1899, 12, 30)).days)


DAY_COUNT_DATES = [
    _serial(*ymd)
    for ymd in [
        (2023, 1, 31), (2023, 2, 28), (2024, 2, 28), (2024, 2, 29), (2024, 3, 31), (2023, 3, 30),
        (2023, 12, 31), (2024, 1, 1), (2020, 2, 29), (2021, 2, 28), (2023, 5, 31), (2023, 6, 30),
        (2025, 8, 15), (2030, 11, 30),
    ]
]  # fmt: skip
BOND_DATES = [
    (_serial(2008, 2, 15), _serial(2017, 11, 15)),
    (_serial(2023, 1, 31), _serial(2026, 7, 31)),
    (_serial(2024, 2, 29), _serial(2034, 8, 31)),
    (_serial(2023, 3, 10), _serial(2023, 9, 1)),
    (_serial(2023, 11, 30), _serial(2024, 2, 29)),
    (_serial(2021, 6, 15), _serial(2051, 6, 15)),
]


def bonds() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for a in DAY_COUNT_DATES:
        for b in DAY_COUNT_DATES:
            add(Case("days360", "DAYS360({1},{2})", (a, b)))
            add(Case("days360", "DAYS360({1},{2},TRUE)", (a, b)))
    for a in DAY_COUNT_DATES:
        for b in DAY_COUNT_DATES:
            for basis in (0.0, 1.0, 2.0, 3.0, 4.0):
                add(Case("yearfrac", "YEARFRAC({1},{2},{3})", (a, b, basis)))
    cases += [
        Case("yearfrac", "YEARFRAC(44927,45000)"),
        Case("yearfrac", "YEARFRAC(44927,45000,5)"),
        Case("yearfrac", "YEARFRAC(44927,45000,-1)"),
        Case("yearfrac", "YEARFRAC(44927,45000,1.9)"),
        Case("yearfrac", "YEARFRAC(44927.9,45000.2,1)"),
        Case("yearfrac", 'YEARFRAC("1/1/2023","3/15/2023",1)'),
        Case("yearfrac", "YEARFRAC(-1,45000)"),
        Case("yearfrac", "YEARFRAC(1,59,1)"),
        Case("yearfrac", "YEARFRAC(1,61,1)"),
        Case("yearfrac", "YEARFRAC(36526,73051,1)"),
        Case("days360", "DAYS360(44927.9,45000.2)"),
        Case("days360", 'DAYS360("1/31/2023","3/31/2023")'),
        Case("days360", "DAYS360(-1,45000)"),
        Case("days360", "DAYS360(1,61)"),
        Case("days360", "DAYS360(44927,45000,2)"),
    ]
    coup = ("COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD")
    for s, m in BOND_DATES:
        for freq in (1.0, 2.0, 4.0):
            for basis in (0.0, 1.0, 2.0, 3.0, 4.0):
                for fn in coup:
                    add(Case(fn.lower(), f"{fn}({{1}},{{2}},{{3}},{{4}})", (s, m, freq, basis)))
                add(Case("price", "PRICE({1},{2},{3},{4},{5},{6},{7})",
                         (s, m, 0.0575, 0.065, 100.0, freq, basis)))  # fmt: skip
                add(Case("yield", "YIELD({1},{2},{3},{4},{5},{6},{7})",
                         (s, m, 0.0575, 95.04287, 100.0, freq, basis)))  # fmt: skip
                add(Case("duration", "DURATION({1},{2},{3},{4},{5},{6})",
                         (s, m, 0.08, 0.09, freq, basis)))  # fmt: skip
                add(Case("mduration", "MDURATION({1},{2},{3},{4},{5},{6})",
                         (s, m, 0.08, 0.09, freq, basis)))  # fmt: skip
    s, m = BOND_DATES[0]
    for rate, yld, red in [(0.0, 0.065, 100.0), (0.0575, 0.0, 100.0), (0.0575, -0.01, 100.0),
                           (0.0575, 0.065, 0.0), (-0.01, 0.065, 100.0), (0.0575, 0.065, 105.0),
                           (0.2, 0.5, 100.0), (0.0575, 10.0, 100.0)]:  # fmt: skip
        add(Case("price", "PRICE({1},{2},{3},{4},{5},2,0)", (s, m, rate, yld, red)))
    for rate, pr, red in [(0.0575, 100.0, 100.0), (0.0575, 50.0, 100.0), (0.0575, 200.0, 100.0),
                          (0.0, 95.0, 100.0), (0.0575, 0.0, 100.0), (0.0575, 95.0, 0.0),
                          (0.2, 30.0, 100.0), (0.0575, 1.0, 100.0)]:  # fmt: skip
        add(Case("yield", "YIELD({1},{2},{3},{4},{5},2,0)", (s, m, rate, pr, red)))
    for fn, args in [("PRICE", "{1},{2},0.0575,0.065,100"), ("YIELD", "{1},{2},0.0575,95,100"),
                     ("MDURATION", "{1},{2},0.08,0.09"), ("DURATION", "{1},{2},0.08,0.09"),
                     ("COUPNUM", "{1},{2}"), ("COUPDAYS", "{1},{2}")]:  # fmt: skip
        for tail in (",2", ",3", ",2,5", ",2,-1", ",2.9,0", ",2,0.5"):
            add(Case(fn.lower(), f"{fn}({args}{tail})", (s, m)))
        add(Case(fn.lower(), f"{fn}({args},2)", (m, s)))
        add(Case(fn.lower(), f"{fn}({args},2)", (s, s)))
        add(Case(fn.lower(), f"{fn}({args},2)", (s + 0.7, m + 0.2)))
        add(Case(fn.lower(), f"{fn}({args},2)", ("2/15/2008", m)))
        add(Case(fn.lower(), f"{fn}({args},2)", (True, m)))
    cases += [
        Case("price", "PRICE(39493,43054,0.0575,0.065,100,2)"),
        Case("yield", "YIELD(39493,43054,0.0575,95.04287,100,2)"),
        Case("mduration", "MDURATION(39493,43054,0.08,0.09,2)"),
        Case("price", "PRICE(39493,39494,0.0575,0.065,100,2,0)"),
        Case("yield", "YIELD(39493,39494,0.0575,99.99,100,2,0)"),
        Case("yield", "YIELD(39493,39600,0.0575,99,100,2,1)"),
        Case("mduration", "MDURATION(39493,43054,0.08,-0.5,2,0)"),
        Case("mduration", "MDURATION(39493,43054,-0.08,0.09,2,0)"),
    ]
    return cases


# Working days and time parts ----------------------------------------------------------------

WORK_TABLE: dict[str, Input] = {
    # holidays: a weekday, a Saturday, a duplicate, a text date, a fraction, a blank, text
    "CB1": _serial(2024, 1, 1), "CB2": _serial(2024, 1, 6), "CB3": _serial(2024, 1, 1),
    "CB4": "1/15/2024", "CB5": _serial(2024, 2, 19) + 0.75, "CB7": "x",
    "CC1": _serial(2024, 1, 1), "CC2": _serial(2024, 1, 15), "CC3": _serial(2024, 12, 25),
    "CD1": _serial(2024, 1, 1), "CD2": NA,
}  # fmt: skip
WORK_STARTS = [
    _serial(2023, 12, 29), _serial(2024, 1, 5), _serial(2024, 1, 6), _serial(2024, 1, 7),
    _serial(2024, 1, 1), _serial(2024, 2, 19) + 0.5,
]  # fmt: skip
WEEKENDS: list[Input] = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0,
                         17.0, 0.0, 8.0, 18.0, "0000011", "1000001", "1111110", "1111111",
                         "0000000", "000001", "00000a1", "abc", True, None, 1.5]  # fmt: skip


def workdays() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for s in WORK_STARTS:
        for days in (0.0, 1.0, -1.0, 5.0, -5.0, 10.0, 2.7, -2.7, 260.0, -260.0):
            add(Case("workday", "WORKDAY({1},{2})", (s, days)))
            add(Case("workday", "WORKDAY({1},{2},CB1:CB5)", (s, days)))
        for e in WORK_STARTS:
            add(Case("networkdays", "NETWORKDAYS({1},{2})", (s, e + 30)))
            add(Case("networkdays", "NETWORKDAYS({1},{2},CC1:CC3)", (s, e + 30)))
            add(Case("networkdays", "NETWORKDAYS({1},{2})", (e + 30, s)))
    for w in WEEKENDS:
        for name in ("WORKDAY.INTL", "_xlfn.WORKDAY.INTL"):
            add(Case("workday.intl", f"{name}({{1}},7,{{2}})", (WORK_STARTS[0], w)))
        for name in ("NETWORKDAYS.INTL", "_xlfn.NETWORKDAYS.INTL"):
            add(Case("networkdays.intl", f"{name}({{1}},{{2}},{{3}})",
                     (WORK_STARTS[0], WORK_STARTS[0] + 40, w)))  # fmt: skip
    for s in WORK_STARTS:
        add(Case("workday.intl", "WORKDAY.INTL({1},-9,7,CB1:CB5)", (s,)))
        add(Case("networkdays.intl", "NETWORKDAYS.INTL({1},{2},11,CC1:CC3)", (s, s + 400)))
    for v in KINDS_ARG:
        add(Case("workday", "WORKDAY({1},3)", (v,)))
        add(Case("workday", "WORKDAY(45292,{1})", (v,)))
        add(Case("networkdays", "NETWORKDAYS({1},45320)", (v,)))
        add(Case("networkdays", "NETWORKDAYS(45292,45320,{1})", (v,)))
    cases += [
        Case("workday", "WORKDAY(45292,3,CB7)"),
        Case("workday", "WORKDAY(45292,3,CD1:CD2)"),
        Case("workday", "WORKDAY(45292,3,{45293,45294})"),
        Case("workday", "WORKDAY(-1,3)"),
        Case("workday", "WORKDAY(2958465,1)"),
        Case("workday", "WORKDAY(1,-1)"),
        Case("workday", "WORKDAY(45292,{1,2})"),
        Case("networkdays", "NETWORKDAYS(45292,45320,CB7)"),
        Case("networkdays", "NETWORKDAYS(45292,45320,CD1:CD2)"),
        Case("networkdays", "NETWORKDAYS(45292,45292)"),
        Case("networkdays", "NETWORKDAYS(45293,45293)"),
        Case("networkdays", "NETWORKDAYS(-1,45320)"),
        Case("networkdays", "NETWORKDAYS(0,10)"),
        Case("networkdays", 'NETWORKDAYS("1/1/2024","2/1/2024")'),
    ]
    times: list[Input] = [0.0, 0.5, 0.75, 0.999999, 0.9999999999, 45000.25, 1 / 3, 2 / 3,
                          0.5 + 1 / 86400 * 0.4999, 0.5 + 1 / 86400 * 0.5, 0.5 + 0.5 / 1440,
                          -0.5, 2958465.99999, 2958466.0, "12:30:45", "3:15 PM",
                          "25:00", "1/2/2024 6:07:08", True, "abc", None, EMPTY_TEXT, NA,
                          0.4999999999999999, 0.041666666666666664]  # fmt: skip
    for fn in ("HOUR", "MINUTE", "SECOND"):
        for v in times:
            add(Case(fn.lower(), f"{fn}({{1}})", (v,)))
    return cases


# Statistics -----------------------------------------------------------------------------------

_STAT_COLUMNS: dict[str, list[Input]] = {
    "DA": [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0],
    "DB": [1.0, "3", True, None, "abc", 4.5, EMPTY_TEXT, 0.0, -1.5, False],
    "DC": [1e9 + 0.1, 1e9 + 0.2, 1e9 + 0.3, 1e9 + 0.4],
    "DD": [1.0, 2.0, 3.0, 4.0, 5.0, 6.0],
    "DE": [2.1, 3.9, 6.2, 7.8, 10.1, 12.2],
    "DF": [5.0, 5.0, 5.0, 5.0, 5.0, 5.0],
    "DG": [1.0, NA, 3.0],
    "DH": [7.0],
    "DI": [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7],
    "DJ": [10.0, 20.0, 20.0, 30.0, 5.0, 20.0],
    "DK": [1.0, "x", 3.0, None, 5.0, True],
    "DL": [3.0, 1.0, 4.0, 1.0, 5.0, 9.0, 2.0, 6.0],
    "DM": [1e-300, 2e-300, 3e-300],
    "DN": [1e300, 1e300, -1e300],
}
STAT_TABLE: dict[str, Input] = {
    f"{col}{i}": v
    for col, values in _STAT_COLUMNS.items()
    for i, v in enumerate(values, 1)
    if v is not None
}


def _col(col: str) -> str:
    return f"{col}1:{col}{len(_STAT_COLUMNS[col])}"


def statistics() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    aggregates = ("STDEV", "STDEVP", "_xlfn.STDEV.S", "_xlfn.STDEV.P", "VAR", "VARP",
                  "_xlfn.VAR.S", "_xlfn.VAR.P", "AVERAGEA", "MINA", "MAXA", "MEDIAN", "SUMSQ",
                  "STDEVA", "STDEVPA", "VARA", "VARPA")  # fmt: skip
    for fn in aggregates:
        tag = fn.removeprefix("_xlfn.").lower()
        for col in ("DA", "DB", "DC", "DG", "DH", "DI", "DK", "DL", "DM", "DN"):
            add(Case(tag, f"{fn}({_col(col)})"))
        cases += [
            Case(tag, f"{fn}(2,4,4,4,5,5,7,9)"),
            Case(tag, f'{fn}(1,"3",TRUE)'),
            Case(tag, f"{fn}({{1,2;3,4}})"),
            Case(tag, f'{fn}({{1,"a",TRUE,2}})'),
            Case(tag, f"{fn}(DA1:DA4,DL1:DL3,10)"),
            Case(tag, f"{fn}(1)"),
            Case(tag, f'{fn}("abc",1,2)'),
            Case(tag, f"{fn}(DA1:DA8*1)"),
        ]
        for v in KINDS_ARG:
            add(Case(tag, f"{fn}({{1}},1,2)", (v,)))
    for x, y in (("DD", "DE"), ("DD", "DF"), ("DD", "DD"), ("DA", "DL"), ("DC", "DA")):
        add(Case("correl", f"CORREL({x}1:{x}6,{y}1:{y}6)"))
    cases += [
        Case("correl", "CORREL(DD1:DD6,DE1:DE5)"),
        Case("correl", "CORREL(DK1:DK6,DD1:DD6)"),
        Case("correl", "CORREL(DG1:DG3,DD1:DD3)"),
        Case("correl", "CORREL({1,2,3},{2,4,7})"),
        Case("correl", "CORREL(DH1,DH1)"),
        Case("correl", "CORREL(DD1:DD2,DE1:DE2)"),
    ]
    for args in ("DE1:DE6,DD1:DD6", "DE1:DE6", "DE1:DE6,DD1:DD6,FALSE", "DE1:DE6,DD1:DD6,TRUE,TRUE",
                 "DF1:DF6,DD1:DD6", "DE1:DE6,DD1:DD6,TRUE,FALSE"):  # fmt: skip
        add(Case("linest", f"LINEST({args})"))
        for r, c in ((1, 2), (2, 1), (2, 2), (3, 1), (3, 2), (4, 1), (4, 2), (5, 1), (5, 2)):
            add(Case("linest", f"INDEX(LINEST({args}),{r},{c})"))
    cases += [
        Case("linest", "INDEX(LINEST({1,2,4},{1,2,3}),1,1)"),
        Case("linest", "LINEST(DE1:DE6,DD1:DD5)"),
        Case("linest", "LINEST(DK1:DK6,DD1:DD6)"),
        Case("linest", "INDEX(LINEST(DC1:DC4,DD1:DD4,TRUE,TRUE),3,1)"),
    ]
    for fn in ("RANK", "_xlfn.RANK.EQ", "_xlfn.RANK.AVG"):
        tag = fn.removeprefix("_xlfn.").lower()
        for v in (20.0, 10.0, 30.0, 5.0, 25.0, 20.5):
            add(Case(tag, f"{fn}({{1}},DJ1:DJ6)", (v,)))
            add(Case(tag, f"{fn}({{1}},DJ1:DJ6,1)", (v,)))
        cases += [
            Case(tag, f"{fn}(3,DK1:DK6)"),
            Case(tag, f"{fn}(3,DK1:DK6,-1)"),
            Case(tag, f"{fn}(1,DG1:DG3)"),
            Case(tag, f'{fn}("20",DJ1:DJ6)'),
            Case(tag, f"{fn}(TRUE,DK1:DK6)"),
        ]
    zs = [-40.0, -38.5, -10.0, -5.0, -3.0, -1.5, -1.0, -0.5, -1e-10, 0.0, 0.5, 1.0, 1.96, 3.0,
          5.0, 8.3, 10.0, 40.0, 0.6744897501960817, -2.3263478740408408]  # fmt: skip
    for z in zs:
        add(Case("normsdist", "NORMSDIST({1})", (z,)))
        add(Case("norm.s.dist", "_xlfn.NORM.S.DIST({1},TRUE)", (z,)))
        add(Case("norm.s.dist", "_xlfn.NORM.S.DIST({1},FALSE)", (z,)))
    normals = [(42.0, 40.0, 1.5), (0.0, 0.0, 1.0), (100.0, 0.0, 1.0), (-3.0, 1.0, 2.0),
               (1.0, 1.0, 1e-10), (1.0, 1.0, 0.0), (1.0, 1.0, -1.0), (5.0, 2.0, 1e10)]  # fmt: skip
    for x, mean, sd in normals:
        for cum in ("TRUE", "FALSE"):
            add(Case("normdist", f"NORMDIST({{1}},{{2}},{{3}},{cum})", (x, mean, sd)))
            add(Case("norm.dist", f"_xlfn.NORM.DIST({{1}},{{2}},{{3}},{cum})", (x, mean, sd)))
    ps = [1e-300, 1e-10, 0.001, 0.02425, 0.025, 0.1, 0.3, 0.5, 0.6, 0.908789, 0.97575, 0.975,
          0.999, 1 - 1e-10, 0.0, 1.0, -0.1, 1.5]  # fmt: skip
    for p in ps:
        add(Case("normsinv", "NORMSINV({1})", (p,)))
        add(Case("norm.s.inv", "_xlfn.NORM.S.INV({1})", (p,)))
        add(Case("norminv", "NORMINV({1},40,1.5)", (p,)))
        add(Case("norm.inv", "_xlfn.NORM.INV({1},40,1.5)", (p,)))
    for v in KINDS_ARG:
        add(Case("normsdist", "NORMSDIST({1})", (v,)))
        add(Case("norminv", "NORMINV({1},0,1)", (v,)))
        add(Case("normdist", "NORMDIST(1,0,1,{1})", (v,)))
    cases += [
        Case("norminv", "NORMINV(0.5,0,0)"),
        Case("norminv", "NORMINV(0.5,0,-1)"),
        Case("normsdist", "NORMSDIST({-1,0,1})"),
        Case("normsdist", "SUM(NORMSDIST({-1,1}))"),
    ]
    return cases


# The rest of the most used: TRANSPOSE, ADDRESS, REPT, RAND, the D functions, REPLACE, CHAR ------

MISC_TABLE: dict[str, Input] = {
    # a database: a header row and six records
    "DP1": "Tree", "DQ1": "Height", "DR1": "Age", "DS1": "Profit",
    "DP2": "Apple", "DQ2": 18.0, "DR2": 20.0, "DS2": 14.0,
    "DP3": "Pear", "DQ3": 12.0, "DR3": 12.0, "DS3": 10.0,
    "DP4": "Cherry", "DQ4": 13.0, "DR4": 14.0, "DS4": 9.0,
    "DP5": "Apple", "DQ5": 14.0, "DR5": 15.0, "DS5": 10.0,
    "DP6": "Pear", "DQ6": 9.0, "DR6": 8.0, "DS6": 8.0,
    "DP7": "Applesauce", "DQ7": 8.0, "DR7": 9.0, "DS7": "n/a",
    # criteria ranges
    "DU1": "Tree", "DU2": "Apple",
    "DV1": "Height", "DV2": ">10",
    "DW1": "Tree", "DX1": "Height", "DW2": "Apple", "DX2": ">10",
    "DY1": "Tree", "DY2": "Apple", "DY3": "Pear",
    "DZ1": "Tree", "DZ2": "=Apple",
    "EA1": "Tree", "EA2": "A*",
    "EB1": "Height",
    "EC1": "Nope", "EC2": "x",
    "ED1": "Calc", "ED2": FormulaInput("=DQ2>DR2"),
    "EE1": "Height", "EF1": "Height", "EE2": ">10", "EF2": "<16",
    "EG1": "tree", "EG2": "pear",
    # cells for TRANSPOSE
    "EI1": 1.0, "EJ1": 2.0, "EK1": 3.0, "EI2": "a", "EJ2": True, "EK2": NA,
}  # fmt: skip
CRITERIA_RANGES = ["DU1:DU2", "DV1:DV2", "DW1:DX2", "DY1:DY3", "DZ1:DZ2", "EA1:EA2", "EB1:EB2",
                   "EC1:EC2", "ED1:ED2", "EE1:EF2", "EG1:EG2"]  # fmt: skip


def misc() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    cases += [
        Case("transpose", "TRANSPOSE(EI1:EK2)"),
        Case("transpose", "INDEX(TRANSPOSE(EI1:EK2),3,1)"),
        Case("transpose", "INDEX(TRANSPOSE(EI1:EK2),2,2)"),
        Case("transpose", "ROWS(TRANSPOSE(EI1:EK2))"),
        Case("transpose", "COLUMNS(TRANSPOSE(EI1:EK2))"),
        Case("transpose", "SUM(TRANSPOSE(EI1:EK1))"),
        Case("transpose", "SUMPRODUCT(TRANSPOSE(EI1:EK1)*{1;2;3})"),
        Case("transpose", "INDEX(TRANSPOSE({1,2,3}),2,1)"),
        Case("transpose", "TRANSPOSE(5)"),
        Case("transpose", 'TRANSPOSE("x")'),
        Case("transpose", "INDEX(TRANSPOSE(EI1:EK2),2,3)"),
        Case("transpose", "INDEX(TRANSPOSE(EI1:EI2),1,2)"),
    ]
    for r, c in ((1.0, 1.0), (5.0, 3.0), (1048576.0, 16384.0), (0.0, 1.0), (1.0, 16385.0),
                 (2.9, 27.9), (100.0, 702.0), (100.0, 703.0)):  # fmt: skip
        add(Case("address", "ADDRESS({1},{2})", (r, c)))
        for kind in (1.0, 2.0, 3.0, 4.0, 0.0, 5.0):
            add(Case("address", "ADDRESS({1},{2},{3})", (r, c, kind)))
            add(Case("address", "ADDRESS({1},{2},{3},FALSE)", (r, c, kind)))
    for sheet in ("Sheet1", "My Sheet", "", "it's", "A1", "[b]x", "Résumé", "1x", "TRUE"):
        add(Case("address", "ADDRESS(2,3,1,TRUE,{1})", (sheet,)))
        add(Case("address", "ADDRESS(2,3,4,FALSE,{1})", (sheet,)))
    for v in KINDS_ARG:
        add(Case("address", "ADDRESS({1},2)", (v,)))
        add(Case("address", "ADDRESS(2,2,1,{1})", (v,)))
    for text, n in (("ab", 3.0), ("ab", 0.0), ("ab", 2.9), ("ab", -1.0), ("", 5.0),
                    ("x", 32767.0), ("x", 32768.0), ("ab", 16384.0), ("ab", 16383.5), (1.5, 2.0),
                    (True, 2.0), ("é", 3.0)):  # fmt: skip
        add(Case("rept", "REPT({1},{2})", (text, n)))
        add(Case("rept", "LEN(REPT({1},{2}))", (text, n)))
    cases += [
        Case("rand", "AND(RAND()>=0,RAND()<1)"),
        Case("rand", "ISNUMBER(RAND())"),
        Case("randbetween", "RANDBETWEEN(3,3)"),
        Case("randbetween", "RANDBETWEEN(5,1)"),
        Case("randbetween", "RANDBETWEEN(1.5,2.5)"),
        Case("randbetween", "RANDBETWEEN(2.5,2.7)"),
        Case("randbetween", "AND(RANDBETWEEN(1,6)>=1,RANDBETWEEN(1,6)<=6)"),
        Case("randbetween", "RANDBETWEEN(-2.5,-2.2)"),
        Case("randbetween", 'RANDBETWEEN("a",2)'),
        Case("randbetween", "INT(RANDBETWEEN(1,10))=RANDBETWEEN(4,4)+0*RAND()-4+RANDBETWEEN(1,1)"),
    ]
    fields = ['"Profit"', "4", '"Age"', '"profit"', "0", "5", '"Nope"', "DS1", "2.9", "TRUE"]
    for fn in ("DSUM", "DAVERAGE", "DCOUNT", "DMAX", "DMIN"):
        for crit in CRITERIA_RANGES:
            add(Case(fn.lower(), f'{fn}(DP1:DS7,"Profit",{crit})'))
        for field in fields:
            add(Case(fn.lower(), f"{fn}(DP1:DS7,{field},DY1:DY3)"))
        cases += [
            Case(fn.lower(), f'{fn}(DP1:DS7,"Tree",DU1:DU2)'),
            Case(fn.lower(), f'{fn}(DP1:DS1,"Profit",DU1:DU2)'),
            Case(fn.lower(), f'{fn}(DP1:DS7,"Profit",DU1)'),
            Case(fn.lower(), f'{fn}(DP2:DS7,"Profit",DU1:DU2)'),
        ]
    cases += [Case("dcount", "DCOUNT(DP1:DS7,,DY1:DY3)"), Case("dsum", "DSUM(DP1:DS7,,DY1:DY3)")]
    for old, start, num, new in (("abcdef", 2.0, 3.0, "XY"), ("abcdef", 1.0, 0.0, "Z"),
                                 ("abcdef", 7.0, 2.0, "Z"), ("abcdef", 10.0, 2.0, "Z"),
                                 ("abcdef", 0.0, 1.0, "Z"), ("abcdef", 2.0, -1.0, "Z"),
                                 ("abcdef", 2.9, 1.9, "Z"), ("abc", 2.0, 100.0, ""),
                                 (12345.0, 2.0, 2.0, 0.0), ("", 1.0, 1.0, "x"),
                                 (True, 1.0, 1.0, "x"), ("abc", 32768.0, 1.0, "x")):  # fmt: skip
        add(Case("replace", "REPLACE({1},{2},{3},{4})", (old, start, num, new)))
    for n in [float(i) for i in range(0, 257)] + [1.5, 65.9, -1.0, 1e10]:
        add(Case("char", "CHAR({1})", (n,)))
    for v in KINDS_ARG:
        add(Case("char", "CHAR({1})", (v,)))
    for text in ("A", "abc", "", "é", "€", "ÿ", " ", "1", "\u2603", "\U0001f600"):
        add(Case("code", "CODE({1})", (text,)))
    add(Case("code", "CODE(CHAR(200))"))
    add(Case("code", "CODE(TRUE)"))
    add(Case("code", "CODE(65)"))
    return cases


# TEXT, DOLLAR, FIXED --------------------------------------------------------------------------

TEXT_VALUES: list[Input] = [
    0.0, 1.0, -1.0, 1234.5678, -1234.5678, 0.5, 0.125, 2.675, 1e15, 123456789012.345, 1e-7,
    45292.75, 0.999999, 60.0, 1.5e20, -0.0001, 12.0, 0.05, 3.14159265358979, 45000.0001,
    "abc", True, None, NA,
]  # fmt: skip
TEXT_FORMATS = [
    "0", "0.00", "#,##0", "#,##0.00", "0%", "0.0%", "0.00E+00", "##0.0E+0", "# ?/?",
    "# ??/??", "$#,##0.00_);($#,##0.00)", "mm/dd/yyyy", "dddd, mmmm d, yyyy", "h:mm AM/PM",
    "[h]:mm:ss", "General", "@", "0;-0;\"zero\"", "[Red]0.0", "*-0", "0_)", "00000",
    "\"Total: \"0", "yyyy-mm-dd hh:mm:ss.000", "mmm-yy", "d-mmm", "0.0,,\"M\"", "#,##0,",
    "[<1000]0;[>=1000]#,##0", "?.??", "0.00e+00", "", "0.0.0", "General;General;General",
    "#.##", ".00", "0 \"days\"", "[$€-2] #,##0.00", "mm:ss.0",
]  # fmt: skip


def textfmt() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for fmt in TEXT_FORMATS:
        for v in TEXT_VALUES:
            add(Case("text", "TEXT({1},{2})", (v, fmt)))
    for v in KINDS_ARG:
        add(Case("text", "TEXT(12.5,{1})", (v,)))
    cases += [
        Case("text", "TEXT(5,0)"),
        Case("text", "TEXT(5.5,0.0)"),
        Case("text", 'TEXT(-5,"0;0")'),
        Case("text", 'TEXT(0,";;;")'),
        Case("text", 'TEXT("12","0.00")'),
        Case("text", 'TEXT("1/2/2024","yyyy")'),
        Case("text", 'TEXT({1.5,2.5},"0")'),
        Case("text", 'TEXT(1234.5,"#,##0.00 ")'),
    ]
    numbers = [0.0, 1.0, -1.0, 1234.5678, -1234.5678, 0.5, 2.675, -2.675, 1.005, 1e15,
               123456789012.345, 1e-10, 9.995, 1.5e20, -0.004, 999999.995, 0.125]  # fmt: skip
    for fn in ("DOLLAR", "FIXED"):
        for x in numbers:
            add(Case(fn.lower(), f"{fn}({{1}})", (x,)))
            for d in (0.0, 1.0, 3.0, -1.0, -2.0, 2.9, -2.9, 10.0, 127.0, 128.0):
                add(Case(fn.lower(), f"{fn}({{1}},{{2}})", (x, d)))
        for v in KINDS_ARG:
            add(Case(fn.lower(), f"{fn}({{1}})", (v,)))
            add(Case(fn.lower(), f"{fn}(1234.5,{{1}})", (v,)))
    for x in numbers:
        add(Case("fixed", "FIXED({1},2,TRUE)", (x,)))
        add(Case("fixed", "FIXED({1},-1,TRUE)", (x,)))
    for v in KINDS_ARG:
        add(Case("fixed", "FIXED(1234.5,1,{1})", (v,)))
    cases += [Case("fixed", "FIXED(1234.5,,TRUE)"), Case("dollar", "DOLLAR(1234.5,)")]
    # A format of a quoted "#" alone: literal text, not a run of # (xlformat 0.3's TEXT).
    cases += [Case("text", 'TEXT(1,"""#""")'), Case("text", 'TEXT(-1,"""#""")')]
    return cases


# Rounding: ties, near-ties and 15- to 17-digit values ------------------------------------------


def rounding_values() -> list[float]:
    ties = [0.5, 1.5, 2.5, 0.05, 0.15, 0.25, 0.35, 1.005, 2.675, 0.125, 0.285, 1.115, 10.235,
            1234.5675, 0.045, 8.345, 1.0005, 5.55, 0.0005, 99.995, 12345.6785, 2.5e-7,
            123456789.125, 1e15 + 0.5, 4.35, 0.615]  # fmt: skip
    values = list(ties)
    for x in (2.675, 1.005, 0.285, 8.345, 99.995, 0.615):
        values += [math.nextafter(x, math.inf), math.nextafter(x, -math.inf)]
    values += [1.23456789012345, 0.1 + 0.2, 1 / 3, 2 / 3, 123.456789012345678, 7.0, 0.0,
               1e-300, 9.99999999999999e20, 5e-11]  # fmt: skip
    return values + [-v for v in values[:20]]


def rounding() -> list[Case]:
    cases = []
    for fn in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
        for x in rounding_values():
            for d in (-2.0, -1.0, 0.0, 1.0, 2.0, 3.0, 4.0, 6.0, 10.0):
                cases.append(Case(fn.lower(), f"{fn}({{1}},{{2}})", (x, d)))
    return cases


MODERN_TABLE: dict[str, Input] = {
    **{f"Z{i}": v for i, v in enumerate((3.0, 1.0, 2.0, 2.0, 5.0, 8.0, -1.0, 2.5), 1)},
    **{f"AA{i}": v for i, v in enumerate((1.0, "x", True, None, 2.0, EMPTY_TEXT, 4.0, False), 1)},
    **{f"AB{i}": v for i, v in enumerate((1.0, NA, 2.0, 3.0), 1)},
    **{
        f"AC{i}": v
        for i, v in enumerate(("apple", "Banana", "cherry", "apple", EMPTY_TEXT, "date"), 1)
    },
    **{f"AD{i}": 10.0 * i for i in range(1, 7)},
    **{f"AE{i}": "abcdef"[i - 1] for i in range(1, 7)},
    **{f"AF{i}": 2.0 * i - 1 for i in range(1, 7)},
    **{f"AG{i}": 13.0 - 2.0 * i for i in range(1, 7)},
    "AI1": 1.0,
    "AJ1": 2.0,
    "AK1": 3.0,
    "AI2": "one",
    "AJ2": "two",
    "AK2": "three",
    "AL1": ",",
    "AL2": ";",
    "AD8": None,
    "AC8": "blank return",
    "AF8": 100.0,
}
_NUM, _MIX, _ERR, _TXT = "$Z$1:$Z$8", "$AA$1:$AA$8", "$AB$1:$AB$4", "$AC$1:$AC$6"
_RET, _RET2, _ASC, _DESC, _NONE = (
    "$AD$1:$AD$6",
    "$AD$1:$AE$6",
    "$AF$1:$AF$6",
    "$AG$1:$AG$6",
    "$AH$1:$AH$5",
)
_CONDS: list[Input] = [">2", "<=2", "<>1", "1", "", "=", "<>", "x", "a*", 1.0, 2.0, True, ">100"]
_F = "_xlfn."


def _criteria_rows(add: Callable[[Case], None]) -> None:
    for c in _CONDS:
        for r in (_NUM, _MIX, _TXT):
            add(Case("averageif", f"AVERAGEIF({r},{{1}})", (c,)))
            add(Case("averageifs", f"AVERAGEIFS({r},{r},{{1}})", (c,)))
            add(Case("maxifs", f"{_F}MAXIFS({r},{r},{{1}})", (c,)))
            add(Case("minifs", f"{_F}MINIFS({r},{r},{{1}})", (c,)))
        add(Case("averageif", f"AVERAGEIF({_TXT},{{1}},{_RET})", (c,)))
        add(Case("averageifs", f"AVERAGEIFS({_RET},{_TXT},{{1}})", (c,)))
        add(Case("averageifs", f'AVERAGEIFS({_RET},{_TXT},{{1}},{_ASC},">3")', (c,)))
        add(Case("maxifs", f"{_F}MAXIFS({_RET},{_TXT},{{1}})", (c,)))
        add(Case("maxifs", f'{_F}MAXIFS({_RET},{_TXT},{{1}},{_ASC},">3")', (c,)))
        add(Case("minifs", f"{_F}MINIFS({_RET},{_TXT},{{1}})", (c,)))
        add(Case("minifs", f'{_F}MINIFS({_RET},{_TXT},{{1}},{_ASC},">3")', (c,)))
    for name, f in (
        ("averageifs", "AVERAGEIFS"),
        ("maxifs", f"{_F}MAXIFS"),
        ("minifs", f"{_F}MINIFS"),
    ):
        add(Case(name, f'{f}({_ERR},{_ERR},">0")'))
        add(Case(name, f'{f}({_ERR},{_ERR},"<>")'))
        add(Case(name, f'{f}({_ERR},{_ERR},"#N/A")'))
        add(Case(name, f'{f}($AD$1:$AD$5,{_TXT},"a*")'))
        add(Case(name, f'{f}({_NONE},{_NONE},"")'))
        add(Case(name, f'{f}({_MIX},{_NUM},">0")'))
    add(Case("averageif", f'AVERAGEIF({_ERR},">0")'))
    add(Case("averageif", f'AVERAGEIF({_ERR},"<>")'))
    add(Case("averageif", f'AVERAGEIF({_ERR},"#N/A")'))
    add(Case("averageif", f'AVERAGEIF({_NONE},"")'))
    add(Case("averageif", f'AVERAGEIF({_TXT},"a*",$AD$3)'))
    add(Case("averageif", f'AVERAGEIF({_TXT},"a*",$XFD$1)'))
    add(Case("averageif", f'AVERAGEIF({_NUM},">0",$AD$1:$AD$2)'))


def _order_rows(add: Callable[[Case], None]) -> None:
    ks: list[Input] = [1.0, 2.0, 8.0, 9.0, 0.0, -1.0, 1.5, 2.9, "2", True, None, "x", NA]
    for fn in ("LARGE", "SMALL"):
        for k in ks:
            add(Case(fn.lower(), f"{fn}({_NUM},{{1}})", (k,)))
        for arg in (
            _MIX,
            _ERR,
            _NONE,
            _TXT,
            "{1,2,3}",
            '{1,"a",TRUE,4}',
            "{1;2;3}",
            "5",
            '"a"',
            '"5"',
            "TRUE",
            "1/0",
            "$Z$1",
            "$AH$1",
            f"({_NUM},{_ASC})",
        ):
            add(Case(fn.lower(), f"{fn}({arg},1)"))
    ps: list[Input] = [
        0.0,
        0.1,
        0.25,
        1.0 / 3,
        0.5,
        0.75,
        0.9,
        0.99,
        1.0,
        -0.1,
        1.1,
        "0.5",
        True,
        False,
        None,
        "x",
        NA,
        1e-10,
        0.123456789,
    ]
    for fn, name in (("PERCENTILE", "percentile"), (f"{_F}PERCENTILE.INC", "percentile.inc")):
        for k in ps:
            add(Case(name, f"{fn}({_NUM},{{1}})", (k,)))
            add(Case(name, f"{fn}({_ASC},{{1}})", (k,)))
        for arg in (
            _MIX,
            _ERR,
            _NONE,
            _TXT,
            "{1,2,3,4}",
            '{1,"a",TRUE,4}',
            "5",
            '"a"',
            "TRUE",
            "1/0",
            "$Z$1",
            "{0.1,0.2,0.3,0.7}",
        ):
            add(Case(name, f"{fn}({arg},0.3)"))
    qs: list[Input] = [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, -1.0, 1.9, 0.5, "2", True, None, "x", NA]
    for fn, name in (("QUARTILE", "quartile"), (f"{_F}QUARTILE.INC", "quartile.inc")):
        for q in qs:
            add(Case(name, f"{fn}({_NUM},{{1}})", (q,)))
        for arg in (_MIX, _ERR, _NONE, "{1,2,3,4}", "5", '"a"', "1/0"):
            add(Case(name, f"{fn}({arg},1)"))
    for fn, name in (("MODE", "mode"), (f"{_F}MODE.SNGL", "mode.sngl")):
        for args in (
            _NUM,
            _MIX,
            _ERR,
            _NONE,
            _TXT,
            _ASC,
            "1,2,2,3",
            "1,2,3",
            "1,1,2,2",
            "2,2,1,1",
            "{1,2,2}",
            "{1,2,3}",
            '"a",1,1',
            "TRUE,1,1",
            '"1",1',
            "1,,1",
            ",1,1",
            f"{_NUM},{_ASC}",
            "1/0,1,1",
            "1,1,1/0",
            "5",
            '"a"',
            "TRUE",
            '{1,"a",1}',
            "2.5,2.5,0.1,0.1",
            "$Z$3:$Z$4,$Z$3",
        ):
            add(Case(name, f"{fn}({args})"))


def _branch_rows(add: Callable[[Case], None]) -> None:
    f = f"{_F}IFS"
    for args in (
        "TRUE,1",
        "FALSE,1",
        "FALSE,1,TRUE,2",
        "TRUE,1,TRUE,2",
        '1>2,"a",1<2,"b"',
        "FALSE,1/0,TRUE,2",
        "TRUE,1/0,TRUE,2",
        "1/0,1,TRUE,2",
        "FALSE,1,1/0,2",
        '"TRUE",1',
        '"x",1',
        '"false",1,TRUE,2',
        "0,1,2,3",
        "$AH$1,1,TRUE,2",
        '"",1,TRUE,2',
        f"TRUE,{_NUM}",
        f"{_NUM}>2,1",
        "TRUE,$AH$1",
        'TRUE,""',
        "{TRUE,FALSE},1",
        "{FALSE,TRUE},1,TRUE,2",
        "FALSE,1,TRUE,$Z$1",
        "NA(),1",
        "ISNA(NA()),NA()",
    ):
        add(Case("ifs", f"{f}({args})"))
    f = f"{_F}SWITCH"
    for args in (
        '1,1,"a"',
        '2,1,"a"',
        '2,1,"a","d"',
        '2,1,"a",2,"b"',
        '"A","a",1',
        '1,"1",1',
        '"1",1,1',
        'TRUE,1,"x"',
        'TRUE,TRUE,"x"',
        '$AH$1,0,"z","n"',
        '$AH$1,"","e","n"',
        '$AH$1,$AH$2,"b","n"',
        '0,$AH$1,"b","n"',
        "1/0,1,2,3",
        "1,1/0,2,3",
        "1,2,1/0,1,4",
        "1,1,4,2,1/0",
        "1,2,3,1/0",
        "1,1,1/0",
        "NA(),NA(),1,2",
        f"$Z$2,{_NUM},1,2",
        f"{_NUM},1,2,3",
        '{1,2},1,"a","z"',
        '1,{1,2},"a","z"',
        "1,1,$Z$1",
        "1,2,3,$Z$1",
        '1.5,1.5,"y"',
        '"a*","apple",1,2',
        '"apple","a*",1,2',
    ):
        add(Case("switch", f"{f}({args})"))


def _text_rows(add: Callable[[Case], None]) -> None:
    tj, cc = f"{_F}TEXTJOIN", f"{_F}CONCAT"
    values = [
        '"a"',
        '"a","b"',
        "1.5",
        "1E+20",
        "0.1+0.2",
        "TRUE",
        "$AH$1",
        '"",""',
        "1/0",
        _TXT,
        _NUM,
        _MIX,
        _ERR,
        _NONE,
        "{1,2;3,4}",
        '"a",$AH$1,"b"',
        "$AI$1:$AK$2",
        '"a",,"b"',
        "1234567.891",
        "-0.5",
        "45720.5",
        "NA()",
    ]
    for v in values:
        add(Case("concat", f"{cc}({v})"))
        for ignore in ("TRUE", "FALSE"):
            add(Case("textjoin", f'{tj}(",",{ignore},{v})'))
    for delim in (
        '""',
        '" - "',
        "1",
        "TRUE",
        "$AH$1",
        "1/0",
        "$AL$1:$AL$2",
        '{",",";"}',
        "{1,2}",
        "$AL$1",
    ):
        add(Case("textjoin", f'{tj}({delim},TRUE,"a","b","c")'))
        add(Case("textjoin", f'{tj}({delim},FALSE,"a","","c","d")'))
    for ignore in ("1", "0", '"x"', '"TRUE"', "$AH$1", "1/0", "$AL$1", "{TRUE,FALSE}"):
        add(Case("textjoin", f'{tj}(",",{ignore},"a","","b")'))
    for n in (32767, 32768):
        add(Case("concat", f'LEN({cc}(REPT("a",{n - 1}),"b"))'))
        add(Case("textjoin", f'LEN({tj}("",TRUE,REPT("a",{n - 1}),"b"))'))
    add(Case("concat", f'{cc}(REPT("a",32767),REPT("a",32767))'))
    add(Case("textjoin", f'{tj}(",",TRUE,REPT("a",32767),"b")'))


def _lookup_rows(add: Callable[[Case], None]) -> None:
    xl, xm = f"{_F}XLOOKUP", f"{_F}XMATCH"
    keys: list[Input] = [
        "apple",
        "APPLE",
        "banana",
        "a*",
        "?pple",
        "zzz",
        "",
        None,
        "~*",
        1.0,
        True,
        NA,
    ]
    for k in keys:
        add(Case("xlookup", f"{xl}({{1}},{_TXT},{_RET})", (k,)))
        add(Case("xmatch", f"{xm}({{1}},{_TXT})", (k,)))
        add(Case("xlookup", f'{xl}({{1}},{_TXT},{_RET},"none")', (k,)))
        for mode in (2.0, -1.0, 1.0):
            add(Case("xlookup", f"{xl}({{1}},{_TXT},{_RET},,{mode:g})", (k,)))
            add(Case("xmatch", f"{xm}({{1}},{_TXT},{mode:g})", (k,)))
        for search in (-1.0, 2.0, -2.0):
            add(Case("xlookup", f"{xl}({{1}},{_TXT},{_RET},,0,{search:g})", (k,)))
            add(Case("xmatch", f"{xm}({{1}},{_TXT},0,{search:g})", (k,)))
    for k in (0.0, 1.0, 4.0, 5.0, 100.0, -1.0, 11.0, 12.0, "5", True, None):
        for mode in (0.0, -1.0, 1.0, 2.0):
            for search in (1.0, -1.0, 2.0, -2.0):
                add(Case("xlookup", f"{xl}({{1}},{_ASC},{_RET},,{mode:g},{search:g})", (k,)))
                add(Case("xmatch", f"{xm}({{1}},{_ASC},{mode:g},{search:g})", (k,)))
                add(Case("xmatch", f"{xm}({{1}},{_DESC},{mode:g},{search:g})", (k,)))
    for args in (
        f"3,{_ASC},{_RET},,3",
        f"3,{_ASC},{_RET},,0,0",
        f"3,{_ASC},{_RET},,0,3",
        f"3,{_ASC},$AD$1:$AD$5",
        f"3,{_ASC},$AD$1:$AE$1",
        f"99,{_ASC},{_RET},1/0",
        f"3,{_ASC},{_RET},1/0",
        f"1/0,{_ASC},{_RET}",
        f"3,{_ASC},{_RET2}",
        f"99,{_ASC},{_RET2}",
        '99,{0},{1},""',
        "2,$AI$1:$AK$1,$AI$2:$AK$2",
        '9,$AI$1:$AK$1,$AI$2:$AK$2,"n"',
        "2,$AI$1:$AK$1,$AI$1:$AK$2",
        "2,{1,2,3},{4,5,6}",
        "2,{1;2;3},{4;5;6}",
        "2,{1,2,3},{4;5;6}",
        '"blank return",$AC$8,$AD$8',
        f'3,{_ASC},{_RET},,"x"',
        f'3,{_ASC},{_RET},,0,"x"',
        f"3,{_ASC},{_RET},,0.5",
        f"3,{_ASC},{_RET},,1.9",
        f"3,{_ASC},{_RET},,TRUE",
        f"3,{_ASC},{_RET},,",
        f"3,{_ASC},{_RET},,,",
        f"99,{_ASC},{_RET},,,",
        "3,5,5",
        "5,5,6",
        '"a","a","b"',
        f"3,$AF$1:$AG$6,{_RET}",
        f'{_NUM},{_ASC},{_RET},"n"',
        f"$Z$1,{_ASC},{_RET}",
        f'$AH$1,$AH$1:$AH$5,{_RET},"n"',
    ):
        add(Case("xlookup", f"{xl}({args})"))
    for args in (
        f"3,{_ASC},3",
        f"3,{_ASC},0,0",
        f"3,{_ASC},0,3",
        "2,$AI$1:$AK$1",
        "2,{1,2,3}",
        "2,{1;2;3}",
        "3,$AF$1:$AG$6",
        "5,5",
        '"a","a"',
        f"{_NUM},{_ASC}",
        "1/0,{1,2}",
        f'3,{_ASC},"x"',
        f"3,{_ASC},0.5",
        f"3,{_ASC},1.9",
        f"3,{_ASC},,",
        f"3,{_ASC},,-1",
        f"$AH$1,{_NONE}",
        "1,$AA$1:$AA$8",
        "#N/A,$AB$1:$AB$4",
    ):
        add(Case("xmatch", f"{xm}({args})"))
    add(Case("xlookup", f"SUM({xl}(3,{_ASC},{_RET}):$AD$6)"))
    add(Case("xlookup", f"{xl}(3,{_ASC},{_RET})+1"))
    add(Case("xlookup", f"ROW({xl}(5,{_ASC},{_RET}))"))
    add(Case("xlookup", f"SUM({xl}(3,{_ASC},{_RET2}))"))
    add(Case("xlookup", f"INDEX({xl}(5,{_ASC},{_RET2}),1,2)"))
    add(Case("xlookup", f"COUNTA({xl}(3,{_ASC},{_RET2}))"))


def modern() -> list[Case]:
    cases: list[Case] = []
    _criteria_rows(cases.append)
    _order_rows(cases.append)
    _branch_rows(cases.append)
    _text_rows(cases.append)
    _lookup_rows(cases.append)
    # How text compares: every pair of these characters, then words around the ones that sort oddly.
    chars = list("~!_-'\"#$%&()*,./:;?@[\\]^`{|}+<=>09aAzZ ")
    for i, x in enumerate(chars):
        for y in chars[i + 1 :]:
            cases.append(Case("compare_text", "{1}<{2}", (x, y)))
    for x, y in (
        ("a-b", "ab"),
        ("a-b", "aa"),
        ("a-b", "ac"),
        ("a'b", "ab"),
        ("a'b", "aa"),
        ("a'b", "ac"),
        ("ab", "a b"),
        ("a b", "aa"),
        ("a_b", "ab"),
        ("a_b", "aa"),
        ("a~", "ab"),
        ("-a", "a"),
        ("-a", "b"),
        ("co-op", "coop"),
        ("coop", "co-op"),
        ("e", "é"),
        ("é", "f"),
        ("z", "é"),
        ("ae", "æ"),
        ("ß", "ss"),
        ("10", "9"),
        ("a10", "a9"),
        ("", " "),
        (" ", "!"),
    ):
        cases.append(Case("compare_text", "{1}<{2}", (x, y)))
        cases.append(Case("compare_text", "{1}={2}", (x, y)))
    return cases


_GUESSES: list[float] = [-0.99, -0.9, -0.5, -0.2, -0.1, -0.01, 0.0, 0.001, 0.01, 0.05,
                         0.1, 0.2, 0.3, 0.5, 1.0, 2.0, 5.0, 10.0, 100.0, 1000.0]  # fmt: skip
_FLOW_SHAPES: list[tuple[float, ...]] = [
    (-100.0, 110.0),
    (-100.0, 200.0),
    (-100.0, 50.0),
    (-100.0, 60.0, 60.0),
    (-1000.0, 100.0, 100.0, 100.0, 1100.0),
    (-1000.0, 300.0, 400.0, 500.0, 200.0, 100.0),
    (-100.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0, 110.0),
    (-1.0, 2.5, -1.5),  # two roots, 0 and 0.5
    (-1000.0, 5000.0, -6000.0, 2000.0),  # several roots
    (-100.0, 230.0, -132.0),  # two roots, 0.1 and 0.2
    (1.0, 1.0, 1.0),  # no root
    (-1.0, -1.0, -1.0),  # no root
    (-1e6, 3e5, 3e5, 3e5, 3e5),
    (-1e-6, 3e-7, 3e-7, 3e-7, 3e-7),
    (-100.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 200.0),
]
_RATE_SHAPES: list[tuple[float, float, float, float, float]] = [
    (12.0, -100.0, 1000.0, 0.0, 0.0),
    (12.0, -100.0, 1000.0, 0.0, 1.0),
    (360.0, -500.0, 100000.0, 0.0, 0.0),
    (10.0, 0.0, -100.0, 200.0, 0.0),
    (5.0, -100.0, 0.0, 600.0, 0.0),
    (2.0, -60.0, 100.0, 0.0, 0.0),
    (1.0, 0.0, -100.0, 110.0, 0.0),
    (24.0, -50.0, 1000.0, -200.0, 0.0),
    (12.0, 100.0, 1000.0, 0.0, 0.0),  # no root
    (0.5, -10.0, 100.0, 0.0, 0.0),
    (120.0, -1.0, 100.0, 0.0, 0.0),
    (3.0, -1e6, 2.5e6, 0.0, 0.0),
]
_XIRR_DATES = [39814.0 + 365.0 * k for k in range(5)]  # 2009 on: no leap day until 2012
_XIRR_SHAPES: list[tuple[tuple[float, ...], tuple[float, ...]]] = [
    ((-100.0, 110.0), (39814.0, 40179.0)),
    ((-100.0, 200.0), (39814.0, 40179.0)),
    ((-100.0, 50.0), (39814.0, 40179.0)),
    ((-100.0, 60.0, 60.0), (39814.0, 40179.0, 40544.0)),
    ((-1000.0, 300.0, 400.0, 500.0), (39814.0, 39900.0, 40100.0, 40400.0)),
    ((-1.0, 2.5, -1.5), (39814.0, 40179.0, 40544.0)),
    ((-100.0, 230.0, -132.0), (39814.0, 40179.0, 40544.0)),
    ((1.0, 1.0), (39814.0, 40179.0)),
    ((-100.0, 110.0), (39814.0, 39815.0)),
    ((-100.0, 110.0), (39814.0, 43000.0)),
]
# Flows with a double root at 0.1: Newton halves the error each step, so where a guess 2^k away
# turns into #NUM! shows the number of iterations and the stopping test.
_DOUBLE_ROOT = (-1.0, 2.2, -1.21)


def iteration() -> list[Case]:
    cases: list[Case] = []
    add = cases.append
    for flows in _FLOW_SHAPES:
        n = len(flows)
        for g in _GUESSES:
            add(Case("irr", f"IRR({{1}}:{{{n}}},{{{n + 1}}})", (*flows, g)))
    for nper, pmt, pv, fv, when in _RATE_SHAPES:
        for g in _GUESSES:
            add(Case("rate", "RATE({1},{2},{3},{4},{5},{6})", (nper, pmt, pv, fv, when, g)))
    for flows, dates in _XIRR_SHAPES:
        n = len(flows)
        f = f"XIRR({{1}}:{{{n}}},{{{n + 1}}}:{{{2 * n}}},{{{2 * n + 1}}})"
        for g in _GUESSES:
            add(Case("xirr", f, (*flows, *dates, g)))
    for k in range(-44, 5):
        for sign in (1.0, -1.0):
            g = 0.1 + sign * 2.0**k
            if g > -1:
                add(Case("irr double root", "IRR({1}:{3},{4})", (*_DOUBLE_ROOT, g)))
                add(Case("xirr double root", "XIRR({1}:{3},{4}:{6},{7})",
                         (*_DOUBLE_ROOT, *_XIRR_DATES[:3], g)))  # fmt: skip
                add(Case("rate double root", "RATE(2,2.2,-1,-1.21,0,{1})", (g,)))
    settle, maturity = 39814.0, 43466.0  # 2009-01-01 to 2019-01-01
    for frequency in (1.0, 2.0, 4.0):
        for basis in (0.0, 1.0, 3.0):
            for coupon in (0.0, 0.01, 0.05, 0.1, 0.5):
                for price in (10.0, 50.0, 90.0, 100.0, 110.0, 200.0, 1000.0):
                    add(Case("yield", "YIELD({1},{2},{3},{4},100,{5},{6})",
                             (settle, maturity, coupon, price, frequency, basis)))  # fmt: skip
    return cases


# Criteria and lookup keys that are blank, as a cell or as an empty text --------------------------

# T holds 0, 5, a blank, the empty text, text, 0 again, a logical and an error; U is what sums.
CRITERIA_TABLE: dict[str, Input] = {
    "T1": 0.0, "T2": 5.0, "T4": "", "T5": "a", "T6": 0.0, "T7": True, "T8": ErrorValue.NA,
    "U1": 1.0, "U2": 2.0, "U3": 4.0, "U4": 8.0, "U5": 16.0, "U6": 32.0, "U7": 64.0, "U8": 128.0,
    "X1": "", "X2": 0.0, "W1": "a",
    # AA and AB: keys and values, the value of key 1 blank; AC1:AD1 and AC2:AD2 the same across
    "AA1": 1.0, "AA2": 2.0, "AB2": 5.0, "AC1": 1.0, "AD1": 2.0, "AD2": 5.0,
}  # fmt: skip


def criteria() -> list[Case]:
    # Y1 is never written: a blank cell. X1 is the empty text, X2 is 0.
    range_t, range_u = "$T$1:$T$8", "$U$1:$U$8"
    keys = ["$Y$1", "$X$1", "$X$2", '""', '"="', '"<>"', "0", '"0"', '"=0"', "$W$1"]
    cases: list[Case] = []
    for key in keys:
        cases += [
            Case("countif", f"COUNTIF({range_t},{key})"),
            Case("sumif", f"SUMIF({range_t},{key},{range_u})"),
            Case("sumif_own", f"SUMIF($U$1:$U$8,{key})"),
            Case("averageif", f"AVERAGEIF({range_t},{key},{range_u})"),
            Case("countifs", f"COUNTIFS({range_t},{key})"),
            Case("sumifs", f"SUMIFS({range_u},{range_t},{key})"),
            Case("averageifs", f"AVERAGEIFS({range_u},{range_t},{key})"),
            Case("maxifs", f"_xlfn.MAXIFS({range_u},{range_t},{key})"),
            Case("minifs", f"_xlfn.MINIFS({range_u},{range_t},{key})"),
        ]
    cases += [
        Case("whole_column", "COUNTIF($T:$T,$Y$1)"),
        Case("all_blank", "COUNTIF($Z$1:$Z$9,$Y$1)"),
        Case("all_blank_sum", "SUMIF($Z$1:$Z$9,$Y$1,$U$1:$U$8)"),
        Case("second_criterion", 'COUNTIFS($T$1:$T$8,$Y$1,$U$1:$U$8,">0")'),
        Case("an_offset", "COUNTIF($T$1:$T$8,OFFSET($Y$1,0,0))"),
        Case("an_index", "COUNTIF($T$1:$T$8,INDEX($Y$1:$Y$2,1))"),
        Case("computed_blank", "COUNTIF($T$1:$T$8,IF(TRUE,$Y$1))"),
        Case("match_exact", "MATCH($Y$1,$T$1:$T$8,0)"),
        Case("match_sorted", "MATCH($Y$1,$T$1:$T$2,1)"),
        Case("vlookup", "VLOOKUP($Y$1,$T$1:$U$8,2,FALSE)"),
        Case("hlookup", "HLOOKUP($Y$1,$T$1:$U$1,1,FALSE)"),
        Case("xlookup", "_xlfn.XLOOKUP($Y$1,$T$1:$T$8,$U$1:$U$8)"),
        Case("xmatch", "_xlfn.XMATCH($Y$1,$T$1:$T$8)"),
    ]
    # What a lookup gives when it finds a blank cell.
    found = "VLOOKUP(1,$AA$1:$AB$2,2,FALSE)"
    cases += [
        Case("found_blank", found),
        Case("found_blank", f"ISBLANK({found})"),
        Case("found_blank", f"ISNUMBER({found})"),
        Case("found_blank", f"ISTEXT({found})"),
        Case("found_blank", f"ISLOGICAL({found})"),
        Case("found_blank", f"ISERROR({found})"),
        Case("found_blank", f"{found}+1"),
        Case("found_blank", f'{found}&""'),
        Case("found_blank", f"LEN({found})"),
        Case("found_blank", f'{found}=""'),
        Case("found_blank", f"{found}=0"),
        Case("found_blank", f"N({found})"),
        Case("found_blank", f'IF(ISBLANK({found}),"na",{found})'),
        Case("found_blank", "ISBLANK(HLOOKUP(1,$AC$1:$AD$2,2,FALSE))"),
        Case("found_blank", "ISBLANK(LOOKUP(1,$AA$1:$AA$2,$AB$1:$AB$2))"),
        Case("found_blank", "ISBLANK(_xlfn.XLOOKUP(1,$AA$1:$AA$2,$AB$1:$AB$2))"),
        Case("found_blank", "ISBLANK(INDEX($AB$1:$AB$2,1))"),
        Case("found_blank", "ISBLANK(CHOOSE(1,$AB$1))"),
        Case("found_blank", "ISBLANK(IF(TRUE,$AB$1))"),
        Case("found_blank", "ISBLANK(OFFSET($AB$1,0,0))"),
        Case("found_blank", "ISBLANK($AB$1)"),
        Case("found_blank", "ISBLANK($X$1)"),
        Case("found_blank", 'ISBLANK("")'),
        Case("found_blank", "ISBLANK(0)"),
        Case("found_blank", "ISBLANK(MATCH(1,$AA$1:$AA$2,0))"),
    ]
    # The range operator and intersection given errors, and a cell meeting itself.
    cases += [
        Case("range_error", "SUM(OFFSET($AB$2,NA(),0):$AB$2)"),
        Case("range_error", "SUM($AB$2:OFFSET($AB$2,NA(),0))"),
        Case("range_error", "SUM(OFFSET($AB$2,NA(),0):OFFSET($AB$2,1/0,0))"),
        Case("range_error", "SUM(OFFSET($AB$2,1/0,0):OFFSET($AB$2,NA(),0))"),
        Case("range_error", "ROWS(OFFSET($AB$2,NA(),0):$AB$2)"),
        Case("range_error", "COUNT(OFFSET($AB$2,NA(),0):$AB$2)"),
        Case("range_error", "SUM(INDEX($AA$1:$AA$2,3):$AA$2)"),
        Case("range_error", "CORREL(OFFSET($AA$1,NA(),0):$AA$2,$AA$1:$AA$2)"),
        Case("meets_itself", "$AA$2 $AA$2"),
        Case("meets_itself", "SUM($AA$2 $AA$2)"),
        Case("meets_itself", 'CONCATENATE("x",$AA$2 $AA$2)'),
        Case("meets_itself", 'CONCATENATE("x",$AA$1:$AB$2 $AA$2)'),
        Case("meets_itself", "$AA$1 $AA$2"),
        Case("meets_itself", "SUM($AA$1 $AA$2)"),
        Case("meets_itself", "SUM($AA$1:$AA$2 $AA$2:$AB$2)"),
        Case("meets_itself", "SUM($AA$1:$AA$2 $AB$1:$AB$2)"),
    ]
    return cases


# The absent built-ins, generated per family -----------------------------------------------------

_EDGES: tuple[Input, ...] = (
    0.0, 1.0, -1.0, 0.5, -0.5, 2.0, -2.0, 0.1, 3.7, -3.7, 10.0, 100.0, 1e-10, 1e-300, 1e6, 1e15,
    1e100,
    1e308, 0.7853981633974483, 1.5707963267948966, 3.141592653589793, 6.283185307179586,
    "3", "abc", "", None, True, False, NA,
)  # fmt: skip
_PAIRS: tuple[Input, ...] = (
    0.0,
    1.0,
    -1.0,
    0.5,
    2.0,
    -2.5,
    3.7,
    10.0,
    1e15,
    "2",
    "abc",
    None,
    True,
    NA,
)


def _file(name: str) -> str:
    """A function's name as a file holds it (capture ``prefixes``)."""
    return PREFIXES.get(name, "") + name


def _one(names: list[str], values: tuple[Input, ...] = _EDGES) -> list[Case]:
    return [Case(n.lower(), f"{_file(n)}({{1}})", (v,)) for n in names for v in values]


def _two(names: list[str], values: tuple[Input, ...] = _PAIRS) -> list[Case]:
    return [
        Case(n.lower(), f"{_file(n)}({{1}},{{2}})", (a, b))
        for n in names
        for a in values
        for b in values
    ]


def math_trig() -> list[Case]:
    names = "ACOS ACOSH ACOT ACOTH ASIN ASINH ATAN ATANH COS COSH COT COTH CSC CSCH SEC SECH SIN"
    names += " SINH TAN TANH DEGREES RADIANS SIGN TRUNC LOG LOG10 SQRTPI"
    cases = _one(names.split())
    cases += _two(["ATAN2", "LOG", "TRUNC"])
    cases += [
        Case("trunc_digits", "TRUNC({1},{2})", (x, d))
        for x in (1234.5678, -1234.5678, 0.1 + 0.2, 2.675, 1e15 + 0.5, 1e-5)
        for d in (-5.0, -2.0, -0.5, 0.0, 0.5, 1.0, 2.0, 3.7, 10.0, 15.0, 20.0, 400.0, -400.0)
    ]
    cases += [
        Case("exact", "EXACT({1},{2})", (a, b))
        for a in ("a", "A", "", " a", "1", 1.0, True, None, NA)
        for b in ("a", "A", "", "a ", "1", 1.0, True, None, NA)
    ]
    cases += [
        Case("time", "TIME({1},{2},{3})", (h, m, s))
        for h in (0.0, 1.0, 12.5, 23.0, 24.0, 25.0, -1.0, 32767.0, 32768.0, "3", None, True, NA)
        for m in (0.0, 30.0, 59.0, 60.0, -30.0, 1.9, 32768.0)
        for s in (0.0, 59.0, 90.0, -5.0, 0.5, 32768.0)
    ]
    return cases


def math_integers() -> list[Case]:
    multiples = (
        0.0,
        1.0,
        -1.0,
        0.5,
        -0.5,
        2.0,
        -2.0,
        0.1,
        -0.1,
        3.7,
        1e15,
        "2",
        "abc",
        None,
        True,
        NA,
    )
    numbers = (
        0.0,
        1.0,
        -1.0,
        2.5,
        -2.5,
        3.0,
        6.7,
        -6.7,
        0.1,
        0.3,
        1.5,
        1e15,
        "2",
        "abc",
        None,
        True,
        NA,
    )
    cases = [
        Case(n.lower(), f"{_file(n)}({{1}},{{2}})", (x, m))
        for n in ("CEILING", "FLOOR", "CEILING.PRECISE", "FLOOR.PRECISE", "ISO.CEILING", "MROUND")
        for x in numbers
        for m in multiples
    ]
    cases += _one(["CEILING.PRECISE", "FLOOR.PRECISE", "ISO.CEILING", "CEILING.MATH", "FLOOR.MATH"])
    cases += [
        Case(n.lower(), f"{_file(n)}({{1}},{{2}},{{3}})", (x, m, mode))
        for n in ("CEILING.MATH", "FLOOR.MATH")
        for x in (2.5, -2.5, 6.7, -6.7, 0.0)
        for m in (1.0, -1.0, 2.0, 0.0, 0.5, None)
        for mode in (0.0, 1.0, -1.0, 0.5, None, True, "abc")
    ]
    cases += _one(["EVEN", "ODD", "FACT", "FACTDOUBLE", "ARABIC"])
    cases += [Case("fact", "FACT({1})", (x,)) for x in (5.0, 5.9, 20.0, 21.0, 170.0, 171.0)]
    cases += [
        Case("factdouble", "FACTDOUBLE({1})", (x,)) for x in (5.0, 6.0, 6.9, 20.0, 300.0, 301.0)
    ]
    cases += _two(["QUOTIENT", "GCD", "LCM", "COMBIN", "COMBINA", "PERMUT", "PERMUTATIONA"])
    cases += [
        Case(n.lower(), f"{_file(n)}({{1}},{{2}})", (a, b))
        for n in ("COMBIN", "COMBINA", "PERMUT", "PERMUTATIONA", "GCD", "LCM", "QUOTIENT")
        for a, b in ((5.0, 2.0), (10.0, 3.0), (5.9, 2.9), (60.0, 30.0), (1030.0, 515.0),
                     (12.0, 18.0),
                     (2.0**53, 3.0), (1e10, 7.0))
    ]  # fmt: skip
    cases += [
        Case("multinomial", "MULTINOMIAL({1},{2},{3})", (a, b, c))
        for a in (1.0, 2.0, 2.9, 0.0, -1.0, "2", None, True, NA)
        for b in (3.0, 0.0)
        for c in (4.0, 200.0)
    ]
    cases += [
        Case("roman", "ROMAN({1},{2})", (x, form))
        for x in (
            0.0,
            1.0,
            4.0,
            49.0,
            99.0,
            499.0,
            999.0,
            1999.0,
            3999.0,
            4000.0,
            -1.0,
            2.9,
            "3",
            None,
            NA,
        )
        for form in (0.0, 1.0, 2.0, 3.0, 4.0, 5.0, True, False, None)
    ]
    cases += _one(["ROMAN"])
    cases += [
        Case("arabic", "_xlfn.ARABIC({1})", (x,))
        for x in (
            "I",
            "IV",
            "MCMXCIX",
            "mcmxcix",
            "  XI ",
            "-X",
            "IIII",
            "VX",
            "IC",
            "MMMM",
            "Z",
            "-",
            "",
        )
    ]
    cases += [
        Case("base", "_xlfn.BASE({1},{2},{3})", (x, r, n))
        for x in (0.0, 7.0, 255.0, 2.0**53 - 1, 2.0**53, -1.0, 7.9, "3", None, NA)
        for r in (2.0, 16.0, 36.0, 1.0, 37.0, 16.9)
        for n in (0.0, 5.0, 255.0, 256.0, -1.0)
    ]
    cases += [
        Case("decimal", "_xlfn.DECIMAL({1},{2})", (x, r))
        for x in ("0", "FF", "ff", " ff", "zz", "1E", "12", 12.0, "", "-1", "1.5", "G", None, NA)
        for r in (2.0, 16.0, 36.0, 1.0, 37.0, 16.9)
    ]
    cases += [
        Case("seriessum", "SERIESSUM({1},{2},{3},$AA$1:$AA$3)", (x, n, m))
        for x in (0.0, 1.0, 2.0, 0.5, "abc", None, NA)
        for n in (0.0, 1.0, 2.5)
        for m in (0.0, 1.0, -1.0)
    ]
    for name in ("SUMX2MY2", "SUMX2PY2", "SUMXMY2"):
        cases += [
            Case(name.lower(), f"{name}($AA$1:$AA$3,$AB$1:$AB$3)"),
            Case(name.lower(), f"{name}($AA$1:$AA$4,$AB$1:$AB$4)"),
            Case(name.lower(), f"{name}($AA$1:$AA$3,$AB$1:$AB$4)"),
            Case(name.lower(), f"{name}($AA$5:$AA$6,$AB$5:$AB$6)"),
            Case(name.lower(), f"{name}({{1,2}},{{3,4}})"),
        ]
    return cases


_INTEGERS_TABLE: dict[str, Input] = {
    "AA1": 1.0, "AA2": 2.0, "AA3": 3.0, "AA4": "x", "AA5": NA, "AA6": 2.0,
    "AB1": 4.0, "AB2": 5.0, "AB3": 6.0, "AB4": 7.0, "AB5": 1.0, "AB6": 3.0,
}  # fmt: skip


SUITES = {
    "number_text": Suite(number_text),
    "text_number": Suite(text_number),
    "compare": Suite(compare),
    "cancellation": Suite(cancellation),
    "dates": Suite(dates),
    "dates_1904": Suite(dates, date1904=True),
    "sums": Suite(sums),
    "functions": Suite(functions, fixed=TABLE),
    "literals": Suite(literals),
    "n_lookup": Suite(n_lookup, fixed=TABLE),
    "finance": Suite(finance, fixed=FIN_TABLE),
    "iterative": Suite(iterative, fixed={**FIN_TABLE, **ITER_TABLE}),
    "bonds": Suite(bonds),
    "workdays": Suite(workdays, fixed=WORK_TABLE),
    "statistics": Suite(statistics, fixed=STAT_TABLE),
    "misc": Suite(misc, fixed=MISC_TABLE),
    "textfmt": Suite(textfmt),
    "rounding": Suite(rounding),
    "modern": Suite(modern, fixed=MODERN_TABLE),
    "criteria": Suite(criteria, fixed=CRITERIA_TABLE),
    "math_trig": Suite(math_trig),
    "math_integers": Suite(math_integers, fixed=_INTEGERS_TABLE),
    "iteration": Suite(iteration),
}
