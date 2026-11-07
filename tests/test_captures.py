"""The engine against results captured from Excel (``tests/captures``, ``CAPTURING.md``).

Each suite becomes one engine sheet laid out as Excel's workbook was: case k's formula in
A{k}, its inputs from B{k} on, and the suite's fixed cells. The engine recalculates everything;
each formula cell must equal what Excel saved, kind and value, bit for bit. A row calling a
function the engine does not implement yet is counted as not evaluated, never skipped
silently; known deviations need a reason.
"""

from __future__ import annotations

import functools
import math
import re
import struct
from pathlib import Path

import pytest
import xlformat

from tests.capture_report import REPORT
from tests.captures import Formula, Row, Suite, load
from xlrecalc._engine import Engine
from xlrecalc._formula._template import cell_axes
from xlrecalc._functions import SPECS
from xlrecalc._numtext import DBL_MIN
from xlrecalc._quirks import Quirks
from xlrecalc._values import Value

# (suite, case) -> reason.
_IPMT_TAIL = (
    "IPMT at a fractional period or late in a long term: Excel's balance computation is not "
    "reproduced in the last bits (13 to 109 units in the last place); other rows within 4"
)
# IRR, XIRR and RATE: Excel's Newton iterates rest on its own NPV, whose last bits are not
# reproduced (see NPV's tolerance), so where its iteration stops differs. Tried (at most five
# hypotheses each): Newton on the rate, on 1 + rate and on 1 / (1 + rate), analytic and
# numerical derivatives, stopping on the step (absolute, relative, returning either iterate)
# or on the NPV, 20 to 100 iterations, damping. Kept: Newton on 1 + rate, halving instead of
# stepping to 1 + rate <= 0, a step below 1E-7.
_ITER_NEAR_ZERO = (
    "IRR where the rate is 0: Excel's last iterate is within 1E-15 of 0 (-6.7E-16, 1.1E-14) "
    "or exactly 0, where a relative tolerance cannot apply"
)
_ITER_OUTCOME = (
    "IRR, RATE: from this guess Excel's iteration reaches a different result (another root, a "
    "rate where this iteration gives #NUM!, or the reverse; a double root left near 1E-7)"
)
_XIRR = (
    "XIRR: Excel stops about 1E-8 (relative) from the root, sometimes after a bracketing "
    "search (results that are short binary fractions), and gives #NUM! for some flows a few "
    "days apart; not reproduced"
)
_ITER_NEAR_ZERO_CASES = (77, 82, 84)
_ITER_OUTCOME_CASES = (
    5,
    16,
    18,
    24,
    34,
    35,
    37,
    41,
    42,
    44,
    45,
    48,
    49,
    51,
    63,
    64,
    65,
    66,
    67,
    68,
    70,
    71,
    72,
    73,
    74,
    79,
    110,
    112,
    123,
    128,
    186,
    187,
    192,
    197,
    200,
    201,
    209,
    214,
    240,
    249,
    252,
    254,
    269,
)
_XIRR_CASES = (
    130,
    131,
    132,
    133,
    134,
    135,
    136,
    137,
    138,
    140,
    141,
    142,
    143,
    148,
    155,
    157,
    161,
    163,
    165,
    166,
    170,
    176,
    178,
    179,
    180,
    181,
    182,
)
_YIELD_ZERO_COUPON = (
    "YIELD of a zero-coupon bond over 20 periods: Excel's iteration stops 1.3E-9 (relative) "
    "from this one, just beyond the relative tolerance"
)
_FEB29_MATURITY = (
    "a bond maturing on February 29: Excel's COUPDAYS in basis 1 counts 366 days for the "
    "period from February 28 of the year before, where the other captured periods follow "
    "EDATE from the previous coupon date"
)
_PRICE_EXTREME = (
    "PRICE at a yield of 1000%: 64 units in the last place from Excel, beyond its tolerance"
)
_LINEST_CANCEL = (
    "LINEST near cancellation (an intercept of -0.02 from means of 7.05 and 3.5, and the "
    "statistics built on it; R squared of values near 1E9): Excel's regression differs by 26 "
    "to 4,607 units in the last place (plain and compensated means, two intercept and "
    "sum-of-squares forms tried)"
)
_NORMAL_TAIL = (
    "the normal distribution's lower tail (z of -2.3 and below): 5 to 80 units in the last "
    "place from Excel, which is closer to the true value there than math.erfc and not Cody's "
    "algorithm either"
)
_VDB_ROUNDING = (
    "VDB over the last period of a long schedule: the book value carried over nine periods "
    "rounds differently from Excel's (7 to 48 units in the last place); a closed form fits "
    "one row and a difference of cumulative values another"
)
_VDB_EDGES = (
    "VDB: Excel gives #DIV/0! for start = end = life at a rate of 1 per period, and "
    "cost - salvage (a negative value) when salvage exceeds cost; not reproduced"
)
_RANDOM = "a random result: RANDBETWEEN(1.5,2.5) is 2 or 3, and two random draws compared"
_COMPUTED_CRITERIA = (
    "a computed criterion (a formula under a name the database lacks, evaluated for each "
    "record) is not supported; it matches no record"
)
_UNION = "a union of ranges is not implemented, and the cell is reported"
_REGEX = "the regular-expression match mode (3) is not implemented, and the cell is reported"
_AT_SIGN = "the @ the formula bar shows in a legacy formula is not shown by FORMULATEXT"
_WEEKNUM_ODD_TYPE = (
    "WEEKNUM with a type its documentation lacks (-1, 22): #NUM! on some weekdays and a week "
    "number on others, with no rule found, where the package gives #NUM!"
)
_DATEDIF_MARCH = (
    "DATEDIF from December 31, 1900 to March 1, 2000 in YD is 61 where every neighbor gives 60"
)
_WEEKNUM_ODD_CASES = (
    476,
    497,
    520,
    543,
    591,
    612,
    614,
    635,
    681,
    704,
    727,
    750,
    773,
    775,
    821,
    865,
    867,
    913,
    934,
    936,
    957,
    1003,
    1026,
    1118,
    1166,
    1189,
    1212,
    1235,
    1256,
    1258,
    1279,
    1302,
    1304,
    1325,
    1327,
    1348,
)
_MULTINOMIAL = (
    "Excel's result carries noise from its logarithm of the gamma function (279.99999999999966)"
)
KNOWN_DEVIATIONS: dict[tuple[str, int], str] = {
    **{
        ("math_integers", case): _MULTINOMIAL
        for case in (3783, 3785, 3787, 3789, 3791, 3793, 3795, 3797, 3803, 3805, 3807, 3809)
    },
    **{("modern", case): _UNION for case in (300, 328)},
    **{
        ("info_family", case): _UNION
        for case in (272, 303, 334, 396, 427, 452, 475, 484, 485, 505, 506)
    },
    ("info_family", 468): _AT_SIGN,
    **{("dates_family", case): _WEEKNUM_ODD_TYPE for case in _WEEKNUM_ODD_CASES},
    ("dates_family", 2004): _DATEDIF_MARCH,
    **{("modern", case): _REGEX for case in (1384, 1417)},
    **{("finance", case): _IPMT_TAIL for case in (359, 367, 368, 371, 373, 397)},
    **{("iterative", case): _ITER_NEAR_ZERO for case in _ITER_NEAR_ZERO_CASES},
    **{("iterative", case): _ITER_OUTCOME for case in _ITER_OUTCOME_CASES},
    **{("iterative", case): _XIRR for case in _XIRR_CASES},
    ("bonds", 2299): _YIELD_ZERO_COUPON,
    **{("bonds", case): _FEB29_MATURITY for case in (1999, 2004, 2006, 2007)},
    ("bonds", 2295): _PRICE_EXTREME,
    **{
        ("statistics", case): _LINEST_CANCEL
        for case in (489, 499, 519, 520, 521, 523, 524, 527, 539, 551)
    },
    **{("finance", case): _VDB_ROUNDING for case in (785, 786, 787)},
    **{("finance", case): _VDB_EDGES for case in (903, 920)},
    **{("misc", case): _RANDOM for case in (183, 188)},
    **{("misc", case): _COMPUTED_CRITERIA for case in (197, 222, 247, 272, 297)},
    **{
        ("statistics", case): _NORMAL_TAIL
        for case in (609, 610, 612, 613, 615, 616, 660, 661, 675, 676, 770)
    },
}
# A listed deviation that matches Excel fails the test, so a stale entry cannot survive, with two
# exceptions. A random result can match by chance. And a row listed because the platform's math
# library (exp, log1p, expm1, pow, erfc) differs from Excel's in the last bits may match on another
# platform's library: it must still differ on the platform the captures' split is measured on
# (docs/coverage.md records it), and elsewhere it counts as listed whether or not it matches, so
# the published counts are the same on every platform.
EXEMPT = frozenset({_RANDOM})
MATH_LIBRARY = frozenset(
    {_IPMT_TAIL, _ITER_NEAR_ZERO, _ITER_OUTCOME, _XIRR, _YIELD_ZERO_COUPON, _PRICE_EXTREME,
     _NORMAL_TAIL}
)  # fmt: skip


def _reference_platform() -> bool:
    from tools.coverage import OUTPUT, platform_label, recorded_platform

    return recorded_platform(OUTPUT.read_text(encoding="utf-8")) == platform_label()


SUITES = [
    "number_text",
    "text_number",
    "compare",
    "cancellation",
    "dates",
    "dates_1904",
    "sums",
    "functions",
    "n_lookup",
    "literals",
    "finance",
    "iterative",
    "bonds",
    "workdays",
    "statistics",
    "misc",
    "textfmt",
    "rounding",
    "modern",
    "criteria",
    "math_trig",
    "math_integers",
    "info_family",
    "dates_family",
]


def column_letters(n: int) -> str:
    letters = ""
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def formula_text(row: Row, k: int) -> str:
    return re.sub(r"\{(\d+)\}", lambda m: f"{column_letters(1 + int(m[1]))}{k}", row.formula)


def same(a: Value, b: Value) -> bool:
    """Equal and of the same kind: the text "#N/A" is not the error, 1.0 is not TRUE."""
    if type(a) is not type(b):
        return False
    if isinstance(a, float) and isinstance(b, float):
        return a == b or (math.isnan(a) and math.isnan(b))
    return a == b


_CALL = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")


# Where a few units in the last place can change a visible digit or a decision: a row that
# rounds, compares or formats a tolerated result must match exactly.
_EXACTING = frozenset({
    "ROUND", "ROUNDUP", "ROUNDDOWN", "MROUND", "INT", "TRUNC", "CEILING", "FLOOR",
    "CEILING.MATH", "FLOOR.MATH", "_XLFN.CEILING.MATH", "_XLFN.FLOOR.MATH", "EVEN", "ODD",
    "TEXT", "FIXED", "DOLLAR", "VALUE", "MATCH", "VLOOKUP", "HLOOKUP", "LOOKUP", "COUNTIF",
    "COUNTIFS", "SUMIF", "SUMIFS", "IF", "RANK", "EXACT",
})  # fmt: skip
_STRING = re.compile(r'"(?:[^"]|"")*"')
_OPERATOR = re.compile(r"[=<>&]")


def _exacting(formula: str, names: list[str]) -> bool:
    return bool(_EXACTING.intersection(names) or _OPERATOR.search(_STRING.sub("", formula)))


def tolerance(formula: str) -> int:
    """The largest tolerance, in units in the last place, of the functions a formula calls.

    0 when the formula also rounds, compares or formats (a comparison operator, ``&``, or a
    function in ``_EXACTING``): a tolerated result there must match exactly.
    """
    names = [n.upper() for n in _CALL.findall(formula)]
    ulps = max((SPECS[n].ulps for n in names if n in SPECS), default=0)
    return 0 if ulps and _exacting(formula, names) else ulps


def relative(formula: str) -> float:
    """The largest relative tolerance of the functions a formula calls, with the same guard."""
    names = [n.upper() for n in _CALL.findall(formula)]
    rel = max((SPECS[n].rel for n in names if n in SPECS), default=0.0)
    return 0.0 if rel and _exacting(formula, names) else rel


def within_relative(got: Value, expected: Value, rel: float) -> bool:
    return (
        rel > 0
        and type(got) is float
        and type(expected) is float
        and math.isfinite(got)
        and math.isfinite(expected)
        and abs(got - expected) <= rel * abs(expected)
    )


def ulps_apart(a: float, b: float) -> int:
    """How many representable steps separate two finite doubles of the same sign."""
    ia, ib = (struct.unpack("<q", struct.pack("<d", x))[0] for x in (a, b))
    return abs(ia - ib)


def within(got: Value, expected: Value, ulps: int) -> bool:
    return (
        ulps > 0
        and type(got) is float
        and type(expected) is float
        and math.isfinite(got)
        and math.isfinite(expected)
        and (got < 0) == (expected < 0)
        and ulps_apart(got, expected) <= ulps
    )


def _input(value: Value) -> Value:
    if isinstance(value, float) and -DBL_MIN < value < DBL_MIN:
        return 0.0  # Excel reads a subnormal number as 0
    return value


def run(suite: Suite) -> tuple[Engine, int]:
    engine = Engine(Quirks.EXCEL, date1904=suite.date1904, today=suite.captured)
    s = engine.add_sheet("cases")
    for ref, value in suite.fixed.items():
        r, c = cell_axes(ref, 0, 0)
        if isinstance(value, Formula):
            engine.set_formula(s, r.value, c.value, value.text[1:])
        else:
            engine.sheets[s].set(r.value, c.value, _input(value))
    for k, row in enumerate(suite.rows, 1):
        for i, value in enumerate(row.inputs, 1):
            if isinstance(value, Formula):
                engine.set_formula(s, k - 1, i, value.text[1:])
            else:
                engine.sheets[s].set(k - 1, i, _input(value))
        engine.set_formula(s, k - 1, 0, formula_text(row, k))
    engine.recalculate(full=True)
    return engine, s


@functools.cache
def outcomes(name: str) -> tuple[dict[str, int], dict[str, int], tuple[str, ...]]:
    """A suite's rows by outcome, why rows were not evaluated, and the rows that differ."""
    suite = load(name)
    engine, s = run(suite)
    counts = dict.fromkeys(("match", "ulps", "rel", "listed", "not evaluated"), 0)
    reasons: dict[str, int] = {}
    failures = []
    reference = _reference_platform()
    for k, row in enumerate(suite.rows, 1):
        template = engine.templates[engine.sheets[s].templates[(k - 1) * 16384]]
        if template.unsupported:
            counts["not evaluated"] += 1
            why = "not implemented: " + ", ".join(sorted(template.unsupported))
            reasons[why] = reasons.get(why, 0) + 1
            continue
        got = engine.get(s, k - 1, 0)
        if same(got, row.expected):
            outcome = "match"
        elif within(got, row.expected, tolerance(row.formula)):
            outcome = "ulps"
        elif within_relative(got, row.expected, relative(row.formula)):
            outcome = "rel"
        else:
            outcome = "differs"
        reason = KNOWN_DEVIATIONS.get((name, row.case))
        if reason is not None:
            may_match = reason in EXEMPT or (reason in MATH_LIBRARY and not reference)
            if outcome == "differs" or may_match:
                counts["listed"] += 1
            else:
                failures.append(
                    f"{row.id} ={formula_text(row, k)}: listed as a known deviation but matches "
                    f"Excel ({outcome}): remove it from KNOWN_DEVIATIONS and QUIRKS.md"
                )
        elif outcome != "differs":
            counts[outcome] += 1
        else:
            failures.append(
                f"{row.id} ={formula_text(row, k)} {row.inputs!r}: "
                f"Excel {row.expected!r}, got {got!r}"
            )
    return counts, reasons, tuple(failures)


@pytest.mark.excel
@pytest.mark.parametrize("name", SUITES)
def test_suite(name):
    counts, reasons, failures = outcomes(name)
    rows = sum(counts.values()) + len(failures)
    detail = "; ".join(f"{v} {k}" for k, v in reasons.items())
    REPORT[name] = (
        f"{name}: {rows} rows, {counts['match']} match, "
        + (f"{counts['ulps']} within a declared tolerance, " if counts["ulps"] else "")
        + (f"{counts['rel']} within a relative tolerance, " if counts["rel"] else "")
        + f"{len(failures)} differ, "
        f"{counts['listed']} known deviations, {counts['not evaluated']} not evaluated"
        + (f" ({detail})" if detail else "")
    )
    assert not failures, f"{len(failures)} rows differ:\n" + "\n".join(failures[:80])


def _figure(text: str) -> int:
    return int(text.replace(",", ""))


@pytest.mark.parametrize("page", ["README.md"])
def test_stated_evidence_is_the_capture_run(page):
    """The evidence figures the README states, from the capture run.

    The changelog's figures are those of their release, and stay as they were.

    Only counts that hold on every platform are published: a row within its declared tolerance
    on one platform may be identical on another, whose math library differs in the last bits.
    """
    total = dict.fromkeys(("match", "ulps", "rel", "listed", "not evaluated"), 0)
    for name in SUITES:
        counts, _, failures = outcomes(name)
        assert not failures
        for key, n in counts.items():
            total[key] += n
    text = " ".join((Path(__file__).parents[1] / page).read_text(encoding="utf-8").split())
    found = re.search(
        r"([\d,]+) formulas computed by Excel.*?([\d,]+) match.*?"
        r"([\d,]+) known deviations.*?([\d,]+) not evaluated",
        text,
    )
    assert found, f"{page}: the evidence sentence is missing"
    stated, matched, listed, skipped = map(_figure, found.groups())
    assert matched == total["match"] + total["ulps"] + total["rel"]
    assert listed == total["listed"]
    assert skipped == total["not evaluated"]
    assert stated == matched + listed + skipped == sum(total.values())


def test_quirks_presets():
    assert Quirks() == Quirks.EXCEL
    assert not any(
        getattr(Quirks.STRICT, f)
        for f in (
            "leap_year_1900",
            "fifteen_digit_precision",
            "comparison_rounding",
            "cancellation_snapping",
        )
    )
    assert Quirks.EXCEL.to_xlformat() == xlformat.Quirks.EXCEL


def test_tolerance_needs_reason():
    from xlrecalc._functions import function

    with pytest.raises(ValueError, match="reason"):
        function("_TEST_TOLERANCE", "number", ulps=2)


def test_relative_tolerance():
    assert within_relative(0.1, 0.1 * (1 + 5e-10), 1e-9)
    assert not within_relative(0.1, 0.1 * (1 + 5e-9), 1e-9)
    assert relative("IRR(A1:A5)") == 1e-9
    assert relative("ROUND(IRR(A1:A5),4)") == 0.0
    for name, spec in SPECS.items():
        if spec.rel:
            assert spec.rel <= 1e-9, name
            assert spec.rel_reason, name


def test_tolerances_are_small_and_explained():
    tolerant = {n: s for n, s in SPECS.items() if s.ulps}
    assert tolerant, "the financial functions declare tolerances"
    for name, spec in tolerant.items():
        assert 0 < spec.ulps <= 8, name
        assert spec.ulps_reason, name


def test_within_tolerance():
    assert within(1.0, math.nextafter(1.0, 2.0), 1)
    assert not within(1.0, math.nextafter(1.0, 2.0), 0)
    assert not within(1.0, 1.0 + 4 * 2.0**-52, 2)
    assert not within(-0.0, 0.0, 4)
    assert tolerance("SUM(PMT(0.05,10,1000))") == 4
    assert tolerance("PV(0.05,10,1000)") == 0
    assert tolerance("ROUND(PMT(0.05,10,1000),2)") == 0
    assert tolerance("PMT(0.05,10,1000)>-129") == 0
    assert tolerance('PMT(0.05,10,1000)&""') == 0
    assert tolerance('TEXT(NPV(0.1,1,2),"0.00")') == 0
    assert tolerance('NPV(0.1,1,2)+LEN("a=b")') == 4


def test_status_needs_note():
    from xlrecalc._functions import function

    with pytest.raises(ValueError, match="partial"):
        function("_TEST_STATUS", "number", status="partial")
    with pytest.raises(ValueError, match="partial"):
        function("_TEST_STATUS", "number", status="beta", note="x")


def test_every_listed_deviation_is_a_captured_row():
    rows = {(name, row.case) for name in SUITES for row in load(name).rows}
    assert sorted(set(KNOWN_DEVIATIONS) - rows) == []


def test_exemptions_are_listed_reasons():
    assert (EXEMPT | MATH_LIBRARY).issubset(KNOWN_DEVIATIONS.values())
