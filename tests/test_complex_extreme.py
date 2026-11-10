"""The complex-number functions on extreme input: a value, never an exception.

Parts from the smallest subnormal to the largest double, both signs and suffixes, malformed and
long texts, and powers up to 1.7E308. Each formula gives text, a finite number or an Excel error,
none raises, and the whole sweep is quick: a whole power is at most 53 squarings, and nothing
else loops over its input but the parse of one text. Run under the watchdog.
"""

from __future__ import annotations

import itertools
import math
import time

from tests.watchdog import bounded

_PARTS = (
    "0",
    "5E-324",
    "1E-300",
    "1E-10",
    "0.5",
    "1",
    "-2.5",
    "1E15",
    "1E20",
    "1E21",
    "1E300",
    "1.7E308",
)
_TEXTS = (
    *_PARTS,
    *(f"{a}+{b}i" for a, b in itertools.product(("1E-300", "1", "1E300", "1.7E308"), repeat=2)),
    "-1.7E308-1.7E308j", "5E-324+5E-324i", "i", "-j", "1e400", "1+2i+3", "abc", "", "  ",
    "1" * 300, "1" * 300 + "+" + "2" * 300 + "i",
)  # fmt: skip
_POWERS = (
    "0", "0.5", "-0.5", "1.5", "-3", "1000", "-1000", "2147483649", "1E15", "9007199254740993",
    "1E300", "-1E300", "1.7E308", "5E-324",
)  # fmt: skip
_UNARY = (
    "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMEXP", "IMLN", "IMLOG10",
    "IMLOG2", "IMREAL", "IMSIN", "IMSQRT", "_xlfn.IMCOSH", "_xlfn.IMCOT", "_xlfn.IMCSC",
    "_xlfn.IMCSCH", "_xlfn.IMSEC", "_xlfn.IMSECH", "_xlfn.IMSINH", "_xlfn.IMTAN",
)  # fmt: skip
_PAIRED = ("IMSUB", "IMDIV", "IMSUM", "IMPRODUCT")
_SECONDS = 60.0  # the whole sweep; it takes a few seconds


def _formulas() -> list[str]:
    texts = [f'"{t}"' for t in _TEXTS]
    out = [f"{name}({t})" for name in _UNARY for t in texts]
    out += [f"IMPOWER({t},{n})" for t in texts for n in _POWERS]
    out += [f"{name}({t},{u})" for name in _PAIRED for t in texts for u in texts]
    out += [f'COMPLEX({a},{b},"{s}")' for a in _PARTS for b in _PARTS for s in "ij"]
    out += [f"COMPLEX(-{a},-{b})" for a in _PARTS for b in _PARTS]
    return out


@bounded(seconds=300.0)
def test_every_formula_is_a_value_and_the_sweep_is_quick():
    from xlrecalc._engine import Engine
    from xlrecalc._values import ErrorValue

    formulas = _formulas()
    engine = Engine()
    sheet = engine.add_sheet("s")
    for row, formula in enumerate(formulas):
        engine.set_formula(sheet, row, 0, formula)
    start = time.perf_counter()
    engine.recalculate(full=True)
    elapsed = time.perf_counter() - start
    assert not engine.report.exceptions, list(engine.report.exceptions.items())[:5]
    bad = []
    for row, formula in enumerate(formulas):
        value = engine.get(sheet, row, 0)
        if not (
            isinstance(value, (ErrorValue, str)) or (type(value) is float and math.isfinite(value))
        ):
            bad.append((formula[:60], repr(value)[:40]))
    assert not bad, bad[:10]
    assert elapsed < _SECONDS, f"{len(formulas)} formulas in {elapsed:.1f} s"
