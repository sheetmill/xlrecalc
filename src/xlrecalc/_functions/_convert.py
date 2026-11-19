"""CONVERT (capture ``convert_family``): a number from one unit of measure to another.

Each unit is in one category (mass, distance, time, pressure, force, energy, power, magnetism,
volume, area, information, speed; temperature apart) and holds two of Excel's own doubles: what
one unit is in the category's base unit, and what one base unit is in it (``CONVERT(1,u,base)``
and ``CONVERT(1,base,u)`` as captured; for 27 units they are not each other's reciprocal in
doubles). A conversion to or from the base uses that factor; between two other units it is
``number * to_base(from) / to_base(to)``, which differs from Excel's result in the last bit or
two for about a fifth of such conversions (QUIRKS.md): Excel likely defines units through other
units, rounding at each step.

A decimal prefix (``k``, ``m``, ``da``, ``e`` for 10...) multiplies the 52 units that take one,
once per power for an area or volume (``km2`` is 1E6 m2); a binary prefix (``ki`` to ``Yi``)
only ``bit`` and ``byte``. A name is matched exactly, case included, first, then as a prefix and
a unit. Temperatures convert by Excel's formula for each pair of scales, not through a common
one.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable
from typing import Any

from xlrecalc._coerce import Context
from xlrecalc._functions._engineering import _strict
from xlrecalc._functions._registry import function
from xlrecalc._values import ErrorValue

NA, NUM = ErrorValue.NA, ErrorValue.NUM

# name: (category, the unit in the base unit, the base unit in the unit), as Excel computes them
_UNITS: dict[str, tuple[str, float, float]] = {
    # g
    "g": ("g", 1.0, 1.0),
    "sg": ("g", 14593.902937206363, 6.852176585679177e-05),
    "lbm": ("g", 453.59237, 0.002204622621848776),
    "u": ("g", 1.660538782e-24, 6.022141794216764e+23),
    "ozm": ("g", 28.349523125, 0.035273961949580414),
    "grain": ("g", 0.06479891, 15.43235835294143),
    "cwt": ("g", 45359.237, 2.2046226218487758e-05),
    "shweight": ("g", 45359.237, 2.2046226218487758e-05),
    "uk_cwt": ("g", 50802.34544, 1.9684130552221215e-05),
    "lcwt": ("g", 50802.34544, 1.9684130552221215e-05),
    "stone": ("g", 6350.29318, 0.00015747304441776972),
    "ton": ("g", 907184.74, 1.102311310924388e-06),
    "uk_ton": ("g", 1016046.9088, 9.842065276110607e-07),
    "LTON": ("g", 1016046.9088, 9.842065276110607e-07),
    "brton": ("g", 1016046.9088, 9.842065276110607e-07),
    # m
    "m": ("m", 1.0, 1.0),
    "mi": ("m", 1609.344, 0.0006213711922373339),
    "Nmi": ("m", 1852.0, 0.0005399568034557236),
    "in": ("m", 0.0254, 39.37007874015748),
    "ft": ("m", 0.3048, 3.2808398950131235),
    "yd": ("m", 0.9144, 1.0936132983377078),
    "ang": ("m", 1e-10, 10000000000.0),
    "ell": ("m", 1.143, 0.8748906386701663),
    "ly": ("m", 9460730472580800.0, 1.0570008340246154e-16),
    "parsec": ("m", 3.0856775812815532e+16, 3.240779289664725e-17),
    "pc": ("m", 3.0856775812815532e+16, 3.240779289664725e-17),
    "Picapt": ("m", 0.0003527777777777778, 2834.6456692913384),
    "Pica": ("m", 0.0003527777777777778, 2834.6456692913384),
    "pica": ("m", 0.004233333333333334, 236.22047244094486),
    "survey_mi": ("m", 1609.3472186944373, 0.0006213699494949495),
    # sec
    "sec": ("sec", 1.0, 1.0),
    "s": ("sec", 1.0, 1.0),
    "yr": ("sec", 31557600.0, 3.168808781402895e-08),
    "day": ("sec", 86400.0, 1.1574074074074073e-05),
    "d": ("sec", 86400.0, 1.1574074074074073e-05),
    "hr": ("sec", 3600.0, 0.0002777777777777778),
    "mn": ("sec", 60.0, 0.016666666666666666),
    "min": ("sec", 60.0, 0.016666666666666666),
    # Pa
    "Pa": ("Pa", 1.0, 1.0),
    "p": ("Pa", 1.0, 1.0),
    "atm": ("Pa", 101325.0, 9.869232667160129e-06),
    "at": ("Pa", 101325.0, 9.869232667160129e-06),
    "mmHg": ("Pa", 133.322, 0.007500637554192106),
    "psi": ("Pa", 6894.757293168362, 0.0001450377377302092),
    "Torr": ("Pa", 133.32236842105263, 0.007500616827041697),
    # N
    "N": ("N", 1.0, 1.0),
    "dyn": ("N", 1e-05, 100000.0),
    "dy": ("N", 1e-05, 100000.0),
    "lbf": ("N", 4.4482216152605005, 0.22480894309971047),
    "pond": ("N", 0.00980665, 101.97162129779282),
    # J
    "J": ("J", 1.0, 1.0),
    "e": ("J", 1e-07, 10000000.0),
    "c": ("J", 4.184, 0.2390057361376673),
    "cal": ("J", 4.1868, 0.23884589662749595),
    "eV": ("J", 1.602176487e-19, 6.241509647120418e+18),
    "ev": ("J", 1.602176487e-19, 6.241509647120418e+18),
    "HPh": ("J", 2684519.537696173, 3.725061359986188e-07),
    "hh": ("J", 2684519.537696173, 3.725061359986188e-07),
    "Wh": ("J", 3600.0, 0.0002777777777777778),
    "wh": ("J", 3600.0, 0.0002777777777777778),
    "flb": ("J", 1.3558179483314003, 0.7375621492772654),
    "BTU": ("J", 1055.05585262, 0.0009478171203133171),
    "btu": ("J", 1055.05585262, 0.0009478171203133171),
    # W
    "W": ("W", 1.0, 1.0),
    "w": ("W", 1.0, 1.0),
    "HP": ("W", 745.6998715822702, 0.0013410220895950279),
    "h": ("W", 745.6998715822702, 0.0013410220895950279),
    "PS": ("W", 735.49875, 0.0013596216173039043),
    # T
    "T": ("T", 1.0, 1.0),
    "ga": ("T", 0.0001, 10000.0),
    # m3
    "m3": ("m3", 1.0, 1.0),
    "m^3": ("m3", 1.0, 1.0),
    "tsp": ("m3", 4.92892159375e-06, 202884.13621105798),
    "tspm": ("m3", 5e-06, 200000.0),
    "tbs": ("m3", 1.4786764781249999e-05, 67628.045403686),
    "oz": ("m3", 2.9573529562499998e-05, 33814.022701843),
    "cup": ("m3", 0.00023658823649999998, 4226.752837730375),
    "pt": ("m3", 0.00047317647299999996, 2113.3764188651876),
    "us_pt": ("m3", 0.00047317647299999996, 2113.3764188651876),
    "uk_pt": ("m3", 0.00056826125, 1759.7539863927022),
    "qt": ("m3", 0.0009463529459999999, 1056.6882094325938),
    "uk_qt": ("m3", 0.0011365225, 879.8769931963511),
    "gal": ("m3", 0.0037854117839999997, 264.17205235814845),
    "uk_gal": ("m3", 0.00454609, 219.96924829908778),
    "l": ("m3", 0.001, 1000.0),
    "L": ("m3", 0.001, 1000.0),
    "lt": ("m3", 0.001, 1000.0),
    "ang3": ("m3", 1e-30, 1e+30),
    "ang^3": ("m3", 1e-30, 1e+30),
    "barrel": ("m3", 0.158987294928, 6.289810770432105),
    "bushel": ("m3", 0.03523907016688, 28.37759325840175),
    "ft3": ("m3", 0.028316846592, 35.31466672148859),
    "ft^3": ("m3", 0.028316846592, 35.31466672148859),
    "in3": ("m3", 1.6387064e-05, 61023.74409473228),
    "in^3": ("m3", 1.6387064e-05, 61023.74409473228),
    "ly3": ("m3", 8.467866646237152e+47, 1.1809349884417084e-48),
    "ly^3": ("m3", 8.467866646237152e+47, 1.1809349884417084e-48),
    "mi3": ("m3", 4168181825.4405794, 2.3991275857892774e-10),
    "mi^3": ("m3", 4168181825.4405794, 2.3991275857892774e-10),
    "yd3": ("m3", 0.764554857984, 1.3079506193143922),
    "yd^3": ("m3", 0.764554857984, 1.3079506193143922),
    "Nmi3": ("m3", 6352182208.0, 1.5742621468581148e-10),
    "Nmi^3": ("m3", 6352182208.0, 1.5742621468581148e-10),
    "Pica3": ("m3", 4.3903956618655696e-11, 22776990435.870632),
    "Pica^3": ("m3", 4.3903956618655696e-11, 22776990435.870632),
    "Picapt3": ("m3", 4.3903956618655696e-11, 22776990435.870632),
    "Picapt^3": ("m3", 4.3903956618655696e-11, 22776990435.870632),
    "GRT": ("m3", 2.8316846592, 0.35314666721488586),
    "regton": ("m3", 2.8316846592, 0.35314666721488586),
    "MTON": ("m3", 1.13267386368, 0.8828666680372147),
    # m2
    "m2": ("m2", 1.0, 1.0),
    "m^2": ("m2", 1.0, 1.0),
    "uk_acre": ("m2", 4046.8564224, 0.0002471053814671653),
    "us_acre": ("m2", 4046.8726098742522, 0.0002471043930466279),
    "ang2": ("m2", 1e-20, 1e+20),
    "ang^2": ("m2", 1e-20, 1e+20),
    "ar": ("m2", 100.0, 0.01),
    "ft2": ("m2", 0.09290304, 10.763910416709722),
    "ft^2": ("m2", 0.09290304, 10.763910416709722),
    "ha": ("m2", 10000.0, 0.0001),
    "in2": ("m2", 0.00064516, 1550.0031000062),
    "in^2": ("m2", 0.00064516, 1550.0031000062),
    "ly2": ("m2", 8.95054210748189e+31, 1.117250763128733e-32),
    "ly^2": ("m2", 8.95054210748189e+31, 1.117250763128733e-32),
    "mi2": ("m2", 2589988.110336, 3.861021585424458e-07),
    "mi^2": ("m2", 2589988.110336, 3.861021585424458e-07),
    "Nmi2": ("m2", 3429904.0, 2.9155334959812285e-07),
    "Nmi^2": ("m2", 3429904.0, 2.9155334959812285e-07),
    "Pica2": ("m2", 1.2445216049382715e-07, 8035216.070432141),
    "Pica^2": ("m2", 1.2445216049382715e-07, 8035216.070432141),
    "Picapt2": ("m2", 1.2445216049382715e-07, 8035216.070432141),
    "Picapt^2": ("m2", 1.2445216049382715e-07, 8035216.070432141),
    "yd2": ("m2", 0.83612736, 1.1959900463010802),
    "yd^2": ("m2", 0.83612736, 1.1959900463010802),
    "Morgen": ("m2", 2500.0, 0.0004),
    # bit
    "bit": ("bit", 1.0, 1.0),
    "byte": ("bit", 8.0, 0.125),
    # m/s
    "m/s": ("m/s", 1.0, 1.0),
    "m/sec": ("m/s", 1.0, 1.0),
    "m/h": ("m/s", 0.0002777777777777778, 3600.0),
    "m/hr": ("m/s", 0.0002777777777777778, 3600.0),
    "mph": ("m/s", 0.44704, 2.2369362920544025),
    "kn": ("m/s", 0.5144444444444445, 1.9438444924406046),
    "admkn": ("m/s", 0.5147733333333333, 1.9426025694156652),
}  # fmt: skip
"""Units by name, keyed to their category's base unit (temperature apart, in ``_SCALES``)."""

_PREFIXED = frozenset(
    {
        "J",
        "K",
        "L",
        "N",
        "Pa",
        "T",
        "W",
        "Wh",
        "ang",
        "ang2",
        "ang3",
        "ang^2",
        "ang^3",
        "ar",
        "at",
        "atm",
        "bit",
        "byte",
        "c",
        "cal",
        "dy",
        "dyn",
        "e",
        "eV",
        "ev",
        "g",
        "ga",
        "kel",
        "l",
        "lt",
        "ly",
        "m",
        "m/h",
        "m/hr",
        "m/s",
        "m/sec",
        "m2",
        "m3",
        "m^2",
        "m^3",
        "mmHg",
        "mph",
        "p",
        "parsec",
        "pc",
        "pond",
        "s",
        "sec",
        "u",
        "uk_pt",
        "w",
        "wh",
    }
)
"""The units that take a decimal prefix."""
_DECIMAL = {
    "Y": 24, "Z": 21, "E": 18, "P": 15, "T": 12, "G": 9, "M": 6, "k": 3, "h": 2, "da": 1, "e": 1,
    "d": -1, "c": -2, "m": -3, "u": -6, "n": -9, "p": -12, "f": -15, "a": -18, "z": -21, "y": -24,
}  # fmt: skip
"""Decimal prefixes and their powers of ten."""
_BINARY = {"ki": 10, "Mi": 20, "Gi": 30, "Ti": 40, "Pi": 50, "Ei": 60, "Zi": 70, "Yi": 80}
"""Binary prefixes and their powers of two, for ``bit`` and ``byte`` only."""

_SCALES = {"C": "C", "cel": "C", "F": "F", "fah": "F", "K": "K", "kel": "K", "Rank": "Rank",
           "Reau": "Reau"}  # fmt: skip
"""Temperature units and their scales."""
_TEMPERATURE: dict[tuple[str, str], Callable[[float], float]] = {
    ("C", "F"): lambda v: v * 1.8 + 32.0,
    ("C", "K"): lambda v: v + 273.15,
    ("C", "Rank"): lambda v: (v + 273.15) * 1.8,
    ("C", "Reau"): lambda v: v * 0.8,
    ("F", "C"): lambda v: (v - 32.0) / 1.8,
    ("F", "K"): lambda v: (v - 32.0) / 1.8 + 273.15,
    ("F", "Rank"): lambda v: v + 459.67,
    ("F", "Reau"): lambda v: (v - 32.0) / 1.8 * 0.8,
    ("K", "C"): lambda v: v - 273.15,
    ("K", "F"): lambda v: (v - 273.15) * 1.8 + 32.0,
    ("K", "Rank"): lambda v: v * 1.8,
    ("K", "Reau"): lambda v: (v - 273.15) * 0.8,
    ("Rank", "C"): lambda v: v / 1.8 - 273.15,
    ("Rank", "F"): lambda v: v - 459.67,
    ("Rank", "K"): lambda v: v / 1.8,
    ("Rank", "Reau"): lambda v: (v / 1.8 - 273.15) * 0.8,
    ("Reau", "C"): lambda v: v * 1.25,
    ("Reau", "F"): lambda v: v * 2.25 + 32.0,
    ("Reau", "K"): lambda v: v * 1.25 + 273.15,
    ("Reau", "Rank"): lambda v: (v * 1.25 + 273.15) * 1.8,
}
"""Excel's formula for each pair of temperature scales (capture ``convert_family``)."""
_PREFIX_IGNORED = frozenset({("Reau", "K"), ("K", "Rank"), ("K", "Reau")})
"""Pairs of scales where Excel ignores a prefix on kelvin: ``CONVERT(-40,"mK","Rank")`` is -72."""
_POWER = re.compile(r"(?:ang|m)\^?([23])")


def _unit(name: str) -> tuple[str, float] | None:
    """A unit name as Excel reads it: the unit and the multiplier its prefix gives, or None."""
    if name in _UNITS or name in _SCALES:
        return name, 1.0
    if name[:2] in _BINARY and name[2:] in ("bit", "byte"):
        return name[2:], float(2 ** _BINARY[name[:2]])
    for size in (2, 1):
        prefix, rest = name[:size], name[size:]
        if prefix in _DECIMAL and rest in _PREFIXED:
            power = _POWER.fullmatch(rest)
            exponent = _DECIMAL[prefix] * (int(power[1]) if power else 1)
            return rest, float(f"1e{exponent}")
    return None


def _temperature(v: float, a: str, ma: float, b: str, mb: float) -> float:
    sa, sb = _SCALES[a], _SCALES[b]
    if (sa, sb) in _PREFIX_IGNORED:
        return _TEMPERATURE[(sa, sb)](v)
    x = v * ma if ma != 1.0 else v
    y = x if sa == sb else _TEMPERATURE[(sa, sb)](x)
    return y / mb if mb != 1.0 else y


def _convert(v: float, a: str, ma: float, b: str, mb: float) -> float | ErrorValue:
    if a in _SCALES or b in _SCALES:
        if a not in _SCALES or b not in _SCALES:
            return NA
        return _temperature(v, a, ma, b, mb)
    (ca, ta, _), (cb, tb, fb) = _UNITS[a], _UNITS[b]
    if ca != cb:
        return NA
    x = v * ma if ma != 1.0 else v
    if a == b:
        y = x
    elif a == ca:
        y = x * fb
    elif b == cb:
        y = x * ta
    else:
        y = x * ta / tb
    return y / mb if mb != 1.0 else y


@function("CONVERT", "value", "value", "value", category="engineering", ulps=2,
          ulps_reason=(
              "Excel likely defines units through other units, rounding at each step; between two "
              "units neither of which is the base, its result can differ in the last bit or two"
          ))  # fmt: skip
def CONVERT(ctx: Context, number: Any, from_unit: Any, to_unit: Any) -> float | ErrorValue:
    """A number from one unit of measure to another."""
    v = _strict(number, ctx)
    if isinstance(v, ErrorValue):
        return v
    if not isinstance(from_unit, str) or not isinstance(to_unit, str):
        return NA
    a, b = _unit(from_unit), _unit(to_unit)
    if a is None or b is None:
        return NA
    try:
        result = _convert(v, a[0], a[1], b[0], b[1])
    except (OverflowError, ZeroDivisionError):
        return NUM
    if isinstance(result, ErrorValue):
        return result
    return result if math.isfinite(result) else NUM
