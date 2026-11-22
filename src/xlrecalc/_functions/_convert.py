"""CONVERT (captures ``convert_pairs``, ``convert_family``): a number from one unit to another.

Each unit is in one category (mass, distance, time, pressure, force, energy, power, magnetism,
volume, area, information, speed; temperature apart), and each category has an internal unit,
often not the visible base: the ångström for distance, its square for area, the litre for
volume, the dyne for force, the erg for energy, the gauss for magnetism, the base unit
elsewhere. Each unit holds one double, how many internal units it is. A conversion
always goes through the internal unit, the same unit and synonyms included, ``(number * c_from) /
c_to``, so ``CONVERT(7.3,"BTU","BTU")`` is 7.300000000000001; the prefixes come last, as one
multiplication by the net power of ten written as a decimal (``1e6`` for ``km`` to ``mm``), or
by the exact power of two for a binary prefix. This gives every captured row bit for bit.

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

# name: (category, how many of the category's internal unit one unit is): Excel's constants
# (captures convert_pairs, convert_family and convert_odd), written as the decimal or the
# expression in doubles whose value they are; psi and Pica^2 as their doubles
_UNITS: dict[str, tuple[str, float]] = {
    # g: in g
    "u": ("g", 1.660538782e-24),
    "grain": ("g", 0.06479891),
    "g": ("g", 1.0),
    "ozm": ("g", 28.349523125),
    "lbm": ("g", 453.59237),
    "stone": ("g", 6350.29318),
    "sg": ("g", 453.59237 * 9.80665 / 0.3048),
    "cwt": ("g", 45359.237),
    "shweight": ("g", 45359.237),
    "lcwt": ("g", 50802.34544),
    "uk_cwt": ("g", 50802.34544),
    "ton": ("g", 907184.74),
    "LTON": ("g", 1016046.9088),
    "brton": ("g", 1016046.9088),
    "uk_ton": ("g", 1016046.9088),
    # m: in 1E-10 m (ångström)
    "ang": ("m", 1.0),
    "Pica": ("m", 254000000.0 / 72),
    "Picapt": ("m", 254000000.0 / 72),
    "pica": ("m", 254000000.0 / 6),
    "in": ("m", 254000000.0),
    "ft": ("m", 3048000000.0),
    "yd": ("m", 9144000000.0),
    "m": ("m", 10000000000.0),
    "ell": ("m", 11430000000.0),
    "mi": ("m", 16093440000000.0),
    "survey_mi": ("m", 6336 / 3937 * 1e13),
    "Nmi": ("m", 18520000000000.0),
    "ly": ("m", 9.4607304725808e+25),
    "parsec": ("m", 3.0856775812815533e+26),
    "pc": ("m", 3.0856775812815533e+26),
    # sec: in s
    "s": ("sec", 1.0),
    "sec": ("sec", 1.0),
    "min": ("sec", 60.0),
    "mn": ("sec", 60.0),
    "hr": ("sec", 3600.0),
    "d": ("sec", 86400.0),
    "day": ("sec", 86400.0),
    "yr": ("sec", 31557600.0),
    # Pa: in Pa
    "Pa": ("Pa", 1.0),
    "p": ("Pa", 1.0),
    "mmHg": ("Pa", 133.322),
    "Torr": ("Pa", 101325 / 760),
    "psi": ("Pa", 6894.757293168362),
    "at": ("Pa", 101325.0),
    "atm": ("Pa", 101325.0),
    # N: in 1E-5 N (dyne)
    "dy": ("N", 1.0),
    "dyn": ("N", 1.0),
    "pond": ("N", 980.665),
    "N": ("N", 100000.0),
    "lbf": ("N", 444822.16152605),
    # J: in 1E-7 J (erg)
    "eV": ("J", 1.602176487e-12),
    "ev": ("J", 1.602176487e-12),
    "e": ("J", 1.0),
    "J": ("J", 10000000.0),
    "flb": ("J", 444822.16152605 * 30.48),
    "c": ("J", 41840000.0),
    "cal": ("J", 41868000.0),
    "BTU": ("J", 10550558526.2),
    "btu": ("J", 10550558526.2),
    "Wh": ("J", 36000000000.0),
    "wh": ("J", 36000000000.0),
    "HPh": ("J", 745.6998715822702 * 3600 * 1e7),
    "hh": ("J", 745.6998715822702 * 3600 * 1e7),
    # W: in W
    "W": ("W", 1.0),
    "w": ("W", 1.0),
    "PS": ("W", 735.49875),
    "HP": ("W", 550 * 4.4482216152605 * 0.3048),
    "h": ("W", 550 * 4.4482216152605 * 0.3048),
    # T: in ga (gauss)
    "T": ("T", 10000.0),
    "ga": ("T", 1.0),
    # bit: in bit
    "bit": ("bit", 1.0),
    "byte": ("bit", 8.0),
    # m/s: in m/s
    "m/h": ("m/s", 1 / 3600),
    "m/hr": ("m/s", 1 / 3600),
    "mph": ("m/s", 0.44704),
    "kn": ("m/s", 1852 / 3600),
    "admkn": ("m/s", 1853.184 / 3600),
    "m/s": ("m/s", 1.0),
    "m/sec": ("m/s", 1.0),
    # m2: in 1E-20 m2 (ångström²)
    "ang2": ("m2", 1.0),
    "ang^2": ("m2", 1.0),
    "Pica2": ("m2", 12445216049382.715),
    "Pica^2": ("m2", 12445216049382.715),
    "Picapt2": ("m2", 12445216049382.715),
    "Picapt^2": ("m2", 12445216049382.715),
    "in2": ("m2", 6.4516e+16),
    "in^2": ("m2", 6.4516e+16),
    "ft2": ("m2", 9.290304e+18),
    "ft^2": ("m2", 9.290304e+18),
    "yd2": ("m2", 8.3612736e+19),
    "yd^2": ("m2", 8.3612736e+19),
    "m2": ("m2", 1e+20),
    "m^2": ("m2", 1e+20),
    "ar": ("m2", 1e+22),
    "Morgen": ("m2", 2.5e+23),
    "uk_acre": ("m2", 4.0468564224e+23),
    "us_acre": ("m2", (6336 / 3937 * 1e3) ** 2 / 640 * 1e20),
    "ha": ("m2", 1e+24),
    "mi2": ("m2", 2.589988110336e+26),
    "mi^2": ("m2", 2.589988110336e+26),
    "Nmi2": ("m2", 3.429904e+26),
    "Nmi^2": ("m2", 3.429904e+26),
    "ly2": ("m2", 8.95054210748189e+51),
    "ly^2": ("m2", 8.95054210748189e+51),
    # m3: in 1E-3 m3 (litre)
    "ang3": ("m3", 1e-27),
    "ang^3": ("m3", 1e-27),
    "Pica3": ("m3", 4.39039566186557e-08),
    "Pica^3": ("m3", 4.39039566186557e-08),
    "Picapt3": ("m3", 4.39039566186557e-08),
    "Picapt^3": ("m3", 4.39039566186557e-08),
    "tsp": ("m3", 0.00492892159375),
    "tspm": ("m3", 0.005),
    "tbs": ("m3", 0.01478676478125),
    "in3": ("m3", 0.016387064),
    "in^3": ("m3", 0.016387064),
    "oz": ("m3", 0.0295735295625),
    "cup": ("m3", 0.2365882365),
    "pt": ("m3", 0.473176473),
    "us_pt": ("m3", 0.473176473),
    "uk_pt": ("m3", 0.56826125),
    "qt": ("m3", 0.946352946),
    "L": ("m3", 1.0),
    "l": ("m3", 1.0),
    "lt": ("m3", 1.0),
    "uk_qt": ("m3", 1.1365225),
    "gal": ("m3", 3.785411784),
    "uk_gal": ("m3", 4.54609),
    "ft3": ("m3", 28.316846592),
    "ft^3": ("m3", 28.316846592),
    "bushel": ("m3", 35.23907016688),
    "barrel": ("m3", 158.987294928),
    "yd3": ("m3", 764.554857984),
    "yd^3": ("m3", 764.554857984),
    "m3": ("m3", 1000.0),
    "m^3": ("m3", 1000.0),
    "MTON": ("m3", 1132.67386368),
    "GRT": ("m3", 2831.6846592),
    "regton": ("m3", 2831.6846592),
    "mi3": ("m3", 4168181825.4405794 * 1000),
    "mi^3": ("m3", 4168181825.4405794 * 1000),
    "Nmi3": ("m3", 6352182208000.0),
    "Nmi^3": ("m3", 6352182208000.0),
    "ly3": ("m3", 9460730472580800.0 ** 3 * 1000),
    "ly^3": ("m3", 9460730472580800.0 ** 3 * 1000),
}  # fmt: skip
"""Units by name, with their category and constant (temperature apart, in ``_SCALES``)."""

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
    ("F", "Reau"): lambda v: (v - 32.0) * 0.8 / 1.8,
    ("K", "C"): lambda v: v - 273.15,
    ("K", "F"): lambda v: (v - 273.15) * 1.8 + 32.0,
    ("K", "Rank"): lambda v: v * 1.8,
    ("K", "Reau"): lambda v: (v - 273.15) * 0.8,
    ("Rank", "C"): lambda v: v / 1.8 - 273.15,
    ("Rank", "F"): lambda v: v - 459.67,
    ("Rank", "K"): lambda v: v / 1.8,
    ("Rank", "Reau"): lambda v: (v / 1.8 - 273.15) * 0.8,
    ("Reau", "C"): lambda v: v / 0.8,
    ("Reau", "F"): lambda v: v / 0.8 * 1.8 + 32.0,
    ("Reau", "K"): lambda v: v / 0.8 + 273.15,
    ("Reau", "Rank"): lambda v: (v / 0.8 + 273.15) * 1.8,
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
    (ka, ca), (kb, cb) = _UNITS[a], _UNITS[b]
    if ka != kb:
        return NA
    x = v * ca / cb  # through the internal unit, also for the same unit
    if ma == mb:
        return x
    ea, eb = round(math.log10(ma)), round(math.log10(mb))
    if float(f"1e{ea}") == ma and float(f"1e{eb}") == mb:
        return x * float(f"1e{ea - eb}")  # the prefixes last, as one decimal power of ten
    return x * (ma / mb)  # a binary prefix: an exact power of two


@function("CONVERT", "value", "value", "value", category="engineering")
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
