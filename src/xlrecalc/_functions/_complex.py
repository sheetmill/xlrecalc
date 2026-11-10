"""Complex numbers as Excel writes them, text such as ``3+4i`` (capture ``complex_family``).

A complex number is text: a real part, an imaginary part with its suffix ``i`` or ``j``, or both;
``i`` alone for 1. Each part is written with 15 significant digits, in digits below 1E21 and
while they take at most 19 places after the point, else in the exponent form with two exponent
digits at least (``E-08``); a part of 0 is left out. The functions compute in polar form where
Excel does (``IMSQRT``, ``IMPOWER``), so their last bits are Excel's: ``IMSQRT(-1)`` is
``6.12323399573677E-17+i``. Text that is no complex number is ``#NUM!``; two suffixes in one
calculation are ``#NUM!`` in ``IMSUB`` and ``IMDIV`` and ``#VALUE!`` in ``IMSUM`` and
``IMPRODUCT``.

The trigonometric functions are #NUM! from a real part of 2^27 and the hyperbolic ones and
``IMEXP`` from an imaginary part of 2^27, as ``SIN`` is, but ``IMCOT`` of a real number from
2^27 is 0; an overflowing ``cosh`` or ``sinh`` is infinite, so ``IMSECH(1E300)`` is 0 and
``IMTAN`` of a huge imaginary part is ``i``; ``IMEXP`` takes e^a as the largest double at most,
and a result whose 15 digits pass it is #NUM!. ``IMTAN`` and ``IMCOT`` use the double-angle
form, ``IMABS`` scales by the larger part, ``IMPOWER`` takes a whole power by squaring and
``IMLOG2`` divides by the double just below ln 2, each fitted to Excel's last digits.

Every row of the capture matches. An empty required argument is #N/A (the registry), a later empty
argument of ``IMSUM`` and ``IMPRODUCT`` is skipped, and a one-argument function called with its
argument empty is refused in a file (``_load.py``), each as Excel does (capture ``empty_args``).
"""

from __future__ import annotations

import math
import re
import sys
from typing import Any

from xlrecalc._coerce import Context, to_number
from xlrecalc._formula import MISSING
from xlrecalc._functions._registry import function
from xlrecalc._numtext import number_to_text
from xlrecalc._values import BLANK, ErrorValue, Ref

NUM, VALUE, DIV0 = ErrorValue.NUM, ErrorValue.VALUE, ErrorValue.DIV0
E = "engineering"
_NUMBER = r"(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?"
_NUMBER_TEXT = re.compile(rf"[+-]?{_NUMBER}")

Complex = tuple[float, float, str | None]  # real, imaginary, suffix (None: no imaginary part)


def _parse(text: str) -> Complex | ErrorValue:
    """A complex number from its text, or #NUM!: spaces at either end are dropped.

    The imaginary part starts at the last sign that follows neither an exponent's ``e`` nor
    another sign, so ``1+-2i`` is 1 and -2; a part of signs alone is 1 with their sign.
    """
    text = text.strip(" ")
    if text == "":
        return NUM
    if text[-1] not in "ij":
        if not _NUMBER_TEXT.fullmatch(text) or not math.isfinite(float(text)):
            return NUM
        return float(text), 0.0, None
    body, suffix = text[:-1], text[-1]
    cut = 0
    for k in range(1, len(body)):
        if body[k] in "+-" and body[k - 1] not in "eE+-":
            cut = k
    real_text, imag_text = body[:cut], body[cut:]
    if real_text and not _NUMBER_TEXT.fullmatch(real_text):
        return NUM
    signs = len(imag_text) - len(imag_text.lstrip("+-"))
    digits = imag_text[signs:]
    if (
        signs > 2
        or (signs == 2 and not real_text)
        or (digits and not re.fullmatch(_NUMBER, digits))
    ):
        return NUM
    imaginary = float(digits) if digits else 1.0
    if not math.isfinite(imaginary) or (real_text and not math.isfinite(float(real_text))):
        return NUM
    if imag_text[:signs].count("-") % 2:
        imaginary = -imaginary
    return float(real_text) if real_text else 0.0, imaginary, suffix


def _arg(v: Any, ctx: Context) -> Complex | ErrorValue:
    """An argument as a complex number: text parsed, a number real, a logical #VALUE!."""
    if isinstance(v, ErrorValue):
        return v
    if isinstance(v, bool):
        return VALUE
    if v is BLANK or v is None:
        return 0.0, 0.0, None
    if isinstance(v, str):
        return _parse(v)
    x = to_number(v, ctx)
    return x if isinstance(x, ErrorValue) else (x, 0.0, None)


def _suffix(*numbers: Complex) -> str | ErrorValue:
    """The suffix of a result: the one the arguments use, ``i`` by default; two are #VALUE!."""
    used = {s for _, _, s in numbers if s is not None}
    if len(used) > 1:  # #NUM! for IMSUB and IMDIV, #VALUE! for IMSUM and IMPRODUCT (capture)
        return NUM
    return used.pop() if used else "i"


def _text(real: float, imaginary: float, suffix: str) -> str | ErrorValue:
    """A complex number as Excel writes it."""
    if not (math.isfinite(real) and math.isfinite(imaginary)):
        return NUM
    re_text, im_text = _number(real), _number(imaginary)
    if math.isinf(float(re_text)) or math.isinf(float(im_text)):  # 15 digits past the largest
        return NUM
    if im_text == "0":
        return re_text
    im_part = {1.0: "", -1.0: "-"}.get(imaginary, im_text) + suffix  # 1 itself, not its text
    if re_text == "0":
        return im_part
    return re_text + ("" if im_part.startswith("-") else "+") + im_part


def _number(x: float) -> str:
    """A part as Excel writes it here, with 15 significant digits.

    In digits below 1E21 and while they take at most 19 places after the point, else in the
    exponent form (``&`` switches sooner both ways): 999999999999999000000 and 1E+21,
    0.0000123456789012345 and 1.23456789012345E-06, 0.0000000000000000001 and 1E-20 (capture
    complex_family).
    """
    text = number_to_text(x)
    if "E" in text:  # 15 significant digits in the exponent form, where ``&`` keeps fewer
        mantissa, exponent = f"{x:.14E}".split("E")
        mantissa = mantissa.rstrip("0").rstrip(".")
        text = f"{mantissa}E{exponent[0]}{int(exponent[1:]):02d}"  # E-08: two digits at least
    if "E+" in text and abs(x) < 1e21:
        mantissa, exponent = text.split("E+")
        sign = "-" if mantissa.startswith("-") else ""
        digits = mantissa.lstrip("-").replace(".", "")
        whole = len(mantissa.lstrip("-").split(".")[0])
        return sign + digits + "0" * (int(exponent) + whole - len(digits))
    if "E-" in text:
        mantissa, exponent = text.split("E-")
        sign = "-" if mantissa.startswith("-") else ""
        digits = mantissa.lstrip("-").replace(".", "")
        if int(exponent) - 1 + len(digits) <= 19:
            return f"{sign}0.{'0' * (int(exponent) - 1)}{digits}"
    return text


def _unary(core: Any) -> Any:
    """A function of one complex number giving one, the suffix kept."""

    def wrapped(ctx: Any, v: Any) -> Any:
        z = _arg(v, ctx)
        if isinstance(z, ErrorValue):
            return z
        try:
            out = core(z[0], z[1])
        except (OverflowError, ValueError, ZeroDivisionError):  # past the double's range
            return NUM
        if isinstance(out, ErrorValue):
            return out
        return _text(out[0], out[1], z[2] or "i")

    return wrapped


def _div(a: float, b: float, c: float, d: float) -> tuple[float, float] | ErrorValue:
    """(a + bi) / (c + di), by Smith's scaling, so that c * c never overflows (1 / 1E300)."""
    if c == 0 and d == 0:
        return NUM
    if abs(c) >= abs(d):
        r = d / c
        den = c + d * r
        return (a + b * r) / den, (b - a * r) / den
    r = c / d
    den = c * r + d
    return (a * r + b) / den, (b * r - a) / den


# Construction and parts ---------------------------------------------------------------------


@function("COMPLEX", "value", "value", "value", min_args=2, category=E)
def COMPLEX(ctx: Any, real: Any, imaginary: Any, suffix: Any = "i") -> Any:
    """A complex number from its parts, with ``i`` or ``j``; a logical part is #VALUE!."""
    parts = []
    for v in (real, imaginary):
        x = VALUE if isinstance(v, bool) else to_number(v, ctx)
        if isinstance(x, ErrorValue):
            return x
        parts.append(x)
    real, imaginary = parts
    if suffix is BLANK or suffix is None or suffix == "":
        suffix = "i"
    if suffix not in ("i", "j"):
        return VALUE
    return _text(real, imaginary, suffix)


def _part(name: str, pick: Any) -> None:
    @function(name, "value", category=E)
    def core(ctx: Any, v: Any) -> Any:
        z = _arg(v, ctx)
        return z if isinstance(z, ErrorValue) else pick(z)


_part("IMREAL", lambda z: z[0])
_part("IMAGINARY", lambda z: z[1])
_part("IMABS", lambda z: _modulus(z[0], z[1]))


def _modulus(a: float, b: float) -> float:
    """|a + bi| as Excel computes it, scaled by the larger part (IMABS(-2.5-1.5j) to the bit)."""
    big, small = (abs(a), abs(b)) if abs(a) >= abs(b) else (abs(b), abs(a))
    return 0.0 if big == 0 else big * math.sqrt(1 + (small / big) ** 2)


@function("IMARGUMENT", "value", category=E)
def IMARGUMENT(ctx: Any, v: Any) -> Any:
    """The angle of a complex number, in radians; #DIV/0! at 0."""
    z = _arg(v, ctx)
    if isinstance(z, ErrorValue):
        return z
    if z[0] == 0 and z[1] == 0:
        return DIV0
    return math.atan2(z[1], z[0])


# One complex number to another -------------------------------------------------------------


def _polar_power(a: float, b: float, n: float, scaled: bool = True) -> tuple[float, float]:
    """R^n (cos n theta + i sin n theta); ``scaled`` takes r by hypot, else as sqrt(a² + b²).

    R^n is Excel's: by squaring for a whole power, a square root for 0.5, e^(n ln r) for any
    other, and a negative power the reciprocal of the positive one, so ``IMPOWER(2,-0.5)`` is
    1/sqrt(2) and ``(0.5+0.5i)^1000`` has Excel's last digits (capture complex_family, all 145
    IMPOWER rows that give a number).
    """
    r = math.hypot(a, b) if scaled else math.sqrt(a * a + b * b)  # not IMABS's form (capture)
    theta = math.atan2(b, a)
    m = abs(n)
    if m.is_integer() and m < 2.0**53:
        rm = _whole_power(r, int(m))
    else:
        rm = math.sqrt(r) if m == 0.5 else math.exp(m * math.log(r)) if r else 0.0
    rn = 1.0 / rm if n < 0 else rm  # ZeroDivisionError at an underflow, OverflowError past
    return rn * math.cos(n * theta), rn * math.sin(n * theta)


def _whole_power(r: float, k: int) -> float:
    """R^k by squaring: at most 53 squarings and as many products."""
    out = 1.0
    while k:
        if k & 1:
            out *= r
        k >>= 1
        if k:
            r *= r
    return out


_LARGEST_ANGLE = 2.0**27  # SIN's and COS's limit in Excel, #NUM! from it (captures math_trig)


def _big(x: float) -> bool:
    return abs(x) >= _LARGEST_ANGLE


def _sinh_(x: float) -> float:
    return math.copysign(math.inf, x) if abs(x) > 710 else math.sinh(x)


def _cosh_(x: float) -> float:
    return math.inf if abs(x) > 710 else math.cosh(x)


def _exp(a: float, b: float) -> tuple[float, float] | ErrorValue:
    """E^a (cos b + i sin b), e^a at most the largest double, as Excel clamps it."""
    if _big(b):
        return NUM
    e = sys.float_info.max if a > 709.78 else math.exp(a)
    return e * math.cos(b), e * math.sin(b)


def _ln(a: float, b: float) -> tuple[float, float] | ErrorValue:
    if a == 0 and b == 0:
        return NUM
    return math.log(math.hypot(a, b)), math.atan2(b, a)


def _sin(a: float, b: float) -> tuple[float, float]:
    return math.sin(a) * _cosh_(b), math.cos(a) * _sinh_(b)


def _cos(a: float, b: float) -> tuple[float, float]:
    return math.cos(a) * _cosh_(b), -math.sin(a) * _sinh_(b)


def _sinh(a: float, b: float) -> tuple[float, float]:
    return _sinh_(a) * math.cos(b), _cosh_(a) * math.sin(b)


def _cosh(a: float, b: float) -> tuple[float, float]:
    return _cosh_(a) * math.cos(b), _sinh_(a) * math.sin(b)


def _tan(a: float, b: float) -> tuple[float, float] | ErrorValue:
    """(sin 2a + i sinh 2b) / (cos 2a + cosh 2b), as Excel computes it: i where b is huge."""
    den = math.cos(2 * a) + _cosh_(2 * b)
    if not math.isfinite(den):
        return 0.0, math.copysign(1.0, b)
    return math.sin(2 * a) / den, _sinh_(2 * b) / den


def _cot(a: float, b: float) -> tuple[float, float] | ErrorValue:
    """(sin 2a - i sinh 2b) / (cosh 2b - cos 2a), as Excel computes it: -i where b is huge."""
    den = _cosh_(2 * b) - math.cos(2 * a)
    if not math.isfinite(den):
        return 0.0, -math.copysign(1.0, b)
    if den == 0:
        return NUM
    return math.sin(2 * a) / den, -_sinh_(2 * b) / den


def _limited(core: Any, on_real: bool) -> Any:
    """#NUM! where the part a sine or cosine takes is 2^27 or more, as SIN gives (capture)."""

    def limited(a: float, b: float) -> Any:
        return NUM if _big(a if on_real else b) else core(a, b)

    return limited


def _ratio(top: Any, bottom: Any) -> Any:
    """Top / bottom; 0 where the bottom overflows (IMSEC of a huge imaginary part, as Excel)."""

    def core(a: float, b: float) -> tuple[float, float] | ErrorValue:
        (p, q), (r, s) = top(a, b), bottom(a, b)
        if not (math.isfinite(r) and math.isfinite(s)):
            return 0.0, 0.0
        return _div(p, q, r, s)

    return core


def _one(a: float, b: float) -> tuple[float, float]:
    return 1.0, 0.0


def _log_base(ln_base: float) -> Any:
    def core(a: float, b: float) -> tuple[float, float] | ErrorValue:
        z = _ln(a, b)
        return z if isinstance(z, ErrorValue) else (z[0] / ln_base, z[1] / ln_base)

    return core


# IMLOG2 divides by the double just below ln 2: IMLOG2(1E20) is 66.4385618977473, one unit above
# the true value's last digit, and all 25 IMLOG2 rows that give a number agree (capture).
_LN2_EXCEL = math.nextafter(math.log(2.0), 0.0)


def _cot_limited(a: float, b: float) -> tuple[float, float] | ErrorValue:
    """IMCOT from a real part of 2^27: 0 for a real number, else #NUM! (capture)."""
    if _big(a):
        return (0.0, 0.0) if b == 0 else NUM
    return _cot(a, b)


_UNARY: dict[str, Any] = {
    "IMCONJUGATE": lambda a, b: (a, -b),
    "IMSQRT": lambda a, b: _polar_power(a, b, 0.5, scaled=False),  # IMSQRT(1E300) is #NUM!
    "IMEXP": _exp,
    "IMLN": _ln,
    "IMLOG10": _log_base(math.log(10.0)),
    "IMLOG2": _log_base(_LN2_EXCEL),
    "IMSIN": _limited(_sin, True),
    "IMCOS": _limited(_cos, True),
    "_xlfn.IMTAN": _limited(_tan, True),
    "_xlfn.IMCOT": _cot_limited,
    "_xlfn.IMSEC": _limited(_ratio(_one, _cos), True),
    "_xlfn.IMCSC": _limited(_ratio(_one, _sin), True),
    "_xlfn.IMSINH": _limited(_sinh, False),
    "_xlfn.IMCOSH": _limited(_cosh, False),
    "_xlfn.IMSECH": _limited(_ratio(_one, _cosh), False),
    "_xlfn.IMCSCH": _limited(_ratio(_one, _sinh), False),
}
for _name, _core in _UNARY.items():
    function(_name, "value", category=E)(_unary(_core))


# Several complex numbers --------------------------------------------------------------------


def _many(args: tuple[Any, ...], ctx: Context) -> list[Complex] | ErrorValue:
    """Every complex number the arguments hold: a range's non-blank cells, an array's elements.

    An empty argument after the first is skipped: ``IMPRODUCT("3+4j",)`` is ``3+4j`` (capture
    empty_args; an empty first one is #N/A, by the registry).
    """
    out: list[Complex] = []
    for arg in args:
        if arg is MISSING:
            continue
        if isinstance(arg, Ref):
            values: list[Any] = arg.values()
        elif isinstance(arg, tuple):
            values = [v for row in arg for v in row]
        else:
            values = [arg]
        for v in values:
            z = _arg(v, ctx)
            if isinstance(z, ErrorValue):
                return z
            out.append(z)
    return out


@function("IMSUM", "raw", variadic=True, category=E)
def IMSUM(ctx: Any, *args: Any) -> Any:
    """The sum of complex numbers."""
    zs = _many(args, ctx)
    if isinstance(zs, ErrorValue):
        return zs
    suffix = _suffix(*zs)
    if isinstance(suffix, ErrorValue):
        return VALUE
    return _text(sum(z[0] for z in zs), sum(z[1] for z in zs), suffix)


@function("IMPRODUCT", "raw", variadic=True, category=E)
def IMPRODUCT(ctx: Any, *args: Any) -> Any:
    """The product of complex numbers."""
    zs = _many(args, ctx)
    if isinstance(zs, ErrorValue):
        return zs
    suffix = _suffix(*zs)
    if isinstance(suffix, ErrorValue) or not zs:
        return VALUE
    a, b = 1.0, 0.0
    for c, d, _ in zs:
        a, b = a * c - b * d, a * d + b * c
    out = _text(a, b, suffix)
    return VALUE if out == NUM else out  # an overflow is #VALUE! here (capture complex_family)


@function("IMSUB", "value", "value", category=E)
def IMSUB(ctx: Any, v: Any, w: Any) -> Any:
    """The difference of two complex numbers."""
    return _binary(ctx, v, w, lambda a, b, c, d: (a - c, b - d))


@function("IMDIV", "value", "value", category=E)
def IMDIV(ctx: Any, v: Any, w: Any) -> Any:
    """The quotient of two complex numbers; #NUM! by 0."""
    return _binary(ctx, v, w, _div)


@function("IMPOWER", "value", "number", category=E)
def IMPOWER(ctx: Any, v: Any, n: float) -> Any:
    """A complex number to a real power, in polar form."""
    z = _arg(v, ctx)
    if isinstance(z, ErrorValue):
        return z
    if z[0] == 0 and z[1] == 0:
        return "0" if n > 0 else NUM
    try:
        real, imaginary = _polar_power(z[0], z[1], n)
    except (OverflowError, ValueError, ZeroDivisionError):  # past the double's range
        return NUM
    return _text(real, imaginary, z[2] or "i")


def _binary(ctx: Any, v: Any, w: Any, op: Any) -> Any:
    x, y = _arg(v, ctx), _arg(w, ctx)
    if isinstance(x, ErrorValue):
        return x
    if isinstance(y, ErrorValue):
        return y
    suffix = _suffix(x, y)
    if isinstance(suffix, ErrorValue):
        return suffix
    out = op(x[0], x[1], y[0], y[1])
    return out if isinstance(out, ErrorValue) else _text(out[0], out[1], suffix)
