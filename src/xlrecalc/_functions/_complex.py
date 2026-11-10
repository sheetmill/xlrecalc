"""Complex numbers as Excel writes them, text such as ``3+4i`` (capture ``complex_family``).

A complex number is text: a real part, an imaginary part with its suffix ``i`` or ``j``, or both;
``i`` alone for 1. Each part is a number as Excel writes it as text (15 significant digits, the
exponent form past them), and a part of 0 is left out. The functions compute in polar form where
Excel does (``IMSQRT``, ``IMPOWER``), so their last bits are Excel's: ``IMSQRT(-1)`` is
``6.12323399573677E-17+i``. Text that is no complex number is ``#NUM!``; two suffixes in one
calculation are ``#NUM!`` in ``IMSUB`` and ``IMDIV`` and ``#VALUE!`` in ``IMSUM`` and
``IMPRODUCT``. A part is written in digits below 1E21 and with 15 significant digits in the
exponent form, which takes two digits at least (``E-08``).

Not imported by the package yet: 935 of the capture's 976 rows match, and the rest (Excel's
argument limits in the trigonometric and hyperbolic functions, ``IMEXP`` at an overflow, a few
last digits) come first.
"""

from __future__ import annotations

import math
import re
from typing import Any

from xlrecalc._coerce import Context, to_number
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
    if im_text == "0":
        return re_text
    im_part = {1.0: "", -1.0: "-"}.get(imaginary, im_text) + suffix  # 1 itself, not its text
    if re_text == "0":
        return im_part
    return re_text + ("" if im_part.startswith("-") else "+") + im_part


def _number(x: float) -> str:
    """A part as Excel writes it here: as ``&`` does, but in digits below 1E21, not 1E+20.

    1E20 is written in full and 1E300 with its exponent (capture complex_family); the switch at
    1E21 is a hypothesis between them.
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
_part("IMABS", lambda z: math.hypot(z[0], z[1]))


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
    """R^n (cos n theta + i sin n theta); ``scaled`` takes r by hypot, else as sqrt(a² + b²)."""
    r = math.hypot(a, b) if scaled else math.sqrt(a * a + b * b)
    theta = math.atan2(b, a)
    rn = r**n
    return rn * math.cos(n * theta), rn * math.sin(n * theta)


def _exp(a: float, b: float) -> tuple[float, float]:
    e = math.exp(a)
    return e * math.cos(b), e * math.sin(b)


def _ln(a: float, b: float) -> tuple[float, float] | ErrorValue:
    if a == 0 and b == 0:
        return NUM
    return math.log(math.hypot(a, b)), math.atan2(b, a)


def _sin(a: float, b: float) -> tuple[float, float]:
    return math.sin(a) * math.cosh(b), math.cos(a) * math.sinh(b)


def _cos(a: float, b: float) -> tuple[float, float]:
    return math.cos(a) * math.cosh(b), -math.sin(a) * math.sinh(b)


def _sinh(a: float, b: float) -> tuple[float, float]:
    return math.sinh(a) * math.cos(b), math.cosh(a) * math.sin(b)


def _cosh(a: float, b: float) -> tuple[float, float]:
    return math.cosh(a) * math.cos(b), math.sinh(a) * math.sin(b)


def _ratio(top: Any, bottom: Any) -> Any:
    def core(a: float, b: float) -> tuple[float, float] | ErrorValue:
        (p, q), (r, s) = top(a, b), bottom(a, b)
        return _div(p, q, r, s)

    return core


def _one(a: float, b: float) -> tuple[float, float]:
    return 1.0, 0.0


def _log_base(base: float) -> Any:
    def core(a: float, b: float) -> tuple[float, float] | ErrorValue:
        z = _ln(a, b)
        return z if isinstance(z, ErrorValue) else (z[0] / math.log(base), z[1] / math.log(base))

    return core


_UNARY: dict[str, Any] = {
    "IMCONJUGATE": lambda a, b: (a, -b),
    "IMSQRT": lambda a, b: _polar_power(a, b, 0.5, scaled=False),  # IMSQRT(1E300) is #NUM!
    "IMEXP": _exp,
    "IMLN": _ln,
    "IMLOG10": _log_base(10.0),
    "IMLOG2": _log_base(2.0),
    "IMSIN": _sin,
    "IMCOS": _cos,
    "_xlfn.IMTAN": _ratio(_sin, _cos),
    "_xlfn.IMCOT": _ratio(_cos, _sin),
    "_xlfn.IMSEC": _ratio(_one, _cos),
    "_xlfn.IMCSC": _ratio(_one, _sin),
    "_xlfn.IMSINH": _sinh,
    "_xlfn.IMCOSH": _cosh,
    "_xlfn.IMSECH": _ratio(_one, _cosh),
    "_xlfn.IMCSCH": _ratio(_one, _sinh),
}
for _name, _core in _UNARY.items():
    function(_name, "value", category=E)(_unary(_core))


# Several complex numbers --------------------------------------------------------------------


def _many(args: tuple[Any, ...], ctx: Context) -> list[Complex] | ErrorValue:
    """Every complex number the arguments hold: a range's non-blank cells, an array's elements."""
    out: list[Complex] = []
    for arg in args:
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
    except (OverflowError, ValueError):  # past the double's range
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
