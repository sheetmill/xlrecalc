"""Text functions: CONCATENATE, LEFT, RIGHT, MID, LEN, FIND, TRIM, UPPER, LOWER, VALUE, TEXT."""

from __future__ import annotations

import math
import re
from decimal import ROUND_HALF_UP, Decimal, localcontext
from typing import Any, cast

import xlformat

from xlrecalc._arrays import to_array
from xlrecalc._coerce import Context, scalar, to_bool, to_number, to_text
from xlrecalc._compare import wildcard
from xlrecalc._formula import MISSING
from xlrecalc._functions._math import _decimal
from xlrecalc._functions._registry import function
from xlrecalc._numtext import text_to_number
from xlrecalc._values import BLANK, ErrorValue, Ref

VALUE = ErrorValue.VALUE
T = "text"
_SPACES = re.compile(" +")


@function("CONCATENATE", "text", min_args=1, variadic=True, category=T)
def CONCATENATE(ctx: Context, *parts: str) -> str | ErrorValue:
    """The texts joined."""
    text = "".join(parts)
    return text if len(text) <= 32767 else VALUE


def _count(n: float, default: int) -> int | ErrorValue:
    k = math.trunc(n)
    return VALUE if k < 0 else k


@function("EXACT", "text", "text", category=T)
def EXACT(ctx: Context, a: str, b: str) -> bool:
    """Whether two texts are the same, case included."""
    return a == b


@function("LEFT", "text", "number", min_args=1, category=T)
def LEFT(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """The first ``n`` characters."""
    k = _count(n, 1)
    return k if isinstance(k, ErrorValue) else text[:k]


@function("RIGHT", "text", "number", min_args=1, category=T)
def RIGHT(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """The last ``n`` characters."""
    k = _count(n, 1)
    if isinstance(k, ErrorValue):
        return k
    return text[max(len(text) - k, 0) :] if k else ""


@function("MID", "text", "number", "number", category=T)
def MID(ctx: Context, text: str, start: float, n: float) -> str | ErrorValue:
    """``n`` characters from position ``start`` (1-based)."""
    s, k = math.trunc(start), math.trunc(n)
    if s < 1 or k < 0:
        return VALUE
    return _characters(_units(text)[s - 1 : s - 1 + k])


@function("LEN", "text", category=T)
def LEN(ctx: Context, text: str) -> float:
    """Number of characters."""
    return float(len(_units(text)))


@function("FIND", "text", "text", "number", min_args=2, category=T)
def FIND(ctx: Context, needle: str, haystack: str, start: float = 1.0) -> float | ErrorValue:
    """Position of ``needle`` in ``haystack`` from ``start``, with case; ``#VALUE!`` if absent."""
    s = math.trunc(start)
    units = _units(haystack)
    if s < 1 or s > len(units) + 1:
        return VALUE
    i = units.find(_units(needle), s - 1)
    return VALUE if i < 0 else float(i + 1)


@function("TRIM", "text", category=T)
def TRIM(ctx: Context, text: str) -> str:
    """Spaces removed at both ends and runs of spaces inside made one."""
    return _SPACES.sub(" ", text.strip(" "))


@function("UPPER", "text", category=T)
def UPPER(ctx: Context, text: str) -> str:
    """Upper case."""
    return text.upper()


@function("LOWER", "text", category=T)
def LOWER(ctx: Context, text: str) -> str:
    """Lower case."""
    return text.lower()


@function("VALUE", "value", category=T)
def VALUE_(ctx: Context, v: Any) -> float | ErrorValue:
    """Text as a number; a logical is ``#VALUE!``, unlike in arithmetic (§4.3)."""
    if isinstance(v, bool):
        return VALUE
    return to_number(v, ctx)


@function("TEXT", "value", "value", category=T)
def TEXT(ctx: Context, v: Any, fmt: Any) -> str | ErrorValue:
    """A value formatted with a number format, by ``xlformat.text_function`` (§8.3).

    Text that reads as a number or a date is formatted as that number, and a logical format is
    ``#VALUE!`` (capture ``textfmt``); xlformat does the rest as ``TEXT`` does it: the empty
    format, a format Excel refuses, a value the format cannot show.
    """
    if v is BLANK or v is MISSING:
        v = 0.0
    elif isinstance(v, str):
        n = text_to_number(v, ctx.quirks, date1904=ctx.date1904, today=ctx.today)
        if n is not None:
            v = n
    if isinstance(fmt, bool):
        return VALUE
    if fmt is BLANK or fmt is MISSING:
        fmt = ""
    fmt = to_text(fmt, ctx)
    if isinstance(fmt, ErrorValue):
        return fmt
    text = xlformat.text_function(v, fmt, date1904=ctx.date1904, quirks=ctx.quirks.to_xlformat())
    if isinstance(text, xlformat.ExcelError):
        return ErrorValue(str(text))
    return text


_MAX_TEXT = 32767  # Excel's longest text in a cell


@function("REPT", "text", "number", category=T)
def REPT(ctx: Context, text: str, times: float) -> str | ErrorValue:
    """A text repeated; ``#VALUE!`` past 32,767 characters."""
    n = math.trunc(times)
    if n < 0 or len(text) * n > _MAX_TEXT:
        return VALUE
    return text * n


@function("REPLACE", "text", "number", "number", "text", category=T)
def REPLACE(ctx: Context, old: str, start: float, count: float, new: str) -> str | ErrorValue:
    """A text with ``count`` characters from ``start`` replaced by ``new``."""
    s, n = math.trunc(start), math.trunc(count)
    if s < 1 or n < 0 or s > _MAX_TEXT:
        return VALUE
    units = _units(old)
    result = units[: s - 1] + _units(new) + units[s - 1 + n :]
    return VALUE if len(result) > _MAX_TEXT else _characters(result)


# Excel for Mac numbers characters in the Mac OS Roman set (capture misc); Excel for Windows
# uses its ANSI code page instead.
@function("CHAR", "number", category=T)
def CHAR(ctx: Context, n: float) -> str | ErrorValue:
    """The character with a code from 1 to 255 (Mac OS Roman on the Mac)."""
    k = math.trunc(n)
    if not 1 <= k <= 255:
        return VALUE
    return bytes([k]).decode("mac_roman")


@function("CODE", "text", category=T)
def CODE(ctx: Context, text: str) -> float | ErrorValue:
    """The code of a text's first character (Mac OS Roman; 95 outside it)."""
    if not text:
        return VALUE
    try:
        return float(text[0].encode("mac_roman")[0])
    except UnicodeEncodeError:
        return 95.0


def _fixed_text(x: float, decimals: float, grouping: bool) -> str | ErrorValue:
    """``x`` rounded as ROUND rounds it, written with ``decimals`` places (none below 0).

    The rounding works on the 15-significant-digit decimal a number shows, so digits past the
    fifteenth are zeros (``FIXED(1.5E20)`` is ``150,000,000,000,000,000,000.00``).
    """
    d = math.trunc(decimals)
    if d > 127:
        return VALUE
    with localcontext() as context:
        context.prec = 500
        try:
            rounded = _decimal(x).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
        except ArithmeticError:
            return VALUE
        if rounded == 0:
            rounded = abs(rounded)
        return format(rounded, f"{',' if grouping else ''}.{max(d, 0)}f")


_MAX_DOLLAR = 255  # capture text_family: a result longer than a number format may be is #VALUE!


@function("DOLLAR", "number", "number", min_args=1, category=T)
def DOLLAR(ctx: Context, x: float, decimals: float = 2.0) -> str | ErrorValue:
    """A number as currency text: ``$1,234.57``, negatives in parentheses.

    With decimals of 0 or more a negative that rounds to zero keeps its parentheses
    (``DOLLAR(-0.004)`` is ``($0.00)``); with fewer it is rounded first (captured).
    """
    text = _fixed_text(x, decimals, True)
    if isinstance(text, ErrorValue):
        return text
    if len(text) > _MAX_DOLLAR:
        return VALUE
    if text.startswith("-"):
        text = text[1:]
    elif not (x < 0 and math.trunc(decimals) >= 0):
        return f"${text}"
    return f"(${text})"


@function("FIXED", "number", "raw", "logical", min_args=1, category=T)
def FIXED(ctx: Context, x: float, decimals: Any = 2.0, no_commas: bool = False) -> Any:
    """A number as text with a fixed number of decimals, grouped unless ``no_commas``.

    An empty decimals argument means 2 here (``FIXED(x,,TRUE)``), where DOLLAR takes 0.
    """
    if decimals is MISSING:
        decimals = 2.0
    d = to_number(scalar(decimals, ctx), ctx)
    if isinstance(d, ErrorValue):
        return d
    return _fixed_text(x, d, not no_commas)


def _flat(arg: Any) -> Any:
    """Values of an argument in order, as a range reads: non-blank cells, or an array's elements."""
    if isinstance(arg, Ref):
        return arg.values()
    if isinstance(arg, tuple):
        return [v for row in arg for v in row]
    return [arg]


@function("_xlfn.CONCAT", "raw", min_args=1, variadic=True, errors=True, category=T)
def CONCAT(ctx: Context, *args: Any) -> str | ErrorValue:
    """The texts of its arguments, ranges and arrays taken cell by cell."""
    parts = []
    size = 0
    for arg in args:
        for v in _flat(arg):
            if v is MISSING or v is BLANK:
                continue
            text = to_text(v, ctx)
            if isinstance(text, ErrorValue):
                return text
            parts.append(text)
            size += len(text)
    return VALUE if size > _MAX_TEXT else "".join(parts)


def _joined(arg: Any, ignore: bool, delimiter_empty: bool) -> Any:
    """An argument's values for ``TEXTJOIN``: blanks count only if empty texts are kept."""
    if isinstance(arg, Ref) and not ignore and not delimiter_empty:
        cells = arg.rows * arg.columns
        if cells > _MAX_TEXT + 1:  # the delimiters alone pass the limit: no need to walk them
            return None
        return [arg.value(r, c) for r in range(arg.rows) for c in range(arg.columns)]
    return _flat(arg)


@function("_xlfn.TEXTJOIN", "raw", "raw", "raw", min_args=3, variadic=True, errors=True, category=T)
def TEXTJOIN(ctx: Context, delimiter: Any, ignore: Any, *args: Any) -> str | ErrorValue:
    """Texts joined by delimiters taken in turn, leaving out empty texts if ``ignore``."""
    keep = to_bool(ignore[0][0] if isinstance(ignore, tuple) else ignore, ctx)
    if isinstance(keep, ErrorValue):
        return keep
    delimiters = []
    for v in _flat(delimiter) if delimiter is not MISSING else ():
        text = to_text(v, ctx)
        if isinstance(text, ErrorValue):
            return text
        delimiters.append(text)
    empty_delimiter = not any(delimiters)
    parts: list[str] = []
    for arg in args:
        values = _joined(arg, keep, empty_delimiter)
        if values is None:
            return VALUE
        for v in values:
            if v is MISSING or v is BLANK:
                text = ""
            else:
                text = to_text(v, ctx)
                if isinstance(text, ErrorValue):
                    return text
            if text or not keep:
                parts.append(text)
    if not parts:
        return ""
    out = [parts[0]]
    for i, part in enumerate(parts[1:]):
        if delimiters:
            out.append(delimiters[i % len(delimiters)])
        out.append(part)
    joined = "".join(out)
    return VALUE if len(joined) > _MAX_TEXT else joined


# Units ---------------------------------------------------------------------------------------
# Excel's text is UTF-16: LEN, MID, FIND and their kin count a character outside the first plane as
# two, and can cut it in two (capture text_family).


def _units(text: str) -> str:
    """The text with each character outside the first plane as its two surrogates."""
    if text.isascii() or max(text) <= "￿":
        return text
    raw = text.encode("utf-16-le")
    return "".join(chr(raw[i] | raw[i + 1] << 8) for i in range(0, len(raw), 2))


def _characters(units: str) -> str:
    """Surrogates back into characters where they pair; a lone one stays."""
    return units.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")


def _limited(text: str) -> str | ErrorValue:
    """A result past the limit of a text is ``#VALUE!``."""
    return VALUE if len(_units(text)) > _MAX_TEXT else text


@function("PROPER", "text", category=T)
def PROPER(ctx: Context, text: str) -> str:
    """A capital for each letter after anything not a letter; the other letters small."""
    out, inside = [], False
    for ch in text:
        changed = ch.lower() if inside else ch.upper()
        out.append(changed if len(changed) == 1 else ch)
        inside = ch.isalpha()
    return "".join(out)


_CONTROLS = dict.fromkeys(range(32))


@function("CLEAN", "text", category=T)
def CLEAN(ctx: Context, text: str) -> str:
    """The text without the characters below 32."""
    return text.translate(_CONTROLS)


@function("T", "raw", errors=True, category=T)
def T_(ctx: Context, v: Any) -> Any:
    """Text as it is, an error as it is, and the empty text for anything else."""
    if isinstance(v, Ref):
        v = v.value(0, 0)
    elif isinstance(v, tuple):
        v = v[0][0]
    return v if isinstance(v, str | ErrorValue) else ""


@function("_xlfn.UNICHAR", "number", category=T)
def UNICHAR(ctx: Context, n: float) -> str | ErrorValue:
    """The character of a code point; ``#N/A`` for a surrogate or a noncharacter."""
    code = math.trunc(n)
    if not 0 < code <= 0x10FFFF:
        return VALUE
    if 0xD800 <= code <= 0xDFFF or 0xFDD0 <= code <= 0xFDEF or code & 0xFFFE == 0xFFFE:
        return ErrorValue.NA
    return chr(code)


@function("_xlfn.UNICODE", "text", category=T)
def UNICODE(ctx: Context, text: str) -> float | ErrorValue:
    """The code point of the first character."""
    return float(ord(text[0])) if text else VALUE


def _search(needle: str, within: str, start: float) -> float | ErrorValue:
    units, s = _units(within), math.trunc(start)
    if s < 1 or s > len(units):
        return VALUE
    tildes = len(needle) - len(needle.rstrip("~"))
    pattern = needle[:-1] if tildes % 2 else needle  # a last ~ with nothing to escape is dropped
    found = wildcard(_units(pattern)).search(units, s - 1)
    return VALUE if found is None else float(found.start() + 1)


@function("SEARCH", "text", "text", "number", min_args=2, category=T)
def SEARCH(ctx: Context, needle: str, within: str, start: float = 1.0) -> float | ErrorValue:
    """Position of ``needle`` from ``start``, without case and with ``?``, ``*`` and ``~``."""
    return _search(needle, within, start)


@function("SEARCHB", "text", "text", "number", min_args=2, category=T)
def SEARCHB(ctx: Context, needle: str, within: str, start: float = 1.0) -> float | ErrorValue:
    """``SEARCH``, in this language setting: no text counts double."""
    return _search(needle, within, start)


@function("FINDB", "text", "text", "number", min_args=2, category=T)
def FINDB(ctx: Context, needle: str, within: str, start: float = 1.0) -> float | ErrorValue:
    """``FIND``, in this language setting: no text counts double."""
    return FIND(ctx, needle, within, start)  # type: ignore[no-any-return]


@function("LENB", "text", category=T)
def LENB(ctx: Context, text: str) -> float:
    """``LEN``, in this language setting: no text counts double."""
    return LEN(ctx, text)  # type: ignore[no-any-return]


@function("LEFTB", "text", "number", min_args=1, category=T)
def LEFTB(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """``LEFT``, in this language setting: no text counts double."""
    return LEFT(ctx, text, n)  # type: ignore[no-any-return]


@function("RIGHTB", "text", "number", min_args=1, category=T)
def RIGHTB(ctx: Context, text: str, n: float = 1.0) -> str | ErrorValue:
    """``RIGHT``, in this language setting: no text counts double."""
    return RIGHT(ctx, text, n)  # type: ignore[no-any-return]


@function("MIDB", "text", "number", "number", category=T)
def MIDB(ctx: Context, text: str, start: float, n: float) -> str | ErrorValue:
    """``MID``, in this language setting: no text counts double."""
    return MID(ctx, text, start, n)  # type: ignore[no-any-return]


@function("REPLACEB", "text", "number", "number", "text", category=T)
def REPLACEB(ctx: Context, text: str, start: float, n: float, new: str) -> str | ErrorValue:
    """``REPLACE``, in this language setting: no text counts double."""
    return REPLACE(ctx, text, start, n, new)  # type: ignore[no-any-return]


@function("SUBSTITUTE", "text", "text", "text", "value", min_args=3, category=T)
def SUBSTITUTE(ctx: Context, text: str, old: str, new: str, instance: Any = None) -> Any:
    """``new`` for each ``old``, with case, or only for the ``instance``-th."""
    k = 0
    if instance is not None:
        n = VALUE if isinstance(instance, bool) else to_number(instance, ctx)
        if isinstance(n, ErrorValue):
            return n
        k = math.trunc(n)
        if k < 1:
            return VALUE
    if not old:
        return text
    if instance is None:  # the size first: a small formula can ask for a billion characters
        size = len(_units(text)) + text.count(old) * (len(_units(new)) - len(_units(old)))
        return VALUE if size > _MAX_TEXT else text.replace(old, new)
    at = -1
    for _ in range(k):
        at = text.find(old, at + 1)
        if at < 0:
            return text
    return _limited(text[:at] + new + text[at + len(old) :])


_ANY_SPACE = re.compile(r"\s+")


@function("_xlfn.NUMBERVALUE", "text", "text", "text", min_args=1, category=T)
def NUMBERVALUE(
    ctx: Context, text: str, decimal: str = ".", group: str | None = None
) -> float | ErrorValue:
    """A number written with any separators, whatever the language setting."""
    if group is None:  # the default gives way to a decimal separator that takes it
        group = "." if decimal[:1] == "," else ","
    if not decimal or not group or decimal[0] == group[0]:
        return VALUE
    body = _ANY_SPACE.sub("", text)
    stripped = body.rstrip("%")
    if not stripped:
        return 0.0
    percents = len(body) - len(stripped)
    whole, point, fraction = stripped.partition(decimal[0])
    if group[0] in fraction:
        return VALUE
    number = text_to_number(
        whole.replace(group[0], "") + (("." + fraction) if point else ""),
        ctx.quirks,
        date1904=ctx.date1904,
        today=ctx.today,
    )
    if not isinstance(number, float):
        return VALUE
    return number / 100.0**percents if percents else number


@function("USDOLLAR", "number", "number", min_args=1, category=T)
def USDOLLAR(ctx: Context, x: float, decimals: float = 2.0) -> str | ErrorValue:
    """``DOLLAR``, in this language setting."""
    return DOLLAR(ctx, x, decimals)  # type: ignore[no-any-return]


def _shown(v: Any, strict: bool, ctx: Context) -> str | ErrorValue:
    """One value as ``VALUETOTEXT`` writes it: in strict form text is in quotes."""
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"' if strict else v
    if isinstance(v, ErrorValue):
        return v.code
    return to_text(v, ctx)


def _format_code(fmt: Any, ctx: Context, top: int) -> int | ErrorValue:
    if fmt is MISSING:
        return 0
    n = to_number(fmt, ctx)
    if isinstance(n, ErrorValue):
        return n
    code = math.trunc(n)
    return code if 0 <= code <= top else VALUE


@function("_xlfn.VALUETOTEXT", "value", "value", min_args=1, errors=True, category=T)
def VALUETOTEXT(ctx: Context, v: Any, fmt: Any = MISSING) -> str | ErrorValue:
    """A value as text, concise (0) or strict (1)."""
    code = _format_code(fmt, ctx, 1)
    if isinstance(code, ErrorValue):
        return code
    shown = _shown(v, code == 1, ctx)
    return shown if isinstance(shown, ErrorValue) else _limited(shown)


@function("_xlfn.ARRAYTOTEXT", "array", "value", min_args=1, errors=True, category=T)
def ARRAYTOTEXT(ctx: Context, v: Any, fmt: Any = MISSING) -> str | ErrorValue:
    """An array as text: comma-separated (0), or as an array constant (1)."""
    if isinstance(v, Ref) and v.is_cell():
        v = v.value(0, 0)
    if v is BLANK or v is MISSING:
        return VALUE
    if isinstance(v, ErrorValue):
        return v
    code = _format_code(fmt, ctx, 1)
    if isinstance(code, ErrorValue):
        return code
    rows = to_array(v) if isinstance(v, Ref) else v if isinstance(v, tuple) else ((v,),)
    if code == 0:
        parts = [_shown(x, False, ctx) for row in rows for x in row]
        return _limited(", ".join(cast(list[str], parts)))
    lines = [",".join(cast(list[str], [_shown(x, True, ctx) for x in row])) for row in rows]
    return _limited("{" + ";".join(lines) + "}")


def _delimiter_at(
    text: str, delimiters: list[str], instance: int, fold: bool, match_end: bool
) -> tuple[int, int] | None:
    """Where the ``instance``-th delimiter sits (start, end), from the end when negative."""
    haystack = text.casefold() if fold else text
    needles = [d.casefold() if fold else d for d in delimiters]
    spots = sorted(
        (at, at + len(d))
        for d in needles
        for at in (m.start() for m in re.finditer(f"(?={re.escape(d)})", haystack))
    )
    if match_end:
        spots = [(0, 0), *spots] if instance < 0 else [*spots, (len(text), len(text))]
    if instance > 0:
        return spots[instance - 1] if instance <= len(spots) else None
    return spots[instance] if -instance <= len(spots) else None


def _before_after(ctx: Context, args: tuple[Any, ...], after: bool) -> str | ErrorValue:
    text, delimiter, instance, mode, end, missing = args
    if isinstance(text, ErrorValue):
        return text
    if isinstance(delimiter, ErrorValue):
        return delimiter
    rows = delimiter if isinstance(delimiter, tuple) else ((delimiter,),)
    delimiters = []
    for row in rows:
        for d in row:
            if isinstance(d, ErrorValue):
                return d
            shown = to_text(d, ctx)
            if isinstance(shown, ErrorValue):
                return shown
            delimiters.append(shown)
    whole = to_text(text, ctx)
    if isinstance(whole, ErrorValue):
        return whole
    k = 1.0 if instance is MISSING else to_number(instance, ctx)
    if isinstance(k, ErrorValue):
        return k
    m = 0.0 if mode is MISSING else to_number(mode, ctx)
    if isinstance(m, ErrorValue):
        return m
    e = 0.0 if end is MISSING else to_number(end, ctx)
    if isinstance(e, ErrorValue):
        return e
    n, m_code, e_code = math.floor(k), math.trunc(m), math.trunc(e)
    if n == 0 or m_code not in (0, 1) or e_code not in (0, 1):
        return VALUE
    if "" in delimiters:
        return (whole if n < 0 else "") if not after else ("" if n < 0 else whole)
    if instance is not MISSING and abs(n) > len(whole) - min(map(len, delimiters)) + 1:
        return VALUE
    spot = _delimiter_at(whole, delimiters, n, bool(m_code), bool(e_code))
    if spot is None:
        return ErrorValue.NA if missing is MISSING else to_text(missing, ctx)
    return whole[spot[1] :] if after else whole[: spot[0]]


@function(
    "_xlfn.TEXTBEFORE", "value", "array", "value", "value", "value", "value",
    min_args=2, errors=True, category=T,
)  # fmt: skip
def TEXTBEFORE(
    ctx: Context, text: Any, delimiter: Any, instance: Any = MISSING, mode: Any = MISSING,
    end: Any = MISSING, missing: Any = MISSING,
) -> Any:  # fmt: skip
    """The text before a delimiter."""
    return _before_after(ctx, (text, delimiter, instance, mode, end, missing), False)


@function(
    "_xlfn.TEXTAFTER", "value", "array", "value", "value", "value", "value",
    min_args=2, errors=True, category=T,
)  # fmt: skip
def TEXTAFTER(
    ctx: Context, text: Any, delimiter: Any, instance: Any = MISSING, mode: Any = MISSING,
    end: Any = MISSING, missing: Any = MISSING,
) -> Any:  # fmt: skip
    """The text after a delimiter."""
    return _before_after(ctx, (text, delimiter, instance, mode, end, missing), True)
