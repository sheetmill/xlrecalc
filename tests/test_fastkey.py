"""The raw-text template key (ARCHITECTURE.md §5.3): equal raw keys mean equal templates."""

from __future__ import annotations

import random
import re

import pytest

from tests.formula_corpus import CORPUS
from xlrecalc._formula import FormulaSyntaxError, Kind, TokenizeError, parse, template_key, tokenize
from xlrecalc._formula._fastkey import raw_key
from xlrecalc._formula._tokens import column_number

POSITIONS = [(0, 0), (4, 2), (99, 26), (5000, 700)]
_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")


def letters(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def moved(text: str, dr: int, dc: int) -> str | None:
    """The formula copied ``dr`` rows and ``dc`` columns away, as Excel's fill does."""
    out = []
    for tok in tokenize(text):
        piece = tok.text
        if tok.kind is Kind.CELL:
            m = _CELL.fullmatch(piece)
            assert m
            c = column_number(m[2]) + (0 if m[1] else dc)
            r = int(m[4]) + (0 if m[3] else dr)
            if not (1 <= c <= 16384 and 1 <= r <= 1048576):
                return None
            piece = f"{m[1]}{letters(c)}{m[3]}{r}"
        elif tok.kind in (Kind.COLUMNS, Kind.ROWS):
            return None  # whole columns and rows are left to the parsed-key test
        out.append(tok.space + piece)
    return "".join(out)


def parsed_key(text: str, row: int, col: int) -> str | None:
    try:
        return template_key(parse(text), row, col)
    except (FormulaSyntaxError, TokenizeError):
        return None


@pytest.mark.parametrize("text", CORPUS, ids=repr)
def test_moving_a_formula_keeps_its_raw_key(text):
    if parsed_key(text, 0, 0) is None:
        pytest.skip("not a formula the parser accepts")
    for r, c in POSITIONS[1:]:
        copy = moved(text, r, c)
        if copy is None:
            continue
        assert raw_key(copy, r, c) == raw_key(text, 0, 0)
        assert parsed_key(copy, r, c) == parsed_key(text, 0, 0)


def test_equal_raw_keys_mean_equal_templates():
    seen: dict[str, str] = {}
    for text in CORPUS:
        for r, c in POSITIONS:
            for copy, row, col in ((text, r, c), (moved(text, r, c), r, c)):
                if copy is None:
                    continue
                parsed = parsed_key(copy, row, col)
                if parsed is None:
                    continue
                raw = raw_key(copy, row, col)
                assert seen.setdefault(raw, parsed) == parsed, (copy, row, col)


def test_references_the_tokenizer_reads_as_names_stay_text():
    assert "XFE1" in raw_key("XFE1+1", 0, 0)
    assert "A1048577" in raw_key("A1048577+1", 0, 0)
    assert '"A1"' in raw_key('"A1"&B1', 0, 0)
    assert "LOG10(" in raw_key("LOG10(A1)", 0, 0)
    assert "'A1'!" in raw_key("'A1'!B2", 0, 0)
    assert raw_key("[1]!'A1'", 0, 0) != raw_key("[1]!'A2'", 1, 0)  # a quoted item is a name
    assert "[A1]" in raw_key("Table1[A1]", 0, 0)
    assert raw_key("A1", 0, 0) != raw_key("A1", 1, 0)
    assert raw_key("A1", 0, 0) == raw_key("A2", 1, 0)
    assert raw_key("$A$1", 0, 0) == raw_key("$A$1", 7, 3)


def _fills(seed: int) -> list[tuple[str, int, int]]:
    """Corpus formulas filled to random cells (most in their own column), and the edge cases."""
    from xlrecalc._formula._fastkey import translate

    rng = random.Random(seed)
    out: list[tuple[str, int, int]] = []
    texts = [t.removeprefix("=") for t in CORPUS] + list(EDGES)
    for text in texts:
        r0, c0 = rng.randint(0, 2000), rng.randint(0, 60)
        for _ in range(12):
            r = rng.choice([rng.randint(0, 1048575), r0 + rng.randint(-3, 3)])
            c = c0 if rng.random() < 0.7 else rng.randint(0, 16383)
            r = min(max(r, 0), 1048575)
            out.append((translate(text, r0, c0, r, c) if rng.random() < 0.8 else text, r, c))
    return out


EDGES = (
    "A1048576+1",  # moved down a row, past the sheet: the text stays text
    "A1048577+1",
    "A1#+1",  # a "#" between runs of digits
    "A#1+1",
    "A$5+B5",  # an absolute row stays
    "3:5+$3:$5",
    "LOG10(A5)",  # digits in a name
    'A5&"A5"',
    "'Sheet 2'!A5",
    "A05+1",
)


@pytest.mark.parametrize("seed", range(3))
def test_the_shape_cache_gives_the_scan_s_key(seed: int) -> None:
    """A formula's raw key is the scan's, whether its shape was met before or not."""
    from xlrecalc._formula import _fastkey

    _fastkey._SHAPES.clear()
    for text, row, col in _fills(seed):
        assert raw_key(text, row, col) == _fastkey._raw_key(text, row, col), (text, row, col)


def test_a_relative_row_moved_off_the_sheet_is_not_its_shape_s_key() -> None:
    from xlrecalc._formula import _fastkey

    _fastkey._SHAPES.clear()
    assert raw_key("A1048575+1", 1048573, 1) == _fastkey._raw_key("A1048575+1", 1048573, 1)
    assert raw_key("A1048576+1", 1048574, 1) == "\x1eR[1]C[-1]\x1e+1"
    assert raw_key("A1048577+1", 1048575, 1) == "A1048577+1"  # past the sheet: a name
    assert raw_key("A1#+1", 0, 0) != raw_key("A#1+1", 0, 0)


def test_the_shape_cache_is_bounded() -> None:
    from xlrecalc._formula import _fastkey

    _fastkey._SHAPES.clear()
    for i in range(_fastkey._SHAPE_LIMIT + 10):
        raw_key("A1+" + "x" * (i % 50) + str(i) + "y" * (i // 50), 0, 0)
        assert len(_fastkey._SHAPES) <= _fastkey._SHAPE_LIMIT


@pytest.mark.parametrize(
    ("first", "second"),
    [("A5+5", "A6+6"), ("A$5+B5", "A$6+B6"), ("$3:$5+A3", "$4:$6+A4")],
)
def test_a_number_that_moves_with_the_row_but_is_not_a_relative_row(
    first: str, second: str
) -> None:
    """A constant or an absolute row that happens to change by the rows moved is another key."""
    from xlrecalc._formula import _fastkey

    _fastkey._SHAPES.clear()
    assert raw_key(first, 4, 1) == _fastkey._raw_key(first, 4, 1)
    assert raw_key(second, 5, 1) == _fastkey._raw_key(second, 5, 1)
    assert raw_key(first, 4, 1) != raw_key(second, 5, 1)
