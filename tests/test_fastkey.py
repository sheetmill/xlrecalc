"""The raw-text template key (ARCHITECTURE.md §5.3): equal raw keys mean equal templates."""

from __future__ import annotations

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
