"""Padded arrays give exactly what dense arrays give (ARCHITECTURE.md §7.7).

A range taller than its data, made into an array where only operators lie between it and a
function taught to read a padded array, is padded: its rows with data, then one fill. Random
sheets and random formulas are computed twice, padded and with every array dense, and must agree
bit for bit. The threshold is lowered so that 200-row ranges are padded and the dense side stays
cheap; the fills include an error (``1/x``), an empty text (``x&""``), a number (``--(x="")``)
and a logical (``x=""``).
"""

from __future__ import annotations

import random

import pytest

import xlrecalc._engine as engine_module
from tests.test_captures import same
from xlrecalc._arrays import Padded, pair
from xlrecalc._engine import Engine
from xlrecalc._quirks import Quirks
from xlrecalc._values import ErrorValue, Ref

ROWS = 200
"""The ranges' height; data stops by row 60."""
RANGES = ["A1:A200", "B1:B200", "C1:C200", "A1:B200", "B1:C200"]  # C is empty
WIDTH = {"A1:A200": 1, "B1:B200": 1, "C1:C200": 1, "A1:B200": 2, "B1:C200": 2}
SCALARS = ["0", "1", "2.5", '""', '"k3"', "TRUE", "-1"]
VALUES: list[object] = [1.0, 2.5, -3.0, 0.0, "k3", "", "x", True, False, ErrorValue.NA]


def _sheet(rng: random.Random, full: bool) -> list[tuple[int, int, object]]:
    """Data in rows 1 to 60 of A and B: half the cells, or (``full``) every one, so that a
    search reaches the fill instead of stopping at a blank among the data."""
    cells = ((r, c) for r in range(60) for c in (0, 1))
    return [(r, c, rng.choice(VALUES)) for r, c in cells if full or rng.random() < 0.5]


ELEMENTWISE = ["ABS", "LEN", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "ISLOGICAL"]


def _operand(rng: random.Random, width: int, depth: int) -> str:
    """An array expression of ``width`` columns: a range under operators, sometimes wrapped in a
    function that is not taught (it must receive a dense array)."""
    ranges = [r for r in RANGES if WIDTH[r] == width]
    if depth == 0:
        return rng.choice(ranges)
    inner = _operand(rng, width, depth - 1)
    form = rng.randrange(12)
    if form == 11:
        return f'(({inner})*1&"")'  # numeric text
    if form == 0:
        return f"-{inner}"
    if form == 1:
        return f'--({inner}="")'
    if form == 2:
        return f"1/{inner}"
    if form == 3:
        return f'({inner}&"")'
    if form == 4:
        return f"({inner})%"
    if form == 5:
        return f"({inner}{rng.choice('+-*/^')}{rng.choice(SCALARS)})"
    if form == 6:
        return f"({inner}{rng.choice(['=', '<>', '<', '>'])}{rng.choice(SCALARS)})"
    if form == 7:
        return f"({inner}*{_operand(rng, width, depth - 1)})"
    if form == 8:
        return f"({inner}={_operand(rng, width, depth - 1)})"
    if form == 9:
        return f"{rng.choice([*ELEMENTWISE, 'N'])}({inner})"  # N is not element-wise: dense
    choice = rng.randrange(4)
    if choice == 0:
        return f'IF({inner}="",1,{inner})'  # two padded arguments
    if choice == 1:
        return f"IF(1/{inner}>0,{inner})"  # an error test, no else: FALSE
    if choice == 2:
        return f'IF({inner}>1,"y",{rng.choice(SCALARS)})'
    return f'IF(ISBLANK({inner}),{inner}&"",-{inner})'


AGGREGATES = ["SUM", "COUNT", "COUNTA", "MIN", "MAX", "AND", "OR"]


def _formula(rng: random.Random) -> str:
    """``SUMPRODUCT`` or an aggregate over operands (each an array formula, see ``_values``)."""
    width = rng.choice((1, 1, 2))
    args = [_operand(rng, width, rng.randrange(4)) for _ in range(rng.choice((1, 1, 2)))]
    roll = rng.random()
    if roll < 0.3:
        return f"SUMPRODUCT({','.join(args)})"
    if roll < 0.55:
        return _lookup(rng, args[0], width)
    name = rng.choice(AGGREGATES)
    if name in ("MIN", "MAX") and rng.random() < 0.5:
        args[0] = f"-{args[0]}"
    return f"{name}({','.join(args)})"


KEYS = ["0", "1", '""', "FALSE", "TRUE", '"k3"', "2.5", "-1", '"x"']
"""Lookup keys; several equal a common fill (0, FALSE, an empty text, TRUE)."""
POSITIONS = ["1", "5", "60", "61", "62", "150", "200", "201", "0"]
"""Rows into a 200-row array: data, the first fill rows, the end, past it, 0 (a whole column)."""


def _lookup(rng: random.Random, operand: str, width: int) -> str:
    """``INDEX``, ``MATCH`` or ``LOOKUP`` over an operand (``MATCH`` over one column)."""
    form = rng.randrange(4)
    if rng.random() < 0.1:
        return f"{rng.choice(['ROWS', 'COLUMNS'])}({operand})"
    if form == 0 or width == 2:
        column = rng.choice(["", ",1", f",{width}"])
        return f"INDEX({operand},{rng.choice(POSITIONS)}{column})"
    if form == 1 and rng.random() < 0.5:
        return f"MATCH({rng.choice(KEYS)},{operand},{rng.choice(['0', '1', '-1'])})"
    if form == 1:
        match, search = rng.choice(["0", "1", "-1", "2"]), rng.choice(["1", "-1", "2", "-2"])
        return f"_xlfn.XMATCH({rng.choice(KEYS)},{operand},{match},{search})"
    if form == 2:
        return f"LOOKUP({rng.choice(KEYS)},{operand})"
    return f"LOOKUP({rng.choice(KEYS)},{operand},B1:B200)"


def _values(cells: list[tuple[int, int, object]], formulas: list[str]) -> list[object]:
    e = Engine(Quirks.EXCEL)
    s = e.add_sheet("d")
    for r, c, v in cells:
        e.sheets[s].set(r, c, v)  # type: ignore[arg-type]
    for i, f in enumerate(formulas):  # each a one-cell legacy array formula: an array context
        e.set_formula(s, i, 5, f, array=(False, (i, 5, i, 5)))
    e.recalculate(full=True)
    return [e.get(s, i, 5) for i in range(len(formulas))]


@pytest.mark.parametrize("seed", range(24))
def test_padded_and_dense_agree(seed: int, monkeypatch: pytest.MonkeyPatch) -> None:
    rng = random.Random(seed)
    cells = _sheet(rng, full=seed % 2 == 1)
    formulas = [_formula(rng) for _ in range(40)]
    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    padded = _values(cells, formulas)
    monkeypatch.setattr(Engine, "arr_padded", Engine.arr)  # every array dense, as before
    dense = _values(cells, formulas)
    wrong = [(f, p, d) for f, p, d in zip(formulas, padded, dense, strict=True) if not same(p, d)]
    assert wrong == []


def test_a_padded_array_is_made_and_read(monkeypatch: pytest.MonkeyPatch) -> None:
    """The comparison above is not vacuous: the padded path runs and ``SUMPRODUCT`` reads it."""
    made: list[Padded] = []
    original = Engine.arr_padded

    def spy(self: Engine, v: object) -> object:
        out = original(self, v)
        if type(out) is Padded and isinstance(v, Ref):  # made from a range, not passed on
            made.append(out)
        return out

    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    monkeypatch.setattr(Engine, "arr_padded", spy)
    cells = [(r, 0, float(r)) for r in range(10)]
    assert _values(cells, ['SUMPRODUCT((A1:A200="")*1)', "SUMPRODUCT(N(A1:A200)*1)"]) == [
        190.0,
        0.0,  # N reads one value of a range (A1), not each
    ]
    assert len(made) == 1  # the first only: N is not taught, so its range is dense


def test_the_fill_rows_add_one_by_one_unless_exact() -> None:
    """A fill of 0.1 over many rows sums as repeated addition does, not as a multiplication."""
    a = Padded(((1.0,),), 10_001, 1, 0.1)
    from xlrecalc._functions._math import _sumproduct_padded

    total = 1.0
    for _ in range(10_000):
        total += 0.1
    assert _sumproduct_padded([a]) == total
    assert _sumproduct_padded([Padded(((0.5,),), 11, 1, 2.0)]) == 20.5  # integers: no shortcut
    assert _sumproduct_padded([Padded(((1.0,),), 11, 1, 2.0)]) == 21.0  # exact: multiplied


def test_two_padded_arrays_of_different_sizes_are_not_paired() -> None:
    a, b = Padded(((1.0,),), 5, 1, 0.0), Padded(((1.0,),), 6, 1, 0.0)
    assert pair(lambda x, y: x + y, a, b) is None
    assert pair(lambda x, y: x + y, a, ((1.0,),)) is None


SEARCHES = [
    "MATCH(0,A1:A200*1,0)",  # the fill's first position
    'MATCH(TRUE,A1:A200="",0)',
    'MATCH("",A1:A200&"",0)',
    "MATCH(0,A1:A200*1,1)",  # a bisection into the fill's run
    "MATCH(5,A1:A200*1,1)",
    "MATCH(0,A1:A200*1,-1)",
    "MATCH(20,-A1:A200,-1)",
    "LOOKUP(0,A1:A200*1,B1:B200)",
    "LOOKUP(2,1/(A1:A200<>0),B1:B200)",
    'LOOKUP(2,1/(A1:A200<>""),A1:A200)',
    "LOOKUP(0,A1:A200*1)",
    "INDEX(A1:A200*1,150)",
    "INDEX(A1:A200*1,0)",
    "_xlfn.XMATCH(0,A1:A200*1,0,-1)",  # the fill's last position
    "_xlfn.XMATCH(0,A1:A200*1)",
    '_xlfn.XMATCH("",A1:A200&"",0,-1)',
    "_xlfn.XMATCH(5,A1:A200*1,1,-1)",
    "_xlfn.XMATCH(5,A1:A200*1,-1)",
    "_xlfn.XMATCH(0,A1:A200*1,0,2)",  # bisection: made dense
    '_xlfn.XMATCH("*",A1:A200&"",2,-1)',
    "SUMPRODUCT(--ISBLANK(A1:A200))",  # element-wise over a range: a fill of TRUE
    "SUMPRODUCT(--ISNUMBER(A1:A200))",
    "SUMPRODUCT(--ISERROR(1/A1:A200))",
    'SUMPRODUCT(LEN(A1:A200&"x"))',
    "SUMPRODUCT(ABS(-A1:A200))",
    "MATCH(TRUE,ISBLANK(A1:A200),0)",
    "SUMPRODUCT(--ISNA(MATCH(A1:A200,B1:B200,0)))",  # MATCH is not element-wise: dense
    'SUM(IF(A1:A200="",1))',  # IF in an array formula: a fill of 1
    'MAX(IF(A1:A200<>"",ROW(A1:A200)))',  # ROW makes no fill: dense
    "SUMPRODUCT(IF(A1:A200>5,A1:A200,0))",
    "COUNT(IF(A1:A200>5,A1:A200))",
    "MATCH(FALSE,IF(A1:A200>0,TRUE),0)",
    "SUMPRODUCT(IFERROR(1/A1:A200,7))",  # lazy forms other than IF: dense
    "SUMPRODUCT(IFNA(A1:A200*1,7))",
    "SUMPRODUCT(CHOOSE(1,A1:A200*1,0))",
    'IF(A1:A200="",1,2)',  # IF read by nothing that takes a padded array: dense
    'N(IF(A1:A200="",1,2))',
    "SUMPRODUCT(IFERROR(A1:A200,7)*1)",  # a bare range first: padded only if the form reads it
    "IF(A1:A200,1,2)",
    "SUM(IFNA(A1:A200,0))",
    "ROWS(A1:A200*1)",
    'COLUMNS(A1:B200&"")',
]


@pytest.mark.parametrize("formula", SEARCHES)
def test_searches_that_reach_the_fill(formula: str, monkeypatch: pytest.MonkeyPatch) -> None:
    """Data with no gap (A1:A10 = 1 to 10) and a result column with a value in every row (B), so a
    position in the fill's run shows in the answer."""
    cells = [(r, 0, float(r + 1)) for r in range(10)] + [(r, 1, 101.0 + r) for r in range(200)]
    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    padded = _values(cells, [formula])
    monkeypatch.setattr(Engine, "arr_padded", Engine.arr)
    assert same(padded[0], _values(cells, [formula])[0])


COUNTS = ["5", "60", "61", "70", "150", "199", "200", "250", "-5", "-61", "-150", "-199", "0"]
"""Row counts for ``TAKE`` and ``DROP``: inside the data, at its end, into the fill, past it."""


def _dynamic(rng: random.Random) -> str:
    """``UNIQUE``, ``SORT``, ``SORTBY``, ``FILTER``, ``TAKE`` or ``DROP`` over operands, bare
    ranges included; their keys and flags often equal the fill (``x=""``, ``--(x="")``)."""
    width = rng.choice((1, 1, 2))
    array = _operand(rng, width, rng.randrange(3))
    form = rng.randrange(6)
    if form == 0:
        extra = rng.choice(["", ",,TRUE", ",TRUE", ",FALSE,FALSE"])
        return f"_xlfn.UNIQUE({array}{extra})"
    if form == 1:
        index, order = rng.randint(1, width), rng.choice(["1", "-1"])
        extra = rng.choice(["", f",{index}", f",{index},{order}", f",{index},{order},FALSE"])
        return f"_xlfn._xlws.SORT({array}{extra})"
    if form == 2:
        keys = [f"{_operand(rng, 1, rng.randrange(3))},{rng.choice(['1', '-1'])}"]
        if rng.random() < 0.4:
            keys.append(f"{_operand(rng, 1, rng.randrange(3))},{rng.choice(['1', '-1'])}")
        return f"_xlfn.SORTBY({array},{','.join(keys)})"
    if form == 3:
        test = _operand(rng, 1, 1 + rng.randrange(2))
        extra = rng.choice(["", ',"none"', ",0"])
        return f"_xlfn._xlws.FILTER({array},{test}{extra})"
    name = rng.choice(["_xlfn.TAKE", "_xlfn.DROP"])
    columns = rng.choice(["", f",{rng.choice(['1', '-1', str(width)])}"])
    return f"{name}({array},{rng.choice(COUNTS)}{columns})"


def _results(
    cells: list[tuple[int, int, object]],
    formulas: list[str],
    limit: int | None = None,
    rows: int = 201,
) -> list[object]:
    """Every cell of each formula's result: each a legacy array formula over ``rows`` by 2."""
    e = Engine(Quirks.EXCEL, max_array_cells=limit)
    s = e.add_sheet("d")
    for r, c, v in cells:
        e.sheets[s].set(r, c, v)  # type: ignore[arg-type]
    for i, f in enumerate(formulas):
        col = 10 + 2 * i
        e.set_formula(s, 0, col, f, array=(False, (0, col, rows - 1, col + 1)))
    e.recalculate(full=True)
    return [
        e.get(s, r, 10 + 2 * i + c)
        for i in range(len(formulas))
        for r in range(rows)
        for c in (0, 1)
    ]


@pytest.mark.parametrize("seed", range(24))
def test_dynamic_functions_padded_and_dense_agree(
    seed: int, monkeypatch: pytest.MonkeyPatch
) -> None:
    """``UNIQUE``, ``SORT``, ``SORTBY``, ``FILTER``, ``TAKE`` and ``DROP`` read a padded array
    and give, cell for cell, what they give over the dense one: the blank counted once, blanks
    and errors in their order, a fill that sorts among the data or passes the filter."""
    rng = random.Random(1000 + seed)
    cells = _sheet(rng, full=seed % 2 == 1)
    formulas = [_dynamic(rng) for _ in range(30)]
    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    padded = _results(cells, formulas)
    monkeypatch.setattr(Engine, "arr_padded", Engine.arr)
    dense = _results(cells, formulas)
    per = 201 * 2
    wrong = [
        formulas[k // per]
        for k, (p, d) in enumerate(zip(padded, dense, strict=True))
        if not same(p, d)
    ]
    assert sorted(set(wrong)) == []


@pytest.mark.parametrize(
    "formula",
    [
        "_xlfn.UNIQUE(A1:A200)",
        '_xlfn.UNIQUE(A1:A200&"",,TRUE)',
        "_xlfn._xlws.SORT(A1:A200)",
        "_xlfn._xlws.SORT(-A1:B200,2,-1)",
        '_xlfn.SORTBY(A1:A200,B1:B200="",-1)',
        '_xlfn._xlws.FILTER(A1:A200,B1:B200="")',
        "_xlfn._xlws.FILTER(A1:B200,ISNUMBER(A1:A200))",
        "_xlfn.TAKE(A1:A200*1,150)",
        '_xlfn.DROP(A1:A200&"",-61)',
    ],
)
def test_each_dynamic_function_reads_a_padded_array(
    formula: str, monkeypatch: pytest.MonkeyPatch
) -> None:
    """The comparison above is not vacuous: each function is given a padded array."""
    made: list[Padded] = []
    original = Engine.arr_padded

    def spy(self: Engine, v: object) -> object:
        out = original(self, v)
        if type(out) is Padded:
            made.append(out)
        return out

    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    monkeypatch.setattr(Engine, "arr_padded", spy)
    _results([(r, c, float(r % 7)) for r in range(10) for c in (0, 1)], [formula])
    assert made


OTHER_ARGUMENTS = [
    '_xlfn.TAKE(A1:A200,--(B1:B200=""))',
    '_xlfn.DROP(A1:B200,1,--(B1:B200=""))',
    '_xlfn.UNIQUE(A1:A200,B1:B200="")',
    '_xlfn.UNIQUE(A1:A200,,B1:B200="")',
    '_xlfn._xlws.SORT(A1:A200,--(B1:B200=""))',
    '_xlfn._xlws.SORT(A1:A200,1,--(B1:B200=""))',
    '_xlfn._xlws.SORT(A1:A200,1,1,B1:B200="")',
    '_xlfn.SORTBY(A1:A200,B1:B200,--(B1:B200=""))',
    '_xlfn._xlws.FILTER(A1:A200,ISNUMBER(A1:A200)*0,B1:B200&"")',
]
"""Padded arrays where the functions read one value (an array there is ``#VALUE!``, padded or
not), a list (made dense), or the result when nothing passes (made dense)."""


def test_other_arguments_padded_and_dense_agree(monkeypatch: pytest.MonkeyPatch) -> None:
    rng = random.Random(7)
    cells = _sheet(rng, full=False)
    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    padded = _results(cells, OTHER_ARGUMENTS)
    monkeypatch.setattr(Engine, "arr_padded", Engine.arr)
    dense = _results(cells, OTHER_ARGUMENTS)
    per = 201 * 2
    wrong = {
        OTHER_ARGUMENTS[k // per]
        for k, (p, d) in enumerate(zip(padded, dense, strict=True))
        if not same(p, d)
    }
    assert wrong == set()


LIMITED = [
    "_xlfn._xlws.SORT(A1:A200)",
    "_xlfn.SORTBY(A1:A200,B1:B200)",
    "_xlfn._xlws.FILTER(A1:A200,ISBLANK(A1:A200)=ISBLANK(A1:A200))",
    "_xlfn.TAKE(A1:A200*1,180)",
    "_xlfn.DROP(A1:A200*1,10)",
    "_xlfn.UNIQUE(A1:A200)",
    "_xlfn.TAKE(A1:A200*1,50)",
]
"""Results over the limit of 150 cells (the first five) and under it (the last two)."""


def test_a_result_past_the_limit_stops_as_a_dense_one_does(monkeypatch: pytest.MonkeyPatch) -> None:
    """A padded array is within the limit where its dense form is not; a result made dense from
    it must stop where the dense computation stops, at the limit on an array's cells."""
    cells = [(r, c, float(r % 7)) for r in range(10) for c in (0, 1)]
    monkeypatch.setattr(engine_module, "_PAD_ROWS", 64)
    padded = _results(cells, LIMITED, limit=150, rows=1)[::2]
    monkeypatch.setattr(Engine, "arr_padded", Engine.arr)
    dense = _results(cells, LIMITED, limit=150, rows=1)[::2]
    first = padded
    assert [same(p, d) for p, d in zip(padded[:5], dense[:5], strict=True)] == [True] * 5
    assert first[5:] == [0.0, 0.0]  # padded, these compute where the dense ones stop
