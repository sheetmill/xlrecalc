"""The dependency index agrees with brute force (ARCHITECTURE.md §6).

Random workbooks mix every reference shape: relative cells and boxes, running totals,
absolute cells, growing ranges (``$A$1:A5``), whole columns and rows, other sheets and defined
names. For random changed cells, the index's dependents must equal the transitive closure found
by checking every formula's boxes directly.
"""

from __future__ import annotations

import random

import pytest

from xlrecalc._engine import Engine, node

SHAPES = [
    "=A{r}*2",
    "=SUM(A{r}:A{r2})",
    "=B{rp}+1",
    "=C{rp}+B{r}",
    "=$A$1+B{r}",
    "=SUM($A$1:A{r})",
    "=SUM(A:A)",
    "=SUM(3:3)",
    "=Other!A{r}*3",
    "=SUM(Data)+D{rp}",
    "=A{r}+$B$3",
    "=SUM(B{r}:C{r})",
    "=E{r2}+1",
]


def build(seed: int) -> Engine:
    rng = random.Random(seed)
    sheets: dict[str, dict[str, object]] = {"S": {}, "Other": {}}
    for r in range(1, 41):
        sheets["S"][f"A{r}"] = float(r)
        sheets["Other"][f"A{r}"] = float(r)
    for col in "BCDE":
        start = rng.randint(1, 5)
        for r in range(start, start + rng.randint(5, 30)):
            if rng.random() < 0.15:
                continue  # gaps break runs
            shape = rng.choice(SHAPES)
            text = shape.format(r=r, rp=max(r - 1, 1), r2=r + rng.randint(0, 3))
            sheets["S"][f"{col}{r}"] = text
    engine = Engine.from_cells(sheets)
    engine.define_name("Data", "S!$A$2:$A$6")
    return engine


def brute(engine: Engine, changed: list[int]) -> set[int]:
    formulas = [(s << 34) | key for s, sheet in enumerate(engine.sheets) for key in sheet.templates]
    boxes = {f: list(engine._boxes(f)) for f in formulas}
    found: set[int] = set()
    queue = list(changed)
    while queue:
        n = queue.pop()
        sheet, rest = n >> 34, n & ((1 << 34) - 1)
        row, col = divmod(rest, 16384)
        for f, bs in boxes.items():
            if f in found:
                continue
            if any(s == sheet and r1 <= row <= r2 and c1 <= col <= c2 for s, r1, c1, r2, c2 in bs):
                found.add(f)
                queue.append(f)
    return found


@pytest.mark.parametrize("seed", range(40))
def test_index_agrees_with_brute_force(seed):
    engine = build(seed)
    rng = random.Random(seed + 1000)
    for _ in range(5):
        sheet = rng.choice([0, 0, 0, 1])
        changed = [
            node(sheet, rng.randint(0, 45), rng.randint(0, 5)) for _ in range(rng.randint(1, 3))
        ]
        assert engine.dependents(changed) == brute(engine, changed)


def test_values_follow_a_change_through_the_index():
    engine = build(7)
    engine.recalculate()
    before = {(r, c): engine.get(0, r, c) for r in range(45) for c in range(1, 5)}
    engine.set_value(0, 2, 0, 1000.0)  # A3
    after_index = {(r, c): engine.get(0, r, c) for r in range(45) for c in range(1, 5)}
    fresh = build(7)
    fresh.sheets[0].set(2, 0, 1000.0)
    fresh.recalculate(full=True)
    expected = {(r, c): fresh.get(0, r, c) for r in range(45) for c in range(1, 5)}
    assert after_index == expected
    assert after_index != before


VOLATILE_SHAPES = ["=SUM(Vol)+B{rp}", "=VOL*2", "=Vol+Plain+C{r}"]


def build_volatile(seed: int) -> tuple[Engine, set[int]]:
    """A random workbook whose formulas use a volatile name too, and a random set of dormant
    cells among those that do: a few, many, or many with cells of another sheet, so that both
    ways of leaving them out of a run are taken."""
    rng = random.Random(seed)
    sheets: dict[str, dict[str, object]] = {"S": {}, "Other": {}}
    for r in range(1, 41):
        sheets["S"][f"A{r}"] = float(r)
        sheets["Other"][f"A{r}"] = float(r)
    users: list[int] = []
    for c, col in enumerate("BCDE", start=1):
        start = rng.randint(1, 5)
        shape = ""
        for r in range(start, start + rng.randint(5, 30)):
            if rng.random() < 0.1:
                continue
            if not shape or rng.random() < 0.2:  # runs of one template, several cells long
                shape = rng.choice(VOLATILE_SHAPES if rng.random() < 0.5 else SHAPES)
            if shape in VOLATILE_SHAPES:
                users.append(node(0, r - 1, c))
            sheets["S"][f"{col}{r}"] = shape.format(r=r, rp=max(r - 1, 1), r2=r + rng.randint(0, 3))
    engine = Engine.from_cells(sheets)
    engine.define_name("Data", "S!$A$2:$A$6")
    engine.define_name("Vol", "OFFSET(S!$A$10,0,0,3,1)+S!$A$20")
    engine.define_name("Plain", "S!$A$30")
    dormant = {n for n in users if rng.random() < 0.6}
    mode = seed % 3
    if mode == 1 and dormant:  # fewer than a run holds: the set is scanned
        dormant = set(rng.sample(sorted(dormant), min(2, len(dormant))))
        awake = sorted(set(users) - dormant)
        for n in rng.sample(awake, min(2, len(awake))):  # an awake cell's place on another sheet
            dormant.add(n | (1 << 34))
    elif mode == 2:  # more than a run holds, in the same columns on another sheet: rows scanned
        dormant |= {node(1, r, c) for r in range(30) for c in range(1, 5)}
    engine.dormant.update(dormant)
    return engine, dormant


def brute_volatile(engine: Engine, changed: list[int], dormant: set[int]) -> set[int]:
    """``brute``, a dormant cell's boxes being those of its own text and its plain names."""
    vol = {(0, 9, 0, 9, 0), (0, 19, 0, 19, 0)}  # Vol's boxes: A10 (OFFSET's reference), A20
    formulas = [(s << 34) | key for s, sheet in enumerate(engine.sheets) for key in sheet.templates]
    boxes = {f: [b for b in engine._boxes(f) if f not in dormant or b not in vol] for f in formulas}
    found: set[int] = set()
    queue = list(changed)
    while queue:
        n = queue.pop()
        sheet, rest = n >> 34, n & ((1 << 34) - 1)
        row, col = divmod(rest, 16384)
        for f, bs in boxes.items():
            if f in found:
                continue
            if any(s == sheet and r1 <= row <= r2 and c1 <= col <= c2 for s, r1, c1, r2, c2 in bs):
                found.add(f)
                queue.append(f)
    return found


@pytest.mark.parametrize("seed", range(40))
def test_a_dormant_cell_has_no_edges_from_a_volatile_name(seed):
    """ARCHITECTURE.md §6.3: a volatile name's references reach no dormant cell; its own text,
    and a plain name, do."""
    engine, dormant = build_volatile(seed)
    rng = random.Random(seed + 2000)
    for _ in range(6):
        changed = [
            node(0, rng.choice([9, 10, 11, 19, 29, rng.randint(0, 45)]), rng.randint(0, 4))
            for _ in range(rng.randint(1, 3))
        ]
        assert engine.dependents(changed) == brute_volatile(engine, changed, dormant)
