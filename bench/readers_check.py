"""Do pandas and openpyxl's data-only mode see the values ``recalculate`` wrote?

    cd bench && uv run python readers_check.py FILE ...

Runs in the bench environment (pandas and openpyxl). Each FILE is a benchmark workbook
(``gen.py``'s layout) after ``recalculate``. The expected values are computed here from column
A with plain left-to-right additions, as Excel sums: B = 2A, C a running total of B, D the sum
of A over the row and the next nine, F1 to F3 the sums of columns B to D. Prints, per reader,
how many formula cells hold the expected value exactly.
"""

from __future__ import annotations

import sys
import time

import openpyxl
import pandas as pd


def expected(a: list[float]) -> dict[tuple[int, int], float]:
    n = len(a)
    out: dict[tuple[int, int], float] = {}
    b = [x * 2 for x in a]
    running = 0.0
    for r in range(n):
        running = b[0] if r == 0 else running + b[r]
        out[r, 1] = b[r]
        out[r, 2] = running
        total = 0.0
        for x in a[r : min(r + 10, n)]:
            total += x
        out[r, 3] = total
    for k, col in enumerate((1, 2, 3)):
        total = 0.0
        for r in range(n):
            total += out[r, col]
        out[k, 5] = total
    return out


def check(name: str, rows: list[list[object]]) -> None:
    a = [float(row[0]) for row in rows]  # type: ignore[arg-type]
    want = expected(a)
    good = sum(1 for (r, c), v in want.items() if rows[r][c] == v)
    print(f"  {name}: {good} of {len(want)} formula cells hold the expected value")


def main(paths: list[str]) -> None:
    for path in paths:
        print(path)
        t = time.perf_counter()
        frame = pd.read_excel(path, header=None)
        rows = frame.where(frame.notna(), None).values.tolist()
        check(f"pandas.read_excel ({time.perf_counter() - t:.1f} s)", rows)
        t = time.perf_counter()
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        rows = [list(r) + [None] * (6 - len(r)) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        check(f"openpyxl data_only ({time.perf_counter() - t:.1f} s)", rows)


if __name__ == "__main__":
    main(sys.argv[1:])
