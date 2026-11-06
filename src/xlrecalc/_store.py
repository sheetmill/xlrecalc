"""The cell store: values and template ids per sheet (ARCHITECTURE.md §7.5).

Keys pack a 0-based position as ``row * 16384 + col``. Each sheet also keeps, per column, the
sorted rows that hold anything, so a box (a whole column, say) yields its non-blank cells
without visiting a million empty positions. The rows are typed arrays (4 bytes a row), not
lists of Python integers (8 bytes of pointer plus, above 256, a 28-byte object).
"""

from __future__ import annotations

import bisect
from array import array
from collections.abc import Iterator

from xlrecalc._values import BLANK, Value

__all__ = ["COLS", "MAX_COL", "MAX_ROW", "Sheet", "TemplateIds"]

COLS = 16384
MAX_ROW = 1048575  # 0-based
MAX_COL = 16383


_NONE = -1
_SLACK = 4096  # dense columns may hold this many unused positions beyond twice their cells


class TemplateIds:
    """The template id of each formula cell, keyed like the values (``row * 16384 + col``).

    Per column, a dense 32-bit array over the span of the column's formula rows, ``-1`` where a
    row holds none: 4 bytes a cell, where a dict costs about 90 with the key's integer. A column
    whose span grows past twice its cells (plus some slack) becomes a dict of row to id.
    """

    __slots__ = ("_base", "_count", "_counts", "_dense", "_sparse")

    def __init__(self) -> None:
        self._dense: dict[int, array[int]] = {}  # col -> ids from row _base[col]
        self._base: dict[int, int] = {}
        self._sparse: dict[int, dict[int, int]] = {}  # col -> {row: id}
        self._counts: dict[int, int] = {}  # col -> formula cells
        self._count = 0

    def at(self, row: int, col: int) -> int | None:
        """The id at (row, col), or ``None``."""
        ids = self._dense.get(col)
        if ids is not None:
            i = row - self._base[col]
            if 0 <= i < len(ids):
                tid = ids[i]
                if tid >= 0:
                    return tid
            return None
        sparse = self._sparse.get(col)
        return None if sparse is None else sparse.get(row)

    def get(self, key: int, default: int | None = None) -> int | None:
        row, col = divmod(key, COLS)
        tid = self.at(row, col)
        return default if tid is None else tid

    def __getitem__(self, key: int) -> int:
        tid = self.get(key)
        if tid is None:
            raise KeyError(key)
        return tid

    def __contains__(self, key: object) -> bool:
        return isinstance(key, int) and self.get(key) is not None

    def __len__(self) -> int:
        return self._count

    def __iter__(self) -> Iterator[int]:
        """Keys, column by column."""
        for col, ids in self._dense.items():
            base = self._base[col]
            for i, tid in enumerate(ids):
                if tid >= 0:
                    yield (base + i) * COLS + col
        for col, rows in self._sparse.items():
            for row in rows:
                yield row * COLS + col

    def keys(self) -> Iterator[int]:
        return iter(self)

    def put(self, row: int, col: int, tid: int) -> bool:
        """Set the id at (row, col); whether the cell held none before."""
        ids = self._dense.get(col)
        if ids is None:
            sparse = self._sparse.get(col)
            if sparse is None:
                self._dense[col] = array("i", (tid,))
                self._base[col] = row
                self._added(col)
                return True
            new = row not in sparse
            sparse[row] = tid
            if new:
                self._added(col)
            return new
        i = row - self._base[col]
        if 0 <= i < len(ids):
            new = ids[i] < 0
            ids[i] = tid
        elif i == len(ids):
            ids.append(tid)
            new = True
        else:
            span = i + 1 if i > 0 else len(ids) - i
            if span > 2 * (self._counts[col] + 1) + _SLACK:
                self._sparsify(col)
                return self.put(row, col, tid)
            if i > 0:
                ids.extend(array("i", (_NONE,)) * (i - len(ids)))
                ids.append(tid)
            else:
                ids[:0] = array("i", (_NONE,)) * -i
                ids[0] = tid
                self._base[col] = row
            new = True
        if new:
            self._added(col)
        return new

    def remove(self, row: int, col: int) -> bool:
        """Clear the id at (row, col); whether there was one."""
        ids = self._dense.get(col)
        if ids is not None:
            i = row - self._base[col]
            if not (0 <= i < len(ids) and ids[i] >= 0):
                return False
            ids[i] = _NONE
        else:
            sparse = self._sparse.get(col)
            if sparse is None or sparse.pop(row, None) is None:
                return False
        self._count -= 1
        self._counts[col] -= 1
        return True

    def _added(self, col: int) -> None:
        self._count += 1
        self._counts[col] = self._counts.get(col, 0) + 1

    def _sparsify(self, col: int) -> None:
        ids = self._dense.pop(col)
        base = self._base.pop(col)
        self._sparse[col] = {base + i: tid for i, tid in enumerate(ids) if tid >= 0}


class Sheet:
    """One worksheet's cells."""

    __slots__ = (
        "columns",
        "filter_mode",
        "formula_columns",
        "hidden_rows",
        "name",
        "templates",
        "values",
    )

    def __init__(self, name: str) -> None:
        self.name = name
        self.hidden_rows: frozenset[int] = frozenset()  # 0-based, as the file left them
        self.filter_mode = False  # a filter is hiding rows: SUBTOTAL 1 to 11 leave out hidden ones
        self.values: dict[int, Value] = {}
        self.templates = TemplateIds()  # formula cells: key -> template id
        self.columns: dict[int, array[int]] = {}  # col -> sorted rows holding a value or formula
        self.formula_columns: dict[int, array[int]] = {}  # col -> sorted rows holding a formula

    def get(self, row: int, col: int) -> Value:
        return self.values.get(row * COLS + col, BLANK)

    def occupy(self, row: int, col: int) -> None:
        _insert(self.columns, row, col)

    def set_template(self, row: int, col: int, template: int | None) -> None:
        if template is None:
            if self.templates.remove(row, col):
                rows = self.formula_columns[col]
                del rows[bisect.bisect_left(rows, row)]
            return
        if self.templates.put(row, col, template):
            _insert(self.formula_columns, row, col)
        self.occupy(row, col)

    def formulas_in(self, r1: int, c1: int, r2: int, c2: int) -> list[tuple[int, int]]:
        """Formula cells in the box."""
        out: list[tuple[int, int]] = []
        cols = self.formula_columns
        for col in range(c1, c2 + 1) if c2 - c1 < len(cols) else sorted(cols):
            rows = cols.get(col)
            if rows and c1 <= col <= c2:
                lo, hi = bisect.bisect_left(rows, r1), bisect.bisect_right(rows, r2)
                out.extend((r, col) for r in rows[lo:hi])
        return out

    def set(self, row: int, col: int, value: Value) -> None:
        key = row * COLS + col
        if value is BLANK:
            self.values.pop(key, None)
        else:
            self.values[key] = value
        self.occupy(row, col)

    def positions(self, r1: int, c1: int, r2: int, c2: int) -> Iterator[tuple[int, int]]:
        """Occupied positions in the box, row-major."""
        found: list[tuple[int, int]] = []
        for col in range(c1, c2 + 1) if c2 - c1 < len(self.columns) else sorted(self.columns):
            if not c1 <= col <= c2:
                continue
            rows = self.columns.get(col)
            if not rows:
                continue
            lo, hi = bisect.bisect_left(rows, r1), bisect.bisect_right(rows, r2)
            if c1 == c2:
                yield from ((r, col) for r in rows[lo:hi])
                return
            found.extend((r, col) for r in rows[lo:hi])
        found.sort()
        yield from found

    def box_values(self, r1: int, c1: int, r2: int, c2: int) -> list[Value]:
        """Non-blank values of the box, row-major: what an aggregate reads, as one list."""
        if c1 == c2:
            rows = self.columns.get(c1)
            if not rows:
                return []
            lo, hi = bisect.bisect_left(rows, r1), bisect.bisect_right(rows, r2)
            get = self.values.get
            return [v for r in rows[lo:hi] if (v := get(r * COLS + c1)) is not None]
        return [v for _, _, v in self.box(r1, c1, r2, c2)]

    def box(self, r1: int, c1: int, r2: int, c2: int) -> Iterator[tuple[int, int, Value]]:
        """Non-blank cells of the box, row-major, as (row, col, value)."""
        values = self.values
        for r, c in self.positions(r1, c1, r2, c2):
            v = values.get(r * COLS + c)
            if v is not None:
                yield r, c, v


def _insert(index: dict[int, array[int]], row: int, col: int) -> None:
    rows = index.get(col)
    if rows is None:
        index[col] = array("i", (row,))  # rows fit in 32 bits (MAX_ROW < 2**20)
    elif row > rows[-1]:
        rows.append(row)
    else:
        i = bisect.bisect_left(rows, row)
        if i == len(rows) or rows[i] != row:
            rows.insert(i, row)
