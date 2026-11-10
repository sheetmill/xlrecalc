# Benchmark summary

Machine: Apple M4 Pro, 24 GB. Every row in the tables below ran on CPython 3.11.18, the
package's own Python, each in a fresh process: the package's runners in its environment, the
compact prototype with the standard library alone, and the other engines (pycel 1.0b30, formulas
1.3.4, xlcalculator 0.5.0) in a temporary environment holding only them and openpyxl 3.1.5. Raw
output: `current_py311.jsonl` (the package's rows) and `baseline_py311.jsonl` (the prototype, the
other engines, and the package before the store changes of "Memory and speed of the store"
below).

Workbook: N data rows, 3N+3 formulas. Columns: A constants; B `=A*2`; C running total `=C(r-1)+B(r)`;
D `=SUM(A r:A r+9)` sliding window; F1..F3 `=SUM(...)` over whole columns.

```bash
cd bench && uv sync
uv run python gen.py 1000 10000 100000 1000000   # data/bench_<N>.xlsx and .pkl
cd ..                                            # the package's runners, from the repository root
uv run python bench/engine_bench.py 10000 [--xlsx]
uv run python bench/memory_breakdown.py 1000000 [--xlsx]
cd bench                                         # the prototype and the other engines, on 3.11
uv run --no-project --python 3.11 python proto_compact.py 10000
uv run --no-project --python 3.11 --with pycel==1.0b30 --with formulas==1.3.4 \
    --with xlcalculator==0.5.0 --with openpyxl==3.1.5 python run_competitor.py pycel 10000
```

Each runner prints one JSON line with timings and the process's memory (`rss.py`: on macOS
the physical footprint, what Activity Monitor shows, as resident memory alone drops whenever the
system compresses pages) at the same points: at the start, after reading the pickled table (for the runners that start from
it), after loading, and the peak. The compact prototype is an upper bound, not a forecast: a
restricted grammar (references, ranges, `+ - * /`, `SUM`), one sheet, no errors or coercion. The
engine, the prototype and pycel compute the same values, Excel's plain left-to-right sums;
formulas sums with compensation (996035.662 against 996035.6619999994).

"Full" = load + evaluate everything. "A1" = change A1 (dirties the whole C chain + summaries).
"Alast" = change A in the last row (15 dependents). Peak = the process's largest footprint.
The engine reads either the pickled table (as the prototype) or openpyxl's xlsx file through its
own reader (as the other engines).

## 10k rows (30,003 formulas)

| engine | full | re-eval A1 | re-eval Alast | peak |
|---|---|---|---|---|
| pycel, xlsx | 4.61 s | 0.06 s | – | 437 MB |
| formulas, xlsx | 24.2 s | 8.0 s | – | 1,855 MB |
| xlcalculator, xlsx (1k rows) | 11.8 s | 11.7 s | – | 1,743 MB |
| compact prototype, pickle | 0.10 s | 0.09 s | 0.01 s | 23 MB |
| xlrecalc, pickle | 0.18 s | 0.04 s | 0.01 s | 31 MB |
| xlrecalc, xlsx | 0.26 s | 0.04 s | 0.01 s | 31 MB |

## 100k rows (300,003 formulas)

| engine | full | re-eval A1 | re-eval Alast | peak |
|---|---|---|---|---|
| pycel, xlsx | 45.4 s | 0.75 s | – | 3,920 MB |
| compact prototype, pickle | 1.22 s | 1.02 s | 0.06 s | 139 MB |
| xlrecalc, pickle | 1.99 s | 0.42 s | 0.07 s | 138 MB |
| xlrecalc, xlsx | 2.69 s | 0.42 s | 0.07 s | 128 MB |

formulas and xlcalculator were not run at this size (their memory extrapolates to more than
15 GB).

## 1M rows (3,000,003 formulas)

| engine | load | eval | re-eval A1 | re-eval Alast | memory the load adds | peak |
|---|---|---|---|---|---|---|
| compact prototype, pickle | 5.65 s | 8.23 s | 11.3 s | 0.92 s | 513 MB | 1,297 MB |
| xlrecalc, pickle | 7.83 s | 13.6 s | 4.99 s | 0.95 s | 486 MB | 1,400 MB |
| xlrecalc, xlsx | 17.2 s | 13.8 s | 5.09 s | 1.01 s | 608 MB | 1,278 MB |

"Memory the load adds" is the footprint after loading less the footprint before it:
after reading the pickled table for the pickle rows (the table, 420 MB, is still held at both
points), at the start for the xlsx row. Peak includes the table for the pickle rows. Computing
the dirty set of "change A1" takes 0.14 ms (1,000,005 cells as 4 runs) against the compact
prototype's 7.58 s; building the dependency index once takes 0.4 s. The engine evaluates
Excel's semantics (coercion, errors, references, snapping), which the prototype does not.

**Where the memory goes** (`memory_breakdown.py 1000000`, traced by `tracemalloc`; each
structure with every object inside it, each object counted once; Python's own
allocations, so below the footprint, which also holds memory freed but not returned to the
system):

| structure | after load | after the full recalculation |
|---|---|---|
| cell values (`Sheet.values`) | 98 MB | 391 MB |
| template id per formula cell (`Sheet.templates`, typed arrays per column) | 12 MB | 12 MB |
| occupied rows per column (`Sheet.columns`, typed arrays) | 16 MB | 16 MB |
| formula rows per column (`Sheet.formula_columns`, typed arrays) | 12 MB | 12 MB |
| stale formulas (`stale`, with the cells' integers) | 229 MB | 0 MB |
| stale formulas whose dependents may not be stale yet (`_unclosed`, a list) | 24 MB | 0 MB |
| templates, template keys, the report, the dependency index | < 1 MB | < 1 MB |

Traced in all: 397 MB after load (both sources, the table dropped), 448 MB after the full
recalculation, 682 MB at the recalculation's peak. The cell values are a dict of float
objects; typed value blocks are planned for a later minor release (ARCHITECTURE.md §7.5).

## Memory and speed of the store

The store, scheduler and aggregate changes measured against the baseline, 1M rows, the same
runners and Python (`baseline_py311.jsonl` before, `current_py311.jsonl` after). The changes,
in order: one stale set (formulas without a cached value recognized by an absent value, the not
yet closed ones in a list, the emptied set cleared after a recalculation); the per-column row
lists as 32-bit typed arrays; the template ids per column in typed arrays; float fast paths in
the arithmetic operators; the scheduler taking its targets one at a time, from the store when
most formulas are stale; aggregates reading a reference's values as one list, the registry
passing uncoerced arguments straight through, and `SUM` of one reference in a single loop.
Every computed value and every written file is unchanged, bit for bit (`bench/golden.py
--large`).

| 1M rows | before | after |
|---|---|---|
| pickle: load / eval / re-eval A1 / re-eval Alast | 8.59 / 18.8 / 7.02 / 2.01 s | 7.83 / 13.6 / 4.99 / 0.95 s |
| xlsx: load / eval / re-eval A1 / re-eval Alast | 25.1 / 19.0 / 7.28 / 2.26 s | 17.2 / 13.8 / 5.09 / 1.01 s |
| footprint the load adds, pickle / xlsx | 1,429 / 1,313 MB | 486 / 608 MB |
| peak footprint, pickle / xlsx | 1,946 / 1,843 MB | 1,400 / 1,278 MB |
| traced after load (table dropped), pickle / xlsx | 962 / 1,062 MB | 397 / 397 MB |
| traced after the full recalculation, pickle / xlsx | 1,037 / 1,137 MB | 448 / 448 MB |
| traced peak of the full recalculation, pickle / xlsx | 1,472 / 1,572 MB | 682 / 682 MB |
| `recalculate(path)`: load / eval / write / total | 24.8 / 18.6 / 21.0 / 64.3 s | 17.2 / 13.2 / 20.2 / 50.6 s |
| `recalculate(path)`: peak footprint | 2,156 MB | 1,163 MB |

The compact prototype's load adds 513 MB with the same measurement and evaluates in 8.2 s; the
engine computes Excel's semantics where the prototype adds floats. Before these changes, a
profile at 100k rows put about 35% of the full recalculation in the scheduler (`_evaluate`,
`precedents`) and about 37% in the sliding `SUM`'s walk through `Ref`, `Sheet.box` and
`positions`, the operators a few percent. What remains of the recalculation's transient memory
(234 MB traced over the 448 MB held after it) is the stale set itself and the growth of the
values dict.

## Earlier designs (not re-run)

Measured on Python 3.13 before the package existed. The rows of the other engines and of the
compact prototype were measured then too; the tables above, on Python 3.11, replace those.

| engine, 10k rows | full | re-eval A1 | re-eval Alast | peak |
|---|---|---|---|---|
| SQLite, phased | 106.9 s | 287.5 s | n/a | 78 MB |
| proto: pure python dicts | 0.37 s | 0.018 s | 0.007 s | 79 MB |
| proto: + sqlite R-tree index | 0.44 s | 0.283 s | 0.006 s | 82 MB |
| proto: values in sqlite | 0.43 s | 0.036 s | 0.009 s | 78 MB |

The "SQLite, phased" design evaluated formulas from SQL in level-synchronous phases, with
operands marshalled through JSON; its code is not included here. At 1k rows: full 1.7 s, A1
re-eval 2.8 s; a 15-cell change costs 27 ms flagging + 20 ms eval (both scale with total formula
count, not with the size of the change). The pure-dicts prototype at 1M rows: load 40.4 s, eval
10.2 s, peak 4,990 MB. Raw output: `proto_10k_100k.jsonl`, `proto_1m.jsonl`,
`competitors_10000.jsonl`, `proto_compact.jsonl`, `engine_10k_100k.jsonl`, `engine_1m.jsonl`,
and `round2.txt` (the SQLite, phased design at 10k rows and pycel at 100k rows,
the runs these tables first came from).

## Structural micro-benchmarks

| measurement                                        | result                                  |
|----------------------------------------------------|-----------------------------------------|
| 4M cell values: dict / sqlite table / dense array  | 148 / 20 / 8.5 bytes per cell           |
| JSON operand marshalling (json+str<->type), 10 operands | 2.1 µs per cell (direct python: 0.05 µs) |
| sqlite -> python UDF call round trip               | 0.06 µs                                 |
| range reverse lookup, 100k ranges: R-tree / python buckets | 8.4 / 5.4 µs per lookup; build 289 / 17 ms |
| floats stored as text `f'{v:1.15}'`                | lossy (0.1+0.2 round-trips to 0.3)      |

## Reading workbooks

`uv run python bench/read_bench.py FILE_OR_DIRECTORY ... [--sample N]`, run from the repository
root in the package's environment: every cell of every worksheet, each reader in a fresh process,
Apple M4 Pro, Python 3.11, openpyxl 3.1.5. The real-world sample is 300 workbooks drawn with seed 1
from a local corpus (`CAPTURING.md` §4). `bench/read_compare.py` compares what the two readers
read, cell by cell.

| files | xlrecalc reader (formulas and values) | openpyxl read-only (formulas) | openpyxl read-only (values) | openpyxl normal |
|---|---|---|---|---|
| bench_10000.xlsx | 0.07 s, 25 MB | 0.12 s, 32 MB | 0.10 s, 32 MB | 0.16 s, 56 MB |
| bench_100000.xlsx | 0.59 s, 25 MB | 0.78 s, 40 MB | 0.71 s, 40 MB | 1.60 s, 279 MB |
| 300 real-world workbooks, 35 MB | 5.9 s, 28 MB | 10.5 s, 118 MB | 10.5 s, 112 MB | 149 s, 1,796 MB |

## Recalculating a file and writing the values back

`uv run python bench/recalc_bench.py N`, run from the repository root in the package's
environment (Python 3.11, Apple M4 Pro): openpyxl's `bench_<N>.xlsx` (no cached values, so every
formula is evaluated) is loaded with the package's reader, evaluated, and written in place by the
surgical writer, which patches every formula cell; then written once more, now holding its
values, which patches nothing. Raw output in `current_py311.jsonl`. The write takes about 7%
longer than when the writer was first measured (18.9 s at 1M rows in alternating runs on the
same machine; the writer's first measurement is in `recalc.jsonl`), the cost of the edits it
now applies while streaming.

| rows | formulas | load | evaluate | write | total | write, nothing changed | file before / after | peak |
|---|---|---|---|---|---|---|---|---|
| 1,000 | 3,003 | 0.02 s | 0.01 s | 0.02 s | 0.05 s | 0.01 s | 0.04 / 0.06 MB | 29 MB |
| 10,000 | 30,003 | 0.15 s | 0.11 s | 0.19 s | 0.46 s | 0.13 s | 0.3 / 0.5 MB | 39 MB |
| 100,000 | 300,003 | 1.48 s | 1.20 s | 1.96 s | 4.64 s | 1.34 s | 3.3 / 5.1 MB | 129 MB |
| 1,000,000 | 3,000,003 | 17.2 s | 13.2 s | 20.2 s | 50.6 s | 13.8 s | 33.7 / 50.7 MB | 1,163 MB |

`bench/readers_check.py`, run in the bench environment, reads each written file with
`pandas.read_excel` and with openpyxl's data-only mode and compares every formula cell with
values computed independently from column A (plain left-to-right sums, as Excel sums): all
3,003, 30,003, 300,003 and 3,000,003 cells equal, with both readers. Excel opens the written
files without a repair, and its own recalculation of them agrees with every written value
(`tests/test_excel_live.py`, up to 100,000 rows; the million-row file was opened by hand with
`capture/open_read.applescript`).
