"""Generate benchmark workbooks.

Columns (N data rows):
  A: constant floats
  B: =A{r}*2                      (one single-cell ref per formula)
  C: =C{r-1}+B{r}  (C1: =B1)      (deep dependency chain / running total)
  D: =SUM(A{r}:A{r+9})            (N distinct overlapping small ranges)
  F1: =SUM(B1:BN)  F2: =SUM(C1:CN)  F3: =SUM(D1:DN)   (summary cells pulling in everything)

Usage: gen.py N [N ...] [--pickle-only]
Writes data/bench_<N>.xlsx and data/bench_<N>.pkl, the same table as a pickled list of rows (with
lower-case function names), which the prototypes load. --pickle-only skips the xlsx; the 1M-row
table is made that way, since openpyxl would take minutes to write it.
"""
import sys, random, pickle
from pathlib import Path
from openpyxl import Workbook

def formulas(N, fn='SUM'):
    rows = []
    for r in range(1, N + 1):
        a = round(random.random() * 100, 3)
        b = f'=A{r}*2'
        c = '=B1' if r == 1 else f'=C{r-1}+B{r}'
        d = f'={fn}(A{r}:A{min(r+9, N)})'
        f = {1: f'={fn}(B1:B{N})', 2: f'={fn}(C1:C{N})', 3: f'={fn}(D1:D{N})'}.get(r)
        rows.append([a, b, c, d, None, f] if f else [a, b, c, d])
    return rows

def main(N, seed=1, pickle_only=False):
    out = Path('data')
    out.mkdir(exist_ok=True)
    if not pickle_only:
        write_xlsx(N, seed, out)
    random.seed(seed)
    with open(out / f'bench_{N}.pkl', 'wb') as f:
        pickle.dump(formulas(N, fn='sum'), f)
    print('wrote', N)

def write_xlsx(N, seed, out):
    random.seed(seed)
    rows = formulas(N)
    wb = Workbook(write_only=True)  # the same cells as normal mode, in a fraction of the time
    ws = wb.create_sheet('Sheet1')
    for row in rows:
        ws.append(row)
    wb.save(out / f'bench_{N}.xlsx')

if __name__ == '__main__':
    args = sys.argv[1:]
    for n in (int(a) for a in args if a != '--pickle-only'):
        main(n, pickle_only='--pickle-only' in args)
