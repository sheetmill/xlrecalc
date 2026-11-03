"""Prototypes that compare plain Python structures with keeping the dependency index or the values in SQLite.

Usage: proto.py <mode> <N>
  mode = py            : values in a dict, dependency index in Python dicts/lists
         sqlite_index  : values in a dict, dependency index in SQLite (R-tree), dirty propagation in SQL
         sqlite_values : dependency index in Python, but cell VALUES live in SQLite and are read per access
                         (approximates what an SQL-driven evaluator pays for reading operands out of the database)

Only supports the benchmark grammar: cell refs, ranges, + - * /, SUM().  Not an Excel parser.
"""
import sys, re, time, json, pickle, resource, sqlite3
from collections import deque

COLS = 16384  # Excel column limit; key = row * COLS + col


def rss_mb():
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1e6


def col_ix(name):
    n = 0
    for ch in name.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


TOK = re.compile(r'\s*(?:(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)|(\d+(?:\.\d+)?)|([A-Za-z_]\w*)|(.))')
REF = re.compile(r'\$?([A-Za-z]+)\$?(\d+)')


def compile_formula(src):
    """Return (callable(get, rng), deps). deps: list of int keys or (c1,r1,c2,r2) tuples."""
    out, deps = [], []
    for ref, num, ident, op in TOK.findall(src[1:]):
        if ref:
            if ':' in ref:
                a, b = ref.split(':')
                (ca, ra), (cb, rb) = REF.match(a).groups(), REF.match(b).groups()
                c1, r1, c2, r2 = col_ix(ca), int(ra) - 1, col_ix(cb), int(rb) - 1
                deps.append((c1, r1, c2, r2))
                out.append(f'rng({c1},{r1},{c2},{r2})')
            else:
                c, r = REF.match(ref).groups()
                key = (int(r) - 1) * COLS + col_ix(c)
                deps.append(key)
                out.append(f'get({key})')
        elif num:
            out.append(num)
        elif ident:
            out.append(ident.upper())
        elif op:
            out.append(op)
    code = compile('lambda get, rng: ' + ''.join(out), '<formula>', 'eval')
    return eval(code, {'SUM': lambda xs: sum(x for x in xs if isinstance(x, (int, float)))}), deps


class PyEngine:
    """Everything in Python structures."""

    def __init__(self):
        self.values = {}          # key -> value
        self.formulas = {}        # key -> callable
        self.deps = {}            # key -> list of deps
        self.rev = {}             # precedent key -> list of dependent formula keys
        self.rev_ranges = {}      # col -> list of (r1, r2, dependent key)   (ranges indexed per column)

    def load(self, table):
        for r, row in enumerate(table):
            for c, v in enumerate(row):
                if v is None:
                    continue
                key = r * COLS + c
                if isinstance(v, str) and v.startswith('='):
                    fn, deps = compile_formula(v)
                    self.formulas[key] = fn
                    self.deps[key] = deps
                    for d in deps:
                        if isinstance(d, int):
                            self.rev.setdefault(d, []).append(key)
                        else:
                            c1, r1, c2, r2 = d
                            for cc in range(c1, c2 + 1):
                                self.rev_ranges.setdefault(cc, []).append((r1, r2, key))
                else:
                    self.values[key] = v

    def get(self, key):
        return self.values.get(key)

    def rng(self, c1, r1, c2, r2):
        vals = self.values
        return [vals.get(r * COLS + c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]

    def dependents_of(self, key):
        r, c = divmod(key, COLS)
        yield from self.rev.get(key, ())
        for r1, r2, dep in self.rev_ranges.get(c, ()):
            if r1 <= r <= r2:
                yield dep

    def dirty_set(self, changed_keys):
        """Transitive dependents of changed cells (BFS over the reverse index)."""
        seen, q = set(), deque(changed_keys)
        while q:
            k = q.popleft()
            for dep in self.dependents_of(k):
                if dep not in seen:
                    seen.add(dep)
                    q.append(dep)
        return seen

    def evaluate(self, targets, restrict=None):
        """Iterative post-order DFS: evaluate `targets` and (transitively) any uncomputed precedents.
        `restrict` is the set of formula keys considered stale (None = all formulas)."""
        formulas, deps, values = self.formulas, self.deps, self.values
        pending = set(restrict) if restrict is not None else set(formulas)
        get, rng = self.get, self.rng
        stack = []
        for t in targets:
            if t in pending:
                stack.append((t, False))
            while stack:
                key, expanded = stack.pop()
                if key not in pending:
                    continue
                if expanded:
                    values[key] = formulas[key](get, rng)
                    pending.discard(key)
                    continue
                stack.append((key, True))
                for d in deps[key]:
                    if isinstance(d, int):
                        if d in pending:
                            stack.append((d, False))
                    else:
                        c1, r1, c2, r2 = d
                        for rr in range(r1, r2 + 1):
                            base = rr * COLS
                            for cc in range(c1, c2 + 1):
                                k2 = base + cc
                                if k2 in pending:
                                    stack.append((k2, False))

    def recalc_all(self):
        self.evaluate(list(self.formulas))

    def set_value(self, key, v):
        self.values[key] = v
        dirty = self.dirty_set([key])
        self.evaluate(list(dirty), restrict=dirty)
        return len(dirty)


class SqliteIndexEngine(PyEngine):
    """Same evaluator, but the dependency index lives in SQLite (R-tree) and dirty propagation is SQL."""

    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect(':memory:')
        self.con.executescript('''
            create virtual table deps_rt using rtree_i32(id, minc, maxc, minr, maxr, +dep integer);
            create table frontier(key integer primary key);
            create table dirty(key integer primary key);
        ''')

    def load(self, table):
        rows = []
        for r, row in enumerate(table):
            for c, v in enumerate(row):
                if v is None:
                    continue
                key = r * COLS + c
                if isinstance(v, str) and v.startswith('='):
                    fn, deps = compile_formula(v)
                    self.formulas[key] = fn
                    self.deps[key] = deps
                    for d in deps:
                        if isinstance(d, int):
                            rr, cc = divmod(d, COLS)
                            rows.append((cc, cc, rr, rr, key))
                        else:
                            c1, r1, c2, r2 = d
                            rows.append((c1, c2, r1, r2, key))
                else:
                    self.values[key] = v
        self.con.executemany('insert into deps_rt(minc,maxc,minr,maxr,dep) values (?,?,?,?,?)', rows)
        self.con.commit()

    def dirty_set(self, changed_keys):
        con = self.con
        con.execute('delete from frontier'); con.execute('delete from dirty')
        con.executemany('insert into frontier values (?)', [(k,) for k in changed_keys])
        while True:
            # dependents of the frontier that are not yet dirty -> new frontier
            con.execute('''
                create temp table nxt as
                select distinct d.dep as key
                from frontier f join deps_rt d
                  on d.minc <= (f.key % 16384) and d.maxc >= (f.key % 16384)
                 and d.minr <= (f.key / 16384) and d.maxr >= (f.key / 16384)
                where d.dep not in (select key from dirty)''')
            n = con.execute('insert into dirty select key from nxt').rowcount
            con.execute('delete from frontier')
            con.execute('insert into frontier select key from nxt')
            con.execute('drop table nxt')
            if not n:
                break
        return {k for (k,) in con.execute('select key from dirty')}


class SqliteValuesEngine(PyEngine):
    """Python dependency index, but values are stored in SQLite and read per access (no JSON)."""

    def __init__(self):
        super().__init__()
        self.con = sqlite3.connect(':memory:')
        self.con.execute('create table cells(key integer primary key, v)')
        self._get = self.con.execute

    def load(self, table):
        super().load(table)
        self.con.executemany('insert into cells values (?,?)', list(self.values.items()))
        self.con.commit()
        self.values = _SqlDict(self.con)   # make writes go to sqlite too

    def get(self, key):
        row = self.con.execute('select v from cells where key=?', (key,)).fetchone()
        return row[0] if row else None

    def rng(self, c1, r1, c2, r2):
        return [v for (v,) in self.con.execute(
            'select v from cells where key between ? and ? and key % 16384 between ? and ? order by key',
            (r1 * COLS + c1, r2 * COLS + c2, c1, c2))]


class _SqlDict:
    def __init__(self, con):
        self.con = con

    def __setitem__(self, key, v):
        self.con.execute('insert or replace into cells values (?,?)', (key, v))

    def get(self, key, default=None):
        row = self.con.execute('select v from cells where key=?', (key,)).fetchone()
        return row[0] if row else default

    def __getitem__(self, key):
        return self.get(key)


def main():
    mode, N = sys.argv[1], int(sys.argv[2])
    with open(f'data/bench_{N}.pkl', 'rb') as f:
        table = pickle.load(f)
    rss0 = rss_mb()
    eng = {'py': PyEngine, 'sqlite_index': SqliteIndexEngine, 'sqlite_values': SqliteValuesEngine}[mode]()
    res = dict(pkg=f'proto_{mode}', N=N)
    t0 = time.perf_counter(); eng.load(table); res['load_s'] = time.perf_counter() - t0
    res['rss_after_load_mb'] = rss_mb() - rss0
    t0 = time.perf_counter(); eng.recalc_all(); res['eval_s'] = time.perf_counter() - t0
    F = lambda r: eng.values.get(r * COLS + 5)
    res['values'] = [F(0), F(1), F(2)]
    res['n_formulas'] = len(eng.formulas)
    # incremental 1: change A1 -> dirties the whole running-total chain + summaries
    t0 = time.perf_counter(); n = eng.set_value(0, 1000.0); res['reeval_A1_s'] = time.perf_counter() - t0
    res['dirty_A1'] = n
    res['values2'] = [F(0), F(1), F(2)]
    # incremental 2: change A in the last row -> small fan-out (B, D window, C last, summaries)
    k = (N - 1) * COLS
    t0 = time.perf_counter(); n = eng.set_value(k, 5.0); res['reeval_Alast_s'] = time.perf_counter() - t0
    res['dirty_Alast'] = n
    # dirty-set computation alone (no evaluation), averaged
    t0 = time.perf_counter()
    for _ in range(20):
        eng.dirty_set([k])
    res['dirty_only_Alast_ms'] = (time.perf_counter() - t0) / 20 * 1000
    t0 = time.perf_counter(); eng.dirty_set([0]); res['dirty_only_A1_ms'] = (time.perf_counter() - t0) * 1000
    res['peak_rss_mb'] = rss_mb()
    print(json.dumps(res), flush=True)


if __name__ == '__main__':
    main()
