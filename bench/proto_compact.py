"""Prototype of the 'compact' design: formulas deduplicated into relative (R1C1-style) templates,
compiled once per template; per-cell state is just (template_id) + value.  No per-cell dependency
lists at all: dependents are derived from templates + geometry.

Usage: proto_compact.py <N>
Same restricted grammar as proto.py.
"""
import sys, re, time, json, pickle, resource
from pathlib import Path
from rss import current_mb, peak_mb as rss_mb
from collections import deque, defaultdict

COLS = 16384



def col_ix(name):
    n = 0
    for ch in name.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


TOK = re.compile(r'\s*(?:(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)|(\d+(?:\.\d+)?)|([A-Za-z_]\w*)|(.))')
REF = re.compile(r'(\$?)([A-Za-z]+)(\$?)(\d+)')


def rel_ref(text, r, c):
    """'B3' at cell (r,c) -> (dr, dc, abs_r, abs_c); $ anchors keep absolute coordinates."""
    ar, cn, arr, rn = REF.match(text).groups()
    rr, cc = int(rn) - 1, col_ix(cn)
    return (rr if arr else rr - r, cc if ar else cc - c, bool(arr), bool(ar))


def to_template(src, r, c):
    """Rewrite a formula into a position-independent template string + list of relative refs."""
    out, refs = [], []
    for ref, num, ident, op in TOK.findall(src[1:]):
        if ref:
            if ':' in ref:
                a, b = ref.split(':')
                ra, rb = rel_ref(a, r, c), rel_ref(b, r, c)
                refs.append(('R', ra, rb))
                out.append(f'rng({_rc(ra)},{_rc(rb)})')
            else:
                ra = rel_ref(ref, r, c)
                refs.append(('C', ra))
                out.append(f'get({_key(ra)})')
        elif num:
            out.append(num)
        elif ident:
            out.append(ident.upper())
        elif op:
            out.append(op)
    return ''.join(out), refs


def _rc(ra):
    dr, dc, absr, absc = ra
    return f"{dr if absr else f'r+{dr}'},{dc if absc else f'c+{dc}'}"


def _key(ra):
    dr, dc, absr, absc = ra
    return f"({dr if absr else f'r+{dr}'})*{COLS}+({dc if absc else f'c+{dc}'})"


class CompactEngine:
    def __init__(self):
        self.values = {}            # key -> value
        self.tmpl_of = {}           # formula cell key -> template id
        self.templates = []         # id -> (callable(get, rng, r, c), refs)
        self.tmpl_ids = {}          # template text -> id
        self.cells_of_tmpl = defaultdict(list)   # id -> [keys]   (only consulted for absolute/large refs)

    def load(self, table):
        for r, row in enumerate(table):
            for c, v in enumerate(row):
                if v is None:
                    continue
                key = r * COLS + c
                if isinstance(v, str) and v.startswith('='):
                    text, refs = to_template(v, r, c)
                    tid = self.tmpl_ids.get(text)
                    if tid is None:
                        code = compile('lambda get, rng, r, c: ' + text, '<t>', 'eval')
                        fn = eval(code, {'SUM': lambda xs: sum(x for x in xs if isinstance(x, (int, float)))})
                        tid = len(self.templates)
                        self.templates.append((fn, refs))
                        self.tmpl_ids[text] = tid
                    self.tmpl_of[key] = tid
                    self.cells_of_tmpl[tid].append(key)
                else:
                    self.values[key] = v

    def get(self, key):
        return self.values.get(key)

    def rng(self, r1, c1, r2, c2):
        vals = self.values
        return [vals.get(r * COLS + c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]

    # --- dependency geometry -------------------------------------------------------------
    def precedents(self, key):
        """Yield precedent keys / (r1,c1,r2,c2) boxes of a formula cell from its template."""
        r, c = divmod(key, COLS)
        for ref in self.templates[self.tmpl_of[key]][1]:
            if ref[0] == 'C':
                dr, dc, absr, absc = ref[1]
                yield ((dr if absr else r + dr) * COLS + (dc if absc else c + dc))
            else:
                (dr1, dc1, a1, b1), (dr2, dc2, a2, b2) = ref[1], ref[2]
                yield ((dr1 if a1 else r + dr1), (dc1 if b1 else c + dc1),
                       (dr2 if a2 else r + dr2), (dc2 if b2 else c + dc2))

    def dependents_of(self, key):
        """Dependents of a cell, derived from templates: for each template ref, the only cells that could
        reference (r,c) through it are at fixed offsets -> probe those keys.  Large/absolute ranges fall
        back to scanning the template's cell list."""
        r, c = divmod(key, COLS)
        tmpl_of = self.tmpl_of
        for tid, (fn, refs) in enumerate(self.templates):
            for ref in refs:
                if ref[0] == 'C':
                    dr, dc, absr, absc = ref[1]
                    if absr or absc:
                        if (absr and dr != r) or (absc and dc != c):
                            continue
                        if absr and absc:
                            yield from self.cells_of_tmpl[tid]
                            continue
                        # one axis absolute: scan template cells on the other axis
                        for k in self.cells_of_tmpl[tid]:
                            rr, cc = divmod(k, COLS)
                            if (absr and cc + dc == c) or (absc and rr + dr == r):
                                yield k
                        continue
                    k = (r - dr) * COLS + (c - dc)
                    if tmpl_of.get(k) == tid:
                        yield k
                else:
                    (dr1, dc1, a1, b1), (dr2, dc2, a2, b2) = ref[1], ref[2]
                    if a1 or a2 or b1 or b2 or (dr2 - dr1 + 1) * (dc2 - dc1 + 1) > 64:
                        for k in self.cells_of_tmpl[tid]:
                            rr, cc = divmod(k, COLS)
                            r1 = dr1 if a1 else rr + dr1; r2 = dr2 if a2 else rr + dr2
                            c1 = dc1 if b1 else cc + dc1; c2 = dc2 if b2 else cc + dc2
                            if r1 <= r <= r2 and c1 <= c <= c2:
                                yield k
                    else:
                        for rr in range(r - dr2, r - dr1 + 1):
                            for cc in range(c - dc2, c - dc1 + 1):
                                k = rr * COLS + cc
                                if tmpl_of.get(k) == tid:
                                    yield k

    def dirty_set(self, changed_keys):
        seen, q = set(), deque(changed_keys)
        while q:
            for dep in self.dependents_of(q.popleft()):
                if dep not in seen:
                    seen.add(dep)
                    q.append(dep)
        return seen

    # --- evaluation --------------------------------------------------------------------------
    def evaluate(self, targets, restrict=None):
        values, tmpl_of, templates = self.values, self.tmpl_of, self.templates
        pending = set(restrict) if restrict is not None else set(tmpl_of)
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
                    r, c = divmod(key, COLS)
                    values[key] = templates[tmpl_of[key]][0](get, rng, r, c)
                    pending.discard(key)
                    continue
                stack.append((key, True))
                for d in self.precedents(key):
                    if isinstance(d, int):
                        if d in pending:
                            stack.append((d, False))
                    else:
                        r1, c1, r2, c2 = d
                        for rr in range(r1, r2 + 1):
                            base = rr * COLS
                            for cc in range(c1, c2 + 1):
                                if base + cc in pending:
                                    stack.append((base + cc, False))

    def recalc_all(self):
        self.evaluate(list(self.tmpl_of))

    def set_value(self, key, v):
        self.values[key] = v
        dirty = self.dirty_set([key])
        self.evaluate(list(dirty), restrict=dirty)
        return len(dirty)


def main():
    N = int(sys.argv[1])
    res = dict(pkg='proto_compact', N=N, python=sys.version.split()[0])
    res['rss_start_mb'] = current_mb()
    with open(Path(__file__).resolve().parent / f'data/bench_{N}.pkl', 'rb') as f:
        table = pickle.load(f)
    rss0 = rss_mb()
    res['rss_table_mb'] = current_mb()
    eng = CompactEngine()
    t0 = time.perf_counter(); eng.load(table); res['load_s'] = time.perf_counter() - t0
    res['rss_after_load_mb'] = current_mb()
    res['peak_after_load_mb'] = rss_mb()
    res['peak_delta_after_load_mb'] = rss_mb() - rss0  # what earlier runs called rss_after_load_mb
    del table
    res['rss_after_load_no_table_mb'] = current_mb()
    res['n_templates'] = len(eng.templates)
    t0 = time.perf_counter(); eng.recalc_all(); res['eval_s'] = time.perf_counter() - t0
    F = lambda r: eng.values.get(r * COLS + 5)
    res['values'] = [F(0), F(1), F(2)]
    res['n_formulas'] = len(eng.tmpl_of)
    t0 = time.perf_counter(); n = eng.set_value(0, 1000.0); res['reeval_A1_s'] = time.perf_counter() - t0
    res['dirty_A1'] = n
    res['values2'] = [F(0), F(1), F(2)]
    k = (N - 1) * COLS
    t0 = time.perf_counter(); n = eng.set_value(k, 5.0); res['reeval_Alast_s'] = time.perf_counter() - t0
    res['dirty_Alast'] = n
    t0 = time.perf_counter()
    for _ in range(20):
        eng.dirty_set([k])
    res['dirty_only_Alast_ms'] = (time.perf_counter() - t0) / 20 * 1000
    t0 = time.perf_counter(); eng.dirty_set([0]); res['dirty_only_A1_ms'] = (time.perf_counter() - t0) * 1000
    res['rss_end_mb'] = current_mb()
    res['peak_rss_mb'] = rss_mb()
    print(json.dumps(res), flush=True)


if __name__ == '__main__':
    main()
