"""Usage: run_competitor.py <pycel|formulas|xlcalculator> <N>
Prints one JSON line: load time, full-eval time, incremental re-eval time, peak RSS (MB)."""
import sys, time, json, resource, traceback
sys.setrecursionlimit(1_000_000)
import threading
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from rss import current_mb, peak_mb as rss_mb
threading.stack_size(512 * 1024 * 1024)

def run(pkg, N):
    path = str(Path(__file__).resolve().parent / f'data/bench_{N}.xlsx')
    targets = ['Sheet1!F1', 'Sheet1!F2', 'Sheet1!F3']
    res = dict(pkg=pkg, N=N, python=sys.version.split()[0])
    res['rss_start_mb'] = current_mb()
    t0 = time.perf_counter()
    if pkg == 'pycel':
        from pycel import ExcelCompiler
        xl = ExcelCompiler(filename=path)
        res['load_s'] = time.perf_counter() - t0
        res['rss_after_load_mb'], res['peak_after_load_mb'] = current_mb(), rss_mb()
        t0 = time.perf_counter()
        res['values'] = [xl.evaluate(t) for t in targets]
        res['eval_s'] = time.perf_counter() - t0
        t0 = time.perf_counter()
        xl.set_value('Sheet1!A1', 1000.0)
        res['values2'] = [xl.evaluate(t) for t in targets]
        res['reeval_s'] = time.perf_counter() - t0
    elif pkg == 'formulas':
        import formulas
        xl = formulas.ExcelModel().loads(path).finish()
        res['load_s'] = time.perf_counter() - t0
        res['rss_after_load_mb'], res['peak_after_load_mb'] = current_mb(), rss_mb()
        t0 = time.perf_counter()
        sol = xl.calculate(outputs=[f"'[bench_{N}.xlsx]SHEET1'!{t.split('!')[1]}" for t in targets])
        res['values'] = [float(sol[k].value[0][0]) for k in sol if any(k.endswith('!' + t.split('!')[1]) for t in targets)]
        res['eval_s'] = time.perf_counter() - t0
        t0 = time.perf_counter()
        sol = xl.calculate(inputs={f"'[bench_{N}.xlsx]SHEET1'!A1": 1000.0},
                           outputs=[f"'[bench_{N}.xlsx]SHEET1'!{t.split('!')[1]}" for t in targets])
        res['values2'] = [float(sol[k].value[0][0]) for k in sol if any(k.endswith('!' + t.split('!')[1]) for t in targets)]
        res['reeval_s'] = time.perf_counter() - t0
    elif pkg == 'xlcalculator':
        from xlcalculator import ModelCompiler, Evaluator
        model = ModelCompiler().read_and_parse_archive(path)
        ev = Evaluator(model)
        res['load_s'] = time.perf_counter() - t0
        res['rss_after_load_mb'], res['peak_after_load_mb'] = current_mb(), rss_mb()
        t0 = time.perf_counter()
        res['values'] = [float(ev.evaluate(t)) for t in targets]
        res['eval_s'] = time.perf_counter() - t0
        t0 = time.perf_counter()
        ev.set_cell_value('Sheet1!A1', 1000.0)
        res['values2'] = [float(ev.evaluate(t)) for t in targets]
        res['reeval_s'] = time.perf_counter() - t0
    res['rss_end_mb'] = current_mb()
    res['peak_rss_mb'] = rss_mb()
    return res

def main():
    pkg, N = sys.argv[1], int(sys.argv[2])
    out = {}
    def target():
        try:
            out.update(run(pkg, N))
        except Exception as ex:
            out.update(pkg=pkg, N=N, error=f'{type(ex).__name__}: {str(ex)[:200]}', peak_rss_mb=rss_mb())
    t = threading.Thread(target=target); t.start(); t.join()
    print(json.dumps(out), flush=True)

if __name__ == '__main__':
    main()
