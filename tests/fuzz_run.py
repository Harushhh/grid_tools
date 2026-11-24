"""Mutation fuzzer for both scripts.   usage: python fuzz_run.py lib|mag N [first_seed] [workers]

Every case: corrupt the real grids in memory, then run the whole script (a) with a random 'user' answering every
question and (b) with no keyboard. The only acceptable outcomes are a correct file or a clean, explained stop
(GridError). Any other exception, a malformed file, a hang or a flood of questions is a bug."""
import os as _os
ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))      # the grid_tools folder
import datetime as dt
import importlib.util
import json
import multiprocessing as mp
import os
import pickle
import random
import re
import sys
import tempfile
import time
import traceback
from collections import Counter

SC = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SC)
import fuzzlib  # noqa: E402
import helpers  # noqa: E402

GA = ROOT
quiet = lambda *a, **k: None
G = {}


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def init(kind):
    G["kind"] = kind
    if kind == "lib":
        G["mod"] = load("lg", GA + r"\liberty_grid_to_config.py")
        G["bases"] = pickle.load(open(os.path.join(SC, "fuzz_lib_bases.pkl"), "rb"))
    else:
        G["mod"] = load("mg", GA + r"\magma_grid_to_config.py")
        G["bases"] = pickle.load(open(os.path.join(SC, "fuzz_mag_bases.pkl"), "rb"))


class FakeSheet:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, values_only=True):
        return iter([tuple(r) for r in self.rows])


class FakeWB:
    def __init__(self, sheets):
        self.sheets = sheets
        self.sheetnames = list(sheets)

    def __getitem__(self, name):
        return FakeSheet(self.sheets[name])


def outcome(fn):
    try:
        path = fn()
        return "ok", path
    except G["mod"].GridError as e:
        return "stopped: " + str(e)[:140].replace("\n", " "), None
    except fuzzlib.Overload as e:
        return "OVERLOAD: " + str(e), None
    except BaseException as e:
        return f"CRASH {type(e).__name__}: {str(e)[:160]}\n{traceback.format_exc()[-1800:]}", None


def run_case(seed):
    mod, kind = G["mod"], G["kind"]
    rng = random.Random(seed)
    t0 = time.time()
    tmp = tempfile.mkdtemp(prefix="fz_")
    F, T = dt.date(2026, 9, 1), dt.date(2030, 3, 31)
    res = {"seed": seed}
    if kind == "lib":
        base = rng.choice(["lib_sep", "lib_sep", "lib_jul_b", "lib_jul_a"])
        sheets, muts = fuzzlib.mutate(G["bases"][base], rng, rng.choice([0, 1, 1, 2, 3, 5]))
        wb = FakeWB(sheets)
        mod.read_workbook = lambda p: wb
        call = lambda ask, out: mod.main("fake.xlsx", F, T, out, None, quiet, ask)
        res["base"] = base
    else:
        motor, m1 = fuzzlib.mutate(G["bases"]["motor"], rng, rng.choice([0, 1, 1, 2, 3]))
        satp, m2 = fuzzlib.mutate(G["bases"]["satp"], rng, rng.choice([0, 1, 1, 2, 3]))
        muts = ["motor:" + m for m in m1] + ["satp:" + m for m in m2]
        data = {"fake_grid.xlsx": motor, "fake_satp.xlsb": satp}
        mod.list_sheets = lambda p: list(data[p])
        mod.read_rows = lambda p, s: data[p][s]
        call = lambda ask, out: mod.main("fake_grid.xlsx", "fake_satp.xlsb", dt.date(2026, 10, 1), T, out, None, quiet, ask)
    res["muts"] = muts

    user = fuzzlib.RandomUser(random.Random(seed + 7))
    ask = mod.Asker(True, user.inp, user.out)
    out = os.path.join(tmp, "o.csv")
    kind_i, path = outcome(lambda: call(ask, out))
    if kind_i == "ok":
        problems = helpers.check_csv(path, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
        if problems:
            kind_i = "BAD OUTPUT: " + "; ".join(problems)
    clean = not muts and res.get("base", "lib_sep") in ("lib_sep", "lib_sep")
    if clean and kind_i == "ok" and user.prompts > 1:
        kind_i = f"BAD OUTPUT: a clean grid triggered {user.prompts - 1} unneeded question(s)"
    clean = not muts and res.get("base", "lib_sep") in ("lib_sep", "lib_sep")
    if clean and kind_i == "ok" and user.prompts > 1:
        kind_i = f"BAD OUTPUT: a clean grid triggered {user.prompts - 1} unneeded question(s)"
    res["interactive"], res["prompts"] = kind_i, user.prompts
    if kind_i.startswith(("CRASH", "BAD", "OVERLOAD")):
        res["transcript"] = "\n".join(user.transcript[-30:])

    kind_n, path = outcome(lambda: call(mod.Asker(False), os.path.join(tmp, "n.csv")))
    if kind_n == "ok":
        problems = helpers.check_csv(path, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
        if problems:
            kind_n = "BAD OUTPUT: " + "; ".join(problems)
    if clean and kind_n != "ok":
        kind_n = "BAD OUTPUT (silent): a clean grid did not run: " + kind_n[:100]
    if clean and kind_n != "ok":
        kind_n = "BAD OUTPUT (silent): a clean grid did not run: " + kind_n[:100]
    res["silent"] = kind_n
    res["secs"] = round(time.time() - t0, 1)
    return res


def bucket(msg):
    """group 'stopped: ...' messages so the summary stays readable"""
    if not msg.startswith("stopped"):
        return msg.split("\n")[0][:150]
    m = msg[9:]
    if re.search(r"geos have no '", m):
        return "stopped: some geos lack the slab (silent run)"
    for key in ("nothing was written", "does not look like", "stopped by you", "nothing to write", "no keyboard input", "not found", "header row", "appears twice", "expected two tables",
                "unexpected value", "duplicate rows", "slab", "is not a file", "cannot open", "does not look like", "empty", "no geo rows"):
        if key in m:
            return "stopped: " + key
    return "stopped: " + m[:60]


if __name__ == "__main__":
    kind, n = sys.argv[1], int(sys.argv[2])
    first = int(sys.argv[3]) if len(sys.argv) > 3 else 1
    workers = int(sys.argv[4]) if len(sys.argv) > 4 else min(6, os.cpu_count() or 2)
    ctx = mp.get_context("spawn")
    t0 = time.time()
    results = []
    with ctx.Pool(workers, initializer=init, initargs=(kind,)) as pool:
        for r in pool.imap_unordered(run_case, range(first, first + n)):
            results.append(r)
            if len(results) % 20 == 0:
                print(f"  ... {len(results)}/{n} cases, {time.time() - t0:.0f}s", flush=True)
    json.dump(results, open(os.path.join(SC, f"fuzz_{kind}_results.json"), "w"), indent=1, default=str)
    for label in ("interactive", "silent"):
        c = Counter(bucket(r[label]) for r in results)
        print(f"\n== {kind}: {label} run, {n} cases ==")
        for k, v in c.most_common():
            print(f"  {v:4d}  {k}")
    bad = [r for r in results if any(r[x].startswith(("CRASH", "BAD", "OVERLOAD")) for x in ("interactive", "silent"))]
    print(f"\nFAILURES: {len(bad)} of {n}  (total {time.time() - t0:.0f}s, avg {sum(r['secs'] for r in results) / n:.1f}s/case, max prompts {max(r['prompts'] for r in results)})")
    seen = set()
    for r in sorted(bad, key=lambda r: r["seed"]):
        which = "interactive" if r["interactive"].startswith(("CRASH", "BAD", "OVERLOAD")) else "silent"
        sig = r[which].split("\n")[0][:110]
        if sig in seen:
            continue
        seen.add(sig)
        print(f"\n--- seed {r['seed']} ({which}) base={r.get('base')} mutations={r['muts']}\n{r[which]}")
        if which == "interactive" and r.get("transcript"):
            print("last output before it:\n" + r["transcript"][-1200:])
