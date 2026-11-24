"""Mutation fuzzer for rsa_gcv4w_comp_to_config.py.   usage: python tests\\fuzz_rsa.py N [first_seed] [workers]

Every case: corrupt the real RSA grid (October or V2) in memory, then run the whole script (a) with a random 'user' answering every
question and (b) with no keyboard. Acceptable outcomes: a correct file, or a clean, explained stop (GridError) about the GRID.
Anything else is a bug: another exception, a malformed file, wrong arithmetic, a hang, a flood of questions, or a stop / question
caused by a NOTE (notes must be flagged in the remarks column, never stop the run)."""
import copy
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

import pandas as pd

SC = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SC)
import fuzzlib  # noqa: E402
import helpers  # noqa: E402

SCRIPT = os.path.join(helpers.ROOT, "rsa_gcv4w_comp_to_config.py")
quiet = lambda *a, **k: None
G = {}


def init():
    spec = importlib.util.spec_from_file_location("rsa", SCRIPT)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    G["mod"] = m
    G["bases"] = pickle.load(open(os.path.join(SC, "fuzz_rsa_bases.pkl"), "rb"))


# ---------------------------------------------------------------- RSA-specific mutations
def _sheet(S, rng):
    return rng.choice(list(S)) if S else None


def _note_rows(rows):
    return [i for i, r in enumerate(rows) if r and isinstance(r[0], str) and re.match(r"\s*\d+\.", r[0])]


def r_notes(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    idx = _note_rows(S[n])
    if not idx:
        return
    i = rng.choice(idx)
    how = rng.randint(0, 5)
    if how == 0:
        del S[n][i]
    elif how == 1:
        S[n][i][0] = re.sub(r"\d+(?:\.\d+)?", lambda m: str(rng.choice([0, 1, 2, 5, 10, 50, 99, 150, -3, 2.5])), S[n][i][0])
    elif how == 2:
        S[n][i][0] = rng.choice(["see circular", "", "3. For Eicher Make - payout will be x% lower", "Eicher 3% lower", None, 7,
                                 "3. For Eicher make - deduction of three", "3. For Eicher make - payout will be reduced by 4%",
                                 "10. This is effective from 22nd Oct 2026", "10. The existing 1st Oct 2026 grid continues",
                                 "10. w.e.f. 5 Nov 2026", "For GVW above 12 ton the payout will be same for old & new."])
    elif how == 3:
        S[n].insert(i, list(S[n][i]))
    elif how == 4:
        S[n][i][0] = S[n][i][0].replace("lower", "higher")
    else:
        S[n].append([rng.choice(["12. Extra 3% for Tata in Gujarat", "Note: everything +1%", "13. Discount for RTOs of Kerala is 70%"])] + [None] * (len(S[n][0]) - 1))


def r_age(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    for r in S[n][3:12]:
        if len(r) > 1 and r[1] in ("Old",) and rng.random() < 0.7:
            r[1] = rng.choice(["New", "Used", "", "OLD", "old ", "N", "Fresh", None, 1, "Old"])


def r_new_rows(S, rng):
    n = _sheet(S, rng)
    if n is None or len(S[n]) < 13:
        return
    out = []
    for r in S[n][3:12]:
        out.append(list(r))
        if rng.random() < 0.6:
            new = list(r)
            new[1] = "New"
            new[2:] = [(v * rng.choice([0.5, 0.9, 1.0, 1.2]) if isinstance(v, (int, float)) else v) for v in new[2:]]
            out.append(new)
    S[n][3:12] = out


def r_band(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    for r in S[n][3:12]:
        if r and rng.random() < 0.3:
            r[0] = rng.choice(["Heavy", "12-20", "x", "12 to 20 Tons", "3.5 - 7.5 T", None, "", "48 Ton and above", "Above 48", "Upto 2.3 Ton",
                               "A. 5 to 2 Ton", "7500 to 12000 kg", "0 to 0", 12, "I. 48+"])


def r_header(S, rng):
    n = _sheet(S, rng)
    if n is None or len(S[n]) < 3:
        return
    h = S[n][2]
    h[0] = rng.choice(["Tonnage Band", "Weight Band", "GVW", "Band", "tonnage  band", "", None, "Tonnage"])
    if len(h) > 1 and rng.random() < 0.5:
        h[1] = rng.choice(["Age", "age", "Vehicle Age", "", None, "AGE "])


def r_title(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    if rng.random() < 0.5:
        if S[n] and S[n][0]:
            S[n][0][0] = rng.choice(["GCV - 4W - SATP", "GCV - 3W - COMP", "PCV - 4W - COMP", "", None, "GCV 4W", "COMP"])
    elif len(S[n]) > 1 and S[n][1]:
        S[n][1][0] = rng.choice(["EV", "CNG", "", None, "NON EV", "NON-EV", "Electric", "NON EV "])


def r_extra_cols(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    w = len(S[n][0]) if S[n] else 0
    for r in S[n]:
        r.extend([None] * 2)
    for _ in range(rng.randint(1, 3)):
        c = w + rng.randint(0, 1)
        for r in S[n]:
            if len(r) <= c:
                r.extend([None] * (c + 1 - len(r)))
        if len(S[n]) > 2:
            S[n][2][c] = rng.choice(["Remarks", "NEW,CITY", "EXTRA CLUSTER", None, "", "REST OF  X"])
            for i in range(3, min(12, len(S[n]))):
                S[n][i][c] = rng.choice([0.2, 0.0, None, "NA", "abc", 0.35])


def r_dup_cluster(S, rng):
    n = _sheet(S, rng)
    if n is None or len(S[n]) < 3 or len(S[n][2]) < 5:
        return
    a, b = rng.sample(range(2, len(S[n][2])), 2)
    S[n][2][b] = S[n][2][a]


def r_scale_rows(S, rng):
    n = _sheet(S, rng)
    if n is None:
        return
    f = rng.choice([100, 0.01, 10, -1])
    for r in S[n][3:12]:
        for c in range(2, len(r)):
            if isinstance(r[c], (int, float)) and not isinstance(r[c], bool) and rng.random() < 0.5:
                r[c] = r[c] * f


def r_zero_row(S, rng):
    n = _sheet(S, rng)
    if n is None or len(S[n]) < 12:
        return
    r = S[n][rng.randint(3, 11)]
    for c in range(2, len(r)):
        r[c] = rng.choice([0, None, "NA", 0.0])


RSA_MUTS = [r_notes, r_age, r_new_rows, r_band, r_header, r_title, r_extra_cols, r_dup_cluster, r_scale_rows, r_zero_row]
POOL = RSA_MUTS * 3 + fuzzlib.MUTATIONS


def mutate(base, rng, how_many):
    S = copy.deepcopy(base)
    chosen = [rng.choice(POOL) for _ in range(how_many)]
    for m in chosen:
        m(S, rng)
    return S, [m.__name__ for m in chosen]


class User(fuzzlib.RandomUser):
    def inp(self, prompt=""):
        if prompt.strip() == ">":
            q = self.question.lower()
            if "weight limits" in q:
                self.prompts += 1
                return self.rng.choice(["7.5-12", "0-2", "48+", "abc", "5", "12-7"])
            if q.startswith("type the"):
                self.prompts += 1
                return self.rng.choice(["2", "86", "abc", "150", "0", "-1", "10"])
        return super().inp(prompt)


def check_values(df, allow_dups=False):
    """arithmetic that must hold for any grid"""
    problems = []
    pi = pd.to_numeric(df["pi_net_rate"], errors="coerce")
    po = pd.to_numeric(df["po_net_rate"], errors="coerce")
    exp = pi.where(pi <= 0, (pi - 7).clip(lower=0)).where(pi > 0, 0.0)
    if ((po - exp).abs() > 1e-6).any():
        problems.append("po_net_rate is not pi_net_rate - 7 (floor 0)")
    if (pi < 0).any():
        problems.append("negative pi_net_rate")
    if df.duplicated(G["mod"].KEY_COLUMNS).any() and not allow_dups:
        problems.append("duplicate keys")
    if not set(zip(df.vehicle_age_min, df.vehicle_age_max)) <= {("1", "99"), ("0", "1")}:
        problems.append("odd age bands")
    if not (df.id == "1").all():
        problems.append("id not 1")
    if "remarks" in df and (((df.remarks != "") & ~df.remarks.str.startswith("Manual Review Required: [")).any()):
        problems.append("a remark that is not a manual-review flag")
    return problems


def outcome(fn):
    mod = G["mod"]
    try:
        return "ok", fn()
    except mod.GridError as e:
        msg = str(e).replace("\n", " ")
        if re.search(r"notes hold text|could not turn|notes do not say|New CV note|effective date|up to \d+ ton", msg):
            return "BAD: a stop caused by a NOTE (notes must be flagged, never stop the run): " + msg[:140], None
        return "stopped: " + msg[:140], None
    except fuzzlib.Overload as e:
        return "OVERLOAD: " + str(e), None
    except BaseException as e:
        return f"CRASH {type(e).__name__}: {str(e)[:160]}\n{traceback.format_exc()[-1800:]}", None


def run_case(seed):
    mod = G["mod"]
    rng = random.Random(seed)
    t0 = time.time()
    tmp = tempfile.mkdtemp(prefix="fzr_")
    which = rng.choice(["rsa_oct", "rsa_v2"])
    F, T = (dt.date(2026, 10, 22) if which == "rsa_v2" else dt.date(2026, 10, 1)), dt.date(2030, 3, 31)
    sheets, muts = mutate(G["bases"][which], rng, rng.choice([0, 1, 1, 2, 3, 4]))
    wb = helpers.FakeWB(sheets)
    mod.read_workbook = lambda p: wb
    call = lambda ask, out: mod.main("fake.xlsx", F, T, out, None, quiet, ask)
    res = {"seed": seed, "muts": muts, "base": which}
    user = User(random.Random(seed + 7))
    ask = mod.Asker(True, user.inp, user.out)
    kind_i, path = outcome(lambda: call(ask, os.path.join(tmp, "o.csv")))

    def verify(kind, path, allow_dups=False):
        if kind == "ok":
            d = pd.read_csv(path, dtype=str, keep_default_na=False)
            problems = helpers.check_csv(path, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS) + check_values(d, allow_dups)
            if problems:
                return "BAD OUTPUT: " + "; ".join(problems)
        return kind

    kind_i = verify(kind_i, path, any("Write all of them anyway" in d for d in ask.decisions))      # the person chose to keep duplicates
    retyped = any("Please type" in line for line in user.transcript)
    if not muts and kind_i == "ok" and user.prompts > 1 and not retyped:
        kind_i = f"BAD OUTPUT: a clean grid triggered {user.prompts - 1} unneeded question(s)"
    res["interactive"], res["prompts"] = kind_i, user.prompts
    if kind_i.startswith(("CRASH", "BAD", "OVERLOAD")):
        res["transcript"] = "\n".join(user.transcript[-30:])
    kind_n, path = outcome(lambda: call(mod.Asker(False), os.path.join(tmp, "n.csv")))
    kind_n = verify(kind_n, path)
    if not muts and kind_n != "ok":
        kind_n = "BAD OUTPUT (silent): a clean grid did not run: " + kind_n[:100]
    res["silent"] = kind_n
    res["secs"] = round(time.time() - t0, 1)
    return res


def bucket(msg):
    if not msg.startswith("stopped"):
        return msg.split("\n")[0][:150]
    m = msg[9:]
    for key in ("nothing was written", "stopped by you", "nothing to write", "no keyboard input", "no GCV 4W comprehensive sheet", "weight band",
                "do not join", "unexpected value", "holds", "duplicate rows", "age label", "cannot read", "Tonnage Band", "can't find",
                "no Age", "Electric", "no cluster", "no weight-band", "has no", "title", "NON EV", "labelled", "repeat a key"):
        if key in m:
            return "stopped: " + key
    return "stopped: " + m[:70]


if __name__ == "__main__":
    n = int(sys.argv[1])
    first = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    workers = int(sys.argv[3]) if len(sys.argv) > 3 else min(6, os.cpu_count() or 2)
    ctx = mp.get_context("spawn")
    t0 = time.time()
    results = []
    with ctx.Pool(workers, initializer=init) as pool:
        for r in pool.imap_unordered(run_case, range(first, first + n)):
            results.append(r)
            if len(results) % 20 == 0:
                print(f"  ... {len(results)}/{n} cases, {time.time() - t0:.0f}s", flush=True)
    json.dump(results, open(os.path.join(SC, "fuzz_rsa_results.json"), "w"), indent=1, default=str)
    for label in ("interactive", "silent"):
        c = Counter(bucket(r[label]) for r in results)
        print(f"\n== rsa: {label} run, {n} cases ==")
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
