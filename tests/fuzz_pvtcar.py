"""Mutation fuzzer for rsa_pvtcar_to_config.py.   usage: python tests\\fuzz_pvtcar.py N [first_seed] [workers]

Every case: corrupt the real October Private Car grid in memory, then run the whole script (a) with a random 'user' answering every
question and (b) with no keyboard. Acceptable outcomes: a correct file, or a clean, explained stop (GridError) about the GRID.
Anything else is a bug: another exception, a malformed file, wrong arithmetic, a hang, a flood of questions, or a stop / question
caused by a NOTE (notes must be flagged in the remarks column, never stop the run)."""
import copy
import datetime as dt
import importlib.util
import json
import multiprocessing as mp
import os
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

SCRIPT = os.path.join(helpers.ROOT, "rsa_pvtcar_to_config.py")
quiet = lambda *a, **k: None
G = {}
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)


def init():
    spec = importlib.util.spec_from_file_location("pvtcar", SCRIPT)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    G["mod"] = m
    G["bases"] = {"oct": helpers.pvtcar_base("oct"), "sep": helpers.pvtcar_base("sep")}


# ---------------------------------------------------------------- Private-Car-specific mutations
def _comp(S, rng):
    names = [n for n in S if n.startswith("Pvt Car- 4W")]
    return rng.choice(names) if names else None


def _table_rows(rows):
    h = next((i for i, r in enumerate(rows) if "Business Type" in r), None)
    return [] if h is None else list(range(h + 1, min(len(rows), h + 16)))


def p_notes(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    idx = [i for i, r in enumerate(S[n]) if len(r) > 1 and isinstance(r[1], str) and re.match(r"\s*\d+\.", r[1])]
    if not idx:
        return
    i = rng.choice(idx)
    how = rng.randint(0, 6)
    if how == 0:
        del S[n][i]
    elif how == 1:
        S[n][i][1] = re.sub(r"\d+(?:\.\d+)?", lambda m: str(rng.choice([0, 1, 2, 3, 5, 28, 99, -2, 2.5])), S[n][i][1])
    elif how == 2:
        S[n][i][1] = rng.choice(["see circular", "", None, 7, "5. Additional two percent payout for the Top Variant Models of Maruti and Tata",
                                 "3. Any other make New car payout will be 30% flat", "3. Any other make New car payout will be reduced",
                                 "5. Extra 4 points for top end variants", "2. Above New Car payout is applicable only for the following makes - Maruti",
                                 "6. Payout is 3% lower for Gujarat", "7. w.e.f. 5 Nov 2026 all payouts change"])
    elif how == 3:
        S[n].insert(i, list(S[n][i]))
    elif how == 4:
        S[n][i][1] = S[n][i][1].replace("Additional", "Reduced").replace("28%", "x%")
    elif how == 5:
        S[n].append([None, rng.choice(["8. Fleet business: +1% on all", "Note: Karnataka -2%", "9. Mumbai Pune payout 1% lower"])] + [None] * (len(S[n][0]) - 2))
    else:
        S[n][i][1] = (S[n][i][1] or "") + " " + "x" * rng.choice([10, 600])


def p_labels(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    for i in _table_rows(S[n]):
        r = S[n][i]
        if len(r) > 3 and rng.random() < 0.25:
            c = rng.choice([1, 2, 3])
            r[c] = rng.choice(["Third Party", "", None, "Upto 1000", "1000-1500", ">1500", "NA", "With NCB", "Without NCB", "Roll over SAOD",
                               "New Car", "Renewal & Roll Over Comp (1+1)", "x", 7, "Below 800 cc", "above 1500"])


def p_header(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    h = next((i for i, r in enumerate(S[n]) if "Business Type" in r), None)
    if h is None:
        return
    j = rng.choice([1, 2, 3])
    S[n][h][j] = rng.choice(["Business Type", "Type", "CC", "Engine", "NCB Status", "NCB", "", None, "business  type"])


def p_stp(S, rng):
    if "Pvt Car STP" not in S:
        return
    rows = S["Pvt Car STP"]
    how = rng.randint(0, 5)
    if how == 0 and len(rows) > 2:
        j = rng.choice([1, 2, 3, 4, 5, 6, 7])
        rows[1][j] = rng.choice(["DIESEL", "Non Diesel", "Petrol", "EV", "", None, "Gasoline", "RTO DIVISION"])
    elif how == 1 and len(rows) > 2:
        j = rng.choice([2, 3, 4, 5, 6, 7])
        rows[2][j] = rng.choice(["Below 1000 CC", "Above 1500 CC", "", None, "x", "1500 to 1000 CC", 5])
    elif how == 2:
        rows.insert(rng.randint(0, len(rows) - 1), [rng.choice(["PC STP GRID: new rule", "x", None])] + [None] * 7)
    elif how == 3:
        i = rng.randint(3, len(rows) - 1)
        rows[i][rng.randint(2, 7)] = rng.choice([None, "NA", "abc", -0.1, 12, 0.5])
    elif how == 4:
        del rows[rng.randint(1, len(rows) - 1)]
    else:
        rows.append(["Extra", "NEW DIVISION", 0.1, 0.2, 0.3, 0.1, 0.2, 0.3])


def p_extra_cols(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    h = next((i for i, r in enumerate(S[n]) if "Business Type" in r), None)
    if h is None:
        return
    w = max(len(r) for r in S[n])
    for r in S[n]:
        r.extend([None] * (w + 2 - len(r)))
    c = w + rng.randint(0, 1)
    S[n][h][c] = rng.choice(["Remarks", "NEW,CITY", "EXTRA CLUSTER", None, "", "REST OF  X", "Mumbai"])
    for i in _table_rows(S[n]):
        S[n][i][c] = rng.choice([0.2, 0.0, None, "NA", "abc", 0.35])


def p_cells(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    f = rng.choice([100, 0.01, 10, -1, 1])
    for i in _table_rows(S[n]):
        r = S[n][i]
        for c in range(4, len(r)):
            if isinstance(r[c], (int, float)) and not isinstance(r[c], bool) and rng.random() < 0.2:
                r[c] = rng.choice([r[c] * f, 0, None, "NA", "abc", "26%", " 0.26 "])


def p_title(S, rng):
    n = _comp(S, rng)
    if n is None:
        return
    new = rng.choice(["Pvt Car- 2W-COMP", "GCV 4W", "", "Sheet9", "Pvt Car- 4W-COMP & SAOD East"])
    S[new if new not in S else new + " (2)"] = S.pop(n)


def p_drive(S, rng):
    """the September 'Top Variant Models Drive' block (and its look-alikes): change, blank or break one of its cells"""
    n = _comp(S, rng)
    if n is None:
        return
    cells = [(i, j) for i, r in enumerate(S[n]) for j, v in enumerate(r)
             if isinstance(v, str) and re.search(r"additional|business share|NCB:|Base payout|Eligible Makes|Top Varient|over & above", v)]
    if not cells:
        return
    i, j = rng.choice(cells)
    S[n][i][j] = rng.choice([None, "", "x", "7% additional", "four percent additional", "Base payout", "With NCB:", "Without NCB:",
                             "Eligible Makes - Maruti", "10% additional", str(S[n][i][j]) + " (revised)", 5, "-2% additional", "75% business share from Top variant models."])


PC_MUTS = [p_notes, p_labels, p_header, p_stp, p_extra_cols, p_cells, p_title, p_drive]
POOL = PC_MUTS * 3 + fuzzlib.MUTATIONS


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
            if "engine-size limits" in q:
                self.prompts += 1
                return self.rng.choice(["1000-1500", "upto 1000", ">1500", "abc", "5", "1500-1000"])
        return super().inp(prompt)


def check_values(df, allow_dups=False):
    """arithmetic that must hold for any grid"""
    problems = []
    mod = G["mod"]
    pi = pd.concat([pd.to_numeric(df["pi_od_rate"], errors="coerce"), pd.to_numeric(df["pi_tp_rate"], errors="coerce")], axis=1).max(axis=1)
    po = pd.concat([pd.to_numeric(df["po_od_rate"], errors="coerce"), pd.to_numeric(df["po_tp_rate"], errors="coerce")], axis=1).max(axis=1)
    exp = (pi - mod.PO_CUT).clip(lower=0).where(pi > 0, 0.0)
    if ((po - exp).abs() > 0.0051).any():
        problems.append("po is not pi - 7 (floor 0)")
    if (pi < 0).any():
        problems.append("negative pi")
    if df.duplicated(mod.KEY_COLUMNS).any() and not allow_dups:
        problems.append("duplicate keys")
    if not set(df.sub_product) <= {"1+1", "SAOD", "1+3", "STP"}:
        problems.append("odd sub_product")
    if not set(df.new_vehicle_makes) <= {mod.MAKES_TOP, mod.MAKES_OTHER, mod.MAKES_STP}:
        problems.append("odd make group")
    if not set(df.is_ncb) <= {"YES", "NO", "NA"}:
        problems.append("odd is_ncb")
    if not set(df.fuel_type) <= set(mod.FUELS):
        problems.append("odd fuel")
    if not (df.id == "1").all():
        problems.append("id not 1")
    if (df.new_rto_list.str.contains(r"[\s,()]")).any():
        problems.append("a space, comma or bracket in new_rto_list")
    if "remarks" in df and (((df.remarks != "") & ~df.remarks.str.startswith("Manual Review Required: [")).any()):
        problems.append("a remark that is not a manual-review flag")
    return problems


def outcome(fn):
    mod = G["mod"]
    try:
        return "ok", fn()
    except mod.GridError as e:
        msg = str(e).replace("\n", " ")
        if re.search(r"\bnotes?\b|Additional|Top Variant|other make", msg):
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
    tmp = tempfile.mkdtemp(prefix="fzp_")
    which = rng.choice(["oct", "sep"])
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
    for key in ("nothing was written", "stopped by you", "nothing to write", "no keyboard input", "no Private Car", "engine-size", "do not join",
                "unexpected value", "holds", "repeat a key", "business type", "NCB label", "cannot read", "can't find", "no STP", "no cluster",
                "has no", "title", "looks like a rate sheet", "fuel", "twice", "STP"):
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
    json.dump(results, open(os.path.join(SC, "fuzz_pvtcar_results.json"), "w"), indent=1, default=str)
    for label in ("interactive", "silent"):
        c = Counter(bucket(r[label]) for r in results)
        print(f"\n== pvtcar: {label} run, {n} cases ==")
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
