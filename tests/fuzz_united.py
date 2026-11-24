"""Mutation fuzzer for united_commission_to_config.py.   usage: python tests\\fuzz_united.py N [first_seed] [workers]

Every case: corrupt the real July United PDF (its cells, as read by pdfplumber, held in memory), then run the whole script (a) with a random 'user'
answering every question and (b) with no keyboard. Acceptable outcomes: a correct file, or a clean, explained stop (GridError) about the DOCUMENT.
Anything else is a bug: another exception, a malformed file, wrong arithmetic, a hang, a flood of questions, or a stop caused by a rule text
(texts must be flagged in the remarks column, never stop the run)."""
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

SCRIPT = os.path.join(helpers.ROOT, "united_commission_to_config.py")
AUG = "aug" in sys.argv[4:5]                         # a 4th argument "aug" fuzzes the August circular instead of the July one
APR = "apr" in sys.argv[4:5]                         # ... "apr" the April circular ('Commission structure.pdf': Maximum proposed headings, merged class cells)
PDF = os.path.join(helpers.INPUTS, "united_apr26.pdf" if APR else "united_aug26.pdf" if AUG else "united_jul26.pdf")
quiet = lambda *a, **k: None
G = {}
F, T = (dt.date(2026, 4, 1), dt.date(2030, 3, 31)) if APR else (dt.date(2026, 8, 1), dt.date(2026, 8, 31)) if AUG else (dt.date(2026, 7, 1), dt.date(2026, 7, 31))


def init():
    spec = importlib.util.spec_from_file_location("united", SCRIPT)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    G["mod"] = m
    G["base"] = m.read_pdf(PDF)


TEXTS = ["NA", "", "-", "12,5%", "abc", "0%", "100%", "-5%", "5", "All States", "Other than above States", "Madhya Pradesh", "Atlantis, Goa", "Haryana and Rajasthan",
         "Tamil Nadu, Pondicherry (Except RTOs listed in Sub Annexure 3)", "OD - 10%, TP - 0%", "OD - 10%", "57.5 %", "0.575", "ALL STATES", "j. GVW > 40000", "PCC > 30 (Except EV)",
         "10 < PCC <= 20 (All Fuel)", "Upto 150 CC", "Taxis", "Motor Trade", "x" * 300, "01-08-2026", "New rule: all payouts 3% lower"]


def _cells(T_, rng):
    allc = [(t, c) for t in T_ for c in t.cells]
    return rng.choice(allc) if allc else (None, None)


def u_text(T_, rng):
    t, c = _cells(T_, rng)
    if c is not None:
        c.text = rng.choice(TEXTS)


def u_percent(T_, rng):
    cs = [(t, c) for t in T_ for c in t.cells if re.fullmatch(r"\d+(\.\d+)?%", c.text)]
    if cs:
        t, c = rng.choice(cs)
        c.text = rng.choice(["0%", "100%", "7%", "7.5 %", "7,5%", "NA", "-", "", "seven", "OD - 5%, TP - 2%", f"{rng.randint(1, 99)}.{rng.randint(0, 9)}%"])


def u_states(T_, rng):
    cs = [(t, c) for t in T_ for c in t.cells if "States" in c.text or "Pradesh" in c.text or "Nadu" in c.text]
    if cs:
        t, c = rng.choice(cs)
        c.text = rng.choice(["All States", "Other than above States", "Goa, Atlantis", "", "Tamil Nadu", "Tamil Nadu, Tamil Nadu", "Karnataka (Except RTOs listed in Sub Annexure 3)",
                             "Uttarakhand, Odisha, Telangana", "West Bengal, Maharashtra, Gujarat, Delhi, Uttarakhand, Punjab, Goa, Jammu & Kashmir", "all   states"])


def u_delete_cells(T_, rng):
    t, _ = _cells(T_, rng)
    if t is not None:
        for _ in range(rng.randint(1, 6)):
            if t.cells:
                t.cells.pop(rng.randrange(len(t.cells)))


def u_move_cell(T_, rng):
    t, c = _cells(T_, rng)
    if c is not None:
        dy = rng.choice([-30, -8, -1, 1, 8, 30])
        c.top, c.bottom = c.top + dy, c.bottom + dy


def u_resize_cell(T_, rng):
    t, c = _cells(T_, rng)
    if c is not None:
        c.x1 = c.x1 + rng.choice([-20, -3, 3, 20])


def u_drop_table(T_, rng):
    if T_:
        T_.pop(rng.randrange(len(T_)))


def u_dup_table(T_, rng):
    if T_:
        T_.append(copy.deepcopy(rng.choice(T_)))


def u_swap_pages(T_, rng):
    if len(T_) >= 2:
        i, j = rng.sample(range(len(T_)), 2)
        T_[i], T_[j] = T_[j], T_[i]


def u_headers(T_, rng):
    hs = [c for t in T_ for c in t.cells if c.text.startswith(("Existing", "Pay-out", "Commission", "States"))]
    if hs:
        rng.choice(hs).text = rng.choice(["", "Old", "Payout", "Pay-out (w.e.f. 01-07-2026)", "Existing Pay-out (w.e.f. 01-04-2026)", "Commission", "Rate", "Pay-out (w.e.f. 31-02-2026)"])


def u_outside(T_, rng):
    if T_:
        rng.choice(T_).outside.append(rng.choice(["All payouts reduced by 2% w.e.f. 15 July 2026", "x", "", "Note: Ambulances 25%"]))


def u_annexure(T_, rng):
    cs = [(t, c) for t in T_ for c in t.cells if "Kgs" in c.text or "%" in c.text and "w.e.f" in c.text or re.fullmatch(r"[A-Z]{2}\s?\d+", c.text)]
    if cs:
        t, c = rng.choice(cs)
        c.text = rng.choice(["Goods Carrying Vehicles 2000 - 3500 Kgs", "Goods Carrying Vehicles 7500 - 12000 Kgs", "w.e.f. 01-07-2026) 12.5%", "", "abc", "XX99", "Pay-out (w"])


def u_shuffle_cells(T_, rng):
    t, _ = _cells(T_, rng)
    if t is not None:
        rng.shuffle(t.cells)


def u_empty_table(T_, rng):
    t, _ = _cells(T_, rng)
    if t is not None:
        t.cells = []


MUTS = [u_text, u_percent, u_states, u_delete_cells, u_move_cell, u_resize_cell, u_drop_table, u_dup_table, u_swap_pages, u_headers, u_outside, u_annexure, u_shuffle_cells,
        u_empty_table]
POOL = MUTS * 2 + [u_percent, u_states, u_text]


def mutate(base, rng, how_many):
    T_ = copy.deepcopy(base)
    chosen = [rng.choice(POOL) for _ in range(how_many)]
    for m in chosen:
        m(T_, rng)
    return T_, [m.__name__ for m in chosen]


def check_values(df, allow_dups=False):
    """arithmetic and vocabulary that must hold for any document"""
    mod = G["mod"]
    problems = []
    pi = pd.concat([pd.to_numeric(df[c], errors="coerce") for c in ("pi_od_rate", "pi_tp_rate", "pi_net_rate")], axis=1).max(axis=1)
    po = pd.concat([pd.to_numeric(df[c], errors="coerce") for c in ("po_od_rate", "po_tp_rate", "po_net_rate")], axis=1).max(axis=1)
    if ((po - (pi - mod.PO_CUT).clip(lower=0).where(pi > 0, 0.0)).abs() > 0.0051).any():
        problems.append("po is not pi - 7 (floor 0)")
    if (pi < 0).any():
        problems.append("negative pi")
    if df.duplicated(mod.KEY_COLUMNS).any() and not allow_dups:
        problems.append("duplicate keys")
    if not set(df.sub_product) <= {"1+1", "1+3", "STP", "SAOD", "1+5"}:
        problems.append("odd sub_product")
    if not set(df.fuel_type) <= set(mod.FUELS_ALL):
        problems.append("odd fuel")
    car_groups = {mod.CAR_RTO_ANNEX4, mod.CAR_RTO_EX_ANNEX4}
    if AUG or APR:
        g = mod.circular_groups(F)
        car_groups |= set(g["car"]) | {n for pair in g["gcv"].values() for n in pair if n}
    if not df.new_rto_list.map(lambda s: s == mod.ALL_INDIA or s in car_groups or s.startswith(mod.RTO_PREFIX) or all(p in mod.STATE_CODES for p in s.split(","))).all():
        problems.append("a new_rto_list that is no ALLINDIA, state-code list or RTO group")
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
        if re.search(r"text outside|rule text|\bnote\b", msg):
            return "BAD: a stop caused by a rule text (it must be flagged, never stop the run): " + msg[:140], None
        return "stopped: " + msg[:140], None
    except fuzzlib.Overload as e:
        return "OVERLOAD: " + str(e), None
    except BaseException as e:
        return f"CRASH {type(e).__name__}: {str(e)[:160]}\n{traceback.format_exc()[-1800:]}", None


def run_case(seed):
    mod = G["mod"]
    rng = random.Random(seed)
    t0 = time.time()
    tmp = tempfile.mkdtemp(prefix="fzu_")
    tbls, muts = mutate(G["base"], rng, rng.choice([0, 1, 1, 2, 3, 4]))

    def call(ask, out):
        # mod.main reads the PDF through open_file; hand it the mutated tables instead
        mod.read_pdf = lambda p: tbls
        return mod.main("fake.pdf", F, T, out, None, quiet, ask)

    res = {"seed": seed, "muts": muts}
    user = fuzzlib.RandomUser(random.Random(seed + 7))
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
    if not muts and kind_i == "ok" and user.prompts > 2 and not retyped:
        kind_i = f"BAD OUTPUT: a clean document triggered {user.prompts - 1} unneeded question(s)"
    res["interactive"], res["prompts"] = kind_i, user.prompts
    if kind_i.startswith(("CRASH", "BAD", "OVERLOAD")):
        res["transcript"] = "\n".join(user.transcript[-30:])
    kind_n, path = outcome(lambda: call(mod.Asker(False), os.path.join(tmp, "n.csv")))
    kind_n = verify(kind_n, path)
    if not muts and kind_n != "ok":
        kind_n = "BAD OUTPUT (silent): a clean document did not run: " + kind_n[:100]
    res["silent"] = kind_n
    res["secs"] = round(time.time() - t0, 1)
    return res


def bucket(msg):
    if not msg.startswith("stopped"):
        return msg.split("\n")[0][:150]
    m = msg[9:]
    for key in ("nothing was written", "stopped by you", "nothing to write", "no keyboard input", "no payout found", "not a percentage", "repeat a key", "duplicate rows",
                "before from_date", "cannot open", "no bordered table"):
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
    json.dump(results, open(os.path.join(SC, "fuzz_united_apr_results.json" if APR else "fuzz_united_aug_results.json" if AUG else "fuzz_united_results.json"), "w"), indent=1, default=str)
    for label in ("interactive", "silent"):
        c = Counter(bucket(r[label]) for r in results)
        print(f"\n== united: {label} run, {n} cases ==")
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
        print(f"\n--- seed {r['seed']} ({which}) mutations={r['muts']}\n{r[which]}")
        if which == "interactive" and r.get("transcript"):
            print("last output before it:\n" + r["transcript"][-1200:])
