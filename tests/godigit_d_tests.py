"""Go Digit: a D (decline) must never beat a real rate of another grid row for the same card (the user, 18 Nov 2026: 'Yes, fix it').
A D in a payout cell of a row that covers only some ages / makes / discount bands means 'not on this row' (Delhi NCR taxi: 'Hyundai' pays 15% for age 0 and says D for age 1+, the row above says
'Hyundai age 1+' = 25%; Pune PCV3W: the 'New' rows say D for 1 yr+, the '1 yr+' row says 67%).
Run:  .venv\\Scripts\\python.exe tests\\godigit_d_tests.py"""
import datetime as dt
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("godigit_agency_to_config.py")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F, T = dt.date(2026, 11, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


mod.MISSING_TABLES = "skip"
HEAD = ["Cluster segment", "Segment", "Body Type", "RTO Cluster", "State", "Make / Variant", "CD1 (OD-disc)"] + ["New", "1 Yr", "2 Yr", "3 Yr", "4 Yr", "5 Yr", "6-10 Yr", "11+ Yr"] * 2
HCV_NOTE = "Note: Volvo, Scania to block. \nOnly for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1"


def book(cv_rows):
    return mod.Book([mod.Sheet("LATEST CV COMBINED GRID", [[], ["", "", "", "", "", "", "", "COMPREHENSIVE"], ["", "", "", "", "", "", "", "CD2"], HEAD] + [list(r) for r in cv_rows]),
                     mod.Sheet("CV Remarks", [["REMARKS FOR HCV"], [HCV_NOTE]])])


def run(rows):
    mod.reset_run_lists()
    return mod.build(None, F, T, quiet, mod.Asker(interactive=False), book=book(rows))


def row(cd1, comp, satp, make="All", cluster="Delhi"):
    """a heavy-vehicle dumper row with its 8 comprehensive and 8 third-party cells"""
    return ("CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", cluster, "Delhi", make, cd1) + tuple(comp) + tuple(satp)


def cards(df, sub, age, fuel="Diesel"):
    d = df[(df.sub_product == sub) & (df.vehicle_age_min == age) & (df.fuel_type == fuel)].sort_values("tariff_max")
    return list(zip(d.tariff_min, d.tariff_max, d.pi_net_rate if sub != "STP" else d.pi_tp_rate))


D8 = ["D"] * 7
# ---------------------------------------------------------------- the Delhi NCR pattern: one row pays 25% for every age, the other 15% for New and D for the rest
df, w = run([row("0-55%", [0.25] * 8, [0.25] * 8), row("0-50%", [0.15] + D8, [0.25] + D8)])
check("D never beats a real rate (same band): the third-party card of age 1 is 25, not 0", cards(df, "STP", 1) == [(0.0, 100.01, 25.0)], cards(df, "STP", 1))
check("D below a real band does not turn the discounts under it into 'no payout': age 1 comprehensive = 25 for 0-55.01 (no zero band 0-50.01)", cards(df, "1+1", 1) == [(0.0, 55.01, 25.0)], cards(df, "1+1", 1))
check("a real rate of another row is kept where the D row has one: age 0 = 15 up to 50.01, then 25 up to 55.01 (two real rows chain)", cards(df, "1+1", 0) == [(0.0, 50.01, 15.0), (50.01, 55.01, 25.0)], cards(df, "1+1", 0))
check("the ignored D is listed (sheet 'D ignored') and said in the run's warnings; nothing is left in 'Lowest rate chosen' as a zero", len(mod.D_IGNORED) > 0 and any("D (decline)" in x for x in w)
      and not any(x["written"] in ("0", "0.0") for x in mod.LAST_LOWEST), (len(mod.D_IGNORED), mod.LAST_LOWEST[:1]))
mod.collect_gaps(df)
check("... and the review workbook gets the sheet 'D ignored' with the rate that stays", "D ignored" in mod.EXTRA_SHEETS and "25 up to 55.01" in set(mod.EXTRA_SHEETS["D ignored"].rate_written), list(mod.EXTRA_SHEETS))

# ---------------------------------------------------------------- the Pune pattern: rows for New (three discount bands, D for the rest) and one row for 1 yr+ (D for New)
df, w = run([row("0–10%", [0.55] + D8, [0.515] + D8), row("10.1–30%", [0.525] + D8, [0.515] + D8), row("30.1–90%", [0.45] + D8, [0.515] + D8), row("1 yr+ cd1 - 0–85%", ["D"] + [0.67] * 7, ["D"] + [0.675] * 7)])
check("Pune: 1 yr+ is ONE card, 67 for 0-85.01 - the D bands of the New rows (up to 10 / 30 / 90) neither cut it nor overlap it", cards(df, "1+1", 1) == [(0.0, 85.01, 67.0)], cards(df, "1+1", 1))
check("Pune: New keeps its three real bands (55 / 52.5 / 45) and the decline above 90 (the 1 yr+ row's D for New, an explicit decline above every real band)",
      cards(df, "1+1", 0) == [(0.0, 10.01, 55.0), (10.1, 30.01, 52.5), (30.1, 90.01, 45.0), (90.01, 100.01, 0.0)], cards(df, "1+1", 0))
check("Pune: third party New = 51.5, 1 yr+ = 67.5 (the D of the other rows never wins)", cards(df, "STP", 0) == [(0.0, 100.01, 51.5)] and cards(df, "STP", 1) == [(0.0, 100.01, 67.5)], (cards(df, "STP", 0), cards(df, "STP", 1)))

# ---------------------------------------------------------------- what stays as it was
df, w = run([row("0-70%", [0.2] * 8, [0.2] * 8), row("70.1-100%", ["D"] * 8, ["D"] * 8)])
check("a D band ABOVE the real band stays (an explicit decline above the cap): 20 up to 70.01, then a zero card up to 100.01", cards(df, "1+1", 2) == [(0.0, 70.01, 20.0), (70.1, 100.01, 0.0)], cards(df, "1+1", 2))
df, w = run([row("0-60%", ["D"] * 8, ["D"] * 8)])
check("a row that is D everywhere is a decline: a zero card, nothing is dropped", cards(df, "1+1", 2) == [(0.0, 60.01, 0.0)] and cards(df, "STP", 2) == [(0.0, 100.01, 0.0)] and not mod.D_IGNORED, (cards(df, "1+1", 2), cards(df, "STP", 2)))
df, w = run([row("0-60%", [0.2] * 8, [0.2] * 8), row("0-60%", [0.25] * 8, [0.25] * 8)])
check("two REAL rates for one card: still the lowest (your rule of 27 Oct), listed in 'Lowest rate chosen'", cards(df, "STP", 2) == [(0.0, 100.01, 20.0)] and len(mod.LAST_LOWEST) > 0 and not mod.D_IGNORED, (cards(df, "STP", 2), mod.LAST_LOWEST[:1]))
df, w = run([row("0-60%", [0.0] * 8, [0.0] * 8), row("0-60%", [0.25] * 8, [0.25] * 8)])
check("a NUMBER 0% is a real rate, not a D: the lowest rule keeps it (0 beats 25)", cards(df, "STP", 2) == [(0.0, 100.01, 0.0)] and not mod.D_IGNORED, cards(df, "STP", 2))
df, w = run([row("0-60%", [None] * 8, [None] * 8), row("0-60%", ["D"] * 8, ["D"] * 8)])
check("a blank cell (a placeholder) never beats a D (an explicit decline): the zero card stays, once", cards(df, "STP", 2) == [(0.0, 100.01, 0.0)], cards(df, "STP", 2))

# ---------------------------------------------------------------- the real workbooks
nov = os.path.join(ROOT, "regression", "inputs", "godigit_nov26.xlsx")
octo = os.path.join(ROOT, "regression", "inputs", "godigit_oct26.xlsx")
if os.path.exists(nov):
    feed = iter(["1"])

    def answer(prompt):
        try:
            return next(feed)
        except StopIteration:
            raise EOFError
    mod.reset_run_lists()
    df, w = mod.build(nov, F, T, quiet, mod.Asker(interactive=True, input_fn=answer, out=lambda *x: None))
    low = pd.DataFrame(mod.LAST_LOWEST)
    di = pd.DataFrame(mod.D_IGNORED)
    check("November: no 'Lowest rate chosen' entry writes 0 any more (it was 116 of 236); 120 real conflicts are left", len(low) == 120 and not low["written"].isin(["0", "0.0"]).any(), (len(low), low["written"].value_counts().head(3).to_dict()))
    check("November: 193 entries had a D ignored (taxi 63, PCV 2W 60, PCV 3W 60, GCV 4W 10)", len(di) == 193 and dict(di.groupby("product").size()) == {"GCV 4W": 10, "PCV 2W": 60, "PCV 3W": 60, "PCV 4W": 63}, dict(di.groupby("product").size()))
    x = df[(df.new_rto_list == "GO-DIGIT_OCT26_PCV4W_UPTO-5_DELHI-NCR") & (df.new_vehicle_makes == "Hyundai") & (df.fuel_type == "Petrol") & (df.cc_max == 1000.01) & (df.vehicle_age_min == 1) & (df.sub_product == "STP")]
    check("November: Delhi NCR taxi Hyundai Petrol age 1 third party = 25 (rows 168 / 1322 of the grid), it was 0", list(x.pi_tp_rate) == [25.0], list(x.pi_tp_rate))
if os.path.exists(octo):
    mod.reset_run_lists()
    df, w = mod.build(octo, dt.date(2026, 10, 15), T, quiet, mod.Asker(interactive=False))
    check("October: the rule changes nothing (53,212 rows, no D ignored): the October grid has no D beside a real rate", len(df) == 53212 and not mod.D_IGNORED, (len(df), len(mod.D_IGNORED)))

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
