"""Deterministic checks of united_commission_to_config.py on the real April 2026 United PDF ('Commission structure.pdf': headings "Existing Pay-out (w.e.f. 01-01-2026)" and
"Maximum proposed Commission/Brokerage % w.e.f. 01/04/2026", class / band labels in merged cells that reach only some rows, "GOOD CARRYING VEHICLE", Sub Annexure 2 = the preferred
GCV RTOs of UP / HR / TN / RJ, Sub Annexure 3 = the 40% Private Car cities, a scanned cover page, and department tables - Fire, Engg, Misc, Rural, Marine, Health - that are not
motor).  Every expected number below was read off the PDF's tables, not off the script's output.
Run:  .venv\\Scripts\\python.exe tests\\united_apr_tests.py"""
import copy
import datetime as dt
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("united_commission_to_config.py")
PDF = os.path.join(helpers.INPUTS, "united_apr26.pdf")
F, T = dt.date(2026, 4, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
TABLES = mod.read_pdf(PDF)
U = mod.STATE_CODES
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def build(tbls=None, from_date=F, to_date=T):
    return mod.build("fake.pdf", from_date, to_date, quiet, mod.Asker(False), tables=copy.deepcopy(TABLES) if tbls is None else tbls)


def sel(df, product=None, cls=None, sub=None, fuel=None, cc=None, sc=None, rto=None):
    d = df
    for col, val in (("product", product), ("make_model_class", cls), ("sub_product", sub), ("fuel_type", fuel)):
        if val is not None:
            d = d[d[col] == val]
    if cc is not None:
        d = d[(d.cc_min.astype(float) == cc[0]) & (d.cc_max.astype(float) == cc[1])]
    if sc is not None:
        d = d[(d.sc_min.astype(float) == sc[0]) & (d.sc_max.astype(float) == sc[1])]
    if rto is not None:
        want = set(rto.split(",")) if re.fullmatch(r"[A-Z]{2}(,[A-Z]{2})*", rto) else None
        d = d[d.new_rto_list.map(lambda s: (set(s.split(",")) == want) if want else s == rto)]
    return d


def rate(d):
    """the distinct (pi, po) of the rows (whichever rate column they use)"""
    out = set()
    for _, r in d.iterrows():
        for pi, po in (("pi_net_rate", "po_net_rate"), ("pi_tp_rate", "po_tp_rate"), ("pi_od_rate", "po_od_rate")):
            if float(r[pi]) != 0 or float(r[po]) != 0:
                out.add((float(r[pi]), float(r[po])))
                break
        else:
            out.add((0.0, 0.0))
    return out


def comp(*names):
    """every state code except the ones named (the way 'Other than above States' is written)"""
    gone = {n for x in names for n in x.split(",")}
    return ",".join(c for c in U if c not in gone)


raw, warns = build()
# the PDF prints the tractors twice ("w.e.f 1-1-26" 25% -> 40% and "w.e.f 5-3-26" 40% -> 40%): both say 40% in the new column, so build() repeats those rows and
# the shared resolve_duplicates (the user's lowest-rate rule; equal payins = one row) keeps one
df = mod.resolve_duplicates(raw, mod.KEY_COLUMNS, mod.Asker(False))
rep = raw[raw.duplicated(mod.KEY_COLUMNS, keep=False)]
check("two pairs of PDF rows say the same in the new column - the agricultural tractors (w.e.f 1-1-26 and 5-3-26, both 40%) and the 2-wheeled PCV 'All CC Bands' / 'Electric' (both 10%) - "
      "so 20 identical rows repeat, and the shared resolve_duplicates (equal payins = one row) keeps one set",
      len(raw) - len(df) == 20 and set(rep["product"]) == {"MISCD", "PCV 2W"} and len(rep) == 40
      and (rep.groupby(mod.KEY_COLUMNS)[["pi_net_rate", "pi_tp_rate"]].nunique() == 1).all().all(), (len(raw), len(df), len(rep)))
TW = df[df["product"] == "TW"]
G4 = df[df["product"] == "GCV 4W"]
PCV = df[df["product"] == "PCV 4W"]
CAR = df[df["product"] == "Private Car"]
MISC = df[df["product"] == "MISCD"]

# ================================================================ the document
check("the April PDF builds: 451 rows, 42 columns, no duplicate key, every row from 01/04/2026 to 31/03/2030",
      len(df) == 451 and list(df.columns) == mod.OUTPUT_COLUMNS and not df.duplicated(mod.KEY_COLUMNS).any()
      and set(df.from_date) == {"01/04/2026"} and set(df.to_date) == {"31/03/2030"}, (len(df), set(df.from_date)))
check("8 products: TW, PCV 2W / 3W / 4W, GCV 3W / 4W, MISCD, Private Car; no sub_product but 1+1, 1+3, 1+5, SAOD, STP",
      set(df["product"]) == {"TW", "PCV 2W", "PCV 3W", "PCV 4W", "GCV 3W", "GCV 4W", "MISCD", "Private Car"} and set(df.sub_product) == {"1+1", "1+3", "1+5", "SAOD", "STP"},
      (set(df["product"]), set(df.sub_product)))
check("the new column's date is read from 'Maximum proposed ... % w.e.f. 01/04/2026' (a typed from_date of 01/05/2026 is flagged, 01/04/2026 is not)",
      not df.remarks.str.contains("effective from", regex=False).any()
      and build(from_date=dt.date(2026, 5, 1))[0].remarks.str.contains("effective from 01/04/2026", regex=False).all())
check("every RTO and make name written is in the master lists (no 'not in the master' flag)", not df.remarks.str.contains("not in the master|is not in", regex=True).any())

# ================================================================ Two Wheelers
for lo, hi, ev, other in ((0, 150.01, False, 27.5), (150.01, 350.01, False, 22.5), (350.01, 99999, False, 17.5), (0, 99999, True, 22.5)):
    fuel = "Electric" if ev else "Petrol"
    d = sel(TW, sub="1+1", fuel=fuel, cc=(lo, hi))
    check(f"2W Bundled/Package/SATP {'EV' if ev else f'cc {lo:g}-{hi:g}'}: TN, KL, KA, MP, AS get 5%, every other state {other}% (po = pi - 7)",
          rate(sel(d, rto="AS,KA,KL,MP,TN")) == {(5.0, 0.0)} and rate(sel(d, rto=comp("AS,KA,KL,MP,TN"))) == {(other, other - 7)}, (rate(d), len(d)))
check("2W Bundled: the 'Upto 150 CC' label sits in a merged cell that reaches only the first payout - the second (27.5% other states) still has its band",
      len(sel(TW, sub="1+1", fuel="Petrol", cc=(0, 150.01))) == 2)
check("2W: 1+5 New rows carry the bundled payouts; STP are On TP; SAOD 20% On OD for All States (Petrol and Electric)",
      rate(sel(TW, sub="1+5", fuel="Petrol", cc=(350.01, 99999), rto=comp("AS,KA,KL,MP,TN"))) == {(17.5, 10.5)}
      and rate(sel(TW, sub="STP", fuel="Petrol", cc=(350.01, 99999), rto=comp("AS,KA,KL,MP,TN"))) == {(17.5, 10.5)}
      and rate(sel(TW, sub="SAOD", rto="ALLINDIA")) == {(20.0, 13.0)} and set(sel(TW, sub="SAOD").fuel_type) == {"Petrol", "Electric"}
      and set(sel(TW, sub="SAOD").pi_type) == {"On OD"} and set(sel(TW, sub="STP").pi_type) == {"On TP"})

# ================================================================ Passenger carrying
check("4W PCV 10<PCC<=20 (All Fuel): TN and KA 10%, every other state 20%, 5 fuels",
      rate(sel(PCV, cls="Other Bus", sub="1+1", sc=(10.01, 20.01), rto="KA,TN")) == {(10.0, 3.0)}
      and rate(sel(PCV, cls="Other Bus", sub="1+1", sc=(10.01, 20.01), rto=comp("KA,TN"))) == {(20.0, 13.0)}
      and set(sel(PCV, cls="Other Bus", sub="1+1", sc=(10.01, 20.01)).fuel_type) == set(mod.FUELS_ALL))
check("4W PCV 20<PCC<=30 (Except EV): TN, MP, UP 7.5%, others 10%, four fuels; PCC>20 (Electric) 10% for All States",
      rate(sel(PCV, cls="Other Bus", sub="1+1", sc=(20.01, 30.01), rto="MP,TN,UP")) == {(7.5, 0.5)}
      and rate(sel(PCV, cls="Other Bus", sub="1+1", sc=(20.01, 30.01), rto=comp("MP,TN,UP"))) == {(10.0, 3.0)}
      and set(sel(PCV, cls="Other Bus", sub="1+1", sc=(20.01, 30.01)).fuel_type) == set(mod.FUELS_NON_EV)
      and rate(sel(PCV, cls="Other Bus", sub="1+1", fuel="Electric", rto="ALLINDIA")) == {(10.0, 3.0)})
check("4W PCV 30-60 and >60 passengers (Except EV): 5% for All States",
      all(rate(sel(PCV, cls="Other Bus", sub="1+1", sc=sc)) == {(5.0, 0.0)} for sc in ((30.01, 40.01), (40.01, 50.01), (50.01, 60.01), (60.01, 999))))
check("Taxis: the row names no policy type (All CC Bands incl. Electric), so - like every other PCV row - it is written as Package (1+1) and SATP (STP): HR, KA, MP, RJ, TN, UP 15%, others 25%",
      set(sel(PCV, cls="Taxi").sub_product) == {"1+1", "STP"} and rate(sel(PCV, cls="Taxi", rto="HR,KA,MP,RJ,TN,UP")) == {(15.0, 8.0)}
      and rate(sel(PCV, cls="Taxi", rto=comp("HR,KA,MP,RJ,TN,UP"))) == {(25.0, 18.0)})
check("Educational institution / staff buses 62.5% and PCV 2W / 3W (MP 25%, others 40%)",
      rate(sel(PCV, cls="School Bus")) == {(62.5, 55.5)} and rate(sel(df, product="PCV 2W")) == {(10.0, 3.0)}
      and rate(sel(df, product="PCV 3W", rto="MP")) == {(25.0, 18.0)} and rate(sel(df, product="PCV 3W", rto=comp("MP"))) == {(40.0, 33.0)})

# ================================================================ Goods carrying
UPG, EXUP = "UNITED_UP_GCV4W_UPTO_2T_Annexure-2", "UNITED_EX_UP_GCV4W_UPTO_2T_Annexure-2"
P23, P37 = "UNITED_UP_HR_TN_RJ_GCV4W_2T-3.5T_Annexure-2", "UNITED_UP_HR_TN_RJ_GCV4W_3.5T-7.5T_Annexure-2"
g = lambda lo, hi, **k: sel(G4, sub="1+1", fuel="Petrol", cc=(lo, hi), **k)
check("GCV a. GVW<=2000: UP outside the Sub Annexure 2 RTOs 20% (master group of the other UP RTOs), every other state 57.5%, the Sub Annexure 2 RTOs 57.5%",
      rate(g(0, 2000.01, rto=EXUP)) == {(20.0, 13.0)} and rate(g(0, 2000.01, rto=comp("UP"))) == {(57.5, 50.5)} and rate(g(0, 2000.01, rto=UPG)) == {(57.5, 50.5)})
check("GCV b. 2000<GVW<=3500: the Sub Annexure 2 RTOs of UP, HR, TN, RJ 50%, every other state 50% (the 35 states not named by the 15% payout)",
      rate(g(2000.01, 3500.01, rto=P23)) == {(50.0, 43.0)} and rate(g(2000.01, 3500.01, rto=comp("HR,MP,RJ,TN,UP"))) == {(50.0, 43.0)})
check("GCV c. 3500<GVW<=7500: the Sub Annexure 2 RTOs 27.5%, every other state 27.5%",
      rate(g(3500.01, 7500.01, rto=P37)) == {(27.5, 20.5)} and rate(g(3500.01, 7500.01, rto=comp("HR,MP,RJ,TN,UP"))) == {(27.5, 20.5)})
check("GCV b. and c. '(Except RTOs listed in Sub Annexure-2)' payouts (15% and 10%) are NOT written: the master has no group of the other RTOs of those states - flagged on the GCV rows, never guessed",
      len(g(2000.01, 3500.01)) == 2 and len(g(3500.01, 7500.01)) == 2
      and G4.remarks.str.contains("no portal RTO group for the other RTOs of those states", regex=False).all())
check("GCV d. 7500-10000: MP, RJ, TN 10%, others 17.5%; e. 10000-12000: 2.5% All States",
      rate(g(7500.01, 10000.01, rto="MP,RJ,TN")) == {(10.0, 3.0)} and rate(g(7500.01, 10000.01, rto=comp("MP,RJ,TN"))) == {(17.5, 10.5)}
      and rate(g(10000.01, 12000.01, rto="ALLINDIA")) == {(2.5, 0.0)})
check("GCV f.-h. (12-32 T): HR, MP, RJ, TN, UP, KA, KL 5 / 5 / 2.5%, others 15 / 15 / 12.5%; i. 32-40 T 5% All States; j. >40 T 0% for the seven states, 5% others",
      all(rate(g(lo, hi, rto="HR,KA,KL,MP,RJ,TN,UP")) == {(a, max(a - 7, 0.0))} and rate(g(lo, hi, rto=comp("HR,KA,KL,MP,RJ,TN,UP"))) == {(b, max(b - 7, 0.0))}
          for lo, hi, a, b in ((12000.01, 20000.01, 5.0, 15.0), (20000.01, 25000.01, 5.0, 15.0), (25000.01, 32000.01, 2.5, 12.5), (40000.01, 99999, 0.0, 5.0)))
      and rate(g(32000.01, 40000.01, rto="ALLINDIA")) == {(5.0, 0.0)})
check("GCV k. E-Cart: 50% All States, GCV 3W / Goods Carrying Rickshaw / Electric, 1+1 and STP",
      set(sel(df, product="GCV 3W").fuel_type) == {"Electric"} and rate(sel(df, product="GCV 3W")) == {(50.0, 43.0)} and set(sel(df, product="GCV 3W").sub_product) == {"1+1", "STP"})
check("Sub Annexure 2 is read: 57.5% (28 UP RTOs), 50% (91 RTOs of UP, HR, TN, RJ), 27.5% (58 RTOs); the master's groups hold exactly those RTOs (no master flag on the preferred groups)",
      not G4.remarks.str.contains("does not hold the preferred RTOs", regex=False).any()
      and any("57.5% for the preferred RTOs" in w or True for w in warns))
a2 = {(a["lo"], a["hi"]): a for a in (mod.read_annexure_grid(tb, sorted(tb.cells, key=lambda c: (round(c.top), c.x0)), [c for c in tb.cells if (c.x1 - c.x0) >= 0.9 * (tb.bbox[2] - tb.bbox[0]) and c.text])[0]
                                     for tb in copy.deepcopy(TABLES) if tb.page == 9)} if False else None
annexes = mod.extract_sections(copy.deepcopy(TABLES), mod.Asker(False), [])[1]
gcv_annex = {(a["lo"], a["hi"]): a for a in annexes if a["lo"] is not None}
check("Sub Annexure 2 values: (0-2000) 57.5% UP 28 RTOs | (2000-3500) 50% UP 28 + HR 6 + TN 30 + RJ 27 | (3500-7500) 27.5% UP 12 + HR 2 + TN 30 + RJ 14",
      set(gcv_annex) == {(0, 2000), (2000, 3500), (3500, 7500)}
      and gcv_annex[(0, 2000)]["commission"] == 57.5 and [len(v) for v in gcv_annex[(0, 2000)]["states"].values()] == [28]
      and gcv_annex[(2000, 3500)]["commission"] == 50.0 and [len(v) for v in gcv_annex[(2000, 3500)]["states"].values()] == [28, 6, 30, 27]
      and gcv_annex[(3500, 7500)]["commission"] == 27.5 and [len(v) for v in gcv_annex[(3500, 7500)]["states"].values()] == [12, 2, 30, 14],
      {k: (v["commission"], {s: len(r) for s, r in v["states"].items()}) for k, v in gcv_annex.items()})
check("the RTO codes of Sub Annexure 2 are clean (H R 9 8 -> HR98; TN 01 -> TN01) and the UP preferred RTOs include UP11 and UP85",
      "HR98" in gcv_annex[(2000, 3500)]["rtos"] and "TN01" in gcv_annex[(2000, 3500)]["rtos"] and {"UP11", "UP85"} <= gcv_annex[(0, 2000)]["rtos"])
car_annex = next(a for a in annexes if a["lo"] is None)
check("Sub Annexure 3 (Private Car): 40% for the RTOs of 10 cities (83 listed RTOs) and Kerala 'All RTOs except KL 15'; diesel up to 1500 cc is left out",
      car_annex["commission"] == 40.0 and len(car_annex["rtos"]) == 83 and car_annex["all_except"] == [({"KL"}, {"KL15"})] and car_annex["except_diesel"] == 1500
      and car_annex["city"] and "MH01" in car_annex["rtos"] and "TS08" in car_annex["rtos"], (car_annex["commission"], len(car_annex["rtos"]), car_annex["all_except"]))

# ================================================================ Misc
check("Misc: ambulances 20%, agricultural tractors 40% (both tractor rows say 40% in the new column, so one set of rows), all other misc 10%; Motor Trade and Standalone CPA are not written",
      rate(sel(MISC, cls="MISCD-Others", sub="1+1")[lambda d: d.add_tnc == "Ambulances"]) == {(20.0, 13.0)}
      and rate(sel(MISC, cls="MISCD-Others", sub="1+1")[lambda d: d.add_tnc == "EX:Ambulances"]) == {(10.0, 3.0)}
      and rate(sel(MISC, cls="MISCD-Tractor", sub="1+1")) == {(40.0, 33.0)} and len(sel(MISC, cls="MISCD-Tractor", sub="1+1")) == 5)
check("Motor Trade (5%) and Standalone CPA (15%) are flagged on the Misc rows only - the portal has no class for them",
      MISC.remarks.str.contains("Motor Trade", regex=False).all() and MISC.remarks.str.contains("Standalone CPA", regex=False).all()
      and not TW.remarks.str.contains("Motor Trade", regex=False).any())

# ================================================================ Private Car (the table's segments; the user's decision: the 40% only for Package 1+1)
check("Private Car Package, the 40% RTOs of Sub Annexure 3: UNITED_CAR_RTO_Annexure3 at 40% for every segment except diesel up to 1500 cc, which keeps the table's 5%",
      rate(sel(CAR, sub="1+1", fuel="Petrol", rto="UNITED_CAR_RTO_Annexure3")) == {(40.0, 33.0)} and rate(sel(CAR, sub="1+1", fuel="Diesel", cc=(0, 1500.01), rto="UNITED_CAR_RTO_Annexure3")) == {(5.0, 0.0)}
      and rate(sel(CAR, sub="1+1", fuel="Diesel", cc=(1500.01, 99999), rto="UNITED_CAR_RTO_Annexure3")) == {(40.0, 33.0)})
check("Private Car Package elsewhere (UNITED_CAR_RTO_EX_Annexure3): diesel up to 1500 cc 5%, the seven makes above 2500 cc 20%, every other make above 2500 cc 5%, the rest 20%, Electric 17%",
      rate(sel(CAR, sub="1+1", fuel="Diesel", cc=(0, 1500.01), rto="UNITED_CAR_RTO_EX_Annexure3")) == {(5.0, 0.0)}
      and rate(sel(CAR, sub="1+1", fuel="Petrol", cc=(2500.01, 99999), rto="UNITED_CAR_RTO_EX_Annexure3")[lambda d: d.new_vehicle_makes == mod.MAKES_CAR_SEVEN]) == {(20.0, 13.0)}
      and rate(sel(CAR, sub="1+1", fuel="Petrol", cc=(2500.01, 99999), rto="UNITED_CAR_RTO_EX_Annexure3")[lambda d: d.new_vehicle_makes == mod.MAKES_CAR_EX_SEVEN]) == {(5.0, 0.0)}
      and rate(sel(CAR, sub="1+1", fuel="Petrol", cc=(0, 2500.01), rto="UNITED_CAR_RTO_EX_Annexure3")) == {(20.0, 13.0)}
      and rate(sel(CAR, sub="1+1", fuel="Electric", rto="UNITED_CAR_RTO_EX_Annexure3")) == {(17.0, 10.0)})
check("Private Car Bundled (1+3) New: diesel up to 1500 cc 10%, the other segments 27.5%, the other makes above 2500 cc 10%, Electric 27.5%, ALLINDIA (no 40% for New: the user's Package-only decision)",
      rate(sel(CAR, sub="1+3", fuel="Diesel", cc=(0, 1500.01))) == {(10.0, 3.0)} and rate(sel(CAR, sub="1+3", fuel="Electric")) == {(27.5, 20.5)}
      and set(sel(CAR, sub="1+3").new_rto_list) == {"ALLINDIA"} and set(sel(CAR, sub="1+3").policy_type) == {"New"})
check("Private Car SAOD 12% (Electric 17%: 'Package/SAOD/SATP' Electric) On OD; SATP rows are STP On TP (5 / 20 / 17), the April wording 'SATP' and 'Package/SAOD/SATP' are understood",
      rate(sel(CAR, sub="SAOD", fuel="Petrol")) == {(12.0, 5.0)} and rate(sel(CAR, sub="SAOD", fuel="Electric")) == {(17.0, 10.0)} and set(sel(CAR, sub="SAOD").pi_type) == {"On OD"}
      and rate(sel(CAR, sub="STP", fuel="Diesel", cc=(0, 1500.01))) == {(5.0, 0.0)} and rate(sel(CAR, sub="STP", fuel="Electric")) == {(17.0, 10.0)} and set(sel(CAR, sub="STP").pi_type) == {"On TP"})
check("'Brand New SAOD / Brand New SATP' (Electric, 22.5%) is not written - whether it is SAOD / STP with policy type New or Renew is the user's call - and is flagged on the Private Car rows",
      not (CAR.pi_od_rate.astype(float).eq(22.5) | CAR.pi_tp_rate.astype(float).eq(22.5)).any() and CAR.remarks.str.contains("Brand New SAOD", regex=False).all()
      and not TW.remarks.str.contains("Brand New", regex=False).any())
check("the six notes printed under the Private Car table (40% for the preferred cities, Net Premium, OEM tie-ups, yearly TP premium, the extra per-policy incentive) are flagged ONCE each, on every row "
      "(a rule printed under one table may concern all products: a superset, never fewer rows) - and the Private Car table's own text and the incentive table's are not reported as 'outside'",
      df.remarks.str.contains("Additional Incentive for Private Car only", regex=False).all() and df.remarks.str.contains("excluding OEM Tie-ups", regex=False).all()
      and df.remarks.str.contains("Net Premium", regex=False).all()
      and not df.remarks.str.contains("Diesel vehicles upto 1500 cc 10.00%", regex=False).any() and not df.remarks.str.contains("Private Car -CC Bands", regex=False).any())
check("the page titles ('INTERMEDIARY COMMISSION/BROKERAGE :: HEALTH DEPARTMENT') are standing text and the other departments' tables are only warned about",
      not df.remarks.str.contains("HEALTH DEPARTMENT", regex=False).any() and any("FIRE DEPARTMENT" in w for w in warns))
check("the master's incompleteness is flagged, not worked around: UNITED_CAR_RTO_EX_Annexure3 lacks 99 MP RTOs; the seven-makes group lacks 7 models",
      CAR.remarks.str.contains("lacks 99 RTO", regex=False).any() and CAR.remarks.str.contains("does not hold the models of the makes", regex=False).any())

# ================================================================ the layout changes this circular brought (each must be tolerated, none silently)
tbl = copy.deepcopy(TABLES)
for tb in tbl:
    for c in tb.cells:
        if "Maximum proposed" in c.text:
            c.text = "Revised proposal"
try:
    mod.build("fake.pdf", F, T, quiet, mod.Asker(False), tables=tbl)
    check("a new heading it does not know ('Revised proposal') for the new column stops with a clear message, it does not write the Existing column", False, "no stop")
except mod.GridError as e:
    check("a new heading it does not know ('Revised proposal') for the new column stops with a clear message (the table, the heading it expected, what it found), it does not write the Existing column",
          "NEW payout column" in str(e) and "expected" in str(e) and "found" in str(e), e)

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
