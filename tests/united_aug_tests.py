"""Deterministic checks of united_commission_to_config.py on the real August 2026 United PDF ('Revised Pay-out' headings, a Private Car table by cc x states with
only a revised commission, Sub Annexure 2 = 31 Rajasthan RTOs for GCV, Sub Annexure 3 = the 40% Private Car RTOs, a scanned cover page).
Run:  .venv\\Scripts\\python.exe tests\\united_aug_tests.py"""
import copy
import datetime as dt
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("united_commission_to_config.py")
PDF = os.path.join(helpers.INPUTS, "united_aug26.pdf")
F, T = dt.date(2026, 8, 1), dt.date(2026, 8, 31)
quiet = lambda *a, **k: None
TABLES = mod.read_pdf(PDF)
U = mod.STATE_CODES
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def build(tbls=None, from_date=F, to_date=T):
    return mod.build("fake.pdf", from_date, to_date, quiet, mod.Asker(False), tables=copy.deepcopy(TABLES) if tbls is None else tbls)


def sel(df, product=None, cls=None, sub=None, fuel=None, cc=None, rto=None):
    d = df
    for col, val in (("product", product), ("make_model_class", cls), ("sub_product", sub), ("fuel_type", fuel)):
        if val is not None:
            d = d[d[col] == val]
    if cc is not None:
        d = d[(d.cc_min.astype(float) == cc[0]) & (d.cc_max.astype(float) == cc[1])]
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
    gone = {n for x in names for n in x.split(",")}
    return ",".join(c for c in U if c not in gone)


df, warns = build()
CAR = df[df["product"] == "Private Car"]
G4 = df[df["product"] == "GCV 4W"]

# ================================================================ the document
check("the August PDF builds: 511 rows (414 + 97 Private Car), 42 columns, no duplicate key, every row from 01/08/2026 to 31/08/2026",
      len(df) == 511 and list(df.columns) == mod.OUTPUT_COLUMNS and not df.duplicated(mod.KEY_COLUMNS).any()
      and set(df.from_date) == {"01/08/2026"} and set(df.to_date) == {"31/08/2026"}, len(df))
check("the new column is the 'Revised Pay-out' one (effective 01-08-2026, existing 01-07-2026): payouts were found in all 5 sections and the Private Car table",
      len(df[df["product"] != "Private Car"]) == 414 and len(CAR) == 97, (len(df), len(CAR)))
check("the scanned cover letter (page 1, no table, an image) is reported as not read, once",
      sum("page 1 has no table" in w and "scanned image" in w for w in warns) == 1, warns)
check("circular_groups: the July constants for 01/07/2026 (and for a PDF without a date), the August groups for 01/08/2026, nothing for another date",
      mod.circular_groups(dt.date(2026, 7, 1))["car"] == (mod.CAR_RTO_ANNEX4, mod.CAR_RTO_EX_ANNEX4) and mod.circular_groups(None) == mod.circular_groups(dt.date(2026, 7, 1))
      and mod.circular_groups(dt.date(2026, 8, 1))["car"] == ("UNITED_CAR_RTO_Annexure3", "UNITED_CAR_RTO_EX_Annexure3")
      and mod.circular_groups(dt.date(2026, 8, 1))["gcv"] == {(2000, 3500): ("UNITED_RJ_GCV4W_2T-3.5T_Annexure-2", "UNITED_RJ_EX_Annexure2_AUG26")}
      and mod.circular_groups(dt.date(2026, 9, 1)) is None)

# ================================================================ Goods Carrying (changed in August)
b = (0, 2000.01)
check("GVW <= 2000: Uttar Pradesh 50 / 43, 'Other than above State' (singular) 60 / 53 on the other 39 state codes - it was 57.5 for All States in July",
      rate(sel(G4, rto="UP", cc=b)) == {(50.0, 43.0)} and rate(sel(G4, rto=comp("UP"), cc=b)) == {(60.0, 53.0)} and len(sel(G4, cc=b)) == 20)
b = (2000.01, 3500.01)
check("2000-3500: HR, MP, TN, UP 15 / 8; KA, KL 45 / 38; every other state 56.5 / 49.5",
      rate(sel(G4, rto="HR,MP,TN,UP", cc=b)) == {(15.0, 8.0)} and rate(sel(G4, rto="KA,KL", cc=b)) == {(45.0, 38.0)}
      and rate(sel(G4, rto=comp("HR", "MP", "TN", "UP", "RJ", "KA", "KL"), cc=b)) == {(56.5, 49.5)})
check("2000-3500 Rajasthan '(Except RTOs listed in Sub Annexure 2 where commission would be 56.5%)': the group of the other Rajasthan RTOs 35 / 28, and the Sub Annexure 2 RTOs group 56.5 / 49.5 "
      "(10 rows each: 5 fuels x 1+1 / STP)",
      rate(sel(G4, rto="UNITED_RJ_EX_Annexure2_AUG26", cc=b)) == {(35.0, 28.0)} and len(sel(G4, rto="UNITED_RJ_EX_Annexure2_AUG26", cc=b)) == 10
      and rate(sel(G4, rto="UNITED_RJ_GCV4W_2T-3.5T_Annexure-2", cc=b)) == {(56.5, 49.5)} and len(sel(G4, rto="UNITED_RJ_GCV4W_2T-3.5T_Annexure-2", cc=b)) == 10
      and not (sel(G4, cc=b).new_rto_list == "RJ").any())
b = (3500.01, 7500.01)
check("3500-7500: HR, MP, PY, RJ, TN, UP 10 / 3 with no annexure group any more (the 'Except RTOs listed' text is only in the Existing column), other states 27.5 / 20.5",
      rate(sel(G4, rto="HR,MP,PY,RJ,TN,UP", cc=b)) == {(10.0, 3.0)} and rate(sel(G4, rto=comp("HR", "MP", "PY", "RJ", "TN", "UP"), cc=b)) == {(27.5, 20.5)}
      and not sel(G4, cc=b).new_rto_list.str.contains("UNITED_").any())
check("GVW > 40000: the other states went from 5 to 2.5 (PO floors at 0); the Haryana ... group stays 0",
      rate(sel(G4, cc=(40000.01, 99999), rto=comp("HR", "MP", "RJ", "TN", "PY", "UP", "KA", "KL"))) == {(2.5, 0.0)}
      and rate(sel(G4, cc=(40000.01, 99999), rto="HR,KA,KL,MP,PY,RJ,TN,UP")) == {(0.0, 0.0)})
check("the other classes are the same as in July: Two Wheelers, passenger carrying, Miscellaneous (this file has no July output to compare - spot checks)",
      rate(sel(df, product="TW", sub="1+1", fuel="Petrol", cc=(0, 150.01), rto="AS,KA,KL,MP,PY,TN")) == {(5.0, 0.0)}
      and rate(sel(df, product="TW", sub="1+1", fuel="Petrol", cc=(0, 150.01), rto=comp("AS", "KA", "KL", "MP", "PY", "TN"))) == {(27.5, 20.5)}
      and rate(sel(df, product="MISCD", cls="MISCD-Tractor", sub="1+1")) == {(40.0, 33.0)})

# ================================================================ Private Car: cc x states, no make groups
check("every Private Car row uses the make group Private_Car_ALL (the July segments by make are gone) and the 5 cc bands of the table",
      set(CAR.new_vehicle_makes) == {"Private_Car_ALL"} and set(zip(CAR.cc_min.astype(float), CAR.cc_max.astype(float)))
      == {(0, 1000.01), (1000.01, 1200.01), (1200.01, 1500.01), (1500.01, 2000.01), (2000.01, 99999), (0, 99999), (1500.01, 99999)})
c13 = sel(CAR, sub="1+3")
check("Bundled (1+3), non-electric: <=1000 All 20 / 13; 1000-1200 KA, HR, MP, TN, PY 20 and the other 35 state codes 40 / 33; 1200-1500 All 20; 1500-2000 All 40; "
      ">2000 CG, HR, KA, MP, PY, RJ, TN, UP 15 / 8 and the other 32 codes 30 / 23 - four fuels each",
      rate(sel(c13, cc=(0, 1000.01), rto="ALLINDIA")) == {(20.0, 13.0)} and rate(sel(c13, cc=(1000.01, 1200.01), rto="HR,KA,MP,PY,TN")) == {(20.0, 13.0)}
      and rate(sel(c13, cc=(1000.01, 1200.01), rto=comp("HR", "KA", "MP", "PY", "TN"))) == {(40.0, 33.0)} and rate(sel(c13, cc=(1200.01, 1500.01), rto="ALLINDIA")) == {(20.0, 13.0)}
      and rate(sel(c13, cc=(1500.01, 2000.01), rto="ALLINDIA")) == {(40.0, 33.0)} and rate(sel(c13, cc=(2000.01, 99999), rto="CG,HR,KA,MP,PY,RJ,TN,UP")) == {(15.0, 8.0)}
      and rate(sel(c13, cc=(2000.01, 99999), rto=comp("CG", "HR", "KA", "MP", "PY", "RJ", "TN", "UP"))) == {(30.0, 23.0)}
      and set(c13[c13.fuel_type != "Electric"].fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"} and len(c13) == 4 * 7 + 1)
check("Package (1+1) and SATP (STP) share the second row group: <=1000 20, 1000-1200 20 / 35, 1200-1500 20, 1500-2000 35, >2000 15 / 30 (TP rate for STP)",
      rate(sel(CAR, sub="STP", fuel="Petrol", cc=(0, 1000.01))) == {(20.0, 13.0)} and rate(sel(CAR, sub="STP", fuel="Petrol", cc=(1000.01, 1200.01), rto=comp("HR", "KA", "MP", "PY", "TN"))) == {(35.0, 28.0)}
      and rate(sel(CAR, sub="STP", fuel="Petrol", cc=(1500.01, 2000.01))) == {(35.0, 28.0)} and rate(sel(CAR, sub="STP", fuel="Petrol", cc=(2000.01, 99999), rto="CG,HR,KA,MP,PY,RJ,TN,UP")) == {(15.0, 8.0)}
      and (sel(CAR, sub="STP").pi_tp_rate.astype(float) > 0).all() and (sel(CAR, sub="STP").pi_net_rate.astype(float) == 0).all())
check("SAOD non-electric: All Vehicles, ALLINDIA, 15 on OD for the four fuels; Electric: 35 for 1+3, 1+1, STP and SAOD",
      len(sel(CAR, sub="SAOD", fuel="Petrol")) == 1 and rate(sel(CAR, sub="SAOD", fuel="Petrol")) == {(15.0, 8.0)} and float(sel(CAR, sub="SAOD", fuel="Petrol").pi_od_rate.iloc[0]) == 15
      and rate(sel(CAR, fuel="Electric", sub="1+3")) == {(35.0, 28.0)} and rate(sel(CAR, fuel="Electric", sub="STP")) == {(35.0, 28.0)} and rate(sel(CAR, fuel="Electric", sub="SAOD")) == {(35.0, 28.0)}
      and float(sel(CAR, sub="SAOD", fuel="Electric").pi_od_rate.iloc[0]) == 35)
p1 = sel(CAR, sub="1+1")
check("Package (CAR_ANNEX4_POLICIES = 1+1) in the 40% RTOs: Petrol / CNG / LPG / Electric all cc and Diesel above 1500 cc at 40 / 33 on the annexure group; the annexure has no row for diesel up to 1500 cc",
      rate(sel(p1, rto="UNITED_CAR_RTO_Annexure3", cc=(0, 99999), fuel="Petrol")) == {(40.0, 33.0)} and set(sel(p1, rto="UNITED_CAR_RTO_Annexure3", cc=(0, 99999)).fuel_type) == {"Petrol", "CNG", "LPG", "Electric"}
      and set(sel(p1, rto="UNITED_CAR_RTO_Annexure3", cc=(1500.01, 99999)).fuel_type) == {"Diesel"} and len(sel(p1, rto="UNITED_CAR_RTO_Annexure3")) == 5)
check("Package rows for 'All' states go to the group of all the other RTOs (UNITED_CAR_RTO_EX_Annexure3) for every fuel the 40% takes over - Petrol / CNG / LPG at every cc, Electric, Diesel above 1500 cc - "
      "and to ALLINDIA for diesel up to 1500 cc (no 40% row there)",
      set(sel(p1, rto="UNITED_CAR_RTO_EX_Annexure3", cc=(0, 1000.01)).fuel_type) == {"Petrol", "CNG", "LPG"} and set(sel(p1, rto="ALLINDIA", cc=(0, 1000.01)).fuel_type) == {"Diesel"}
      and set(sel(p1, rto="UNITED_CAR_RTO_EX_Annexure3", cc=(1500.01, 2000.01)).fuel_type) == {"Petrol", "CNG", "LPG", "Diesel"} and set(sel(p1, rto="UNITED_CAR_RTO_EX_Annexure3", cc=(0, 99999)).fuel_type) == {"Electric"}
      and not CAR[CAR.sub_product != "1+1"].new_rto_list.str.contains("UNITED_").any())
state_rows = p1[~p1.new_rto_list.isin(["ALLINDIA", "UNITED_CAR_RTO_EX_Annexure3", "UNITED_CAR_RTO_Annexure3"])]
over = state_rows[state_rows.remarks.str.contains("covers the whole state(s)", regex=False)]
check("a Package row for some states cannot leave the 40% RTOs out (the master has no 'these states except Sub Annexure 3' group): it is written for the states and flagged - exactly the "
      "Petrol / CNG / LPG rows of 1000-1200 cc (2 state sets x 3 fuels) and all four fuels above 2000 cc (2 x 4) = 14 rows, and the note names the states that hold 40% RTOs",
      len(over) == 14 and set(zip(over.cc_min.astype(float), over.cc_max.astype(float))) == {(1000.01, 1200.01), (2000.01, 99999)} and "Diesel" not in set(over[over.cc_min.astype(float) == 1000.01].fuel_type)
      and over.remarks.str.contains("Sub Annexure 3 RTOs of KA, TN", regex=False).any() and len(state_rows) - len(over) == 2
      and set(state_rows[~state_rows.index.isin(over.index)].fuel_type) == {"Diesel"}, (len(over), len(state_rows)))
check("3+3 (Long Term Package) and Long term Liability have no sub_product: not written, each flagged once on every Private Car row (97 rows) and warned about",
      CAR.remarks.str.contains("Long Term Package (3+3) has no sub_product", regex=False).all() and CAR.remarks.str.contains("Long term Liability has no sub_product", regex=False).all()
      and sum("Long Term Package (3+3)" in w for w in warns) == 1 and sum("Long term Liability" in w for w in warns) == 1 and not df[df["product"] != "Private Car"].remarks.str.contains("3+3", regex=False).any())

# ================================================================ the master lists
check("master: the Sub Annexure 2 group holds 27 of the circular's 31 Rajasthan RTOs (13 missing, 9 more) -> flagged on its 10 rows; the group of the other Rajasthan RTOs is exact -> no flag",
      df[df.new_rto_list == "UNITED_RJ_GCV4W_2T-3.5T_Annexure-2"].remarks.str.contains("does not hold the preferred RTOs listed in Sub Annexure 2", regex=False).all()
      and not df[df.new_rto_list == "UNITED_RJ_EX_Annexure2_AUG26"].remarks.str.contains("does not hold", regex=False).any())
check("master: UNITED_CAR_RTO_Annexure3 lacks RJ06 RJ07 RJ19 RJ27 and has AP39 AP40 -> flagged on its 5 rows; UNITED_CAR_RTO_EX_Annexure3 lacks the 99 MP RTOs -> flagged on its rows",
      sel(df, rto="UNITED_CAR_RTO_Annexure3").remarks.str.contains("only in the circular ['RJ06', 'RJ07', 'RJ19', 'RJ27']", regex=False).all()
      and sel(df, rto="UNITED_CAR_RTO_EX_Annexure3").remarks.str.contains("lacks 99 RTO(s)", regex=False).all())

# ================================================================ the setting the user can change
old = mod.CAR_ANNEX4_POLICIES
mod.CAR_ANNEX4_POLICIES = ("1+1", "1+3", "STP")
try:
    dfx, _ = build()
finally:
    mod.CAR_ANNEX4_POLICIES = old
cx = dfx[dfx["product"] == "Private Car"]
check("CAR_ANNEX4_POLICIES = (1+1, 1+3, STP): the annexure group gets 40 / 33 rows for 1+3 and STP as well (5 + 5 + 5 rows), their 'All states' rows move to the other-RTOs group, SAOD stays on ALLINDIA",
      len(sel(cx, rto="UNITED_CAR_RTO_Annexure3")) == 15 and set(sel(cx, rto="UNITED_CAR_RTO_Annexure3").sub_product) == {"1+1", "1+3", "STP"}
      and rate(sel(cx, sub="1+3", rto="UNITED_CAR_RTO_Annexure3", cc=(0, 99999), fuel="Petrol")) == {(40.0, 33.0)} and set(sel(cx, sub="SAOD").new_rto_list) == {"ALLINDIA"}
      and "UNITED_CAR_RTO_EX_Annexure3" in set(sel(cx, sub="STP").new_rto_list) and not cx.duplicated(mod.KEY_COLUMNS).any())

# ================================================================ changes to the document
t = copy.deepcopy(TABLES)
for tb in t:
    if tb.page == 7:
        tb.cells = [c for c in tb.cells if c.text != "56.5%"]
d2, w2 = build(t)
check("Sub Annexure 2 without its commission cell: the 56.5% preferred-RTO rows are not written (-10), a warning says so, the rest is unchanged",
      len(d2) == len(df) - 10 and any("no commission % found" in w for w in w2) and len(d2[d2.new_rto_list == "UNITED_RJ_EX_Annexure2_AUG26"]) == 10)
t = copy.deepcopy(TABLES)
for tb in t:
    for c in tb.cells:
        if c.page == 6 and c.text == "Non Electric - SAOD":
            c.text = "Non Electric - Something New"
d3, w3 = build(t)
check("an unknown Private Car policy type: its rows are not written (-4), flagged on the Private Car rows, a warning says so; no question, no crash",
      len(d3[d3["product"] == "Private Car"]) == len(CAR) - 4 and any("Something New" in w for w in w3) and d3[d3["product"] == "Private Car"].remarks.str.contains("Something New", regex=False).all())
t = copy.deepcopy(TABLES)
for tb in t:
    for c in tb.cells:
        if c.page == 6 and c.text == ">2000 CC":
            c.text = ">2000 CC sedan"
d4, w4 = build(t)
check("an unknown cc segment ('>2000 CC sedan'): that segment is not written for either non-electric policy group, a warning says so",
      len(d4[d4["product"] == "Private Car"]) < len(CAR) and any("sedan" in w for w in w4))
d5, w5 = build(from_date=dt.date(2026, 9, 1), to_date=dt.date(2026, 9, 30))
check("typing 01/09/2026 against the PDF's 01-08-2026: every row says so in remarks, nothing else changes",
      d5.remarks.str.contains("effective from 01/08/2026", regex=False).all() and len(d5) == len(df))

bad = [n for n, ok, _ in results if not ok]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
