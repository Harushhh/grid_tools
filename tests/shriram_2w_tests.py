"""Deterministic checks of shriram_2w_to_config.py: the readers of the payout / age / body / condition cells, the RTO-group lookup by content, small
grids (held in memory) through build(), and the real October 2026 two-wheeler grid.
Run:  .venv\\Scripts\\python.exe tests\\shriram_2w_tests.py"""
import datetime as dt
import os
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("shriram_2w_to_config.py")
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


ask = mod.Asker(interactive=False)
M = mod.load_masters(ask)

# ================================================================ cell readers
check("NPO: plain 25", mod.parse_npo(25) == ("plain", 25.0) and mod.parse_npo("32.5") == ("plain", 32.5))
check("NPO: '35 ON OD' is On OD 35", mod.parse_npo("35 ON OD") == ("od", 35.0))
check("NPO: IRDA is the IRDA cap", mod.parse_npo("IRDA") == ("irda", mod.IRDA_RATE))
check("NPO: 'Rs. 250' / 'Rs.  150' are fixed rupee amounts", mod.parse_npo("Rs. 250") == ("flat", 250.0) and mod.parse_npo("Rs.  150") == ("flat", 150.0))
check("NPO: WITHDRAW and free text are not a payout", mod.parse_npo("WITHDRAW") is None and mod.parse_npo("as per grid") is None)
check("age: '6 to 15' = ages 6..15 = 5.01 - 15.01", mod.parse_age("6 to 15") == (5.01, 15.01))
check("age: unreadable is None", mod.parse_age("NEW") is None)
check("body: BODY TYPE / POLICY TYPE are swapped in the New Business sheet", mod.body_of("BUNDLE", "BIKE") == ("BIKE", "BUNDLE") and mod.body_of("BUNDLE", "EV") == ("EV", "BUNDLE")
      and mod.body_of("BUNDLE", "SCOOTY/MOPED") == ("SCOOTER", "BUNDLE") and mod.body_of("SCOOTER/SCOOTY") == ("SCOOTER", None))
inc, exc, un = mod.scope_of("HR38/HR51/HR55 are excluded")
check("condition: 'HR38/HR51/HR55 are excluded'", inc is None and exc == {"HR38", "HR51", "HR55"} and not un, (inc, exc, un))
inc, exc, un = mod.scope_of("Excluding OD-01,13,15,22,33,34,36")
check("condition: Odisha exclusions also take the OR twin", "OD13" in exc and "OR13" in exc and len(exc) == 14 and inc is None, exc)
inc, exc, un = mod.scope_of("MP-04/09/10/12/15/19/37/38/40/42")
check("condition: a bare code list is an inclusion", inc == {f"MP{n}" for n in "04 09 10 12 15 19 37 38 40 42".split()} and not exc, inc)
inc, exc, un = mod.scope_of("UP - 14, 16, 37, HR -26, 38, 51, 55")
check("condition: the NCR list (two states) is an inclusion", inc == {"UP14", "UP16", "UP37", "HR26", "HR38", "HR51", "HR55"}, inc)
inc, exc, un = mod.scope_of("(EXCLUDING NCR RTO - UP 14,16,37, UP-75,82,84)")
check("condition: '(EXCLUDING NCR RTO - UP 14,16,37, UP-75,82,84)'", exc == {"UP14", "UP16", "UP37", "UP75", "UP82", "UP84"}, exc)
inc, exc, un = mod.scope_of("Excluding MH-48 & MH-01")
check("condition: 'Excluding MH-48 & MH-01'", exc == {"MH48", "MH01"}, exc)
check("condition: 'DL- ALL RTOs', 'UPTO 150', '180CC Vehicle only' need no action", all(mod.scope_of(x) == (None, set(), "") for x in
      ("DL- ALL RTOs", "UPTO 150", "ABOVE 150", "Less than or equal to 180CC Vehicle only", "")))
check("condition: text with no RTO code is returned as unread (it gets flagged)", mod.scope_of("only for cases booked before 15th")[2] != "")
flags = []
check("makes: grid names -> the master's group names", mod.makes_of("RE", flags) == "ROYAL" and mod.makes_of("ATHER, BAJAJ, HERO, HONDA, PIAGGIO", flags) == "ATHER,BAJAJ,HERO,HONDA,PIAGGIO"
      and mod.makes_of("ATHER,SIMPLE ENERGY", flags) == "ATHER,SIMPLE_ENERGY" and mod.makes_of("AMPERE,BAJAJ AUTO,OLA", flags) == "AMPERE,BAJAJ,OLA" and not flags, flags)
mod.makes_of("ZERO MOTORCYCLES", flags)
check("makes: an unknown make is written as typed and flagged", len(flags) == 1 and "ZERO MOTORCYCLES" in flags[0], flags)

# ================================================================ RTO lookup
w = []
ctx = mod.Ctx(w, ("01/10/2026", "31/03/2030"), mod.read_declines(None, ask, w), M)
rto = lambda label, inc=None, exc=None: mod.rto_for(label, inc, exc, ctx)
check("RTO: West Bengal is its state code", rto("WEST BENGAL") == ("WB", []))
check("RTO: Delhi is DL", rto("DELHI") == ("DL", []))
check("RTO: AP/Telangana", rto("AP/TELANGANA") == ("AP,TG,TS", []))
check("RTO: Odisha is OD,OR", rto("ODISHA") == ("OD,OR", []))
check("RTO: Odisha minus its listed RTOs is the master group", rto("ODISHA", None, mod.scope_of("Excluding OD-01,13,15,22,33,34,36")[1]) == ("OD_OR_EX_01_13_15_22_33_34_36", []))
inc = mod.scope_of("UP - 14, 16, 37, HR -26, 38, 51, 55")[0]
check("RTO: NCR limited to its seven RTOs is the NCR group of the master", rto("NCR", inc) == ("SHRIRAM_NCR_HR26_MAY26", []), rto("NCR", inc))
check("RTO: Delhi/NCR without HR 38/51/55", rto("DELHI/NCR", None, {"HR38", "HR51", "HR55"}) == ("SHRIRAM_JUNE26_Delhi_NCR_HR_38_51_55", []))
got = rto("MAHARSHTRA", None, {"MH01", "MH48"})
check("RTO: Maharashtra minus MH01 / MH48 = Mumbai without them + the rest of Maharashtra", got == ("SHRIRAM_JUNE26_Mumbai_EX_MH_01_48,SHRIRAM_JUNE26_ROM", []), got)
check("RTO: MP loses the pan-India declined MP 06 / 07 / 30", rto("MP")[0] in ("SHRIRAM_MP_TW_MAY26", "MP_EX_06_07_30") and rto("MP")[1] == [], rto("MP"))
got = rto("RAJASTHAN")
check("RTO: Rajasthan minus RJ05 (pan-India declined) has no master group: whole state written, gap named", got[0] == "RJ" and got[1] and "RJ" in got[1][0], got)
check("RTO: an unknown state label is None (the row is listed as not written)", rto("ATLANTIS") is None)
got = rto("All except(DL,NCR, MP,CG,KL,TN,KA,HR,AR,MN,ML,MZ,NL,SK)")
toks = got[0].split(",")
check("RTO: 'All except(...)' names the other states, not the excluded ones", {"AP", "GJ", "WB", "RJ"} <= set(toks) and not ({"DL", "MP", "CG", "KL", "TN", "KA", "HR", "NL", "ML", "AR", "SK", "MZ", "MN"} & set(toks)), got)
check("RTO: 'All except' with NCR takes UP 14 / 16 / 37 out of UP (a UP group, not UP)", "UP" not in toks and any(t.startswith("SHRIRAM_UP_TW") or t.startswith("UP_") for t in toks), toks)

# ================================================================ small grids through build()
NB_HDR = ["STATE", "MANUFACTURER", "BODY TYPE", "POLICY TYPE", "CONDITION", "DIS %", "Average Net PO\n(Converted)", "UW'S CONDITION", "CHANGE DATE"]
RO_HDR = ["PRODUCT CATEGORY", "RTO STATE", "BODY TYPE", "MANUFACTURER", "CC", "AGE", "DISCOUNT (%)", "NPO(%) / FIXED", "POLICY TYPE", "CONDITION", "BLOCKED CONDITION"]
RO = lambda state, body, make, npo, ptype, dis="", cc="-", cond="": ["Two Wheeler ( 1 +1)", state, body, make, cc, "6 to 15", dis, npo, ptype, cond, ""]


def wb_of(nb=(), ro=()):
    sheets = {}
    if nb:
        sheets["New Business - 2W"] = [NB_HDR] + list(nb)
    if ro:
        sheets["ROLLOVER 2W (1+1)"] = [["GENERAL SLAB FOR BROKER TW (1+1) AGE 6 TO 15 YRS"] + [""] * 10, RO_HDR] + list(ro)
    return helpers.FakeWB(sheets)


def run(wb):
    return mod.build("x.xlsx", F, T, log=quiet, ask=mod.Asker(interactive=False), wb=wb)


NB = lambda state, make, body, npo, dis=50, cond="", uw="", change="": [state, make, "BUNDLE", body, cond, dis, npo, uw, change]
df, w = run(wb_of(nb=[NB("BIHAR", "HERO,HONDA,SUZUKI,BAJAJ", "BIKE", 45, 50, "Less than or equal to 180CC Vehicle only"),
                      NB("BIHAR", "ATHER, BAJAJ, HERO, HONDA, PIAGGIO", "EV", 40, 45),
                      NB("BIHAR", "HERO,HONDA,SUZUKI", "SCOOTY/MOPED", 45, 50, "", "", "Payout reduced w.e.f 1st Oct"),
                      NB("HIMACHAL", "BAJAJ", "BIKE", "WITHDRAW", 40, "Less than or equal to 180CC Vehicle only", "", "Grid withdrawn w.e.f 1st Oct"),
                      NB("DELHI/NCR", "BAJAJ", "BIKE", 0, 0, "Less than or equal to 180CC Vehicle only", "HR38/HR51/HR55 are excluded")]))
check("NB: bike, EV (two classes), scooty, zero-payout rows = 1 + 2 + 1 + 1 rows; the withdrawn row is not written", len(df) == 5 and set(df.sub_product) == {"1+5"} and set(df.policy_type) == {"New"}, len(df))
b = df[(df.make_model_class == "BIKE") & (df.fuel_type == "Petrol") & (df.new_rto_list == "BR")].iloc[0]
check("NB: bike <=180cc: Petrol, cc 0-180.01, age 0-99, On Net 45, tariff_max 50.01, po 38", (b.cc_min, b.cc_max, b.vehicle_age_min, b.vehicle_age_max) == (0, 180.01, 0, 99)
      and (b.pi_type, b.pi_net_rate, b.tariff_max, b.tariff_min, b.po_net_rate, b.po_type) == ("On Net", 45, 50.01, 0, 38, "On Net"), b.to_dict())
e = df[df.fuel_type == "Electric"]
check("NB: EV is written once per class, Electric, cc 0-99999", sorted(e.make_model_class) == ["BIKE", "SCOOTER"] and set(e.cc_max) == {99999}, e[["make_model_class", "cc_max"]].values.tolist())
z = df[df.new_rto_list == "SHRIRAM_JUNE26_Delhi_NCR_HR_38_51_55"].iloc[0]
check("NB: DIS 0 with payout 0 = tariff 0-100.01, payout 0", (z.pi_net_rate, z.po_net_rate, z.tariff_max) == (0, 0, 100.01), z.to_dict())
check("NB: 'Payout reduced w.e.f' needs no flag", not df.remarks.str.contains("Payout reduced").any())
check("NB: the withdrawn row is listed as not written", any(r["reason"].startswith("withdrawn") for r in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN)
check("NB: every row has the constant columns of the portal's export", set(df.insurer_vertical) == {"Not Applicable"} and set(df["product"]) == {"TW"} and set(df.user_id) == {2966}
      and set(df.sc_max) == {999} and set(df.is_ncb) == {"NA"})

df, w = run(wb_of(ro=[RO("BIHAR", "BIKE", "ALL EXCL. TVS", 32.5, "P", 40), RO("BIHAR", "BIKE", "ALL EXCL. TVS", 25, "L"),
                      RO("HARYANA", "BIKE", "ALL EXCL. TVS", "IRDA", "P", 40, "-", "(EXCLUDING NCR RTO HR 26, HR 38, HR 51, HR 55)"),
                      RO("HARYANA", "BIKE", "ALL EXCL. TVS", "IRDA", "L", "", "-", "(EXCLUDING NCR RTO HR 26, HR 38, HR 51, HR 55)"),
                      RO("HARYANA", "EV", "ALL EXCL. TVS", "35 ON OD", "P", 40, "-", "(EXCLUDING NCR RTO HR 26, HR 38, HR 51, HR 55)"),
                      RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 250", "P", 50), RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 150", "L"),
                      RO("BIHAR", "SCOOTER/SCOOTY", "TVS", 25, "P", 40), RO("BIHAR", "SCOOTER/SCOOTY", "TVS", 25, "L"),
                      RO("WEST BENGAL", "BIKE", "ALL EXCL. TVS", 35, "P", 50, "UPTO 150", "UPTO 150"), RO("WEST BENGAL", "BIKE", "ALL EXCL. TVS", 25, "P", 50, "ABOVE 150", "ABOVE 150")]))
check("rollover: 1+1 for P, STP for L, both Renew; EV doubles; IRDA on STP is not written, Rs. amounts are flat cards", sorted(df.sub_product.unique()) == ["1+1", "STP"] and set(df.policy_type) == {"Renew"}
      and len(df) == 11, len(df))                                      # BR bike P+L, HR bike P (IRDA), HR EV P x2 classes, AP bike P+L (Rs.), TVS scooter P+L, WB x2 cc bands
p = df[(df.new_rto_list == "BR") & (df.sub_product == "1+1") & (df.new_vehicle_makes == "TW_ALL_EX_TVS")].iloc[0]
check("rollover P: On Net 32.5, tariff_max 40.01, age 5.01-15.01, po 25.5, bike T&C", (p.pi_type, p.pi_net_rate, p.tariff_max, p.vehicle_age_min, p.vehicle_age_max, p.po_net_rate) == ("On Net", 32.5, 40.01, 5.01, 15.01, 25.5)
      and p.add_tnc == mod.TNC_BIKE, p.to_dict())
s = df[(df.new_rto_list == "BR") & (df.sub_product == "STP") & (df.new_vehicle_makes == "TW_ALL_EX_TVS")].iloc[0]
check("rollover L: On TP 25, tariff 0-100.01, po 18", (s.pi_type, s.pi_tp_rate, s.pi_net_rate, s.tariff_min, s.tariff_max, s.po_tp_rate) == ("On TP", 25, 0, 0, 100.01, 18), s.to_dict())
h = df[(df.sub_product == "1+1") & (df.fuel_type == "Petrol") & df.new_rto_list.str.contains("HR")].iloc[0]
check("rollover IRDA on P = On OD 15, po 8; Haryana without the NCR RTOs is a master group", (h.pi_type, h.pi_od_rate, h.po_od_rate) == ("On OD", 15, 8) and "HR" in h.new_rto_list, h.to_dict())
o = df[(df.fuel_type == "Electric") & (df.sub_product == "1+1")]
check("rollover '35 ON OD' = On OD 35 (po 28), once per EV class", len(o) == 2 and set(o.pi_type) == {"On OD"} and set(o.pi_od_rate) == {35} and set(o.po_od_rate) == {28})
t = df[df.new_vehicle_makes == "TVS_EX_APACHE"]
check("rollover TVS = the master group TVS without Apache; scooters carry no T&C", len(t) == 2 and set(t.make_model_class) == {"SCOOTER"} and set(t.add_tnc) == {""}, t[["make_model_class", "add_tnc"]].values.tolist())
wbk = df[df.new_rto_list == "WB"]
check("rollover cc UPTO 150 / ABOVE 150 -> 0-150.01 / 150.01-99999", sorted(zip(wbk.cc_min, wbk.cc_max)) == [(0, 150.01), (150.01, 99999)], wbk[["cc_min", "cc_max"]].values.tolist())
nw = pd.DataFrame(mod.LAST_NOT_WRITTEN)
check("rollover: the only row not written is the IRDA on STP one", len(nw) == 1 and nw.reason.str.contains("IRDA on STP").all(), nw.reason.tolist())
fp = df[(df.new_rto_list == "AP") & (df.sub_product == "1+1")].iloc[0]
fs = df[(df.new_rto_list == "AP") & (df.sub_product == "STP")].iloc[0]
check("rollover 'Rs. 250' on P = a flat card: On Net, rates 0, pi_flat_amount 250, po_flat_amount 175 (70%), tariff_max 50.01", (fp.pi_type, fp.pi_net_rate, fp.pi_od_rate, fp.pi_tp_rate, fp.pi_flat_amount, fp.po_flat_amount, fp.po_type, fp.tariff_max)
      == ("On Net", 0, 0, 0, 250, 175, "On Net", 50.01), fp.to_dict())
check("rollover 'Rs. 150' on L = a flat card: On TP, rates 0, flat 150 -> po 105, tariff 0-100.01", (fs.pi_type, fs.pi_tp_rate, fs.pi_net_rate, fs.pi_flat_amount, fs.po_flat_amount, fs.tariff_min, fs.tariff_max) == ("On TP", 0, 0, 150, 105, 0, 100.01), fs.to_dict())
check("flat cards say so in the remarks", df[df.pi_flat_amount > 0].remarks.str.contains("fixed rupee payout").all() and not df[df.pi_flat_amount == 0].remarks.str.contains("fixed rupee payout").any())
mod.WRITE_FLAT_AMOUNTS = False
d2, w2 = run(wb_of(ro=[RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 250", "P", 50), RO("BIHAR", "BIKE", "ALL EXCL. TVS", 25, "P", 40)]))
check("WRITE_FLAT_AMOUNTS = False: the Rs. row is listed as not written instead", len(d2) == 1 and any("fixed rupee" in r["reason"] for r in mod.LAST_NOT_WRITTEN))
mod.WRITE_FLAT_AMOUNTS = True
mod.FLAT_PI_TYPES["pkg"], saved = "On OD", mod.FLAT_PI_TYPES["pkg"]
d3, w3 = run(wb_of(ro=[RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 250", "P", 50)]))
check("the pi_type of a flat card comes from FLAT_PI_TYPES (SAOD = On OD, 1+3 = On Net are there too)", d3.pi_type.iloc[0] == "On OD" and mod.FLAT_PI_TYPES["saod"] == "On OD" and mod.FLAT_PI_TYPES["b13"] == "On Net")
mod.FLAT_PI_TYPES["pkg"] = saved
df, w = run(wb_of(ro=[RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 250", "P", 50), RO("ANDHRA PRADESH", "BIKE", "ALL EXCL. TVS", "Rs. 300", "P", 50)]))
check("two flat cards of one key with different amounts: the LOWER amount is written (250, not 300), and the choice is listed", len(df) == 1 and df.pi_flat_amount.iloc[0] == 250
      and len(mod.LAST_LOWEST) == 1 and mod.LAST_LOWEST[0]["written"] == "250" and any("LOWEST" in x for x in w), (len(df), mod.LAST_LOWEST))
df, w = run(wb_of(ro=[RO("BIHAR", "BIKE", "ALL EXCL. TVS", 32.5, "P", 40)]))
check("rollover: the make group the master lacks is flagged on its rows", df.remarks.str.contains("TW_ALL_EX_TVS").sum() == (df.new_vehicle_makes == "TW_ALL_EX_TVS").sum() > 0)

df, w = run(wb_of(nb=[NB("BIHAR", "HERO", "BIKE", 5, 50), NB("BIHAR", "HERO", "BIKE", 5, 50)]))
check("two identical grid rows are one upload row", len(df) == 1)
df, w = run(wb_of(nb=[NB("BIHAR", "HERO", "BIKE", 5, 50), NB("BIHAR", "HERO", "BIKE", 9, 50)]))
check("two grid rows with different payouts for one key: only the LOWEST is written (5, not 9)", len(df) == 1 and df.pi_net_rate.iloc[0] == 5 and mod.LAST_LOWEST[0]["rates_found"] in ("5 / 9", "9 / 5"), mod.LAST_LOWEST)
df, w = run(wb_of(nb=[NB("BIHAR", "HERO", "BIKE", 5, 50, "", "some new rule nobody told the script")]))
check("an unreadable condition is flagged, not guessed", len(df) == 1 and "some new rule nobody told the script" in df.remarks.iloc[0])
df, w = run(wb_of(nb=[NB("ATLANTIS", "HERO", "BIKE", 5, 50), NB("BIHAR", "HERO", "BIKE", 5, 50)]))
check("an unknown state is not written and is listed", len(df) == 1 and any(r["state"] == "ATLANTIS" for r in mod.LAST_NOT_WRITTEN))
df, w = run(wb_of(nb=[NB("BIHAR", "HERO", "BIKE", 5, 50, change="Payout reduced w.e.f 1st Dec")]))
check("a note that names a month after from_date is warned about", any("after from_date" in x for x in w), w)
try:
    run(helpers.FakeWB({"Sheet1": [["a", "b"], [1, 2]]}))
    check("a workbook without the two sheets stops with a clear message", False)
except mod.GridError as e:
    check("a workbook without the two sheets stops with a clear message", "Shriram two-wheeler grid" in str(e), str(e))

# ================================================================ the real October 2026 grid
REAL = os.path.join(helpers.INPUTS, "shriram_2w_oct26.xlsx")
BROKER = os.path.join(helpers.INPUTS, "shriram_2w_oct26_broker-grid.xlsx")
if os.path.exists(REAL):
    df, w = mod.build(REAL, F, T, log=quiet, ask=ask, broker_path=BROKER if os.path.exists(BROKER) else None)
    check("October grid: 543 rows (235 x 1+1, 99 x 1+5, 209 x STP)", len(df) == 543 and df.groupby("sub_product").size().to_dict() == {"1+1": 235, "1+5": 99, "STP": 209}, df.groupby("sub_product").size().to_dict())
    flat = df[df.pi_flat_amount > 0]
    check("October grid: the 10 Rs. grid rows are 14 flat cards (EV twice): STP On TP, 1+1 On Net, rates 0, payout amount = 70% of the payin amount",
          len(flat) == 14 and set(flat[flat.sub_product == "STP"].pi_type) == {"On TP"} and set(flat[flat.sub_product == "1+1"].pi_type) == {"On Net"}
          and (flat[["pi_od_rate", "pi_tp_rate", "pi_net_rate", "po_od_rate", "po_tp_rate", "po_net_rate"]] == 0).all().all() and ((flat.pi_flat_amount * 0.7 - flat.po_flat_amount).abs() < 1e-9).all()
          and sorted(flat.pi_flat_amount.unique()) == [100, 150, 180, 200, 250], (len(flat), sorted(flat.pi_flat_amount.unique())))
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
    po = df[["po_od_rate", "po_tp_rate", "po_net_rate"]].max(axis=1)
    check("October grid: po = max(pi - 7, 0) on every row", ((pi - 7).clip(lower=0) - po).abs().max() < 1e-9)
    check("October grid: dates 01/10/2026 - 31/03/2030", set(df.from_date) == {"01/10/2026"} and set(df.to_date) == {"31/03/2030"})
    check("October grid: EV rows are Electric, the others Petrol; classes BIKE / SCOOTER only", set(df.fuel_type) == {"Petrol", "Electric"} and set(df.make_model_class) == {"BIKE", "SCOOTER"})
    check("October grid: rollover age 5.01-15.01, new business 0-99", set(zip(df[df.sub_product != "1+5"].vehicle_age_min, df[df.sub_product != "1+5"].vehicle_age_max)) == {(5.01, 15.01)}
          and set(zip(df[df.sub_product == "1+5"].vehicle_age_min, df[df.sub_product == "1+5"].vehicle_age_max)) == {(0, 99)})
    check("October grid: STP rows are On TP with tariff 0-100.01", set(df[df.sub_product == "STP"].pi_type) == {"On TP"} and set(df[df.sub_product == "STP"].tariff_max) == {100.01})
    check("October grid: no key repeats", not df.duplicated(mod.KEY_COLUMNS).any())
    check("October grid: Andhra Pradesh has ALL EXCL. TVS rows (flat cards) and TVS rows",
          ((df.new_rto_list == "AP") & (df.new_vehicle_makes == "TW_ALL_EX_TVS")).any() and ((df.new_rto_list == "AP") & (df.new_vehicle_makes == "TVS_EX_APACHE")).any())
    check("October grid: 8 grid rows gave no upload row (7 IRDA on STP, 1 withdrawn)", len(mod.LAST_NOT_WRITTEN) == 8, len(mod.LAST_NOT_WRITTEN))
    check("October grid: the SAOD sheet is mentioned, not converted", any("TW SAOD" in x for x in w) and not (df.sub_product == "SAOD").any())
    out = os.path.join(tempfile.mkdtemp(prefix="shriram2w_t_"), "out.csv")
    mod.main(REAL, F, T, out_path=out, ask=mod.Asker(interactive=False), log=quiet, broker_path=BROKER if os.path.exists(BROKER) else None)
    problems = helpers.check_csv(out, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
    check("October grid: the written file has the 42 columns, numbers only in numeric columns, DD/MM/YYYY dates", problems == [], problems)
    rev = mod.review_path(out)
    book = pd.read_excel(rev, sheet_name=None, dtype=str, keep_default_na=False) if os.path.exists(rev) else {}
    check("October grid: ONE review workbook next to the upload CSV, a sheet each for the summary, the RTO groups, the make groups, the rows not written, the manual-review notes",
          {"Summary", "RTO groups to create", "Make groups to create", "Not written", "Manual review notes"} <= set(book) and os.path.exists(out), sorted(book))
    check("October grid: the sheet 'Make groups to create' names TW_ALL_EX_TVS with the entries to put in it; the RTO sheet has group, rows flagged and RTO codes",
          "TW_ALL_EX_TVS" in set(book.get("Make groups to create", pd.DataFrame({"group_needed": []})).group_needed) and list(book["RTO groups to create"].columns) == ["group_needed", "rows_flagged", "items"])
    nw = book.get("Not written", pd.DataFrame())
    check("October grid: the sheet 'Not written' lists the grid rows with sheet, grid row, payout and reason", len(nw) == 8 and {"sheet", "grid_row", "payout", "reason"} <= set(nw.columns))
    check("October grid: the Summary sheet says which file to upload, the row count and the warnings", "Upload THIS file" in " ".join(book["Summary"]["item"]) and "Warning" in set(book["Summary"]["item"]) and any(r.startswith("543") for r in book["Summary"]["value"]))
    check("October grid: no separate side files are written any more (the folder holds the CSV and the one workbook)", sorted(os.listdir(os.path.dirname(out))) == sorted(["out.csv", os.path.basename(rev)]), os.listdir(os.path.dirname(out)))
else:
    print("(real October grid not found - run learn.py case first)")

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
