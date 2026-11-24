"""Deterministic checks of indusind_grid_to_config.py (IndusInd General = Reliance General, commercial-vehicle grid): every header of the real November grid read by its words against a
hand-checked table, cells of the real workbook against the rows written (read again with openpyxl, independently of the script), the real grid against the user's own export (11,304 of its rows
reproduced; the 34 diesel-taxi rows that follow the grid and not the export; Indore), the conditional readings (a qualifier, the make groups of a make label, the Car Carrier guard - only when the
wording / the number is in the grid), cells that are not payouts, the structure of the table (twin region columns, twin payout columns, banner, extra header row, two sheets), the notes (blocks,
regions a table is keyed by, standing / resolved notes), the RTO-group check, the dates and the run from the command line.
Run:  .venv\\Scripts\\python.exe tests\\indusind_tests.py"""
import datetime as dt
import json
import os
import subprocess
import sys
import tempfile

os.environ["GRID_RTO_NAMES"] = "cluster"                   # the script's own default (the user's rule of 30 Oct 2026): the suite tests the mode the user runs

import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
import pandas as pd  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402
import shiftlib as L  # noqa: E402                         # insert() moves merged ranges with the cells, as Excel does

mod = helpers.load_script("indusind_grid_to_config.py")
ROOT = helpers.ROOT
F, T = dt.date(2026, 11, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


SRC = os.path.join(helpers.INPUTS, "indusind_nov26.xlsx")
REF = os.path.join(helpers.REFS, "indusind_nov26_rates_export.csv")
ALL5, ND4 = ["Petrol", "Diesel", "CNG", "LPG", "Electric"], ["Petrol", "CNG", "LPG", "Electric"]
COL = openpyxl.utils.column_index_from_string


def real(data_only=True):
    return openpyxl.load_workbook(SRC, data_only=data_only)


def mini(keep=(5, 6, 7, 8, 9, 23), tweak=None):
    """the real workbook with only some of the 73 region rows (Mumbai, Nagpur, Pune, Ahmednagar and Aurangabad [both under the parent ROM], Indore) - every other sheet as it is"""
    wb = real()
    ws = wb["Nov'26"]
    for r in range(77, 4, -1):
        if r not in keep:
            ws.delete_rows(r)
    if tweak:
        tweak(wb)
    return wb


LAST = {}


def build(wb, names="cluster", ask=None, f=F, t=T, interactive=False):
    """(frame, warnings, asker): the frame is what build returns (the region names are still the script's group names: `bare` gives the cluster names)"""
    mod.RTO_NAMES = names
    try:
        a = ask or mod.Asker(interactive=interactive)
        df, w = mod.build(None, f, t, quiet, a, wb=wb)
    finally:
        mod.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")
    LAST["ask"], LAST["warnings"] = a, w
    return df, w, a


def bare(df):
    mod.RTO_NAMES = "cluster"
    return mod.apply_cluster_names(df)


def stops(fn, *words):
    """the message of the GridError / NeedsAnswer that fn() raises ('' when it does not stop)"""
    try:
        fn()
    except mod.GridError as e:
        return str(e)
    return ""


def sel(df, **kw):
    m = pd.Series(True, index=df.index)
    for k, v in kw.items():
        m &= (df[k] == v)
    return df[m]


def pi_of(df):
    return df[["pi_net_rate", "pi_tp_rate", "pi_od_rate"]].astype(float).max(axis=1)


def sheet(wb, name="Nov'26"):
    return wb[name]


def find_cell(ws, text):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and text.lower() in c.value.lower():
                return c
    raise AssertionError(f"no cell with {text!r}")


# ================================================================ the real grid, header by header
wb = real()
sh = mod.Sheet(wb["Nov'26"])
anchor = mod.find_anchor(sh)
tbl = mod.RateTable(sh, *anchor)
check("the header row is found by its words: 'RTO Region' twice (B and C) on row 4", anchor == (4, [2, 3]), anchor)
check("the table has 73 data rows, the first on Excel row 5, the header band is rows 2-4", tbl.ok and len(tbl.data) == 73 and tbl.first == 5 and tbl.top == 2, (tbl.first, tbl.top, len(tbl.data)))
tbl.pick_region(quiet, mod.Asker(interactive=False))
check("the cluster column is the one with 73 distinct names (C), the parent column (B) is not used", tbl.rcol == 3 and tbl.parent == 2, (tbl.rcol, tbl.parent))
good, refused, empty = mod.classify_columns(tbl)
by = {c["letter"]: specs for c, specs in good}
check("every one of the 31 payout columns D..AH is understood, none refused, none empty (AI / AJ have no header and no number)", len(good) == 31 and not refused and not empty and sorted(by, key=COL) == [openpyxl.utils.get_column_letter(i) for i in range(4, 35)],
      (len(good), refused, [c["letter"] for c in empty]))


def spec_of(letter, i=0):
    s = by[letter][i]
    return (s["product"], s["klass"], s["makes"], s["fuels"], s["subs"], s["zd"], tuple(s["age"]), tuple(s["cc"]), tuple(s["sc"]), s["tnc"]) if not s["deal"] else ("DEAL",)


TAXI, BUS = "PCV 4W", "PCV 4W"
EXPECT = {
    "D": ("PCV 4W", "School Bus", "BUS_ALL", ALL5, ["1+1", "STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "E": ("PCV 3W", "Passenger Carrying Rickshaw", "PCV_3W_ALL", ND4, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "F": ("PCV 3W", "Passenger Carrying Rickshaw", "PCV_3W_ALL", ND4, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "G": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["1+1"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "H": ("PCV 4W", "Taxi", "TAXI_ALL", ["Diesel"], ["1+1", "STP"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "I": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["STP"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "J": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["1+1"], "YES", (0, 99), (0, 99999), (0, 6.01), ""),
    "K": ("GCV 4W", "NA", "TATA,Maruti,Mahindra", ALL5, ["1+1"], "NA", (0, 99), (0, 2000.01), (0, 999), ""),
    "L": ("GCV 4W", "NA", "TATA,Maruti,Mahindra", ALL5, ["STP"], "NA", (0, 99), (0, 2000.01), (0, 999), ""),
    "M": ("GCV 4W", "NA", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", ALL5, ["1+1"], "NA", (0, 99), (2000.01, 2500.01), (0, 999), ""),
    "N": ("GCV 4W", "NA", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", ALL5, ["STP"], "NA", (0, 99), (2000.01, 2500.01), (0, 999), ""),
    "O": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["1+1"], "NA", (0, 99), (2500.01, 3500.01), (0, 999), ""),
    "P": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["STP"], "NA", (0, 99), (2500.01, 3500.01), (0, 999), ""),
    "Q": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["1+1"], "NA", (0, 99), (3500.01, 99999), (0, 999), ""),
    "R": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["STP"], "NA", (0, 99), (3500.01, 99999), (0, 999), ""),
    "S": ("GCV 3W", "Goods Carrying Rickshaw", "GCV_3W_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "T": ("GCV 3W", "Goods Carrying Rickshaw", "GCV_3W_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "U": ("PCV 4W", "Taxi", "TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG", ALL5, ["1+1"], "NO", (0, 99), (0, 99999), (6.01, 8.01), ""),
    "V": ("PCV 4W", "Taxi", "TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG", ALL5, ["STP"], "NO", (0, 99), (0, 99999), (6.01, 8.01), ""),
    "Y": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["1+1", "STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "Z": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), "Flat Bed"),
    "AA": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), "Flat Bed"),
    "AB": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), "JCB, L&T,Caterpillar"),
    "AC": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), "JCB, L&T,Caterpillar"),
    "AD": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 1), (0, 99999), (0, 999), "New,Agriculture Tractor"),
    "AE": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["1+1"], "NA", (1, 99), (0, 99999), (0, 999), "Old,Agriculture Tractor"),
    "AF": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
}
for letter, want in EXPECT.items():
    check(f"column {letter}: {by[letter][0]['segment'][:34]!r} / {by[letter][0]['chain'][-1][:22]!r} -> product, class, make group, fuels, policy types, nil dep, age, weight, seats, T&C", spec_of(letter) == want, (spec_of(letter), want))
check("column W holds TWO products: other buses over 8 seats (BUS_ALL, 8-999) and 7+1 taxis of the other makes (the master's EX group, 6.01-8.01); X the same for STP",
      [spec_of("W", i) for i in (0, 1)] == [("PCV 4W", "Other Bus", "BUS_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (8, 999), ""),
                                          ("PCV 4W", "Taxi", "TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (6.01, 8.01), "")]
      and [by["X"][i]["subs"] for i in (0, 1)] == [["STP"], ["STP"]], [spec_of("W", 0), spec_of("W", 1)])
check("columns AG / AH (Employee Pick Up, 'Only Deal Based') are read as deal based: no rows", by["AG"][0]["deal"] and by["AH"][0]["deal"])
check("the two make labels carry their model-level qualifier (K, L, M, N); no other column has one",
      sorted(l for l, specs in by.items() if any(s.get("qualifiers") for s in specs)) == ["K", "L", "M", "N"]
      and "Only Jeeto & Supro" in by["K"][0]["qualifiers"][0] and "other than Jeeto & Supro" in by["M"][0]["qualifiers"][0], [(l, s.get("qualifiers")) for l, specs in by.items() for s in specs if s.get("qualifiers")])
check("only the Car Carrier column (no T&C, MISCD-Others) is a guarded catch-all", [l for l, specs in by.items() if any(s.get("guard") for s in specs)] == ["Y"])

# ================================================================ the brain on wordings that are NOT in November's grid (a reusable, header-driven reader: the same words give the same meaning anywhere)
for text, want in ((" 2 - 2.5 k", (2000.01, 2500.01)), (" above 3.5k", (3500.01, 99999)), (" 7.5 t - 12 t", (7500.01, 12000.01)), (" 7.5 - 12 tn", (7500.01, 12000.01)), (" >12 t", (12000.01, 99999)),
                   (" 12t - 20t", (12000.01, 20000.01)), (" above 40 tons", (40000.01, 99999)), ("gcv < 2.0 k", (0, 2000.01)), (" upto 3.5 t", (0, 3500.01))):
    got = mod.band_of(text)
    check(f"weight band {text.strip()!r} -> {want}", got is not None and got[0] == want, got)
check("a text with no weight words has no band", mod.band_of(" gcv special") is None)
for text, want in (("<6 str", (0, 6.01)), ("7+1", (6.01, 8.01)), ("4+1", (3.01, 5.01)), ("6+1", (5.01, 7.01)), ("> 8 str", (8, 999)), ("<= 4 seater", (0, 4.01))):
    got = mod.seats_of(text)
    check(f"seats {text!r} -> {want}", got is not None and got[0] == want, got)
check("fuels: 'Petrol + CNG + Battery' = Petrol, CNG, LPG (CNG rides with LPG), Electric; 'Diesel'; 'All Fuel Type' = all five; 'Non diesel' = four; anything else is None",
      mod.fuels_of("Petrol + CNG + Battery") == ND4 and mod.fuels_of("Diesel") == ["Diesel"] and mod.fuels_of("All Fuel Type") == ALL5 and mod.fuels_of("Non diesel") == ND4 and mod.fuels_of("Petrol + Hydrogen") is None)
check("makes: a list in any separator, aliases (M&M, Maruti Suzuki, Ashok Leyland, AL), unknown words are returned as the rest",
      mod.split_makes("Maruti, Toyota & Mahindra") == (["maruti", "toyota", "mahindra"], []) and mod.split_makes("M&M / Ashok Leyland and AL") == (["mahindra", "ashok"], [])
      and mod.split_makes("Tata / Piaggio") == (["tata"], ["piaggio"]))


def chain_of(*texts):
    return mod.classify_chain(list(texts), {"makes": {}})


def spec1(*texts):
    s = chain_of(*texts)[0]
    return s["product"], s["klass"], s["makes"], s["fuels"], s["subs"], s["zd"], tuple(s["age"]), tuple(s["cc"]), tuple(s["sc"]), s["tnc"]


check("another tonnage band is read by its words: 'GCV 7.5T - 12T' / 'COMP' -> GCV 4W 7500.01-12000.01, every make, Comp = 1+1", spec1("GCV 7.5T - 12T", "COMP") == ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["1+1"], "NA", (0, 99), (7500.01, 12000.01), (0, 999), ""))
check("'GCV Above 12 T' / 'STP' -> the open band 12000.01-99999, STP", spec1("GCV Above 12 T", "STP")[7] == (12000.01, 99999) and spec1("GCV Above 12 T", "STP")[4] == ["STP"])
check("a make label of other makes: 'TATA / Ashok Leyland' -> the master's single-make groups 'TATA,Ashok'", spec1("TATA / Ashok Leyland", "GCV < 2.0 K", "COMP")[2] == "TATA,Ashok")
mod.MAKE_DEFINITIONS.clear()
check("'Others TATA / Eicher' -> GCV_4W_ALL_EX_TATA_EICHER, registered so the 'Make groups to create' sheet can say how it is made", spec1("Others TATA / Eicher", "GCV 2 - 2.5 K", "COMP")[2] == "GCV_4W_ALL_EX_TATA_EICHER"
      and mod.MAKE_DEFINITIONS["GCV_4W_ALL_EX_TATA_EICHER"] == ("GCV_4W_ALL", ["TATA", "EICHER"]))
mod.MAKE_DEFINITIONS.clear()
check("a 4+1 taxi with nil dep and petrol + CNG: seats 3.01-5.01, is_zd YES, Petrol / CNG / LPG", spec1("PCV Taxi 4+1 ( Only ND )", "Comp - Others ( Petrol + CNG )")[5:10] == ("YES", (0, 99), (0, 99999), (3.01, 5.01), "") and spec1("PCV Taxi 4+1 ( Only ND )", "Comp - Others ( Petrol + CNG )")[3] == ["Petrol", "CNG", "LPG"])
check("the same words in another ORDER of the chain (the leaf above the segment) give the same columns: the role of a text comes from its words, not from its level",
      spec1("COMP", "GCV < 2.0 K") == spec1("GCV < 2.0 K", "COMP") and spec1("TATA / Maruti", "COMP", "GCV < 2.0 K") == spec1("GCV < 2.0 K", "TATA / Maruti", "COMP"))
check("a make label with no brackets has no qualifier; with another wording ('( Only Jeeto )') it has one; 'Below 2T other than ...' is one too",
      chain_of("TATA / Maruti", "GCV < 2.0 K", "COMP")[0]["qualifiers"] == [] and chain_of("TATA / Maruti ( Only Jeeto )", "GCV < 2.0 K", "COMP")[0]["qualifiers"] == ["Only Jeeto"]
      and chain_of("Others TATA ( Mahindra Below 2T other than Jeeto )", "GCV < 2.0 K", "COMP")[0]["qualifiers"] == ["Mahindra Below 2T other than Jeeto"])
for chain, what in ((("GCV < 2.0 K", "COMP + RSA"), "a leaf word nobody knows"), (("Crane Hire", "COMP"), "a product nobody knows"), (("School Bus", "SAOD"), "SAOD (not a policy type this script reads)"),
                    (("MISD CPM (Only JCB, Hitachi and Caterpillar)", "COMP"), "other body words (no T&C is known for them)"), (("PCV 3W (Non diesel)", "Diesel - Comp"), "a leaf fuel outside the segment's fuels"),
                    (("GCV < 2.0 K", "COMP", "Remarks"), "a header text that no rule reads"), (("GCV Special", "COMP"), "a GCV title with no weight band"), (("PCV Taxi 7+1 ( Gold )", "COMP"), "a bracket nobody understands"),
                    (("School Bus 18 seats", "COMP"), "a number left over in a title"), (("School Bus", "GCV 3W", "COMP"), "two products in one header"), (("School Bus",), "no Comp / STP wording"),
                    (("Others Tata / Piaggio", "GCV 2 - 2.5 K", "COMP"), "a make label with an unknown make")):
    try:
        chain_of(*chain)
        check(f"refused: {what}: {' / '.join(chain)}", False, "it was read")
    except mod.Refuse as e:
        check(f"refused: {what}: {' / '.join(chain)} - the reason says what was not understood", len(str(e)) > 15, str(e))
check("'Other makes' needs a make list of the same product before it: with one it is the complement, without one the column is refused",
      mod.classify_chain(["PCV 7+1 ( Other makes )", "COMP"], {"makes": {("taxi", (6.01, 8.01)): ["maruti", "toyota"]}})[0]["makes"] == "TAXI_ALL_EX_MARUTI_TOYOTA")
try:
    mod.classify_chain(["PCV 7+1 ( Other makes )", "COMP"], {"makes": {}})
    check("'Other makes' with nothing before it is refused", False)
except mod.Refuse:
    check("'Other makes' with nothing before it is refused", True)
mod.MAKE_DEFINITIONS.clear()
check("a deal-based product is not a rate card: its specs are marked deal and give no rows", chain_of("Employee Pick Up ( Only Deal Based )", "COMP")[0]["deal"])
check("the leaf words are used up one by one: 'Comp - Others ( Petrol + CNG + Battery )', 'Diesel - Comp & STP', 'Comp/STP', 'Comp (Non fresh)', 'All'",
      mod.parse_leaf("Comp - Others ( Petrol + CNG + Battery )") == dict(subs=["1+1"], fuels=ND4, age=None, age_lo=None, age_hi=None)
      and mod.parse_leaf("Diesel - Comp & STP") == dict(subs=["1+1", "STP"], fuels=["Diesel"], age=None, age_lo=None, age_hi=None)
      and mod.parse_leaf("Comp/STP")["subs"] == ["1+1", "STP"] and mod.parse_leaf("Comp\n(Non fresh)") == dict(subs=["1+1"], fuels=None, age="nonfresh", age_lo=None, age_hi=None)
      and mod.parse_leaf("All")["subs"] == ["1+1", "STP"])

# ================================================================ the real grid, run
df_full, warns, ask = build(real(), interactive=False)
bf = bare(df_full)
check("the real grid gives 11,461 rows (73 regions x 157), 7,970 of them zero cards", len(df_full) == 11461 and int((pi_of(df_full) == 0).sum()) == 7970, (len(df_full), int((pi_of(df_full) == 0).sum())))
check("every region gives 157 rows", set(bf.groupby("new_rto_list").size()) == {157} and bf["new_rto_list"].nunique() == 73)
check("the cluster name is the region exactly as the grid writes it (MUMBAI, Indore, Ladakh, GUJ 1, HP AND JAMMU, UP1), no prefix", {"MUMBAI", "Indore", "Ladakh", "GUJ 1", "HP AND JAMMU", "UP1"} <= set(bf["new_rto_list"]) and not bf["new_rto_list"].str.contains("RELIANCE").any())
check("the fixed columns: id 1, the standardised insurer name, Not Applicable, Renew, user 2966, no veh_use, ncb / cpa NA, tariff 0-100.01, dates DD/MM/YYYY",
      set(df_full["id"]) == {1} and set(df_full["insurance_company"]) == {"INDUSIND GENERAL INSURANCE COMPANY LIMITED"} and set(df_full["insurer_vertical"]) == {"Not Applicable"}
      and set(df_full["policy_type"]) == {"Renew"} and set(df_full["user_id"]) == {2966} and set(df_full["veh_use"]) == {""} and set(df_full["is_ncb"]) == {"NA"} and set(df_full["is_cpa"]) == {"NA"}
      and set(df_full["tariff_min"]) == {0} and set(df_full["tariff_max"]) == {100.01} and set(df_full["from_date"]) == {"01/11/2026"} and set(df_full["to_date"]) == {"31/03/2030"})
check("1+1 is On Net (the rate in pi_net / po_net), STP is On TP (pi_tp / po_tp); po = max(pi - 7, 0) on every row",
      (sel(df_full, sub_product="1+1")["pi_type"] == "On Net").all() and (sel(df_full, sub_product="STP")["pi_type"] == "On TP").all()
      and (sel(df_full, sub_product="1+1")["pi_tp_rate"] == 0).all() and (sel(df_full, sub_product="STP")["pi_net_rate"] == 0).all()
      and np.allclose((pi_of(df_full) - 7).clip(lower=0), df_full[["po_net_rate", "po_tp_rate"]].astype(float).max(axis=1)) and (df_full["po_type"] == df_full["pi_type"]).all())
check("STP rows carry NA in is_zd (and is_ncb) - the user's rule", (sel(df_full, sub_product="STP")["is_zd"] == "NA").all() and (sel(df_full, sub_product="STP")["is_ncb"] == "NA").all())
check("39 rows have a payin above 0 and up to 7: their payout is floored to 0", int(((pi_of(df_full) > 0) & (df_full[["po_net_rate", "po_tp_rate"]].astype(float).max(axis=1) == 0)).sum()) == 39)
check("no text cell of the upload has a leading / trailing space (the user's own export had two make names with one)", all((df_full[c].astype(str) == df_full[c].astype(str).str.strip()).all() for c in ("new_vehicle_makes", "new_rto_list", "add_tnc", "make_model_class")))
check("the make groups written: the master's own taxi groups (the user's decision of 20 Nov 2026), TATA,Maruti,Mahindra and its EX group, the family groups",
      sorted(df_full["new_vehicle_makes"].unique()) == sorted(["BUS_ALL", "PCV_3W_ALL", "TAXI_ALL", "TATA,Maruti,Mahindra", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", "GCV_4W_ALL", "GCV_3W_ALL",
                                                                   "TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG", "TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG", "MISCD_ALL"]), sorted(df_full["new_vehicle_makes"].unique()))
masters = mod.load_masters(mod.Asker(interactive=False))
names_known = set(masters.makes)
check("every make group written (and every single make of a list) is in the make / model master: no name is flagged", all(t in names_known for v in df_full["new_vehicle_makes"].unique() for t in v.split(",")), [t for v in df_full["new_vehicle_makes"].unique() for t in v.split(",") if t not in names_known])

# ---- cells of the real workbook against the rows (read again with openpyxl, independent of the script)
ws0 = wb["Nov'26"]
row_of = {str(ws0.cell(r, 3).value).strip(): r for r in range(5, 78)}


def cell(region, col):
    return ws0.cell(row_of[region], COL(col)).value


def rows_for(region, **kw):
    return sel(bf, new_rto_list=region, **kw)


def one(region, **kw):
    r = rows_for(region, **kw)
    assert len(r) == 1, (region, kw, len(r))
    return r.iloc[0]


for region, col, kw, sub in (
        ("MUMBAI", "D", dict(make_model_class="School Bus", fuel_type="Petrol", sub_product="1+1"), "pi_net_rate"),
        ("MUMBAI", "D", dict(make_model_class="School Bus", fuel_type="Electric", sub_product="STP"), "pi_tp_rate"),
        ("MUMBAI", "K", dict(new_vehicle_makes="TATA,Maruti,Mahindra", fuel_type="Diesel", sub_product="1+1"), "pi_net_rate"),
        ("MUMBAI", "L", dict(new_vehicle_makes="TATA,Maruti,Mahindra", fuel_type="CNG", sub_product="STP"), "pi_tp_rate"),
        ("MUMBAI", "G", dict(make_model_class="Taxi", new_vehicle_makes="TAXI_ALL", fuel_type="Petrol", sub_product="1+1", is_zd="NO"), "pi_net_rate"),
        ("MUMBAI", "H", dict(make_model_class="Taxi", new_vehicle_makes="TAXI_ALL", fuel_type="Diesel", sub_product="STP"), "pi_tp_rate"),
        ("MUMBAI", "J", dict(make_model_class="Taxi", new_vehicle_makes="TAXI_ALL", fuel_type="LPG", sub_product="1+1", is_zd="YES"), "pi_net_rate"),
        ("PUNE", "AB", dict(add_tnc="JCB, L&T,Caterpillar", fuel_type="Petrol", sub_product="1+1"), "pi_net_rate"),
        ("PUNE", "AD", dict(add_tnc="New,Agriculture Tractor", fuel_type="Petrol"), "pi_net_rate"),
        ("PUNE", "AE", dict(add_tnc="Old,Agriculture Tractor", fuel_type="Petrol"), "pi_net_rate"),
        ("NAGPUR", "E", dict(make_model_class="Passenger Carrying Rickshaw", fuel_type="CNG", sub_product="1+1"), "pi_net_rate"),
        ("Indore", "D", dict(make_model_class="School Bus", fuel_type="Diesel", sub_product="1+1"), "pi_net_rate")):
    want = round(cell(region, col) * 100, 6)
    got = float(one(region, **kw)[sub])
    check(f"{region} {col}: the grid cell {cell(region, col)!r} -> payin {want:g} in the row ({kw.get('make_model_class', kw.get('add_tnc', kw.get('new_vehicle_makes')))}, {kw['fuel_type']})", abs(got - want) < 1e-9, (got, want))
check("a payout is the payin - 7: Mumbai school bus 72.5 -> 65.5, Mumbai diesel taxi 25 -> 18",
      float(one("MUMBAI", make_model_class="School Bus", fuel_type="Petrol", sub_product="1+1")["po_net_rate"]) == 65.5 and float(one("MUMBAI", make_model_class="Taxi", new_vehicle_makes="TAXI_ALL", fuel_type="Diesel", sub_product="STP")["po_tp_rate"]) == 18)
check("a cell of float noise is cleaned: 0.47500000000000003 -> 47.5, 0.21250000000000002 -> 21.25 (no 47.50000000000001)", {47.5, 21.25} <= set(pi_of(df_full).round(6)) and not any(abs(x * 1e6 - round(x * 1e6)) > 1e-3 for x in pi_of(df_full)))
check("Employee Pick Up (AG, AH) gives no row although 4 regions have a payout in it", not (bf["make_model_class"].isin(["Employee Pick Up"])).any() and sum(1 for r in range(5, 78) if (ws0.cell(r, COL("AG")).value or 0) > 0) == 4)
check("Car Carrier is zero everywhere in November: its rows are written (zero cards) and NO guard note is flagged",
      not any("Car Carrier" in (r or "") for r in df_full["remarks"]) and len(sel(df_full, make_model_class="MISCD-Others", add_tnc="")) == 73 * 2 * 5)

# ---- against the user's own export of the same grid
d2 = df_full.copy()
mod.RTO_NAMES = "cluster"
lines = []
mod.EXTRA_SHEETS.clear()
got = mod.diff_against_reference(d2, REF, lines.append)
check("against the user's export: 157 new rows (Indore), 79 removed (their earlier Private Car rows), 34 changed (the diesel taxi) - and 11,270 rows equal in every column", got == (157, 79, 34), got)
ref = pd.read_csv(REF, dtype=str, keep_default_na=False)
changed = pd.DataFrame(mod.LAST_CHANGES)
check("the 34 rows that differ are ALL Diesel Taxi (1+1 and STP) in 17 regions", len(changed) == 34 and set(changed["fuel"]) == {"Diesel"} and set(changed["make_model_class"]) == {"Taxi"} and changed["region"].nunique() == 17, changed.head(3).to_dict("records"))
nxt = {str(ws0.cell(r, 3).value).strip(): ws0.cell(r + 1, COL("H")).value for r in range(5, 77)}
back = {str(ws0.cell(r, 3).value).strip().upper(): r for r in range(5, 78)}
shifted = 0
for rec in changed.drop_duplicates("region").itertuples():
    r = back[rec.region.upper()]
    below = ws0.cell(r + 1, COL("H")).value if r < 77 else None
    shifted += below is not None and abs(float(str(rec.earlier_upload).split(" | ")[0]) - below * 100) < 1e-6
check("the user's diesel values of those regions are the diesel value of the NEXT grid row (a one-row offset in their file); the script follows the grid", shifted >= 15, shifted)
check("the 157 new rows are all Indore, the 79 removed rows are all Private Car (the earlier set of 13-10-2026)", any("Indore" in x for x in lines) and any("79 x" in x or "Private Car" in x for x in lines), lines[-8:])
check("the changed rows are listed in the run's log (one line each) and in the review sheet 'Changed vs earlier upload'", sum(1 for x in lines if "earlier upload" in x and "grid" in x) >= 34 and "Changed vs earlier upload" in mod.EXTRA_SHEETS and len(mod.EXTRA_SHEETS["Changed vs earlier upload"]) == 34)

# ================================================================ the make groups, by the words of a label
check("a make label gives a list (TATA,Maruti,Mahindra), 'Others' + the makes gives the master's EX group, no label = the family group",
      mod.make_group("gcv4w", ["tata", "maruti", "mahindra"], False) == "TATA,Maruti,Mahindra" and mod.make_group("gcv4w", ["tata", "maruti", "mahindra"], True) == "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA"
      and mod.make_group("gcv4w", [], False) == "GCV_4W_ALL" and mod.make_group("taxi", [], False) == "TAXI_ALL")
check("the taxi make set is the master's own group (the user's choice), the complement its EX group; another set is composed from the makes and registered for the 'groups to create' list",
      mod.make_group("taxi", ["maruti", "toyota", "mahindra", "kia", "mg"], False) == "TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG" and mod.make_group("taxi", ["mg", "kia", "maruti", "toyota", "mahindra"], True) == "TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG"
      and mod.make_group("taxi", ["maruti", "toyota"], True) == "TAXI_ALL_EX_MARUTI_TOYOTA" and mod.MAKE_DEFINITIONS.get("TAXI_ALL_EX_MARUTI_TOYOTA") == ("TAXI_ALL", ["Maruti", "Toyota"])
      and mod.make_group("taxi", ["maruti", "toyota"], False) == "Maruti,Toyota")
mod.MAKE_DEFINITIONS.clear()

# ================================================================ conditional readings: only when the wording is in the grid
def drop_labels(wb):
    ws = wb["Nov'26"]
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row == 2 and mr.min_col > 3:
            ws.unmerge_cells(str(mr))
    for c in ("K2", "M2"):
        ws[c].value = None


d, w, _ = build(mini(tweak=drop_labels))
check("no make label above the GCV <2 T / 2-2.5 T columns: both are written for ALL makes (GCV_4W_ALL) and nothing is flagged for a qualifier", set(sel(d, product="GCV 4W", cc_min=0)["new_vehicle_makes"]) == {"GCV_4W_ALL"}
      and not any("model-level" in (r or "") for r in d["remarks"]) and "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA" not in set(d["new_vehicle_makes"]) and "TATA,Maruti,Mahindra" not in set(d["new_vehicle_makes"]))


def reword_label(wb):
    wb["Nov'26"]["K2"].value = "TATA / Maruti / Mahindra\n( Only Jeeto )"


d, w, _ = build(mini(tweak=reword_label))
hit = d[d["remarks"].str.contains("( Only Jeeto )", regex=False)]
check("a qualifier in ANOTHER wording is flagged too, on the rows of that column only (K, L: 5 fuels x 2 policy types x 6 regions)", len(hit) == 60 and set(hit["cc_max"].astype(float)) == {2000.01} and "Only Jeeto" in hit["remarks"].iloc[0], (len(hit), hit["remarks"].iloc[0][:150] if len(hit) else ""))


def car_carrier_money(wb):
    ws = wb["Nov'26"]
    ws["Y6"].value = 0.1                                                  # Nagpur: Car Carrier pays 10%


d, w, _ = build(mini(tweak=car_carrier_money))
g = d[d["remarks"].str.contains("Car Carrier", regex=False)]
check("Car Carrier with a payout above 0 in one region: its rows (that region, 5 fuels x 2 policy types) are flagged, the others are not; the note names the region", len(g) == 10 and set(bare(g)["new_rto_list"]) == {"NAGPUR"} and "NAGPUR" in g["remarks"].iloc[0], (len(g), g["remarks"].iloc[0][:200] if len(g) else ""))


def deal_based(wb):
    ws = wb["Nov'26"]
    ws["AG3"].value = "Employee Pick Up"


d, w, a = build(mini(tweak=deal_based))
check("'Employee Pick Up' without 'Only Deal Based' is a product this script has no mapping for: the columns are refused (warned, listed), every other row is written", any("Employee Pick Up" in x and "no mapping" in x for x in w) and len(d) == 6 * 157, (len(d), [x for x in w if "Employee" in x]))


def unknown_segment(wb):
    ws = wb["Nov'26"]
    ws["Q3"].value = "GCV Above 3.5K Special Permit"


d, w, a = build(mini(tweak=unknown_segment))
check("a segment with a word nobody understands refuses just that block (Q, R): warned with the column and the words, nothing guessed, the other columns are all written",
      any("column Q" in x and "special" in x.lower() for x in w) and any("column R" in x for x in w) and len(d) == 6 * 157 - 6 * 10 and not (d["cc_min"].astype(float) == 3500.01).any(), (len(d), [x[:150] for x in w]))
check("the refused columns are listed in the review sheet 'Unread parts' (NOT READ), the deal columns in 'Not written'", "Unread parts" in mod.EXTRA_SHEETS and set(mod.EXTRA_SHEETS["Unread parts"]["where"]) >= {"column Q", "column R"} and "Not written" in mod.EXTRA_SHEETS)


def unknown_leaf(wb):
    wb["Nov'26"]["O4"].value = "COMP + RSA"


d, w, a = build(mini(tweak=unknown_leaf))
check("a leaf with a word nobody understands ('COMP + RSA') refuses that column", any("column O" in x and "rsa" in x.lower() for x in w) and len(d) == 6 * 157 - 6 * 5, (len(d), [x[:120] for x in w]))


def no_other_list(wb):
    ws = wb["Nov'26"]
    ws["U3"].value = "PCV Taxi 7+1 ( Only NND ) All Fuel Type"        # no make list under the 7+1 taxi block: 'Other makes' (W) has nothing to be the complement of


d, w, a = build(mini(tweak=no_other_list))
check("'Other makes' with no make list before it for the same product refuses its column (W, X) - the other bus rows of that column are not written either, and it is said",
      any("column W" in x and "Other makes" in x for x in w) and not (d["make_model_class"] == "Other Bus").any(), [x[:160] for x in w if "column W" in x])
check("... and the taxi 7+1 columns without a make list (U, V) are written for ALL makes (TAXI_ALL, seats 6.01-8.01)", set(sel(d, sc_min=6.01)["new_vehicle_makes"]) == {"TAXI_ALL"} and len(sel(d, sc_min=6.01)) == 6 * 10)

# ================================================================ cells
d, w, a = build(mini())
check("a mini grid of six regions gives 6 x 157 rows", len(d) == 942)


def text_cell(wb):
    wb["Nov'26"]["K6"].value = "approx 65"


msg = stops(lambda: build(mini(tweak=text_cell)))
check("a payout cell that holds a text stops a run that cannot ask, naming the cell", "K6" in msg and "approx 65" in msg, msg)


def word_cells(wb):
    ws = wb["Nov'26"]
    ws["K6"].value, ws["L6"].value, ws["M6"].value, ws["N6"].value = "NA", None, "Declined", "-"


d, w, a = build(mini(tweak=word_cells))
nag = sel(bare(d), new_rto_list="NAGPUR", new_vehicle_makes="TATA,Maruti,Mahindra")
check("NA / a blank / Declined / '-' in a payout cell are zero cards (the declined entry), not a stop", len(nag) == 10 and (pi_of(nag) == 0).all() and any("blank payout" in x for x in a.decisions), (len(nag), a.decisions))


def big_cell(wb):
    wb["Nov'26"]["K6"].value = 65


msg = stops(lambda: build(mini(tweak=big_cell)))
check("a payout of 65 in a grid of fractions (0.65 = 65%) is put to the person, never taken as 6,500%", "K6" in msg and "65" in msg, msg)


def negative_cell(wb):
    wb["Nov'26"]["K6"].value = -0.1


check("a negative payout is put to the person", "K6" in stops(lambda: build(mini(tweak=negative_cell))))

# ================================================================ the structure of the table
d0, _, _ = build(mini())


def same(df_a, df_b):
    key = lambda x: x.sort_values(["new_rto_list", "product", "make_model_class", "new_vehicle_makes", "sub_product", "fuel_type", "is_zd", "add_tnc", "cc_min", "sc_min", "vehicle_age_min"]).reset_index(drop=True).drop(columns=["remarks"])
    return key(df_a).equals(key(df_b))


def swap_regions(wb):
    ws = wb["Nov'26"]
    for r in range(5, 11):
        ws.cell(r, 2).value, ws.cell(r, 3).value = ws.cell(r, 3).value, ws.cell(r, 2).value


d, _, a = build(mini(tweak=swap_regions))
check("the two 'RTO Region' columns swapped: the one with the more distinct names is still the cluster (decided by the values, not by the position)", same(d, d0) and any("names the clusters" in x for x in a.decisions), a.decisions)


def regions_differ(wb):
    ws = wb["Nov'26"]
    for r in range(5, 11):
        ws.cell(r, 2).value = f"OTHER {r}"                              # B: as many distinct names as C, other names


msg = stops(lambda: build(mini(tweak=regions_differ)))
check("two 'RTO Region' columns with the same number of different names: which names the clusters is not in the document - the run stops", "RTO Region" in msg and "which one" in msg, msg)


def no_parent(wb):
    wb["Nov'26"]["B4"].value = "Parent"                                  # only C says 'RTO Region' now


d, _, a = build(mini(tweak=no_parent))
check("one 'RTO Region' column only: used, nothing to decide", same(d, d0) and not any("names the clusters" in x for x in a.decisions))


def extra_header_row(wb):
    ws = wb["Nov'26"]
    ws.insert_rows(4)                                                    # a blank row inside the header band, between the segment titles and the leaf row
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 4:
            ws.unmerge_cells(str(mr))
            ws.merge_cells(start_row=mr.min_row + 1, start_column=mr.min_col, end_row=mr.max_row + 1, end_column=mr.max_col)


d, _, _ = build(mini(tweak=extra_header_row))
check("a blank row inside the header band changes nothing", same(d, d0))


def banner(wb):
    ws = wb["Nov'26"]
    L.insert(ws, "rows", 1, 1)
    ws["D1"].value = "Rates for November 2026 - brokers premier channel"
    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=34)


d, w, a = build(mini(tweak=banner))
check("a banner merged over every payout column is a NOTE, not a header level: the rows are the same, the banner is flagged on every row", same(d, d0) and (d["remarks"].str.contains("brokers premier channel", regex=False)).all(), (len(d), len(d0)))


def twin_payout(wb):
    ws = wb["Nov'26"]
    L.insert(ws, "cols", COL("K"), 1)                                    # a copy of K (GCV < 2 T COMP) with the same heading and HIGHER numbers, inside the merged label
    for r in range(1, 5):
        ws.cell(r, COL("K")).value = ws.cell(r, COL("K") + 1).value
    for r in range(5, 10):
        v = ws.cell(r, COL("K") + 1).value
        ws.cell(r, COL("K")).value = None if v is None else min(v + 0.1, 0.95)


d, w, a = build(mini(tweak=twin_payout))
d = mod.report(d, w, quiet, a)                                            # the duplicate rule is applied by report(): the same entry from two columns keeps the lowest payin
low_rows = sel(bare(d), new_rto_list="MUMBAI", new_vehicle_makes="TATA,Maruti,Mahindra", sub_product="1+1", fuel_type="Petrol")
check("a payout column there twice under one heading: the LOWEST payin is written (the user's rule of 12 Nov 2026) and the choice is listed", len(low_rows) == 1 and float(low_rows["pi_net_rate"].iloc[0]) == 67.5 and len(mod.LAST_LOWEST) >= 5
      and len(d) == 942, (len(low_rows), len(mod.LAST_LOWEST), len(d)))


def hidden_copy(wb):
    cp = wb.copy_worksheet(wb["Nov'26"])
    cp.title = "Nov'26 old"
    for row in cp.iter_rows(min_row=5):
        for c in row:
            if isinstance(c.value, float):
                c.value = c.value * 0.5
    cp.sheet_state = "hidden"


d, _, _ = build(mini(tweak=hidden_copy))
check("a hidden copy of the table (an old version) is never read", same(d, d0))


def visible_copy(wb):
    cp = wb.copy_worksheet(wb["Nov'26"])
    cp.title = "Nov'26 old"


msg = stops(lambda: build(mini(tweak=visible_copy)))
check("a visible copy of the table: which sheet is this month's is not in the document - a question (a stop when none can be asked)", "Nov'26 old" in msg and "same table" in msg, msg)
answers = iter(["1"])
d, _, a = build(mini(tweak=visible_copy), ask=mod.Asker(interactive=True, input_fn=lambda p: next(answers), out=quiet))
check("... answered 'the first sheet': that sheet is written, once", same(d, d0) and len(d) == 942)


def blank_headers(ws, c_from, c_to):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= 4 and not (mr.max_col < c_from or mr.min_col > c_to):
            ws.unmerge_cells(str(mr))
    for r in range(1, 5):
        for c in range(c_from, c_to + 1):
            ws.cell(r, c).value = None


def second_table(wb):
    """the table split in two sheets: the first keeps the headers of D..R, the copy those of S..AH (the other columns have no header on that sheet): different products on each"""
    ws = wb["Nov'26"]
    cp = wb.copy_worksheet(ws)
    cp.title = "Second"
    blank_headers(ws, COL("S"), COL("AH"))
    blank_headers(cp, COL("D"), COL("R"))


d, _, _ = build(mini(tweak=second_table))
check("two sheets that hold DIFFERENT products (the table split in two): both are read and every row is written once", same(d, d0) and len(d) == 942, (len(d), len(d0)))


def no_table(wb):
    wb["Nov'26"]["B4"].value, wb["Nov'26"]["C4"].value = "Parent", "Locale"      # no header says a region any more


msg = stops(lambda: build(mini(tweak=no_table)))
check("no sheet with a region header over payout columns: the run stops and says what it expected and what it found", "expected" in msg and "RTO Region" in msg and "found" in msg and "Visible sheets" in msg, msg)

# ================================================================ the notes
d, w, a = build(mini())
bd = bare(d)
kp = d[d["remarks"].str.contains("Kaali Pili", regex=False)]
check("Kaali Pili Taxi names Mumbai, Delhi and Kolkata but the mini grid has only Mumbai: a key that is no region of the grid cannot be placed, so the note keeps to the taxi rows of EVERY region (never fewer)",
      len(kp) == 6 * 34 and set(kp["make_model_class"]) == {"Taxi"}, (len(kp), set(kp["make_model_class"])))
kf = df_full[df_full["remarks"].str.contains("Kaali Pili", regex=False)]
check("... with the whole grid, where all three are regions, the note is keyed: the TAXI rows of MUMBAI, DELHI and KOLKATA only (3 x 34 rows)", len(kf) == 102 and set(bf.loc[kf.index, "new_rto_list"]) == {"MUMBAI", "DELHI", "KOLKATA"} and set(kf["make_model_class"]) == {"Taxi"}, (len(kf), set(bf.loc[kf.index, "new_rto_list"])))
sb = d[d["remarks"].str.contains("School Bus - Volume", regex=False)]
check("the School Bus volume slabs are flagged on the School Bus rows only (6 regions x 2 x 5 fuels)", len(sb) == 60 and set(sb["make_model_class"]) == {"School Bus"}, (len(sb), set(sb["make_model_class"])))
check("the School Bus terms & conditions (900 characters) are split in parts of at most 500; every part starts with the block's title (so it says what it is about and is scoped by it)",
      all(len(x) <= mod.MAX_NOTE_CHARS + 3 for x in (mod.LAST_NOTES and [t for t, _c in mod.LAST_NOTES])) and sum("Terms & Conditions for School Bus" in t for t, _c in mod.LAST_NOTES) >= 2)
lcv = d[d["remarks"].str.contains("LCV Contest", regex=False)]
check("the LCV contest is flagged on the goods vehicle rows (GCV 4W and 3W) only", len(lcv) > 0 and set(lcv["product"]) == {"GCV 4W", "GCV 3W"} and len(lcv) == len(d[d["product"].isin(["GCV 4W", "GCV 3W"])]), (len(lcv), set(lcv["product"])))
check("the hygiene guideline and 'Specific Deals' name no product: they are flagged on EVERY row (never fewer rows than the note may affect)", d["remarks"].str.contains("Guidelines for Hygiene", regex=False).all() and d["remarks"].str.contains("Specific Deals", regex=False).all())
check("the TW Scooter contest names two-wheelers only: no row of it is written, so nothing is flagged - and it is listed in 'Notes read' (no rows affected)",
      not d["remarks"].str.contains("TW - Scooter", regex=False).any() and any(v["status"] == "no rows affected" and "TW" in v["text"] for v in mod.NOTES_SEEN.values()), [v["status"] for v in mod.NOTES_SEEN.values()])
check("the (hidden) title 'PO Grid for Jun'26 FY 26-27' is a standing note: not flagged, and said so in 'Notes read'; the dates are never read from it",
      not d["remarks"].str.contains("PO Grid for", regex=False).any() and any("PO Grid" in v["text"] and v["status"].startswith("no action") for v in mod.NOTES_SEEN.values()), [v["text"][:30] for v in mod.NOTES_SEEN.values()])


def other_title(wb):
    wb["Nov'26"]["A1"].value = "PO Grid for Dec'27 FY 27-28"


d, _, _ = build(mini(tweak=other_title))
check("another month in the title changes nothing: from_date and to_date are the ones given, the title is not read", same(d, d0) and set(d["from_date"]) == {"01/11/2026"} and not d["remarks"].str.contains("PO Grid", regex=False).any())


def rules_under_table(wb):
    ws = wb["Nov'26"]
    ws["C20"].value = "Payout on GCV below 2 T will be reduced by 5% for Eicher make"
    ws["C21"].value = "Prepared by"


d, _, _ = build(mini(tweak=rules_under_table))
note = d[d["remarks"].str.contains("reduced by 5%", regex=False)]
check("a sentence under the table that names a product is flagged on the rows of that product; a short label ('Prepared by') is not a note", len(note) == len(d[d["product"].isin(["GCV 4W", "GCV 3W"])]) and not d["remarks"].str.contains("Prepared by", regex=False).any(), (len(note),))


def release_sheet(wb):
    wb.create_sheet("Release").append(["Release notes - nothing to read here"])


d, w, _ = build(mini(tweak=release_sheet))
check("a sheet of text with no sentence about payouts, slabs or deals is not a rule sheet: nothing is flagged, it is warned about and listed ('Unread parts')",
      same(d, d0) and any("'Release'" in x or "Release" in x for x in w) and "Release" in set(mod.EXTRA_SHEETS.get("Unread parts", pd.DataFrame({"sheet": []}))["sheet"]), w)
tmp = tempfile.mkdtemp(prefix="indusind_")
mod.RTO_NAMES = "cluster"
known = [t for t, _c in mod.LAST_NOTES]
resolved = os.path.join(tmp, "resolved.json")
json.dump(["Specific Deals; Wherever Specific deals are taken VLI/Contest/Annual Deals will not be applicable"], open(resolved, "w"))
os.environ["GRID_STANDING_NOTES_FILE"] = resolved
mod._RESOLVED = None
try:
    d, _, _ = build(mini())
finally:
    os.environ.pop("GRID_STANDING_NOTES_FILE", None)
    mod._RESOLVED = None
check("a note the person resolved in the web app (standing-notes file) is not flagged again", not d["remarks"].str.contains("Specific Deals", regex=False).any() and d["remarks"].str.contains("Guidelines for Hygiene", regex=False).all())

# ================================================================ the RTO groups
d, w, a = build(mini())
ind = d[bare(d)["new_rto_list"] == "Indore"]
check("Indore has no RTO group in the master: its 157 rows are written and flagged 'no RTO group'; the other five regions are not", len(ind) == 157 and ind["remarks"].str.contains("no RTO group in the master for the region 'Indore'", regex=False).all()
      and not d[bare(d)["new_rto_list"] != "Indore"]["remarks"].str.contains("no RTO group", regex=False).any())
grp = mod.EXTRA_SHEETS["RTO groups to create"]
check("'RTO groups to create' names the group to make, the rows flagged and the parent region the grid puts Indore under", len(grp) == 1 and grp.iloc[0]["group_needed"] == "RELIANCE_SEPT26_Indore" and grp.iloc[0]["rows_flagged"] == 157 and "MADHYA PRADESH" in grp.iloc[0]["items"], grp.to_dict("records"))
check("TRIVANDRUM is found by the master's own spelling (Thiruvananthapuram) and is not flagged", "TRIVANDRUM" not in " ".join(grp["group_needed"]) and not df_full[bf["new_rto_list"] == "TRIVANDRUM"]["remarks"].str.contains("no RTO group", regex=False).any())
check("the whole grid: exactly one region (Indore) has no group", df_full["remarks"].str.contains("no RTO group in the master", regex=False).sum() == 157 and int(sel(bf, new_rto_list="Indore").shape[0]) == 157)
old = mod.RTO_GROUP_PREFIX
mod.RTO_GROUP_PREFIX = "NOSUCHINSURER_"
try:
    d, w, a = build(mini())
finally:
    mod.RTO_GROUP_PREFIX = old
check("a prefix that matches no group of the master is stale: the check is skipped with a WARNING, nothing is flagged for a missing group", any("prefix is probably out of date" in x for x in w) and not d["remarks"].str.contains("no RTO group", regex=False).any(), w)
d, w, a = build(mini(), names="group")
check("group mode (GRID_RTO_NAMES=group): new_rto_list is the portal's group name RELIANCE_SEPT26_<region> and `check_names` does the master check; the bare names are the default", set(d["new_rto_list"]) == {"RELIANCE_SEPT26_MUMBAI", "RELIANCE_SEPT26_NAGPUR", "RELIANCE_SEPT26_PUNE", "RELIANCE_SEPT26_AHMEDNAGAR", "RELIANCE_SEPT26_AURANGABAD", "RELIANCE_SEPT26_Indore"}, set(d["new_rto_list"]))

# ================================================================ a small table of its own, and a stop that says WHY
def tiny(title="Tractor - (Agricultural without Trailer)", leaves=("Comp (Fresh)", "Comp (Non fresh)", "STP"), regions=("ALPHA", "BETA", "GAMMA", "DELTA", "EPSILON", "ZETA"), banner=None):
    """a minimal workbook: one visible sheet, 'Zone' | 'RTO Region' | the payout columns under ONE merged segment title, the regions below"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Grid"
    r0 = 1
    if banner:
        ws.cell(1, 3).value = banner
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=2 + len(leaves))
        r0 = 2
    ws.cell(r0, 3).value = title
    ws.merge_cells(start_row=r0, start_column=3, end_row=r0, end_column=2 + len(leaves))
    ws.cell(r0 + 1, 1).value, ws.cell(r0 + 1, 2).value = "Zone", "RTO Region"
    for j, leaf in enumerate(leaves):
        ws.cell(r0 + 1, 3 + j).value = leaf
    for i, region in enumerate(regions):
        ws.cell(r0 + 2 + i, 1).value, ws.cell(r0 + 2 + i, 2).value = "West", region
        for j in range(len(leaves)):
            ws.cell(r0 + 2 + i, 3 + j).value = 0.1 * (j + 1) + 0.01 * i
    return wb


d, w, a = build(tiny())
check("a table of its own whose payout columns all sit under ONE merged segment title is read (a title shared by every column is a header, not a banner): 3 columns x 5 fuels x 6 regions",
      len(d) == 3 * 5 * 6 and set(d["make_model_class"]) == {"MISCD-Tractor"} and set(d["add_tnc"]) == {"New,Agriculture Tractor", "Old,Agriculture Tractor", ""}, (len(d), set(d["add_tnc"])))
d, w, a = build(tiny(banner="Rates for November 2026 - brokers premier channel"))
check("... and a real banner above it (no product, no Comp / STP, no make) is still a note: flagged on every row, the rows are the same", len(d) == 90 and d["remarks"].str.contains("brokers premier channel", regex=False).all(), (len(d),))


def every_header_unknown(wb):
    ws = wb["Nov'26"]
    for c in range(COL("D"), COL("AH") + 1):
        ws.cell(4, c).value = f"{ws.cell(4, c).value} Premier"


msg = stops(lambda: build(mini(tweak=every_header_unknown)))
check("a table whose every payout column is refused stops with the reason of the first refusals (column and words), not just 'nothing found'", "expected a visible sheet" in msg and "no column read" in msg and "premier" in msg.lower() and "column D" in msg, msg)


# ================================================================ rows / columns hidden in Excel and blank cells: said, never a silent surprise
def hide_things(wb):
    ws = wb["Nov'26"]
    ws.row_dimensions[7].hidden = True                                   # Pune
    ws.column_dimensions["AD"].hidden = True                             # tractor, fresh


d, w, a = build(mini(tweak=hide_things))
check("rows / columns hidden in Excel that hold payouts are written like the others and WARNED about (a hidden region is not a silent surprise)",
      len(d) == 942 and any("hidden in Excel" in x and "Excel row 7" in x for x in w) and any("hidden in Excel" in x and "column(s) AD" in x for x in w), w)
check("the real grid has no hidden payout row or column (its hidden row 1 is the title) and no blank payout cell: neither warning appears", not any("hidden in Excel" in x or "blank" in x for x in warns), warns)


def many_blanks(wb):
    ws = wb["Nov'26"]
    for r in range(5, 9):                                                  # four of the six regions, columns D..R: every column keeps numbers in the other two rows
        for c in range(COL("D"), COL("S")):
            ws.cell(r, c).value = None


d, w, a = build(mini(tweak=many_blanks))
check("nearly half of the payout cells blank (formulas whose saved values are missing?) is WARNED about, with the fix (open, save, run again)", any("payout cells are blank" in x and "save it" in x for x in w), w)
d, w, a = build(mini(tweak=word_cells))
check("a few blank / NA cells are only counted in the run's decisions, not warned about", not any("payout cells are blank" in x for x in w) and any("blank payout" in x for x in a.decisions), (w, a.decisions))


# ================================================================ the sentences of the sheets of other products are listed, not applied
d, w, a = build(mini())
other = [v for v in mod.NOTES_SEEN.values() if v["status"].startswith("not applied")]
check("the sentences of the TW Scooter / Private Car / Short Term Taxi sheets are listed in 'Notes read' as 'not applied' (the rules of a product to come are not lost): e.g. the TW grid exclusion and the short-term-taxi volume enabler",
      len(other) >= 10 and any("Grid not applicable for Bike Comp" in v["text"] for v in other) and any("Volume payout is applicable only for other than Diesel" in v["text"] for v in other)
      and any("TW Scooter" in v["why"] for v in other) and any("Short Term Taxi" in v["why"] for v in other), (len(other), [v["text"][:40] for v in other][:5]))
check("... and they flag no row (they are about products that are not written)", not d["remarks"].str.contains("Bike Comp", regex=False).any() and not d["remarks"].str.contains("Volume payout", regex=False).any())

# ================================================================ dates and the command line
check("to_date before from_date stops before anything is read", "before" in stops(lambda: mod.build(None, dt.date(2026, 11, 1), dt.date(2026, 10, 1), quiet, mod.Asker(interactive=False), wb=mini())))
out = os.path.join(tmp, "out", "INDUSIND_UPLOAD_2026-11.csv")
env = dict(os.environ, PYTHONIOENCODING="utf-8")
p = subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.join(ROOT, "indusind_grid_to_config.py"), "--grid", SRC, "--from-date", "01/11/2026", "--to-date", "31/03/2030", "--out", out, "--no-prompt"],
                   cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
check("the command line writes the CSV and ONE review workbook next to it (Summary, Not written, RTO groups to create, Manual review notes, Notes read)",
      p.returncode == 0 and os.path.exists(out) and os.path.exists(out.replace("UPLOAD", "REVIEW").replace(".csv", ".xlsx")), (p.returncode, p.stdout[-300:], p.stderr[-300:]))
book = pd.read_excel(out.replace("UPLOAD", "REVIEW").replace(".csv", ".xlsx"), sheet_name=None, dtype=str, keep_default_na=False)
check("the review workbook lists the three other-product sheets and the two deal-based columns in 'Not written'", {"TW Scooter", "Private Car", "Short Term Taxi"} <= set(book["Not written"]["sheet"]) and {"column AG", "column AH"} <= set(book["Not written"]["where"]), book["Not written"].to_dict("records"))
csv = pd.read_csv(out, dtype=str, keep_default_na=False)
check("the file written is the upload: 42 columns in the portal's order, 11,461 rows, the bare cluster names", list(csv.columns) == mod.OUTPUT_COLUMNS and len(csv) == 11461 and "MUMBAI" in set(csv["new_rto_list"]), (list(csv.columns)[:5], len(csv)))
p = subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.join(ROOT, "indusind_grid_to_config.py"), "--grid", SRC, "--out", out, "--no-prompt"], cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
check("without --from-date / --to-date a run that cannot ask stops: the dates are given on every run", p.returncode != 0 and "--from-date" in (p.stdout + p.stderr) and "STOPPED" in (p.stdout + p.stderr), p.stdout + p.stderr)
p = subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.join(ROOT, "indusind_grid_to_config.py"), "--grid", os.path.join(helpers.INPUTS, "sbi_nov26.xlsx"), "--from-date", "01/11/2026", "--to-date", "31/03/2030", "--out", out, "--no-prompt"],
                   cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
check("another insurer's workbook stops with a message that says what was expected and what was found (no traceback)", p.returncode != 0 and "STOPPED - expected" in p.stderr + p.stdout and "Traceback" not in p.stderr + p.stdout, (p.stdout + p.stderr)[-400:])

# ================================================================ OCTOBER 2026 - 'Reliance - PO Grid Broking Premier 11-Oct 26': another generation of the same grid (the user, 23 Nov 2026)
# The same script reads both months: a header chain may name its product twice (a group title above a sub-title), says nil dep in words, lists fuels in the leaf, limits ages ('< 5 Years'), has a style
# column (Kaali Peeli), deal columns of any product and Private Car columns inside the commercial-vehicle sheet.  The answer key is the user's own October export (92): every payin and payout equal.
OCT = os.path.join(helpers.INPUTS, "indusind_oct26.xlsx")
OCT_REF = os.path.join(helpers.REFS, "indusind_oct26_rates_export.csv")
FO, TO = dt.date(2026, 10, 13), dt.date(2026, 10, 31)       # the dates of the user's October upload: the case's own - the script takes the dates on every run


def oreal():
    return openpyxl.load_workbook(OCT, data_only=True)


def omini(keep=(5, 6, 8, 9, 34, 41, 47, 70), tweak=None):
    """the real October workbook with only some of the 72 region rows (Mumbai, Nagpur, Ahmednagar, Aurangabad, Trivandrum, Delhi, Leh & Lahdhk, Kolkata) - every other sheet as it is"""
    wb = oreal()
    ws = wb["Oct 26"]
    for r in range(76, 4, -1):
        if r not in keep:
            ws.delete_rows(r)
    if tweak:
        tweak(wb)
    return wb


OL = [openpyxl.utils.get_column_letter(i) for i in range(4, 34)]            # D..AG
osh = mod.Sheet(oreal()["Oct 26"])
oanchor = mod.find_anchor(osh)
otbl = mod.RateTable(osh, *oanchor)
otbl.pick_region(quiet, mod.Asker(interactive=False))
check("October: the header row is found by its words - only column C says a region ('RTO Region FY 26-27'; B says 'RTO Fy 26 -27'), on row 4", oanchor == (4, [3]), oanchor)
check("October: 72 data rows from Excel row 5, the header band is rows 2-4 (row 1 is the visible title: no payout column under it)", otbl.ok and len(otbl.data) == 72 and otbl.first == 5 and otbl.top == 2, (otbl.first, otbl.top, len(otbl.data)))
check("October: the cluster column is C (72 distinct names); the PARENT column B is found by its values although its header says no region (27 distinct names)", otbl.rcol == 3 and otbl.parent == 2, (otbl.rcol, otbl.parent))
check("October: the stray backtick over 'School Bus' is noise, not a header level: the chain of D is School Bus / All", otbl.chain[COL("D")] == ["School Bus", "All"], otbl.chain[COL("D")])
ogood, orefused, oempty = mod.classify_columns(otbl)
oby = {c["letter"]: specs for c, specs in ogood}
check("October: every one of the 30 payout columns D..AG is understood, none refused, none empty", len(ogood) == 30 and not orefused and not oempty and sorted(oby, key=COL) == OL, (len(ogood), orefused[:2], [c["letter"] for c in oempty]))


def ospec(letter, i=0):
    s = oby[letter][i]
    return (s["product"], s["klass"], s["makes"], s["fuels"], s["subs"], s["zd"], tuple(s["age"]), tuple(s["cc"]), tuple(s["sc"]), s["tnc"]) if not (s["deal"] or s["skip"]) else ("SKIP",)


def row10(prod, klass, mk, fuels, sub, zd, age, cc, sc, tnc):
    return (prod, klass, mk, fuels, [sub] if isinstance(sub, str) else sub, zd, age, cc, sc, tnc)


OEXPECT = {
    "D": ("PCV 4W", "School Bus", "BUS_ALL", ALL5, ["1+1", "STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "E": ("GCV 4W", "NA", "TATA,Maruti,Mahindra", ALL5, ["1+1"], "NA", (0, 99), (0, 2000.01), (0, 999), ""),
    "F": ("GCV 4W", "NA", "TATA,Maruti,Mahindra", ALL5, ["STP"], "NA", (0, 99), (0, 2000.01), (0, 999), ""),
    "G": ("GCV 4W", "NA", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", ALL5, ["1+1"], "NA", (0, 99), (0, 2500.01), (0, 999), ""),
    "H": ("GCV 4W", "NA", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", ALL5, ["STP"], "NA", (0, 99), (0, 2500.01), (0, 999), ""),
    "I": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["1+1"], "NA", (0, 99), (2500.01, 3500.01), (0, 999), ""),
    "J": ("GCV 4W", "NA", "GCV_4W_ALL", ALL5, ["STP"], "NA", (0, 99), (2500.01, 3500.01), (0, 999), ""),
    "K": ("PCV 3W", "Passenger Carrying Rickshaw", "PCV_3W_ALL", ND4, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "L": ("PCV 3W", "Passenger Carrying Rickshaw", "PCV_3W_ALL", ND4, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "M": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["1+1"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "N": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["STP"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "O": ("PCV 4W", "Taxi", "TAXI_ALL", ["Diesel"], ["1+1"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "P": ("PCV 4W", "Taxi", "TAXI_ALL", ["Diesel"], ["STP"], "NO", (0, 99), (0, 99999), (0, 6.01), ""),
    "Q": ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["1+1"], "YES", (0, 99), (0, 99999), (0, 6.01), ""),
    "R": ("PCV 4W", "Taxi", "Maruti,Toyota,KIA,MG", ALL5, ["1+1"], "NO", (0, 99), (0, 99999), (6.01, 8.01), ""),
    "S": ("PCV 4W", "Taxi", "Maruti,Toyota,KIA,MG", ALL5, ["STP"], "NO", (0, 99), (0, 99999), (6.01, 8.01), ""),
    "T": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 1), (0, 99999), (0, 999), "New,Agriculture Tractor"),
    "U": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["1+1"], "NA", (1, 5.01), (0, 99999), (0, 999), "Old,Agriculture Tractor"),
    "V": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["1+1"], "NA", (5.01, 99), (0, 99999), (0, 999), "Old,Agriculture Tractor"),
    "W": ("MISCD", "MISCD-Tractor", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), ""),
    "X": ("PCV 4W", "Taxi", "TAXI_ALL", ALL5, ["1+1", "STP"], "NA", (0, 99), (0, 99999), (0, 999), "Kaali Peeli"),
    "Z": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), "Flat Bed"),
    "AA": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), "Flat Bed"),
    "AB": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["1+1"], "NA", (0, 99), (0, 99999), (0, 999), "JCB, L&T,Caterpillar"),
    "AC": ("MISCD", "MISCD-Others", "MISCD_ALL", ALL5, ["STP"], "NA", (0, 99), (0, 99999), (0, 999), "JCB, L&T,Caterpillar"),
}
for letter, want in OEXPECT.items():
    check(f"October column {letter}: {oby[letter][0]['segment'][:38]!r} / {oby[letter][0]['chain'][-1][:26]!r} -> product, class, make group, fuels, policy types, nil dep, age, weight, seats, T&C",
          ospec(letter) == want, (ospec(letter), want))
check("October columns Y and AG are DEAL columns ('Employee Pick up (Only Deal Based)', 'Car Carrier (Only deal basis)'): read, not written; the wording is kept for the 'Not written' text",
      oby["Y"][0]["deal"] and oby["AG"][0]["deal"] and oby["AG"][0]["seg"]["deal_words"].lower() == "only deal basis" and oby["Y"][0]["seg"]["deal_words"].lower() == "only deal based")
check("October columns AD, AE, AF ('PC STP Non-Diesel' x three cc bands) are PRIVATE CAR columns inside the commercial-vehicle sheet: understood and NOT written (the user's decision 5)",
      all(oby[l][0]["skip"] and oby[l][0]["other"] == "Private Car" for l in ("AD", "AE", "AF")), [oby[l][0].get("skip") for l in ("AD", "AE", "AF")])
check("October column X is the Kaali Peeli style: the T&C is the style, the rows are flagged as one ambiguity; no other column has a style", [l for l, sp in oby.items() if any(s.get("style") for s in sp)] == ["X"] and oby["X"][0]["style"] == "Kaali Peeli")
check("October: the qualifier columns are E, F, G, H (the same wording as November); Car Carrier is a deal now, so no column is a guarded catch-all",
      sorted(l for l, sp in oby.items() if any(s.get("qualifiers") for s in sp)) == ["E", "F", "G", "H"] and not [l for l, sp in oby.items() if any(s.get("guard") for s in sp)])
check("October 'Others' band (G-H, 'GCV < 2.5 K') is 0-2500.01 as the grid says (the user's export has the same band), not November's 2000.01-2500.01", ospec("G")[7] == (0, 2500.01) and ospec("E")[7] == (0, 2000.01))

# ---- the new wordings by themselves (a reusable brain: the same words give the same meaning anywhere)
check("a region header is any SHORT cell with 'region' / 'cluster' among a few words - 'RTO Region FY 26-27 (New)' too (the old match stopped at three extra words) - but not a code column or a sentence",
      mod.region_header("RTO Region FY 26-27") and mod.region_header("RTO Region FY 2026-27 (New)") and mod.region_header("Cluster") and mod.region_header("RTO Region")
      and not mod.region_header("Region_Code") and not mod.region_header("RTO Fy 26 -27") and not mod.region_header("Region wise payout applicable for all the policies issued in the month"))
check("makes: 'Kia& MG' (the & has a space on one side only) splits, 'M&M' and 'L&T' do not", mod.split_makes("Maruti,Toyota,Kia& MG") == (["maruti", "toyota", "kia", "mg"], []) and mod.split_makes("M&M") == (["mahindra"], []))
check("two product texts of one header that AGREE are merged: a group title + a sub-title + a leaf with the fuels = taxi <6 seats, no nil dep, Petrol / CNG / LPG / Electric, Comp",
      spec1("PCV TAXI <6 Seater", "PCV TAXI < 6Str Without Nil Dep (Others)", "Comp- Other -(Battery,CNG, Petrol)") == ("PCV 4W", "Taxi", "TAXI_ALL", ND4, ["1+1"], "NO", (0, 99), (0, 99999), (0, 6.01), ""))
check("... in any order of the levels, and '(Diesel Only)' + 'STP- Diesel' = diesel STP",
      spec1("Comp- Other -(Battery,CNG, Petrol)", "PCV TAXI < 6Str Without Nil Dep (Others)", "PCV TAXI <6 Seater") == spec1("PCV TAXI <6 Seater", "PCV TAXI < 6Str Without Nil Dep (Others)", "Comp- Other -(Battery,CNG, Petrol)")
      and spec1("PCV TAXI <6 Seater", "PCV TAXI < 6Str Without Nil Dep (Diesel Only)", "STP- Diesel")[3:5] == (["Diesel"], ["STP"]))
check("nil dep in words: 'Nil Dep' = YES, 'Without Nil Dep' = NO, 'NND' = NO, '(Only without Nil Dep)' = NO, '(Only ND)' = YES",
      spec1("PCV TAXI < 6Str Nil Dep", "COMP")[5] == "YES" and spec1("PCV TAXI < 6Str Without Nil Dep", "COMP")[5] == "NO" and spec1("PCV TAXI < 6Str NND", "COMP")[5] == "NO"
      and spec1("PCV Taxi 7+1STR (Only without Nil Dep)", "COMP")[5] == "NO" and spec1("PCV Taxi 4+1 ( Only ND )", "COMP")[5] == "YES")
check("'7+1STR' (no space) and 'Kia& MG' in one title: seats 6.01-8.01, the four makes as the master's single-make groups (a make set the master has no group for)",
      spec1("PCV Taxi 7+1STR", "PCV Taxi 7+1STR (Only without Nil Dep) (Maruti,Toyota,Kia& MG) All Fuel Type", "COMP")[2:] == ("Maruti,Toyota,KIA,MG", ALL5, ["1+1"], "NO", (0, 99), (0, 99999), (6.01, 8.01), ""))
check("'All Make/All Ages' and 'All Makes' say 'no restriction': the family group, every age", spec1("GCV 2.5K-3.5K", "All Make/All Ages", "COMP")[2] == "GCV_4W_ALL" and spec1("GCV 2.5K-3.5K", "All Makes", "COMP")[6] == (0, 99))
for chain, what in ((("PCV TAXI <6 Seater", "PCV TAXI 7+1 Seater", "COMP"), "two seat bands in one header"), (("PCV TAXI <6 Seater", "GCV < 2.0 K", "COMP"), "two products in one header"),
                    (("PCV TAXI Nil Dep", "PCV TAXI Without Nil Dep", "COMP"), "nil dep and without nil dep"), (("PCV Taxi 7+1 (Maruti, Toyota)", "PCV Taxi 7+1 (Kia, MG)", "COMP"), "two make lists"),
                    (("PCV TAXI (Others)", "COMP"), "'(Others)' with no fuels in the Comp / STP header"), (("PCV TAXI <6 Seater (Others)", "Comp - Diesel"), "'(Others)' beside diesel"),
                    (("PCV TAXI (Diesel Only)", "Comp - Petrol"), "a fuel the segment does not allow"), (("GCV 2.5K-3.5K", "TATA", "All Make/All Ages", "COMP"), "'All Make' beside a make label"),
                    (("School Bus", "Kaali Peeli", "COMP"), "a style on a product that has none"), (("School Bus", "Comp (Fresh) < 5 Years"), "a fresh column with an age limit"),
                    (("School Bus", "Comp > 5 Years > 6 Years"), "two lower age limits"), (("PCV TAXI", "PC STP Non-Diesel", "COMP"), "a Private Car text beside a taxi text"),
                    (("MISD CPM", "COMP"), "a body-type column whose body words are nowhere in the header"), (("PCV TAXI <6 Seater", "PCV TAXI < 6Str Gold", "COMP"), "a word nobody knows in a merged title")):
    try:
        chain_of(*chain)
        check(f"October: refused - {what}: {' / '.join(chain)}", False, "it was read")
    except mod.Refuse as e:
        check(f"October: refused - {what}: {' / '.join(chain)}", len(str(e)) > 15, str(e))


def age_of(text):
    s = chain_of("Tractor", text)[0]
    return tuple(s["age"]), s["tnc"]


for text, want in (("Comp (Fresh)", ((0, 1), "New,Agriculture Tractor")), ("Comp (New)", ((0, 1), "New,Agriculture Tractor")), ("Comp (Non fresh) < 5 Years", ((1, 5.01), "Old,Agriculture Tractor")),
                   ("Comp (Non fresh) > 5 Years", ((5.01, 99), "Old,Agriculture Tractor")), ("Comp (Non fresh) <= 3 yrs", ((1, 3.01), "Old,Agriculture Tractor")),
                   ("Comp (Non fresh) >= 7 years", ((7, 99), "Old,Agriculture Tractor")), ("Comp (Non fresh) below 4 years", ((1, 4.01), "Old,Agriculture Tractor")),
                   ("Comp (Non fresh) above 8 years", ((8.01, 99), "Old,Agriculture Tractor")), ("Comp < 5 Years", ((0, 5.01), "")), ("Comp > 5 Years", ((5.01, 99), "")), ("Comp (Non fresh)", ((1, 99), "Old,Agriculture Tractor"))):
    check(f"age words: {text!r} -> {want[0]} and T&C {want[1]!r} (the standard slab convention: Fresh / New = 0-1, '< N years' up to N + 0.01, '> N years' from N + 0.01)", age_of(text) == want, age_of(text))
check("'only deal basis' / 'only deal based' make a column a deal whatever the product (a GCV title too): no rows", chain_of("GCV < 2.0 K (Only deal basis)", "COMP")[0]["deal"] and chain_of("Car Carrier (Only deal basis)", "Comp/STP")[0]["deal"]
      and chain_of("Employee Pick up", "Employee Pick up (Only Deal Based)", "Comp/STP")[0]["deal"])
kp = [chain_of(*c)[0] for c in (("PCV TAXI KAALI PEELI", "Kaali Peeli", "Comp/STP"), ("PCV TAXI KALI PILI", "Comp/STP"), ("PCV TAXI", "Kalli Pilli", "STP"))]
check("Kaali Peeli in any spelling (Kaali Peeli, Kali Pili, Kalli Pilli) on a taxi: the T&C 'Kaali Peeli', every fuel, every seat count, no nil dep; a style is flagged by the run, a taxi without it is not",
      all(s["style"] == "Kaali Peeli" and s["tnc"] == "Kaali Peeli" and s["fuels"] == ALL5 and tuple(s["sc"]) == (0, 999) and s["zd"] == "NA" for s in kp) and kp[0]["subs"] == ["1+1", "STP"]
      and not chain_of("PCV TAXI <6 Str", "COMP")[0]["style"])
check("a Private Car or two-wheeler column is understood and NOT written, whatever its other words ('<1000CC' nobody reads): the reason is said, nothing is guessed",
      chain_of("PC STP Non-Diesel", "<1000CC", "All Makes")[0]["skip"].startswith("another product (Private Car)") and chain_of("TW Scooter", "Comp")[0]["other"] == "Two Wheeler"
      and chain_of("Private Car", "Diesel / EV", "Comp")[0]["other"] == "Private Car")
check("a stray closing bracket is punctuation, only the words are read: 'Agricultural Tractor W/o Trailer)' is a tractor", spec1("Agricultural Tractor W/o Trailer)", "Tractor - (Agricultural without Trailer)", "Comp (Fresh)")[1] == "MISCD-Tractor")

# ---- the real October grid, run
dfo, wo, aso = build(oreal(), interactive=False, f=FO, t=TO)
bo = bare(dfo)
full_rto_sheet = mod.EXTRA_SHEETS["RTO groups to create"].copy()
full_decisions = list(aso.decisions)
check("the real October grid gives 8,784 rows = 72 regions x 122 rows, 5,369 of them zero cards", len(dfo) == 8784 and set(bo.groupby("new_rto_list").size()) == {122} and bo["new_rto_list"].nunique() == 72
      and int((pi_of(dfo) == 0).sum()) == 5369, (len(dfo), int((pi_of(dfo) == 0).sum())))
check("the dates are the ones given to the run (13/10/2026 - 31/10/2026): nothing is read from the workbook's title ('PO Grid for 11th Oct')", set(dfo["from_date"]) == {"13/10/2026"} and set(dfo["to_date"]) == {"31/10/2026"})
check("the cluster name is the region exactly as the grid writes it (MUMBAI, 'LEH & LAHDHK', TRIVANDRUM, UP1, GUJ 1), no prefix",
      {"MUMBAI", "LEH & LAHDHK", "TRIVANDRUM", "UP1", "GUJ 1"} <= set(bo["new_rto_list"]) and not bo["new_rto_list"].str.contains("RELIANCE").any())
check("po = max(pi - 7, 0) on every row (39 rows have a payin above 0 and up to 7), 1+1 On Net, STP On TP, STP rows NA in is_zd / is_ncb, the other fixed columns as in November",
      np.allclose((pi_of(dfo) - 7).clip(lower=0).where(pi_of(dfo) > 0, 0), dfo[["po_net_rate", "po_tp_rate"]].astype(float).max(axis=1)) and int(((pi_of(dfo) > 0) & (dfo[["po_net_rate", "po_tp_rate"]].astype(float).max(axis=1) == 0)).sum()) == 39
      and (sel(dfo, sub_product="STP")["is_zd"] == "NA").all() and (sel(dfo, sub_product="STP")["is_ncb"] == "NA").all() and set(dfo["user_id"]) == {2966} and set(dfo["policy_type"]) == {"Renew"})
check("every text name written is in the master and none has a leading / trailing space: the make groups are the family groups, TATA,Maruti,Mahindra and its EX group, Maruti,Toyota,KIA,MG",
      sorted(dfo["new_vehicle_makes"].unique()) == sorted(["BUS_ALL", "PCV_3W_ALL", "TAXI_ALL", "TATA,Maruti,Mahindra", "GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", "GCV_4W_ALL", "MISCD_ALL", "Maruti,Toyota,KIA,MG"])
      and all(t in names_known for v in dfo["new_vehicle_makes"].unique() for t in v.split(",")) and all((dfo[c].astype(str) == dfo[c].astype(str).str.strip()).all() for c in ("new_vehicle_makes", "add_tnc", "new_rto_list")))

# ---- every cell of the October workbook against the rows (read again with openpyxl, independent of the script)
wso = oreal()["Oct 26"]
live = [l for l in OL if not any(s["deal"] or s["skip"] for s in oby[l])]
by_region = {k: g for k, g in bo.groupby("new_rto_list")}


def col_rows(g, spec):
    """the rows of one region that a column's card describes (the hand-checked OEXPECT row: product, class, make group, fuels, policy types, nil dep, age, weight, seats, T&C)"""
    prod, klass, mk, fuels, subs, zd, age, cc, sc, tnc = spec
    m = ((g["product"] == prod) & (g["make_model_class"] == klass) & (g["new_vehicle_makes"] == mk) & g["fuel_type"].isin(fuels) & g["sub_product"].isin(subs) & (g["add_tnc"] == tnc)
         & np.isclose(g["vehicle_age_min"].astype(float), age[0]) & np.isclose(g["vehicle_age_max"].astype(float), age[1]) & np.isclose(g["cc_min"].astype(float), cc[0]) & np.isclose(g["cc_max"].astype(float), cc[1])
         & np.isclose(g["sc_min"].astype(float), sc[0]) & np.isclose(g["sc_max"].astype(float), sc[1]) & ((g["sub_product"] == "STP") | (g["is_zd"] == zd)))
    return g[m]


wrong, bad_count = [], []
for r in range(5, 77):
    region = " ".join(str(wso.cell(r, 3).value).split())
    for l in live:
        got = col_rows(by_region[region], OEXPECT[l])
        want = round(float(wso.cell(r, COL(l)).value or 0) * 100, 6)
        if len(got) != len(OEXPECT[l][3]) * len(OEXPECT[l][4]):
            bad_count.append((region, l, len(got)))
        elif not np.allclose(pi_of(got).astype(float), want, atol=1e-9):
            wrong.append((region, l, want, pi_of(got).tolist()[:3]))
check(f"every payout cell of the {len(live)} written columns x 72 regions is in its rows as the cell x 100, in exactly the rows its card describes (1,800 cells read again with openpyxl, independent of the script)",
      not wrong and not bad_count and len(live) == 25, (wrong[:3], bad_count[:3], len(live)))
mum = sel(bo, new_rto_list="MUMBAI")
check("Mumbai by hand: school bus 72.5, GCV < 2 T 67.5, 'Others' 60 / 57.5, PCV 3W 52.5, taxi no nil dep 55 / diesel 25, nil dep 40, 7+1 zero, tractor fresh 22.5 / < 5 y 22.5 / > 5 y 17.5, Kaali Peeli 40, JCB 20, flat bed 0",
      all(abs(float(x) - y) < 1e-9 for x, y in ((sel(mum, make_model_class="School Bus", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 72.5),
                                                  (sel(mum, new_vehicle_makes="TATA,Maruti,Mahindra", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 67.5),
                                                  (sel(mum, new_vehicle_makes="GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 60),
                                                  (sel(mum, new_vehicle_makes="GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA", sub_product="STP", fuel_type="Petrol")["pi_tp_rate"].iloc[0], 57.5),
                                                  (sel(mum, product="PCV 3W", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 52.5),
                                                  (sel(mum, make_model_class="Taxi", is_zd="NO", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 55),
                                                  (sel(mum, make_model_class="Taxi", is_zd="NO", sub_product="1+1", fuel_type="Diesel", sc_max=6.01)["pi_net_rate"].iloc[0], 25),
                                                  (sel(mum, make_model_class="Taxi", is_zd="YES", sub_product="1+1")["pi_net_rate"].iloc[0], 40),
                                                  (sel(mum, add_tnc="Kaali Peeli", sub_product="STP", fuel_type="Petrol")["pi_tp_rate"].iloc[0], 40),
                                                  (sel(mum, add_tnc="JCB, L&T,Caterpillar", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 20),
                                                  (sel(mum, add_tnc="Flat Bed", sub_product="1+1", fuel_type="Petrol")["pi_net_rate"].iloc[0], 0)))
      and sorted(sel(mum, make_model_class="MISCD-Tractor", sub_product="1+1", fuel_type="Petrol")[["vehicle_age_min", "pi_net_rate"]].astype(float).values.tolist()) == [[0, 22.5], [1, 22.5], [5.01, 17.5]])
check("the tractor rows carry the standard age slabs: fresh 0-1 (T&C New), non fresh < 5 y 1-5.01, > 5 y 5.01-99 (T&C Old), STP 0-99 with no T&C",
      {(float(a), float(b), t) for a, b, t in sel(bo, make_model_class="MISCD-Tractor", sub_product="1+1")[["vehicle_age_min", "vehicle_age_max", "add_tnc"]].values.tolist()} ==
      {(0.0, 1.0, "New,Agriculture Tractor"), (1.0, 5.01, "Old,Agriculture Tractor"), (5.01, 99.0, "Old,Agriculture Tractor")}
      and set(sel(bo, make_model_class="MISCD-Tractor", sub_product="STP")["add_tnc"]) == {""})
check("Kaali Peeli: 720 rows (72 regions x 5 fuels x 1+1 / STP), 40 in Mumbai, Delhi, Kolkata and 0 in the other 69 regions, the T&C 'Kaali Peeli'",
      len(sel(bo, add_tnc="Kaali Peeli")) == 720 and {r for r, g in sel(bo, add_tnc="Kaali Peeli").groupby("new_rto_list") if (pi_of(g) == 40).all()} == {"MUMBAI", "DELHI", "KOLKATA"}
      and int((pi_of(sel(bo, add_tnc="Kaali Peeli")) == 0).sum()) == 69 * 10)
kp_rows = sel(dfo, add_tnc="Kaali Peeli")
check("the Kaali Peeli rows are FLAGGED as an ambiguity (no vehicle-type field; the ordinary taxi cards also cover these taxis) and no other row carries that flag",
      kp_rows["remarks"].str.contains("the style 'Kaali Peeli' has no field in the upload", regex=False).all() and int(dfo["remarks"].str.contains("the style 'Kaali Peeli'", regex=False).sum()) == 720)
check("the M & HCV grid (GVW 12 T and above, a table of region expressions in 'Enabler, PO Guidelines & T&C') is NOT written (the user's decision 3): it stays a flagged note on the GCV 4W rows - and only on them",
      not (dfo["cc_min"].astype(float) >= 12000).any() and not dfo["new_vehicle_makes"].str.contains("ASHOK", case=False).any()
      and set(sel(dfo, product="GCV 4W")["remarks"].str.contains("M & HCV Grid", regex=False)) == {True} and not dfo[dfo["product"] != "GCV 4W"]["remarks"].str.contains("M & HCV Grid", regex=False).any()
      and int(dfo["remarks"].str.contains("M & HCV Grid", regex=False).sum()) == 72 * 30)
check("the hygiene guidelines flag every row and STAY flagged (the user, 23 Nov 2026); the qualifier columns flag their 1,440 rows; all 8,784 rows carry a flag",
      dfo["remarks"].str.contains("HYGIENE GUIDELINES", regex=False).all() and int(dfo["remarks"].str.contains("model-level condition", regex=False).sum()) == 1440 and (dfo["remarks"] != "").all())
nw = pd.DataFrame(mod.LAST_NOT_WRITTEN)
check("'Not written' lists the three other-product sheets, the two deal columns (Employee Pick up, Car Carrier) and the three Private Car columns of the CV sheet (AD, AE, AF)",
      {"TW SCOOTER GRID", "PRIVATE CAR COMP, SAOD & STP", "Short Term Taxi"} <= set(nw["sheet"]) and {"column Y", "column AG", "column AD", "column AE", "column AF"} <= set(nw["where"])
      and any("Only deal basis" in w for w in nw["why"]) and any("another product (Private Car)" in w for w in nw["why"]), nw[["sheet", "where"]].values.tolist())
check("a deal column is not counted as unread: no refusal, no INCOMPLETE line, the only warning is the make / model list sheet nobody reads", not any(x.startswith("INCOMPLETE") for x in wo) and len(wo) == 1 and "Make Model wise segment" in wo[0], wo)

# ---- the workbook's own RTO List: hints and the codes of the groups to create
rl = mod.read_rto_list(mod.Sheet(oreal()["RTO List"]))
regions72 = [" ".join(str(wso.cell(r, 3).value).split()) for r in range(5, 77)]
check("the RTO List sheet is read by its headers (a code column and a city column): 1,501 RTO codes under 72 regions", rl is not None and sum(map(len, rl.values())) == 1501 and len(rl) == 72, (None if rl is None else (sum(map(len, rl.values())), len(rl))))
check("a sheet that is no RTO list gives None (the contests sheet, the make / model list, a rate table)", all(mod.read_rto_list(mod.Sheet(oreal()[n])) is None for n in ("Contest", "Make Model wise segment", "Oct 26", "Enabler, PO Guidelines & T&C")))
check("'did you mean' hints: the two regions the RTO List spells another way are paired with their spelling - LEH & LAHDHK = Ladakh, TRIVANDRUM = Thiruvananthapuram - and nothing else gets a hint",
      mod.region_hints(regions72, rl) == {"LEH & LAHDHK": "Ladakh", "TRIVANDRUM": "Thiruvananthapuram"}, mod.region_hints(regions72, rl))
check("a hint needs a likeness: a grid region that is in no way like the RTO List's free name gets none (no wild guess by elimination), a new region the list lacks too gets none",
      mod.region_hints(["MUMBAI", "ZZZZ"], {"Mumbai": ["MH-01"] * 20, "Pune": ["MH-12"] * 20}) == {} and mod.region_hints(["MUMBAI", "INDORE"], {"Mumbai": ["MH-01"] * 20, "Bhopal": ["MP-04"] * 20}) == {}
      and mod.region_hints(["MUMBAI", "PUNEE"], {"Mumbai": ["MH-01"] * 20, "Pune": ["MH-12"] * 20}) == {"PUNEE": "Pune"})
leh = bo[bo["new_rto_list"] == "LEH & LAHDHK"]["remarks"]
check("the region the master has no group for says what the workbook's RTO List calls it and which master group exists: 'did you mean Ladakh ... RELIANCE_SEPT26_Ladakh'; the written name stays the grid's",
      len(leh) == 122 and leh.str.contains("did you mean 'Ladakh'", regex=False).all() and leh.str.contains("RELIANCE_SEPT26_Ladakh", regex=False).all())
check("TRIVANDRUM has its group through the alias (the master spells it Thiruvananthapuram): no 'no RTO group' flag; the decisions list both spellings the RTO List writes differently",
      not bo[bo["new_rto_list"] == "TRIVANDRUM"]["remarks"].str.contains("no RTO group", regex=False).any() and any("LEH & LAHDHK = Ladakh, TRIVANDRUM = Thiruvananthapuram" in x for x in full_decisions), full_decisions[-3:])
check("'RTO groups to create' carries the RTO codes of the region (from the workbook's RTO List, under the spelling it hints at)", "rto_codes" in full_rto_sheet.columns and len(full_rto_sheet) == 1
      and set(x.strip() for x in full_rto_sheet["rto_codes"].iloc[0].split(",")) == set(rl["Ladakh"]) and "Ladakh" in full_rto_sheet["items"].iloc[0], full_rto_sheet.to_dict("records"))
d, w, a = build(omini(), f=FO, t=TO)
ok_pairs = {"LEH & LAHDHK = Ladakh", "TRIVANDRUM = Thiruvananthapuram"}
said = [x for x in a.decisions if "spells differently" in x]
check("a hint is never wrong: with only 8 of the 72 regions in the grid (the RTO List then has many free names) the hints made are right, or there are none",
      all(set(p.strip() for p in x.split(": ", 1)[1].split(",")) <= ok_pairs for x in said), said)


def drop_rto_list(wb):
    del wb["RTO List"]


d2, w2, a2 = build(omini(tweak=drop_rto_list), f=FO, t=TO)
bd2 = bare(d2)
leh2 = bd2[bd2["new_rto_list"] == "LEH & LAHDHK"]["remarks"].iloc[0]
check("without the RTO List sheet the check is as before: the same 'no RTO group' flag with no hint, and no codes column", "did you mean" not in leh2 and "no RTO group in the master for the region 'LEH & LAHDHK' (looked for RELIANCE_SEPT26_LEH_LAHDHK): create the group in the portal" in leh2
      and "rto_codes" not in mod.EXTRA_SHEETS["RTO groups to create"].columns and mod.REGION_HINTS == {}, leh2[:200])
build(oreal(), f=FO, t=TO)                                                # the whole grid: its hints are in force
check("the hints also fold the grid's spelling into the earlier upload's for the comparison: LEH & LAHDHK = RELIANCE_SEPT26_Ladakh, TRIVANDRUM = RELIANCE_SEPT26_Thiruvananthapuram", mod.rto_key("LEH & LAHDHK") == mod.rto_key("RELIANCE_SEPT26_Ladakh")
      and mod.rto_key("TRIVANDRUM") == mod.rto_key("RELIANCE_SEPT26_Thiruvananthapuram"))

# ---- the INCOMPLETE line: a half-read grid never passes for a whole one
def break_header(wb):
    wb["Oct 26"]["I4"] = "COMP + RSA"


d3, w3, a3 = build(omini(tweak=break_header), f=FO, t=TO)
check("one unreadable header (a word nobody knows) is REFUSED, the other 29 columns are written, and the first warning says INCOMPLETE: 1 of 30 payout columns were not read (column I)",
      w3 and w3[0].startswith("INCOMPLETE: 1 of 30 payout columns") and "column(s) I" in w3[0] and len(d3) == 8 * 117 and "COMP + RSA" in " ".join(w3), (w3[:2], len(d3)))
mod.RUN_WARNINGS.clear()
mod.report(d3, w3, quiet, mod.Asker(interactive=False))
check("the INCOMPLETE line is the FIRST warning of the run's report (the run log, and the review Summary which lists RUN_WARNINGS in order)", bool(mod.RUN_WARNINGS) and mod.RUN_WARNINGS[0].startswith("INCOMPLETE: 1 of 30 payout columns"), mod.RUN_WARNINGS[:2])
mod.RUN_WARNINGS.clear()
check("a grid with every column read has no INCOMPLETE line", not any(x.startswith("INCOMPLETE") for x in w))

# ---- against the user's October export (92): the answer key
d4 = dfo.copy()
lines = []
mod.EXTRA_SHEETS.clear()
mod.RTO_NAMES = "cluster"
build(oreal(), f=FO, t=TO)                                                # the hints of this grid are in force for the comparison
got = mod.diff_against_reference(d4, OCT_REF, lines.append)
check("against the user's October export: 0 new rows, 181 removed (their 30 M & HCV rows, 72 nil-dep DIESEL taxi rows, 79 earlier Private Car rows), 0 changed: every payin and payout of the 8,784 rows is theirs",
      got == (0, 181, 0), (got, lines[:3]))
check("the removed rows are named: 30 x GCV 4W 1+1 (M & HCV), 72 x PCV 4W 1+1 (nil dep diesel), 64 / 5 / 10 Private Car", any("30 x GCV 4W 1+1" in x and "72 x PCV 4W 1+1" in x and "Private Car" in x for x in lines), lines[-4:])
pairs = mod.EXTRA_SHEETS["Other names vs earlier upload"]
check("4,608 rows pair by the card but name it another way (listed): Jeeto / Supro groups 2 x 720, 'JCB,Caterpillar' 720, PCV_4W_ALL for BUS_ALL 720 and for TAXI_ALL 720 (Kaali Peeli) + 1,008 (with the T&C 'Non,Kalli Pilli')",
      int(pairs["rows"].sum()) == 4608 and sorted(pairs["rows"]) == [720, 720, 720, 720, 720, 1008] and set(pairs["makes_earlier_upload"]) >= {"TATA,Maruti,JEETO_SUPRO", "GCV_4W_ALL_EX_TATA_MARUTI_JEETO_SUPRO", "PCV_4W_ALL"}
      and "Non,Kalli Pilli" in set(pairs["tnc_earlier_upload"]) and "JCB,Caterpillar" in set(pairs["tnc_earlier_upload"]), pairs.to_dict("records")[:2])
check("... and a region that is only spelt another way ('LEH & LAHDHK' / Ladakh) is not a difference: no region is listed as missing but the earlier Private Car 'ALLINDIA'", any("regions in the earlier upload but not in this grid (1): ['allindia']" in x for x in lines) and not any("regions not in the earlier upload" in x for x in lines), lines[-3:])

# ---- the command line on the October grid: the dates the person gives, the earlier upload as the answer key, the review workbook
outo = os.path.join(tmp, "out", "INDUSIND_UPLOAD_2026-10.csv")
p = subprocess.run([sys.executable, "-I", "-X", "utf8", os.path.join(ROOT, "indusind_grid_to_config.py"), "--grid", OCT, "--from-date", "13/10/2026", "--to-date", "31/10/2026", "--ref", OCT_REF, "--out", outo, "--no-prompt"],
                   cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
check("the command line converts the October grid: exit 0, 30 columns read, 0 refused, the comparison says 0 new / 181 removed / 0 changed", p.returncode == 0 and "30 read, 0 refused" in p.stdout and "0 new rows, 181 removed rows, 0 rows with changed values" in p.stdout, (p.returncode, p.stdout[-400:], p.stderr[-300:]))
csvo = pd.read_csv(outo, dtype=str, keep_default_na=False)
check("the file written is the upload: 42 columns in the portal's order, 8,784 rows, the dates given to the run", list(csvo.columns) == mod.OUTPUT_COLUMNS and len(csvo) == 8784 and set(csvo["from_date"]) == {"13/10/2026"} and set(csvo["to_date"]) == {"31/10/2026"})
booko = pd.read_excel(outo.replace("UPLOAD", "REVIEW").replace(".csv", ".xlsx"), sheet_name=None, dtype=str, keep_default_na=False)
check("the review workbook has 'Other names vs earlier upload', 'Not written', 'RTO groups to create' (with the RTO codes), 'Unread parts' (the RTO List 'left out on purpose', the make / model list NOT READ) and the Summary",
      {"Summary", "Other names vs earlier upload", "Not written", "RTO groups to create", "Unread parts", "Manual review notes", "Notes read"} <= set(booko) and "rto_codes" in booko["RTO groups to create"].columns
      and any(s == "RTO List" and st == "left out on purpose" for s, st in zip(booko["Unread parts"]["sheet"], booko["Unread parts"]["status"])), list(booko))

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
