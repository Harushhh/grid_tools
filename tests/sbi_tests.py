"""Deterministic checks of sbi_grid_to_config.py: the text scanner (every rule found by its words, reworded notes, a note that has gone), the header-driven reading (the same file shifted /
renamed / with a new column), the slab (none for two-wheelers), the fuels a note declines ('In PCV 3W EV Decline'), the NCB rules, High-End, the AP 'non doable RTOs' group, the declined-cluster
cross-check, the notes that are flagged (short ones too), and the real October grid against the user's own upload.
Run:  .venv\\Scripts\\python.exe tests\\sbi_tests.py"""
import datetime as dt
import os
import sys

import numpy as np
import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("sbi_grid_to_config.py")
F, T = dt.date(2026, 10, 3), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


SRC = os.path.join(helpers.INPUTS, "sbi_oct26.xlsx")
SRC_WB = openpyxl.load_workbook(SRC, data_only=True)
S0 = SRC_WB["GCV _PCV _Pvt Car_TW"]
SRC_ROW = {str(S0.cell(r, 3).value).strip(): r for r in range(6, 60)}
COL = openpyxl.utils.column_index_from_string


def src(cluster, col):
    return S0.cell(SRC_ROW[cluster], COL(col)).value


# ================================================================ a small workbook made of the real one
CLUSTERS = ("MH - M", "AP - Rest", "DL - NCR", "TN - C", "CH - Rest", "GA")


def mini(clusters=CLUSTERS, row_off=0, col_off=0, rate_name="GCV _PCV _Pvt Car_TW", tweak=None, notes=True):
    ws = openpyxl.Workbook()
    out = ws.active
    out.title = rate_name
    for r in range(1, 6):                                          # the five header rows
        for c in range(1, S0.max_column + 1):
            v = S0.cell(r, c).value
            if v is not None:
                out.cell(r + row_off, c + col_off, v)
    for mr in S0.merged_cells.ranges:
        if mr.max_row <= 5:
            out.merge_cells(start_row=mr.min_row + row_off, start_column=mr.min_col + col_off, end_row=mr.max_row + row_off, end_column=mr.max_col + col_off)
    r_out = 6 + row_off
    for cl in clusters:
        for c in range(1, S0.max_column + 1):
            v = S0.cell(SRC_ROW[cl], c).value
            if v is not None:
                out.cell(r_out, c + col_off, v)
        r_out += 1
    if notes:
        r_out += 1
        for r in range(61, S0.max_row + 1):
            for c in range(1, S0.max_column + 1):
                v = S0.cell(r, c).value
                if v is not None:
                    out.cell(r_out + (r - 61), c + col_off, v)
    # the other sheets
    he = SRC_WB["PV SAOD- HIGHEND Enabler"]
    w = ws.create_sheet("PV SAOD- HIGHEND Enabler")
    for i, row in enumerate(he.iter_rows(values_only=True), 1):
        if i <= 4 or (row[0] in clusters):
            w.append(list(row))
    mp = SRC_WB["RTO Cluster Mapping"]
    w = ws.create_sheet("RTO Cluster Mapping")
    for i, row in enumerate(mp.iter_rows(values_only=True), 1):
        if i == 1 or (row[3] in clusters):
            w.append(list(row))
    for name in ("PV SATP Decline make and Model", "Car, Taxi Declined list"):
        w = ws.create_sheet(name)
        for row in SRC_WB[name].iter_rows(values_only=True):
            w.append(list(row))
    ws.create_sheet("Sheet1").sheet_state = "hidden"
    if tweak:
        tweak(ws)
    return ws


def note_cell(wb, pattern):
    """the cell (row, col) of the rate sheet whose text matches"""
    import re
    sh = wb.worksheets[0]
    for row in sh.iter_rows():
        for c in row:
            if isinstance(c.value, str) and re.search(pattern, c.value, re.I | re.S):
                return sh, c
    raise AssertionError(f"no note like {pattern}")


LAST_ASK = None                                                    # the Asker of the last build: its `decisions` are what the run told the person


def build(wb, slab=("1L-3L", "3L-4.99L"), names="group", answers=None, interactive=False, f=F):
    global LAST_ASK
    mod.SLAB_ARGS["pc"], mod.SLAB_ARGS["cv"] = slab
    mod.RTO_NAMES = names
    out = []
    ask = LAST_ASK = mod.Asker(interactive=True, input_fn=lambda prompt: next(answers), out=quiet) if answers is not None else mod.Asker(interactive=interactive)
    try:
        df, w = mod.build(None, f, T, quiet, ask, wb=wb)
    finally:
        mod.SLAB_ARGS["pc"] = mod.SLAB_ARGS["cv"] = None
        mod.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")
    return df, w


def sel(df, **kw):
    m = pd.Series(True, index=df.index)
    for k, v in kw.items():
        m &= (df[k] == v)
    return df[m]


def rate(df):
    return float(df[["pi_net_rate", "pi_tp_rate", "pi_od_rate"]].astype(float).max(axis=1).iloc[0])


G = lambda c: "SBI_OCT26_" + c.replace(" ", "_")

# ================================================================ pure helpers
b = mod.bands_from_labels(["Upto 999 CC", "1000 CC-1499 CC", "Above 1500 CC"])
check("cc bands of the taxi labels are contiguous: 0-999.01, 999.01-1499.01, 1499.01-open", b == {"Upto 999 CC": (0.0, 999.01), "1000 CC-1499 CC": (999.01, 1499.01), "Above 1500 CC": (1499.01, 99999)}, b)
b = mod.bands_from_labels(["C. 1500+", "A. 0-1000", "B. 1001-1500"])
check("cc bands of the car SATP labels (given out of order): 0-1000.01, 1000.01-1500.01, 1500.01-open", b == {"A. 0-1000": (0.0, 1000.01), "B. 1001-1500": (1000.01, 1500.01), "C. 1500+": (1500.01, 99999)}, b)
names = ["CH - R", "CH - Rest", "UP - GV", "UP - Rest 1", "UP - Rest 2", "North East 1", "AS", "GA"]
got = mod.clusters_in("Upto 5 Year:-CH - Rest, AS, UP GV, North East 1, UP - Rest 2.", names)
check("a note's cluster list: 'CH - R' is not read inside 'CH - Rest', the 'UP GV' typo is UP - GV, 'AS' is a name only as a whole word", got == {"CH - Rest", "AS", "UP - GV", "North East 1", "UP - Rest 2"}, got)
check("'as per' is not the cluster AS (case-sensitive)", mod.clusters_in("rates as per the grid, GA allowed", names) == {"GA"})
sec = mod.decline_sections('For "GCV 3W" Declined RTO Cluster:- Upto 5 Year:-TN - Rest, AS. Above 5 Year :-TN - Rest (Rest All Cluster Allowed)')
check("a decline note lists clusters per age: 'Upto 5 Year' (new + mid) and 'Above 5 Year' (old)", len(sec) == 2 and sec[0][0] == {"new", "mid", "upto5"} and sec[1][0] == {"old"}, [s[0] for s in sec])
check("without an age heading a note holds one list for every age", mod.decline_sections("Declin Cluster :- AS,CH - R")[0][0] == mod.ALL_AGES)
segs, states = mod.note_scope("additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only", {"DL", "MH", "TN"})
check("note scope: '2.5 to 3.5' without the 't' is the 2.5-3.5 t segment, 'DL' the Delhi clusters", segs == {"gcv_25_35"} and states == {"DL"}, (segs, states))
segs, states = mod.note_scope("In the 12T to 40T & 20T to 40T segment, the TN-C payout grid is applicable for RTOs TN20, TN54", {"DL", "MH", "TN"})
check("note scope: '12T to 40T' covers the 12-20 t and 20-40 t segments; TN20 names the TN clusters", segs == {"gcv_12_20", "gcv_20_40"} and states == {"TN"}, (segs, states))
segs, states = mod.note_scope("a completely new rule nobody has seen", {"DL"})
check("a note that names no segment and no state may affect every row", segs is None and states is None)
check("cluster prefixes: 'MH - M' -> MH, 'GA' -> GA, 'North East 1' -> NE", [mod.cluster_prefix(x) for x in ("MH - M", "GA", "North East 1")] == ["MH", "GA", "NE"])

# ================================================================ the text scanner on the real notes
wb = mini()
n0, h0, c0, _ = mod.rate_sheet(wb, mod.Asker(), [], quiet)
table = mod.RateTable(wb, n0, h0, c0)
notes = mod.collect_notes(table)
R = mod.read_rules(notes, mod.Asker(), [], quiet)
check("rules: general GCV ages 20 / 25, Private Car 15 / 25, PCV 3W 15 / 25, taxi 15 / 20, bus 25, TW 10 / 15, tractor 25",
      (R["age.gcv"], R["age.pc.comp"], R["age.pc.satp"], R["age.pcv3w"], R["age.taxi"], R["age.bus"], R["age.tw.comp"], R["age.tw.satp"], R["age.tractor"]) == ((20, 25), 15, 25, (15, 25), (15, 20), 25, 10, 15, 25), R)
check("rules: Private Car non-NCB cut 5, High-End non-NCB zero, taxi non-NCB cut 3, car SATP (3, ages 1-10, below 1 zero), TW cc 500",
      (R["pc.noncb_cut"], R["he.noncb_zero"], R["taxi.noncb_cut"], R["pc.satp"], R["tw.cc_max"]) == (5.0, True, 3.0, (3.0, 1, 10, 1), 500), R)
check("rules: the two-wheeler makes (RE = ROYAL, 'KTM Honda' without a comma), the races T&C, 'Electric declined for all GCV', the six non-doable AP RTOs",
      R["tw.makes"] == ["HERO", "HONDA", "YAMAHA", "TVS", "SUZUKI", "ROYAL", "BAJAJ", "KTM", "ATHER"] and R["tw.makes_unknown"] == [] and R["tw.tnc"].startswith("Vehicles participating in races")
      and R["gcv.no_electric"] is True and R["ap.nondoable"] == ["AP02", "AP03", "AP04", "AP21", "AP26", "AP27"], (R["tw.makes"], R["ap.nondoable"]))
check("rules: the fuels the notes decline are read by ONE reader (October: only GCV Electric) and 'In PCV 3W EV Decline' is not in the October grid", R["fuel_declines"] == {"gcv": {"Electric"}}, R["fuel_declines"])
check("rules: 'additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only' is read as +2 points, state DL, 2.5 to 3.5 t", [(u["pts"], u["state"], u["lo"], u["hi"]) for u in R["uplifts"]] == [(2.0, "DL", 2.5, 3.5)], R["uplifts"])
check("rules: the grid's own period is 03/10/2026 - 31/10/2026", R["period"] == (dt.date(2026, 10, 3), dt.date(2026, 10, 31)), R["period"])
check("rules: both slab notes are read (Pvt Car 1L-3L -1, 50K-1L -2, 15K-50K -3; CV 3L-4.99L -1, 1L-2.99L -2, 15K-99K -3) and the '<15K' IRDA slab is named",
      R["slabs.pc"] == [("1L-3L", 1.0), ("50K-1L", 2.0), ("15K-50K", 3.0)] and R["slabs.cv"] == [("3L-4.99L", 1.0), ("1L-2.99L", 2.0), ("15K-99K", 3.0)] and R["slab_irda.pc"] == "15K" and R["slab_irda.cv"] == "15K", (R["slabs.pc"], R["slabs.cv"]))
decl = mod.read_declines(notes, [cl for _, cl, _ in table.rows])
check("the 'Declined cluster' notes are read (GCV 3W, up to 2.5 t, 2.5-3.5 t, 12-20, 20-40, above 40 allowed-list, PCV 3W, taxi, car SATP x2, tractor)",
      {d["seg"] for d in decl} == {"gcv3w", "gcv_u25", "gcv_25_35", "gcv_12_20", "gcv_20_40", "gcv_40", "pcv3w", "taxi", "pc_satp_p", "pc_satp_d", "miscd"}, sorted(d["seg"] for d in decl))

# reworded numbers follow the note; a note that has gone is ASKED about
def reword(old, new):
    def tweak(w):
        sh, c = note_cell(w, old)
        c.value = c.value.replace(*new) if isinstance(new, tuple) else new
    return tweak


w2 = mini(tweak=reword(r"reduction of 5%", ("reduction of 5%", "reduction of 6%")))
df6, _ = build(w2)
mh_no = sel(df6, new_rto_list=G("MH - M"), product="Private Car", sub_product="1+1", make_model_class="Car", new_vehicle_makes=mod.MK_PC, is_ncb="NO", fuel_type="Petrol")
check("a note that says 6% instead of 5% makes the non-NCB payin 30 - 1 - 6 = 23 (nothing is hard-coded)", rate(mh_no) == 23.0, rate(mh_no))
w3 = mini(tweak=reword(r"PCV Taxi Comp", ("Upto15 year", "Upto 12 year")))
df12, _ = build(w3)
check("a note that says 12 years instead of 15 for taxi comprehensive gives age 0-12.01", set(sel(df12, product="PCV 4W", make_model_class="Taxi", sub_product="1+1").vehicle_age_max) == {12.01})
w4 = mini(tweak=reword(r"A reduction of 5%", "(the rule has been removed)"))
try:
    build(w4)
    stopped = ""
except mod.NeedsAnswer as e:
    stopped = str(e)
check("a rule whose note has gone stops with a question when nothing can be asked", "Private Car non-NCB reduction" in stopped, stopped)
dfq, wq = build(w4, answers=iter(["1"]))
mh_no2 = sel(dfq, new_rto_list=G("MH - M"), product="Private Car", sub_product="1+1", make_model_class="Car", new_vehicle_makes=mod.MK_PC, is_ncb="NO", fuel_type="Petrol")
check("... and with 'use last month's value' it goes on with 5 and says so", rate(mh_no2) == 24.0 and any("was not found" in x and "last month's value (5.0) is used" in x for x in wq), (rate(mh_no2), wq[:2]))

w5 = mini(tweak=reword(r"^Pvt Car Slab", "Pvt Car Slab- Agency: the discount structure changed, ask your relationship manager"))
try:
    build(w5, slab=(None, "3L-4.99L"))
    q = ""
except mod.NeedsAnswer as e:
    q = str(e)
check("a slab note that cannot be read is never taken for 'no slab': it stops with a question", "slab note has no slab this script can read" in q, q)
d5, w5w = build(w5, slab=("none", "3L-4.99L"))
check("... unless --pc-slab none says the grid is to be written as printed", rate(sel(d5, new_rto_list=G("MH - M"), product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")) == 30.0)
w6 = mini(tweak=reword(r"Electric declined for all GCV", "Fuel types: the GCV note about electric vehicles has been removed"))
try:
    build(w6)
    q = ""
except mod.NeedsAnswer as e:
    q = str(e)
check("the note that Electric is declined for all GCV is a rule: when it has gone the script asks instead of writing Electric rows", "Electric vehicles are declined for all GCV" in q, q)

# ================================================================ the October values, checked against the workbook cells themselves
df, warns = build(mini())
check("a build of six clusters has no duplicate keys and the po is pi - 7", not df.duplicated(mod.KEY_COLUMNS).any() and ((df.pi_net_rate.astype(float) - df.po_net_rate.astype(float)).between(6.99, 7.01) | (df.po_net_rate.astype(float) == 0)).all())
mh = G("MH - M")
e34 = sel(df, new_rto_list=mh, product="GCV 4W", sub_product="1+1", make_model_class="Truck", new_vehicle_makes=mod.MK_BELOW35, fuel_type="Diesel", vehicle_age_min=0, vehicle_age_max=1)
check("GCV up to 2.0 t, new, comprehensive: the cell minus the CV slab (65 -> 64), On Net, add_tnc New, po 57", rate(e34) == src("MH - M", "E") - 1 and e34.iloc[0].pi_type == "On Net" and e34.iloc[0].add_tnc == "New" and float(e34.iloc[0].po_net_rate) == 57.0, e34.iloc[0].to_dict())
tata = sel(df, new_rto_list=mh, product="GCV 4W", sub_product="STP", make_model_class="Trailer", new_vehicle_makes="TATA", fuel_type="CNG", cc_min=2500.01, cc_max=3000.01, vehicle_age_min=5.01)
check("GCV 2.5-3.0 t Tata, SATP above 5 years: the cell minus 1, On TP, age to 25.01 (the October upload had 20.01)", rate(tata) == src("MH - M", "AH") - 1 and tata.iloc[0].pi_type == "On TP" and tata.iloc[0].vehicle_age_max == 25.01, tata.iloc[0].to_dict())
big = sel(df, new_rto_list=mh, product="GCV 4W", new_vehicle_makes="Mahindra", cc_min=40000.01)
check("above 40 t of a cluster the grid leaves blank: a ZERO card per age group, class and fuel (not dropped)", len(big) == 2 * 2 * 5 * 4 and (big[["pi_net_rate", "pi_tp_rate"]].astype(float).sum(axis=1) == 0).all(), len(big))
dec = sel(df, new_rto_list=mh, product="GCV 4W", new_vehicle_makes=mod.MK_GCV_ALL)
check("3.5-12 t ('Pan India decline'): GCV_4W_ALL zero cards in the three weight bands", set(dec.cc_min) == {3500.01, 5000.01, 7500.01} and (dec.pi_net_rate.astype(float) + dec.pi_tp_rate.astype(float) == 0).all())
check("GCV rows have the four non-electric fuels (the note 'Electric declined for all GCV') and the five classes", set(sel(df, product="GCV 4W").fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"} and len(set(sel(df, product="GCV 4W", make_model_class="Truck").make_model_class)) == 1
      and set(sel(df, product="GCV 4W").make_model_class) == {"Pick Up Van", "Tanker", "Dumper Tipper", "Trailer", "Truck", "Goods Carrying Tractor"})
g3 = sel(df, new_rto_list=mh, product="GCV 3W", sub_product="1+1", vehicle_age_min=1, fuel_type="Diesel")
check("GCV 3W (a block inside the first block of the sheet): upto-5-year comprehensive = the cell - 1", rate(g3) == src("MH - M", "S") - 1 and g3.iloc[0].make_model_class == "Goods Carrying Rickshaw", g3.iloc[0].to_dict())
trc = sel(df, new_rto_list=mh, product="GCV 4W", make_model_class="Goods Carrying Tractor", sub_product="STP", fuel_type="Petrol")
check("GCV tractor (one 'Comp /SATP' column) serves 1+1 and STP", rate(trc) == src("MH - M", "AO") - 1 and len(sel(df, new_rto_list=mh, make_model_class="Goods Carrying Tractor")) == 8, trc.iloc[0].to_dict())
pcy = sel(df, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")
pcn = sel(df, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="NO", fuel_type="Petrol")
check("Private Car comprehensive 'other than High-End': NCB yes = cell - Pvt Car slab (30 -> 29), no = 5 less (24), On OD, age 0-15.01", rate(pcy) == src("MH - M", "BY") - 1 and rate(pcn) == src("MH - M", "BY") - 6
      and pcy.iloc[0].pi_type == "On OD" and pcy.iloc[0].vehicle_age_max == 15.01, (rate(pcy), rate(pcn)))
pcs = sel(df, new_rto_list=mh, product="Private Car", sub_product="SAOD", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Diesel")
check("Private Car SAOD takes the SAOD column (27 -> 26)", rate(pcs) == src("MH - M", "BZ") - 1, rate(pcs))
satp = lambda **kw: sel(df, new_rto_list=mh, product="Private Car", sub_product="STP", **kw)
old, mid, under = satp(fuel_type="Petrol", cc_max=1000.01, vehicle_age_min=10.01), satp(fuel_type="Petrol", cc_max=1000.01, vehicle_age_min=1), satp(fuel_type="Petrol", cc_max=1000.01, vehicle_age_min=0)
check("Private Car SATP: 10+ years = cell - 1, ages 1-10 = 3 less again, below 1 year = ZERO (the grid's own note)",
      rate(old) == src("MH - M", "DB") - 1 and rate(mid) == src("MH - M", "DB") - 4 and rate(under) == 0.0 and (old.iloc[0].vehicle_age_max, mid.iloc[0].vehicle_age_max, under.iloc[0].vehicle_age_max) == (25.01, 10.01, 1.0),
      (rate(old), rate(mid), rate(under)))
check("Private Car SATP fuels: Petrol / Hybrid / EV = Petrol + Electric; Diesel / CNG / LPG = those three; cc bands 0-1000.01, 1000.01-1500.01, 1500.01-open",
      set(satp(cc_max=1000.01, vehicle_age_min=10.01).fuel_type) == {"Petrol", "Electric", "Diesel", "CNG", "LPG"} and sorted(set(satp().cc_max)) == [1000.01, 1500.01, 99999] and set(satp().new_vehicle_makes) == {mod.MK_PC_TP})
pv = sel(df, new_rto_list=mh, product="PCV 3W", sub_product="1+1", fuel_type="Petrol", vehicle_age_min=1)
check("PCV 3W: non-diesel non-new comprehensive = CC - 1, add_tnc Old, sc 0-4.01 ('3+1'), age 1-15.01", rate(pv) == src("MH - M", "CC") - 1 and pv.iloc[0].add_tnc == "Old" and pv.iloc[0].sc_max == 4.01 and pv.iloc[0].vehicle_age_max == 15.01, pv.iloc[0].to_dict())
tx = lambda ncb, zd, cc: sel(df, new_rto_list=mh, product="PCV 4W", make_model_class="Taxi", sub_product="1+1", is_ncb=ncb, is_zd=zd, cc_min=cc, fuel_type="Petrol")
check("Taxi: non-nil-dep with NCB = cell - 1, nil-dep = its own column, no NCB = 3 less; sc 0-7.01 ('6+1 & 4+1'); SATP age to 20.01",
      rate(tx("YES", "NO", 0)) == src("MH - M", "CJ") - 1 and rate(tx("YES", "YES", 0)) == src("MH - M", "CI") - 1 and rate(tx("NO", "NO", 0)) == src("MH - M", "CJ") - 4 and tx("YES", "NO", 0).iloc[0].sc_max == 7.01
      and set(sel(df, product="PCV 4W", make_model_class="Taxi", sub_product="STP").vehicle_age_max) == {20.01}, rate(tx("YES", "NO", 0)))
sb = sel(df, new_rto_list=mh, product="PCV 4W", make_model_class="School Bus", sub_product="1+1", fuel_type="Petrol")
check("School bus: 12-17 seats = 12.01-18.01, 18+ = 18.01-999; the cells minus 1; age to 25.01", sorted(zip(sb.sc_min, sb.sc_max, sb.pi_net_rate.astype(float))) == [(12.01, 18.01, src("MH - M", "DJ") - 1), (18.01, 999, src("MH - M", "DH") - 1)], sb[["sc_min", "sc_max", "pi_net_rate"]])
tw = lambda sub, k: sel(df, new_rto_list=mh, product="TW", sub_product=sub, make_model_class=k, fuel_type="Petrol")
check("Two-wheeler: the cells AS PRINTED (no slab: 30, 30, 30), SAOD = its column (On OD), 1+1 and STP share the '1+1 & SATP' column; cc 0-500.01; ages 10.01 / 10.01 / 15.01; the races T&C", rate(tw("SAOD", "BIKE")) == src("MH - M", "CR") and rate(tw("1+1", "SCOOTER")) == src("MH - M", "CU")
      and rate(tw("STP", "SCOOTER")) == src("MH - M", "CU") and src("MH - M", "CR") > 1 and tw("SAOD", "BIKE").iloc[0].pi_type == "On OD" and tw("1+1", "BIKE").iloc[0].cc_max == 500.01
      and (tw("1+1", "BIKE").iloc[0].vehicle_age_max, tw("SAOD", "BIKE").iloc[0].vehicle_age_max, tw("STP", "BIKE").iloc[0].vehicle_age_max) == (10.01, 10.01, 15.01) and tw("1+1", "BIKE").iloc[0].add_tnc.startswith("Vehicles participating in races"))
check("Two-wheeler makes: bike group has ROYAL, scooter group does not", "ROYAL" in tw("1+1", "BIKE").iloc[0].new_vehicle_makes and "ROYAL" not in tw("1+1", "SCOOTER").iloc[0].new_vehicle_makes and set(sel(df, product="TW").fuel_type) == {"Petrol", "Electric"})
mt = lambda klass, sub, tnc: sel(df, new_rto_list=mh, product="MISCD", make_model_class=klass, sub_product=sub, add_tnc=tnc, fuel_type="Petrol")
check("Tractor / harvester: new / old comprehensive and SATP; the harvester carries 'Harvester,EX Trailer'", rate(mt("MISCD-Tractor", "1+1", "New")) == src("MH - M", "CV") - 1 and rate(mt("MISCD-Others", "1+1", "Old,Harvester,EX Trailer")) == src("MH - M", "CY") - 1
      and rate(mt("MISCD-Tractor", "STP", "")) == src("MH - M", "CZ") - 1 and set(sel(df, product="MISCD").fuel_type) == set(mod.FUELS_ALL))

# ================================================================ the slab (asked; any answer follows the note)
dfn, _ = build(mini(), slab=("none", "none"))
check("slab 'none' writes the grid as printed (E 65, BY 30, CR 30)", rate(sel(dfn, new_rto_list=mh, product="GCV 4W", sub_product="1+1", make_model_class="Truck", new_vehicle_makes=mod.MK_BELOW35, fuel_type="Diesel", vehicle_age_min=0, vehicle_age_max=1)) == 65.0
      and rate(sel(dfn, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")) == 30.0 and rate(sel(dfn, new_rto_list=mh, product="TW", sub_product="SAOD", make_model_class="BIKE", fuel_type="Petrol")) == 30.0)
dfs, _ = build(mini(), slab=("50K-1L", "1L-2.99L"))
tw_saod = lambda frame: rate(sel(frame, new_rto_list=mh, product="TW", sub_product="SAOD", make_model_class="BIKE", fuel_type="Petrol"))
check("slab 50K-1L (Private Car) and 1L-2.99L (CV) take 2 points each, in their own family - and the two-wheeler stays AS PRINTED (30): the Pvt Car slab is not its slab (the user, 3 Nov 2026)",
      rate(sel(dfs, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")) == 28.0 and tw_saod(dfs) == 30.0
      and rate(sel(dfs, new_rto_list=mh, product="PCV 4W", make_model_class="School Bus", sub_product="1+1", fuel_type="Petrol", sc_min=18.01)) == 63.0)
dfm, _ = build(mini(), slab=("1L-3L", "none"))
check("the families are independent: CV with no slab (65), Private Car with 1L-3L (29), the two-wheeler as printed (30) whatever the Private Car slab is", rate(sel(dfm, new_rto_list=mh, product="GCV 4W", sub_product="1+1", make_model_class="Truck", new_vehicle_makes=mod.MK_BELOW35, fuel_type="Diesel", vehicle_age_min=0, vehicle_age_max=1)) == 65.0
      and rate(sel(dfm, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")) == 29.0 and tw_saod(dfm) == 30.0 and tw_saod(dfn) == 30.0 and tw_saod(df) == 30.0)
check("... every two-wheeler row (1+1, SAOD, STP, bike and scooter) equals its grid cell with the Private Car slab at 1L-3L, 50K-1L and none: not one row is lowered",
      all(tw_saod(f_) == 30.0 for f_ in (df, dfs, dfm, dfn)) and all(rate(sel(df, new_rto_list=mh, product="TW", sub_product=s, make_model_class=k, fuel_type="Petrol")) == rate(sel(dfn, new_rto_list=mh, product="TW", sub_product=s, make_model_class=k, fuel_type="Petrol"))
          for s in ("1+1", "SAOD", "STP") for k in ("BIKE", "SCOOTER")))
check("... and the run tells the person, in its decisions (they go to the report), that the two-wheelers have no slab", any("no premium slab for two-wheelers" in d for d in LAST_ASK.decisions), LAST_ASK.decisions)
d1, _ = build(mini(), answers=iter(["2", "2"]), slab=(None, None))
check("the slab is ASKED when no flag is given: option 2 and 2 = 1L-3L and 3L-4.99L", rate(sel(d1, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")) == 29.0)
try:
    build(mini(), slab=(None, None))
    q = ""
except mod.NeedsAnswer as e:
    q = str(e)
check("... and with no keyboard it stops with that question - which is about the Private Car grid only, never the two-wheeler (the user, 3 Nov 2026)", "premium slab for the Private Car grid" in q and "Two-Wheeler" not in q and "TW" not in q, q)
try:
    build(mini(), slab=("9L", "none"))
    q = ""
except mod.GridError as e:
    q = str(e)
check("a slab that is not in the note is refused with the names that are", "9L" in q and "1L-3L" in q, q)
# a payout of zero stays zero, a small one is floored at zero
ws_z = mini(tweak=lambda w: setattr(w.worksheets[0].cell(7 + 0, 5 + 0), "value", 0.5))
dz, _ = build(ws_z)
z = sel(dz, new_rto_list=G("AP - Rest"), product="GCV 4W", sub_product="1+1", make_model_class="Truck", new_vehicle_makes=mod.MK_BELOW35, fuel_type="Diesel", vehicle_age_min=0, vehicle_age_max=1)
check("a payout below the slab (0.5 - 1) is floored at 0, never negative", rate(z) == 0.0)

# ================================================================ the fuels a note declines (3 Nov 2026: the November grid says 'In PCV 3W EV Decline'; the Electric rows were written anyway)
class _N:                                                          # the least a note needs: its label, the text a rule has not used up
    def __init__(self, label, text):
        self.label, self.rest, self.used = label, text, False


def declines(label, text):
    n = _N(label, text)
    return mod.read_fuel_declines([n]), n


applied = [("PCV 3W", "In PCV 3W EV Decline", {"pcv3w": {"Electric"}}), ("", "In PCV 3W EV Decline", {"pcv3w": {"Electric"}}), ("PCV 3W", "EV Decline", {"pcv3w": {"Electric"}}),
           ("PCV Taxi", "Electric vehicles are not allowed", {"taxi": {"Electric"}}), ("GCV", "For GCV Electric vehical declined", {"gcv": {"Electric"}}),
           ("GCV", "( Fuel type Electric declined for all GCV )", {"gcv": {"Electric"}}), ("School Bus", "Decline CNG for School Bus", {"bus": {"CNG"}}), ("Tw", "EV declined in TW", {"tw": {"Electric"}}),
           ("Tractor & Harvester (excluding Trailer)", "Electric not allowed", {"miscd": {"Electric"}}), ("Pvt Car SATP", "Diesel is declined", {"pc_satp": {"Diesel"}})]
wrong = [(lab, txt, declines(lab, txt)[0]) for lab, txt, want in applied if declines(lab, txt)[0] != want]
check("the fuel-decline reader applies 'In PCV 3W EV Decline' (also without a label, or with the product only in the label), 'Electric vehicles are not allowed', 'Decline CNG for School Bus' ... for ANY product and fuel",
      not wrong and len(applied) == 10, wrong)
refused = [("PCV 3W", "Mahindra EV declined"), ("Pvt Car SATP", "Declin Cluster (Petrol / Hybrid / EV) :- AS,CH - R,CH - Rest"), ("Pvt Car SATP", "Declin Cluster (Diesel / CNG / LPG) :- AS,CH - R"), ("", "EV Decline"),
           ("PCV 3W", "In PCV 3W EV Decline in Delhi"), ("PCV 3W", "EV business is not allowed in Chennai"), ("Tw", "All Imported vehicles Decline"), ("PCV 3W", "PCV 3W Decline Cluster :- TN - Rest, North East 2, AS")]
leaked = [(lab, txt) for lab, txt in refused if declines(lab, txt)[0] or declines(lab, txt)[1].used or declines(lab, txt)[1].rest != txt]
check("... and never guesses: a make before the fuel ('Mahindra EV declined'), a place after it, the cluster lists '(Petrol / Hybrid / EV)', or no product at all are NOT read (they stay unread notes, flagged)", not leaked and len(refused) == 8, leaked)
n_ = declines("PCV 3W", "In PCV 3W EV Decline")[1]
check("... a note that was read is used up completely (nothing of 'In PCV 3W EV Decline' is left to be flagged as a fragment)", n_.used and n_.rest.strip() == "", n_.rest)


def pcv3w_ev_note(w):
    sh = w.worksheets[0]
    sh.cell(sh.max_row + 1, 3, "PCV 3W")
    sh.cell(sh.max_row, 4, "In PCV 3W EV Decline")


dev, wev = build(mini(tweak=pcv3w_ev_note))
p3 = sel(dev, product="PCV 3W")
p3_before = sel(df, product="PCV 3W")
check("THE NOVEMBER BUG: with the note 'In PCV 3W EV Decline' no PCV 3W row has the fuel Electric; Petrol, CNG, LPG and Diesel are written as before (only the Electric rows of 1+1 / STP x new / non-new x clusters are gone)",
      len(p3_before) > 0 and "Electric" in set(p3_before.fuel_type) and len(p3) > 0 and "Electric" not in set(p3.fuel_type) and set(p3.fuel_type) == {"Petrol", "CNG", "LPG", "Diesel"}
      and len(p3_before) - len(p3) == len(sel(p3_before, fuel_type="Electric")) == 2 * 2 * len(CLUSTERS), (len(p3_before), len(p3)))
rest_before, rest_now = df[df["product"] != "PCV 3W"], dev[dev["product"] != "PCV 3W"]
check("... nothing else changes: every other product has exactly the rows it had, Electric included (taxi, school bus, two-wheeler, tractor, Private Car), at the same rates",
      len(rest_now) == len(rest_before) and all("Electric" in set(sel(dev, product=p).fuel_type) for p in ("PCV 4W", "TW", "MISCD", "Private Car"))
      and mod.to_csv_frame(rest_now).drop(columns="remarks").reset_index(drop=True).equals(mod.to_csv_frame(rest_before).drop(columns="remarks").reset_index(drop=True)))
check("... the note is APPLIED, not flagged (no 'Manual Review Required: In PCV 3W EV Decline' on any row), and the run says what it did in its decisions",
      not dev.remarks.str.contains("EV Decline", regex=False).any() and any("Electric declined for PCV 3W" in d for d in LAST_ASK.decisions) and any("Electric declined for GCV" in d for d in LAST_ASK.decisions), LAST_ASK.decisions)
check("... a grid that still has the note gives no 'the last grid declined Electric' warning; the October grid (no note) does - the rows are written, the person is told",
      not any("declined Electric" in x for x in wev) and any("PCV 3W" in x and "declined Electric" in x and "ARE written" in x for x in warns), (wev[:2], warns[:2]))


def taxi_ev_note(w):
    sh = w.worksheets[0]
    sh.cell(sh.max_row + 1, 3, "PCV Taxi")
    sh.cell(sh.max_row, 4, "EV not allowed")


dtx, _ = build(mini(tweak=taxi_ev_note))
check("the same reader serves the next product: 'EV not allowed' under a PCV Taxi label removes Electric from the taxi rows only (no code for taxis was written)",
      "Electric" not in set(sel(dtx, product="PCV 4W", make_model_class="Taxi").fuel_type) and len(sel(dtx, product="PCV 4W", make_model_class="Taxi")) > 0
      and "Electric" in set(sel(dtx, product="PCV 3W").fuel_type) and "Electric" in set(sel(dtx, product="PCV 4W", make_model_class="School Bus").fuel_type))


def reworded_ev_note(w):                                           # a wording the reader does not know: it must be FLAGGED, never dropped (the old 25-character floor dropped it in silence)
    sh = w.worksheets[0]
    sh.cell(sh.max_row + 1, 3, "PCV 3W")
    sh.cell(sh.max_row, 4, "Only petrol cars here")


drw, _ = build(mini(tweak=reworded_ev_note))
fl = drw[drw.remarks.str.contains("Only petrol cars here", regex=False)]
check("a SHORT note no rule reads (21 characters - the shared helper drops anything under 25) is flagged on the rows it may affect, never dropped in silence",
      len(fl) > 0 and set(fl["product"]) == {"PCV 3W"} and not df.remarks.str.contains("Only petrol cars here", regex=False).any(), (len(fl), set(fl["product"])))
check("... while the headings 'Slab:-' / 'Notes:' and the fragment left of a note a rule has read ('In PCV 3W allowed') are still not flagged",
      not df.remarks.str.contains(r"Slab:-|Notes:|In PCV 3W allowed", regex=True).any() and not dev.remarks.str.contains(r"Slab:-|Notes:|In PCV 3W allowed", regex=True).any())

# ================================================================ 'additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only' (3 Nov 2026: the user added the 2 points by hand; "only for GCV 4W, cc 2500.01-3500.01")
DL_NOTE = r"DL Registration"
NO_NOTE = "All the Grids are applicable subject to meeting UW guideline"         # a standing note: the DL sentence replaced by something that changes nothing
d_no, w_no = build(mini(tweak=reword(DL_NOTE, NO_NOTE)))
PI = ["pi_od_rate", "pi_tp_rate", "pi_net_rate"]
ka, kb = mod.key_of(df), mod.key_of(d_no)
pa, pb = df[PI].astype(float).sum(axis=1).set_axis(ka), d_no[PI].astype(float).sum(axis=1).set_axis(kb)
delta = (pa - pb.reindex(pa.index)).round(4)
in_band = ((df.new_rto_list == G("DL - NCR")) & (df["product"] == "GCV 4W") & (df.cc_min >= 2500.01) & (df.cc_max <= 3500.01)).set_axis(ka)
check("the same workbook with and without the note has the same rows (the note changes a rate, never which rows exist)", len(df) == len(d_no) and set(ka) == set(kb) and in_band.sum() == 480, (len(df), len(d_no), int(in_band.sum())))
paid = in_band & (pb.reindex(pa.index) > 0)
check("DL - NCR, GCV 4W, every weight band inside 2500.01-3500.01 (2500.01-3000.01, 3000.01-3500.01, 2500.01-3500.01), Comp AND SATP: +2 points on every paid row",
      paid.sum() > 0 and set(delta[paid]) == {2.0} and set(df[in_band.values].cc_min) == {2500.01, 3000.01} and set(df[in_band.values].cc_max) == {3000.01, 3500.01}
      and set(df[in_band.values].sub_product) == {"1+1", "STP"} and {(a, b) for a, b in zip(df[in_band.values].cc_min, df[in_band.values].cc_max)} == {(2500.01, 3000.01), (3000.01, 3500.01), (2500.01, 3500.01)}, (int(paid.sum()), set(delta[paid])))
check("... and NOTHING else moves: not the other weight segments of DL - NCR, not its GCV 3W / tractor / Private Car / two-wheeler rows, not the 2.5-3.5 t rows of the other clusters",
      (~in_band).sum() > 0 and set(delta[~in_band]) == {0.0}, set(delta[~in_band]))
a_row = sel(df, new_rto_list=G("DL - NCR"), product="GCV 4W", sub_product="STP", make_model_class="Trailer", new_vehicle_makes="TATA", fuel_type="CNG", cc_min=2500.01, cc_max=3000.01, vehicle_age_min=5.01)
check("one cell read from the workbook: DL - NCR, Tata 2.5-3.0 t SATP above 5 years = the cell - the CV slab 1 + the note's 2 (and po = pi - 7); the same cell in MH - M carries no extra",
      src("DL - NCR", "AH") > 1 and rate(a_row) == src("DL - NCR", "AH") + 1 and float(a_row.iloc[0].po_tp_rate) == src("DL - NCR", "AH") - 6 and rate(tata) == src("MH - M", "AH") - 1, (rate(a_row), src("DL - NCR", "AH")))
check("the CV slab is still taken off first: with no CV slab the same row is the cell + 2",
      rate(sel(build(mini(), slab=("1L-3L", "none"))[0], new_rto_list=G("DL - NCR"), product="GCV 4W", sub_product="STP", make_model_class="Trailer", new_vehicle_makes="TATA", fuel_type="CNG", cc_min=2500.01, cc_max=3000.01, vehicle_age_min=5.01)) == src("DL - NCR", "AH") + 2)


def blank_dl_cell(w):                                              # the grid pays nothing for that Tata 2.5-3.0 t SATP cell of DL - NCR
    w.worksheets[0].cell(6 + CLUSTERS.index("DL - NCR"), COL("AH")).value = None


d_blank, _ = build(mini(tweak=blank_dl_cell))
z_row = sel(d_blank, new_rto_list=G("DL - NCR"), product="GCV 4W", sub_product="STP", make_model_class="Trailer", new_vehicle_makes="TATA", fuel_type="CNG", cc_min=2500.01, cc_max=3000.01, vehicle_age_min=5.01)
check("a blank cell stays a ZERO card: the note's points are never added to a declined entry", len(z_row) == 1 and rate(z_row) == 0.0, rate(z_row))
d3, _ = build(mini(tweak=reword(DL_NOTE, ("additional 2%", "additional 3%"))))
check("the number is the note's: 'additional 3%' gives +3 (nothing is hard-coded)", rate(sel(d3, new_rto_list=G("DL - NCR"), product="GCV 4W", sub_product="STP", make_model_class="Trailer", new_vehicle_makes="TATA", fuel_type="CNG", cc_min=2500.01, cc_max=3000.01, vehicle_age_min=5.01)) == src("DL - NCR", "AH") + 2)
build(mini())
check("the run tells the person what it did, and that the cluster also holds other states' RTOs (HR26 ..., so the whole-cluster choice stays visible)",
      any("+2 points on GCV 4W 2.5-3.5 t for DL registrations" in d and "WHOLE cluster DL - NCR" in d and "10 RTOs of other states" in d and "HR26" in d for d in LAST_ASK.decisions), LAST_ASK.decisions)


def dl_wording(text):                                              # the note in words the reader does not know / with a qualifier after it
    return reword(DL_NOTE, text)


for what, text in (("another wording", "additional 2% for DL vehicles in the 2.5 to 3.5 segment"), ("a qualifier after the sentence", "additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only but not for Tata")):
    dw, ww = build(mini(tweak=dl_wording(text)))
    kw_ = mod.key_of(dw)
    pw = dw[PI].astype(float).sum(axis=1).set_axis(kw_)
    fl_ = dw[dw.remarks.str.contains("additional 2%", regex=False)]
    check(f"{what}: NOT guessed - no points added, and the note is flagged on the Delhi 2.5-3.5 t GCV 4W rows (never applied half, never dropped)",
          set((pw - pb.reindex(pw.index)).round(4)) == {0.0} and len(fl_) > 0 and set(fl_.new_rto_list) == {G("DL - NCR")} and set(fl_["product"]) == {"GCV 4W"} and set(fl_.cc_min) <= {2500.01, 3000.01}, (what, len(fl_)))
check("... and a grid WITHOUT the note warns that the last grid gave the 2 points ('DL registrations'); the grid with it does not",
      any("DL registrations" in x and "2 extra points" in x and "written without them" in x for x in w_no) and not any("DL registrations" in x for x in warns), (w_no[:2], warns[:2]))
d_ga, w_ga = build(mini(tweak=reword(DL_NOTE, ("DL Registration", "KL Registration"))))
fl_ga = d_ga[d_ga.remarks.str.contains("KL Registration", regex=False)]
check("a state that no cluster holds (KL): not applied, warned about, and the note flagged - the rows are the no-note rows",
      any("KL registrations" in x and "no cluster holds KL RTOs" in x for x in w_ga) and len(fl_ga) > 0 and set((d_ga[PI].astype(float).sum(axis=1).set_axis(mod.key_of(d_ga)) - pb.reindex(mod.key_of(d_ga))).round(4)) == {0.0}, w_ga[:2])
other_label = _N("Pvt Car", "additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only")
check("the sentence is a GCV rule: under another product's label it is not read", mod.read_uplifts([other_label]) == [] and len(mod.read_uplifts([_N("GCV", "additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only")])) == 1)

# ================================================================ High-End
he = lambda sub, make, ncb, fuel: sel(df, new_rto_list=mh, product="Private Car", sub_product=sub, new_vehicle_makes=make, is_ncb=ncb, fuel_type=fuel)
check("High-End enabler: BMW 5 Series 20 as given (no slab) with NCB yes; NO = 0 (the note 'Zero payout ... High-End NON NCB')", rate(he("1+1", "SBI_MAY26_BMW_5_Series", "YES", "Diesel")) == 20.0 and rate(he("1+1", "SBI_MAY26_BMW_5_Series", "NO", "Diesel")) == 0.0 and rate(he("SAOD", "SBI_MAY26_BMW_5_Series", "YES", "Petrol")) == 20.0)
check("High-End fuels follow the sheet's words: 'Petrol / Hybrid' = Petrol + Electric, 'Diesel / CNG / LPG' = Diesel + CNG + LPG; 8 Series (Petrol / Hybrid only) has no Diesel / CNG",
      set(sel(df, new_rto_list=mh, new_vehicle_makes="SBI_MAY26_BMW_5_Series", sub_product="1+1").fuel_type) == {"Petrol", "Electric", "Diesel", "CNG", "LPG"} and set(sel(df, new_rto_list=mh, new_vehicle_makes="SBI_MAY26_BMW_8_Series", sub_product="1+1").fuel_type) == {"Petrol", "Electric"})
check("High-End: only the clusters of the enabler sheet get rows (AP - Rest has none), the group name is  SBI_MAY26_<make>_<model>", len(sel(df, new_rto_list=G("AP - Rest"), new_vehicle_makes="SBI_MAY26_BMW_5_Series")) == 0 and "SBI_MAY26_JLR_XE_/_XF_Series" in set(df.new_vehicle_makes))
check("High-End 'MB Class' has no make group in the master: the rows are written AND flagged (never mapped to a lookalike)", df[df.new_vehicle_makes == "SBI_MAY26_Mercedes_MB_Class"].remarks.str.contains("not in the make / model master").all()
      and len(df[df.new_vehicle_makes == "SBI_MAY26_Mercedes_MB_Class"]) > 0)

# ================================================================ the AP 'non doable RTOs' group
ap = sel(df, product="GCV 4W", new_vehicle_makes="Mahindra", cc_min=2500.01, sub_product="1+1", make_model_class="Truck", fuel_type="Diesel", vehicle_age_min=0)
check("2.5-3.5 t of AP - Rest goes to the group that is AP - Rest without AP02 / AP03 / AP04 / AP21 / AP26 / AP27 (found by its RTO codes in the master)", set(ap[ap.new_rto_list.str.contains("AP_-_Rest")].new_rto_list) == {"SBI_OCT26_AP_-_Rest_2.5_3.5T"}, set(ap.new_rto_list))
other = sel(df, product="GCV 4W", new_vehicle_makes="Mahindra", cc_min=12000.01, sub_product="1+1", make_model_class="Truck", fuel_type="Diesel", vehicle_age_min=0)
check("... and every other segment of AP - Rest keeps the plain cluster group", set(other[other.new_rto_list.str.contains("AP_-_Rest")].new_rto_list) == {"SBI_OCT26_AP_-_Rest"})
def to_cluster(frame):
    mod.RTO_NAMES = "cluster"
    try:
        return mod.apply_cluster_names(frame)
    finally:
        mod.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")


dfc, _ = build(mini(), names="cluster")
txt = to_cluster(dfc)
spc = txt[txt.new_rto_list.str.contains(" - Condition: ", regex=False)]
check("cluster mode: the AP rows carry the cluster name, the condition in plain English and the exact part; the other rows the bare cluster",
      len(spc) > 0 and spc.new_rto_list.str.startswith("AP - Rest - Condition: AP02/AP03/AP04/AP21/AP26/AP27 are not allowed RTOs for GCV 2.5 t-3.5 t - Part: AP - Rest without AP02/AP03/AP04/AP21/AP26/AP27 - Product: GCV 4W").all()
      and set(txt[~txt.new_rto_list.str.contains(" - Condition: ", regex=False)].new_rto_list) == set(CLUSTERS), sorted(set(txt.new_rto_list))[:3])

# ================================================================ header-driven: the same file shifted, renamed, with a new column
base = mod.to_csv_frame(df)
shifted, _ = build(mini(row_off=3, col_off=2, rate_name="Oct 26 Rates (final)"))
check("the same grid moved 3 rows down and 2 columns right on a renamed sheet gives the same upload (no cell address anywhere)", mod.to_csv_frame(shifted).equals(base))


wn = mini()
sh = wn.worksheets[0]
sh.cell(5, 117, "Extra scheme")                                      # a numeric column the script has no rule for, to the right of the table
for r in range(6, 6 + len(CLUSTERS)):
    sh.cell(r, 117, 9)
dfx, wx = build(wn)
check("a new numeric column that no rule reads is warned about and listed (nothing written for it, nothing else changes)", any("Extra scheme" in x for x in wx) and mod.to_csv_frame(dfx).equals(base) and any("no rule for this column" in str(r) for r in mod.LAST_NOT_WRITTEN), [x[:90] for x in wx][:3])

wh = mini()
wh.worksheets[0].cell(5, 117, "Brand new product")                  # labelled, but empty: still reported
dfh, wh2 = build(wh)
check("a labelled column without any number that no rule reads is reported too", any("Brand new product" in x and "is labelled but" in x for x in wh2) and mod.to_csv_frame(dfh).equals(base), [x[:90] for x in wh2][:3])

# ================================================================ the clusters and the sheets
dup = mini(clusters=CLUSTERS + ("GA",))
try:
    build(dup)
    q = ""
except mod.NeedsAnswer as e:
    q = str(e)
check("a cluster that appears twice in the rate sheet is asked about, never merged silently", "appears twice" in q, q)
dfd, wd = build(dup, answers=iter(["2"]))
check("... and the answer picks one row: the same clusters as without the repeat", set(dfd.new_rto_list) == set(df.new_rto_list) and len(dfd) == len(df))
wm = mini()
mp = wm["RTO Cluster Mapping"]
mp.append(["XX", "Nowhere", "XX01", "XX - New", "RC 99"])
dfm2, wm2 = build(wm)
check("a cluster of the RTO list that the rate sheet does not have is warned about", any("XX - New" in x for x in wm2), [x[:80] for x in wm2])
wf = mini()
wf["Car, Taxi Declined list"].append([None, "Skoda", "All models"])
dff, wf2 = build(wf)
check("a changed declined make / model list is warned about (the make groups need a look)", any("declined make / model lists" in x and "changed" in x for x in wf2), [x[:80] for x in wf2])
check("... and the unchanged lists give no such warning", not any("declined make / model lists" in x for x in warns), warns[:3])
dfp, wp = build(mini(), f=dt.date(2026, 11, 1))
check("a from_date that is not the start of the grid's own period is warned about", any("03/10/2026" in x and "01/11/2026" in x for x in wp), [x[:100] for x in wp])
check("the matching from_date gives no such warning", not any("applies from" in x for x in warns))

# ================================================================ blank / text cells, cross-checks and flagged notes
ws_t = mini(tweak=lambda w: setattr(w.worksheets[0].cell(6, COL("BY")), "value", "D"))
dt_, _ = build(ws_t)
t_ = sel(dt_, new_rto_list=mh, product="Private Car", sub_product="1+1", new_vehicle_makes=mod.MK_PC, is_ncb="YES", fuel_type="Petrol")
check("a text in a rate cell ('D') is written as 0 and the rows are flagged with the cell address", rate(t_) == 0.0 and "BY6 holds the text 'D'" in t_.iloc[0].remarks, t_.iloc[0].remarks)


def pay_a_declined(w):
    sh = w.worksheets[0]
    row = 6 + CLUSTERS.index("CH - Rest")
    for col in ("AS", "AT"):
        sh.cell(row, COL(col)).value = 12                            # '12T to 20T' is declined for CH - Rest by the note, the grid now pays
dp, _ = build(mini(tweak=pay_a_declined))
check("the note declines a cluster but the grid pays there: flagged on exactly that cluster and segment", dp[dp.remarks.str.contains("declines 'CH - Rest' but the grid pays", regex=False)].new_rto_list.unique().tolist() == [G("CH - Rest")]
      and set(dp[dp.remarks.str.contains("declines 'CH - Rest' but the grid pays", regex=False)]["product"]) == {"GCV 4W"}, dp.remarks[dp.remarks != ""].unique()[:3])
check("... and an unchanged grid has no such flag for a cluster the notes decline", not df.remarks.str.contains("declines 'CH - Rest'", regex=False).any())
check("the DL +2% note (2.5-3.5 t) is APPLIED since 3 Nov 2026 (see the section on it): no row says 'Manual Review Required: additional 2%- DL Registration ...' any more", not df.remarks.str.contains("DL Registration", regex=False).any())
tnr = df[df.remarks.str.contains("TN‑C|TN-C", regex=True)]
check("the TN-C RTO note is flagged on the Tamil Nadu clusters' 12-40 t rows only", len(tnr) > 0 and set(tnr.new_rto_list) == {G("TN - C")} and set(tnr.cc_min) == {12000.01, 20000.01}, (set(tnr.new_rto_list), set(tnr.cc_min)))
check("a standing note (make declines, 'All the Grids are applicable ...') is never flagged", not df.remarks.str.contains("imported makes|All the Grids", regex=True).any())


def new_note(w):
    sh = w.worksheets[0]
    sh.cell(sh.max_row + 1, 4, "A brand new SBI rule: policies of vehicles older than 7 years get 4% less")


dnn, _ = build(mini(tweak=new_note))
flag = dnn.remarks.str.contains("brand new SBI rule", regex=False)
check("a new text nobody has seen is flagged on EVERY row (it names no segment and no state)", flag.all() and len(dnn) == len(df), flag.mean())


def new_taxi_note(w):
    sh = w.worksheets[0]
    sh.cell(sh.max_row + 1, 3, "PCV Taxi")
    sh.cell(sh.max_row, 4, "A new taxi condition: WB clusters get 2% less on comprehensive")


dtn, _ = build(mini(clusters=CLUSTERS + ("WB - K",), tweak=new_taxi_note))
fl = dtn[dtn.remarks.str.contains("new taxi condition", regex=False)]
check("a new text that names the taxi and WB is flagged on the WB clusters' taxi rows only", len(fl) > 0 and set(fl["product"]) == {"PCV 4W"} and set(fl.new_rto_list) == {G("WB - K")}, (set(fl["product"]), set(fl.new_rto_list)))

dhe, _ = build(mini(tweak=lambda w: w["PV SAOD- HIGHEND Enabler"].append(["Bonanza valid only for policies issued till 20th Oct"])))
fl = dhe[dhe.remarks.str.contains("Bonanza valid only", regex=False)]
check("a new line on the High-End sheet is flagged on the High-End rows only; its title and 'Pvt Car Slabs to Apply' (no action, the user) are not",
      len(fl) > 0 and fl.new_vehicle_makes.str.startswith("SBI_MAY26_").all() and not (fl.new_vehicle_makes == mod.MK_PC).any() and set(fl["product"]) == {"Private Car"}
      and not df.remarks.str.contains("Slabs to Apply|Bonnazza", regex=True).any(), len(fl))

# ================================================================ the cluster mode names and the output frame
out = txt
check("cluster mode (the default): new_rto_list is the cluster as the grid writes it ('MH - M'), not the portal's group name", "MH - M" in set(out.new_rto_list) and not out.new_rto_list.str.startswith("SBI_OCT26_").any())
check("every row has the constants of the portal's export (id 1, company, Not Applicable, user_id 2966, tariff 0-100.01, dates DD/MM/YYYY)", set(df.id) == {1} and set(df.insurance_company) == {"SBI GENERAL INSURANCE COMPANY LIMITED"} and set(df.insurer_vertical) == {"Not Applicable"}
      and set(df.user_id) == {2966} and set(df.tariff_max) == {100.01} and set(df.from_date) == {"03/10/2026"} and set(df.to_date) == {"31/03/2030"} and set(df.policy_type) == {"Renew"})

# ================================================================ the real October grid against the user's own upload
REF = os.path.join(helpers.REFS, "sbi_oct26_rates_export.csv")
if os.path.exists(REF):
    real = openpyxl.load_workbook(SRC, data_only=True)
    dr, wr = build(real)
    ref = pd.read_csv(REF, dtype=str, keep_default_na=False)
    mine = mod.to_csv_frame(dr)
    mine["_k"], ref["_k"] = mod.key_of(mine), mod.key_of(ref)
    mg = mine.merge(ref.drop_duplicates("_k"), on="_k", how="outer", suffixes=("", "_ref"), indicator=True)
    both = mg[mg["_merge"] == "both"]
    pi = ["pi_od_rate", "pi_tp_rate", "pi_net_rate", "po_od_rate", "po_tp_rate", "po_net_rate"]
    d_in = sum(pd.to_numeric(both[c]) - pd.to_numeric(both[c + "_ref"]) for c in pi[:3]).round(4)
    d_out = sum(pd.to_numeric(both[c]) - pd.to_numeric(both[c + "_ref"]) for c in pi[3:]).round(4)
    moved = both[(d_in != 0) | (d_out != 0)]
    # the October upload took the Private Car slab off every two-wheeler payout; the user (3 Nov 2026): there is NO slab for two-wheelers, so those rows are the grid cell as printed (+1).  The October upload also
    # left out the 'additional 2% - DL Registration Vehicle 2.5 to 3.5' points that the user added by hand in November: the GCV 4W 2.5-3.5 t rows of DL - NCR are +2.
    tw_moved, gcv_moved = moved[moved["product"] == "TW"], moved[moved["product"] == "GCV 4W"]
    check("THE REAL GRID: 106,644 rows, 94,176 keys equal the user's October upload; the ONLY payin / payout differences are the 248 two-wheeler rows now written as printed (+1 each: no slab for two-wheelers) and the 400 GCV 4W "
          "2.5-3.5 t rows of DL - NCR (shared with the upload) that now carry the note's 2 points (+2 each)",
          len(dr) == 106644 and len(both) == 94176 and len(moved) == 648 and len(tw_moved) == 248 and len(gcv_moved) == 400 and set(d_in[tw_moved.index]) == {1.0} and set(d_out[tw_moved.index]) == {1.0}
          and set(d_in[gcv_moved.index]) == {2.0} and set(d_out[gcv_moved.index]) == {2.0} and set(gcv_moved["new_rto_list"]) == {G("DL - NCR")}
          and set(gcv_moved["cc_min"].astype(float)) == {2500.01, 3000.01} and set(gcv_moved["cc_max"].astype(float)) == {3000.01, 3500.01},
          (len(dr), len(both), len(moved), len(tw_moved), len(gcv_moved)))
    band = mine[(mine.new_rto_list == G("DL - NCR")) & (mine["product"] == "GCV 4W") & (mine.cc_min.astype(float) >= 2500.01) & (mine.cc_max.astype(float) <= 3500.01)]
    off = band[~band["_k"].isin(set(both["_k"]))]
    check("... the DL - NCR 2.5-3.5 t band has 480 rows: 400 moved, and the other 80 are the SATP-above-5-years rows whose age cap (25.01, not the upload's 20.01) gives them no key in the upload - none is a zero card",
          len(band) == 480 and len(off) == 80 and set(off.sub_product) == {"STP"} and set(off.vehicle_age_max.astype(float)) == {25.01} and (off[pi[:3]].astype(float).sum(axis=1) > 0).all(), (len(band), len(off)))
    tw_all = both[both["product"] == "TW"]
    tw_same = tw_all[~tw_all.index.isin(moved.index)]
    check("... the other 400 of the upload's 648 two-wheeler rows are zero cards (a blank / zero cell: nothing to add back) and unchanged",
          len(tw_all) == 648 and len(tw_same) == 400 and (sum(pd.to_numeric(tw_same[c]) for c in pi[:3]) == 0).all(), (len(tw_all), len(tw_same)))
    gone = mg[mg["_merge"] == "right_only"]
    by_kind = gone.groupby(["product_ref", "sub_product_ref"]).size().to_dict()
    check("... the rows only the upload has are exactly: GCV 4W SATP with the 20.01 age cap, the High-End rows with NCB 'NA', and the 0-10 year car SATP band",
          by_kind == {("GCV 4W", "STP"): 4320, ("Private Car", "1+1"): 1649, ("Private Car", "SAOD"): 1649, ("Private Car", "STP"): 810}, by_kind)
    spot = [("MH - M", "E", dict(product="GCV 4W", sub_product="1+1", make_model_class="Truck", new_vehicle_makes=mod.MK_BELOW35, fuel_type="Diesel", vehicle_age_min="0.0", vehicle_age_max="1.0")),
            ("OD", "BX", dict(product="GCV 4W", sub_product="STP", make_model_class="Truck", new_vehicle_makes="TATA,Ashok", fuel_type="Diesel", vehicle_age_min="5.01", cc_min="40000.01")),
            ("KA - B", "CP", dict(product="PCV 4W", make_model_class="Taxi", sub_product="1+1", is_ncb="YES", is_zd="NO", fuel_type="Petrol", cc_min="1499.01")),
            ("GJ - A", "CU", dict(product="TW", sub_product="STP", make_model_class="SCOOTER", fuel_type="Petrol")),
            ("OD", "DK", dict(product="PCV 4W", make_model_class="School Bus", sub_product="STP", fuel_type="Petrol", sc_min="12.01"))]
    mtxt = mod.to_csv_frame(dr)
    ok = True
    for cl, col, kw in spot:
        rows = mtxt[(mtxt.new_rto_list == G(cl))]
        for k, v in kw.items():
            rows = rows[rows[k].astype(str).isin([str(v), str(v).replace(".0", "")])]
        got = float(rows[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).max(axis=1).iloc[0]) if len(rows) else None
        want = (src(cl, col) or 0) if kw["product"] == "TW" else max((src(cl, col) or 0) - 1, 0)           # a two-wheeler has no slab
        if got != want:
            ok = False
            print("   spot check", cl, col, got, want)
    check("five cells read straight from the workbook (Mumbai E, Odisha BX, Bangalore CP, Gujarat CU, Odisha DK) are in the output minus the slab - the two-wheeler cell (Gujarat CU) as printed", ok)

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
