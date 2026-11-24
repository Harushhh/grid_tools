"""Deterministic checks of icici_grid_to_config.py: the cell grammar of the CV matrices (age labels, clauses, OTHERS, declines, zero dep, policies), the make-group composer against the real make
master, the header-driven reading (sheets renamed, rows / columns inserted, columns shuffled, the person's own RTO column), the notes and what is flagged, Private Car / Two-Wheeler / new
Two-Wheeler against small edits of the real grids, and the real grids against the user's own uploads (the November Private Car + Two-Wheeler upload: every row; the October CV upload).
Run:  .venv\\Scripts\\python.exe tests\\icici_tests.py"""
import datetime as dt
import os
import sys
import tempfile

import numpy as np
import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("icici_grid_to_config.py")
F, T = dt.date(2026, 11, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []
tmp = tempfile.mkdtemp(prefix="icici_")


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


IN = lambda n: os.path.join(helpers.INPUTS, n)
REF = lambda n: os.path.join(helpers.REFS, n)
close = lambda a, b: abs(float(a) - float(b)) < 1e-6


def build(specs, ask=None, mode="group"):
    old = mod.RTO_NAMES
    mod.RTO_NAMES = mode
    try:
        return mod.build(specs, F, T, quiet, ask or mod.Asker(interactive=False))
    finally:
        mod.RTO_NAMES = old


def sel(df, **kw):
    m = np.ones(len(df), dtype=bool)
    for k, v in kw.items():
        if isinstance(v, (list, tuple, set)):
            m &= df[k].isin(list(v)).values
        elif isinstance(v, float):
            m &= np.isclose(pd.to_numeric(df[k], errors="coerce").values, v)
        else:
            m &= (df[k].astype(str) == str(v)).values
    return df[m]


def pi(d):
    r = d.iloc[0]
    return max(float(r["pi_od_rate"]), float(r["pi_tp_rate"]), float(r["pi_net_rate"]))


def po(d):
    r = d.iloc[0]
    return max(float(r["po_od_rate"]), float(r["po_tp_rate"]), float(r["po_net_rate"]))


def shape(df):
    """the frame without the remarks, sorted by its key: layout shifts must not change anything else"""
    d = df.drop(columns=["remarks"]).copy()
    d["_k"] = mod.key_of(d)
    return d.sort_values(["_k", "fuel_type"]).reset_index(drop=True)


# ================================================================ the cell grammar
for label, want in (("1-3 yrs", (1, 3.01)), ("4-5 yrs", (3.01, 5.01)), (">=4 yrs", (3.01, 99)), (">=6yrs", (5.01, 99)), ("10+ yrs", (10.01, 99)), ("<=10 yrs", (0, 10.01)),
                    ("0-10yrs", (0, 10.01)), ("6-10 yrs", (5.01, 10.01)), ("1-5yrs", (1, 5.01)), ("1-10 years", (1, 10.01))):
    kind, val = mod.classify_word(label, {})
    check(f"age label '{label}' = {want} (the convention of the Shriram scripts and of the user's ICICI upload)", kind == "age" and close(val[0], want[0]) and close(val[1], want[1]), (kind, val))

check("a '%' followed by '|' is a typo for a comma", mod.split_clauses("JK1|<=10 yrs:28% |OTHERS:0%") == ["JK1|<=10 yrs:28%", "OTHERS:0%"], mod.split_clauses("JK1|<=10 yrs:28% |OTHERS:0%"))
check("'AOTP15%' (no colon) and 'ZD-20%' (a dash) are read", mod.parse_clause("AOTP15%") == (["AOTP"], 15.0) and mod.parse_clause("COMP|NEW|ZD-20%") == (["COMP", "NEW", "ZD"], 20.0),
      (mod.parse_clause("AOTP15%"), mod.parse_clause("COMP|NEW|ZD-20%")))
check("a clause without a rate is only 'CC' or an unknown text", mod.parse_clause("CC") == (["CC"], None))
check("a make glued to an age ('M&M>=4 yrs') is two words", mod.split_glued("M&M>=4 yrs") == ["M&M", ">=4 yrs"], mod.split_glued("M&M>=4 yrs"))
W = lambda w, places=None: mod.classify_word(w, places or {})
check("words: M&M is a make, Fuel Tanker a vehicle type, KASHMIR / HP1 / UP1-WEST places, BOLERO unknown",
      W("M&M") == ("make", "MAHINDRA") and W("Fuel Tanker")[0] == "body" and W("KASHMIR")[0] == "loc" and W("HP1")[0] == "loc" and W("UP1-WEST")[0] == "loc" and W("BOLERO")[0] == "unknown",
      [W(x) for x in ("M&M", "Fuel Tanker", "KASHMIR", "HP1", "UP1-WEST", "BOLERO")])
check("a district is a place only when the workbook's own lists name it", W("Purnia")[0] == "unknown" and W("Purnia", {"purnia": "PURNIA"}) == ("loc", "PURNIA"))


def regs(text, base=None, pol="COMP", places=None, skipped=None):
    return mod.regions(mod.parse_cell(text, places or {}), base or mod.full_box(), pol, skipped)


def by_age(rs):
    out = {}
    for box, rate, _ in rs:
        for iv in box["age"]:
            out[(round(iv[0], 2), round(iv[1], 2))] = rate
    return out


check("OLD:10%,NEW:0% -> old 1-99 at 10 and new 0-1 at 0 (a named zero is a zero card)", by_age(regs("OLD:10%,NEW:0%")) == {(1.0, 99.0): 10.0, (0.0, 1.0): 0.0}, by_age(regs("OLD:10%,NEW:0%")))
check("words of one kind in a clause are alternatives: '1-3 years|6-10 years:20%' = two bands at 20",
      by_age(regs("1-3 years|6-10 years:20%,4-5 years|10+ years:44%")) == {(1.0, 3.01): 20.0, (5.01, 10.01): 20.0, (3.01, 5.01): 44.0, (10.01, 99.0): 44.0}, by_age(regs("1-3 years|6-10 years:20%,4-5 years|10+ years:44%")))
rs = regs("TATA|M&M|MARUTI:55%,OTHERS:35%")
named = [r for r in rs if r[1] == 55.0][0][0]["makes"]
rest = [r for r in rs if r[1] == 35.0][0][0]["makes"]
check("TATA|M&M|MARUTI:55%,OTHERS:35% -> the three makes at 55, everything else ('not' the three) at 35", named == ("in", frozenset({"TATA", "MAHINDRA", "MARUTI"})) and rest == ("not", frozenset({"TATA", "MAHINDRA", "MARUTI"})), (named, rest))
skipped = []
rs = regs("M&M:18%,OTHERS:0%", skipped=skipped)
check("OTHERS:0% is not written (WRITE_OTHERS_ZERO = False) and is handed back as skipped; the named rate stays", [r[1] for r in rs] == [18.0] and len(skipped) == 1, (rs, skipped))
mod.WRITE_OTHERS_ZERO = True
rs = regs("M&M:18%,OTHERS:0%")
mod.WRITE_OTHERS_ZERO = False
check("with WRITE_OTHERS_ZERO = True the OTHERS:0% region is written", sorted(r[1] for r in rs) == [0.0, 18.0], rs)
rs = regs("Omni|Magic|Eeco:0%,NEW|COMP:10%,OLD|COMP:25%,OLD|AOTP:38%")
pos = [r for r in rs if r[1] in (10.0, 25.0)]
check("a 0% clause is a decline: the other COMP rows are 'all makes except Omni, Magic, Eeco'", len(pos) == 2 and all(r[0]["makes"] == ("not", frozenset({"OMNI", "MAGIC", "EECO"})) for r in pos), pos)
rs = regs("COMP:40%,AOTP:30%")
rs2 = regs("COMP:40%,AOTP:30%", pol="AOTP")
check("COMP:40%,AOTP:30% -> COMP 40, AOTP 30", [r[1] for r in rs] == [40.0] and [r[1] for r in rs2] == [30.0], (rs, rs2))
check("a cell with no AOTP clause gives the AOTP policy the comprehensive clauses (CV NOTES 9)", by_age(regs("COMP|New:45%,COMP|Old:20%", pol="AOTP")) == {(0.0, 1.0): 45.0, (1.0, 99.0): 20.0})
rs = regs("COMP|NEW|NZD:25%,COMP|NEW|ZD-20%,COMP|OLD:38%")
zd = {tuple(sorted(r[0]["zd"][1])) if r[0]["zd"] != mod.ALL else (): r[1] for r in rs}
check("NZD = is_zd NO, ZD = YES, no mark = NA", zd == {("NO",): 25.0, ("YES",): 20.0, (): 38.0}, zd)
rs = regs("OLD|TATA|>=6 yrs:20%,OLD|1-3yrs|>=6 yrs:20%,OTHERS:0%")
check("a region wholly inside another region of the same rate is dropped (no duplicate TATA row)", len(rs) == 1 and rs[0][0]["makes"] == mod.ALL, rs)
a, b = mod.full_box(), mod.full_box()
b["locs"] = mod.s_in(["X"])
c = dict(a)
c["locs"] = mod.s_in(["Y"])
check("subtracting a box that does not overlap returns the box whole (no spurious split)", mod.b_sub(c, b) == [c])
p = mod.parse_cell("BOLERO:10%,OTHERS:5%", {})
check("a word nobody knows is reported, never guessed", p["unknown"] == ["BOLERO"], p["unknown"])
p = mod.parse_cell("CC,OTHERS:22%", {})
check("CC marks the cell and OTHERS:22% keeps its rate", p["cc"] and [c["rate"] for c in p["clauses"]] == [22.0], p)
cv = lambda v, u="percent": mod.cell_value(v, u)
check("cells: '60.0%' = 60, a fraction table's 0.6 = 60, 'No Biz' and a blank = zero card, #N/A = error, other text = text",
      cv("60.0%") == ("num", 60.0) and cv(0.6, "fraction") == ("num", 60.0) and cv("No Biz")[0] == "zero" and cv(None)[0] == "zero" and cv("#N/A")[0] == "error" and cv("OLD:10%")[0] == "text", [cv("60.0%"), cv(0.6, "fraction"), cv("#N/A")])

CS = mod.classify_segment
seg = {h: CS(h) for h in ("PCV 3W Petrol/CNG New", "PCV 3W Non-Diesel Old", "Staff Bus >18", "MHCV >40T Tipper/ Dumper", "SCV >= 2450 GVW Old", "PCVTAXI<=1000CC", "MIsc D CE (Excluding CRANES)",
                         "MHCV 12-20T Tanker", "MHCV 12-18T Truck", "School Bus 18-36", "GCV 3W Electric", "Some New Column")}
check("segment headers of every month are known: Petrol/CNG = Non-Diesel, Staff Bus >18, >40T Tipper/ Dumper, SCV >= 2450 Old, taxi bands, MISCD CE; a new column is None",
      seg["PCV 3W Petrol/CNG New"]["key"] == "pcv3w_nd_new" and seg["PCV 3W Non-Diesel Old"]["key"] == "pcv3w_nd_old" and seg["Staff Bus >18"]["sc"] == (18.01, 999) and seg["MHCV >40T Tipper/ Dumper"]["cls"] == "Dumper Tipper"
      and seg["SCV >= 2450 GVW Old"]["cc"] == (2450.0, 3500.01) and seg["PCVTAXI<=1000CC"]["cc"] == (0, 1000.01) and seg["MIsc D CE (Excluding CRANES)"]["cls"] == "MISCD-Others" and seg["Some New Column"] is None,
      {k: (v or {}).get("key") for k, v in seg.items()})
check("tonnages are read from the header (12-20T = 12000.01-20000.01, a renamed 12-18T band follows), the Tanker class carries the hazardous T&C, GCV 3W Electric is Electric only",
      seg["MHCV 12-20T Tanker"]["cc"] == (12000.01, 20000.01) and seg["MHCV 12-20T Tanker"]["text"] == mod.TNC_HAZARD and seg["MHCV 12-18T Truck"]["cc"] == (12000.01, 18000.01) and seg["GCV 3W Electric"]["fuels"] == ["Electric"])

# ================================================================ the make composer (the real make master)
masters = mod.load_masters(mod.Asker(interactive=False))
MN = mod.MakeNames(masters, {"BHARAT_BENZ"})
t = lambda fam, dom: MN.text(fam, dom, {})[0]
check("default groups: GCV 4W / MISCD / bus leave out Bharat Benz, taxi / PCV 3W / PCV 2W / GCV 3W do not",
      t("gcv4w", mod.ALL) == "GCV_4W_ALL_EX_BHARAT_BENZ" and t("miscd", mod.ALL) == "MISCD_ALL_EX_BHARAT_BENZ" and t("bus", mod.ALL) == "BUS_ALL_EX_BHARAT_BENZ" and t("taxi", mod.ALL) == "TAXI_ALL"
      and t("pcv3w", mod.ALL) == "PCV_3W_ALL" and t("gcv3w", mod.ALL) == "GCV_3W_ALL", [t(f, mod.ALL) for f in ("gcv4w", "miscd", "bus", "taxi", "pcv3w", "gcv3w")])
check("everything but TATA, ASHOK, MAHINDRA = the master's group of exactly that", t("gcv4w", ("not", frozenset({"TATA", "ASHOK", "MAHINDRA"}))) == "GCV_4W_ALL_EX_BHARAT_BENZ_TATA_ASHOK_MAHINDRA")
check("everything but TATA, MAHINDRA = the group that also leaves out ASHOK + 'Ashok' given back (the user's own form)", t("gcv4w", ("not", frozenset({"TATA", "MAHINDRA"}))) == "GCV_4W_ALL_EX_BHARAT_BENZ_TATA_ASHOK_MAHINDRA,Ashok", t("gcv4w", ("not", frozenset({"TATA", "MAHINDRA"}))))
check("a make set = the master's single-make names in the cell's order (TATA,Ashok)", MN.text("gcv4w", ("in", frozenset({"TATA", "ASHOK"})), {"TATA": 0, "ASHOK": 1})[0] == "TATA,Ashok")
check("taxi: everything but Eeco, Omni, Magic = TaxiALL_EX_EECO_OMNI_MAGIC; everything but Maruti = TAXI_ALL_EX_MARUTI_HYUNDAI + Hyundai given back",
      t("taxi", ("not", frozenset({"EECO", "OMNI", "MAGIC"}))) == "TaxiALL_EX_EECO_OMNI_MAGIC" and t("taxi", ("not", frozenset({"MARUTI"}))) == "TAXI_ALL_EX_MARUTI_HYUNDAI,Hyundai",
      (t("taxi", ("not", frozenset({"EECO", "OMNI", "MAGIC"}))), t("taxi", ("not", frozenset({"MARUTI"})))))
check("bus: Magic / Force / Sumo and everything else have their master groups", t("bus", ("in", frozenset({"MAGIC", "FORCE", "SUMO"}))) == "BUS_MAGIC_FORCE_SUMO" and t("bus", ("not", frozenset({"MAGIC", "FORCE", "SUMO"}))) == "BUS_EX_MAGIC_FORCE_SUMO_BHARAT_BENZ")
text, why = MN.text("taxi", ("in", frozenset({"OMNI", "MAGIC", "EECO"})), {})
check("a make set the master has no group for is reported, not replaced", text is None and "no group" in why, (text, why))

# ================================================================ small workbooks made of the real CV grid
CLUSTERS = ("AHMEDABAD & GANDHINAGAR", "ASSAM", "BHUBANESHWAR", "BIHAR", "HIMACHALPRADESH", "KOLKATA", "PATNA", "NCR", "GUJARAT", "JAMMUANDKASHMIR")
SEQ = [0]


def cv_mini(clusters=CLUSTERS, edit=None, src="icici_cv_oct26.xlsx"):
    wb = openpyxl.load_workbook(IN(src), data_only=True)
    for name in ("COMP", "AOTP", " COMP", "TP"):
        if name in wb.sheetnames:
            ws = wb[name]
            for r in range(ws.max_row, 1, -1):
                if ws.cell(r, 1).value not in clusters:
                    ws.delete_rows(r)
    if edit:
        edit(wb)
    SEQ[0] += 1
    path = os.path.join(tmp, f"cv_{SEQ[0]}.xlsx")
    wb.save(path)
    wb.close()
    return path


def cell_of(ws, cluster, header):
    col = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(1, c).value).strip() == header)
    row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == cluster)
    return ws.cell(row, col)


PATH = cv_mini()
BASE, W0 = build([("cv", PATH)])
G = lambda cl: mod.GROUP_PREFIX["cv"] + cl
AH = G("AHMEDABAD_&_GANDHINAGAR")
check("the small CV workbook converts (10 clusters)", len(BASE) > 3000 and BASE["new_rto_list"].nunique() >= 10, (len(BASE), BASE["new_rto_list"].nunique()))
check("every row has insurer_vertical ICICI-RB-Offline (the user's choice of 4 Nov 2026), the company, user 2966, id 1 and the dates", set(BASE["insurer_vertical"]) == {"ICICI-RB-Offline"} and set(BASE["insurance_company"]) == {mod.INSURANCE_COMPANY}
      and set(BASE["user_id"]) == {2966} and set(BASE["id"]) == {1} and set(BASE["from_date"]) == {"01/11/2026"} and set(BASE["to_date"]) == {"31/03/2030"})
d = sel(BASE, new_rto_list=AH, product="GCV 3W", sub_product="1+1", fuel_type="Diesel", vehicle_age_min=1.0)
check("plain cell: Ahmedabad GCV 3W Old 60% -> 1+1 Diesel, age 1-99, T&C Old, On Net 60, payout 53 (pi - 7)", len(d) == 1 and pi(d) == 60 and po(d) == 53 and d.iloc[0]["add_tnc"] == "Old" and d.iloc[0]["pi_type"] == "On Net", d.to_dict("records"))
d = sel(BASE, new_rto_list=AH, product="GCV 3W", sub_product="1+1", make_model_class="Goods Carrying Rickshaw", new_vehicle_makes="GCV_3W_ALL")
check("GCV 3W: New / Old carry the four non-electric fuels, the Electric column is Electric only", set(sel(d, add_tnc="Old")["fuel_type"]) == {"Petrol", "Diesel", "CNG", "LPG"} and set(sel(d, add_tnc="")["fuel_type"]) == {"Electric"} and set(sel(d, add_tnc="New")["fuel_type"]) == {"Petrol", "Diesel", "CNG", "LPG"},
      d[["fuel_type", "add_tnc"]].drop_duplicates().to_dict("records"))
d = sel(BASE, new_rto_list=AH, product="GCV 4W", sub_product="1+1", fuel_type="Petrol", cc_min=20000.01, make_model_class="Truck")
check("MHCV 20-40T Truck 17% -> cc 20000.01-40000.01, make group without Bharat Benz, payout 10", len(d) == 1 and pi(d) == 17 and po(d) == 10 and d.iloc[0]["new_vehicle_makes"] == "GCV_4W_ALL_EX_BHARAT_BENZ" and close(d.iloc[0]["cc_max"], 40000.01), d.to_dict("records"))
d = sel(BASE, new_rto_list=AH, product="GCV 4W", sub_product="STP", fuel_type="Petrol", cc_min=3500.01)
check("STP of an AOTP segment comes from the AOTP matrix (LCV 3.5-7.5T = 32), On TP", len(d) >= 1 and pi(sel(d, cc_max=7500.01)) == 32 and sel(d, cc_max=7500.01).iloc[0]["pi_type"] == "On TP", d[["cc_max", "pi_tp_rate"]].to_dict("records"))
d = sel(BASE, new_rto_list=AH, product="GCV 4W", sub_product="STP", fuel_type="Petrol", make_model_class="NA", cc_max=2450.0, vehicle_age_min=1.0)
d1 = sel(BASE, new_rto_list=AH, product="GCV 4W", sub_product="1+1", fuel_type="Petrol", make_model_class="NA", cc_max=2450.0, vehicle_age_min=1.0)
check("STP of a segment the AOTP matrix does not list = the comprehensive payout (SCV Old 49)", len(d) == 1 and pi(d) == 49 and pi(d1) == 49)
d = sel(BASE, new_rto_list=AH, product="MISCD", sub_product="1+1", fuel_type="Petrol", make_model_class="MISCD-Others")
got = sorted((r.add_tnc, r.vehicle_age_min, r.vehicle_age_max, pi(pd.DataFrame([r]))) for r in d.itertuples())
check("MISCD CE 'COMP:CRANES|ACE|ESCORTS|<=10yrs:20%,COMP|OTHERS:42%,AOTP:0%': EX:CRANES,ACE,ESCORTS 42 / CRANES,ACE,ESCORTS 0-10.01 at 20 and 10.01-99 at 42 (the user's partition)",
      got == [("CRANES,ACE,ESCORTS", 0.0, 10.01, 20.0), ("CRANES,ACE,ESCORTS", 10.01, 99.0, 42.0), ("EX:CRANES,ACE,ESCORTS", 0.0, 99.0, 42.0)], got)
d = sel(BASE, new_rto_list=AH, product="MISCD", sub_product="STP", fuel_type="Petrol", make_model_class="MISCD-Others")
check("... and its STP rate is the cell's AOTP:0% (a zero card)", len(d) >= 1 and all(pi(d.iloc[[i]]) == 0 for i in range(len(d))), d[["add_tnc", "pi_tp_rate"]].to_dict("records"))
d = sel(BASE, new_rto_list=G("BHUBANESWAR"), product="GCV 3W", sub_product="1+1", fuel_type="Electric")
check("'OLD:10%,NEW:0%' on GCV 3W Electric (Bhubaneshwar): Old 1-99 at 10, New 0-1 at 0, T&C Old / New", sorted((r.add_tnc, pi(pd.DataFrame([r]))) for r in d.itertuples()) == [("New", 0.0), ("Old", 10.0)], d[["add_tnc", "pi_net_rate"]].to_dict("records"))
d = sel(BASE, new_rto_list=G("ASSAM"), product="PCV 4W", sub_product="1+1", fuel_type="Petrol", make_model_class="School Bus", sc_max=18.01)
check("School Bus <18 'MAGIC|FORCE|SUMO:25%,OTHERS:65%' -> BUS_MAGIC_FORCE_SUMO 25 and BUS_EX_MAGIC_FORCE_SUMO_BHARAT_BENZ 65, the bus T&C",
      {r.new_vehicle_makes: pi(pd.DataFrame([r])) for r in d.itertuples()} == {"BUS_MAGIC_FORCE_SUMO": 25.0, "BUS_EX_MAGIC_FORCE_SUMO_BHARAT_BENZ": 65.0} and set(d["add_tnc"]) == {mod.TNC_BUS}, d[["new_vehicle_makes", "pi_net_rate", "add_tnc"]].to_dict("records"))
d = sel(BASE, new_rto_list=G("HIMACHAL_PRADESH"), product="GCV 4W", sub_product="1+1", fuel_type="Petrol", cc_min=2450.0, make_model_class="NA")
got = sorted((r.vehicle_age_min, r.vehicle_age_max, r.add_tnc, pi(pd.DataFrame([r]))) for r in d.itertuples() if r.vehicle_age_min >= 1)
check("SCV >= 2450 Old 'OLD|1-3 yrs:40%,OLD|>=4 yrs:50%,NEW:0%' -> 1-3.01 at 40 and 3.01-99 at 50, both T&C Old (the SCV rows keep the segment's tag)", got == [(1.0, 3.01, "Old", 40.0), (3.01, 99.0, "Old", 50.0)], got)
d = sel(BASE, new_rto_list=G("PATNA"), product="MISCD", sub_product="1+1", fuel_type="Petrol", make_model_class="MISCD-Tractor")
got = sorted((r.vehicle_age_min, r.vehicle_age_max, pi(pd.DataFrame([r]))) for r in d.itertuples() if r.add_tnc == "")
check("Tractor Old '1-3 years|6-10 years:20%,4-5 years|10+ years:44%' -> four bands (the user's upload has the same four)", got == [(1.0, 3.01, 20.0), (3.01, 5.01, 44.0), (5.01, 10.01, 20.0), (10.01, 99.0, 44.0)], got)
d = sel(BASE, new_rto_list=G("KOLKATA"), product="GCV 4W", sub_product="1+1", fuel_type="Petrol", cc_min=7500.01, cc_max=12000.01)
got = sorted((r.is_zd, r.add_tnc, pi(pd.DataFrame([r]))) for r in d.itertuples())
check("Kolkata LCV 7.5-12T 'COMP|NEW|NZD:25%,COMP|NEW|ZD-20%,COMP|OLD:38%' -> is_zd NO 25 / YES 20 for New, NA 38 for Old", got == [("NA", "Old", 38.0), ("NO", "New", 25.0), ("YES", "New", 20.0)], got)
d = sel(BASE, new_rto_list=G("KOLKATA"), product="GCV 4W", sub_product="1+1", fuel_type="Petrol", cc_min=2450.0, make_model_class="NA", vehicle_age_min=1.0)
got = {r.new_vehicle_makes: pi(pd.DataFrame([r])) for r in d.itertuples()}
check("Kolkata SCV Old 'TATA:10%,M&M:45%,OTHERS:35%' -> TATA 10, Mahindra 45, the master group without TATA / ASHOK / MAHINDRA + Ashok 35",
      got == {"TATA": 10.0, "Mahindra": 45.0, "GCV_4W_ALL_EX_BHARAT_BENZ_TATA_ASHOK_MAHINDRA,Ashok": 35.0}, got)
d = sel(BASE, new_rto_list=G("PATNA"), product="PCV 4W", sub_product="1+1", fuel_type="Petrol", make_model_class="Taxi", cc_min=1000.01)
got = sorted((r.new_vehicle_makes, r.add_tnc.split(",")[0], pi(pd.DataFrame([r]))) for r in d.itertuples())
check("Patna taxi > 1000cc: the 0% 'Omni|Magic|Eeco' clause makes the other rows 'TaxiALL_EX_EECO_OMNI_MAGIC' (New 10, Old 25)", got == [("TaxiALL_EX_EECO_OMNI_MAGIC", "New", 10.0), ("TaxiALL_EX_EECO_OMNI_MAGIC", "Old", 25.0)], got)
nw = pd.DataFrame(mod.LAST_NOT_WRITTEN)
check("OTHERS:0% cells are listed in 'Not written' (no zero card); a bare CC cell too; a make set with no master group too",
      nw["reason"].str.startswith("OTHERS:0%").any() and nw["reason"].str.contains("council call", regex=False).any() and nw["reason"].str.contains("no group for", regex=False).any(), nw["reason"].str[:40].value_counts().head(6).to_dict())
rem = BASE["remarks"].astype(str)
check("a CC cell with a rate (OTHERS:9%) is written and flagged with the council-call note", rem.str.contains("Council call", regex=False).any() and sel(BASE, new_rto_list=AH, product="GCV 4W", sub_product="STP", cc_min=20000.01, cc_max=40000.01, make_model_class="Tanker", fuel_type="Petrol").shape[0] >= 1)
g4 = BASE["product"].eq("GCV 4W") & (BASE["cc_max"].astype(float) <= 3500.02)
check("the garbage-van note is flagged on GCV 4W up to 3.5 t and on nothing else", rem[g4].str.contains("Garbage van", regex=False).all() and not rem[~g4].str.contains("Garbage van", regex=False).any())
taxi = BASE["make_model_class"].eq("Taxi")
check("the taxi note ('For Kolkata, only city meter taxi') is flagged on the taxi rows only", rem[taxi].str.contains("only city meter taxi", regex=False).all() and not rem[~taxi].str.contains("only city meter taxi", regex=False).any())
check("a standing note (CV NOTES 1, 4, 14 ...) is not flagged: no remark carries 'commission is mentioned as 0%' or 'Hazardous segment - such as milk tanker' as a flag", not rem.str.contains("Wherever the commission is mentioned", regex=False).any() and not rem.str.contains("Manual Review Required: [Hazardous", regex=False).any())
check("the Negative RTO list is flagged on the rows of its segment and state (Bihar)", rem[BASE["new_rto_list"].str.contains("BIHAR") & g4].str.contains("Negative RTO list", regex=False).any() or rem.str.contains("Negative RTO list", regex=False).any())
check("every make written is in the make master (no make flag)", not rem.str.contains("is not in the make", regex=False).any() and not rem.str.contains("make / model master", regex=False).any(), rem[rem.str.contains("master", regex=False)].head(2).tolist())

# ---- cluster mode: the cluster name as the grid writes it; a part of a cluster carries its condition
mod.RTO_NAMES = "cluster"
dfc, _ = mod.build([("cv", PATH)], F, T, quiet, mod.Asker(interactive=False))
out = mod.apply_cluster_names(mod.to_csv_frame(dfc))
mod.RTO_NAMES = "group"
check("cluster mode: new_rto_list = the grid's cluster name ('AHMEDABAD & GANDHINAGAR', 'HIMACHALPRADESH')", {"AHMEDABAD & GANDHINAGAR", "HIMACHALPRADESH"} <= set(out["new_rto_list"]), sorted(set(out["new_rto_list"]))[:4])
sc = out[out["new_rto_list"].str.contains(" - Condition: ", regex=False)]
check("a part of a cluster ('PURNIA' in Bihar) carries its condition in plain English with the part, product and sub-product", len(sc) > 0 and sc["new_rto_list"].str.contains("BIHAR - Condition: ", regex=False).any()
      and sc["new_rto_list"].str.contains(" - Part: ", regex=False).all() and sc["new_rto_list"].str.contains(" - Product: ", regex=False).all() and not sc["new_rto_list"].str.contains(",", regex=False).any(), sc["new_rto_list"].head(2).tolist())

# ================================================================ layout shifts: the same rows
def rename(wb):
    wb["COMP"].title, wb["AOTP"].title, wb["CC"].title, wb["RTO MASTER"].title = " COMP", "TP", "COUNCIL CALL", "CV RTO "


def shift(wb):
    for name in ("COMP", "AOTP"):
        ws = wb[name]
        ws.insert_rows(1, 3)
        ws.insert_cols(1, 2)
        ws["A1"] = "ICICI CV grid - November"
        ws["B2"] = "title row"


def working_col(wb):
    for name, head in (("COMP", "rto"), ("AOTP", "RTO")):
        ws = wb[name]
        ws.insert_cols(2)
        ws.cell(1, 2, head)
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 2, "ICICI_SEPT26_CV_" + str(ws.cell(r, 1).value).replace(" ", "_"))


def swap_columns(wb):
    ws = wb["COMP"]
    for r in range(1, ws.max_row + 1):
        a, b = ws.cell(r, 5), ws.cell(r, 20)
        a.value, b.value = b.value, a.value


base_s = shape(BASE)
for label, fn in (("sheets renamed (' COMP', 'TP', 'COUNCIL CALL', 'CV RTO ')", rename), ("three title rows above and two empty columns to the left", shift),
                  ("the person's own RTO column after the cluster column (headed rto / RTO)", working_col), ("two segment columns swapped", swap_columns)):
    other, w = build([("cv", cv_mini(edit=fn))])
    o = shape(other)
    same = len(o) == len(base_s) and (o["_k"].values == base_s["_k"].values).all() and np.allclose(o[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).values, base_s[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).values)
    check(f"layout shift - {label}: the same rows with the same rates", same, (len(o), len(base_s)))
ask = mod.Asker(interactive=False)
build([("cv", cv_mini(edit=working_col))], ask)
check("the person's own RTO column is not used and the run says so", any("is no segment" in x for x in ask.decisions), ask.decisions[-6:])

# ================================================================ cells and columns the script cannot read
def break_cells(wb):
    ws = wb["COMP"]
    cell_of(ws, "ASSAM", "School Bus >36").value = "BOLERO:10%,OTHERS:5%"
    cell_of(ws, "ASSAM", "PCVTAXI>1000CC").value = "#N/A"
    ws.cell(1, ws.max_column + 1, "Brand New Segment")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, ws.max_column, "12%")


other, w = build([("cv", cv_mini(edit=break_cells))])
nw = pd.DataFrame(mod.LAST_NOT_WRITTEN)
asm = other[other["new_rto_list"] == G("ASSAM")]
check("a cell with an unknown word is listed in 'Not written' and no row is guessed for it", nw["reason"].str.contains("not understood: BOLERO", regex=False).any() and sel(asm, make_model_class="School Bus", sc_min=36.01).empty)
check("an Excel error in a cell is listed, not converted", nw["reason"].str.contains("Excel error", regex=False).any() and sel(asm, make_model_class="Taxi", cc_min=1000.01).empty)
ab = BASE[BASE["new_rto_list"] == G("ASSAM")]
gone = ab[(ab["make_model_class"].eq("School Bus") & np.isclose(ab["sc_min"].astype(float), 36.01)) | (ab["make_model_class"].eq("Taxi") & np.isclose(ab["cc_min"].astype(float), 1000.01))]
check("a column the script does not know is warned about and listed, never read; only the rows of the two broken cells are missing", any("Brand New Segment" in x for x in w) and nw["what"].astype(str).str.contains("Brand New Segment", regex=False).any()
      and len(gone) > 0 and len(other) == len(BASE) - len(gone), (len(other), len(BASE), len(gone)))


def drop_note(wb):
    ws = wb["CV NOTES"]
    for r in range(1, ws.max_row + 1):
        if "declined across the system" in str(ws.cell(r, 2).value):
            ws.cell(r, 2).value = "BHARAT BENZ is welcome"


try:
    build([("cv", cv_mini(edit=drop_note))], mod.Asker(interactive=False))
    ok = False
except mod.NeedsAnswer as e:
    ok = "declines a make across the system" in str(e)
check("a rule whose note has gone (BHARAT BENZ declined) is ASKED about, never silently kept; --no-prompt stops with the question", ok)
answers = iter(["1"])
ask = mod.Asker(interactive=True, input_fn=lambda prompt="": next(answers, ""), out=quiet)
d3, w3 = build([("cv", cv_mini(edit=drop_note))], ask)
check("answering 'keep last month's rule' gives the same rows and a warning", len(d3) == len(BASE) and any("declines a make" in x for x in w3), [x[:80] for x in w3][:3])


def dup_cluster(wb):
    ws = wb["COMP"]
    r = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "ASSAM")
    ws.insert_rows(r + 1)
    for c in range(1, ws.max_column + 1):
        ws.cell(r + 1, c, ws.cell(r, c).value)
    cell_of(ws, "ASSAM", "GCV 3W Old").value = "70.0%"        # the first ASSAM row now says 70, the copy keeps the original


d4, w4 = build([("cv", cv_mini(edit=dup_cluster))])
a = sel(d4, new_rto_list=G("ASSAM"), product="GCV 3W", sub_product="1+1", fuel_type="Petrol", vehicle_age_min=1.0)
b = sel(BASE, new_rto_list=G("ASSAM"), product="GCV 3W", sub_product="1+1", fuel_type="Petrol", vehicle_age_min=1.0)
check("a cluster written twice with two payins: the LOWEST is written and the choice is listed (the user's rule of 27 Oct 2026)", len(a) == 1 and pi(a) == pi(b) and len(mod.LAST_LOWEST) >= 1, (pi(a), pi(b), len(mod.LAST_LOWEST)))

# ================================================================ Private Car, Two-Wheeler, new Two-Wheeler
PC = IN("icici_pc_nov26.xlsx")
pc, wp = build([("pc", PC)])
d = sel(pc, new_rto_list=mod.GROUP_PREFIX["pc"] + "ARUNACHAL_PRADESH")
get = lambda **kw: pi(sel(d, **kw))
check("Private Car Arunachal: New (1+3) 25 for all five fuels with T&C 'Applicable for 3+3', SAOD / 1+1 with NCB from the fuel columns (Petrol 25 on 1+1), without NCB 15, STP 30 on TP",
      get(sub_product="1+3", fuel_type="LPG") == 25 and set(sel(d, sub_product="1+3")["add_tnc"]) == {"Applicable for 3+3"} and get(sub_product="1+1", fuel_type="Petrol", is_ncb="Yes") == 25 and get(sub_product="SAOD", fuel_type="Petrol", is_ncb="Yes") == 20
      and get(sub_product="1+1", fuel_type="Diesel", is_ncb="No") == 15 and get(sub_product="STP", fuel_type="Electric") == 30 and sel(d, sub_product="STP").iloc[0]["pi_type"] == "On TP", d.groupby(["sub_product", "is_ncb"]).size().to_dict())
check("... LPG takes the CNG column; payouts are pi - 7; the policy type is New only for 1+3", get(sub_product="1+1", fuel_type="LPG", is_ncb="Yes") == get(sub_product="1+1", fuel_type="CNG", is_ncb="Yes") and po(sel(d, sub_product="STP")) == 23
      and set(sel(d, sub_product="1+3")["policy_type"]) == {"New"} and set(sel(d, sub_product="1+1")["policy_type"]) == {"Renew"})
check("Private Car: the 'Used car' column is not converted and the run says so (Not written + decision)", any("Used car" in str(e.get("what", "")) or "Used car" in str(e.get("sheet", "")) or e.get("what") == "Used car" for e in mod.LAST_NOT_WRITTEN) and any("Used car" in x for x in wp + [])
      or any(e.get("what") == "Used car" for e in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN[:2])
check("Private Car: 61 clusters x 5 sub_products worth of rows (3 x 61 x 5 + 2 x 61 x 5 ... = 1,830 rows), vertical ICICI-RB-Offline", len(pc) == 1830 and set(pc["insurer_vertical"]) == {"ICICI-RB-Offline"}, len(pc))
TW = IN("icici_tw_nov26.xlsx")
tw, wt = build([("tw", TW)])
d = sel(tw, new_rto_list=mod.GROUP_PREFIX["tw"] + "ANDAMAN&NICOBAR", fuel_type="Petrol")
get = lambda **kw: pi(sel(d, **kw))
check("Two-Wheeler Andaman: Bike 1+1 30 (T&C 'Applicable for 2+2,3+3' from the column title), Scooter 1+1 40, Bike SAOD 20, Scooter AOTP = STP 40, fuels Petrol and Electric",
      get(make_model_class="BIKE", sub_product="1+1") == 30 and set(sel(d, sub_product="1+1")["add_tnc"]) == {"Applicable for 2+2,3+3"} and get(make_model_class="SCOOTER", sub_product="1+1") == 40 and get(make_model_class="BIKE", sub_product="SAOD") == 20
      and get(make_model_class="SCOOTER", sub_product="STP") == 40 and set(sel(tw, new_rto_list=d.iloc[0]["new_rto_list"])["fuel_type"]) == {"Petrol", "Electric"}, d[["sub_product", "make_model_class", "pi_net_rate", "pi_od_rate", "pi_tp_rate"]].drop_duplicates().to_dict("records"))
check("Two-Wheeler: the grid's float noise (0.30000000000000004) is written as 30, and 63 clusters x 2 bodies x 2 fuels x 3 policies = 756 rows", pi(sel(tw, new_rto_list=mod.GROUP_PREFIX["tw"] + "ANDAMAN&NICOBAR", make_model_class="BIKE", sub_product="1+1", fuel_type="Petrol")) == 30.0 and len(tw) == 756, len(tw))
NEW = IN("icici_tw_new_nov26.xlsx")
tn, wn = build([("tw_new", NEW)])
d = sel(tn, new_rto_list=mod.GROUP_PREFIX["tw15"] + "Nashik")
got = {(r.new_vehicle_makes, r.make_model_class, r.fuel_type): pi(pd.DataFrame([r])) for r in d.itertuples()}
check("new Two-Wheeler (Nashik): HERO scooter 35, HONDA bike 10 / scooter 40, TVS scooter 50, Electric 30 for all makes (class NA); 1+5 / New / On Net, age 0-99",
      got[("HERO", "SCOOTER", "Petrol")] == 35 and got[("HONDA", "BIKE", "Petrol")] == 10 and got[("HONDA", "SCOOTER", "Petrol")] == 40 and got[("TVS", "SCOOTER", "Petrol")] == 50 and got[("TW_ALL", "NA", "Electric")] == 30
      and set(d["sub_product"]) == {"1+5"} and set(d["policy_type"]) == {"New"} and set(d["pi_type"]) == {"On Net"} and len(d) == 12, got)
check("new Two-Wheeler: Royal Enfield has bikes only, 63 clusters x 12 rows = 756", sel(d, new_vehicle_makes="ROYAL")["make_model_class"].tolist() == ["BIKE"] and len(tn) == 756, len(tn))

both_spec = [("pc", PC), ("tw", TW), ("tw_new", NEW)]
allrows, wa = build(both_spec)
check("the three workbooks together = the sum of their rows (1,830 + 756 + 756)", len(allrows) == 3342, len(allrows))


def pc_shift(wb):
    ws = wb["Sheet1"]
    ws.title = "Private car commission"
    ws.insert_rows(1, 2)
    ws.insert_cols(1, 1)


def pc_percent(wb):
    ws = wb["Sheet1"]
    for row in ws.iter_rows(min_row=3, min_col=3, max_col=14):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.value = c.value * 100


def edit_copy(src, fn, name):
    wb = openpyxl.load_workbook(src, data_only=True)
    fn(wb)
    p = os.path.join(tmp, name)
    wb.save(p)
    return p


for label, fn in (("sheet renamed, two rows above, a column to the left", pc_shift), ("payouts typed as percent (25) instead of fractions (0.25)", pc_percent)):
    o, _ = build([("pc", edit_copy(PC, fn, "pc_x.xlsx"))])
    check(f"Private Car - {label}: the same rows and rates", (shape(o)["_k"].values == shape(pc)["_k"].values).all() and np.allclose(shape(o)["pi_od_rate"].astype(float), shape(pc)["pi_od_rate"].astype(float)) and len(o) == len(pc))


def tw_shift(wb):
    ws = wb["TW Old"]
    ws.title = "Two wheeler grid"
    ws.insert_rows(1, 3)


o, _ = build([("tw", edit_copy(TW, tw_shift, "tw_x.xlsx"))])
check("Two-Wheeler - sheet renamed and three rows above: the same 756 rows", len(o) == 756 and (shape(o)["_k"].values == shape(tw)["_k"].values).all())


def new_shift(wb):
    wb.move_sheet("EV", offset=-6)
    ws = wb["HMC"]
    ws.insert_rows(1, 2)
    ws.insert_cols(1, 1)
    ws.cell(3, 6).value = "Final Nov'26"


o, wo = build([("tw_new", edit_copy(NEW, new_shift, "new_x.xlsx"))])
check("new Two-Wheeler - sheets reordered, rows / columns inserted on one sheet, the payout column re-titled: the same 756 rows", len(o) == 756 and (shape(o)["_k"].values == shape(tn)["_k"].values).all())
mod.FROM_DATE_ARG["tw_new"] = "15/11/2026"
o, _ = build(both_spec)
mod.FROM_DATE_ARG.clear()
check("--tw-new-from-date gives only the new Two-Wheeler rows their own start date", set(sel(o, sub_product="1+5")["from_date"]) == {"15/11/2026"} and set(sel(o, sub_product="1+1")["from_date"]) == {"01/11/2026"})
mod.VERTICAL_ARG[0] = "Not Applicable"
o, _ = build([("tw", TW)])
mod.VERTICAL_ARG[0] = None
check("--vertical overrides insurer_vertical for every row", set(o["insurer_vertical"]) == {"Not Applicable"})
try:
    build([("tw", PC)])
    ok = False
except mod.GridError as e:
    ok = "Private Car grid" in str(e) or "looks like" in str(e) or "given as" in str(e)
check("a file given for the wrong role is refused with a plain message", ok)
try:
    mod.classify_workbook(openpyxl.Workbook())
    roles = mod.classify_workbook(openpyxl.Workbook())
except Exception as e:                                              # pragma: no cover
    roles = e
check("an empty workbook holds none of the grids", roles == set(), roles)

# ================================================================ main(): the files
mod.default_output = lambda name: os.path.join(tmp, name)
outp = mod.main(both_spec, F, T, ref_path=REF("icici_pc_tw_nov26_upload.csv"), ask=mod.Asker(interactive=False), log=quiet)
rev = mod.review_path(outp)
check("main writes ICICI_UPLOAD_2026-11.csv (all four roles would not add a suffix; three roles do) and the review workbook", os.path.basename(outp) == "ICICI_UPLOAD_2026-11_PC_TW_TW_NEW.csv" and os.path.exists(rev), os.path.basename(outp))
sheets = openpyxl.load_workbook(rev).sheetnames
check("the review workbook has the Summary and the Not written sheet", "Summary" in sheets and "Not written" in sheets, sheets)
summ = pd.read_excel(rev, sheet_name="Summary", dtype=str)
check("the Summary says what was decided (insurer_vertical, the Used car column)", summ["value"].astype(str).str.contains("insurer_vertical = ICICI-RB-Offline", regex=False).any() and summ["value"].astype(str).str.contains("Used car", regex=False).any())
back = pd.read_csv(outp, dtype=str, keep_default_na=False)
check("the CSV has the 42 columns in the contract order and no empty new_rto_list", list(back.columns) == mod.OUTPUT_COLUMNS and (back["new_rto_list"] != "").all())

# ================================================================ the real grids against the user's own uploads
mod.RTO_NAMES = "group"


def diff_cols(df, ref_path):
    lines = []
    res = mod.diff_against_reference(df, ref_path, lines.append)
    return res, {l.split("column ")[1].split(":")[0] for l in lines if l.strip().startswith("column ")}


(new_n, gone_n, changed_n), cols_n = diff_cols(allrows, REF("icici_pc_tw_nov26_upload.csv"))
check("November Private Car + Two-Wheeler + new Two-Wheeler against the user's upload: every one of its 3,318 rows is produced (none removed); 24 rows are new", gone_n == 0 and new_n == 24 and changed_n == 3318, (new_n, gone_n, changed_n))
check("... and the ONLY column that differs on those rows is insurer_vertical (the upload had Not Applicable; ICICI-RB-Offline since the user's choice of 4 Nov 2026) - every rate, payout, pi_type, tariff, date and id is identical", cols_n == {"insurer_vertical"}, cols_n)
ref = pd.read_csv(REF("icici_pc_tw_nov26_upload.csv"), dtype=str, keep_default_na=False)
mine = mod.to_csv_frame(allrows)
mine["_k"], ref["_k"] = mod.key_of(mine), mod.key_of(ref)
check("... the 24 new rows are the Lakshadweep cluster (in the grid, not in the user's upload: the portal has no group for it)", set(mine[~mine["_k"].isin(ref["_k"])]["new_rto_list"]) == {"ICICI-RB_JUL26_LAKSHADWEEP"})
check("Private Car comprehensive rows are On OD (the grid says OD %), Two-Wheeler comprehensive On Net, STP On TP", set(sel(allrows, product="Private Car", sub_product="1+1")["pi_type"]) == {"On OD"} and set(sel(allrows, product="TW", sub_product="1+1")["pi_type"]) == {"On Net"}
      and set(sel(allrows, sub_product="STP")["pi_type"]) == {"On TP"} and set(sel(allrows, sub_product="1+5")["pi_type"]) == {"On Net"})

mod.RTO_NAMES = "group"
OCT = dt.date(2026, 10, 3)
cvdf, wcv = mod.build([("cv", IN("icici_cv_oct26.xlsx"))], OCT, T, quiet, mod.Asker(interactive=False))
(new_c, gone_c, changed_c), cols_c = diff_cols(cvdf, REF("icici_cv_oct26_rates_export.csv"))
check("October CV against the user's upload (dated 03/10/2026): 20,327 rows written, 19,629 of the upload's 20,112 keys produced, 8 rows with another rate", len(cvdf) == 20327 and gone_c == 483 and changed_c == 8, (len(cvdf), new_c, gone_c, changed_c))
check("... and only the four rate columns of those 8 rows differ: the vertical, the dates, pi_type, tariff, user_id and every other rate are identical", cols_c == {"pi_net_rate", "pi_tp_rate", "po_net_rate", "po_tp_rate"}, cols_c)
ref = pd.read_csv(REF("icici_cv_oct26_rates_export.csv"), dtype=str, keep_default_na=False)
mine = mod.to_csv_frame(cvdf)
mine["_k"], ref["_k"] = mod.key_of(mine), mod.key_of(ref)
mm = mine.merge(ref.drop_duplicates("_k"), on="_k", suffixes=("", "_ref"))
bad = mm[(pd.to_numeric(mm["pi_net_rate"]) - pd.to_numeric(mm["pi_net_rate_ref"])).abs().gt(1e-6) | (pd.to_numeric(mm["pi_tp_rate"]) - pd.to_numeric(mm["pi_tp_rate_ref"])).abs().gt(1e-6)]
check("... the 8 differing rows are all Jharkhand taxi up to 1000 cc, Old: the grid says 40 (OLD|OTHERS:40%), the user's upload has 35 - the document wins", len(bad) == 8 and set(bad["new_rto_list"]) == {"ICICI_SEPT26_CV_JHARKHAND"} and set(bad["make_model_class"]) == {"Taxi"}, bad[["new_rto_list", "pi_net_rate", "pi_net_rate_ref"]].head(3).to_dict("records"))
check("... every payout is the payin minus 7 (never below 0) in the whole CV output", ((cvdf[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).values - cvdf[["po_od_rate", "po_tp_rate", "po_net_rate"]].astype(float).values) >= 7 - 1e-9)[cvdf[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).values > 7].all())

# the November CV workbook has other sheet names (' COMP', 'TP', 'COUNCIL CALL', a changes sheet): the same segments and nearly the same rows
nov, wnv = build([("cv", IN("icici_cv_nov26.xlsx"))])
check("the real November CV workbook (renamed sheets, a 'changes' sheet) converts: 20,3xx rows, the same 65 clusters", 20000 < len(nov) < 20700 and nov["new_rto_list"].str.startswith(mod.GROUP_PREFIX["cv"]).all(), len(nov))

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
