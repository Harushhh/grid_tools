"""Deterministic checks of godigit_agency_to_config.py: the readers of the payout / discount cells, small workbooks (held in memory) through build() - one rule at a time - the password-protected
workbook, and the real October 2026 workbook against the user's October upload.
Run:  .venv\\Scripts\\python.exe tests\\godigit_tests.py"""
import datetime as dt
import io
import os
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("godigit_agency_to_config.py")
F, T = dt.date(2026, 10, 15), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


ask = mod.Asker(interactive=False)
M = mod.load_masters(ask)

# ================================================================ the payout cells (CD2)
check("CD2: a fraction is a percent, float noise removed", mod.cd2(0.07500000000000001) == ("rate", [7.5]) and mod.cd2(0.325) == ("rate", [32.5]))
check("CD2: '10%' and '29%/26%' (two discount bands)", mod.cd2("10%") == ("rate", [10.0]) and mod.cd2("29%/26%") == ("rate", [29.0, 26.0]))
check("CD2: D is a decline, an Excel error is a decline, blank is blank", mod.cd2("D") == ("D",) and mod.cd2("#VALUE!") == ("err",) and mod.cd2(None) == ("blank",) and mod.cd2("  ") == ("blank",))
check("CD2: 'With Addon: 10%/7.5%Without Addon: 2.5%' (no space before Without)", mod.cd2("With Addon: 10%/7.5%Without Addon: 2.5%") == ("addon", [10.0, 7.5], [2.5]))
check("CD2: '2.5%; TATA-With Addon: ...' is the TATA cell", mod.cd2("2.5%; TATA-With Addon: 10%/7.5%Without Addon: 7.5%") == ("tata", 2.5, [10.0, 7.5], [7.5]))
check("CD2: any other text is raw (it gets flagged, never guessed)", mod.cd2("as per policy")[0] == "raw")

# ================================================================ the discount texts (CD1)
sl = lambda v: mod.cd1_slots(v)[0]
check("CD1: '0-60%' (en dash too), '60%' and 0.85 are one cap", sl("0–60%") == [("cap", [60.0])] * 8 and sl("60%") == [("cap", [60.0])] * 8 and sl(0.85) == [("cap", [85.0])] * 8)
check("CD1: nothing named = the whole range", sl(None) == [("cap", [100.0])] * 8)
x = sl("New/1yr: D | 2-5yr: 80% | 6yr+: 85%/90%")
check("CD1: a cap per age group, D = declined, '85%/90%' = two bands", x[:2] == [("D",)] * 2 and x[2:6] == [("cap", [80.0])] * 4 and x[6:] == [("cap", [85.0, 90.0])] * 2, x)
x = sl("New/1yr: D | 2-5yr: 80%")
check("CD1: an age group the text does not name continues the previous cap (6yr+ after 2-5yr: 80%)", x[6:] == [("cap", [80.0])] * 2, x)
x = sl("2-5yr: 85% | 6yr+: 90%")
check("CD1: an age group before the first named one is declined (New/1yr)", x[:2] == [("D",)] * 2, x)
x = sl("New/1yr: 60% | 2-4yr: 85% | 5yr+: with NCB/break in > 90 days : 90%, without NCB : 80%")
check("CD1: the NCB / break-in clause", x[5] == ("ncb", 90.0, 80.0) and x[4] == ("cap", [85.0]), x)
x = sl("With Addon: 85%/90%\nWithout Addon: 60%")
check("CD1: with / without add-on", x[0] == ("addon", [85.0, 90.0], [60.0]), x)
x = sl("Age 0: 60%\nAge>=1: 75%")
check("CD1: 'Age 0: 60% / Age>=1: 75%'", x[0] == ("cap", [60.0]) and x[1:] == [("cap", [75.0])] * 7, x)
x = sl("60% for age 0-1, 85% for age 2-5, 90% for age 6+")
check("CD1: '60% for age 0-1, 85% for age 2-5, 90% for age 6+'", x[1] == ("cap", [60.0]) and x[5] == ("cap", [85.0]) and x[7] == ("cap", [90.0]), x)
x, tata = mod.cd1_slots("New/1yr: 60% | 2-5yr: 80% | 6yr+: 85%/90%; TATA age 0 - With Addon: 85%/90%\nWithout Addon: 60%")
check("CD1: the 'TATA age 0 - ...' clause is split off", x is not None and x[0] == ("cap", [60.0]) and tata and tata.startswith("With Addon"), (x, tata))
check("CD1: a text it does not know is None (the row is flagged)", sl("only for hilly regions") is None)

# ================================================================ weights, bodies, two-wheeler / car helpers
check("weight bands of the segment text", mod.gvw_of("GCV4 12 to 20T") == (12000.01, 20000.01) and mod.gvw_of("GCV4 44T+") == (44000.01, 99999) and mod.gvw_of("GCV4 upto 1.6T") == (0, 1600.01)
      and mod.gvw_of("GCV4 40T+") == (40000.01, 99999) and mod.gvw_of("GCV4 nonsense") is None)
check("body type -> class and T&C", mod.body_class("Non-Dumper/Tipper", "All") == ("Truck", "") and mod.body_class("Dumper/Tipper", "All") == ("Dumper Tipper", "")
      and mod.body_class("Oil Tanker", "All") == ("Tanker", "Oil Tanker") and mod.body_class("Port Trailer", "All") == ("Trailer", "Port Trailer") and mod.body_class("LCV", "All make") == ("NA", "")
      and mod.body_class("—", "Oil Tanker") == ("Tanker", "Oil Tanker"))
check("TW/PC discount limit: 0.4 = 40, 0 = no limit, negative = declined", mod.cd1_cap(0.4) == (40.0, "") and mod.cd1_cap(0) == (100.0, "") and mod.cd1_cap(-2) == (None, "") and mod.cd1_cap(None) == (100.0, ""))
cap, note = mod.cd1_cap(0.4979)
mod.CD1_ROUNDED.clear()
check("TW/PC discount limit: a cell the sheet shows as a whole percent ('0%') is written as the sheet shows it - 49.9915 -> 50, 39.9951 -> 40, 60.2443 -> 60, 49.79 -> 50 - silently",
      [mod.cd1_cap(v, True) for v in (0.499915, 0.399951, 0.600244, 0.4979)] == [(50.0, "")] * 1 + [(40.0, ""), (60.0, ""), (50.0, "")], [mod.cd1_cap(v, True) for v in (0.499915, 0.399951, 0.600244, 0.4979)])
check("TW/PC discount limit: a cell with decimals ('0.00%') is written as it is - 67.5 stays 67.5 (the first build wrote 68); any other value keeps its own digits like every other rate",
      mod.cd1_cap(0.675, False) == (67.5, "") and mod.cd1_cap(0.499915, False) == (mod.rnd(49.9915), ""), (mod.cd1_cap(0.675, False), mod.cd1_cap(0.499915, False)))
check("TW/PC discount limit: Excel shows 67.5 in a '0%' cell as 68 (half up), so the limit is 68 there", mod.cd1_cap(0.675, True) == (68.0, ""))
check("TW/PC discount limit: format unknown (an old .xls): a multiple of half a point is a real limit (67.5), formula noise is the whole percent the sheet shows",
      mod.cd1_cap(0.675) == (67.5, "") and mod.cd1_cap(0.4979) == (50.0, "") and mod.cd1_cap(0.29) == (29.0, "") and mod.cd1_cap(0.4) == (40.0, ""), (mod.cd1_cap(0.675), mod.cd1_cap(0.4979)))
_mode = mod.DISCOUNT_ROUNDING
mod.DISCOUNT_ROUNDING = "down"
cd_down = mod.cd1_cap(0.4979, True)
mod.DISCOUNT_ROUNDING = "nearest"
cd_near, cd_note = mod.cd1_cap(0.4979), mod.cd1_cap(0.4979)[1]
cd_half = mod.cd1_cap(0.675)
mod.DISCOUNT_ROUNDING = _mode
check("TW/PC discount limit: DISCOUNT_ROUNDING = 'down' (always the whole percent below) and 'nearest' (the first build, said so on the rows, 67.5 -> 68) stay available",
      cd_down == (49.0, "") and cd_near[0] == 50.0 and "49.79" in cd_note and cd_half[0] == 68.0, (cd_down, cd_near, cd_half))
import io as _io
import openpyxl as _ox
_wb = _ox.Workbook()
_ws = _wb.active
_ws["A1"], _ws["B1"], _ws["C1"] = 0.499915, 0.675, 0.4
_ws["A1"].number_format, _ws["B1"].number_format, _ws["C1"].number_format = "0%", "0.00%", "0%"
_buf = _io.BytesIO()
_wb.save(_buf)
_book = mod._book_openpyxl(_buf.getvalue())
_sh = _book.sheets[0]
check("the workbook reader keeps which cells the sheet shows as whole percents ('0%' yes, '0.00%' no)", _sh.shows_whole(1, 1) is True and _sh.shows_whole(1, 2) is False and _sh.shows_whole(1, 3) is True, _sh.whole_pct)
check("a sheet read without formats (an old .xls) says 'not known'", mod.Sheet("x", [[0.5]]).shows_whole(1, 1) is None)
check("1+5 segments: 'MC <=155', 'MC >155', '< =350', '> 350', '< 3 KW', '3-7 KW'", [mod.tw15_seg(s) for s in ("MC <=155", "MC >155", "< =350", "> 350", "< 3 KW", ">7 KW", "3-7 KW", "SCOOTER")]
      == ["mc155", "mc155_", "350", "350_", "3kw", "kw7", "37kw", "scooter"])
check("1+5 targets: Bajaj 3-7 KW = Electric scooter 3.01-7.01; Others MC <=155 = the 'others' bike group; an unknown pair = nothing",
      mod.tw15_targets("BAJAJ", "3-7 KW") == [("SCOOTER", "BAJAJ", (3.01, 7.01), ["Electric"])]
      and mod.tw15_targets("Others", "MC <=155")[0][1] == mod.BIKES_OTHER and mod.tw15_targets("HERO MOTOCORP", "All") == [])
check("PC cc bands", mod.pc_cc("petrol", "<1000") == (0, 1000.01) and mod.pc_cc("petrol", "1000-1500") == (1000.01, 1500.01) and mod.pc_cc("cng", None) == (0, 99999) and mod.pc_cc("diesel", ">1500") == (1500.01, 99999))
check("the HCV note of the sheet 'CV Remarks' is read",
      mod.read_hcv_rule(["Note: Volvo, Scania to block. \nOnly for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1"]) == dict(blocked=True, cap=70.0, body="nondumpertipper"))


# ================================================================ small workbooks through build()
mod.MISSING_TABLES = "skip"        # most small books below have no two-wheeler / car sheet: the script would ask whether to write the file without them (tested further down with "ask")
HEAD =["Cluster segment", "Segment", "Body Type", "RTO Cluster", "State", "Make / Variant", "CD1 (OD-disc)"] + ["New", "1 Yr", "2 Yr", "3 Yr", "4 Yr", "5 Yr", "6-10 Yr", "11+ Yr"] * 2
HCV_NOTE = "Note: Volvo, Scania to block. \nOnly for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1"


def book(cv_rows, remarks=True, extra=()):
    sheets = [mod.Sheet("LATEST CV COMBINED GRID", [[], ["", "", "", "", "", "", "", "COMPREHENSIVE"], ["", "", "", "", "", "", "", "CD2"], HEAD] + [list(r) for r in cv_rows])]
    if remarks:
        sheets.append(mod.Sheet("CV Remarks", [["REMARKS FOR SCHOOL BUS"], ["1. Short Term Policies - CD2 reduction of 25% on the above Rate"], ["REMARKS FOR STAFF BUS"],
                                               ["Note:"], ["For Seating Capacity up to 10-seater, CD1_85% & CD2_20% in all open cohorts."], ["REMARKS FOR HCV"], [HCV_NOTE]]))
    sheets.extend(extra)
    sheets.append(mod.Sheet("Hidden old grid", [["x"]], visible=False))
    return mod.Book(sheets)


def run(b):
    mod.reset_run_lists()
    df, warnings = mod.build(None, F, T, quiet, mod.Asker(interactive=False), book=b)
    return df, warnings


def rows_of(df, **kw):
    m = pd.Series(True, index=df.index)
    for k, v in kw.items():
        m &= (df[k] == v)
    return df[m]


NA8 = [None] * 8
hcv_dump = ("CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", "BAD GJ", "Gujarat", "All", "New/1yr: 60% | 2-4yr: 85% | 5yr+: 85%/90%", 0.25, 0.25, 0.25, 0.25, 0.25, "29%/26%", "29%/26%", "10%/10%",
            0.325, 0.325, 0.325, 0.325, 0.325, 0.41, 0.41, 0.41)
df, w = run(book([hcv_dump]))
d = rows_of(df, fuel_type="Diesel", sub_product="1+1")
check("HCV dumper: 5 fuels, comp 1+1 On Net, STP On TP, PO = PI - 7", set(df.fuel_type) == {"Petrol", "Diesel", "CNG", "LPG", "Electric"} and set(df.pi_type) == {"On Net", "On TP"}
      and (rows_of(df, sub_product="STP").pi_tp_rate - rows_of(df, sub_product="STP").po_tp_rate).round(4).isin([7.0]).all())
check("HCV: the make group without Volvo (CV Remarks 'Volvo, Scania to block') and the RTO group in the master's spelling",
      set(df.new_vehicle_makes) == {"GCV_4W_ALL_EX_VOLVO"} and set(df.new_rto_list) == {"GO-DIGIT_OCT26_HCV_BAD_GJ"}, (set(df.new_vehicle_makes), set(df.new_rto_list)))
check("discount caps by age group: New/1yr 60.01, 2-4yr 85.01", set(d[d.vehicle_age_min.isin([0, 1])].tariff_max) == {60.01} and set(d[d.vehicle_age_min.isin([2, 3, 4])].tariff_max) == {85.01})
a5 = d[(d.vehicle_age_min == 5)].sort_values("tariff_max")
check("a pair '29%/26%' with '85%/90%' = two bands that chain (0-85.01 at 29, 85.01-90.01 at 26)", list(zip(a5.tariff_min, a5.tariff_max, a5.pi_net_rate)) == [(0.0, 85.01, 29.0), (85.01, 90.01, 26.0)], a5[["tariff_min", "tariff_max", "pi_net_rate"]])
a7 = d[(d.vehicle_age_min == 10.01)]
check("equal pair '10%/10%' = one band up to the higher cap", list(zip(a7.tariff_min, a7.tariff_max, a7.pi_net_rate)) == [(0.0, 90.01, 10.0)], a7[["tariff_min", "tariff_max", "pi_net_rate"]])
check("the first age column carries add_tnc 'New'; the SATP rows are 0-100.01 with is_ncb / is_zd NA",
      set(d[d.vehicle_age_min == 0].add_tnc) == {"New"} and set(d[d.vehicle_age_min == 1].add_tnc) == {""}
      and set(rows_of(df, sub_product="STP").tariff_max) == {100.01} and set(rows_of(df, sub_product="STP").is_zd) == {"NA"})
check("age bands: New 0-1 ... 5 Yr 5-6, 6-10 Yr 6-10.01, 11+ Yr 10.01-99", sorted(set(zip(d.vehicle_age_min, d.vehicle_age_max))) == [(0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6), (6, 10.01), (10.01, 99)])

truck = ("CV HCV", "GCV4 12 to 20T", "Non-Dumper/Tipper", "BAD GJ", "Gujarat", "All", "New/1yr: D | 2-4yr: D | 5yr+: 85%", "D", "D", "D", "D", "D", 0.075, 0.075, 0.075,
         "D", "D", "D", "D", "D", 0.325, 0.325, 0.325)
df, w = run(book([truck]))
d = rows_of(df, fuel_type="Diesel", sub_product="1+1")
check("Truck: Mahindra / Bharat Benz rows at 70% CD1 and every other make (without Volvo) at the grid's own bands",
      set(d.new_vehicle_makes) == {"GCV_4W_ALL_EX_VOLVO_TATA_MAHINDRA_BHARAT,TATA", "Mahindra,Bharat_Benz"}
      and set(d[d.new_vehicle_makes == "Mahindra,Bharat_Benz"].tariff_max) == {70.01} and set(d[d.new_vehicle_makes != "Mahindra,Bharat_Benz"][d.vehicle_age_min == 5].tariff_max) == {85.01},
      d.groupby("new_vehicle_makes").tariff_max.unique())
z = d[(d.vehicle_age_min <= 4) & (d.new_vehicle_makes != "Mahindra,Bharat_Benz")]
check("D (and a declined age group) = a zero card over the whole discount range (100.01)", len(z) == 5 and (z.pi_net_rate == 0).all() and set(z.tariff_max) == {100.01})
s = rows_of(df, fuel_type="Diesel", sub_product="STP")
check("Truck SATP: one make group for all makes but Volvo; D = zero card", set(s.new_vehicle_makes) == {"GCV_4W_ALL_EX_VOLVO"} and (s[s.vehicle_age_min <= 4].pi_tp_rate == 0).all() and len(s[s.vehicle_age_min <= 4]) == 5)

errrow = ("CV HCV", "GCV4 20 to 40T", "Dumper/Tipper", "BAD GJ", "Gujarat", "All", "New/1yr: 60% | 2-4yr: 85% | 5yr+: 85%/90%", 0.2, 0.2, 0.2, 0.2, 0.2, "#VALUE!", "#VALUE!", "#VALUE!", 0.3, 0.3, 0.3, 0.3, 0.3, 0.4, 0.4, 0.4)
df, w = run(book([errrow]))
e = rows_of(df, fuel_type="Diesel", sub_product="1+1", vehicle_age_min=5)
check("an Excel error cell (#VALUE!) = a zero card over the whole range, and the run says how many", list(zip(e.tariff_min, e.tariff_max, e.pi_net_rate)) == [(0.0, 100.01, 0.0)] and any("#VALUE!" in x for x in w), (e, w))

eload = [("CV GCV3(EV/Non-EV)", "E-Loaders", "—", "BAD GJ", "Gujarat", "All make", "0–50%") + (0.3,) * 6 + (0.4, 0.4) + (0.275,) * 8,
         ("CV GCV3(EV/Non-EV)", "E-Loaders", "—", "BAD GJ", "Gujarat", "All make", "0–85%") + (0.15,) * 6 + (0.25, 0.25) + (0.275,) * 8]
df, w = run(book(eload))
c0 = rows_of(df, sub_product="1+1", vehicle_age_min=0).sort_values("tariff_max")
check("E-Loaders: only Electric; two grid rows of one entry chain their bands (0-50.01 at 30, 50.01-85.01 at 15), add_tnc 'New,Loaders'",
      set(df.fuel_type) == {"Electric"} and list(zip(c0.tariff_min, c0.tariff_max, c0.pi_net_rate)) == [(0.0, 50.01, 30.0), (50.01, 85.01, 15.0)] and set(c0.add_tnc) == {"New,Loaders"}, c0[["tariff_min", "tariff_max", "pi_net_rate"]])
st = rows_of(df, sub_product="STP", vehicle_age_min=0)
check("the same SATP payin from two grid rows is ONE row", len(st) == 1 and st.pi_tp_rate.iloc[0] == 27.5 and not mod.LAST_LOWEST, (len(st), mod.LAST_LOWEST))

two = [("CV GCV3(EV/Non-EV)", "GCV3", "—", "BAD GJ", "Gujarat", "All make", "0–60%") + (0.1,) * 8 + (0.1,) * 8,
       ("CV GCV3(EV/Non-EV)", "GCV3", "—", "BAD GJ", "Gujarat", "All make", "0–60%") + (0.2,) * 8 + (0.2,) * 8]
df, w = run(book(two))
check("two grid rows, one entry, two payins: the LOWEST (10) is written and the entry is listed", set(rows_of(df, sub_product="1+1").pi_net_rate) == {10.0} and len(mod.LAST_LOWEST) > 0
      and mod.LAST_LOWEST[0]["rates_found"] == "10 / 20" and mod.LAST_LOWEST[0]["written"] == "10", mod.LAST_LOWEST[:1])

pa = ("CV PCV2W(All) 3W Non-electric", "PCV3W non-diesel", "—", "BAD GJ", "Gujarat", "All make", "0–60%", "D", "D", "D", 0.07, 0.07, 0.07, None, None, "D", "D", "D", 0.07, 0.07, 0.07, None, None)
pb = ("CV PCV2W(All) 3W Non-electric", "PCV3W non-diesel", "—", "BAD GJ", "Gujarat", "All make", "0–85%", "D", "D", "D", None, None, None, 0.6, 0.6, "D", "D", "D", None, None, None, 0.6, 0.6)
df, w = run(book([pa, pb]))
s2 = rows_of(df, product="PCV 3W", fuel_type="Diesel", sub_product="STP").sort_values("vehicle_age_min")
check("PCV 2W and PCV 3W get the same rates; fuels are the four non-electric ones", set(df["product"]) == {"PCV 2W", "PCV 3W"} and set(df.fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"})
check("a blank cell is a zero placeholder that never beats a real rate of another grid row (SATP 3-5 yr = 7, 6+ = 60), D = zero card",
      dict(zip(s2.vehicle_age_min, s2.pi_tp_rate)) == {0: 0.0, 1: 0.0, 2: 0.0, 3: 7.0, 4: 7.0, 5: 7.0, 6: 60.0, 10.01: 60.0}, dict(zip(s2.vehicle_age_min, s2.pi_tp_rate)))
cm = rows_of(df, product="PCV 3W", fuel_type="Diesel", sub_product="1+1")
check("the bands of the two grid rows chain, both with a full set of ages (0-60.01 then 60.01-85.01)", sorted(set(zip(cm.tariff_min, cm.tariff_max))) == [(0.0, 60.01), (60.01, 85.01)]
      and len(cm) == 16 and cm[(cm.tariff_max == 85.01) & (cm.vehicle_age_min == 6)].pi_net_rate.iloc[0] == 60.0)

taxi = [("CV PCV4(Upto 5 seating)", "Taxi upto 5 seater non-diesel < 1000 cc", "With Add on cover", "Ahmedabad", "Gujarat", "Maruti & Hyundai (excl WagonR Model)", "0-45%") + (0.075,) * 16,
        ("CV PCV4(Upto 5 seating)", "Taxi upto 5 seater non-diesel < 1000 cc", "Without Add on cover", "Ahmedabad", "Gujarat", "Maruti & Hyundai (excl WagonR Model)", "0-50%") + (0.075,) * 16]
df, w = run(book(taxi))
t1 = rows_of(df, sub_product="1+1", vehicle_age_min=2, fuel_type="Petrol")
check("Taxi: add-on cover = is_zd YES / NO, each with its own band from 0; SATP carries no ZD (one row, NA)",
      sorted(zip(t1.is_zd, t1.tariff_min, t1.tariff_max)) == [("NO", 0.0, 50.01), ("YES", 0.0, 45.01)] and set(rows_of(df, sub_product="STP").is_zd) == {"NA"}
      and len(rows_of(df, sub_product="STP", vehicle_age_min=2, fuel_type="Petrol")) == 1, t1[["is_zd", "tariff_min", "tariff_max"]])
check("Taxi: sc 0-5.01, class Taxi, non-diesel fuels CNG / LPG / Petrol, make group in the master's spelling",
      set(df.sc_max) == {5.01} and set(df.make_model_class) == {"Taxi"} and set(df.fuel_type) == {"CNG", "LPG", "Petrol"} and set(df.new_vehicle_makes) == {"MARUTI_EX_WAGONR,Hyundai"} and set(df.cc_max) == {1000.01},
      (set(df.new_vehicle_makes), set(df.fuel_type)))

staff = ("Staff Bus", "Staff Bus", "CORPORATE (SELF-USE) ", "ROM", "MAHARASHTRA", ">10", "60%") + (0.4,) * 16
school = ("School Bus", "School Bus", "IN SCHOOL'S NAME · Vol ≤ ₹3L", "Bangalore", "Karnataka", "8 & above", "85%") + (0.3,) * 16
df, w = run(book([staff, school]))
check("Staff Bus: Other Bus, BUS_ALL, seating 10.01-999, add_tnc = the body text, NO Electric (the sheet says non-electric buses)",
      set(df.fuel_type) == {"CNG", "Diesel", "LPG", "Petrol"} and set(df.make_model_class) == {"Other Bus"} and set(df.sc_min) == {10.01} and "CORPORATE (SELF-USE)" in set(df.add_tnc.str.replace("New,", "")))
check("School Bus is not written and is listed (WRITE_SCHOOL_BUS = False)", not df.new_rto_list.str.contains("SCHOOL").any() and any("School Bus" in x["segment"] for x in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN[:1])
check("the staff-bus note 'up to 10-seater' is no flag on the staff-bus rows (they are written for MORE than 10 seats): it is listed once in 'Not written' (10 Nov 2026; tests/godigit_cd1_tests.py has the rest)",
      (df.remarks == "").all() and len(df) > 0 and any(x["segment"] == "Staff Bus up to 10 seats" for x in mod.LAST_NOT_WRITTEN), (df.remarks.unique()[:2], mod.LAST_NOT_WRITTEN[:1]))

odd = ("CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", "BAD GJ", "Gujarat", "All", "only for hilly regions", 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3)
df, w = run(book([odd]))
check("a discount text it does not know: written for the whole range (100.01) and flagged on exactly those rows",
      set(df.tariff_max) == {100.01} and df.remarks.str.contains("only for hilly regions").all(), df.remarks.iloc[0][:120])

tata = ("CV HCV", "GCV4 12 to 20T", "Non-Dumper/Tipper", "BAD GJ", "Gujarat", "All", "New/1yr: 60% | 2-5yr: 80% | 6yr+: 85%/90%; TATA age 0 - With Addon: 85%/90%\nWithout Addon: 60%",
        "2.5%; TATA-With Addon: 10%/7.5%Without Addon: 7.5%", 0.025, 0.13, 0.13, 0.13, 0.13, "23%/21%", "23%/21%", 0.1, 0.1, 0.21, 0.21, 0.21, 0.21, 0.325, 0.325)
df, w = run(book([tata]))
t0 = rows_of(df, fuel_type="Diesel", sub_product="1+1", vehicle_age_min=0, new_vehicle_makes="TATA").sort_values(["is_zd", "tariff_max"])
check("TATA age-0 clause: Tata gets is_zd YES (0-85.01 at 10, 85.01-90.01 at 7.5) and NO (0-60.01 at 7.5); the other makes keep the base 2.5",
      list(zip(t0.is_zd, t0.tariff_min, t0.tariff_max, t0.pi_net_rate)) == [("NO", 0.0, 60.01, 7.5), ("YES", 0.0, 85.01, 10.0), ("YES", 85.01, 90.01, 7.5)]
      and set(rows_of(df, fuel_type="Diesel", sub_product="1+1", vehicle_age_min=0, new_vehicle_makes="GCV_4W_ALL_EX_VOLVO_TATA_MAHINDRA_BHARAT").pi_net_rate) == {2.5}, t0[["is_zd", "tariff_min", "tariff_max", "pi_net_rate"]])

# an RTO group the master lacks: kept by the name it would have, flagged, listed
miss = ("CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", "NO SUCH CLUSTER", "Gujarat", "All", "60%") + (0.25,) * 8 + (0.3,) * 8
mapping = mod.Sheet("Latest CV RTO Mapping", [["RTO Code", "City", "CV GCV3(EV/Non-EV)", "CV LCV", "CV MCV", "CV HCV"], ["GJ01", "A", "x", "x", "x", "NO SUCH CLUSTER"], ["GJ02", "B", "x", "x", "x", "NO SUCH CLUSTER"]])
df, w = run(book([miss], extra=[mapping]))
mod.collect_gaps(df)
gap = mod.EXTRA_SHEETS.get("RTO groups to create")
check("a cluster whose RTO group the master lacks: the name it would have, flagged on its rows, listed WITH the RTO codes of the mapping sheet",
      set(df.new_rto_list) == {"GO-DIGIT_OCT26_HCV_NO_SUCH_CLUSTER"} and df.remarks.str.contains("is not in the RTO master").all() and gap is not None and gap.iloc[0]["items"] == "GJ01,GJ02", gap)

# ---------------------------------------------------------------- two-wheelers and cars (small sheets)
tw11 = mod.Sheet("Latest 2W 1+1 & SATP", [["", "", "1+1", "", "SATP"], ["", "", "", "", "With CPA", "Without CPA"], ["Agency/PB Clusters", "Agency/PB Seg", "CD1", "CD2", "CD2", "CD2"],
                                          ["GJ_ABS", "SC/EV", 0.6, 0.59, 0.635, 0.61], ["GJ_Bad", "MC_180-350_RE", -2, 0.365, 0.285, 0.26], ["GJ_Bad", "MC <= 180 Others", 0, 0.315, 0.265, 0.24],
                                          ["GJ_Bad", "MC_new_text", 0.1, 0.1, 0.1, 0.1]])
df, w = run(book([hcv_dump], extra=[tw11]))
tw = df[df["product"] == "TW"]
sc = rows_of(tw, new_rto_list="GO-DIGIT_OCT26_2W_1+1_TP_GJ_ABS")
check("TW 1+1: scooter SC/EV both fuels, 0-60.01 at 59; SATP with CPA 63.5 / without 61, tariff 0-100.01",
      set(rows_of(sc, sub_product="1+1").fuel_type) == {"Petrol", "Electric"} and set(rows_of(sc, sub_product="1+1").pi_net_rate) == {59.0} and set(rows_of(sc, sub_product="1+1").tariff_max) == {60.01}
      and sorted(zip(rows_of(sc, sub_product="STP", fuel_type="Petrol").is_cpa, rows_of(sc, sub_product="STP", fuel_type="Petrol").pi_tp_rate)) == [("NO", 61.0), ("YES", 63.5)])
re_ = rows_of(tw, new_rto_list="GO-DIGIT_OCT26_2W_1+1_TP_GJ_BAD", make_model_class="BIKE", new_vehicle_makes="ROYAL", fuel_type="Petrol")
check("TW: a negative CD1 declines the comprehensive card (0 at 100.01) but the SATP rates stay; CD1 0 = no limit (100.01) with the real rate",
      set(rows_of(re_, sub_product="1+1").pi_net_rate) == {0.0} and set(rows_of(re_, sub_product="1+1").tariff_max) == {100.01} and sorted(set(rows_of(re_, sub_product="STP").pi_tp_rate)) == [26.0, 28.5]
      and set(rows_of(tw, new_vehicle_makes="BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,TVS,SUZUKI", sub_product="1+1").pi_net_rate) == {31.5}
      and set(rows_of(tw, new_vehicle_makes="BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,TVS,SUZUKI", sub_product="1+1").tariff_max) == {100.01})
check("TW: a segment it does not know is listed, not guessed", any("MC_new_text" in x["segment"] for x in mod.LAST_NOT_WRITTEN) if False else any(x.get("segment") == "MC_new_text" for x in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN)

pc_comp = mod.Sheet("PVT Car Comp+SAOD", [[], [None, None, None, None, "SAOD", None, "Comp"], [None, None, "Cluster", "Volume Slab", "Non NCB", "NCB", "Non NCB", "NCB"],
                                          [None, None, "BLR", "0-25 K", 17.5, 21, 19.5, 23], [None, None, "BLR", "1-3 Lac", 18.5, 23, 21, 26]])
pc_tp = mod.Sheet(" PVT Car TP", [[], [], [None, None, None, "Cluster", "Segment", "Age", "MaxCD2", "Notes"], [None, None, None, "Bangalore", "CNG", "All", 0, None],
                                  [None, None, None, "Bangalore", "Petrol<1000", "<10", 0.155, None], [None, None, None, "Bangalore", "Petrol<1000", ">10", 0.185, "Few RTOs for 10+ age are declined"],
                                  [None, None, None, "Mumbai2", "Diesel>1500", "All", 0.2, None]])
df, w = run(book([hcv_dump], extra=[pc_comp, pc_tp]))
pc = df[df["product"] == "Private Car"]
c11 = rows_of(pc, sub_product="1+1", fuel_type="Petrol")
check("Private Car 1+1 / SAOD: the volume slab of the deal only (1-3 Lac), NCB YES / NO, On OD, one row per non-electric fuel",
      sorted(zip(c11.is_ncb, c11.pi_od_rate)) == [("NO", 21.0), ("YES", 26.0)] and sorted(zip(rows_of(pc, sub_product="SAOD", fuel_type="Diesel").is_ncb, rows_of(pc, sub_product="SAOD", fuel_type="Diesel").pi_od_rate)) == [("NO", 18.5), ("YES", 23.0)]
      and set(rows_of(pc, sub_product="1+1").fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"} and set(c11.new_rto_list) == {"GO-DIGIT_PC_1+1_OCT26_BLR"}, c11[["new_rto_list", "is_ncb", "pi_od_rate"]])
tp = rows_of(pc, sub_product="STP")
p1 = rows_of(tp, fuel_type="Petrol", cc_max=1000.01).sort_values("vehicle_age_min")
check("Private Car SATP: Petrol<1000 <10 / >10 = ages 0-10.01 / 10.01-99; Electric takes the Petrol rate, LPG the CNG rate, notes are flagged",
      list(zip(p1.vehicle_age_min, p1.vehicle_age_max, p1.pi_tp_rate)) == [(0, 10.01, 15.5), (10.01, 99, 18.5)] and set(rows_of(tp, fuel_type="Electric", cc_max=1000.01).pi_tp_rate) == {15.5, 18.5}
      and set(rows_of(tp, fuel_type="LPG").pi_tp_rate) == {0.0} and tp.remarks.str.contains("Few RTOs for 10+ age are declined", regex=False).any(), p1[["vehicle_age_min", "pi_tp_rate"]])
check("Private Car SATP: the grid's 'Mumbai2' is the master's 'Mumbai_2' (CLUSTER_NAME_EXCEPTIONS)", "GO-DIGIT_PC_TP_OCT26_Mumbai_2" in set(tp.new_rto_list), set(tp.new_rto_list))

# ================================================================ tables found by what their headers SAY
# (10 Nov 2026: Go Digit renamed every sheet of the November workbook, put 'Avg CD2' in front of 'Max CD2', moved the car note column and added a block to the left of the car sheet; a reader that
# counts cells from a header read the wrong column without a word.  These small sheets rebuild each of those shapes.)
_rule = mod.PAYOUT_MEASURE
mod.PAYOUT_MEASURE = None          # the question path first (None = ask); the user's standing answer (PAYOUT_MEASURE = "avg") is tested after the November workbook


def asker(*picks):
    """an interactive Asker that answers the numbered questions with `picks`; `seen` = every prompt and line it showed"""
    seen, it = [], iter(picks)

    def fn(prompt):
        seen.append("PROMPT " + prompt)
        try:
            return next(it)
        except StopIteration:
            raise EOFError
    return mod.Asker(interactive=True, input_fn=fn, out=lambda *x: seen.append(" ".join(map(str, x)))), seen


def run_ask(b, ask):
    mod.reset_run_lists()
    return mod.build(None, F, T, quiet, ask, book=b)


def at(df, cluster):
    """the rows of one cluster whatever the RTO name mode (this suite runs in 'group' mode: GO-DIGIT_..._<cluster>)"""
    return df[df["new_rto_list"].str.upper().str.endswith("_" + cluster.upper())]


def sheet(name, above, head, rows):
    return mod.Sheet(name, [list(r) for r in above] + [list(head)] + [list(r) for r in rows])


T11_ROWS = [("GJ_ABS", "SC/EV", 0.6, 0.65, 0.59, 0.7, 0.635), ("GJ_Bad", "MC <= 180 Others", 0.4, 0.35, 0.315, 0.3, 0.265)]       # cluster, segment, CD1, 1+1 Max, 1+1 Avg, SATP Max, SATP Avg
T15_ROWS = [("GJ_ABS", "BAJAJ", "3-7 KW", 0.2, 0.325, 0.35, "Slab Net 5"), ("GJ_ABS", "HONDA", "MC <=155", 0.4, 0.6, 0.65, "Slab OD+Add on")]   # cluster, make, segment, CD1, Avg, Max, formula


def t11_nov(name="New 2W Comp & SATP wef 1st Nov"):
    return sheet(name, [[None, None, None, "1+1"], [None, None, None, None, None, None, "SATP"]], [None, "Agency/PB Clusters", "Agency/PB Seg", "CD1", "Max CD2", "Avg CD2", "Max CD2", "Avg CD2"],
                 [[c + s, c, s, *rest] for c, s, *rest in T11_ROWS])


def t15_nov(name="2W Grid 1+5"):
    return sheet(name, [[None, None, None, None, "1+5"]], ["Key", "Agency/PB Clusters", "Make", "Agency/PB Seg", "CD1", "Avg CD2", "Max CD2", "Formula Type"], [[c + m + s, c, m, s, *rest] for c, m, s, *rest in T15_ROWS])


# ---- the exact November failure: the 1+1 sheet renamed, the 1+5 sheet sharing its headers, two payout columns, no CPA split
nonint_msg = ""
try:
    run(book([hcv_dump], extra=[t11_nov(), t15_nov()]))
except mod.GridError as e:
    nonint_msg = str(e)
check("November 1+1 / 1+5 sheets: found by their headers (no 'more than one sheet fits'); two payout columns without a keyboard stop with a question that names both",
      "Max CD2" in nonint_msg and "Avg CD2" in nonint_msg and "more than one sheet" not in nonint_msg, nonint_msg)
for pick, label, want in (("2", "Avg", dict(c11=59.0, s11=63.5, o11=31.5, os11=26.5, b15=32.5, h15=60.0)), ("1", "Max", dict(c11=65.0, s11=70.0, o11=35.0, os11=30.0, b15=35.0, h15=65.0))):
    a, seen = asker(pick)
    df, w = run_ask(book([hcv_dump], extra=[t11_nov(), t15_nov()]), a)
    tw = df[df["product"] == "TW"]
    sc = rows_of(at(tw, "GJ_ABS"), make_model_class="SCOOTER", fuel_type="Petrol", sub_product="1+1")
    st = rows_of(at(tw, "GJ_ABS"), make_model_class="SCOOTER", fuel_type="Petrol", sub_product="STP")
    oth = rows_of(at(tw, "GJ_Bad"), fuel_type="Petrol")
    check(f"November 1+1 / SATP with the '{label} CD2' answer: SC/EV 1+1 payin {want['c11']} up to CD1 60.01; SATP {want['s11']} On TP in ONE row with is_cpa NA (the grid no longer splits by CPA)",
          len(sc) == 1 and sc.pi_net_rate.iloc[0] == want["c11"] and sc.tariff_max.iloc[0] == 60.01 and len(st) == 1 and st.pi_tp_rate.iloc[0] == want["s11"] and st.is_cpa.iloc[0] == "NA" and st.pi_type.iloc[0] == "On TP",
          (sc[["pi_net_rate", "tariff_max"]].to_dict("records"), st[["pi_tp_rate", "is_cpa"]].to_dict("records")))
    check(f"... and a second entry ('{label}'): 1+1 {want['o11']}, SATP {want['os11']}", set(rows_of(oth, sub_product="1+1").pi_net_rate) == {want["o11"]} and set(rows_of(oth, sub_product="STP").pi_tp_rate) == {want["os11"]},
          (rows_of(oth, sub_product="1+1").pi_net_rate.unique(), rows_of(oth, sub_product="STP").pi_tp_rate.unique()))
    bj = rows_of(tw, sub_product="1+5", new_vehicle_makes="BAJAJ")
    hd = rows_of(tw, sub_product="1+5", new_vehicle_makes="HONDA")
    check(f"November 1+5 with the '{label} CD2' answer: Bajaj {want['b15']} On Net (Slab Net 5), Honda {want['h15']} On OD (Slab OD+Add on) - the formula type is read from ITS column, not from the payout figure next to it",
          set(bj.pi_net_rate) == {want["b15"]} and set(bj.pi_type) == {"On Net"} and set(hd.pi_od_rate) == {want["h15"]} and set(hd.pi_type) == {"On OD"} and len(bj) > 0 and len(hd) > 0,
          (bj[["pi_net_rate", "pi_type"]].drop_duplicates().to_dict("records"), hd[["pi_od_rate", "pi_type"]].drop_duplicates().to_dict("records")))
    check(f"the payout question is asked ONCE for both two-wheeler tables ('{label}')", sum("Type 1-2" in s for s in seen) == 1, [s for s in seen if "Type" in s])
    check(f"the answer is written into the decisions of the run ('{label}')", any(f"{label} CD2" in d for d in a.decisions), a.decisions)
    check("the renamed sheets are named in the warnings", any("New 2W Comp & SATP wef 1st Nov" in x and "last month's sheet" in x for x in w) and any("2W Grid 1+5" in x and "last month's sheet" in x for x in w), [x[:90] for x in w][:4])
    check("the SATP payout that is no longer split by CPA is said in the warnings", any("not split into 'With CPA'" in x for x in w), [x[:90] for x in w][:5])

# ---- a column moving one cell to the right, a spacer column, an inserted column: the same rows
def t15_layout(order, name="2W 1+5"):
    names = {"c": "Agency/PB Clusters", "m": "Make", "s": "Agency/PB Seg", "d": "CD1", "p": "Max CD2", "f": "Formula Type", "_": None}
    data = [dict(c="GJ_ABS", m="BAJAJ", s="3-7 KW", d=0.2, p=0.35, f="Slab Net 5"), dict(c="GJ_ABS", m="HONDA", s="MC <=155", d=0.4, p=0.65, f="Slab OD+Add on")]
    return sheet(name, [], [names[k] for k in order], [[d.get(k) for k in order] for d in data])


base15 = run(book([hcv_dump], extra=[t15_layout("cmsdpf")]))[0]
base15 = base15[base15["sub_product"] == "1+5"].reset_index(drop=True)
check("1+5 sheet in last month's layout: Bajaj 35 On Net, Honda 65 On OD", len(base15) > 0 and set(base15[base15.new_vehicle_makes == "BAJAJ"].pi_net_rate) == {35.0} and set(base15[base15.new_vehicle_makes == "HONDA"].pi_od_rate) == {65.0})
for order in ("_cmsdpf", "__cmsdpf", "cmsd_pf", "cm_sdpf", "cmsdp_f", "cmsdpf_", "_c_m_s_d_p_f", "fpdsmc", "cfpdsm"):
    got = run(book([hcv_dump], extra=[t15_layout(order)]))[0]
    got = got[got["sub_product"] == "1+5"].reset_index(drop=True)
    check(f"1+5 columns in the order '{order}' ('_' = a blank column): exactly the same rows as last month's layout", got.equals(base15), (len(got), len(base15)))

PC_TP_ROWS = [("Bangalore", "Petrol<1000", "<10", 0.155, 0.12, ""), ("Bangalore", "Petrol<1000", ">10", 0.185, 0.15, "Few RTOs for 10+ age are declined"), ("Mumbai2", "Diesel>1500", "All", 0.2, 0.17, "")]


def tp_layout(order, name=" PVT Car TP", head_spell=None):
    names = {"c": "Cluster", "s": (head_spell or {}).get("s", "Segment"), "a": "Age", "p": "Max CD2", "x": "Avg CD2", "n": "Note", "_": None}
    data = [dict(c=c, s=s, a=a, p=p, x=x, n=n or None) for c, s, a, p, x, n in PC_TP_ROWS]
    return sheet(name, [], [names[k] for k in order], [[d.get(k) for k in order] for d in data])


def tp_rows(order, ask=None, **kw):
    b = book([hcv_dump], extra=[tp_layout(order, **kw)])
    df, w = (run(b) if ask is None else run_ask(b, ask))
    return df[(df["product"] == "Private Car")].sort_values(["new_rto_list", "fuel_type", "vehicle_age_min"]).reset_index(drop=True), w


tp_base, _ = tp_rows("casp" + "n")
check("car TP in last month's layout: Cluster | Segment | Age | Max CD2 | Note - the note is a note", len(tp_base) > 0 and tp_base.remarks.str.contains("Few RTOs for 10+ age are declined", regex=False).any() and set(tp_base[tp_base.new_rto_list == "GO-DIGIT_PC_TP_OCT26_Bangalore"].pi_tp_rate) <= {15.5, 18.5, 0.0}, tp_base.new_rto_list.unique())
for order in ("_casp" + "n", "casp_n", "casp" + "n_", "nasp" + "c"):
    got, _ = tp_rows(order)
    check(f"car TP columns in the order '{order}': exactly the same rows", got.equals(tp_base), (len(got), len(tp_base)))
a, seen = asker("2")
tp_nov, w = tp_rows("caspxn", a)
tp_avg = tp_nov[(tp_nov.new_rto_list == "GO-DIGIT_PC_TP_OCT26_Bangalore") & (tp_nov.fuel_type == "Petrol") & (tp_nov.cc_max == 1000.01)].sort_values("vehicle_age_min")
check("car TP with 'Avg CD2' in front of the note (the November shape) and the Avg answer: payin from Avg CD2 (12 / 15), the note stays the note, and no row carries a number as its note",
      list(tp_avg.pi_tp_rate) == [12.0, 15.0] and tp_nov.remarks.str.contains("Few RTOs for 10+ age are declined", regex=False).any() and not tp_nov.remarks.str.contains(r"grid note: \d+(?:\.\d+)?(?:\s|$|\|)", regex=True).any(),
      (list(tp_avg.pi_tp_rate), tp_nov.remarks.unique()[:3]))
tp_typo, w = tp_rows("casp" + "n", head_spell=dict(s="Segement"))
check("a typo in a header ('Segement') is read as 'Segment' and the run says so; the rows are the same", tp_typo.equals(tp_base) and any("Segement" in x and "is read as" in x for x in w), [x[:100] for x in w][:4])

# ---- car comprehensive / SAOD: a second block on the left with its own cluster column, rows in another order, and one volume slab
def pcc_nov(slabs=("> 5 Lacs",), shuffle=True, name="New 4W Comp_SAOD "):
    clusters = ["Andaman", "Bad_UP", "Bihar"]
    rate = {"Andaman": (22, 22, 25, 25), "Bad_UP": (19.5, 24, 22, 27), "Bihar": (19.5, 24, 22, 28)}
    nb_order = clusters
    slab_order = clusters[1:] + clusters[:1] if shuffle else clusters                       # the slab block lists the clusters in another order than the new-business block
    head = ["Cluster", "HYUNDAI_NB", "OTHERS_NB", None, "Cluster", "Volume Slab", "Non NCB", "NCB", "Non NCB", "NCB"]
    above = [[], [None, "NB"], [None, "1+3; 90:10", None, None, None, None, "SAOD", None, "Comp"]]
    rows = []
    for i, nb in enumerate(nb_order):
        for sl in slabs:
            rows.append([nb, 0.26, 0.25, None, slab_order[i], sl, *rate[slab_order[i]]])
    return sheet(name, above, head, rows), rate


pcc, rate = pcc_nov()
a, seen = asker("1")
df, w = run_ask(book([hcv_dump], extra=[pcc]), a)
pcdf = df[df["product"] == "Private Car"]
ok = True
for cl, (sn, sy, cn, cy) in rate.items():
    got = {(r.sub_product, r.is_ncb): r.pi_od_rate for r in rows_of(at(pcdf, cl), fuel_type="Petrol").itertuples()}
    ok &= got == {("SAOD", "NO"): sn, ("SAOD", "YES"): sy, ("1+1", "NO"): cn, ("1+1", "YES"): cy}
check("car comp / SAOD with a new-business block on the left whose cluster rows run in another order: every cluster gets the rates of ITS OWN row (the cluster column of the 'Volume Slab' block, not the first 'Cluster' on the row)", ok and len(pcdf) > 0, pcdf.new_rto_list.unique())
check("... the missing volume slab '1-3 Lac' was ASKED about (not guessed): the sheet's only slab is offered, the choice is a decision of the run and a warning",
      any("1-3 Lac" in s and "> 5 Lacs" in s for s in seen) and any("volume slab" in d and "> 5 Lacs" in d for d in a.decisions) and any("1-3 Lac" in x and "your choice" in x for x in w), (a.decisions, [s[:80] for s in seen if "slab" in s][:3]))
try:
    run(book([hcv_dump], extra=[pcc_nov()[0]]))
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("... without a keyboard it stops and names the slab it wanted and the slabs the sheet has", "1-3 Lac" in msg and "> 5 Lacs" in msg, msg)
a, seen = asker("2")
df, w = run_ask(book([hcv_dump], extra=[pcc_nov()[0]]), a)
check("... answering 'none' leaves the private-car 1+1 / SAOD rows out and says so", not (df["product"] == "Private Car").any() and any("leave the private-car 1+1 / SAOD rows out" in x for x in w), [x[:90] for x in w][:3])
many = pcc_nov(slabs=("0-25 K", "1-3 Lac", "> 5 Lacs"), shuffle=False)[0]
df, w = run(book([hcv_dump], extra=[many]))
check("with the slab of the deal in the sheet nothing is asked and only its rows are written (last month's behaviour)", set(df[df["product"] == "Private Car"].new_rto_list.str.replace(r"^.*_OCT26_", "", regex=True).str.upper()) == {"ANDAMAN", "BAD_UP", "BIHAR"} and not any("volume slab" in x for x in w))

# ---- 2W SAOD with new columns in the middle
SAOD_HEAD = ["Cluster", "Segment", "Make", "Min CD1", "Max CD1 for NoBreakIn", "Max CD1 for BreakIn", "CD2", "Status No Break in", "Status with Break in", "Max 1 year SAOD CD2", "Max 2 year SAOD CD2"]
saod = sheet("2W SAOD", [[]], SAOD_HEAD, [["GJ_ABS", "MC <155", "Other than Honda and Bajaj", 0, 0.8, 0.55, 0.225, "STP", "STP", 0.4, 0.35]])
df, w = run(book([hcv_dump], extra=[saod]))
sd = rows_of(df, sub_product="SAOD", fuel_type="Petrol").sort_values("tariff_max")
check("2W SAOD with three new columns: found by its headers; the 1-year payout 40 On OD with the No BreakIn (80.01) and BreakIn (55.01) caps",
      list(zip(sd.tariff_max, sd.pi_od_rate)) == [(55.01, 40.0), (80.01, 40.0)] and set(sd.add_tnc) == {"BreakIn", "No BreakIn"}, sd[["tariff_max", "pi_od_rate", "add_tnc"]])
unread = [x for x in w if "are not read by this script" in x]
check("... the new columns nobody reads are listed in the warnings; the columns left out on purpose (Min CD1, 2nd year) are not",
      len(unread) == 1 and all(t in unread[0] for t in ("'CD2'", "'Status No Break in'", "'Status with Break in'")) and "Min CD1" not in unread[0] and "2 year" not in unread[0], unread)

# ---- a table that is not there stops with what was expected and what was found
mod.MISSING_TABLES = "ask"
try:
    run(book([hcv_dump]))
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("a missing table stops: it says what header row was expected and which sheets the workbook has", "expected a header row that says" in msg and "Visible sheets" in msg and "LATEST CV COMBINED GRID" in msg, msg)
check("... and the closest sheet is not a sheet that merely has a 'CD2' somewhere", "closest sheet is 'LATEST CV COMBINED GRID'" not in msg, msg)
near = sheet("Latest 2W 1+1 & SATP", [[None, None, "1+1"]], ["Agency/PB Clusters", "Agency/PB Seg", "CD2"], [["GJ_ABS", "SC/EV", 0.5]])
try:
    run(book([hcv_dump], extra=[near]))
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("a near miss is named: the closest sheet, the header it lacks", "closest sheet is 'Latest 2W 1+1 & SATP'" in msg and "CD1" in msg.split("it lacks")[1][:40], msg)
a, seen = asker("2", "2", "2", "2", "2")
df, w = run_ask(book([hcv_dump]), a)
check("a missing table can be left out when the person says so: a NOT FOUND warning for each, the rest of the file is written",
      sum("NOT FOUND" in x for x in w) == 5 and set(df["product"]) <= {"GCV 4W"}, [x[:70] for x in w][:6])
a, seen = asker("1")
try:
    run_ask(book([hcv_dump]), a)
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("... and 'stop' really stops", "was not found" in msg and "stopped by you" in msg, msg)
mod.MISSING_TABLES = "skip"

# ---- new blank rows above a table (title rows, a banner) and blank rows inside it change nothing
def pad(sh, n, inside=0):
    rows = [list(r) for r in sh.rows]
    if inside:
        rows = rows[:-1] + [[] for _ in range(inside)] + rows[-1:]
    return mod.Sheet(sh.name, [[] for _ in range(n)] + rows)


def built(extra, answers=("2", "1")):
    a, _ = asker(*answers)
    df, w = run_ask(book([hcv_dump], extra=extra), a)
    return df[df["product"].isin(["TW", "Private Car"])].reset_index(drop=True)


fixtures = lambda: [t11_nov(), t15_nov(), tp_layout("caspxn"), saod, pcc_nov()[0]]
plain = built(fixtures())
check("sanity: the shifted-layout fixtures give two-wheeler and private-car rows of every kind", {("TW", "1+1"), ("TW", "STP"), ("TW", "1+5"), ("TW", "SAOD"), ("Private Car", "1+1"), ("Private Car", "SAOD"), ("Private Car", "STP")} <=
      set(zip(plain["product"], plain.sub_product)), sorted(set(zip(plain["product"], plain.sub_product))))
for n, inside in ((7, 0), (12, 0), (3, 2)):
    shifted = built([pad(s_, n, inside) for s_ in fixtures()])
    check(f"{n} blank rows above every table{' and ' + str(inside) + ' inside it' if inside else ''}: exactly the same rows", shifted.equals(plain), (len(shifted), len(plain)))
# the commercial-vehicle grid too (its header sits on row 4 in the sheets of both months)
cv_plain = run(book([hcv_dump]))[0]
cv_book = mod.Book([pad(mod.Sheet("LATEST CV COMBINED GRID", [[], [], [], HEAD, list(hcv_dump)]), 9)] + [s_ for s_ in book([hcv_dump]).sheets if s_.name != "LATEST CV COMBINED GRID"])
check("9 blank rows above the commercial-vehicle header: exactly the same rows", run(cv_book)[0].equals(cv_plain) and len(cv_plain) > 0)

# ---- a layout it does not understand stops with what was expected and what was found
rows11 = [["GJ_ABS", "SC/EV", 0.6, 0.5, 0.55, 0.6]]
unpaired = sheet("Latest 2W 1+1 & SATP", [[None, None, "1+1", None, "SATP"], [None, None, None, None, "With CPA"]], ["Agency/PB Clusters", "Agency/PB Seg", "CD1", "CD2", "CD2"], [r[:5] for r in rows11])
untitled = sheet("Latest 2W 1+1 & SATP", [], ["Agency/PB Clusters", "Agency/PB Seg", "CD1", "CD2", "CD2", "CD2"], rows11)
for what, sh_, want in (("a SATP payout with only a 'With CPA' column", unpaired, "either one column or one 'With CPA' and one 'Without CPA'"), ("payout columns with no '1+1' / 'SATP' title above them", untitled, "expected the titles '1+1' and 'SATP'")):
    try:
        run(book([hcv_dump], extra=[sh_]))
        msg = ""
    except mod.GridError as e:
        msg = str(e)
    check(f"two-wheeler 1+1 sheet with {what}: stops, saying what was expected and what was found (never read in column order)", want in msg and "found" in msg, msg)

# ---- the commercial-vehicle grid: titles over the age groups, the payout heading, a new column
satp_first = [["", "", "", "", "", "", "", "THIRD-PARTY / SATP"] + [""] * 7 + ["COMPREHENSIVE"], ["", "", "", "", "", "", "", "Max CD2"] + [""] * 7 + ["Max CD2"]]
row = ("CV GCV3(EV/Non-EV)", "GCV3", "—", "BAD GJ", "Gujarat", "All make", "0–60%") + (0.3,) * 8 + (0.1,) * 8
sh = mod.Sheet("LATEST CV COMBINED GRID", [[]] + satp_first + [HEAD + ["Brand new column"], list(row)])
df, w = run(mod.Book([sh]))
check("CV grid with the third-party group FIRST and titles over the groups: read by the titles (comprehensive 10, third-party 30)",
      set(rows_of(df, sub_product="1+1").pi_net_rate) == {10.0} and set(rows_of(df, sub_product="STP").pi_tp_rate) == {30.0} and any("third-party group first" in x for x in w), [x[:80] for x in w][:3])
check("... the 'Max CD2' heading (earlier grids said 'CD2') and a new column nobody reads are both said in the warnings",
      any("headed 'Max CD2'" in x for x in w) and any("'Brand new column'" in x and "are not read" in x for x in w), [x[:100] for x in w][:5])
same = [["", "", "", "", "", "", "", "COMPREHENSIVE"] + [""] * 7 + ["COMPREHENSIVE"], []]
try:
    mod.read_cv_grid(mod.Book([mod.Sheet("LATEST CV COMBINED GRID", [[]] + same + [HEAD, list(row)])]), mod.Asker(interactive=False), quiet)
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("two groups with the same title stop with what was expected", "one comprehensive and one third-party" in msg, msg)

# ---- the CV remarks: found by their headings, and the age-band table is checked instead of flagged
AGE_TABLE = [["System Age", "Age Band"], ["New", "Up to 11 months"], ["1 Yr", "12 months onwards to 1 year 11 months"], ["2 Yr", "1 year 12 months onwards to 2 years 11 months"], ["3 Yr", "2 years 12 months onwards to 3 years 11 months"],
             ["4 Yr", "3 years 12 months onwards to 4 years 11 months"], ["5 Yr", "4 years 12 months onwards to 5 years 11 months"], ["6–10 Yr", "5 years 12 months onwards to 10 years 11 months"], ["11+ Yr", "10 years 12 months onwards "]]
rem = lambda table: mod.Sheet("Remarks for CV", [["REMARKS FOR HCV"], [HCV_NOTE], [], ["PfB for age bracket"]] + table)
sec = mod.read_cv_remarks(mod.Book([rem(AGE_TABLE)]), mod.Asker(interactive=False))
check("the remarks sheet is found by its 'REMARKS FOR' headings (renamed), the age-band lines are not notes", list(sec) == ["hcv", "_age_table"] and len(sec["hcv"]) == 1 and len(sec["_age_table"]) == 8, {k: len(v) for k, v in sec.items()})
w2 = []
mod.check_age_table(sec["_age_table"], w2, quiet)
check("the sheet's own age-band table agrees with the bands of this script (whole years): no warning", w2 == [], w2)
changed = [list(r) for r in AGE_TABLE]
changed[7][1] = "5 years 12 months onwards to 9 years 11 months"
sec2 = mod.read_cv_remarks(mod.Book([rem(changed)]), mod.Asker(interactive=False))
mod.check_age_table(sec2["_age_table"], w2, quiet)
check("a band that changed in the sheet is a warning that names it", len(w2) == 1 and "6–10 Yr" in w2[0] and "need a look" in w2[0], w2)
df, w = run(mod.Book([mod.Sheet("LATEST CV COMBINED GRID", [[], [], [], HEAD, list(hcv_dump)]), rem(AGE_TABLE)]))
check("... and none of the age-band lines is flagged on the heavy-vehicle rows", len(df) > 0 and not df.remarks.str.contains("months onwards", regex=False).any() and not df.remarks.str.contains("Up to 11", regex=False).any(), df.remarks.unique()[:2])

# ---- the RTO mapping sheets: by what they say
mp = lambda name, head, rows: mod.Sheet(name, [list(head)] + [list(r) for r in rows])
cvmap = mp("RTO Mapping", ["RTO Code", "City", "CV HCV", "CV LCV"], [["GJ01", "Ahmedabad", "BAD GJ", "x"], ["GJ02", "Surat", "BAD GJ", "x"]])
fourw = mod.Sheet("4W RTO Mapping - New", [[], [None, "RTO Code", "4WTP", "4WOD"], [None, "AN01", "Andaman", "Goa_Andaman"], [None, "AN02", "Andaman", "Goa_Andaman"]])
names_col = mp("Some table", ["RTO", "Cluster"], [["WL BAD GJ", "BAD GJ"], ["WL BAD GJ", "BAD GJ"]])
m = mod.read_all_mappings(mod.Book([cvmap, fourw, names_col]), mod.Asker(interactive=False))
check("RTO mapping sheets are found by their headers (renamed, the code column not first): codes per (column, cluster); a column headed RTO that holds the person's names is not a mapping",
      m.get(("cvhcv", "badgj")) == ["GJ01", "GJ02"] and m.get(("4wtp", "andaman")) == ["AN01", "AN02"] and not any("wl" in k[1] for k in m), sorted(m)[:6])

# ---- the real November 2026 workbook (Go Digit's 'communication grid': 13 sheets, all visible, every one renamed or redesigned)
nov = os.path.join(helpers.INPUTS, "godigit_nov26.xlsx")
if os.path.exists(nov):
    import openpyxl as _ox2
    nb = mod.read_workbook(nov)
    ask0 = mod.Asker(interactive=False)
    warn = []
    wanted = [(mod.SPEC_CV, "CV Grid", 4), (mod.SPEC_TW11, "New 2W Comp & SATP wef 1st Nov", 3), (mod.SPEC_TW15, "2W Grid 1+5", 2), (mod.SPEC_SAOD, "2W SAOD", 2), (mod.SPEC_PCC, "New 4W Comp_SAOD ", 4), (mod.SPEC_PCT, "4W TP Consolidated", 1)]
    got = [(h.sheet.name, h.hr) for h in (mod.find_table(nb, s, ask0, warn, quiet) for s, _, _ in wanted)]
    check("November workbook: every table is found by its headers, on the right sheet and header row", got == [(n, r) for _, n, r in wanted], got)
    msg = ""
    try:
        mod.build(nov, dt.date(2026, 11, 1), T, quiet, mod.Asker(interactive=False))
    except mod.GridError as e:
        msg = str(e)
    check("November workbook without a keyboard: stops at the payout question (it names Max CD2 and Avg CD2), not at a sheet that 'fits twice'", "Max CD2" in msg and "Avg CD2" in msg and "more than one sheet" not in msg, msg)
    out = {}
    for pick, measure in (("2", "avg"), ("1", "max")):
        a, seen = asker(pick, "1")
        mod.reset_run_lists()
        out[measure], wn = mod.build(nov, dt.date(2026, 11, 1), T, quiet, a)
        out[measure + "_w"] = wn
    d = out["avg"]
    byp = d.groupby(["product", "sub_product"]).size().to_dict()
    check("November workbook: the car comp / SAOD rows are 28 clusters x 2 NCB x 4 fuels = 224 each; the 1+1 and SATP two-wheeler rows are one per entry and fuel (704 each: 351 entries + the 'MC <= 180 Hero/Honda/TVS' row of WB_Bad, read since 10 Nov 2026)",
          byp.get(("Private Car", "1+1")) == 224 and byp.get(("Private Car", "SAOD")) == 224 and byp.get(("TW", "1+1")) == 704 and byp.get(("TW", "STP")) == 704, byp)
    check("November workbook: no key is written twice, no row carries a number as its 'grid note', every file column is filled", not d.duplicated(mod.KEY_COLUMNS).any() and not d.remarks.str.contains(r"grid note: \d+(?:\.\d+)?(?:\s|$|\|)", regex=True).any())
    check("November workbook: a different answer (Max CD2) changes only the payouts of the two-wheeler and car third-party tables, not their number or keys",
          len(out["max"]) == len(d) and set(out["max"][out["max"]["product"] == "GCV 4W"].pi_net_rate) == set(d[d["product"] == "GCV 4W"].pi_net_rate)
          and (out["max"].pi_tp_rate.sum() != d.pi_tp_rate.sum()) and (out["max"].pi_net_rate.sum() != d.pi_net_rate.sum()))
    check("November workbook: the warnings name what changed (renamed sheets, the heading 'Max CD2', the SATP CPA split, the unread SAOD columns, the slab)",
          all(any(k in x for x in out["avg_w"]) for k in ("last month's sheet", "headed 'Max CD2'", "not split into 'With CPA'", "'Status No Break in'", "volume slab '1-3 Lac'")), [x[:60] for x in out["avg_w"]][:12])
    # independent spot checks straight from the cells (openpyxl), for the answer 'avg'
    wbx = _ox2.load_workbook(nov, data_only=True)
    ws = wbx["New 2W Comp & SATP wef 1st Nov"]
    r = next(r for r in range(4, ws.max_row + 1) if ws.cell(r, 2).value == "Andaman" and ws.cell(r, 3).value == "MC <= 180 Hero/Honda")
    x = at(d[d["product"] == "TW"], "Andaman")
    x = x[(x.new_vehicle_makes == "HERO,HONDA") & (x.fuel_type == "Petrol")]
    check(f"November spot check, 1+1 sheet row {r}: Andaman MC<=180 Hero/Honda - Avg CD2 {ws.cell(r, 6).value} -> {round(ws.cell(r, 6).value * 100, 4)} (1+1), SATP Avg {ws.cell(r, 8).value} -> {round(ws.cell(r, 8).value * 100, 4)}",
          abs(x[x.sub_product == "1+1"].pi_net_rate.iloc[0] - round(ws.cell(r, 6).value * 100, 4)) < 1e-6 and abs(x[x.sub_product == "STP"].pi_tp_rate.iloc[0] - round(ws.cell(r, 8).value * 100, 4)) < 1e-6)
    ws = wbx["2W Grid 1+5"]
    r = next(r for r in range(3, ws.max_row + 1) if ws.cell(r, 8).value == "Slab OD+Add on" and isinstance(ws.cell(r, 6).value, (int, float)))
    x = at(d[d["sub_product"] == "1+5"], ws.cell(r, 2).value)
    x = x[(x.new_vehicle_makes == "ROYAL") & (x.cc_min == 0) & (x.cc_max == 350.01)]
    check(f"November spot check, 1+5 sheet row {r}: {ws.cell(r, 2).value} Royal Enfield <=350 'Slab OD+Add on' - Avg {ws.cell(r, 6).value} written On OD (a position reader wrote it On Net)", len(x) > 0 and (x.pi_type == "On OD").all() and (abs(x.pi_od_rate - round(ws.cell(r, 6).value * 100, 4)) < 1e-6).all())
    ws = wbx["New 4W Comp_SAOD "]
    r = next(r for r in range(5, 33) if ws.cell(r, 11).value == "GJ_Good")
    x = at(d[(d["product"] == "Private Car") & d.sub_product.isin(["1+1", "SAOD"])], "GJ_Good")
    x = x[x.fuel_type == "Petrol"]
    check(f"November spot check, car sheet row {r}: GJ_Good takes the rates of ITS row of the slab block ({[ws.cell(r, c).value for c in (13, 14, 15, 16)]}), not those of row 13 where the new-business block lists it",
          {(q.sub_product, q.is_ncb): q.pi_od_rate for q in x.itertuples()} == {("SAOD", "NO"): ws.cell(r, 13).value, ("SAOD", "YES"): ws.cell(r, 14).value, ("1+1", "NO"): ws.cell(r, 15).value, ("1+1", "YES"): ws.cell(r, 16).value})
    ws = wbx["4W TP Consolidated"]
    r = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value == "Andaman" and ws.cell(r, 2).value == "Diesel<1500" and ws.cell(r, 3).value == "<10")
    x = at(d[(d["product"] == "Private Car") & (d.sub_product == "STP")], "Andaman")
    x = x[(x.fuel_type == "Diesel") & (x.cc_max == 1500.01) & (x.vehicle_age_max == 10.01)]
    check(f"November spot check, car TP row {r}: Andaman Diesel<1500 <10 - Avg CD2 {ws.cell(r, 5).value} -> {round(ws.cell(r, 5).value * 100, 4)}, and no number as its grid note", len(x) == 1 and abs(x.pi_tp_rate.iloc[0] - round(ws.cell(r, 5).value * 100, 4)) < 1e-6 and "grid note" not in x.remarks.iloc[0])
else:
    print("note: regression/inputs/godigit_nov26.xlsx is missing - the November workbook checks were skipped")

# ---- the user's standing answer (10 Nov 2026): where a table has both, the payin is 'Avg CD2' - applied without asking, said once
mod.PAYOUT_MEASURE = _rule
check("the shipped setting is the user's answer of 10 Nov 2026: 'Avg CD2'", mod.PAYOUT_MEASURE == "avg", mod.PAYOUT_MEASURE)
for setting, label, want in (("avg", "Avg", dict(c11=59.0, s11=63.5, b15=32.5, h15=60.0)), ("max", "Max", dict(c11=65.0, s11=70.0, b15=35.0, h15=65.0))):
    mod.PAYOUT_MEASURE = setting
    a, seen = asker()                                              # no answers at all: a question would end the run
    df, w = run_ask(book([hcv_dump], extra=[t11_nov(), t15_nov()]), a)
    tw = df[df["product"] == "TW"]
    sc = rows_of(at(tw, "GJ_ABS"), make_model_class="SCOOTER", fuel_type="Petrol")
    bj, hd = rows_of(tw, sub_product="1+5", new_vehicle_makes="BAJAJ"), rows_of(tw, sub_product="1+5", new_vehicle_makes="HONDA")
    check(f"PAYOUT_MEASURE = '{setting}': both tables are read from the '{label} CD2' column without a question (1+1 {want['c11']}, SATP {want['s11']}, 1+5 Bajaj {want['b15']} / Honda {want['h15']})",
          not any("Type 1-" in s for s in seen) and set(rows_of(sc, sub_product="1+1").pi_net_rate) == {want["c11"]} and set(rows_of(sc, sub_product="STP").pi_tp_rate) == {want["s11"]}
          and set(bj.pi_net_rate) == {want["b15"]} and set(hd.pi_od_rate) == {want["h15"]}, [s for s in seen if "Type" in s])
    check(f"... and the rule is said ONCE in the decisions of the run ('{label}')", sum("payout column" in d and f"'{label} CD2'" in d for d in a.decisions) == 1, a.decisions)
mod.PAYOUT_MEASURE = _rule
odd11 = sheet("New 2W Comp & SATP", [[None, None, None, "1+1", None, None, "SATP"]], [None, "Agency/PB Clusters", "Agency/PB Seg", "CD1", "Max CD2", "CD2", "CD2"], [[None, "GJ_ABS", "SC/EV", 0.6, 0.65, 0.59, 0.7]])
try:
    run(book([hcv_dump], extra=[odd11]))
    msg = ""
except mod.GridError as e:
    msg = str(e)
check("a table whose two payout columns are not the pair the rule covers ('Max CD2' + 'CD2') is a question, not a guess", "Max CD2" in msg and "payin" in msg, msg)
if os.path.exists(nov):
    msg = ""
    try:
        mod.build(nov, dt.date(2026, 11, 1), T, quiet, mod.Asker(interactive=False))
    except mod.GridError as e:
        msg = str(e)
    check("November workbook with the shipped rule, without a keyboard: the payout question is gone; the run stops at the private-car slab (it names '1-3 Lac' and the slab the sheet has)", "1-3 Lac" in msg and "> 5 Lacs" in msg and "payin" not in msg, msg)
    a, seen = asker("1")
    mod.reset_run_lists()
    d2, w2 = mod.build(nov, dt.date(2026, 11, 1), T, quiet, a)
    check("... and with the slab answered it writes exactly the 'Avg CD2' rows of the earlier run", d2.equals(out["avg"]) and len(d2) > 0, (len(d2), len(out.get("avg", []))))

# ================================================================ reading the workbook
tmp = tempfile.mkdtemp(prefix="godigit_")
import openpyxl  # noqa: E402
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "LATEST CV COMBINED GRID"
for r in [[], HEAD, ["CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", "BAD GJ", "Gujarat", "All", "60%"] + [0.25] * 8 + [0.3] * 8]:
    ws.append(r)
hid = wb.create_sheet("Old")
hid.sheet_state = "hidden"
hid.append(["x"])
plain = os.path.join(tmp, "plain.xlsx")
wb.save(plain)
b = mod.read_workbook(plain)
check("an .xlsx is read as plain values; a hidden sheet is marked not visible", b.sheetnames == ["LATEST CV COMBINED GRID", "Old"] and [s.visible for s in b.sheets] == [True, False] and b["LATEST CV COMBINED GRID"].cell(3, 7) == "60%")
try:
    mod.read_workbook(__file__)
    ok = False
except mod.GridError:
    ok = True
check("something that is not a workbook stops with a plain message", ok)
try:
    import msoffcrypto
    from msoffcrypto.format.ooxml import OOXMLFile
    have_lib = True
except ImportError:
    have_lib = False
if have_lib:
    enc = os.path.join(tmp, "protected.xls")                                   # the insurer's file: an encrypted .xlsx called .xls
    with open(plain, "rb") as fh, open(enc, "wb") as out:
        OOXMLFile(fh).encrypt("Secret@1", out)
    b2 = mod.read_workbook(enc, "Secret@1")
    check("a password-protected workbook (an encrypted .xlsx called .xls) is read with the password", b2.sheetnames == b.sheetnames and b2["LATEST CV COMBINED GRID"].cell(3, 7) == "60%")
    for pw, what in ((None, "no password"), ("wrong", "a wrong password")):
        try:
            mod.read_workbook(enc, pw)
            ok = False
        except mod.PasswordNeeded:
            ok = True
        check(f"the protected workbook with {what} says so (PasswordNeeded)", ok)
    opener = mod.make_opener(mod.Asker(interactive=False), None)
    try:
        opener(enc)
        ok = False
    except mod.GridError as e:
        ok = "--password" in str(e) and "GRID_PASSWORD" in str(e)
    check("without a keyboard the missing password is a plain message that names --password", ok)
    check("the opener takes the password given", mod.make_opener(mod.Asker(interactive=False), "Secret@1")(enc).sheetnames == b.sheetnames)
else:
    print("note: msoffcrypto-tool is not installed - the password-protected workbook checks were skipped")

check("a hidden sheet is never taken for the grid (same name words, but hidden)", mod.find_sheet(mod.Book([mod.Sheet("Latest CV Combined", [["x"]], visible=False), mod.Sheet("LATEST CV COMBINED GRID", [["y"]])]),
      mod.CV_SHEET_WORDS, (), "x", mod.Asker(interactive=False)).name == "LATEST CV COMBINED GRID")
try:
    mod.read_cv_grid(mod.Book([mod.Sheet("LATEST CV COMBINED GRID", [HEAD[:7] + ["New", "1 Yr"], ["CV HCV"]])]), mod.Asker(interactive=False), quiet)
    ok = False
except mod.GridError as e:
    ok = "age columns" in str(e)
check("a grid whose age columns changed stops with a question about them", ok)

# ================================================================ the CSV and the review workbook of a run
mod.reset_run_lists()
out = os.path.join(tmp, "o", "GODIGIT_UPLOAD_2026-10.csv")
res = mod.main(plain, F, T, out, None, quiet, mod.Asker(interactive=False))
chk = helpers.check_csv(out, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
check("main() writes a clean CSV (42 columns, numbers, dates DD/MM/YYYY, remarks only as Manual Review flags)", res == out and not chk, chk)
o = pd.read_csv(out, dtype=str, keep_default_na=False)
check("constant columns: id 1, the company, user 2966, dates 15/10/2026 - 31/03/2030", set(o.id) == {"1"} and set(o.insurance_company) == {"GO DIGIT GENERAL INSURANCE LTD"} and set(o.user_id) == {"2966"}
      and set(o.from_date) == {"15/10/2026"} and set(o.to_date) == {"31/03/2030"} and set(o.insurer_vertical) == {"Not Applicable"})
check("the default output goes to output/godigit/", mod.OUTPUT_FOLDER == "godigit" and mod.default_output("GODIGIT_UPLOAD_2026-10.csv").replace("\\", "/").endswith("/output/godigit/GODIGIT_UPLOAD_2026-10.csv"))

# ================================================================ the real October 2026 workbook against the user's October upload
grid = os.path.join(helpers.INPUTS, "godigit_oct26.xlsx")
ref = os.path.join(helpers.REFS, "godigit_oct26_rates_export.csv")
if os.path.exists(grid):
    mod.reset_run_lists()
    msgs = []
    df, w = mod.build(grid, F, T, lambda *a, **k: msgs.append(" ".join(map(str, a))), mod.Asker(interactive=False))
    by = df.groupby("product").size().to_dict()
    check("October workbook: 53,212 rows; Private Car 579, TW 2,044, GCV 4W 38,125 (counts per product; the October car third-party note '1000-1500CC and <10 age is declined' splits one row into two since 10 Nov 2026 - the declined slice gets no card, exactly as in the user's upload: +2)", len(df) == 53212 and by.get("Private Car") == 579 and by.get("TW") == 2044 and by.get("GCV 4W") == 38125, by)
    missing = sorted(set(df.new_rto_list) - set(M.rto))
    check("every RTO group written is in the master except the clusters the master lacks (18, listed with their names)", len(missing) == 18 and all(n.startswith("GO-DIGIT_OCT26_") for n in missing), missing[:5])
    bad_makes = sorted({t for v in df.new_vehicle_makes.unique() for t in v.split(",")} - set(M.makes))
    check("every make group written is in the master", not bad_makes, bad_makes)
    check("no card is written twice (the key is unique)", not df.duplicated(mod.KEY_COLUMNS).any())
    check("hidden sheets are not read: the run says 14 are hidden", any("14 hidden" in m for m in msgs), msgs[:3])
    if os.path.exists(ref):
        new, removed, changed = mod.diff_against_reference(df, ref, quiet)
        check("against the user's October upload: all but a handful of its rows are reproduced with the same rates (removed <= 700, changed <= 2)", removed <= 700 and changed <= 2, (new, removed, changed))
        check("... and the rows that are new are the ones listed (SC/EV petrol, RJ 1+5, other fuels of the car cards, clusters the user did not upload, zero cards) - under 6,500", new <= 6500, new)
else:
    print("note: regression/inputs/godigit_oct26.xlsx is missing - the real-workbook checks were skipped")

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
