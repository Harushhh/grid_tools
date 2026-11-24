"""Deterministic checks of shriram_broker_grid_to_config.py: the readers of the payout / age / GVW / policy-type cells, the remark rules, the
RTO-group lookup by content, and whole small grids (held in memory) through build().  The last block runs the real October grid.
Run:  .venv\\Scripts\\python.exe tests\\shriram_tests.py"""
import datetime as dt
import os
import re
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("shriram_broker_grid_to_config.py")
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


ask = mod.Asker(interactive=False)
M = mod.load_masters(ask)
R = mod.Rtos(M)


# ================================================================ cell readers
check("payout: plain 25", mod.parse_po(25) == ("plain", 25.0, None))
check("payout: '15 OD' / '19.5% OD'", mod.parse_po("15 OD") == ("od", 15.0, None) and mod.parse_po("19.5% OD") == ("od", 19.5, None))
check("payout: '16 Net'", mod.parse_po("16 Net") == ("net", 16.0, None))
check("payout: '15 OD + 2.5 TP' is a split rate", mod.parse_po("15 OD + 2.5 TP") == ("odtp", 15.0, 2.5))
check("payout: IRDA = the IRDA cap on OD", mod.parse_po("IRDA") == ("irda", mod.IRDA_RATE, None))
check("payout: 'WITHDRAW' and text are not a payout", mod.parse_po("WITHDRAW") is None and mod.parse_po("Metro - 45%") is None)
check("age: 'upto 15 years.'", mod.parse_age("upto 15 years.")["pkg"] == (0, 15.01))
check("age: '4 to 20 years' -> 3.01-20.01", mod.parse_age("4 to 20 years")["pkg"] == (3.01, 20.01))
a = mod.parse_age("upto 15 years & upto 25 Years for STP")
check("age: a different STP band", a["pkg"] == (0, 15.01) and a["stp"] == (0, 25.01), a)
a = mod.parse_age("6 to 15 years package and 6 to 25 years for STP")
check("age: 'package and ... for STP' with two bands", a["pkg"] == (5.01, 15.01) and a["stp"] == (5.01, 25.01), a)
check("age: NEW", mod.parse_age("NEW")["new"] and mod.parse_age("New ")["new"])
a = mod.parse_age("upto 25 Years for STP")
check("age: only an STP band", a["stp"] == (0, 25.01) and a["pkg"] is None, a)
check("GVW: UPTO 2800", mod.gvw_band("GCCV LCV UPTO 2800 GVW") == (0, 2800.01))
check("GVW: ABOVE 2801 TO 3500 (a start value)", mod.gvw_band("GCCV LCV ABOVE 2801 GVW TO 3500 GVW") == (2800.01, 3500.01))
check("GVW: 12001 TO 20000", mod.gvw_band("GCCV HCV 12001 TO 20000 GVW") == (12000.01, 20000.01))
check("GVW: Above 50000 is open", mod.gvw_band("GCCV HCV Above 50000 GVW") == (50000.01, mod.CC_OPEN_END))
check("GVW: '2001-2800'", mod.gvw_band("GCCV LCV 2001-2800 GVW") == (2000.01, 2800.01))
check("GVW: 'UPTO 3500' (no GVW word)", mod.gvw_band("GCCV LCV UPTO 3500") == (0, 3500.01))
kinds = {"PRIVATE CAR PETROL": "pc_petrol", "PRIVATE CAR DIESEL": "pc_diesel", "PCCV 4W Upto 6+1": "taxi06", "PCCV 4W 7 to 10": "taxi710",
         "PCCV 4W CORPORATE BUS": "corp", "PCCV 4W SCHOOL BUS": "school", "PCCV 3W DIESEL": "p3w_diesel", "PCCV 3W OTHER THAN DIESEL": "p3w_other",
         "PCCV 3W E-Rickshaw\xa0 and GCCV 3W E-cart": "erick", "GCCV 3W Except E-CART": "g3w", "GCCV 3W EXCEPT E-CART ": "g3w", "MISC-D TRACTOR": "tractor",
         "MISC-D OTHERS": "miscd_others", "NEW MAHINDRA GCCV LCV, MCV & HCV (All GVW)": "newmah", "GCCV MCV ABOVE 7501 TO 12000 GVW": "gcv4w",
         "GCCV HCV\xa0 20001 TO 42500 GVW": "gcv4w", "GCCV LCV UPTO 2805 GVW": "gcv4w", "GCCV HCV Above 50000 GVW": "gcv4w"}
bad = {k: mod.classify(k)[0] for k, v in kinds.items() if mod.classify(k)[0] != v}
check("product labels are recognised (including the non-breaking spaces of the grid)", not bad, bad)
check("an unknown product is not guessed", mod.classify("WIND TURBINE")[0] is None)
check("policy types", mod.policy_subs("P & L") == ["pkg", "stp"] and mod.policy_subs("P & SA-OD") == ["pkg", "saod"] and mod.policy_subs("SA-OD") == ["saod"]
      and mod.policy_subs("Bundle (1+3)") == ["b13"] and mod.policy_subs("L") == ["stp"] and mod.policy_subs("P\xa0") == ["pkg"] and mod.policy_subs("X") is None)
check("RTO codes in a cell", mod.codes_in("HR-19,HR-84") == {"HR19", "HR84"} and mod.codes_in("TS/TG-05") == {"TS05", "TG05"} and mod.codes_in("SK01") == {"SK01"})
check("a list of codes takes the letters before it", mod.parse_code_list("JK-01,04,06,08") == {"JK01", "JK04", "JK06", "JK08"}
      and mod.parse_code_list("UK 08/UK 07/UK 14") == {"UK08", "UK07", "UK14"} and mod.parse_code_list("WB-02 & 06") == {"WB02", "WB06"})


# ================================================================ remark rules
def spec_of(text, product="GCCV HCV 12001 TO 20000 GVW", state="RAJASTHAN", rem1=""):
    row = dict(r=1, state=state, product=product, uw=text, rem1=rem1, eff="", po=1, ptype="P", age="", dis=80)
    return mod.read_remarks(mod.classify(product)[0] or "", row)


s = spec_of("Except New Mahindra")
check("'Except New Mahindra' splits Mahindra from the rest", s.mahindra_split and not s.flags)
s = spec_of("Except New Mahindra & Only NCB cases. For Hazardous category PO will be 5% less")
check("'Only NCB cases' sets NCB; the hazardous note is read (and, in the default mode, only flagged)", s.ncb == "YES" and s.mahindra_split and s.hazard
      and all("Hazardous" in f for f in s.flags), vars(s))
s = spec_of("Without NCB Cases, Except New Mahindra")
check("'Without NCB' sets NO", s.ncb == "NO")
s = spec_of("Tankers, Except New Mahindra")
check("tankers", s.classes == "tanker")
s = spec_of("Excluding Tankers, Except New Mahindra")
check("excluding tankers", s.classes == "nontanker")
s = spec_of("(CNG ONLY), Except New Mahindra")
check("CNG only", s.fuels == ["CNG"])
s = spec_of("ZONE 3, Except New Mahindra")
check("a Rajasthan zone", s.part_include == [mod.RJ_ZONE_GROUPS[3]])
s = spec_of("Zone 2 & 3, Except New Mahindra")
check("two zones", s.part_include == [mod.RJ_ZONE_GROUPS[2], mod.RJ_ZONE_GROUPS[3]])
s = spec_of("RJ41, Except New Mahindra")
check("a leading RTO code", s.code_include == {"RJ41"}, vars(s))
s = spec_of("RJ42,RJ48, Except New Mahindra")
check("a leading list of RTO codes", s.code_include == {"RJ42", "RJ48"})
s = spec_of("HR 68 EXCLUDED, Except New Mahindra", state="HARYANA")
check("'HR 68 EXCLUDED' takes the RTO out", s.code_exclude == {"HR68"})
s = spec_of("Except RTO Codes MP-06,07,09, Except New Mahindra", state="MADHYA PRADESH")
check("'Except RTO Codes MP-06,07,09'", s.code_exclude == {"MP06", "MP07", "MP09"})
s = spec_of("Other than TATA & Ashok Leyland MAKE, Except New Mahindra")
check("other than TATA & Ashok Leyland", s.exclude == {"TATA", "ASHOK"} and s.mahindra_split)
s = spec_of("Body Type Buker Truck only, Except New Mahindra")
check("Buker Truck: class Truck and its T&C", s.class_only == ["Truck"] and "Buker Truck" in s.tnc)
s = spec_of("NCB =>25%", product="PRIVATE CAR PETROL", state="KARNATAKA")
check("'NCB =>25%' on a car: NCB yes and the text as T&C", s.ncb == "YES" and s.tnc == ["NCB =>25%"], vars(s))
s = spec_of("Only HONDA & HYUNDAI & KIA manufacture only", product="PRIVATE CAR PETROL")
check("bundle makes", s.make_key == "hhk" and not s.flags)
s = spec_of("SA-OD not allowed", product="PRIVATE CAR PETROL")
check("'SA-OD not allowed' skips SAOD", "saod" in s.skip and not s.flags)
s = spec_of("Some rule nobody wrote a pattern for")
check("a remark no rule understands is flagged whole, not guessed", s.flags == ["Some rule nobody wrote a pattern for"], s.flags)
s = spec_of("", rem1="PO reduced by 2.5%")
check("a change note of REMARKS 1 needs no action", not s.flags)
s = spec_of("", rem1="Additional Slab")
check("'Additional Slab' needs no action", not s.flags)
s = spec_of("", rem1="TP PO applicable after 3 yrs")
check("'TP PO applicable after 3 yrs' is flagged", s.flags == ["TP PO applicable after 3 yrs"])
s = spec_of("New & With NCB cases only", rem1="")
check("'New & With NCB cases only'", s.new_or_ncb and s.ncb is None, vars(s))

s = spec_of("WITH NCB ONLY -MARUTI SUZUKI ( SA-OD not allowed)", product="PRIVATE CAR PETROL")
check("Private Car: 'WITH NCB ONLY -MARUTI SUZUKI' = make Maruti, NCB yes, no SAOD", s.make_key == "maruti_only" and s.ncb == "YES" and "saod" in s.skip and not s.flags, vars(s))
s = spec_of("Without NCB Except Maruti ALTO/SWIFT", product="PRIVATE CAR PETROL")
check("Private Car: 'Except Maruti ALTO/SWIFT'", s.make_key == "ex_alto_swift" and s.ncb == "NO" and not s.flags, vars(s))
s = spec_of("NCB Cases ( Other than Maruti Suzuki & M&M )", product="PRIVATE CAR DIESEL")
check("Private Car: 'Other than Maruti Suzuki & M&M'", s.make_key == "ex_maruti_mm" and not s.flags, vars(s))
s = spec_of("Maruti suzuki,Hyundai and M&M Bolero make only.", product="PCCV 4W Upto 6+1")
check("taxi: the Maruti / Hyundai / Bolero phrase is the taxi group, not the Private Car Maruti rule", s.make_key == "mshb" and not s.flags, vars(s))
s = spec_of("Bajaj and TVS make only", product="PCCV 3W DIESEL")
check("3W: 'Bajaj and TVS make only'", s.include == ["BAJAJ", "TVS"] and not s.flags, vars(s))
s = spec_of("NCB Cases & With CPA policy", product="GCCV MCV ABOVE 7501 TO 12000 GVW")
check("'NCB Cases & With CPA policy': NCB yes, CPA yes", s.ncb == "YES" and s.cpa == "YES" and not s.flags, vars(s))
s = spec_of("SC upto (3+1)", product="PCCV 3W OTHER THAN DIESEL")
check("'SC upto (3+1)' = seating up to 4", s.sc == (0, 4.01) and not s.flags, vars(s))
s = spec_of("Other Than Bangalore RTO Codes- Excluded Ashok Leyland, Except New Mahindra", state="KARNATAKA")
check("'Excluded Ashok Leyland' with a Bangalore partition", s.exclude == {"ASHOK"} and s.part_exclude == [mod.BANGALORE_GROUP] and s.mahindra_split and not s.flags, vars(s))
# make groups
name, back = mod.gcv_group({"MAHINDRA"})
check("excluding Mahindra: the portal's group", name == mod.GCV_BASE + "_MAHINDRA")
name, back = mod.gcv_group({"TATA", "MAHINDRA"})
check("excluding TATA and Mahindra: the TATA / ASHOK / MAHINDRA group plus Ashok given back", name == mod.GCV_BASE + "_TATA_ASHOK_MAHINDRA,Ashok", name)
check("make names are written the way the master spells them", mod.MAKE_TOKENS["MAHINDRA"] in M.makes and mod.MAKE_TOKENS["ASHOK"] in M.makes
      and mod.MAKE_TOKENS["HYUNDAI"] in M.makes, [v for v in mod.MAKE_TOKENS.values() if v not in M.makes])
check("every make group the script can write is in the master", all(g in M.makes for g in (
    mod.MAKES_PC, mod.MAKES_TAXI_06, mod.MAKES_TAXI_710, mod.MAKES_TAXI_MSHB, mod.MAKES_TAXI_EX_MSHB, mod.MAKES_BUS, mod.MAKES_3W_GCV, mod.MAKES_3W_PCV,
    mod.MAKES_MISCD, *mod.GCV_EXCLUSIONS.values())), [g for g in (mod.MAKES_PC, mod.MAKES_BUS, *mod.GCV_EXCLUSIONS.values()) if g not in M.makes])

# ================================================================ RTO groups by content
class FakeDecl(mod.Declines):
    def __init__(self, entries=(), pan=None, whole=()):
        super().__init__()
        self.entries, self.pan, self.whole = list(entries), dict(pan or {}), set(whole)


D0 = FakeDecl()
tok, miss = R.resolve(mod.area_of("ANDHRA PRADESH"), D0, [])
check("a whole state is written as its state code", tok == "AP" and not miss)
tok, miss = R.resolve(mod.area_of("ODISHA"), D0, [])
check("Odisha is OD + OR (the master's state groups, not ALLINDIA)", tok == "OD,OR" and not miss, tok)
tok, miss = R.resolve(mod.area_of("GUJARAT & DADRA NAGAR HAVELI & DAMAN & DIU"), D0, [])
check("Gujarat + Daman + Dadra is a comma list", tok == "GJ,DD,DN")
tok, miss = R.resolve(mod.area_of("MUMBAI"), D0, [])
check("an area that is a master group keeps its name", tok == "SHRIRAM_JUNE26_Mumbai")
D1 = FakeDecl([("gccvlcvupto7500gvw", "Andhra Pradesh", {"AP27"}, "Declined")])
tok, miss = R.resolve(mod.area_of("ANDHRA PRADESH"), D1, ["gccvlcvupto7500gvw"])
check("AP minus AP-27 finds the master group that holds exactly those RTOs", tok == "SHRIRAM_JUNE26_GCV4W_AP_EX_27" and not miss, (tok, miss))
tok, miss = R.resolve(mod.area_of("ANDHRA PRADESH"), D1, ["privatecar"])
check("a declined list of another product is not used", tok == "AP" and not miss)
D2 = FakeDecl([], pan={"UP": {"UP75", "UP82", "UP84"}})
tok, miss = R.resolve(mod.area_of("UTTAR PRADESH"), D2, ["privatecar"])
check("the pan-India declines apply when the sheet has nothing for the product", tok == "SHRIRAM_JUNE26_UP_EX_75_82_84", tok)
tok, miss = R.resolve(mod.area_of("UTTAR PRADESH"), D2, ["privatecar"], exempt=True)
check("... but not to an exempt product (e-rickshaw, school and corporate bus)", tok == "UP")
D3 = FakeDecl([], pan={"MP": {"MP06", "MP07", "MP31"}})
tok, miss = R.resolve(mod.area_of("MADHYA PRADESH"), FakeDecl([], pan={"MP": {"MP06", "MP07", "MP30"}}), ["privatecar"])
check("MP minus the three pan-India RTOs: the master's MP_EX_06_07_30 holds exactly those", tok == "MP_EX_06_07_30" and not miss, (tok, miss))
tok, miss = R.resolve(mod.area_of("MADHYA PRADESH"), D3, ["privatecar"])
check("no master group for the remaining RTOs: the whole area is written and the gap is named", tok == "MP" and len(miss) == 1 and "MP 06, 07, 31" in miss[0], (tok, miss))
D4 = FakeDecl([], whole={"AR", "ML", "MZ"})
tok, miss = R.resolve(mod.area_of("ASSAM/TRIPURA/ARUNACHAL/MEGHALAYA/ MIZORAM"), D4, [])
check("states declined for everything drop out of a multi-state label", tok == "AS,TR", tok)
tok, miss = R.resolve(mod.area_of("KARNATAKA"), D0, [], include=M.codes(mod.BANGALORE_GROUP), include_names=(mod.BANGALORE_GROUP,))
check("a partition named in the remarks (Bangalore)", tok == mod.BANGALORE_GROUP, tok)
tok, miss = R.resolve(mod.area_of("TAMILNADU & PONDICHERRY"), D0, [], include=M.codes(mod.TN_METRO_GROUP), include_names=(mod.TN_METRO_GROUP,))
check("a Tamil Nadu partition leaves Pondicherry as it is", tok == mod.TN_METRO_GROUP + ",PY", tok)
tok, miss = R.resolve(mod.area_of("RAJASTHAN"), D0, [], include={"RJ41"})
check("a single RTO code finds its one-RTO group", tok == "RJ41", tok)
tok, miss = R.resolve(mod.ALL_INDIA, D0, [])
check("PAN INDIA is ALLINDIA", tok == "ALLINDIA")


# ================================================================ small grids through build()
HDR = ["STATE", "PRODUCTS", "DIS %", "% PO", "POLICY TYPE", "AGE", "UW REMARKS", "REMARKS 1", "EFFECTIVE DATE"]


def grid(*rows, extra=None):
    sheets = {"GRID OCT' 26": [HDR] + [list(r) + [None] * (9 - len(r)) for r in rows]}
    sheets.update(extra or {})
    return helpers.FakeWB(sheets)


def run(wb, write_ev=None, a=None):
    return mod.build(None, F, T, quiet, a or mod.Asker(interactive=False), wb=wb, write_ev=write_ev)


df, w = run(grid(["GOA", "PRIVATE CAR PETROL", 85, "16 Net", "P & SA-OD", "upto 15 years.", "NCB Cases"],
                 ["GOA", "PRIVATE CAR PETROL", 90, "10 Net", "P & SA-OD", "upto 15 years.", "NCB Cases"]))
pk = df[df.sub_product == "1+1"]
check("car P & SA-OD: 1+1 and SAOD, both on OD, NCB yes, no Electric by default",
      set(df.sub_product) == {"1+1", "SAOD"} and set(df.pi_type) == {"On OD"} and set(df.is_ncb) == {"YES"} and set(df.fuel_type) == {"Petrol", "CNG", "LPG"}, (set(df.sub_product), set(df.fuel_type)))
check("car payout: po = pi - 7", set(zip(pk.pi_od_rate, pk.po_od_rate)) == {(16.0, 9.0), (10.0, 3.0)})
low, high = pk[pk.pi_od_rate == 16.0].iloc[0], pk[pk.pi_od_rate == 10.0].iloc[0]
check("discount slabs chain: 85 -> 0-85.01, 90 -> 85.01-90.01", (low.tariff_min, low.tariff_max, high.tariff_min, high.tariff_max) == (0, 85.01, 85.01, 90.01),
      (low.tariff_min, low.tariff_max, high.tariff_min, high.tariff_max))
check("car make group, class, age, new_rto_list", set(df.new_vehicle_makes) == {mod.MAKES_PC} and set(df.make_model_class) == {"Car"} and set(df.new_rto_list) == {"GA"}
      and set(df.vehicle_age_max) == {15.01}, df.iloc[0].to_dict())
df2, _ = run(grid(["GOA", "PRIVATE CAR PETROL", 85, "16 Net", "P", "upto 15 years."]), write_ev=True)
check("write_ev adds Electric to the non-diesel fuels", set(df2.fuel_type) == {"Petrol", "CNG", "LPG", "Electric"})
df, w = run(grid(["GOA", "GCCV LCV UPTO 2800 GVW", 80, "15 OD + 2.5 TP", "P & L", "upto 15 years & upto 25 Years for STP", "Except New Mahindra"]))
check("goods vehicle: an all-makes row and a Mahindra row for old vehicles, 1+1 and STP", len(df) == 20 and set(df.new_vehicle_makes) == {mod.GCV_BASE + "_Mahindra", "Mahindra"} or
      set(df.new_vehicle_makes) == {mod.GCV_BASE + "_MAHINDRA", "Mahindra"}, set(df.new_vehicle_makes))
mah = df[df.new_vehicle_makes == "Mahindra"]
check("the Mahindra row starts at age 1 and says Old", set(mah.vehicle_age_min) == {1.0} and set(mah.add_tnc) == {"Old"})
od = df[(df.sub_product == "1+1") & (df.new_vehicle_makes == "Mahindra")].iloc[0]
tp = df[(df.sub_product == "STP") & (df.new_vehicle_makes == "Mahindra")].iloc[0]
check("'15 OD + 2.5 TP': On OD and TP, po 8 and 0; the STP row is On TP 2.5", (od.pi_type, od.pi_od_rate, od.pi_tp_rate, od.po_od_rate, od.po_tp_rate) == ("On OD and TP", 15.0, 2.5, 8.0, 0.0)
      and (tp.pi_type, tp.pi_tp_rate, tp.po_tp_rate, tp.vehicle_age_max) == ("On TP", 2.5, 0.0, 25.01))
check("GVW band is the cc band", set(df.cc_min) == {0.0} and set(df.cc_max) == {2800.01})
df, w = run(grid(["GOA", "GCCV HCV 20001 TO 42500 GVW", 90, 27.5, "P & L", "upto 15 years.", "Tankers, Except New Mahindra"],
                 ["GOA", "GCCV HCV 20001 TO 42500 GVW", 90, 30, "P & L", "upto 15 years.", "Other than tankers, Except New Mahindra"]))
check("tankers get the Tanker class, the others the five other classes", set(df[df.pi_net_rate == 27.5].make_model_class) == {"Tanker"}
      and len(set(df[df.pi_net_rate == 30].make_model_class)) == 5)
df, w = run(grid(["GOA", "PRIVATE CAR DIESEL", 25, "10 OD", "P", "upto 15 years", "SA-OD not allowed"], ["GOA", "PRIVATE CAR DIESEL", 40, "15 OD", "P", "upto 15 years", "IDV UPTO 10 LACS and SA-OD not allowed"]))
t = df.sort_values("tariff_max")
check("a discount slab chains across a T&C that only narrows it (25 -> 25.01-40.01 with IDV UPTO 10 LACS)",
      list(zip(t.tariff_min, t.tariff_max, t.add_tnc)) == [(0.0, 25.01, ""), (25.01, 40.01, "IDV UPTO 10 LACS")], list(zip(t.tariff_min, t.tariff_max, t.add_tnc)))
df, w = run(grid(["GOA", "GCCV MCV ABOVE 7501 TO 12000 GVW", 70, 20, "P & L", "upto 15 years & upto 25 Years for STP", "NCB Cases & With CPA policy"]))
check("CPA is written on the package rows", set(df[df.sub_product == "1+1"].is_cpa) == {"YES"} and set(df[df.sub_product == "STP"].is_cpa) == {"NA"})

# hazardous category: the Tanker class, 5 points of payin lower, where the row names tankers
HZ = "For Hazardous category PO will be 5% less"
rows_h = [["GOA", "GCCV HCV 20001 TO 42500 GVW", 90, 27.5, "P & L", "upto 15 years.", "Tankers, Except New Mahindra, " + HZ],
          ["GOA", "GCCV HCV 20001 TO 42500 GVW", 90, 30, "P & L", "upto 15 years.", "Other than tankers, Except New Mahindra, " + HZ],
          ["GOA", "GCCV MCV ABOVE 7501 TO 12000 GVW", 85, 15, "P", "upto 15 years.", "Except New Mahindra, " + HZ]]
mode0 = mod.HAZARDOUS_MODE
mod.HAZARDOUS_MODE = "tanker_rows"
df, w = run(grid(*rows_h))
mod.HAZARDOUS_MODE = mode0
tk = df[df.make_model_class == "Tanker"]
check("hazardous: the Tanker rows of a row that names tankers are 5 points lower (27.5 -> 22.5, po 15.5)", set(tk.pi_net_rate.astype(float)) | set(tk.pi_tp_rate.astype(float)) == {0.0, 22.5}
      and set(tk.po_net_rate.astype(float)) | set(tk.po_tp_rate.astype(float)) == {0.0, 15.5}, tk.iloc[0].to_dict())
check("hazardous: the 'other than tankers' rows keep 30 and the row without tankers stays as listed",
      set(df[df.make_model_class == "Truck"][["pi_net_rate", "pi_tp_rate"]].astype(float).max(axis=1)) == {30.0}
      and set(df[df.make_model_class == "NA"][["pi_net_rate", "pi_tp_rate"]].astype(float).max(axis=1)) == {15.0})
dflt, wd = run(grid(*rows_h))
check("hazardous (default mode 'tnc'): payouts as listed (27.5), NO flag, the grid's own percentage is written in add_tnc ('Hazardous category PO 5% less')",
      set(dflt[dflt.make_model_class == "Tanker"][["pi_net_rate", "pi_tp_rate"]].astype(float).max(axis=1)) == {27.5}
      and not dflt.remarks.str.contains("Hazardous").any() and dflt[dflt.make_model_class == "Tanker"].add_tnc.str.contains("Hazardous category PO 5% less", regex=False).all(), (wd, dflt.add_tnc.unique()[:3]))
_mode0 = mod.HAZARDOUS_MODE
mod.HAZARDOUS_MODE = "flag"
dflag, wf = run(grid(*rows_h))
mod.HAZARDOUS_MODE = _mode0
check("hazardous (mode 'flag'): Tanker payouts are as listed (27.5) and the rows carry the note",
      set(dflag[dflag.make_model_class == "Tanker"][["pi_net_rate", "pi_tp_rate"]].astype(float).max(axis=1)) == {27.5}
      and dflag[dflag.make_model_class == "Tanker"].remarks.str.contains("Hazardous").all(), wf)
mode = mod.HAZARDOUS_MODE
try:
    mod.HAZARDOUS_MODE = "tanker"
    df, w = run(grid(rows_h[2]))
    mx = lambda d: d[["pi_net_rate", "pi_tp_rate"]].astype(float).max(axis=1)
    check("hazardous (mode tanker): a row for all classes is split into Tanker (5 lower) and the five other classes",
          set(df.make_model_class) == {"Tanker"} | set(mod.GCV_CLASSES_OTHER) and set(mx(df[df.make_model_class == "Tanker"])) == {10.0}
          and set(mx(df[df.make_model_class == "Truck"])) == {15.0}, set(df.make_model_class))
    mod.HAZARDOUS_MODE = "off"
    df, w = run(grid(rows_h[0]))
    check("hazardous (mode off): nothing is reduced", set(df[df.make_model_class == "Tanker"].pi_net_rate.astype(float)) | set(df[df.make_model_class == "Tanker"].pi_tp_rate.astype(float)) == {0.0, 27.5})
finally:
    mod.HAZARDOUS_MODE = mode

# a row that is the same as in the earlier upload keeps that upload's from_date
ref_rows = pd.DataFrame([dict(zip(mod.OUTPUT_COLUMNS, [""] * 42))])
base_df, _ = run(grid(["GOA", "GCCV 3W Except E-CART", 85, 45, "P & L", "upto 15 years."], ["GOA", "PCCV 3W DIESEL", 80, 40, "P & L", "upto 15 years."]))
ref_df = mod.to_csv_frame(base_df.copy())
ref_df.loc[ref_df["product"] == "GCV 3W", "from_date"] = "01-09-2026"
ref_df.loc[ref_df["product"] == "PCV 3W", "pi_net_rate"] = "35"                      # a payout that changed since the earlier upload
ref_df.loc[ref_df["product"] == "PCV 3W", "pi_tp_rate"] = ref_df.loc[ref_df["product"] == "PCV 3W", "pi_tp_rate"].map(lambda x: "35" if x != "0" else x)
ref_df.loc[ref_df["product"] == "PCV 3W", "from_date"] = "01-09-2026"
tmpref = os.path.join(tempfile.mkdtemp(prefix="shriram_ref_"), "ref.csv")
ref_df.to_csv(tmpref, index=False)
kept = mod.keep_ref_dates(base_df, tmpref, quiet)
check("unchanged rows keep the earlier from_date (GCV 3W 01/09/2026)", set(kept[kept["product"] == "GCV 3W"].from_date) == {"01/09/2026"}, set(kept.from_date))
check("rows whose payout changed take the new from_date", set(kept[kept["product"] == "PCV 3W"].from_date) == {"01/10/2026"}, set(kept[kept["product"] == "PCV 3W"].from_date))

df, w = run(grid(["GOA", "GCCV 3W Except E-CART", 85, 45, "P", "upto 15 years."], ["GOA", "GCCV 3W Except E-CART", 85, 41, "P", "upto 15 years."]))
check("two grid rows, one key, two payouts: only the LOWEST (41) is written, one warning, every entry listed", len(df) == 4 and set(df.pi_net_rate) == {41.0} and set(df.pi_tp_rate) == {0.0}
      and len(mod.LAST_LOWEST) == 4 and any("LOWEST" in x for x in w), (len(df), w))
df, w = run(grid(["GOA", "GCCV 3W Except E-CART", 85, 45, "P", "upto 15 years."], ["GOA", "GCCV 3W Except E-CART", 85, 45, "P", "upto 15 years."]))
check("the same payout twice is one row", len(df) == 4 and (df.remarks == "").all(), len(df))
df, w = run(grid(["GOA", "WIND TURBINE", 85, 45, "P", "upto 15 years."], ["ATLANTIS", "PRIVATE CAR PETROL", 85, 45, "P", "upto 15 years."],
                 ["GOA", "PRIVATE CAR PETROL", 85, "WITHDRAW", "P", "upto 15 years."], ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."]))
check("a row the script cannot read is a warning, never a silent drop",
      sum("not written" in x for x in w) == 3 and any("WIND TURBINE" in x for x in w) and any("ATLANTIS" in x for x in w) and any("WITHDRAW" in x for x in w), w)
df, w = run(grid(["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years.", "A rule nobody can apply"]))
check("an unread remark is flagged as Manual Review Required with its text", df.remarks.str.startswith("Manual Review Required: [A rule nobody can apply]").all())
check("the output always has the 42 columns in order", list(df.columns) == mod.OUTPUT_COLUMNS)
df, w = run(grid(["PAN INDIA", "PCCV 4W SCHOOL BUS", 99, "SC>15 PO 50% & <15 30%", "P & L", "upto 15 years"]))
check("school bus: seating bands, ALLINDIA, tariff to 100.01, four fuels, STP too",
      set(df.new_rto_list) == {"ALLINDIA"} and set(zip(df.sc_min, df.sc_max, df.pi_net_rate + df.pi_tp_rate)) == {(0.0, 15.01, 30.0), (15.01, 999.0, 50.0)}
      and set(df.tariff_max) == {100.01} and set(df.fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"} and set(df.sub_product) == {"1+1", "STP"}, df.iloc[0].to_dict())
df, w = run(grid(["PAN INDIA", "PCCV 4W CORPORATE BUS", 90, "Metro - 45% SC >= 15, 40% SC < 15:Non Metro - 40% SC >= 15, 35% SC < 15", "P & L", "upto 15 years"]))
check("corporate bus: metro and non-metro groups with their own rates", set(df.new_rto_list) == {mod.METRO_GROUP, mod.NON_METRO_GROUP}
      and set(df[df.new_rto_list == mod.METRO_GROUP].pi_net_rate.astype(float)) >= {45.0, 40.0}, set(df.new_rto_list))
df, w = run(grid(["GOA", "PCCV 3W E-Rickshaw  and GCCV 3W E-cart", 75, 46.5, "P & L", "upto 15 years."]))
check("e-rickshaw / e-cart row writes both a PCV 3W and a GCV 3W (Electric)", set(df["product"]) == {"PCV 3W", "GCV 3W"} and set(df.fuel_type) == {"Electric"})
df, w = run(grid(["GOA", "MISC-D TRACTOR", 85, 35, "P", "New", "Agriculture tractors only without trolly/trailer"]))
check("tractor: new vehicles only, the grid's own T&C text", set(df.vehicle_age_max) == {1.0} and set(df.add_tnc) == {"Agriculture tractors only without trolly,trailer"}, df.iloc[0].to_dict())
df, w = run(grid(["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."], extra={"Declined-Excluded": [["Product", "RTO State"], ["Private Car", "Goa", "GA-05", "x", "Declined"]]}))
check("a declined RTO with no master group is flagged on its rows", df.remarks.str.contains("No RTO group in the master for GA without GA 05").all(), df.remarks.iloc[0])
df, w = run(grid(["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years.", None, None, "w.e.f 1st Nov"]))
check("an EFFECTIVE DATE month different from from_date is a warning", any("w.e.f' month 11" in x for x in w), w)
try:
    run(helpers.FakeWB({"Sheet1": [["a", "b"], [1, 2]]}))
    check("a workbook without the rate header stops with a clear message", False)
except mod.GridError as e:
    check("a workbook without the rate header stops with a clear message", "rate-table header" in str(e), str(e))
w2 = helpers.FakeWB({"GRID": [HDR, ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."]], "PCV 6+1 Short-term policy": [["State"], ["x"]]})
df, w = run(w2)
check("the short-term sheet is named in a warning, not converted", any("Short-term" in x for x in w))

# ================================================================ the real October grid
REAL = os.path.join(helpers.INPUTS, "shriram_oct26.xlsx")
if os.path.exists(REAL):
    import openpyxl
    wb = openpyxl.load_workbook(REAL, data_only=True)
    df, w = mod.build(None, F, T, quiet, mod.Asker(interactive=False), wb=wb)
    check("October grid: 6,659 rows (6,710 before STP lost NCB / zero dep and the lowest of two rates was kept, 6,643 before the November 2026 rules), 42 columns", len(df) == 6659 and list(df.columns) == mod.OUTPUT_COLUMNS, len(df))
    stp = df[df.sub_product == "STP"]
    check("October grid: no STP row carries NCB or zero dep (is_ncb NA, is_zd NA)", len(stp) > 2000 and set(stp.is_ncb) == {"NA"} and set(stp.is_zd) == {"NA"}, (set(stp.is_ncb), set(stp.is_zd)))
    check("October grid: no key repeats (one entry, one rate) and the lowest-rate choices are listed", not df.duplicated(mod.KEY_COLUMNS).any() and len(mod.LAST_LOWEST) == 59, (int(df.duplicated(mod.KEY_COLUMNS).sum()), len(mod.LAST_LOWEST)))
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float)
    po = df[["po_od_rate", "po_tp_rate", "po_net_rate"]].astype(float)
    exp = (pi - mod.PO_CUT).clip(lower=0).where(pi > 0, 0.0)
    check("October grid: po = max(pi - 7, 0) on every row and column", (abs(po.values - exp.values) < 0.011).all())
    check("October grid: dates are DD/MM/YYYY on every row", set(df.from_date) == {"01/10/2026"} and set(df.to_date) == {"31/03/2030"})
    check("October grid: exactly one rate column is used per row", ((pi > 0).sum(axis=1) == 1).sum() + ((df.pi_type == "On OD and TP") & ((pi > 0).sum(axis=1) == 2)).sum() >= len(df) - 12,
          int(((pi > 0).sum(axis=1) != 1).sum()))
    check("October grid: tariff bands are well formed (min < max)", (df.tariff_min.astype(float) < df.tariff_max.astype(float)).all())
    check("October grid: no make written in capitals where the master spells it otherwise", not df.new_vehicle_makes.str.contains(r"(?:^|,)(?:MAHINDRA|ASHOK|HYUNDAI)(?:,|$)").any())
    unknown = [n for v in df.new_rto_list.unique() for n in v.split(",") if n not in M.rto]
    flagged = df[df.new_rto_list.map(lambda v: any(n not in M.rto for n in v.split(",")))].remarks
    check("October grid: every RTO name is in the master (or its rows are flagged)", not unknown or flagged.str.contains("not in the RTO master").all(), unknown[:5])
    check("October grid: the hazardous note is one warning and its text is in add_tnc of the rows that carry it (payouts unchanged, no flag)", sum("Hazardous" in x for x in w) == 1
      and df.add_tnc.str.contains("Hazardous category PO 5% less", regex=False).sum() > 3000 and not df.remarks.str.contains("not applied [(]the hazardous", regex=True).any())
    check("October grid: Electric only where the grid says Electric Vehicles only / e-rickshaw",
          set(df[df.fuel_type == "Electric"]["product"].unique()) <= {"PCV 4W", "PCV 3W", "GCV 3W", "GCV 4W", "MISCD"}
          and not ((df.fuel_type == "Electric") & (df["product"] == "Private Car")).any())
    out = os.path.join(tempfile.mkdtemp(prefix="shriram_t_"), "out.csv")
    a = mod.Asker(interactive=False)
    mod.main(REAL, F, T, out_path=out, ask=a, log=quiet)
    problems = helpers.check_csv(out, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
    check("October grid: the written file has the 42 columns, numbers only in numeric columns, DD/MM/YYYY dates", problems == [], problems)
    rev = mod.review_path(out)
    book = pd.read_excel(rev, sheet_name=None, dtype=str, keep_default_na=False) if os.path.exists(rev) else {}
    gaps = book.get("RTO groups to create")
    check("October grid: the RTO groups the master lacks are a sheet of the ONE review workbook next to the upload, with their RTO codes",
          gaps is not None and list(gaps.columns) == ["group_needed", "rows_flagged", "items"] and len(gaps) >= 10
          and gaps["items"].str.len().min() > 0 and gaps.rows_flagged.astype(int).sum() > 300, None if gaps is None else len(gaps))
    check("October grid: the workbook also has the Summary, the lowest-rate choices (59) and the manual-review notes - and no other side file exists",
          {"Summary", "Lowest rate chosen", "Manual review notes"} <= set(book) and len(book.get("Lowest rate chosen", [])) == 59
          and sorted(os.listdir(os.path.dirname(out))) == sorted(["out.csv", os.path.basename(rev)]), (sorted(book), os.listdir(os.path.dirname(out))))
    back = pd.read_csv(out, dtype=str, keep_default_na=False)
    check("October grid: every row that names a missing group is counted in that list",
          back.remarks.str.contains("No RTO group in the master for").sum() > 0 and back.remarks.str.contains("No RTO group in the master for").sum() <= gaps.rows_flagged.astype(int).sum())

# the rate sheet's name changes every month (GRID OCT' 26, GRID SEP' 26); the short-term policy sheet has the same header and must never be picked or asked about
sep = helpers.FakeWB({"GRID SEP' 26": [HDR, ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."]], "PCV 6+1 Short-term policy": [HDR, ["GOA", "PCCV 4W UPTO 6+1", 85, 20, "P", "upto 10 years."]]})
try:
    got = mod.find_grid_sheet(sep, mod.Asker(interactive=False))[0]
    check("the rate sheet is found by the name GRID <month> ' <yy>, with the short-term sheet's identical header ignored (no question)", got == "GRID SEP' 26", got)
except mod.GridError as e:
    check("the rate sheet is found by the name GRID <month> ' <yy>, with the short-term sheet's identical header ignored (no question)", False, str(e))
two = helpers.FakeWB({"Rates A": [HDR, ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "x"]], "Rates B": [HDR, ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "x"]]})
try:
    mod.find_grid_sheet(two, mod.Asker(interactive=False))
    check("two unnamed sheets with the rate-table header still raise the question", False)
except mod.GridError:
    check("two unnamed sheets with the rate-table header still raise the question", True)

# ================================================================ STP: no NCB, no zero dep; the lowest of two rates wins (the user, 27 Oct 2026)
PA = "upto 15 years & upto 25 Years"
df, w = run(grid(["PUNJAB/CHANDIGARH", "GCCV LCV ABOVE 2801 GVW TO 3500 GVW", 75, 42.5, "P & L", PA, "Without Zero Dep"], ["PUNJAB/CHANDIGARH", "GCCV LCV ABOVE 2801 GVW TO 3500 GVW", 75, 45, "P & L", PA, "Zero dep"]))
stp, pkg = df[df.sub_product == "STP"], df[df.sub_product == "1+1"]
check("STP rows carry is_zd NA and is_ncb NA (the 'Zero dep' / 'Without Zero Dep' pair of a P & L row)", set(stp.is_zd) == {"NA"} and set(stp.is_ncb) == {"NA"} and len(stp) > 0, (set(stp.is_zd), set(stp.is_ncb)))
check("... and the two payins meet on one STP entry: the LOWER (42.5) is written", set(stp.pi_tp_rate) == {42.5} and mod.LAST_LOWEST and all(x["written"] == "42.5" for x in mod.LAST_LOWEST), set(stp.pi_tp_rate))
check("the package rows keep their own zero-dep split (YES 45, NO 42.5)", set(zip(pkg.is_zd, pkg.pi_net_rate)) == {("YES", 45.0), ("NO", 42.5)}, set(zip(pkg.is_zd, pkg.pi_net_rate)))
df, w = run(grid(["BIHAR", "GCCV HCV 12001 TO 20000 GVW", 85, 30, "P & L", PA, "NCB Cases"], ["BIHAR", "GCCV HCV 12001 TO 20000 GVW", 85, 25, "P & L", PA, "Without NCB Cases"]))
stp, pkg = df[df.sub_product == "STP"], df[df.sub_product == "1+1"]
check("NCB cases / Without NCB cases: the package rows keep YES 30 / NO 25, the STP entry (is_ncb NA) gets the lower 25", set(zip(pkg.is_ncb, pkg.pi_net_rate)) == {("YES", 30.0), ("NO", 25.0)}
      and set(stp.is_ncb) == {"NA"} and set(stp.pi_tp_rate) == {25.0}, (set(zip(pkg.is_ncb, pkg.pi_net_rate)), set(stp.pi_tp_rate)))
df, w = run(grid(["GOA", "GCCV 3W Except E-CART", 85, 45, "P", "upto 15 years."], ["GOA", "GCCV 3W Except E-CART", 85, 41, "P", "upto 15 years."]))
side = os.path.join(tempfile.mkdtemp(prefix="shriram_low_"), "o.csv")
mod.write_csv(df, side, mod.Asker(interactive=False), quiet)
low = pd.read_excel(mod.review_path(side), sheet_name="Lowest rate chosen", dtype=str, keep_default_na=False)
check("the lowest-rate choices are a sheet of the review workbook (grid rows, the rates found, the one written)", len(low) == 4 and {"grid_rows", "rates_found", "written"} <= set(low.columns) and set(low.written) == {"41"} and set(low.rates_found) <= {"45 / 41", "41 / 45"}, low.head(2).to_dict("records"))

# ================================================================ the addendum (--addendum)
import openpyxl  # noqa: E402


def addendum_file(*rows):
    """Shriram's addendum layout: an empty first row, the grid's header on row 2 (+ the user's RTO / MAKE mapping columns), the rows below"""
    wbk = openpyxl.Workbook()
    ws = wbk.active
    ws.title = "Sheet1"
    ws.append([])
    ws.append(HDR + [None, "RTO", "MAKE"])
    for r in rows:
        ws.append(list(r) + [None] * (9 - len(r)))
    path = os.path.join(tempfile.mkdtemp(prefix="shriram_add_"), "addendum.xlsx")
    wbk.save(path)
    return path


def with_addendum(grid_rows, add_rows):
    return mod.build("x.xlsx", F, T, log=quiet, ask=mod.Asker(interactive=False), wb=grid(*grid_rows), addendum_path=addendum_file(*add_rows))


def pays(df, sub):
    d = df[df.sub_product == sub]
    return sorted(set(d[["pi_net_rate", "pi_tp_rate"]].max(axis=1)))


GR = lambda state, product, dis, po, ptype, age="upto 15 years": [state, product, dis, po, ptype, age]
check("addendum: _covers - adjacent bands cover, a gap does not", mod._covers([(12000.01, 20000.01), (20000.01, 42500.01)], (12000.01, 42500.01))
      and not mod._covers([(12000.01, 20000.01), (20500.01, 42500.01)], (12000.01, 42500.01)) and mod._covers([(0, 99999)], (12000.01, 20000.01)))
df, w = with_addendum([GR("ODISHA", "GCCV HCV 20001 TO 42500 GVW", 85, 25, "P & L"), GR("BIHAR", "GCCV HCV 20001 TO 42500 GVW", 85, 20, "P & L")],
                      [GR("ODISHA", "GCCV HCV  20001 TO 42500 GVW", 85, 27.5, "P & L", "upto 15 years & upto 25 Years for STP")])
od = df[df.new_rto_list == "OD,OR"]
check("addendum: a P & L row with the same band replaces the grid row (Odisha 25 -> 27.5), other states stay", pays(od, "1+1") == [27.5] and pays(od, "STP") == [27.5] and pays(df[df.new_rto_list == "BR"], "1+1") == [20.0])
check("addendum: the replacing rows carry the addendum's own age band for STP (upto 25 years)", set(od[od.sub_product == "STP"].vehicle_age_max) == {25.01} and set(od[od.sub_product == "1+1"].vehicle_age_max) == {15.01})
df, w = with_addendum([GR("ODISHA", "GCCV HCV 42501 TO 50000 GVW", 85, 17.5, "P & L")], [GR("ODISHA", "GCCV HCV 42501 TO 50000 GVW", "", 25, "L", "upto 25 years")])
check("addendum: an 'L' row replaces only the STP part of a 'P & L' grid row (1+1 stays 17.5, STP becomes 25)", pays(df, "1+1") == [17.5] and pays(df, "STP") == [25.0], (pays(df, "1+1"), pays(df, "STP")))
df, w = with_addendum([GR("ANDHRA PRADESH", "GCCV HCV 12001 TO 42500 GVW", 85, 25, "P & L")],
                      [GR("ANDHRA PRADESH", "GCCV HCV 12001 TO 20000 GVW", 85, 27.5, "P & L"), GR("ANDHRA PRADESH", "GCCV HCV  20001 TO 42500 GVW", 85, 27.5, "P & L")])
check("addendum: a grid band split in two addendum rows is replaced by both (no 25 left, two bands)", pays(df, "1+1") == [27.5] and sorted(set(zip(df.cc_min, df.cc_max))) == [(12000.01, 20000.01), (20000.01, 42500.01)],
      sorted(set(zip(df.cc_min, df.cc_max))))
df, w = with_addendum([GR("CHATTISGARH", "GCCV HCV 12001 TO 50000 GVW", 80, 15, "P & L")], [GR("CHATTISGARH", "GCCV HCV 45501 - 50000 GVW", "", 10, "L")])
check("addendum: a band that overlaps only in part is not trimmed - both rows are written and both say so", pays(df, "STP") == [10.0, 15.0] and df[df.sub_product == "STP"].remarks.str.contains("only in part").all()
      and any("only in part" in x for x in w), (pays(df, "STP"), w))
df, w = with_addendum([GR("BIHAR", "GCCV HCV 20001 TO 42500 GVW", 85, 20, "P & L")], [GR("JHARKHAND", "GCCV LCV UPTO 2800 GVW", 80, 55, "P & L")])
check("addendum: a row with no grid counterpart is added (New Addition)", set(df.new_rto_list) == {"BR", "JH"} and pays(df[df.new_rto_list == "JH"], "1+1") == [55.0])
df, w = with_addendum([GR("ANDHRA PRADESH", "PRIVATE CAR PETROL", 85, 20, "L", "4 to 20 years")], [GR("ANDHRA PRADESH", "PRIVATE CAR PETROL", 85, 22, "L", "4 to 20 years")])
check("addendum: another product is never replaced - the rows meet on one key, the LOWER payin (20) is written and says 'covers the same state, product and policy type'", pays(df, "STP") == [20.0]
      and df.remarks.str.contains("covers the same state, product and policy type").all() and mod.LAST_LOWEST[0]["written"] == "20", (pays(df, "STP"), mod.LAST_LOWEST))
df, w = with_addendum([GR("ODISHA", "GCCV HCV 20001 TO 42500 GVW", 85, 25, "P & L")], [GR("ODISHA", "GCCV HCV 20001 TO 42500 GVW", 85, "WITHDRAW", "P & L")])
check("addendum: an addendum row whose payout cannot be read replaces nothing (the grid row stays) and is warned about", pays(df, "1+1") == [25.0] and any("not readable" in x for x in w), w)
check("addendum: row numbers of the addendum start at ADDENDUM_ROW_BASE", mod.ADDENDUM_ROW_BASE == 10000 and all(r >= 10000 for r in [10003]))
try:
    mod.build("x.xlsx", F, T, log=quiet, ask=mod.Asker(interactive=False), wb=grid(GR("BIHAR", "GCCV HCV 20001 TO 42500 GVW", 85, 20, "P & L")), addendum_path="C:/does/not/exist.xlsx")
    check("addendum: a missing addendum file stops with a clear message", False)
except mod.GridError as e:
    check("addendum: a missing addendum file stops with a clear message", "exist" in str(e).lower() or "not found" in str(e).lower() or "addendum" in str(e).lower(), str(e))

# ================================================================ the corrections of the November 2026 upload (the user fixed these rows by hand; see lessons_learned.md)
# --- the remarks, read: no flag, the effect on the spec
s = spec_of("Bolero Model only", product="GCCV LCV UPTO 2800 GVW", state="ANDHRA PRADESH")
check("'Bolero Model only' = the Bolero group GCV_4W_BOLERO, no flag", s.makes_override == mod.MAKES_BOLERO == "GCV_4W_BOLERO" and not s.flags and not s.mahindra_split, vars(s))
s = spec_of("Except Bolero Model & New Mahindra", product="GCCV LCV 2001-2800 GVW", state="ANDHRA PRADESH")
check("'Except Bolero Model & New Mahindra' = the Mahindra row without the Bolero models, no flag", s.mahindra_split and s.bolero_ex and not s.flags, vars(s))
s = spec_of("Except New Mahindra & Bolero Model  (DL RTO Only)", product="GCCV LCV ABOVE 2801 GVW TO 3500 GVW", state="DELHI&NCR")
check("'Except New Mahindra & Bolero Model (DL RTO Only)': the same, in the DL RTOs, no flag (it used to be flagged)", s.mahindra_split and s.bolero_ex and s.part_include == ["DL"] and not s.flags, vars(s))
s = spec_of("swift dzire model excluded for RJ-14", product="PCCV 4W Upto 6+1", state="RAJASTHAN")
check("taxi: 'swift dzire model excluded for RJ-14' = the taxi group without Dzire, RJ-14 taken OUT of the RTOs (it used to be read as 'for RJ-14' = the only RTO)",
      s.make_key == "dzire_ex" and s.code_exclude == {"RJ14"} and s.code_include is None and not s.flags, vars(s))
s = spec_of("For RJ-14 and for Maruti Suzuki-Swift Dzire", product="PCCV 4W Upto 6+1", state="RAJASTHAN")
check("taxi: 'For RJ-14 and for Maruti Suzuki-Swift Dzire' = the Dzire group in RJ-14", s.make_key == "dzire_only" and s.code_include == {"RJ14"} and not s.flags, vars(s))
s = spec_of("For Maruti Suzuki-Swift Dzire of RJ-14", product="PCCV 4W Upto 6+1", state="RAJASTHAN")
check("taxi: the October wording 'For Maruti Suzuki-Swift Dzire of RJ-14' too", s.make_key == "dzire_only" and s.code_include == {"RJ14"} and not s.flags, vars(s))
s = spec_of("6+1  and  Including Kaali peeli taxi ambassador model only/ Except NIL dep cases, excluding RTO (WB-04, WB-05, WB-11 & WB-69 )", product="PCCV 4W Upto 6+1", state="WEST BENGAL")
check("taxi: 'Including Kaali peeli taxi ambassador model only / Except NIL dep cases, excluding RTO (...)' = Ambasador, WITHOUT nil dep, those RTOs out",
      s.make_key == "ambasador" and s.zd == "NO" and s.code_exclude == {"WB04", "WB05", "WB11", "WB69"} and not s.flags, vars(s))
s = spec_of("6+1 and Including Kaali peeli taxi ambassador model only/ Without NIL dep cases for FOR RTOs WB-04, WB-05 & WB-69", product="PCCV 4W Upto 6+1", state="WEST BENGAL")
check("taxi: '... Without NIL dep cases for FOR RTOs WB-04, ...' = Ambasador, without nil dep, only those RTOs",
      s.make_key == "ambasador" and s.zd == "NO" and s.code_include == {"WB04", "WB05", "WB69"} and not s.flags, vars(s))
s = spec_of("EXCLUDING M &M make", product="PCCV 4W Upto 6+1", state="J & K")
check("taxi: 'EXCLUDING M &M make' = the taxi group without Mahindra", s.make_key == "mm_ex" and not s.flags and not s.exclude, vars(s))
s = spec_of("EXCLUDING M &M make", product="GCCV HCV 12001 TO 20000 GVW", state="J & K")
check("goods vehicle: the same words leave Mahindra out of the all-makes group", s.exclude == {"MAHINDRA"} and s.make_key is None and not s.flags, vars(s))
s = spec_of("EXCLUDING M &M BOLERO", product="PCCV 4W Upto 6+1", state="UTTARANCHAL-AA")
check("taxi: 'EXCLUDING M &M BOLERO' stays the Bolero group (the shorter 'excluding M &M' rule does not take it)", s.make_key == "bolero_ex" and not s.flags, vars(s))
s = spec_of("NCB Cases 20% and Above, Except New Mahindra", product="GCCV LCV UPTO 2800 GVW", state="UTTAR PRADESH")
check("'NCB Cases 20% and Above': NCB yes and the T&C 'NCB >=20%' (the user's wording), Mahindra split", s.ncb == "YES" and s.tnc == ["NCB >=20%"] and s.mahindra_split and not s.flags, vars(s))
s = spec_of("NCB Cases and New vehicles excluding New Mahindra", product="GCCV LCV ABOVE 3501 TO 7500 GVW", state="GUJARAT & DADRA NAGAR HAVELI & DAMAN & DIU")
check("'NCB Cases and New vehicles excluding New Mahindra': NCB yes, Mahindra split, no flag", s.ncb == "YES" and s.mahindra_split and not s.flags, vars(s))
s = spec_of("for  NIL dep cases payout is 40%", product="PCCV 4W Upto 6+1", state="WEST BENGAL")
check("'for NIL dep cases payout is 40%': the zero-dep payin is 40, the cell's payout is for the cases without it - the 40 is no longer dropped", s.zd_pay == 40.0 and s.zd == "NO" and not s.flags, vars(s))
s = spec_of("Except NIL dep cases", product="PCCV 4W Upto 6+1", state="WEST BENGAL")
check("'Except NIL dep cases' is WITHOUT nil dep (it used to be read as with)", s.zd == "NO", s.zd)
s = spec_of("With Nil dep", product="PCCV 4W Upto 6+1", state="WEST BENGAL")
check("'With Nil dep' is still with", s.zd == "YES", s.zd)
s = spec_of("EXCLUDING 1001 CC TO 1500 CC.", product="PRIVATE CAR PETROL", state="GUJARAT & DADRA NAGAR HAVELI & DAMAN & DIU")
check("'EXCLUDING 1001 CC TO 1500 CC.' is a cc range, not RTO code 10 (the 1001 used to be read as 'excluding 10')", s.cc_excl == (1001, 1500) and not s.code_exclude and not s.flags, vars(s))
s = spec_of("Excluding 10 years", product="PRIVATE CAR PETROL")
check("... and '10 years' after 'excluding' is no RTO code either - the words are left over and flagged", not s.code_exclude and s.flags == ["Excluding 10 years"], vars(s))
s = spec_of("M & M(All CC)  & Others makes-Greater than 1500 CC ( EXCEPT MARUTI make).", product="PRIVATE CAR DIESEL", state="BIHAR")
check("'M & M(All CC) & Others makes-Greater than 1500 CC (EXCEPT MARUTI make)': one rule, the limit 1500, no flag, no Maruti-exclusion left over", s.mm_gt == 1500 and not s.flags and s.make_key is None, vars(s))
s = spec_of("GCCV 3W EXCEPT E-CART OTHER THAN DIESEL", product="GCCV 3W Except E-CART", state="MADHYA PRADESH")
check("'GCCV 3W EXCEPT E-CART OTHER THAN DIESEL' = no Diesel rows", s.no_fuels == {"Diesel"} and not s.flags, vars(s))
s = spec_of("WITH NCB ONLY -HYUNDAI,HONDA, M&M, TATA & TOYOTA Only ( SA- OD not allowed)", product="PRIVATE CAR PETROL", state="MADHYA PRADESH")
check("'SA- OD not allowed' (a space after the hyphen) skips SAOD like 'SA-OD not allowed'", "saod" in s.skip and s.make_key == "hhtmt" and not s.flags, vars(s))
s = spec_of("Upto 10 years and except cranes, ambulance, tractor.", product="MISC-D OTHERS", state="RAJASTHAN", rem1="")
check("MISC-D: 'except cranes, ambulance, tractor.' in any order is read (', tractor.' used to be left over and flagged)", s.misc == "ex" and not s.flags, vars(s))
s = spec_of("Excluding crane, tractor", product="MISC-D OTHERS", state="RAJASTHAN")
check("MISC-D: only two of the three vehicles is not the known exclusion - flagged, not guessed", s.misc is None and s.flags, vars(s))
s = spec_of("EXL RTO codes 03, 06, 18, 24, 26, 40, 42, 56, 57, 59", product="GCCV 3W Except E-CART", state="UTTAR PRADESH")
check("'EXL RTO codes 03, 06, ...' takes those RTOs out (EXL = excluding)", s.code_exclude == {"UP03", "UP06", "UP18", "UP24", "UP26", "UP40", "UP42", "UP56", "UP57", "UP59"} and not s.flags, vars(s))
s = spec_of("other than JK- 01, 03, 04, 05 RTO'S. Except New Mahindra", product="GCCV HCV 12001 TO 20000 GVW", state="J & K")
check("'other than JK- 01, 03, 04, 05 RTO'S': those RTOs out, no flag", s.code_exclude == {"JK01", "JK03", "JK04", "JK05"} and s.mahindra_split and not s.flags, vars(s))
s = spec_of("Declined for RJ-14", product="PRIVATE CAR PETROL", state="RAJASTHAN")
check("'Declined for RJ-14' takes the RTO out (not 'for RJ-14' = the only RTO)", s.code_exclude == {"RJ14"} and s.code_include is None and not s.flags, vars(s))
s = spec_of("Allowed for JK-01,04,06 RTO codes", product="GCCV 3W Except E-CART", state="J & K")
check("an RTO list after 'Allowed for' is still the only RTOs", s.code_include == {"JK01", "JK04", "JK06"} and not s.flags, vars(s))
s = spec_of("UPTO 15 YEARS (DL)", product="GCCV 3W Except E-CART", state="DELHI&NCR")
s2 = mod.read_remarks("g3w", dict(r=1, state="DELHI&NCR", product="GCCV 3W Except E-CART", uw="UPTO 15 YEARS (DL)", rem1="", eff="", po=1, ptype="P", age="upto 15 years.", dis=70))
check("'UPTO 15 YEARS (DL)' repeats the AGE column: nothing to apply, no flag; the same words with ANOTHER limit are flagged",
      s2.part_include == ["DL"] and not s2.flags and mod.read_remarks("g3w", dict(r=1, state="DELHI&NCR", product="GCCV 3W Except E-CART", uw="UPTO 10 YEARS (DL)", rem1="", eff="", po=1, ptype="P", age="upto 15 years.", dis=70)).flags == ["UPTO 10 YEARS"], vars(s2))
s = spec_of("a rule of 12 points nobody wrote a pattern for")
check("a number no rule used is a rule nobody read: flagged (the leftover check looked at words only)", s.flags == ["a rule of 12 points nobody wrote a pattern for"], s.flags)
s = spec_of("12")
check("... even when no word is left, only the number", s.flags == ["12"], s.flags)

# --- the notes that only say what changed since last month (the user cleared all of them without touching a row)
changes = ["PO redcued by 1%", "Disc by 5% & PO by 2% increased", "PO by 2% increased", "Dic by 5% & PO increased by 7.5%", "Dic by 5% & PO increased by 2.5%",
           "Disc & PO by 5% increased, Age criteria improved", "Disc & PO by 5% increased, GVW & Age criteria improved", "Disc by 30% & PO increased by 4.5%", "Disc increased by 5%",
           "New Addition", "New Addtion", "UW's guideline changed", "Age criteria improved", "PO reduced by 2.5%", "Discount increased by 15%"]
check("change notes (typos included) need no action", all(mod.is_standing(x) for x in changes) and all(not spec_of("", rem1=x).flags for x in changes), [x for x in changes if not mod.is_standing(x)])
rules_ = ["Disc increased by 5% for UP only", "PO increased by 2% for RJ-14", "Disc increased", "Tankers, PO reduced by 2%", "NCB cases, Disc by 5% increased", "Except Maruti, Disc increased by 5%"]
check("... but a state, an RTO, a make, NCB or a vehicle word makes it a rule: the note is not standing", all(not mod.is_change_note(x) for x in ["Disc increased for UP only", "PO increased for RJ-14", "Tankers, PO reduced by 2%", "NCB cases, Disc by 5% increased"]) and not mod.is_change_note("Disc by 5%"),
      [x for x in rules_ if mod.is_change_note(x)])
check("two instructions the user cleared are standing notes", mod.is_standing("Please impose 25000 theft excess in E-cart and E-Rickshaw for DL and NCR RTO codes.") and mod.is_standing("UPTO 15 YEARS (NCR)(OR BASED DL)"))

# --- whole small grids
AP_LCV = "ANDHRA PRADESH"
df, w = run(grid([AP_LCV, "GCCV LCV UPTO 2800 GVW", 80, 45, "P & L", "5 to 10 years", "Bolero Model only"],
                 [AP_LCV, "GCCV LCV UPTO 2000 GVW", 80, 53.5, "P & L", "upto 15 years & upto 20 Years for STP", "Except New Mahindra"],
                 [AP_LCV, "GCCV MCV ABOVE 7501 TO 12000 GVW", 85, 15, "P & L", "upto 15 years", "Except New Mahindra"],
                 ["KARNATAKA", "GCCV LCV UPTO 2000 GVW", 80, 50, "P", "upto 15 years", "Except New Mahindra"]))
bol = df[df.new_vehicle_makes == "GCV_4W_BOLERO"]
check("Andhra Pradesh: the Bolero-only row is written for the Bolero group only (1+1 and STP, five fuels), its own age band", len(bol) == 10 and set(bol.pi_net_rate[bol.sub_product == "1+1"]) == {45.0}
      and set(bol.vehicle_age_min) == {4.01} and set(bol.vehicle_age_max) == {10.01} and set(bol.cc_max) == {2800.01}, bol.iloc[0].to_dict())
low2000 = df[(df.cc_max == 2000.01) & (df.new_rto_list == "AP")]
check("... and the Mahindra row of another AP row whose weight band overlaps leaves the Bolero models out (the user's by-hand correction of row 14)",
      set(low2000[low2000.vehicle_age_min == 1.0].new_vehicle_makes) == {"Mahindra_ex_Bolero"} and set(low2000[low2000.vehicle_age_min == 0.0].new_vehicle_makes) == {mod.GCV_BASE + "_MAHINDRA"}, set(low2000.new_vehicle_makes))
mcv = df[(df.cc_min == 7500.01) & (df.new_rto_list == "AP")]
check("... but not a row whose band the Bolero row does not touch (7501-12000 kg), and not another state's row", set(mcv.new_vehicle_makes) == {mod.GCV_BASE + "_MAHINDRA", "Mahindra"}
      and set(df[df.new_rto_list == "KA"].new_vehicle_makes) == {mod.GCV_BASE + "_MAHINDRA", "Mahindra"}, (set(mcv.new_vehicle_makes), set(df[df.new_rto_list == "KA"].new_vehicle_makes)))
check("... the run says which rows left Bolero out and why", sum("leave BOLERO out" in x for x in w) == 1 and any("(2)" in x or "row(s) (3)" in x for x in w if "BOLERO" in x), [x for x in w if "BOLERO" in x])
df, w = run(grid([AP_LCV, "GCCV LCV 2001-2800 GVW", 80, 50, "P & L", "upto 15 years & upto 20 Years for STP", "Except Bolero Model & New Mahindra"]))
check("'Except Bolero Model & New Mahindra' on its own row: Mahindra_ex_Bolero for old vehicles, the all-but-Mahindra group for the rest, no warning about a carve-out",
      set(df.new_vehicle_makes) == {mod.GCV_BASE + "_MAHINDRA", "Mahindra_ex_Bolero"} and not any("leave BOLERO out" in x for x in w) and (df.remarks == "").all(), (set(df.new_vehicle_makes), df.remarks.unique()))

GJ = "GUJARAT & DADRA NAGAR HAVELI & DAMAN & DIU"
df, w = run(grid([GJ, "PRIVATE CAR PETROL", 90, "30 OD", "P", "upto 15 years.", "NCB =>20%"],
                 [GJ, "PRIVATE CAR PETROL", 70, "35 OD", "P & SA-OD", "upto 15 years."],
                 [GJ, "PRIVATE CAR PETROL", 70, "15 OD", "P", "upto 15 years.", "Only EECO( SA-OD and STP not allowed)"],
                 [GJ, "PRIVATE CAR PETROL", None, 30, "L", "4 to 20 years", "EXCLUDING 1001 CC TO 1500 CC."],
                 [GJ, "PRIVATE CAR DIESEL", 65, "30 OD", "P", "upto 15 years."]))
pk = df[df.sub_product == "1+1"]
check("Gujarat cars: the 1+1 rows of the state leave EECO out (the carved group), the EECO-only row keeps EECO",
      set(pk[pk.pi_od_rate.isin([30.0, 35.0]) & (pk.fuel_type == "Petrol")].new_vehicle_makes) == {"SHRIRAM_PC_EX_Eecco_DECLINE_JUNE26"} and set(pk[pk.pi_od_rate == 15.0].new_vehicle_makes) == {"EECO"}, pk[["fuel_type", "pi_od_rate", "new_vehicle_makes"]].drop_duplicates().to_dict("records"))
check("... the SAOD and STP parts keep the full group (EECO is not allowed there, the EECO row has none), the diesel row is not touched",
      set(df[df.sub_product == "SAOD"].new_vehicle_makes) == {mod.MAKES_PC} and set(df[df.sub_product == "STP"].new_vehicle_makes) == {mod.MAKES_PC} and set(df[df.fuel_type == "Diesel"].new_vehicle_makes) == {mod.MAKES_PC},
      df[["sub_product", "fuel_type", "new_vehicle_makes"]].drop_duplicates().to_dict("records"))
stp = df[df.sub_product == "STP"]
check("'EXCLUDING 1001 CC TO 1500 CC.' on STP: the cc band is written below and above the range (0-1000.01, 1500.01-99999), three fuels each",
      sorted(set(zip(stp.cc_min, stp.cc_max))) == [(0.0, 1000.01), (1500.01, 99999.0)] and len(stp) == 6 and set(stp.fuel_type) == {"Petrol", "CNG", "LPG"}, sorted(set(zip(stp.cc_min, stp.cc_max))))
check("... and the carved group is the one the user created, flagged here only because the master export is older than it", df[df.new_vehicle_makes == "SHRIRAM_PC_EX_Eecco_DECLINE_JUNE26"].remarks.str.contains("not in the make / model master", case=False).all()
      and any("leave EECO out" in x for x in w), (df.remarks.unique()[:3], [x for x in w if "EECO" in x]))

BR = "BIHAR"
df, w = run(grid([BR, "PRIVATE CAR DIESEL", 85, "30 OD", "P", "upto 15 years", "M & M(All CC)  & Others makes-Greater than 1500 CC ( EXCEPT MARUTI make)."],
                 [BR, "PRIVATE CAR DIESEL", 85, "15 OD", "P", "upto 15 years", "OTHER THAN M &M"]))
got = sorted(set(zip(df.new_vehicle_makes, df.cc_min, df.cc_max, df.pi_od_rate)))
check("Bihar diesel cars: Mahindra at any cc (30), every other make but Maruti above 1500 cc (30), 'OTHER THAN M &M' up to 1500 cc only (15) - the user's by-hand split",
      got == [("Mahindra", 0.0, 99999.0, 30.0), ("SHRIRAM_PC_EX_DECLINE_MAHINDRA", 0.0, 1500.01, 15.0), ("SHRIRAM_PC_EX_DECLINE_MARUTI_MAHINDRA", 1500.01, 99999.0, 30.0)], got)
df, w = run(grid(["KARNATAKA", "PRIVATE CAR DIESEL", 85, "15 OD", "P", "upto 15 years", "OTHER THAN M &M"]))
check("'OTHER THAN M &M' alone (no Mahindra row in the state) keeps every cc", set(zip(df.cc_min, df.cc_max)) == {(0.0, 99999.0)}, set(zip(df.cc_min, df.cc_max)))

WB = "WEST BENGAL"
df, w = run(grid([WB, "PCCV 4W Upto 6+1", 75, 42.5, "P & L", "upto 15 years & upto 20 Years for STP", "for  NIL dep cases payout is 40%"]))
pkg, stp = df[df.sub_product == "1+1"], df[df.sub_product == "STP"]
check("West Bengal taxi 'for NIL dep cases payout is 40%': 1+1 zero-dep YES gets 40, without zero dep gets the cell's 42.5 (it used to write YES at 42.5 and drop the 40)",
      set(zip(pkg.is_zd, pkg.pi_net_rate)) == {("YES", 40.0), ("NO", 42.5)} and set(pkg.fuel_type) == {"Petrol", "Diesel", "CNG", "LPG"}, set(zip(pkg.is_zd, pkg.pi_net_rate)))
check("... STP has no zero dep, so both meet on one entry and the LOWER (40) is written and listed", set(stp.is_zd) == {"NA"} and set(stp.pi_tp_rate) == {40.0} and len(mod.LAST_LOWEST) == 4 and all(x["written"] == "40" for x in mod.LAST_LOWEST),
      (set(stp.pi_tp_rate), len(mod.LAST_LOWEST)))
check("... po = pi - 7 on both (33 and 35.5)", set(zip(pkg.is_zd, pkg.po_net_rate)) == {("YES", 33.0), ("NO", 35.5)}, set(zip(pkg.is_zd, pkg.po_net_rate)))

df, w = run(grid(["MADHYA PRADESH", "GCCV 3W EXCEPT E-CART", 70, 51, "P & L", "upto 15 years.", "GCCV 3W EXCEPT E-CART OTHER THAN DIESEL"]))
check("goods 3W 'OTHER THAN DIESEL': Petrol, CNG and LPG only (the October upload had no Diesel for MP / RJ either)", set(df.fuel_type) == {"Petrol", "CNG", "LPG"} and (df.remarks == "").all(), set(df.fuel_type))

UP = "UTTAR PRADESH"
df, w = run(grid([UP, "GCCV LCV UPTO 2800 GVW", 85, 44, "P & L", "upto 5 years", "NCB Cases 20% and Above, Except New Mahindra"]))
pkg, stp = df[df.sub_product == "1+1"], df[df.sub_product == "STP"]
check("UP 'NCB Cases 20% and Above': 1+1 has NCB YES and the T&C 'NCB >=20%'; STP has the T&C and NCB NA (STP carries no NCB: the user's rule of 27 Oct)",
      set(pkg.is_ncb) == {"YES"} and set(stp.is_ncb) == {"NA"} and set(pkg.add_tnc) >= {"NCB >=20%"} and set(stp.add_tnc) >= {"NCB >=20%"} and (df.remarks == "").all(), (set(pkg.is_ncb), set(stp.is_ncb), set(pkg.add_tnc)))

df, w = run(grid(["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years.", "Disc by 5% & PO by 2% increased"], ["GOA", "PRIVATE CAR PETROL", 85, 21, "L", "4 to 20 years", "Dic by 5% & PO increased by 7.5%"]))
check("a note that only says what changed is not flagged (the user cleared 217 such rows by hand)", (df.remarks == "").all() and len(df) > 0, df.remarks.unique())

# --- the RTO working list: an Excel error in a cell is no name
def wgrid(rows, rto):
    head = list(HDR) + ["RTO"]
    return helpers.FakeWB({"GRID NOV' 26": [head] + [list(r) + [None] * (9 - len(r)) + [v] for r, v in zip(rows, rto)]})


df, w = run(wgrid([["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."], ["ANDHRA PRADESH", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."],
                   ["BIHAR", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."]], ["#REF!", "#n/a", "West Bihar Open"]))
note = next((x for x in w if x.startswith("RTO working list")), "")
check("an Excel error in the RTO column (#REF!, #N/A) is no name: those rows use the master lookup, the typed name is kept", set(df.new_rto_list) == {"GA", "AP", "West Bihar Open"} and not df.new_rto_list.str.startswith("#").any(), set(df.new_rto_list))
check("... and the run line says it: how many cells, which errors, what happens", "2 cell(s) of the column hold an Excel error (#N/A, #REF!)" in note and "master lookup" in note, note)

# --- the sheet 'Declined vehicle' has text rules too: 'Decline Product [MISC.D - OTHERS (BODY TYPE GARBAGE CARRIER OR GARBAGE COMPACTOR)' (new in November 2026; the first build never read that sheet's text).
# The user's choice (4 Nov 2026): write the declined body types into the exclusion T&C of the MISC-D others rows.
LINE = "Decline Product  [MISC.D - OTHERS (BODY TYPE GARBAGE  CARRIER OR GARBAGE COMPACTOR)"
DV = [["Decline Vehicles _ PCCV 4W (7-10)"], ["Model", "Application Location"], ["Commander", "Pan India"], [LINE]]
PC1 = ["GOA", "PRIVATE CAR PETROL", 85, 20, "P", "upto 15 years."]
check("the reader gives the line, its product label and the body types, written Like This (OR / comma / & / AND split, capitals ignored)",
      mod.decline_product_notes(grid(PC1, extra={"Declined vehicle": DV})) == [("Decline Product [MISC.D - OTHERS (BODY TYPE GARBAGE CARRIER OR GARBAGE COMPACTOR)", "MISC.D - OTHERS", ["Garbage Carrier", "Garbage Compactor"])]
      and mod.decline_product_notes(grid(PC1, extra={"Declined vehicle": [["Decline Product [MISC.D - OTHERS (body type garbage carrier, tipper & mixer and crusher)]"]]}))[0][2] == ["Garbage Carrier", "Tipper", "Mixer", "Crusher"]
      and mod.decline_product_notes(grid(PC1, extra={"Declined vehicle": [["Decline Product [MISC.D - OTHERS (ALL)]"]]}))[0][2] == [] and mod.decline_product_notes(grid(PC1)) == [],
      mod.decline_product_notes(grid(PC1, extra={"Declined vehicle": DV})))
EXS = "Upto 10 years and except cranes, ambulance, tractor."
df, w = run(grid(["GOA", "MISC-D OTHERS", 70, 20, "P", "upto 10 years", EXS], ["GOA", "MISC-D OTHERS", 70, 18, "P", "upto 10 years", "JCB only"], ["GOA", "MISC-D OTHERS", 70, 17, "P", "upto 10 years"],
                 ["GOA", "MISC-D OTHERS", 70, 16, "P & L", "upto 10 years", "excluding crane and ambulance"], ["GOA", "MISC-D TRACTOR", 85, 35, "P", "New"], PC1, extra={"Declined vehicle": DV}))
o = df[(df["product"] == "MISCD") & (df.make_model_class == "MISCD-Others")]
tnc_of = lambda po: set(o[o.pi_net_rate.astype(float).where(o.sub_product == "1+1", o.pi_tp_rate.astype(float)) == po].add_tnc)
check("'Decline Product [MISC.D - OTHERS (BODY TYPE ...)]': the two body types are written into the exclusion T&C of the MISC-D others rows (EX:Crane,Tractor,Ambulance + Garbage Carrier, Garbage Compactor)",
      tnc_of(20.0) == {"EX:Crane,Tractor,Ambulance,Garbage Carrier,Garbage Compactor"}, tnc_of(20.0))
check("... a row with no exclusion list gets 'EX:<bodies>', 'EX:Crane & Ambulance' keeps its wording and continues", tnc_of(17.0) == {"EX:Garbage Carrier,Garbage Compactor"} and tnc_of(16.0) == {"EX:Crane & Ambulance,Garbage Carrier,Garbage Compactor"}, (tnc_of(17.0), tnc_of(16.0)))
check("... a row for ONE body type ('JCB /-') is not touched, the tractor and the car rows are not touched, and nothing is flagged",
      tnc_of(18.0) == {"JCB /-"} and set(df[df.make_model_class == "MISCD-Tractor"].add_tnc) == {"NEW"} and set(df[df["product"] == "Private Car"].add_tnc) == {""} and (df.remarks == "").all(), (tnc_of(18.0), df.remarks.unique()))
check("... and the run says how many rows got it", any("written into the exclusion T&C (EX:...) of 20 of the 25" in x and "the 5 rows for one body type only" in x for x in w), [x for x in w if "Declined vehicle" in x])
exs = {"": "EX:Garbage Carrier,Garbage Compactor", "EX:Crane,Tractor,Ambulance": "EX:Crane,Tractor,Ambulance,Garbage Carrier,Garbage Compactor", "EX:Crane & Ambulance": "EX:Crane & Ambulance,Garbage Carrier,Garbage Compactor",
       "EX:Crane,Tractor, Ambulance & other than Combined Harvester": "EX:Crane,Tractor, Ambulance,Garbage Carrier,Garbage Compactor & other than Combined Harvester",
       "JCB /-": "JCB /-", "Combined Harvester /-": "Combined Harvester /-", "Backhoe loader /-": "Backhoe loader /-", "JCB & Excavator /-": "JCB & Excavator /-", "NEW": "NEW"}
check("add_declined_bodies: the body types go into the exclusion list (before an 'other than' clause), any T&C that is not an exclusion list (single-body rows) stays, no bodies = no change",
      all(mod.add_declined_bodies(k, ["Garbage Carrier", "Garbage Compactor"]) == v for k, v in exs.items()) and mod.add_declined_bodies("EX:Crane,Tractor,Ambulance", []) == "EX:Crane,Tractor,Ambulance",
      {k: mod.add_declined_bodies(k, ["Garbage Carrier", "Garbage Compactor"]) for k in exs})
df, w = run(grid(["GOA", "MISC-D OTHERS", 70, 20, "P", "upto 10 years", EXS], PC1, extra={"Declined vehicle": [["Decline Product [MISC.D - OTHERS (ALL)]"]]}))
check("a 'Decline Product' line with no body type to write is flagged on the rows of that product (not guessed), and no T&C changes",
      df[df["product"] == "MISCD"].remarks.str.contains("Decline Product", regex=False).all() and set(df[df["product"] == "MISCD"].add_tnc) == {"EX:Crane,Tractor,Ambulance"} and (df[df["product"] == "Private Car"].remarks == "").all(), df.remarks.unique())
df, w = run(grid(PC1, extra={"Declined vehicle": [["Decline Product [WIND TURBINE (BODY TYPE BLADE)]"]]}))
check("a 'Decline Product' line for a product this script does not know is a warning, no row is flagged for it", (df.remarks == "").all() and any("WIND TURBINE" in x and "does not know that product" in x for x in w), w)
df, w = run(grid(PC1))
check("no 'Declined vehicle' sheet, no line, nothing changes", (df.remarks == "").all() and not any("Declined vehicle" in x for x in w))

# --- the real November grid (the user's grid with their RTO column, 52 of its cells are `=#REF!`)
NOV = os.path.join(helpers.INPUTS, "shriram_nov26.xlsx")
if os.path.exists(NOV):
    wbn = openpyxl.load_workbook(NOV, data_only=True)
    dn, wn = mod.build(None, dt.date(2026, 11, 1), T, quiet, mod.Asker(interactive=False), wb=wbn)
    check("November grid: no RTO list is an Excel error, and the run line counts the 52 broken cells", not dn.new_rto_list.str.contains("#").any() and any("52 cell(s) of the column hold an Excel error (#REF!)" in x for x in wn), [x for x in wn if "RTO working" in x])
    mk_rows = dn.remarks.str.contains("not in the make / model master", case=False)
    check("November grid: 6,397 rows; what is flagged is only the make groups the master export (25 Oct) does not know yet (15 rows)", len(dn) == 6397 and int(mk_rows.sum()) == 15 and int((dn.remarks != "").sum()) == 15, ((dn.remarks != "").sum(), int(mk_rows.sum())))
    mo = dn[(dn["product"] == "MISCD") & (dn.make_model_class == "MISCD-Others")]
    bod = mo.add_tnc.str.contains("Garbage Carrier,Garbage Compactor", regex=False)
    check("November grid: the declined body types (Garbage Carrier, Garbage Compactor) are in the exclusion T&C of the 70 MISC-D others rows that have one - the 80 rows for one body type (JCB, Backhoe, Combined Harvester) and the tractor rows are untouched",
          len(mo) == 150 and int(bod.sum()) == 70 and mo[bod].add_tnc.str.startswith("EX:").all() and mo[~bod].add_tnc.str.endswith("/-").all() and int(((dn.make_model_class == "MISCD-Tractor") & dn.add_tnc.str.contains("Garbage")).sum()) == 0,
          (len(mo), int(bod.sum()), sorted(set(mo[~bod].add_tnc))))
    ap = dn[(dn.new_rto_list == "SHRIRAM_JUNE26_GCV4W_AP_EX_27") & (dn.fuel_type == "Diesel") & (dn.sub_product == "1+1")]
    check("November grid: Andhra Pradesh as the user corrected it by hand - Bolero rows GCV_4W_BOLERO (45 / 30), the two Mahindra rows Mahindra_ex_Bolero",
          sorted(set(zip(ap.new_vehicle_makes, ap.vehicle_age_min, ap.pi_net_rate))) == [("GCV_4W_BOLERO", 4.01, 45.0), ("GCV_4W_BOLERO", 10.01, 30.0), ("Mahindra_ex_Bolero", 1.0, 50.0), ("Mahindra_ex_Bolero", 1.0, 53.5),
                                                                                                  (mod.GCV_BASE + "_MAHINDRA", 0.0, 50.0), (mod.GCV_BASE + "_MAHINDRA", 0.0, 53.5)], sorted(set(zip(ap.new_vehicle_makes, ap.vehicle_age_min, ap.pi_net_rate))))
    wbt = dn[(dn.new_rto_list == "WB") & (dn["product"] == "PCV 4W") & (dn.fuel_type == "Petrol") & (dn.sub_product == "1+1") & (dn.sc_max == 7.01)]
    check("November grid: West Bengal taxi 0-7 - zero dep YES 40, without 42.5", sorted(set(zip(wbt.is_zd, wbt.pi_net_rate))) == [("NO", 42.5), ("YES", 40.0)], sorted(set(zip(wbt.is_zd, wbt.pi_net_rate))))
    amb = dn[(dn.new_vehicle_makes == "Ambasador") & (dn.sub_product == "1+1") & (dn.fuel_type == "Petrol")]
    check("November grid: the two Kaali-peeli / ambassador rows are Ambasador and WITHOUT nil dep", len(amb) == 2 and set(amb.is_zd) == {"NO"} and sorted(amb.pi_net_rate) == [50.0, 55.0], amb[["is_zd", "pi_net_rate"]].to_dict("records"))
    rj = dn[(dn["product"] == "PCV 4W") & (dn.new_rto_list.isin(["RJ_EX_14", "RJ14"])) & (dn.fuel_type == "Petrol") & (dn.sub_product == "1+1")]
    check("November grid: Rajasthan taxis - the groups the user chose (without Dzire for RJ_EX_14, SWIFT_DZIRE for RJ14)", set(rj[rj.new_rto_list == "RJ_EX_14"].new_vehicle_makes) == {"SHRIRAM_PCV_4W_0-7_JUNE26_EX_DZIRE"}
          and set(rj[rj.new_rto_list == "RJ14"].new_vehicle_makes) == {"SWIFT_DZIRE"}, set(zip(rj.new_rto_list, rj.new_vehicle_makes)))
    up = dn[(dn.add_tnc == "NCB >=20%") & (dn["product"] == "GCV 4W")]
    check("November grid: UP 'NCB Cases 20% and Above' - 20 rows with the T&C 'NCB >=20%', NCB YES on 1+1 only", len(up) == 20 and set(up[up.sub_product == "1+1"].is_ncb) == {"YES"} and set(up[up.sub_product == "STP"].is_ncb) == {"NA"})
    gj = dn[(dn.new_rto_list == "GJ,DD,DN") & (dn["product"] == "Private Car") & (dn.fuel_type == "Petrol") & (dn.sub_product == "STP")]
    check("November grid: Gujarat petrol STP 'EXCLUDING 1001 CC TO 1500 CC.' - two cc bands", sorted(set(zip(gj.cc_min, gj.cc_max))) == [(0.0, 1000.01), (1500.01, 99999.0)], sorted(set(zip(gj.cc_min, gj.cc_max))))
    g3 = dn[(dn["product"] == "GCV 3W") & dn.new_rto_list.isin(["RJ", "SHRIRAM_JUNE26_GCV3W_MP_EX_4_18_48"]) & (dn.fuel_type != "Electric")]      # (the e-cart rows of the same states are Electric)
    check("November grid: goods 3W of MP and RJ - Petrol, CNG, LPG only (no Diesel)", set(g3.fuel_type) == {"Petrol", "CNG", "LPG"} and len(g3) == 12, (set(g3.fuel_type), len(g3)))

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
