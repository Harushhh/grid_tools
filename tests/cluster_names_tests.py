"""Checks of the RTO name rule of 30 Oct 2026: new_rto_list carries ONLY the cluster name as the grid writes it ("West UP Open", "DL_Delhi", "GJ_ABS"); the portal's group
name (insurer + month + segment prefix) and single RTO codes are not written, and the RTO names are not checked against the RTO master (the person maps and groups the RTOs
himself).  GRID_RTO_NAMES=group brings the earlier behaviour back (the other suites run in it).
Run:  .venv\\Scripts\\python.exe tests\\cluster_names_tests.py"""
import os
import sys

os.environ["GRID_RTO_NAMES"] = "cluster"                  # before any script is loaded
import pandas as pd  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


SCRIPTS = ["liberty_grid_to_config.py", "magma_grid_to_config.py", "rsa_gcv4w_comp_to_config.py", "rsa_pvtcar_to_config.py", "united_commission_to_config.py",
           "shriram_broker_grid_to_config.py", "shriram_2w_to_config.py", "godigit_agency_to_config.py", "sbi_grid_to_config.py", "icici_grid_to_config.py", "indusind_grid_to_config.py"]
mods = {s: helpers.load_script(s) for s in SCRIPTS}
check("every script is in cluster mode by default", all(m.RTO_NAMES == "cluster" for m in mods.values()))

# ---- the shared block: apply_cluster_names
M = mods["liberty_grid_to_config.py"]
M.CLUSTER_OF.clear()
M.rto_cluster("GO-DIGIT_OCT26_HCV_West_UP_Open", "West UP Open")
M.rto_cluster("GO-DIGIT_OCT26_2W_1+1_TP_GJ_ABS", "GJ_ABS")
M.rto_cluster("GO-DIGIT_OCT26_HCV_Delhi_NCR", "Delhi\n  NCR,  North")
frame = pd.DataFrame({"new_rto_list": ["GO-DIGIT_OCT26_HCV_West_UP_Open", "GO-DIGIT_OCT26_2W_1+1_TP_GJ_ABS", "GO-DIGIT_OCT26_HCV_Delhi_NCR", "ALLINDIA", "GJ,DD,DN", "MH"],
                      "product": ["a"] * 6})
out = M.apply_cluster_names(frame)
check("registered group names become the grid's cluster name", list(out["new_rto_list"][:2]) == ["West UP Open", "GJ_ABS"], list(out["new_rto_list"]))
check("a cluster name keeps the grid's spelling; line breaks and runs of spaces collapse, a comma becomes '_' (new_rto_list is itself a comma list)", out["new_rto_list"][2] == "Delhi NCR_ North", out["new_rto_list"][2])
check("ALLINDIA and state codes are left alone", list(out["new_rto_list"][3:]) == ["ALLINDIA", "GJ,DD,DN", "MH"], list(out["new_rto_list"][3:]))
check("the frame the script keeps is not changed", frame["new_rto_list"][0] == "GO-DIGIT_OCT26_HCV_West_UP_Open")
M.RTO_NAMES = "group"
check("in group mode nothing is changed", M.apply_cluster_names(frame)["new_rto_list"].tolist() == frame["new_rto_list"].tolist())
M.RTO_NAMES = "cluster"

# ---- a name nobody registered: the insurer prefix of the script is stripped (Shriram, United)
S = mods["shriram_broker_grid_to_config.py"]
S.CLUSTER_OF.clear()
sf = pd.DataFrame({"new_rto_list": ["SHRIRAM_JUNE26_Mumbai", "SHRIRAM_JUNE26_Mumbai,MH", "DL_Delhi", "GJ,DD,DN"], "product": ["a"] * 4})
check("Shriram: the insurer / month prefix of a master cluster group is dropped, other names stay", S.apply_cluster_names(sf)["new_rto_list"].tolist() == ["Mumbai", "Mumbai,MH", "DL_Delhi", "GJ,DD,DN"],
      S.apply_cluster_names(sf)["new_rto_list"].tolist())
U = mods["united_commission_to_config.py"]
U.CLUSTER_OF.clear()
uf = pd.DataFrame({"new_rto_list": ["UNITED_JULY26_Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP", "UNITED_CAR_RTO_Annexure3", "UP,TN", "ALLINDIA"], "product": ["a"] * 4})
check("United: the insurer / month prefix of an annexure group is dropped; state codes stay", U.apply_cluster_names(uf)["new_rto_list"].tolist() == ["Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP", "CAR_RTO_Annexure3", "UP,TN", "ALLINDIA"],
      U.apply_cluster_names(uf)["new_rto_list"].tolist())

# ---- two portal groups with one cluster name and otherwise the same key: say so
M.CLUSTER_OF.clear()
M.RUN_WARNINGS.clear()
cols = {c: ["x", "x"] for c in M.KEY_COLUMNS if c != "new_rto_list"}
cols["new_rto_list"] = ["A_HCV_Mumbai", "B_HCV_Mumbai"]
M.rto_cluster("A_HCV_Mumbai", "Mumbai")
M.rto_cluster("B_HCV_Mumbai", "Mumbai")
M.apply_cluster_names(pd.DataFrame(cols))
check("rows that now share a key are reported as a warning", any("share one key" in w for w in M.RUN_WARNINGS), M.RUN_WARNINGS)

# ---- the name builders of the scripts
mg = mods["magma_grid_to_config.py"]
mg.CLUSTER_OF.clear()
check("Magma: 'DL NCR' -> the group name inside, the bare cluster in the CSV", mg.rto_name("DL NCR") == "MAGMA_APR26_DL_NCR" and mg.apply_cluster_names(pd.DataFrame({"new_rto_list": ["MAGMA_APR26_DL_NCR"]}))["new_rto_list"][0] == "DL NCR")
rg = mods["rsa_gcv4w_comp_to_config.py"]
rg.CLUSTER_OF.clear()
name = rg.rto_name("West UP Open")
check("RSA GCV: 'West UP Open' stays 'West UP Open'", rg.apply_cluster_names(pd.DataFrame({"new_rto_list": [name]}))["new_rto_list"][0] == "West UP Open", name)
rp = mods["rsa_pvtcar_to_config.py"]
rp.CLUSTER_OF.clear()
name = rp.rto_name("Baroda (Gujarat)", rp.COMP_RTO_PART)
check("RSA Private Car: bracket text dropped, the grid's name kept, not upper-cased or renamed", rp.apply_cluster_names(pd.DataFrame({"new_rto_list": [name]}))["new_rto_list"][0] == "Baroda", name)
lb = mods["liberty_grid_to_config.py"]
lb.CLUSTER_OF.clear()
lb._CTX["prefix"] = lb.RTO_PREFIX[lb.CV_LOW_SHEET]
name = lb.rto_name("ASSAM - 1 GK")
lb._CTX["prefix"] = None
check("Liberty: a geo of a sheet -> the bare geo; the Pan India rows stay ALLINDIA", lb.apply_cluster_names(pd.DataFrame({"new_rto_list": [name, "ALLINDIA"]}))["new_rto_list"].tolist() == ["ASSAM - 1 GK", "ALLINDIA"], name)

# ---- the master check no longer looks at RTO names, and still looks at makes
class Rev:
    def __init__(self):
        self.items = []

    def flag(self, text, rows=None):
        self.items.append(text)


for s in SCRIPTS:
    m = mods[s]
    masters = m.Masters({"MH": ["MH01"]}, {"TATA": ["Tata Ace"]})
    df = pd.DataFrame({"new_rto_list": ["West UP Open", "GJ_ABS"], "new_vehicle_makes": ["TATA", "NOT_A_MAKE"]})
    rev = Rev()
    n = m.check_names(df, rev, masters, log=lambda *a, **k: None)
    check(f"{s}: RTO cluster names are not flagged, a make the master lacks still is", n == 1 and len(rev.items) == 1 and "new_vehicle_makes" in rev.items[0], rev.items)
    m.RTO_NAMES = "group"
    rev = Rev()
    n = m.check_names(df, rev, masters, log=lambda *a, **k: None)
    m.RTO_NAMES = "cluster"
    check(f"{s}: GRID_RTO_NAMES=group checks the RTO names again", n == 3, (n, rev.items))

# ---- Shriram: single RTOs are not written
for s in ("shriram_broker_grid_to_config.py", "shriram_2w_to_config.py"):
    m = mods[s]
    check(f"{s}: RJ41 / 'RJ14,RJ41' are single RTOs, a state code or a cluster group is not",
          m._single_rtos("RJ41") and m._single_rtos("RJ14,RJ41") and not m._single_rtos("RJ") and not m._single_rtos("SHRIRAM_JUNE26_Mumbai") and not m._single_rtos("GJ,DD,DN") and not m._single_rtos(""))

# ---- the condition in the RTO cell (the user's rule of 30 Oct 2026: complex grids write the cluster name AND its condition in plain English, with the part, the product, the sub-product)
for s in SCRIPTS:
    m = mods[s]
    sc = m.RtoScope("MH", "the insurer declines some RTOs of this state for this product so they are left out", "MH without MH01, MH02")
    cell = m.rto_cell(sc)
    check(f"{s}: rto_cell writes name, condition and part, with no comma in the sentences",
          cell == "MH - Condition: the insurer declines some RTOs of this state for this product so they are left out - Part: MH without MH01 / MH02" and "," not in cell, cell)
    check(f"{s}: a scope without a condition stays the plain name", m.rto_cell(m.RtoScope("MH")) == "MH" and m.rto_cell("GJ,DD,DN") == "GJ,DD,DN")
    m.RTO_NAMES = "group"
    cell_g = m.rto_cell(sc)
    m.RTO_NAMES = "cluster"
    check(f"{s}: group mode never writes a condition", cell_g == "MH", cell_g)
    row = dict(product="GCV 4W", sub_product="1+1", policy_type="Renew", fuel_type="Diesel", vehicle_age_min=0, vehicle_age_max=15.0, cc_min=0.0, cc_max=99999, sc_min=0, sc_max=999,
               new_vehicle_makes="TATA,Ashok", make_model_class="NA", tariff_min=0, tariff_max=100.01, is_ncb="Yes", is_cpa="NA", is_zd="NA", add_tnc="Old")
    det = m.rto_details(row)
    check(f"{s}: rto_details lists product, sub-product, policy, fuel, age, cc, sc, makes, discount band, NCB and extra terms in plain words; 'NA' is left out; no comma",
          det == "Product: GCV 4W - Sub-product: 1+1 - Policy type: Renew - Fuel: Diesel - Vehicle age: 0 to 15 years - Engine cc: 0 to 99999 - Seating capacity: 0 to 999 - Make group: TATA / Ashok - Discount band: 0 to 100.01 - NCB: Yes - Extra terms: Old"
          and "," not in det, det)
check("plain_list: sorted, no commas", mods["liberty_grid_to_config.py"].plain_list(["MH02", "MH01", "MH02"]) == "MH01/MH02")

U = mods["united_commission_to_config.py"]
for name, expect in [
        ("UNITED_JULY26_Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP", ("only in the preferred RTOs listed in Sub Annexure 3", "only the RTOs listed in Sub Annexure 3 for UP TN MP (2 to 3.5 tonnes goods vehicles)")),
        ("UNITED_JULY26_Sub Annexure3_GCV4W_2T-3.5T_EX_UP_TN_MP", ("all the other RTOs of the same states", "UP TN MP without the RTOs listed in Sub Annexure 3 (2 to 3.5 tonnes goods vehicles)")),
        ("UNITED_RJ_EX_Annexure2_AUG26", ("all the other RTOs of the same states", "RJ without the RTOs listed in Sub Annexure 2")),
        ("UNITED_CAR_RTO_Annexure3", ("only in the RTOs of the identified cities listed in Sub Annexure 3", "only the RTOs of the identified cities listed in Sub Annexure 3")),
        ("UNITED_CAR_RTO_EX_Annexure4_JULY26", ("all the other RTOs", "all RTOs without the RTOs of the identified cities listed in Sub Annexure 4"))]:
    got = U.rto_condition(name, "GCV 4W", "1+1")
    check(f"United: the annexure group {name} is explained in plain English", got is not None and expect[0] in got[0] and got[1] == expect[1], got)
check("United: state codes and ALLINDIA have no condition", U.rto_condition("UP,TN", "GCV 4W", "1+1") is None and U.rto_condition("ALLINDIA", "Private Car", "1+1") is None)
U.CLUSTER_OF.clear()
uf = pd.DataFrame({"new_rto_list": ["UNITED_JULY26_Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP", "UP,TN", "UNITED_JULY26_Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP"], "product": ["GCV 4W", "GCV 4W", "GCV 4W"],
                   "sub_product": ["1+1", "1+1", "STP"]})
uf["cc_min"], uf["cc_max"], uf["sc_min"], uf["sc_max"] = [0, 0, 0], [99999, 99999, 99999], [0, 0, 0], [999, 999, 999]
out = U.apply_cluster_names(uf)["new_rto_list"].tolist()
check("United: the CSV cell = cluster name + condition + part + the row's product, sub-product, cc and seating capacity; state codes stay; two sub-products give two texts",
      out[0].startswith("Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP - Condition: ") and " - Product: GCV 4W - Sub-product: 1+1 - " in out[0] and out[0].endswith("Engine cc: 0 to 99999 - Seating capacity: 0 to 999")
      and out[1] == "UP,TN" and "Sub-product: STP" in out[2], out)

# Shriram: the area minus the RTOs the insurer declines -> the state is the name, the condition says which RTOs
class Declines:
    def __init__(self, gone):
        self.gone = set(gone)

    def decline(self, labels, letters, exempt=False):
        return {c for c in self.gone if c[:2] in letters}


for s in ("shriram_broker_grid_to_config.py", "shriram_2w_to_config.py"):
    m = mods[s]
    masters = m.Masters({"MH": ["MH01", "MH02", "MH03"], "RJ": ["RJ01", "RJ41"], "RJ41": ["RJ41"], "ALLINDIA": ["MH01", "MH02", "MH03", "RJ01", "RJ41"]}, {"TATA": ["x"]})
    rt = m.Rtos(masters)
    sc, missing = rt.resolve([("codes", ["MH"])], Declines({"MH02"}), frozenset())
    check(f"{s}: a state minus a declined RTO -> the state code, the condition and the exact part, nothing 'missing in the master'",
          str(sc) == "MH" and "declines some RTOs" in sc.condition and sc.part == "MH without MH02" and missing == [], (str(sc), sc.condition, sc.part, missing))
    check(f"{s}: the cell reads as one plain sentence", m.rto_cell(sc) ==
          "MH - Condition: the insurer declines some RTOs of this state for this product so they are left out - Part: MH without MH02", m.rto_cell(sc))
    sc2, _ = rt.resolve([("codes", ["MH"])], Declines(set()), frozenset())
    check(f"{s}: a whole state has no condition and stays 'MH'", m.rto_cell(sc2) == "MH" and not sc2.condition)
    sc3, _ = rt.resolve([("codes", ["RJ"])], Declines({"RJ01"}), frozenset())
    check(f"{s}: a single RTO group (RJ41) is not the name: the state is, and the part says what is left", str(sc3) == "RJ" and "RJ01" in sc3.part, (str(sc3), sc3.part))
    m.RTO_NAMES = "group"
    sc4, missing4 = rt.resolve([("codes", ["MH"])], Declines({"MH02"}), frozenset())
    cell4 = m.rto_cell(sc4)
    m.RTO_NAMES = "cluster"
    check(f"{s}: group mode still reports what the master lacks, and writes no condition", str(sc4) == "MH" and missing4 and cell4 == "MH", (str(sc4), missing4, cell4))

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
