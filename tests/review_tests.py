"""The shared ManualReview safeguard (remarks column) in all three scripts: the class itself, then how Liberty and Magma use it.
Run:  .venv\\Scripts\\python.exe tests\\review_tests.py"""
import copy
import datetime as dt
import os
import sys
import time

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

lib = helpers.load_script("liberty_grid_to_config.py")
mag = helpers.load_script("magma_grid_to_config.py")
rsa = helpers.load_script("rsa_gcv4w_comp_to_config.py")
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


# ================================================================ the class (identical in all three scripts)
for label, m in (("liberty", lib), ("magma", mag), ("rsa", rsa)):
    MR = m.ManualReview
    df = pd.DataFrame({"a": range(5), "remarks": [""] * 5})
    r = MR()
    counts, n = r.apply(df)
    check(f"[{label}] no flags: remarks untouched, nothing reported", df.remarks.tolist() == [""] * 5 and counts == [] and n == 0)

    df = pd.DataFrame({"a": range(5), "remarks": [""] * 5})
    r = MR()
    r.flag("note A, with a comma", np.array([True, True, False, False, False]))
    r.flag("note B", lambda d: (d.a >= 1).to_numpy() & (d.a <= 2))
    r.flag("note C for all")
    r.flag("note A, with a comma", np.array([True, False, False, False, False]))       # the same note again: not written twice on a row
    counts, n = r.apply(df)
    check(f"[{label}] format: 'Manual Review Required: [A] | [B] | [C]', joined with ' | ', unaffected rows empty or only C",
          df.remarks[0] == "Manual Review Required: [note A, with a comma] | [note C for all]"
          and df.remarks[1] == "Manual Review Required: [note A, with a comma] | [note B] | [note C for all]"
          and df.remarks[2] == "Manual Review Required: [note B] | [note C for all]"
          and df.remarks[4] == "Manual Review Required: [note C for all]", df.remarks.tolist())
    check(f"[{label}] counts: A 2 rows (the repeat adds none), B 2, C 5; 5 rows flagged", dict(counts) == {"note A, with a comma": 2, "note B": 2, "note C for all": 5} and n == 5, counts)

    df = pd.DataFrame({"remarks": ["old text", None, float("nan"), "", "  "]})
    r = MR()
    r.flag("X")
    r.apply(df)
    check(f"[{label}] existing remarks are kept and joined; None / NaN / blank are handled",
          df.remarks.tolist() == ["old text | Manual Review Required: [X]", "Manual Review Required: [X]", "Manual Review Required: [X]",
                                  "Manual Review Required: [X]", "Manual Review Required: [X]"], df.remarks.tolist())

    df = pd.DataFrame({"a": range(3)})
    r = MR()
    r.flag("no remarks column yet")
    r.apply(df)
    check(f"[{label}] a frame without a remarks column gets one", list(df.columns) == ["a", "remarks"] and df.remarks.str.startswith("Manual Review Required: [").all())

    empty = pd.DataFrame({"a": [], "remarks": []})
    r = MR()
    r.flag("x")
    counts, n = r.apply(empty)
    check(f"[{label}] an empty frame is fine", n == 0 and len(empty) == 0)

    r = MR()
    r.flag("  two   spaces\nand a\tline break  ")
    r.flag("   ")
    r.flag("")
    check(f"[{label}] text is tidied (one line, single spaces); blank notes are ignored", len(r.items) == 1 and r.items[0][0] == "two spaces and a line break")

    df = pd.DataFrame({"a": range(3), "remarks": [""] * 3})
    r = MR()
    r.flag("matches nothing", np.array([False, False, False]))
    out = []
    r.report(df, out.append)
    check(f"[{label}] a flag that matches no row is not reported and writes nothing", out == [] and (df.remarks == "").all())

    df = pd.DataFrame({"a": range(2), "remarks": ["", ""]})
    r = MR()
    r.flag("x" * 5000)
    r.apply(df)
    check(f"[{label}] a very long note is kept whole (the portal's remarks field is unlimited text)", len(df.remarks[0]) > 5000)

    big = pd.DataFrame({"a": np.arange(70000), "remarks": [""] * 70000})
    r = MR()
    for k in range(12):
        r.flag(f"note {k}", lambda d, k=k: (d.a % 12 == k).to_numpy() | (d.a % 5 == 0).to_numpy())
    t0 = time.time()
    r.apply(big)
    check(f"[{label}] 70,000 rows x 12 notes takes seconds, not minutes", time.time() - t0 < 15, f"{time.time() - t0:.1f}s")

    keys = m.unknown_notes(["short", "A long enough rule that is new", "a LONG enough   rule that is NEW", "The standing disclaimer, as written"],
                           ["the standing disclaimer"])
    check(f"[{label}] unknown_notes: short cells are labels, standing phrases (any case/spacing) are skipped, repeats collapse", keys == ["A long enough rule that is new"], keys)
    # the 25-character floor dropped 'In PCV 3W EV Decline' (20 characters) in SBI without a word (3 Nov 2026): a short text of two words or more that says decline / exclude / not allowed is a rule
    keys = m.unknown_notes(["In PCV 3W EV Decline", "TN not allowed", "Declined", "Priority Cities", "UPTO 50K", "Base + 0.5%", "PC STP GRID", "EV excluded in UP", "ok", "Except"], ["EV excluded in UP"])
    check(f"[{label}] unknown_notes: a short text that says decline / exclude / not allowed is a rule (flagged), a label, a one-word cell value and a standing phrase are not",
          keys == ["In PCV 3W EV Decline", "TN not allowed"], keys)

# ================================================================ Liberty
_LB = {}


def lib_base(which="sep"):
    if which not in _LB:
        _LB[which] = helpers.liberty_base(which)
    return copy.deepcopy(_LB[which])


def run_lib(sheets):
    lib.read_workbook = lambda p: helpers.FakeWB(sheets)
    return lib.build("fake.xlsx", dt.date(2026, 9, 1), dt.date(2030, 3, 31), quiet, lib.Asker(False))


df, warns = run_lib(lib_base())
fl = df[df.remarks != ""]
check("Liberty Sep: only the HCV3 / 'highlighted in green' note is flagged, on the CV >7.5T sheet rows (5,880)",
      len(fl) == 5880 and fl.remarks.str.contains("HCV3 upto 43.5T", regex=False).all() and set(fl["product"]) == {"GCV 4W", "MISCD"},
      (len(fl), fl.remarks.iloc[0][:80] if len(fl) else ""))
check("Liberty Sep: the standing notes (declined models disclaimer, SLABS..., ADVISED..., One grid...) are not flagged",
      not df.remarks.str.contains("declined models|SLABS|ADVISED|One grid", regex=True).any())

S = lib_base()
S["PC SATP"][1][0] = "EV rates are 10% lower than the Petrol grid for all GEOs"
df2, _ = run_lib(S)
note = df2[df2.remarks.str.contains("EV rates are 10% lower", regex=False)]
expected = df2[(df2["product"] == "Private Car") & (df2.sub_product == "STP")]
check("Liberty: a NEW note on 'PC SATP' is flagged on exactly the Private Car STP rows (that sheet) and nowhere else",
      len(note) == len(expected) > 0 and set(note.index) == set(expected.index), (len(note), len(expected)))
S = lib_base()
for r in S["PC SATP"][:5]:
    if r and isinstance(r[0], str) and r[0].startswith("ADVISED"):
        r[0] = "advised   to offer LESSER grid wherever   feasible"
df3, _ = run_lib(S)
check("Liberty: a standing note with different spacing / capitals is still standing (flag count unchanged)", int((df3.remarks != "").sum()) == 5880)
S = lib_base()
for r in S["PC SATP"][:5]:
    if r and isinstance(r[0], str) and r[0].startswith("SLABS"):
        r[0] = "SLABS MENTIONED HERE CAN BE EXCEEDED"
df4, _ = run_lib(S)
check("Liberty: a standing note whose MEANING changed (CANNOT -> CAN) is new and gets flagged", df4.remarks.str.contains("CAN BE EXCEEDED", regex=False).any())
check("Liberty: the rates are untouched by the flags (same numbers with and without the extra note)",
      df.drop(columns="remarks").equals(df2.drop(columns="remarks")))
try:
    run_lib(lib_base())
    check("Liberty: runs with no keyboard (a note never needs an answer)", True)
except Exception as e:
    check("Liberty: runs with no keyboard (a note never needs an answer)", False, repr(e))

# the RTO master is global: Liberty writes the master's group name (sheet prefix + geo, spaces as _) and flags a geo the master lacks
names = set(df["new_rto_list"])
check("Liberty Sep: every new_rto_list is ALLINDIA or <sheet prefix> + geo with '_' for spaces (3 prefixes), none left in the grid's own spelling",
      all(n == "ALLINDIA" or n.startswith(("LIBERTY_PC_TW_APR26_", "LIBERTY_GCV_4W_<7.5T_APR26_", "LIBERTY_GCV_4W>7.5T_APR26_")) for n in names) and not any(" " in n for n in names)
      and "LIBERTY_PC_TW_APR26_WEST_BENGAL_-_5_R" in names and "LIBERTY_GCV_4W_<7.5T_APR26_BIHAR" in names and "LIBERTY_GCV_4W>7.5T_APR26_BIHAR" in names, sorted(names)[:3])
bad = df[df.remarks.str.contains("not in the RTO master", regex=False)]
check("Liberty Sep: the master flags exactly the two CV >7.5T geos it lacks (ASSAM: the master has 'Assam'; TELANGANA: the master has a trailing '_'), and nothing else",
      set(bad["new_rto_list"]) == {"LIBERTY_GCV_4W>7.5T_APR26_ASSAM", "LIBERTY_GCV_4W>7.5T_APR26_TELANGANA"}
      and bad[bad.new_rto_list.str.endswith("ASSAM")].remarks.str.contains("the master spells it 'LIBERTY_GCV_4W>7.5T_APR26_Assam'", regex=False).all()
      and bad[bad.new_rto_list.str.endswith("TELANGANA")].remarks.str.contains("the master spells it 'LIBERTY_GCV_4W>7.5T_APR26_TELANGANA_'", regex=False).all()
      and (df[~df.index.isin(bad.index)].remarks.str.contains("RTO master", regex=False) == False).all(), sorted(set(bad["new_rto_list"])))
S = lib_base()
for ws_rows in S["PC SATP"]:
    if ws_rows and ws_rows[0] == "WEST BENGAL - 5 R":
        ws_rows[0] = "WEST BENGAL - 9 Z"
df5, _ = run_lib(S)
new_geo = df5[df5.new_rto_list == "LIBERTY_PC_TW_APR26_WEST_BENGAL_-_9_Z"]
check("Liberty: a geo the master has never heard of is written under the sheet's prefix and flagged (PC STP rows only), never mapped to another group",
      len(new_geo) > 0 and set(new_geo["sub_product"]) == {"STP"} and new_geo.remarks.str.contains("is not in the RTO master", regex=False).all()
      and not df5.new_rto_list.eq("LIBERTY_PC_TW_APR26_WEST_BENGAL_-_5_R").loc[df5["product"] == "Private Car"].any(), len(new_geo))

# ================================================================ Magma
GUIDE = [["SL No.", "Note:-"], [1, "For Volume consideration refer Additional Changes sheet"],
         [4, "For 2W, outgo will reduce by 2%, if CPA taken"], [8, "For Pvt Car any Non Diesel will have same rate as per Petrol"],
         [9, "The below RTOs are blocked for all Products on account of high TP claim :"], [None, "RTO Location Code *"], [None, "MP-49--NARSINGHPUR"],
         [10, "The following make and model combinations will be considered under GCV 20 to 40T with a highest GVW of up to 43500T"],
         [None, "Make", "Model"], [None, "TATA", "LPT 3723"], [None, "ASHOK LEYLAND", "AL 3718"],
         [11, "Below Vehicle Body Type from CE and TP-CE to be considered as below-"], ["BIZ MIX 3 (Agency)", "Vehicle Body Type", "Revised Bix Mix as per Grid"],
         ["CE", "CRANE", "CE-CONSTRUCTION EQ"]]
NOTES = [[None, None, None], [None, "UW Cluster", "LOB", "Remarks"], [None, "TL1", "Tractor Old & Harvester Old", "Each slab will be reduced by 1%"],
         [None, "KA1", "Tractor Old", "Top 4 Slab will be same & last slab will be -1"]]
SATP_NOTE = [["Note:-"], ["Mentioned grid is on overall motor buisness volume of 1st year earned premium"],
             ["CPA Premium will not be considered for outgo if less than Rs 450;Except for TW"], ["The below Body Types Belong to CE"],
             ["*-List of body Types belongs to CE."], ["CRANE", "VIBRATORY COMPACTOR"], ["MOTOR GRADER", "TRUCK CRANE"]]
_MG = {}


def run_mag(guide=None, notes=None, satp_note=None, break_sheet=None):
    if "b" not in _MG:
        _MG["b"] = helpers.magma_base()
    motor, satp = copy.deepcopy(_MG["b"][0]), copy.deepcopy(_MG["b"][1])
    motor["Generic Guidelines"] = copy.deepcopy(guide if guide is not None else GUIDE)
    motor["Notes"] = copy.deepcopy(notes if notes is not None else NOTES)
    satp["Note"] = copy.deepcopy(satp_note if satp_note is not None else SATP_NOTE)
    data = {"fake_grid.xlsx": motor, "fake_satp.xlsb": satp}
    mag.list_sheets = lambda p: list(data[p])

    def read(p, s):
        if s == break_sheet:
            raise OSError("cannot read this sheet")
        return data[p][s]
    mag.read_rows = read
    return mag.build("fake_grid.xlsx", "fake_satp.xlsb", dt.date(2026, 10, 1), dt.date(2030, 3, 31), quiet, mag.Asker(False))


df, _ = run_mag()
check("Magma: the standing guidelines (volume, CPA, blocked RTOs, make list, body types, slab rules ...) are not flagged and make/model tables are not mistaken for rules",
      (df.remarks == "").all(), df.remarks[df.remarks != ""].head(2).tolist())
g = copy.deepcopy(GUIDE) + [[12, "For GCV 3W, outgo will reduce by 3% for all clusters"]]
df, _ = run_mag(guide=g)
fl = df[df.remarks != ""]
check("Magma: a NEW guideline in the comprehensive grid's sheet is flagged on the comprehensive rows only (not on STP rows)",
      len(fl) > 0 and fl.sub_product.ne("STP").all() and (df[df.sub_product != "STP"].remarks != "").all() and (df[df.sub_product == "STP"].remarks == "").all(), (len(fl),))
df, _ = run_mag(satp_note=SATP_NOTE + [["Tractor TP rates are not applicable for the state of Kerala"]])
fl = df[df.remarks != ""]
check("Magma: a NEW note in the SATP grid's sheet is flagged on the STP rows only", len(fl) > 0 and fl.sub_product.eq("STP").all() and (df[df.sub_product == "STP"].remarks != "").all())
g = [list(r) for r in GUIDE]
g[2] = [4, "For 2W, outgo will reduce by 3%, if CPA taken"]
df, _ = run_mag(guide=g)
check("Magma: a standing guideline whose NUMBER changed (2% -> 3%) is new and gets flagged", df.remarks.str.contains("reduce by 3%", regex=False).any())
base_m = run_mag()[0]
check("Magma: every new_rto_list is MAGMA_APR26_ + the cluster with '_' for spaces (the RTO master's name) and the master flags none of the 114 clusters",
      base_m.new_rto_list.str.startswith("MAGMA_APR26_").all() and not base_m.new_rto_list.str.contains(" ").any() and base_m.new_rto_list.nunique() == 114
      and "MAGMA_APR26_AP" in set(base_m.new_rto_list) and (base_m.remarks == "").all(), sorted(set(base_m.new_rto_list))[:3])
_mb = helpers.magma_base()
_mrows = copy.deepcopy(_mb[0][mag.MOTOR_SHEET])
for row in _mrows:
    for j in range(min(3, len(row))):
        if row[j] == "BR":
            row[j] = "ZZ 9"
_saved = _MG["b"]
_MG["b"] = ({**_mb[0], mag.MOTOR_SHEET: _mrows}, _mb[1])
df_z, _ = run_mag()
_MG["b"] = _saved
zz = df_z[df_z.new_rto_list == "MAGMA_APR26_ZZ_9"]
check("Magma: a cluster the master has never heard of (BR renamed ZZ 9) is written as MAGMA_APR26_ZZ_9 on its motor rows and flagged; no other cluster is flagged",
      len(zz) > 0 and zz.remarks.str.contains("new_rto_list 'MAGMA_APR26_ZZ_9' is not in the RTO master", regex=False).all()
      and (df_z[df_z.new_rto_list != "MAGMA_APR26_ZZ_9"].remarks == "").all(), len(zz))
df, _ = run_mag(notes=NOTES + [[None, "KA9", "Tractor New", "Each slab will be reduced by 2%"]])
check("Magma: a new remark in the 'Remarks' column of the Notes table is flagged", df.remarks.str.contains("reduced by 2%", regex=False).any())
df, _ = run_mag(break_sheet="Generic Guidelines")
check("Magma: a guideline sheet that cannot be read is skipped - the run does not stop", len(df) > 0)

# ================================================================ the master lists (shared block, identical in all five scripts)
import tempfile  # noqa: E402

pvt = helpers.load_script("rsa_pvtcar_to_config.py")
uni = helpers.load_script("united_commission_to_config.py")
for label, m in (("liberty", lib), ("magma", mag), ("rsa", rsa), ("pvtcar", pvt), ("united", uni)):
    masters = m.Masters({"ALLINDIA": ["AN01", "MH01"], "MH": ["MH01"], "GRP_X": ["MH01", "MH02"]}, {"TATA": ["TATA NANO"], "Ashok": ["ASHOK LEYLAND X"], "Private_Car_ALL": ["A B"]})
    df = pd.DataFrame({"new_rto_list": ["ALLINDIA", "MH,GRP_X", "ZZ", "MH,QQ", "", "GRP_X"],
                       "new_vehicle_makes": ["TATA", "TATA,Ashok", "TATA,ASHOK", "NEWMAKE", "Private_Car_All", "Private_Car_ALL"], "remarks": [""] * 6})
    r = m.ManualReview()
    n = m.check_names(df, r, masters, quiet)
    r.apply(df)
    rem = df.remarks.tolist()
    check(f"[{label}] check_names: names in the master are not flagged (rows 0, 1 and 5 of the RTO column); unknown ones are, a comma list name by name, an empty list too",
          "new_rto_list" not in rem[0] and "new_rto_list" not in rem[1] and "new_rto_list" not in rem[5] and "new_rto_list 'ZZ' is not in the RTO master" in rem[2]
          and "name(s) 'QQ' (in 'MH,QQ')" in rem[3] and "new_rto_list '' is not in the RTO master" in rem[4], rem)
    check(f"[{label}] check_names: a make name that differs from the master only in capitals is flagged with the master's spelling as a hint, never replaced; an unknown make is flagged",
          "the master spells it 'Ashok'" in rem[2] and "the master spells it 'Private_Car_ALL'" in rem[4] and "'NEWMAKE'" in rem[3] and "new_vehicle_makes" not in rem[0] + rem[1] + rem[5]
          and df.new_vehicle_makes.tolist()[2] == "TATA,ASHOK", rem)
    check(f"[{label}] check_names returns the number of rows flagged: rows 2, 3, 4 (RTO) and 2, 3, 4 (makes) = 3 rows", n == 6, n)
    check(f"[{label}] Masters.near: a spelling that differs only in capitals gives the master's; the same spelling or no match gives None", masters.near("makes", "ASHOK") == "Ashok"
          and masters.near("makes", "Ashok") is None and masters.near("makes", "nothing") is None and masters.codes("MH") == {"MH01"} and masters.codes("zzz") == set())

    check(f"[{label}] Masters.near also sees through a stray leading / trailing '_' or space (the master has 'LIBERTY_..._TELANGANA_'), and still only as a hint",
          masters.near("rto", "GRP_X_") == "GRP_X" and masters.near("rto", " grp_x ") == "GRP_X" and masters.near("rto", "GRP_X") is None and masters.near("rto", "GRP_Y") is None)

    import openpyxl
    with tempfile.TemporaryDirectory() as tmp:
        def write(name, rows, age):
            wb = openpyxl.Workbook()
            wb.active.append(["ID", "NAME", "CLUSTER"])
            for i, (nm, items) in enumerate(rows.items(), 1):
                wb.active.append([i, nm, items])
            path = os.path.join(tmp, name)
            wb.save(path)
            os.utime(path, (time.time() - age, time.time() - age))
        write("rto_master (1).xlsx", {"OLD": "AA01"}, 500)
        write("RTO_MASTER (3).xlsx", {"NEW": "MH01,MH02", "ALLINDIA": "MH01"}, 100)
        write("make_model_master.xlsx", {"TATA": "TATA NANO,TATA NEXON"}, 100)
        write("notes.xlsx", {"x": "y"}, 10)
        got = m.load_masters(m.Asker(False), folder=tmp)
        check(f"[{label}] load_masters: the newest file with 'rto' + 'master' (case ignored) and the make / model one are read; other workbooks are ignored",
              set(got.rto) == {"NEW", "ALLINDIA"} and got.rto["NEW"] == ["MH01", "MH02"] and got.makes == {"TATA": ["TATA NANO", "TATA NEXON"]}
              and [os.path.basename(p) for p in got.sources] == ["RTO_MASTER (3).xlsx", "make_model_master.xlsx"], got.sources)
    with tempfile.TemporaryDirectory() as empty:
        e = None
        try:
            m.load_masters(m.Asker(False), folder=empty)
        except m.GridError as ex:
            e = ex
        check(f"[{label}] load_masters: no master workbook in the folder -> without a keyboard the run stops with a message that says so", e is not None and "not found" in str(e), e)
        with tempfile.TemporaryDirectory() as good:
            for nm, rows in (("rto_master.xlsx", {"ALLINDIA": "AN01"}), ("make_model_master.xlsx", {"TATA": "TATA NANO"})):
                wb = openpyxl.Workbook()
                wb.active.append(["ID", "NAME", "CLUSTER"])
                for i, (a, b) in enumerate(rows.items(), 1):
                    wb.active.append([i, a, b])
                wb.save(os.path.join(good, nm))
            answers = iter(["C:\\no\\such\\folder", good])
            got = m.load_masters(m.Asker(True, lambda prompt="": next(answers), quiet), folder=empty)
            check(f"[{label}] load_masters: with a keyboard it asks for the folder, asks again for a path that is no folder, and then loads", set(got.rto) == {"ALLINDIA"} and set(got.makes) == {"TATA"})

bad = [n for n, ok, _ in results if not ok]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
