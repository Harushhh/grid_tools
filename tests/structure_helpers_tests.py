"""The small helpers of the structure-shift hardening (11 Nov 2026), checked one by one and fast - the slow proof is tests/structure_shift_tests.py:

  Magma     slab_of / slab_sheet (a sheet named '<=3 Lacs v2' is the '< =3 lac' slab), find_uw_header (the 'UW Cluster' cell in any column)
  ICICI     tw_oem_key / one_sheet_per_oem (the OEM of a new Two-Wheeler sheet from the words of its name), twin_groups / twin_blank (the same heading twice: the lowest payin)
  United    _is_new_head (more wordings of the new payout column's heading)
  RSA GCV   find_header ('Tonnage Band' in any column; a note that mentions tonnages is no header)
  Liberty   sheet_by_columns / blocks_of (a sheet that is not called as the settings say is found by its columns; the 'Geo Segment' header in any column)

Run:  .venv\\Scripts\\python.exe tests\\structure_helpers_tests.py"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402
import openpyxl  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"), flush=True)


def raises(fn, exc):
    try:
        fn()
    except exc as e:
        return e
    except Exception as e:                                     # another exception type is not the one wanted
        return None if not isinstance(e, exc) else e
    return None


# ---------------------------------------------------------------- Magma
mag = helpers.load_script("magma_grid_to_config.py")
slabs = {"< =3 lac": (0.0, 3.0), "<=3 Lacs v2": (0.0, 3.0), "Upto 3 lakh": (0.0, 3.0), "below 3 lac": (0.0, 3.0), ">3 to <=5": (3.0, 5.0), ">3 to <=5 lac": (3.0, 5.0),
         "> 3 to < =5 Lacs": (3.0, 5.0), ">5 to <=8": (5.0, 8.0), "> 8 lac": (8.0, None), "Sheet1": None, "Notes": None}
bad = {n: (mag.slab_of(n), want) for n, want in slabs.items() if mag.slab_of(n) != want}
check("Magma slab_of: the premium slab a sheet name says, in any spelling (11 names)", not bad, bad)
names = ["Generic Guidelines", "<=3 Lacs v2", ">3 to <=5 v2", ">5 to <=8 v2"]
check("Magma slab_sheet: '< =3 lac' is the sheet whose name says that slab ('<=3 Lacs v2'), '>3 to <=5' the one for 3-5", mag.slab_sheet(names, "< =3 lac") == "<=3 Lacs v2" and mag.slab_sheet(names, ">3 to <=5") == ">3 to <=5 v2")
check("Magma slab_sheet: the sheet called exactly the wanted name wins (None = the usual lookup), two sheets with one slab are not guessed",
      mag.slab_sheet(["< =3 lac", "<=3 Lacs v2"], "< =3 lac") is None and mag.slab_sheet(["<=3 Lacs v2", "upto 3 lakh"], "< =3 lac") is None)
check("Magma find_uw_header: 'UW Cluster' in column A, in column D, with spaces / capitals changed, absent",
      mag.find_uw_header([[None], ["x"], ["UW Cluster/Outgo on", "a"]]) == (2, 0) and mag.find_uw_header([[None] * 4, [None, None, None, "uw  cluster"]]) == (1, 3)
      and mag.find_uw_header([["a", "b"], ["c"]]) == (None, 0))

# ---------------------------------------------------------------- ICICI
ic = helpers.load_script("icici_grid_to_config.py")
oems = {"HMC": "hmc", "HMC v2": "hmc", "Hero (Sept)": "hmc", "HMSI": "hmsi", "Honda": "hmsi", "TVS": "tvs", "Suzuki - Oct": "suzuki", "Yamaha": "yamaha", "RE": "re", "Royal Enfield": "re",
        "EV": "ev", "Sheet1": "", "Hero Honda": "", "Notes": ""}
bad = {n: (ic.tw_oem_key(n), want) for n, want in oems.items() if ic.tw_oem_key(n) != want}
check("ICICI tw_oem_key: the OEM of a new Two-Wheeler sheet from the words of its name (14 names; two OEM words = none, not a guess)", not bad, bad)


class Ctx:
    def __init__(self):
        self.warnings = []


c = Ctx()
kept = ic.one_sheet_per_oem(c, [("HMC", [], "", 0), ("HMC (old)", [], "", 0), ("TVS", [], "", 0)])
check("ICICI one_sheet_per_oem: 'HMC' and a left-in 'HMC (old)': the sheet called exactly the OEM is used, a warning says so", [k[0] for k in kept] == ["HMC", "TVS"] and len(c.warnings) == 1, (kept, c.warnings))
e = raises(lambda: ic.one_sheet_per_oem(Ctx(), [("HMC v2", [], "", 0), ("HMC v3", [], "", 0)]), ic.GridError)
check("ICICI one_sheet_per_oem: two sheets that say HMC and none called exactly HMC STOP the run, naming both", e is not None and "HMC v2" in str(e) and "HMC v3" in str(e) and "expected one sheet per OEM" in str(e), e)
g = ic.twin_groups([(2, "gcv"), (3, "scv"), (6, "gcv"), (7, "tp")])
check("ICICI twin_groups: the same heading in columns C and G is one group of two (both columns are read, the lowest payin is written); the others are no twins", g == {2: (2, 6), 6: (2, 6)}, g)
check("ICICI twin_groups: headings that differ ('SCV >= 2450' / 'SCV <= 2450') are no twins", ic.twin_groups([(2, "scv >= 2450 gvw"), (3, "scv <= 2450 gvw")]) == {})
tw = {2: (2, 6), 6: (2, 6)}
check("ICICI twin_blank: a blank cell of a twin is no payin when the other twin has a number (it must not become a zero card that wins) ...",
      ic.twin_blank({2: 0.4, 6: None}, tw, 6) and ic.twin_blank({2: 0.4, 6: ""}, tw, 6) and ic.twin_blank({2: None, 6: 0.3}, tw, 2))
check("... but a zero is a payin (kept), two blanks are read as before, and a column without a twin is never skipped",
      not ic.twin_blank({2: 0.4, 6: 0}, tw, 6) and not ic.twin_blank({2: None, 6: None}, tw, 2) and not ic.twin_blank({3: None}, tw, 3))

# ---------------------------------------------------------------- United
un = helpers.load_script("united_commission_to_config.py")
heads = {"Pay-out (w.e.f. 01-07-2026)": True, "Revised Pay-out (w.e.f. 01-08-2026)": True, "Maximum proposed Commission/Brokerage % w.e.f. 01/04/2026": True, "New Pay-out": True,
         "Proposed Commission (w.e.f. 26-10-2026)": True, "Revised Brokerage %": True, "Updated Payout": True, "Existing Pay-out (w.e.f. 01-04-2026)": False, "Commission": False,
         "States (based on Registration of Vehicle)": False, "S No": False, "Class of Vehicles": False, "Updated rate": False, "Revised proposal": False}
bad = {t: want for t, want in heads.items() if un._is_new_head(t) != want}
check("United _is_new_head: 7 wordings of the new column's heading are read, 7 other headings are not", not bad, bad)

# ---------------------------------------------------------------- RSA GCV
rg = helpers.load_script("rsa_gcv4w_comp_to_config.py")
hdr = ["Tonnage Band", "Age", "AMRAVATI", "AURANGABAD", "KONKAN"]
check("RSA GCV find_header: 'Tonnage Band' in column A, in column D (three blank columns left of the table), as '(row, column)'",
      rg.find_header([["GCV"], ["NON EV"], hdr]) == (2, 0) and rg.find_header([[None] * 5, [None, None, None, *hdr]]) == (1, 3))
check("RSA GCV find_header: a long note that mentions tonnages is no header, and a sheet without one gives (None, 0)",
      rg.find_header([["The tonnages below are for old vehicles only, see notes", "a", "b", "c"]]) == (None, 0) and rg.find_header([["x", "y"]]) == (None, 0))

# ---------------------------------------------------------------- Liberty
lib = helpers.load_script("liberty_grid_to_config.py")


def sheet(wb, name, header, rows, left=0):
    ws = wb.create_sheet(name)
    for r, row in enumerate([[None] * left + header] + [[None] * left + x for x in rows], start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(r, c).value = v
    return ws


tw_cols = [c[0] for c in lib.TW_COMP]
wb = openpyxl.Workbook()
wb.remove(wb.active)
sheet(wb, "Rates 2026", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.2, 0.2, 0.15, 0.15], ["GOA", "50k-1.5L", 0.25, 0.2, 0.15, 0.1]])
sheet(wb, "Other", ["Geo Segment", "Volume Slab", "Something", "Else"], [["DELHI", "1L", 1, 2]])
expect = [(None, c) for c in tw_cols]
check("Liberty sheet_by_columns: the one sheet whose columns are the role's columns is found although its name is another", lib.sheet_by_columns(wb, "Geo segment", True, None, expect, wanted=lib.PCTW_SHEET) == "Rates 2026")
asker = lib.Asker(False)
blocks = lib.blocks_of(wb, lib.PCTW_SHEET, "Geo segment", True, None, asker, "the TW rows", expect=expect)
check("Liberty blocks_of: reads that sheet (the run says which sheet it used)", blocks and any(c[1] == tw_cols[0] for b in blocks for c in b.cols), blocks)
wb2 = openpyxl.Workbook()
wb2.remove(wb2.active)
sheet(wb2, "A", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.2, 0.2, 0.15, 0.15]])
sheet(wb2, "B", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.3, 0.2, 0.15, 0.15]])
check("Liberty sheet_by_columns: two sheets that have the columns equally well are NOT guessed (None -> the usual question)", lib.sheet_by_columns(wb2, "Geo segment", True, None, expect, wanted=lib.PCTW_SHEET) is None)
wb3 = openpyxl.Workbook()
wb3.remove(wb3.active)
sheet(wb3, "CV <7.5T Comp & SATP v2", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.2, 0.2, 0.15, 0.15]])
sheet(wb3, "CV >7.5T Comp & SATP v2", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.3, 0.2, 0.15, 0.15]])
check("Liberty sheet_by_columns: sheets with the same columns are told apart by the wanted name that is still part of theirs ('CV <7.5T Comp & SATP' -> '... v2' with '<')",
      lib.sheet_by_columns(wb3, "Geo segment", True, None, expect, wanted="CV <7.5T Comp & SATP") == "CV <7.5T Comp & SATP v2")
wb4 = openpyxl.Workbook()
wb4.remove(wb4.active)
sheet(wb4, "PC TW_Comp and SAOD", ["Geo Segment", "Volume Slab", *tw_cols], [["DELHI", "50k-1.5L", 0.2, 0.2, 0.15, 0.15]], left=3)
b4 = lib.blocks_of(wb4, "PC TW_Comp and SAOD", "Geo segment", True, None, asker, "the TW rows", expect=expect)
check("Liberty _blocks: the 'Geo Segment' header in column D (three blank columns left of the table) is found, the columns are read", b4 and b4[0].geo_col == 3 and any(c[1] == tw_cols[0] for c in b4[0].cols), b4)

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
