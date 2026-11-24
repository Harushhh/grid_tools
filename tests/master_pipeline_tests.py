"""The two-phase RTO / make master pipeline (the user, 31 Oct 2026).
Phase 1 (RTO_NAMES = "master", Shriram): a scope whose RTO codes are exactly a master group gets the master's name; a scope the master lacks is NOT guessed and does not stop the run - its cell holds
'New changes found, new mapping needs to be created: <Sheet Name> - <RTOs>'.  Phase 2 (master_builder.py): reads those warnings and writes the two-column Excel that adds the groups.
Run:  .venv\\Scripts\\python.exe tests\\master_pipeline_tests.py"""
import importlib.util
import os
import shutil
import subprocess
import sys
import tempfile
from collections import Counter

os.environ["GRID_RTO_NAMES"] = "master"                   # before any script is loaded
import numpy as np  # noqa: E402
import openpyxl  # noqa: E402
import pandas as pd  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

ROOT = helpers.ROOT
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


S = helpers.load_script("shriram_broker_grid_to_config.py")
S2 = helpers.load_script("shriram_2w_to_config.py")
check("the Shriram scripts default to the master mode", S.RTO_NAMES == "master" and S2.RTO_NAMES == "master")

# ================================================================ the warning text
text = S.master_text("Pin code Metro", ["AP02", "AP01", "AP03"])
check("master_text: 'New changes found, new mapping needs to be created: <Sheet> - <RTOs>'", text == "New changes found, new mapping needs to be created: Pin code Metro - AP02/AP01/AP03", text)
cell = "GA,DD," + S.master_text("New Business - 2W", ["TW_ALL_EX_TVS"]) + "," + S.master_text("PAN India", ["RJ01", "RJ02"])
names, warns = S.split_cell(cell)
check("split_cell: the names, then every warning with its sheet (a ' - ' inside the sheet name is fine), its RTO codes or its group name",
      names == ["GA", "DD"] and [(w["sheet"], w["items"], w["name"]) for w in warns] == [("New Business - 2W", [], "TW_ALL_EX_TVS"), ("PAN India", ["RJ01", "RJ02"], "")], (names, warns))
mt = S.master_text("New Business - 2W", name="TW_ALL_EX_TVS", base="TW_ALL", excluded=["TVS"])
check("a make group with a definition: '<Sheet> - <GROUP> = <BASE> except <GROUP>|<GROUP>'", mt == "New changes found, new mapping needs to be created: New Business - 2W - TW_ALL_EX_TVS = TW_ALL except TVS"
      and S.split_cell(mt)[1] == [dict(sheet="New Business - 2W", items=[], name="TW_ALL_EX_TVS", base="TW_ALL", excluded=["TVS"])], S.split_cell(mt))
check("split_cell: a plain cell has no warning", S.split_cell("GJ,DD,DN") == (["GJ", "DD", "DN"], []))


# ================================================================ phase 1 on a data frame: names and makes
class Review:
    def __init__(self):
        self.flags = []

    def flag(self, text, rows=None):
        self.flags.append((text, rows))


S.reset_run_lists()
masters = S.Masters({"MH": ["MH01", "MH02", "MH03"], "MH_EX_02": ["MH01", "MH03"], "GA": ["GA01"]}, {"TATA": ["Tata Ace"], "Mahindra": ["Bolero"]})
df = pd.DataFrame({"new_rto_list": ["MH_EX_02", "ga", "GA," + S.master_text("Declined-Excluded", ["RJ01"]), "NOPE"], "new_vehicle_makes": ["tata", "Mahindra", "NEWMAKE", "TATA,OTHERMAKE"]})
S.MASTER_SHEET = "GRID OCT' 26"
rev = Review()
n = S.check_names(df, rev, masters, log=lambda *a, **k: None)
check("a name the master has stays; a different capital letter is the master's spelling", df["new_rto_list"][0] == "MH_EX_02" and df["new_rto_list"][1] == "GA" and df["new_vehicle_makes"][0] == "TATA" and df["new_vehicle_makes"][1] == "Mahindra", df.to_dict("list"))
check("an RTO name the master lacks becomes the plain-text warning, and the warnings already in a cell are kept",
      df["new_rto_list"][3] == "New changes found, new mapping needs to be created: GRID OCT' 26 - NOPE" and df["new_rto_list"][2] == "GA,New changes found, new mapping needs to be created: Declined-Excluded - RJ01", df["new_rto_list"].tolist())
check("a make the master lacks becomes the warning too; the makes it has stay in front", df["new_vehicle_makes"][2] == "New changes found, new mapping needs to be created: GRID OCT' 26 - NEWMAKE"
      and df["new_vehicle_makes"][3] == "TATA,New changes found, new mapping needs to be created: GRID OCT' 26 - OTHERMAKE", df["new_vehicle_makes"].tolist())
check("the run goes on: the rows with a warning are counted and flagged, the new mappings are listed on their own sheet", n == 2 and len(rev.flags) == 1 and "New mappings to create" in S.EXTRA_SHEETS
      and set(S.EXTRA_SHEETS["New mappings to create"]["kind"]) == {"RTO", "make / model"}, (n, list(S.EXTRA_SHEETS)))


# ================================================================ phase 1 in Shriram's RTO resolver
class Declines:
    """what Declines gives resolve(): the RTOs taken out of a state, and the sheet they came from"""

    def __init__(self, gone, sheet="Declined-Excluded"):
        self.gone, self.sheet, self.last_sheets = set(gone), sheet, set()

    def decline(self, labels, letters, exempt=False):
        self.last_sheets = {self.sheet}
        return {c for c in self.gone if c[:2] in letters}


for mod in (S, S2):
    ms = mod.Masters({"MH": ["MH01", "MH02", "MH03", "MH04"], "MH_EX_02": ["MH01", "MH03", "MH04"], "RJ": ["RJ01", "RJ02", "RJ03"], "RJ41": ["RJ41"],
                      "ALLINDIA": ["MH01", "MH02", "MH03", "MH04", "RJ01", "RJ02", "RJ03"]}, {"TATA": ["x"]})
    rt = mod.Rtos(ms)
    sc, missing = rt.resolve([("codes", ["MH"])], Declines({"MH02"}), frozenset(), sheet="GRID OCT' 26")
    check(f"{mod.__name__}: the RTOs left are exactly a master group -> the master's name, no warning", str(sc) == "MH_EX_02" and missing == [], (str(sc), missing))
    sc, missing = rt.resolve([("codes", ["MH"])], Declines({"MH01"}), frozenset(), sheet="GRID OCT' 26")
    check(f"{mod.__name__}: no master group has exactly these RTOs -> the plain-text warning with the sheet and the RTOs; nothing is widened to the whole state",
          str(sc) == "New changes found, new mapping needs to be created: Declined-Excluded - MH02/MH03/MH04" and missing == [], (str(sc), missing))
    sc, _ = rt.resolve([("codes", ["MH"]), ("codes", ["RJ"])], Declines({"MH01"}), frozenset())
    check(f"{mod.__name__}: a whole state next to an unmapped one: the name first, the warning last", str(sc) == "RJ,New changes found, new mapping needs to be created: Declined-Excluded - MH02/MH03/MH04"
          or str(sc).endswith("MH02/MH03/MH04") and str(sc).startswith("RJ,"), str(sc))
    sc, _ = rt.resolve([("codes", ["RJ"])], Declines({"RJ01"}), frozenset(), include={"RJ41", "RJ02", "RJ03"})
    check(f"{mod.__name__}: a single-RTO group is not a name for a scope of several RTOs", "RJ41" not in str(sc).split("New changes")[0], str(sc))
    sc, _ = rt.resolve([("codes", ["MH"])], Declines(set()), frozenset())
    check(f"{mod.__name__}: a whole state stays the state's name", str(sc) == "MH", str(sc))

S2.reset_run_lists()
dfm = pd.DataFrame({"new_rto_list": ["GA", "GA"], "new_vehicle_makes": ["TW_ALL_EX_TVS", "TATA"]})
S2.MASTER_SHEET = "New Business - 2W + ROLLOVER 2W (1+1)"
S2.check_names(dfm, Review(), S2.Masters({"GA": ["GA01"]}, {"TATA": ["x"], "TW_ALL": ["BAJAJ CHETAK", "TVS XL"], "TVS": ["TVS XL"]}), log=lambda *a, **k: None)
check("Shriram 2W: the make group the master lacks carries its definition (all two-wheeler makes except TVS) in the warning",
      dfm["new_vehicle_makes"][0] == "New changes found, new mapping needs to be created: New Business - 2W + ROLLOVER 2W (1+1) - TW_ALL_EX_TVS = TW_ALL except TVS" and dfm["new_vehicle_makes"][1] == "TATA", dfm["new_vehicle_makes"].tolist())

# ================================================================ phase 2: master_builder.py
spec = importlib.util.spec_from_file_location("master_builder", os.path.join(ROOT, "master_builder.py"))
B = importlib.util.module_from_spec(spec)
spec.loader.exec_module(B)
AP = [f"AP{n:02d}" for n in range(1, 40)]
TG = [f"TG{n:02d}" for n in range(1, 12)]
master = {"AP": AP, "TG": TG, "SHRIRAM_OCT26_AP_X": ["AP01", "AP02"]}
sets = {frozenset(set(AP) - {"AP27"}): {"sheets": {"Declined-Excluded"}, "rows": 3},
        frozenset({"AP05", "AP06", "AP07"}): {"sheets": {"Pin code Metro"}, "rows": 2},
        frozenset({"AP08", "AP09"}): {"sheets": {"Pin code Metro"}, "rows": 1},
        frozenset({"AP01", "AP02"}): {"sheets": {"PAN India"}, "rows": 1},
        frozenset(TG[:3] + AP[:2]): {"sheets": {"Zone"}, "rows": 1}}
rows, notes = B.build(sets, master, "shriram", "oct26")
by = {c: n for n, c in rows}
names = [n for n, _ in rows]
check("master_builder: the RTOs left out of the state name the group (shriram_oct26_ap_ex_27)", "shriram_oct26_ap_ex_27" in names, names)
check("master_builder: a sheet that names a kind of area gives the condition (shriram_oct26_ap_metro), a second one with the same words gets _2", "shriram_oct26_ap_metro" in names and "shriram_oct26_ap_metro_2" in names, names)
check("master_builder: a set the master already has is not added again, and says so", all(c != "AP01,AP02" for _, c in rows) and any("already the master group" in x for x in notes), (rows, notes))
check("master_builder: several states -> all the states in the region word", any(n.startswith("shriram_oct26_ap_tg_") or n.startswith("shriram_oct26_tg_ap_") for n in names), names)
check("master_builder: every name is lowercase, distinct, letters / digits / '_' only - no spaces, no commas", all(B.NAME_OK.fullmatch(n) and n == n.lower() for n in names) and len(set(names)) == len(names), names)
check("master_builder: every rto_cluster is comma-separated RTO codes with no spaces", all(B.CLUSTER_OK.fullmatch(c) for _, c in rows) and all(" " not in c for _, c in rows), rows)
B.check_rows(rows)
for bad in ([("Has Space", "AP01")], [("a,b", "AP01")], [("ok_name", "AP01, AP02")], [("ok_name", "AP01"), ("ok_name", "AP02")], [("UPPER", "AP01")]):
    try:
        B.check_rows(bad)
        check(f"master_builder refuses {bad}", False)
    except B.BuilderError:
        check(f"master_builder refuses {bad}", True)
check("master_builder: the month and the insurer come from the file (from_date, insurance_company)", B.month_word(Counter({"01-10-2026": 9, "01/09/2026": 2})) == "oct26" and B.insurer_word(Counter({"SHRIRAM GENERAL INSURANCE CO. LTD.": 4})) == "shriram")

mm = {"TW_ALL": ["BAJAJ CHETAK", "TVS XL", "HERO SPLENDOR", " HONDA ACTIVA"], "TVS": ["TVS XL", "TVS SCOOTY"], "Mahindra": ["BOLERO"], "HERO_ONLY": ["HERO SPLENDOR"]}
mg = {"TW_ALL_EX_TVS": {"base": "TW_ALL", "excluded": ["TVS"], "sheets": {"New Business - 2W"}, "rows": 3}, "NO_DEFINITION": {"base": "", "excluded": [], "sheets": {"x"}, "rows": 1},
      "SHRIRAM_OCT26_HERO": {"base": "TW_ALL", "excluded": ["TVS", "tw_all"], "sheets": {"x"}, "rows": 1}, "NOT_THERE": {"base": "NO_SUCH_GROUP", "excluded": [], "sheets": {"x"}, "rows": 1}}
mrows, mnotes = B.build_makes(mg, mm, "shriram", "oct26")
check("make master: the group is its base group in the make master minus the groups taken out - the master's own entries, trimmed, in the master's order",
      ("shriram_oct26_tw_all_ex_tvs", "BAJAJ CHETAK,HERO SPLENDOR,HONDA ACTIVA") in mrows, mrows)
check("make master: a warning with no definition, or whose base group the master lacks, or that would be empty, is listed and not built",
      len(mrows) == 1 and sum("not built" in n for n in mnotes) == 3, (mrows, mnotes))
check("make master: the entries are comma-separated with no space around the commas; a model's own inner spaces stay", all(", " not in c and " ," not in c for _, c in mrows) and "HERO SPLENDOR" in mrows[0][1])
B.check_rows(mrows, "make")
for bad_make in ([("ok_name", "A B, C")], [("ok_name", "A B ,C")], [("ok_name", "A,,B")], [("Not Lower", "A")], [("ok_name", "")]):
    try:
        B.check_rows(bad_make, "make")
        check(f"make master refuses {bad_make}", False)
    except B.BuilderError:
        check(f"make master refuses {bad_make}", True)
check("make names drop the insurer and month words the group's name already has", B.build_makes({"SHRIRAM_OCT26_TW_X": {"base": "TW_ALL", "excluded": ["TVS"], "sheets": {"x"}, "rows": 1}}, mm, "shriram", "oct26")[0][0][0] == "shriram_oct26_tw_x")

tmp = tempfile.mkdtemp(prefix="masterpipe-")
out = os.path.join(tmp, "add.xlsx")
B.write_xlsx(rows, out)
wb = openpyxl.load_workbook(out)
got = list(wb.active.iter_rows(values_only=True))
check("the Excel file has ONE sheet with exactly the two columns rto_name, rto_cluster", wb.sheetnames == ["rto_master_additions"] and list(got[0]) == ["rto_name", "rto_cluster"] and len(got) == len(rows) + 1 and all(len(r) == 2 for r in got), (wb.sheetnames, got[:2]))
check("free_path never overwrites an existing file", B.free_path(out) != out and not os.path.exists(B.free_path(out)))

# ================================================================ the loop: converter -> warnings -> master_builder -> new master -> converter has no warnings left
INP = os.path.join(ROOT, "regression", "inputs")
env = dict(os.environ, PYTHONIOENCODING="utf-8", GRID_RTO_NAMES="master")


def converter(master_dir, out_csv):
    args = [sys.executable, "-I", "shriram_broker_grid_to_config.py", "--grid", os.path.join(INP, "shriram_sep26.xlsx"), "--from-date", "01/09/2026", "--to-date", "31/03/2030", "--no-prompt", "--out", out_csv]
    return subprocess.run(args, cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900, env=dict(env, **({"GRID_MASTER_DIR": master_dir} if master_dir else {})))


csv1 = os.path.join(tmp, "first.csv")
p = converter(None, csv1)
d1 = pd.read_csv(csv1, dtype=str, keep_default_na=False) if p.returncode == 0 else None
check("phase 1 on the real September grid: the run is not broken by missing mappings", d1 is not None and len(d1) == 6564, (p.stdout + p.stderr)[-300:])
if d1 is not None:
    w = d1["new_rto_list"].str.contains(B.NEW_MAPPING, regex=False)
    check("phase 1: the rows whose scope has no master group carry the plain-text warning; all the others carry master names", int(w.sum()) > 100 and not d1.loc[~w, "new_rto_list"].str.contains("New changes").any(), int(w.sum()))
    xl = os.path.join(tmp, "from_csv.xlsx")
    code = B.main(["--csv", csv1, "--master-dir", os.path.join(ROOT, "master"), "--out", xl])
    wb = openpyxl.load_workbook(xl)
    added = [list(r) for r in wb.active.iter_rows(values_only=True)][1:]
    check("phase 2 on that CSV: one valid new group per distinct set of RTO codes", code == 0 and len(added) >= 5 and all(B.NAME_OK.fullmatch(n) and B.CLUSTER_OK.fullmatch(c) for n, c in added), len(added))
    # a master workbook = the real one + the new groups
    md = os.path.join(tmp, "newmaster")
    os.makedirs(md)
    shutil.copy(os.path.join(ROOT, "master", "make_model_master.xlsx"), md)
    mw = openpyxl.load_workbook(os.path.join(ROOT, "master", "rto_master.xlsx"))
    ws = mw.worksheets[0]
    last = max((r[0].value or 0) for r in ws.iter_rows(min_row=2) if isinstance(r[0].value, (int, float)))
    for i, (n, c) in enumerate(added, 1):
        ws.append([last + i, n, c])
    mw.save(os.path.join(md, "rto_master.xlsx"))
    csv2 = os.path.join(tmp, "second.csv")
    p2 = converter(md, csv2)
    d2 = pd.read_csv(csv2, dtype=str, keep_default_na=False) if p2.returncode == 0 else None
    left = int(d2["new_rto_list"].str.contains(B.NEW_MAPPING, regex=False).sum()) if d2 is not None else -1
    check("the loop closes: with the new groups in the master the same grid gives no RTO warning and the same number of rows", d2 is not None and left == 0 and len(d2) == len(d1), (left, (p2.stdout + p2.stderr)[-300:]))
    if d2 is not None:
        same = [c for c in d1.columns if c not in ("new_rto_list", "remarks") and not (d1[c] == d2[c]).all()]
        check("... and nothing but the RTO cells (and the remarks about them) changed", not same, same)
        check("... and the cells now hold the new lowercase names", d2["new_rto_list"].str.contains(r"shriram_sep26_", regex=True).any())

# ---- the same loop for the two-wheeler grid: RTO sets and the make group TW_ALL_EX_TVS
G2, B2 = os.path.join(INP, "shriram_2w_oct26.xlsx"), os.path.join(INP, "shriram_2w_oct26_broker-grid.xlsx")


def converter2(master_dir, out_csv):
    args = [sys.executable, "-I", "shriram_2w_to_config.py", "--grid", G2, "--broker-grid", B2, "--from-date", "01/10/2026", "--to-date", "31/03/2030", "--no-prompt", "--out", out_csv]
    return subprocess.run(args, cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900, env=dict(env, **({"GRID_MASTER_DIR": master_dir} if master_dir else {})))


c1 = os.path.join(tmp, "tw_first.csv")
q = converter2(None, c1)
t1 = pd.read_csv(c1, dtype=str, keep_default_na=False) if q.returncode == 0 else None
check("2W phase 1: the run goes on, the make group the master lacks is a warning with its definition", t1 is not None and t1["new_vehicle_makes"].str.contains("TW_ALL_EX_TVS = TW_ALL except TVS", regex=False).any(), (q.stdout + q.stderr)[-300:])
if t1 is not None:
    rx, mx = os.path.join(tmp, "tw_rto.xlsx"), os.path.join(tmp, "tw_make.xlsx")
    B.main(["--csv", c1, "--master-dir", os.path.join(ROOT, "master"), "--out", rx, "--make-out", mx])
    mw = openpyxl.load_workbook(mx)
    mrows2 = [list(r) for r in mw.active.iter_rows(values_only=True)]
    check("2W phase 2: the make Excel has one sheet, the two columns make_model_name / make_model_cluster, and the group TW_ALL minus TVS (lowercase name)",
          mw.sheetnames == ["make_master_additions"] and mrows2[0] == ["make_model_name", "make_model_cluster"] and len(mrows2) == 2 and mrows2[1][0] == "shriram_oct26_tw_all_ex_tvs" and ",TVS XL" not in "," + mrows2[1][1], mrows2[:2])
    md2 = os.path.join(tmp, "newmaster2")
    os.makedirs(md2)
    for src, extra_file in ((os.path.join(ROOT, "master", "rto_master.xlsx"), rx), (os.path.join(ROOT, "master", "make_model_master.xlsx"), mx)):
        w2 = openpyxl.load_workbook(src)
        ws2 = w2.worksheets[0]
        last = max((r[0].value or 0) for r in ws2.iter_rows(min_row=2) if isinstance(r[0].value, (int, float)))
        if os.path.exists(extra_file):
            for i, r in enumerate(list(openpyxl.load_workbook(extra_file).active.iter_rows(min_row=2, values_only=True)), 1):
                ws2.append([last + i, r[0], r[1]])
        w2.save(os.path.join(md2, os.path.basename(src)))
    c2 = os.path.join(tmp, "tw_second.csv")
    q2 = converter2(md2, c2)
    t2 = pd.read_csv(c2, dtype=str, keep_default_na=False) if q2.returncode == 0 else None
    left_rto = int(t2["new_rto_list"].str.contains(B.NEW_MAPPING, regex=False).sum()) if t2 is not None else -1
    undefined = {"KOMAKI", "OLA", "SIMPLE_ENERGY"}                    # new electric makes of the October grid: the master has no model list for them, so no definition can be built
    left_make = {w["name"] for v in t2["new_vehicle_makes"] for w in B.split_cell(v)[1]} if t2 is not None else {"?"}
    check("2W: with the new RTO and make groups in the masters the same grid gives no RTO warning, the TW_ALL_EX_TVS warning is gone (found by its entries), the same number of rows",
          t2 is not None and left_rto == 0 and "TW_ALL_EX_TVS" not in left_make and len(t2) == len(t1), (left_rto, left_make, (q2.stdout + q2.stderr)[-300:]))
    check("... and the only warnings left are the makes nobody can define from the master (their models are not in it): they are listed by master_builder, not built", left_make == undefined, left_make)

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
