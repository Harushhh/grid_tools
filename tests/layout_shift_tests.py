"""Structure-shift checks (the user, 30 Oct 2026: insurers move rows, columns and sheet names around; the extraction must find tables by what they SAY, not where they are).
Part 1 - the locator helpers of the shared block on small in-memory sheets.
Part 2 - the real October 2026 Shriram grids (broker and two-wheeler) are shifted the way an insurer might (rows above, columns inserted or swapped, sheets renamed, title rows) and
the converter must write the SAME CSV as for the unshifted grid. Merged cells move with the cells, as they do in Excel.
Run:  .venv\\Scripts\\python.exe tests\\layout_shift_tests.py"""
import os
import subprocess
import sys
import tempfile

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

ROOT = helpers.ROOT
PY = sys.executable
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


# ================================================================ part 1: the helpers
M = helpers.load_script("shriram_broker_grid_to_config.py")
check("head_match: the same words, any case, spaces and punctuation", M.head_match("RTO  Code", "rto code") and M.head_match("Product ", "product") and M.head_match("RTO-STATE", "rto state"))
check("head_match: a synonym's words inside a short cell ('RTO State' says 'state'), never inside a long sentence",
      M.head_match("RTO State", "state") and not M.head_match("the statement of the state of affairs this month", "state") and not M.head_match("Statement date", "state"))
check("head_match: other words do not match", not M.head_match("RTO Name", "rto code") and not M.head_match(None, "product") and not M.head_match("", "product"))
rows = [["Declined and Excluded RTOS (", None, None], [None, None, None], [None, "Product", "RTO Code"], [None, "GCCV 3W", "UP-93"]]
check("locate_table: finds the header under title rows and blank rows, in any column", M.locate_table(rows, {"product": ["product"], "code": ["rto code"]}, ("product", "code")) == (2, {"product": 1, "code": 2}))
check("locate_table: None when a required header is missing", M.locate_table(rows, {"product": ["product"], "code": ["rto code"], "x": ["premium"]}, ("product", "x")) is None)
wb = helpers.FakeWB({"Rates": [["a", "b"], [1, 2]], "Whatever it is called": rows, "Other": [["Product", "RTO Code"], ["TW", "no codes here"], ["TW", "none"]]})
found = M.find_tables(wb, M.DECLINED_FIELDS, M.DECLINED_REQUIRED, M._has_codes("code"))
check("find_tables: the sheet is found by its headers, whatever its name; a table whose 'code' column holds no codes is refused", [f[0] for f in found] == ["Whatever it is called"], [f[0] for f in found])
w = []
top = M.pick_table(found, "Declined-Excluded", "the declined RTOs", None, w, lambda *a, **k: None)
check("pick_table: a table on an unusual sheet is used AND reported in the warnings", top and top[0] == "Whatever it is called" and any("found by its column headers" in x for x in w), (top and top[0], w))
check("pick_table: nothing found gives None", M.pick_table([], "x", "y") is None)
check("sheet_like: the usual name, spaces and capitals ignored; else the only sheet that has the words; two candidates give None",
      M.sheet_like(helpers.FakeWB({"Metro  RTO codes": []}), "Metro RTO Codes", ("metro",)) == "Metro  RTO codes"
      and M.sheet_like(helpers.FakeWB({"Metro list": [], "Zone": []}), "Metro RTO Codes", ("metro",)) == "Metro list"
      and M.sheet_like(helpers.FakeWB({"Metro A": [], "Metro B": []}), "Metro RTO Codes", ("metro",)) is None)
check("the master export is read by its headers: a name / list column order change is found", True)
tmp = tempfile.mkdtemp(prefix="layoutshift-")
mp = os.path.join(tmp, "rto_master.xlsx")
w2 = openpyxl.Workbook()
w2.active.append(["RTO CLUSTER", "SOMETHING", "RTO NAME"])
w2.active.append(["MH01,MH02", "x", "Mumbai"])
w2.save(mp)
check("_read_master: columns found by header (cluster list before the name)", M._read_master(mp) == {"Mumbai": ["MH01", "MH02"]}, M._read_master(mp))


# ================================================================ part 2: shifted workbooks
def insert(ws, kind, idx, amount=1):
    """insert rows / columns the way Excel does: merged cells move with the cells (openpyxl alone leaves them where they were and loses the merged text)"""
    merged = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in list(ws.merged_cells.ranges)]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    (ws.insert_rows if kind == "rows" else ws.insert_cols)(idx, amount=amount)
    for r0, c0, r1, c1 in merged:
        if kind == "rows":
            if r0 >= idx:
                r0, r1 = r0 + amount, r1 + amount
            elif r1 >= idx:
                r1 += amount
        elif c0 >= idx:
            c0, c1 = c0 + amount, c1 + amount
        elif c1 >= idx:
            c1 += amount
        ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)


def swap_cols(ws, a, b):
    for r in range(1, ws.max_row + 1):
        va, vb = ws.cell(r, a).value, ws.cell(r, b).value
        ws.cell(r, a).value, ws.cell(r, b).value = vb, va


def run(script, args, out):
    p = subprocess.run([PY, "-I", script, *args, "--from-date", "01/10/2026", "--to-date", "31/03/2030", "--no-prompt", "--out", out], cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=900, env=dict(os.environ, PYTHONIOENCODING="utf-8", GRID_RTO_NAMES="cluster"))
    return p, (pd.read_csv(out, dtype=str, keep_default_na=False) if p.returncode == 0 and os.path.exists(out) else None)


INP = os.path.join(ROOT, "regression", "inputs")
RATE, DECL, PAN, METRO = "GRID OCT' 26", "Declined-Excluded", "PAN India", "Metro RTO Codes"
BROKER_SHIFTS = [
    ("rows above the rate table", lambda wb: insert(wb[RATE], "rows", 1, 6)),
    ("columns left of the rate table", lambda wb: insert(wb[RATE], "cols", 1, 2)),
    ("a column inserted inside the rate table", lambda wb: insert(wb[RATE], "cols", 4)),
    ("two rate columns swapped", lambda wb: swap_cols(wb[RATE], 2, 3)),
    ("the rate sheet renamed", lambda wb: setattr(wb[RATE], "title", "Rates for October 2026")),
    ("the side sheets renamed", lambda wb: (setattr(wb[DECL], "title", "Declined RTOs and exclusions"), setattr(wb[PAN], "title", "Pan India declines"), setattr(wb[METRO], "title", "Metro codes"))),
    ("a column inserted in the side sheets", lambda wb: (insert(wb[DECL], "cols", 1), insert(wb[PAN], "cols", 1))),
    ("title rows above the side tables", lambda wb: (insert(wb[DECL], "rows", 1), insert(wb[PAN], "rows", 1, 2))),
    ("everything at once", lambda wb: (insert(wb[RATE], "rows", 1, 6), insert(wb[RATE], "cols", 1, 2), insert(wb[DECL], "cols", 1), insert(wb[PAN], "rows", 1, 2), setattr(wb[METRO], "title", "Metro codes"))),
]
out0 = os.path.join(tmp, "broker_base.csv")
p, base = run("shriram_broker_grid_to_config.py", ["--grid", os.path.join(INP, "shriram_oct26.xlsx")], out0)
check("Shriram broker: the unshifted October grid runs", base is not None and len(base) > 6000, (p.stdout + p.stderr)[-300:])
if base is not None:
    for label, fn in BROKER_SHIFTS:
        wb = openpyxl.load_workbook(os.path.join(INP, "shriram_oct26.xlsx"), data_only=True)
        fn(wb)
        g = os.path.join(tmp, "shifted.xlsx")
        wb.save(g)
        o = os.path.join(tmp, "shifted.csv")
        if os.path.exists(o):
            os.remove(o)
        p, got = run("shriram_broker_grid_to_config.py", ["--grid", g], o)
        check(f"Shriram broker: {label} -> the same CSV", got is not None and got.equals(base),
              (p.stdout + p.stderr)[-250:] if got is None else f"{len(got)} rows vs {len(base)}; columns that differ: {[c for c in base.columns if len(got) != len(base) or not (got[c] == base[c]).all()][:5]}")

# the two-wheeler script reads the broker grid's PAN India sheet (--broker-grid): shift that file
G2 = os.path.join(INP, "shriram_2w_oct26.xlsx")
B2 = os.path.join(INP, "shriram_2w_oct26_broker-grid.xlsx")
out0 = os.path.join(tmp, "tw_base.csv")
p, base2 = run("shriram_2w_to_config.py", ["--grid", G2, "--broker-grid", B2], out0)
check("Shriram 2W: the unshifted October grid runs", base2 is not None and len(base2) > 400, (p.stdout + p.stderr)[-300:])
if base2 is not None:
    TW_SHIFTS = [
        ("rows above and a column left of the New Business sheet", G2, lambda wb: (insert(wb["New Business - 2W"], "rows", 1, 5), insert(wb["New Business - 2W"], "cols", 1, 2))),
        ("a column inserted in the Rollover sheet", G2, lambda wb: insert(wb["ROLLOVER 2W (1+1)"], "cols", 3)),
        ("the Rollover sheet renamed", G2, lambda wb: setattr(wb["ROLLOVER 2W (1+1)"], "title", "Rollover two wheelers")),
        ("the broker grid's PAN India sheet renamed and given a column and title rows", B2, lambda wb: (insert(wb[PAN], "cols", 1), insert(wb[PAN], "rows", 1, 2), setattr(wb[PAN], "title", "Pan India declines"))),
    ]
    for label, src, fn in TW_SHIFTS:
        wb = openpyxl.load_workbook(src, data_only=True)
        fn(wb)
        g = os.path.join(tmp, "tw_shifted.xlsx")
        wb.save(g)
        o = os.path.join(tmp, "tw_shifted.csv")
        if os.path.exists(o):
            os.remove(o)
        args = ["--grid", g, "--broker-grid", B2] if src == G2 else ["--grid", G2, "--broker-grid", g]
        p, got = run("shriram_2w_to_config.py", args, o)
        check(f"Shriram 2W: {label} -> the same CSV", got is not None and got.equals(base2),
              (p.stdout + p.stderr)[-250:] if got is None else f"{len(got)} rows vs {len(base2)}; columns that differ: {[c for c in base2.columns if len(got) != len(base2) or not (got[c] == base2[c]).all()][:5]}")

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
