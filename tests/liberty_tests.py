"""Tests for grid_automation/liberty_grid_to_config.py ."""
import os as _os
ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))      # the grid_tools folder
import builtins
import datetime as dt
import importlib.util
import os
import tempfile

import openpyxl
import pandas as pd

SCRATCH = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "liberty_grid_to_config.py")
GRID = os.path.join(ROOT, "regression", "inputs", "liberty_sep26.xlsx")
REF = os.path.join(ROOT, "regression", "references", "liberty_sep26_rates_export.csv")
F, T = dt.date(2026, 9, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None

spec = importlib.util.spec_from_file_location("lg", SRC)
lg = importlib.util.module_from_spec(spec)
spec.loader.exec_module(lg)


def edit_copy(name, sheet, fn):
    wb = openpyxl.load_workbook(GRID, data_only=True)
    fn(wb[sheet])
    p = os.path.join(SCRATCH, name)
    wb.save(p)
    return p


def find_row(ws, text, col=1, slab=None, slab_col=2):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == text:
            return r
    raise KeyError(text)


def row_with_slab(ws, geo, slab, geo_col=1):
    r0 = find_row(ws, geo, geo_col)
    for r in range(r0, r0 + 8):
        if ws.cell(r, geo_col + 1).value == slab:
            return r
    raise KeyError((geo, slab))


def col_of(ws, row, text, nth=1):
    n = 0
    for c in range(1, ws.max_column + 1):
        if ws.cell(row, c).value == text:
            n += 1
            if n == nth:
                return c
    raise KeyError(text)


def expect_error(label, fn, needle):
    try:
        fn()
    except lg.GridError as e:
        print(("PASS" if needle in str(e) else "FAIL"), label, "->", str(e)[:160])
    else:
        print("FAIL", label, "-> no error raised")


base, base_w = lg.build(GRID, F, T, quiet)
print("baseline rows", len(base), "| warnings", base_w)
KEY = lg.KEY_COLUMNS


def changed(df):
    m = base.merge(df, on=KEY, suffixes=("", "_n"), how="outer", indicator=True)
    both = m[m["_merge"] == "both"]
    mask = pd.Series(False, index=both.index)
    for c in ["pi_od_rate", "pi_tp_rate", "pi_net_rate", "po_od_rate", "po_tp_rate", "po_net_rate"]:
        mask |= (both[c] - both[c + "_n"]).abs() > 1e-9
    return both[mask], m[m["_merge"] != "both"]


# 1. TW Bike 1+1, DELHI - NCR, slab 50k-1.5L, +1 point
def bump_tw(ws):
    r, c = row_with_slab(ws, "DELHI - NCR", lg.PCTW_SLAB), col_of(ws, 4, "Bike 1+1")
    ws.cell(r, c).value += 0.01
df, w = lg.build(edit_copy("t1.xlsx", lg.PCTW_SHEET, bump_tw), F, T, quiet)
ch, other = changed(df)
ok = len(ch) == 2 and len(other) == 0 and set((ch["pi_net_rate_n"] - ch["pi_net_rate"]).round(6)) == {1.0} and set((ch["po_net_rate_n"] - ch["po_net_rate"]).round(6)) == {1.0}
print("PASS" if ok else "FAIL", "TW Bike 1+1 +1:", len(ch), "rows changed (Petrol+Electric), others", len(other))

# 2. a CV cell of the used slab moves 5 fuels; the other slab changes nothing
def bump_cv(slab):
    def f(ws):
        r = row_with_slab(ws, "BIHAR", slab)
        c = col_of(ws, 4, "Upto 2.5T", 1)
        ws.cell(r, c).value += 0.01
    return f
df, _ = lg.build(edit_copy("t2.xlsx", lg.CV_LOW_SHEET, bump_cv(lg.CV_SLAB)), F, T, quiet)
ch, other = changed(df)
print("PASS" if len(ch) == 5 and len(other) == 0 and set(ch["new_rto_list"]) == {"LIBERTY_GCV_4W_<7.5T_APR26_BIHAR"} and set(ch["sub_product"]) == {"1+1"} else "FAIL", "CV '< 1.5L' cell +1:", len(ch), "rows (5 fuels)")
df, _ = lg.build(edit_copy("t2b.xlsx", lg.CV_LOW_SHEET, bump_cv("> 1.5L")), F, T, quiet)
ch, other = changed(df)
print("PASS" if len(ch) == 0 and len(other) == 0 else "FAIL", "CV '> 1.5L' (unused slab) cell changed:", len(ch), "rows changed")

# 3. grid Private Car columns are not used and do not trigger a warning
def bump_pc(ws):
    r, c = row_with_slab(ws, "DELHI - NCR", lg.PCTW_SLAB), col_of(ws, 4, "PC Comprehensive Petrol")
    ws.cell(r, c).value = 0.99
df, w = lg.build(edit_copy("t3.xlsx", lg.PCTW_SHEET, bump_pc), F, T, quiet)
ch, other = changed(df)
print("PASS" if len(ch) == 0 and len(other) == 0 and w == [] else "FAIL", "PC columns in grid ignored:", len(ch), "changed rows | warnings", w)

# 4. SATP 'NA' -> a number
def na_to_num(ws):
    hdr = next(r for r in range(1, 10) if ws.cell(r, 1).value == "Geo_Segment")
    r = find_row(ws, "ANDHRA PRADESH - 3 R")
    ws.cell(r, col_of(ws, hdr, "PC < 1000 CC", 1)).value = 0.30
df, _ = lg.build(edit_copy("t4.xlsx", lg.PC_SATP_SHEET, na_to_num), F, T, quiet)
hit = df[(df.new_rto_list == "LIBERTY_PC_TW_APR26_ANDHRA_PRADESH_-_3_R") & (df["product"] == "Private Car") & (df.sub_product == "STP") & (df.fuel_type == "Petrol") & (df.cc_max == 1000.01)].iloc[0]
print("PASS" if (hit.pi_tp_rate, hit.po_tp_rate) == (30.0, 23.0) else "FAIL", "PC SATP NA->0.30:", hit.pi_tp_rate, hit.po_tp_rate)

# 5. bad cell values stop the run with the address
def txt_comp(ws):
    ws.cell(row_with_slab(ws, "BIHAR", lg.CV_SLAB), col_of(ws, 4, "GCV 3W", 1)).value = "TBD"
expect_error("text in a comprehensive cell", lambda: lg.build(edit_copy("t5.xlsx", lg.CV_LOW_SHEET, txt_comp), F, T, quiet), "unexpected value 'TBD'")
def na_comp(ws):
    ws.cell(row_with_slab(ws, "BIHAR", lg.CV_SLAB), col_of(ws, 4, "GCV 3W", 1)).value = "NA"
expect_error("NA in a comprehensive cell", lambda: lg.build(edit_copy("t5b.xlsx", lg.CV_LOW_SHEET, na_comp), F, T, quiet), "expected a number")

# 6. layout guards
def rename_header(ws):
    ws.cell(4, col_of(ws, 4, "GCV 3W", 1)).value = "GCV 3W (new)"
expect_error("renamed column", lambda: lg.build(edit_copy("t6.xlsx", lg.CV_LOW_SHEET, rename_header), F, T, quiet), "not found in the header rows")
def rename_slab(ws):
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value == lg.CV_SLAB:
            ws.cell(r, 2).value = "< 2L"
expect_error("slab renamed", lambda: lg.build(edit_copy("t6b.xlsx", lg.CV_LOW_SHEET, rename_slab), F, T, quiet), "not found. Slabs in the sheet")
wb = openpyxl.load_workbook(GRID, data_only=True)
wb["TW SATP"].title = "TW SATP v2"
p0 = os.path.join(SCRATCH, "t6c0.xlsx"); wb.save(p0)
df0, _ = lg.build(p0, F, T, quiet)       # 11 Nov 2026 (structure shifts): a lone renamed sheet is the one sheet with that role's columns (`sheet_by_columns`): found, no error
c0 = changed(df0)
print("PASS" if len(df0) == len(base) and c0[0].empty and c0[1].empty else "FAIL", "sheet renamed (the only sheet with its columns): found by its columns - same rows, same rates")
wb.copy_worksheet(wb["TW SATP v2"]).title = "TW SATP v3"      # a copy left in: two sheets have the columns equally well, which one is this month's is not guessed
p = os.path.join(SCRATCH, "t6c.xlsx"); wb.save(p)
expect_error("sheet renamed, a copy of it left in", lambda: lg.build(p, F, T, quiet), "sheet 'TW SATP' not found")

# 7. a new numeric column nobody mapped -> warning
def new_col(ws):
    c = ws.max_column + 1
    ws.cell(4, c).value = "NEW LOB"
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value:
            ws.cell(r, c).value = 0.10
df, w = lg.build(edit_copy("t7.xlsx", lg.PCTW_SHEET, new_col), F, T, quiet)
print("PASS" if any("NEW LOB" in x for x in w) else "FAIL", "unmapped column warning:", w)

# 8. dates
answers = iter(["1-9-26", "32/13/2026", "", "01/09/2026", "31/03/2030"])
asked = []
f, t = lg.ask_date_range(lambda p: (asked.append(p), next(answers))[1])
print("PASS" if (f, t) == (F, T) and len(asked) == 5 else "FAIL", "date prompt re-asks on bad input:", len(asked), "prompts")
answers = iter(["01/04/2027", "01/09/2026", "01/09/2026", "31/03/2030"])
f, t = lg.ask_date_range(lambda p: next(answers))
print("PASS" if (f, t) == (F, T) else "FAIL", "reversed range is re-asked")
expect_error("to before from", lambda: lg.build(GRID, dt.date(2026, 9, 1), dt.date(2026, 8, 31), quiet), "is before from_date")
df, _ = lg.build(GRID, dt.date(2026, 11, 5), dt.date(2031, 3, 9), quiet)
print("PASS" if set(df["from_date"]) == {"05/11/2026"} and set(df["to_date"]) == {"09/03/2031"} else "FAIL", "dates in every row:", set(df["from_date"]), set(df["to_date"]))

# 9. output file shape
out = lg.main(GRID, F, T, os.path.join(SCRATCH, "lib_final.csv"), None, quiet)
o = pd.read_csv(out, dtype=str, keep_default_na=False)
r = pd.read_csv(REF, dtype=str, keep_default_na=False, nrows=3)
print("PASS" if list(o.columns) == list(r.columns) else "FAIL", "header identical to rates_export:", len(o.columns), "columns")
print("PASS" if set(o["id"]) == {"1"} and not (o == "nan").any().any() and set(o["from_date"]) == {"01/09/2026"} and set(o["to_date"]) == {"31/03/2030"} else "FAIL",
      "id all 1, no 'nan' text, dates", set(o["from_date"]), set(o["to_date"]))
names = set(o["new_rto_list"])
families = ("LIBERTY_PC_TW_APR26_", "LIBERTY_GCV_4W_<7.5T_APR26_", "LIBERTY_GCV_4W>7.5T_APR26_")
print("PASS" if all(n == "ALLINDIA" or n.startswith(families) for n in names) and not any(" " in n for n in names) and "ALLINDIA" in names
      and "LIBERTY_PC_TW_APR26_WEST_BENGAL_-_5_R" in names and "LIBERTY_PC_TW_APR26_CHANDIGARH(TRICITY)" in names and "LIBERTY_GCV_4W>7.5T_APR26_BIHAR" in names
      and "LIBERTY_GCV_4W_<7.5T_APR26_BIHAR" in names else "FAIL",
      "master group names (sheet prefix + geo, spaces as _):", len(names), "distinct, sample", sorted(names)[:4])
pan = o[o.new_rto_list == "ALLINDIA"]
print("PASS" if len(pan) == 10 and set(pan[pan.sub_product == "1+1"]["pi_od_rate"]) == {"27"} and set(pan[pan.sub_product == "SAOD"]["pi_od_rate"]) == {"25"}
      and set(pan[pan.sub_product == "1+1"]["po_od_rate"]) == {"20"} and set(pan[pan.sub_product == "SAOD"]["po_od_rate"]) == {"18"} else "FAIL",
      "Pan India PC rows:", len(pan), "| comp pi/po", set(pan[pan.sub_product == "1+1"][["pi_od_rate", "po_od_rate"]].apply(tuple, axis=1)), "| SAOD", set(pan[pan.sub_product == "SAOD"][["pi_od_rate", "po_od_rate"]].apply(tuple, axis=1)))

# 10. Colab flow with fake upload + typed dates
class FakeFiles:
    downloaded = None
    def upload(self):
        return {os.path.basename(GRID): open(GRID, "rb").read(), os.path.basename(REF): open(REF, "rb").read()}
    def download(self, name):
        if str(name).lower().endswith(".csv"):        # the review workbook is downloaded too (27 Oct 2026): this test reads the upload
            FakeFiles.downloaded = name
typed = iter(["1/9/26", "01/09/2026", "31/03/2030", ""])      # last "" = Enter at "Write the file now? [Y/n]"
builtins_input = builtins.input
builtins.input = lambda prompt="": (print(prompt + (a := next(typed))), a)[1]
spec2 = importlib.util.spec_from_file_location("lg2", SRC)
lg2 = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(lg2)
lg2.colab_files = FakeFiles()
cwd = os.getcwd()
os.chdir(tempfile.mkdtemp(prefix="lib_colab_"))
lg2.default_output = lambda filename, _d=tempfile.mkdtemp(prefix="lib_colab_out_"): os.path.join(_d, filename)   # the test must never write into output (it overwrote the files LIBERTY_UPLOAD_2026-09.csv / MAGMA_UPLOAD_2026-10.csv there)
lg2.colab_main()
c = pd.read_csv(FakeFiles.downloaded, dtype=str, keep_default_na=False)
os.chdir(cwd)
builtins.input = builtins_input
print("PASS" if len(c) == 23690 and set(c["from_date"]) == {"01/09/2026"} else "FAIL", "colab flow:", FakeFiles.downloaded, len(c), "rows")
