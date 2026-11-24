"""Interactive-mode tests for grid_automation/magma_grid_to_config.py .
Every question is answered by a scripted 'keyboard'."""
import os as _os
ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))      # the grid_tools folder
import builtins
import datetime as dt
import importlib.util
import os
import re
import tempfile

import openpyxl
import pandas as pd
from pyxlsb import open_workbook

SCRATCH = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "magma_grid_to_config.py")
GRID = os.path.join(ROOT, "regression", "inputs", "magma_oct26_comprehensive.xlsx")
SATP = os.path.join(ROOT, "regression", "inputs", "magma_oct26_satp.xlsb")
REF = os.path.join(ROOT, "regression", "references", "magma_oct26_reupload.csv")
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None


def load(name):
    spec = importlib.util.spec_from_file_location(name, SRC)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


mg = load("mg")


class Opt:
    def __init__(self, needle):
        self.needle = needle


class Opts:
    """pick several menu options (typed as 3,4,5): one needle per option"""
    def __init__(self, needles):
        self.needles = needles


class Opts:
    """pick several menu options (typed as 3,4,5): one needle per option"""
    def __init__(self, needles):
        self.needles = needles


class Script:
    def __init__(self, steps):
        self.steps, self.menu, self.log = list(steps), [], []

    def out(self, s):
        s = str(s)
        self.log.append(s)
        if s.startswith("\n? "):
            self.menu = []
        else:
            m = re.match(r"\s+(\d+)\) (.*)", s)
            if m:
                self.menu.append((int(m.group(1)), m.group(2)))

    def inp(self, prompt=""):
        if not self.steps:
            raise AssertionError(f"ran out of scripted answers at {prompt!r}; last output:\n" + "\n".join(self.log[-12:]))
        step = self.steps.pop(0)
        if isinstance(step, Opt):
            for n, label in self.menu:
                if step.needle.lower() in label.lower():
                    return str(n)
            raise AssertionError(f"no menu option containing {step.needle!r}; menu = {self.menu}")
        if isinstance(step, Opts):
            nums = []
            for needle in step.needles:
                hit = [n for n, label in self.menu if needle.lower() in label.lower()]
                if len(hit) != 1:
                    raise AssertionError(f"{needle!r} matches {len(hit)} options in {self.menu}")
                nums.append(str(hit[0]))
            return ",".join(nums)
        return step

    def asker(self, mod=mg):
        return mod.Asker(interactive=True, input_fn=self.inp, out=self.out)

    def done(self):
        return not self.steps

    def questions(self):
        return sum(1 for s in self.log if s.startswith("\n? "))


def motor_copy(name, edit):
    wb = openpyxl.load_workbook(GRID, data_only=True)
    edit(wb)
    p = os.path.join(SCRATCH, name)
    wb.save(p)
    return p


def satp_xlsx(name, edit=None):
    with open_workbook(SATP) as wb:
        with wb.get_sheet(mg.SATP_SHEET) as sh:
            rows = [[c.v for c in r] for r in sh.rows()]
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = mg.SATP_SHEET
    for r in rows:
        ws.append(r)
    if edit:
        edit(out)
    p = os.path.join(SCRATCH, name)
    out.save(p)
    return p


def row_of(ws, cluster):
    return next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == cluster)


def col_of(ws, header_row, text):
    return next(c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value == text)


def same(a, b):
    """compare the way the CSV will show them (numbers rounded to 4 decimals): re-saving a workbook with openpyxl adds ~1e-15 noise"""
    drop = lambda d: mg.to_csv_frame(d).drop(columns="id").reset_index(drop=True)      # id is 1..N, so it shifts when rows are added
    return drop(a).equals(drop(b))


def check(ok, label, extra=""):
    print("PASS" if ok else "FAIL", label, extra)


base_satp = satp_xlsx("i_satp_base.xlsx")
base, _ = mg.build(GRID, base_satp, F, T, quiet)
print("baseline rows", len(base))
MOTOR, SATPS = mg.MOTOR_SHEET, mg.SATP_SHEET

# ---------------------------------------------------------------- 1. sheet renamed
def rename_satp_sheet(wb):
    wb[SATPS].title = "Third party grid"
s = Script([Opt("'Third party grid'")])
df, _ = mg.build(GRID, satp_xlsx("i1.xlsx", rename_satp_sheet), F, T, quiet, s.asker())
check(same(df, base) and s.done() and s.questions() == 1, "SATP sheet renamed: pick it -> identical output, 1 question")
s = Script([Opt("leave out")])
df, _ = mg.build(GRID, satp_xlsx("i1b.xlsx", rename_satp_sheet), F, T, quiet, s.asker())
check(len(df) == 23826 and not (df["sub_product"] == "STP").any(), "SATP sheet renamed: leave out -> motor rows only", f"({len(df)} rows)")
def rename_motor_sheet(wb):
    wb[MOTOR].title = "Comprehensive grid"
s = Script([Opt("'Comprehensive grid'")])
df, _ = mg.build(motor_copy("i1c.xlsx", rename_motor_sheet), base_satp, F, T, quiet, s.asker())
check(same(df, base) and s.done(), "motor sheet renamed: pick it -> identical output")

# 11 Nov 2026 (structure shifts, `slab_of`): a sheet whose NAME still says the premium slab in other words ('<=2 Lakhs', '< =3 lakh', 'upto 3 lakh v2') is the sheet - read without a question, said in the decisions
def rename_satp_words(wb):
    wb[SATPS].title = "<=2 Lakhs"
s = Script([])
a = s.asker()
df, _ = mg.build(GRID, satp_xlsx("i1d.xlsx", rename_satp_words), F, T, quiet, a)
check(same(df, base) and s.questions() == 0 and any("premium slab" in d for d in a.decisions), "SATP sheet '<=2 Lakhs' (the slab in other words): used without a question, identical output, the decision is recorded")
def rename_motor_words(wb):
    wb[MOTOR].title = "upto 3 lakh v2"
s = Script([])
a = s.asker()
df, _ = mg.build(motor_copy("i1e.xlsx", rename_motor_words), base_satp, F, T, quiet, a)
check(same(df, base) and s.questions() == 0 and any("premium slab" in d for d in a.decisions), "motor sheet 'upto 3 lakh v2' (the slab in other words): used without a question, identical output")
def two_sheets_one_slab(wb):
    wb[SATPS].title = "<=2 Lakhs"
    wb.create_sheet("upto 2 lakh copy")
s = Script([Opt("'<=2 Lakhs'")])
df, _ = mg.build(GRID, satp_xlsx("i1f.xlsx", two_sheets_one_slab), F, T, quiet, s.asker())
check(same(df, base) and s.done() and s.questions() == 1, "two sheets that say the same slab are NOT guessed: one question, the person picks the sheet, identical output")

# ---------------------------------------------------------------- 2. SATP column missing
def hide_tp3w(wb):
    ws = wb[SATPS]
    ws.cell(3, col_of(ws, 3, "TP GCV 3W")).value = "TP GCV 3W (renamed)"
p = satp_xlsx("i2.xlsx", hide_tp3w)
g3 = lambda d: d[(d["product"] == "GCV 3W") & (d["sub_product"] == "STP")]
s = Script([Opt("rate 0")])
df, _ = mg.build(GRID, p, F, T, quiet, s.asker())
check(len(df) == len(base) and len(g3(df)) == 570 and (g3(df)["pi_tp_rate"] == 0).all() and (g3(df)["po_tp_rate"] == 0).all() and s.done(),
      "SATP column missing -> rate 0: same row count, GCV 3W STP all 0")
s = Script([Opt("leave these rows out")])
df, _ = mg.build(GRID, p, F, T, quiet, s.asker())
check(len(df) == len(base) - 570 and len(g3(df)) == 0, "SATP column missing -> leave out: 570 fewer rows")
s = Script([Opt("every column"), Opt("TP GCV <=2.5 T")])
a = s.asker()
df, w = mg.build(GRID, p, F, T, quiet, a)
want = base[(base["product"] == "GCV 4W") & (base["sub_product"] == "STP") & (base["cc_max"] == 2500.01) & (base["fuel_type"] == "Petrol")].set_index("new_rto_list")["pi_tp_rate"]
got = g3(df)[g3(df)["fuel_type"] == "Petrol"].set_index("new_rto_list")["pi_tp_rate"]
check(len(df) == len(base) and (got == want.reindex(got.index)).all(), "SATP column missing -> use another column: GCV 3W STP takes that column's rates", f"warnings={w}")

# ---------------------------------------------------------------- 3. motor column missing
def hide_gcv3w(wb):
    ws = wb[MOTOR]
    ws.cell(3, col_of(ws, 3, "GCV 3W")).value = "GCV 3W (renamed)"
p = motor_copy("i3.xlsx", hide_gcv3w)
m3 = lambda d: d[(d["product"] == "GCV 3W") & (d["sub_product"] == "1+1")]
s = Script([Opt("rate 0")])
df, _ = mg.build(p, base_satp, F, T, quiet, s.asker())
check(len(df) == len(base) and len(m3(df)) == 570 and (m3(df)["pi_net_rate"] == 0).all() and (m3(df)["pi_type"] == "On Net").all(), "motor column missing -> rate 0")
s = Script([Opt("leave these rows out")])
df, _ = mg.build(p, base_satp, F, T, quiet, s.asker())
check(len(df) == len(base) - 570 and len(m3(df)) == 0, "motor column missing -> leave out: 570 fewer rows")

# ---------------------------------------------------------------- 4. repeated cluster
def dup_cluster(wb):
    ws = wb[MOTOR]
    ws.cell(row_of(ws, "BR"), 1).value = "AP"
p = motor_copy("i4.xlsx", dup_cluster)
s = Script([Opt("first row")])
df, _ = mg.build(p, base_satp, F, T, quiet, s.asker())
ap = lambda d: d[(d["new_rto_list"] == "MAGMA_APR26_AP") & (d["product"] == "GCV 4W") & (d["sub_product"] == "1+1") & (d["cc_max"] == 2500.01) & (d["fuel_type"] == "Petrol") & (d["vehicle_age_min"] == 0)]["pi_net_rate"].tolist()
br = lambda d: d[(d["new_rto_list"] == "MAGMA_APR26_BR") & (d["product"] == "GCV 4W") & (d["sub_product"] == "1+1") & (d["cc_max"] == 2500.01) & (d["fuel_type"] == "Petrol") & (d["vehicle_age_min"] == 0)]["pi_net_rate"].tolist()
check(len(df) == len(base) - 209 and len(ap(base)) > 0 and len(br(base)) > 0 and ap(df) == ap(base), "repeated cluster -> first row: AP keeps its own rates, BR motor rows gone")
s = Script([Opt("last row")])
df, _ = mg.build(p, base_satp, F, T, quiet, s.asker())
check(len(df) == len(base) - 209 and ap(df) == br(base) and ap(df) != ap(base), "repeated cluster -> last row: AP takes the later row's rates")
try:
    mg.build(p, base_satp, F, T, quiet)
    check(False, "repeated cluster without a keyboard")
except mg.GridError as e:
    check("duplicate cluster names" in str(e), "repeated cluster, no Asker: stops with the old message")

# ---------------------------------------------------------------- 5. header row lost
uw_row = None
def lose_uw(wb):
    global uw_row
    ws = wb[MOTOR]
    uw_row = next(r for r in range(1, 12) if str(ws.cell(r, 1).value or "").upper().startswith("UW CLUSTER"))
    ws.cell(uw_row, 1).value = "Clusters"
p = motor_copy("i5.xlsx", lose_uw)
s = Script(["abc", "2", str(uw_row)])
a = s.asker()
df, _ = mg.build(p, base_satp, F, T, quiet, a)
check(same(df, base) and s.done() and any("header row" in d for d in a.decisions), f"'UW Cluster' row lost: bad answers re-asked, row {uw_row} accepted -> identical output")

# ---------------------------------------------------------------- 6. unreadable cells (motor and SATP)
def bad_motor(wb):
    ws = wb[MOTOR]
    ws.cell(row_of(ws, "BR"), col_of(ws, 3, "GCV 3W")).value = "TBD"
s = Script([Opt("Use 0")])
a = s.asker()
df, _ = mg.build(motor_copy("i6.xlsx", bad_motor), base_satp, F, T, quiet, a)
br3 = df[(df["new_rto_list"] == "BR") & (df["product"] == "GCV 3W") & (df["sub_product"] == "1+1")]
check(len(df) == len(base) and (br3["pi_net_rate"] == 0).all() and len(a.tallies["cells that were not rates"]) == 1, "unreadable motor cell -> use 0", f"({len(br3)} rows)")
def bad_satp(wb):
    ws = wb[SATPS]
    c = col_of(ws, 3, "TP GCV 3W")
    ws.cell(row_of(ws, "BR"), c).value = "n/a"
    ws.cell(row_of(ws, "AP"), c).value = "n/a"
s = Script([Opt("Use 0"), "y"])           # the 2nd 'n/a': "same kind of problem - do that for every other cell like it?" -> yes
a = s.asker()
df, _ = mg.build(GRID, satp_xlsx("i6b.xlsx", bad_satp), F, T, quiet, a)
check(s.done() and s.questions() == 1 and len(a.tallies["cells that were not rates"]) == 2 and len(df) == len(base), "two unreadable SATP cells + 'same for every cell like it' -> one menu, one yes/no")

# ---------------------------------------------------------------- 7. which file is which
neutral_a = motor_copy("file_one.xlsx", lambda wb: None)
neutral_b = satp_xlsx("file_two.xlsx")
s = Script([])
check(mg.assign_files([neutral_b, neutral_a], s.asker()) == (neutral_a, neutral_b) and s.questions() == 0, "files told apart by their sheets, whichever order and whatever the names")
twin = motor_copy("file_three.xlsx", lambda wb: None)
s = Script([Opt("file_three.xlsx")])
g, st = mg.assign_files([neutral_a, twin], s.asker())
check(g == twin and st == neutral_a and s.questions() == 1, "two files that look alike: asks which is the comprehensive grid")
try:
    mg.assign_files([neutral_a], Script([]).asker())
    check(False, "one file only")
except mg.GridError as e:
    check("exactly two" in str(e), "only one file given: stops saying two are needed")

# ---------------------------------------------------------------- 8. confirmation, q, EOF
out = os.path.join(SCRATCH, "i8_out.csv")
if os.path.exists(out):
    os.remove(out)
try:
    mg.main(GRID, base_satp, F, T, out, None, quiet, Script(["n"]).asker())
    check(False, "declining the confirmation")
except mg.GridError as e:
    check("nothing was written" in str(e) and not os.path.exists(out), "n at 'write the file now?' writes nothing")
s = Script([""])
mg.main(GRID, base_satp, F, T, out, None, quiet, s.asker())
check(os.path.exists(out) and len(pd.read_csv(out)) == 39102 and s.questions() == 0, "clean grids in interactive mode: no questions, Enter writes the file", "(39,102 rows)")
try:
    mg.build(GRID, satp_xlsx("i8q.xlsx", rename_satp_sheet), F, T, quiet, Script(["q"]).asker())
    check(False, "q stops")
except mg.GridError as e:
    check("stopped by you" in str(e), "q at a question stops the run")
def eof(prompt=""):
    raise EOFError
try:
    mg.build(GRID, satp_xlsx("i8e.xlsx", rename_satp_sheet), F, T, quiet, mg.Asker(interactive=True, input_fn=eof, out=quiet))
    check(False, "EOF")
except mg.GridError as e:
    check("no keyboard input" in str(e), "closed keyboard (EOF) stops with a clear message")
try:
    mg.build(GRID, satp_xlsx("i8n.xlsx", rename_satp_sheet), F, T, quiet)
    check(False, "no Asker")
except mg.NeedsAnswer as e:
    check("sheet '<=2 Lacs' not found" in str(e), "no Asker given: stops with the old message")

# ---------------------------------------------------------------- 10. several columns at once: Electric PCV 3W takes the New and Old columns with their age bands
BASE_TNC = mg.BASE_TNC
def hide_tp_pcv3w_electric(wb):
    ws = wb[SATPS]
    ws.cell(3, col_of(ws, 3, "TP-PCV-3W Electric")).value = "TP-PCV-3W Electric (renamed)"
p = satp_xlsx("i10.xlsx", hide_tp_pcv3w_electric)
pcv_stp = lambda d, fuel: d[(d["product"] == "PCV 3W") & (d["sub_product"] == "STP") & (d["fuel_type"] == fuel)]
s = Script([Opt("every column"), Opts(["TP-PCV-3W New", "TP-PCV-3W Old"])])
a = s.asker()
df, w = mg.build(GRID, p, F, T, quiet, a)
ele = pcv_stp(df, "Electric")
new_r = pcv_stp(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_tp_rate"]
old_r = pcv_stp(base, "Petrol")[lambda d: d["vehicle_age_min"] == 1].set_index("new_rto_list")["pi_tp_rate"]
e_new = ele[ele["vehicle_age_max"] == 1].set_index("new_rto_list")
e_old = ele[ele["vehicle_age_min"] == 1].set_index("new_rto_list")
not_e = lambda d: d[~((d["product"] == "PCV 3W") & (d["sub_product"] == "STP") & (d["fuel_type"] == "Electric"))]
check(len(ele) == 228 and len(df) == len(base) - 114 + 228 and (e_new["pi_tp_rate"] == new_r.reindex(e_new.index)).all()
      and (e_old["pi_tp_rate"] == old_r.reindex(e_old.index)).all() and set(e_new["add_tnc"]) == {"New," + BASE_TNC} and set(e_old["add_tnc"]) == {"Old," + BASE_TNC}
      and same(not_e(df), not_e(base)) and any("Electric (renamed)" in x and "NOT in the output" in x for x in w) and s.done()
      and any("one set of rows per column" in d for d in a.decisions),
      "SATP Electric column missing -> New + Old picked together: Electric copies both, with their age bands and terms; nothing else changes", f"({len(ele)} Electric rows)")
def hide_motor_pcv3w_electric(wb):
    ws = wb[MOTOR]
    ws.cell(3, col_of(ws, 3, "PCV 3W Electric")).value = "PCV 3W Electric (renamed)"
p = motor_copy("i10b.xlsx", hide_motor_pcv3w_electric)
pcv_m = lambda d, fuel: d[(d["product"] == "PCV 3W") & (d["sub_product"] == "1+1") & (d["fuel_type"] == fuel)]
s = Script([Opt("every column"), Opts(["PCV 3W New", "PCV 3W Old"])])
df, w = mg.build(p, base_satp, F, T, quiet, s.asker())
ele = pcv_m(df, "Electric")
new_r = pcv_m(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_net_rate"]
old_r = pcv_m(base, "Petrol")[lambda d: d["vehicle_age_min"] == 1].set_index("new_rto_list")["pi_net_rate"]
e_new, e_old = ele[ele["vehicle_age_max"] == 1].set_index("new_rto_list"), ele[ele["vehicle_age_min"] == 1].set_index("new_rto_list")
not_m = lambda d: d[~((d["product"] == "PCV 3W") & (d["sub_product"] == "1+1") & (d["fuel_type"] == "Electric"))]
check(len(ele) == 228 and len(df) == len(base) - 114 + 228 and (e_new["pi_net_rate"].round(4) == new_r.reindex(e_new.index).round(4)).all()
      and (e_old["pi_net_rate"].round(4) == old_r.reindex(e_old.index).round(4)).all() and (e_new["pi_type"] == pcv_m(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_type"].reindex(e_new.index)).all()      # IRDA cells stay On OD
      and same(not_m(df), not_m(base)) and s.done(),
      "motor Electric column missing -> New + Old picked together: Electric copies both with their age bands", f"({len(ele)} Electric rows)")

# the typed forms 1,2 / 1-2 work in the list menu too (New and Old are not the first two columns, so only check the parsing here)
for text, want in (("3,4,5", [3, 4, 5]), ("3-5", [3, 4, 5]), (" 1 , 3 - 4 ", [1, 3, 4]), ("2,2,1", [2, 1]), ("5;6", [5, 6])):
    check(mg.parse_picks(text, 11) == want, f"parse_picks({text!r}) == {want}")
for bad in ("", "0", "12", "3-1", "a", "1,,x", "-2"):
    try:
        mg.parse_picks(bad, 11)
        check(False, f"parse_picks({bad!r}) refused")
    except ValueError:
        check(True, f"parse_picks({bad!r}) refused")

# refused: a column that cannot be mirrored (two segments share its label) or mixing in 'rate 0'
s = Script([Opt("every column"), Opts(["TP GCV 20T-40T Age<5", "TP-PCV-3W New"]), Opts(["TP-PCV-3W New", "TP-PCV-3W Old"])])
p2 = satp_xlsx("i10c.xlsx", hide_tp_pcv3w_electric)
df, _ = mg.build(GRID, p2, F, T, quiet, s.asker())
check(any("age / cc bands" in l for l in s.log) and len(pcv_stp(df, "Electric")) == 228 and s.done(), "a column shared by two segments can't be combined: refused, then re-asked")
s = Script([Opts(["every column", "rate 0"]), Opt("rate 0")])
df, _ = mg.build(GRID, p2, F, T, quiet, s.asker())
check(any("can only be picked on their own" in l for l in s.log) and len(pcv_stp(df, "Electric")) == 114 and s.done(), "'every column' + 'rate 0' together: refused, then re-asked")
# a menu that takes one answer refuses commas
s = Script(["1,2", Opt("'Third party grid'")])
df, _ = mg.build(GRID, satp_xlsx("i10d.xlsx", rename_satp_sheet), F, T, quiet, s.asker())
check(same(df, base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "sheet question (one answer): '1,2' refused, then one number works")

# ---------------------------------------------------------------- 10. several columns at once: Electric PCV 3W takes the New and Old columns with their age bands
BASE_TNC = mg.BASE_TNC
def hide_tp_pcv3w_electric(wb):
    ws = wb[SATPS]
    ws.cell(3, col_of(ws, 3, "TP-PCV-3W Electric")).value = "TP-PCV-3W Electric (renamed)"
p = satp_xlsx("i10.xlsx", hide_tp_pcv3w_electric)
pcv_stp = lambda d, fuel: d[(d["product"] == "PCV 3W") & (d["sub_product"] == "STP") & (d["fuel_type"] == fuel)]
s = Script([Opt("every column"), Opts(["TP-PCV-3W New", "TP-PCV-3W Old"])])
a = s.asker()
df, w = mg.build(GRID, p, F, T, quiet, a)
ele = pcv_stp(df, "Electric")
new_r = pcv_stp(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_tp_rate"]
old_r = pcv_stp(base, "Petrol")[lambda d: d["vehicle_age_min"] == 1].set_index("new_rto_list")["pi_tp_rate"]
e_new = ele[ele["vehicle_age_max"] == 1].set_index("new_rto_list")
e_old = ele[ele["vehicle_age_min"] == 1].set_index("new_rto_list")
not_e = lambda d: d[~((d["product"] == "PCV 3W") & (d["sub_product"] == "STP") & (d["fuel_type"] == "Electric"))]
check(len(ele) == 228 and len(df) == len(base) - 114 + 228 and (e_new["pi_tp_rate"] == new_r.reindex(e_new.index)).all()
      and (e_old["pi_tp_rate"] == old_r.reindex(e_old.index)).all() and set(e_new["add_tnc"]) == {"New," + BASE_TNC} and set(e_old["add_tnc"]) == {"Old," + BASE_TNC}
      and same(not_e(df), not_e(base)) and any("Electric (renamed)" in x and "NOT in the output" in x for x in w) and s.done()
      and any("one set of rows per column" in d for d in a.decisions),
      "SATP Electric column missing -> New + Old picked together: Electric copies both, with their age bands and terms; nothing else changes", f"({len(ele)} Electric rows)")
def hide_motor_pcv3w_electric(wb):
    ws = wb[MOTOR]
    ws.cell(3, col_of(ws, 3, "PCV 3W Electric")).value = "PCV 3W Electric (renamed)"
p = motor_copy("i10b.xlsx", hide_motor_pcv3w_electric)
pcv_m = lambda d, fuel: d[(d["product"] == "PCV 3W") & (d["sub_product"] == "1+1") & (d["fuel_type"] == fuel)]
s = Script([Opt("every column"), Opts(["PCV 3W New", "PCV 3W Old"])])
df, w = mg.build(p, base_satp, F, T, quiet, s.asker())
ele = pcv_m(df, "Electric")
new_r = pcv_m(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_net_rate"]
old_r = pcv_m(base, "Petrol")[lambda d: d["vehicle_age_min"] == 1].set_index("new_rto_list")["pi_net_rate"]
e_new, e_old = ele[ele["vehicle_age_max"] == 1].set_index("new_rto_list"), ele[ele["vehicle_age_min"] == 1].set_index("new_rto_list")
not_m = lambda d: d[~((d["product"] == "PCV 3W") & (d["sub_product"] == "1+1") & (d["fuel_type"] == "Electric"))]
check(len(ele) == 228 and len(df) == len(base) - 114 + 228 and (e_new["pi_net_rate"].round(4) == new_r.reindex(e_new.index).round(4)).all()
      and (e_old["pi_net_rate"].round(4) == old_r.reindex(e_old.index).round(4)).all() and (e_new["pi_type"] == pcv_m(base, "Petrol")[lambda d: d["vehicle_age_max"] == 1].set_index("new_rto_list")["pi_type"].reindex(e_new.index)).all()      # IRDA cells stay On OD
      and same(not_m(df), not_m(base)) and s.done(),
      "motor Electric column missing -> New + Old picked together: Electric copies both with their age bands", f"({len(ele)} Electric rows)")

# the typed forms 1,2 / 1-2 work in the list menu too (New and Old are not the first two columns, so only check the parsing here)
for text, want in (("3,4,5", [3, 4, 5]), ("3-5", [3, 4, 5]), (" 1 , 3 - 4 ", [1, 3, 4]), ("2,2,1", [2, 1]), ("5;6", [5, 6])):
    check(mg.parse_picks(text, 11) == want, f"parse_picks({text!r}) == {want}")
for bad in ("", "0", "12", "3-1", "a", "1,,x", "-2"):
    try:
        mg.parse_picks(bad, 11)
        check(False, f"parse_picks({bad!r}) refused")
    except ValueError:
        check(True, f"parse_picks({bad!r}) refused")

# refused: a column that cannot be mirrored (two segments share its label) or mixing in 'rate 0'
s = Script([Opt("every column"), Opts(["TP GCV 20T-40T Age<5", "TP-PCV-3W New"]), Opts(["TP-PCV-3W New", "TP-PCV-3W Old"])])
p2 = satp_xlsx("i10c.xlsx", hide_tp_pcv3w_electric)
df, _ = mg.build(GRID, p2, F, T, quiet, s.asker())
check(any("age / cc bands" in l for l in s.log) and len(pcv_stp(df, "Electric")) == 228 and s.done(), "a column shared by two segments can't be combined: refused, then re-asked")
s = Script([Opts(["every column", "rate 0"]), Opt("rate 0")])
df, _ = mg.build(GRID, p2, F, T, quiet, s.asker())
check(any("can only be picked on their own" in l for l in s.log) and len(pcv_stp(df, "Electric")) == 114 and s.done(), "'every column' + 'rate 0' together: refused, then re-asked")
# a menu that takes one answer refuses commas
s = Script(["1,2", Opt("'Third party grid'")])
df, _ = mg.build(GRID, satp_xlsx("i10d.xlsx", rename_satp_sheet), F, T, quiet, s.asker())
check(same(df, base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "sheet question (one answer): '1,2' refused, then one number works")

# ---------------------------------------------------------------- 11. spaces and capitals in header labels do not matter
def respace_motor(wb):
    ws = wb[MOTOR]
    for r in (2, 3):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v in ("75-150cc", "150-350cc", "GCV 3W", "PCV 3W New", "Tractor Old"):
                ws.cell(r, c).value = v.replace("-", " - ").upper()
s = Script([])
df, w = mg.build(motor_copy("i11.xlsx", respace_motor), base_satp, F, T, quiet, s.asker())
check(same(df, base) and s.questions() == 0 and w == [], "headers with other spaces / capitals ('75 - 150CC', 'GCV 3W' upper-case) -> matched, no question, identical output")

# ---------------------------------------------------------------- 11. spaces and capitals in header labels do not matter
def respace_motor(wb):
    ws = wb[MOTOR]
    for r in (2, 3):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v in ("75-150cc", "150-350cc", "GCV 3W", "PCV 3W New", "Tractor Old"):
                ws.cell(r, c).value = v.replace("-", " - ").upper()
s = Script([])
df, w = mg.build(motor_copy("i11.xlsx", respace_motor), base_satp, F, T, quiet, s.asker())
check(same(df, base) and s.questions() == 0 and w == [], "headers with other spaces / capitals ('75 - 150CC', 'GCV 3W' upper-case) -> matched, no question, identical output")

# ---------------------------------------------------------------- 9. Colab flow: files uploaded in the 'wrong' order with neutral names, a sheet renamed, dates, confirm
class FakeFiles:
    downloaded = None
    def upload(self):
        return {"b.xlsx": open(satp_xlsx("i9_satp.xlsx", rename_satp_sheet), "rb").read(),
                "a.xlsx": open(GRID, "rb").read(),
                "last month.csv": open(REF, "rb").read()}
    def download(self, name):
        FakeFiles.downloaded = name
sc = Script(["1-10-26", "01/10/2026", "31/03/2030", Opt("'Third party grid'"), ""])      # dates first, then the sheet question while reading, then Enter to confirm
real_input, real_print = builtins.input, builtins.print
builtins.input, builtins.print = sc.inp, lambda *a, **k: sc.out(" ".join(map(str, a)))
try:
    mg2 = load("mg2")
    mg2.colab_files = FakeFiles()
    cwd = os.getcwd()
    os.chdir(tempfile.mkdtemp(prefix="mag_colab_i_"))
    mg2.default_output = lambda filename, _d=tempfile.mkdtemp(prefix="mag_colab_out_"): os.path.join(_d, filename)   # the test must never write into output (it overwrote the files LIBERTY_UPLOAD_2026-09.csv / MAGMA_UPLOAD_2026-10.csv there)
    mg2.colab_main()
    c = pd.read_csv(FakeFiles.downloaded, dtype=str, keep_default_na=False)
    os.chdir(cwd)
finally:
    builtins.input, builtins.print = real_input, real_print
diff_line = next((l for l in sc.log if "vs last month.csv" in l), "")
check(len(c) == 39102 and set(c["from_date"]) == {"01/10/2026"} and sc.done() and "0 new rows, 0 removed rows, 0 rows with changed values" in diff_line,
      "Colab flow: files sorted out by content, sheet question answered, dates typed, downloaded, 0 differences vs last month's upload", f"({len(c)} rows)")
