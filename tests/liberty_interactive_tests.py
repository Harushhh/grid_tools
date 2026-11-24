"""Interactive-mode tests for grid_automation/liberty_grid_to_config.py .
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

SCRATCH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib_i")      # own folder: the Magma tests use the same file names
os.makedirs(SCRATCH, exist_ok=True)
SRC =os.path.join(ROOT, "liberty_grid_to_config.py")
SRC_M = os.path.join(ROOT, "magma_grid_to_config.py")
GRID = os.path.join(ROOT, "regression", "inputs", "liberty_sep26.xlsx")
JUL_B = os.path.join(ROOT, "regression", "inputs", "liberty_jul26.xlsx")
JUL_A = os.path.join(ROOT, "regression", "inputs", "liberty_jul26_v1.xlsx")
F, T = dt.date(2026, 9, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


lg = load("lg", SRC)


class Opt:
    """pick the menu option whose text contains `needle`"""
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
    """A scripted keyboard: literal strings are typed as they are, Opt(...) picks a menu entry by its text."""
    def __init__(self, steps):
        self.steps, self.menu, self.log, self.prompts = list(steps), [], [], []

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
        self.prompts.append(prompt)
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

    def asker(self, mod=lg):
        return mod.Asker(interactive=True, input_fn=self.inp, out=self.out)

    def done(self):
        return not self.steps

    def questions(self):
        return sum(1 for s in self.log if s.startswith("\n? "))


def edit_copy(name, edit, src=GRID):
    wb = openpyxl.load_workbook(src, data_only=True)
    edit(wb)
    p = os.path.join(SCRATCH, name)
    wb.save(p)
    return p


def same(a, b):
    return a.reset_index(drop=True).equals(b.reset_index(drop=True))


def check(ok, label, extra=""):
    print("PASS" if ok else "FAIL", label, extra)


base, _ = lg.build(GRID, F, T, quiet)
print("baseline rows", len(base))

# ---------------------------------------------------------------- 1. sheet renamed -> pick the renamed sheet
def rename_tw(wb):
    wb["TW SATP"].title = "TW SATP v2"
    wb.copy_worksheet(wb["TW SATP v2"]).title = "TW SATP v3"      # a copy left in (11 Nov 2026): two sheets have the columns equally well, so the sheet is ASKED about; a lone renamed sheet is found by its columns
s = Script([Opt("TW SATP v2")])
df, _ = lg.build(edit_copy("i1.xlsx", rename_tw), F, T, quiet, s.asker())
check(same(df, base) and s.done() and s.questions() == 1, "sheet renamed: pick it -> identical output, 1 question")

# 11 Nov 2026 (structure shifts, `sheet_by_columns`): the lone renamed sheet is found by its columns - no question, identical output, the decision is recorded
def rename_tw_alone(wb):
    wb["TW SATP"].title = "TW SATP v2"
s = Script([])
a = s.asker()
df, _ = lg.build(edit_copy("i1a.xlsx", rename_tw_alone), F, T, quiet, a)
check(same(df, base) and s.questions() == 0 and any("its columns are those of" in d and "TW SATP" in d for d in a.decisions), "sheet renamed (the only sheet with its columns): found by its columns, no question, identical output, the decision is recorded")

# ---------------------------------------------------------------- 2. sheet renamed -> leave that part out
s = Script([Opt("leave out")])
a = s.asker()
df, _ = lg.build(edit_copy("i2.xlsx", rename_tw), F, T, quiet, a)
gone = len(base) - len(df)
check(gone == 888 and not ((df["product"] == "TW") & (df["sub_product"] == "STP")).any(), "sheet renamed: leave out -> TW STP rows gone", f"({gone} fewer)")

# ---------------------------------------------------------------- 3. missing column (PC SATP EV): zero / skip / use another
def hide_ev(wb):
    ws = wb["PC SATP"]
    for row in ws.iter_rows():
        for c in row:
            if c.value == "All KWH":
                c.value = "All KWH (gone)"
pc_ev = lambda d: d[(d["product"] == "Private Car") & (d["sub_product"] == "STP") & (d["fuel_type"] == "Electric")]
p = edit_copy("i3.xlsx", hide_ev)
s = Script([Opt("rate 0")])
df, _ = lg.build(p, F, T, quiet, s.asker())
e = pc_ev(df)
check(len(df) == len(base) and len(e) == 72 and (e["pi_tp_rate"] == 0).all() and s.done(), "missing EV column -> rate 0: same row count, Electric rates all 0", f"({len(e)} rows)")
s = Script([Opt("leave these rows out")])
df, _ = lg.build(p, F, T, quiet, s.asker())
check(len(df) == len(base) - 72 and len(pc_ev(df)) == 0, "missing EV column -> leave out: 72 fewer rows")
s = Script([Opt("All KWH (gone)")])
a = s.asker()
df, w = lg.build(p, F, T, quiet, a)
check(same(df, base) and w == [] and s.done() and any("All KWH" in d for d in a.decisions), "missing EV column -> 'it is the renamed column' -> identical output to the original, no warning")
s = Script([Opt("every column"), Opt("Petrol / PC > 1500 CC")])
a = s.asker()
df, w = lg.build(p, F, T, quiet, a)
e = pc_ev(df).set_index("new_rto_list")["pi_tp_rate"]
pet = df[(df["product"] == "Private Car") & (df["sub_product"] == "STP") & (df["fuel_type"] == "Petrol") & (df["cc_max"] == 99999)].set_index("new_rto_list")["pi_tp_rate"]
check(len(df) == len(base) and (e == pet.reindex(e.index)).all() and s.done() and s.questions() == 2, "missing EV column -> 'show every column', pick Petrol >1500: EV takes that column's rates", f"(questions: {s.questions()})")

# ---------------------------------------------------------------- 4. slab renamed -> choose the new slab
def rename_slab(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    for row in ws.iter_rows():
        for c in row:
            if c.value == lg.CV_SLAB:
                c.value = "< 2L"
s = Script([Opt("'< 2L'")])
df, _ = lg.build(edit_copy("i4.xlsx", rename_slab), F, T, quiet, s.asker())
check(same(df, base) and s.done() and s.questions() == 1, "slab renamed: pick the new one -> identical output, asked once for both tables")

# ---------------------------------------------------------------- 5. header row not found -> type the row number
def lose_header(wb):
    wb["PC TW_Comp and SAOD"]["A4"].value = "Geografie"
s = Script(["abc", "9999", "2", "4"])        # not a number, past the end, blank in column A, then the real row
a = s.asker()
df, _ = lg.build(edit_copy("i5.xlsx", lose_header), F, T, quiet, a)
check(same(df, base) and s.done() and any("header row" in d for d in a.decisions), "header row lost: bad answers re-asked, row number accepted -> identical output")

# ---------------------------------------------------------------- 6. unreadable cells
def bad_cell(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    r0 = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "BIHAR")
    r = next(r for r in range(r0, r0 + 8) if ws.cell(r, 2).value == lg.CV_SLAB)
    c = next(c for c in range(1, ws.max_column + 1) if ws.cell(4, c).value == "GCV 3W")      # first one = the 'New' age group
    ws.cell(r, c).value = "TBD"
s = Script([Opt("Use 0")])
a = s.asker()
df, _ = lg.build(edit_copy("i6.xlsx", bad_cell), F, T, quiet, a)
bad = df[(df["new_rto_list"] == "LIBERTY_GCV_4W_<7.5T_APR26_BIHAR") & (df["product"] == "GCV 3W") & (df["sub_product"] == "1+1") & (df["vehicle_age_max"] == 1.01)]
check(len(df) == len(base) and len(bad) > 0 and (bad["pi_net_rate"] == 0).all() and len(a.tallies.get("cells that were not rates", [])) >= 1,
      "unreadable cell -> use 0: row count same, that rate is 0, listed in the summary", f"({len(bad)} rows)")
# two bad cells + 'all' -> asked once
def two_bad(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    c = next(c for c in range(1, ws.max_column + 1) if ws.cell(4, c).value == "GCV 3W")
    n = 0
    for rr in range(5, ws.max_row):
        if ws.cell(rr, 2).value == lg.CV_SLAB and n < 2:
            ws.cell(rr, c).value = "TBD"
            n += 1
s = Script([Opt("Use 0"), "y"])           # the 2nd 'TBD': "same kind of problem - do that for every other cell like it?" -> yes
a = s.asker()
df, _ = lg.build(edit_copy("i6b.xlsx", two_bad), F, T, quiet, a)
check(s.done() and s.questions() == 1 and len(a.tallies["cells that were not rates"]) == 2, "two unreadable cells + 'same for every cell like it' -> one menu, one yes/no, both used as 0")

# ---------------------------------------------------------------- 7. q stops, no keyboard stops, declined confirm writes nothing
try:
    lg.build(edit_copy("i7.xlsx", rename_tw), F, T, quiet, Script(["q"]).asker())
    check(False, "q stops the run")
except lg.GridError as e:
    check("stopped by you" in str(e), "q at a question stops the run:", str(e))

def eof(prompt=""):
    raise EOFError
try:
    lg.build(edit_copy("i7b.xlsx", rename_tw), F, T, quiet, lg.Asker(interactive=True, input_fn=eof, out=quiet))
    check(False, "EOF stops")
except lg.GridError as e:
    check("no keyboard input" in str(e), "closed keyboard (EOF) stops with a clear message:", str(e))

out = os.path.join(SCRATCH, "i7_out.csv")
if os.path.exists(out):
    os.remove(out)
s = Script(["n"])
try:
    lg.main(GRID, F, T, out, None, quiet, s.asker())
    check(False, "declining the confirmation")
except lg.GridError as e:
    check("nothing was written" in str(e) and not os.path.exists(out), "answering n to 'write the file now?' writes nothing:", str(e))
s = Script(["", ])
lg.main(GRID, F, T, out, None, quiet, s.asker())
check(os.path.exists(out) and len(pd.read_csv(out)) == len(base), "Enter at 'write the file now?' writes the file")

# no question needed -> a clean workbook asks nothing, even in interactive mode
s = Script([""])
lg.main(GRID, F, T, os.path.join(SCRATCH, "i7c.csv"), None, quiet, s.asker())
check(s.questions() == 0 and s.done(), "clean workbook in interactive mode: no questions except the final confirmation")

# ---------------------------------------------------------------- 8. real July grids, answered interactively
class Rules(Script):
    """answers every menu with the first rule whose text appears in one of its options"""
    def __init__(self, needles):
        super().__init__([])
        self.needles, self.asked = needles, []

    def inp(self, prompt=""):
        self.asked.append(self.menu_title if hasattr(self, "menu_title") else "")
        for needle in self.needles:
            for n, label in self.menu:
                if needle.lower() in label.lower():
                    return str(n)
        raise AssertionError(f"no rule matches {self.menu}")

    def out(self, s):
        if str(s).startswith("\n? "):
            self.menu_title = str(s)[3:90]
        super().out(s)


for tag, path in (("JULY-B (4 slabs per geo)", JUL_B), ("JULY-A (one slab per geo)", JUL_A)):
    s = Rules(["no such column: write these rows with rate 0", "Leave these geos out"])
    a = s.asker()
    df, w = lg.build(path, dt.date(2026, 7, 1), dt.date(2030, 3, 20), quiet, a)
    e = pc_ev(df)
    titles = " | ".join(t[:60] for t in s.asked)
    # GOA and the 3 MAHARASHTRA geos write the slab as '<1.5L' (no space): they are matched, not asked about
    check(len(df) == 23010 and len(e) == 72 and (e["pi_tp_rate"] == 0).all() and w == [] and df.duplicated(lg.KEY_COLUMNS).sum() == 0 and s.questions() == 1,
          f"{tag}: only the Electric column is asked about ('<1.5L' vs '< 1.5L' is matched) -> {len(df)} rows, no duplicates, no warnings", f"[{s.questions()} question: {titles}]")

# ---------------------------------------------------------------- 9. non-interactive still stops with the question as the message
try:
    lg.build(edit_copy("i9.xlsx", hide_ev), F, T, quiet)
    check(False, "non-interactive stop")
except lg.NeedsAnswer as e:
    check("column ('EV', 'All KWH') not found" in str(e), "no Asker given: stops with the old message")

# ---------------------------------------------------------------- 10. Colab flow: upload -> menu answers -> dates -> confirm -> download
class FakeFiles:
    downloaded = None
    def upload(self):
        return {"Liberty grid renamed.xlsx": open(edit_copy("i10.xlsx", rename_tw), "rb").read()}
    def download(self, name):
        if str(name).lower().endswith(".csv"):        # the review workbook is downloaded too (27 Oct 2026): this test reads the upload
            FakeFiles.downloaded = name
sc = Script(["1/9/26", "01/09/2026", "31/03/2030", Opt("TW SATP v2"), ""])      # dates first, then the sheet question while reading, then Enter to confirm
sc.menu = []
real_input, real_print = builtins.input, builtins.print
builtins.input, builtins.print = sc.inp, lambda *a, **k: sc.out(" ".join(map(str, a)))
try:
    lg2 = load("lg2", SRC)
    lg2.colab_files = FakeFiles()
    cwd = os.getcwd()
    os.chdir(tempfile.mkdtemp(prefix="lib_colab_i_"))
    lg2.default_output = lambda filename, _d=tempfile.mkdtemp(prefix="lib_colab_out_"): os.path.join(_d, filename)   # the test must never write into output (it overwrote the files LIBERTY_UPLOAD_2026-09.csv / MAGMA_UPLOAD_2026-10.csv there)
    lg2.colab_main()
    c = pd.read_csv(FakeFiles.downloaded, dtype=str, keep_default_na=False)
    os.chdir(cwd)
finally:
    builtins.input, builtins.print = real_input, real_print
# the sheet question comes first (Asker created before the upload), dates after
check(len(c) == len(base) and set(c["from_date"]) == {"01/09/2026"} and sc.done(), "Colab flow with a renamed sheet: question answered, dates typed, file downloaded", f"({len(c)} rows)")

# ---------------------------------------------------------------- 12. several columns at once: Electric takes the Petrol columns, each with its own cc band
def electric_matches_petrol(df):
    pc = df[(df["product"] == "Private Car") & (df["sub_product"] == "STP")]
    pet = pc[pc["fuel_type"] == "Petrol"][["new_rto_list", "cc_min", "cc_max", "pi_tp_rate", "po_tp_rate"]]
    ele = pc[pc["fuel_type"] == "Electric"][["new_rto_list", "cc_min", "cc_max", "pi_tp_rate", "po_tp_rate"]]
    m = ele.merge(pet, on=["new_rto_list", "cc_min", "cc_max"], suffixes=("_e", "_p"), how="outer", indicator=True)
    return (len(ele) == len(pet) == 216 and (m["_merge"] == "both").all() and (m["pi_tp_rate_e"] == m["pi_tp_rate_p"]).all()
            and (m["po_tp_rate_e"] == m["po_tp_rate_p"]).all() and sorted(ele["cc_max"].unique()) == [1000.01, 1500.01, 99999.0])
not_ev = lambda d: d[~((d["product"] == "Private Car") & (d["sub_product"] == "STP") & (d["fuel_type"] == "Electric"))]
p = edit_copy("i12.xlsx", hide_ev)
s = Script([Opt("Electric same as Petrol")])
a = s.asker()
df, w = lg.build(p, F, T, quiet, a)
check(electric_matches_petrol(df) and len(df) == len(base) - 72 + 216 and same(not_ev(df), not_ev(base)) and s.done()
      and any("All KWH" in x and "NOT in the output" in x for x in w)          # the unused (renamed) EV column is still reported as unmapped
      and any("one set of rows per column" in d for d in a.decisions),
      "EV missing -> 'Electric same as Petrol' (one choice): Electric = Petrol rates in all 3 cc bands, nothing else changes", f"({len(df)} rows)")
for typed in ("1,2,3", "1-3", " 1 - 3 ", "3, 1-2"):
    s = Script([Opt("every column"), typed])
    df, w = lg.build(p, F, T, quiet, s.asker())
    check(electric_matches_petrol(df) and same(not_ev(df), not_ev(base)) and s.done(), f"EV missing -> show every column, type {typed!r} -> same as Petrol")
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "Petrol / PC 1000-1500 CC", "Petrol / PC > 1500 CC"])])
df, w = lg.build(p, F, T, quiet, s.asker())
check(electric_matches_petrol(df) and s.done(), "EV missing -> several columns picked by name in the list menu -> same as Petrol")

# a first-menu multi-pick of plain column suggestions works too (two of the suggested columns)
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "Diesel / PC < 1000 CC"])])
df, w = lg.build(p, F, T, quiet, s.asker())
e = pc_ev(df)
out = lg.report(df, [], quiet)
e2 = pc_ev(out)
vsum = lambda x: x[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).sum(axis=1)
want = sorted(vsum(e).groupby([e[c] for c in lg.KEY_COLUMNS]).min().values)
check(len(e) == 144 and set(e["cc_max"]) == {1000.01} and s.done() and len(e2) == 72 and sorted(vsum(e2).values) == want and not out.duplicated(lg.KEY_COLUMNS).any(),
      "two columns with the same cc band -> both read, then the duplicate rule keeps the LOWER payin of each entry (72 rows, nothing asked)", f"({len(e)} -> {len(e2)} Electric rows)")

# ---------------------------------------------------------------- 13. combinations that are not allowed are refused and re-asked
s = Script([Opts(["every column", "rate 0"]), Opt("rate 0")])
a = s.asker()
df, _ = lg.build(p, F, T, quiet, a)
check(any("can only be picked on their own" in l for l in s.log) and len(pc_ev(df)) == 72 and (pc_ev(df)["pi_tp_rate"] == 0).all() and s.done(),
      "'every column' + 'rate 0' together: refused with a message, then re-asked")
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "All KWH (gone)"]), Opts(["Petrol / PC < 1000 CC", "Petrol / PC 1000-1500 CC", "Petrol / PC > 1500 CC"])])
df, _ = lg.build(p, F, T, quiet, s.asker())
check(any("don't know the band" in l for l in s.log) and electric_matches_petrol(df) and s.done(), "a column whose band is unknown can't be combined: refused, then re-asked")
s = Script([Opt("every column"), "1,99", "0", "3-1", "abc", "1,2,3"])
df, _ = lg.build(p, F, T, quiet, s.asker())
check(electric_matches_petrol(df) and s.done() and sum("Please type a number from 1 to" in l for l in s.log) == 4, "bad multi answers ('1,99', '0', '3-1', 'abc') are each refused, then '1,2,3' works")

# ---------------------------------------------------------------- 14. questions that take one answer refuse commas
s = Script(["1,2", Opt("'< 2L'")])
df, _ = lg.build(edit_copy("i14.xlsx", rename_slab), F, T, quiet, s.asker())
check(same(df, base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "slab question (one answer): '1,2' refused, then one number works")
def hide_gcv3w_cv(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    c = next(c for c in range(1, ws.max_column + 1) if ws.cell(4, c).value == "GCV 3W")       # only the first one: one question
    ws.cell(4, c).value = "GCV 3W (renamed)"
s = Script(["1,2", Opt("rate 0")])
df, _ = lg.build(edit_copy("i14b.xlsx", hide_gcv3w_cv), F, T, quiet, s.asker())
check(len(df) == len(base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "CV column question (no bands to copy): commas refused, single answer works")

# ---------------------------------------------------------------- 12. several columns at once: Electric takes the Petrol columns, each with its own cc band
def electric_matches_petrol(df):
    pc = df[(df["product"] == "Private Car") & (df["sub_product"] == "STP")]
    pet = pc[pc["fuel_type"] == "Petrol"][["new_rto_list", "cc_min", "cc_max", "pi_tp_rate", "po_tp_rate"]]
    ele = pc[pc["fuel_type"] == "Electric"][["new_rto_list", "cc_min", "cc_max", "pi_tp_rate", "po_tp_rate"]]
    m = ele.merge(pet, on=["new_rto_list", "cc_min", "cc_max"], suffixes=("_e", "_p"), how="outer", indicator=True)
    return (len(ele) == len(pet) == 216 and (m["_merge"] == "both").all() and (m["pi_tp_rate_e"] == m["pi_tp_rate_p"]).all()
            and (m["po_tp_rate_e"] == m["po_tp_rate_p"]).all() and sorted(ele["cc_max"].unique()) == [1000.01, 1500.01, 99999.0])
not_ev = lambda d: d[~((d["product"] == "Private Car") & (d["sub_product"] == "STP") & (d["fuel_type"] == "Electric"))]
p = edit_copy("i12.xlsx", hide_ev)
s = Script([Opt("Electric same as Petrol")])
a = s.asker()
df, w = lg.build(p, F, T, quiet, a)
check(electric_matches_petrol(df) and len(df) == len(base) - 72 + 216 and same(not_ev(df), not_ev(base)) and s.done()
      and any("All KWH" in x and "NOT in the output" in x for x in w)          # the unused (renamed) EV column is still reported as unmapped
      and any("one set of rows per column" in d for d in a.decisions),
      "EV missing -> 'Electric same as Petrol' (one choice): Electric = Petrol rates in all 3 cc bands, nothing else changes", f"({len(df)} rows)")
for typed in ("1,2,3", "1-3", " 1 - 3 ", "3, 1-2"):
    s = Script([Opt("every column"), typed])
    df, w = lg.build(p, F, T, quiet, s.asker())
    check(electric_matches_petrol(df) and same(not_ev(df), not_ev(base)) and s.done(), f"EV missing -> show every column, type {typed!r} -> same as Petrol")
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "Petrol / PC 1000-1500 CC", "Petrol / PC > 1500 CC"])])
df, w = lg.build(p, F, T, quiet, s.asker())
check(electric_matches_petrol(df) and s.done(), "EV missing -> several columns picked by name in the list menu -> same as Petrol")

# a first-menu multi-pick of plain column suggestions works too (two of the suggested columns)
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "Diesel / PC < 1000 CC"])])
df, w = lg.build(p, F, T, quiet, s.asker())
e = pc_ev(df)
out = lg.report(df, [], quiet)
e2 = pc_ev(out)
vsum = lambda x: x[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).sum(axis=1)
want = sorted(vsum(e).groupby([e[c] for c in lg.KEY_COLUMNS]).min().values)
check(len(e) == 144 and set(e["cc_max"]) == {1000.01} and s.done() and len(e2) == 72 and sorted(vsum(e2).values) == want and not out.duplicated(lg.KEY_COLUMNS).any(),
      "two columns with the same cc band -> both read, then the duplicate rule keeps the LOWER payin of each entry (72 rows, nothing asked)", f"({len(e)} -> {len(e2)} Electric rows)")

# ---------------------------------------------------------------- 13. combinations that are not allowed are refused and re-asked
s = Script([Opts(["every column", "rate 0"]), Opt("rate 0")])
a = s.asker()
df, _ = lg.build(p, F, T, quiet, a)
check(any("can only be picked on their own" in l for l in s.log) and len(pc_ev(df)) == 72 and (pc_ev(df)["pi_tp_rate"] == 0).all() and s.done(),
      "'every column' + 'rate 0' together: refused with a message, then re-asked")
s = Script([Opt("every column"), Opts(["Petrol / PC < 1000 CC", "All KWH (gone)"]), Opts(["Petrol / PC < 1000 CC", "Petrol / PC 1000-1500 CC", "Petrol / PC > 1500 CC"])])
df, _ = lg.build(p, F, T, quiet, s.asker())
check(any("don't know the band" in l for l in s.log) and electric_matches_petrol(df) and s.done(), "a column whose band is unknown can't be combined: refused, then re-asked")
s = Script([Opt("every column"), "1,99", "0", "3-1", "abc", "1,2,3"])
df, _ = lg.build(p, F, T, quiet, s.asker())
check(electric_matches_petrol(df) and s.done() and sum("Please type a number from 1 to" in l for l in s.log) == 4, "bad multi answers ('1,99', '0', '3-1', 'abc') are each refused, then '1,2,3' works")

# ---------------------------------------------------------------- 14. questions that take one answer refuse commas
s = Script(["1,2", Opt("'< 2L'")])
df, _ = lg.build(edit_copy("i14.xlsx", rename_slab), F, T, quiet, s.asker())
check(same(df, base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "slab question (one answer): '1,2' refused, then one number works")
def hide_gcv3w_cv(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    c = next(c for c in range(1, ws.max_column + 1) if ws.cell(4, c).value == "GCV 3W")       # only the first one: one question
    ws.cell(4, c).value = "GCV 3W (renamed)"
s = Script(["1,2", Opt("rate 0")])
df, _ = lg.build(edit_copy("i14b.xlsx", hide_gcv3w_cv), F, T, quiet, s.asker())
check(len(df) == len(base) and s.done() and any("Please type a number from 1 to" in l for l in s.log), "CV column question (no bands to copy): commas refused, single answer works")

# ---------------------------------------------------------------- 15. spaces and capitals in slab names and header labels do not matter
def respace(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    n = 0
    for row in ws.iter_rows():
        for c in row:
            if c.value == lg.CV_SLAB:
                c.value = ["<1.5L", " < 1.5L ", "<1.5 l", "<  1.5L"][n % 4]
                n += 1
    ws2 = wb["PC TW_Comp and SAOD"]
    for row in ws2.iter_rows():
        for c in row:
            if c.value == lg.PCTW_SLAB:
                c.value = "50K-1.5L" if (c.row % 2) else "50k - 1.5L"
    ws3 = wb["CV >7.5T Comp & SATP"]
    for r in range(1, 8):
        for c in range(1, ws3.max_column + 1):
            v = ws3.cell(r, c).value
            if v in ("7.5 - 12T", "12 - 20T"):
                ws3.cell(r, c).value = v.replace(" ", "")
s = Script([])
df, w = lg.build(edit_copy("i15.xlsx", respace), F, T, quiet, s.asker())
check(same(df, base) and s.questions() == 0 and w == [], "slab names '<1.5L' / ' < 1.5L ' / '50K-1.5L' and headers '7.5-12T' (other spaces, capitals) -> matched, no question, identical output")

# ---------------------------------------------------------------- 15. spaces and capitals in slab names and header labels do not matter
def respace(wb):
    ws = wb["CV <7.5T Comp & SATP"]
    n = 0
    for row in ws.iter_rows():
        for c in row:
            if c.value == lg.CV_SLAB:
                c.value = ["<1.5L", " < 1.5L ", "<1.5 l", "<  1.5L"][n % 4]
                n += 1
    ws2 = wb["PC TW_Comp and SAOD"]
    for row in ws2.iter_rows():
        for c in row:
            if c.value == lg.PCTW_SLAB:
                c.value = "50K-1.5L" if (c.row % 2) else "50k - 1.5L"
    ws3 = wb["CV >7.5T Comp & SATP"]
    for r in range(1, 8):
        for c in range(1, ws3.max_column + 1):
            v = ws3.cell(r, c).value
            if v in ("7.5 - 12T", "12 - 20T"):
                ws3.cell(r, c).value = v.replace(" ", "")
s = Script([])
df, w = lg.build(edit_copy("i15.xlsx", respace), F, T, quiet, s.asker())
check(same(df, base) and s.questions() == 0 and w == [], "slab names '<1.5L' / ' < 1.5L ' / '50K-1.5L' and headers '7.5-12T' (other spaces, capitals) -> matched, no question, identical output")

# ---------------------------------------------------------------- 11. the two scripts share one question layer
def shared(path):
    t = open(path, encoding="utf-8").read()
    a = re.search(r"^# %% -+ QUESTIONS\s*$", t, re.M).start()
    b = re.search(r"^# %% -+ READING THE ", t, re.M).start()
    return t[a:b]
check(shared(SRC) == shared(SRC_M), "question layer is identical in the Liberty and Magma scripts")
