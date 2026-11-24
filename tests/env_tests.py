"""Environment / input-handling resilience tests for both scripts (files, dates, typing, uploads, Ctrl+C, console)."""
import os as _os
ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))      # the grid_tools folder
import builtins
import datetime as dt
import importlib.util
import io
import os
import random
import re
import shutil
import sys
import tempfile

SC = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SC)
import pandas as pd  # noqa: E402
from fuzzlib import RandomUser  # noqa: E402

GA = ROOT
SEP = os.path.join(ROOT, "regression", "inputs", "liberty_sep26.xlsx")
REF = os.path.join(ROOT, "regression", "references", "liberty_sep26_rates_export.csv")
MGRID = os.path.join(ROOT, "regression", "inputs", "magma_oct26_comprehensive.xlsx")
MSATP = os.path.join(SC, "i_satp_base.xlsx")           # the SATP grid saved as .xlsx (made by the earlier tests)
if not os.path.exists(MSATP):                          # made here from the real SATP .xlsb when the file is not there yet
    import openpyxl
    from pyxlsb import open_workbook
    _wb = openpyxl.Workbook()
    _ws = _wb.active
    _ws.title = "<=2 Lacs"
    with open_workbook(os.path.join(ROOT, "regression", "inputs", "magma_oct26_satp.xlsb")) as _sb, _sb.get_sheet("<=2 Lacs") as _sh:
        for _row in _sh.rows():
            _ws.append([c.v for c in _row])
    _wb.save(MSATP)
MREF = os.path.join(ROOT, "regression", "references", "magma_oct26_reupload.csv")
quiet = lambda *a, **k: None
TMP = tempfile.mkdtemp(prefix="env_")
F, T = dt.date(2026, 9, 1), dt.date(2030, 3, 31)
FAILS = []


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


lg, mg = load("lg", GA + r"\liberty_grid_to_config.py"), load("mg", GA + r"\magma_grid_to_config.py")


def check(ok, label, extra=""):
    print("PASS" if ok else "FAIL", label, extra)
    if not ok:
        FAILS.append(label)


class Keys:
    """scripted keyboard; a step is a string, or ('opt', needle) to pick a menu entry by its text"""
    def __init__(self, steps):
        self.steps, self.menu, self.log = list(steps), {}, []

    def out(self, s):
        s = str(s)
        self.log.append(s)
        if s.startswith("\n? "):
            self.menu = {}
        m = re.match(r"\s+(\d+)\) (.*)", s)
        if m:
            self.menu[int(m.group(1))] = m.group(2)

    def inp(self, prompt=""):
        if not self.steps:
            raise AssertionError(f"out of answers at {prompt!r}; last output: {self.log[-6:]}")
        step = self.steps.pop(0)
        if isinstance(step, tuple):
            hit = [n for n, l in self.menu.items() if step[1].lower() in l.lower()]
            assert len(hit) >= 1, f"no option {step[1]!r} in {self.menu}"
            return str(hit[0])
        return step

    def ask(self, mod):
        return mod.Asker(True, self.inp, self.out)


# ============================================================ 1. dates
for mod, name in ((lg, "Liberty"), (mg, "Magma")):
    good = {"01/10/2026": (2026, 10, 1), "1/10/2026": (2026, 10, 1), "01-10-2026": (2026, 10, 1), "01.10.2026": (2026, 10, 1),
            "2026-10-01": (2026, 10, 1), "1 Oct 2026": (2026, 10, 1), "01-Oct-2026": (2026, 10, 1), "1 october 2026": (2026, 10, 1),
            " 31/03/2030 ": (2030, 3, 31), "29/02/2028": (2028, 2, 29), "\ufeff01/10/2026": (2026, 10, 1)}
    ok = all(mod.parse_ddmmyyyy(k) == dt.date(*v) for k, v in good.items())
    bad = ["", " ", "1/10/26", "32/01/2026", "00/10/2026", "29/02/2027", "10/13/2026", "abc", "2026/13/01", "1 Foo 2026", "01/10/26 ", "12345678", "1//2026", None, 20261001]
    refused = 0
    for b in bad:
        try:
            mod.parse_ddmmyyyy(b)
        except ValueError:
            refused += 1
        except Exception as e:                       # anything but ValueError is a bug
            print("   unexpected", type(e).__name__, repr(b))
    check(ok and refused == len(bad), f"{name}: date formats accepted / refused cleanly", f"({len(good)} good, {refused}/{len(bad)} refused)")
    rng, bugs = random.Random(5), 0
    alphabet = "0123456789/-. :abcJanOctZ٠\n\t"
    for _ in range(20000):
        s = "".join(rng.choice(alphabet) for _ in range(rng.randint(0, 14)))
        try:
            d = mod.parse_ddmmyyyy(s)
            assert isinstance(d, dt.date)
        except ValueError:
            pass
        except Exception:
            bugs += 1
    check(bugs == 0, f"{name}: 20,000 random strings -> a date or a ValueError, nothing else")
    t = dt.date(2026, 10, 22)
    check(mod.date_notes(dt.date(2026, 10, 1), dt.date(2030, 3, 31), t) == []
          and len(mod.date_notes(dt.date(2026, 10, 1), dt.date(2026, 10, 5), t)) == 2          # already past + only 5 days
          and any("past" in n for n in mod.date_notes(dt.date(2020, 1, 1), dt.date(2021, 1, 1), t))
          and any("only" in n for n in mod.date_notes(dt.date(2026, 10, 1), dt.date(2026, 10, 10), t))
          and any("years long" in n for n in mod.date_notes(dt.date(2026, 10, 1), dt.date(2045, 1, 1), t)), f"{name}: unusual date ranges are noticed")
    # typed dates: bad ones re-asked, reversed re-asked, unusual ones confirmed (and a 'no' asks again)
    answers = iter(["x", "1/10/26", "2026-10-01", "31/03/2030"])
    f, t2 = mod.ask_date_range(lambda p: next(answers))
    check((f, t2) == (dt.date(2026, 10, 1), dt.date(2030, 3, 31)), f"{name}: typed dates - garbage re-asked, ISO accepted")
    keys = Keys(["n"])                                        # unusual range -> 'no' -> asked again (the 2nd range is normal)
    answers = iter(["01/10/2019", "05/10/2019", "01/10/2026", "31/03/2030"])
    f, t2 = mod.ask_date_range(lambda p: next(answers), keys.ask(mod))
    check((f, t2) == (dt.date(2026, 10, 1), dt.date(2030, 3, 31)) and keys.steps == [], f"{name}: unusual dates are confirmed; 'n' asks for them again")
    def ctrl_c(p):
        raise KeyboardInterrupt
    try:
        mod.ask_date("from_date", ctrl_c)
        check(False, f"{name}: Ctrl+C at a date prompt")
    except mod.GridError as e:
        check("stopped by you" in str(e), f"{name}: Ctrl+C at a date prompt stops cleanly")

# ============================================================ 2. typing garbage at questions
for mod, name in ((lg, "Liberty"), (mg, "Magma")):
    rng, bugs = random.Random(9), 0
    for _ in range(20000):
        n = rng.randint(1, 15)
        s = "".join(rng.choice("0123456789,-; \t?qa") for _ in range(rng.randint(0, 12)))
        try:
            p = mod.parse_picks(s, n)
            assert p and all(1 <= i <= n for i in p) and len(set(p)) == len(p)
        except ValueError:
            pass
        except Exception:
            bugs += 1
    check(bugs == 0, f"{name}: 20,000 random answers to parse_picks -> valid picks or ValueError")
    opts = [("a", 1), ("b", 2), ("c", 3)]
    k = Keys(["", "zzz", "0", "4", "1.5", "-1", "1,2", "?", "3"])
    a = k.ask(mod)
    check(a.choose("p", opts) == 3 and k.steps == [], f"{name}: menu refuses blanks/garbage/ranges, shows details on '?', takes a valid number")
    k = Keys(["9", "2-1", "1,,x", "3,1"])
    a = k.ask(mod)
    check(a.choose("p", opts, multi=True) == [3, 1] and k.steps == [], f"{name}: multi menu refuses 9 / 2-1 / 1,,x and takes '3,1' in the order typed")
    for exc, want in ((EOFError, "no keyboard"), (KeyboardInterrupt, "stopped by you")):
        def boom(p, exc=exc):
            raise exc
        try:
            mod.Asker(True, boom, quiet).choose("p", opts)
            check(False, f"{name}: {exc.__name__}")
        except mod.GridError as e:
            check(want in str(e), f"{name}: {exc.__name__} at a question -> clean stop ({e})")
    try:
        mod.Asker(True, lambda p: "1", quiet).choose("nothing to choose from", [])
        check(False, f"{name}: empty menu")
    except mod.NeedsAnswer as e:
        check("nothing to choose from" in str(e), f"{name}: a menu with no options stops with the message instead of looping")

# ============================================================ 3. files that will not open / cannot be written
def junk(name, content):
    p = os.path.join(TMP, name)
    open(p, "wb").write(content)
    return p

bad_files = {"empty.xlsx": b"", "text.xlsx": b"this is not a workbook", "fake.xls": b"\xd0\xcf\x11\xe0 old excel", "zip.xlsx": b"PK\x03\x04garbage"}
for fname, content in bad_files.items():
    p = junk(fname, content)
    k = Keys([("opt", "different file"), SEP])
    lg_wb = lambda x: x
    try:
        out = lg.main(p, F, T, os.path.join(TMP, "L_" + fname + ".csv"), None, quiet, lg.Asker(False))
        check(False, f"Liberty: {fname} without a keyboard")
    except lg.NeedsAnswer as e:
        check("cannot open" in str(e), f"Liberty: {fname} (no keyboard) -> 'cannot open ...' not a traceback")
    except Exception as e:
        check(False, f"Liberty: {fname} (no keyboard) crashed with {type(e).__name__}: {e}")
    try:
        mg.main(p, MSATP, F, T, os.path.join(TMP, "M_" + fname + ".csv"), None, quiet, mg.Asker(False))
        check(False, f"Magma: {fname} without a keyboard")
    except mg.NeedsAnswer as e:
        check("cannot open" in str(e), f"Magma: {fname} (no keyboard) -> 'cannot open ...'")
    except Exception as e:
        check(False, f"Magma: {fname} (no keyboard) crashed with {type(e).__name__}: {e}")
# a missing path and a directory
for target in (os.path.join(TMP, "nope.xlsx"), TMP):
    try:
        lg.main(target, F, T, os.path.join(TMP, "x.csv"), None, quiet, lg.Asker(False))
        check(False, "Liberty: missing path")
    except lg.NeedsAnswer as e:
        check("cannot open" in str(e), f"Liberty: {'directory' if os.path.isdir(target) else 'missing file'} -> clean question")
    except Exception as e:
        check(False, f"Liberty: {target} crashed with {type(e).__name__}: {e}")
# interactive: damaged file -> 'a different file' -> real workbook -> finished
k = Keys([("opt", "different file"), SEP, ""])
p = junk("zip2.xlsx", b"PK\x03\x04garbage")
out = lg.main(p, F, T, os.path.join(TMP, "L_ok.csv"), None, quiet, k.ask(lg))
check(os.path.exists(out) and len(pd.read_csv(out)) == 23690 and k.steps == [], "Liberty: damaged file -> pick another file -> 23,690 rows written")
k = Keys([("opt", "different file"), MGRID, ""])
out = mg.main(p, MSATP, dt.date(2026, 10, 1), T, os.path.join(TMP, "M_ok.csv"), None, quiet, k.ask(mg))
check(os.path.exists(out) and len(pd.read_csv(out, low_memory=False)) == 39102 and k.steps == [], "Magma: damaged file -> pick another file -> 39,102 rows written")
# a bad path typed for the replacement is refused and re-asked
k = Keys([("opt", "different file"), "C:\\nowhere\\x.xlsx", TMP, SEP, ""])
out = lg.main(p, F, T, os.path.join(TMP, "L_ok2.csv"), None, quiet, k.ask(lg))
check(os.path.exists(out) and k.steps == [], "Liberty: a wrong path / a folder typed as the replacement is refused and asked again")

# output that cannot be written
blocker = junk("blocker.txt", b"x")
for mod, name, call in ((lg, "Liberty", lambda o, a: lg.main(SEP, F, T, o, None, quiet, a)),
                        (mg, "Magma", lambda o, a: mg.main(MGRID, MSATP, dt.date(2026, 10, 1), T, o, None, quiet, a))):
    bad_out = os.path.join(blocker, "sub", "o.csv")             # a folder cannot be made under a file
    try:
        call(bad_out, mod.Asker(False))
        check(False, f"{name}: unwritable output without a keyboard")
    except mod.NeedsAnswer as e:
        check("cannot write" in str(e), f"{name}: unwritable output (no keyboard) -> 'cannot write ...'")
    except Exception as e:
        check(False, f"{name}: unwritable output crashed {type(e).__name__}: {e}")
    k = Keys(["", ("opt", "different file name"), os.path.join(TMP, name + "_renamed")])      # confirm, then the write fails
    out = call(bad_out, k.ask(mod))
    check(out.endswith(".csv") and os.path.exists(out) and k.steps == [], f"{name}: unwritable output -> new name -> written ({os.path.basename(out)})")
    # output into a folder that does not exist yet is simply created
    deep = os.path.join(TMP, "a", "b", "c", name + ".csv")
    check(os.path.exists(call(deep, mod.Asker(False))), f"{name}: output folder that does not exist is created")

# comparison files that are not usable
refs = {"empty.csv": b"", "binary.csv": bytes(range(256)) * 40, "wrongcols.csv": b"a,b,c\n1,2,3\n", "latin1.csv": "x,y\n\xe9\xe8,2\n".encode("latin-1"),
        "onlyheader.csv": b"new_rto_list,product\n"}
for rname, content in refs.items():
    rp = junk(rname, content)
    for mod, name, call in ((lg, "Liberty", lambda rp, a: lg.main(SEP, F, T, os.path.join(TMP, "r.csv"), rp, quiet, a)),
                            (mg, "Magma", lambda rp, a: mg.main(MGRID, MSATP, dt.date(2026, 10, 1), T, os.path.join(TMP, "r.csv"), rp, quiet, a))):
        try:
            out = call(rp, mod.Asker(False))
            check(os.path.exists(out), f"{name}: unusable comparison file {rname} -> skipped, upload still written")
        except Exception as e:
            check(False, f"{name}: comparison file {rname} crashed {type(e).__name__}: {e}")
k = Keys([("opt", "without the comparison"), ""])
out = lg.main(SEP, F, T, os.path.join(TMP, "r2.csv"), os.path.join(TMP, "wrongcols.csv"), quiet, k.ask(lg))
check(os.path.exists(out) and k.steps == [], "Liberty: broken comparison file (interactive) -> 'carry on without it' -> written")

# ============================================================ 4. Colab uploads
class Up:
    """fake google.colab.files: a list of uploads, one per call"""
    def __init__(self, batches, fail_download=False):
        self.batches, self.calls, self.downloaded, self.fail_download = list(batches), 0, None, fail_download

    def upload(self):
        self.calls += 1
        item = self.batches.pop(0) if self.batches else {}
        if isinstance(item, Exception):
            raise item
        return item

    def download(self, name):
        if self.fail_download:
            raise Exception("Failed to fetch")
        self.downloaded = name


def run_colab(mod, batches, steps, fail_download=False):
    k = Keys(steps)
    real_in, real_pr, real_cwd = builtins.input, builtins.print, os.getcwd()
    builtins.input, builtins.print = k.inp, lambda *a, **kw: k.out(" ".join(map(str, a)))
    try:
        m2 = load("c_" + mod.__name__, GA + (r"\liberty_grid_to_config.py" if mod is lg else r"\magma_grid_to_config.py"))
        m2.colab_files = Up(batches, fail_download)
        os.chdir(tempfile.mkdtemp(prefix="colab_", dir=TMP))
        m2.colab_main()
        return m2.colab_files, k
    finally:
        builtins.input, builtins.print = real_in, real_pr
        os.chdir(real_cwd)


sep_b, ref_b = open(SEP, "rb").read(), open(REF, "rb").read()
up, k = run_colab(lg, [{}, {"notes.txt": b"hi"}, {"Sep26.xlsx": sep_b, "ref.csv": ref_b}], ["y", "y", "01/09/2026", "31/03/2030", ""])
check(up.downloaded and up.calls == 3 and k.steps == [], "Liberty Colab: empty upload, then a .txt, then the real files -> asked to upload again each time, then finished")
up, k = run_colab(lg, [{"a.xlsx": sep_b, "b.xlsx": sep_b}], [("opt", "a.xlsx"), "01/09/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Liberty Colab: two workbooks uploaded -> asks which one")
up, k = run_colab(lg, [{}], ["n"])
check(up.downloaded is None and any("STOPPED" in l for l in k.log), "Liberty Colab: declining to upload again stops with a plain message")
up, k = run_colab(lg, [{"Sep.xlsx": sep_b}], ["01/09/2026", "31/03/2030", "n"])
check(up.downloaded is None and any("nothing was written" in l for l in k.log), "Liberty Colab: 'n' at the final question -> nothing downloaded")
mg_b, ms_b = open(MGRID, "rb").read(), open(MSATP, "rb").read()
up, k = run_colab(mg, [{"one.xlsx": mg_b}, {"two.xlsx": ms_b, "last.csv": open(MREF, "rb").read()}], ["y", "01/10/2026", "31/03/2030", ""])
check(up.downloaded and up.calls == 2 and k.steps == [], "Magma Colab: only one grid uploaded -> asks for more, then sorts them out and finishes")
up, k = run_colab(mg, [{"one.xlsx": mg_b, "two.xlsx": ms_b, "three.xlsx": mg_b}], [("opt", "one.xlsx"), ("opt", "two.xlsx"), "01/10/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Magma Colab: three workbooks -> asks which is comprehensive and which is SATP")

# --- the Colab web page itself failing ("Failed to fetch" from the upload window / the download call)
fetch_err = Exception("Failed to fetch")
up, k = run_colab(lg, [fetch_err], [("opt", "Files pane"), SEP, "", "01/09/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Liberty Colab: upload window fails ('Failed to fetch') -> use a file from the Files pane -> finished")
up, k = run_colab(lg, [fetch_err, {"Sep.xlsx": sep_b}], [("opt", "again"), "01/09/2026", "31/03/2030", ""])
check(up.downloaded and up.calls == 2 and k.steps == [], "Liberty Colab: upload window fails -> 'try again' -> works the second time")
up, k = run_colab(lg, [fetch_err, fetch_err], [("opt", "again"), ("opt", "Files pane"), "C:\\no\\such.xlsx", SEP, "", "01/09/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Liberty Colab: fails twice, a wrong path is refused, then a real path works")
up, k = run_colab(lg, [{"Sep.xlsx": sep_b}], ["01/09/2026", "31/03/2030", ""], fail_download=True)
check(up.downloaded is None and any("The file IS written" in l for l in k.log) and not any("STOPPED" in l for l in k.log),
      "Liberty Colab: the download call fails -> no traceback, says the file IS written and where")
up, k = run_colab(mg, [fetch_err], [("opt", "Files pane"), MGRID, MSATP, "", "01/10/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Magma Colab: upload window fails -> both grids given from the Files pane -> finished")
up, k = run_colab(mg, [{"one.xlsx": mg_b, "two.xlsx": ms_b}], ["01/10/2026", "31/03/2030", ""], fail_download=True)
check(up.downloaded is None and any("The file IS written" in l for l in k.log), "Magma Colab: the download call fails -> says the file IS written and where")

# --- the Colab web page itself failing ("Failed to fetch" from the upload window / the download call)
fetch_err = Exception("Failed to fetch")
up, k = run_colab(lg, [fetch_err], [("opt", "Files pane"), SEP, "", "01/09/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Liberty Colab: upload window fails ('Failed to fetch') -> use a file from the Files pane -> finished")
up, k = run_colab(lg, [fetch_err, {"Sep.xlsx": sep_b}], [("opt", "again"), "01/09/2026", "31/03/2030", ""])
check(up.downloaded and up.calls == 2 and k.steps == [], "Liberty Colab: upload window fails -> 'try again' -> works the second time")
up, k = run_colab(lg, [fetch_err, fetch_err], [("opt", "again"), ("opt", "Files pane"), "C:\\no\\such.xlsx", SEP, "", "01/09/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Liberty Colab: fails twice, a wrong path is refused, then a real path works")
up, k = run_colab(lg, [{"Sep.xlsx": sep_b}], ["01/09/2026", "31/03/2030", ""], fail_download=True)
check(up.downloaded is None and any("The file IS written" in l for l in k.log) and not any("STOPPED" in l for l in k.log),
      "Liberty Colab: the download call fails -> no traceback, says the file IS written and where")
up, k = run_colab(mg, [fetch_err], [("opt", "Files pane"), MGRID, MSATP, "", "01/10/2026", "31/03/2030", ""])
check(up.downloaded and k.steps == [], "Magma Colab: upload window fails -> both grids given from the Files pane -> finished")
up, k = run_colab(mg, [{"one.xlsx": mg_b, "two.xlsx": ms_b}], ["01/10/2026", "31/03/2030", ""], fail_download=True)
check(up.downloaded is None and any("The file IS written" in l for l in k.log), "Magma Colab: the download call fails -> says the file IS written and where")

# ============================================================ 5. unexpected errors, console, Ctrl+C
for mod, name, script in ((lg, "Liberty", "liberty_grid_to_config.py"), (mg, "Magma", "magma_grid_to_config.py")):
    m = load("cr_" + name, GA + "\\" + script)
    def explode(*a, **k):
        raise RuntimeError("boom inside the script")
    m.build = explode
    argv = [script, "--no-prompt", "--from-date", "01/09/2026", "--to-date", "31/03/2030"] + (
        ["--grid", SEP] if name == "Liberty" else ["--grid", MGRID, "--satp", MSATP])
    old_argv, old_cwd = sys.argv, os.getcwd()
    sys.argv = argv
    os.chdir(TMP)
    try:
        m.cli()
        check(False, f"{name}: unexpected exception")
    except SystemExit as e:
        text = str(e.code)
        files = [f for f in os.listdir(TMP) if "_error_" in f]
        check("unexpected" in text and "boom" in text and "Traceback" not in text and files, f"{name}: unexpected exception -> one-line message + details file", f"({files[-1:] if files else None})")
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)
    # a console that cannot print the character
    class Ascii(io.TextIOWrapper):
        pass
    buf = io.BytesIO()
    old = sys.stdout
    sys.stdout = io.TextIOWrapper(buf, encoding="ascii", errors="strict")
    verdict = None
    try:
        mod.say("geo: Ünïcode ✓ Ω")
        sys.stdout.flush()
        verdict = (b"geo:" in buf.getvalue(), "")
    except Exception as e:
        verdict = (False, type(e).__name__)
    finally:
        sys.stdout = old
    check(verdict[0], f"{name}: a console that cannot show accents / check marks does not crash the script", verdict[1])

print("\nFAILURES:", len(FAILS))
for f in FAILS:
    print("  -", f)
shutil.rmtree(TMP, ignore_errors=True)
