"""Structure-shift suite (11 Nov 2026): every converter either writes the SAME upload for a grid whose structure changed in a way that does not change its meaning, or STOPS with a message that says what
it expected and what it found - never a different file in silence, never a bare traceback.  (The user, 11 Nov 2026: "update our scripts for all insurers to prepare for potential structural changes,
like the one we encountered with Go Digit".)

The operators are in tests/shiftlib.py: sheets renamed / reordered / added / hidden copies, blank rows and columns above, left of and inside the table, title rows, whitespace noise in the headers, an extra
column at the right edge, two columns swapped with their headers, a blank row inside the data, a rate column there twice.  Each is applied to the real grids of regression\\inputs (every sheet at once)
and the script runs with --no-prompt (a new question is a stop, never a wrong answer).  The outcome is one of

    IDENTICAL   the same bytes                                   REORDERED  the same rows in another order (an upload's row order means nothing)
    REMARKS     only the remarks column differs (a banner above the table is a text the script flags, as it should)
    STOPPED     a clear 'STOPPED - ...' message (no traceback): the grid says something the script cannot place without guessing
    CRASH       a traceback                                      DIFFERENT  another file in silence  <- the one outcome that is never allowed

By default IDENTICAL and REORDERED pass.  EXPECT below lists, per operator (and per script where it differs), the other outcomes that are the RIGHT answer and why.

Run:  .venv\\Scripts\\python.exe tests\\structure_shift_tests.py                         all subjects x all operators (about 40-60 minutes with 4 workers)
      SHIFT_QUICK=1 .venv\\Scripts\\python.exe tests\\structure_shift_tests.py             the small grids only (about 15 minutes)
      SHIFT_ONLY=liberty,magma SHIFT_OPS=cols_left,reorder  ...                           a selection      SHIFT_WORKERS=2   fewer processes"""
import concurrent.futures as cf
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shiftlib as L  # noqa: E402

ROOT = L.ROOT
CASES = {c["name"]: c for c in json.load(open(os.path.join(ROOT, "regression", "cases.json"), encoding="utf-8"))["cases"]}

# subject -> (regression case, the argument that carries the file to shift)
SUBJECTS = {
    "liberty": ("liberty_sep26", "--grid"),
    "rsa_gcv4w": ("rsa_gcv4w_oct26", "--grid"),
    "rsa_pvtcar": ("rsa_pvtcar_oct26", "--grid"),
    "magma": ("magma_oct26", "--grid"),
    "shriram_broker": ("shriram_nov26", "--grid"),
    "shriram_2w": ("shriram_2w_oct26", "--grid"),
    "shriram_2w_broker": ("shriram_2w_oct26", "--broker-grid"),
    "sbi": ("sbi_oct26", "--grid"),
    "icici_cv": ("icici_cv_nov26", "--cv-grid"),
    "icici_pc": ("icici_pc_tw_nov26", "--pc-grid"),
    "icici_tw": ("icici_pc_tw_nov26", "--tw-grid"),
    "icici_tw_new": ("icici_pc_tw_nov26", "--tw-new-grid"),
    "godigit": ("godigit_oct26", "--grid"),
    "indusind": ("indusind_nov26", "--grid"),
    "indusind_oct": ("indusind_oct26", "--grid"),
}
QUICK = ["rsa_pvtcar", "shriram_2w", "shriram_2w_broker", "icici_pc", "icici_tw", "icici_tw_new", "magma"]
GOOD = {"IDENTICAL", "REORDERED"}

# (subject or "*", operator) -> (outcomes that are right there IN ADDITION to GOOD, why).  The first matching key wins; the subject's own key before the "*" one.
EXPECT = {
    # a banner row above the table is a text the script cannot tell from a rule: it is flagged for manual review, the rates are the same
    ("*", "title_rows"): ({"REMARKS"}, "a banner above the table is flagged as a note (the rates are the same)"),
    # every sheet called Sheet1 ... SheetN: where the NAME of a sheet is its meaning (which RTO group, which product, which OEM, which premium slab) there is nothing to read it from
    ("shriram_broker", "rename_harsh"): ({"STOPPED"}, "two sheets look like the rate table: which one is not in the document"),
    ("shriram_2w", "rename_harsh"): ({"STOPPED"}, "two sheets look like the Rollover sheet"),
    ("magma", "rename_harsh"): ({"STOPPED"}, "the premium slab of a sheet is its name ('< =3 lac', '>3 to <=5' ...)"),
    ("icici_cv", "rename_harsh"): ({"STOPPED"}, "the notes sheet is told from the rate matrices by its name"),
    ("icici_tw_new", "rename_harsh"): ({"STOPPED"}, "the OEM of a new Two-Wheeler sheet is its name (HMC, TVS ...)"),
    # last month's sheet left in next to this month's (same clusters): which one is current is not in the document
    ("rsa_gcv4w", "decoy_sheet"): ({"STOPPED"}, "two sheets hold the same clusters: expected one sheet per region"),
    ("rsa_pvtcar", "decoy_sheet"): ({"STOPPED"}, "two sheets hold the same clusters: expected one sheet per region"),
    ("indusind", "decoy_sheet"): ({"STOPPED"}, "two visible sheets hold the same commercial-vehicle table: which one is this month's is not in the document"),
    # the October workbook has a TABLE of text (the M & HCV grid, in the guidelines sheet): a new column or a blank row inside it changes the text of the note that is quoted in remarks - the rows and the rates are the same
    ("indusind_oct", "extra_col_right"): ({"REMARKS"}, "the new column lands inside the text of the M & HCV table: the note quoted in remarks changes, the rates are the same"),
    ("indusind_oct", "blank_row_top"): ({"REMARKS"}, "a blank row inside the M & HCV table splits its note in two: the notes quoted in remarks change, the rates are the same"),
    ("indusind_oct", "decoy_sheet"): ({"STOPPED"}, "two visible sheets hold the same commercial-vehicle table: which one is this month's is not in the document"),
    # a column that is there twice under one heading (the copy has HIGHER numbers): which one is meant is not in the document.  Where it is a PAYIN the script writes the LOWEST of the two (the user,
    # 12 Nov 2026: ICICI, Shriram, RSA) - the original's output, IDENTICAL on either side; where it is no payin (DIS %) or the script cannot tell (Magma, SBI, Go Digit) it STOPS
    ("*", "dup_col"): ({"STOPPED"}, "two columns with one heading: the lowest payin is written (IDENTICAL here), a column that is no payin stops - never one picked in silence"),
    ("*", "dup_col_front"): ({"STOPPED"}, "the same, the copy on the left (a reader that takes the first column is wrong here)"),
}


def target(case, flag):
    a = case["args"]
    return a[a.index(flag) + 1]


def allowed_for(subject, op):
    for key in ((subject, op), ("*", op)):
        if key in EXPECT:
            return GOOD | EXPECT[key][0]
    return GOOD


def base_task(subject):
    wd = L.new_workdir()
    try:
        return subject, L.run_case(CASES[SUBJECTS[subject][0]], {}, wd)
    finally:
        shutil.rmtree(wd, ignore_errors=True)


def shift_task(subject, op, base):
    case = CASES[SUBJECTS[subject][0]]
    flag = SUBJECTS[subject][1]
    wd = L.new_workdir()
    try:
        src = os.path.join(ROOT, target(case, flag))
        dst = os.path.join(wd, "shifted.xlsx")
        ok, desc, renames = L.shifted_copy(src, [op], dst)
        if not ok:
            return subject, op, dict(kind="n/a", msg=desc)
        got = L.run_case(case, {target(case, flag): dst}, wd, undo=L.UNDO.get(op), renames=renames)
        kind = L.classify(base, got)
        res = dict(kind=kind)
        if kind in ("STOPPED", "CRASH"):
            res["msg"] = L.message(got)
            res["clear"] = "STOPPED -" in got["text"] and "Traceback" not in got["text"]
        elif kind in ("DIFFERENT", "REMARKS"):
            res["msg"] = f"{got.get('rows')} rows vs {base.get('rows')}"
        return subject, op, res
    finally:
        shutil.rmtree(wd, ignore_errors=True)


results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"), flush=True)


if __name__ == "__main__":
    only = os.environ.get("SHIFT_ONLY")
    subjects = only.split(",") if only else (QUICK if os.environ.get("SHIFT_QUICK") else list(SUBJECTS))
    ops = os.environ.get("SHIFT_OPS", "").split(",") if os.environ.get("SHIFT_OPS") else list(L.OPERATORS)
    workers = int(os.environ.get("SHIFT_WORKERS", "4"))
    with cf.ProcessPoolExecutor(max_workers=workers) as ex:
        bases = dict(f.result() for f in cf.as_completed([ex.submit(base_task, s) for s in subjects]))
        for s in subjects:
            check(f"{s}: the unshifted grid runs (the baseline of its shifts)", "sha" in bases[s], L.message(bases[s]))
        futs = [ex.submit(shift_task, s, o, bases[s]) for s in subjects if "sha" in bases[s] for o in ops]
        for f in cf.as_completed(futs):
            s, o, res = f.result()
            kind = res["kind"]
            if kind == "n/a":
                continue
            ok = kind in allowed_for(s, o) and (kind != "STOPPED" or res.get("clear"))
            check(f"{s} / {o} ({L.OPERATORS[o][0]}): {kind}" + (f" - {res['msg']}" if res.get("msg") and kind in ("STOPPED", "CRASH") else ""),
                  ok, f"allowed {sorted(allowed_for(s, o))}; {res.get('msg', '')}")
    bad = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
    sys.exit(1 if bad else 0)
