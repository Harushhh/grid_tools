"""Twin columns (the user, 12 Nov 2026: 'Lowest payin'): a PAYIN column that is there twice under ONE heading is not a stop - both columns are read and the user's lowest-rate rule of 27 Oct 2026
keeps the LOWEST payin of each entry (and lists it in the sheet 'Lowest rate chosen'); a blank cell of one twin is no payin.  Cases, each on the real grid of regression\\inputs:

  ICICI Private Car / CV / Two-Wheeler   a column copied beside itself (right and left) with LOWER numbers must give the output of the grid whose original column carries those lower numbers
  Shriram broker '% PO', two-wheeler 'Average Net PO (Converted)'   the same, and with HIGHER numbers the output is the original's (identical)
  ICICI new Two-Wheeler                  a second payout column with the SAME heading: the same two checks; with ANOTHER heading ('Last month') the run STOPS (which month is meant is not in the sheet)
  Shriram 'DIS %'                        the same heading twice for something that is no payin: STOPS (there is no 'lowest' of two discount bands)

The generic shift suite (tests/structure_shift_tests.py, operators dup_col / dup_col_front with a HIGHER copy) covers the same grids' output identity; RSA settles a repeated cluster column by the same rule.

Run:  .venv\\Scripts\\python.exe tests\\twin_column_tests.py"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl  # noqa: E402
import shiftlib as L  # noqa: E402

ROOT = L.ROOT
CASES = {c["name"]: c for c in json.load(open(os.path.join(ROOT, "regression", "cases.json"), encoding="utf-8"))["cases"]}
SAME = {"IDENTICAL", "REORDERED", "REMARKS"}          # the same rows and rates (the remarks column may carry a note)
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"), flush=True)


def arg(case, flag):
    a = case["args"]
    return a[a.index(flag) + 1]


def named_twin(text, sheet=None, where="back", mode="higher"):
    """a mutation: the first column whose header cell says `text` (in `sheet`, else in any sheet) copied beside itself with the same heading texts; mode higher / lower / replaced (no copy: the
    original column carries the lower numbers)"""
    def mutate(wb):
        for ws in wb.worksheets:
            if sheet and ws.title != sheet:
                continue
            for hr in range(1, 8):
                col = next((c.column for c in ws[hr] if isinstance(c.value, str) and c.value.strip().lower() == text.lower()), None)
                if col is None:
                    continue
                nums = [r for r in range(hr + 1, ws.max_row + 1) if isinstance(ws.cell(r, col).value, (int, float)) and not isinstance(ws.cell(r, col).value, bool)]
                if len(nums) < 5:
                    continue
                fraction = all(0 <= ws.cell(r, col).value <= 1 for r in nums)
                if mode == "replaced":
                    for r in nums:
                        ws.cell(r, col).value = ws.cell(r, col).value * 0.5
                    return True
                head = [(r, ws.cell(r, col).value) for r in range(1, hr + 1) if ws.cell(r, col).value not in (None, "")]
                new, old = (col, col + 1) if where == "front" else (col + 1, col)
                L.insert(ws, "cols", new, 1)
                for r, v in head:
                    ws.cell(r, new).value = v
                for r in nums:
                    v = ws.cell(r, old).value
                    ws.cell(r, new).value = v * 0.5 if mode == "lower" else ((v + 1) / 2 if fraction else v + 1)
                return True
        return False
    return mutate


def generic_twin(mode, where="back"):
    """the first inner numeric column of every sheet (shiftlib's duplicate operator)"""
    ops = {"replaced": lambda wb: L.op_twin_lower_replaced(wb), "lower": lambda wb: L.op_twin_lower(wb, where == "front"), "higher": lambda wb: L.op_dup_col(wb, where == "front")}
    return ops[mode]


def payout_twin(same_heading, mode="higher", where="back"):
    """ICICI new Two-Wheeler: a second numeric column right of the payout column of every sheet (same heading or 'Last month')"""
    def mutate(wb):
        done = False
        for ws in wb.worksheets:
            head = bike = None
            for r in range(1, 15):
                for c in ws[r]:
                    if isinstance(c.value, str) and c.value.strip().lower().startswith("bike"):
                        head, bike = r, c.column
                        break
                if head:
                    break
            if head is None:
                continue
            pay = bike + 1
            nums = [r for r in range(head + 1, ws.max_row + 1) if isinstance(ws.cell(r, pay).value, (int, float))]
            if len(nums) < 3:
                continue
            fraction = all(0 <= ws.cell(r, pay).value <= 1 for r in nums)
            if mode == "replaced":
                for r in nums:
                    ws.cell(r, pay).value = ws.cell(r, pay).value * 0.5
                done = True
                continue
            new, old = (pay, pay + 1) if where == "front" else (pay + 1, pay)
            L.insert(ws, "cols", new, 1)
            for r in range(1, head + 1):
                if ws.cell(r, old).value not in (None, ""):
                    ws.cell(r, new).value = ws.cell(r, old).value if same_heading else ("Last month" if r == head else None)
            for r in nums:
                v = ws.cell(r, old).value
                ws.cell(r, new).value = v * 0.5 if mode == "lower" else ((v + 1) / 2 if fraction else v + 1)
            done = True
        return done
    return mutate


def run(case_name, flag, mutate, work, tag):
    case = CASES[case_name]
    src = arg(case, flag)
    wb = openpyxl.load_workbook(os.path.join(ROOT, src), data_only=True)
    if not mutate(wb):
        return None
    dst = os.path.join(work, tag + ".xlsx")
    wb.save(dst)
    sub = os.path.join(work, tag)
    os.makedirs(sub, exist_ok=True)
    return L.run_case(case, {src: dst}, sub)


def lowest_cases(label, case_name, flag, make, work):
    """higher copy (both sides) = the original's output; lower copy (both sides) = the output of the grid whose original column carries the lower numbers"""
    base_dir = os.path.join(work, label.replace(" ", "_") + "_base")
    os.makedirs(base_dir, exist_ok=True)
    base = L.run_case(CASES[case_name], {}, base_dir)
    for where in ("back", "front"):
        got = run(case_name, flag, make("higher", where), work, f"{label}_{where}_higher".replace(" ", "_"))
        if got is None:
            check(f"{label}: the operator could be applied ({where}, higher)", False, "not applicable")
            continue
        kind = L.classify(base, got)
        check(f"{label}: the column twice, the copy on the {where} with HIGHER numbers: the output is the original's ({kind})", kind in SAME, L.message(got))
    ref = run(case_name, flag, make("replaced", "back"), work, f"{label}_replaced".replace(" ", "_"))
    check(f"{label}: the grid whose column carries the lower numbers runs", ref is not None and "sha" in ref, "" if ref is None else L.message(ref))
    if ref is None or "sha" not in ref:
        return
    for where in ("back", "front"):
        got = run(case_name, flag, make("lower", where), work, f"{label}_{where}_lower".replace(" ", "_"))
        kind = L.classify(ref, got) if got is not None else "n/a"
        check(f"{label}: the column twice, the copy on the {where} with LOWER numbers: the lowest payin is written = the output of the grid with the lower numbers ({kind})", kind in SAME,
              "" if got is None else L.message(got))


def main():
    work = L.new_workdir()
    try:
        lowest_cases("ICICI Private Car", "icici_pc_tw_nov26", "--pc-grid", lambda mode, where: generic_twin(mode, where), work)
        lowest_cases("ICICI CV", "icici_cv_nov26", "--cv-grid", lambda mode, where: generic_twin(mode, where), work)
        lowest_cases("ICICI Two-Wheeler", "icici_pc_tw_nov26", "--tw-grid", lambda mode, where: generic_twin(mode, where), work)
        lowest_cases("ICICI new Two-Wheeler", "icici_pc_tw_nov26", "--tw-new-grid", lambda mode, where: payout_twin(True, mode, where), work)
        lowest_cases("Shriram broker % PO", "shriram_nov26", "--grid", lambda mode, where: named_twin("% PO", None, where, mode), work)
        lowest_cases("Shriram two-wheeler New Business", "shriram_2w_oct26", "--grid", lambda mode, where: named_twin("Average Net PO\n(Converted)", "New Business - 2W", where, mode), work)
        lowest_cases("Shriram two-wheeler Rollover", "shriram_2w_oct26", "--grid", lambda mode, where: named_twin("NPO(%) / FIXED", "ROLLOVER 2W (1+1)", where, mode), work)
        # another heading: which month is meant is not in the sheet
        got = run("icici_pc_tw_nov26", "--tw-new-grid", payout_twin(False), work, "newtw_other_heading")
        check("ICICI new Two-Wheeler: a second payout column with ANOTHER heading ('Last month') STOPS and says what it expected", got is not None and got["rc"] != 0 and "expected ONE payout column" in got["text"]
              and "Traceback" not in got["text"], "" if got is None else L.message(got))
        # no 'lowest' of two discount bands
        got = run("shriram_nov26", "--grid", named_twin("DIS %", None, "back", "higher"), work, "shriram_dis")
        check("Shriram: 'DIS %' there twice STOPS (a discount band is no payin: there is no lowest) and names the columns", got is not None and got["rc"] != 0 and "with the heading 'dis'" in got["text"]
              and "Traceback" not in got["text"], "" if got is None else L.message(got))
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
    bad = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
    sys.exit(1 if bad else 0)
