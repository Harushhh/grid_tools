"""Structure-shift operators and a runner for the real grids (11 Nov 2026).

The user (10-11 Nov 2026): the insurers change their grids every month - sheets renamed, columns moved, rows added - and the Go Digit November workbook broke a script that read by position.  Every converter
must either give the SAME upload for a shifted grid or STOP with a message that says what it expected and what it found; it must never write a different file in silence.

`OPERATORS` change a workbook the way an insurer might without changing what its tables MEAN (blank rows or columns, title rows, renamed or reordered sheets, extra sheets, whitespace noise in the headers, an
extra column at the right edge, two columns swapped together with their headers).  `run_subject` runs a script on the shifted grid exactly as the regression run does (but with --no-prompt: a new question is a
stop, never a wrong answer) and classifies the result:
    IDENTICAL   the same bytes as the unshifted run            (resilient)
    REORDERED   the same rows in another order                 (the portal does not care: acceptable)
    REMARKS     only the remarks column differs                (the flags changed: look at them)
    STOPPED     no file, a message                             (fail-safe if the message says what was expected and found)
    CRASH       no file, a Python traceback                    (fail-safe, but no diagnostic)
    DIFFERENT   a file whose values differ                     (SILENT CORRUPTION - the thing to remove)
Merged cells move with the cells, as they do in Excel (openpyxl alone leaves them where they were)."""
import hashlib
import io
import os
import re
import subprocess
import sys
import tempfile

import openpyxl
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable


# ---------------------------------------------------------------- cell-level helpers
def insert(ws, kind, idx, amount=1):
    """insert rows / columns the way Excel does: merged cells move with the cells"""
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


def used_range(ws):
    """(first row, first col, last row, last col) of the cells that hold a value, or None for an empty sheet"""
    rows = [c.row for row in ws.iter_rows() for c in row if c.value not in (None, "")]
    if not rows:
        return None
    cols = [c.column for row in ws.iter_rows() for c in row if c.value not in (None, "")]
    return min(rows), min(cols), max(rows), max(cols)


def header_rows(ws, limit=15, top=3):
    """the rows of the first `limit` that hold the most text cells (>= 3): the header rows of the table, as far as a script-free guess can tell"""
    counts = []
    for r in range(1, min(ws.max_row, limit) + 1):
        n = sum(1 for c in ws[r] if isinstance(c.value, str) and c.value.strip())
        if n >= 3:
            counts.append((-n, r))
    return [r for _, r in sorted(counts)[:top]]


def merged_columns(ws):
    cols = set()
    for m in ws.merged_cells.ranges:
        cols.update(range(m.min_col, m.max_col + 1))
    return cols


# ---------------------------------------------------------------- the operators (each takes a workbook, returns True when it changed something)
def op_rename_mild(wb):
    seen, renames = set(), {}
    for ws in wb.worksheets:
        old = ws.title
        new = (old.rstrip() + " v2")[:31]
        while new in seen:
            new = new[:-1] + "x"
        seen.add(new)
        ws.title = new
        renames[new] = old
    wb._renames = {**getattr(wb, "_renames", {}), **renames}
    return True


def op_rename_harsh(wb):
    renames = {}
    for i, ws in enumerate(wb.worksheets):
        old = ws.title
        ws.title = f"Sheet{i + 1}"
        renames[ws.title] = old
    wb._renames = {**getattr(wb, "_renames", {}), **renames}
    return True


def op_reorder(wb):
    if len(wb.worksheets) < 2:
        return False
    wb.move_sheet(wb.worksheets[0], offset=len(wb.worksheets) - 1)
    return True


def op_rows_above(wb):
    for ws in wb.worksheets:
        insert(ws, "rows", 1, 4)
    return True


def op_title_rows(wb):
    for ws in wb.worksheets:
        insert(ws, "rows", 1, 2)
        ws["A1"] = "Rate grid - November 2026"
        ws["A2"] = "For internal use only"
    return True


def op_cols_left(wb):
    for ws in wb.worksheets:
        insert(ws, "cols", 1, 2)
    return True


def op_col_inside(wb):
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        if u and u[3] >= 4:
            insert(ws, "cols", 3, 1)
            changed = True
    return changed


def op_col_swap(wb):
    """two adjacent columns swapped WITH their headers (every row of the two columns); columns that belong to a merged range are left alone"""
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        if not u:
            continue
        bad = merged_columns(ws)
        for a in range(max(u[1] + 2, 3), u[3]):
            if a in bad or a + 1 in bad:
                continue
            body = [(ws.cell(r, a).value, ws.cell(r, a + 1).value) for r in range(u[0], u[2] + 1)]
            if sum(1 for x, y in body if x not in (None, "") and y not in (None, "")) < 3:
                continue
            for r in range(1, ws.max_row + 1):
                va, vb = ws.cell(r, a).value, ws.cell(r, a + 1).value
                ws.cell(r, a).value, ws.cell(r, a + 1).value = vb, va
            changed = True
            break
    return changed


def op_header_noise(wb):
    changed = False
    for ws in wb.worksheets:
        for r in header_rows(ws):
            for c in ws[r]:
                v = c.value
                if isinstance(v, str) and v.strip() and not re.fullmatch(r"[\d.,%\s-]+", v):
                    c.value = "  ".join(v.strip().split(" ")) + " "
                    changed = True
    return changed


def op_extra_sheets(wb):
    wb.create_sheet("Notes", 0)["A1"] = "Release notes - nothing to read here"
    wb.create_sheet("Summary")["A1"] = "Summary of the month"
    ws = wb.create_sheet("Old version")
    ws["A1"] = "last month"
    ws.sheet_state = "hidden"
    return True


def op_extra_col_right(wb):
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        rows = header_rows(ws)
        if not u or not rows:
            continue
        col = u[3] + 1
        if "MergedCell" in type(ws.cell(rows[0], col)).__name__:              # the right edge lies inside a merged header: put the new column behind it
            col = max([m.max_col for m in ws.merged_cells.ranges] + [col]) + 1
        ws.cell(rows[0], col).value = "Comments"
        for r in range(rows[0] + 1, min(rows[0] + 4, u[2]) + 1):
            if "MergedCell" not in type(ws.cell(r, col)).__name__:
                ws.cell(r, col).value = "see mail"
        changed = True
    return changed


def op_blank_row_mid(wb):
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        rows = header_rows(ws)
        if not u or not rows or u[2] - rows[-1] < 10:
            continue
        insert(ws, "rows", (rows[-1] + u[2]) // 2, 1)
        changed = True
    return changed


def op_blank_row_top(wb):
    """a blank row three rows under the header: inside the table body, wherever the table ends"""
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        rows = header_rows(ws)
        if not u or not rows or u[2] - rows[-1] < 8:
            continue
        insert(ws, "rows", rows[-1] + 3, 1)
        changed = True
    return changed


def _primary(wb):
    """the visible sheet with the most cells that hold a value"""
    best, size = None, 0
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        u = used_range(ws)
        if u and (u[2] - u[0] + 1) * (u[3] - u[1] + 1) > size:
            best, size = ws, (u[2] - u[0] + 1) * (u[3] - u[1] + 1)
    return best


def _decoy(wb, hidden):
    """a copy of the biggest sheet with every number scaled by 0.9: last month's sheet left in the workbook (visible) or an old version nobody deleted (hidden).
    A script that reads the copy instead of the real sheet writes other rates."""
    src = _primary(wb)
    if src is None:
        return False
    cp = wb.copy_worksheet(src)
    cp.title = (src.title.rstrip() + (" (old)" if not hidden else " (hidden old)"))[:31]
    for row in cp.iter_rows():
        for c in row:
            if isinstance(c.value, float) or (isinstance(c.value, int) and not isinstance(c.value, bool) and abs(c.value) > 1):
                c.value = c.value * 0.9
    if hidden:
        cp.sheet_state = "hidden"
    return True


def _dup(wb, front=False, mode="higher"):
    """a column that is there TWICE: a copy of an inner data column is inserted right of it (front=True: LEFT of it) WITH the same heading and HIGHER numbers (a new 'Max' / 'Avg' pair under one
    heading, as Go Digit's November grid had two CD2 columns).  Which of the two is meant is not in the document, so a script must never pick one in silence - the first of two same-headed
    columns is right with the copy on the right and wrong with it on the left.  The user's answer (12 Nov 2026): where the column is a PAYIN the LOWEST of the two is written (the copy is higher,
    so the output is IDENTICAL to the original's, on either side); where it is not (a discount band) the script stops"""
    changed = False
    for ws in wb.worksheets:
        u = used_range(ws)
        rows = header_rows(ws)
        if not u or not rows or u[3] < 5:
            continue
        bottom = max(rows)
        bad = set()                                         # only a merged HEADER makes a column unfit (a note merged across the sheet under the table does not)
        for m in ws.merged_cells.ranges:
            if m.min_row <= bottom and m.max_col > m.min_col:
                bad.update(range(m.min_col, m.max_col + 1))
        for a in range(u[1] + 2, u[3]):
            if a in bad or a + 1 in bad:
                continue
            nums = [r for r in range(bottom + 1, u[2] + 1) if isinstance(ws.cell(r, a).value, (int, float)) and not isinstance(ws.cell(r, a).value, bool)]
            if len(nums) < 5:
                continue
            head = [(r, ws.cell(r, a).value) for r in range(1, bottom + 1) if ws.cell(r, a).value not in (None, "")]
            fraction = all(0 <= ws.cell(r, a).value <= 1 for r in nums)               # a table of fractions (0.25 = 25%) must stay one
            if mode == "lower_replaced":                                              # no twin: the ORIGINAL column carries the lower numbers (what a twin with lower numbers must give)
                for r in nums:
                    ws.cell(r, a).value = ws.cell(r, a).value * 0.5
                changed = True
                break
            insert(ws, "cols", a if front else a + 1, 1)
            new, old_col = (a, a + 1) if front else (a + 1, a)
            for r, v in head:
                ws.cell(r, new).value = v
            for r in nums:
                v = ws.cell(r, old_col).value
                if mode == "lower":
                    ws.cell(r, new).value = v * 0.5                                   # LOWER than the original: the lowest-rate rule writes the copy's numbers
                else:
                    ws.cell(r, new).value = (v + 1) / 2 if fraction else v + 1        # HIGHER than the original (a fraction stays a fraction): the lowest-rate rule keeps the original
            changed = True
            break
    return changed


def op_dup_col(wb, front=False):
    return _dup(wb, front, "higher")


def op_twin_lower(wb, front=False):
    return _dup(wb, front, "lower")


def op_twin_lower_replaced(wb):
    return _dup(wb, False, "lower_replaced")


def op_dup_col_front(wb):
    return op_dup_col(wb, front=True)


def op_decoy_sheet(wb):
    return _decoy(wb, hidden=False)


def op_decoy_hidden(wb):
    return _decoy(wb, hidden=True)


OPERATORS = {
    "rename_mild": ("every sheet renamed (' v2' added)", op_rename_mild),
    "rename_harsh": ("every sheet called Sheet1, Sheet2 ...", op_rename_harsh),
    "reorder": ("the first sheet moved to the end", op_reorder),
    "rows_above": ("4 blank rows above every sheet", op_rows_above),
    "title_rows": ("2 title rows above every sheet", op_title_rows),
    "cols_left": ("2 blank columns left of every sheet", op_cols_left),
    "col_inside": ("a blank column after the 2nd column of every sheet", op_col_inside),
    "col_swap": ("two adjacent columns swapped with their headers", op_col_swap),
    "header_noise": ("double spaces and a trailing space in the header cells", op_header_noise),
    "extra_sheets": ("two unrelated sheets and a hidden one added", op_extra_sheets),
    "extra_col_right": ("a 'Comments' column at the right edge", op_extra_col_right),
    "blank_row_mid": ("a blank row in the middle of the data", op_blank_row_mid),
    "blank_row_top": ("a blank row three rows under the header (inside the table)", op_blank_row_top),
    "decoy_sheet": ("a visible copy of the biggest sheet with other numbers (last month's, left in)", op_decoy_sheet),
    "decoy_hidden": ("a hidden copy of the biggest sheet with other numbers", op_decoy_hidden),
    "dup_col": ("a rate column there twice under the same heading (other numbers)", op_dup_col),
    "dup_col_front": ("the same, the copy LEFT of the original", op_dup_col_front),
}


# ---------------------------------------------------------------- running a script
def values_hash(path):
    frame = pd.read_csv(path, dtype=str, keep_default_na=False)
    if "remarks" in frame.columns:
        frame["remarks"] = ""
    buf = io.StringIO()
    frame.to_csv(buf, index=False, lineterminator="\n")
    return hashlib.sha256(buf.getvalue().encode("utf-8")).hexdigest()


def sorted_hash(path, drop_remarks=True, fold_case=False):
    """sha256 of the CSV's rows sorted (the row order of an upload means nothing), remarks blanked unless drop_remarks is False"""
    frame = pd.read_csv(path, dtype=str, keep_default_na=False)
    if drop_remarks and "remarks" in frame.columns:
        frame["remarks"] = ""
    if "id" in frame.columns:                       # some scripts number their rows: the number follows the order, it is not data
        frame = frame.drop(columns=["id"])
    if fold_case:                                   # the project compares RTO and make names without regard to capitals ('NCR' = 'Ncr'): the spelling of a cluster that two sheets spell differently follows the sheet order
        for c in ("new_rto_list", "new_vehicle_makes"):
            if c in frame.columns:
                frame[c] = frame[c].str.lower()
    lines = sorted("|".join(r) for r in frame.itertuples(index=False, name=None))
    return hashlib.sha256(chr(10).join(lines).encode("utf-8")).hexdigest()


def undo_renames(text, renames):
    """a sheet name echoed in the output ('New changes found ...: PAN India - RJ01/...') is the old name again before the comparison"""
    for new in sorted(renames, key=len, reverse=True):
        text = text.replace(new, renames[new])
    return text


UNDO = {"rename_mild": undo_renames, "rename_harsh": undo_renames}      # operators whose effect on the output text is known and taken out again


def run_case(case, swaps, workdir, timeout=1800, undo=None, renames=None, script=None):
    """run the script of a regression case with the files in `swaps` ({original relative path: new path}) instead, --no-prompt, into workdir.
    -> dict(rc, sha, values, rows, text, warnings)"""
    args = [swaps.get(a, a) for a in case["args"]]
    out = os.path.join(workdir, "out.csv")
    cmd = [PY, "-I", script or case["script"], *args, "--no-prompt", "--out", out]
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    if script:                                      # a development copy of the script somewhere else: it still reads the master lists of this project
        env["GRID_MASTER_DIR"] = os.path.join(ROOT, "master")
    p = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env)
    text = p.stdout + "\n" + p.stderr
    res = dict(rc=p.returncode, text=text, warnings=sorted({ln.strip() for ln in text.splitlines() if ln.strip().startswith("WARNING")}))
    if p.returncode == 0 and os.path.exists(out) and undo:
        raw = open(out, "rb").read().decode("utf-8", errors="replace")
        open(out, "wb").write(undo(raw, renames or {}).encode("utf-8"))
    if p.returncode == 0 and os.path.exists(out):
        res["sha"] = hashlib.sha256(open(out, "rb").read()).hexdigest()
        res["values"] = values_hash(out)
        res["sorted_all"] = sorted_hash(out, drop_remarks=False)
        res["sorted_values"] = sorted_hash(out)
        res["sorted_ci"] = sorted_hash(out, drop_remarks=False, fold_case=True)
        res["rows"] = len(pd.read_csv(out, dtype=str, keep_default_na=False))
    return res


def classify(base, got):
    """IDENTICAL / REMARKS / STOPPED / CRASH / DIFFERENT"""
    if got["rc"] != 0 or "sha" not in got:
        return "CRASH" if "Traceback" in got["text"] else "STOPPED"
    if got["sha"] == base["sha"]:
        return "IDENTICAL"
    if got["sorted_all"] == base["sorted_all"] or got["sorted_ci"] == base["sorted_ci"]:
        return "REORDERED"
    if got["values"] == base["values"] or got["sorted_values"] == base["sorted_values"]:
        return "REMARKS"
    return "DIFFERENT"


def message(got, width=260):
    """the last meaningful lines of a failed run"""
    lines = [ln.strip() for ln in got["text"].splitlines() if ln.strip()]
    keep = [ln for ln in lines if "STOPPED" in ln or "Error" in ln or "expected" in ln.lower()] or lines[-2:]
    return " | ".join(keep[-2:])[:width]


def shifted_copy(src, op_names, dst):
    """load `src` (values only), apply the operators in order, save to `dst`; -> (ok, description, {new sheet name: old sheet name} of the renames)"""
    wb = openpyxl.load_workbook(src, data_only=True)
    done = []
    for name in op_names:
        try:
            if OPERATORS[name][1](wb):
                done.append(name)
        except Exception as e:                      # an operator that cannot be applied to this workbook is not a finding about the script
            return False, f"operator {name} not applicable: {type(e).__name__}: {e}", {}
    if not done:
        return False, "operator not applicable", {}
    wb.save(dst)
    return True, "+".join(done), dict(getattr(wb, "_renames", {}))


def new_workdir():
    return tempfile.mkdtemp(prefix="shift-")
