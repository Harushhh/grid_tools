"""Black-box audit: which parts of a grid does a converter actually USE?  (18 Nov 2026, the user: "at some places there is some make or other conditions which is not applied,
also it is not mentioned anywhere".)

The Go Digit audit found its silent errors by tracing every wording to the rows written.  This tool finds the same class of error WITHOUT reading the script: it changes the cells of
a regression grid and watches the upload.

    python tests/unread_audit.py <case> --mode sheets         # every sheet, all its cells changed at once: is the sheet read at all?
    python tests/unread_audit.py <case> --mode columns        # every column of every sheet, all its cells changed at once
    python tests/unread_audit.py <case> --mode values         # every distinct text value, all the cells that carry it changed at once
    python tests/unread_audit.py <case> --mode rows           # every row of the named sheet(s), all its cells changed at once (--sheet)

A change = a number x 1.1 + 0.0123, a text + ' ZQX'.  The converter is run exactly as the regression run does it (same arguments, same answers) and the result is compared with the
unchanged run:
    READ        the upload changed (or the run stopped / crashed): the converter noticed the cell
    MENTIONED   the upload is identical but the run's own text changed (a warning that names the cell): it is not applied, but it is not silent
    SILENT      nothing changed anywhere: the converter never looked at the cell, or looked and ignored it.  THE FINDINGS.
Several workbooks of one case (ICICI: four, Shriram two-wheeler: grid + broker grid) are all audited.  A text that the upload echoes (a cluster name in new_rto_list) is READ by
definition and is skipped in the 'values' mode.  PDF and .xlsb inputs are not audited (United / the Magma SATP grid).

The results are kept as JSON next to the scratch files; `--summary` prints the findings only."""
import argparse
import concurrent.futures as cf
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

import openpyxl
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable
MARK = " ZQX"


def load_case(name):
    cases = json.load(open(os.path.join(ROOT, "regression", "cases.json"), encoding="utf-8"))["cases"]
    for c in cases:
        if c["name"] == name:
            return c
    raise SystemExit(f"no case named {name}")


def workbook_args(case):
    """[(position in args, path)] of the .xlsx inputs of a case"""
    out = []
    for i, a in enumerate(case["args"]):
        if isinstance(a, str) and a.lower().endswith((".xlsx", ".xlsm")) and os.path.exists(os.path.join(ROOT, a)):
            out.append((i, a))
    return out


def run(case, swaps, workdir, timeout=1800):
    """run the case's script with the files of `swaps` ({arg position: path}) instead; -> dict(rc, text, rows (sorted row strings), sha)"""
    args = list(case["args"])
    for i, p in swaps.items():
        args[i] = p
    out = os.path.join(workdir, "out.csv")
    cmd = [PY, "-I", "-X", "utf8", case["script"], *args, "--out", out]
    answers = case.get("answers") or []
    if not answers:
        cmd.append("--no-prompt")
    env = dict(os.environ, PYTHONIOENCODING="utf-8", GRID_MASTER_DIR=os.path.join(ROOT, "master"))
    p = subprocess.run(cmd, cwd=ROOT, env=env, input="\n".join(answers) + "\n", capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)
    text = (p.stdout + "\n" + p.stderr).replace(workdir, "<W>").replace(workdir.replace("\\", "/"), "<W>")
    res = dict(rc=p.returncode, text=text)
    if p.returncode == 0 and os.path.exists(out):
        frame = pd.read_csv(out, dtype=str, keep_default_na=False)
        if "id" in frame.columns:
            frame = frame.drop(columns=["id"])
        res["rows"] = sorted("|".join(r) for r in frame.itertuples(index=False, name=None))
        res["sha"] = hashlib.sha256("\n".join(res["rows"]).encode("utf-8")).hexdigest()
        rv = os.path.join(workdir, "out_review.xlsx")                          # the review workbook: a change there is a MENTION (a list the person is shown)
        if os.path.exists(rv):
            wbr = openpyxl.load_workbook(rv, read_only=True, data_only=True)
            parts = []
            for ws in wbr.worksheets:
                parts.append("## " + ws.title)
                for row in ws.iter_rows(values_only=True):
                    parts.append("|".join("" if v is None else str(v) for v in row).replace(workdir, "<W>"))
            wbr.close()
            res["review"] = hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()
    return res


def normalise_text(text):
    """the run's own text without what changes between two identical runs (timings, scratch paths)"""
    keep = []
    for ln in text.splitlines():
        if re.search(r"\d+(\.\d+)?\s*(s|sec|seconds)\b", ln) and "Reading" in ln:
            continue
        keep.append(ln.rstrip())
    return "\n".join(keep)


def cells_of(ws):
    """{(row, col): value} of the non-empty cells of a sheet"""
    out = {}
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            if isinstance(v, bool):
                continue
            out[(c.row, c.column)] = v
    return out


def change(v):
    if isinstance(v, (int, float)):
        return v * 1.1 + 0.0123
    return str(v) + MARK


_WB = {}


def make_copy(src, dst, sheet, coords):
    """save a copy of `src` with the cells changed.  The workbook is loaded once per worker; the cells are put back after the save, so the next unit starts from the original."""
    wb = _WB.get(src)
    if wb is None:
        wb = _WB[src] = openpyxl.load_workbook(os.path.join(ROOT, src), data_only=True)
    ws = wb[sheet]
    saved = []
    for (r, c) in coords:
        cell = ws.cell(r, c)
        saved.append((cell, cell.value))
        cell.value = change(cell.value)
    try:
        try:
            wb.save(dst)
        except ValueError:                                          # "I/O operation on closed file": openpyxl closes the archive of the loaded file on the first save - load it again
            _WB.pop(src, None)
            wb2 = openpyxl.load_workbook(os.path.join(ROOT, src), data_only=True)
            ws2 = wb2[sheet]
            for (r, c) in coords:
                ws2.cell(r, c).value = change(ws2.cell(r, c).value)
            wb2.save(dst)
            return
    finally:
        for cell, v in saved:
            cell.value = v


def classify(base, got, base_text):
    if got["rc"] != 0 or "sha" not in got:
        return "READ", "the run stopped: " + " | ".join([ln.strip() for ln in got["text"].splitlines() if ln.strip()][-1:])[:160]
    if got["sha"] != base["sha"]:
        a, b = set(base["rows"]), set(got["rows"])
        return "READ", f"{len(a - b)} rows gone, {len(b - a)} new"
    if got.get("review") != base.get("review"):
        return "MENTIONED", "the review workbook changed (a list the person is shown)"
    if normalise_text(got["text"]) != base_text:
        a = set(base_text.splitlines())
        extra = [ln.strip() for ln in normalise_text(got["text"]).splitlines() if ln not in a and ln.strip()]
        return "MENTIONED", (extra[0][:200] if extra else "the run's text changed")
    return "SILENT", ""


def label(v):
    return re.sub(r"\s+", " ", str(v)).strip()[:60]


_G = {}


def _init(case, base_path):
    """worker start: the case and the unchanged run's result (a pickle: 20,000 row strings are not sent with every task)"""
    import pickle
    _G["case"] = case
    _G["base"], _G["base_text"] = pickle.load(open(base_path, "rb"))
    _G["base"]["rows"] = set(_G["base"]["rows"])


def job(unit):
    src_index, src, sheet, coords, desc, tag = unit
    case, base, base_text = _G["case"], _G["base"], _G["base_text"]
    wd = tempfile.mkdtemp(prefix="unread-")
    try:
        dst = os.path.join(wd, os.path.basename(src))          # the same file name: the run prints it
        make_copy(src, dst, sheet, coords)
        got = run(case, {src_index: dst}, wd)
        status, detail = classify(base, got, base_text)
        return dict(workbook=os.path.basename(src), sheet=sheet, unit=desc, cells=len(coords), status=status, detail=detail, tag=tag)
    except Exception as e:                                                   # a unit that cannot be tried is not a finding about the converter
        return dict(workbook=os.path.basename(src), sheet=sheet, unit=desc, cells=len(coords), status="ERROR", detail=f"{type(e).__name__}: {e}"[:200], tag=tag)
    finally:
        shutil.rmtree(wd, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("case")
    ap.add_argument("--mode", choices=["sheets", "columns", "values", "rows"], default="columns")
    ap.add_argument("--sheet", action="append", help="only these sheets")
    ap.add_argument("--jobs", type=int, default=8)
    ap.add_argument("--min-len", type=int, default=3, help="values mode: shortest text tried")
    ap.add_argument("--limit", type=int, default=0, help="stop after this many units (a trial)")
    ap.add_argument("--out", default="")
    ap.add_argument("--hidden", action="store_true", help="also audit hidden sheets")
    ap.add_argument("--workbook", action="append", help="only workbooks whose path contains this text")
    a = ap.parse_args()
    case = load_case(a.case)
    books = workbook_args(case)
    if a.workbook:
        books = [(i, p) for i, p in books if any(w in p for w in a.workbook)]
    if not books:
        raise SystemExit("this case has no .xlsx input to audit")
    t0 = time.time()
    wd = tempfile.mkdtemp(prefix="unread-base-")
    base = run(case, {}, wd)
    shutil.rmtree(wd, ignore_errors=True)
    if base["rc"] != 0 or "sha" not in base:
        raise SystemExit("the unchanged run failed:\n" + base["text"][-600:])
    base_text = normalise_text(base["text"])
    echo = "\n".join(base["rows"]).lower()
    print(f"baseline: {len(base['rows']):,} rows in {time.time() - t0:.0f}s; {len(books)} workbook(s)", flush=True)
    units = []
    for i, src in books:
        wb = openpyxl.load_workbook(os.path.join(ROOT, src), data_only=True)
        for ws in wb.worksheets:
            if a.sheet and ws.title not in a.sheet:
                continue
            if ws.sheet_state != "visible" and not a.hidden:
                continue
            cells = cells_of(ws)
            if not cells:
                continue
            if a.mode == "sheets":
                sample = [label(v) for v in list(cells.values())[:4]]
                units.append((i, src, ws.title, list(cells), f"the whole sheet ({len(cells):,} cells; {'; '.join(sample)})", "sheet"))
            elif a.mode == "columns":
                by = {}
                for (r, c), v in cells.items():
                    by.setdefault(c, []).append((r, c))
                for c, coords in sorted(by.items()):
                    sample = [label(cells[x]) for x in sorted(coords)[:3]]
                    units.append((i, src, ws.title, coords, f"column {openpyxl.utils.get_column_letter(c)} (e.g. {'; '.join(sample)})", "col"))
            elif a.mode == "rows":
                by = {}
                for (r, c), v in cells.items():
                    by.setdefault(r, []).append((r, c))
                for r, coords in sorted(by.items()):
                    sample = [label(cells[x]) for x in sorted(coords)[:4]]
                    units.append((i, src, ws.title, coords, f"row {r} ({'; '.join(sample)})", "row"))
            else:
                by = {}
                for (r, c), v in cells.items():
                    if isinstance(v, str) and len(v.strip()) >= a.min_len and not re.fullmatch(r"[-+0-9.,% ]+", v.strip()):
                        by.setdefault(re.sub(r"\s+", " ", v.strip()), []).append((r, c))
                for text, coords in sorted(by.items(), key=lambda kv: kv[0].lower()):
                    if text.lower() in echo:                                # the upload carries this very text: it is read
                        continue
                    units.append((i, src, ws.title, coords, f"{text[:90]!r} x{len(coords)}", "val"))
    if a.limit:
        units = units[:a.limit]
    print(f"{len(units)} units to try ({a.mode}) with {a.jobs} workers", flush=True)
    import pickle
    base_path = os.path.join(tempfile.gettempdir(), f"unread_base_{os.getpid()}.pkl")
    pickle.dump((base, base_text), open(base_path, "wb"))
    results = []
    with cf.ProcessPoolExecutor(max_workers=a.jobs, initializer=_init, initargs=(case, base_path)) as ex:
        futs = [ex.submit(job, (i, src, sh, co, d, tg)) for i, src, sh, co, d, tg in units]
        done = 0
        for f in cf.as_completed(futs):
            results.append(f.result())
            done += 1
            if done % 25 == 0:
                print(f"  {done}/{len(units)} done ({time.time() - t0:.0f}s)", flush=True)
    try:
        os.remove(base_path)
    except OSError:
        pass
    order = {"SILENT": 0, "MENTIONED": 1, "ERROR": 2, "READ": 3}
    results.sort(key=lambda r: (order.get(r["status"], 9), r["workbook"], r["sheet"], r["unit"]))
    out = a.out or os.path.join(tempfile.gettempdir(), f"unread_{a.case}_{a.mode}.json")
    json.dump(results, open(out, "w", encoding="utf-8"), indent=1)
    counts = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    print(f"\n{a.case} / {a.mode}: {counts}  ({time.time() - t0:.0f}s)  -> {out}")
    for r in results:
        if r["status"] in ("SILENT", "MENTIONED", "ERROR"):
            print(f"  {r['status']:<9} [{r['workbook']}] {r['sheet']!r}: {r['unit']}  cells={r['cells']}" + (f"  -> {r['detail']}" if r["detail"] else ""))


if __name__ == "__main__":
    main()
