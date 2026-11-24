"""United India structure-shift tests (11 Nov 2026).  The circular is a PDF (or, for a scan, a transcription .json of its tables), so the mutations are made on the October 2026 transcription
(regression/inputs/united_oct26.json: 6 tables, every Motor layout): the same circular with its structure changed in a way that does not change what it says must give the SAME upload, and a table
the script finds but cannot read must STOP the run with what was expected and what was found - never be left out of the upload with a warning.

  heading wordings      the new column called 'Revised Payout', 'New Pay-out', 'Proposed Commission', 'Revised Brokerage %' ...   -> identical
  unknown heading       'Updated rate (w.e.f. ...)'                                                                              -> STOPPED, says what it expected
  tables reordered      the tables in the opposite order                                                                         -> the same rows (any order)
  extra table in front  a Fire department table (no Existing / new Pay-out columns)                                              -> identical
  extra column right    a 'Remarks' column after the last one (the full-width titles widen with it)                                -> identical
  extra column left     a 'Sr' column before the first one                                                                         -> identical
  Private Car heading   the Private Car table with a heading the script does not know                                              -> STOPPED

Run:  .venv\\Scripts\\python.exe tests\\united_shift_tests.py"""
import copy
import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import shiftlib as L  # noqa: E402

ROOT = L.ROOT
SOURCE = os.path.join(ROOT, "regression", "inputs", "united_oct26.json")
CASE = {"script": "united_commission_to_config.py", "args": ["--grid", "regression/inputs/united_oct26.json", "--from-date", "26/10/2026", "--to-date", "31/03/2030"]}
doc = json.load(open(SOURCE, encoding="utf-8"))
results = []
SCRIPT = os.environ.get("UNITED_SCRIPT")          # a development copy of the script to try instead of the real one


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"), flush=True)


def cells_of(table):
    for row in table["rows"]:
        for i, cell in enumerate(row):
            yield row, i, cell


def retext(d, old, new):
    """every cell whose text starts with `old` gets `new` in its place (the date part stays)"""
    out = copy.deepcopy(d)
    n = 0
    for t in out["tables"]:
        for row, i, cell in cells_of(t):
            text = cell if isinstance(cell, str) else cell["t"]
            if text.startswith(old):
                new_text = new + text[len(old):]
                row[i] = new_text if isinstance(cell, str) else dict(cell, t=new_text)
                n += 1
    return out, n


def op_reorder_tables(d):
    out = copy.deepcopy(d)
    out["tables"].reverse()
    return out


def op_extra_table_front(d):
    out = copy.deepcopy(d)
    out["tables"].insert(0, {"page": 2, "columns": 4, "rows": [[{"t": "INTERMEDIARY COMMISSION/BROKERAGE FIRE DEPARTMENT", "cs": 4}], ["Sr", "Class of Business", "Existing", "Revised"],
                                                              ["1", "Fire - Standard policies", "10%", "12.5%"], ["2", "Engineering", "7.5%", "7.5%"]]})
    return out


def header_row(table):
    """index of the row that holds the column headings (the one with 'S No'); the rows above it are full-width titles"""
    for r, row in enumerate(table["rows"]):
        if any((c if isinstance(c, str) else c["t"]).lower().replace(" ", "") in ("sno", "sr") for c in row):
            return r
    return 1


def op_extra_col_right(d):
    out = copy.deepcopy(d)
    for t in out["tables"]:
        old = t["columns"]
        t["columns"] = old + 1
        for r, row in enumerate(t["rows"]):
            for i, cell in enumerate(row):
                if not isinstance(cell, str) and int(cell.get("cs", 1)) == old:
                    cell["cs"] = old + 1                                       # a full-width title is still full width
        h = header_row(t)
        t["rows"][h].append({"t": "Remarks", "rs": len(t["rows"]) - h})          # a column down the table from its header row on
    return out


def op_extra_col_left(d):
    out = copy.deepcopy(d)
    for t in out["tables"]:
        old = t["columns"]
        t["columns"] = old + 1
        for row in t["rows"]:
            for i, cell in enumerate(row):
                if not isinstance(cell, str) and int(cell.get("cs", 1)) == old:
                    cell["cs"] = old + 1
        h = header_row(t)
        t["rows"][h].insert(0, {"t": "Sr", "rs": len(t["rows"]) - h})            # a column down the table from its header row on, under the full-width titles
    return out


def run(d, workdir, name):
    path = os.path.join(workdir, name + ".json")
    json.dump(d, open(path, "w", encoding="utf-8"), ensure_ascii=False)
    sub = os.path.join(workdir, name)
    os.makedirs(sub, exist_ok=True)
    return L.run_case(CASE, {"regression/inputs/united_oct26.json": path}, sub, script=SCRIPT)


def main():
    work = tempfile.mkdtemp(prefix="united_shift_")
    try:
        os.makedirs(os.path.join(work, "base"), exist_ok=True)
        base = L.run_case(CASE, {}, os.path.join(work, "base"), script=SCRIPT)
        check("united: the unshifted transcription runs", "sha" in base, L.message(base))
        if "sha" not in base:
            return
        good = {"IDENTICAL", "REORDERED"}
        for label, new in (("Revised Payout", "Revised Payout"), ("New Pay-out", "New Pay-out"), ("Proposed Commission", "Proposed Commission"),
                           ("Revised Brokerage %", "Revised Brokerage %"), ("Updated Payout", "Updated Payout")):
            d, n = retext(doc, "Revised Pay-out", new)
            got = run(d, work, "head_" + label.replace(" ", "_").replace("%", "pct"))
            kind = L.classify(base, got)
            check(f"united / the new column headed '{label}' ({n} headings): {kind}", n > 0 and kind in good, L.message(got))
        d, n = retext(doc, "Revised Pay-out", "Updated rate")
        got = run(d, work, "head_unknown")
        msg = L.message(got)
        check("united / an unknown heading of the new column STOPS and says what it expected", L.classify(base, got) == "STOPPED" and "expected" in got["text"] and "NEW payout column" in got["text"], msg)
        for name, fn in (("tables reordered", op_reorder_tables), ("extra table in front", op_extra_table_front), ("extra column right", op_extra_col_right),
                         ("extra column left", op_extra_col_left)):
            got = run(fn(doc), work, name.replace(" ", "_"))
            kind = L.classify(base, got)
            check(f"united / {name}: {kind}", kind in good, L.message(got))
        d = copy.deepcopy(doc)
        for t in d["tables"]:                                                    # the Private Car table: its column headings are not the ones the script knows
            for row, i, cell in cells_of(t):
                text = cell if isinstance(cell, str) else cell["t"]
                if text in ("Policy Type", "Segment"):
                    row[i] = "Kind" if text == "Policy Type" else "Band"
        got = run(d, work, "car_headings")
        check("united / a Private Car table with unknown column headings STOPS and says what it expected", L.classify(base, got) == "STOPPED" and "Private Car" in got["text"] and "expected" in got["text"], L.message(got))
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
    bad = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
    sys.exit(1 if bad else 0)
