"""Print the names of the two master workbooks (RTO groups, make / model groups) as JSON, found the way the converters find them:
the newest .xlsx in the master folder whose name has 'rto' + 'master' (resp. 'make' + 'model' + 'master').

    python dump_masters.py <master folder>
"""
import json
import os
import sys

import openpyxl


def find(folder, words):
    try:
        names = [n for n in os.listdir(folder) if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$") and all(w in n.lower() for w in words)]
    except OSError:
        return None
    return max((os.path.join(folder, n) for n in names), key=os.path.getmtime) if names else None


def names_of(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    return [str(r[1]).strip() for r in rows[1:] if r and len(r) >= 3 and r[1] is not None]


def main():
    folder = sys.argv[1]
    rto, mk = find(folder, ("rto", "master")), find(folder, ("make", "model", "master"))
    sys.stdout.reconfigure(encoding="utf-8")
    json.dump({"rto": names_of(rto) if rto else [], "makes": names_of(mk) if mk else [], "files": [rto, mk]}, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    main()
