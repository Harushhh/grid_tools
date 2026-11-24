"""Shared setup of the test suites: load a script by file name, and rebuild - in memory, from the known-good grids in
regression\\inputs - the sheets the tests then change one thing at a time. (Nothing here is stored as a pickle: the grids are
the single source.)"""
import importlib.util
import math
import os

import pandas as pd

# The suites test the converters in the group-name mode (new_rto_list = the portal's group name, checked against the RTO master); cluster_names_tests.py sets "cluster" before it loads a script.
os.environ.setdefault("GRID_RTO_NAMES", "group")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
INPUTS = os.path.join(ROOT, "regression", "inputs")
REFS = os.path.join(ROOT, "regression", "references")


def load_script(file):
    spec = importlib.util.spec_from_file_location(file[:-3], os.path.join(ROOT, file))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


class FakeSheet:
    def __init__(self, rows, state="visible"):
        self.rows, self.sheet_state = rows, state

    def iter_rows(self, values_only=True):
        return iter([tuple(r) for r in self.rows])


class FakeWB:
    """Stands in for an openpyxl workbook: sheet names, sheet_state and iter_rows(values_only=True) are all the scripts use.
    hidden = names of the sheets that are hidden."""

    def __init__(self, sheets, hidden=()):
        self.sheets, self.sheetnames, self.hidden = sheets, list(sheets), set(hidden)

    def __getitem__(self, name):
        return FakeSheet(self.sheets[name], "hidden" if name in self.hidden else "visible")


def _trim(rows):
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    width = max(len(r) for r in rows)
    rows = [list(r) + [None] * (width - len(r)) for r in rows]
    while all(r[-1] is None for r in rows):
        rows = [r[:-1] for r in rows]
    return rows


def sheets_of(path, trim=False, read_only=False):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=read_only)
    out = {}
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        out[ws.title] = _trim(rows) if trim else rows
    return out


def rsa_base(which="oct"):
    """The RSA workbook as {sheet: rows}: 'oct' = the October grid (V1), 'v2' = the grid effective 22 Oct."""
    return sheets_of(os.path.join(INPUTS, {"oct": "rsa_gcv4w_oct26.xlsx", "v2": "rsa_gcv4w_v2_oct22.xlsx"}[which]), trim=True)


def pvtcar_base(which="oct"):
    """The RSA Private Car workbook as {sheet: rows}: 'oct' = October 2026 (V1, flat "Additional 2%" note), 'sep' = September 2026
    (V1, the "Top Variant Models Drive" block)."""
    return sheets_of(os.path.join(INPUTS, {"oct": "rsa_pvtcar_oct26.xlsx", "sep": "rsa_pvtcar_sep26.xlsx"}[which]), trim=True)


def liberty_base(which="sep"):
    """The Liberty workbook as {sheet: rows} (loaded the way the script loads it: not read-only, so merged cells are clean)."""
    return sheets_of(os.path.join(INPUTS, {"sep": "liberty_sep26.xlsx", "jul": "liberty_jul26.xlsx"}[which]))


def magma_base():
    """(comprehensive sheets, SATP sheets) as {sheet: rows}: only the two grid sheets the script reads."""
    import openpyxl
    from pyxlsb import open_workbook
    wb = openpyxl.load_workbook(os.path.join(INPUTS, "magma_oct26_comprehensive.xlsx"), data_only=True, read_only=True)
    motor = {"< =3 lac": [list(r) for r in wb["< =3 lac"].iter_rows(values_only=True)]}
    with open_workbook(os.path.join(INPUTS, "magma_oct26_satp.xlsb")) as sb:
        with sb.get_sheet("<=2 Lacs") as sh:
            satp = {"<=2 Lacs": [[c.v for c in row] for row in sh.rows()]}
    return motor, satp


def check_csv(path, columns, numeric):
    """Format problems of a written upload file (empty list = fine)."""
    d = pd.read_csv(path, dtype=str, keep_default_na=False)
    problems = []
    if list(d.columns) != list(columns):
        problems.append("header differs")
    if len(d) == 0:
        problems.append("no rows")
    low = d.drop(columns=["new_rto_list"], errors="ignore").astype(str).apply(lambda s: s.str.lower())      # a geo may really be called "inf"
    for bad in ("nan", "inf", "-inf", "none"):
        if (low == bad).any().any():
            problems.append(f"'{bad}' text in the file")
    for c in numeric:
        if c in d:
            v = pd.to_numeric(d[c], errors="coerce")
            if v.isna().any() or not v.map(math.isfinite).all():
                problems.append(f"column {c} has non-numbers")
    for c in ("from_date", "to_date"):
        if c in d and not d[c].str.fullmatch(r"\d{2}/\d{2}/\d{4}").all():
            problems.append(f"{c} not DD/MM/YYYY")
    if "remarks" in d:
        odd = d.remarks[(d.remarks != "") & ~d.remarks.str.startswith("Manual Review Required: [")]
        if len(odd):
            problems.append(f"{len(odd)} remarks that are not a 'Manual Review Required: [..]' flag")
    return problems
