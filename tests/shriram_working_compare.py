"""Compare what shriram_broker_grid_to_config.py maps for each GRID ROW with the user's own mapping of the same row.

The user's working file (D:\\harsh\\grids\\OCT26\\shriram\\Shriram_-_OCT26_BROKER_GRID.xlsx, 27 MB, sheet "GRID OCT' 26 (W)") carries, beside every grid row, the columns
RTO | MAKE | product | class | cc from/to | SC FROM/TO | age from/to | TARIFF FROM/TO | fuel | CONDITION | TNC | NCB | CPA | ZD.  Those 797 rows are saved in
regression\\references\\shriram_oct26_working_mapping.csv (made by --export). This is the second answer key of Shriram: it shows the user's DECISIONS row by row, before they
are expanded into upload rows.

    .venv\\Scripts\\python.exe tests\\shriram_working_compare.py            # tally of the fields where the script and the user differ, with examples
    .venv\\Scripts\\python.exe tests\\shriram_working_compare.py --export   # (re)make the reference CSV from the working file (needs the 27 MB file)

Known, accepted differences (see AGENT_CONTEXT.md 7.6): the user's Mahindra rows are separate sheet rows without the 'Old' T&C (it is added when the upload rows are made - the comparison
ignores a leading 'Old'), the tariff chain is made after the rows are built (tariff is not compared), Delhi & NCR Private Car leaves out HR26, tractors with age NEW and no remark are
0-99 with no T&C, Kerala 'UPTO 2501-2805' was typed 2800 by the user."""
import collections
import datetime as dt
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("shriram_broker_grid_to_config.py")
WORKING = r"D:\harsh\grids\OCT26\shriram\Shriram_-_OCT26_BROKER_GRID.xlsx"
REF = os.path.join(helpers.REFS, "shriram_oct26_working_mapping.csv")
GRID = os.path.join(helpers.INPUTS, "shriram_oct26.xlsx")
COLS = ["STATE", "PRODUCTS", "DIS %", "% PO", "POLICY TYPE", "AGE", "UW REMARKS", "REMARKS 1", "RTO", "MAKE", "class", "cc from", "cc to", "SC FROM", "SC TO", "age from", "age to",
        "TARIFF FROM", "TATIFF TO", "fuel", "CONDITION", "TNC", "NCB", "CPA", "ZD"]
OTHERS = set(mod.GCV_CLASSES_OTHER)


def n(v):
    return "" if v is None or (isinstance(v, float) and v != v) else re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def numtxt(v):
    t = n(v)
    try:
        return f"{float(t):g}"
    except ValueError:
        return t


def gkey(state, prod, dis, po, pt, age, uw, rem):
    return "|".join([n(state).lower(), n(prod).lower(), numtxt(dis), numtxt(po).lower(), n(pt).lower(), n(age).lower(), n(uw).lower(), n(rem).lower()])


def export():
    import openpyxl
    wb = openpyxl.load_workbook(WORKING, read_only=True, data_only=True)
    rows = [list(r) for r in wb["GRID OCT' 26 (W)"].iter_rows(values_only=True)]
    wb.close()
    hdr = [n(x) for x in rows[0]]
    keep = [hdr.index(c) for c in COLS]
    out = [[n(r[i]) if i < len(r) else "" for i in keep] for r in rows[1:] if n(r[hdr.index("STATE")]) != ""]
    pd.DataFrame(out, columns=COLS).to_csv(REF, index=False, encoding="utf-8")
    print(f"{len(out)} rows -> {REF}")


def mine_by_key():
    import openpyxl
    wb = openpyxl.load_workbook(GRID, data_only=True)
    ask = mod.Asker(interactive=False)
    grid, warns = mod.read_grid(wb, ask, lambda *a, **k: None)
    d = mod.Declines()
    d.read(wb, warns)
    ctx = mod.Ctx(warns, ("01/10/2026", "31/03/2030"), d, mod.load_masters(ask), False)
    out = collections.defaultdict(list)
    for g in grid:
        for r in mod.convert_row(g, ctx):
            out[gkey(g["state"], g["product"], g["dis"], g["po"], g["ptype"], g["age"], g["uw"], g["rem1"])].append(r)
    return out


def variants(rows):
    """one tuple per (rto, makes, class family, cc, sc, ncb, cpa, zd, tnc) for the package rows (STP only when there are no package rows)"""
    d = {}
    for r in rows:
        cf = "TANKER" if r["make_model_class"] == "Tanker" else "EX TANKER" if r["make_model_class"] in OTHERS else r["make_model_class"]
        k = (r["new_rto_list"], r["new_vehicle_makes"].lower(), cf, float(r["cc_min"]), float(r["cc_max"]), float(r["sc_min"]), float(r["sc_max"]), r["is_ncb"], r["is_cpa"],
             r["is_zd"], r["add_tnc"], r["_sub"] == "stp")
        d.setdefault(k, set()).add((float(r["vehicle_age_min"]), float(r["vehicle_age_max"])))
    return d


def fl(v):
    try:
        return round(float(n(v)), 2)
    except ValueError:
        return None


def main(show=3):
    if "--export" in sys.argv:
        return export()
    ref = pd.read_csv(REF, dtype=str, keep_default_na=False)
    mine = mine_by_key()
    tally, examples, compared, equal = collections.Counter(), collections.defaultdict(list), 0, 0
    for _, r in ref.iterrows():
        key = gkey(r["STATE"], r["PRODUCTS"], r["DIS %"], r["% PO"], r["POLICY TYPE"], r["AGE"], r["UW REMARKS"], r["REMARKS 1"])
        if key not in mine:
            continue
        compared += 1
        V = variants(mine[key])
        pk = {k: a for k, a in V.items() if not k[-1]} or V
        u = dict(rto=r["RTO"].lower(), make=r["MAKE"].lower(), cls=r["class"], cc=(fl(r["cc from"]) or 0, fl(r["cc to"]) or 0), sc=(fl(r["SC FROM"]) or 0, fl(r["SC TO"]) or 0),
                 af=fl(r["age from"]) or 0, at=[float(x) for x in re.findall(r"\d+(?:\.\d+)?", r["age to"])], tnc="" if r["TNC"] in ("0", "nan") else r["TNC"].lower(),
                 ncb=(r["NCB"] or "NA").upper(), cpa=(r["CPA"] or "NA").upper(), zd=(r["ZD"] or "NA").upper())
        best, score = None, -1
        for k, a in pk.items():
            eq = {"rto": k[0].lower() == u["rto"], "make": k[1] == u["make"], "class": k[2].lower() == u["cls"].lower() or (k[2] in ("Passenger Carrying Rickshaw", "Goods Carrying Rickshaw") and k[2] in u["cls"]),
                  "cc": abs(k[3] - u["cc"][0]) < 1.5 and abs(k[4] - u["cc"][1]) < 1.5, "sc": abs(k[5] - u["sc"][0]) < 0.02 and abs(k[6] - u["sc"][1]) < 0.03,
                  "age": abs(min(x[0] for x in a) - u["af"]) < 1.02 and bool(u["at"]) and abs(max(x[1] for x in a) - u["at"][0]) < 0.02,
                  "tnc": re.sub(r"^old,?", "", k[10].lower()) == re.sub(r"^old,?", "", u["tnc"]), "ncb": k[7] == u["ncb"], "cpa": k[8] == u["cpa"], "zd": k[9] == u["zd"]}
            if sum(eq.values()) > score:
                best, score = (k, eq), sum(eq.values())
        k, eq = best
        equal += all(eq.values())
        for f, ok in eq.items():
            if not ok:
                tally[f] += 1
                if len(examples[f]) < show:
                    examples[f].append((r["STATE"], r["PRODUCTS"][:28], r["UW REMARKS"][:40], k, {x: r[x] for x in ("RTO", "MAKE", "class", "TNC", "NCB", "CPA", "ZD")}))
    print(f"grid rows of the working sheet compared: {compared}; equal on rto / make / class / cc / sc / age / tnc / ncb / cpa / zd: {equal} ({equal / compared:.0%})")
    print("fields that differ (rows):", dict(tally))
    for f, ex in examples.items():
        for e in ex:
            print(f"  [{f}] {e[0]} | {e[1]} | {e[2]} | mine {e[3][:4]} ncb/cpa/zd {e[3][7:10]} tnc {e[3][10]!r} | user {e[4]}")
    return compared, equal


if __name__ == "__main__":
    main()
