"""Compare the Shriram two-wheeler script's per-grid-row mapping with the mapping the user typed beside the rows of their September 2026 working file.

    python tests/shriram_2w_working_compare.py [--working <working .xlsx>] [--broker-grid <broker grid .xlsx>] [--show 40]

The working file ("shriram_-_2W_NEWROLLOVER_SEPT_26.xlsx") keeps Shriram's two rate sheets and adds, to the right of each row, the user's own mapping:
  New Business - 2W : rto | make | cc from | cc to | CLASS                          (columns after the CHANGE DATE column)
  ROLLOVER 2W (1+1) : RTO | MAKE | PRODUCT (class) | AGE FROM | AGE TO | CC FROM | CC TO | POLICY TYPE
For every grid row the script's rows (convert_nb / convert_ro) are collected and the fields compared; a row the user left unmapped (blank RTO) is not counted.
It is a development tool, not part of QUICK: the September grid carries BLOCKED rows and payouts that the October script does not write, and the user's mapping
is a draft (it names groups the master of 25 Oct does not have).  A printed disagreement is a question, not necessarily a bug.
"""
import argparse
import importlib.util
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
spec = importlib.util.spec_from_file_location("shriram_2w", os.path.join(ROOT, "shriram_2w_to_config.py"))
S = importlib.util.module_from_spec(spec)
spec.loader.exec_module(S)

WORKING = r"D:\harsh\grids\SEPT 26\Shriram\shriram_-_2W_NEWROLLOVER_SEPT_26.xlsx"
BROKER = r"C:\Users\Harsj\Downloads\Shriram_-_OCT26_BROKER_GRID (1).xlsx"


def text(v):
    return "" if v is None else str(v).replace("\n", " ").strip()


def names(token):
    return sorted(t.strip().lower() for t in text(token).split(",") if t.strip())


def codes(masters, token):
    out = set()
    for t in names(token):
        found = {k: v for k, v in masters.rto.items() if k.lower() == t}
        for v in found.values():
            out |= set(v)
    return out


def user_mapping(rows, kind):
    """{grid row number: dict of the user's mapping} read from the columns right of the grid"""
    hdr = rows[0] if kind == "nb" else rows[0]
    out = {}
    if kind == "nb":
        col = {text(h).lower(): i for i, h in enumerate(hdr) if text(h)}
        for n, r in enumerate(rows[1:], start=2):
            out[n] = dict(rto=text(r[col["rto"]]), make=text(r[col["make"]]), cc=(text(r[col["cc from"]]), text(r[col["cc to"]])), cls=text(r[col["class"]]))
    else:
        first = next(i for i, h in enumerate(hdr) if text(h).upper() == "RTO")
        names_ = [text(h).upper() for h in hdr[first:]]
        idx = {n: first + i for i, n in enumerate(names_)}
        for n, r in enumerate(rows[2:], start=3):
            out[n] = dict(rto=text(r[idx["RTO"]]), make=text(r[idx["MAKE"]]), cls=text(r[idx["PRODUCT"]]), age=(text(r[idx["AGE FROM"]]), text(r[idx["AGE TO"]])),
                          cc=(text(r[idx["CC FROM"]]), text(r[idx["CC TO"]])), sub=text(r[first + len(names_) - 1]))
    return out


def num(x):
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--working", default=WORKING)
    ap.add_argument("--broker-grid", default=BROKER)
    ap.add_argument("--show", type=int, default=30)
    a = ap.parse_args()
    ask = S.Asker(interactive=False)
    wb = S.read_workbook(a.working)
    masters = S.load_masters(ask)
    warnings = []
    ctx = S.Ctx(warnings, ("01/10/2026", "31/03/2030"), S.read_declines(a.broker_grid if os.path.isfile(a.broker_grid) else None, ask, warnings), masters)
    nb, _ = S.read_new_business(wb, ask, lambda *x: None)
    ro, _ = S.read_rollover(wb, ask, lambda *x: None)
    mine = {}
    for g in nb:
        mine[("nb", g["r"])] = S.convert_nb(g, ctx)
    for g in ro:
        mine[("ro", g["r"])] = S.convert_ro(g, ctx)
    theirs = {("nb", n): v for n, v in user_mapping(S.sheet_rows(wb, "New Business - 2W"), "nb").items()}
    theirs.update({("ro", n): v for n, v in user_mapping(S.sheet_rows(wb, "ROLLOVER 2W (1+1)"), "ro").items()})
    tally = {}
    shown = 0
    for key in sorted(mine):
        t = theirs.get(key)
        if not t or not t["rto"]:
            tally["not mapped by the user"] = tally.get("not mapped by the user", 0) + 1
            continue
        rows = mine[key]
        if not rows:
            tally["no row from the script (skipped / blocked)"] = tally.get("no row from the script (skipped / blocked)", 0) + 1
            continue
        r0 = rows[0]
        diffs = []
        mine_rto = set()
        for tok in r0["new_rto_list"].split(","):
            mine_rto |= masters.codes(tok.strip())
        user_rto = set()
        for tok in t["rto"].split(","):
            user_rto |= masters.codes(tok.strip())
        known = []
        mine_tokens = [x for x in names(r0["new_rto_list"]) if x != "tg"]           # the script writes Telangana as TG,TS; the working file as TS
        if mine_rto != user_rto and mine_tokens != names(t["rto"]):
            diffs.append(f"rto {r0['new_rto_list']} vs {t['rto']}")
        elif names(r0["new_rto_list"]) != names(t["rto"]) and mine_tokens == names(t["rto"]):
            known.append("Telangana written as TG,TS")
        if t["make"] == "#N/A":
            known.append("make not typed in the working file (#N/A)")
        elif names(r0["new_vehicle_makes"]) == ["tvs_ex_apache"] and names(t["make"]) == ["tvs"]:
            known.append("TVS_EX_APACHE (the decline sheet) instead of TVS")
        elif names(r0["new_vehicle_makes"]) != names(t["make"]):
            diffs.append(f"make {r0['new_vehicle_makes']} vs {t['make']}")
        classes = sorted({r["make_model_class"] for r in rows})
        if classes != sorted(x.strip() for x in t["cls"].split(",") if x.strip()):
            diffs.append(f"class {classes} vs {t['cls']}")
        if key[0] == "ro":
            if (num(r0["vehicle_age_min"]), num(r0["vehicle_age_max"])) != (num(t["age"][0]), num(t["age"][1])):
                diffs.append(f"age {(r0['vehicle_age_min'], r0['vehicle_age_max'])} vs {t['age']}")
            want = "STP" if r0["sub_product"] == "STP" else "1+1"
            if want not in t["sub"]:
                diffs.append(f"policy {want} vs {t['sub']}")
        if (num(r0["cc_min"]), num(r0["cc_max"])) != (num(t["cc"][0]), num(t["cc"][1])):
            diffs.append(f"cc {(r0['cc_min'], r0['cc_max'])} vs {t['cc']}")
        tally["equal" if not diffs else "differ"] = tally.get("equal" if not diffs else "differ", 0) + 1
        for k in known:
            tally["  known difference: " + k] = tally.get("  known difference: " + k, 0) + 1
        for d in diffs:
            f = d.split()[0]
            tally["  field " + f] = tally.get("  field " + f, 0) + 1
        if diffs and shown < a.show:
            shown += 1
            print(f"{key[0]} row {key[1]}: " + "; ".join(diffs))
    print()
    for k, v in sorted(tally.items()):
        print(f"{v:>5}  {k}")


if __name__ == "__main__":
    main()
