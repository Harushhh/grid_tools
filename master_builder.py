"""master_builder.py - phase 2 of the RTO / make master pipeline: read the plain-text warnings an upload CSV carries and write the Excel files that add the missing groups to the masters.

Phase 1 (the converters, RTO_NAMES = "master", Shriram): a scope whose RTO codes are exactly a group of the master gets the master's name; one the master lacks is left in the cell as
    RTOs:   New changes found, new mapping needs to be created: <Sheet Name> - <RTO/RTO/RTO...>
    makes:  New changes found, new mapping needs to be created: <Sheet Name> - <GROUP> = <BASE GROUP> except <GROUP>|<GROUP>      (or only <GROUP> when the script does not know how it is made)
Phase 2 (this script): finds those warnings, one new group per distinct set, and writes a workbook with exactly two columns and one sheet:
  RTO master   rto_name      lowercase [insurer]_[month]_[region/condition], e.g. shriram_oct26_ap_ex_27 - letters, digits and '_' only: no spaces, no commas
               rto_cluster   the RTO codes, comma-separated with no spaces, e.g. AP01,AP02,AP03
  make master  make_model_name     lowercase [insurer]_[month]_[condition], e.g. shriram_oct26_tw_all_ex_tvs - the same rule
               make_model_cluster  the make-model entries, comma-separated with no space around the commas (a model's own name keeps its inner spaces, as in the master: TVS SCOOTY)
Nothing is invented. An RTO name is built from the insurer and month of the file, the state(s) of the codes and the condition (the RTOs left out of the state, or the kind of sheet - metro, zone ...).
A make group is built only when the warning says how it is made (its base group in the make master minus the groups taken out): the entries are the master's own. A make warning with no
definition is listed, not built. A set the master already has under some name is not added again. A name already in the master, or used twice, gets a _2 / _3 suffix. An output file that exists is never
overwritten.

    python master_builder.py --csv output\\shriram\\SHRIRAM_UPLOAD_2026-10.csv
    python master_builder.py --csv a.csv --csv b.csv --insurer shriram --month oct26 --out new_rto_groups.xlsx --make-out new_make_groups.xlsx
"""
import argparse
import glob
import os
import re
import sys
from collections import Counter

NEW_MAPPING = "New changes found, new mapping needs to be created: "           # the phrase the converters write (shared block: NEW_MAPPING)
CODE = re.compile(r"[A-Z]{2}\d{1,3}")
CODE_LIST = re.compile(r"[A-Z]{2}\d{1,3}(?:/[A-Z]{2}\d{1,3})*")
NAME_OK = re.compile(r"[a-z0-9_]+")
CLUSTER_OK = re.compile(r"[A-Z]{2}\d{1,3}(,[A-Z]{2}\d{1,3})*")
GROUPING_WORDS = ("metro", "zone", "city", "cities", "ncr", "hill", "coastal", "tier", "cluster", "region")        # a sheet that names a kind of area: it names the condition
HERE = os.path.dirname(os.path.abspath(__file__))
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


class BuilderError(Exception):
    pass


# ------------------------------------------------------------------ reading the warnings
def split_cell(cell):
    """([names the cell holds], [warning dicts {sheet, items, name, base, excluded}]) - the same grammar as the converters' shared block."""
    head, *tails = str(cell).split(NEW_MAPPING)
    names = [t.strip() for t in head.split(",") if t.strip()]
    warns = []
    for tail in tails:
        tail = tail.strip().rstrip(",")
        if " = " in tail:
            left, _, right = tail.partition(" = ")
            sheet, _, name = left.rpartition(" - ")
            base, _, ex = right.partition(" except ")
            warns.append(dict(sheet=sheet.strip(), items=[], name=name.strip(), base=base.strip(), excluded=[x.strip() for x in ex.split("|") if x.strip()]))
        else:
            sheet, _, rest = tail.rpartition(" - ")
            if CODE_LIST.fullmatch(rest.strip()):
                warns.append(dict(sheet=sheet.strip(), items=[i.strip() for i in rest.split("/") if i.strip()], name="", base="", excluded=[]))
            else:
                warns.append(dict(sheet=sheet.strip(), items=[], name=rest.strip(), base="", excluded=[]))
    return names, warns


def read_csv_rows(path):
    import csv
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def collect(csv_paths):
    """-> (rto_sets, make_groups, facts).  rto_sets = {frozenset(codes): {"sheets": set, "rows": int}};  make_groups = {name: {"base", "excluded", "sheets", "rows"}}."""
    rto_sets, makes, insurers, dates = {}, {}, Counter(), Counter()
    for path in csv_paths:
        for r in read_csv_rows(path):
            insurers[(r.get("insurance_company") or "").strip()] += 1
            dates[(r.get("from_date") or "").strip()] += 1
            for w in split_cell(r.get("new_rto_list", ""))[1]:
                codes = frozenset(i for i in w["items"] if CODE.fullmatch(i))
                if codes:
                    d = rto_sets.setdefault(codes, {"sheets": set(), "rows": 0})
                    d["sheets"].add(w["sheet"])
                    d["rows"] += 1
            for w in split_cell(r.get("new_vehicle_makes", ""))[1]:
                if w["name"]:
                    d = makes.setdefault(w["name"], {"base": w["base"], "excluded": w["excluded"], "sheets": set(), "rows": 0})
                    d["sheets"].add(w["sheet"])
                    d["rows"] += 1
    return rto_sets, makes, {"insurers": insurers, "dates": dates}


# ------------------------------------------------------------------ the masters
def read_master(path):
    """{group name: [codes / entries]} of a portal master export; the name and list columns are found by their headers"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    head = next((i for i, r in enumerate(rows[:6]) if any(isinstance(c, str) and "name" in c.lower() for c in r) and any(isinstance(c, str) and ("cluster" in c.lower() or "codes" in c.lower() or "entries" in c.lower()) for c in r)), 0)
    name_col = next(j for j, c in enumerate(rows[head]) if isinstance(c, str) and "name" in c.lower())
    list_col = next(j for j, c in enumerate(rows[head]) if isinstance(c, str) and ("cluster" in c.lower() or "codes" in c.lower() or "entries" in c.lower()))
    return {str(r[name_col]).strip(): [x.strip() for x in str(r[list_col] or "").split(",") if x.strip()] for r in rows[head + 1:] if r and len(r) > max(name_col, list_col) and r[name_col] is not None}


def find_master(folder, words):
    hits = [p for p in glob.glob(os.path.join(folder, "*.xls*")) if not os.path.basename(p).startswith("~$") and all(w in os.path.basename(p).lower() for w in words)]
    return max(hits, key=os.path.getmtime) if hits else None


# ------------------------------------------------------------------ names
def slug(text):
    return re.sub(r"_+", "_", re.sub(r"[^a-z0-9]+", "_", str(text).lower())).strip("_")


def month_word(dates):
    """'oct26' from the most common from_date (DD/MM/YYYY or DD-MM-YYYY)"""
    for d, _ in dates.most_common():
        m = re.fullmatch(r"(\d{2})[/-](\d{2})[/-](\d{4})", d or "")
        if m and 1 <= int(m.group(2)) <= 12:
            return f"{MONTHS[int(m.group(2)) - 1]}{m.group(3)[2:]}"
    raise BuilderError("no from_date found in the file(s): give --month (for example oct26)")


def insurer_word(insurers):
    for name, _ in insurers.most_common():
        w = re.findall(r"[a-z0-9]+", name.lower())
        if w:
            return w[0]
    raise BuilderError("no insurance_company found in the file(s): give --insurer (for example shriram)")


def condition_of(codes, sheets, master):
    """(region, condition) words of a set of RTO codes: the state(s), and either the kind of area the sheet names (metro, zone ...) or the RTOs left out of / the only RTOs of the state."""
    states = sorted({c[:2] for c in codes})
    region = "_".join(s.lower() for s in states) if len(states) <= 3 else f"multi{len(states)}"
    sheet_words = {w for s in sheets for w in re.findall(r"[a-z]+", s.lower())}
    for w in GROUPING_WORDS:
        if w in sheet_words:
            return region, w
    digits = lambda cs: "_".join(re.sub(r"^[A-Z]+", "", c) for c in sorted(cs))
    if len(states) == 1:
        state = master.get(states[0], [])
        if state and set(codes) <= set(state):
            left_out = set(state) - set(codes)
            if 0 < len(left_out) <= 12:
                return region, f"ex_{digits(left_out)}"
        if len(codes) <= 8:
            return region, f"only_{digits(codes)}"
    return region, f"{len(codes)}_rtos"


def unique_name(base, taken):
    name, n = base, 1
    while name in taken:
        n += 1
        name = f"{base}_{n}"
    taken.add(name)
    return name


def build(rto_sets, master, insurer, month):
    """-> ([(name, codes)], [notes]): one new RTO group per set the master does not have yet"""
    by_content = {frozenset(c): n for n, c in master.items()}
    taken = {n.lower() for n in master}
    rows, notes = [], []
    for codes in sorted(rto_sets, key=lambda s: (sorted(s)[0], len(s), sorted(s))):
        if codes in by_content:
            notes.append(f"{len(codes)} RTOs ({sorted(codes)[0]} ...) are already the master group '{by_content[codes]}': not added")
            continue
        region, cond = condition_of(codes, rto_sets[codes]["sheets"], master)
        rows.append((unique_name(slug(f"{insurer}_{month}_{region}_{cond}"), taken), ",".join(sorted(codes))))
    return rows, notes


def build_makes(make_groups, make_master, insurer, month):
    """-> ([(name, entries)], [notes]): one new make group per warning that says how the group is made (its base group in the make master minus the groups taken out)"""
    low = {n.lower(): n for n in make_master}
    by_content = {frozenset(e.strip() for e in v): n for n, v in make_master.items()}
    taken = {n.lower() for n in make_master}
    rows, notes = [], []
    for name in sorted(make_groups):
        g = make_groups[name]
        if not g["base"]:
            notes.append(f"make group '{name}' ({g['rows']} rows): the converter does not say how it is made (no base group): not built - give it a definition (MAKE_DEFINITIONS)")
            continue
        base = low.get(g["base"].lower())
        out = [low.get(x.lower()) for x in g["excluded"]]
        if base is None or None in out:
            notes.append(f"make group '{name}': its base group '{g['base']}' or a group taken out of it is not in the make master: not built")
            continue
        gone = {e.strip() for x in out for e in make_master[x]}
        entries = list(dict.fromkeys(e.strip() for e in make_master[base] if e.strip() and e.strip() not in gone))
        if not entries:
            notes.append(f"make group '{name}' would be empty: not built")
            continue
        if frozenset(entries) in by_content:
            notes.append(f"make group '{name}' is already the master group '{by_content[frozenset(entries)]}': not added")
            continue
        words = [w for w in slug(name).split("_") if w and w != insurer and w != month]
        rows.append((unique_name(slug(f"{insurer}_{month}_{'_'.join(words)}"), taken), ",".join(entries)))
    return rows, notes


# ------------------------------------------------------------------ the output and its strict rules
def check_rows(rows, kind="rto"):
    """the strict rules, enforced on the output itself"""
    seen = set()
    for name, cluster in rows:
        if not NAME_OK.fullmatch(name) or name != name.lower():
            raise BuilderError(f"the name '{name}' is not lowercase letters, digits and '_' only")
        if name in seen:
            raise BuilderError(f"the name '{name}' is not distinct")
        seen.add(name)
        if kind == "rto":
            if not CLUSTER_OK.fullmatch(cluster):
                raise BuilderError(f"the cluster of '{name}' is not comma-separated RTO codes without spaces")
        else:
            entries = cluster.split(",")
            if not cluster or any(e == "" or e != e.strip() for e in entries):
                raise BuilderError(f"the cluster of '{name}' has an empty entry or a space around a comma")


def write_xlsx(rows, path, columns=("rto_name", "rto_cluster"), kind="rto"):
    import openpyxl
    check_rows(rows, kind)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "rto_master_additions" if kind == "rto" else "make_master_additions"
    ws.append(list(columns))
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def free_path(path):
    """never overwrite a file the person may have uploaded already"""
    base, ext = os.path.splitext(path)
    n, out = 1, path
    while os.path.exists(out):
        n += 1
        out = f"{base}_{n}{ext}"
    return out


def default_out(insurer, month, what):
    stamp = f"20{month[-2:]}-{MONTHS.index(month[:3]) + 1:02d}" if re.fullmatch(r"[a-z]{3}\d{2}", month) else month
    return os.path.join(HERE, "output", insurer, f"{insurer.upper()}_MASTER_{what}_ADD_{stamp}.xlsx")


# ------------------------------------------------------------------ main
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", action="append", required=True, help="an upload CSV carrying 'New changes found ...' warnings (repeat for several files)")
    ap.add_argument("--master-dir", default=os.path.join(HERE, "master"), help="the folder with the exported master workbooks (default: master\\)")
    ap.add_argument("--insurer", help="the insurer word of the names (default: the first word of insurance_company)")
    ap.add_argument("--month", help="the month word of the names, e.g. oct26 (default: from the from_date of the file)")
    ap.add_argument("--out", help="the RTO Excel file to write (default: output\\<insurer>\\<INSURER>_MASTER_RTO_ADD_<yyyy-mm>.xlsx)")
    ap.add_argument("--make-out", help="the make / model Excel file to write (default: output\\<insurer>\\<INSURER>_MASTER_MAKE_ADD_<yyyy-mm>.xlsx)")
    a = ap.parse_args(argv)
    for p in a.csv:
        if not os.path.exists(p):
            raise BuilderError(f"{p} does not exist")
    rto_sets, make_groups, facts = collect(a.csv)
    insurer = slug(a.insurer) if a.insurer else insurer_word(facts["insurers"])
    month = slug(a.month) if a.month else month_word(facts["dates"])
    print(f"{len(rto_sets)} distinct set(s) of RTO codes and {len(make_groups)} make group(s) need a new mapping")
    # RTO groups
    rto_path = find_master(a.master_dir, ("rto", "master"))
    master = read_master(rto_path) if rto_path else {}
    if not rto_path and rto_sets:
        print(f"WARNING: no RTO master workbook in {a.master_dir}: names cannot be checked against it and 'ex_' conditions cannot be worked out")
    rows, notes = build(rto_sets, master, insurer, month)
    for n in notes:
        print("  note:", n)
    if rows:
        out = free_path(a.out or default_out(insurer, month, "RTO"))
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        write_xlsx(rows, out)
        print(f"Wrote {len(rows)} new RTO group(s) to {out}")
        for name, cluster in rows:
            print(f"  {name}  ({cluster.count(',') + 1} RTOs: {cluster[:60]}{'...' if len(cluster) > 60 else ''})")
    elif rto_sets:
        print("No new RTO group to add.")
    # make / model groups
    if make_groups:
        mm_path = find_master(a.master_dir, ("make", "model", "master"))
        make_master = read_master(mm_path) if mm_path else {}
        if not mm_path:
            print(f"WARNING: no make / model master workbook in {a.master_dir}: the groups cannot be built")
        mrows, mnotes = build_makes(make_groups, make_master, insurer, month)
        for n in mnotes:
            print("  note:", n)
        if mrows:
            mout = free_path(a.make_out or default_out(insurer, month, "MAKE"))
            os.makedirs(os.path.dirname(os.path.abspath(mout)), exist_ok=True)
            write_xlsx(mrows, mout, ("make_model_name", "make_model_cluster"), "make")
            print(f"Wrote {len(mrows)} new make / model group(s) to {mout}")
            for name, cluster in mrows:
                print(f"  {name}  ({cluster.count(',') + 1} entries: {cluster[:60]}{'...' if len(cluster) > 60 else ''})")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BuilderError as e:
        print(f"STOPPED - {e}")
        sys.exit(2)
