"""Look at an insurer grid without running the converter: show its layout, remember it, and spot when it changed.

    python grid_doctor.py dump  <grid>                  compact view of every sheet: size, header cells, odd text values
    python grid_doctor.py check <grid> [--against NAME]  compare with the closest remembered layout (exit 1 = it changed)
    python grid_doctor.py snapshot <grid> --name NAME    remember this layout in layouts\\NAME.json (do this after a good month)
    python grid_doctor.py list                           remembered layouts

Reads .xlsx/.xlsm (openpyxl) and .xlsb (pyxlsb). Only cell text near the top and the distinct non-numeric words in the
body are kept, never rates, so a snapshot is small and a diff is short enough to read in one glance.
"""
import argparse
import difflib
import glob
import json
import os
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
LAYOUTS = os.path.join(HERE, "layouts")
HEADER_ROWS = 10          # rows from the top that count as "header"
SCAN_ROWS = 400           # body rows scanned for odd text values
MAX_TOKENS = 80
NOTE_MIN = 28             # a text cell this long is a note / guideline, not a label
MAX_NOTES = 60


def col_letter(n):                       # 1 -> A
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def read_sheets(path):
    """Yield (sheet name, rows_kept, total_rows, total_cols); rows_kept = first SCAN_ROWS rows as lists of values."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsb":
        from pyxlsb import open_workbook
        with open_workbook(path) as wb:
            for name in wb.sheets:
                kept, nrows, ncols = [], 0, 0
                with wb.get_sheet(name) as sh:
                    for row in sh.rows():
                        nrows += 1
                        ncols = max(ncols, len(row))
                        if nrows <= SCAN_ROWS:
                            kept.append([c.v for c in row])
                yield name, kept, nrows, ncols
    elif ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            kept, nrows, ncols = [], 0, 0
            for row in ws.iter_rows(values_only=True):
                nrows += 1
                ncols = max(ncols, len(row))
                if nrows <= SCAN_ROWS:
                    kept.append(list(row))
            yield ws.title, kept, nrows, ncols
        wb.close()
    else:
        sys.exit(f"cannot read {path}: expected .xlsx, .xlsm or .xlsb")


def is_text(v):
    return isinstance(v, str) and v.strip() != ""


PLACEHOLDERS = {"NA", "N/A", "IRDA", "BLOCK"}


def looks_like_data(row):
    """A rate row: three or more cells that are numbers or placeholders such as NA / IRDA / Block."""
    n = 0
    for v in row:
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            n += 1
        elif is_text(v) and v.strip().upper() in PLACEHOLDERS:
            n += 1
    return n >= 3


def header_height(rows):
    """Rows above the first rate row (at least 1, at most HEADER_ROWS)."""
    for i, row in enumerate(rows[:HEADER_ROWS]):
        if looks_like_data(row):
            return max(1, i)
    return min(HEADER_ROWS, len(rows))


def layout_of(path):
    sheets = {}
    for name, rows, nrows, ncols in read_sheets(path):
        height = header_height(rows)
        header = {}
        for r, row in enumerate(rows[:height], 1):
            for c, v in enumerate(row, 1):
                if is_text(v):
                    header[f"{col_letter(c)}{r}"] = " ".join(v.split())[:60]
        geos = {" ".join(row[0].split()) for row in rows[height:] if row and is_text(row[0])}
        tokens = set()
        for row in rows[height:]:
            for v in row[1:]:
                if is_text(v):
                    t = " ".join(v.split())
                    if t not in geos and len(t) <= 30:
                        tokens.add(t)
        notes = [" ".join(v.split())[:600] for row in rows[height:] for v in row if is_text(v) and len(v.strip()) >= NOTE_MIN]
        sheets[name] = {"rows": nrows, "cols": ncols, "header": header, "tokens": sorted(tokens)[:MAX_TOKENS], "notes": notes[:MAX_NOTES]}
    return {"file": os.path.basename(path), "sheets": sheets}


def split_coord(coord):
    letters = "".join(ch for ch in coord if ch.isalpha())
    num = 0
    for ch in letters:
        num = num * 26 + ord(ch) - 64
    return int("".join(ch for ch in coord if ch.isdigit())), num      # (row, col)


def load_layouts():
    out = {}
    for p in sorted(glob.glob(os.path.join(LAYOUTS, "*.json"))):
        out[os.path.splitext(os.path.basename(p))[0]] = json.load(open(p, encoding="utf-8"))
    return out


def similarity(new, old):
    """0..1: how many sheet names and header texts two layouts share."""
    shared = set(new["sheets"]) & set(old["sheets"])
    if not shared:
        return 0.0
    names = len(shared) / len(set(new["sheets"]) | set(old["sheets"]))
    heads = []
    for s in shared:
        a, b = set(new["sheets"][s]["header"].values()), set(old["sheets"][s]["header"].values())
        heads.append(len(a & b) / max(1, len(a | b)))
    return names * sum(heads) / len(heads)


def closest(new, layouts):
    best, score = None, 0.0
    for key, lay in layouts.items():
        s = similarity(new, lay)
        if s > score:
            best, score = key, s
    return best, score


def few(items, n=6):
    items = list(items)
    return f"{items[:n]}" + (f" (+{len(items) - n} more)" if len(items) > n else "")


def word_diff(a, b, context=4):
    """'... payout will be [3% -> 4%] lower ...' : only the words that changed, with a little context around them."""
    wa, wb = a.split(), b.split()
    parts = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, wa, wb, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        before = " ".join(wa[max(0, i1 - context):i1])
        after = " ".join(wa[i2:i2 + context])
        old_w, new_w = " ".join(wa[i1:i2]) or "(nothing)", " ".join(wb[j1:j2]) or "(nothing)"
        parts.append(f"...{before} [{old_w} -> {new_w}] {after}...")
    return " ".join(parts)[:420]


def diff_notes(old, new, limit=6):
    """What changed in the long text cells under a table: edited, new and removed notes."""
    gone = [x for x in old if x not in new]
    added = [x for x in new if x not in old]
    out, pairs = [], []
    for x in list(gone):
        close = difflib.get_close_matches(x, added, n=1, cutoff=0.6)
        if close:
            pairs.append((x, close[0]))
            gone.remove(x)
            added.remove(close[0])
    for a, b in pairs[:limit]:
        out.append(f"note edited: {word_diff(a, b)}")
    if len(pairs) > limit:
        out.append(f"... and {len(pairs) - limit} more edited notes")
    for x in added[:limit]:
        out.append(f"new note: {x[:200]}")
    for x in gone[:limit]:
        out.append(f"note removed: {x[:200]}")
    return out


def diff(old, new, limit=12):
    """Return a list of human-readable differences between two layouts (empty = same)."""
    out = []
    gone, added = sorted(set(old["sheets"]) - set(new["sheets"])), sorted(set(new["sheets"]) - set(old["sheets"]))
    if gone:
        out.append(f"sheets no longer there: {gone}")
    if added:
        out.append(f"new sheets: {added}")
    for name in sorted(set(old["sheets"]) & set(new["sheets"])):
        o, n = old["sheets"][name], new["sheets"][name]
        notes = []
        if o["cols"] != n["cols"]:
            notes.append(f"columns {o['cols']} -> {n['cols']}")
        if abs(o["rows"] - n["rows"]) > max(3, 0.02 * o["rows"]):
            notes.append(f"rows {o['rows']} -> {n['rows']}")
        oh, nh = o["header"], n["header"]
        by_text_old, by_text_new = {}, {}
        for k, v in oh.items():
            by_text_old.setdefault(v, []).append(k)
        for k, v in nh.items():
            by_text_new.setdefault(v, []).append(k)
        c_old, c_new = Counter(oh.values()), Counter(nh.values())      # counted, so one of two equal headers being renamed shows up
        removed = list(dict.fromkeys((c_old - c_new).elements()))
        new_t = list(dict.fromkeys((c_new - c_old).elements()))
        moved = [(t, by_text_old[t][0], by_text_new[t][0]) for t in by_text_old
                 if t in by_text_new and len(by_text_old[t]) == 1 and len(by_text_new[t]) == 1
                 and by_text_old[t][0] != by_text_new[t][0]]
        if moved:
            deltas = Counter((split_coord(b)[0] - split_coord(a)[0], split_coord(b)[1] - split_coord(a)[1]) for _, a, b in moved)
            (dr, dc), cnt = deltas.most_common(1)[0]
            if cnt >= 0.8 * len(moved):
                notes.append(f"{len(moved)} header cells moved together by {dr:+d} row(s), {dc:+d} column(s)")
            else:
                notes.append(f"{len(moved)} header cells moved, e.g. " + "; ".join(f"'{t}' {a}->{b}" for t, a, b in moved[:limit // 3]))
        if removed:
            notes.append(f"{len(removed)} header text(s) gone: {few(removed)}")
        if new_t:
            notes.append(f"{len(new_t)} new header text(s): {few(new_t)}")
        if "notes" in o and "notes" in n:
            notes.extend(diff_notes(o["notes"], n["notes"]))
        ot, nt = set(o["tokens"]), set(n["tokens"])
        if nt - ot:
            notes.append(f"new text value(s) in the body: {few(sorted(nt - ot))}")
        if ot - nt:
            notes.append(f"text value(s) no longer in the body: {few(sorted(ot - nt))}")
        if notes:
            out.append(f"sheet '{name}':")
            out.extend("    - " + x for x in notes)
    return out


def cmd_dump(a):
    lay = layout_of(a.grid)
    print(f"FILE {lay['file']}")
    for name, s in lay["sheets"].items():
        print(f"\nSHEET '{name}'  {s['rows']} rows x {s['cols']} cols")
        by_row = {}
        for k, v in s["header"].items():
            r, c = split_coord(k)
            by_row.setdefault(r, []).append((c, f"{k}={v}"))
        for r in sorted(by_row):
            cells = [t for _, t in sorted(by_row[r])]
            shown = " | ".join(cells[:14]) + (f" | ... (+{len(cells) - 14} more)" if len(cells) > 14 else "")
            print(f"  row {r}: {shown}")
        if s["tokens"]:
            print(f"  body text values: {s['tokens']}")
        for note in s.get("notes", [])[:12]:
            print(f"  note: {note[:300]}")


def cmd_snapshot(a):
    lay = layout_of(a.grid)
    os.makedirs(LAYOUTS, exist_ok=True)
    path = os.path.join(LAYOUTS, a.name + ".json")
    json.dump(lay, open(path, "w", encoding="utf-8"), indent=1)
    print(f"saved {path}  ({len(lay['sheets'])} sheets)")


def cmd_check(a):
    new = layout_of(a.grid)
    layouts = load_layouts()
    if not layouts:
        sys.exit("no remembered layouts yet - run 'snapshot' on a grid that worked")
    key, score = (a.against, 1.0) if a.against else closest(new, layouts)
    if key not in layouts:
        sys.exit(f"no remembered layout named {key}; have: {sorted(layouts)}")
    if score < 0.3:
        print(f"No remembered layout looks like this file (best match '{key}', similarity {score:.0%}). Treat it as a new insurer or a new format.")
        sys.exit(1)
    changes = diff(layouts[key], new)
    print(f"{new['file']}  vs remembered layout '{key}' ({layouts[key]['file']})")
    if not changes:
        print("Same layout: sheets, header cells, sizes and text values all match.")
        return
    print("The layout changed:")
    for c in changes:
        print("  " + c)
    sys.exit(1)


def cmd_list(_):
    for k, v in load_layouts().items():
        print(f"{k:<28} from {v['file']}  ({len(v['sheets'])} sheets)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in (("dump", cmd_dump), ("check", cmd_check), ("snapshot", cmd_snapshot)):
        p = sub.add_parser(name)
        p.add_argument("grid")
        if name == "check":
            p.add_argument("--against")
        if name == "snapshot":
            p.add_argument("--name", required=True)
        p.set_defaults(fn=fn)
    sub.add_parser("list").set_defaults(fn=cmd_list)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
