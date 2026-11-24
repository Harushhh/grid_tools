"""The person's RTO working list (the user, 31 Oct 2026, Shriram and Go Digit only): a column headed RTO next to the rate table is written into new_rto_list EXACTLY as typed - no master
lookup, no condition text, no check against the master; a row without a value keeps the old lookup; a grid without the column behaves as before.
Run:  .venv\\Scripts\\python.exe tests\\working_list_tests.py"""
import datetime as dt
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


# ================================================================ the shared helpers (the same text in all eight scripts)
SCRIPTS = ["liberty_grid_to_config.py", "magma_grid_to_config.py", "rsa_gcv4w_comp_to_config.py", "rsa_pvtcar_to_config.py", "united_commission_to_config.py",
           "shriram_broker_grid_to_config.py", "shriram_2w_to_config.py", "godigit_agency_to_config.py", "sbi_grid_to_config.py"]
for s in SCRIPTS:
    m = helpers.load_script(s)
    ok = (m.working_value("  West UP Open ") == "West UP Open" and m.working_value(12.0) == "12" and m.working_value("MH01, MH02 ,MH03") == "MH01,MH02,MH03"
          and m.working_value("West\nUP") == "West UP" and all(m.working_value(x) == "" for x in (None, "", "   ", "NA", "n/a", "-", "--", "Nil", float("nan"))))
    check(f"{s}: working_value keeps the text as typed, drops spaces around commas, and a blank / NA / '-' / nan is no record", ok)
    m.working_reset()
    errs = [m.working_value(x) for x in ("#REF!", " #n/a ", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!")]
    check(f"{s}: an Excel error in the cell (#REF!, #N/A, #VALUE! ...) is no record - counted for the run line (the 52 `=#REF!` cells of the November 2026 grid were written as the RTO list `#REF!`) - and a name that merely starts with # is kept",
          errs == [""] * 7 and sorted(m.WORKING_ERRORS) == sorted(["#REF!", "#N/A", "#VALUE!", "#NAME?", "#DIV/0!", "#NULL!", "#NUM!"]) and m.working_value("#12 Cluster") == "#12 Cluster", (errs, m.WORKING_ERRORS))
    m.working_reset()
    check(f"{s}: a new run forgets the errors of the last one", m.WORKING_ERRORS == [])
    check(f"{s}: working_header accepts RTO, 'RTO working list', 'RTO list', 'RTO name' (any case / spacing) and not RTO Cluster, RTO Code, RTO State",
          all(m.working_header(x) for x in ("RTO", " rto ", "RTO Working List", "RTO_LIST", "RTO name", "Working RTO")) and not any(m.working_header(x) for x in ("RTO Cluster", "RTO Code", "RTO STATE", "State", None, "")))

# ================================================================ Shriram broker grid
B = helpers.load_script("shriram_broker_grid_to_config.py")
HDR = ["STATE", "PRODUCTS", "DIS %", "% PO", "POLICY TYPE", "AGE", "UW REMARKS", "REMARKS 1", "EFFECTIVE DATE"]
GOA = ["GOA", "PRIVATE CAR PETROL", 85, "16 Net", "P & SA-OD", "upto 15 years.", "NCB Cases"]
AP = ["ANDHRA PRADESH", "PRIVATE CAR PETROL", 85, "16 Net", "P", "upto 15 years."]


def bgrid(rows, rto=None, header="RTO", at=1, extra=None):
    """the grid sheet; `rto` = one value per row, in a column headed `header` inserted at index `at`"""
    head = list(HDR)
    body = [list(r) + [None] * (9 - len(r)) for r in rows]
    if rto is not None:
        head.insert(at, header)
        for r, v in zip(body, rto):
            r.insert(at, v)
    sheets = {"GRID OCT' 26": [head] + body}
    sheets.update(extra or {})
    return helpers.FakeWB(sheets)


def brun(wb, names="group", write_ev=None):
    B.RTO_NAMES = names
    try:
        return B.build(None, F, T, quiet, B.Asker(interactive=False), wb=wb, write_ev=write_ev)
    finally:
        B.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "master")


def note_of(w):
    return next((x for x in w if x.startswith("RTO working list")), "")


for mode in ("group", "master", "cluster"):
    df, w = brun(bgrid([GOA, AP], ["  West Goa Open ", None]), mode)
    goa = df[df.sub_product.isin(["1+1", "SAOD"])]
    check(f"broker [{mode}]: a row with a name gets it exactly as typed (outer spaces trimmed), on every row it makes", set(df[df.new_rto_list == "West Goa Open"].sub_product) == {"1+1", "SAOD"} and len(df[df.new_rto_list == "West Goa Open"]) == 6, df.new_rto_list.value_counts())
    check(f"broker [{mode}]: nothing is flagged or rewritten for a working-list name (no 'New changes found', no 'not in the master')",
          not df[df.new_rto_list == "West Goa Open"].remarks.str.contains("mapping|master|Manual Review", case=False).any() and not df.new_rto_list.str.contains("New changes found").any(), df.remarks.unique()[:3])
    check(f"broker [{mode}]: a blank cell keeps today's lookup (Andhra Pradesh = AP)", set(df[df.new_rto_list != "West Goa Open"].new_rto_list) == {"AP"}, df.new_rto_list.unique())
n = note_of(w)
check("broker: one run line - rows that carry a name, the Excel row of the blank one", "1 grid row(s) carry a name in column 'RTO'" in n and "Excel row 3" in n and "master lookup" in n, n)

base, wb0 = brun(bgrid([GOA, AP]))
check("broker: a grid without the column behaves as before - no run line, the master lookup", note_of(wb0) == "" and set(base.new_rto_list) == {"GA", "AP"}, (note_of(wb0), base.new_rto_list.unique()))
df, _ = brun(bgrid([GOA, AP], [None, None]))
check("broker: the column present but empty = the lookup for every row (same rows as without the column)", df.drop(columns=["remarks"]).equals(base.drop(columns=["remarks"])))
df, w = brun(bgrid([GOA, AP, AP], ["NA", "-", "n/a"]))
check("broker: NA / '-' / n/a say 'no record': the lookup is used", set(df.new_rto_list) == {"GA", "AP"} and "3 row(s) have no value" in note_of(w), note_of(w))
df, w = brun(bgrid([["ATLANTIS", "PRIVATE CAR PETROL", 85, "16 Net", "P", "upto 15 years."], ["ATLANTIS", "PRIVATE CAR PETROL", 80, "15 Net", "P", "upto 15 years."]], ["Lost City", None]))
check("broker: a state the script does not know is still written when the row has a name, and skipped (as before) when it has none", set(df.new_rto_list) == {"Lost City"} and len(df) == 3 and any("state 'ATLANTIS'" in x for x in w), (len(df), w))
df, w = brun(bgrid([AP, AP[:2] + [90] + AP[3:]], ["Cluster One", "Cluster Two"]))
check("broker: two grid rows with two names keep their own", set(df.new_rto_list) == {"Cluster One", "Cluster Two"})
df, w = brun(bgrid([AP], ["MH01, MH02"]))
check("broker: a list in the cell is the list MH01,MH02 (spaces around the comma are formatting)", set(df.new_rto_list) == {"MH01,MH02"})
df, w = brun(bgrid([GOA], ["Odd Name (not in any master) 7"]), "master")
check("broker [master]: a name that is in no master stays as typed - it does not become the 'New changes found' warning", set(df.new_rto_list) == {"Odd Name (not in any master) 7"} and not df.new_rto_list.str.contains("New changes").any())
for head in ("RTO WORKING LIST", "rto list", "RTO Name", "Working RTO"):
    df, w = brun(bgrid([AP], ["Named By Header"], header=head))
    check(f"broker: the header '{head}' names the column", set(df.new_rto_list) == {"Named By Header"} and "carry a name in column" in note_of(w), note_of(w))
for at in (0, 4, 9):
    df, w = brun(bgrid([AP], ["Moved Column"], at=at))
    check(f"broker: the column can be anywhere in the table (index {at})", set(df.new_rto_list) == {"Moved Column"})
for head in ("RTO Cluster", "RTO Code", "RTO STATE"):
    df, w = brun(bgrid([AP], ["Not Mine"], header=head))
    check(f"broker: a column headed '{head}' is not the working list", set(df.new_rto_list) == {"AP"} and note_of(w) == "")
# the addendum is not read for it
wb_ad = bgrid([AP], ["Addendum Name"])
rows_ad, _ = B.read_grid(wb_ad, B.Asker(interactive=False), quiet, working=False)
check("broker: the addendum file is not read for the column (its own RTO / MAKE mapping columns stay ignored)", rows_ad and all(g["rto"] == "" for g in rows_ad))
# a new run forgets the last one
brun(bgrid([GOA], ["Gone After This Run"]))
check("broker: a run remembers its names (so nothing rewrites or checks them) ...", "Gone After This Run" in B.WORKING_RTO and len(B.WORKING_SEEN) == 1)
brun(bgrid([GOA]))
check("broker: ... and forgets them when the next run starts; a grid without the column leaves nothing behind", B.WORKING_RTO == set() and all(not e["header"] for e in B.WORKING_SEEN))
# buses
SCHOOL = ["PAN INDIA", "PCCV 4W SCHOOL BUS", 99, "SC>15 PO 50% & <15 30%", "P & L", "upto 15 years"]
CORP = ["PAN INDIA", "PCCV 4W CORPORATE BUS", 90, "Metro - 45% SC >= 15, 40% SC < 15:Non Metro - 40% SC >= 15, 35% SC < 15", "P & L", "upto 15 years"]
df, w = brun(bgrid([SCHOOL], ["School Cluster"]))
check("broker: a school bus row (one RTO scope) takes the name", set(df.new_rto_list) == {"School Cluster"})
df, w = brun(bgrid([CORP], ["Corp Cluster"]))
check("broker: a corporate bus row has two scopes (metro / non-metro, different payouts): one name cannot say both - the master groups stay and the run line says so",
      set(df.new_rto_list) == {B.METRO_GROUP, B.NON_METRO_GROUP} and "corporate bus" in note_of(w), (set(df.new_rto_list), note_of(w)))
# a remark about RTO codes does not override the name
df, w = brun(bgrid([["MADHYA PRADESH", "PRIVATE CAR PETROL", 85, "16 Net", "P", "upto 15 years.", "EXL RTO codes 03, 06, 18"]], ["MP My List"]))
check("broker: a remark that names RTO codes does not change a working-list name (the column is used strictly)", set(df.new_rto_list) == {"MP My List"})

# ================================================================ Shriram two-wheeler grid
W = helpers.load_script("shriram_2w_to_config.py")
NB_HDR = ["STATE", "MANUFACTURER", "BODY TYPE", "POLICY TYPE", "CONDITION", "DIS %", "Average Net PO\n(Converted)", "UW'S CONDITION", "CHANGE DATE"]
RO_HDR = ["PRODUCT CATEGORY", "RTO STATE", "BODY TYPE", "MANUFACTURER", "CC", "AGE", "DISCOUNT (%)", "NPO(%) / FIXED", "POLICY TYPE", "CONDITION", "BLOCKED CONDITION"]
NB = lambda state, make, body, npo, dis=50, cond="": [state, make, "BUNDLE", body, cond, dis, npo, "", ""]
RO = lambda state, body, make, npo, ptype, dis="", cc="-", cond="": ["Two Wheeler ( 1 +1)", state, body, make, cc, "6 to 15", dis, npo, ptype, cond, ""]


def wgrid(nb=(), ro=(), nb_rto=None, ro_rto=None, header="RTO"):
    sheets = {}
    if nb:
        h = list(NB_HDR) + ([header] if nb_rto is not None else [])
        sheets["New Business - 2W"] = [h] + [list(r) + ([v] if nb_rto is not None else []) for r, v in zip(nb, nb_rto or [None] * len(nb))]
    if ro:
        h = list(RO_HDR) + ([header] if ro_rto is not None else [])
        sheets["ROLLOVER 2W (1+1)"] = [["GENERAL SLAB FOR BROKER TW (1+1) AGE 6 TO 15 YRS"] + [""] * 10, h] + [list(r) + ([v] if ro_rto is not None else []) for r, v in zip(ro, ro_rto or [None] * len(ro))]
    return helpers.FakeWB(sheets)


def wrun(wb, names="group"):
    W.RTO_NAMES = names
    try:
        return W.build("x.xlsx", F, T, log=quiet, ask=W.Asker(interactive=False), wb=wb)
    finally:
        W.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "master")


nb_rows = [NB("BIHAR", "HERO,HONDA,SUZUKI,BAJAJ", "BIKE", 45, 50), NB("DELHI/NCR", "BAJAJ", "BIKE", 30, 50)]
ro_rows = [RO("BIHAR", "BIKE", "ALL EXCL. TVS", 32.5, "P", 40), RO("BIHAR", "BIKE", "ALL EXCL. TVS", 25, "L"), RO("ATLANTIS", "BIKE", "ALL EXCL. TVS", 20, "P", 40)]
for mode in ("group", "master", "cluster"):
    df, w = wrun(wgrid(nb_rows, ro_rows, ["Bihar NB List", None], ["Bihar RO List", "Bihar RO List", "Lost City"]), mode)
    check(f"2W [{mode}]: New Business and Rollover rows take the names exactly as typed; a blank cell keeps the lookup",
          {"Bihar NB List", "Bihar RO List", "Lost City"} <= set(df.new_rto_list) and len(set(df.new_rto_list)) >= 4 and not df.new_rto_list.str.contains("New changes").any(), df.new_rto_list.unique())
    check(f"2W [{mode}]: a state the script does not know is written when its row has a name", "Lost City" in set(df.new_rto_list))
    check(f"2W [{mode}]: nothing is flagged about the working-list names", not df[df.new_rto_list.isin(["Bihar NB List", "Bihar RO List", "Lost City"])].remarks.str.contains("new_rto_list", case=False).any()
          and not df.new_rto_list.str.contains("New changes").any(), df.remarks.unique()[:3])
n = note_of(w)
check("2W: the run line names both sheets, the counts and the blank row", "'New Business - 2W'" in n and "'ROLLOVER 2W (1+1)'" in n and "1 grid row(s) carry a name" in n and "2 grid row(s) carry a name" not in n.split("|")[0], n)
b2, w2 = wrun(wgrid(nb_rows, ro_rows[:2]))
check("2W: no column = as before (no run line); the Rollover header 'RTO STATE' is the state, not the working list", note_of(w2) == "" and "BR" in set(b2.new_rto_list), (note_of(w2), b2.new_rto_list.unique()))
df, w = wrun(wgrid(nb_rows, ro_rows[:2], None, ["Only RO", None]))
check("2W: a sheet without the column is named in the run line and keeps the lookup", "'New Business - 2W' has no RTO working-list column" in note_of(w) and "Only RO" in set(df.new_rto_list), note_of(w))

# ================================================================ Go Digit
G = helpers.load_script("godigit_agency_to_config.py")
G.MISSING_TABLES = "skip"                                   # the small books below have no two-wheeler / car sheets: the script would ask whether to write the file without them
HEAD =["Cluster segment", "Segment", "Body Type", "RTO Cluster", "State", "Make / Variant", "CD1 (OD-disc)"] + ["New", "1 Yr", "2 Yr", "3 Yr", "4 Yr", "5 Yr", "6-10 Yr", "11+ Yr"] * 2
REMARKS = [["REMARKS FOR SCHOOL BUS"], ["1. Short Term Policies - CD2 reduction of 25% on the above Rate"], ["REMARKS FOR STAFF BUS"], ["Note:"], ["REMARKS FOR HCV"],
           ["A new condition for the HCV rows that nobody has explained, only for hilly regions"]]
hcv = lambda cluster, rto=None: ["CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", cluster, "Gujarat", "All", "New/1yr: 60% | 2-4yr: 85% | 5yr+: 85%/90%", 0.25, 0.25, 0.25, 0.25, 0.25, "29%/26%", "29%/26%", "10%/10%",
                                 0.325, 0.325, 0.325, 0.325, 0.325, 0.41, 0.41, 0.41] + ([rto] if rto is not None else [])


def gbook(rows, with_col=True):
    head = HEAD + (["RTO"] if with_col else [])
    return G.Book([G.Sheet("LATEST CV COMBINED GRID", [[], [], [], head] + rows), G.Sheet("CV Remarks", REMARKS)])


def grun(b):
    G.reset_run_lists()
    return G.build(None, dt.date(2026, 10, 15), T, quiet, G.Asker(interactive=False), book=b)


for mode in ("group", "cluster"):
    G.RTO_NAMES = mode
    df, w = grun(gbook([hcv("BAD GJ", "Gujarat HCV List"), hcv("Mumbai", None)]))
    G.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")
    got = set(df.new_rto_list)
    check(f"Go Digit [{mode}]: the named row carries the name exactly; the blank row keeps the cluster lookup", "Gujarat HCV List" in got and len(got) == 2 and "Gujarat HCV List" not in (got - {"Gujarat HCV List"}), got)
    check(f"Go Digit [{mode}]: the working-list name is not flagged as 'not in the master'", not df[df.new_rto_list == "Gujarat HCV List"].remarks.str.contains("not in the", case=False).any(), df.remarks.unique()[:2])
check("Go Digit: the run line counts the rows that carry a name and lists the blank row", "carry a name in column 'RTO'" in note_of(w) and "1 grid row(s)" in note_of(w) and "Excel row" in note_of(w), note_of(w))
named = df[df.new_rto_list == "Gujarat HCV List"]
check("Go Digit: a note printed in the HCV section of 'CV Remarks' is still flagged on the HCV rows that carry a working-list name (the scope is the product, not the RTO name)",
      len(named) > 0 and named.remarks.str.contains("hilly regions").all(), named.remarks.unique()[:2])
d0, w0 = grun(gbook([hcv("BAD GJ"), hcv("Mumbai")], with_col=False))
check("Go Digit: no column = as before (no run line, the cluster lookup)", note_of(w0) == "" and not any("working" in x.lower() for x in w0))

# the real October workbook: an RTO column in every table
path = os.path.join(helpers.INPUTS, "godigit_oct26.xlsx")
TABLES = [  # (sheet, keys of the header row as the reader finds it, cluster column(s) -> column of the RTO header, tag)
    ("LATEST CV COMBINED GRID", ("clustersegment", "rtocluster"), [(3, 23)], "CV"),
    ("Latest 2W 1+1 & SATP", ("agencypbclusters", "agencypbseg", "cd1"), [(0, 6)], "T11"),
    ("2W 1+5", ("agencypbclusters", "make"), [(3, 9), (10, 16)], "T15"),
    ("Latest 2W SAOD", ("cluster", "segment", "make"), [(1, 11)], "SAOD"),
    ("PVT Car Comp+SAOD", ("cluster", "volumeslab"), [(2, 8)], "PCC"),
    (" PVT Car TP", ("cluster", "segment", "age"), [(9, 14)], "PCT"),       # the header row of this sheet names the table at column 9 (the reader's block)
]


def put(sh, r, c, v):
    row = sh.rows[r - 1]
    row.extend([None] * (c + 1 - len(row)))
    row[c] = v


def with_working(book, only=None, header="RTO"):
    for sheet, keys, blocks, tag in TABLES:
        if only and tag not in only:
            continue
        sh = next(s for s in book.sheets if s.name == sheet)
        hr = G.find_header_row(sh, keys)
        for cluster_col, rto_col in blocks:
            put(sh, hr, rto_col, header)
            for r in range(hr + 1, len(sh.rows) + 1):
                cl = G.norm(sh.rows[r - 1][cluster_col]) if cluster_col < len(sh.rows[r - 1]) else ""
                if cl:
                    seg = (G.norm(sh.rows[r - 1][0]) + " ") if tag == "CV" else ""          # the CV groups differ by segment: so must the typed names, or two rows would share one key
                    put(sh, r, rto_col, f"WL {tag} {seg}{cl}")
    return book


def keyed(df, remarks=False):
    cols = [c for c in df.columns if c not in ("new_rto_list", "from_date", "to_date") and (remarks or c != "remarks")]
    return df[cols].astype(str).agg("|".join, axis=1).sort_values().reset_index(drop=True)


G.RTO_NAMES = "cluster"                                      # the RTO names are not checked in this mode: the remarks can only differ if a flag lost its rows
b0 = G.read_workbook(path)
base, wbase = grun(b0)
check("Go Digit (October workbook): without the column, no run line", note_of(wbase) == "")
mod_book = with_working(G.read_workbook(path))
df, w = grun(mod_book)
n = note_of(w)
check("Go Digit (October workbook): every table with the column is named in the run line", all(x in n for x in ("LATEST CV COMBINED GRID", "Latest 2W 1+1 & SATP", "2W 1+5", "Latest 2W SAOD", "PVT Car Comp+SAOD", "PVT Car TP")), n[:400])
check("Go Digit (October workbook): every row of the upload carries a working-list name (WL <table> <cluster>) - no group name is left", df.new_rto_list.str.startswith("WL ").all(), df[~df.new_rto_list.str.startswith("WL ")].new_rto_list.unique()[:5])
check("Go Digit (October workbook): the names are exactly the typed ones - 'WL <table> <cluster>' style", any(x.startswith("WL CV ") for x in df.new_rto_list) and any(x.startswith("WL PCC ") for x in df.new_rto_list) and any(x.startswith("WL PCT ") for x in df.new_rto_list)
      and any(x.startswith("WL T15 ") for x in df.new_rto_list) and any(x.startswith("WL SAOD ") for x in df.new_rto_list) and any(x.startswith("WL T11 ") for x in df.new_rto_list), sorted(set(df.new_rto_list))[:6])
check("Go Digit (October workbook): the rates and every other column are exactly those of the run without the column (only the RTO name differs)", len(df) == len(base) and keyed(df).equals(keyed(base)), (len(df), len(base)))
check("Go Digit (October workbook): every flag - also the notes of the 'CV Remarks' sections, scoped by product - reaches the same rows as without the column",
      keyed(df, True).equals(keyed(base, True)), (int(df.remarks.str.len().gt(0).sum()), int(base.remarks.str.len().gt(0).sum())))
part = with_working(G.read_workbook(path), only={"CV"})
dp, wp = grun(part)
check("Go Digit (October workbook): a column in only one table: that table's rows are named, the others keep the group names, and the run line says which tables have no column",
      dp.new_rto_list.str.startswith("WL CV").any() and (~dp.new_rto_list.str.startswith("WL ")).any() and "has no RTO working-list column" in note_of(wp), note_of(wp)[:300])

# an RTO column between the first and the last column of a table block (it would take the place of one of the table's own columns): refused with a warning, never read as data
inside = G.read_workbook(path)
sh15 = next(s for s in inside.sheets if s.name == "2W 1+5")
h15 = G.find_header_row(sh15, ("agencypbclusters", "make"))
for inside_col in (8, 15):                                  # inside the first block (columns 3-8) and inside the second (10-15), whichever is the titled one that is read
    free = next((rr for rr in range(1, h15) if (sh15.rows[rr - 1][inside_col] if len(sh15.rows[rr - 1]) > inside_col else None) is None), 1)
    put(sh15, free, inside_col, "RTO")                      # it would shift the cells if it were read as part of the table
di, wi = grun(inside)
G.RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")
check("Go Digit: an RTO column INSIDE the columns of a table block is not used - a warning says where to put it", any("inside the columns of a table block" in x for x in wi), [x[:120] for x in wi][:3])

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
