"""Deterministic checks of rsa_gcv4w_comp_to_config.py on the real October grid (held in memory) with targeted changes to the
GRID (cells, headers, bands, clusters, sheets). How the NOTES are read is in rsa_tests_notes.py.
Run:  .venv\\Scripts\\python.exe tests\\rsa_tests.py"""
import copy
import datetime as dt
import itertools
import os
import re
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("rsa_gcv4w_comp_to_config.py")
REF = os.path.join(helpers.REFS, "rsa_gcv4w_oct26_rates_export.csv")
BASE = helpers.rsa_base("oct")
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
TMP = tempfile.mkdtemp(prefix="rsa_t_")
counter = itertools.count()
WEST = "GCV Comp 4W - WEST"
EAST = "GCV Comp 4W - EAST"
NORTH = "GCV Comp 4W - NORTH"
SOUTH = "GCV Comp 4W - SOUTH"
FakeWB = helpers.FakeWB


class Unexpected(Exception):
    pass


class Smart:
    """Answers questions from rules [(question substring, menu label substring | 'text:...' | 'y' | 'n')]; any menu that
    matches no rule raises Unexpected (so a test notices a question it did not expect)."""

    def __init__(self, rules=()):
        self.rules, self.question, self.labels, self.asked = list(rules), "", {}, []

    def out(self, s):
        s = str(s)
        if s.startswith("\n? "):
            self.question, self.labels = s[3:], {}
            self.asked.append(self.question)
        else:
            m = re.match(r"\s+(\d+)\) (.*)", s)
            if m:
                self.labels[int(m.group(1))] = m.group(2)

    def inp(self, prompt=""):
        if "[Y/n]" in prompt or "[y/N]" in prompt:
            self.asked.append(prompt)
            for q, a in self.rules:
                if q in prompt and a in ("y", "n"):
                    return a
            return "y"
        for q, a in self.rules:
            if q in self.question:
                if prompt.startswith("   Type 1-"):
                    for i, lab in self.labels.items():
                        if a.lower() in lab.lower():
                            return str(i)
                    raise Unexpected(f"rule {q!r}: no option containing {a!r} in {self.labels}")
                if a.startswith("text:"):
                    return a[5:]
        raise Unexpected(f"unexpected question: {self.question!r} {list(self.labels.values())}")


def base():
    return copy.deepcopy(BASE)


def build(sheets, rules=(), interactive=True):
    mod.read_workbook = lambda p: FakeWB(sheets)
    user = Smart(rules)
    ask = mod.Asker(True, user.inp, user.out) if interactive else mod.Asker(False)
    df, warnings = mod.build("fake.xlsx", F, T, quiet, ask)
    return df, warnings, user


def run_main(sheets, rules=(), interactive=True, ref=None):
    mod.read_workbook = lambda p: FakeWB(sheets)
    user = Smart(rules)
    ask = mod.Asker(True, user.inp, user.out) if interactive else mod.Asker(False)
    out = os.path.join(TMP, f"o{next(counter)}.csv")
    path = mod.main("fake.xlsx", F, T, out, ref, quiet, ask)
    return path, user


def find_col(rows, head, text):
    return next(j for j, v in enumerate(rows[head]) if v == text)


def pick(df, cluster, cc_min, make, tmin=None, age=None, zd=None, fuel="Diesel"):
    d = df[(df.new_rto_list == mod.RTO_PREFIX + cluster) & (df.cc_min.astype(float) == cc_min) & (df.new_vehicle_makes == make)
           & (df.fuel_type == fuel)]
    if tmin is not None:
        d = d[d.tariff_min.astype(float) == tmin]
    if age is not None:
        d = d[d.vehicle_age_min.astype(int) == age]
    if zd is not None:
        d = d[d.is_zd == zd]
    assert len(d) == 1, f"expected one row, got {len(d)} for {cluster} {cc_min} {make} {tmin} {age} {zd}"
    return float(d.pi_net_rate.iloc[0]), float(d.po_net_rate.iloc[0])


results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def raises(fn, exc=mod.GridError):
    try:
        fn()
    except exc as e:
        return e
    return None


TATA, EICH, OTH = mod.MAKE_TATA_AL, mod.MAKE_EICHER, mod.MAKE_OTHERS
WEST_RTOS = {"ROYAL_SEPT26_" + n for n in ("AMRAVATI", "AURANGABAD", "KONKAN", "MUMBAI", "NAGPUR", "NASHIK", "PUNE", "REST_OF_MAHARASHTRA",
                                          "DADRA_AND_NAGAR_HAVELI", "DAMAN_&_DIU", "GOA", "LAKSHADWEEP", "AHMEDABAD", "BHARUCH", "GANDHINAGAR",
                                          "JAMNAGAR", "MEHSANA", "NAVSARI", "RAJKOT", "REST_OF_GUJARAT", "SURAT", "VADODARA", "VALSAD")}
GUJARAT = ("AHMEDABAD", "BHARUCH", "GANDHINAGAR", "JAMNAGAR", "MEHSANA", "NAVSARI", "RAJKOT", "REST_OF_GUJARAT", "SURAT", "VADODARA", "VALSAD")

# ---------------------------------------------------------------- 0. the real grid
df, warns, user = build(base())
check("clean grid: no question at all", not user.asked, user.asked)
check("clean grid: 67,392 rows", len(df) == 67392, len(df))
check("clean grid: no warnings", not warns, warns)
logs = []
res = mod.diff_against_reference(df, REF, logs.append)
check("clean grid vs your October upload: 33,696 new / 33,732 removed / 0 changed - the removed are the 36 Electric rows plus the 33,696 TATA,ASHOK rows, which are written "
      "TATA,Ashok (the make / model master's spelling); rates are unchanged (the remarks column is not compared)", res == (33696, 33732, 0), res)
check("every row id 1, 42 columns, DD/MM/YYYY", (df.id == 1).all() and list(df.columns) == mod.OUTPUT_COLUMNS
      and df.from_date.eq("01/10/2026").all() and df.to_date.eq("31/03/2030").all())
path, u = run_main(base())
check("written CSV passes the format checks", not helpers.check_csv(path, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS))
check("clean run asks only the final confirmation", len(u.asked) == 1 and "Write the file now" in u.asked[0], u.asked)

# PUNE (WEST, column I), band D (7.5-12 T, cc_min 7500.01): base 0.19
S = base()
r = S[WEST]
col = find_col(r, 2, "PUNE")
check("hand-check PUNE 7.5-12T TATA high band Old no Nil Dep = 19", pick(df, "PUNE", 7500.01, TATA, 81.01, 1, "NO") == (19.0, 12.0))
check("   ... Nil Dep YES = 17", pick(df, "PUNE", 7500.01, TATA, 81.01, 1, "YES") == (17.0, 10.0))
check("   ... New = 17 (19 - 2)", pick(df, "PUNE", 7500.01, TATA, 81.01, 0, "NA") == (17.0, 10.0))
check("   ... low discount band +2 = 21", pick(df, "PUNE", 7500.01, TATA, 0.0, 1, "NO") == (21.0, 14.0))
check("   ... Eicher -3 = 16, other makes -5 = 14", pick(df, "PUNE", 7500.01, EICH, 0.0, 1, "NO")[0] == 16.0 and pick(df, "PUNE", 7500.01, OTH, 0.0, 1, "NO")[0] == 14.0)
check("   ... PO = PI - 7, floor 0 (3.5-7.5T base 5: PI 5 -> PO 0)", pick(df, "PUNE", 3500.01, TATA, 81.01, 1, "NO") == (5.0, 0.0))
check("   ... New of a cell at 10 or below stays: 3.5-7.5T New = 5", pick(df, "PUNE", 3500.01, TATA, 81.01, 0, "NA")[0] == 5.0)
check("EAST new CV cut is 1 (BHAGALPUR 12-20T 30 -> New 29)", pick(df, "BHAGALPUR", 12000.01, TATA, 85.01, 0, "NA")[0] == 29.0)
check("zero cell stays zero everywhere (EAST 48+T = 0 incl. +2 band)", all(pick(df, "PATNA", 48000.01, mk, t, a, z)[0] == 0.0
      for mk, t in ((TATA, 85.01), (TATA, 0.0), (EICH, 0.0), (OTH, 0.0)) for a, z in ((1, "YES"), (1, "NO"), (0, "NA"))))
check("cluster names: commas, double spaces, DADRA exception",
      {"ROYAL_SEPT26_JORHAT_DIBRUGARH_TINSUKIA", "ROYAL_SEPT26_REST_OF__TELANGANA", "ROYAL_SEPT26_DADRA_AND_NAGAR_HAVELI",
       "ROYAL_SEPT26_ANDAMAN_&_NICOBAR_ISLANDS"} <= set(df.new_rto_list))

# ---------------------------------------------------------------- 1. one cell changes
S = base()
S[WEST][3 + 3][col] = 0.25
d2, _, u = build(S)
check("a changed cell moves every dependent row (0.25: 25 / 23 / New 23 / Eicher 22 / others 20)",
      [pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0], pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "YES")[0],
       pick(d2, "PUNE", 7500.01, TATA, 81.01, 0, "NA")[0], pick(d2, "PUNE", 7500.01, EICH, 0.0, 1, "NO")[0],
       pick(d2, "PUNE", 7500.01, OTH, 0.0, 1, "NO")[0]] == [25.0, 23.0, 23.0, 22.0, 20.0] and not u.asked)
check("... and nothing else moved (4 make/tier combos x 3 variants x 4 fuels = 48 rows)",
      int((d2.pi_net_rate.astype(float) != df.pi_net_rate.astype(float)).sum()) == 48)

# ---------------------------------------------------------------- 2. odd cells
S = base(); S[WEST][3 + 3][col] = "NA"
d2, _, u = build(S)
check("'NA' in a rate cell = no rate, no question", not u.asked and pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 0.0, u.asked)
S = base(); S[WEST][3 + 3][col] = "abc"
e = raises(lambda: build(S, interactive=False))
check("text in a rate cell with no keyboard: stops with the question", isinstance(e, mod.NeedsAnswer) and "abc" in str(e), e)
d2, _, u = build(S, [("holds", "Use 0")])
check("text in a rate cell, answer 'use 0'", pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 0.0 and len(u.asked) == 1, u.asked)
S = base(); S[WEST][3 + 3][col] = 19
d2, _, u = build(S, [("above 1", "divide by 100")])
check("19 instead of 0.19: asked, 'divide by 100' gives the same as the grid", pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 19.0 and len(u.asked) == 1, u.asked)
S = base(); S[WEST][3 + 3][col] = None
d2, _, u = build(S, [("empty cell", "Use 0")])
check("an empty rate cell is asked about", pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 0.0 and len(u.asked) == 1, u.asked)
S = base(); S[WEST][3 + 3][col] = -0.05
d2, _, u = build(S, [("negative", "Use 0")])
check("a negative rate is asked about", pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 0.0 and len(u.asked) == 1, u.asked)

# ---------------------------------------------------------------- 3. header / layout
S = base(); S[WEST][2][0] = "Weight Band"
d2, _, u = build(S)
check("header 'Weight Band' still found, no question", len(d2) == 67392 and not u.asked, u.asked)
S = base(); S[WEST][2][0] = "Band"
d2, _, u = build(S, [("looks like a GCV 4W sheet", "Tell me"), ("Which row", "text:3")])
check("header 'Band': asks which row, then works", len(d2) == 67392 and len(u.asked) == 2, u.asked)
d2, _, u = build(S, [("looks like a GCV 4W sheet", "Leave this sheet out")])
check("header 'Band': leave that sheet out -> 3 regions", len(d2) == 67392 - 23 * 432, len(d2))
S = base(); S[WEST][2][1] = "Business"
d2, w, u = build(S, [("no 'Age' column", "Old")])
check("'Age' header lost: asked what the rows are; the old Age column becomes a no-rates header warning",
      len(u.asked) == 1 and any("no rates" in x for x in w) and len(d2) == 67392, (u.asked, w))
S = base()
for i in range(3, 12):
    S[EAST][i][1] = "Used"
d2, _, u = build(S, [("neither Old nor New", "Old")])
check("age label 'Used' is asked once (not 9 times)", len(u.asked) == 1 and len(d2) == 67392, u.asked)

# explicit New rows: put a New row under every Old row of WEST
S = base()
rows = S[WEST]
new_rows = []
for i in range(3, 12):
    old = rows[i]
    new = list(old)
    new[1] = "New"
    new[col] = 0.30                       # PUNE New = 30 in the grid
    new_rows.append((i + 1, new))
for k, (pos, new) in enumerate(new_rows):
    rows.insert(pos + k, new)
d2, _, u = build(S)
check("explicit New rows are used instead of the 2-point rule (PUNE 7.5-12T New TATA = 30)",
      pick(d2, "PUNE", 7500.01, TATA, 81.01, 0, "NA")[0] == 30.0 and len(d2) == 67392 and not u.asked, (u.asked,))
check("   ... Old rows unaffected", pick(d2, "PUNE", 7500.01, TATA, 81.01, 1, "NO")[0] == 19.0)


# ---------------------------------------------------------------- 4. notes (the full set of wordings is in rsa_tests_notes.py)
def note_row(rows, needle):
    return next(i for i, r in enumerate(rows) if isinstance(r[0], str) and needle in r[0])


S = base(); i = note_row(S[WEST], "For Eicher Make")
S[WEST][i][0] = "3. For Eicher Make - payout will be 4% lower"
d2, _, u = build(S)
check("a changed note (Eicher 4%) is applied without asking", pick(d2, "PUNE", 7500.01, EICH, 0.0, 1, "NO")[0] == 15.0 and not u.asked, u.asked)
S = base(); i = note_row(S[WEST], "For Eicher Make"); S[WEST][i][0] = "3. Eicher: see circular"
d2, _, u = build(S, interactive=False)
eich_rows = d2[(d2.new_vehicle_makes == EICH) & d2.new_rto_list.map(lambda x: x in WEST_RTOS)]
check("unreadable Eicher note, no keyboard: the run does NOT stop, nothing is asked", len(d2) == 67392 and not u.asked)
check("... Eicher rate is left uncut (19, not 16) and ONLY the WEST Eicher rows are flagged",
      pick(d2, "PUNE", 7500.01, EICH, 0.0, 1, "NO")[0] == 19.0 and (eich_rows.remarks != "").all()
      and int(((d2.remarks != "") & d2.new_rto_list.isin(WEST_RTOS)).sum()) == len(eich_rows) == 23 * 9 * 12
      and "payout reduction for Eicher make not found" in eich_rows.remarks.iloc[0]
      and eich_rows.remarks.iloc[0].startswith("Manual Review Required: ["), eich_rows.remarks.iloc[0][:200])
d3, _, u = build(S, [])
check("... the same with a keyboard: still nothing asked", len(u.asked) == 0 and d3.remarks.equals(d2.remarks))
S = base(); i = note_row(S[EAST], "All new CV"); S[EAST][i][0] = "1. All new CV - payout will be 2% lower than the current grid, If the current payout is less than 12% then same payout as current grid will be maintained"
d2, _, u = build(S)
check("EAST note edited to 2% / 12%: both read (BHAGALPUR 12-20T 30 -> New 28; 3.5-7.5T stays)",
      pick(d2, "BHAGALPUR", 12000.01, TATA, 85.01, 0, "NA")[0] == 28.0 and not u.asked, u.asked)
S = base(); i = note_row(S[WEST], "at max")
S[WEST][i][0] = "5. Discounts for Tata & AL makes will be at max  88%  for all tonnages, for discount less than  83% additional 3% payout"
d2, _, u = build(S)
check("discount note edited (88 / 83 / +3): tariff bands and extra follow",
      pick(d2, "PUNE", 7500.01, TATA, 83.01, 1, "NO")[0] == 19.0 and pick(d2, "PUNE", 7500.01, TATA, 0.0, 1, "NO")[0] == 22.0
      and set(d2[(d2.new_rto_list == "ROYAL_SEPT26_PUNE") & (d2.new_vehicle_makes == TATA)].tariff_max.astype(float)) == {83.01, 88.01}, u.asked)
S = base(); S[WEST].insert(24, ["11. Extra 3% payout for Tata in Gujarat"] + [None] * (len(S[WEST][0]) - 1))
d2, _, u = build(S, interactive=False)
guj = d2[d2.remarks.str.contains("Extra 3% payout for Tata in Gujarat", regex=False)]
check("a new, unknown note naming Gujarat does not stop the run and is flagged on the Gujarat clusters only (all makes)",
      len(d2) == 67392 and not u.asked and len(guj) == 11 * 432 and set(guj.new_rto_list) == {"ROYAL_SEPT26_" + c for c in GUJARAT}
      and (d2[d2.new_rto_list == "ROYAL_SEPT26_PUNE"].remarks == "").all(), (len(guj), u.asked))
check("the October North notes 7-11 are flagged in remarks, not warnings", not build(base())[1])

# ---------------------------------------------------------------- 5. clusters / columns / sheets
S = base(); r = S[WEST]
for row in r[:12]:
    row.append(None)
r[0][-1], r[1][-1], r[2][-1] = "WEST", "Gujarat", "NEWTOWN, ZONE 1"
for i in range(3, 12):
    r[i][-1] = 0.2
d2, _, u = build(S)
check("a new cluster column is included (+432 rows) with its comma name converted", len(d2) == 67392 + 432 and
      "ROYAL_SEPT26_NEWTOWN__ZONE_1" in set(d2.new_rto_list) and not u.asked, (len(d2), u.asked))
S = base()
for row in S[WEST]:
    del row[col]
d2, _, u = build(S)
check("a removed cluster just disappears (-432 rows), no question", len(d2) == 67392 - 432 and not u.asked)
S = base()
for row in S[WEST][:12]:
    row.append(None)
S[WEST][2][-1] = "Remarks"
d2, w, u = build(S)
check("a header with no rates under it ('Remarks') is a warning, not a cluster", len(d2) == 67392 and any("Remarks" in x for x in w), w)
S = base()
for row in S[WEST][:12]:
    row.append(None)
for i in range(3, 12):
    S[WEST][i][-1] = 0.1
d2, w, u = build(S)
check("rates in a column with no cluster name are a warning", any("no cluster name" in x for x in w), w)
S = base(); S[WEST][2][col + 1] = S[WEST][2][col]
path, u = run_main(S, [])
out = pd.read_csv(path, dtype=str, keep_default_na=False)
check("the same cluster name twice: no question about the repeated key (only the final 'write the file' confirmation), the lower payin of each entry is kept -> 155 clusters x 432 rows",
      len(out) == 155 * 432 and not any("repeat a key" in q or "duplicate" in q for q in u.asked), (len(out), u.asked))
S = base(); del S[NORTH]
d2, _, u = build(S)
check("a missing sheet (NORTH) is fine: other three regions", len(d2) == 67392 - 49 * 432 and not u.asked)
S = {f"Sheet{i}": rows for i, rows in enumerate(base().values())}
for rows in S.values():
    rows[2][0] = "x"
    rows[0][0] = "x"
e = raises(lambda: build(S))
check("no usable sheet at all: clean stop", e is not None and "no GCV 4W comprehensive sheet" in str(e), e)

# ---------------------------------------------------------------- 6. bands, title, fuel
S = base()
for i in range(3, 12):
    if S[EAST][i][0].startswith("D."):
        S[EAST][i][0] = "D. Heavy"
e = raises(lambda: build(S, interactive=False))
check("unreadable band label, no keyboard: stops with the question", isinstance(e, mod.NeedsAnswer) and "weight band" in str(e), e)
d2, _, u = build(S, [("weight limits", "text:7.5-12")])
check("unreadable band label: typed limits give the same rows", len(d2) == 67392 and len(u.asked) == 1, u.asked)
S = base()
for i in range(3, 12):
    if S[WEST][i][0].startswith("C."):
        S[WEST][i][0] = "C. 3.5 to 8 Ton"
d2, _, u = build(S, [("do not join up", "Carry on")])
check("bands that no longer join up are asked about", len(u.asked) == 1 and "do not join" in u.asked[0], u.asked)
S = base(); S[WEST][0][0] = "GCV - 4W - SATP"
S = {("GCV Satp 4W - WEST" if k == WEST else k): v for k, v in S.items()}
d2, _, u = build(S, [("title", "Leave this sheet out")])
check("a sheet titled SATP instead of COMP is asked about", len(u.asked) == 1 and len(d2) == 67392 - 23 * 432, (u.asked, len(d2)))
S = base(); S[WEST][1][0] = "EV"
d2, _, u = build(S, [("NON EV", "Electric rows only")])
check("a sheet labelled EV instead of NON EV is asked about (Electric rows only: 23 clusters x 108)",
      d2.fuel_type.eq("Electric").sum() == 23 * 108 and len(d2) == 67392 - 23 * 432 + 23 * 108, (u.asked, d2.fuel_type.value_counts().to_dict()))
S = base()
for i in range(3, 12):
    S[SOUTH][i][0] = S[SOUTH][i][0].replace("Ton", "T")
d2, _, u = build(S)
check("'Ton' written as 'T' is read without a question", len(d2) == 67392 and not u.asked)
S = base()
for i in range(3, 12):
    S[WEST][i][0] = re.sub(r"[A-I]\.\s*", "", S[WEST][i][0])
d2, _, u = build(S)
check("band labels without 'A. ' prefixes are read without a question", len(d2) == 67392 and not u.asked, u.asked)
S = base()
S[WEST][11][0] = "I. 48 Ton and above"
d2, _, u = build(S)
check("'48 Ton and above' is read as the open top band", len(d2) == 67392 and not u.asked, (len(d2), u.asked))

# ---------------------------------------------------------------- 7. whole-run behaviours
S = base()
mod.RTO_PREFIX = "ROYAL_OCT26_"
d2, _, u = build(S)
mod.RTO_PREFIX = "ROYAL_SEPT26_"
check("RTO prefix setting changes every new_rto_list", d2.new_rto_list.str.startswith("ROYAL_OCT26_").all())
e = raises(lambda: mod.build("x.xlsx", dt.date(2026, 10, 1), dt.date(2026, 9, 1), quiet, mod.Asker()))
check("to_date before from_date is refused", e is not None and "before" in str(e), e)
mod.read_workbook = lambda p: (_ for _ in ()).throw(OSError("file is open in Excel"))
e = None
try:
    mod.build("fake.xlsx", F, T, quiet, mod.Asker(False))
except mod.GridError as ex:
    e = ex
check("a file that cannot be opened is a clean stop without a keyboard", isinstance(e, mod.NeedsAnswer), e)

bad = [n for n, ok, _ in results if not ok]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
