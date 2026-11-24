"""How rsa_gcv4w_comp_to_config.py reads the plain-text notes - and what it does with the notes it cannot apply: they are
never guessed at, never asked about, never stop the run; the rows they may affect carry 'Manual Review Required: [..]'
in the remarks column.   Run:  .venv\\Scripts\\python.exe tests\\rsa_tests_notes.py"""
import copy
import datetime as dt
import io
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("rsa_gcv4w_comp_to_config.py")
quiet = lambda *a, **k: None
WEST, EAST, NORTH, SOUTH = ("GCV Comp 4W - " + r for r in ("WEST", "EAST", "NORTH", "SOUTH"))
TATA, EICH, OTH = mod.MAKE_TATA_AL, mod.MAKE_EICHER, mod.MAKE_OTHERS
PRE = "Manual Review Required: ["
_CACHE = {}


class Unexpected(Exception):
    pass


class Smart:
    """Records every question; answers none (a note must never cause one)."""

    def __init__(self):
        self.asked = []

    def out(self, s):
        if str(s).startswith("\n? "):
            self.asked.append(str(s)[3:])

    def inp(self, prompt=""):
        if "[Y/n]" in prompt or "[y/N]" in prompt:
            return "y"
        raise Unexpected(f"unexpected question: {self.asked[-1] if self.asked else prompt}")


def base(name="rsa_oct"):
    if name not in _CACHE:
        _CACHE[name] = helpers.rsa_base("oct" if name == "rsa_oct" else "v2")
    return copy.deepcopy(_CACHE[name])


def build(sheets, interactive=True, from_date=dt.date(2026, 10, 1), logs=None):
    mod.read_workbook = lambda p: helpers.FakeWB(sheets)
    user = Smart()
    ask = mod.Asker(True, user.inp, user.out) if interactive else mod.Asker(False)
    df, warnings = mod.build("fake.xlsx", from_date, dt.date(2030, 3, 31), (logs.append if logs is not None else quiet), ask)
    for c in ("pi_net_rate", "cc_min", "tariff_min", "tariff_max"):
        df[c] = df[c].astype(float)
    return df, warnings, user


def row_of(df, cluster, cc_min, make, tmin, age, zd, fuel="Diesel"):
    d = df[(df.new_rto_list == mod.RTO_PREFIX + cluster) & (df.cc_min == cc_min) & (df.new_vehicle_makes == make)
           & (df.fuel_type == fuel) & (df.vehicle_age_min.astype(int) == age) & (df.is_zd == zd)]
    d = d[d.tariff_min > 0] if tmin == "hi" else d[d.tariff_min == 0]
    assert len(d) == 1, (len(d), cluster, cc_min, make, tmin, age, zd)
    return d.iloc[0]


def pick(df, *args, **kw):
    return float(row_of(df, *args, **kw).pi_net_rate)


def flagged(df, part):
    return df[df.remarks.str.contains(part, regex=False)]


results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def note_row(rows, needle):
    return next(i for i, r in enumerate(rows) if isinstance(r[0], str) and needle in r[0])


WEST_CLUSTERS = [str(c).replace(" ", "_") for c in base()[WEST][2][2:] if c]
WEST_CLUSTERS = [("DADRA_AND_NAGAR_HAVELI" if c == "DADRA_&_NAGAR_HAVELI" else c) for c in WEST_CLUSTERS]
WEST_RTOS = {"ROYAL_SEPT26_" + c for c in WEST_CLUSTERS}


def in_west(df):
    return df.new_rto_list.isin(WEST_RTOS)


# ================================================================ the real V2 grid
logs = []
df, warns, u = build(base("rsa_v2"), from_date=dt.date(2026, 10, 22), logs=logs)
text = "\n".join(logs)
check("V2 grid, from_date 22/10: nothing is asked, 67,392 rows, no warnings", not u.asked and len(df) == 67392 and not warns, (u.asked, warns))
check("V2: EAST 'up to 12 ton 1% lower, above 12 ton same' read", "new CV -1 up to 12 T, 0 above 12 T" in text and "effective 22/10/2026" in text)
check("V2: NORTH / SOUTH 'up to 12 ton 2%, above 12 ton 1%' read", text.count("new CV -2 up to 12 T, -1 above 12 T") == 2)
check("V2: WEST 'existing grid continues' recognised and written (setting 'include')", "WEST': its note says the existing grid continues - written anyway" in text)
north = df[df.remarks != ""]
ids = set(re.findall(r"\[(\d+)\.", " ".join(north.remarks)))
check("V2: only NORTH rows are flagged, with notes 7, 8, 9, 10, 11 and 16 (21,168 rows)",
      len(north) == 21168 and ids == {"7", "8", "9", "10", "11", "16"} and len(north.new_rto_list.unique()) == 49, (len(north), ids))
check("V2: every remark starts 'Manual Review Required: [' and unflagged rows are empty", ((df.remarks == "") | df.remarks.str.startswith(PRE)).all())
check("V2: the log says how many rows are flagged", "MANUAL REVIEW - 21,168 row(s)" in text)
check("V2: PATNA (EAST) 12-20T: grid 31 -> Old 31, New 31 (no cut above 12 T)", (pick(df, "PATNA", 12000.01, TATA, "hi", 1, "NO"), pick(df, "PATNA", 12000.01, TATA, "hi", 0, "NA")) == (31.0, 31.0))
check("V2: LUCKNOW (NORTH) 12-20T: 13 -> New 12 (-1) ; 20-40T: 23 -> 22",
      (pick(df, "LUCKNOW", 12000.01, TATA, "hi", 0, "NA"), pick(df, "LUCKNOW", 20000.01, TATA, "hi", 0, "NA")) == (12.0, 22.0))
check("V2: CHENNAI (SOUTH) 7.5-12T: 19 -> New 17 (-2) ; 5 stays 5",
      (pick(df, "CHENNAI", 7500.01, TATA, "hi", 0, "NA"), pick(df, "CHENNAI", 12000.01, TATA, "hi", 0, "NA")) == (17.0, 5.0))
g = row_of(df, "GWALIOR", 40000.01, TATA, "hi", 1, "NO")
e = row_of(df, "GWALIOR", 40000.01, EICH, "lo", 1, "NO")
check("V2: a Gwalior TATA row carries notes 7, 9, 11 and a Gwalior EICHER row notes 8, 10, 11 (joined with ' | ')",
      re.findall(r"\[(\d+)\.", g.remarks) == ["7", "9", "11"] and re.findall(r"\[(\d+)\.", e.remarks) == ["8", "10", "11"] and g.remarks.count(" | ") == 2,
      (g.remarks[:60], e.remarks[:60]))
check("V2: a Jaipur row also carries note 16 (all makes, Rajasthan); an Amritsar (Punjab) row does not",
      "[16." in row_of(df, "JAIPUR", 40000.01, TATA, "hi", 1, "NO").remarks and "[16." not in row_of(df, "AMRITSAR", 40000.01, TATA, "hi", 1, "NO").remarks)

# from_date 01/10 with the V2 grid: no question, no stop; the rows are flagged where the note says another date
d2, _, u = build(base("rsa_v2"), interactive=False, from_date=dt.date(2026, 10, 1))
pat = row_of(d2, "PATNA", 12000.01, TATA, "hi", 1, "NO")
pune = row_of(d2, "PUNE", 12000.01, TATA, "hi", 1, "NO")
check("V2 typed with from_date 01/10 and no keyboard: it runs; rows carry 01/10/2026; EAST/NORTH/SOUTH rows are flagged 'effective from 22/10/2026', WEST rows are not",
      not u.asked and set(d2.from_date) == {"01/10/2026"} and "effective from 22/10/2026" in pat.remarks and pune.remarks == "", (pat.remarks[:120], pune.remarks))

mod.UNCHANGED_SHEETS = "skip"
d2, _, u = build(base("rsa_v2"), from_date=dt.date(2026, 10, 22))
check("UNCHANGED_SHEETS = 'skip': WEST left out (-23 x 432 rows)", len(d2) == 67392 - 23 * 432 and not u.asked, len(d2))
mod.UNCHANGED_SHEETS = "include"

# ================================================================ other wordings of the same rule (V1 grid, WEST: PUNE 7.5-12T cell = 19)
VARIANTS = [
    ("3. For Eicher Make - payout will be reduced by 4%", 4), ("3. For Eicher make - deduction of three", 3),
    ("3. For Eicher Make - payout will be three percent lower", 3), ("3. For Eicher make payout is 4 points less", 4),
    ("3. Deduction of 4% for Eicher make", 4), ("3. For Eicher Make - payout will be lower by 4%", 4),
    ("3. For Eicher Make - payout will be 2.5% lower", 2.5), ("3. For Eicher Make - payout will be 4% higher", -4),
    ("3. For EICHER MAKE: payout will be 4 % lower", 4), ("3. For Eicher make - the payout will be 4% lower than the current grid", 4),
]
for phrase, cut in VARIANTS:
    S = base(); i = note_row(S[WEST], "For Eicher Make"); S[WEST][i][0] = phrase
    d2, w, u = build(S)
    got = pick(d2, "PUNE", 7500.01, EICH, "lo", 1, "NO")
    check(f"Eicher wording: {phrase[3:]!r} -> cut {cut:g}, nothing flagged in WEST", got == 19.0 - cut and not u.asked and not w
          and (d2[in_west(d2)].remarks == "").all(), (got, u.asked, w))
S = base(); i = note_row(S[WEST], "For Eicher Make"); S[WEST][i][0] = "3. For Eicher and Tata makes - payout will be 4% lower"
d2, w, u = build(S)
check("one note for two make groups (Eicher and Tata -4)", pick(d2, "PUNE", 7500.01, EICH, "lo", 1, "NO") == 15.0 and pick(d2, "PUNE", 7500.01, TATA, "hi", 1, "NO") == 15.0 and not u.asked)
for phrase, cut in [("8. Business sourced with Nil Dep - deduction of 3", 3), ("8. Business sourced with Nil Dep - payout will be 3% less", 3),
                    ("8. Business sourced with Nil Dep - payout will be three percent lower", 3), ("8. With Nil Dep payout will be lower by 3%", 3)]:
    S = base(); i = note_row(S[WEST], "with Nil Dep"); S[WEST][i][0] = phrase
    d2, w, u = build(S)
    got = pick(d2, "PUNE", 7500.01, TATA, "hi", 1, "YES")
    check(f"Nil Dep wording: {phrase[3:]!r} -> {cut:g}", got == 19.0 - cut and not u.asked and not w and (d2[in_west(d2)].remarks == "").all(), (got, u.asked, w))
for phrase, cut in [("1. All new CV - deduction of 3%, If the current payout is less than 10% then same payout as current grid will be maintained", 3),
                    ("1. All new CV - payout will be three percent lower than the current grid, If the current payout is less than 10% then same payout as current grid will be maintained", 3),
                    ("1. All new CV payout will be reduced by 1.5%. (If the current payout is below 10% then same payout will be maintained)", 1.5)]:
    S = base(); i = note_row(S[WEST], "All new CV"); S[WEST][i][0] = phrase
    d2, w, u = build(S)
    got = pick(d2, "PUNE", 7500.01, TATA, "hi", 0, "NA")
    check(f"New CV wording: {phrase[3:45]!r}... -> {cut:g}", got == 19.0 - cut and not u.asked and not w and (d2[in_west(d2)].remarks == "").all(), (got, u.asked, w))
S = base(); i = note_row(S[WEST], "at max")
S[WEST][i][0] = "5. Discounts for Tata & AL makes will be at maximum 88% for all tonnages, for discount below 83% additional 3% payout"
d2, w, u = build(S)
check("discount note with 'maximum' / 'below'", pick(d2, "PUNE", 7500.01, TATA, "hi", 1, "NO") == 19.0 and not u.asked and (d2[in_west(d2)].remarks == "").all())
S = base(); i = note_row(S[WEST], "This is effective")
for phrase in ("10. w.e.f. 01/10/2026", "10. Effective date - 1 October 2026", "10. Applicable from 01-10-2026", "10. This grid is valid till 31 Mar 2027"):
    S[WEST][i][0] = phrase
    d2, _, u = build(S)
    check(f"effective-date wording {phrase[4:]!r} read: nothing asked, nothing flagged", not u.asked and (d2[in_west(d2)].remarks == "").all())

# ================================================================ notes that cannot be applied: flagged, never asked, never stop
S = base(); i = note_row(S[WEST], "For Eicher Make"); S[WEST][i][0] = "3. For Eicher & Mahindra makes - payout will be 4% lower"
d2, w, u = build(S, interactive=False)
wr = d2[in_west(d2)]
check("an unknown make in the Eicher note: the run does not stop; Eicher stays uncut (19) and is flagged; the note's own text is flagged on the whole sheet",
      not u.asked and pick(d2, "PUNE", 7500.01, EICH, "lo", 1, "NO") == 19.0 and len(wr) == 23 * 432 and (wr.remarks != "").all()
      and wr.remarks.str.contains("Eicher & Mahindra makes", regex=False).all()
      and flagged(wr, "payout reduction for Eicher make not found").new_vehicle_makes.eq(EICH).all(), (len(wr), u.asked))

S = base(); i = note_row(S[WEST], "All new CV")
S[WEST][i][0] = "1. All new CV upto 12 ton - payout will be 2% lower than the current grid, If the current payout is less than 10% then same payout as current grid will be maintained"
d2, w, u = build(S, interactive=False)
fl = flagged(d2, "only gives a reduction for up to 12 ton")
check("'up to 12 ton' with nothing said above it: no stop; bands up to 12 T are cut (-2), heavier bands are not and their New rows are flagged",
      not u.asked and pick(d2, "PUNE", 7500.01, TATA, "hi", 0, "NA") == 17.0
      and pick(d2, "PUNE", 20000.01, TATA, "hi", 0, "NA") == pick(d2, "PUNE", 20000.01, TATA, "hi", 1, "NO")
      and len(fl) == 23 * 5 * 4 * 4 and fl.add_tnc.eq("New").all() and set(fl.cc_min) == {12000.01, 20000.01, 40000.01, 45000.01, 48000.01}, (len(fl), u.asked))

S = base(); i = note_row(S[WEST], "All new CV")
S[WEST][i][0] = "1. All new CV - payout will be 2% lower than the current grid. For GVW above 15 ton - payout will be 1% lower. If the current payout is less than 10% then same payout as current grid will be maintained"
d2, _, u = build(S, interactive=False)
fl = flagged(d2, "crosses or falls between")
check("'all new CV 2%' + 'above 15 ton 1%': 7.5-12 T cut 2, 20-40 T cut 1, the 12-20 T band (it crosses 15 ton) is uncut and its New rows are flagged",
      not u.asked and pick(d2, "PUNE", 7500.01, TATA, "hi", 0, "NA") == 17.0
      and pick(d2, "PUNE", 20000.01, TATA, "hi", 0, "NA") == pick(d2, "PUNE", 20000.01, TATA, "hi", 1, "NO") - 1
      and pick(d2, "PUNE", 12000.01, TATA, "hi", 0, "NA") == pick(d2, "PUNE", 12000.01, TATA, "hi", 1, "NO")
      and len(fl) == 23 * 1 * 4 * 4 and set(fl.cc_min) == {12000.01}, (len(fl), u.asked))
S = base(); i = note_row(S[WEST], "All new CV")
S[WEST][i][0] = "1. All new CV upto 12 ton - payout will be 2% lower than the current grid. For GVW above 15 ton - payout will be 1% lower. (If the current payout is less than 10% then same payout will be maintained)"
d2, _, u = build(S, interactive=False)
fl = flagged(d2, "crosses or falls between")
check("limits that disagree (12 and 15): no stop; the band between them is uncut and flagged", not u.asked and len(fl) == 23 * 4 * 4 and set(fl.cc_min) == {12000.01})
S = base()
for r in S[WEST][3:12]:
    if r[0].startswith("D."):
        r[0] = "D. 7.5 to 14 Ton"
    if r[0].startswith("E."):
        r[0] = "E. 14 to 20 Ton"
i = note_row(S[WEST], "All new CV")
S[WEST][i][0] = "1. All new CV upto 12 ton - payout will be 2% lower than the current grid. For GVW above 12 ton - payout will be 1% lower. (If the current payout is less than 10% then same payout will be maintained)"
d2, _, u = build(S, interactive=False)
fl = flagged(d2, "crosses or falls between")
check("a weight band straddling the 12-ton limit is uncut and flagged; the next band takes the 'above' cut", not u.asked and len(fl) == 23 * 4 * 4 and set(fl.cc_min) == {7500.01}
      and pick(d2, "PUNE", 14000.01, TATA, "hi", 0, "NA") == pick(d2, "PUNE", 14000.01, TATA, "hi", 1, "NO") - 1)

S = base(); i = note_row(S[WEST], "This is effective")
S[WEST][i][0] = "10. This is effective from 31/02/2026"
d2, _, u = build(S, interactive=False)
check("an impossible effective date is not silently skipped: the text is flagged on the whole WEST sheet", not u.asked and (d2[in_west(d2)].remarks.str.contains("31/02/2026", regex=False)).all())
S = base(); S[WEST][i][0] = "10. Effective from 1st Nov 2026"
d2, _, u = build(S, interactive=False)
check("a different effective date than from_date: no stop; the WEST rows are flagged", not u.asked and (d2[in_west(d2)].remarks.str.contains("effective from 01/11/2026", regex=False)).all())
S = base(); S[WEST].insert(24, ["11. For GVW up to 3 ton - extra 1% payout for Tata"] + [None] * (len(S[WEST][0]) - 1))
d2, _, u = build(S, interactive=False)
check("a brand-new kind of rule (extra 1% up to 3 ton): not applied, nothing stops, flagged on the sheet",
      not u.asked and (d2[in_west(d2)].remarks.str.contains("extra 1% payout for Tata", regex=False)).all() and pick(d2, "PUNE", 7500.01, TATA, "hi", 1, "NO") == 19.0)
S = base(); S[WEST].insert(24, ["11. Extra 3% payout for Tata in Gujarat"] + [None] * (len(S[WEST][0]) - 1))
d2, _, u = build(S, interactive=False)
guj = flagged(d2, "Extra 3% payout for Tata in Gujarat")
check("a note naming Gujarat is flagged on the Gujarat clusters only", len(guj) == 11 * 432 and not u.asked and (d2[d2.new_rto_list == "ROYAL_SEPT26_PUNE"].remarks == "").all())


# a rule the notes do not give: a neutral value (no cut / no extra / one discount band) and a flag on the rows it would have touched
def without(sheet_rows, needle):
    del sheet_rows[note_row(sheet_rows, needle)]


S = base(); without(S[WEST], "at max  86%")
d2, w, u = build(S, interactive=False)
t = d2[in_west(d2) & (d2.new_vehicle_makes == TATA)]
check("Tata discount note missing: one band 0-100.01, no +2 extra, Tata rows flagged (and only those)",
      not u.asked and set(zip(t.tariff_min, t.tariff_max)) == {(0.0, 100.01)} and pick(d2, "PUNE", 7500.01, TATA, "lo", 1, "NO") == 19.0
      and t.remarks.str.contains("maximum discount and the extra payout for Tata & AL makes", regex=False).all()
      and (d2[in_west(d2) & (d2.new_vehicle_makes != TATA)].remarks == "").all())
S = base(); without(S[WEST], "at max 81%")
d2, w, u = build(S, interactive=False)
o = d2[in_west(d2) & (d2.new_vehicle_makes != TATA)]
check("Eicher/others discount note missing: one band up to 100.01 and those rows flagged", not u.asked and set(o.tariff_max) == {100.01} and o.remarks.str.contains("maximum discount for Eicher and other makes", regex=False).all())
S = base(); without(S[WEST], "with Nil Dep")
d2, w, u = build(S, interactive=False)
yes = d2[in_west(d2) & (d2.is_zd == "YES")]
check("Nil Dep note missing: Nil Dep rows are not reduced (19) and flagged; non-Nil-Dep rows are not", not u.asked and pick(d2, "PUNE", 7500.01, TATA, "hi", 1, "YES") == 19.0
      and yes.remarks.str.contains("Nil Dep business not found", regex=False).all() and (d2[in_west(d2) & (d2.is_zd != "YES")].remarks == "").all())
S = base(); i = note_row(S[WEST], "All new CV"); S[WEST][i][0] = "1. All new CV - payout will be 2% lower than the current grid"
d2, w, u = build(S, interactive=False)
new_rows = d2[in_west(d2) & (d2.add_tnc == "New")]
check("the 'cells of 10% or less stay' clause missing: the cut is applied to every cell and the New rows are flagged",
      not u.asked and pick(d2, "PUNE", 3500.01, TATA, "hi", 0, "NA") == 3.0 and new_rows.remarks.str.contains("payout at or below which New CV stays unreduced", regex=False).all()
      and (d2[in_west(d2) & (d2.add_tnc != "New")].remarks == "").all())
S = base(); without(S[WEST], "All new CV")
d2, w, u = build(S, interactive=False)
new_rows = d2[in_west(d2) & (d2.add_tnc == "New")]
check("New CV note missing: New = Old and every New row of the sheet is flagged", not u.asked and pick(d2, "PUNE", 7500.01, TATA, "hi", 0, "NA") == 19.0
      and new_rows.remarks.str.contains("New CV payout reduction was not found", regex=False).all() and (d2[in_west(d2) & (d2.add_tnc != "New")].remarks == "").all())

# the remarks survive a trip through CSV (commas, quotes, line breaks inside a note)
S = base(); i = note_row(S[NORTH], "Additional 2% payout")
S[NORTH][i][0] = '11. Discounts - "Rajasthan", MP: 50%, see\nattached sheet; Jaipur is 10% (no extra) for the RTOs named'
d2, _, u = build(S, interactive=False)
buf = io.StringIO()
mod.to_csv_frame(d2).to_csv(buf, index=False, lineterminator="\n")
back = pd.read_csv(io.StringIO(buf.getvalue()), dtype=str, keep_default_na=False)
check("remarks with commas, quotes and line breaks come back unchanged from a CSV", back.remarks.tolist() == d2.remarks.tolist() and back.remarks.str.contains('"Rajasthan", MP: 50%', regex=False).any())
check("... and the CSV still has the 42 upload columns", list(back.columns) == mod.OUTPUT_COLUMNS)

bad = [n for n, ok, _ in results if not ok]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
