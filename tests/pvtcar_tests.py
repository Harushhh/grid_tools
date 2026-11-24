"""Deterministic checks of rsa_pvtcar_to_config.py on the real October Private Car grid (held in memory): the rates, the naming,
how the notes are read and flagged, and the questions asked when the GRID itself is odd.
Run:  .venv\\Scripts\\python.exe tests\\pvtcar_tests.py"""
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

mod = helpers.load_script("rsa_pvtcar_to_config.py")
REF = os.path.join(helpers.REFS, "rsa_pvtcar_oct26_rates_export.csv")
BASE = helpers.pvtcar_base()
F, T = dt.date(2026, 10, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
TMP = tempfile.mkdtemp(prefix="pvtcar_t_")
counter = itertools.count()
WEST, NORTH, EAST, SOUTH = ("Pvt Car- 4W-COMP & SAOD " + r for r in ("West", "North", "East", "South"))
STP, TOPLIST = "Pvt Car STP", "Top End Variant List"
FakeWB = helpers.FakeWB
P = mod.RTO_PREFIX + mod.COMP_RTO_PART
PS = mod.RTO_PREFIX + mod.STP_RTO_PART


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


def build(sheets, rules=(), interactive=True, hidden=()):
    mod.read_workbook = lambda p: FakeWB(sheets, hidden)
    user = Smart(rules)
    ask = mod.Asker(True, user.inp, user.out) if interactive else mod.Asker(False)
    df, warnings = mod.build("fake.xlsx", F, T, quiet, ask)
    return df, warnings, user


def run_main(sheets, rules=(), ref=None):
    mod.read_workbook = lambda p: FakeWB(sheets)
    user = Smart(rules)
    out = os.path.join(TMP, f"o{next(counter)}.csv")
    path = mod.main("fake.xlsx", F, T, out, ref, quiet, mod.Asker(True, user.inp, user.out))
    return path, user


def head(rows):
    return next(i for i, r in enumerate(rows) if "Business Type" in r)


def col_of(rows, name):
    return next(j for j, v in enumerate(rows[head(rows)]) if v == name)


def trow(rows, kind, cc, ncb):
    """index of the table row: kind = 'Comp' / 'SAOD' / 'New', cc as written in the grid, ncb as written"""
    h = head(rows)
    return next(i for i in range(h + 1, h + 20) if rows[i][1] and kind.lower() in str(rows[i][1]).lower()
                and not (kind == "Comp" and "saod" in str(rows[i][1]).lower())
                and str(rows[i][2]).strip() == cc and str(rows[i][3]).strip() == ncb)


def note_row(rows, text):
    return next(i for i, r in enumerate(rows) for v in r if isinstance(v, str) and text in v)


def note_cell(rows, text):
    i = note_row(rows, text)
    return i, next(j for j, v in enumerate(rows[i]) if isinstance(v, str) and text in v)


def add_note(rows, text):
    """a new line of text under the table (column B), as Royal would add a note"""
    rows.append([None, text] + [None] * (len(rows[0]) - 2))
    return rows


def pick(df, cluster, sub, make, cc_min, ncb, fuel="Petrol", stp=False):
    d = df[(df.new_rto_list == (PS if stp else P) + cluster) & (df.sub_product == sub) & (df.new_vehicle_makes == make)
           & (df.cc_min.astype(float) == cc_min) & (df.is_ncb == ncb) & (df.fuel_type == fuel)]
    assert len(d) == 1, f"expected one row, got {len(d)} for {cluster} {sub} {make} {cc_min} {ncb} {fuel}"
    r = d.iloc[0]
    return (float(r.pi_tp_rate), float(r.po_tp_rate)) if stp else (float(r.pi_od_rate), float(r.po_od_rate))


TOP, OTH, ALL = mod.MAKES_TOP, mod.MAKES_OTHER, mod.MAKES_STP
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def raises(fn, exc=Exception):
    try:
        fn()
    except exc as e:
        return e
    return None


# ================================================================ the October grid
df0, warn0, user0 = build(base())
check("October grid builds with no question and no warning", user0.asked == [] and warn0 == [], (user0.asked, warn0))
check("19,620 rows: 237 clusters, 5 fuels, 42 columns", len(df0) == 19620 and df0.new_rto_list.nunique() == 237 and set(df0.fuel_type) == set(mod.FUELS)
      and list(df0.columns) == mod.OUTPUT_COLUMNS, (len(df0), df0.new_rto_list.nunique()))
d = mod.diff_against_reference(df0, REF, quiet)
check("against the October upload: 2,190 new rows (405 for the 3 North-East clusters + 1,785 New rows for other makes), 0 removed, 0 changed", d == (2190, 0, 0), d)
check("every cell of the October upload is reproduced (17,430 matching keys)", d[1] == 0 and d[2] == 0)
check("no duplicate keys", int(df0.duplicated(mod.KEY_COLUMNS).sum()) == 0)
check("id = 1, user_id 2966, dates DD/MM/YYYY, age 0-99, tariff 0-100.01, sc 0-999 on every row",
      (df0.id == 1).all() and (df0.user_id == 2966).all() and (df0.from_date == "01/10/2026").all() and (df0.to_date == "31/03/2030").all()
      and (df0.vehicle_age_min == 0).all() and (df0.vehicle_age_max == 99).all() and (df0.tariff_max == 100.01).all() and (df0.sc_max == 999).all())

# ---- rates
check("Mumbai Comp With NCB <=1000: top-variant group 28 / 21, every other make 26 / 19",
      pick(df0, "MUMBAI", "1+1", TOP, 0, "YES") == (28, 21) and pick(df0, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19))
check("Mumbai SAOD Without NCB >1500: 19 / 12 for the top variants, 17 / 10 for the rest",
      pick(df0, "MUMBAI", "SAOD", TOP, 1500.01, "NO") == (19, 12) and pick(df0, "MUMBAI", "SAOD", OTH, 1500.01, "NO") == (17, 10))
check("Goa Comp With NCB: 23.5 -> 25.5 for the top variants, PO 16.5 / 18.5 (half points are kept)",
      pick(df0, "GOA", "1+1", OTH, 0, "YES") == (23.5, 16.5) and pick(df0, "GOA", "1+1", TOP, 0, "YES") == (25.5, 18.5))
check("New car: the top-variant group gets 30 / 32 / 35 as in the grid (no +2), is_ncb NA, policy_type New",
      [pick(df0, "MUMBAI", "1+3", TOP, c, "NA") for c in (0, 1000.01, 1500.01)] == [(30, 23), (32, 25), (35, 28)]
      and set(df0[df0.sub_product == "1+3"].new_vehicle_makes) == {TOP, OTH} and set(df0[df0.sub_product == "1+3"].policy_type) == {"New"})
check("Comp and SAOD are policy_type Renew; STP is Renew too", set(df0[df0.sub_product.isin(["1+1", "SAOD"])].policy_type) == {"Renew"} and set(df0[df0.sub_product == "STP"].policy_type) == {"Renew"})
check("every fuel gets the same rate (Electric included)", df0[df0.sub_product != "STP"].groupby(["new_rto_list", "sub_product", "new_vehicle_makes", "cc_min", "is_ncb"])["pi_od_rate"].nunique().max() == 1)
check("engine-size bands are 0-1000.01 / 1000.01-1500.01 / 1500.01-99999",
      set(zip(df0.cc_min, df0.cc_max)) == {(0, 1000.01), (1000.01, 1500.01), (1500.01, 99999)})
check("comp rows are 'On OD' (pi_od_rate / po_od_rate, TP 0); STP rows are 'On TP' (pi_tp_rate / po_tp_rate, OD 0)",
      (df0[df0.sub_product != "STP"].pi_type == "On OD").all() and (df0[df0.sub_product != "STP"].pi_tp_rate == 0).all()
      and (df0[df0.sub_product == "STP"].pi_type == "On TP").all() and (df0[df0.sub_product == "STP"].pi_od_rate == 0).all()
      and (df0[df0.sub_product == "STP"].po_od_rate == 0).all())
pi = df0[["pi_od_rate", "pi_tp_rate"]].max(axis=1)
po = df0[["po_od_rate", "po_tp_rate"]].max(axis=1)
check("PO = PI - 7, never below 0, zero stays zero", ((po - (pi - 7).clip(lower=0)).abs() < 0.0051).all() and (po[pi == 0] == 0).all())
check("make groups: top variants / other / Private_Car_ALL for STP", set(df0.new_vehicle_makes) == {TOP, OTH, ALL} and set(df0[df0.sub_product == "STP"].new_vehicle_makes) == {ALL})

# ---- STP
check("STP Andaman: Diesel 1000-1500 = 35 / 28 and Petrol <1000 = 32 / 25; Electric / CNG / LPG take the Non Diesel column",
      pick(df0, "ANDAMAN_and_NICOBAR", "STP", ALL, 1000.01, "NA", "Diesel", True) == (35, 28)
      and all(pick(df0, "ANDAMAN_and_NICOBAR", "STP", ALL, 0, "NA", f, True) == (32, 25) for f in ("Petrol", "CNG", "LPG", "Electric")))
check("STP rates are rounded to 2 decimals: Daman Diesel 1000-1500 = 13.04 / 6.04 (grid 0.1303743...); Jaipur Petrol <1000 = 7.02 / 0.02",
      pick(df0, "DAMAN_AND_DIU", "STP", ALL, 1000.01, "NA", "Diesel", True) == (13.04, 6.04) and pick(df0, "JAIPUR", "STP", ALL, 0, "NA", "Petrol", True) == (7.02, 0.02))
check("STP: a payin of 5 gives payout 0 (floored); a zero cell stays 0 / 0",
      pick(df0, "GUNTUR", "STP", ALL, 0, "NA", "Petrol", True) == (8, 1) and pick(df0, "DAMAN_AND_DIU", "STP", ALL, 0, "NA", "Diesel", True) == (0, 0)
      and pick(df0, "CHITTOOR", "STP", ALL, 1000.01, "NA", "Petrol", True) == (5, 0))
check("STP has 118 divisions x 3 bands x 5 fuels = 1,770 rows", int((df0.sub_product == "STP").sum()) == 1770)

# ---- naming
names = set(df0.new_rto_list)
check("naming: Baroda -> VADODARA, Delhi NCR -> Delhi_NCR, Mysore, Vijayawada, Vishakapatnam -> VISAKHAPATANAM, Ernakulam -> ERNAKULAM/KOCHI, Bhuvaneshwar -> BHUBANESWAR",
      all(P + n in names for n in ("VADODARA", "Delhi_NCR", "Mysore", "Vijayawada", "VISAKHAPATANAM", "ERNAKULAM/KOCHI", "BHUBANESWAR", "ANDAMAN_and_NICOBAR")), "")
check("naming: text in brackets is dropped, a name over two lines is joined, a double space is kept",
      all(P + n in names for n in ("KOLHAPUR", "VARANASI", "MANGALORE", "HUBLI", "REST_OF_KARNATAKA", "REST_OF__TELANGANA")), "")
check("naming: commas become underscores (North-East divisions); STP names use PVT_STP_CAR_ and the STP sheet's own spelling",
      PS + "ARUNACHAL_PRADESH_MANIPUR_NAGALAND_TRIPURA" in names and P + "JORHAT_DIBRUGARH_TINSUKIA" in names and PS + "Delhi_NCR" in names
      and PS + "Mysore" in names and PS + "REST_OF__TELANGANA" in names and PS + "ERNAKULAM/KOCHI" in names)
check("no cluster name has a space, a comma or a bracket", not any(re.search(r"[\s,()]", n) for n in names))

# ---- remarks in the October grid
fl = df0[df0.remarks != ""]
check("October: nothing is flagged - every note is read (extra, other-make New rate, makes) or standing (VLI tables, titles)",
      len(fl) == 0, (len(fl), fl.remarks.iloc[0][:100] if len(fl) else ""))

# ================================================================ notes: how they are read
def with_note(sheet, find, new):
    s = base()
    i, j = note_cell(s[sheet], find)
    s[sheet][i][j] = new
    return s


df, _, u = build(with_note(WEST, "Additional 2% payout", "5. Additional 3% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda (List attached)"))
check("'Additional 3%' note -> the top-variant group gets +3 in that sheet, the other makes do not; other sheets still +2",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (29, 22) and pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19)
      and pick(df, "LUCKNOW", "1+1", TOP, 0, "YES") == (26, 19) and u.asked == [], pick(df, "MUMBAI", "1+1", TOP, 0, "YES"))
df, _, _ = build(with_note(WEST, "Additional 2% payout", "5. Extra two percent payout for the top end variants of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda"))
check("another wording ('Extra two percent ... top end variants') is read the same: +2", pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (28, 21) and not df[df.remarks.str.contains("Top Variant", regex=False)].shape[0])
df, _, _ = build(with_note(WEST, "Additional 2% payout", "5. Additional 2.5% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda"))
check("a decimal extra (2.5%) is read: 28.5", pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (28.5, 21.5))

s = base()
i, j = note_cell(s[WEST], "Additional 2% payout")
s[WEST][i][j] = None
df, _, _ = build(s)
w = df[df._sheet == WEST] if "_sheet" in df else None
fl = df[df.remarks.str.contains("Top Variant Models' note was not found", regex=False)]
check("the extra note missing from West: no extra (top variants = the cell), and the West top-variant Comp / SAOD rows are flagged (not New, not other makes, not other sheets)",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (26, 19) and len(fl) == 17 * 12 * 5 and set(fl.new_vehicle_makes) == {TOP} and set(fl.sub_product) == {"1+1", "SAOD"}
      and fl.new_rto_list.isin([P + "MUMBAI", P + "SURAT"]).any() and not fl.new_rto_list.isin([P + "LUCKNOW"]).any(), len(fl))

s = base()
add_note(s[WEST], "6. Additional 4% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda")
df, _, _ = build(s)
fl = df[df.remarks.str.contains("two notes give different values", regex=False)]
check("two notes with different 'Additional N%' values: neither is guessed, no extra is added, the whole West sheet (2,550 rows) is flagged",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (26, 19) and len(fl) == 17 * 30 * 5 and fl.remarks.str.contains("Additional 4%", regex=False).all()
      and not fl.new_rto_list.isin([P + "LUCKNOW"]).any(), len(fl))

# ---- other makes' New car payout (written by default: OTHER_MAKES_NEW = "write")
check("default: New rows for other makes at the note's flat 28 / PO 21, every cc band, every fuel, is_ncb NA, policy_type New; nothing flagged",
      all(pick(df0, "MUMBAI", "1+3", OTH, c, "NA") == (28, 21) for c in (0, 1000.01, 1500.01))
      and len(df0[(df0.sub_product == "1+3") & (df0.new_vehicle_makes == OTH)]) == 119 * 3 * 5
      and set(df0[(df0.sub_product == "1+3") & (df0.new_vehicle_makes == OTH)].policy_type) == {"New"} and (df0.remarks == "").all())
df, _, _ = build(with_note(WEST, "Any other make New car payout", "3. Any other make New car payout will be 30% flat"))
check("the rate comes from the note (30) for that sheet, 28 elsewhere; the top-variant New rows are not touched",
      pick(df, "MUMBAI", "1+3", OTH, 0, "NA") == (30, 23) and pick(df, "LUCKNOW", "1+3", OTH, 0, "NA") == (28, 21) and pick(df, "MUMBAI", "1+3", TOP, 0, "NA") == (30, 23))
df, _, _ = build(with_note(WEST, "Any other make New car payout", "3. Any other make New car payout is twenty five percent flat"))
check("another wording ('is twenty five percent flat') is not understood as a number: no other-make New rows for West, flagged (see below)",
      df[(df.sub_product == "1+3") & (df.new_vehicle_makes == OTH) & df.new_rto_list.isin([P + "MUMBAI"])].empty)
s = base()
i, j = note_cell(s[WEST], "Any other make New car payout")
s[WEST][i][j] = None
df, _, _ = build(s)
fl = df[df.remarks != ""]
check("the note missing from West -> no other-make New rows for West, its New rows are flagged, the other sheets are written",
      df[(df.sub_product == "1+3") & (df.new_vehicle_makes == OTH) & df.new_rto_list.isin([P + "MUMBAI"])].empty and len(fl) == 17 * 3 * 5 and set(fl.sub_product) == {"1+3"}
      and fl.remarks.str.contains("was not found in the notes", regex=False).all() and pick(df, "LUCKNOW", "1+3", OTH, 0, "NA") == (28, 21), len(fl))
df, _, _ = build(with_note(WEST, "Any other make New car payout", "3. Any other make New car payout will be reduced"))
fl = df[df.remarks != ""]
check("the note present but without a number ('will be reduced'): no rows for other makes in West; its New rows say the rate was not found AND the whole sheet carries the unread note",
      df[(df.sub_product == "1+3") & (df.new_vehicle_makes == OTH) & df.new_rto_list.isin([P + "MUMBAI"])].empty
      and fl.remarks.str.contains("was not found in the notes", regex=False).any() and fl.remarks.str.contains("will be reduced", regex=False).any()
      and len(fl) == 17 * 27 * 5 and not fl.new_rto_list.isin([P + "LUCKNOW"]).any(), len(fl))
s = base()
add_note(s[WEST], "6. Any other make New car payout will be 31% flat")
df, _, _ = build(s)
fl = df[df.remarks.str.contains("two notes give different values", regex=False)]
check("two notes with different other-make rates: neither is used (no rows for other makes in West), the whole West sheet is flagged",
      df[(df.sub_product == "1+3") & (df.new_vehicle_makes == OTH) & df.new_rto_list.isin([P + "MUMBAI"])].empty and len(fl) == 17 * 27 * 5, len(fl))

old = mod.OTHER_MAKES_NEW
mod.OTHER_MAKES_NEW = "flag"
try:
    df, _, _ = build(base())
    fl = df[df.remarks != ""]
    check("flag mode (the October behaviour): no New rows for other makes; the New rows (1,785) carry the 28% note and nothing else is flagged",
          df[(df.sub_product == "1+3") & (df.new_vehicle_makes == OTH)].empty and len(fl) == 1785 and set(fl.sub_product) == {"1+3"}
          and fl.remarks.str.contains("28% flat", regex=False).all() and fl.remarks.str.startswith("Manual Review Required: [").all(), len(fl))
    d = mod.diff_against_reference(df, REF, quiet)
    check("flag mode reproduces the October upload exactly apart from the 3 North-East clusters (405 new rows, 0 removed, 0 changed)", d == (405, 0, 0), d)
finally:
    mod.OTHER_MAKES_NEW = old

# ---- makes named in the notes
df, _, _ = build(with_note(WEST, "applicable only for the following makes", "2. Above New Car payout is applicable only for the following makes - Maruti/Tata/Mahindra/Hyundai/Toyota/MG & Skoda."))
fl = df[df.remarks.str.contains("names the makes", regex=False)]
check("a make dropped from the New car note (Kia): the top-variant rows of that sheet (15 table rows x 17 clusters x 5 fuels) are flagged to check the make group; rates unchanged",
      len(fl) == 15 * 17 * 5 and set(fl.new_vehicle_makes) == {TOP} and pick(df, "MUMBAI", "1+3", TOP, 0, "NA") == (30, 23), len(fl))
df, _, _ = build(with_note(WEST, "Additional 2% payout", "5. Additional 2% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG, Skoda and Honda (List attached)"))
check("a make added to the top-variant note (Honda) is flagged too", df.remarks.str.contains("honda", regex=False).any())
df, _, _ = build(with_note(WEST, "Additional 2% payout", "5. Additional 2% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and �koda (List attached)"))
check("the mangled 'S' of Skoda in the grid text ('?koda') is understood: no makes flag", not df.remarks.str.contains("names the makes", regex=False).any())

# ---- other text rules
s = base()
add_note(s[WEST], "6. For Mumbai and Pune the payout will be 2% lower than the grid")
df, _, u = build(s)
fl = df[df.remarks.str.contains("For Mumbai and Pune", regex=False)]
check("a new rule that names clusters is flagged on those clusters only (both make groups, every row of them: 2 x 150), never guessed or asked about",
      u.asked == [] and set(fl.new_rto_list) == {P + "MUMBAI", P + "PUNE"} and len(fl) == 2 * 30 * 5 and pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19), (len(fl), set(fl.new_rto_list)))
s = base()
add_note(s[NORTH], "6. All cities get 1% extra for fleet business")
df, _, _ = build(s)
fl = df[df.remarks.str.contains("fleet business", regex=False)]
check("a new rule that names no cluster flags the whole sheet (North: 29 clusters)", fl.new_rto_list.nunique() == 29 and not fl.new_rto_list.isin([P + "MUMBAI"]).any(), fl.new_rto_list.nunique())
s = base()
add_note(s[SOUTH], "6. Karnataka business will be 2% lower")
df, _, _ = build(s)
fl = df[df.remarks.str.contains("Karnataka business", regex=False)]
check("a state in the text brings in its clusters (Belagavi, Tumkur ...) and all Priority-City columns (the sheet does not say which state they are in: Bangalore, Mysore, Chennai ...), not the other states' clusters (Calicut)",
      {P + "BELAGAVI", P + "TUMKUR", P + "BANGALORE", P + "Mysore", P + "CHENNAI"} <= set(fl.new_rto_list) and P + "CALICUT" not in set(fl.new_rto_list), sorted(set(fl.new_rto_list))[:6])
df, _, _ = build(with_note(WEST, "Base (Payout shown", "Base (Payout shown as per Grid above) - revised"))
check("a standing phrase still matches when the text around it changes (phrase match, not whole-cell)", not df.remarks.str.contains("revised", regex=False).any())
df, _, _ = build(with_note(WEST, "volume linked payouts applicable only for Mumbai", "Pvt Car Roll-over (Comp+SAOD) volume linked payouts applicable only for Mumbai"))
check("volume-linked headings are standing notes: not flagged", not df.remarks.str.contains("volume linked", regex=False).any())
s = base()
s[STP].insert(0, ["PC STP GRID: Diesel above 1500 cc payout reduced by 5% in Kerala from 25 Oct"] + [None] * (len(s[STP][0]) - 1))
df, _, _ = build(s)
fl = df[df.remarks.str.contains("reduced by 5%", regex=False)]
check("a new text above the STP table is flagged on the STP rows only (1,770)", len(fl) == 1770 and set(fl.sub_product) == {"STP"}, len(fl))

# ================================================================ grid problems are asked, not guessed
s = base()
r = trow(s[WEST], "Comp", "Upto 1000", "With NCB")
s[WEST][r][1] = "Roll-over Third Party Plus"
df, _, u = build(s, [("know the business type", "Leave")])
check("an unknown business type asks; 'leave out' drops those rows (only that engine size / NCB row)", len(u.asked) >= 1 and len(df) < len(df0) and (df.sub_product == "1+1").sum() < (df0.sub_product == "1+1").sum(), u.asked[:1])
e = raises(lambda: build(s, interactive=False), mod.GridError)
check("without a keyboard the same problem stops with the question as the message", e is not None and "unknown business type" in str(e), e)
df, _, u = build(s, [("know the business type", "Comprehensive")])
check("... or answered 'Comprehensive': that row is written as Comp (1+1) and the output equals the October one",
      len(df) == len(df0) and pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19))

s = base()
r = trow(s[WEST], "Comp", "Upto 1000", "With NCB")
s[WEST][r][2] = "Small cars"
df, _, u = build(s, [("engine-size limits", "text:upto 1000")])
check("an engine size it cannot read asks once for the limits; typed limits are used", len(u.asked) == 1 and len(df) == len(df0) and pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19), u.asked)

s = base()
for sheet in (WEST, NORTH, EAST, SOUTH):
    h = head(s[sheet])
    for i in range(h + 1, h + 16):
        s[sheet][i][2] = {"Upto 1000": "Below 1000 cc", "1000-1500": "1000 to 1500 CC", ">1500": "Above 1500 cc"}.get(s[sheet][i][2], s[sheet][i][2])
df, _, u = build(s)
check("other wordings of the engine sizes (Below 1000 cc / 1000 to 1500 CC / Above 1500 cc) read the same, no question, identical output", u.asked == [] and df.equals(df0), u.asked)

s = base()
r = trow(s[WEST], "Comp", "Upto 1000", "With NCB")
s[WEST][r][3] = "Loyalty bonus"
df, _, u = build(s, [("NCB label", "With NCB")])
check("an unknown NCB label asks", len(u.asked) >= 1 and "NCB label" in u.asked[0], u.asked[:1])

s = base()
h = head(s[WEST])
s[WEST].insert(h + 4, list(s[WEST][h + 1]))
df, _, u = build(s, [("twice", "first")])
check("the same table row twice asks which to keep", len(u.asked) == 1 and "twice" in u.asked[0] and len(df) == len(df0), u.asked)

s = base()
r = trow(s[WEST], "Comp", "Upto 1000", "With NCB")
c = col_of(s[WEST], "Mumbai")
s[WEST][r][c] = "to be advised"
df, _, u = build(s, [("not a rate", "Use 0")])
check("a rate cell holding text asks; 0 is used for it (Mumbai Comp With NCB <=1000: 0 for both groups)",
      len(u.asked) == 1 and pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (0, 0) and pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (0, 0), u.asked)
s[WEST][r][c] = 0
df, _, u = build(s)
check("a zero cell stays zero for the top-variant group too (no extra added to nothing)", pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (0, 0) and u.asked == [])
s[WEST][r][c] = "NA"
df, _, u = build(s)
check("the word NA in a rate cell is a zero, without a question", pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (0, 0) and u.asked == [])
s[WEST][r][c] = 26
df, _, u = build(s, [("above 1", "Read it as a percent")])
check("26 typed instead of 0.26 asks and, read as a percent, gives 26 (not 2,600)", pick(df, "MUMBAI", "1+1", OTH, 0, "YES") == (26, 19), u.asked)

s = base()
del s[STP]
df, _, u = build(s, [("without STP", "Carry on")])
check("no STP sheet: asks, and carries on with Comp / SAOD / New only", len(u.asked) == 1 and (df.sub_product != "STP").all() and len(df) == len(df0) - 1770, u.asked)
e = raises(lambda: build(s, interactive=False), mod.GridError)
check("... and stops without a keyboard", e is not None and "STP" in str(e), e)

s = base()
del s[WEST], s[NORTH], s[EAST], s[SOUTH]
e = raises(lambda: build(s), mod.GridError)
check("no Comp / SAOD sheet at all stops with a clear message", e is not None and "no Private Car" in str(e), e)

s = base()
h = head(s[EAST])
s[EAST][h][3] = "Status"
df, _, u = build(s, [("looks like a rate sheet", "Leave")])
check("a sheet that looks like a rate sheet but lost its header asks to leave it out; the other regions are written",
      len(u.asked) == 1 and not df.new_rto_list.isin([P + "PATNA"]).any() and df.new_rto_list.isin([P + "MUMBAI"]).any() and df.new_rto_list.isin([P + "LUCKNOW"]).any(), u.asked)

df, _, u = build(base(), hidden=[SOUTH])
check("a hidden sheet is not read", not df.new_rto_list.isin([P + "CHENNAI"]).any() and df.new_rto_list.isin([P + "MUMBAI"]).any() and u.asked == [])

s = base()
s[WEST] = [r for r in s[WEST] if r[2] != "1000-1500"]
df, _, u = build(s, [("do not join", "Carry on")])
check("engine-size bands that do not join up (the middle band is gone from West) ask whether to carry on",
      len(u.asked) == 1 and "do not join" in u.asked[0] and len(df) < len(df0), u.asked)
e = raises(lambda: build(s, [("do not join", "Stop")]), mod.GridError)
check("... and 'stop' stops", e is not None and "do not join" in str(e), e)

s = base()
h = head(s[WEST])
j = col_of(s[WEST], "RAJKOT")
s[WEST][h][j] = None
df, warns, _ = build(s)
check("a column with rates but no cluster name is left out with a warning (its rates are NOT in the output)", any("no cluster name" in w for w in warns) and not df.new_rto_list.isin([P + "RAJKOT"]).any(), warns)
s = base()
h = head(s[WEST])
s[WEST][h][-1] = "NEW CLUSTER"
for i in range(h + 1, h + 16):
    s[WEST][i][-1] = None
df, warns, _ = build(s)
check("a named column without rates is left out with a warning", any("no rates" in w for w in warns), warns)

s = base()
h = head(s[WEST])
j = col_of(s[WEST], "Pune")
s[WEST][h][j] = "Mumbai"
mod.read_workbook = lambda p: FakeWB(s)
user = Smart([])
out, rep = mod.build("fake.xlsx", F, T, quiet, mod.Asker(True, user.inp, user.out))
dd = mod.report(out, rep, quiet, mod.Asker(True, user.inp, user.out))
check("... the repeated clusters are settled when the file is reported: nothing is asked, the lower payin stays, no duplicate key is left", int(dd.duplicated(mod.KEY_COLUMNS).sum()) == 0 and not user.asked
      and int(out.duplicated(mod.KEY_COLUMNS).sum()) > 0, user.asked)

s = base()
for row in s[STP]:
    if row and row[2] == "DIESEL":
        row[2] = "Gasoline"
df, _, u = build(s, [("which fuels the STP columns", "Diesel only")])
check("an unknown fuel heading in the STP sheet asks which fuels it is for", len(u.asked) == 1 and "fuel" in u.asked[0].lower(), u.asked)
s = base()
for row in s[STP]:
    if row and row[2] == "DIESEL":
        row[2] = "Diesel"
    if row and row[5] == "Non Diesel":
        row[5] = "Petrol & CNG"
df, warns, u = build(s)
check("an STP fuel heading naming a fuel ('Petrol & CNG' -> Petrol and CNG only) covers just those; the uncovered fuels are warned about",
      set(df[df.sub_product == "STP"].fuel_type) == {"Diesel", "Petrol", "CNG"} and any("no columns for" in w for w in warns), warns)

e = raises(lambda: mod.build("fake.xlsx", T, F, quiet, mod.Asker(False)), mod.GridError)
check("to_date before from_date stops", e is not None and "before from_date" in str(e), e)

old = mod.RTO_PREFIX
mod.RTO_PREFIX = "ROYAL_NOV26_"
try:
    df, _, _ = build(base())
    check("the prefix setting changes every new_rto_list (comp and STP)", df.new_rto_list.str.startswith("ROYAL_NOV26_PVT_").all() and df.new_rto_list.nunique() == 237)
finally:
    mod.RTO_PREFIX = old

# ================================================================ September: the "Top Variant Models Drive" block
SEP_REF = os.path.join(helpers.REFS, "rsa_pvtcar_sep26_rates_export.csv")
SEP = helpers.pvtcar_base("sep")


def sep():
    return copy.deepcopy(SEP)


def cell_equals(rows, text):
    """(row, column) of the first cell that is exactly `text` (so '2% additional' is not found inside 'Base + 2% additional')"""
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if isinstance(v, str) and v.strip() == text:
                return i, j
    raise KeyError(text)


def distinct_rates(d):
    """{(cluster, sub_product, is_ncb, fuel, make group): {(pi, po)}} - the September upload has no engine-size bands, so rates are compared per key"""
    out = {}
    pi = pd.to_numeric(d["pi_od_rate"]) + pd.to_numeric(d["pi_tp_rate"])
    po = pd.to_numeric(d["po_od_rate"]) + pd.to_numeric(d["po_tp_rate"])
    for k, a, b in zip(zip(d.new_rto_list, d.sub_product, d.is_ncb, d.fuel_type, d.new_vehicle_makes), pi, po):
        out.setdefault(k, set()).add((round(float(a), 2), round(float(b), 2)))
    return out


dfs, warns_s, us = build(sep())
check("September grid builds with no question and no warning: 19,620 rows, nothing flagged (the Drive block is read, its other lines are standing notes)",
      us.asked == [] and warns_s == [] and len(dfs) == 19620 and (dfs.remarks == "").all(), (us.asked, warns_s, len(dfs), int((dfs.remarks != "").sum())))
mine_s, ref_s = distinct_rates(dfs), distinct_rates(pd.read_csv(SEP_REF, dtype=str, keep_default_na=False))
common = set(mine_s) & set(ref_s)
check("September: the same 6,540 (cluster, sub_product, NCB, fuel, make group) keys as your September upload, and every rate of it is reproduced (0 differences, STP rounding included)",
      set(mine_s) - set(ref_s) == set() and len(common) == 6540 and sum(1 for k in common if mine_s[k] != ref_s[k]) == 0,
      (len(common), len(set(mine_s) - set(ref_s)), [(k, mine_s[k], ref_s[k]) for k in common if mine_s[k] != ref_s[k]][:2]))
check("September: the top-variant extra is 4 points With NCB and 2 points Without NCB (the 'Base payout' lines), for Comp and SAOD; New rows get no extra",
      pick(dfs, "MUMBAI", "1+1", TOP, 0, "YES")[0] - pick(dfs, "MUMBAI", "1+1", OTH, 0, "YES")[0] == 4
      and pick(dfs, "MUMBAI", "1+1", TOP, 0, "NO")[0] - pick(dfs, "MUMBAI", "1+1", OTH, 0, "NO")[0] == 2
      and pick(dfs, "MUMBAI", "SAOD", TOP, 0, "YES")[0] - pick(dfs, "MUMBAI", "SAOD", OTH, 0, "YES")[0] == 4
      and pick(dfs, "MUMBAI", "1+3", TOP, 0, "NA") == (30, 23) and pick(dfs, "MUMBAI", "1+3", OTH, 0, "NA") == (28, 21), "")
check("September Mumbai Comp (your upload: With NCB 31 / 27, Without NCB 28 / 26)",
      pick(dfs, "MUMBAI", "1+1", TOP, 0, "YES") == (31, 24) and pick(dfs, "MUMBAI", "1+1", OTH, 0, "YES") == (27, 20)
      and pick(dfs, "MUMBAI", "1+1", TOP, 0, "NO") == (28, 21) and pick(dfs, "MUMBAI", "1+1", OTH, 0, "NO") == (26, 19))

s = sep()
i, j = cell_equals(s[WEST], "4% additional")
s[WEST][i][j] = "5% additional"
df, _, _ = build(s)
check("the Drive block is read each month: '5% additional' With NCB -> +5 in that sheet only (Mumbai 32), the other sheets still +4, Without NCB still +2",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (32, 25) and pick(df, "MUMBAI", "1+1", TOP, 0, "NO") == (28, 21)
      and pick(df, "LUCKNOW", "1+1", TOP, 0, "YES") == pick(dfs, "LUCKNOW", "1+1", TOP, 0, "YES") and (df.remarks == "").all())
s = sep()
s[WEST][i][j] = "four percent additional"
df, _, _ = build(s)
check("'four percent additional' is read as 4: identical output", df.equals(dfs))

s = sep()
a, b = cell_equals(s[WEST], "2% additional")
s[WEST][a][b] = None
df, _, _ = build(s)
fl = df[df.remarks.str.contains("Top Variant Models' note was not found", regex=False)]
check("the Without-NCB extra missing from West: its Without-NCB top-variant Comp / SAOD rows get no extra and are flagged (3 + 3 table rows x 17 clusters x 5 fuels); With NCB is still +4",
      pick(df, "MUMBAI", "1+1", TOP, 0, "NO") == pick(df, "MUMBAI", "1+1", OTH, 0, "NO") and pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (31, 24)
      and len(fl) == 6 * 17 * 5 and set(fl.is_ncb) == {"NO"} and set(fl.new_vehicle_makes) == {TOP} and fl.remarks.str.contains("is_ncb NO only", regex=False).all(), len(fl))

s = sep()
add_note(s[WEST], "5. Additional 2% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda")
df, _, _ = build(s)
fl = df[df.remarks.str.contains("two notes give different values", regex=False)]
check("a flat 'Additional 2%' note next to a Drive block that says 4 / 2: they disagree, so neither is used and the whole West sheet (2,550 rows) is flagged",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == pick(df, "MUMBAI", "1+1", OTH, 0, "YES") and len(fl) == 17 * 30 * 5 and not fl.new_rto_list.isin([P + "LUCKNOW"]).any(), len(fl))
s = sep()
add_note(s[WEST], "5. Additional 2% payout for the Top Variant Models of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia, MG and Skoda")
s[WEST][i][j] = "2% additional"
df, _, _ = build(s)
check("... but when the flat note and the Drive block say the same (2 and 2 / 2) there is no conflict: +2 for both NCB, nothing flagged",
      pick(df, "MUMBAI", "1+1", TOP, 0, "YES")[0] - pick(df, "MUMBAI", "1+1", OTH, 0, "YES")[0] == 2 and (df.remarks == "").all(), int((df.remarks != "").sum()))

s = sep()
a, b = note_cell(s[WEST], "50% business share from Top variant models")
s[WEST][a][b] = "75% business share from Top variant models."
df, _, _ = build(s)
fl = df[df.remarks != ""]
check("a business-share step with another percentage (75%) is a new line: flagged on the whole West sheet; the rates do not move",
      len(fl) == 17 * 30 * 5 and fl.remarks.str.contains("75% business share", regex=False).all() and not fl.new_rto_list.isin([P + "LUCKNOW"]).any()
      and pick(df, "MUMBAI", "1+1", TOP, 0, "YES") == (31, 24), len(fl))
s = sep()
a, b = note_cell(s[WEST], "Eligible Makes")
s[WEST][a][b] = "Eligible Makes - Maruti, Tata, Mahindra, Hyundai, Toyota and MG."
df, _, _ = build(s)
fl = df[df.remarks != ""]
check("'Eligible Makes' without Kia and Skoda: the top-variant rows of that sheet (15 table rows x 17 clusters x 5 fuels) are flagged to check the make group",
      len(fl) == 15 * 17 * 5 and set(fl.new_vehicle_makes) == {TOP} and fl.remarks.str.contains("names the makes", regex=False).all(), len(fl))
s = sep()
a, b = note_cell(s[WEST], "Eligible Makes")
s[WEST].append([None, "A new line: Honda variants are eligible as well from 15 Sep."] + [None] * (len(s[WEST][0]) - 2))
df, _, _ = build(s)
fl = df[df.remarks != ""]
check("a new line of text under the Drive block is flagged on the whole sheet", len(fl) == 17 * 30 * 5 and fl.remarks.str.contains("Honda variants", regex=False).all(), len(fl))
s = sep()
a, b = cell_equals(s[WEST], "With NCB:")
s[WEST][a][b] = "With NCB"
df, _, _ = build(s)
check("the heading written without its colon ('With NCB') is read the same", df.equals(dfs))

# ---- rounding: from the full-precision cell, a true half rounds up
s = base()
r = next(i for i, row in enumerate(s[STP]) if row[1] == "REST OF GUJARAT")
s[STP][r][5] = 0.3312502263284594
s[STP][r][7] = 0.09724959438700262
s[STP][r][3] = 0.33125
df, _, _ = build(s)
check("STP rounding follows your upload: 0.3312502 -> 33.13 (PO 26.13), 0.0972496 -> 9.72 (PO 2.72), and a true half 0.33125 -> 33.13 (not Python's 33.12)",
      pick(df, "REST_OF_GUJARAT", "STP", ALL, 0, "NA", "Petrol", True) == (33.13, 26.13) and pick(df, "REST_OF_GUJARAT", "STP", ALL, 1500.01, "NA", "Petrol", True) == (9.72, 2.72)
      and pick(df, "REST_OF_GUJARAT", "STP", ALL, 1000.01, "NA", "Diesel", True) == (33.13, 26.13), "")
s = base()
r = trow(s[WEST], "Comp", "Upto 1000", "Without NCB")
c = col_of(s[WEST], "Mumbai")
s[WEST][r][c] = 0.285
df, _, _ = build(s)
check("0.285 (28.499999999999996 in binary) is the 28.5 it was meant to be; 0.28125 is a true half -> 28.13 (PO 21.13)",
      pick(df, "MUMBAI", "1+1", OTH, 0, "NO") == (28.5, 21.5), pick(df, "MUMBAI", "1+1", OTH, 0, "NO"))
s[WEST][r][c] = 0.28125
df, _, _ = build(s)
check("... 0.28125 -> 28.13", pick(df, "MUMBAI", "1+1", OTH, 0, "NO") == (28.13, 21.13), pick(df, "MUMBAI", "1+1", OTH, 0, "NO"))

# ================================================================ output file
path, user = run_main(base(), ref=REF)
problems = helpers.check_csv(path, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
check("the written file has the 42 columns, numbers only in numeric columns, DD/MM/YYYY dates, remarks only as Manual Review flags", problems == [], problems)
back = pd.read_csv(path, dtype=str, keep_default_na=False)
check("the written file has 19,620 rows and re-reads to the same make groups", len(back) == 19620 and set(back.new_vehicle_makes) == {TOP, OTH, ALL})
check("no question was asked on the way to the file", all("[Y/n]" in q for q in user.asked), user.asked)

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
