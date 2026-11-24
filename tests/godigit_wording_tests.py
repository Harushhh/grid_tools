"""Go Digit: the WORDS of the make / body / segment columns and the parts of the workbook nothing reads (the user, 18 Nov 2026: "at some places there is some make or other conditions which is not
applied, also it is not mentioned anywhere").  Every wording is read by an exact table or flagged; a sheet / column / note nothing reads is listed - never silent.
Run:  .venv\\Scripts\\python.exe tests\\godigit_wording_tests.py"""
import datetime as dt
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("godigit_agency_to_config.py")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
F, T = dt.date(2026, 11, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


mod.MISSING_TABLES = "skip"
HEAD = ["Cluster segment", "Segment", "Body Type", "RTO Cluster", "State", "Make / Variant", "CD1 (OD-disc)"] + ["New", "1 Yr", "2 Yr", "3 Yr", "4 Yr", "5 Yr", "6-10 Yr", "11+ Yr"] * 2
HCV_NOTE = "Note: Volvo, Scania to block. \nOnly for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1"


def book(cv_rows, extra=()):
    sheets = [mod.Sheet("LATEST CV COMBINED GRID", [[], ["", "", "", "", "", "", "", "COMPREHENSIVE"], ["", "", "", "", "", "", "", "CD2"], HEAD] + [list(r) for r in cv_rows]),
              mod.Sheet("CV Remarks", [["REMARKS FOR HCV"], [HCV_NOTE]])]
    sheets.extend(extra)
    return mod.Book(sheets)


def run(b):
    mod.reset_run_lists()
    return mod.build(None, F, T, quiet, mod.Asker(interactive=False), book=b)


def row(seg, bseg, body, make, cd1="0-70%", comp=0.3, satp=0.3, cluster="Delhi"):
    return (seg, bseg, body, cluster, "Delhi", make, cd1) + (comp,) * 8 + (satp,) * 8


def only(df, **kw):
    m = pd.Series(True, index=df.index)
    for k, v in kw.items():
        m &= (df[k] == v)
    return df[m]


# ---------------------------------------------------------------- the readings, one wording at a time
def lcv(make, body="LCV", seg="GCV4 1.6T To 2.5T"):
    return row("CV LCV", seg, body, make)


for make, group in (("Mahindra and AL", "Mahindra,Ashok"), ("All(except Tata,Mahindra and AL)", "GCV_4W_ALL_EX_TATA_MAHINDRA_ASHOK"), ("Non Mahindra", "GCV_4W_ALL_EX_MAHINDRA"),
                    ("TATA Intra", "DIGIT_CV_TATAINTRA"), ("Tata Ace", "DIGIT_CV_TATAACE"), ("Non-Tata", "GCV_4W_ALL_EX_TATA"), ("All make", "GCV_4W_ALL"), ("All", "GCV_4W_ALL"), ("Mahindra", "Mahindra")):
    df, w = run(book([lcv(make)]))
    check(f"LCV make '{make}' -> {group}, read (no flag that says it is not read)", set(df.new_vehicle_makes) == {group} and not df.remarks.str.contains("not read by this script").any(), (set(df.new_vehicle_makes), df.remarks.iloc[0][:150]))

df, w = run(book([lcv("Mahindra and AL"), lcv("All(except Tata,Mahindra and AL)")]))
check("two LCV rows of one cluster no longer collapse into one card (the make words split them): both make groups are written with their own payout", set(df.new_vehicle_makes) == {"Mahindra,Ashok", "GCV_4W_ALL_EX_TATA_MAHINDRA_ASHOK"})

df, w = run(book([lcv("Hindustan Motors only")]))
check("an LCV make no table reads: written for all makes AND flagged on exactly its rows", set(df.new_vehicle_makes) == {"GCV_4W_ALL"} and df.remarks.str.contains("make 'Hindustan Motors only' is not read by this script").all(), df.remarks.iloc[0][:200])
key = ("CV LCV", "GCV4 1.6T To 2.5T", "LCV", "Hindustan Motors only")
check("... and listed in the review sheet 'Wording readings' as NOT APPLIED", key in mod.WORDINGS and mod.WORDINGS[key]["unread"] and (mod.wording_readings_frame().iloc[0]["status"].startswith("NOT APPLIED")), mod.WORDINGS.get(key))

df, w = run(book([row("CV MCV", "GCV4 3.5 To 7.5T", "—", "All make Petrol/CNG")]))
check("MCV 'All make Petrol/CNG': Petrol, CNG, LPG only (the Diesel and Electric rows of the old reading are gone)", set(df.fuel_type) == {"Petrol", "CNG", "LPG"} and not df.remarks.str.contains("not read").any(), set(df.fuel_type))
df, w = run(book([row("CV MCV", "GCV4 7.5 to 12T", "—", "All make CNG")]))
check("MCV 'All make CNG': CNG and LPG (LPG takes the CNG rate)", set(df.fuel_type) == {"CNG", "LPG"}, set(df.fuel_type))
df, w = run(book([row("CV MCV", "GCV4 7.5 to 12T", "Oil Tanker CNG", "All")]))
check("MCV body 'Oil Tanker CNG': class Tanker, T&C Oil Tanker, CNG and LPG", set(df.fuel_type) == {"CNG", "LPG"} and set(df.make_model_class) == {"Tanker"} and df.add_tnc.str.contains("Oil Tanker").all(), (set(df.fuel_type), set(df.make_model_class)))

# ---------------------------------------------------------------- CPA
gc = lambda body: row("CV GCV3(EV/Non-EV)", "GCV3", body, "All make")
for body, want in (("with CPA", "YES"), ("w/o CPA", "NO"), ("without CPA", "NO"), ("—", "NA")):
    df, w = run(book([gc(body)]))
    check(f"GCV3 body '{body}' -> is_cpa {want}", set(df.is_cpa) == {want} and not df.remarks.str.contains("not read").any(), set(df.is_cpa))
pc = lambda body, seg="PCV3W non-diesel", make="All make", cluster="Delhi": row("CV PCV2W(All) 3W Non-electric", seg, body, make, cluster=cluster)
df, w = run(book([pc("with CPA"), pc("without CPA")]))
check("PCV 3W 'with CPA' / 'without CPA' rows of one cluster are two cards (YES / NO), not one lowest-rate card", set(df.is_cpa) == {"YES", "NO"} and len(only(df, product="PCV 3W", sub_product="1+1", is_cpa="YES", fuel_type="Petrol")) == 8)

er = lambda body: row("PCV 3W Electric", "E-Rickshaw", body, "", cd1="0-70%")
df, w = run(book([er("With CPA and No TPPD reduction"), er("With CPA and TPPD reduction"), er("With No CPA and only TPPD reduction"), er("With No CPA and No TPPD reduction. Consider all RTOs as Good UK for E-Rickshaw")]))
check("E-Rickshaw: the CPA part of the body text is applied (YES / NO)", set(df.is_cpa) == {"YES", "NO"}, set(df.is_cpa))
check("E-Rickshaw: the TPPD part cannot be applied: flagged on the rows, with no flag lost when the variants collapse into one card",
      df.remarks.str.contains("TPPD reduction' part is not applied").all() and df.remarks.str.contains("Consider all RTOs as Good UK for E-Rickshaw': the sentence").any(), df.remarks.drop_duplicates().tolist()[:3])

# ---------------------------------------------------------------- the 3W makes
df, w = run(book([pc("—", "PCV3W diesel", "Bajaj and TVS only"), pc("—", "PCV3W diesel", "Bajaj RE", ), pc("—", "PCV3W diesel", "Other than Bajaj")]))
m2 = set(only(df, product="PCV 2W").new_vehicle_makes)
m3 = set(only(df, product="PCV 3W").new_vehicle_makes)
check("PCV 3W diesel makes: 'Bajaj and TVS only' = BAJAJ,TVS (2W) / PCV_3W_BAJAJ_TVS (3W); 'Other than Bajaj' = ..._EX_BAJAJ; 'Bajaj RE' only for the three-wheeler",
      m2 == {"BAJAJ,TVS", "PCV_2W_ALL_EX_BAJAJ"} and m3 == {"PCV_3W_BAJAJ_TVS", "PCV_3W_ALL_EX_BAJAJ", "PCV_3W_BAJAJ_RE"}, (m2, m3))
check("... the groups the master lacks are listed with their definition, not guessed (sheet 'Make groups to create')",
      mod.MAKE_DEFINITIONS["PCV_3W_BAJAJ_TVS"][0] == "PCV_3W_ALL" and "PCV_2W_ALL_EX_BAJAJ" in mod.MAKE_DEFINITIONS and "GCV_4W_ALL_EX_MAHINDRA" in mod.MAKE_DEFINITIONS)
df, w = run(book([pc("—", "PCV3W diesel", "All excluding Piaggio")]))
check("'All excluding Piaggio' (October's wording) still gives the _EX_PIAGGAO groups of both products", set(df.new_vehicle_makes) == {"PCV_2W_ALL_EX_PIAGGAO", "PCV_3W_ALL_EX_PIAGGAO"}, set(df.new_vehicle_makes))
df, w = run(book([pc("—", "PCV2 - without ZD", "All", cluster="GOA"), pc("—", "PCV3W non-diesel", "All make", cluster="GOA"), pc("—", "PCV3W non-diesel", "All make", cluster="PUNE")]))
two = only(df, product="PCV 2W")
check("'PCV2 - without ZD' is a two-wheeler row: PCV 2W only, is_zd NO on the comprehensive cards (NA on STP), nothing left unread",
      set(only(two, new_rto_list="GO-DIGIT_OCT26_PCV2W_3W_NON-EV_GOA").sub_product) == {"1+1", "STP"} and set(only(two, sub_product="1+1", new_rto_list="GO-DIGIT_OCT26_PCV2W_3W_NON-EV_GOA").is_zd) == {"NO"}
      and set(only(two, sub_product="STP").is_zd) == {"NA"} and not df.remarks.str.contains("not read|not applied").any(), (set(two.new_rto_list), set(two.is_zd)))
check("a cluster with its own PCV2W row: its PCV3W rows are for PCV 3W only (GOA); a cluster without one keeps both products, as in the October upload (PUNE)",
      set(two.new_rto_list) == {"GO-DIGIT_OCT26_PCV2W_3W_NON-EV_GOA", "GO-DIGIT_OCT26_PCV2W_3W_NON-EV_PUNE"} and set(only(df, product="PCV 3W").new_rto_list) == set(two.new_rto_list)
      and len(only(two, new_rto_list="GO-DIGIT_OCT26_PCV2W_3W_NON-EV_GOA")) == 16 * 4 and len(only(df, product="PCV 3W", new_rto_list="GO-DIGIT_OCT26_PCV2W_3W_NON-EV_GOA")) == 16 * 4, (len(two), len(df)))
df, w = run(book([row("CV Tractor", "Tractor Electric and Low HP", "—", "All make")]))
check("'Tractor Electric and Low HP': flagged (the upload has no horsepower field)", df.remarks.str.contains("electric / low-horsepower condition is not applied").all())

# ---------------------------------------------------------------- taxis
tx = lambda seg, body, make="All", cd1="0-50%": row("CV PCV4(Upto 5 seating)", seg, body, make, cd1=cd1)
df, w = run(book([tx("Taxi upto 5 seater diesel > 1000 cc & <=1500 cc", "With Add on cover", "Maruti & Hyundai"), tx("Taxi upto 5 seater diesel > 1000 cc & <= 1500 cc", "Without Add on cover", "Maruti")]))
check("taxi cc band '> 1000 cc & <=1500 cc' (no space) = 1000.01-1500.01, like the spaced '<= 1500 cc'", set(zip(df.cc_min, df.cc_max)) == {(1000.01, 1500.01)}, set(zip(df.cc_min, df.cc_max)))
df, w = run(book([tx("Taxi upto 5 seater All < 1000 cc", "With Add on cover")]))
check("taxi segment 'All' fuels = every fuel (it said Petrol, CNG, LPG only)", set(df.fuel_type) == {"Petrol", "Diesel", "CNG", "LPG", "Electric"} and not df.remarks.str.contains("not read").any(), set(df.fuel_type))
df, w = run(book([tx("Taxi upto 5 seater", "Diesel")]))
check("taxi body 'Diesel' on a segment with no fuel word = Diesel (it was written for CNG, LPG, Petrol)", set(df.fuel_type) == {"Diesel"}, set(df.fuel_type))
df, w = run(book([tx("Taxi upto 5 seater", "Non diesel/ Electric")]))
check("taxi body 'Non diesel/ Electric' = CNG, LPG, Petrol and Electric", set(df.fuel_type) == {"CNG", "LPG", "Petrol", "Electric"}, set(df.fuel_type))
df, w = run(book([tx("Taxi upto 5 seater sleeper > 1000 cc", "With Add on cover")]))
check("a taxi segment no pattern reads is flagged, written with the default cc / fuels", df.remarks.str.contains("segment 'Taxi upto 5 seater sleeper > 1000 cc' is not read by this script").all())

# ---------------------------------------------------------------- the flags of a dropped duplicate are not lost
base = dict(k="a", _z=False, pi_net_rate=0.0, pi_tp_rate=0.0, pi_od_rate=0.0, product="P", sub_product="1+1", new_rto_list="r", new_vehicle_makes="m", fuel_type="Petrol", make_model_class="c",
            vehicle_age_min=0, vehicle_age_max=1, cc_min=0, cc_max=1, tariff_max=70.01)
two = pd.DataFrame([dict(base, _flags=("flag one",), pi_net_rate=10.0, _r=1), dict(base, _flags=("flag two",), pi_net_rate=20.0, _r=2)])
out = mod.settle_duplicates(two, ["k"], [])
check("settle_duplicates: the row that stays (the lowest payin) carries the flags of the row dropped for it", len(out) == 1 and out.iloc[0]["pi_net_rate"] == 10.0 and set(out.iloc[0]["_flags"]) == {"flag one", "flag two"}, out[["pi_net_rate", "_flags"]])

# ---------------------------------------------------------------- the real November workbook: every wording read or flagged, every unread part named
xlsx = os.path.join(ROOT, "regression", "inputs", "godigit_nov26.xlsx")
if os.path.exists(xlsx):
    feed = iter(["1"])

    def answer(prompt):
        try:
            return next(feed)
        except StopIteration:
            raise EOFError
    mod.reset_run_lists()
    ask = mod.Asker(interactive=True, input_fn=answer, out=lambda *x: None)
    df, w = mod.build(xlsx, F, T, quiet, ask)
    mod.collect_gaps(df)
    wd = mod.wording_readings_frame()
    notapplied = set(zip(wd[wd.status != "read"].segment, wd[wd.status != "read"].body))
    check("November: every wording of the CV grid's segment / body / make columns has a reading (221 distinct)", len(wd) == 221, len(wd))
    check("November: the wordings nothing applies are exactly the five known ones (Tractor Electric and Low HP, the four E-Rickshaw bodies) - a new one fails here and is flagged on its rows",
          notapplied == {("Tractor Electric and Low HP", "—"), ("E-Rickshaw", "With CPA and No TPPD reduction"), ("E-Rickshaw", "With CPA and TPPD reduction"),
                         ("E-Rickshaw", "With No CPA and No TPPD reduction. Consider all RTOs as Good UK for E-Rickshaw"), ("E-Rickshaw", "With No CPA and only TPPD reduction")}, notapplied)
    flagged_all = all(df.remarks.str.contains(t.split(":")[0].split("'")[1], regex=False).any() for t in wd[wd.status != "read"].not_applied.str.split(" \\| ").str[0])
    check("November: each not-applied wording is on the remarks of rows of the upload", flagged_all)
    parts = pd.DataFrame(mod.UNREAD_PARTS)
    sheets_unread = set(parts[parts["where"] == "the whole sheet"].sheet)
    check("November: the sheet '2W Grid 5+5' (9,591 cells, no table of this script) is named as not read", sheets_unread == {"2W Grid 5+5"}, sheets_unread)
    car = parts[(parts.sheet == "New 4W Comp_SAOD ") & (parts.status == "NOT READ")]
    check("November: the car sheet's new-business blocks (columns A:I and Z:AC) and its three notes are named as not read", {"columns A:I", "columns Z:AC", "column T", "column AE"} <= set(car["where"]), set(car["where"]))
    left = parts[parts.status == "left out on purpose"]
    check("November: Min CD1 and the 2nd-4th year SAOD payouts are listed as left out on purpose, with the reason", set(left.holds) == {"Min CD1", "Max 2 year SAOD CD2", "Max 3rd year SAOD CD2", "Max 4th year SAOD CD2"} and (left.why != "").all(), set(left.holds))
    check("November: mapping sheets and the remarks sheet are not reported as unread", not (set(parts.sheet) & {"RTO Mapping", "4W RTO Mapping - New", "4W COMP", "4W SAOD", "2W RTO", "Remarks for CV"}), set(parts.sheet))
    check("the unread parts are warnings of the run too (the Summary of the review workbook)", any("2W Grid 5+5" in x and "not read" in x for x in w) and any("columns A:I" in x for x in w))
    notes = pd.DataFrame(mod.NOTES_READ)
    st = lambda s: notes[notes.text.str.contains(s, regex=False)].status.tolist()
    check("November: every remark sentence is listed with what became of it - HCV note applied, the 'up to 10-seater' note not applied, the short-term policy notes a standing decision",
          st("Volvo, Scania to block") == ["applied"] and st("up to 10-seater") == ["not applied - listed in 'Not written'"] and st("Short-term policies will be 5% lesser") == ["no action (standing decision)"]
          and (notes[notes.status == "no action (standing decision)"].why != "").all(), notes.status.value_counts().to_dict())
    check("the review workbook gets the three new sheets", {"Wording readings", "Unread parts", "Notes read"} <= set(mod.EXTRA_SHEETS), list(mod.EXTRA_SHEETS))
else:
    print("SKIP the November workbook checks (regression/inputs/godigit_nov26.xlsx is missing)")

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
