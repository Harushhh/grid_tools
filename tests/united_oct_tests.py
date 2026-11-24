"""Checks of united_commission_to_config.py on the October 2026 United circular, which arrived as a SCAN (7 pages of pictures, no text, no tables):
the PDF itself must stop with a message that says what to do, and its transcription (regression/inputs/united_oct26.json, the Motor tables read
off the page images) must build the right rows. Also the transcription reader on its own (row / column spans, bad files).
Run:  .venv\\Scripts\\python.exe tests\\united_oct_tests.py"""
import copy
import datetime as dt
import json
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("united_commission_to_config.py")
PDF = os.path.join(helpers.INPUTS, "united_oct26.pdf")
TRANSCRIPTION = os.path.join(helpers.INPUTS, "united_oct26.json")
F, T = dt.date(2026, 10, 26), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
U = mod.STATE_CODES
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def build(path=TRANSCRIPTION, from_date=F, to_date=T):
    return mod.build(path, from_date, to_date, quiet, mod.Asker(False))


def sel(df, product=None, cls=None, sub=None, fuel=None, cc=None, sc=None, rto=None, tnc=None):
    d = df
    for col, val in (("product", product), ("make_model_class", cls), ("sub_product", sub), ("fuel_type", fuel), ("add_tnc", tnc)):
        if val is not None:
            d = d[d[col] == val]
    if cc is not None:
        d = d[(d.cc_min.astype(float) == cc[0]) & (d.cc_max.astype(float) == cc[1])]
    if sc is not None:
        d = d[(d.sc_min.astype(float) == sc[0]) & (d.sc_max.astype(float) == sc[1])]
    if rto is not None:
        want = set(rto.split(",")) if re.fullmatch(r"[A-Z]{2}(,[A-Z]{2})*", rto) else None
        d = d[d.new_rto_list.map(lambda s: (set(s.split(",")) == want) if want else s == rto)]
    return d


def rate(d):
    out = set()
    for _, r in d.iterrows():
        for pi, po in (("pi_net_rate", "po_net_rate"), ("pi_tp_rate", "po_tp_rate"), ("pi_od_rate", "po_od_rate")):
            if float(r[pi]) != 0 or float(r[po]) != 0:
                out.add((float(r[pi]), float(r[po])))
                break
        else:
            out.add((0.0, 0.0))
    return out


def comp(*names):
    gone = {n for x in names for n in x.split(",")}
    return ",".join(c for c in U if c not in gone)


def error_of(fn):
    try:
        fn()
    except mod.GridError as e:
        return str(e)
    return None


# ================================================================ the scan itself
import shutil  # noqa: E402

tbl = mod.read_pdf(PDF)
check("the scanned October PDF is read through its transcription: found by the SHA-256 of the PDF (source_sha256 in regression/inputs/united_oct26.json), 6 tables, and the TableList says where they came from",
      tbl.transcribed_from is not None and os.path.basename(tbl.transcribed_from) == "united_oct26.json" and len(tbl) == 6, tbl.transcribed_from)
lg = []
dfp, wp = mod.build(PDF, F, T, lg.append, mod.Asker(False))
dfj, wj = build()
check("building from the PDF itself (what the user runs) gives exactly the rows of its transcription and says in the log that it used the transcription",
      dfp.equals(dfj) and any("This PDF is a scan: using its transcription" in str(x) for x in lg), len(dfp))
other = os.path.join(tempfile.mkdtemp(prefix="united_oct_"), "scan_copy.pdf")
shutil.copyfile(PDF, other)
check("a renamed copy of the same scan is still recognised (it is the file's bytes that count, not its name)", mod.find_transcription(other) == mod.find_transcription(PDF) is not None)
with open(other, "ab") as fh:
    fh.write(b"%% another scan %%" + bytes([10]))
msg = error_of(lambda: mod.read_pdf(other))
check("a DIFFERENT scan (other bytes) with no transcription stops with a message that says it is a scan, what to do, and that a transcription can be given instead (or found by its SHA-256)",
      msg is not None and "this PDF is a scan" in msg and "source_sha256" in msg and "7 of its pages" in msg and mod.find_transcription(other) is None, msg)

# ================================================================ the transcription reader
def transcribe(doc):
    d = tempfile.mkdtemp(prefix="united_oct_")
    p = os.path.join(d, "t.json")
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(doc, fh)
    return p


tb = mod.read_pdf(transcribe({"tables": [{"page": 9, "columns": 3, "rows": [[{"t": "title", "cs": 3}], [{"t": "A", "rs": 2}, "b1", "c1"], ["b2", "c2"], ["A2", "b3", {"t": "c3", "rs": 1}]]}]}))[0]
by = {c.text: c for c in tb.cells}
check("transcription reader: a rowspan keeps its column for the rows below (b2 / c2 sit in columns 1 and 2 of row 3), a colspan covers the width, boxes are 100 x 20, the page number is kept",
      (by["A"].x0, by["A"].x1, by["A"].top, by["A"].bottom) == (0, 100, 20, 60) and (by["b2"].x0, by["b2"].top) == (100, 40) and (by["c2"].x0, by["c2"].top) == (200, 40)
      and (by["title"].x0, by["title"].x1) == (0, 300) and (by["A2"].x0, by["A2"].top) == (0, 60) and tb.page == 9 and tb.bbox == (0.0, 0.0, 300.0, 80.0), by)
for what, doc in (("a cell running past the declared columns", {"tables": [{"columns": 2, "rows": [["a", "b", "c"]]}]}),
                  ("a table without 'columns'", {"tables": [{"rows": [["a"]]}]}), ("no tables", {"tables": []}), ("a cell that is not text", {"tables": [{"columns": 1, "rows": [[5]]}]})):
    m = error_of(lambda doc=doc: mod.read_pdf(transcribe(doc)))
    check(f"transcription reader: {what} -> a clean stop that names the transcription, no traceback", m is not None and ("transcription" in m), m)
d0 = tempfile.mkdtemp(prefix="united_oct_")
bad = os.path.join(d0, "bad.json")
with open(bad, "w") as fh:
    fh.write("{not json")
m = error_of(lambda: mod.read_pdf(bad))
check("transcription reader: a file that is not JSON -> a clean stop", m is not None and "cannot read the transcription" in m, m)

# ================================================================ the October circular
df, warns = build()
CAR, G4, G3 = df[df["product"] == "Private Car"], df[df["product"] == "GCV 4W"], df[df["product"] == "GCV 3W"]
check("the October transcription builds: 496 rows, 42 columns, no duplicate key, every row from 26/10/2026 to 31/03/2030 (the dates the user typed)",
      len(df) == 496 and list(df.columns) == mod.OUTPUT_COLUMNS and not df.duplicated(mod.KEY_COLUMNS).any() and set(df.from_date) == {"26/10/2026"} and set(df.to_date) == {"31/03/2030"}, len(df))
check("the 'Revised Pay-out (w.e.f. 26-10-2026)' column is the new one: no flag about a different effective date",
      not df.remarks.str.contains("effective from", regex=False).any())

# GCV: changed in October (22.5 -> 27.5, 15 -> 25) and unchanged
check("GCV 7500-10000 kg: other states 22.5 -> 27.5; Tamil Nadu + Pondicherry 5, Rajasthan 15 unchanged",
      rate(sel(G4, cc=(7500.01, 10000.01), rto=comp("PY,TN", "RJ"))) == {(27.5, 20.5)} and rate(sel(G4, cc=(7500.01, 10000.01), rto="PY,TN")) == {(5.0, 0.0)}
      and rate(sel(G4, cc=(7500.01, 10000.01), rto="RJ")) == {(15.0, 8.0)})
check("GCV 20000-25000 kg: other states 15 -> 25 (PO 18); the 8-state list stays 5",
      rate(sel(G4, cc=(20000.01, 25000.01), rto=comp("HR", "KA", "KL", "MP", "PY", "RJ", "TN", "UP"))) == {(25.0, 18.0)}
      and rate(sel(G4, cc=(20000.01, 25000.01), rto="HR,KA,KL,MP,PY,RJ,TN,UP")) == {(5.0, 0.0)})
check("GCV <= 2000: Uttar Pradesh 50 / 43 and 'Other than above State' 60 / 53; > 40000: 0 for the 8 states, 2.5 elsewhere; E-Cart 50 on GCV 3W",
      rate(sel(G4, cc=(0, 2000.01), rto="UP")) == {(50.0, 43.0)} and rate(sel(G4, cc=(0, 2000.01), rto=comp("UP"))) == {(60.0, 53.0)}
      and rate(sel(G4, cc=(40000.01, 99999), rto="HR,KA,KL,MP,PY,RJ,TN,UP")) == {(0.0, 0.0)} and rate(sel(G4, cc=(40000.01, 99999), rto=comp("HR", "KA", "KL", "MP", "PY", "RJ", "TN", "UP"))) == {(2.5, 0.0)}
      and rate(G3) == {(50.0, 43.0)})
b = (2000.01, 3500.01)
check("GCV 2000-3500: Kerala + Karnataka 45 and the other states 56.5 as before; the MP / TN / UP / HR 15 and the Rajasthan 35 payouts say '(Except RTOs listed in Sub Annexure 2 ...)' but this PDF "
      "has no Sub Annexure 2 - an annexure that is not in the document is NOT applied (user's rule, 26 Oct 2026; nothing is carried over from August), so they are written for the whole states, "
      "10 rows each, and flagged on exactly those 20 rows",
      rate(sel(G4, cc=b, rto="KA,KL")) == {(45.0, 38.0)} and rate(sel(G4, cc=b, rto=comp("HR", "MP", "TN", "UP", "RJ", "KA", "KL"))) == {(56.5, 49.5)}
      and rate(sel(G4, cc=b, rto="HR,MP,TN,UP")) == {(15.0, 8.0)} and len(sel(G4, cc=b, rto="HR,MP,TN,UP")) == 10 and rate(sel(G4, cc=b, rto="RJ")) == {(35.0, 28.0)} and len(sel(G4, cc=b, rto="RJ")) == 10
      and not sel(G4, cc=b).new_rto_list.str.contains("UNITED_", regex=False).any()
      and set(G4[G4.remarks.str.contains("not in this document, so it is not applied", regex=False)].index) == set(sel(G4, cc=b, rto="HR,MP,TN,UP").index) | set(sel(G4, cc=b, rto="RJ").index)
      and len(G4[G4.remarks.str.contains("not in this document, so it is not applied", regex=False)]) == 20 and not G3.remarks.ne("").any()
      and any("applied without its annexure" in w for w in warns))
check("3500-7500: Haryana, MP, Pondicherry, Rajasthan, TN, UP 10 / 3 (the 'Except RTOs listed in Sub Annexure 3' is only in the Existing column), other states 27.5 / 20.5",
      rate(sel(G4, cc=(3500.01, 7500.01), rto="HR,MP,PY,RJ,TN,UP")) == {(10.0, 3.0)} and rate(sel(G4, cc=(3500.01, 7500.01), rto=comp("HR", "MP", "PY", "RJ", "TN", "UP"))) == {(27.5, 20.5)})

# passenger carrying: the new wording
bus = df[(df["product"] == "PCV 4W") & (df.make_model_class == "Other Bus")]
check("'Seating Capacity <= 10 / 10 < .. <= 20 / 20 < .. <= 30 / > 30 / > 20 (Electric)' is read like the old 'PCC' bands: sc bands 6.01-10.01, 10.01-20.01, 20.01-30.01, 30.01-999, and Electric 20.01-999",
      set(zip(bus.sc_min.astype(float), bus.sc_max.astype(float))) == {(6.01, 10.01), (10.01, 20.01), (20.01, 30.01), (30.01, 999.0), (20.01, 999.0)}
      and rate(sel(bus, sc=(6.01, 10.01), rto="HR,RJ")) == {(7.5, 0.5)} and rate(sel(bus, sc=(20.01, 999), fuel="Electric")) == {(10.0, 3.0)})
taxi = sel(df, product="PCV 4W", cls="Taxi")
check("Taxis: one row 'Both Package policy & Stand-Alone Third Party policy (SATP)' gives 1+1 AND STP: 15 -> 10 in the 7 states, other states 25 -> 35",
      set(taxi.sub_product) == {"1+1", "STP"} and rate(sel(taxi, rto="HR,KA,MP,PY,RJ,TN,UP")) == {(10.0, 3.0)} and rate(sel(taxi, rto=comp("HR", "KA", "MP", "PY", "RJ", "TN", "UP"))) == {(35.0, 28.0)}
      and len(taxi) == 20)
check("'Two Wheeler (PCV)' (the class was 'Two Wheeled (PCV)' before) is read: PCV 2W 10 / 3 on ALLINDIA, 1+1 and STP",
      rate(sel(df, product="PCV 2W")) == {(10.0, 3.0)} and set(sel(df, product="PCV 2W").sub_product) == {"1+1", "STP"} and len(sel(df, product="PCV 2W")) == 10)

# Private Car: the October table (existing + revised commission, policy types merged)
r1 = sel(CAR, sub="1+1")
check("Private Car, the first policy group (Bundled 1+3 / Package / SATP, non-electric) gives 1+3, 1+1 and STP the SAME revised rates: <=1000 17.5; 1000-1200 KA, HR, MP, TN, PY 17.5 and others 35; "
      "1200-1500 15; 1500-2000 35; >2000 the 8 states 15 and others 30 (the Existing column - 20, 20 / 40, 20, 40 - is ignored)",
      all(rate(sel(CAR, sub=s, cc=(0, 1000.01), rto="ALLINDIA", fuel="Petrol")) == {(17.5, 10.5)} and rate(sel(CAR, sub=s, cc=(1000.01, 1200.01), rto="HR,KA,MP,PY,TN", fuel="Petrol")) == {(17.5, 10.5)}
          and rate(sel(CAR, sub=s, cc=(1000.01, 1200.01), rto=comp("HR", "KA", "MP", "PY", "TN"), fuel="Petrol")) == {(35.0, 28.0)}
          and rate(sel(CAR, sub=s, cc=(1200.01, 1500.01), rto="ALLINDIA", fuel="Petrol")) == {(15.0, 8.0)} and rate(sel(CAR, sub=s, cc=(1500.01, 2000.01), rto="ALLINDIA", fuel="Petrol")) == {(35.0, 28.0)}
          and rate(sel(CAR, sub=s, cc=(2000.01, 99999), rto="CG,HR,KA,MP,PY,RJ,TN,UP", fuel="Petrol")) == {(15.0, 8.0)}
          and rate(sel(CAR, sub=s, cc=(2000.01, 99999), rto=comp("CG", "HR", "KA", "MP", "PY", "RJ", "TN", "UP"), fuel="Petrol")) == {(30.0, 23.0)} for s in ("1+3", "1+1", "STP")))
check("'Non Electric & Electric - SAOD' is written for all five fuels at 25 on OD (it was 15 / Electric 35); Electric Cars (Bundled / Package / SATP) 35 for 1+3, 1+1 and STP - no Electric SAOD row of its own",
      len(sel(CAR, sub="SAOD")) == 5 and set(sel(CAR, sub="SAOD").fuel_type) == {"Petrol", "Diesel", "CNG", "LPG", "Electric"} and rate(sel(CAR, sub="SAOD")) == {(25.0, 18.0)}
      and (sel(CAR, sub="SAOD").pi_od_rate.astype(float) == 25).all()
      and all(rate(sel(CAR, sub=s, fuel="Electric")) == {(35.0, 28.0)} and len(sel(CAR, sub=s, fuel="Electric")) == 1 for s in ("1+3", "1+1", "STP")))
check("the October PDF carries no 40% annexure (and its cover letter has no 40% rule): no annexure group is written for Private Car, the 1+1 rows stay on the plain scope (no EX group), and every "
      "Private Car row says the annexure was not found",
      not CAR.new_rto_list.str.contains("UNITED_").any() and CAR.remarks.str.contains("the annexure of the 40% RTOs was not found", regex=False).all() and len(CAR) == 92
      and any("no portal RTO group is known for the circular effective 26/10/2026" in w for w in warns))
check("3+3 (Long Term Package) and Long term Liability: not written, flagged on every Private Car row, warned about once each",
      CAR.remarks.str.contains("Long Term Package (3+3) has no sub_product", regex=False).all() and CAR.remarks.str.contains("Long term Liability has no sub_product", regex=False).all()
      and sum("Long Term Package (3+3)" in w for w in warns) == 1)
check("TW (unchanged), Misc, Motor Trade / Standalone CPA behave as before: 26 TW rows, TW 5+5 / Motor Trade / CPA flagged and not written",
      len(df[df["product"] == "TW"]) == 26 and not df["product"].eq("MISCD").sum() == 0 and any("Motor Trade" in w for w in warns) and any("Long Term Package (5+5)" in w for w in warns))

# a change to the transcription is handled like a change to a PDF
doc = json.load(open(TRANSCRIPTION, encoding="utf-8"))
d2 = copy.deepcopy(doc)
for t in d2["tables"]:
    for row in t["rows"]:
        for i, cell in enumerate(row):
            if cell == "Taxis" or (isinstance(cell, dict) and cell["t"] == "Taxis"):
                row[i] = {"t": "Cabs and shuttles", "rs": 2}
df2, w2 = build(transcribe(d2))
check("a class label the script does not know ('Cabs and shuttles'): that payout is not written (-20 rows), a warning says so, nothing crashes",
      len(df2) == len(df) - 20 and any("Cabs and shuttles" in w for w in w2))

bad = [n for n, ok, _ in results if not ok]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
