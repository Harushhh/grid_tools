"""Deterministic checks of united_commission_to_config.py on the real July 2026 United PDF (read once, held in memory) with targeted changes to the
DOCUMENT (cell texts, headers, annexure tables, text outside the tables): the rates, the state lists, the flags and the questions asked.
Run:  .venv\\Scripts\\python.exe tests\\united_tests.py"""
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

mod = helpers.load_script("united_commission_to_config.py")
PDF = os.path.join(helpers.INPUTS, "united_jul26.pdf")
REF = os.path.join(helpers.REFS, "united_jul26_rates_export.csv")
F, T = dt.date(2026, 7, 1), dt.date(2026, 7, 31)
quiet = lambda *a, **k: None
TMP = tempfile.mkdtemp(prefix="united_t_")
counter = itertools.count()
TABLES = mod.read_pdf(PDF)
U = mod.STATE_CODES


class Unexpected(Exception):
    pass


class Smart:
    """Answers questions from rules [(question substring, menu label substring | 'y' | 'n')]; a menu that matches no rule raises Unexpected."""

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
            if q in self.question and prompt.startswith("   Type 1-"):
                for i, lab in self.labels.items():
                    if a.lower() in lab.lower():
                        return str(i)
                raise Unexpected(f"rule {q!r}: no option containing {a!r} in {self.labels}")
        raise Unexpected(f"unexpected question: {self.question!r} {list(self.labels.values())}")


def tables():
    return copy.deepcopy(TABLES)


def build(tbls, rules=(), interactive=True, from_date=F, to_date=T):
    user = Smart(rules)
    ask = mod.Asker(True, user.inp, user.out) if interactive else mod.Asker(False)
    df, warnings = mod.build("fake.pdf", from_date, to_date, quiet, ask, tables=tbls)
    return df, warnings, user


def find(tbls, page, text, x_min=None, x_max=None, exact=True):
    hits = [c for t in tbls if t.page == page for c in t.cells
            if (c.text == text if exact else text in c.text) and (x_min is None or c.x0 >= x_min) and (x_max is None or c.x1 <= x_max)]
    assert hits, f"no cell {text!r} on page {page}"
    return hits[0]


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
        want = set(rto.split(",")) if isinstance(rto, str) and re.fullmatch(r"[A-Z]{2}(,[A-Z]{2})*", rto) else None
        d = d[d.new_rto_list.map(lambda s: (set(s.split(",")) == want) if want else s == rto)]
    return d


def one(df, **kw):
    d = sel(df, **kw)
    assert len(d) == 1, f"expected one row, got {len(d)} for {kw}"
    return d.iloc[0]


def rate(r):
    """(pi, po) of whichever rate column the row uses"""
    for pi, po in (("pi_net_rate", "po_net_rate"), ("pi_tp_rate", "po_tp_rate"), ("pi_od_rate", "po_od_rate")):
        if float(r[pi]) != 0 or float(r[po]) != 0:
            return float(r[pi]), float(r[po])
    return 0.0, 0.0


def comp(*names):
    return ",".join(c for c in U if c not in set(n for x in names for n in x.split(",")))


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


# ================================================================ the July PDF
df0, warn0, user0 = build(tables())
check("the July PDF builds with no question", user0.asked == [], user0.asked)
check("457 rows (404 + 53 Private Car), 42 columns, no duplicate key", len(df0) == 457 and list(df0.columns) == mod.OUTPUT_COLUMNS and not df0.duplicated(mod.KEY_COLUMNS).any(), len(df0))
d = mod.diff_against_reference(df0, REF, quiet)
check("against the July upload: 367 keys match, the 5 Annexure 4 rows differ only in from_date (01/07/2026 in the PDF, 01-04-2026 in the upload); 90 rows only here, 43 only there (explained below)", d == (90, 43, 5), d)
check("9 warnings: Sub Annexure 2, the scanned cover page (page 1, not read), the 4 payouts not written (5+5 x2, Motor Trade, CPA) and the 3 Private Car 3+3 payouts",
      len(warn0) == 9 and sum("page 1 has no table" in w for w in warn0) == 1, warn0)
check("every row: id 1, user_id empty, DD/MM/YYYY dates, age 0-99, tariff 0-100.01, is_ncb / is_cpa / is_zd NA, insurer United",
      (df0.id == 1).all() and (df0.user_id == "").all() and (df0[df0["product"] != "Private Car"].from_date == "01/07/2026").all() and (df0.to_date == "31/07/2026").all() and (df0.vehicle_age_max == 99).all()
      and (df0.tariff_max == 100.01).all() and (df0.is_ncb == "NA").all() and (df0.insurance_company == "UNITED INDIA INSURANCE CO LTD").all())
pi = df0[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
po = df0[["po_od_rate", "po_tp_rate", "po_net_rate"]].max(axis=1)
check("po = pi - 7, never below 0, zero stays zero", ((po - (pi - 7).clip(lower=0)).abs() < 0.0051).all() and (po[pi == 0] == 0).all())
check("1+1 is On Net (pi_net), STP is On TP (pi_tp), SAOD is On OD, the one 'OD - 10%, TP - 0%' row is On OD and TP",
      (df0[df0.sub_product == "1+1"].pi_type.isin(["On Net", "On OD and TP"])).all() and (df0[df0.sub_product == "STP"].pi_type == "On TP").all()
      and (df0[df0.sub_product == "SAOD"].pi_type == "On OD").all() and int((df0.pi_type == "On OD and TP").sum()) == 4)

# ---- goods carrying
G = dict(product="GCV 4W", cls="NA")
check("GCV a. GVW <= 2000: All States -> ALLINDIA 57.5 / 50.5, 1+1 and STP, five fuels",
      rate(one(df0, **G, sub="1+1", fuel="Diesel", cc=(0, 2000.01), rto="ALLINDIA")) == (57.5, 50.5) and rate(one(df0, **G, sub="STP", fuel="CNG", cc=(0, 2000.01), rto="ALLINDIA")) == (57.5, 50.5))
b = (2000.01, 3500.01)
check("GCV b. 2000-3500 kg: 56.5 for every state except MP / TN / UP, 15 (po 8) for their other RTOs, 56.5 for the Sub Annexure 3 preferred RTOs - 1+1 and STP",
      rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=b, rto=comp("MP", "TN", "UP"))) == (56.5, 49.5)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=b, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_EX_UP_TN_MP")) == (15, 8)
      and rate(one(df0, **G, sub="STP", fuel="Petrol", cc=b, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP")) == (56.5, 49.5)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=b, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP")) == (56.5, 49.5))
c = (3500.01, 7500.01)
check("GCV c. 3500-7500 kg: 27.5 elsewhere, 10 for the other RTOs of HR / MP / RJ / TN / UP, 27.5 for the preferred RTOs, and Pondicherry (named in the PDF, no RTOs in the annexure) 10 on its own",
      rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=c, rto=comp("HR", "MP", "RJ", "TN", "PY", "UP"))) == (27.5, 20.5)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=c, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_3.5T-7.5T_EX_HR_MP_RJ_TN_UP")) == (10, 3)
      and rate(one(df0, **G, sub="STP", fuel="Petrol", cc=c, rto="PY")) == (10, 3)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=c, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_3.5T-7.5T_HR_MP_RJ_TN_UP")) == (27.5, 20.5))
check("GCV d. 7500-10000: TN, PY 5 (po 0), Rajasthan 15, every other state 22.5", rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=(7500.01, 10000.01), rto="PY,TN")) == (5, 0)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=(7500.01, 10000.01), rto="RJ")) == (15, 8)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=(7500.01, 10000.01), rto=comp("TN", "PY", "RJ"))) == (22.5, 15.5))
check("GCV j. GVW > 40000: the 8 states 0, the others 5 (po floored to 0); open band cc_max 99999",
      rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=(40000.01, 99999), rto="HR,MP,RJ,TN,PY,UP,KA,KL")) == (0, 0)
      and rate(one(df0, **G, sub="1+1", fuel="Petrol", cc=(40000.01, 99999), rto=comp("HR", "MP", "RJ", "TN", "PY", "UP", "KA", "KL"))) == (5, 0))
check("E-Cart: GCV 3W / Goods Carrying Rickshaw, Electric only, ALLINDIA 50 / 43", len(sel(df0, product="GCV 3W")) == 2
      and rate(one(df0, product="GCV 3W", sub="1+1", fuel="Electric")) == (50, 43) and rate(one(df0, product="GCV 3W", sub="STP", fuel="Electric")) == (50, 43))
check("GCV 4W has 10 weight bands x 5 fuels in the bands that have one payout", set(zip(df0[df0["product"] == "GCV 4W"].cc_min, df0[df0["product"] == "GCV 4W"].cc_max)) ==
      {(0, 2000.01), (2000.01, 3500.01), (3500.01, 7500.01), (7500.01, 10000.01), (10000.01, 12000.01), (12000.01, 20000.01), (20000.01, 25000.01), (25000.01, 32000.01),
       (32000.01, 40000.01), (40000.01, 99999)})

# ---- passenger carrying
B = dict(product="PCV 4W", cls="Other Bus")
check("Bus 6 < PCC <= 10 (All Fuel): HR, RJ 7.5 (po 0.5), others 30 (po 23); sc 6.01-10.01, five fuels",
      rate(one(df0, **B, sub="1+1", fuel="Electric", sc=(6.01, 10.01), rto="HR,RJ")) == (7.5, 0.5) and rate(one(df0, **B, sub="STP", fuel="Diesel", sc=(6.01, 10.01), rto=comp("HR", "RJ"))) == (30, 23))
check("Bus 10 < PCC <= 20: TN, PY, KA 7.5, others 30; 20 < PCC <= 30 (Except EV): TN, PY, MP, UP 5 (po 0), others 10 (po 3) - four fuels, no Electric",
      rate(one(df0, **B, sub="1+1", fuel="Petrol", sc=(10.01, 20.01), rto="TN,PY,KA")) == (7.5, 0.5)
      and rate(one(df0, **B, sub="1+1", fuel="Petrol", sc=(20.01, 30.01), rto="TN,PY,MP,UP")) == (5, 0)
      and rate(one(df0, **B, sub="1+1", fuel="Petrol", sc=(20.01, 30.01), rto=comp("TN", "PY", "MP", "UP"))) == (10, 3)
      and sel(df0, **B, fuel="Electric", sc=(20.01, 30.01)).empty)
r = one(df0, **B, sub="1+1", fuel="Petrol", sc=(30.01, 999))
check("Bus PCC > 30 (Except EV): 'OD - 10%, TP - 0%' -> 1+1 is On OD and TP with pi_od 10 / po_od 3 and pi_tp 0; the STP row is On TP 0",
      r.pi_type == "On OD and TP" and float(r.pi_od_rate) == 10 and float(r.po_od_rate) == 3 and float(r.pi_tp_rate) == 0 and r.new_rto_list == "ALLINDIA"
      and float(one(df0, **B, sub="STP", fuel="Petrol", sc=(30.01, 999)).pi_tp_rate) == 0)
check("Bus PCC > 20 (Electric Vehicles): Electric only, sc 20.01-999, 10 / 3, 1+1 and STP",
      rate(one(df0, **B, sub="1+1", fuel="Electric", sc=(20.01, 999))) == (10, 3) and rate(one(df0, **B, sub="STP", fuel="Electric", sc=(20.01, 999))) == (10, 3)
      and sel(df0, **B, sc=(20.01, 999)).fuel_type.nunique() == 1)
T_ = dict(product="PCV 4W", cls="Taxi")
check("Taxi: the Package row is 1+1 only (15 for HR / KA / RJ / TN / PY / UP / MP, 25 others), the SATP row is STP only (15 / 30)",
      rate(one(df0, **T_, sub="1+1", fuel="Petrol", rto="HR,KA,RJ,TN,PY,UP,MP")) == (15, 8) and rate(one(df0, **T_, sub="1+1", fuel="Petrol", rto=comp("HR", "KA", "RJ", "TN", "PY", "UP", "MP"))) == (25, 18)
      and rate(one(df0, **T_, sub="STP", fuel="Petrol", rto=comp("HR", "KA", "RJ", "TN", "PY", "UP", "MP"))) == (30, 23) and len(sel(df0, **T_)) == 20)
check("School bus: All States -> ALLINDIA 62.5 / 55.5, add_tnc 'Educational Institution, Staff & School Buses', class School Bus, 1+1 and STP",
      rate(one(df0, product="PCV 4W", cls="School Bus", sub="1+1", fuel="Petrol", rto="ALLINDIA")) == (62.5, 55.5)
      and (sel(df0, cls="School Bus").add_tnc == mod.TNC_SCHOOL).all() and len(sel(df0, cls="School Bus")) == 10)
W3 = dict(product="PCV 3W", cls="Passenger Carrying Rickshaw")
check("Three wheeler (PCV): Madhya Pradesh 25 (po 18), WB / MH / GJ / DL / UK / UA / PB / GA / JK 60, every other state 40 (MP not in that list)",
      rate(one(df0, **W3, sub="1+1", fuel="Petrol", rto="MP")) == (25, 18) and rate(one(df0, **W3, sub="1+1", fuel="Petrol", rto="WB,MH,GJ,DL,UK,UA,PB,GA,JK")) == (60, 53)
      and rate(one(df0, **W3, sub="STP", fuel="Petrol", rto=comp("MP", "WB", "MH", "GJ", "DL", "UK", "UA", "PB", "GA", "JK"))) == (40, 33))
check("Two wheeler (PCV): 10 / 3 ALLINDIA", rate(one(df0, product="PCV 2W", sub="1+1", fuel="Petrol")) == (10, 3) and len(sel(df0, product="PCV 2W")) == 10)

# ---- two wheelers
OTH = comp("TN", "PY", "KL", "KA", "MP", "AS")
check("TW Upto 150 CC: Petrol, cc 0-150.01: 27.5 / 20.5 for the other states, 5 / 0 for TN, PY, KL, KA, MP, AS - 1+1, STP and 1+5 (New)",
      rate(one(df0, product="TW", sub="1+1", fuel="Petrol", cc=(0, 150.01), rto=OTH)) == (27.5, 20.5) and rate(one(df0, product="TW", sub="STP", fuel="Petrol", cc=(0, 150.01), rto="TN,PY,KL,KA,MP,AS")) == (5, 0)
      and rate(one(df0, product="TW", sub="1+5", fuel="Petrol", cc=(0, 150.01), rto=OTH)) == (27.5, 20.5) and one(df0, product="TW", sub="1+5", fuel="Petrol", cc=(0, 150.01), rto=OTH).policy_type == "New")
check("TW 150-350 and above 350 CC: 22.5 / 17.5; Electric Vehicles: Electric only, all cc, 22.5",
      rate(one(df0, product="TW", sub="1+1", fuel="Petrol", cc=(150.01, 350.01), rto=OTH)) == (22.5, 15.5) and rate(one(df0, product="TW", sub="1+1", fuel="Petrol", cc=(350.01, 99999), rto=OTH)) == (17.5, 10.5)
      and rate(one(df0, product="TW", sub="STP", fuel="Electric", cc=(0, 99999), rto=OTH)) == (22.5, 15.5) and sel(df0, product="TW", fuel="Electric", cc=(0, 150.01)).empty)
check("TW SAOD: All States 20 / 13, Petrol and Electric, cc 0-99999", len(sel(df0, product="TW", sub="SAOD")) == 2 and set(sel(df0, product="TW", sub="SAOD").fuel_type) == {"Petrol", "Electric"}
      and rate(one(df0, product="TW", sub="SAOD", fuel="Petrol")) == (20, 13) and one(df0, product="TW", sub="SAOD", fuel="Petrol").pi_type == "On OD")
check("TW: no row for the Long Term Package (5+5), and no TW row other than 1+1 / STP / 1+5 / SAOD", set(sel(df0, product="TW").sub_product) == {"1+1", "STP", "1+5", "SAOD"})

# ---- miscellaneous
check("MISCD: Ambulances 20 (tnc 'Ambulances'), all other 10 (tnc 'EX:Ambulances'), Agricultural Tractors 40 (class MISCD-Tractor); ALLINDIA; 1+1 and STP; five fuels",
      rate(one(df0, product="MISCD", cls="MISCD-Others", sub="1+1", fuel="Petrol", tnc="Ambulances")) == (20, 13) and rate(one(df0, product="MISCD", cls="MISCD-Others", sub="STP", fuel="Petrol", tnc="EX:Ambulances")) == (10, 3)
      and rate(one(df0, product="MISCD", cls="MISCD-Tractor", sub="1+1", fuel="Petrol")) == (40, 33) and len(sel(df0, product="MISCD")) == 30)
check("Motor Trade and Standalone CPA are not written (no class in the portal's upload)", set(sel(df0, product="MISCD").make_model_class) == {"MISCD-Others", "MISCD-Tractor"})

# ---- remarks
fl = df0[df0.remarks != ""]
check("flags: 26 TW rows carry the 5+5 payouts, the 30 MISCD rows carry Motor Trade and Standalone CPA; nothing else is flagged",
      len(fl) == 109 and set(fl["product"]) == {"TW", "MISCD", "Private Car"} and fl[fl["product"] == "TW"].remarks.str.contains("5+5", regex=False).all()
      and len(fl[fl["product"] == "TW"]) == 26 and len(fl[fl["product"] == "MISCD"]) == 30 and len(fl[fl["product"] == "Private Car"]) == 53
      and fl[fl["product"] == "MISCD"].remarks.str.contains("Motor Trade").all() and fl[fl["product"] == "MISCD"].remarks.str.contains("Standalone CPA").all()
      and fl.remarks.str.startswith("Manual Review Required: [").all(), (len(fl), sorted(set(fl["product"]))))

# ---- the four differences from the July upload
check("the school bus rate is 62.5 for ALL states in the PDF (the July upload had 9 states): the 9-state list is not in the output", sel(df0, cls="School Bus", rto="DL,GA,GJ,JK,MH,PB,UA,UK,WB").empty)
check("Pondicherry 3500-7500 kg 10% and the STP rows of the Sub Annexure 3 preferred RTOs are in the output (the July upload had neither)", len(sel(df0, **G, cc=c, rto="PY")) == 10
      and len(sel(df0, **G, sub="STP", rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP")) == 5)

# ================================================================ the document changes
def heads(tbls):
    return [c for t in tbls for c in t.cells if c.text.startswith("Pay-out (w.e.f.")]


t = tables()
for h in heads(t):
    h.text = h.text.replace("01-07-2026", "01-09-2026")
df, warns, _ = build(t)
check("the PDF's new column dated 01-09-2026 while you typed 01/07/2026: every row is flagged, nothing guessed or asked; no portal groups are known for that circular, so the rows that need "
      "one are not written (GCV 40 + Private Car Package 30 = -70) and a warning says so",
      (df.remarks.str.contains("effective from 01/09/2026", regex=False)).all() and len(df) == len(df0) - 70 and any("no portal RTO group is known for the circular effective 01/09/2026" in w for w in warns),
      (len(df), len(df0)))
df, _, _ = build(tables(), from_date=dt.date(2026, 8, 1), to_date=dt.date(2026, 8, 31))
check("... the same the other way round: from_date 01/08/2026 against a PDF dated 01-07-2026", df.remarks.str.contains("effective from 01/07/2026", regex=False).all())

t = tables()
find(t, 4, "a. GVW <= 2000").text = "a. GVW up to 2000"
df, warns, u = build(t)
fl = df[df.remarks.str.contains("up to 2000", regex=False)]
check("an unknown band label ('a. GVW up to 2000'): that payout is not written (-10 rows), the GCV rows are flagged (GCV 4W and 3W), a warning says so, no question",
      len(df) == len(df0) - 10 and u.asked == [] and set(fl["product"]) == {"GCV 4W", "GCV 3W"} and any("up to 2000" in w for w in warns), (len(df), len(fl)))

t = tables()
c0 = find(t, 3, "Haryana, Rajasthan", x_min=468)
c0.text = "Haryana, Atlantis"
df, _, u = build(t)
check("a states text with a word that is no state ('Atlantis'): that payout is not written (-10 rows) and the passenger rows are flagged",
      len(df) == len(df0) - 10 and df.remarks.str.contains("Atlantis", regex=False).any() and u.asked == [])

t = tables()
find(t, 3, "7.5%", x_min=400).text = "seven and a half"
df, _, u = build(t, [("read as a percentage", "Leave this payout out")])
check("a commission that is not a percentage asks once and, left out, is not written (and is flagged)", len(u.asked) == 1 and len(df) == len(df0) - 10 and df.remarks.str.contains("seven and a half", regex=False).any(), u.asked)
e = raises(lambda: build(t, interactive=False), mod.GridError)
check("... without a keyboard the same problem stops with the question as the message", e is not None and "not a percentage" in str(e), e)

t = tables()
find(t, 3, "7.5%", x_min=400).text = "NA"
df, _, u = build(t)
check("a July commission of NA means no payout for that case: nothing written for it, nothing flagged, no question", len(df) == len(df0) - 10 and u.asked == [] and len(df[df.remarks != ""]) == 109)

t = tables()
find(t, 9, "w.e.f. 01-07-2026) 56.5%", exact=False).text = "w.e.f. 01-07-2026) 60.0%"
df, _, _ = build(t)
check("Sub Annexure 3 commission read from its merged cell: 60% -> the preferred-RTO rows get 60 (po 53), 1+1 and STP",
      rate(one(df, **G, sub="1+1", fuel="Petrol", cc=b, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP")) == (60, 53)
      and rate(one(df, **G, sub="STP", fuel="Petrol", cc=b, rto=mod.RTO_PREFIX + "Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP")) == (60, 53))

t = [tb for tb in tables() if tb.page not in (9, 10)]
df, warns, _ = build(t)
plain = sel(df, **G, cc=b, rto="MP,TN,UP")
check("the Sub Annexure 3 table of 2000-3500 kg missing from the PDF (an annexure that is not in the document is NOT applied, user's rule of 26 Oct 2026): the '(Except RTOs ...)' payout is written "
      "for the whole states MP, TN, UP at 15 / 8 (10 rows) and flagged on exactly those rows; the preferred-RTO rows and the EX group are gone (-10 vs the full PDF), a warning says so",
      len(df) == len(df0) - 10 and len(plain) == 10 and rate(plain.iloc[0]) == (15, 8) and plain.remarks.str.contains("not in this document, so it is not applied", regex=False).all()
      and df.remarks.str.contains("not in this document, so it is not applied", regex=False).sum() == 10 and not df.new_rto_list.str.contains("2T-3.5T", regex=False).any()
      and any("applied without its annexure" in w for w in warns), (len(df), len(plain)))

old_groups = dict(mod.ANNEX3_GROUPS)
mod.ANNEX3_GROUPS.pop((3500, 7500))
try:
    df, warns, _ = build(tables())
    check("no portal RTO group configured for a Sub Annexure 3 band (3500-7500): warned, those rows are not written, nothing crashes", any("ANNEX3_GROUPS" in w for w in warns) and len(df) < len(df0), len(df))
finally:
    mod.ANNEX3_GROUPS.clear()
    mod.ANNEX3_GROUPS.update(old_groups)

t = tables()
for tb in t:
    if tb.page == 6:
        tb.cells = [c for c in tb.cells if not c.text.startswith("Commission")]
e = raises(lambda: build(t), mod.GridError)
check("a table whose Commission / States sub-columns are missing STOPS the run and says what it expected and what it found (11 Nov 2026: it used to be left out of the upload with a warning)",
      e is not None and "Commission" in str(e) and "sub-columns" in str(e) and "expected" in str(e) and "page 6" in str(e), e)

t = tables()
t[0].outside.append("New rule: every payout is 2% higher for business sourced through POSPs from 15 July")
df, _, _ = build(t)
check("a new rule-sized text outside the tables is flagged on every row", df.remarks.str.contains("every payout is 2% higher", regex=False).all())
old_standing = list(mod.STANDING_NOTES)
mod.STANDING_NOTES.append("business sourced through POSPs")
try:
    df, _, _ = build(t)
    check("... unless it is a standing note", not df.remarks.str.contains("every payout is 2% higher", regex=False).any())
finally:
    mod.STANDING_NOTES[:] = old_standing

e = raises(lambda: build([]), mod.GridError)
check("a PDF without any table stops with a clear message", e is not None and "no payout found" in str(e), e)
e = raises(lambda: mod.build("fake.pdf", T, F, quiet, mod.Asker(False), tables=tables()), mod.GridError)
check("to_date before from_date stops", e is not None and "before from_date" in str(e), e)

t = tables()
find(t, 3, "PCC <= 10 (All Fuel)").text = "PCC <= 10 (Any Fuel)"
df, _, _ = build(t)
check("an unknown fuel wording in a passenger band: that band is not written and flagged, not guessed", len(df) == len(df0) - 20 and df.remarks.str.contains("Any Fuel", regex=False).any(), len(df))

t = tables()
find(t, 5, "Haryana, Madhya Pradesh, Rajasthan, Tamil Nadu, Pondicherry, Uttar Pradesh (Except RTOs listed in Sub Annexure 3)", x_min=468, exact=False).text = \
    "Haryana, Madhya Pradesh, Rajasthan, Tamil Nadu, Uttar Pradesh (Except RTOs listed in Sub Annexure 3)"
df, _, _ = build(t)
check("Pondicherry taken out of the 3500-7500 kg text: its own row disappears (-10 rows)", len(df) == len(df0) - 10 and sel(df, **G, cc=c, rto="PY").empty)

bad = os.path.join(TMP, "not_a.pdf")
open(bad, "w").write("this is not a pdf")
e = raises(lambda: mod.build(bad, F, T, quiet, mod.Asker(False)), Exception)
check("a file that is not a PDF fails with an explained stop or a retry question, not a traceback",
      e is not None and (isinstance(e, mod.GridError)), repr(e)[:120])

# ================================================================ Private Car: the table's segments, Sub Annexure 4, the master lists
ALLM, SEVENM, EXSM = mod.MAKES_CAR_ALL, mod.MAKES_CAR_SEVEN, mod.MAKES_CAR_EX_SEVEN
EXR, A4R = mod.CAR_RTO_EX_ANNEX4, mod.CAR_RTO_ANNEX4
car = df0[df0["product"] == "Private Car"]


def crow(df, sub, fuel, cc, makes, rto):
    d = df[(df["product"] == "Private Car") & (df.sub_product == sub) & (df.fuel_type == fuel) & (df.cc_min.astype(float) == cc[0]) & (df.cc_max.astype(float) == cc[1])
           & (df.new_vehicle_makes == makes) & (df.new_rto_list == rto)]
    assert len(d) == 1, f"expected one Private Car row, got {len(d)} for {sub} {fuel} {cc} {makes} {rto}"
    return rate(d.iloc[0])


def covers(rows, fuel, cc, cls):
    """how many rows cover a vehicle: fuel, cc, and make class 'seven' (one of the 7 makes) or 'other'"""
    return int(sum(1 for _, r in rows.iterrows() if r.fuel_type == fuel and float(r.cc_min) < cc <= float(r.cc_max)
                   and (r.new_vehicle_makes == ALLM or (r.new_vehicle_makes == SEVENM and cls == "seven") or (r.new_vehicle_makes == EXSM and cls == "other"))))


check("Private Car: 53 rows (1+1 20, 1+3 14, SAOD 5, STP 14), class Car, 1+3 is policy_type New, the rest Renew, On Net / On OD / On TP",
      len(car) == 53 and car.sub_product.value_counts().to_dict() == {"1+1": 20, "1+3": 14, "SAOD": 5, "STP": 14} and set(car.make_model_class) == {"Car"}
      and set(car[car.sub_product == "1+3"].policy_type) == {"New"} and set(car[car.sub_product != "1+3"].policy_type) == {"Renew"}
      and set(car[car.sub_product.isin(["1+1", "1+3"])].pi_type) == {"On Net"} and set(car[car.sub_product == "SAOD"].pi_type) == {"On OD"} and set(car[car.sub_product == "STP"].pi_type) == {"On TP"})
check("segment 'diesel vehicles upto 1500 CC': Diesel only, cc 0-1500.01, all makes - Package 5 / 0 (EX Annexure 4 RTOs), New 10 / 3 and SATP 5 / 0 (ALLINDIA)",
      crow(df0, "1+1", "Diesel", (0, 1500.01), ALLM, EXR) == (5, 0) and crow(df0, "1+3", "Diesel", (0, 1500.01), ALLM, "ALLINDIA") == (10, 3) and crow(df0, "STP", "Diesel", (0, 1500.01), ALLM, "ALLINDIA") == (5, 0)
      and set(car[(car.cc_max.astype(float) == 1500.01) & (car.cc_min.astype(float) == 0) & (car.new_rto_list != A4R)].fuel_type) == {"Diesel"})
check("segment 'above 2500 CC except the 7 makes': the other-makes group, all four fuels (Diesel too), Package 5, New 10, SATP 5",
      all(crow(df0, "1+1", f, (2500.01, 99999), EXSM, EXR) == (5, 0) for f in ("Petrol", "Diesel", "CNG", "LPG"))
      and all(crow(df0, "1+3", f, (2500.01, 99999), EXSM, "ALLINDIA") == (10, 3) for f in ("Petrol", "Diesel", "CNG", "LPG"))
      and crow(df0, "STP", "Diesel", (2500.01, 99999), EXSM, "ALLINDIA") == (5, 0))
check("segment 'other than above 2 segments' = the 7 makes above 2500 cc (all fuels), every other non-diesel up to 2500 cc, diesel from 1500 to 2500 cc: Package 20, New 27.5, SATP 20",
      crow(df0, "1+1", "Diesel", (2500.01, 99999), SEVENM, EXR) == (20, 13) and crow(df0, "1+1", "Petrol", (0, 2500.01), ALLM, EXR) == (20, 13) and crow(df0, "1+1", "Diesel", (1500.01, 2500.01), ALLM, EXR) == (20, 13)
      and crow(df0, "1+3", "CNG", (2500.01, 99999), SEVENM, "ALLINDIA") == (27.5, 20.5) and crow(df0, "1+3", "LPG", (0, 2500.01), ALLM, "ALLINDIA") == (27.5, 20.5)
      and crow(df0, "STP", "Diesel", (1500.01, 2500.01), ALLM, "ALLINDIA") == (20, 13))
parts = [("1+1", EXR), ("1+1", A4R), ("1+3", "ALLINDIA"), ("STP", "ALLINDIA")]
bad = [(sub, rto, f, cc, cls, covers(car[(car.sub_product == sub) & (car.new_rto_list == rto)], f, cc, cls))
       for sub, rto in parts for f in ("Petrol", "Diesel", "CNG", "LPG") for cc in (500, 1000, 1500, 1501, 2000, 2500, 2501, 4000) for cls in ("seven", "other")
       if covers(car[(car.sub_product == sub) & (car.new_rto_list == rto)], f, cc, cls) != 1]
check("the segments partition every non-electric vehicle: for Package (EX group and Annexure 4 group), New and SATP each combination of fuel, cc and make is covered by exactly one row", not bad, bad[:4])
check("SAOD: all vehicles 12 / 5 for the four non-electric fuels (ALLINDIA, all cc); Electric: New 27.5, Package 17 (EX group), SAOD 17, SATP 17 - one row each, all business types",
      all(crow(df0, "SAOD", f, (0, 99999), ALLM, "ALLINDIA") == (12, 5) for f in ("Petrol", "Diesel", "CNG", "LPG")) and crow(df0, "1+3", "Electric", (0, 99999), ALLM, "ALLINDIA") == (27.5, 20.5)
      and crow(df0, "1+1", "Electric", (0, 99999), ALLM, EXR) == (17, 10) and crow(df0, "SAOD", "Electric", (0, 99999), ALLM, "ALLINDIA") == (17, 10)
      and crow(df0, "STP", "Electric", (0, 99999), ALLM, "ALLINDIA") == (17, 10))
check("Sub Annexure 4: Package in the 40% RTOs - Petrol / CNG / LPG / Electric all cc and Diesel above 1500 cc at 40 / 33, Diesel up to 1500 cc at the table's 5 / 0; New, SAOD and SATP get no Annexure 4 row",
      all(crow(df0, "1+1", f, (0, 99999), ALLM, A4R) == (40, 33) for f in ("Petrol", "CNG", "LPG", "Electric")) and crow(df0, "1+1", "Diesel", (1500.01, 99999), ALLM, A4R) == (40, 33)
      and crow(df0, "1+1", "Diesel", (0, 1500.01), ALLM, A4R) == (5, 0) and set(car[car.new_rto_list == A4R].sub_product) == {"1+1"} and len(car[car.new_rto_list == A4R]) == 6)
check("dates: a payout that is the same in the Existing and the new column keeps the Existing date (01/04/2026); the Sub Annexure 4 40% rows take from_date (01/07/2026)",
      (car[car.new_rto_list != A4R].from_date == "01/04/2026").all() and (car[(car.new_rto_list == A4R) & (car.pi_net_rate.astype(float) == 40)].from_date == "01/07/2026").all()
      and (car.to_date == "31/07/2026").all() and len(car[(car.new_rto_list == A4R) & (car.pi_net_rate.astype(float) == 5)]) == 1 and (car[(car.new_rto_list == A4R) & (car.pi_net_rate.astype(float) == 5)].from_date == "01/04/2026").all())
check("Package rows use the EX Annexure 4 group; New, SAOD and SATP use ALLINDIA (the user decided on 26 Oct 2026: Sub Annexure 4 covers Package 1+1 only) with no remark about the scope",
      set(car[(car.sub_product == "1+1") & (car.new_rto_list != A4R)].new_rto_list) == {EXR} and set(car[car.sub_product != "1+1"].new_rto_list) == {"ALLINDIA"}
      and len(car[car.sub_product != "1+1"]) == 33 and not car.remarks.str.contains("does not say which policy types", regex=False).any())
check("Long Term Package (3+3) has no sub_product in the portal's upload: not written, its three payouts are warned about and flagged on every Private Car row",
      "3+3" not in set(df0.sub_product) and car.remarks.str.contains("3+3", regex=False).all() and sum("3+3" in w for w in warn0) == 3)
check("master check: the EX Annexure 4 group lacks the 99 Madhya Pradesh RTOs that ALLINDIA has -> flagged on the rows that use it; the 7-makes group lacks 7 models, the other-makes group lacks CITROEN",
      car[car.new_rto_list == EXR].remarks.str.contains("lacks 99 RTO", regex=False).all() and car[car.new_rto_list == EXR].remarks.str.contains("MP 99", regex=False).all()
      and car[car.new_vehicle_makes == SEVENM].remarks.str.contains("lacks 7", regex=False).all() and car[car.new_vehicle_makes == EXSM].remarks.str.contains("CITROEN", regex=False).all()
      and not car[car.new_rto_list == "ALLINDIA"].remarks.str.contains("lacks 99 RTO", regex=False).any())
check("master check: the Sub Annexure 3 and 4 RTO groups hold exactly the RTOs the circular lists (no flag about them)", not df0.remarks.str.contains("does not hold the preferred", regex=False).any()
      and not df0.remarks.str.contains("does not hold the RTOs of Sub Annexure 4", regex=False).any() and not df0.remarks.str.contains("does not hold the other RTOs", regex=False).any())

old = mod.CAR_ANNEX4_POLICIES
mod.CAR_ANNEX4_POLICIES = ("1+1", "SAOD")
try:
    df, _, _ = build(tables())
    cr = df[df["product"] == "Private Car"]
    check("CAR_ANNEX4_POLICIES = (1+1, SAOD): SAOD gets the Annexure 4 rows (40 / 33; diesel up to 1500 cc at the table's 12 / 5) and its main rows move to the EX group; New and SATP stay on ALLINDIA",
          crow(df, "SAOD", "Petrol", (0, 99999), ALLM, A4R) == (40, 33) and crow(df, "SAOD", "Diesel", (0, 1500.01), ALLM, A4R) == (12, 5) and crow(df, "SAOD", "Petrol", (0, 99999), ALLM, EXR) == (12, 5)
          and set(cr[cr.sub_product.isin(["1+3", "STP"])].new_rto_list) == {"ALLINDIA"}
          and not cr[cr.sub_product == "SAOD"].remarks.str.contains("does not say which policy types", regex=False).any())
finally:
    mod.CAR_ANNEX4_POLICIES = old

t = tables()
find(t, 13, "e.f. 01-07-2026) 40.0%", exact=False).text = "e.f. 01-07-2026) 45.0%"
df, _, _ = build(t)
check("the Sub Annexure 4 commission is read from the circular (45% -> the Annexure 4 rows get 45 / 38)", crow(df, "1+1", "Petrol", (0, 99999), ALLM, A4R) == (45, 38))

t = [tb for tb in tables() if tb.page != 13]
df, warns, _ = build(t)
cr = df[df["product"] == "Private Car"]
check("Sub Annexure 4 missing from the PDF: no Annexure 4 rows (-6), flagged on the Package rows, every other Private Car row is still written",
      len(cr) == 47 and A4R not in set(cr.new_rto_list) and cr[cr.sub_product == "1+1"].remarks.str.contains("Sub Annexure 4 (the 40% RTOs) was not found", regex=False).all(), len(cr))

t = tables()
find(t, 7, "Other than above 2 segments", x_min=100, x_max=270).text = "Everything else"
df, warns, u = build(t)
cr = df[df["product"] == "Private Car"]
check("a segment label that is not one of the table's: its payouts are not written, flagged on the Private Car rows and warned about; no question",
      len(cr) < 53 and cr.remarks.str.contains("Everything else", regex=False).any() and any("Everything else" in w for w in warns) and u.asked == [])

t = tables()
find(t, 7, "Package", x_min=60, x_max=150).text = "Package Plus"
df, warns, _ = build(t)
check("an unknown policy type: its payouts are not written and flagged, the other policy types are written",
      len(df[df["product"] == "Private Car"]) < 53 and df.remarks.str.contains("Package Plus", regex=False).any() and (df["product"] == "Private Car").any())

# ---- the master lists
old_make = mod.MAKES_CAR_SEVEN
mod.MAKES_CAR_SEVEN = "CAR_Tata_Maruti_Mahindra_Toyota_Hyundai_Honda_KIA"
try:
    df, _, _ = build(tables())
    check("a make group the master spells differently only in capitals is flagged with the master's spelling as a hint - not replaced by a guess",
          df.remarks.str.contains("not in the make / model master", regex=False).any() and df.remarks.str.contains("the master spells it 'CAR_Tata_Maruti_Mahindra_Toyota_Hyundai_Honda_Kia'", regex=False).any()
          and (df.new_vehicle_makes == "CAR_Tata_Maruti_Mahindra_Toyota_Hyundai_Honda_KIA").any())
finally:
    mod.MAKES_CAR_SEVEN = old_make
old_rto = mod.CAR_RTO_ANNEX4
mod.CAR_RTO_ANNEX4 = "UNITED_CAR_RTO_Annexure4_JULY27"
try:
    df, _, _ = build(tables())
    fl = df[df.new_rto_list == "UNITED_CAR_RTO_Annexure4_JULY27"]
    check("an RTO group that is not in the RTO master (a new month's name nobody created) is flagged on its rows (6), still written",
          len(fl) == 6 and fl.remarks.str.contains("is not in the RTO master", regex=False).all())
finally:
    mod.CAR_RTO_ANNEX4 = old_rto
old_cls = dict(mod.CLASSES["gcv4w"])
mod.CLASSES["gcv4w"]["makes"] = "gcv_4w_all"
try:
    df, _, _ = build(tables())
    fl = df[df.new_vehicle_makes == "gcv_4w_all"]
    check("a make group in the wrong capitals on 210 GCV 4W rows is flagged with the master's spelling ('GCV_4W_ALL')", len(fl) == 210 and fl.remarks.str.contains("the master spells it 'GCV_4W_ALL'", regex=False).all(), len(fl))
finally:
    mod.CLASSES["gcv4w"].update(old_cls)

t = tables()
find(t, 9, "MP01").text = ""
df, _, _ = build(t)
grp = mod.RTO_PREFIX + mod.ANNEX3_GROUPS[(2000, 3500)][0]
fl = df[df.new_rto_list == grp]
check("a Sub Annexure 3 RTO the master group has but the circular's list does not (MP01 taken out of the PDF): flagged on the rows of that group", len(fl) > 0 and fl.remarks.str.contains("does not hold", regex=False).all()
      and fl.remarks.str.contains("MP01", regex=False).all(), (len(fl), fl.remarks.iloc[0][:200] if len(fl) else ""))

e = raises(lambda: mod.load_masters(mod.Asker(False), folder=TMP), mod.GridError)
check("master workbooks not found in the folder: without a keyboard the run stops with the question as the message", e is not None and "not found" in str(e), e)
user = Smart([("holds the two master", "x")])
mst = mod.load_masters(mod.Asker(False))
check("the masters load from the master folder (5,828 RTO groups, 595 make / model groups) and know ALLINDIA's 3,663 RTOs", len(mst.rto) == 5828 and len(mst.makes) == 595 and len(mst.codes("ALLINDIA")) == 3663)

# ================================================================ duplicates and the written file
two = tables() + tables()
mod_df, w = mod.build("fake.pdf", F, T, quiet, mod.Asker(False), tables=two)
user = Smart([])
dd = mod.report(mod_df, w, quiet, mod.Asker(True, user.inp, user.out))
check("the same table twice repeats every key: settled when reported (identical payins = one row), nothing is asked, no duplicate is left", int(mod_df.duplicated(mod.KEY_COLUMNS).sum()) > 0 and int(dd.duplicated(mod.KEY_COLUMNS).sum()) == 0
      and not user.asked and len(dd) == len(df0), len(dd))

out = os.path.join(TMP, f"o{next(counter)}.csv")
df, warns = mod.build("fake.pdf", F, T, quiet, mod.Asker(False), tables=tables())
mod.write_csv(mod.to_csv_frame(df), out, mod.Asker(False), quiet)
problems = helpers.check_csv(out, mod.OUTPUT_COLUMNS, mod.NUMERIC_COLUMNS)
back = pd.read_csv(out, dtype=str, keep_default_na=False)
check("the written file has the 42 columns, numbers only in numeric columns, DD/MM/YYYY dates, remarks only as Manual Review flags, 457 rows", problems == [] and len(back) == 457, problems)

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
