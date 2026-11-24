"""
SBI General Insurance broker grid -> SBI upload CSV
====================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (NN).csv") from the workbook SBI General circulates for the Agency / C & D broker channel
("Agent_C__D_Cat._Broker_Grid_revision__Oct26_W.e.f_3rd_Oct.xlsx").  Rates come only from that workbook.

What the workbook holds and how it becomes rows (checked against the user's October 2026 upload, 102,604 rows, from_date 03/10/2026)
-----------------------------------------------------------------------------------------------------------------------------------
Hidden sheets are older versions (the hidden 'Sheet1' is one) and are never read.  EVERYTHING is found by what the headers and the notes SAY, never by a cell address: SBI moves columns and rows.

* The rate sheet (the visible sheet whose header row says 'RTO Cluster Name' | 'State' and that holds dozens of numeric columns) is one wide table: one row per RTO cluster (54 in October), then a
  block of columns per product.  The header is 5 rows of merged cells (product, segment, fuel / age, age / cc band, Comp / SATP); the script reads the merged ranges and picks every block by its words:
    GCV 4W up to 2.0 t (all makes) / 2.0-2.5 t Tata / 2.0-2.5 t other makes, GCV 3W, 2.5-3.5 t Mahindra / Tata 2.5-3.0 t / Ashok Leyland 2.5-3.5 t + Tata 3.0-3.5 t, GCV tractor, 3.5-12 t (declined),
    12-20 t, 20-40 t, > 40 t (each: other makes = Mahindra, Tata & Ashok Leyland), Private Car comp / SAOD 'other than High-End', PCV 3W (diesel / non-diesel, new / non-new), PCV Taxi (3 cc bands:
    comp nil-dep / comp non-nil-dep / SATP), TW (SAOD, 1+1 & SATP; bike / scooter), Agricultural Tractor & Harvester, Private Car SATP (10+ years, 3 cc bands, two fuel groups), School Bus (18+ / 12-17 seats).
  Each block has an age column set (New, Upto 5 / Age 1-5, Above 5) x (Comp, SATP).  Comp = 1+1 / On Net, SATP = STP / On TP, SAOD = On OD.  A blank or zero cell is a declined entry: a ZERO card is
  written (the user's October upload has one for every cluster).  A column the script has no rule for is listed (sheet 'Not written') and warned about - never guessed.
* 'PV SAOD- HIGHEND Enabler': the payouts of the High-End models by RTO cluster (the user, 1 Nov 2026: the 20 given there are the rates of the High-End make / models; the small-car rates are in the rate sheet).
  Model cluster -> the portal's make group  SBI_MAY26_<make>_<model>  (a group the make master lacks is flagged, e.g. 'MB Class').  The note 'Zero payout will be applicable on High-End NON NCB policies'
  makes two rows per entry: is_ncb YES = the payout, NO = 0.  'Petrol / Hybrid' = Petrol + Electric, 'Diesel / CNG / LPG' = those three fuels.
* 'RTO Cluster Mapping' (RTO code -> cluster): used to check that the clusters of the rate sheet and of this list agree and to name the cluster minus the RTOs a note takes out (AP - Rest without AP02 / AP03 /
  AP04 / AP21 / AP26 / AP27 for GCV 2.5-3.5 t: 'Non doable RTOs').
* 'PV SATP Decline make and Model' and 'Car, Taxi Declined list': the declined makes / models.  The portal's SBI make groups already leave them out; the script fingerprints both lists and flags the Private
  Car / taxi rows when SBI changes a list (the make groups then need a look).

The notes under the rate table are read by a TEXT SCANNER (`read_rules`): every rule is found by its words (a label such as 'Pvt Car' / 'PCV Taxi' / 'Tw' + keywords), its number is taken from the note, and the
words it understood are used up.  Rules read: the premium slabs (below), vehicle age limits per product (comp / SAOD and SATP), Private Car non-NCB cut (-5 on non-High-End comp and SAOD), High-End non-NCB = 0,
taxi non-NCB / new cut (-3), Private Car SATP cut (-3 for ages 1-10, zero below 1 year), TW cc limit and make list, the fuels a note declines ('Fuel type Electric declined for all GCV', 'In PCV 3W EV Decline':
one reader for every product and fuel, `read_fuel_declines`; no row is written for a declined fuel), 'additional 2% - DL Registration Vehicle 2.5 to 3.5 Segment Only' (`read_uplifts`: N extra points on the
GCV 4W weight bands inside that tonnage, Comp and SATP, in the clusters that hold that state's RTOs - the whole cluster), the grid's validity dates, the non-doable AP RTOs.
Any other text is checked against STANDING_NOTES; what remains is flagged  Manual Review Required: [text]  in `remarks` on the rows it may affect (scoped by the segment / state it names, else every row).
The 'Declined cluster' notes are CROSS-CHECKED with the grid (a cluster the note declines must have no payout; a cluster with no payout that the note does not decline is flagged).

Premium slab (asked every month, the user's choice of 1 Nov 2026): the grid is for the top slab; 'Pvt Car Slab' (1L-3L -1, 50K-1L -2, 15K-50K -3) and 'CV Slab' (3L-4.99L -1, 1L-2.99L -2, 15K-99K -3)
lower EVERY payout of that family by the slab's points.  Private Car follows the Private Car slab, everything else (GCV, PCV, school bus, tractor) the CV slab; the High-End Enabler rates are as given.
TWO-WHEELERS HAVE NO SLAB (the user, 3 Nov 2026): the grid has no two-wheeler slab note, so their payouts are written as printed.  October and November: both slabs -1.  (--pc-slab / --cv-slab answer the questions.)

Differences between the document and the user's October upload (listed, the document wins): GCV 2.5-3.5 t SATP above 5 years goes up to 25 years like every other SATP row (the upload had 20);
Private Car SATP gets a 0-1 year band at zero payout; High-End fuels follow the sheet's words (Diesel / CNG / LPG together, Petrol / Hybrid = Petrol + Electric); High-End non-NCB rows carry 0.

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a sheet or header that is not where it should be, a note it relies on that has
gone, dates that look wrong - it stops, asks you a specific question (type the number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every answer is listed at
the end of the run, and you confirm before the file is written.  It never asks you for a rate.  With --no-prompt (or when there is no keyboard) it asks nothing and stops with the question as the error
message instead.  If something goes wrong that it did not foresee it says so in one line and saves the technical details to a *_error_*.txt file.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run sbi_grid_to_config.py`) and run it.  Upload the workbook (and, optionally, an earlier upload CSV to get a diff report); it asks for the two
slabs, from_date and to_date (DD/MM/YYYY), writes the dates into every row, and downloads SBI_UPLOAD_<yyyy-mm>.csv.  The master workbooks (RTO / make-model) must be in a `master` folder next to the script.

Command line
------------
    python sbi_grid_to_config.py [--grid <workbook.xlsx>] [--from-date 03/10/2026 --to-date 31/03/2030] [--pc-slab 1L-3L --cv-slab 3L-4.99L] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the workbook, the dates, the slabs, the earlier upload - is asked for; --pc-slab / --cv-slab take a slab name from the note or 'none' for the grid as printed)
"""
import argparse
import datetime as dt
import difflib
import getpass
import hashlib
import io
import math
import os
import re
import sys
import time
import traceback

import numpy as np
import pandas as pd

try:  # only present inside Google Colab
    from google.colab import files as colab_files  # type: ignore
    IN_COLAB = True
except Exception:
    colab_files = None
    IN_COLAB = False


# %% ---------------------------------------------------------------- SETTINGS
INSURANCE_COMPANY = "SBI GENERAL INSURANCE COMPANY LIMITED"
USER_ID = 2966
ROW_ID = 1                                       # the portal's export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 2
CC_OPEN_END = 99999
AGE_OPEN_END = 99
SC_RANGE = (0, 999)
TARIFF_FULL = 100.01                             # no discount band: every row is 0-100.01
MAX_NOTE_CHARS = 500
HEADER_SCAN_ROWS = 12

GROUP_PREFIX = "SBI_OCT26_"                      # the month the portal's RTO groups were created: <prefix><cluster, spaces -> '_'>  (change it when the master gets groups under a new prefix)
HE_GROUP_PREFIX = "SBI_MAY26_"                   # the High-End make groups: <prefix><make>_<model>  (SBI_MAY26_BMW_5_Series)

# the make groups of the portal's make / model master (the master's spelling; check_names flags a name it lacks)
MK_GCV_ALL, MK_GCV3W, MK_PCV3W, MK_MISCD, MK_BUS = "GCV_4W_ALL", "GCV_3W_ALL", "PCV_3W_ALL", "MISCD_ALL", "BUS_ALL"
MK_BELOW35, MK_BELOW35_EX_TATA = "SBI_Below_3.5T_GCV4W_MAKES", "SBI_Below_3.5T_GCV4W_MAKES_EX_TATA"
MK_TATA, MK_ASHOK, MK_MAHINDRA = "TATA", "Ashok", "Mahindra"
MK_TAXI, MK_PC, MK_PC_TP = "SBI_MAY26_TAXI", "SBI_MAY26_Non-High-End", "SBI_JUNE26_PC_TP"
TW_MAKE_NAMES = {"hero": "HERO", "honda": "HONDA", "yamaha": "YAMAHA", "tvs": "TVS", "suzuki": "SUZUKI", "re": "ROYAL", "royal enfield": "ROYAL", "bajaj": "BAJAJ", "ktm": "KTM", "ather": "ATHER"}
TW_MAKE_ORDER = ["HERO", "HONDA", "YAMAHA", "TVS", "SUZUKI", "ROYAL", "BAJAJ", "KTM", "ATHER"]      # the order of the user's October upload
SCOOTER_WITHOUT = {"ROYAL"}                      # Royal Enfield builds no scooters: the scooter rows leave it out (the October upload)

GCV_CLASSES = ["Pick Up Van", "Tanker", "Dumper Tipper", "Trailer", "Truck"]     # the grid does not say the body type: the same payout is written for each class (the October upload)
FUELS_ALL = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NO_EV = ["Petrol", "Diesel", "CNG", "LPG"]  # GCV: 'Fuel type Electric declined for all GCV' (read from the notes)
FUELS_TW = ["Petrol", "Electric"]
FUELS_NON_DIESEL_3W = ["Petrol", "CNG", "LPG", "Electric"]
TNC_NEW, TNC_OLD, TNC_HARVESTER = "New", "Old", "Harvester,EX Trailer"
MIN_NOTE_CHARS = 8                               # a WHOLE text that no rule read is a heading ('Notes:', 'Slab:-') below this length and a rule above it, until a standing note says otherwise.  The shared helper's
                                                 # default of 25 dropped 'In PCV 3W EV Decline' (20 characters) without a word - it was neither applied nor flagged (the user, 3 Nov 2026)

# a rule whose note has gone from the grid is ASKED about, with last month's value as the first option
LAST_KNOWN = {"age.gcv": (20, 25), "age.pc.comp": 15, "age.pc.satp": 25, "age.pcv3w": (15, 25), "age.taxi": (15, 20), "age.bus": 25, "age.tw.comp": 10, "age.tw.satp": 15, "age.tractor": 25,
              "pc.noncb_cut": 5.0, "he.noncb_zero": True, "taxi.noncb_cut": 3.0, "pc.satp": (3.0, 1, 10, 1), "tw.cc_max": 500, "gcv.no_electric": True}
# the fuels the last grid declined, per product key (FUEL_PRODUCTS), besides GCV Electric above: when a grid no longer says so the run WARNS that those rows are written again (a question here would also
# be asked of every older grid).  November 2026: 'In PCV 3W EV Decline'.
LAST_FUEL_DECLINES = {"pcv3w": ("Electric",)}
# the extra points the last grid gave to a state's registrations in a tonnage range ('additional 2% - DL Registration Vehicle 2.5 to 3.5 Segment Only'): (state, from t, to t) -> points; the run WARNS when one is gone
LAST_UPLIFTS = {("DL", 2.5, 3.5): 2.0}

# the declined make / model lists ('PV SATP Decline make and Model', 'Car, Taxi Declined list') as the SBI make groups were built from them (October 2026); a change is flagged
DECLINE_FINGERPRINT = "0f1c7aa6adb9"

STANDING_NOTES = [                               # texts that need no action (the rule is inside the grid's cells, the portal's groups handle it, or it is an information line)
    "All the Grids are applicable subject to meeting UW guideline",
    "Make / GVW change endorsement would require Head Agency approval",
    "In case of any intentional wrong booking on Make & GVW entire PO will be forfeited",
    "Sourcing from permitted RTOs only. For details refer UW guidelines",
    "Other existing RTO State definitions continue",
    "Management reserved the rights to withdraw grid by giving one week notice",
    # declined makes: the SBI make groups of the portal already leave them out (a changed list is a NEW text and is flagged)
    "All imported makes & Volvo, MAN, AMW, Mercedes, Eicher (above 3.5T), Isuzu, Bharat Benz, Hyundai, SCANIA",
    "For above 12 T allowed make are Tata, AL & Mahindra only",
    "Mahindra Supro is added and considered under the Grid of Upto 2.0 T-All makes & 2.0-2.5T Tata makes",
    "Pan India decline all cluster all make model",
    "Decline Make and model and U/W guidline also attached",
    "PVT CAR SATP Declined Make & Model list attached",
    # two-wheeler exclusions the portal cannot evaluate (the user's October upload carries only the first one as a T&C)
    "Vehicles built /used for Driving tuition, Demo, Airside operations or for special purposes",
    "All Imported vehicles Decline",
    "Vehicle Carrying geographical extension to Pakistan, Maldives & Sri Lanka",
    "1+5 Tw not allowed",
    # school bus: already inside the grid's cells / the header bands
    "School Bus rates to apply on UW accepted proposals as LoB type",
    "Payout of 2% to be deducted on 12+1 seater, full payout to be applied for 18+1-seater and above",
    "staff bus U/W acceptance only",
    "Seating capacity 12+1 and above",
    "Below 12 seating capacity school bus is not allowed",
    "Agency, C & D type of Broker Note",
    # the High-End enabler sheet: its title and the line 'Pvt Car Slabs to Apply' (the user, 1 Nov 2026: the 20 given there are the High-End rates as printed - no slab on them)
    "Bonnazza for HIGHEND Private Car",
    "Pvt Car Slabs to Apply",
]

# A note that is neither understood by the text scanner nor standing is flagged on the rows it may affect: the segments and states it names (`note_scope`), else EVERY row.
EXCEL_EXTS = (".xlsx", ".xlsm")

OUTPUT_COLUMNS = [
    "id", "insurance_company", "insurer_vertical", "product", "sub_product", "policy_type", "fuel_type",
    "make_model_class", "new_vehicle_makes", "new_rto_list", "vehicle_age_min", "vehicle_age_max", "cc_min",
    "cc_max", "sc_min", "sc_max", "pi_od_rate", "pi_tp_rate", "pi_tp_2", "pi_tp_3", "pi_tp_4", "pi_tp_5",
    "pi_net_rate", "pi_flat_amount", "pi_vli", "pi_type", "tariff_min", "tariff_max", "is_ncb", "is_cpa",
    "is_zd", "from_date", "to_date", "user_id", "veh_use", "add_tnc", "remarks", "po_type", "po_od_rate",
    "po_tp_rate", "po_net_rate", "po_flat_amount",
]
NUMERIC_COLUMNS = [
    "vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "sc_min", "sc_max", "pi_od_rate", "pi_tp_rate",
    "pi_tp_2", "pi_tp_3", "pi_tp_4", "pi_tp_5", "pi_net_rate", "pi_flat_amount", "pi_vli", "tariff_min",
    "tariff_max", "po_od_rate", "po_tp_rate", "po_net_rate", "po_flat_amount",
]
KEY_COLUMNS = [
    "new_vehicle_makes", "new_rto_list", "product", "sub_product", "policy_type", "fuel_type", "vehicle_age_min",
    "vehicle_age_max", "make_model_class", "is_ncb", "is_cpa", "is_zd", "cc_min", "cc_max", "sc_min", "sc_max", "add_tnc",
    "tariff_min", "tariff_max",
]

# the user, 1 Nov 2026: every row takes the from_date typed for the run - rows of the earlier upload do NOT keep their old date
KEEP_REF_DATES = False

OUTPUT_FOLDER = "sbi"                # the folder of this insurer under output\ (default_output)


# %% ---------------------------------------------------------------- QUESTIONS
# (this block, down to the "READING THE ..." heading, is identical in the other script)
def say(*parts):
    """print() that never fails on characters the console cannot show"""
    text = " ".join(str(p) for p in parts)
    try:
        print(text)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "ascii"
        print(text.encode(enc, "replace").decode(enc))


class GridError(Exception):
    """The input does not look the way this script expects, or you stopped it; nothing is written."""


class NeedsAnswer(GridError):
    """A question came up but this run cannot ask it (no keyboard, or --no-prompt). The message is the question."""


# ---- service mode: the same script driven by a web application instead of a person at the keyboard.  It is switched on by environment variables only, so a run without them is
# exactly what it was.  GRID_SERVICE=1: a question is not asked; it is written to GRID_QUESTIONS_FILE and, unless GRID_ANSWERS_FILE holds its answer, the run stops with NeedsAnswer.
# GRID_REPORT_FILE: the run's report as JSON.  GRID_ALIASES_FILE: names to replace before the master check.  GRID_STANDING_NOTES_FILE: more notes that need no action.
import difflib as _difflib
import hashlib as _hashlib
import json as _json

_SERVICE = {"questions": {}, "answers": None}
ALIAS_USED = []                   # the name mappings this run applied: kind, name in the grid, name written, rows


def service_on():
    return os.environ.get("GRID_SERVICE") == "1"


def _env_json(name, default):
    path = os.environ.get(name)
    if not path:
        return default
    try:
        with open(path, encoding="utf-8") as fh:
            return _json.load(fh)
    except (OSError, ValueError):
        return default


def service_reset():
    _SERVICE["questions"], _SERVICE["answers"] = {}, None


def _fold_text(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def question_fp(kind, problem, labels):
    """The fingerprint of a question: its wording without the numbers and file names that change every month (the web app matches past answers by it)."""
    t = re.sub(r"[\w\-./:()]*\.(?:xlsx|xlsm|xlsb|xls|csv|pdf|json)\b", "<file>", _fold_text(problem))
    t = re.sub(r"\d+", "#", t)
    ls = "|".join(re.sub(r"\d+", "#", _fold_text(l)) for l in labels)
    raw = kind + "|" + t + "|" + ls
    return _hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def question_id(kind, problem, labels):
    raw = kind + "|" + _fold_text(problem) + "|" + "|".join(_fold_text(l) for l in labels)
    return _hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def service_dump():
    path = os.environ.get("GRID_QUESTIONS_FILE")
    if not path:
        return
    try:
        with open(path, "w", encoding="utf-8") as fh:
            _json.dump({"script": os.path.basename(str(globals().get("__file__", ""))), "questions": list(_SERVICE["questions"].values())}, fh, ensure_ascii=False)
    except OSError:
        pass


def service_question(kind, problem, question, short, labels, multi=False, blank_ok=False):
    """Register a question (once per id) and write the questions file."""
    qid = question_id(kind, problem, labels)
    entry = _SERVICE["questions"].get(qid)
    if entry is None:
        entry = dict(id=qid, fp=question_fp(kind, problem, labels), kind=kind, problem=str(problem), question=str(question or problem), short=str(short or ""),
                     multi=bool(multi), blank_ok=bool(blank_ok), options=[dict(n=i, label=str(l)) for i, l in enumerate(labels, 1)], status="pending", answer=None, invalid=False)
        _SERVICE["questions"][qid] = entry
        service_dump()
    return entry


def service_answer(qid):
    if _SERVICE["answers"] is None:
        data = _env_json("GRID_ANSWERS_FILE", {})
        _SERVICE["answers"] = {str(a.get("id")): a for a in ((data or {}).get("answers") or []) if isinstance(a, dict)}
    return _SERVICE["answers"].get(qid)


def service_mark(entry, status, answer=None, invalid=False):
    entry["status"], entry["answer"], entry["invalid"] = status, answer, invalid
    service_dump()


def apply_aliases(df):
    """GRID_ALIASES_FILE = {"rto": {"name in the grid": "master group"}, "makes": {...}}: the person's own mapping of a name the master lacks (made in the web app), applied before the
    master check.  A comma list is mapped name by name.  Returns the number of cells changed."""
    aliases = _env_json("GRID_ALIASES_FILE", {})
    if not isinstance(aliases, dict) or not aliases:
        return 0
    can_dup = all(c in df.columns for c in KEY_COLUMNS)
    dup_before = bool(df.duplicated(KEY_COLUMNS).any()) if can_dup else None
    changed = 0
    for col, kind in (("new_rto_list", "rto"), ("new_vehicle_makes", "makes")):
        table = {Masters._fold(k): v for k, v in (aliases.get(kind) or {}).items() if k and v}
        if not table or col not in df.columns or (kind == "rto" and RTO_NAMES == "cluster"):      # cluster names are not mapped to master groups
            continue

        def swap(value, table=table):
            v = str(value)
            if Masters._fold(v) in table:
                return table[Masters._fold(v)]
            parts = [p.strip() for p in v.split(",")]
            if any(Masters._fold(p) in table for p in parts):
                return ",".join(table.get(Masters._fold(p), p) for p in parts)
            return v
        new = df[col].map(swap)
        diff = new != df[col]
        if diff.any():
            for old in sorted(set(df.loc[diff, col])):
                ALIAS_USED.append(dict(kind=kind, name=old, written=swap(old), rows=int((df[col] == old).sum())))
            df[col] = new
            changed += int(diff.sum())
    if changed and dup_before is False and df.duplicated(KEY_COLUMNS).any():
        RUN_WARNINGS.append("a name mapping made two cards identical (the same key): check the upload before using it")
    return changed


def service_report(frame, csv_path, sheets, ask):
    """GRID_REPORT_FILE: what the run produced, as JSON for the web app (counts, warnings, decisions, the lists of the review workbook)."""
    path = os.environ.get("GRID_REPORT_FILE")
    if not path:
        return

    def records(df, cap=3000):
        return _json.loads(df.head(cap).to_json(orient="records", force_ascii=False))

    def grid_note(text):
        """the grid note (as unknown_notes returned it) that this flag was made from - '' for a flag made some other way (a check of the script itself)"""
        kf = note_key(text)
        if kf in GRID_NOTES:
            return GRID_NOTES[kf]
        return next((t for kt, t in GRID_NOTES.items() if len(kf) >= 40 and kt.startswith(kf)), "")     # a long note is cut short in the flag
    remarks = frame["remarks"].astype(str) if "remarks" in frame.columns else pd.Series([""] * len(frame))
    have = lambda c: c in frame.columns and len(frame) > 0
    data = dict(csv=os.path.basename(csv_path), review=os.path.basename(review_path(csv_path)) if sheets else "", rows=int(len(frame)),
                flagged=int(remarks.str.contains("Manual Review Required", regex=False).sum()),
                from_date=str(frame["from_date"].iloc[0]) if have("from_date") else "", to_date=str(frame["to_date"].iloc[0]) if have("to_date") else "",
                warnings=list(RUN_WARNINGS), decisions=list(getattr(ask, "decisions", [])), aliases_used=list(ALIAS_USED), notes_silenced=list(NOTES_SILENCED.values()),
                notes=[dict(note=t, rows_flagged=int(c), grid_note=grid_note(t)) for t, c in LAST_NOTES],
                products=[dict(product=str(p), sub_product=str(sp), rows=int(n)) for (p, sp), n in frame.groupby(["product", "sub_product"]).size().items()] if {"product", "sub_product"} <= set(frame.columns) else [],
                sheets={name: records(f) for name, f in sheets.items()}, sheet_rows={name: int(len(f)) for name, f in sheets.items()})
    try:
        with open(path, "w", encoding="utf-8") as fh:
            _json.dump(data, fh, ensure_ascii=False)
    except OSError:
        pass


class Asker:
    """Puts questions to the person running the script - in the terminal or in Colab - and remembers the answers.

    With interactive=False nothing is asked: choose() / text() raise NeedsAnswer(problem) instead, so a run that
    cannot ask stops with a message that says what is wrong.  Typing q (or Ctrl+C) at any question stops the run,
    typing ? shows the details of the question."""

    def __init__(self, interactive=False, input_fn=None, out=say):
        self.interactive, self.out, self._input = interactive, out, input_fn or input
        self.memory = {}        # memo key -> the answer, so the same question is never asked twice
        self.decisions = []     # one plain-language line per answer, printed at the end of the run
        self.tallies = {}       # what -> [details], for one answer that covers many cells
        self.last_label = ""
        if service_on():
            service_reset()     # a new run of the web app: the questions of an earlier run are not this run's

    def _read(self, prompt):
        try:
            answer = str(self._input(prompt)).replace("﻿", "").strip()      # ﻿: BOM that PowerShell puts on piped text
        except EOFError:
            raise NeedsAnswer("a question came up but there is no keyboard input to answer it")
        except KeyboardInterrupt:
            raise GridError("stopped by you")
        if answer.lower() in ("q", "quit", "stop"):
            raise GridError("stopped by you")
        return answer

    def choose(self, problem, options, question=None, short=None, memo=None, accept=None, record=True,
               multi=False, check=None):
        """options = [(label, value)]; returns the chosen value - or, with multi=True, the list of chosen values
        (the person types 3,4,5 or 3-5).  `problem` is the message used when no question can be asked; `question`
        is what the person sees; `memo` makes the answer reusable; check(values) may raise ValueError to ask again."""
        if memo is not None and memo in self.memory and (accept is None or accept(self.memory[memo])):
            return self.memory[memo]
        if service_on() and options:
            return self._service_choose(problem, options, question, short, memo, record, multi, check)
        if not self.interactive or not options:
            raise NeedsAnswer(problem)
        self.out(f"\n? {question or problem}")
        for i, (label, _) in enumerate(options, 1):
            self.out(f"   {i}) {label}")
        n = len(options)
        hint = f"Type 1-{n}, or several like 3,4,5 or 3-5" if multi else f"Type 1-{n}"
        while True:
            a = self._read(f"   {hint} (q = stop, ? = details): ")
            if a in ("?", "h", "help"):
                self.out(f"   {problem}")
                continue
            try:
                if multi:
                    picks = parse_picks(a, n)
                elif a.isdigit() and 1 <= int(a) <= n:
                    picks = [int(a)]
                else:
                    raise ValueError(f"Please type a number from 1 to {n}, or q to stop.")
                values = [options[i - 1][1] for i in picks]
                if multi and check:
                    check(values)
                break
            except ValueError as e:
                self.out(f"   {e}")
        self.last_label = "; ".join(options[i - 1][0] for i in picks)
        value = values if multi else values[0]
        if record:
            self.decide(f"{short or question or problem} -> {self.last_label}")
        if memo is not None:
            self.memory[memo] = value
        return value

    def text(self, problem, question=None, validate=None, blank_ok=False):
        """Free text. validate(text) returns the value or raises ValueError(message) to ask again."""
        if service_on():
            return self._service_text(problem, question, validate, blank_ok)
        if not self.interactive:
            raise NeedsAnswer(problem)
        self.out(f"\n? {question or problem}")
        while True:
            a = self._read("   > ")
            if not a and blank_ok:
                return ""
            if a in ("?", "h", "help"):
                self.out(f"   {problem}")
                continue
            try:
                if not a:
                    raise ValueError("Please type an answer, or q to stop.")
                return validate(a) if validate else a
            except ValueError as e:
                self.out(f"   {e}")

    def confirm(self, question, default=True):
        """Yes / no. Without a keyboard the default is used."""
        if not self.interactive:
            return default
        while True:
            a = self._read(f"{question} [{'Y/n' if default else 'y/N'}]: ").lower()
            if not a:
                return default
            if a in ("y", "yes"):
                return True
            if a in ("n", "no"):
                return False
            self.out("   Please type y or n (q = stop).")

    def _service_choose(self, problem, options, question, short, memo, record, multi, check):
        """Service mode: the answer comes from GRID_ANSWERS_FILE ({"id": ..., "pick": [2]}); without it the question is written to GRID_QUESTIONS_FILE and the run stops."""
        entry = service_question("choose", problem, question, short, [label for label, _ in options], multi=multi)
        a = service_answer(entry["id"])
        picks = None
        if isinstance(a, dict) and "pick" in a:
            raw = a["pick"] if isinstance(a["pick"], list) else [a["pick"]]
            try:
                picks = [int(x) for x in raw]
            except (TypeError, ValueError):
                picks = None
            if picks is not None and (not picks or any(i < 1 or i > len(options) for i in picks) or (not multi and len(picks) != 1)):
                picks = None
        values = [options[i - 1][1] for i in picks] if picks else None
        if values is not None and multi and check:
            try:
                check(values)
            except ValueError:
                values = None
        if values is None:
            service_mark(entry, "pending", invalid=a is not None)        # an answer that no longer fits is asked again
            raise NeedsAnswer(problem)
        service_mark(entry, "answered", answer=a)
        self.last_label = "; ".join(options[i - 1][0] for i in picks)
        value = values if multi else values[0]
        if record:
            self.decide(f"{short or question or problem} -> {self.last_label}")
        if memo is not None:
            self.memory[memo] = value
        return value

    def _service_text(self, problem, question, validate, blank_ok):
        entry = service_question("text", problem, question, None, [], blank_ok=blank_ok)
        a = service_answer(entry["id"])
        if isinstance(a, dict) and "text" in a:
            t = str(a["text"]).strip()
            try:
                if not t and blank_ok:
                    service_mark(entry, "answered", answer=a)
                    return ""
                if not t:
                    raise ValueError("empty")
                value = validate(t) if validate else t
                service_mark(entry, "answered", answer=a)
                return value
            except ValueError:
                pass
        service_mark(entry, "pending", invalid=a is not None)
        raise NeedsAnswer(problem)

    def decide(self, text):
        self.decisions.append(text)

    def tally(self, what, detail):
        self.tallies.setdefault(what, []).append(detail)

    def summary(self, log=say):
        if not (self.decisions or self.tallies):
            return
        log("\nWhat you decided in this run:")
        for d in self.decisions:
            log(f"  - {d}")
        for what, items in self.tallies.items():
            more = " ..." if len(items) > 5 else ""
            log(f"  - {what}: {len(items)} (first: {', '.join(items[:5])}{more})")


def parse_picks(text, n):
    """'3,4,5' / '3-5' / '1, 3-4' -> [3, 4, 5] / [3, 4, 5] / [1, 3, 4]: each between 1 and n, no repeats"""
    hint = f"Please type a number from 1 to {n}, or several like 3,4,5 or 3-5, or q to stop."
    picks = []
    for part in re.split(r"[,;\s]+", re.sub(r"\s*-\s*", "-", text.strip())):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not m:
            raise ValueError(hint)
        lo, hi = int(m.group(1)), int(m.group(2) or m.group(1))
        if lo < 1 or lo > hi or hi > n:
            raise ValueError(hint)
        picks += [i for i in range(lo, hi + 1) if i not in picks]
    if not picks:
        raise ValueError(hint)
    return picks


def key_text(key):
    """('Petrol', 'PC < 1000 CC') -> 'Petrol / PC < 1000 CC'"""
    return " / ".join(str(p) for p in (key if isinstance(key, tuple) else (key,)) if p)


def row_number(text, last, first=1):
    """'7' -> 7, checked against the rows the sheet has"""
    try:
        n = int(str(text).strip())
    except ValueError:
        raise ValueError("Please type the row number, for example 4.")
    if not first <= n <= last:
        raise ValueError(f"The sheet has rows {first} to {last} - please type one of those.")
    return n


def _name_key(s):
    return re.sub(r"[^a-z0-9<>=+]+", "", str(s).lower())


def resolve_sheet(names, wanted, ask, optional=False, what="", skip=()):
    """The sheet called `wanted`. If the workbook has no such sheet, ask which one to use
    (or, when `optional`, to leave that part out - returns None). Sheets in `skip` are not offered."""
    names = [n for n in names if n not in skip]
    if wanted in names:
        return wanted
    same = [n for n in names if _name_key(n) == _name_key(wanted)]
    if len(same) == 1:
        return same[0]
    options = [(f"Use the sheet '{n}'", n) for n in difflib.get_close_matches(wanted, names, n=max(1, len(names)), cutoff=0.0)]
    if optional:
        options.append((f"This workbook has no such sheet: leave out {what or wanted}", None))
    return ask.choose(f"sheet '{wanted}' not found in the workbook. Sheets: {names}", options,
                      question=f"I can't find the sheet '{wanted}'" + (f" ({what})" if what else "") + ". What should I do?",
                      short=f"sheet '{wanted}'", memo=("sheet", wanted, tuple(skip)),
                      accept=lambda v: v is None or v in names)


def clean_path(text):
    """A path typed, pasted or dragged into the terminal. Understands what VS Code / PowerShell put there:
    "C:\\a b\\x.xlsx", 'C:\\a b\\x.xlsx', & 'C:\\a b\\July''26\\x.xlsx' (a leading '& ' and an apostrophe written
    as ''), plain paths and file:/// links; expands ~ and %VARS%."""
    t = str(text).replace("\ufeff", "").strip()
    if t.startswith("& "):
        t = t[2:].strip()
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "'\"":
        quote, t = t[0], t[1:-1]
        if quote == "'":
            t = t.replace("''", "'")
    t = t.strip().strip('"').strip("'")
    if t.lower().startswith("file:///"):
        t = t[8:].replace("/", os.sep)
    return os.path.expanduser(os.path.expandvars(t))


def ask_file(ask, problem, question, exts=None, blank_ok=False):
    """Ask for the path of a file that exists; '' when blank_ok and nothing was typed."""
    def check(a):
        p = clean_path(a)
        if not os.path.isfile(p):
            raise ValueError(f"'{p}' is not a file - check the path and try again (q = stop).")
        if exts and not p.lower().endswith(exts):
            raise ValueError(f"expected a file ending in {' / '.join(exts)}")
        return p
    return ask.text(problem, question=question, validate=check, blank_ok=blank_ok)


def open_file(ask, path, what, opener, exts=None):
    """(path, opener(path)). If the file cannot be opened - damaged, password protected, open in Excel, an old
    .xls, no library - ask whether to retry, or to use another file, instead of crashing."""
    while True:
        try:
            return path, opener(path)
        except GridError:
            raise
        except Exception as e:                                                       # BadZipFile, OSError, KeyError, ...
            reason = f"{type(e).__name__}: {e}"
            name = os.path.basename(path)
            pick = ask.choose(f"cannot open {what} {name}: {reason}",
                              [("Try the same file again (for example after closing it in Excel)", "retry"),
                               ("Use a different file (for example the same workbook saved as .xlsx)", "other")],
                              question=f"I can't open {what} '{name}' ({reason}). What now?", record=False)
            if pick == "other":
                path = ask_file(ask, f"no other {what} given", f"Path of the file to use instead of '{name}':", exts=exts)
                ask.decide(f"{what}: could not open '{name}' -> used '{os.path.basename(path)}' instead")


# ---- one cell at a time
def shown(v, limit=60):
    """repr() of a cell value, cut short so a huge text cannot flood the screen"""
    s = repr(v)
    return s if len(s) <= limit else s[: limit - 4] + "...'"


def cell_issue(ask, issue, addr, v, problem, question, options, tally, once=False):
    """A cell that is not a clean rate. The first time a kind of problem (`issue`) turns up it asks what to do
    (options = [(label, action)]); the next time it asks once whether to do the same for this cell and for every
    other cell like it. The same cell is only ever asked about once."""
    seen = ask.memory.setdefault(("issue cells",), {})
    if (issue, addr) in seen:
        return seen[(issue, addr)]
    policy = ask.memory.get(("policy", issue))
    if policy and not policy["all"] and ask.confirm(
            f"{addr} holds {shown(v)}: the same kind of problem as before, where you chose \"{policy['label']}\". "
            f"Do that here and for every other cell like it?", True):
        policy["all"] = True
    if policy and policy["all"]:
        action = policy["action"]
    else:
        action = ask.choose(problem, options, question=question, record=False)
        ask.memory[("policy", issue)] = {"action": action, "label": ask.last_label, "all": once}
    seen[(issue, addr)] = action
    ask.tally(tally, f"{addr} ({shown(v, 30)})")
    return action


def read_number(ask, addr, v, unit, expected, zero_words=(), blank_zero=False, specials=()):
    """One raw grid cell -> (kind, number) in the grid's own unit (unit 'fraction': 0.25 means 25%; unit 'percent':
    25 means 25%).  kind is 'num', 'zero' (a word such as NA / Block, or a blank, meaning no rate), or 'special'
    (one of `specials`, e.g. IRDA, with the word as the value).  Anything that is not a clean rate - text, an
    error such as #N/A, a number that is negative or on the wrong scale - is put to the person."""
    problem = f"{addr}: unexpected value {shown(v)} (expected {expected})"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if math.isfinite(v):
            return _in_range(ask, addr, float(v), unit, problem)
        v = None                                                  # nan / inf behave like an empty cell
    if v is None or (isinstance(v, str) and not v.replace(" ", " ").strip()):
        if blank_zero:
            return "zero", 0.0
        return _unreadable(ask, addr, v, problem, specials, "an empty cell")
    if isinstance(v, str):
        s = v.replace(" ", " ").strip()
        up = s.upper()
        if up in specials:
            return "special", up
        if up in zero_words:
            return "zero", 0.0
        m = re.fullmatch(r"([+-]?\d+(?:[.,]\d+)?)\s*(%?)", s)
        if m:
            act = cell_issue(
                ask, "text-number", addr, v, problem,
                f"{addr} holds {shown(v)}: a number typed as text. Should text like this be read as a number?",
                [(f"Read text like {shown(v)} as a number (a % sign divides by 100 in fraction grids)", "read"),
                 ("Do not - treat text like this as an unreadable cell", "no")],
                "numbers typed as text", once=True)
            if act == "read":
                x = float(m.group(1).replace(",", "."))
                if m.group(2) and unit == "fraction":
                    x /= 100
                return _in_range(ask, addr, x, unit, problem)
    return _unreadable(ask, addr, v, problem, specials, None)


def _unreadable(ask, addr, v, problem, specials, blank_name):
    """Neither a number nor an accepted word: use 0 (or one of the special words), or stop. Asked once per kind of text."""
    token = "(empty)" if blank_name else str(v).strip().upper()
    options = [("Use 0 (no rate) for this cell", "zero")] + [(f"Treat it as {w}", w) for w in specials]
    kind = blank_name or f"{shown(v)}, which is not a rate"
    act = cell_issue(ask, ("unreadable", token), addr, v, problem, f"{addr} holds {kind}. What should I use?",
                     options, "cells that were not rates")
    return ("special", act) if act in specials else ("zero", 0.0)


def _in_range(ask, addr, x, unit, problem):
    """A number: check it is a believable rate for the grid's unit and ask about the odd ones."""
    big = 1.0 if unit == "fraction" else 100.0
    if x < 0:
        act = cell_issue(ask, "negative", addr, x, f"{problem} - a negative rate",
                         f"{addr} holds {x:g}, a negative rate. What should I use?",
                         [("Use 0 (no rate)", "zero"), ("Use the negative number as it is", "keep")],
                         "negative rates")
        return ("zero", 0.0) if act == "zero" else ("num", x)
    if x > big:
        opts = [("Use it as it is", "keep"), ("Use 0 (no rate)", "zero")]
        if unit == "fraction":
            opts.insert(0, ("Read it as a percent: divide by 100 (so 25 means 25%)", "scale"))
        act = cell_issue(ask, "too-big", addr, x, f"{problem} - {x:g} is above {big:g}",
                         f"{addr} holds {x:g}, which is above {big:g} ({'a rate over 100%' if unit == 'fraction' else 'over 100%'}). "
                         f"What should I use?", opts, "rates above the usual range")
        return {"scale": ("num", x / 100), "zero": ("zero", 0.0)}.get(act, ("num", x))
    if unit == "percent" and 0 < x < 1:
        act = cell_issue(ask, "fraction", addr, x, f"{problem} - {x:g} looks like a fraction",
                         f"{addr} holds {x:g}. This grid writes 25% as 25, so {x:g} looks like a fraction ({x * 100:g}%). "
                         f"What should I use?",
                         [("Read it as a fraction: multiply by 100", "scale"), ("Use it as it is", "keep"),
                          ("Use 0 (no rate)", "zero")], "fractions in a percent grid")
        return {"scale": ("num", x * 100), "zero": ("zero", 0.0)}.get(act, ("num", x))
    return "num", x


# ---- dates
_MONTHS = {m: i for i, m in enumerate(("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}


def parse_ddmmyyyy(text):
    """'01/10/2026', '1-10-2026', '01.10.2026', '1 Oct 2026' or '2026-10-01' -> date (day first, 4-digit year).
    Raises ValueError with a readable message for anything else."""
    t = str(text).replace("﻿", "").strip()
    nice = ValueError(f"'{t}' is not a valid date - use DD/MM/YYYY, for example 01/10/2026")
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", t)
    if m:
        y, mo, d = (int(x) for x in m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[-/.\s](\d{1,2})[-/.\s](\d{4})", t)
        if m:
            d, mo, y = (int(x) for x in m.groups())
        else:
            m = re.fullmatch(r"(\d{1,2})[-/.\s]*([A-Za-z]{3,9})[-/.,\s]*(\d{4})", t)
            if not m or m.group(2)[:3].lower() not in _MONTHS:
                raise nice
            d, mo, y = int(m.group(1)), _MONTHS[m.group(2)[:3].lower()], int(m.group(3))
    try:
        return dt.date(y, mo, d)
    except ValueError:
        raise nice


def date_notes(from_date, to_date, today=None):
    """Reasons a pair of dates looks unusual (empty list when they look normal)."""
    today = today or dt.date.today()
    notes = []
    if to_date < today:
        notes.append(f"to_date {to_date:%d/%m/%Y} is already in the past")
    if abs((from_date - today).days) > 400:
        notes.append(f"from_date {from_date:%d/%m/%Y} is more than a year from today")
    span = (to_date - from_date).days + 1
    if span < 28:
        notes.append(f"the period is only {span} day(s) long")
    if span > 366 * 8:
        notes.append(f"the period is {span // 365} years long")
    return notes


def check_dates(from_date, to_date, ask, log=say):
    """True when the dates may be used: odd-looking ones are put to the person (or just noted when nobody can answer)."""
    notes = date_notes(from_date, to_date)
    if not notes:
        return True
    if ask is not None and ask.interactive:
        return ask.confirm("Heads-up: " + "; ".join(notes) + ". Is that what you want?", True)
    log("Note: " + "; ".join(notes))
    return True


def ask_date(label, input_fn=input):
    while True:
        try:
            return parse_ddmmyyyy(input_fn(f"Enter {label} (DD/MM/YYYY): "))
        except ValueError as e:
            say(f"  {e}")
        except EOFError:
            raise GridError(f"{label} was not entered (no keyboard input available) - pass --from-date and --to-date")
        except KeyboardInterrupt:
            raise GridError("stopped by you")


def ask_date_range(input_fn=input, ask=None):
    """Ask for from_date and to_date until both are valid, to_date is not before from_date and - if they look
    unusual (see date_notes) - you have confirmed them."""
    while True:
        from_date = ask_date("from_date", input_fn)
        to_date = ask_date("to_date", input_fn)
        if to_date < from_date:
            say("  to_date must not be before from_date - please enter both again.")
            continue
        say(f"Using from_date {from_date:%d/%m/%Y} and to_date {to_date:%d/%m/%Y}")
        if check_dates(from_date, to_date, ask):
            return from_date, to_date


# ---- output
LAST_LOWEST = []          # entries that were given more than one payin: the lowest was kept (write_csv writes the list next to the CSV)


def payin_value(frame):
    """The payin of every row as one number: OD + TP + Net + flat amount ('On OD and TP' adds its two parts; the other types fill only one of them)."""
    cols = [c for c in ("pi_od_rate", "pi_tp_rate", "pi_net_rate", "pi_flat_amount") if c in frame.columns]
    if not cols:
        return pd.Series(0.0, index=frame.index)
    return frame[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)


_MRR = "Manual Review Required: "


def _remark_parts(text):
    """(the plain text, [the notes]) of a remarks cell: the notes are the '[...]' items after 'Manual Review Required: ' (ManualReview.apply writes them joined with ' | ')."""
    text = "" if text is None or (isinstance(text, float) and text != text) else str(text)
    i = text.find(_MRR)
    if i < 0:
        return text.strip(" |"), []
    return text[:i].strip(" |"), [n for n in re.split(r"(?<=\]) \| (?=\[)", text[i + len(_MRR):]) if n.strip()]


def merge_remarks(keep, dropped):
    """The remarks of the row that stays + whatever the remarks of the rows dropped for it say that its own do not: a flag is never lost with a duplicate (the user, 18 Nov 2026:
    a condition that is not applied must not vanish because another row of the same key was written instead)."""
    before, notes = _remark_parts(keep)
    plain = [before] if before else []
    changed = False
    for d in dropped:
        b, ns = _remark_parts(d)
        if b and b not in plain:
            plain.append(b)
            changed = True
        for n in ns:
            if n not in notes:
                notes.append(n)
                changed = True
    if not changed:
        return keep
    return " | ".join(plain + ([_MRR + " | ".join(notes)] if notes else []))


def resolve_duplicates(df, key_cols, ask):
    """Rows that repeat the same key are one entry given more than once.  The same payin twice is one row; different payins keep the LOWEST - the user's rule (27 Oct 2026):
    "if the rate grid contains two different rates for the same entry, always select the lower rate".  Nothing is asked.  Every entry whose payins differed is listed in
    LAST_LOWEST (write_csv writes the list next to the CSV) and noted among the answers printed at the end of the run.  The row that stays carries the manual-review notes of the rows
    dropped for it (merge_remarks)."""
    global LAST_LOWEST
    LAST_LOWEST = []
    dup = df.duplicated(key_cols, keep=False)
    if not dup.any():
        return df
    value = payin_value(df)
    drop = []
    merged = {}
    for _, idx in df[dup].groupby(key_cols, sort=False).groups.items():
        idx = list(idx)
        best = min(idx, key=lambda i: (value[i], idx.index(i)))          # the lowest payin; of equals the first
        drop.extend(i for i in idx if i != best)
        if "remarks" in df.columns:
            now = merge_remarks(df.at[best, "remarks"], [df.at[i, "remarks"] for i in idx if i != best])
            if now != df.at[best, "remarks"]:
                merged[best] = now
        found = list(dict.fromkeys(f"{value[i]:g}" for i in idx))
        if len(found) > 1:
            row = df.loc[best]
            entry = {lab: str(row[c]) for lab, c in (("product", "product"), ("sub_product", "sub_product"), ("rto", "new_rto_list"), ("makes", "new_vehicle_makes"), ("fuel", "fuel_type"),
                                                     ("make_model_class", "make_model_class"), ("tariff_max", "tariff_max")) if c in df.columns}
            entry.update(rates_found=" / ".join(found), written=f"{value[best]:g}")
            LAST_LOWEST.append(entry)
    if LAST_LOWEST:
        ask.decide(f"{len(LAST_LOWEST)} entries were given more than one payin: the lowest was written (listed in the review workbook next to the upload)")
    if merged:
        df = df.copy()
        for i, text in merged.items():
            df.at[i, "remarks"] = text
        ask.decide(f"{len(merged)} entries kept the manual-review notes of the rows that were dropped for them")
    return df.drop(index=drop).reset_index(drop=True)


def write_csv(frame, path, ask, log=say):
    """Write the upload CSV (the portal takes only .csv) and, next to it, the review workbook (write_review); if the CSV cannot be written (open in Excel, no such drive, no permission) ask what to do."""
    frame = apply_cluster_names(frame)
    while True:
        try:
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            frame.to_csv(path, index=False, lineterminator="\n", encoding="utf-8")
            break
        except OSError as e:
            pick = ask.choose(f"cannot write {path}: {e}",
                              [("Try again (for example after closing the file in Excel)", "retry"),
                               ("Write to a different file name", "other")],
                              question=f"I can't write '{path}' ({e}). What now?", record=False)
            if pick == "other":
                path = ask.text("no other output name given", question="New file name (or path) for the CSV:",
                                validate=lambda a: clean_path(a) if clean_path(a).lower().endswith(".csv") else clean_path(a) + ".csv")
                ask.decide(f"output file -> {path}")
    write_review(frame, path, ask, log)
    return path


def compare_safely(ask, differ, df, ref_path, log=say):
    """Run the comparison with an earlier upload; if that fails (wrong file, odd columns) ask whether to skip it."""
    while ref_path:
        try:
            return differ(df, ref_path, log)
        except GridError:
            raise
        except Exception as e:
            if not ask.interactive:
                log(f"Comparison with {os.path.basename(ref_path)} skipped: {type(e).__name__}: {e}")
                return None
            pick = ask.choose(f"cannot compare with {ref_path}: {e}",
                              [("Carry on without the comparison", "skip"), ("Use a different earlier upload", "other")],
                              question=f"The comparison with '{os.path.basename(ref_path)}' failed ({type(e).__name__}: {e}). What now?",
                              record=False)
            if pick == "skip":
                return None
            ref_path = ask_file(ask, "no other file given", "Path of the earlier upload (.csv) to compare with:", exts=(".csv",))
    return None


# ---- where the output goes, and the ONE review workbook written next to the upload CSV
EXTRA_SHEETS = {}         # sheet name -> DataFrame: lists the person should look at (groups to create, grid rows not written ...); reset by ManualReview()
RUN_WARNINGS = []         # the warnings of this run, for the Summary sheet
LAST_NOTES = []           # (note, rows flagged) of this run's manual-review report
NOTES_SILENCED = {}       # note key -> note text: the notes the person's own "no action needed" list (web app) kept from being flagged in this run
GRID_NOTES = {}           # note key -> note text: the grid notes unknown_notes handed back in this run (the only flags a "no action needed" entry can silence)
NOTES_SEEN = {}           # note key -> dict(text, status, why): every note unknown_notes looked at in this run - the review sheet 'Notes read' (a script that writes its own 'Notes read' sheet keeps it)


def reset_run_lists():
    global EXTRA_SHEETS, RUN_WARNINGS, LAST_NOTES, _RESOLVED
    EXTRA_SHEETS, RUN_WARNINGS, LAST_NOTES = {}, [], []
    _RESOLVED = None
    ALIAS_USED.clear()
    NOTES_SILENCED.clear()
    GRID_NOTES.clear()
    NOTES_SEEN.clear()
    CLUSTER_OF.clear()


# ---- RTO names (the user's rule of 30 Oct 2026): new_rto_list carries ONLY the cluster name as the grid writes it ("West UP Open", "DL_Delhi", "GJ_ABS"), never the
# portal's group name (insurer + month + segment prefix) and never single RTO codes. The person maps the clusters and groups the RTOs and makes in the portal himself, so the
# RTO names are not checked against the RTO master any more. "group" (GRID_RTO_NAMES=group) gives the earlier behaviour: the master's group names, checked and flagged.
RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "cluster")      # "cluster" | "group" | "master" (Shriram: the master's name, or a plain-text warning)
CLUSTER_OF = {}           # the name the script builds for a cluster (group name, code list) -> the cluster name of the grid ("GO-DIGIT_OCT26_HCV_Mumbai" -> "Mumbai")

# ---- the person's RTO working list (the user, 31 Oct 2026, Shriram and Go Digit only: "I will prepare a working list for RTO and attach it directly within the grid. Please use those exact
# records and columns"). A column headed RTO (or one of WORKING_HEADERS) next to a rate table holds the RTO name of each grid row. The converter writes it into new_rto_list EXACTLY as typed:
# no master lookup, no condition text, no check against the master. A row whose cell is blank (or says NA) keeps today's lookup; a table without the column behaves as it always did.
WORKING_HEADERS = ("rto", "rtoworkinglist", "rtoworking", "rtolist", "rtoname", "workingrto")      # label_key of the header texts ("RTO Cluster" and "RTO Code" are other columns)
WORKING_NONE = ("na", "n/a", "nil", "none", "-", "--")                                              # a cell that says there is no record
# an Excel error shown in a cell (a formula whose sheet or cell is gone: `=#REF!`) is not a name either (the user's November 2026 grid had 52 of them: they were written as the RTO list `#REF!`)
WORKING_ERROR = re.compile(r"#(?:REF!|N/A|VALUE!|NAME\?|DIV/0!|NULL!|NUM!|SPILL!|CALC!|GETTING_DATA)", re.I)
WORKING_RTO = set()       # the names written from the column in this run: not checked against the master, never rewritten
WORKING_SEEN = []         # one entry per table read: dict(sheet, header, with_value, blank = [Excel rows with no value], unused = [(Excel row, why)])
WORKING_ERRORS = []       # the Excel errors found instead of a name in this run (one text per cell)


def working_header(cell):
    """True when a header cell names the working-list column"""
    return re.sub(r"[^a-z0-9]+", "", str(cell if cell is not None else "").lower()) in WORKING_HEADERS     # its own key: label_key means another thing in some scripts


def working_value(cell):
    """the text of a working-list cell as typed (outer spaces and line breaks trimmed, spaces around a comma dropped, a whole number without '.0'); '' for a blank cell, a 'no record' mark
    or an Excel error (`#REF!`: counted in WORKING_ERRORS and said in the run line)"""
    if cell is None or cell != cell:
        return ""
    if isinstance(cell, float) and cell.is_integer():
        cell = int(cell)
    text = re.sub(r"\s*,\s*", ",", re.sub(r"\s*[\r\n]+\s*", " ", str(cell)).strip())          # 'MH01, MH02' is the list MH01,MH02
    if WORKING_ERROR.fullmatch(text):
        WORKING_ERRORS.append(text.upper())
        return ""
    return "" if text.lower() in WORKING_NONE else text


def working_rto(text):
    """the name of a row from the working list: remembered so that nothing rewrites or checks it, and returned as it is"""
    WORKING_RTO.add(text)
    return text


def working_reset():
    """a new run: forget the names and the tables of the last one (called before the grid is read)"""
    WORKING_RTO.clear()
    WORKING_SEEN.clear()
    WORKING_ERRORS.clear()


def working_table(sheet, header, with_value, blank):
    """remember one table that was read: `header` = the column's header text, or None when the table has no such column; `with_value` = grid rows that carry a name; `blank` = Excel rows with none"""
    entry = dict(sheet=sheet, header=header, with_value=with_value, blank=list(blank), unused=[])
    WORKING_SEEN.append(entry)
    return entry


def working_unused(sheet, row, why):
    """a grid row of `sheet` whose working-list name could not be used (said in the run line)"""
    for e in WORKING_SEEN:
        if e["sheet"] == sheet and e["header"]:
            if (row, why) not in e["unused"]:
                e["unused"].append((row, why))
            return


def working_note():
    """the run line about the working list; '' when no table of this run has the column (the converter then works exactly as before)"""
    if not any(e["header"] for e in WORKING_SEEN):
        return ""
    parts = []
    for e in WORKING_SEEN:
        if not e["header"]:
            parts.append(f"'{e['sheet']}' has no RTO working-list column: the master lookup is used for all its rows")
            continue
        bit = f"'{e['sheet']}': {e['with_value']:,} grid row(s) carry a name in column '{e['header']}' (written exactly as typed)"
        if e["blank"]:
            bit += (f"; {len(e['blank']):,} row(s) have no value there and use the master lookup (Excel row {', '.join(str(r) for r in e['blank'][:8])}"
                    f"{' ...' if len(e['blank']) > 8 else ''})")
        if e["unused"]:
            bit += "; not used for " + "; ".join(f"row {r} ({why})" for r, why in e["unused"][:5]) + (" ..." if len(e["unused"]) > 5 else "")
        parts.append(bit)
    if WORKING_ERRORS:
        kinds = ", ".join(sorted(set(WORKING_ERRORS)))
        parts.append(f"{len(WORKING_ERRORS):,} cell(s) of the column hold an Excel error ({kinds}) instead of a name - read as no record, those rows use the master lookup (fix or clear the cells in the grid)")
    return "RTO working list - " + " | ".join(parts)


def cluster_label(text):
    """A cluster name as the grid writes it: line breaks and runs of spaces become one space, and a comma becomes '_' (new_rto_list is itself a comma-separated list)."""
    return re.sub(r"\s+", " ", str(text)).strip().replace(",", "_")


def rto_cluster(name, cluster):
    """Remember that `name` (what the script builds for new_rto_list) stands for the grid's `cluster`; returns `name`, so it wraps the line that builds it."""
    CLUSTER_OF[name] = cluster_label(cluster)
    return name


class RtoScope(str):
    """The RTO scope of a row (the string is the cluster name or the state codes) together with the condition that narrows it: Shriram's 'all RTOs of MH except ...',
    United's 'only the preferred RTOs of Sub Annexure 3'. `rto_cell` writes both into new_rto_list (user's rule of 30 Oct 2026: for the complex grids the RTO cell holds the cluster
    name and its condition in plain English, with the part, the product and the sub-product it applies to)."""

    def __new__(cls, name, condition="", part=""):
        s = super().__new__(cls, name)
        s.condition, s.part = condition, part
        return s


def plain_list(items):
    """['MH02', 'MH01'] -> 'MH01/MH02': codes in one phrase WITHOUT commas (new_rto_list is itself a comma-separated list)"""
    return "/".join(sorted({str(i).strip() for i in items if str(i).strip()}))


def _tidy(s):
    """a sentence for the RTO cell: no commas (they would split new_rto_list), one space between words"""
    return re.sub(r"\s+", " ", re.sub(r"\s*,\s*", " / ", str(s))).strip()


def rto_cell(scope):
    """'<cluster> - Condition: <plain English> - Part: <the exact part>' for a scope that carries a condition (cluster mode); the plain name otherwise. `apply_cluster_names` adds what the
    row itself says (product, sub-product, vehicle details ...) when the file is written, after every later change of the row."""
    cond, part = getattr(scope, "condition", ""), getattr(scope, "part", "")
    if RTO_NAMES != "cluster" or not (cond or part):
        return str(scope)
    return f"{scope} - Condition: {_tidy(cond)} - Part: {_tidy(part)}"


def _n(x):
    try:
        return f"{float(x):.10g}"
    except (TypeError, ValueError):
        return str(x).strip()


def rto_details(row):
    """What the condition of an RTO cell applies to, in plain words, from the row's own columns (the user, 30 Oct 2026: 'the exact part, product and sub-product ... and more details like cc
    min and max, sc min and max'): product, sub-product, policy type, fuel, vehicle age, engine cc, seating capacity, make group, class, discount band, NCB / CPA / zero dep, extra terms.
    A value that says nothing ('NA', blank) is left out.  Joined with ' - '; no commas."""
    g = lambda k: ("" if row.get(k) is None else str(row.get(k)).strip())
    span = lambda lo, hi, unit="": f"{_n(row.get(lo))} to {_n(row.get(hi))}{unit}" if g(lo) != "" and g(hi) != "" else ""
    items = [("Product", g("product")), ("Sub-product", g("sub_product")), ("Policy type", g("policy_type")), ("Fuel", g("fuel_type")),
             ("Vehicle age", span("vehicle_age_min", "vehicle_age_max", " years")), ("Engine cc", span("cc_min", "cc_max")), ("Seating capacity", span("sc_min", "sc_max")),
             ("Make group", g("new_vehicle_makes")), ("Make class", g("make_model_class")), ("Discount band", span("tariff_min", "tariff_max")),
             ("NCB", g("is_ncb")), ("CPA", g("is_cpa")), ("Zero dep", g("is_zd")), ("Extra terms", g("add_tnc"))]
    return " - ".join(f"{k}: {_tidy(v)}" for k, v in items if v and v.upper() != "NA")


def apply_cluster_names(frame):
    """The CSV frame with new_rto_list = the cluster names (RTO_NAMES == "cluster"); a copy, the frame the script keeps is not touched. A value that is a comma list
    (state codes) is mapped name by name; a name nothing registered stays as it is. A script that defines `rto_condition(value, product, sub_product)` (United: the groups of
    the annexures) gets the plain-English condition written after the name as well (`rto_cell`)."""
    if RTO_NAMES != "cluster" or "new_rto_list" not in frame.columns:
        return frame
    out = frame.copy()
    strip = globals().get("RTO_STRIP")                      # a script whose names come from the master / a table (Shriram, United) gives the regex of its insurer prefix

    def one(t):
        t = t.strip()
        if t in CLUSTER_OF:
            return CLUSTER_OF[t]
        return cluster_label(re.sub(strip, "", t)) if strip and re.match(strip, t) else t
    table = {}
    for v in out["new_rto_list"].astype(str).unique():
        table[v] = v if v in WORKING_RTO else (CLUSTER_OF[v] if v in CLUSTER_OF else ",".join(one(t) for t in v.split(",")))
    original = out["new_rto_list"].astype(str)
    out["new_rto_list"] = original.map(table)
    hook = globals().get("rto_condition")
    if hook and {"product", "sub_product"} <= set(out.columns):
        combos = pd.DataFrame({"v": original, "p": out["product"].astype(str), "s": out["sub_product"].astype(str)}).drop_duplicates()
        text = {}
        for v, p, s in combos.itertuples(index=False):
            got = hook(v, p, s)
            if got:
                text[(v, p, s)] = rto_cell(RtoScope(table[v], *got))
        if text:
            out["new_rto_list"] = [text.get((v, p, s), cur) for v, p, s, cur in zip(original, out["product"].astype(str), out["sub_product"].astype(str), out["new_rto_list"])]
    has = out["new_rto_list"].str.contains(" - Condition: ", regex=False)
    if has.any():                                           # the condition applies to this row's product, sub-product, vehicle details ...: say so, from the row as it is written
        out.loc[has, "new_rto_list"] = [f"{cell} - {rto_details(row)}" for cell, (_, row) in zip(out.loc[has, "new_rto_list"], out.loc[has].iterrows())]
    keys = [c for c in globals().get("KEY_COLUMNS", []) if c in out.columns]
    if keys and out.duplicated(keys).sum() > frame.duplicated(keys).sum():
        RUN_WARNINGS.append(f"{int(out.duplicated(keys).sum() - frame.duplicated(keys).sum())} rows share one key now that new_rto_list holds only the cluster name "
                            f"(two portal groups had the same cluster name): look at them before uploading")
    return out


def add_sheet(name, frame):
    """A list for the review workbook (one sheet of that name); frames added under one name are stacked."""
    if frame is None or not len(frame):
        return
    EXTRA_SHEETS[name] = pd.concat([EXTRA_SHEETS[name], frame], ignore_index=True) if name in EXTRA_SHEETS else frame.reset_index(drop=True)


# ---- what no reader used (the user, 18 Nov 2026: "at some places there is some make or other conditions which is not applied, also it is not mentioned anywhere"): the review sheet 'Unread parts'
UNREAD_COLUMNS = ["sheet", "where", "holds", "cells", "status", "why"]


def note_unread(sheet, where, holds, cells, status, why):
    """One line of the review sheet 'Unread parts': a part of the grid (a whole sheet, columns, rows) that is NOT in the upload - status 'NOT READ' (nothing reads it: also a warning) or
    'left out on purpose' (a decision of the script, with its reason)."""
    add_sheet("Unread parts", pd.DataFrame([dict(sheet=sheet, where=where, holds=holds, cells=cells, status=status, why=why)], columns=UNREAD_COLUMNS))


def describe_rows(rows, head=8):
    """(the short texts at the top, the long texts) of a sheet's rows: what the person reading 'Unread parts' needs to recognise the sheet"""
    top, notes = [], []
    for i, row in enumerate(rows):
        for v in row:
            t = re.sub(r"\s+", " ", str(v)).strip() if isinstance(v, str) else ""
            if not t:
                continue
            if i < head and len(t) < 25 and t not in top and len(top) < 10:
                top.append(t[:48])
            elif len(t) >= 25 and t not in notes and len(notes) < 3:
                notes.append(t[:160])
    return top, notes


def unread_sheet(name, rows, status, why, warnings=None):
    """A whole sheet that no reader of this script used: 'Unread parts' (and, for 'NOT READ', a warning - the run says it).  `rows` = the sheet's rows (lists of cell values)."""
    cells = sum(1 for r in rows for v in r if v is not None and str(v).strip() != "")
    if not cells:
        return
    top, notes = describe_rows(rows)
    holds = "; ".join(top[:10] + [f"text: {n}" for n in notes[:2]])
    note_unread(name, "the whole sheet", holds, cells, status, why)
    if status == "NOT READ" and warnings is not None:
        warnings.append(f"sheet '{name}' ({cells:,} filled cells; it says {holds[:160]}) is not read by this script: nothing from it is in the upload")


def default_output(filename):
    """output\\<insurer>\\<filename> next to the scripts, one folder per insurer (in Colab the current folder: the file is downloaded)."""
    if IN_COLAB:
        return filename
    here = os.path.dirname(os.path.abspath(globals().get("__file__", "scripts")))
    return os.path.join(here, "output", OUTPUT_FOLDER, filename)


def review_path(csv_path):
    """SHRIRAM_UPLOAD_2026-10.csv -> SHRIRAM_REVIEW_2026-10.xlsx (any other name: <name>_review.xlsx)"""
    base = os.path.splitext(csv_path)[0]
    return (base.replace("_UPLOAD", "_REVIEW") if "_UPLOAD" in base else base + "_review") + ".xlsx"


def write_review(frame, csv_path, ask, log=say):
    """ONE workbook next to the upload CSV with everything the person should look at, a sheet each: Summary, the RTO / make groups to create, the grid rows that gave no row,
    the entries where the lowest of different payins was written, the manual-review notes.  Written only when there is something to show; if it cannot be written the CSV is still complete."""
    sheets = dict(EXTRA_SHEETS)
    if LAST_LOWEST:
        sheets["Lowest rate chosen"] = pd.DataFrame(LAST_LOWEST)
    if LAST_NOTES:
        sheets["Manual review notes"] = pd.DataFrame(sorted(LAST_NOTES, key=lambda x: -x[1]), columns=["note", "rows_flagged"])
    if NOTES_SEEN and "Notes read" not in sheets:                          # every note the run looked at and what became of it (flagged ones first)
        sheets["Notes read"] = pd.DataFrame(sorted(NOTES_SEEN.values(), key=lambda d: (not d["status"].startswith("FLAGGED"), d["status"], d["text"])), columns=["text", "status", "why"])
    service_report(frame, csv_path, sheets, ask)
    if not sheets:
        return None
    flagged = int(frame["remarks"].astype(str).str.contains("Manual Review Required", regex=False).sum()) if "remarks" in frame.columns else 0
    dates = f"{frame['from_date'].iloc[0]} - {frame['to_date'].iloc[0]}" if {"from_date", "to_date"} <= set(frame.columns) and len(frame) else ""
    info = [("Upload THIS file (the portal accepts only .csv)", os.path.basename(csv_path)), ("Rows", f"{len(frame):,}"), ("from_date - to_date", dates),
            ("Rows flagged for manual review", f"{flagged:,}"), ("Sheets in this workbook", ", ".join(["Summary"] + list(sheets)))]
    info += [("Warning", w) for w in RUN_WARNINGS] + [("Decided in this run", d) for d in getattr(ask, "decisions", [])]
    path = review_path(csv_path)
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            pd.DataFrame(info, columns=["item", "value"]).to_excel(xw, sheet_name="Summary", index=False)
            for name, f in sheets.items():
                f.to_excel(xw, sheet_name=name[:31], index=False)
            for ws in xw.book.worksheets:
                for col in ws.columns:
                    width = max((len(str(c.value)) for c in col[:300] if c.value is not None), default=8)
                    ws.column_dimensions[col[0].column_letter].width = max(10, min(width + 2, 90))
                ws.freeze_panes = "A2"
    except OSError as e:
        log(f"The review workbook could not be written ({e}); the upload CSV is complete.")
        return None
    log(f"Review workbook ({len(sheets) + 1} sheets: {', '.join(['Summary'] + list(sheets))}): {path}")
    return path


# ---- text rules that cannot be applied
class ManualReview:
    """Notes and text rules the script could not turn into logic. It never guesses at them, never asks about them and never
    stops for them: the rows they may affect get  Manual Review Required: [the note text]  in the `remarks` column of the
    upload, so each can be taken up with the insurer and those rates mapped by hand afterwards."""
    PREFIX = "Manual Review Required: "

    def __init__(self):
        self.items = []             # (note text, the rows it may affect: None = every row, a bool array, or a function of the frame)
        reset_run_lists()           # a new run: the review workbook starts empty

    def flag(self, text, rows=None):
        text = re.sub(r"\s+", " ", str(text)).strip()
        if text and fold_key(text) in resolved_keys():          # the person explained it or said it needs no action (web app memory): never flagged again, whatever number it carries
            NOTES_SILENCED.setdefault(note_key(text), text)
            return
        if text:
            self.items.append((text, rows))

    def apply(self, df):
        """Write the flags into df['remarks'] and return ([(note text, rows flagged)], rows flagged in all)."""
        n = len(df)
        tags, counts = [""] * n, {}
        for text, rows in self.items:
            mask = np.ones(n, dtype=bool) if rows is None else np.asarray(rows(df) if callable(rows) else rows, dtype=bool)
            tag = f"[{text}]"
            hit = [i for i in np.flatnonzero(mask) if tag not in tags[i]]
            for i in hit:
                tags[i] = (tags[i] + " | " if tags[i] else "") + tag
            counts[text] = counts.get(text, 0) + len(hit)
        existing = df["remarks"].fillna("").astype(str).tolist() if "remarks" in df.columns else [""] * n
        df["remarks"] = [(e.strip() + " | " if e.strip() else "") + self.PREFIX + t if t else e for e, t in zip(existing, tags)]
        return list(counts.items()), sum(1 for t in tags if t)

    def report(self, df, log=say):
        """apply() and say in plain words what was flagged. Returns the number of rows flagged."""
        global LAST_NOTES
        counts, flagged = self.apply(df)
        shown = [(t, c) for t, c in counts if c]
        LAST_NOTES = list(shown)
        if shown:
            log(f"\nMANUAL REVIEW - {flagged:,} row(s) now say 'Manual Review Required' in the remarks column "
                f"({len(shown)} note(s) this script could not apply; it did not guess):")
            for t, c in shown:
                log(f"  {c:>8,} rows: {t[:220]}{'...' if len(t) > 220 else ''}")
        return flagged


def note_key(text):
    """A note without spaces, punctuation and capitals, to compare it with known phrases."""
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


_STATE_CODE = r"\b(?:AN|AP|AR|AS|BR|CG|CH|DD|DL|DN|GA|GJ|HP|HR|JH|JK|KA|KL|LA|LD|MH|ML|MN|MP|MZ|NL|OD|OR|PB|PY|RJ|SK|TG|TN|TR|TS|UA|UK|UP|WB)\s*-?\s*\d{1,3}\b"
_CODE_LIST = r"\b\d{1,3}(?:\s*[,/&]\s*\d{1,3}){2,}\b"
_RESOLVED = None            # the keys of the notes the person resolved in the web app (GRID_STANDING_NOTES_FILE), read once per run


def fold_note(text):
    """A note with its numbers folded to '#': 'PO will be 5% less' and 'PO will be 7% less' are one note (the user, 31 Oct 2026: the condition is the same, only the percentage changes).
    The numbers that ARE the condition stay: an RTO code ('RJ-14', 'MH 12') and a list of three or more numbers ('EXL RTO codes 03, 06, 18'), so a different RTO is a different note."""
    t = re.sub(r"\s+", " ", str(text))
    keep = []

    def hold(m):
        keep.append(m.group(0))
        return "\x01" + "k" * len(keep) + "\x01"
    t = re.sub(_STATE_CODE, hold, re.sub(_CODE_LIST, hold, t))
    t = re.sub(r"\d+(?:\.\d+)?", "#", t)
    return re.sub(r"\x01(k+)\x01", lambda m: keep[len(m.group(1)) - 1], t)


def fold_key(text):
    return note_key(fold_note(text))


def resolved_keys():
    """{fold_key} of the notes in GRID_STANDING_NOTES_FILE: what the person explained or said needs no action. Empty outside the web app."""
    global _RESOLVED
    path = os.environ.get("GRID_STANDING_NOTES_FILE", "")
    stamp = (path, os.path.getmtime(path), os.path.getsize(path)) if path and os.path.exists(path) else (path, 0, 0)
    if _RESOLVED is None or _RESOLVED[0] != stamp:                      # read again when the file changed (a long-lived process, the tests)
        _RESOLVED = (stamp, {fold_key(x) for x in (_env_json("GRID_STANDING_NOTES_FILE", []) or [])})
    return _RESOLVED[1]


NOTE_MIN_CHARS = 25         # a text shorter than this is a label ("Priority Cities", "UPTO 50K", "PC STP GRID"), not a rule ...
SHORT_RULE_CHARS = 8        # ... unless it is two words or more and SAYS it is one: a decline / exclusion word makes even a short text a rule ("In PCV 3W EV Decline", 20 characters, was dropped in silence in SBI)
SHORT_RULE_WORDS = re.compile(r"\b(?:declin\w*|not\s+allowed|not\s+applicable|not\s+eligible|exclud\w*|except|restricted|blocked|banned|discontinued|withdrawn)\b", re.I)


# What a note says BEYOND the standing phrase that covers it (the user, 18 Nov 2026: "at some places there is some make or other conditions which is not applied, also it is not mentioned anywhere").  A standing
# phrase is matched by CONTAINMENT, so 'CPA Premium will not be considered for outgo if less than Rs 450' also covered 'CPA Premium ... Rs 450;Except for TW' and 'The following make and model combinations
# will be considered under GCV 20 to 40T with a highest GVW of up to 43500T' - the exception and the 43500 were decided by nobody.  The extra text is looked at: a number or a word of exception / limit in it
# makes the note a NEW note (flagged); other extra words are listed in the review sheet 'Notes read' with the standing phrase and the extra text.
STANDING_EXTRA_RX = re.compile(r"\d|\b(?:except\w*|exclud\w*|excl|only|not|no|never|unless|if|declin\w*|block\w*|exempt\w*|above|below|upto|up\s+to|more\s+than|less\s+than|other\s+than|max\w*|min\w*|without|apart|allowed|eligible)\b", re.I)


def standing_rest(text, phrases):
    """(the standing phrases that really sit in `text`, what the text says beyond them).  A phrase is found as WORDS (spaces and punctuation between them do not matter, a word may not start in the
    middle of another one), so 'Note' is not found in 'not eligible' although the letters are there once the spaces are taken out.  The rest is '' when nothing is left but a list number or a
    colon; None when no phrase is found at all (the key matched by accident)."""
    rest, found = text, []
    for ph in sorted(phrases, key=lambda x: -len(str(x))):
        words = re.findall(r"[a-z0-9]+", str(ph).lower())
        if not words:
            continue
        m = re.search(r"(?<![a-z0-9])" + r"[^a-z0-9]*".join(re.escape(w) for w in words), rest, re.I)
        if m:
            found.append(ph)
            rest = rest[:m.start()] + " " + rest[m.end():]
    if not found:
        return [], None
    rest = re.sub(r"^\s*\d{1,3}\s*[.)]\s*", "", rest)                     # the note's own number: '10.'
    rest = re.sub(r"[^A-Za-z0-9%.+<>=/&-]+", " ", rest).strip(" .,;:-")
    return found, (rest if len(re.sub(r"[^A-Za-z0-9]", "", rest)) >= 4 else "")


def see_note(key, text, status, why):
    """remember a note and what became of it for the review sheet 'Notes read' (the first decision about a note stands)"""
    NOTES_SEEN.setdefault(key, dict(text=text[:400], status=status, why=why[:400]))


def unknown_notes(texts, standing, min_len=NOTE_MIN_CHARS):
    """The texts (cells long enough to be a rule rather than a label, or short but saying decline / exclude / not allowed) that none of the `standing` phrases accounts for.
    `standing` lists notes already known not to need action: a disclaimer, a rule that is already inside the grid's cells,
    a rule the portal handles in its make / RTO groups. Anything else is a new text rule and gets flagged.  A standing phrase that sits INSIDE a longer note covers only its own words:
    when the rest of the note holds a number or a word of exception / limit (STANDING_EXTRA_RX) the note is flagged like any other (standing_extra).  Every note looked at is remembered in
    NOTES_SEEN (the review sheet 'Notes read')."""
    topics = [s for s in standing if isinstance(s, re.Pattern)]            # a compiled pattern = a TOPIC: every note it finds is a standing decision as a whole
    known = [(note_key(s), s) for s in standing if not isinstance(s, re.Pattern)]
    known = [(k, s) for k, s in known if k]                                # a phrase with no letters would match every note
    why_of = globals().get("STANDING_WHY", {})                             # an optional {standing phrase or pattern: why it needs no action}: shown in 'Notes read'
    approved = resolved_keys()                              # notes the person resolved in the web app: that text (numbers folded), never a longer rule that merely contains it
    out, seen = [], set()
    for t in texts:
        t = re.sub(r"\s+", " ", str(t)).strip()
        k = note_key(t)
        if len(t) < min_len and not (len(t) >= SHORT_RULE_CHARS and len(t.split()) >= 2 and SHORT_RULE_WORDS.search(t)):         # one word ('Declined') is a cell value, not a note
            continue
        if k in seen:
            continue
        topic = next((tp for tp in topics if tp.search(t)), None)
        if topic is not None:
            see_note(k, t, "no action (standing decision about the whole topic)", f"standing topic: {topic.pattern[:100]}" + (f" | {why_of[topic]}" if topic in why_of else ""))
            continue
        if globals().get("STANDING_STRICT", False):                          # an audited script (18 Nov 2026): a standing phrase covers its own words only
            found, extra = standing_rest(t, [sp for sk, sp in known if sk in k])
        else:                                                                # a script not audited yet: the old rule, a standing phrase covers every note that contains it
            old_hit = [sp for sk, sp in known if sk in k]
            found, extra = ([max(old_hit, key=lambda x: len(note_key(x)))], "") if old_hit else ([], None)
        flagged_extra, phrase = False, (found[0] if found else "")
        if found:
            if not extra:
                see_note(k, t, "no action (standing decision)", f"standing note: {phrase[:160]}" + (f" | {why_of[phrase]}" if phrase in why_of else ""))
                continue
            if not STANDING_EXTRA_RX.search(extra):
                see_note(k, t, "no action (standing decision; the note says more, with no condition in it)", f"standing note: {phrase[:100]} | the rest of the note: {extra[:200]}")
                continue
            flagged_extra = True
        if fold_key(t) in approved:
            NOTES_SILENCED.setdefault(k, t)
            see_note(k, t, "no action (you resolved it in the web app)", "explained or marked 'no action needed' in the app's memory")
            continue
        out.append(t)
        seen.add(k)
        GRID_NOTES.setdefault(k, t)
        see_note(k, t, "FLAGGED: says more than the standing note that matches it" if flagged_extra else "FLAGGED: no rule of this script applies it",
                 (f"the standing note '{phrase[:100]}' covers only its own words; the rest of the note is a condition nobody has decided on: {extra[:200]}" if flagged_extra
                  else "Manual Review Required on the rows it may affect"))
    return out


# ---- the portal's master lists: the authority for RTO group and make / model group names
# Two workbooks exported from the portal (RTO name -> RTO codes, make model name -> make model entries) live in the `master` folder next to the
# scripts; the newest file whose name has "rto" + "master" (resp. "make" + "model" + "master") is used, so a fresh download can simply be dropped in.
# A name a script writes into new_rto_list / new_vehicle_makes that is not in them is flagged on its rows - never replaced by a guess.
MASTER_DIR = os.path.join(os.path.dirname(os.path.abspath(globals().get("__file__", "scripts"))), "master")


class Masters:
    """name -> list of RTO codes (rto) / of make model entries (makes), as exported from the portal."""

    def __init__(self, rto, makes, sources=()):
        self.rto, self.makes, self.sources = rto, makes, tuple(sources)
        self._loose = {"rto": {self._fold(k): k for k in rto}, "makes": {self._fold(k): k for k in makes}}

    @staticmethod
    def _fold(name):
        return str(name).strip(" _").lower()

    def near(self, kind, name):
        """the master's spelling of a name that differs from it only in capitals or in a stray leading / trailing '_' or space
        (a hint for the person, never applied)"""
        found = self._loose[kind].get(self._fold(name))
        return found if found and found != name else None

    def codes(self, name):
        return set(self.rto.get(name, ()))


# ---- finding a table by what it SAYS, not by where it is (the user, 30 Oct 2026: insurers move things around - rows, columns, sheet names - and the output must not change).
# A table is the row among the first rows whose cells say the wanted headers ("Product", "RTO Code" ...), in any column, below any title rows; its sheet is found by those
# headers too (the sheet's name only breaks a tie). A table that cannot be found, or is found twice, is NEVER silently skipped: it is asked about, or reported.
def _hwords(text):
    return re.findall(r"[a-z0-9]+", str(text if text is not None else "").lower())


def head_match(cell, synonym):
    """True when a header cell says `synonym`: the same words (spaces, punctuation and capitals ignored), or all the synonym's words in a short cell ('RTO State' says 'state')."""
    if not isinstance(cell, str) or not cell.strip():
        return False
    w, s = _hwords(cell), _hwords(synonym)
    return bool(s) and ("".join(w) == "".join(s) or (set(s) <= set(w) and len(w) <= len(s) + 3))


def locate_table(rows, fields, required, scan=15):
    """(header row index, {field: column}) of the best header row among the first `scan` rows: the row that says every `required` field and the most `fields`
    ({field: [synonyms]}); None when no row does."""
    best = None
    for i, r in enumerate(rows[:scan]):
        cols = {}
        for j, cell in enumerate(r):
            for f, syns in fields.items():
                if f not in cols and any(head_match(cell, s) for s in syns):
                    cols[f] = j
                    break
        if all(f in cols for f in required) and (best is None or len(cols) > len(best[1])):
            best = (i, cols)
    return best


def find_tables(wb, fields, required, check=None, scan=15, exclude=()):
    """Every sheet of the workbook that holds the table, best first: [(sheet name, header index, {field: column}, rows below the header)].  `check(body, cols)` may refuse a table
    whose columns do not hold what their headers say (the values are the second witness); `exclude` = sheet names to leave out (already taken by another table)."""
    found = []
    for name in wb.sheetnames:
        if name in exclude:
            continue
        try:
            rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
        except Exception:                                   # a chart sheet or an unreadable sheet holds no table
            continue
        hit = locate_table(rows, fields, required, scan)
        if hit and (check is None or check(rows[hit[0] + 1:], hit[1])):
            found.append((name, hit[0], hit[1], rows[hit[0] + 1:]))
    found.sort(key=lambda x: -len(x[2]))
    return found


def sheet_like(wb, usual, must=()):
    """The sheet called `usual` (spaces, capitals and punctuation ignored), else the ONE sheet whose name has all the `must` words ('metro' for 'Metro RTO Codes'); None when there is
    none or when several fit (a name that fits two sheets is no evidence)."""
    key = "".join(_hwords(usual))
    for n in wb.sheetnames:
        if "".join(_hwords(n)) == key:
            return n
    hits = [n for n in wb.sheetnames if must and set(_hwords(" ".join(must))) <= set(_hwords(n))]
    return hits[0] if len(hits) == 1 else None


def pick_table(found, usual, what, ask=None, warnings=None, log=say):
    """The one table of `find_tables` to use: the only one; else the best (most headers, then the sheet whose name is closest to the `usual` name); two equally good
    ones are asked about (ask=None: the closest name wins and a warning says so).  Says on the log, and in the warnings when it is not on the usual sheet, where it found it."""
    if not found:
        return None
    ratio = lambda n: difflib.SequenceMatcher(None, "".join(_hwords(n)), "".join(_hwords(usual))).ratio()
    ranked = sorted(found, key=lambda x: (-len(x[2]), -ratio(x[0])))
    top = ranked[0]
    tie = [x for x in ranked if len(x[2]) == len(top[2]) and abs(ratio(x[0]) - ratio(top[0])) < 0.05]
    if len(tie) > 1 and ask is not None and getattr(ask, "interactive", True) is not None:
        pick = ask.choose(f"more than one sheet holds {what}: {', '.join(repr(x[0]) for x in tie)}",
                          [(f"sheet '{x[0]}' (header on row {x[1] + 1})", x[0]) for x in tie], question=f"Which sheet holds {what}?", short=f"sheet of {what}")
        top = next(x for x in tie if x[0] == pick)
    elif len(tie) > 1 and warnings is not None:
        warnings.append(f"more than one sheet holds {what} ({', '.join(repr(x[0]) for x in tie)}): '{top[0]}' is used")
    if " ".join(_hwords(top[0])) != " ".join(_hwords(usual)):
        msg = f"{what}: found by its column headers on sheet '{top[0]}' (header on row {top[1] + 1}); the usual sheet name is '{usual}'"
        log("  " + msg)
        if warnings is not None:
            warnings.append(msg)
    return top


def _read_master(path):
    """{group name: [codes / entries]} of a portal master export. The name and the list columns are found by their headers ('RTO NAME' / 'MAKE MODEL NAME', 'RTO CLUSTER' /
    'MAKE MODEL CLUSTER'); an export that has neither is read as the portal writes it today: ID, name, list."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()                                          # a read-only workbook keeps the file open (and locked on Windows) until closed
    hit = locate_table([list(r) for r in rows], {"name": ["rto name", "make model name", "name"], "list": ["rto cluster", "make model cluster", "cluster", "make model entries", "rto codes"]},
                       ("name", "list"), scan=6)
    head, name_col, list_col = (hit[0], hit[1]["name"], hit[1]["list"]) if hit else (0, 1, 2)
    out = {}
    for r in rows[head + 1:]:
        if r and len(r) > max(name_col, list_col) and r[name_col] is not None:
            out[str(r[name_col]).strip()] = [x.strip() for x in str(r[list_col] or "").split(",") if x.strip()]
    return out


def _find_master(folder, words):
    try:
        names = [n for n in os.listdir(folder) if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$") and all(w in n.lower() for w in words)]
    except OSError:
        return None
    return max((os.path.join(folder, n) for n in names), key=os.path.getmtime) if names else None


_MASTERS = {}


def load_masters(ask, folder=None):
    """The Masters found in `folder` (default: the master folder next to the scripts). If a file is missing the person is asked for the folder."""
    if os.environ.get("GRID_MASTER_API_URL"):
        return masters_from_api()
    folder = os.path.abspath(folder or os.environ.get("GRID_MASTER_DIR") or MASTER_DIR)       # GRID_MASTER_DIR: a fresher export than the one in master\\
    while True:
        rto_path, mm_path = _find_master(folder, ("rto", "master")), _find_master(folder, ("make", "model", "master"))
        if rto_path and mm_path:
            break
        missing = [what for what, p in (("the RTO master (a file with 'rto' and 'master' in its name)", rto_path),
                                        ("the make / model master (a file with 'make', 'model' and 'master' in its name)", mm_path)) if not p]
        def as_folder(a):
            p = clean_path(a)
            if not os.path.isdir(p):
                raise ValueError(f"'{p}' is not a folder (q = stop).")
            return p
        folder = ask.text(f"{' and '.join(missing)} not found in {folder}",
                          question=f"I can't find {' and '.join(missing)} in the folder '{folder}'. Type the path of the folder that holds the two master "
                                   f"workbooks exported from the portal:", validate=as_folder)
        ask.decide(f"master folder -> {folder}")
    key = (rto_path, mm_path, os.path.getmtime(rto_path), os.path.getmtime(mm_path))
    if key not in _MASTERS:
        try:
            _MASTERS[key] = Masters(_read_master(rto_path), _read_master(mm_path), (rto_path, mm_path))
        except Exception as e:
            raise GridError(f"cannot read the master workbooks ({os.path.basename(rto_path)}, {os.path.basename(mm_path)}): {type(e).__name__}: {e}")
    return _MASTERS[key]


# ---- RTO_NAMES == "master" (the user, 31 Oct 2026): a scope whose RTO codes are exactly a group of the master gets the master's name; one the master lacks is NEVER guessed and never stops
# the run: its cell holds a plain-text warning that the master-building agent (master_builder.py) reads and turns into the new group. The same for a make / model name the master lacks.
NEW_MAPPING = "New changes found, new mapping needs to be created: "
MASTER_SHEET = "the grid"          # the sheet(s) the make names come from, for the warning text: the script sets it (no commas: they would split the list)
MAKE_DEFINITIONS = {}              # a make group the master may lack -> (the master group it is made from, [the master groups taken out of it]); a script that knows one lists it here


def master_text(sheet, items=(), name=None, base=None, excluded=()):
    """The plain-text warning of a name the master lacks.
    RTOs:  'New changes found, new mapping needs to be created: <Sheet Name> - <CODE/CODE/...>'
    Makes: '... : <Sheet Name> - <GROUP> = <BASE GROUP> except <GROUP>|<GROUP>'  when the script knows how the group is made (MAKE_DEFINITIONS), else '... : <Sheet Name> - <GROUP>'.
    (A make-model entry may hold a '/', so the makes use '|'.)"""
    if name is None:
        return f"{NEW_MAPPING}{sheet} - {'/'.join(str(i).strip() for i in items if str(i).strip())}"
    return f"{NEW_MAPPING}{sheet} - {name}" + (f" = {base}" + (f" except {'|'.join(excluded)}" if excluded else "") if base else "")


def split_cell(cell):
    """([names the cell holds], [warning dicts]) - a warning is {sheet, items (RTO codes), name, base, excluded} (the grammar of `master_text`); the warnings come last in a cell."""
    head, *tails = str(cell).split(NEW_MAPPING)
    names = [t.strip() for t in head.split(",") if t.strip()]
    warns = []
    for tail in tails:
        tail = tail.strip().rstrip(",")
        if " = " in tail:                                   # a make group with its definition
            left, _, right = tail.partition(" = ")
            sheet, _, name = left.rpartition(" - ")
            base, _, ex = right.partition(" except ")
            warns.append(dict(sheet=sheet.strip(), items=[], name=name.strip(), base=base.strip(), excluded=[x.strip() for x in ex.split("|") if x.strip()]))
        else:
            sheet, _, rest = tail.rpartition(" - ")
            if CODE_LIST_RX.fullmatch(rest.strip()):        # RTO codes
                warns.append(dict(sheet=sheet.strip(), items=[i.strip() for i in rest.split("/") if i.strip()], name="", base="", excluded=[]))
            else:                                           # a name with no known definition (an RTO name or a make group)
                warns.append(dict(sheet=sheet.strip(), items=[], name=rest.strip(), base="", excluded=[]))
    return names, warns


CODE_LIST_RX = re.compile(r"[A-Z]{2}\d{1,3}(?:/[A-Z]{2}\d{1,3})*")


def make_group_by_content(masters, definition):
    """The master's make group whose entries are exactly `definition` = (base group, [groups taken out of it]) worked out in the master, or None (no such group, or the base is not in the master)."""
    low = {k.lower(): k for k in masters.makes}
    base = low.get(str(definition[0]).lower())
    out = [low.get(str(x).lower()) for x in definition[1]]
    if base is None or None in out:
        return None
    gone = {e.strip() for x in out for e in masters.makes[x]}
    want = frozenset(e.strip() for e in masters.makes[base] if e.strip() and e.strip() not in gone)
    if not want:
        return None
    return next((n for n in sorted(masters.makes) if frozenset(e.strip() for e in masters.makes[n] if e.strip()) == want), None)


def check_master_names(df, review, masters, log=say):
    """RTO_NAMES == "master": keep every RTO / make name the master has (in the master's own spelling), turn every name it lacks into the plain-text warning (`master_text`), list what is
    new on the sheet 'New mappings to create' and flag those rows. The run never stops for a missing mapping."""
    columns = {"new_rto_list": ({k.lower(): k for k in masters.rto}, "RTO"), "new_vehicle_makes": ({k.lower(): k for k in masters.makes}, "make / model")}
    new = {}                                                # (kind, sheet, what) -> rows
    for col, (low, label) in columns.items():
        table, found = {}, {}
        for v in df[col].astype(str).unique():
            if col == "new_rto_list" and v in WORKING_RTO:
                table[v] = v                                # the person's own RTO name (working list): written as typed, never checked
                continue
            names, warns = split_cell(v)
            kept, unknown = [], []
            for n in names:
                if n in low.values():
                    kept.append(n)
                elif n.lower() in low:
                    kept.append(low[n.lower()])             # the master's spelling IS the mapped name
                else:
                    unknown.append(n)
            for n in unknown:                               # one warning per name: each is its own group to create
                d = MAKE_DEFINITIONS.get(n) if label != "RTO" else None
                hit = make_group_by_content(masters, d) if d else None
                if hit:                                     # the master has a group with exactly the entries this one is defined by (master_builder.py made it): its name
                    kept.append(hit)
                    continue
                warns.append(dict(sheet=MASTER_SHEET, items=[], name=n, base=d[0] if d else "", excluded=list(d[1]) if d else []))
            parts = [master_text(w["sheet"], w["items"], w["name"] or None, w["base"], w["excluded"]) for w in warns]
            table[v] = ",".join(kept + parts)
            found[v] = [(label, w["sheet"], w["name"] or ",".join(w["items"]) + (f" = {w['base']}" if w["base"] else "")) for w in warns]
        df[col] = df[col].astype(str).map(table)
        for v, keys in found.items():
            n = int((df[col] == table[v]).sum())
            for k in keys:
                new[k] = new.get(k, 0) + n
    if new:
        add_sheet("New mappings to create", pd.DataFrame([dict(kind=k, sheet=s, items=i, rows=n) for (k, s, i), n in sorted(new.items())]))
    has = df["new_rto_list"].astype(str).str.contains(NEW_MAPPING, regex=False) | df["new_vehicle_makes"].astype(str).str.contains(NEW_MAPPING, regex=False)
    if has.any():
        review.flag("a new mapping needs to be created: the RTO / make cell holds the plain-text warning 'New changes found, new mapping needs to be created: <sheet> - <items>'",
                    lambda d: np.array(d["new_rto_list"].astype(str).str.contains(NEW_MAPPING, regex=False) | d["new_vehicle_makes"].astype(str).str.contains(NEW_MAPPING, regex=False), dtype=bool))
    log(f"Master lists ({len(masters.rto):,} RTO groups, {len(masters.makes):,} make / model groups): "
        + (f"{len(new)} new mapping(s) needed - {int(has.sum()):,} row(s) hold the plain-text warning instead of a name (sheet 'New mappings to create'; master_builder.py turns them into the new groups)"
           if new else "every RTO and make name written is in them"))
    return int(has.sum())


def masters_from_api():
    """The master source of the future (the user, 31 Oct 2026: 'we will connect to an RTO Master and a Make/Model Master via API'). Not connected yet: this build reads the portal's
    exported workbooks in master\\.  When the API exists, build `Masters(rto, makes)` here from it - {RTO NAME: [codes]} and {MAKE MODEL NAME: [entries]}, exactly what `_read_master` returns -
    and nothing else in the scripts changes (`load_masters` is the only door)."""
    raise GridError("GRID_MASTER_API_URL is set, but the master API is not connected yet: unset it to read the exported master workbooks in the master folder.")


def check_names(df, review, masters, log=say):
    """Flag, on the rows that carry it, every RTO group / make group name of the upload that is not in the master lists. A comma list is checked
    name by name. Nothing is replaced: the person adds the missing group to the portal's master and the mapping by hand.  RTO_NAMES == "master": `check_master_names`."""
    if RTO_NAMES == "master":
        return check_master_names(df, review, masters, log)
    apply_aliases(df)                                       # the person's own name mappings (web app), before the master is asked
    flagged = 0
    missing = {"rto": {}, "makes": {}}                      # name -> rows that carry it (the lists of the review workbook)
    for col, kind, label in (("new_rto_list", "rto", "RTO master"), ("new_vehicle_makes", "makes", "make / model master")):
        if kind == "rto" and RTO_NAMES == "cluster":        # new_rto_list holds cluster names: the person maps and groups the RTOs himself
            continue
        known = masters.rto if kind == "rto" else masters.makes
        for value in sorted(df[col].astype(str).unique()):
            if value in known or (kind == "rto" and value in WORKING_RTO):
                continue
            tokens = [t.strip() for t in value.split(",") if t.strip()]
            bad = [t for t in tokens if t not in known]
            if tokens and not bad:
                continue
            for t in (bad or [value]):
                missing[kind][t] = missing[kind].get(t, 0) + int((df[col].astype(str) == value).sum())
            hints = [f"'{t}' - the master spells it '{masters.near(kind, t)}'" for t in bad if masters.near(kind, t)]
            what = f"{col} '{value}' is not in the {label}" if (not tokens or tokens == [value]) else (
                f"{col} has the name(s) {', '.join(repr(t) for t in bad)} (in '{value}') that are not in the {label}")
            review.flag(what + (f" ({'; '.join(hints)})" if hints else ""), lambda d, c=col, v=value: np.array(d[c].astype(str) == v, dtype=bool))
            flagged += int((df[col].astype(str) == value).sum())
    for kind, sheet in (("rto", "RTO groups to create"), ("makes", "Make groups to create")):
        if missing[kind]:
            rows = [dict(group_needed=t, rows_flagged=n, items=(f"the master spells it '{masters.near(kind, t)}'" if masters.near(kind, t) else ""))
                    for t, n in sorted(missing[kind].items(), key=lambda kv: (-kv[1], kv[0]))]
            if os.environ.get("GRID_REPORT_FILE"):             # the web app offers the closest master names
                names = list(masters.rto if kind == "rto" else masters.makes)
                for r in rows:
                    r["suggestions"] = ", ".join(_difflib.get_close_matches(r["group_needed"], names, n=5, cutoff=0.6))
            add_sheet(sheet, pd.DataFrame(rows))
    log(f"Master lists ({len(masters.rto):,} RTO groups, {len(masters.makes):,} make / model groups): "
        + (("every make name written is in them" if RTO_NAMES == "cluster" else "every RTO and make name written is in them") if not flagged
           else f"{flagged:,} row(s) carry a name that is NOT in them (flagged in remarks)")
        + ("; new_rto_list holds the cluster names of the grid (not checked: you map and group the RTOs in the portal)" if RTO_NAMES == "cluster" else ""))
    return flagged


def colab_upload(ask):
    """colab_files.upload() -> {name: bytes}. If Colab's upload window itself fails ("Failed to fetch": a lost
    connection, a browser extension or ad blocker, a restarted runtime) it offers to use files you put in the
    Colab Files pane instead, or to try the window again."""
    while True:
        try:
            return colab_files.upload() or {}
        except (GridError, KeyboardInterrupt):
            raise
        except Exception as e:
            reason = f"{type(e).__name__}: {e}"
            pick = ask.choose(
                f"the Colab upload window failed: {reason}",
                [("Use files I have put in the Files pane (folder icon on the left) - I will type their paths", "paths"),
                 ("Try the upload window again", "retry")],
                question=f"Colab's upload window failed ({reason}). That is a problem between your browser and Colab "
                         f"(connection, a browser extension, an ad blocker), not with your grid. What now?",
                record=False)
            if pick == "retry":
                continue
            got = {}
            while True:
                p = ask_file(ask, "no file given",
                             "Path of a file in the Files pane (for example /content/Sep26_Branch_Grid_v3.xlsx), "
                             "or press Enter when you have given them all:",
                             exts=EXCEL_EXTS + (".csv",), blank_ok=True)
                if not p:
                    break
                with open(p, "rb") as fh:
                    got[os.path.basename(p)] = fh.read()
            if got:
                return got


def colab_download(path):
    """colab_files.download(path) - and the review workbook next to it when there is one; if Colab cannot start a download, say where the file is instead of crashing."""
    for p in (path, review_path(path)):
        if p != path and not os.path.exists(p):
            continue
        try:
            colab_files.download(p)
        except Exception as e:
            say(f"\nColab could not start the download ({type(e).__name__}: {e}).\n"
                f"The file IS written: {os.path.abspath(p)}\n"
                f"Open the Files pane (folder icon on the left), find it there and download it from there.")


def crash_report(script, extra=""):
    """An unexpected error: save the technical details to a file next to where you ran the script and say where."""
    try:
        path = f"{script}_error_{time.strftime('%Y%m%d_%H%M%S')}.txt"
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(f"{extra}\nargv: {sys.argv}\npython: {sys.version}\n\n{traceback.format_exc()}")
        return path
    except Exception:
        return None


# %% ---------------------------------------------------------------- READING THE WORKBOOK
USUAL_RATE_SHEET = "GCV _PCV _Pvt Car_TW"        # the usual name: only a hint for the log (the sheets are found by their headers)


def _text(v):
    return "" if v is None else str(v).replace("\xa0", " ").replace("​", "").strip()


def norm(v):
    """A cell as one line of text: no line breaks, no double spaces, no non-breaking spaces."""
    return re.sub(r"\s+", " ", _text(v)).strip()


def label_key(s):
    """A label without spaces, punctuation and capitals: 'NCB Status' and 'ncb  status' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def num(v):
    """A cell as a float, or None when it is not a plain number (blank, 'D', text)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%?\s*", _text(v))
    return float(m.group(1)) if m else None


def round_half_up(x, places=RATE_DECIMALS):
    q = 10 ** places
    return math.floor(float(x) * q + 0.5 + 1e-9) / q


def rnd(x):
    """a percent value from a cell: 0.07500000000000001 * 100 -> 7.5"""
    return round(float(x) + 1e-9, 4)


def read_workbook(path):
    """Open with values only. NOT read-only mode: the header cells are merged and only the normal mode keeps the merged ranges."""
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def hidden_sheets(wb):
    return {ws.title for ws in wb.worksheets if ws.sheet_state != "visible"}


def col_letter(c):
    from openpyxl.utils import get_column_letter
    return get_column_letter(c)


def rate_sheet(wb, ask, warnings, log=say):
    """(sheet name, header index, {field: column}, body): the visible sheet whose header says 'RTO Cluster Name' | 'State' and that has rate columns (dozens of numeric columns) beside them."""
    def has_rates(body, cols):
        width = max((len(r) for r in body[:80]), default=0)
        n = 0
        for c in range(cols["state"] + 1, width):
            if sum(1 for r in body[:80] if c < len(r) and isinstance(r[c], (int, float)) and not isinstance(r[c], bool)) >= 3:
                n += 1
        return n >= 12
    found = find_tables(wb, {"cluster": ["rto cluster name", "rto cluster"], "state": ["state"]}, ("cluster", "state"), check=has_rates, exclude=hidden_sheets(wb))
    top = pick_table(found, USUAL_RATE_SHEET, "the SBI rate grid", ask, warnings, log)
    if top is None:
        raise GridError("no visible sheet has the rate-grid header 'RTO Cluster Name' | 'State' with rate columns beside it - is this the SBI broker grid?")
    return top


class RateTable:
    """The SBI rate sheet: the five header rows (merged cells carried down into every column they cover), the cluster rows, and the cells. Blocks are found by header WORDS
    (`find_cols`), never by a column letter.  Header levels of a column: top (row 1), sub (row 2), row3, row4 (the row above the leaf: age / cc band / fuel group), leaf (the row with 'RTO Cluster Name')."""

    def __init__(self, wb, name, hdr, cols):
        self.ws, self.name, self.hdr = wb[name], name, hdr             # hdr = 0-based index of the leaf header row
        ws = self.ws
        self.c_cluster, self.c_state = cols["cluster"] + 1, cols["state"] + 1
        merged = {}
        for mr in ws.merged_cells.ranges:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merged[(r, c)] = (mr.min_row, mr.min_col)
        self.width = ws.max_column
        self.levels = {c: [norm(ws.cell(*merged.get((r, c), (r, c))).value) for r in range(1, hdr + 2)] for c in range(1, self.width + 1)}
        # the chain of a column: its distinct header texts from the top to the leaf - empty rows above the header and a merged cell repeated down several rows do not matter
        self.chain = {}
        for c, lv in self.levels.items():
            ch = []
            for x in lv:
                if x and (not ch or ch[-1] != x):
                    ch.append(x)
            self.chain[c] = ch
        self.claimed = set()
        self.rows = self._cluster_rows()
        self.last_row = self.rows[-1][0] if self.rows else hdr + 1

    def _cluster_rows(self):
        out = []
        for r in range(self.hdr + 2, self.ws.max_row + 1):
            cl, st = norm(self.ws.cell(r, self.c_cluster).value), norm(self.ws.cell(r, self.c_state).value)
            if not cl:
                break
            if not st and not any(isinstance(self.ws.cell(r, c).value, (int, float)) for c in range(self.c_state + 1, self.width + 1)):
                break                                                  # the first text row under the table ('Slab:-', 'Notes:')
            out.append((r, cl, st))
        return out

    def lv(self, c, which):
        """a header level of a column, counted from the chain of its distinct texts: top = first, leaf = last, row4 = the one above the leaf (age / cc band / fuel group), row3 = the one above that,
        sub = the second of a chain of three or more (the segment title of the GCV blocks)"""
        ch = self.chain[c]
        n = len(ch)
        return {"top": ch[0] if n else "", "sub": ch[1] if n >= 3 else "", "row3": ch[-3] if n >= 3 else "", "row4": ch[-2] if n >= 2 else "", "leaf": ch[-1] if n else ""}[which]

    def find_cols(self, **pats):
        """the rate columns (right of the State column) whose header levels match every given regex (top= sub= row3= row4= leaf=), in sheet order"""
        out = []
        for c in range(self.c_state + 1, self.width + 1):
            if all(re.search(p, self.lv(c, k), re.I) for k, p in pats.items() if p is not None):
                out.append(c)
        return out

    def cell(self, r, c):
        """(number or None, a flag text when the cell holds text that is not a number)"""
        raw = self.ws.cell(r, c).value
        v = num(raw)
        if v is None and norm(raw):
            return None, f"cell {col_letter(c)}{r} holds the text '{norm(raw)[:30]}', not a number: written as 0"
        return v, ""

    def path(self, c):
        return " / ".join(self.chain[c])

    def numeric_columns(self):
        """rate columns that hold at least one number in a cluster row"""
        out = []
        for c in range(self.c_state + 1, self.width + 1):
            if any(isinstance(self.ws.cell(r, c).value, (int, float)) for r, _, _ in self.rows):
                out.append(c)
        return out


# ---- the notes under the table
class Note:
    """One text of the sheet: its label (the product word in the cluster column) and its text. `rest` = the text with the words a rule has understood taken out."""

    def __init__(self, r, label, text):
        self.r, self.label, self.text = r, norm(label), norm(text)
        self.rest, self.used = self.text, False


def collect_notes(table):
    """Every text of the sheet that is not part of the rate table: the title above it and every row under the last cluster row (label | text)."""
    ws, notes, seen = table.ws, [], set()
    for r in range(1, table.hdr + 1):
        for c in (table.c_cluster, table.c_state):
            t = norm(ws.cell(r, c).value)
            if len(t) > 12 and t not in seen:
                seen.add(t)
                notes.append(Note(r, "title", t))
    for r in range(table.last_row + 1, ws.max_row + 1):
        cells = [(c, norm(ws.cell(r, c).value)) for c in range(1, table.width + 1)]
        cells = [(c, t) for c, t in cells if t]
        long_ = [(c, t) for c, t in cells if len(t) >= 3]
        if not long_:
            continue
        body = max(long_, key=lambda x: len(x[1]))
        label = " ".join(t for c, t in cells if c != body[0] and c <= table.c_state).strip(" *")      # the product word in the cluster column ('Tw', 'Pvt Car'); a bare '*' is no label
        notes.append(Note(r, label, body[1]))
    return notes


def take(note, m):
    """the words of `m` are understood: take them out of the note's rest"""
    note.used = True
    note.rest = (note.rest[:m.start()] + " " + note.rest[m.end():]).strip()


def first_note(notes, label, text):
    """(note, match) of the first note whose label matches `label` (None = any) and whose unused text matches `text`"""
    lab, pat = (re.compile(label, re.I) if label else None), re.compile(text, re.I | re.S)
    for n in notes:
        if lab is not None and not lab.search(n.label):
            continue
        m = pat.search(n.rest)
        if m:
            return n, m
    return None, None


# ---- the TEXT SCANNER: every rule the notes state, found by its words and read as a number (the user, 1 Nov 2026: 'a dedicated text-scanner function to read all grid notes, map the rules and apply
# them'; no row or cell is ever addressed).  A rule whose note has gone is ASKED about (last month's value first), never silently kept.
PERIOD_RX = r"from\s+(\d{1,2})(?:st|nd|rd|th)?\s+([a-z]{3,9})\s*'?\s*(\d{2,4})\s+to\s+(\d{1,2})(?:st|nd|rd|th)?\s+([a-z]{3,9})\s*'?\s*(\d{2,4})"
SLAB_RX = re.compile(r"for\s+(?P<name>[<>]?\s*\d+(?:\.\d+)?\s*[kl]?\s*(?:-\s*\d+(?:\.\d+)?\s*[kl]?)?)\s+grid\s+would\s+be\s+lower\s+by\s*(?P<cut>\d+(?:\.\d+)?)\s*%", re.I)
IRDA_SLAB_RX = re.compile(r"for\s+less\s+than\s*(?P<name>\d+\s*[kl]).*$", re.I | re.S)


def _date_of(d, mon, y):
    y = int(y)
    return dt.date(y + 2000 if y < 100 else y, dt.datetime.strptime(mon[:3].title(), "%b").month, int(d))


def _age_pair(m):
    return int(m.group(1)), int(m.group(2))


# (rule name, label regex, text regex, converter, what it is - for the question when the note is gone)
RULE_TABLE = [
    ("age.gcv", None, r"comprehensive\s+sourcing\s+is\s+allowed\s+upto\s*(\d+)\s*years?.*?satp\s+sourcing\s+is\s+allowed\s+upto\s*(\d+)\s*years?", _age_pair, "the vehicle age limits of the commercial grids (comprehensive / SATP)"),
    ("age.pc.comp", r"^pvt car$", r"comp\s+and\s+saod\s+age\s+upto\s*(\d+)\s*years?\s+allowed", lambda m: int(m.group(1)), "the Private Car comprehensive / SAOD age limit"),
    ("age.pc.satp", r"^pvt car$", r"satp\s+age\s+upto\s*(\d+)\s*years?\s+allowed", lambda m: int(m.group(1)), "the Private Car SATP age limit"),
    ("age.pcv3w", r"^pcv 3w$", r"comp\W+upto\s*(\d+)\s*year.*?satp\s+upto\s*(\d+)\s*year", _age_pair, "the PCV 3W age limits (comprehensive / SATP)"),
    ("age.taxi", r"^pcv taxi$", r"comp\W+upto\s*(\d+)\s*year.*?satp\s+upto\s*(\d+)\s*year", _age_pair, "the PCV taxi age limits (comprehensive / SATP)"),
    ("age.bus", r"school bus|staff bus", r"upto\s*(\d+)\s*years?\s+only\s+allowed", lambda m: int(m.group(1)), "the school / staff bus age limit"),
    ("age.tw.comp", r"^tw$", r"comp\s+and\s+saod\s+age\s+upto\s*(\d+)\s*years?\s+allowed", lambda m: int(m.group(1)), "the two-wheeler comprehensive / SAOD age limit"),
    ("age.tw.satp", r"^tw$", r"satp\s+age\s+upto\s*(\d+)\s*years?\s+allowed", lambda m: int(m.group(1)), "the two-wheeler SATP age limit"),
    ("age.tractor", r"tractor", r"comprehensive\s*&\s*satp\s+sourcing\s+is\s+allowed\s+upto\s*(\d+)\s*years?", lambda m: int(m.group(1)), "the tractor / harvester age limit"),
    ("pc.noncb_cut", r"^pvt car$", r"reduction\s+of\s*(\d+(?:\.\d+)?)\s*%\s+grid\s+will\s+be\s+applied\s+on\s+private\s+car\s+non\s*ncb\s+comprehensive\s+and\s+saod", lambda m: float(m.group(1)),
     "the Private Car non-NCB reduction (comprehensive and SAOD, non High-End)"),
    ("he.noncb_zero", r"^pvt car$", r"zero\s+payout\s+will\s+be\s+applicable\s+on\s+high-?\s*end\s+non\s*ncb\s+policies", lambda m: True, "the zero payout of High-End non-NCB policies"),
    ("taxi.noncb_cut", r"^pcv taxi$", r"(?:for\s+pcv\s+taxi\s*\(.*?\)\s*grid\s*-?\s*)?non\s*ncb\s*/\s*new\s+grid\s+would\s+be\s+lesser\s+by\s*(\d+(?:\.\d+)?)\s*%", lambda m: float(m.group(1)), "the taxi non-NCB / new vehicle reduction"),
    ("pc.satp", r"^pvt car satp$", r"grid\s+would\s+be\s+lesser\s+by\s*(\d+(?:\.\d+)?)\s*%\s+for\s+vehicle\s+age\s*(\d+)\s*-\s*(\d+)\s*years?\s*&?\s*below\s*(\d+)\s*year\s*age\s*grid\s+is\s+zero",
     lambda m: (float(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))), "the Private Car SATP reduction by vehicle age (and zero below 1 year)"),
    ("tw.cc_max", r"^tw$", r"cubic\s+capacity\s*>\s*(\d+)\s*cc\s*\(?decline", lambda m: int(m.group(1)), "the two-wheeler engine limit (models above it are declined)"),
]


def rule_gone(ask, warnings, name, what):
    """A rule whose note has gone from the grid is ASKED about, last month's value first (the user, 1 Nov 2026) - never silently kept.  -> the value to use (None = the rule is not applied)."""
    last = LAST_KNOWN.get(name)
    shown = "yes" if last is True else last
    pick = ask.choose(f"the note about {what} was not found in the grid",
                      [(f"Use last month's value: {shown}", "last"), ("Do not apply it", "none")],
                      question=f"I can't find the note about {what} (it used to say {shown}). Use last month's value, or leave it out?", short=what, memo=f"rule {name}")
    warnings.append(f"the note about {what} was not found: " + (f"last month's value ({shown}) is used" if pick == "last" else "the rule is not applied"))
    return last if pick == "last" else None


# ---- the fuels a note declines.  The November 2026 grid says 'In PCV 3W EV Decline' and the Electric rows were written anyway: the old rule only knew 'Electric declined for all GCV' (the user, 3 Nov 2026:
# 'the conditions explicitly state that EVs are declined').  ONE reader for every product and fuel, so the next product that declines a fuel needs no new code.  A note is cut into clauses (at brackets, dashes,
# semicolons, stars and full stops) and a clause that says NOTHING but '<fuel> declined / not allowed' - with the product in the clause or in the note's label - is applied and used up.  Anything more
# specific ('Mahindra EV declined' names a make) is not guessed: it stays an unread note and is flagged on the rows it may affect.
FUEL_WORDS = {"ev": "Electric", "electric": "Electric", "diesel": "Diesel", "cng": "CNG", "lpg": "LPG", "petrol": "Petrol"}
FUEL_PRODUCTS = [                                # product words -> the product key that fuel_key() gives for a segment; the first match wins
    (r"pcv\s*3w", "pcv3w"), (r"gcv", "gcv"), (r"taxis?\b", "taxi"), (r"school\s*bus|staff\s*bus", "bus"), (r"\btw\b|two[- ]?wheelers?", "tw"), (r"tractors?|harvesters?", "miscd"),
    (r"(?:pvt|private)\s*car\s*satp", "pc_satp"), (r"(?:pvt|private)\s*car", "pc"),
]
FUEL_PRODUCT_NAMES = {"gcv": "GCV", "pcv3w": "PCV 3W", "taxi": "PCV taxi", "bus": "school / staff bus", "tw": "two-wheeler", "miscd": "tractor / harvester", "pc": "Private Car", "pc_satp": "Private Car SATP"}
_FP = r"(?P<prod>pcv\s*3w|gcv(?:\s*3w)?|(?:pcv\s*)?taxis?|school\s*bus|staff\s*bus|tw|two[- ]?wheelers?|tractors?|harvesters?|(?:pvt|private)\s*car(?:\s*satp)?)"
_FF = r"(?P<fuel>ev|electric|diesel|cng|lpg|petrol)(?:\s+vehi(?:cle|cal)s?)?"
_FNO = r"(?:(?:is|are|will\s+be)\s+)?(?:declin\w*|delcin\w*|not\s+allowed)"
FUEL_DECLINE_RX = [re.compile(p, re.I) for p in (
    rf"(?:(?:in|for)\s+)?(?:the\s+)?{_FP}\s+(?:fuel\s+type\s+)?{_FF}\s+{_FNO}",                                   # In PCV 3W EV Decline
    rf"(?:fuel\s+type\s+)?{_FF}\s+{_FNO}(?:\s+(?:for|in)\s+(?:all\s+|the\s+)?{_FP})?",                           # Fuel type Electric declined for all GCV   /   EV declined
    rf"(?:declin\w*|delcin\w*)\s+(?:of\s+)?(?:all\s+)?{_FF}(?:\s+(?:for|in)\s+(?:all\s+|the\s+)?{_FP})?",        # Decline EV for PCV 3W
)]


def fuel_product(*texts):
    """the product key (FUEL_PRODUCTS) that the first of the texts names, or None"""
    for t in texts:
        for rx, key in FUEL_PRODUCTS:
            if t and re.search(rx, t, re.I):
                return key
    return None


def fuel_key(seg):
    """the product key the fuel declines are kept under, for a segment of the build (GCV_SEGS ... 'school_18')"""
    if seg in GCV_SEGS:
        return "gcv"
    return {"pc_he": "pc", "pc_satp_p": "pc_satp", "pc_satp_d": "pc_satp", "school_12": "bus", "school_18": "bus"}.get(seg, seg)


def read_fuel_declines(notes):
    """-> {product key: {fuels}} that the notes decline; the clauses read are used up"""
    out = {}
    for n in notes:
        for clause in re.split(r"[()*;|–—]|\.(?=\s|$)|\s-\s", n.rest):
            c = clause.strip(" :-_.")
            for rx in FUEL_DECLINE_RX if c else ():
                m = rx.fullmatch(c)
                if m:
                    key = fuel_product(m.groupdict().get("prod") or "") or fuel_product(n.label)
                    if key:
                        out.setdefault(key, set()).add(FUEL_WORDS[m.group("fuel").lower()])
                        n.used, n.rest = True, n.rest.replace(clause, " ", 1)
                    break
    return out


# ---- extra points for one state's registrations in a tonnage range: 'additional 2%- DL Registration Vehical 2.5 to 3.5 Segment Only' (label GCV).  The user (3 Nov 2026): the note is only for the 2.5 t to 3.5 t
# segment, i.e. GCV 4W rows whose weight band lies inside 2500.01-3500.01 (2500.01-3000.01, 3000.01-3500.01 or 2500.01-3500.01), Comp and SATP; they had added the 2 points by hand to the whole DL - NCR cluster.
# The note is only READ here; Ctx finds the clusters that hold the state's RTOs and uses the note up when it can apply it (else it stays an unread note and is flagged).
UPLIFT_RX = re.compile(r"additional\s*(?P<pts>\d+(?:\.\d+)?)\s*%\s*-?\s*(?P<state>[a-z]{2})\s+registration\s+vehi\w*\s+(?P<lo>\d+(?:\.\d+)?)\s*t?\s*(?:to|-)\s*(?P<hi>\d+(?:\.\d+)?)\s*t?\s*segment\s+only", re.I)


def read_uplifts(notes):
    """-> [{pts, state, lo, hi, _note}] the extra points a GCV note gives to the weight bands between `lo` and `hi` tonnes in one state's registrations.  The WHOLE note must be that sentence: a
    qualifier after it ('... Segment Only but not for Tata') is not guessed - the note stays unread and is flagged."""
    out = []
    for n in notes:
        m = UPLIFT_RX.fullmatch(n.rest.strip(" *.:-")) if re.search(r"\bgcv\b", n.label, re.I) else None
        if m:
            out.append(dict(pts=float(m.group("pts")), state=m.group("state").upper(), lo=float(m.group("lo")), hi=float(m.group("hi")), _note=n))
    return out


def read_rules(notes, ask, warnings, log=say):
    """-> dict of the rules the notes state (see RULE_TABLE and below); every note a rule understood is marked used and its words taken out."""
    R = {}
    for name, label, text, conv, what in RULE_TABLE:
        n, m = first_note(notes, label, text)
        if n is not None:
            R[name] = conv(m)
            take(n, m)
            continue
        R[name] = rule_gone(ask, warnings, name, what)
    # the fuels the notes decline, for every product (see above).  GCV Electric is a rule of its own: when its note has gone the person is asked, as for the other rules.
    R["fuel_declines"] = read_fuel_declines(notes)
    if "Electric" in R["fuel_declines"].get("gcv", ()):
        R["gcv.no_electric"] = True
    else:
        R["gcv.no_electric"] = rule_gone(ask, warnings, "gcv.no_electric", "the note that Electric vehicles are declined for all GCV")
        if R["gcv.no_electric"]:
            R["fuel_declines"].setdefault("gcv", set()).add("Electric")
    R["uplifts"] = read_uplifts(notes)
    # the validity period of the grid: the title and the first note say 'from 3rd Oct 26 to 31st Oct 26'
    R["period"] = None
    for n in notes:
        m = re.search(PERIOD_RX, n.rest, re.I)
        if m:
            try:
                p = (_date_of(m.group(1), m.group(2), m.group(3)), _date_of(m.group(4), m.group(5), m.group(6)))
            except ValueError:
                continue
            R["period"] = R["period"] or p
            take(n, m)
    # the two-wheeler make list: 'Declined Makes - All Makes other than Hero, Yamaha, TVS, Suzuki, RE, Bajaj, KTM Honda and Ather'
    n, m = first_note(notes, r"^tw$", r"declined\s+makes\W*all\s+makes\s+other\s+than\s+(?P<tail>.+?)\s*\.?\s*$")
    R["tw.makes"] = None
    if n:
        tail = m.group("tail").lower()
        found = [v for k, v in sorted(TW_MAKE_NAMES.items(), key=lambda kv: -len(kv[0])) if re.search(r"(?<![a-z])" + re.escape(k) + r"(?![a-z])", tail)]
        left = re.sub(r"(?<![a-z])(" + "|".join(re.escape(k) for k in TW_MAKE_NAMES) + r")(?![a-z])", " ", tail)
        left = re.sub(r"\band\b|[,&.]", " ", left).split()
        R["tw.makes"] = [v for v in TW_MAKE_ORDER if v in set(found)] + sorted(set(found) - set(TW_MAKE_ORDER))
        R["tw.makes_unknown"] = left
        take(n, m)
    # the T&C text of the two-wheeler rows: the user's October upload carries the races exclusion in add_tnc
    n, m = first_note(notes, r"^tw$", r"vehicles\s+participating\s+in\s+races.*")
    R["tw.tnc"] = n.text.strip(" *.") if n else ""
    if n:
        take(n, m)
    # 'Non doable RTOs' of a segment: RTOs the grid does not pay for (AP02 / AP03 ...): the segment gets the cluster minus these
    n, m = first_note(notes, None, r"2\.5\s*t\s*-\s*3\.5\s*t\s*gcv\W*\s*non\s*doable\s*rtos\W*(?P<tail>.+)")
    R["ap.nondoable"] = re.findall(r"\b[A-Z]{2}\d{2}\b", m.group("tail")) if n else []
    if n:
        take(n, m)
    # premium slabs: 'Pvt Car Slab- ... For 1L-3L Grid would be lower by 1% , ...'   /   'CV Slab- ...'
    for fam, word in (("pc", r"pvt\s*car"), ("cv", r"cv")):
        R["slabs." + fam], R["slab_irda." + fam] = [], None
        for n in notes:
            if not re.match(rf"\s*{word}\s*slab", n.rest, re.I):
                continue
            R["slabs." + fam] = [(re.sub(r"\s+", "", m.group("name")).upper(), float(m.group("cut"))) for m in SLAB_RX.finditer(n.rest)]
            irda = IRDA_SLAB_RX.search(n.rest)
            R["slab_irda." + fam] = re.sub(r"\s+", "", irda.group("name")).upper() if irda else None
            n.rest = IRDA_SLAB_RX.sub(" ", SLAB_RX.sub(" ", n.rest))
            n.used = True
    return R


# ---- the cluster lists of the 'Declined cluster' notes, cross-checked with the grid
ALL_AGES = {"new", "mid", "old", "upto5", "all"}
AGE_HEAD = re.compile(r"(new\s*&\s*upto\s*5\s*years?|upto\s*5\s*years?\s*&\s*above\s*5\s*years?|upto\s*5\s*years?|above\s*5\s*years?)\s*:?-?", re.I)
DECLINE_WORD = re.compile(r"(declin\w*|delcin\w*)\s*(rto\s*)?cluster", re.I)
CLUSTER_ALIASES = {"UP GV": "UP - GV"}


def clusters_in(text, names):
    """the cluster names (case-sensitive, as the grid writes them) that a text lists; the longest name first, so 'CH - R' is never read inside 'CH - Rest'"""
    t = text
    for a, b in CLUSTER_ALIASES.items():
        t = t.replace(a, b)
    found = set()
    for n in sorted(names, key=len, reverse=True):
        pat = r"(?<![A-Za-z0-9])" + r"\s*-\s*".join(r"\s+".join(re.escape(w) for w in part.split()) for part in re.split(r"\s*-\s*", n)) + r"(?![A-Za-z0-9])"
        if re.search(pat, t):
            found.add(n)
            t = re.sub(pat, " ", t)
    return found


def decline_sections(text):
    """[(age keys, text)] - a note may list its clusters per age ('Upto 5 Year:-... Above 5 Year :-...')"""
    heads = list(AGE_HEAD.finditer(text))
    if not heads:
        return [(ALL_AGES, text)]
    out = []
    for i, h in enumerate(heads):
        end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
        lab = h.group(1).lower()
        ages = ALL_AGES if ("upto" in lab and "above" in lab) else ({"new", "mid", "upto5"} if lab.startswith(("new", "upto")) else {"old"})
        out.append((ages, text[h.end():end]))
    return out


# note -> segment: (regex on label + text, segment key)
DECLINE_SEGMENTS = [
    (r'"\s*gcv\s*3w\s*"', "gcv3w"), (r'upto\s*2\.5\s*t\s*gcv"?\s*-?\s*declin', "gcv_u25"), (r'2\.5\s*t\s*-\s*3\.5\s*t\s*gcv"?\s*declin', "gcv_25_35"),
    (r'12\s*t\s*-\s*20\s*t\s*gcv', "gcv_12_20"), (r'20\s*t\s*-\s*40\s*t\s*gcv', "gcv_20_40"), (r'above\s*40\s*t', "gcv_40"),
    (r"^pcv 3w\b", "pcv3w"), (r"^pcv taxi\b", "taxi"), (r"^pvt car satp\b.*petrol", "pc_satp_p"), (r"^pvt car satp\b.*diesel", "pc_satp_d"), (r"tractor", "miscd"),
]


def read_declines(notes, names):
    """-> [{seg, text, sections: [(ages, mode 'declined'|'allowed', {clusters})]}] from the notes that list declined (or, for > 40 t, allowed) clusters; the notes are marked used"""
    out = []
    for n in notes:
        if not DECLINE_WORD.search(n.text):
            continue
        probe = f"{n.label} {n.text}"
        seg = next((s for rx, s in DECLINE_SEGMENTS if re.search(rx, probe, re.I)), None)
        if seg is None:
            continue
        secs = []
        for ages, body in decline_sections(n.text):
            mode = "allowed" if re.search(r"rest\s*(all)?\s*(cluster)?\s*declin", body, re.I) else "declined"
            secs.append((ages, mode, clusters_in(body, names)))
        out.append(dict(seg=seg, text=n.text, sections=secs))
        n.used, n.rest = True, ""
    return out


def read_state_declines(notes):
    """'All India RTOs allowed other than RJ, KL and MP which are fully declined' (the 12-17 seater school bus) -> [(segment, {state prefixes})]"""
    out = []
    for n in notes:
        m = re.search(r"all\s+india\s+rtos\s+allowed\s+other\s+than\s+(?P<tail>.+?)\s+which\s+are\s+fully\s+declined", n.rest, re.I)
        if m and re.search(r"12\s*to\s*17", n.label + " " + n.text, re.I):
            out.append(("school_12", set(re.findall(r"\b[A-Z]{2}\b", m.group("tail")))))
            take(n, m)
    return out


def read_pan_declines(notes):
    """'For "3.5T-5.0T GCV": Pan India decline all cluster all make model' -> [segment keys whose every cluster must be declined]"""
    out = []
    for n in notes:
        m = re.search(r"pan\s*india\s*declin\w*\s+all\s+cluster\s+all\s+make\s+model", n.rest, re.I)
        if not m:
            continue
        if re.search(r"3\.5\s*t\s*-\s*5\.0\s*t", n.text):
            out.append("gcv_35_50")
        elif re.search(r"5\.0\s*t\s*-\s*7\.5\s*t", n.text):
            out.append("gcv_50_75")
        elif re.search(r"7\.5\s*t\s*-\s*12", n.text):
            out.append("gcv_75_120")
        else:
            continue
        take(n, m)
    return out


# ---- the other sheets: RTO code -> cluster, the High-End enabler, the declined make / model lists
def read_mapping(wb, exclude, warnings, log=say):
    """({cluster: {RTO codes}}, {RTO code: cluster}) of the sheet that lists RTO code | RTO cluster name (found by its headers); ({}, {}) with a warning when there is none"""
    found = find_tables(wb, {"code": ["rto code"], "cluster": ["rto cluster name", "rto cluster"], "state": ["state code", "state"]}, ("code", "cluster"), exclude=set(exclude) | hidden_sheets(wb))
    if not found:
        warnings.append("no sheet lists RTO code | RTO cluster name (the usual 'RTO Cluster Mapping'): the clusters of the grid cannot be checked against it and an RTO-level note cannot be applied")
        return {}, {}
    name, _, cols, body = found[0]
    by_cluster, by_code = {}, {}
    for r in body:
        if len(r) > max(cols["code"], cols["cluster"]) and norm(r[cols["code"]]) and norm(r[cols["cluster"]]):
            code, cl = norm(r[cols["code"]]).upper(), norm(r[cols["cluster"]])
            for piece in re.split(r"[/\s]+", code):
                if piece:
                    by_cluster.setdefault(cl, set()).add(piece)
                    by_code.setdefault(piece, cl)
    log(f"RTO list: sheet '{name}', {len(by_cluster)} clusters, {len(by_code)} RTO codes")
    return by_cluster, by_code


def read_high_end(wb, exclude, warnings, log=say):
    """([(cluster, make cluster, model cluster, package payout, SAOD payout)], sheet name, the texts of the sheet outside the table) of the High-End enabler sheet (found by its headers); (None, '', [])"""
    fields = {"cluster": ["rto cluster name"], "code": ["model cluster code"], "make": ["make cluster"], "model": ["model cluster"], "package": ["package"], "saod": ["saod"]}
    found = find_tables(wb, fields, ("cluster", "make", "model", "package", "saod"), exclude=set(exclude) | hidden_sheets(wb))
    if not found:
        return None, "", []
    name, hdr, cols, body = found[0]
    rows, texts = [], []
    for r in body:
        if len(r) > max(cols.values()) and norm(r[cols["cluster"]]) and norm(r[cols["model"]]):
            rows.append((norm(r[cols["cluster"]]), norm(r[cols["make"]]), norm(r[cols["model"]]), num(r[cols["package"]]), num(r[cols["saod"]])))
        else:
            texts.extend(norm(c) for c in r if len(norm(c)) >= 3)                     # a line under the table
    texts.extend(norm(c) for r in wb[name].iter_rows(min_row=1, max_row=hdr, values_only=True) for c in r if len(norm(c)) >= 3)         # the lines above it
    log(f"High-End enabler: sheet '{name}', {len(rows)} rows")
    return rows, name, texts


def decline_lists_fingerprint(wb, exclude):
    """one fingerprint of the declined make / model lists (the sheets with a Make | Model table), to notice when SBI changes them"""
    found = find_tables(wb, {"make": ["make", "make name"], "model": ["model", "model name"]}, ("make", "model"), exclude=set(exclude) | hidden_sheets(wb))
    text = "|".join(sorted(label_key(" ".join(norm(c) for c in row)) for name, _, _, body in found for row in body if any(norm(c) for c in row)))
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12] if text else "", [f[0] for f in found]


# %% ---------------------------------------------------------------- BUILD
LAST_GAPS = {}            # RTO groups the master lacks for this run: group name -> the RTO codes the grid's mapping sheet puts in it (a sheet of the review workbook, see collect_gaps)
LAST_NOT_WRITTEN = []     # grid columns / rows that gave no upload row: the sheet 'Not written' of the review workbook
SPECIAL_GROUPS = {}       # RTO group that is a cluster minus some RTOs -> (condition, part) in plain English (cluster mode writes them after the cluster name, see rto_condition)

AGE_NEW, AGE_MID, AGE_OLD, AGE_UPTO5 = r"^new", r"upto\s*5\s*year|age\s*1\s*-\s*5", r"above\s*5", r"^age\s*upto\s*5"
LEAF_COMP, LEAF_SATP, LEAF_BOTH = r"^comp(?!.*(satp|stp))", r"^(satp|stp)\b", r"^comp\s*/\s*satp"

SEG_NAMES = {"gcv_u25": "GCV up to 2.5 t", "gcv3w": "GCV 3W", "gcv_25_35": "GCV 2.5 t-3.5 t", "gcv_tractor": "GCV tractor", "gcv_35_50": "GCV 3.5 t-5 t", "gcv_50_75": "GCV 5 t-7.5 t",
             "gcv_75_120": "GCV 7.5 t-12 t", "gcv_12_20": "GCV 12 t-20 t", "gcv_20_40": "GCV 20 t-40 t", "gcv_40": "GCV above 40 t", "pc": "Private Car comp / SAOD", "pc_he": "Private Car High-End",
             "pc_satp_p": "Private Car SATP (petrol)", "pc_satp_d": "Private Car SATP (diesel)", "pcv3w": "PCV 3W", "taxi": "PCV taxi", "tw": "two-wheeler", "miscd": "tractor / harvester",
             "school_18": "school bus 18+ seats", "school_12": "school bus 12-17 seats"}
SEG_TAG = {"gcv_25_35": "2.5_3.5T"}                 # the tail of the portal's group name for a cluster minus RTOs in that segment (SBI_OCT26_AP_-_Rest_2.5_3.5T)
GCV_SEGS = {"gcv_u25", "gcv3w", "gcv_25_35", "gcv_tractor", "gcv_35_50", "gcv_50_75", "gcv_75_120", "gcv_12_20", "gcv_20_40", "gcv_40"}
# premium slab families: Private Car uses the 'Pvt Car Slab' ("pc"), two-wheelers have none ("tw" = 0), every other product the 'CV Slab' ("cv"); the emitters name the family in ctx.rate(v, family)

# the GCV blocks: (segment, top-level header regex, sub-segment regex, [(make group, cc_min, cc_max)], age columns)
GCV_BLOCKS = [
    ("gcv_u25", r"upto\s*2\.5\s*t\s*gcv", r"upto\s*2\.0\s*t.*all\s*makes", [(MK_BELOW35, 0, 2000.01), (MK_TATA, 2000.01, 2500.01)], "full"),
    ("gcv_u25", r"upto\s*2\.5\s*t\s*gcv", r"other\s*than\s*tata", [(MK_BELOW35_EX_TATA, 2000.01, 2500.01)], "full"),
    ("gcv_25_35", r"2\.5\s*t\s*-\s*3\.5\s*t\s*gcv", r"mahindra", [(MK_MAHINDRA, 2500.01, 3500.01)], "full"),
    ("gcv_25_35", r"2\.5\s*t\s*-\s*3\.5\s*t\s*gcv", r"2\.5\s*-\s*3\.0\s*t\s*tata", [(MK_TATA, 2500.01, 3000.01)], "full"),
    ("gcv_25_35", r"2\.5\s*t\s*-\s*3\.5\s*t\s*gcv", r"ashok\s*leyland", [(MK_ASHOK, 2500.01, 3500.01), (MK_TATA, 3000.01, 3500.01)], "full"),
    ("gcv_12_20", r"12\s*t\s*to\s*20\s*t.*other\s*makes", None, [(MK_MAHINDRA, 12000.01, 20000.01)], "full"),
    ("gcv_12_20", r"12\s*t\s*to\s*20\s*t.*tata", None, [(MK_TATA + "," + MK_ASHOK, 12000.01, 20000.01)], "full"),
    ("gcv_20_40", r"20\s*t\s*to\s*40\s*t.*other\s*makes", None, [(MK_MAHINDRA, 20000.01, 40000.01)], "full"),
    ("gcv_20_40", r"20\s*t\s*to\s*40\s*t.*tata", None, [(MK_TATA + "," + MK_ASHOK, 20000.01, 40000.01)], "full"),
    ("gcv_40", r">\s*40\s*t.*other\s*makes", None, [(MK_MAHINDRA, 40000.01, CC_OPEN_END)], "split5"),
    ("gcv_40", r">\s*40\s*t.*tata", None, [(MK_TATA + "," + MK_ASHOK, 40000.01, CC_OPEN_END)], "split5"),
]
GCV_DECLINED = [("gcv_35_50", r"3\.5\s*t\s*-\s*5\.0\s*t", 3500.01, 5000.01), ("gcv_50_75", r"5\.0\s*t\s*-\s*7\.5\s*t", 5000.01, 7500.01), ("gcv_75_120", r"7\.5\s*t\s*-\s*12\s*t", 7500.01, 12000.01)]


def group_name(cluster):
    return GROUP_PREFIX + re.sub(r"\s+", "_", cluster.strip())


def rto_condition(value, product, sub_product):
    """cluster mode (apply_cluster_names): the plain-English condition of an RTO group that is a cluster minus some RTOs, else None"""
    return SPECIAL_GROUPS.get(value)


def bands_from_labels(labels):
    """Contiguous bands from labels such as 'Upto 999 CC', '1000 CC-1499 CC', 'Above 1500 CC' / 'A. 0-1000', 'B. 1001-1500', 'C. 1500+' -> {label: (min, max)}: each band ends at its upper limit + 0.01
    and the next one starts there (the user's convention), the last one is open (99999)."""
    parsed = []
    for lab in labels:
        nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", lab)]
        if not nums:
            continue
        opened = bool(re.search(r"above|over|\+\s*$", lab, re.I))
        lo = nums[0] if (len(nums) > 1 or opened) else 0.0
        hi = None if opened else nums[-1]
        parsed.append((lo, hi, lab))
    parsed.sort(key=lambda x: (x[0], x[1] is None))
    out, prev = {}, 0.0
    for lo, hi, lab in parsed:
        mx = CC_OPEN_END if hi is None else round(hi + 0.01, 2)
        out[lab] = (prev, mx)
        prev = mx
    return out


class Ctx:
    """What the emitters share: the rate table, the rules read from the notes, the slab of each family, the rows and the cross-check book."""

    def __init__(self, table, rules, slabs, mapping, masters, warnings):
        self.t, self.R, self.slab, self.warnings, self.masters = table, rules, slabs, warnings, masters
        self.by_cluster, self.by_code = mapping
        self.rows, self.not_written, self.gaps = [], [], {}
        self.zero = {}                     # (segment, age key) -> {cluster: True when the grid pays nothing there}
        self.excl = {}                     # segment -> {cluster: {RTO codes the segment leaves out}}
        self.clusters = [cl for _, cl, _ in table.rows]
        self.plain = {}                    # cluster -> its group name
        for code in rules.get("ap.nondoable") or []:
            cl = self.by_code.get(code)
            if cl:
                self.excl.setdefault("gcv_25_35", {}).setdefault(cl, set()).add(code)
        by_content = {}
        for name, codes in masters.rto.items():
            by_content.setdefault(frozenset(c.upper() for c in codes), []).append(name)
        self.by_content = by_content
        # the extra points of the notes ('additional 2% - DL Registration Vehicle 2.5 to 3.5 Segment Only'): the clusters that hold the state's RTOs get them on the GCV 4W weight bands inside the tonnage.
        # A cluster is the portal's unit, so the WHOLE cluster is written (the user did the same by hand); a note whose state no cluster holds is not applied and stays an unread note (flagged).
        self.uplifts = []
        for u in rules.get("uplifts", []):
            hit = {cl for cl, codes in self.by_cluster.items() if any(re.match(rf"{u['state']}\d", c) for c in codes)} or {cl for cl in self.clusters if cluster_prefix(cl) == u["state"]}
            if not hit:
                warnings.append(f"a note gives {u['pts']:g} extra points to {u['state']} registrations in {u['lo']:g}-{u['hi']:g} t but no cluster holds {u['state']} RTOs: not applied (flagged)")
                continue
            u["_note"].used, u["_note"].rest = True, ""
            self.uplifts.append(dict(u, clusters=hit, cc=(round(u["lo"] * 1000 + 0.01, 2), round(u["hi"] * 1000 + 0.01, 2))))

    # ---- numbers
    def uplift(self, cluster, ccmin, ccmax):
        """the extra points the notes give a GCV 4W weight band in a cluster: the band must lie inside the note's tonnage (2.5 to 3.5 = 2500.01-3500.01, 2500.01-3000.01 or 3000.01-3500.01)"""
        return sum(u["pts"] for u in self.uplifts if cluster in u["clusters"] and ccmin >= u["cc"][0] - 0.005 and ccmax <= u["cc"][1] + 0.005)

    def cap(self, key, which):
        """the maximum vehicle age (+0.01) of a rule: key 'age.xxx', which = 0 (comprehensive) / 1 (SATP)"""
        v = self.R.get(key)
        if isinstance(v, tuple):
            v = v[which]
        return round(float(v) + 0.01, 2) if v is not None else AGE_OPEN_END

    def rate(self, v, fam, extra=0.0):
        """the payin of a grid cell: the cell minus the premium slab of its family minus `extra`, never below 0; a blank or zero cell stays 0 (a declined entry)"""
        if v is None or v <= 0:
            return 0.0
        return max(v - self.slab[fam] - extra, 0.0)

    # ---- rows
    def fuels_allowed(self, seg, fuels):
        """the fuels of a row less the ones the notes decline for its product ('In PCV 3W EV Decline'): no row is written for a declined fuel"""
        gone = self.R.get("fuel_declines", {}).get(fuel_key(seg), ())
        return [f for f in fuels if f not in gone]

    def add(self, seg, cluster, *, product, sub, klass, makes, fuels, amin, amax, pit, rate, ccmin=0, ccmax=CC_OPEN_END, scmin=SC_RANGE[0], scmax=SC_RANGE[1], ncb="NA", zd="NA", tnc="", agekey="all",
            scope=None, flags=()):
        fuels = self.fuels_allowed(seg, fuels)
        if not fuels:
            return
        self.rows.append(dict(product=product, sub_product=sub, policy_type="Renew", make_model_class=klass, new_vehicle_makes=makes, new_rto_list=scope or self.group(cluster, seg),
                              vehicle_age_min=amin, vehicle_age_max=amax, cc_min=ccmin, cc_max=ccmax, sc_min=scmin, sc_max=scmax, is_ncb=ncb, is_cpa="NA", is_zd=zd, add_tnc=tnc, pi_type=pit,
                              pi_od_rate=rate if pit == "On OD" else 0.0, pi_tp_rate=rate if pit == "On TP" else 0.0, pi_net_rate=rate if pit == "On Net" else 0.0, tariff_min=0.0, tariff_max=TARIFF_FULL,
                              _fuels=tuple(fuels), _seg=seg, _cl=cluster, _age=agekey, _flags=tuple(flags)))

    def group(self, cluster, seg):
        """the portal's RTO group of a cluster; for a segment that leaves some RTOs out, the group that holds the cluster minus those RTOs"""
        out = self.excl.get(seg, {}).get(cluster)
        if not out:
            return self.plain.get(cluster) or self.plain.setdefault(cluster, rto_cluster(group_name(cluster), cluster))
        want = frozenset(c.upper() for c in self.by_cluster.get(cluster, ())) - frozenset(out)
        hits = sorted(n for n in self.by_content.get(want, []) if n.startswith(GROUP_PREFIX)) if want else []
        name = hits[0] if hits else f"{group_name(cluster)}_{SEG_TAG.get(seg, seg)}"
        if not hits:
            self.gaps[name] = sorted(want)
        SPECIAL_GROUPS[name] = (f"{'/'.join(sorted(out))} are not allowed RTOs for {SEG_NAMES.get(seg, seg)}", f"{cluster} without {'/'.join(sorted(out))}")
        return rto_cluster(name, cluster)

    def cell(self, r, c):
        v, flag = self.t.cell(r, c)
        return v, ((flag,) if flag else ())

    def mark(self, seg, agekey, cluster, v):
        """remember whether the grid pays nothing for a cluster in (segment, age): true only if every column of it is blank or zero"""
        d = self.zero.setdefault((seg, agekey), {})
        d[cluster] = d.get(cluster, True) and (v is None or v <= 0)

    def miss(self, seg, what):
        self.warnings.append(f"{SEG_NAMES.get(seg, seg)}: {what} - no rows were written for it")
        self.not_written.append(dict(sheet=self.t.name, column="", segment=SEG_NAMES.get(seg, seg), reason=what))

    # ---- columns
    def pair(self, cols, age_rx, seg, what):
        """(comp column, satp column) of one age in a block; a 'Comp /SATP' column serves both. A missing one is reported and its rows are not written."""
        sel = [c for c in cols if age_rx is None or re.search(age_rx, self.t.lv(c, "row4"), re.I)]
        comp = [c for c in sel if re.search(LEAF_COMP, self.t.lv(c, "leaf"), re.I)]
        satp = [c for c in sel if re.search(LEAF_SATP, self.t.lv(c, "leaf"), re.I)]
        both = [c for c in sel if re.search(LEAF_BOTH, self.t.lv(c, "leaf"), re.I)]
        if both and not comp and not satp:
            comp = satp = both
        self.t.claimed.update(sel)
        for name, lst in (("Comp", comp), ("SATP", satp)):
            if len(lst) > 1:
                raise GridError(f"{what}: {len(lst)} columns say {name} ({', '.join(col_letter(c) for c in lst)}) - the header of this block has changed")
        return (comp[0] if comp else None), (satp[0] if satp else None)

    def one(self, cols, what, **pats):
        """the one column among `cols` whose header levels match (None if none; an error if several)"""
        hit = [c for c in cols if all(re.search(p, self.t.lv(c, k), re.I) for k, p in pats.items())]
        if len(hit) > 1:
            raise GridError(f"{what}: {len(hit)} columns match ({', '.join(col_letter(c) for c in hit)}) - the header of this block has changed")
        if hit:
            self.t.claimed.add(hit[0])
        return hit[0] if hit else None


def emit_gcv(ctx):
    """GCV 4W blocks (up to 2.5 t ... above 40 t), the 3W block, the tractor column and the declined 3.5-12 t columns"""
    fuels = FUELS_NO_EV if ctx.R.get("gcv.no_electric") else FUELS_ALL
    cap_c, cap_s = ctx.cap("age.gcv", 0), ctx.cap("age.gcv", 1)
    for seg, top, sub, bands, ages in GCV_BLOCKS:
        cols = ctx.t.find_cols(top=top, sub=sub)
        if not cols:
            ctx.miss(seg, f"the block '{top}'{' / ' + sub if sub else ''} is not in the grid")
            continue
        defs = ([("new", AGE_NEW, 0, 1, TNC_NEW), ("mid", AGE_MID, 1, 5.01, ""), ("old", AGE_OLD, 5.01, None, "")] if ages == "full" else [("upto5", AGE_UPTO5, 0, 5.01, ""), ("old", AGE_OLD, 5.01, None, "")])
        for agekey, rx, amin, amax0, tnc in defs:
            comp, satp = ctx.pair(cols, rx, seg, f"{SEG_NAMES[seg]} {agekey}")
            for sub_p, col, pit, cap in (("1+1", comp, "On Net", cap_c), ("STP", satp, "On TP", cap_s)):
                if col is None:
                    ctx.miss(seg, f"no {'Comp' if sub_p == '1+1' else 'SATP'} column for the age group '{agekey}'")
                    continue
                for r, cl, _ in ctx.t.rows:
                    v, fl = ctx.cell(r, col)
                    ctx.mark(seg, agekey, cl, v)
                    for makes, ccmin, ccmax in bands:
                        rate = ctx.rate(v, "cv", -ctx.uplift(cl, ccmin, ccmax))             # the CV slab off, the note's extra points (DL 2.5-3.5 t) on; a blank / zero cell stays a zero card
                        for klass in GCV_CLASSES:
                            ctx.add(seg, cl, product="GCV 4W", sub=sub_p, klass=klass, makes=makes, fuels=fuels, amin=amin, amax=amax0 if amax0 is not None else cap, pit=pit, rate=rate,
                                    ccmin=ccmin, ccmax=ccmax, tnc=tnc, agekey=agekey, flags=fl)
    # GCV 3W sits inside the first block of the sheet
    cols = ctx.t.find_cols(top=r"upto\s*2\.5\s*t\s*gcv", sub=r"^gcv\s*3w")
    if not cols:
        ctx.miss("gcv3w", "the block 'GCV 3W' is not in the grid")
    for agekey, rx, amin, amax0, tnc in [("new", AGE_NEW, 0, 1, TNC_NEW), ("mid", AGE_MID, 1, 5.01, ""), ("old", AGE_OLD, 5.01, None, "")] if cols else []:
        comp, satp = ctx.pair(cols, rx, "gcv3w", f"GCV 3W {agekey}")
        for sub_p, col, pit, cap in (("1+1", comp, "On Net", cap_c), ("STP", satp, "On TP", cap_s)):
            if col is None:
                ctx.miss("gcv3w", f"no {'Comp' if sub_p == '1+1' else 'STP'} column for the age group '{agekey}'")
                continue
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col)
                ctx.mark("gcv3w", agekey, cl, v)
                ctx.add("gcv3w", cl, product="GCV 3W", sub=sub_p, klass="Goods Carrying Rickshaw", makes=MK_GCV3W, fuels=fuels, amin=amin, amax=amax0 if amax0 is not None else cap, pit=pit,
                        rate=ctx.rate(v, "cv"), tnc=tnc, agekey=agekey, flags=fl)
    # the GCV tractor: one Comp /SATP column
    cols = ctx.t.find_cols(top=r"gcv\s*tractor")
    if not cols:
        ctx.miss("gcv_tractor", "the column 'GCV Tractor' is not in the grid")
    else:
        comp, satp = ctx.pair(cols, None, "gcv_tractor", "GCV tractor")
        for sub_p, col, pit, cap in (("1+1", comp, "On Net", cap_c), ("STP", satp, "On TP", cap_s)):
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col) if col else (None, ())
                ctx.mark("gcv_tractor", "all", cl, v)
                ctx.add("gcv_tractor", cl, product="GCV 4W", sub=sub_p, klass="Goods Carrying Tractor", makes=MK_BELOW35, fuels=fuels, amin=0, amax=cap, pit=pit, rate=ctx.rate(v, "cv"), flags=fl)
    # 3.5-12 t: 'Pan India decline all cluster all make model' - the columns are blank; a zero card is written for each cluster
    for seg, top, ccmin, ccmax in GCV_DECLINED:
        cols = ctx.t.find_cols(top=top)
        if not cols:
            ctx.miss(seg, f"the column '{top}' is not in the grid")
            continue
        comp, satp = ctx.pair(cols, None, seg, SEG_NAMES[seg])
        for sub_p, col, pit, cap in (("1+1", comp, "On Net", cap_c), ("STP", satp, "On TP", cap_s)):
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col) if col else (None, ())
                ctx.mark(seg, "all", cl, v)
                for klass in GCV_CLASSES:
                    ctx.add(seg, cl, product="GCV 4W", sub=sub_p, klass=klass, makes=MK_GCV_ALL, fuels=fuels, amin=0, amax=cap, pit=pit, rate=ctx.rate(v, "cv"), ccmin=ccmin, ccmax=ccmax, flags=fl)


def emit_pc(ctx):
    """Private Car comprehensive and SAOD 'other than High-End' (non-NCB: the note's cut), found by its headers"""
    cols = ctx.t.find_cols(top=r"pvt\s*car.*comp.*saod.*other\s*than\s*high")
    if not cols:
        ctx.miss("pc", "the block 'Pvt Car-COMP & SAOD Other than Highend' is not in the grid")
        return
    comp = ctx.one(cols, "Private Car comprehensive", leaf=r"\bcomp\b")
    saod = ctx.one(cols, "Private Car SAOD", leaf=r"saod")
    cut = ctx.R.get("pc.noncb_cut")
    cap = ctx.cap("age.pc.comp", 0)
    for sub_p, col in (("1+1", comp), ("SAOD", saod)):
        if col is None:
            ctx.miss("pc", f"no {'comprehensive' if sub_p == '1+1' else 'SAOD'} column")
            continue
        for r, cl, _ in ctx.t.rows:
            v, fl = ctx.cell(r, col)
            ctx.mark("pc", "all", cl, v)
            for ncb, extra in ((("YES", 0.0), ("NO", cut)) if cut is not None else (("NA", 0.0),)):
                ctx.add("pc", cl, product="Private Car", sub=sub_p, klass="Car", makes=MK_PC, fuels=FUELS_ALL, amin=0, amax=cap, pit="On OD", rate=ctx.rate(v, "pc", extra), ncb=ncb, flags=fl)


def emit_high_end(ctx, he_rows):
    """The High-End enabler: package and SAOD payouts per RTO cluster and model cluster, as given (no slab)."""
    if not he_rows:
        return
    cap = ctx.cap("age.pc.comp", 0)
    zero_nonncb = ctx.R.get("he.noncb_zero")
    known = set(ctx.clusters)
    for cl, make, model, pkg, saod in he_rows:
        if cl not in known:
            ctx.warnings.append(f"High-End enabler: the cluster '{cl}' is not in the rate sheet - its rows are written anyway")
        base = re.sub(r"\s*\(.*?\)\s*", " ", model).strip()
        name = HE_GROUP_PREFIX + re.sub(r"\s+", "_", f"{make} {base}")
        if re.search(r"petrol\s*/\s*hybrid", model, re.I):
            fuels = ["Petrol", "Electric"]
        elif re.search(r"diesel\s*/\s*cng\s*/\s*lpg", model, re.I):
            fuels = ["Diesel", "CNG", "LPG"]
        else:
            ctx.warnings.append(f"High-End enabler: the model cluster '{model}' does not say its fuels (Petrol / Hybrid or Diesel / CNG / LPG) - not written")
            ctx.not_written.append(dict(sheet="High-End enabler", column="", segment=f"{make} {model}", reason="no fuel in the model cluster text"))
            continue
        for sub_p, val in (("1+1", pkg), ("SAOD", saod)):
            for ncb in (("YES", "NO") if zero_nonncb else ("NA",)):
                rate = 0.0 if (ncb == "NO") else (0.0 if val is None else float(val))
                ctx.add("pc_he", cl, product="Private Car", sub=sub_p, klass="Car", makes=name, fuels=fuels, amin=0, amax=cap, pit="On OD", rate=rate, ncb=ncb)


def emit_pcv3w(ctx):
    cols = ctx.t.find_cols(top=r"pcv\s*3w")
    if not cols:
        ctx.miss("pcv3w", "the block 'PCV 3W' is not in the grid")
        return
    m = re.search(r"carrying\s*capacity\s*(\d+)\s*\+\s*(\d+)", ctx.t.lv(cols[0], "top"), re.I)
    scmax = round(int(m.group(1)) + int(m.group(2)) + 0.01, 2) if m else SC_RANGE[1]
    cap_c, cap_s = ctx.cap("age.pcv3w", 0), ctx.cap("age.pcv3w", 1)
    for fuels, grp in ((FUELS_NON_DIESEL_3W, r"non\s*diesel"), (["Diesel"], r"^diesel")):
        sub_cols = [c for c in cols if re.search(grp, ctx.t.lv(c, "row3"), re.I)]
        for agekey, rx, amin, tnc in (("new", r"^new", 0, TNC_NEW), ("old", r"^non\s*new", 1, TNC_OLD)):
            comp, satp = ctx.pair(sub_cols, rx, "pcv3w", f"PCV 3W {grp} {agekey}")
            for sub_p, col, pit, cap in (("1+1", comp, "On Net", cap_c), ("STP", satp, "On TP", cap_s)):
                if col is None:
                    ctx.miss("pcv3w", f"no {'Comp' if sub_p == '1+1' else 'SATP'} column for {fuels[0]} {agekey}")
                    continue
                for r, cl, _ in ctx.t.rows:
                    v, fl = ctx.cell(r, col)
                    ctx.mark("pcv3w", agekey, cl, v)
                    ctx.add("pcv3w", cl, product="PCV 3W", sub=sub_p, klass="Passenger Carrying Rickshaw", makes=MK_PCV3W, fuels=fuels, amin=amin, amax=1 if agekey == "new" else cap, pit=pit,
                            rate=ctx.rate(v, "cv"), scmax=scmax, tnc=tnc, agekey=agekey, flags=fl)


def emit_taxi(ctx):
    cols = ctx.t.find_cols(top=r"pcv\s*taxi")
    if not cols:
        ctx.miss("taxi", "the block 'PCV Taxi' is not in the grid")
        return
    sizes = [int(a) + int(b) for a, b in re.findall(r"(\d+)\s*\+\s*(\d+)", ctx.t.lv(cols[0], "top"))]
    scmax = round(max(sizes) + 0.01, 2) if sizes else SC_RANGE[1]
    labels = list(dict.fromkeys(ctx.t.lv(c, "row4") for c in cols))
    bands = bands_from_labels(labels)
    cut = ctx.R.get("taxi.noncb_cut")
    cap_c, cap_s = ctx.cap("age.taxi", 0), ctx.cap("age.taxi", 1)
    for lab in labels:
        if lab not in bands:
            continue
        ccmin, ccmax = bands[lab]
        bc = [c for c in cols if ctx.t.lv(c, "row4") == lab]
        nil = ctx.one(bc, f"taxi {lab} nil-dep comprehensive", leaf=r"^comp\W*nill?\s*dep")
        non = ctx.one(bc, f"taxi {lab} non-nil-dep comprehensive", leaf=r"non\s*nill?\s*dep")
        satp = ctx.one(bc, f"taxi {lab} SATP", leaf=r"^satp")
        for r, cl, _ in ctx.t.rows:
            if nil is None or non is None:
                ctx.miss("taxi", f"{lab}: a comprehensive column (nil dep / non nil dep) is missing")
                break
            for ncb, extra in ((("YES", 0.0), ("NO", cut)) if cut is not None else (("NA", 0.0),)):
                for zd, col in (("YES", nil), ("NO", non)):
                    v, fl = ctx.cell(r, col)
                    ctx.mark("taxi", lab, cl, v)
                    ctx.add("taxi", cl, product="PCV 4W", sub="1+1", klass="Taxi", makes=MK_TAXI, fuels=FUELS_ALL, amin=0, amax=cap_c, pit="On Net", rate=ctx.rate(v, "cv", extra),
                            ccmin=ccmin, ccmax=ccmax, scmax=scmax, ncb=ncb, zd=zd, agekey=lab, flags=fl)
            if satp is not None:
                v, fl = ctx.cell(r, satp)
                ctx.add("taxi", cl, product="PCV 4W", sub="STP", klass="Taxi", makes=MK_TAXI, fuels=FUELS_ALL, amin=0, amax=cap_s, pit="On TP", rate=ctx.rate(v, "cv"), ccmin=ccmin, ccmax=ccmax,
                        scmax=scmax, agekey=lab, flags=fl)


def emit_tw(ctx):
    cols = ctx.t.find_cols(top=r"^tw\b")
    if not cols:
        ctx.miss("tw", "the block 'TW' is not in the grid")
        return
    saod_cols = [c for c in cols if re.search(r"saod", ctx.t.lv(c, "row4"), re.I)]
    pkg_cols = [c for c in cols if re.search(r"1\s*\+\s*1|package", ctx.t.lv(c, "row4"), re.I)]
    makes = ctx.R.get("tw.makes") or TW_MAKE_ORDER
    ccmax = round(float(ctx.R["tw.cc_max"]) + 0.01, 2) if ctx.R.get("tw.cc_max") else CC_OPEN_END
    cap_c, cap_s = round(float(ctx.R["age.tw.comp"]) + 0.01, 2) if ctx.R.get("age.tw.comp") else AGE_OPEN_END, round(float(ctx.R["age.tw.satp"]) + 0.01, 2) if ctx.R.get("age.tw.satp") else AGE_OPEN_END
    for klass, word in (("BIKE", r"bike"), ("SCOOTER", r"scooter")):
        sc = ctx.one(saod_cols, f"two-wheeler SAOD {klass}", leaf=word)
        pk = ctx.one(pkg_cols, f"two-wheeler 1+1 & SATP {klass}", leaf=word)
        mk = ",".join(m for m in makes if klass == "BIKE" or m not in SCOOTER_WITHOUT)
        for sub_p, col, pit, cap in (("1+1", pk, "On Net", cap_c), ("SAOD", sc, "On OD", cap_c), ("STP", pk, "On TP", cap_s)):
            if col is None:
                ctx.miss("tw", f"no column for {klass} {sub_p}")
                continue
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col)
                ctx.mark("tw", "all", cl, v)
                ctx.add("tw", cl, product="TW", sub=sub_p, klass=klass, makes=mk, fuels=FUELS_TW, amin=0, amax=cap, pit=pit, rate=ctx.rate(v, "tw"), ccmax=ccmax, tnc=ctx.R.get("tw.tnc") or "", flags=fl)


def emit_miscd(ctx):
    cols = ctx.t.find_cols(top=r"agricultural\s*tractor")
    if not cols:
        ctx.miss("miscd", "the block 'Agricultural Tractor & Harvester' is not in the grid")
        return
    cap = ctx.cap("age.tractor", 0)
    for klass, word, tnc_new, tnc_old, tnc_stp in (("MISCD-Tractor", r"^tractor", TNC_NEW, TNC_OLD, ""),
                                                   ("MISCD-Others", r"^harvester", f"{TNC_NEW},{TNC_HARVESTER}", f"{TNC_OLD},{TNC_HARVESTER}", TNC_HARVESTER)):
        new = ctx.one(cols, f"{klass} new comprehensive", leaf=word, row4=r"^new\s*comp")
        old = ctx.one(cols, f"{klass} non-new comprehensive", leaf=word, row4=r"^non\s*new\s*comp")
        satp = ctx.one(cols, f"{klass} SATP", leaf=word, row4=r"^satp")
        for sub_p, col, pit, amin, amax, tnc, agekey in (("1+1", new, "On Net", 0, 1, tnc_new, "new"), ("1+1", old, "On Net", 1, cap, tnc_old, "old"), ("STP", satp, "On TP", 0, cap, tnc_stp, "all")):
            if col is None:
                ctx.miss("miscd", f"no column for {klass} {sub_p} {agekey}")
                continue
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col)
                ctx.mark("miscd", agekey, cl, v)
                ctx.add("miscd", cl, product="MISCD", sub=sub_p, klass=klass, makes=MK_MISCD, fuels=FUELS_ALL, amin=amin, amax=amax, pit=pit, rate=ctx.rate(v, "cv"), tnc=tnc, agekey=agekey, flags=fl)


def emit_pc_satp(ctx):
    cols = ctx.t.find_cols(top=r"pvt\s*car\s*satp")
    if not cols:
        ctx.miss("pc_satp_p", "the block 'Pvt Car SATP' is not in the grid")
        return
    sat = ctx.R.get("pc.satp")                       # (cut, from age, to age, below age)
    cap = ctx.cap("age.pc.satp", 0)
    labels = list(dict.fromkeys(ctx.t.lv(c, "leaf") for c in cols))
    bands = bands_from_labels(labels)
    for seg, grp, fuels in (("pc_satp_p", r"petrol", ["Petrol", "Electric"]), ("pc_satp_d", r"diesel", ["Diesel", "CNG", "LPG"])):
        gcols = [c for c in cols if re.search(grp, ctx.t.lv(c, "row4"), re.I)]
        for lab, (ccmin, ccmax) in bands.items():
            col = ctx.one(gcols, f"Private Car SATP {grp} {lab}", leaf=re.escape(lab))
            if col is None:
                ctx.miss(seg, f"no column for the cc band '{lab}'")
                continue
            # the printed values are for vehicles of 10+ years; ages 1-10 are lower by the note's cut, below 1 year the grid is zero
            if sat:
                cut, lo, hi, under = sat
                ages = [(0, float(under), 0.0, True, "under1"), (float(lo), round(hi + 0.01, 2), cut, False, "mid"), (round(hi + 0.01, 2), cap, 0.0, False, "old")]
            else:
                ages = [(0, cap, 0.0, False, "all")]
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col)
                for amin, amax, extra, zero, agekey in ages:
                    if agekey in ("old", "all"):
                        ctx.mark(seg, "all", cl, v)
                    ctx.add(seg, cl, product="Private Car", sub="STP", klass="Car", makes=MK_PC_TP, fuels=fuels, amin=amin, amax=amax, pit="On TP", rate=0.0 if zero else ctx.rate(v, "pc", extra),
                            ccmin=ccmin, ccmax=ccmax, agekey=agekey, flags=fl)


def emit_school(ctx):
    cols = ctx.t.find_cols(top=r"school\s*bus\s*-\s*\d+")
    if not cols:
        ctx.miss("school_18", "the blocks 'School Bus' are not in the grid")
        return
    tops = list(dict.fromkeys(ctx.t.lv(c, "top") for c in cols))
    starts = sorted((int(re.search(r"school\s*bus\s*-\s*(\d+)", t, re.I).group(1)), t) for t in tops)
    cap = ctx.cap("age.bus", 0)
    for i, (lo, top) in enumerate(starts):
        scmin = round(lo + 0.01, 2)
        scmax = round(starts[i + 1][0] + 0.01, 2) if i + 1 < len(starts) else SC_RANGE[1]
        seg = "school_18" if i + 1 == len(starts) else "school_12"
        bc = [c for c in cols if ctx.t.lv(c, "top") == top]
        comp = ctx.one(bc, f"school bus {top} comprehensive", leaf=r"^comp")
        satp = ctx.one(bc, f"school bus {top} SATP", leaf=r"^satp")
        for sub_p, col, pit in (("1+1", comp, "On Net"), ("STP", satp, "On TP")):
            if col is None:
                ctx.miss(seg, f"no {'Comp' if sub_p == '1+1' else 'SATP'} column")
                continue
            for r, cl, _ in ctx.t.rows:
                v, fl = ctx.cell(r, col)
                ctx.mark(seg, "all", cl, v)
                ctx.add(seg, cl, product="PCV 4W", sub=sub_p, klass="School Bus", makes=MK_BUS, fuels=FUELS_ALL, amin=0, amax=cap, pit=pit, rate=ctx.rate(v, "cv"), scmin=scmin, scmax=scmax, flags=fl)


# ---- the notes that are neither understood nor standing: flagged on the rows they may affect
SEG_WORDS = [
    (r"upto\s*2\.?5\s*t|up\s*to\s*2\.5\s*t|upto-?\s*2\s*t\b", {"gcv_u25", "gcv_tractor"}), (r"2\.5\s*t?\s*(?:to|-)\s*3\.5\s*t?", {"gcv_25_35"}), (r"3\.5\s*t?\s*(?:to|-)\s*5(?:\.0)?\s*t?\b", {"gcv_35_50"}),
    (r"(?<![\d.])5(?:\.0)?\s*t?\s*(?:to|-)\s*7\.5\s*t?", {"gcv_50_75"}), (r"7\.5\s*t?\s*(?:to|-)\s*12(?:\.0)?\s*t?", {"gcv_75_120"}), (r"12\s*t?(?:on)?\s*(?:to|-)\s*20\s*t?", {"gcv_12_20"}),
    (r"12\s*t?(?:on)?\s*(?:to|-)\s*40\s*t?", {"gcv_12_20", "gcv_20_40"}), (r"(?<![\d.])20\s*t?(?:on)?\s*(?:to|-)\s*40\s*t?", {"gcv_20_40"}), (r"(?:above|>)\s*40\s*t", {"gcv_40"}),
    (r"gcv\s*3w", {"gcv3w"}), (r"pcv\s*3w", {"pcv3w"}), (r"taxi", {"taxi"}), (r"school\s*bus|staff\s*bus", {"school_12", "school_18"}), (r"\btw\b|two[- ]?wheeler", {"tw"}),
    (r"tractor|harvester", {"miscd", "gcv_tractor"}), (r"pvt\s*car|private\s*car", {"pc", "pc_he", "pc_satp_p", "pc_satp_d"}),
    (r"models\s+3718\s+and\s+3723", {"gcv_20_40", "gcv_40"}),                     # 'the 20 t-40 t grid applies to these models': they are the above-40 t vehicles
]
STATE_WORDS = {"j&k": "JK", "delhi": "DL", "chennai": "TN", "tamil": "TN", "kolkat": "WB", "bengal": "WB", "bangalore": "KA", "mumbai": "MH", "uttar pradesh": "UP", "haryana": "HR", "punjab": "PB"}


def cluster_prefix(cluster):
    """'MH - M' -> 'MH', 'GA' -> 'GA', 'North East 1' -> 'NE'"""
    m = re.match(r"([A-Z]{2})\b", cluster)
    return m.group(1) if m else "".join(w[0] for w in cluster.split()[:2]).upper()


def note_scope(text, prefixes):
    """(segment keys or None, state prefixes or None) a note may affect: the segments and states it names; None = not named (every one)"""
    segs = set()
    for rx, s in SEG_WORDS:
        if re.search(rx, text, re.I):
            segs |= s
    if re.search(r"\bgcv\b", text, re.I) and not (segs & GCV_SEGS):
        segs |= GCV_SEGS
    states = {p for p in re.findall(r"(?<![A-Za-z])([A-Z]{2})(?=\d|\b|\s*-)", text) if p in prefixes}
    states |= {v for k, v in STATE_WORDS.items() if k in text.lower() and v in prefixes}
    return (segs or None), (states or None)


def settle_duplicates(df, key_cols, warnings):
    """Rows with the same key: the same payin twice is one row; two grid cells giving one entry different payins: the LOWEST is written (the user's rule, 27 Oct 2026); each such entry is listed in LAST_LOWEST."""
    global LAST_LOWEST
    LAST_LOWEST = []
    keys = df[key_cols].astype(str).agg("|".join, axis=1)
    dup = keys.duplicated(keep=False)
    if not dup.any():
        return df
    value = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].astype(float).sum(axis=1)
    drop = []
    for _, idx in df[dup].groupby(keys[dup], sort=False).groups.items():
        idx = list(idx)
        best = min(idx, key=lambda i: (value[i], idx.index(i)))
        drop.extend(i for i in idx if i != best)
        found = list(dict.fromkeys(fmt(value[i]) for i in idx))
        if len(found) > 1:
            row = df.loc[best]
            LAST_LOWEST.append(dict(product=row["product"], sub_product=row["sub_product"], rto=row["new_rto_list"], makes=row["new_vehicle_makes"], fuel=row["fuel_type"],
                                    age=f"{fmt(row['vehicle_age_min'])}-{fmt(row['vehicle_age_max'])}", payins=" / ".join(found), written=fmt(value[best])))
    if LAST_LOWEST:
        warnings.append(f"{len(LAST_LOWEST)} entries had two different payins in the grid: the lowest is written (listed in the sheet 'Lowest rate chosen')")
    return df.drop(index=drop)


def fill_columns(df, dates):
    """the constant columns, the payout (pi - 7, never below 0) and the column order"""
    df = df.copy()
    for c in ("pi_od_rate", "pi_tp_rate", "pi_net_rate"):
        df[c] = df[c].astype(float).map(lambda x: round_half_up(x))
        df[c.replace("pi_", "po_")] = df[c].map(lambda x: round_half_up(max(x - PO_CUT, 0.0)) if x > 0 else 0.0)
    df["id"], df["insurance_company"], df["insurer_vertical"] = ROW_ID, INSURANCE_COMPANY, "Not Applicable"
    for c in ("pi_tp_2", "pi_tp_3", "pi_tp_4", "pi_tp_5", "pi_flat_amount", "pi_vli", "po_flat_amount"):
        df[c] = 0.0
    df["po_type"], df["from_date"], df["to_date"], df["user_id"], df["veh_use"], df["remarks"] = df["pi_type"], dates[0], dates[1], USER_ID, "", ""
    df["tariff_min"] = df["tariff_min"].astype(float)
    return df


def _short(text):
    t = re.sub(r"\s+", " ", str(text)).strip()
    return t if len(t) <= MAX_NOTE_CHARS else t[:MAX_NOTE_CHARS - 3] + "..."


SLAB_ARGS = {"pc": None, "cv": None}      # --pc-slab / --cv-slab: an answer given on the command line


def choose_slab(ask, fam, label, slabs, irda, warnings):
    """The premium slab of a family, asked every month (the user, 1 Nov 2026). Returns the points the slab lowers every payout by."""
    if not slabs:
        preset = SLAB_ARGS.get(fam)
        if preset is not None and label_key(preset) in ("none", "0"):
            return 0.0
        pick = ask.choose(f"the {label} slab note has no slab this script can read",
                          [("Write the grid as printed (no slab reduction)", "none"), ("Stop", "stop")],
                          question=f"The grid has no readable '{label} Slab' note (For <range> Grid would be lower by N%). Write the {label} grid as printed, without a slab reduction?",
                          short=f"{label} slab note", memo=f"slab note {fam}")
        if pick == "stop":
            raise GridError(f"stopped by you - the {label} slab note could not be read")
        warnings.append(f"no readable '{label} Slab' note was found in the grid: the {label} payouts are written as printed (no slab reduction)")
        return 0.0
    options = [("No reduction - the grid as printed (a premium above the slabs listed)", 0.0)] + [(f"{name}: the grid is lower by {cut:g}%", cut) for name, cut in slabs]
    preset = SLAB_ARGS.get(fam)
    if preset is not None:
        key = label_key(preset)
        for (lab, val), name in zip(options, ["none"] + [n for n, _ in slabs]):
            if key in (label_key(name), label_key(lab)) or (key == "0" and val == 0.0):
                return val
        raise GridError(f"--{fam}-slab '{preset}' is not one of: none, {', '.join(n for n, _ in slabs)}")
    extra = f" The slab '<{irda}' (only IRDA / 10% on SATP) is not built." if irda else ""
    return ask.choose(f"the premium slab for the {label} grid is needed",
                      options, question=f"Which premium slab is your business in for the {label} grid this month? A slab lowers EVERY payout of these products by its points (the grid's own note lists them).{extra}",
                      short=f"{label} premium slab", memo=f"slab {fam}")


def build(grid_path, from_date, to_date, log=say, ask=None, wb=None):
    """Read the workbook and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    from_date / to_date are datetime.date values; they are written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever the grid itself is missing something. The notes are never asked about: what they say is applied, and
    what cannot be applied is flagged in the remarks column (see ManualReview)."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    if wb is None:
        log(f"Reading {os.path.basename(str(grid_path))}...")
        grid_path, wb = open_file(ask, grid_path, "the SBI workbook", read_workbook, EXCEL_EXTS)
    hidden = hidden_sheets(wb)
    log(f"Sheets: {len(wb.sheetnames) - len(hidden)} visible, {len(hidden)} hidden (hidden sheets are older versions or helpers and are not read)")
    warnings = []
    masters = load_masters(ask)
    review = ManualReview()
    SPECIAL_GROUPS.clear()
    name, hdr, cols, _ = rate_sheet(wb, ask, warnings, log)
    table = RateTable(wb, name, hdr, cols)
    seen = {}
    for r, cl, _ in table.rows:
        if cl in seen:
            pick = ask.choose(f"the cluster '{cl}' appears twice in the rate sheet (rows {seen[cl]} and {r})", [(f"Use the first (row {seen[cl]})", "first"), (f"Use the second (row {r})", "second")],
                              question=f"The cluster '{cl}' is in the rate sheet twice (rows {seen[cl]} and {r}). Which one is right?", short=f"cluster {cl}")
            if pick == "first":
                table.rows = [x for x in table.rows if x[0] != r]
            else:
                table.rows = [x for x in table.rows if x[0] != seen[cl]]
        seen[cl] = r
    log(f"Rate sheet '{name}': {len(table.rows)} clusters, header row {hdr + 1}")
    notes = collect_notes(table)
    rules = read_rules(notes, ask, warnings, log)
    if rules.get("tw.makes") is None:
        pick = ask.choose("the note listing the allowed two-wheeler makes was not found", [("Use last month's list: " + ", ".join(TW_MAKE_ORDER), "last"), ("Stop", "stop")],
                          question="I can't find the note 'Declined Makes - All Makes other than ...' for the two-wheelers. Use last month's make list?", short="two-wheeler makes", memo="rule tw.makes")
        if pick == "stop":
            raise GridError("stopped by you - the two-wheeler make note is missing")
        rules["tw.makes"] = list(TW_MAKE_ORDER)
        warnings.append("the two-wheeler make note was not found: last month's make list is used")
    if rules.get("tw.makes_unknown"):
        warnings.append(f"the two-wheeler make note names {rules['tw.makes_unknown']} which this script does not know: they are not in the make group (add them to TW_MAKE_NAMES)")
    names = [cl for _, cl, _ in table.rows]
    declines = read_declines(notes, names)
    state_declines = read_state_declines(notes)
    pan = read_pan_declines(notes)
    mapping = read_mapping(wb, {name}, warnings, log)
    he_rows, he_name, he_texts = read_high_end(wb, {name}, warnings, log)
    for text in he_texts:
        n = Note(0, "High-End enabler", text)
        n.segs = {"pc_he"}                                                  # a text of that sheet concerns the High-End rows
        notes.append(n)
    if he_rows is None:
        warnings.append("no sheet holds the High-End enabler (RTO Cluster Name | Make Cluster | Model Cluster | Package | SAOD): no High-End rows were written")
    if mapping[0]:
        only_grid = sorted(set(names) - set(mapping[0]))
        only_map = sorted(set(mapping[0]) - set(names))
        if only_grid:
            warnings.append(f"cluster(s) of the rate sheet that the RTO list does not know: {only_grid}")
        if only_map:
            warnings.append(f"cluster(s) of the RTO list that the rate sheet does not have: {only_map}")
    finger, lists = decline_lists_fingerprint(wb, {name, he_name})
    if lists and finger != DECLINE_FINGERPRINT:
        warnings.append(f"the declined make / model lists ({', '.join(lists)}) changed since the SBI make groups were built (fingerprint {finger}, expected {DECLINE_FINGERPRINT}): check the make groups")
    if not lists:
        warnings.append("the declined make / model lists (Make | Model tables) were not found")
    if rules.get("period"):
        a, b = rules["period"]
        if from_date != a:
            warnings.append(f"the grid says it applies from {a:%d/%m/%Y} to {b:%d/%m/%Y}; from_date is {from_date:%d/%m/%Y}")
        else:
            log(f"The grid's own period: {a:%d/%m/%Y} to {b:%d/%m/%Y} (from_date matches)")
    # 'Pvt Car Slab' is the Private Car slab and 'CV Slab' the commercial-vehicle one; the grid has NO two-wheeler slab, so the two-wheeler payouts are written as printed (the user, 3 Nov 2026: 'there is no
    # slab deduction applicable for two-wheelers' - the November upload took the Private Car slab off every two-wheeler payout).  A 'TW Slab' note would be a new text: flagged, not guessed.
    slabs = {"pc": choose_slab(ask, "pc", "Private Car", rules.get("slabs.pc"), rules.get("slab_irda.pc"), warnings),
             "cv": choose_slab(ask, "cv", "commercial-vehicle", rules.get("slabs.cv"), rules.get("slab_irda.cv"), warnings),
             "tw": 0.0}
    for key, fuels in sorted(rules["fuel_declines"].items()):
        ask.decide(f"Per the grid's notes: {', '.join(sorted(fuels))} declined for {FUEL_PRODUCT_NAMES[key]} - no such rows are written")
    for key, fuels in LAST_FUEL_DECLINES.items():
        lost = sorted(set(fuels) - rules["fuel_declines"].get(key, set()))
        if lost:
            warnings.append(f"{FUEL_PRODUCT_NAMES[key]}: the last grid declined {', '.join(lost)} but this grid has no such note - those rows ARE written; check with SBI that the restriction is lifted")
    ctx = Ctx(table, rules, slabs, mapping, masters, warnings)
    for u in ctx.uplifts:
        others = sorted(c for cl in u["clusters"] for c in mapping[0].get(cl, ()) if not re.match(rf"{u['state']}\d", c))
        ask.decide(f"Per the grid's notes: +{u['pts']:g} points on GCV 4W {u['lo']:g}-{u['hi']:g} t for {u['state']} registrations - written for the WHOLE cluster {', '.join(sorted(u['clusters']))}"
                   + (f", which also holds {len(others)} RTOs of other states ({', '.join(others[:6])}{' ...' if len(others) > 6 else ''})" if others else ""))
    have = {(u["state"], u["lo"], u["hi"]) for u in rules["uplifts"]}
    for (state, lo, hi), pts in LAST_UPLIFTS.items():
        if (state, lo, hi) not in have:
            warnings.append(f"{state} registrations: the last grid gave {pts:g} extra points on GCV 4W {lo:g}-{hi:g} t but this grid has no such note - those rows are written without them; check with SBI")
    emit_gcv(ctx)
    emit_pc(ctx)
    emit_high_end(ctx, he_rows)
    emit_pcv3w(ctx)
    emit_taxi(ctx)
    emit_tw(ctx)
    emit_miscd(ctx)
    emit_pc_satp(ctx)
    emit_school(ctx)
    if any(r["_seg"] == "tw" for r in ctx.rows):
        ask.decide("Per the grid: no premium slab for two-wheelers (the slab notes are 'Pvt Car Slab' and 'CV Slab') - the two-wheeler payouts are written as printed")
    holds = set(table.numeric_columns())
    for c in range(table.c_state + 1, table.width + 1):
        if table.chain[c] and c not in table.claimed:
            warnings.append(f"column {col_letter(c)} ('{table.path(c)}') " + ("holds numbers but" if c in holds else "is labelled but")
                            + " no rule of this script reads it: nothing is written for it")
            ctx.not_written.append(dict(sheet=name, column=col_letter(c), segment=table.path(c), reason="no rule for this column"))
    if not ctx.rows:
        raise GridError("nothing to write - no block of the rate sheet could be read")

    df = pd.DataFrame(ctx.rows)
    df = df.explode("_fuels").rename(columns={"_fuels": "fuel_type"}).reset_index(drop=True)
    df["_flags"] = df["_flags"].map(tuple)
    df = settle_duplicates(df, KEY_COLUMNS, warnings).reset_index(drop=True)
    df = fill_columns(df, dates)

    flag_texts = {}
    for i, fl in enumerate(df["_flags"]):
        for text in fl:
            flag_texts.setdefault(text, []).append(i)
    for text, idx in sorted(flag_texts.items(), key=lambda kv: -len(kv[1])):
        mask = np.zeros(len(df), dtype=bool)
        mask[idx] = True
        review.flag(_short(text), mask)
    seg_of, cl_of, age_of = df["_seg"].values, df["_cl"].values, df["_age"].values
    prefixes = {cluster_prefix(c) for c in names}

    def mask_of(segs=None, clusters=None, states=None, ages=None):
        m = np.ones(len(df), dtype=bool)
        if segs is not None:
            m &= np.isin(seg_of, list(segs))
        if clusters is not None:
            m &= np.isin(cl_of, list(clusters))
        if states is not None:
            m &= np.array([cluster_prefix(c) in states for c in cl_of], dtype=bool)
        if ages is not None:
            m &= np.isin(age_of, list(ages))
        return m
    # the declined-cluster notes against the grid
    for d in declines:
        seg = d["seg"]
        have = {a for (s, a) in ctx.zero if s == seg}
        for ages, mode, named in d["sections"]:
            use = {a for a in have if ages == ALL_AGES or a in ages or a == "all"}
            if not use:
                continue
            zero = {cl for cl in names if all(ctx.zero[(seg, a)].get(cl, True) for a in use)}
            paid = (named - zero) if mode == "declined" else ((set(names) - named) - zero)
            unlisted = (zero - named) if mode == "declined" else (named & zero)
            for cl in sorted(paid):
                review.flag(_short(f"{SEG_NAMES.get(seg, seg)}: the note " + ("declines" if mode == "declined" else "allows only the listed clusters, not") + f" '{cl}' but the grid pays there"),
                            mask_of({seg}, {cl}, None, use | {"under1"}))
            for cl in sorted(unlisted):
                review.flag(_short(f"{SEG_NAMES.get(seg, seg)}: the grid pays nothing for '{cl}' but the note " + ("does not decline it (the rest is allowed)" if mode == "declined" else "allows it")),
                            mask_of({seg}, {cl}, None, use | {"under1"}))
    for seg, states in state_declines:
        for cl in names:
            if cluster_prefix(cl) in states and not ctx.zero.get((seg, "all"), {}).get(cl, True):
                review.flag(_short(f"{SEG_NAMES.get(seg, seg)}: the note fully declines {', '.join(sorted(states))} but the grid pays for '{cl}'"), mask_of({seg}, {cl}))
    for seg in pan:
        for cl in names:
            if not ctx.zero.get((seg, "all"), {}).get(cl, True):
                review.flag(_short(f"{SEG_NAMES.get(seg, seg)}: the note declines every cluster but the grid pays for '{cl}'"), mask_of({seg}, {cl}))
    # the notes nobody understood: standing ones are dropped, the rest flagged on the rows they may affect
    leftovers = []
    for n in notes:
        text = n.rest if n.used else n.text
        text = re.sub(r"\s+", " ", text).strip(" *.-:")
        if not n.used or re.search(r"\d", text):
            leftovers.append((n, text))
    unknown = (unknown_notes([t for n, t in leftovers if not n.used], STANDING_NOTES, min_len=MIN_NOTE_CHARS)       # a whole text no rule read: even a short one is a rule until a standing note says otherwise
               + unknown_notes([t for n, t in leftovers if n.used], STANDING_NOTES, min_len=25))                                # what is left of a text a rule has read ('In PCV 3W allowed'): only a long one is worth a flag
    for n, text in leftovers:
        if text in unknown:
            segs, states = note_scope(f"{n.label} {text}", prefixes)
            review.flag(_short(text), mask_of(getattr(n, "segs", None) or segs, None, states))
    global LAST_GAPS, LAST_NOT_WRITTEN
    LAST_GAPS = dict(ctx.gaps)
    LAST_NOT_WRITTEN = list(ctx.not_written)
    globals()["MASTER_SHEET"] = name
    df = df.drop(columns=[c for c in df.columns if c.startswith("_")])
    check_names(df, review, masters, log)                                # every RTO / make group written must be in the master lists
    review.report(df, log)                                               # fills the remarks column
    return df[OUTPUT_COLUMNS], warnings


# %% ---------------------------------------------------------------- OUTPUT AND CHECKS
def fmt(x):
    x = round(float(x), 4)
    return str(int(x)) if x == int(x) else repr(x)


def to_csv_frame(df):
    out = df.copy()
    for c in NUMERIC_COLUMNS:
        out[c] = out[c].map(fmt)
    return out


def key_of(d, cols=None):
    """the comparison key of a frame of upload rows: numbers to 4 decimals, make and RTO names without regard to capitals or order ('MAHINDRA' of an old upload = 'Mahindra' of the master)"""
    d = d.copy()
    for c in ("vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "sc_min", "sc_max", "tariff_min", "tariff_max"):
        d[c] = pd.to_numeric(d[c], errors="coerce").map(lambda x: f"{x:.4f}")
    for c in ("new_vehicle_makes", "new_rto_list"):
        d[c] = d[c].astype(str).str.lower().map(lambda s: ",".join(sorted(t.strip() for t in s.split(","))))
    d["add_tnc"] = d["add_tnc"].astype(str).str.replace(r"[^\x00-\x7f]", "-", regex=True).str.replace(" - ", "-", regex=False)      # the October upload lost the dash of 'CONTRACT CARRIAGE - TRANSPORTER'
    return d[cols or KEY_COLUMNS].astype(str).agg("|".join, axis=1)


def keep_ref_dates(df, ref_path, log=say):
    """A row that is the same (key and payin) as a row of the earlier upload keeps that row's from_date: the grid has not changed for it since then."""
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    need = KEY_COLUMNS + ["pi_type", "pi_od_rate", "pi_tp_rate", "pi_net_rate", "from_date"]
    if any(c not in ref.columns for c in need):
        return df
    mine = to_csv_frame(df)
    cols = KEY_COLUMNS + ["pi_type", "pi_od_rate", "pi_tp_rate", "pi_net_rate"]
    for d in (mine, ref):
        for c in ("pi_od_rate", "pi_tp_rate", "pi_net_rate"):
            d[c] = pd.to_numeric(d[c], errors="coerce").map(lambda x: f"{x:.4f}")
    mine["_k"], ref["_k"] = key_of(mine, cols), key_of(ref, cols)
    their = ref.drop_duplicates("_k").set_index("_k")["from_date"].astype(str).str.replace("-", "/")
    old = mine["_k"].map(their)
    new_from = dt.datetime.strptime(df["from_date"].iloc[0], "%d/%m/%Y").date()

    def earlier(x):
        try:
            return dt.datetime.strptime(x, "%d/%m/%Y").date() < new_from
        except (TypeError, ValueError):
            return False
    use = old.map(earlier)
    if use.any():
        df = df.copy()
        df.loc[use.values, "from_date"] = old[use].values
        log(f"{int(use.sum()):,} rows are the same as in {os.path.basename(ref_path)} and keep its earlier from_date "
            f"({', '.join(f'{k}: {v}' for k, v in old[use].value_counts().items())})")
    return df


def diff_against_reference(df, ref_path, log=say):
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    mine = to_csv_frame(df)
    missing = [c for c in KEY_COLUMNS if c not in ref.columns]
    if missing:
        raise ValueError(f"the earlier upload has no column(s) {missing} - it does not look like an upload file")

    mine["_k"], ref["_k"] = key_of(mine), key_of(ref)
    if ref["_k"].duplicated().any():
        log(f"  note: reference has {int(ref['_k'].duplicated().sum())} duplicate keys")
    merged = mine.merge(ref.drop_duplicates("_k"), on="_k", how="outer", suffixes=("", "_ref"), indicator=True)
    only_new = merged[merged["_merge"] == "left_only"]
    only_ref = merged[merged["_merge"] == "right_only"]
    both = merged[merged["_merge"] == "both"]
    changed = pd.Series(False, index=both.index)
    detail = {}
    for c in [c for c in OUTPUT_COLUMNS if c not in ("id", "remarks") and c in ref.columns and c + "_ref" in both.columns]:
        if c in KEY_COLUMNS and c not in NUMERIC_COLUMNS:
            continue                                   # part of the key: equal by construction
        if c in NUMERIC_COLUMNS + ["user_id"]:
            d = (pd.to_numeric(both[c], errors="coerce") - pd.to_numeric(both[c + "_ref"], errors="coerce")).abs() > 1e-6
        elif c in ("from_date", "to_date"):
            d = both[c].astype(str).str.replace("-", "/") != both[c + "_ref"].astype(str).str.replace("-", "/")
        else:
            d = both[c].astype(str) != both[c + "_ref"].astype(str)
        if d.any():
            detail[c] = int(d.sum())
        changed |= d
    log(f"  vs {os.path.basename(ref_path)}: {len(only_new)} new rows, {len(only_ref)} removed rows, "
        f"{int(changed.sum())} rows with changed values (of {len(both)} matching keys)")
    for c, n in detail.items():
        ex = both[both[c].astype(str) != both[c + "_ref"].astype(str)][["_k", c, c + "_ref"]].head(3).values.tolist()
        log(f"    column {c}: {n} differ, e.g. {ex}")
    if len(only_ref):
        by = only_ref.groupby(["product_ref" if "product_ref" in only_ref else "product",
                               "sub_product_ref" if "sub_product_ref" in only_ref else "sub_product"]).size()
        log("    removed rows are: " + "; ".join(f"{n} x {p} {s}" for (p, s), n in by.head(8).items()))
    cl_new = sorted(set(mine["new_rto_list"]) - set(ref["new_rto_list"]))
    cl_old = sorted(set(ref["new_rto_list"]) - set(mine["new_rto_list"]))
    if cl_new:
        log(f"    RTO lists not in the earlier upload ({len(cl_new)}): {cl_new[:6]}{' ...' if len(cl_new) > 6 else ''} - their RTO groups must exist in the portal")
    if cl_old:
        log(f"    RTO lists in the earlier upload but not in this grid ({len(cl_old)}): {cl_old[:6]}{' ...' if len(cl_old) > 6 else ''}")
    return len(only_new), len(only_ref), int(changed.sum())


def collect_gaps(df):
    """For the review workbook: the RTO groups the master lacks WITH the RTO codes the grid's mapping sheet puts in them, and the grid rows that gave no upload row."""
    sheet = EXTRA_SHEETS.get("RTO groups to create")
    if sheet is not None and len(sheet):
        sheet = sheet.copy()
        sheet["items"] = [",".join(LAST_GAPS.get(g, ())) or ("the workbook's RTO mapping sheet has no RTO under this cluster name" if g in LAST_GAPS else it) for g, it in zip(sheet["group_needed"], sheet["items"])]
        EXTRA_SHEETS["RTO groups to create"] = sheet
    if LAST_NOT_WRITTEN:
        add_sheet("Not written", pd.DataFrame(LAST_NOT_WRITTEN).drop_duplicates())


def unique_warnings(warnings):
    """each warning once, with how many times it came up"""
    seen, order = {}, []
    for w in warnings:
        if w not in seen:
            order.append(w)
        seen[w] = seen.get(w, 0) + 1
    return [w if seen[w] == 1 else f"{w} (x{seen[w]})" for w in order]


def report(df, warnings, log=say, ask=None):
    """Print what is about to be written. Rows that repeat a key were settled in build (settle_duplicates). Returns the frame to write."""
    ask = ask or Asker()
    log(f"\nRows to write: {len(df):,}   ({df['new_rto_list'].nunique()} RTO lists)")
    counts = df.groupby(["product", "sub_product"]).size()
    for (prod, sub), n in counts.items():
        log(f"  {prod:<12} {sub:<5} {n:>7,}")
    log(f"Rows that repeat a key (two grid rows, two payouts: only the lowest is written): {int(df.duplicated(KEY_COLUMNS).sum())}")
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
    po = df[["po_od_rate", "po_tp_rate", "po_net_rate"]].max(axis=1)
    log(f"PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {int(((pi > 0) & (po == 0)).sum())}")
    for w in unique_warnings(warnings):
        log(f"WARNING: {w}")
        RUN_WARNINGS.append(w)
    return df


def main(grid_path, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None):
    ask = ask or Asker()
    df, warnings = build(grid_path, from_date, to_date, log, ask)
    df = report(df, warnings, log, ask)
    if ref_path and KEEP_REF_DATES:
        try:
            df = keep_ref_dates(df, ref_path, log)
        except Exception as e:                                       # the comparison below reports an unreadable earlier upload
            log(f"from_date of unchanged rows not carried over: {type(e).__name__}: {e}")
    diff = compare_safely(ask, diff_against_reference, df, ref_path, log)
    ask.summary(log)
    if ask.interactive:
        notes = []
        if warnings:
            notes.append("some rows or sheets are not in the output (see WARNING above)")
        if diff and (diff[0] + diff[1]) > 0.25 * len(df):
            notes.append(f"{diff[0]} rows are new and {diff[1]} rows are gone compared with the earlier upload")
        question = ("Heads-up: " + "; ".join(notes) + ". Write the file anyway?" if notes
                    else "Everything look right? Write the file now?")
        if not ask.confirm(question, True):
            raise GridError("stopped by you - nothing was written")
    collect_gaps(df)
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"SBI_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
    log(f"\nWrote {out_path}")
    return out_path


# %% ---------------------------------------------------------------- ENTRY POINTS
def collect_uploads(ask):
    """Colab: upload the files; asks again if nothing usable came up. Returns (workbook, earlier upload or None)."""
    while True:
        uploaded = colab_upload(ask)
        books, csvs, ignored = [], [], []
        for name, data in (uploaded or {}).items():
            try:
                with open(name, "wb") as fh:
                    fh.write(data)
            except OSError as e:
                raise GridError(f"cannot save the uploaded file {name}: {e}")
            low = name.lower()
            (csvs if low.endswith(".csv") else books if low.endswith(EXCEL_EXTS) else ignored).append(name)
        if ignored:
            say(f"Ignored (neither a workbook nor a CSV): {ignored}")
        if books:
            break
        if not ask.confirm("No workbook (.xls / .xlsx) came with that upload. Upload again?", True):
            raise GridError("no workbook was uploaded")
    grid = books[0] if len(books) == 1 else ask.choose(
        "more than one workbook was uploaded", [(n, n) for n in books],
        question="You uploaded more than one workbook. Which one is this month's SBI grid?", short="workbook")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the SBI broker grid (.xlsx) and, optionally, an earlier upload (.csv) to compare with")
    try:
        grid, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("sbi_grid_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="SBI General broker grid -> SBI upload CSV")
    ap.add_argument("--grid", help="SBI workbook (.xlsx) - asked for if left out")
    ap.add_argument("--pc-slab", help="the premium slab of the Private Car grid (two-wheelers have none), as named in the grid's 'Pvt Car Slab' note (1L-3L ...) or 'none' - asked for if left out")
    ap.add_argument("--cv-slab", help="the premium slab of the commercial-vehicle grid, as named in the grid's 'CV Slab' note (3L-4.99L ...) or 'none' - asked for if left out")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    SLAB_ARGS["pc"], SLAB_ARGS["cv"] = a.pc_slab, a.cv_slab
    try:
        grid = a.grid or ask_file(ask, "no workbook given - pass --grid <workbook.xls>",
                                  "Path of the SBI workbook (.xlsx) - you can drag the file into this window:", exts=EXCEL_EXTS)
        grid = clean_path(grid)
        if not os.path.isfile(grid):
            raise GridError(f"workbook not found: {grid}")
        ref = a.ref
        if ref is None and ask.interactive:
            try:
                ref = ask_file(ask, "", "Earlier upload CSV to compare with (path), or press Enter to skip:",
                               exts=(".csv",), blank_ok=True) or None
            except NeedsAnswer:                   # no keyboard: the comparison is optional, just skip it
                ref = None
        if ref:
            ref = clean_path(ref)
        if ref and not os.path.isfile(ref):
            say(f"Note: the earlier upload {ref} was not found - skipping the comparison.")
            ref = None
        if a.from_date and a.to_date:
            from_date, to_date = parse_ddmmyyyy(a.from_date), parse_ddmmyyyy(a.to_date)
            if to_date < from_date:
                raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
            if not check_dates(from_date, to_date, ask):
                raise GridError("stopped by you - check the dates")
        elif ask.interactive:
            from_date, to_date = ask_date_range(ask=ask)
        else:
            sys.exit("STOPPED - pass --from-date and --to-date (DD/MM/YYYY); they cannot be asked for here.")
        main(grid, from_date, to_date, a.out, ref, ask=ask)
    except (GridError, ValueError) as e:
        sys.exit(f"STOPPED - {e}")
    except KeyboardInterrupt:
        sys.exit("STOPPED - stopped by you. No file was written.")
    except Exception as e:
        saved = crash_report("sbi_grid_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
