"""
ICICI Lombard rate grids (CV, Private Car, Two-Wheeler) -> ICICI upload CSV
============================================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (NN).csv") from the workbooks ICICI Lombard circulates every month.  Rates come only from those workbooks.
Give it any of the four workbooks (`--cv-grid`, `--pc-grid`, `--tw-grid`, `--tw-new-grid`, or `--grid` once per file: each workbook is recognised by what its sheets SAY, not by its name):

* CV grid  ("CV Grid <month> v2")      sheets: the comprehensive matrix (COMP / CV_Comp / ' COMP'), the third-party matrix (AOTP / CV_AOTP / TP), the Council Call list (CC / COUNCIL CALL),
                                       the RTO sub-cluster list (RTO MASTER / CV RTO), the Negative RTO list, CV NOTES.  One row per RTO cluster, one column per segment (GCV 3W, SCV < / >= 2450 kg,
                                       LCV, MHCV by tonnage and body, MISCD CE, tractor, PCV 3W, school / staff bus by capacity, taxi, PCV 2W).  Comprehensive = 1+1 / On Net, AOTP = STP / On TP.
* Private Car grid ("Pvt Car Grid")   one sheet, one row per cluster: New (1+3, 3+3) | SAOD with NCB | Comprehensive with NCB (Petrol, CNG, Diesel, Electric each) | Act only (Net GWP) |
                                       SAOD and Comprehensive without NCB | Used car (not converted).  All payouts are on OD except the Act-only column (On TP).
* Two-Wheeler grid ("TW Grid")        sheet 'TW Old': Bike / Scooter x Comprehensive (1+1; the same grid serves 2+2 and 3+3), SAOD, AOTP.
* New two-wheeler grid ("NEW 15 ...") one sheet per OEM (HMC = Hero, HMSI = Honda, TVS, Suzuki, Yamaha, RE = Royal Enfield, EV): cluster x Bike / Scooter, the payout of a new two-wheeler (1+5).

What a cell becomes (checked against the user's October 2026 CV upload, 20,112 rows, and their November 2026 Private Car + Two-Wheeler upload, 3,318 rows)
--------------------------------------------------------------------------------------------------------------------------------------------------------------
* A plain percentage is one rate card per fuel; `No Biz`, `0%` and a blank are ZERO cards (the user's uploads have one for every cluster).  Fuels: goods and bus segments carry all five,
  `GCV 3W Electric` / `PCV 3W Electric` / `PCVTAXI_ELECTRIC` carry Electric only, Non-Diesel carries Petrol / CNG / LPG, Diesel carries Diesel, taxi carries the four non-electric fuels.
* The text inside a CV cell ("OLD:10%,NEW:0%", "TATA|M&M:45%,OTHERS:35%", "OLD|1-3 yrs:40%,OLD|>=4 yrs:50%,NEW:0%", "COMP:40%,AOTP:30%", "COMP|NEW|NZD:25%,COMP|NEW|ZD:20%") is read by a small
  grammar (`parse_cell`): clauses are separated by commas, the words of a clause by `|`, the rate follows `:`.  A word is a policy (COMP / AOTP), a transaction (NEW = age 0-1, tag `New`;
  OLD = age 1-99, tag `Old`), an age band, a zero-dep mark (ZD = YES, NZD = NO), a make, a vehicle type that no make group can carry (CRANES, ACE, ESCORTS, Fuel Tanker ... go to add_tnc),
  a place (a sub-cluster code such as HP1, a district) or OTHERS = everything that no other clause of the cell covers.  A clause is a box in (policy, zero-dep, place, body, make, age);
  OTHERS is the box minus the others (`regions`).  Age labels follow the convention of the Shriram scripts: 'a-b yrs' = (a-1)+0.01 .. b+0.01, '10+ yrs' = 10.01 .. 99, '>=4 yrs' = 3.01 .. 99.
  A make set is written as the master's group (`BUS_MAGIC_FORCE_SUMO`), as a list of the master's single-make groups (`TATA,Mahindra,Maruti`) or - for 'everything but' - as the master's
  `..._EX_...` group plus the makes it has to give back (`GCV_4W_ALL_EX_BHARAT_BENZ_TATA_ASHOK_MAHINDRA,Ashok`).  A word the script does not know, a set no master group can express, `CC`
  (council call) and Excel errors are NEVER guessed: the part is listed in the sheet 'Not written' of the review workbook (with the Council Call entries that may apply) and in the warnings.
* STP rows: a segment the AOTP matrix lists (LCV, MHCV) takes its STP payout from there; every other segment takes the comprehensive payout, unless its cell says AOTP:n% (CV NOTES, note 9).
* Default make groups leave out Bharat Benz (CV NOTES: 'BHARAT BENZ is declined across the system', read from the note).
* Private Car: Comprehensive and SAOD with NCB are the Petrol / CNG / Diesel / Electric columns (LPG takes the CNG column); without NCB one payout for all fuels; 1+3 is the 'New' column for all fuels
  (add_tnc 'Applicable for 3+3'); STP is the Act-only column on TP.  Two-Wheeler: Bike / Scooter x Petrol / Electric, 1+1 (add_tnc 'Applicable for 2+2,3+3', read from the column title), SAOD, STP;
  the new-two-wheeler grid is 1+5 by OEM (Electric = all makes, class NA).  po = pi - 7 (never below 0).

RTO names: new_rto_list carries only the cluster name as the grid writes it (the user's rule of 30 Oct 2026); a part of a cluster ('BIHAR except PURNIA', 'HP1') carries its condition in plain English.
GRID_RTO_NAMES=group writes the portal's group names (ICICI_SEPT26_CV_<cluster>, ICICI-RB_PC_JUL26_<cluster>, ICICI-RB_JUL26_<cluster>) - the comparison with an earlier upload uses them.

Interactive
-----------
The script checks the workbooks while it reads them.  Whenever something about the GRID it needs is missing, unclear or odd - a sheet or header that is not where it should be, a note it relies on that
has gone, dates that look wrong - it stops, asks you a specific question (type the number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every answer is listed
at the end of the run, and you confirm before the file is written.  It never asks you for a rate.  With --no-prompt (or when there is no keyboard) it asks nothing and stops with the question as the
error message instead.  If something goes wrong that it did not foresee it says so in one line and saves the technical details to a *_error_*.txt file.

Command line
------------
    python icici_grid_to_config.py --cv-grid <CV.xlsx> --pc-grid <PC.xlsx> --tw-grid <TW.xlsx> --tw-new-grid <NEW.xlsx> [--from-date 01/11/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the dates, the earlier upload - is asked for; at least one workbook is needed; `--cv-from-date` etc. give one workbook its own start date; `--vertical` overrides insurer_vertical)
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
INSURANCE_COMPANY = "ICICI LOMBARD GENERAL INSURANCE CO LTD"
USER_ID = 2966
ROW_ID = 1                                       # the portal's export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 8                                # the user's Private Car upload keeps the grid's value to 8 decimals (0.19910490492 -> 19.91049049, po 12.91049049); rounded once, half up
CC_OPEN_END = 99999
AGE_OPEN_END = 99
SC_RANGE = (0, 999)
TARIFF_FULL = 100.01                             # no discount band: every row is 0-100.01
MAX_NOTE_CHARS = 500
MIN_NOTE_CHARS = 8                               # a WHOLE text that no rule read is a rule until a standing note says otherwise (SBI lesson F-39: a short note can decide which rows exist)
HEADER_SCAN_ROWS = 12

# insurer_vertical: ICICI-RB-Offline on EVERY row, Private Car and Two-Wheeler included (the user, 4 Nov 2026).  Their CV uploads of September and October carry it, their August and November Private Car /
# Two-Wheeler uploads carried Not Applicable; the user settled it.  The dict stays per product so a change is one line; --vertical overrides it for every row.
INSURER_VERTICAL = {"cv": "ICICI-RB-Offline", "pc": "ICICI-RB-Offline", "tw": "ICICI-RB-Offline"}

# the portal's RTO groups of the user's uploads (only for GRID_RTO_NAMES=group and for the comparison with an earlier upload; the month in a prefix is the month the group was created)
GROUP_PREFIX = {"cv": "ICICI_SEPT26_CV_", "pc": "ICICI-RB_PC_JUL26_", "tw": "ICICI-RB_JUL26_", "tw15": "ICICI-RB_JUL26_"}
UP_TWO = "UP-WEST,{p}UP-EAST"                    # the user maps the single 'Uttar Pradesh' row of the Private Car grid and of the new two-wheeler grid to these two groups ...
UP_ALL = "UP-WEST,{p}UP-EAST,{p}LUCKNOW,{p}KANPUR,{p}VARANASI,{p}ALLAHABAD"      # ... and the one of the two-wheeler 1+1 / SAOD / STP grid to these six
_TW_NAMES = {"andamannicobar": "ANDAMAN&NICOBAR", "bhubaneshwarcuttack": "BHUBANESHWAR_&_CUTTAK", "chhattisgarh": "CHATTISGARH", "nashik": "Nashik", "vapi": "Vapi", "vishakapattnam": "VISAKHAPATNAM"}
GROUP_NAME_EXCEPTIONS = {                        # label_key of the grid's cluster -> the portal's group name after the prefix (everything else: capitals, spaces -> '_')
    "cv": {"andhrapradesh": "ANDHRA_PRADESH", "arunachalpradesh": "ARUNACHAL_PRADESH", "bhubaneshwar": "BHUBANESWAR", "chhattisgarh": "CHATTISGARH", "himachalpradesh": "HIMACHAL_PRADESH",
           "jammuandkashmir": "JAMMU_AND_KASHMIR", "madhyapradesh": "MADHYA_PRADESH", "nashik": "Nashik", "tamilnadu": "TAMIL_NADU", "vapi": "Vapi", "vishakapattnam": "VISAKHAPATNAM",
           "westbengal": "WEST_BENGAL"},
    "pc": {"chhattisgarh": "CHATTISGARH", "nasik": "Nashik", "vishakapattnam": "VISAKHAPATNAM", "vapi": "Vapi", "uttarpradesh": UP_TWO},
    "tw": dict(_TW_NAMES, uttarpradesh=UP_ALL),
    "tw15": dict(_TW_NAMES, uttarpradesh=UP_TWO),
}

FUELS_ALL = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NON_EV = ["Petrol", "Diesel", "CNG", "LPG"]
FUELS_NON_DIESEL = ["Petrol", "CNG", "LPG"]
FUELS_TW = ["Petrol", "Electric"]

# fixed T&C texts of the user's uploads (add_tnc is part of the key; the portal matches on them).  They are the user's wording of CV NOTES 4, 13 and 14; the run warns when the note is gone.
TNC_HAZARD = "Hazardous segment - such as milk tanker, water tanker, chemical tanker is declined"
TNC_BUS = "Contractor owned buses not to source. Also <=12 Seater vehicle will be consider in School Cab /Van NILL PO"
TNC_STAFF = "Staff Bus"
TNC_TAXI = "OLA & UBER/ TOURIST TAXI, Self driven Rental vehicles is declined in Taxi"
TNC_MISCD = "EX CRANES"                          # the segment 'MIsc D CE (Excluding CRANES)'
TAXI_SEATS = 7                                   # CV NOTES 13: 'PCV TAXI refers to upto 7 seater'
TNC_MULTI_YEAR = "Applicable for {}"             # 'Bike Comp (1+1)/(2+2)/(3+3)' -> sub_product 1+1, add_tnc 'Applicable for 2+2,3+3'

# the make master's single-make names for the makes a CV cell can name (the master's spelling; check_names flags a name it lacks)
MAKE_NAME = {"TATA": "TATA", "MAHINDRA": "Mahindra", "MARUTI": "Maruti", "ASHOK": "Ashok", "EICHER": "EICHER", "BAJAJ": "BAJAJ", "PIAGGIO": "PIAGGIO", "TVS": "TVS", "CITROEN": "CITROEN",
             "HYUNDAI": "Hyundai", "TATAACE": "DIGIT_CV_TATAACE"}
TW_MAKES = {"hmc": ("HERO", "Hero MotoCorp"), "hmsi": ("HONDA", "Honda"), "tvs": ("TVS", "TVS"), "suzuki": ("SUZUKI", "Suzuki"), "yamaha": ("YAMAHA", "Yamaha"), "re": ("ROYAL", "Royal Enfield")}
TW_EV_SHEET = "ev"                               # the EV sheet: Electric, all makes (TW_ALL), class NA
TW_ALL = "TW_ALL"

# Standing notes: texts that need no action (the rule is inside the grid's cells, the portal's groups handle it, or it is information for the sales team).  Matching ignores capitals and punctuation;
# a changed NUMBER or word makes the note new and it is flagged.
STANDING_NOTES = [
    # CV NOTES
    "Wherever the commission is mentioned as 0%/Decline, there would be no amount released",
    "Short period tractor policy will be at 0% commission",
    "Any segment not mentioned in the grid will be at 0% commission",
    "Any segment not mentioned in the grid will be evaluated separately for specific locations",
    "Strategic Fleet calls to be taken on case to case basis",
    "For Payout applicable for AOTP policies, please refer to sheet",
    "For Broking channel, if segment level call is in RTO mapped to EMG location",
    "Current PO is analyzed and approved basis running(Current) discount, any change in discount will lead to change in PO",
    "School Bus and Staff Bus refers to carrying capacity not seating capacity",       # note 4: the whole note (staff bus under 18, contractor owned buses, 12 seaters) is the T&C text of the bus rows
    "Hazardous segment - such as milk tanker, water tanker, chemical tanker is declined",      # note 14: the T&C text of the Tanker rows
    "Notes.", "Description", "Applied filters",       # the footer of a Power BI export in the November CV notes sheet
    # TW NOTES
    "New TW grid represents the maximum permissible levels",
    "Grid is evaluated basis current system discount",
    "Min expceted CPA attachment",
    "TW Grid for 1+1,2+2,3+3, SAOD, AOTP is applicable for all makes and all fuel types",
    "Team needs to exercise caution for cohorts wherein doability is linked to only certain district /RTOs",
    "For EV New (1+5) grid, there is no differentiation in bikes and scooters",
    "Bike OD discount to be zero, for deviation, please check with RUW",
    "Scooter discount upto 30%. For deviation, please check with RUW",
    "Min expceted CPA attachment is 80%",
    "IMID level deviation call will be considered with business commitment",
    "For multi year TW Comprehensive, same grid will be applicable as single year TW Comprehensive",
    "For multi year TW SAOD, same grid will be applicable as single year TW SAOD",
    "For 1+5 & 5+5 TW, New TW grid will be applicable",
]

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

# 'OTHERS:0%' means everything the cell's other clauses leave out is declined.  The user's October upload has NO zero card for it (a decline by absence; a Council Call row the user adds by hand
# for that scope would collide with one), while a plain 'No Biz' / '0%' cell and a named 'NEW:0%' do have one.  False = leave the OTHERS:0% cards out (they are listed in the sheet 'Not written');
# True = write them.
WRITE_OTHERS_ZERO = False

# every row takes the from_date typed for the run (the user's uploads carry one date)
KEEP_REF_DATES = False

OUTPUT_FOLDER = "icici"              # the folder of this insurer under output\ (default_output)


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


# %% ---------------------------------------------------------------- READING THE WORKBOOKS
def _text(v):
    return "" if v is None else str(v).replace("\xa0", " ").replace("​", "").strip()


def norm(v):
    """A cell as one line of text: no line breaks, no double spaces, no non-breaking spaces."""
    return re.sub(r"\s+", " ", _text(v)).strip()


def label_key(s):
    """A label without spaces, punctuation and capitals: 'NCB Status' and 'ncb  status' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def hkey(s):
    """A header without spaces and capitals but WITH its signs ('MHCV >40T Tipper/ Dumper' -> 'mhcv>40ttipper/dumper'): the signs say which band a column is."""
    return re.sub(r"\s+", "", _text(s).lower())


def num(v):
    """A cell as a float, or None when it is not a plain number (blank, 'No Biz', text)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%?\s*", _text(v))
    return float(m.group(1)) if m else None


def round_half_up(x, places=RATE_DECIMALS):
    q = 10 ** places
    return math.floor(float(x) * q + 0.5 + 1e-9) / q


def read_workbook(path):
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def hidden_sheets(wb):
    return {ws.title for ws in wb.worksheets if ws.sheet_state != "visible"}


def col_letter(c):
    from openpyxl.utils import get_column_letter
    return get_column_letter(c)


def twin_groups(columns):
    """`columns` = [(0-based column, the key that says what the column holds)].  Columns that hold the same thing (the SAME heading twice: a new column beside the old one) -> {column: the columns
    of its group} for every column that has a twin.  Which twin is this month's is not in the grid; the user's answer (12 Nov 2026) is the LOWEST payin of them: the rows of every twin are written
    and settle_duplicates keeps the lowest of each entry and lists it in 'Lowest rate chosen'."""
    by = {}
    for j, key in columns:
        by.setdefault(key, []).append(j)
    return {j: tuple(js) for js in by.values() if len(js) > 1 for j in js}


def blank_cell(v):
    return v is None or (isinstance(v, str) and not v.strip())


def twin_blank(vals, twins, c):
    """True when the cell of column c is blank and another column of its twin group has something in this row: a blank is no payin, so it must not become a zero card that wins as the lowest"""
    group = twins.get(c)
    return bool(group) and blank_cell(vals.get(c)) and any(not blank_cell(vals.get(o)) for o in group if o != c)


def sheet_values(wb, name):
    """the rows of a sheet as lists of values (a chart sheet or an unreadable sheet has none)"""
    try:
        return [list(r) for r in wb[name].iter_rows(values_only=True)]
    except Exception:
        return []


def clean_cluster(v):
    """a cluster name as the grid writes it, without the footnote stars ('Himachal Pradesh**')"""
    return norm(re.sub(r"\*+", "", _text(v)))


def group_name(family, cluster):
    """The portal's RTO group of a cluster in the user's uploads (family cv / pc / tw): prefix + the name in capitals with '_' for spaces, a few spelled as the user's groups are
    (GROUP_NAME_EXCEPTIONS).  Used for GRID_RTO_NAMES=group and for the comparison with an earlier upload; the CSV itself carries the cluster name."""
    prefix = GROUP_PREFIX[family]
    c = norm(re.sub(r"\(.*?\)", "", clean_cluster(cluster)))          # 'VAPI( Incl Diu, Daman, silvassa & Valsad)' is the group Vapi
    special = GROUP_NAME_EXCEPTIONS[family].get(label_key(c))
    if special:
        return prefix + special.replace("{p}", prefix)
    return prefix + re.sub(r"\s+", "_", c.upper())


# ---------------------------------------------------------------------------------------------------------------------------------- the CV segments
def _kg(t):
    return round(float(t) * 1000 + 0.01, 2)


def _seg(key, product, cls, family, cc=(0, CC_OPEN_END), sc=SC_RANGE, fuels=None, age=None, tag="", text="", split=True, tag_always=False):
    """tag_always: the segment's New / Old tag stays on every row of the segment, also on a row with an age band (the user's SCV rows: 'Old' with 1-3.01); for the other segments the tag is
    written only when the row's age is exactly the new (0-1) or the old (1-99) range"""
    return dict(key=key, product=product, cls=cls, family=family, cc=cc, sc=sc, fuels=list(fuels or FUELS_ALL), age=age, tag=tag, text=text, split=split, tag_always=tag_always)


GCV_CLASS = {"tanker": "Tanker", "tipper": "Dumper Tipper", "tipper/dumper": "Dumper Tipper", "dumper": "Dumper Tipper", "truck": "Truck", "trailer": "Trailer"}


def classify_segment(header):
    """The CV segment a matrix column stands for (None for a column this script does not know).  Found by the WORDS and signs of the header, tonnages and seat bands are read from it, so a
    renamed band ('MHCV 12-18T') or a new spelling ('PCV 3W Petrol/CNG New' for 'Non-Diesel New') still lands in the right segment.  `split` = the cell may be split NEW / OLD."""
    h = hkey(header)
    if not h:
        return None
    m = re.match(r"^gcv3w(new|old)$", h)
    if m:
        new = m.group(1) == "new"
        return _seg(f"gcv3w_{m.group(1)}", "GCV 3W", "Goods Carrying Rickshaw", "gcv3w", fuels=FUELS_NON_EV, age=(0, 1) if new else (1, AGE_OPEN_END), tag="New" if new else "Old", split=False)
    if re.match(r"^gcv3w(electric|ev)$", h):
        return _seg("gcv3w_ev", "GCV 3W", "Goods Carrying Rickshaw", "gcv3w", fuels=["Electric"])
    m = re.match(r"^scv(<|>=|>)(\d+)gvw(new|old)$", h)
    if m:
        n, new = float(m.group(2)), m.group(3) == "new"
        cc = (0, n) if m.group(1) == "<" else (n, _kg(3.5))
        return _seg(f"scv{m.group(1)}{m.group(2)}_{m.group(3)}", "GCV 4W", "NA", "gcv4w", cc=cc, age=(0, 1) if new else (1, AGE_OPEN_END), tag="New" if new else "Old", split=False, tag_always=True)
    m = re.match(r"^lcv(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)t$", h)
    if m:
        return _seg(f"lcv{m.group(1)}-{m.group(2)}", "GCV 4W", "NA", "gcv4w", cc=(_kg(m.group(1)), _kg(m.group(2))))
    m = re.match(r"^mhcv(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)t(tanker|tipper/?dumper|tipper|dumper|truck|trailer)$", h)
    if m:
        cls = GCV_CLASS[m.group(3)]
        return _seg(f"mhcv{m.group(1)}-{m.group(2)}_{cls}", "GCV 4W", cls, "gcv4w", cc=(_kg(m.group(1)), _kg(m.group(2))), text=TNC_HAZARD if cls == "Tanker" else "")
    m = re.match(r"^mhcv>(\d+(?:\.\d+)?)t(tanker|tipper/?dumper|tipper|dumper|truck|trailer)$", h)
    if m:
        cls = GCV_CLASS[m.group(2)]
        return _seg(f"mhcv>{m.group(1)}_{cls}", "GCV 4W", cls, "gcv4w", cc=(_kg(m.group(1)), CC_OPEN_END), text=TNC_HAZARD if cls == "Tanker" else "")
    if re.match(r"^misc-?dce", h):
        return _seg("miscd_ce", "MISCD", "MISCD-Others", "miscd", text=TNC_MISCD)
    m = re.match(r"^tractor(new|old)$", h)
    if m:
        new = m.group(1) == "new"
        return _seg(f"tractor_{m.group(1)}", "MISCD", "MISCD-Tractor", "miscd", age=(0, 1) if new else (1, AGE_OPEN_END), tag="New" if new else "Old", split=False)
    m = re.match(r"^pcv3w(?:non-diesel|petrol/cng|nondiesel)(new|old)$", h)
    if m:
        new = m.group(1) == "new"
        return _seg(f"pcv3w_nd_{m.group(1)}", "PCV 3W", "Passenger Carrying Rickshaw", "pcv3w", fuels=FUELS_NON_DIESEL, age=(0, 1) if new else (1, AGE_OPEN_END), tag="New" if new else "Old", split=False)
    if re.match(r"^pcv3wothers\(diesel\)$|^pcv3wdiesel$", h):
        return _seg("pcv3w_diesel", "PCV 3W", "Passenger Carrying Rickshaw", "pcv3w", fuels=["Diesel"])
    if re.match(r"^pcv3w(electric|ev)$", h):
        return _seg("pcv3w_ev", "PCV 3W", "Passenger Carrying Rickshaw", "pcv3w", fuels=["Electric"])
    m = re.match(r"^(school|staff)bus(<18|18-36|>36|>18)$", h)
    if m:
        school = m.group(1) == "school"
        sc = {"<18": (0, 18.01), "18-36": (18.01, 36.01), ">36": (36.01, SC_RANGE[1]), ">18": (18.01, SC_RANGE[1])}[m.group(2)]
        return _seg(f"{m.group(1)}bus{m.group(2)}", "PCV 4W", "School Bus" if school else "Other Bus", "bus", sc=sc, text=TNC_BUS if school else f"{TNC_STAFF},{TNC_BUS}")
    if re.match(r"^pcvtaxi_?(electric|ev)$", h):
        return _seg("taxi_ev", "PCV 4W", "Taxi", "taxi", sc=(0, TAXI_SEATS), fuels=["Electric"], text=TNC_TAXI)
    m = re.match(r"^pcvtaxi(<=|>)(\d+)cc$", h)
    if m:
        cc = (0, float(m.group(2)) + 0.01) if m.group(1) == "<=" else (float(m.group(2)) + 0.01, CC_OPEN_END)
        return _seg(f"taxi{m.group(1)}{m.group(2)}", "PCV 4W", "Taxi", "taxi", cc=cc, sc=(0, TAXI_SEATS), fuels=FUELS_NON_EV, text=TNC_TAXI)
    if re.match(r"^pcv\(?2w\)?$", h):
        return _seg("pcv2w", "PCV 2W", "PCV 2W", "pcv2w")
    return None


# ---------------------------------------------------------------------------------------------------------------------------------- the cell grammar
AGE_RX = re.compile(r"(?P<op>>=|<=|>|<)?\s*(?P<a>\d+(?:\.\d+)?)\s*(?:(?P<plus>\+)|\s*(?:-|to)\s*(?P<b>\d+(?:\.\d+)?))?\s*(?:yrs?|years?)\b", re.I)
MAKE_ALIASES = {      # label of a make word in a cell (letters, digits and & only) -> the token it stands for
    "tata": "TATA", "tatamotors": "TATA", "m&m": "MAHINDRA", "mm": "MAHINDRA", "mahindra": "MAHINDRA", "maruti": "MARUTI", "marutisuzuki": "MARUTI",
    "al": "ASHOK", "ashok": "ASHOK", "ashokleyland": "ASHOK", "eicher": "EICHER", "bajaj": "BAJAJ", "piaggio": "PIAGGIO", "piaggao": "PIAGGIO", "tvs": "TVS",
    "citroen": "CITROEN", "hyundai": "HYUNDAI", "omni": "OMNI", "magic": "MAGIC", "eeco": "EECO", "ecco": "EECO", "force": "FORCE", "sumo": "SUMO", "tataace": "TATAACE",
}
BODY_WORDS = {        # vehicle-type words that no make group can carry: they travel in add_tnc ('CRANES,ACE,ESCORTS' / 'EX:CRANES,ACE,ESCORTS')
    "ace": "ACE", "escorts": "ESCORTS", "cranes": "CRANES", "crane": "CRANES", "jcb": "JCB", "fueltanker": "Fuel Tanker", "bulker": "Bulker", "backhoeloader": "Backhoe Loader",
    "drillingrig": "Drilling Rig", "mobileplant": "Mobile Plant", "transitmixer": "Transit Mixer", "concretemixer": "Concrete Mixer",
}
REGION_WORDS = {"kashmir": "KASHMIR", "jammu": "JAMMU", "delhi": "DELHI", "ncr": "NCR"}      # parts of a cluster that no RTO list names as a district
POLICY_WORDS = {"comp": "COMP", "aotp": "AOTP"}
ZD_WORDS = {"zd": "YES", "nzd": "NO"}
SUBCLUSTER_RX = re.compile(r"[a-z]{2,3}\d(?:-[a-z]+)?")      # HP1, HR2, GJ1, WB2, JK1, UP1-WEST


def wlabel(s):
    """a word of a cell: letters, digits and & only, lower case"""
    return re.sub(r"[^a-z0-9&]+", "", _text(s).lower())


def age_interval(op, a, plus, b):
    """The portal band of a grid age label (the convention of the Shriram scripts, found again in the user's ICICI upload): 'a-b yrs' = the a-th to the b-th year of life = (a-1)+0.01 .. b+0.01
    (year 1 starts at 1: a new vehicle, age 0-1, is the NEW case), 'a+ yrs' = a+0.01 .. 99, '>=n yrs' = (n-1)+0.01 .. 99, '<=n yrs' = 0 .. n+0.01."""
    a = float(a)
    lo_of = lambda n: 0.0 if n <= 0 else (1.0 if n <= 1 else round(n - 1 + 0.01, 2))
    if b is not None:
        return lo_of(a), round(float(b) + 0.01, 2)
    if plus:
        return round(a + 0.01, 2), float(AGE_OPEN_END)
    if op == ">=":
        return lo_of(a), float(AGE_OPEN_END)
    if op == "<=":
        return 0.0, round(a + 0.01, 2)
    if op == ">":
        return round(a + 0.01, 2), float(AGE_OPEN_END)
    if op == "<":
        return 0.0, a
    return lo_of(a), round(a + 0.01, 2)


def split_clauses(text):
    """a cell's text -> clause strings.  Commas separate clauses; a '%' followed by '|' is a typo for a comma ('JK1|<=10 yrs:28% |OTHERS:0%')."""
    t = re.sub(r"\s+", " ", _text(text))
    t = re.sub(r"%\s*\|", "%,", t)
    return [c.strip() for c in t.split(",") if c.strip()]


RATE_END = re.compile(r"^(?P<head>.*?)\s*(?:[:\-]\s*|\s+)?(?P<rate>\d+(?:\.\d+)?)\s*%\s*$")


def parse_clause(c):
    """'OLD|TATA|AL:12%' -> (['OLD', 'TATA', 'AL'], 12.0); 'CC' -> (['CC'], None); 'AOTP15%' -> (['AOTP'], 15.0); 'COMP|NEW|ZD-20%' -> (['COMP', 'NEW', 'ZD'], 20.0)"""
    m = RATE_END.match(c)
    if not m:
        return [t.strip() for t in re.split(r"[|:]", c) if t.strip()], None
    head = m.group("head").strip(" :|-")
    return [t.strip() for t in re.split(r"[|:]", head) if t.strip()], float(m.group("rate"))


def split_glued(tok):
    """'M&M>=4 yrs' -> ['M&M', '>=4 yrs'] (a make glued to an age)"""
    m = AGE_RX.search(tok)
    if not m or (m.start() == 0 and m.end() == len(tok.strip())):
        return [tok]
    out = [x for x in (tok[:m.start()].strip(), tok[m.start():m.end()].strip(), tok[m.end():].strip()) if x]
    return out


def classify_word(tok, places):
    """(kind, value) of one word of a cell; kind in policy / txn / age / zd / cc / others / make / body / loc / unknown.  `places` = the place vocabulary {wlabel: NAME} of the workbook."""
    k = wlabel(tok)
    if k in POLICY_WORDS:
        return "policy", POLICY_WORDS[k]
    if k in ("new", "old"):
        return "txn", k.upper()
    if k in ("newold", "oldnew"):
        return "txn", "ALL"
    if k in ZD_WORDS:
        return "zd", ZD_WORDS[k]
    if k == "cc":
        return "cc", None
    if k in ("others", "other"):
        return "others", None
    m = AGE_RX.fullmatch(tok.strip())
    if m:
        return "age", age_interval(m.group("op"), m.group("a"), m.group("plus"), m.group("b"))
    if k in MAKE_ALIASES:
        return "make", MAKE_ALIASES[k]
    if k in BODY_WORDS:
        return "body", BODY_WORDS[k]
    if k in places:
        return "loc", places[k]
    if k in REGION_WORDS:
        return "loc", REGION_WORDS[k]
    if SUBCLUSTER_RX.fullmatch(re.sub(r"[^a-z0-9-]+", "", tok.lower())):
        return "loc", re.sub(r"\s+", "", tok).upper()
    return "unknown", tok.strip()


# ---- sets over a discrete dimension (makes, places, body words, zero dep, policy): ALL, ("in", {...}) or ("not", {...})
ALL = ("all", frozenset())


def s_in(items):
    return ("in", frozenset(items))


def s_and(a, b):
    if a == ALL:
        return b
    if b == ALL:
        return a
    (ka, sa), (kb, sb) = a, b
    if ka == "in" and kb == "in":
        return ("in", sa & sb)
    if ka == "in":
        return ("in", sa - sb)
    if kb == "in":
        return ("in", sb - sa)
    return ("not", sa | sb)


def s_neg(a):
    if a == ALL:
        return ("in", frozenset())
    k, s = a
    return ("not", s) if k == "in" else ("in", s)


def s_empty(a):
    return a[0] == "in" and not a[1]


# ---- age: a list of disjoint [lo, hi) intervals
def a_and(a, b):
    out = []
    for lo1, hi1 in a:
        for lo2, hi2 in b:
            lo, hi = max(lo1, lo2), min(hi1, hi2)
            if hi > lo + 1e-9:
                out.append((lo, hi))
    return sorted(out)


def a_union(items):
    """the union of age bands as disjoint, sorted intervals"""
    out = []
    for lo, hi in sorted(items):
        if out and lo <= out[-1][1] + 1e-9:
            out[-1] = (out[-1][0], max(out[-1][1], hi))
        else:
            out.append((lo, hi))
    return out


def a_neg(a, universe=((0.0, float(AGE_OPEN_END)),)):
    cur = list(universe)
    for lo2, hi2 in sorted(a):
        nxt = []
        for lo, hi in cur:
            if hi2 <= lo or lo2 >= hi:
                nxt.append((lo, hi))
                continue
            if lo2 > lo:
                nxt.append((lo, lo2))
            if hi2 < hi:
                nxt.append((hi2, hi))
        cur = nxt
    return [(lo, hi) for lo, hi in cur if hi > lo + 1e-9]


# ---- boxes: the conditions of a clause in (policy, zero dep, place, body, make, age).  Subtracting goes through the dimensions in this order, age LAST, so that a make set keeps one row per age band
# (the user's own partition: 'everything but CRANES / ACE / ESCORTS' for all ages + 'CRANES / ACE / ESCORTS' above 10 years).
DIMS = ("pol", "zd", "locs", "body", "makes", "age")


def full_box():
    return {"pol": ALL, "zd": ALL, "locs": ALL, "body": ALL, "makes": ALL, "age": [(0.0, float(AGE_OPEN_END))]}


def b_and(a, b):
    return {d: (a_and(a[d], b[d]) if d == "age" else s_and(a[d], b[d])) for d in DIMS}


def b_empty(b):
    return not b["age"] or any(s_empty(b[d]) for d in DIMS if d != "age")


def b_sub(a, b):
    """a minus b, as disjoint boxes (a itself, whole, when they do not overlap)"""
    if b_empty(b_and(a, b)):
        return [a]
    pieces, prefix = [], dict(a)
    for d in DIMS:
        if d == "age":
            rest, inter = a_and(prefix["age"], a_neg(b["age"])), a_and(prefix["age"], b["age"])
            keep = bool(rest)
        else:
            rest, inter = s_and(prefix[d], s_neg(b[d])), s_and(prefix[d], b[d])
            keep = not s_empty(rest)
        if keep:
            piece = dict(prefix)
            piece[d] = rest
            if not b_empty(piece):
                pieces.append(piece)
        prefix = dict(prefix)
        prefix[d] = inter
        if b_empty(prefix):
            break
    return pieces


def parse_cell(text, places):
    """-> dict(clauses = [dict(box, rate, others, raw)], unknown = [words nobody knows], cc = True when the cell names the council call, order = {word: position})
    A clause's box takes every dimension its words name; OTHERS is resolved by `regions`."""
    clauses, unknown, cc, order = [], [], False, {}
    for raw in split_clauses(text):
        toks, rate = parse_clause(raw)
        toks = [x for t in toks for x in split_glued(t)]
        box, others, is_cc = full_box(), False, False
        found = {"policy": [], "zd": [], "make": [], "body": [], "loc": [], "txn": [], "age": []}
        for t in toks:
            kind, val = classify_word(t, places)
            if kind in found:
                found[kind].append(val)
            elif kind == "cc":
                is_cc = True
            elif kind == "others":
                others = True
            else:
                unknown.append(val)
            if kind in ("make", "body", "loc"):
                order.setdefault(val, len(order))
        # words of one kind inside a clause are alternatives (TATA|M&M, 1-3 years|6-10 years); the kinds are conditions that all hold (OLD|TATA|1-3 yrs)
        txns = [v for v in found["txn"] if v != "ALL"]
        if txns and len(set(txns)) == 1:
            box["age"] = a_and(box["age"], [(0.0, 1.0)] if txns[0] == "NEW" else [(1.0, float(AGE_OPEN_END))])
        if found["age"]:
            box["age"] = a_and(box["age"], a_union(found["age"]))
        for dim, kind in (("pol", "policy"), ("zd", "zd"), ("makes", "make"), ("body", "body"), ("locs", "loc")):
            if found[kind]:
                box[dim] = s_in(found[kind])
        cc = cc or is_cc
        if rate is None:
            if not is_cc:
                unknown.append(raw)
            continue
        clauses.append(dict(box=box, rate=rate, others=others, raw=raw))
    return dict(clauses=clauses, unknown=unknown, cc=cc, order=order)


def regions(parsed, base, pol, skipped=None):
    """The regions (box, rate, clause text) one policy ('COMP' / 'AOTP') gets from a parsed cell inside the segment's base box.  A clause tagged with the other policy does not apply; for AOTP,
    when no clause is tagged AOTP the comprehensive clauses are used (the STP payout of a segment the AOTP matrix does not list is the comprehensive payout - CV NOTES, note 9).
    OTHERS = the clause's own conditions inside the base box minus every named clause of the cell.  `skipped` collects the OTHERS:0% regions that are not written (WRITE_OTHERS_ZERO)."""
    skipped = skipped if skipped is not None else []
    tag = lambda c: c["box"]["pol"]
    if pol == "AOTP" and not any(tag(c) == s_in(["AOTP"]) for c in parsed["clauses"]):
        use = list(parsed["clauses"])
    else:
        use = [c for c in parsed["clauses"] if tag(c) in (ALL, s_in([pol]))]
    use = [dict(c, box=dict(c["box"], pol=ALL)) for c in use]
    named = [c for c in use if not c["others"]]
    declines = [c for c in named if c["rate"] == 0]
    out = []
    for c in use:
        region = [b_and(c["box"], base)]
        if c["others"]:
            for n in named:
                region = [p for r in region for p in b_sub(r, n["box"])]
            if c["rate"] == 0 and not WRITE_OTHERS_ZERO:
                skipped.extend((r, c["raw"]) for r in region if not b_empty(r))
                continue
        elif c["rate"] != 0:
            # a 0% clause is a decline: what it declines is taken out of the other clauses too ('Omni|Magic|Eeco:0%' -> the other rows are 'all taxi makes except Omni, Magic, Eeco')
            cut = region
            for d in declines:
                cut = [p for r in cut for p in b_sub(r, d["box"])]
            region = cut or region                       # a rate inside a wider 0% clause stays: the clause that names it is the more specific one
        out.extend((r, c["rate"], c["raw"]) for r in region if not b_empty(r))
    # a region that lies wholly inside another region of the same rate is a duplicate ('OLD|TATA|>=6 yrs:20%' inside 'OLD|1-3yrs|>=6 yrs:20%' for all makes)
    keep = []
    for i, (r, rate, raw) in enumerate(out):
        if any(j != i and rate == out[j][1] and not b_sub(r, out[j][0]) and (b_sub(out[j][0], r) or j < i) for j in range(len(out))):
            continue
        keep.append((r, rate, raw))
    return keep


EXCEL_ERROR = re.compile(r"#(?:REF!|N/A|VALUE!|NAME\?|DIV/0!|NULL!|NUM!)", re.I)


def cell_value(v, unit="percent"):
    """(kind, payload) of a CV matrix cell: ('num', percent) / ('zero', 0) for No Biz or a blank / ('text', str) / ('error', str) for an Excel error.
    `unit` is the unit of the matrix's plain numbers ('fraction': 0.6 means 60%; a text with a % sign always means percent)."""
    if v is None or (isinstance(v, str) and not v.replace(" ", " ").strip()):
        return "zero", 0.0
    if isinstance(v, bool):
        return "error", str(v)
    if isinstance(v, (int, float)):
        return "num", float(v) * (100 if unit == "fraction" else 1)
    s = _text(v)
    if EXCEL_ERROR.fullmatch(s):
        return "error", s
    if label_key(s) in ("nobiz", "nobusiness", "decline", "declined", "na", "nil", "d"):
        return "zero", 0.0
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(%?)", s)
    if m:
        return "num", float(m.group(1)) * (100 if unit == "fraction" and not m.group(2) else 1)
    return "text", s


# ---------------------------------------------------------------------------------------------------------------------------------- the CV workbook
class CvMatrix:
    """One matrix sheet of the CV grid: the header row, the cluster column, the segment columns (found by `classify_segment`), the columns it did not know and the cluster rows.
    `unit` = 'fraction' when every plain number of the matrix is at most 1 (0.6 = 60%), else 'percent'; a text such as '60.0%' says its unit itself."""

    twins = {}                                           # column -> the columns that carry the SAME heading (twin_groups), set by find_cv_matrices

    def __init__(self, name, rows, hdr, cluster_col, cols, unknown, kind, ignored=()):
        self.name, self.hdr, self.cluster_col, self.cols, self.unknown_cols, self.kind = name, hdr, cluster_col, cols, unknown, kind   # cols = [(column index, header text, segment)]
        self.ignored = list(ignored)                         # labelled columns that are no segment and are not used (the person's own RTO column)
        self.rows, self.bad = [], []                         # rows = (excel row, cluster, {column index: raw value}); bad = Excel rows whose cluster cell is an Excel error
        for i in range(hdr + 1, len(rows)):
            r = rows[i]
            raw = r[cluster_col] if cluster_col < len(r) else None
            cl = "" if isinstance(raw, str) and EXCEL_ERROR.fullmatch(raw.strip()) else clean_cluster(raw)
            has_values = any((c < len(r) and r[c] is not None) for c, _, _ in cols)
            if not cl:
                if isinstance(raw, str) and EXCEL_ERROR.fullmatch(raw.strip()) and has_values:
                    self.bad.append(i + 1)
                continue                                     # a blank row inside the matrix does not end it: the text row under the table does
            if not has_values:
                if self.rows:
                    break                                    # the first text row under the table
                continue
            self.rows.append((i + 1, cl, {c: r[c] for c, _, _ in cols if c < len(r)}))
        nums = [float(v) for _, _, vals in self.rows for v in vals.values() if isinstance(v, (int, float)) and not isinstance(v, bool)]
        self.unit = "fraction" if nums and max(nums) <= 1.0 + 1e-9 else "percent"


def pick_cluster_column(rows, hdr, first, name_hint=""):
    """The column that holds the cluster names: of the labelled header cells left of the first segment, the one whose header says cluster / segment and whose values are text - not an 'RTO'
    column (the person's own RTO group names) and not a column of numbers.  -> (column, [labelled columns that are not used])"""
    cands = [j for j, h in enumerate(rows[hdr][:first]) if isinstance(h, str) and h.strip()]
    if not cands:
        return 0, []

    def score(j):
        h = hkey(rows[hdr][j])
        vals = [r[j] for r in rows[hdr + 1:hdr + 40] if j < len(r) and r[j] is not None]
        text = sum(1 for v in vals if isinstance(v, str) and not re.fullmatch(r"[\d.\s]+", v)) / max(len(vals), 1)
        sc = (3 if "cluster" in h else 0) + (2 if "segment" in h else 0) + (-6 if working_header(rows[hdr][j]) else 0) + (-5 if text < 0.5 else 0)
        return (sc, -j)
    best = max(cands, key=score)
    return best, [(j, norm(rows[hdr][j])) for j in cands if j != best]


def find_cv_matrices(wb, skip=()):
    """Every visible sheet that holds a CV segment matrix, best first: a header row with at least six cells that are known segments.  -> [CvMatrix]"""
    found = []
    for name in wb.sheetnames:
        if name in skip or wb[name].sheet_state != "visible":
            continue
        rows = sheet_values(wb, name)
        best = None
        for i, r in enumerate(rows[:HEADER_SCAN_ROWS]):
            hits = [(j, norm(h), classify_segment(h)) for j, h in enumerate(r) if isinstance(h, str) and h.strip()]
            hits = [x for x in hits if x[2]]
            if len(hits) >= 6 and (best is None or len(hits) > len(best[1])):
                best = (i, hits)
        if not best:
            continue
        i, hits = best
        twins = twin_groups([(j, " ".join(h.lower().split())) for j, h, _ in hits])
        first = hits[0][0]
        cluster_col, ignored = pick_cluster_column(rows, i, first)
        taken = {j for j, _, _ in hits} | {j for j, _ in ignored}
        unknown = [(j, norm(h)) for j, h in enumerate(rows[i]) if j != cluster_col and j not in taken and isinstance(h, str) and h.strip()]
        low = label_key(name)
        kind = "aotp" if ("aotp" in low or re.search(r"(^|[^a-z])tp$", name.strip().lower()) or low == "tp") and "comp" not in low else "comp"
        matrix = CvMatrix(name, rows, i, cluster_col, hits, unknown, kind, ignored)
        matrix.twins = twins
        found.append(matrix)
    found.sort(key=lambda m: (-len(m.cols), m.name))
    return found


def pick_cv_matrices(wb, ask, warnings, log=say):
    """(comprehensive matrix, third-party matrix or None).  The comprehensive one is the sheet that lists the most segments; the third-party one is the other sheet (its name says AOTP / TP)."""
    found = find_cv_matrices(wb)
    if not found:
        return None, None
    comp = next((m for m in found if m.kind == "comp"), found[0])
    others = [m for m in found if m is not comp]
    aotp = next((m for m in others if m.kind == "aotp"), others[0] if others else None)
    if len(others) > 1:
        warnings.append(f"more than one sheet holds a third-party (AOTP) matrix ({', '.join(repr(m.name) for m in others)}): '{aotp.name}' is used")
    for m in (comp, aotp):
        if m is not None:
            log(f"  CV matrix '{m.name}' ({m.kind.upper()}): {len(m.rows)} clusters x {len(m.cols)} segments, header on row {m.hdr + 1}")
    return comp, aotp


class CcEntry(dict):
    """one Council Call entry (a row of the CC sheet)"""


def read_council_calls(wb, skip=()):
    """The Council Call list (CC / COUNCIL CALL / CC Call): state, tonnage, age, transaction, make, location, policy, COA %, status, ID.  Only for the review sheet - an entry is never turned into a rate."""
    found = find_tables(wb, {"state": ["state"], "tonnage": ["tonnage"], "age": ["age"], "txn": ["transaction"], "make": ["manufacturer"], "model": ["model"], "loc": ["rto location"],
                             "pol": ["comp tp", "comp/ tp"], "coa": ["coa", "coa %"], "status": ["council call status", "status"], "fuel": ["fuel type"], "id": ["council call id"],
                             "remark": ["remark", "remarks"]}, ("state", "tonnage", "coa"), exclude=set(skip) | hidden_sheets(wb))
    if not found:
        return None, []
    name, hdr, cols, body = found[0]
    out = []
    for k, r in enumerate(body):
        g = lambda f: norm(r[cols[f]]) if f in cols and cols[f] < len(r) and r[cols[f]] is not None else ""
        if not g("state") and not g("tonnage"):
            continue
        coa = r[cols["coa"]] if cols["coa"] < len(r) else None
        out.append(CcEntry(row=hdr + 2 + k, id=g("id"), state=g("state"), tonnage=g("tonnage"), age=g("age"), txn=g("txn"), make=g("make"), model=g("model"), loc=g("loc"), pol=g("pol"),
                           coa=(round_half_up(float(coa) * 100 if isinstance(coa, (int, float)) and coa <= 1 else float(coa)) if isinstance(coa, (int, float)) else None),
                           status=g("status"), fuel=g("fuel"), remark=g("remark")))
    return name, out


def read_places(wb):
    """The place vocabulary of the CV workbook {wlabel: NAME}: the districts and sub-cluster codes of its own reference sheets (the RTO sub-cluster list, the Negative RTO list, the RTO code list and the
    Council Call locations).  A word of a cell that is none of these and no make / age / policy word is NOT a place - the cell is listed, never guessed."""
    places, subs = {}, {}

    def add(x):
        x = norm(x)
        k = wlabel(x)
        if x and len(k) > 1 and k not in ("all", "na", "other", "others", "none"):
            places.setdefault(k, x.upper())
    sources = 0
    for name in wb.sheetnames:
        if wb[name].sheet_state != "visible":
            continue
        rows = sheet_values(wb, name)[:9000]
        if not rows:
            continue
        hit = locate_table(rows, {"loc": ["rto location", "rto_city", "city"], "district": ["district"], "sub": ["rto sub cluster"], "cluster": ["rto cluster", "rto_cluster"], "seg": ["segment"],
                                  "mm": ["make model hp"]}, (), scan=25)
        if hit and len(hit[1]) >= 2 and ({"loc", "district", "sub"} & set(hit[1])):
            sources += 1
            h, cols = hit
            for r in rows[h + 1:]:
                for f in ("loc", "district", "sub"):
                    if f in cols and cols[f] < len(r) and r[cols[f]]:
                        add(r[cols[f]])
                if {"sub", "loc", "cluster"} <= set(cols) and max(cols.values()) < len(r) and r[cols["sub"]] and r[cols["loc"]]:
                    subs.setdefault((label_key(r[cols["cluster"]]), norm(r[cols["sub"]]).upper()), set()).add(norm(r[cols["loc"]]).title())
        hit = locate_table(rows, {"loc": ["rto location"], "status": ["council call status"]}, ("loc", "status"), scan=25)
        if hit:
            h, cols = hit
            for r in rows[h + 1:]:
                if cols["loc"] < len(r) and r[cols["loc"]]:
                    for part in re.split(r"[,/;]|\s{2,}", str(r[cols["loc"]])):
                        add(part)
                    for part in re.split(r"[,\s/;]+", str(r[cols["loc"]])):
                        add(part)
    return places, subs, sources


def read_negative_list(wb):
    """The Negative RTO list: rows of (segment, state, cluster, district, RTO, status).  Only rows whose current status is still negative are kept."""
    found = find_tables(wb, {"segment": ["segment"], "state": ["rto state"], "cluster": ["rto cluster"], "district": ["district"], "loc": ["rto location"], "old": ["old status", "old_status"],
                             "cur": ["current status", "current_status"]}, ("segment", "state"), exclude=hidden_sheets(wb))
    if not found:
        return None, []
    name, hdr, cols, body = found[0]
    out = []
    for r in body:
        g = lambda f: norm(r[cols[f]]) if f in cols and cols[f] < len(r) and r[cols[f]] is not None else ""
        if not g("segment") or not g("state"):
            continue
        status = g("cur") or g("old")
        if status and re.search(r"okay\s*to\s*open|open", status, re.I) and not re.search(r"negative", status, re.I):
            continue
        out.append(dict(segment=g("segment"), state=g("state"), cluster=g("cluster"), district=g("district"), rto=g("loc"), status=status))
    return name, out


def read_notes(wb, skip=()):
    """The notes of the workbook: [(sheet, row, text)].  A sheet whose name says 'notes' gives every text of it (one note per cell)."""
    out = []
    for name in wb.sheetnames:
        if name in skip or wb[name].sheet_state != "visible" or "note" not in label_key(name):
            continue
        for i, r in enumerate(sheet_values(wb, name)):
            for c in r:
                if isinstance(c, str) and len(norm(c)) >= MIN_NOTE_CHARS and label_key(c) not in ("notes", "description", "note"):
                    out.append((name, i + 1, norm(c)))
    return out


# the clusters of each state (label_key): a Council Call entry or a Negative RTO list entry names a STATE, the matrices name clusters (a city or a whole state)
STATE_CLUSTERS = {
    "gujarat": ["gujarat", "ahmedabadgandhinagar", "baroda", "rajkot", "surat", "vapi"], "maharashtra": ["maharashtra", "mumbai", "pune", "nagpur", "nashik"],
    "karnataka": ["karnataka", "bangalore", "mysore", "mangalore"], "tamilnadu": ["tamilnadu", "chennai", "coimbatore"], "andhrapradesh": ["andhrapradesh", "vijaywada", "vishakapattnam"],
    "telangana": ["telangana", "hyderabad"], "odisha": ["odisha", "bhubaneshwar", "cuttack"], "westbengal": ["westbengal", "kolkata"], "rajasthan": ["rajasthan", "jaipur"],
    "madhyapradesh": ["madhyapradesh", "indore"], "kerala": ["kerala", "cochin"], "punjab": ["punjab", "ludhiana"], "bihar": ["bihar", "patna"], "jharkhand": ["jharkhand", "ranchi"],
    "assam": ["assam", "guwahati"], "uttarakhand": ["uttarakhand", "dehradun"], "uttarpradesh": ["upeast", "upwest", "allahabad", "kanpur", "lucknow", "varanasi", "uttarpradesh"],
    "delhi": ["ncr"], "ncr": ["ncr"], "chandigarh": ["chandigarh"], "haryana": ["haryana"],
}


def cluster_in_state(cluster, state):
    """True when a matrix cluster belongs to the state a Council Call / Negative RTO entry names (spelling, spaces and the 'AND' of 'JAMMU AND KASHMIR' do not matter)"""
    c, s = label_key(cluster), label_key(state)
    if not c or not s:
        return False
    if c == s or c in STATE_CLUSTERS.get(s, ()):
        return True
    return len(s) >= 6 and (c.startswith(s[:7]) or s.startswith(c[:7]))


# ---------------------------------------------------------------------------------------------------------------------------------- the Private Car grid
PC_KINDS = [      # (kind, regex on the group title, what it is)
    ("noncb", r"non-?ncb", "SAOD and comprehensive without NCB"),
    ("used", r"usedcar", "Used car"),
    ("new", r"new.*1\+3|new\(1\+3|^new", "New (1+3, 3+3)"),
    ("saod_ncb", r"saod", "SAOD with NCB"),
    ("comp_ncb", r"comp", "Comprehensive with NCB"),
    ("stp", r"actonly|netgwp|thirdparty|tponly|^act", "Act only (STP)"),
]
FUEL_WORDS = {"petrol": "Petrol", "cng": "CNG", "diesel": "Diesel", "electric": "Electric", "ev": "Electric", "lpg": "LPG"}


class PcTable:
    twins = {}                                           # column -> the columns of the same (group, fuel) under the SAME heading (twin_groups), set by find_pc_grid

    def __init__(self, name, hdr, cluster_col, cat_col, cols, extra, rows):
        self.name, self.hdr, self.cluster_col, self.cat_col, self.cols, self.extra, self.rows = name, hdr, cluster_col, cat_col, cols, extra, rows   # cols = [(column, kind, fuel, title)]


def find_pc_grid(wb, ask, warnings, log=say):
    """The Private Car sheet: a two-row header whose lower row names fuels (Petrol / CNG / Diesel / Electric) and whose upper row holds the group titles (New (1+3, 3+3), SAOD with NCB ...).
    Returns a PcTable or None."""
    for name in wb.sheetnames:
        if wb[name].sheet_state != "visible":
            continue
        rows = sheet_values(wb, name)
        for i, r in enumerate(rows[:HEADER_SCAN_ROWS]):
            fuels = [j for j, c in enumerate(r) if isinstance(c, str) and label_key(c) in FUEL_WORDS]
            if len(fuels) < 3:
                continue
            titles = rows[i - 1] if i > 0 else r
            carried, last = {}, ""
            for j in range(len(r)):
                t = norm(titles[j]) if j < len(titles) else ""
                last = t or last
                carried[j] = last
            lab = [j for j, c in enumerate(r) if isinstance(c, str) and re.search(r"rto", c, re.I) and re.search(r"state|city|location|cluster", c, re.I)]
            cluster_col = lab[0] if lab else 1
            cat_col = next((j for j, c in enumerate(r) if isinstance(c, str) and re.search(r"categor", c, re.I)), None)
            cols, extra = [], []
            for j in range(len(r)):
                if j in (cluster_col, cat_col) or not carried[j]:
                    continue
                sub = norm(r[j]) if j < len(r) else ""
                if not sub and not any(isinstance(row[j], (int, float)) for row in rows[i + 1:i + 20] if j < len(row)):
                    continue
                title = hkey(carried[j])
                kind = next((k for k, rx, _ in PC_KINDS if re.search(rx, title)), None)
                fuel = FUEL_WORDS.get(label_key(sub), "all" if re.search(r"all", sub, re.I) else None)
                if kind and fuel:
                    cols.append((j, kind, fuel, carried[j]))
                elif sub or carried[j]:
                    extra.append((j, f"{carried[j]} / {sub}".strip(" /")))
            data = []
            for k in range(i + 1, len(rows)):
                row = rows[k]
                if not any(v not in (None, "") for v in row):
                    continue                                 # a blank row inside the table does not end it
                cl = clean_cluster(row[cluster_col]) if cluster_col < len(row) else ""
                if not cl:
                    if data:
                        break
                    continue
                if not any(isinstance(row[c], (int, float)) for c, _, _, _ in cols if c < len(row)):
                    if data:
                        break
                    continue
                data.append((k + 1, cl, norm(row[cat_col]) if cat_col is not None and cat_col < len(row) else "", {c: row[c] for c, _, _, _ in cols if c < len(row)},
                             {c: row[c] for c, _ in extra if c < len(row)}))
            if data and any(kd == "comp_ncb" for _, kd, _, _ in cols):
                log(f"  Private Car sheet '{name}': {len(data)} clusters, header on rows {i}-{i + 1}")
                table = PcTable(name, i, cluster_col, cat_col, cols, extra, data)
                table.twins = twin_groups([(j, (kd, fu)) for j, kd, fu, _ in cols])
                return table
    return None


# ---------------------------------------------------------------------------------------------------------------------------------- the Two-Wheeler grids
TW_OLD_KINDS = [("comp", r"comp"), ("saod", r"saod"), ("aotp", r"aotp|\btp\b")]


class TwOld:
    twins = {}                                           # column -> the columns under the SAME heading (twin_groups), set by find_tw_old

    def __init__(self, name, hdr, cluster_col, cols, rows):
        self.name, self.hdr, self.cluster_col, self.cols, self.rows = name, hdr, cluster_col, cols, rows       # cols = [(column, body 'BIKE' / 'SCOOTER', kind, title)]


def find_tw_old(wb, ask, warnings, log=say):
    """The two-wheeler sheet: a header row with 'Bike ...' and 'Scooter ...' columns (Comp, SAOD, AOTP) and the RTO category / location columns."""
    for name in wb.sheetnames:
        if wb[name].sheet_state != "visible":
            continue
        rows = sheet_values(wb, name)
        for i, r in enumerate(rows[:HEADER_SCAN_ROWS]):
            cols = []
            for j, c in enumerate(r):
                if not isinstance(c, str):
                    continue
                h = hkey(c)
                body = "BIKE" if h.startswith("bike") else "SCOOTER" if h.startswith("scooter") else None
                kind = next((k for k, rx in TW_OLD_KINDS if re.search(rx, h)), None)
                if body and kind:
                    cols.append((j, body, kind, norm(c)))
            if len({b for _, b, _, _ in cols}) < 2 or len(cols) < 4:
                continue
            lab = [j for j, c in enumerate(r) if isinstance(c, str) and re.search(r"rto", c, re.I) and re.search(r"location|state|cluster", c, re.I)]
            cluster_col = lab[0] if lab else 1
            data = []
            for k in range(i + 1, len(rows)):
                row = rows[k]
                if not any(v not in (None, "") for v in row):
                    continue                                 # a blank row inside the table does not end it
                cl = clean_cluster(row[cluster_col]) if cluster_col < len(row) else ""
                if not cl or not any(isinstance(row[c], (int, float)) for c, _, _, _ in cols if c < len(row)):
                    if data:
                        break
                    continue
                data.append((k + 1, cl, {c: row[c] for c, _, _, _ in cols if c < len(row)}))
            if data:
                log(f"  Two-Wheeler sheet '{name}': {len(data)} clusters x {len(cols)} columns, header on row {i + 1}")
                table = TwOld(name, i, cluster_col, cols, data)
                table.twins = twin_groups([(j, " ".join(ti.lower().split())) for j, _, _, ti in cols])
                return table
    return None


def find_tw_new(wb, ask, warnings, log=say):
    """The new two-wheeler sheets (one per OEM): every sheet with the header RTO Cluster | RTO State | Zone | Bike / Scooter | <payout column, titled with the month>.  -> [(sheet, rows, payout column header)]"""
    found = find_tables(wb, {"cluster": ["rto cluster"], "state": ["rto state"], "zone": ["zone"], "body": ["bike scooter", "bike / scooter", "bike"]}, ("cluster", "body"), exclude=hidden_sheets(wb))
    out = []
    for name, hdr, cols, body in found:
        last = max(cols.values())
        cand = [j for j in range(last + 1, last + 6) if sum(1 for r in body if j < len(r) and isinstance(r[j], (int, float))) >= max(3, len(body) // 3)]
        if not cand:
            warnings.append(f"sheet '{name}' has the new two-wheeler header but no payout column to the right of it: not read")
            continue
        if len(cand) > 1:                                  # two numeric columns right of 'Bike / Scooter'
            heads = sheet_values(wb, name)[hdr]
            texts = [" ".join(norm(heads[j] if j < len(heads) else "").lower().split()) for j in cand]
            if len(set(texts)) > 1:                        # different headings (last month's and this month's?): which one is the payout is not in the sheet
                raise GridError(f"new Two-Wheeler sheet '{name}': expected ONE payout column to the right of 'Bike / Scooter'; found {len(cand)}: "
                                f"{', '.join(col_letter(j + 1) + ' ' + repr(norm(heads[j]) if j < len(heads) else '') for j in cand)} - which of them is this month's is not in the sheet, so nothing is written")
            warnings.append(f"new Two-Wheeler sheet '{name}': the payout column '{norm(heads[cand[0]])}' is there {len(cand)} times (columns {', '.join(col_letter(j + 1) for j in cand)}): "
                            f"the LOWEST payin of them is written for each entry (listed in the sheet 'Lowest rate chosen')")
        rows = []
        for k, r in enumerate(body):
            if cols["cluster"] < len(r) and clean_cluster(r[cols["cluster"]]):
                vals = [(j, r[j] if j < len(r) else None) for j in cand]
                if len(cand) > 1 and any(not blank_cell(v) for _, v in vals):
                    vals = [(j, v) for j, v in vals if not blank_cell(v)]                   # a blank twin cell is no payin
                for j, v in vals:
                    rows.append((hdr + 2 + k, clean_cluster(r[cols["cluster"]]), norm(r[cols["body"]]), v, j))
        title = norm(sheet_values(wb, name)[hdr][cand[0]])
        out.append((name, rows, title, cand[0]))
    if out:
        log(f"  New two-wheeler sheets: {', '.join(x[0] for x in out)}")
    return out


def classify_workbook(wb):
    """What a workbook holds, by what its sheets say: a set of 'cv' / 'pc' / 'tw' / 'tw_new'"""
    roles = set()
    if find_cv_matrices(wb):
        roles.add("cv")
    sink = []
    if find_pc_grid(wb, None, sink, lambda *a, **k: None):
        roles.add("pc")
    if find_tw_old(wb, None, sink, lambda *a, **k: None):
        roles.add("tw")
    if find_tw_new(wb, None, sink, lambda *a, **k: None):
        roles.add("tw_new")
    return roles


# %% ---------------------------------------------------------------- BUILD
# the make groups of the portal's make / model master (the master's spelling; check_names flags a name it lacks)
MK_BIKE, MK_SCOOTER = "BIKEALL", "SCOOTERALL"
MK_PC = "Private_Car_ALL"
GROUP_TOKENS = {"BHARAT_BENZ": "BHARAT_BENZ", "TATA": "TATA", "ASHOK": "ASHOK", "MAHINDRA": "MAHINDRA", "EICHER": "EICHER", "MARUTI": "MARUTI", "HYUNDAI": "HYUNDAI", "CITROEN": "CITROEN",
                "PIAGGAO": "PIAGGIO", "PIAGGIO": "PIAGGIO", "BAJAJ": "BAJAJ", "TVS": "TVS", "EECO": "EECO", "OMNI": "OMNI", "MAGIC": "MAGIC", "FORCE": "FORCE", "SUMO": "SUMO"}
# family -> (the master's 'all makes' group, the pattern of its 'all makes except ...' groups, whether the makes the notes decline are left out of the default)
MAKE_FAMILIES = {
    "gcv4w": ("GCV_4W_ALL", r"GCV_4W_ALL_EX", True),
    "gcv3w": ("GCV_3W_ALL", r"GCV_3W_ALL_EX", False),
    "miscd": ("MISCD_ALL", r"MISCD_ALL_EX", True),
    "pcv3w": ("PCV_3W_ALL", r"PCV_3W_ALL_EX", False),
    "bus": ("BUS_ALL", r"BUS_ALL_EX", True),
    "taxi": ("TAXI_ALL", r"TAXI_?ALL_EX", False),
    "pcv2w": ("PCV_2W_ALL", r"PCV_2W_ALL_EX", False),
}
POSITIVE_GROUPS = {("bus", frozenset({"MAGIC", "FORCE", "SUMO"})): "BUS_MAGIC_FORCE_SUMO"}      # a make set the master has a group for
LAST_KNOWN = {"declined_makes": ("BHARAT_BENZ",), "aotp_copy": True}      # a rule whose note has gone from the grid is ASKED about, with last month's value as the first option


def group_token_set(suffix):
    """'TATA_ASHOK_BHARAT_BENZ' -> {TATA, ASHOK, BHARAT_BENZ}; None when a word of the group's name is not a make this script knows"""
    out = []
    for t in suffix.upper().replace("BHARAT_BENZ", "BHARATBENZ").split("_"):
        t = t.replace("BHARATBENZ", "BHARAT_BENZ")
        if t not in GROUP_TOKENS:
            return None
        out.append(GROUP_TOKENS[t])
    return frozenset(out)


class MakeNames:
    """Turns the makes of a CV cell into the master's names.  A set of makes -> the master's group for exactly that set, or the master's single-make names joined by commas; 'everything but these' ->
    the master's  <family>_EX_<makes>  group (found by its NAME's words, the smallest one that leaves out at least those makes and the declined ones) plus the makes it would leave out too many,
    written as extra names (the user's own form: GCV_4W_ALL_EX_BHARAT_BENZ_TATA_ASHOK_MAHINDRA,Ashok).  What the master cannot express is reported, never replaced by a lookalike."""

    def __init__(self, masters, declined):
        self.masters, self.declined = masters, frozenset(declined)
        self.lower = {k.lower(): k for k in masters.makes}
        self._ex = {}

    def single(self, token):
        n = MAKE_NAME.get(token)
        return self.lower.get(n.lower()) if n else None

    def default(self, family):
        base, pattern, uses = MAKE_FAMILIES[family]
        if uses and self.declined:
            name = f"{base}_EX_{'_'.join(sorted(self.declined))}"
            return self.lower.get(name.lower(), name)
        return self.lower.get(base.lower(), base)

    def exclusions(self, family):
        if family in self._ex:
            return self._ex[family]
        base, pattern, uses = MAKE_FAMILIES[family]
        out = []
        for name in self.masters.makes:
            m = re.match(rf"^{pattern}(?:_(.+))?$", name, re.I)
            m2 = re.match(r"^BUS_EX_(.+)_BHARAT_BENZ$", name, re.I) if family == "bus" else None
            if m2:
                toks = group_token_set(m2.group(1))
                if toks is not None:
                    out.append((name, toks | {"BHARAT_BENZ"}))
                continue
            if not m or not m.group(1):
                continue
            toks = group_token_set(m.group(1))
            if toks is not None:
                out.append((name, toks))
        self._ex[family] = out
        return out

    def text(self, family, dom, order):
        """-> (the new_vehicle_makes text, '') or (None, why it cannot be expressed)"""
        if dom == ALL:
            return self.default(family), ""
        kind, toks = dom
        if kind == "in":
            g = POSITIVE_GROUPS.get((family, frozenset(toks)))
            if g and g.lower() in self.lower:
                return self.lower[g.lower()], ""
            names = []
            for t in sorted(toks, key=lambda x: order.get(x, 99)):
                n = self.single(t)
                if not n:
                    return None, f"the make master has no group for {', '.join(sorted(toks))}"
                names.append(n)
            return ",".join(names), ""
        base, pattern, uses = MAKE_FAMILIES[family]
        want = frozenset(toks) | (self.declined if uses else frozenset())
        best = None
        for name, ex in self.exclusions(family):
            if ex >= want and (best is None or len(ex) < len(best[1])):
                best = (name, ex)
        if not best:
            return None, f"the make master has no group for all makes except {', '.join(sorted(toks))}"
        names = []
        for t in sorted(best[1] - want, key=lambda x: order.get(x, 99)):
            n = self.single(t)
            if not n:
                return None, f"the master group {best[0]} also leaves out {t}, which has no group of its own to give back"
            names.append(n)
        return ",".join([best[0]] + names), ""


# ---------------------------------------------------------------------------------------------------------------------------------- the rows
def fmt_age(x):
    return f"{x:g}"


class Ctx:
    """What the emitters share: the rows, the things not written, the council-call cells, the warnings."""

    def __init__(self, masters, warnings, ask, dates):
        self.masters, self.warnings, self.ask, self.dates = masters, warnings, ask, dates
        self.rows, self.not_written, self.cc_cells = [], [], []
        self.makes = None
        self.places, self.subs = {}, {}
        self.cc_entries = []
        self.unknown_cols = []
        self.zero_skipped = 0
        self.copy_off = False
        self.negative = []

    def add(self, role, **kw):
        row = dict(role=role, policy_type=kw.pop("policy_type", "Renew"), is_ncb=kw.pop("is_ncb", "NA"), is_cpa="NA", is_zd=kw.pop("is_zd", "NA"), tariff_min=0.0, tariff_max=TARIFF_FULL,
                   sc_min=SC_RANGE[0], sc_max=SC_RANGE[1], cc_min=0, cc_max=CC_OPEN_END, vehicle_age_min=0, vehicle_age_max=AGE_OPEN_END, add_tnc="", _flags=(), _fuels=list(FUELS_ALL))
        row.update(kw)
        self.rows.append(row)
        return row

    def skip(self, role, sheet, where, what, text, reason):
        self.not_written.append(dict(workbook=role, sheet=sheet, where=where, what=what, cell=text, reason=reason))


PI_TYPE = {"1+1": "On Net", "1+3": "On OD", "1+5": "On Net", "SAOD": "On OD", "STP": "On TP"}
PI_TYPE_PC = dict(PI_TYPE, **{"1+1": "On OD"})      # the Private Car grid's payouts are OD % (the column titles say so), also the comprehensive ones; only the Act-only column is on TP
CC_FLAG = "Council call: the grid cell says CC - the Council Call list (not converted) may add or replace rates for this cluster and segment; see the sheet 'Council call cells' of the review workbook"


def set_rate(row, rate):
    """the payin lands in the one rate column the pi_type names"""
    pt = row["pi_type"]
    for c in ("pi_od_rate", "pi_tp_rate", "pi_net_rate"):
        row[c] = 0.0
    row[{"On OD": "pi_od_rate", "On TP": "pi_tp_rate", "On Net": "pi_net_rate"}[pt]] = round_half_up(rate)


# ---------------------------------------------------------------------------------------------------------------------------------- Private Car
def percent_of(ask, addr, v, unit):
    """a payout cell -> percent (a blank or a word such as 'No Biz' is a zero card)"""
    kind, x = read_number(ask, addr, v, unit, "a fraction such as 0.25" if unit == "fraction" else "a percent such as 25",
                          zero_words=("NA", "N/A", "-", "NO BIZ", "NO BUSINESS", "DECLINE", "DECLINED", "NIL", "D"), blank_zero=True)
    return 0.0 if kind != "num" else (x * 100 if unit == "fraction" else x)


def table_unit(values):
    nums = [float(v) for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
    return "fraction" if nums and max(nums) <= 1.0 + 1e-9 else "percent"


def tenures(title):
    """'Bike Comp (1+1)/(2+2)/(3+3)' -> ('1+1', ['2+2', '3+3']);  'New (1+3,3+3)' -> ('1+3', ['3+3'])"""
    found = re.findall(r"\d\s*\+\s*\d", title)
    found = [re.sub(r"\s+", "", f) for f in found]
    return (found[0], found[1:]) if found else (None, [])


def emit_pc(ctx, t, ask):
    """Private Car: see the module docstring.  Comprehensive and SAOD with NCB have a column per fuel (LPG takes the CNG column); the others are one column for every fuel."""
    unit = table_unit([v for _, _, _, vals, _ in t.rows for v in vals.values()])
    col = {}
    for c, kind, fuel, _ in t.cols:
        col.setdefault((kind, fuel), []).append(c)               # a twin column (the same heading twice) is a second column of the same entry: its rows are written too, the lowest payin stays
    titles = {kind: title for _, kind, _, title in t.cols}
    for _, kind, _, title in t.cols:
        if kind == "used":
            ctx.warnings.append(f"Private Car column '{title}' (Used car) is not converted: it has no portal sub_product (listed in the sheet 'Not written')")
            ctx.skip("pc", t.name, "all clusters", "Used car", "", "no portal sub_product for a used-car payout")
    for extra_col, label in t.extra:
        ctx.skip("pc", t.name, f"column {col_letter(extra_col + 1)}", label, "", "column not read by this script")
    if t.extra or any(kind == "used" for _, kind, _, _ in t.cols):
        ask.decide("Private Car: " + "; ".join([f"the '{title}' column has no portal sub_product and is not converted" for _, kind, _, title in t.cols if kind == "used"]
                                               + [f"column '{label}' is not read (every cluster is written, as in your uploads)" for _, label in t.extra]))
    for row, cluster, cat, vals, extra in t.rows:
        grp = group_name("pc", cluster)
        rto = rto_cluster(grp, cluster)
        def get(kind, fuel):
            """the payins of one entry: one per column that holds it (a blank twin cell is no payin)"""
            return [percent_of(ask, f"{t.name}!{col_letter(c + 1)}{row}", vals[c], unit) for c in col.get((kind, fuel), []) if not twin_blank(vals, t.twins, c)]

        def put(sub, rate, fuels, ncb, tnc=""):
            r = ctx.add("pc", product="Private Car", sub_product=sub, policy_type="New" if sub == "1+3" else "Renew", make_model_class="Car", new_vehicle_makes=MK_PC, new_rto_list=rto,
                        is_ncb=ncb, add_tnc=tnc, pi_type=PI_TYPE_PC[sub], _fuels=fuels, _cl=cluster, _seg=f"pc_{sub}", _fam="pc")
            set_rate(r, rate)
        for a in get("new", "all"):
            sub, rest = tenures(titles.get("new", ""))
            put(sub or "1+3", a, FUELS_ALL, "NA", TNC_MULTI_YEAR.format(",".join(rest)) if rest else "")
        for kind, sub in (("saod_ncb", "SAOD"), ("comp_ncb", "1+1")):
            per = {f: get(kind, f) for f in ("Petrol", "CNG", "Diesel", "Electric")}
            if per["CNG"]:
                per["LPG"] = per["CNG"]
            for f, rates in per.items():
                for rate in rates:
                    put(sub, rate, [f], "Yes")
        for n in get("noncb", "all"):
            for sub in ("SAOD", "1+1"):
                put(sub, n, FUELS_ALL, "No")
        for s in get("stp", "all"):
            put("STP", s, FUELS_ALL, "NA")


# ---------------------------------------------------------------------------------------------------------------------------------- Two-Wheeler
def emit_tw_old(ctx, t, ask):
    unit = table_unit([v for _, _, vals in t.rows for v in vals.values()])
    for row, cluster, vals in t.rows:
        grp = group_name("tw", cluster)
        rto = rto_cluster(grp, cluster)
        for c, body, kind, title in t.cols:
            if c not in vals or twin_blank(vals, t.twins, c):
                continue
            rate = percent_of(ask, f"{t.name}!{col_letter(c + 1)}{row}", vals[c], unit)
            main, rest = tenures(title)
            sub = {"comp": main or "1+1", "saod": "SAOD", "aotp": "STP"}[kind]
            tnc = TNC_MULTI_YEAR.format(",".join(rest)) if kind == "comp" and rest else ""
            r = ctx.add("tw", product="TW", sub_product=sub, make_model_class=body, new_vehicle_makes=MK_BIKE if body == "BIKE" else MK_SCOOTER, new_rto_list=rto, add_tnc=tnc,
                        pi_type=PI_TYPE[sub], _fuels=FUELS_TW, _cl=cluster, _seg=f"tw_{kind}_{body}", _fam="tw")
            set_rate(r, rate)


TW_OEM_WORDS = {"hmc": "hmc", "hero": "hmc", "heromotocorp": "hmc", "hmsi": "hmsi", "honda": "hmsi", "tvs": "tvs", "suzuki": "suzuki", "yamaha": "yamaha", "re": "re", "royal": "re", "enfield": "re",
                "royalenfield": "re", "ev": "ev", "electric": "ev"}


def tw_oem_key(name):
    """The OEM key of a new Two-Wheeler sheet: its name as it is ('HMC', 'EV'), else the ONE OEM word the name contains ('HMC v2', 'Hero (Sept)', 'Royal Enfield'); '' when it holds none or two"""
    key = label_key(name)
    if key == TW_EV_SHEET or key in TW_MAKES:
        return key
    found = {TW_OEM_WORDS[w] for w in re.findall(r"[a-z]+", str(name).lower()) if w in TW_OEM_WORDS} | ({TW_OEM_WORDS[key]} if key in TW_OEM_WORDS else set())
    return found.pop() if len(found) == 1 else ""


def one_sheet_per_oem(ctx, sheets):
    """Two sheets that say the same OEM (this month's and a copy left in): the one called exactly the OEM's name is used; if none is, the run stops - which one is current is not for a script to guess"""
    by_key = {}
    for item in sheets:
        by_key.setdefault(tw_oem_key(item[0]), []).append(item)
    drop = []
    for k, items in by_key.items():
        if k and len(items) > 1:
            exact = [it for it in items if label_key(it[0]) == k]
            if len(exact) != 1:
                raise GridError(f"the sheets {[i[0] for i in items]} all say {k.upper()} - expected one sheet per OEM; hide or delete the one that is not this month's grid")
            ctx.warnings.append(f"the sheets {[i[0] for i in items]} all say {k.upper()}: only the sheet called exactly '{exact[0][0]}' is used")
            drop += [it for it in items if it is not exact[0]]
    return [it for it in sheets if not any(it is d for d in drop)]


def emit_tw_new(ctx, sheets, ask):
    """One sheet per OEM: HMC = Hero, HMSI = Honda, TVS, Suzuki, Yamaha, RE = Royal Enfield (Petrol, bike and scooter), EV = Electric for all makes (class NA).  A sheet of another name is reported."""
    if sheets:
        ask.decide("new Two-Wheeler (1+5) sheets " + ", ".join(x[0] for x in sheets) + ": payout column '" + "' / '".join(dict.fromkeys(x[2] for x in sheets)) + "'; their rows start on "
                   + FROM_DATE_ARG.get("tw_new", ctx.dates[0]) + " (--tw-new-from-date gives them another start date when ICICI's differs)")
    sheets = one_sheet_per_oem(ctx, sheets)
    converted = 0
    for name, rows, title, pay_col in sheets:
        key = tw_oem_key(name)
        ev = key == TW_EV_SHEET
        if not ev and key not in TW_MAKES:
            ctx.warnings.append(f"sheet '{name}' has the new two-wheeler layout but its name is not an OEM this script knows (HMC, HMSI, TVS, Suzuki, Yamaha, RE, EV): not converted")
            ctx.skip("tw_new", name, "all rows", "OEM sheet", "", "sheet name is not a known OEM")
            continue
        converted += 1
        unit = table_unit([v for _, _, _, v, _ in rows])
        seen = set()
        for row, cluster, body, v, pay in rows:
            b = label_key(body)
            cls = "BIKE" if b == "bike" else "SCOOTER" if b in ("scooter", "scooty") else "NA" if b in ("bikescooter", "bikeandscooter") else None
            if cls is None or (cluster, cls, pay) in seen:
                if cls is None:
                    ctx.skip("tw_new", name, f"row {row}", cluster, body, "body type is neither Bike, Scooter nor Bike & Scooter")
                else:
                    ctx.skip("tw_new", name, f"row {row}", cluster, body, "the same cluster and body type appears twice on this sheet: the first row is used")
                continue
            seen.add((cluster, cls, pay))
            rate = percent_of(ask, f"{name}!{col_letter(pay + 1)}{row}", v, unit)
            grp = group_name("tw15", cluster)
            r = ctx.add("tw_new", product="TW", sub_product="1+5", policy_type="New", make_model_class=cls if not ev else "NA", new_vehicle_makes=TW_ALL if ev else TW_MAKES[key][0],
                        new_rto_list=rto_cluster(grp, cluster), pi_type="On Net", _fuels=["Electric"] if ev else ["Petrol"], _cl=cluster, _seg=f"tw15_{key}_{cls}", _fam="tw")
            set_rate(r, rate)
    if sheets and not converted:
        raise GridError(f"expected sheets named for an OEM (HMC, HMSI, TVS, Suzuki, Yamaha, RE, EV) in the new Two-Wheeler workbook; found {[x[0] for x in sheets]}: none of them could be read, "
                        f"so no new Two-Wheeler row would be written")


# ---------------------------------------------------------------------------------------------------------------------------------- CV
def seg_base_box(seg):
    b = full_box()
    if seg["age"]:
        b["age"] = [(float(seg["age"][0]), float(seg["age"][1]))]
    return b


def place_label(name, cluster, subs, seg):
    """a place of a cell for the plain-English condition: a sub-cluster code with the districts the workbook's RTO list gives it"""
    got = sorted(subs.get((label_key(cluster), name.upper()), ()))
    return f"{name} ({' / '.join(got[:12])}{' ...' if len(got) > 12 else ''})" if got else name


def cv_rto(ctx, cluster, dom, order, seg):
    """new_rto_list of a region: the cluster, or the cluster minus / plus the places the cell names (cluster mode: the plain-English condition; group mode: the user's name form)"""
    if dom == ALL:
        return rto_cluster(group_name("cv", cluster), cluster)
    kind, names = dom
    names = sorted(names, key=lambda n: order.get(n, 99))
    shown = ", ".join(place_label(n, cluster, ctx.subs, seg) for n in names)
    if kind == "in":
        cond, part = f"the grid pays this rate only in {', '.join(names)}", f"only {shown}"
        gname = GROUP_PREFIX["cv"] + "_".join(re.sub(r"\W+", "_", n) for n in names)
    else:
        cond, part = f"{', '.join(names)} {'has' if len(names) == 1 else 'have'} a rate of their own; this row is the rest of the cluster", f"{clean_cluster(cluster)} without {shown}"
        gname = group_name("cv", cluster) + "_EX_" + "_".join(re.sub(r"\W+", "_", n) for n in names)
    if RTO_NAMES == "cluster":
        return rto_cell(RtoScope(cluster_label(cluster), cond, part))
    return rto_cluster(gname, f"{cluster} ({part})")


def body_text(dom, order):
    if dom == ALL:
        return ""
    kind, items = dom
    s = ",".join(sorted(items, key=lambda n: order.get(n, 99)))
    return s if kind == "in" else f"EX:{s}"


def cv_tnc(seg, age, body, order):
    """add_tnc of a CV row: the New / Old tag (when the row's age is exactly the new or the old range), the vehicle-type words of the cell, the segment's fixed T&C text.
    The words of a cell replace 'EX CRANES' (the segment's own wording of the same thing)."""
    tag = seg["tag"] if seg.get("tag_always") else "New" if age == (0.0, 1.0) else "Old" if age == (1.0, float(AGE_OPEN_END)) else ""
    parts = [tag] if tag else []
    if body:
        parts.append(body)
        if seg["text"] and seg["text"] != TNC_MISCD:
            parts.append(seg["text"])
    elif seg["text"]:
        parts.append(seg["text"])
    return ",".join(parts)


def emit_cv_cell(ctx, cluster, header, seg, raw, pol, sheet, addr, unit):
    """The rows of one cluster x segment cell for one policy (pol COMP -> 1+1, AOTP -> STP)."""
    kind, payload = cell_value(raw, unit)
    sub = "1+1" if pol == "COMP" else "STP"
    where = f"{addr} {cluster} / {header}"
    if kind == "error":
        ctx.skip("cv", sheet, where, header, payload, "the cell holds an Excel error")
        return
    if kind in ("num", "zero"):
        parsed = dict(clauses=[dict(box=full_box(), rate=payload, others=False, raw=str(payload))], unknown=[], cc=False, order={})
    else:
        parsed = parse_cell(payload, ctx.places)
        if parsed["unknown"]:
            ctx.skip("cv", sheet, where, header, payload, "not understood: " + " / ".join(dict.fromkeys(parsed["unknown"])))
            return
    flags = ()
    if parsed["cc"]:
        flags = (CC_FLAG,)
        ctx.cc_cells.append(dict(cluster=cluster, segment=header, policy=sub, cell=payload, sheet=sheet))
    declined_rest = []
    regs = regions(parsed, seg_base_box(seg), pol, declined_rest)
    if declined_rest:
        ctx.skip("cv", sheet, where, header, payload, "OTHERS:0% - everything else is declined: no zero card is written (WRITE_OTHERS_ZERO = False, as in the user's uploads)")
    if not regs:
        if parsed["cc"]:
            ctx.skip("cv", sheet, where, header, payload, "council call: the grid gives no fixed rate")
        elif kind == "text":
            ctx.skip("cv", sheet, where, header, payload, "the cell's clauses leave nothing inside this segment")
        return
    for box, rate, raw_clause in regs:
        text, why = ctx.makes.text(seg["family"], box["makes"], parsed["order"])
        if text is None:
            ctx.skip("cv", sheet, where, header, payload, f"{why} (clause '{raw_clause}')" + (" - a 0% decline: the other rows of this cell already leave it out" if rate == 0 else ""))
            if rate == 0:
                ctx.zero_skipped += 1
            continue
        zd = "NA"
        if box["zd"] != ALL and len(box["zd"][1]) == 1 and box["zd"][0] == "in":
            zd = next(iter(box["zd"][1]))
        body = body_text(box["body"], parsed["order"])
        for lo, hi in box["age"]:
            r = ctx.add("cv", product=seg["product"], sub_product=sub, make_model_class=seg["cls"], new_vehicle_makes=text, new_rto_list=cv_rto(ctx, cluster, box["locs"], parsed["order"], seg),
                        vehicle_age_min=lo, vehicle_age_max=hi, cc_min=seg["cc"][0], cc_max=seg["cc"][1], sc_min=seg["sc"][0], sc_max=seg["sc"][1], is_zd=zd,
                        add_tnc=cv_tnc(seg, (lo, hi), body, parsed["order"]), pi_type=PI_TYPE[sub], _fuels=list(seg["fuels"]), _cl=cluster, _seg=seg["key"], _fam=seg["family"], _flags=flags)
            set_rate(r, rate)


def emit_cv(ctx, comp, aotp, ask):
    """Every cell of the comprehensive matrix -> 1+1 rows; the STP rows come from the third-party matrix for the segments it lists and from the comprehensive cell for the others."""
    aotp_cols = {}
    aotp_rows = {}
    if aotp is not None:
        for c, h, seg in aotp.cols:
            aotp_cols.setdefault(seg["key"], []).append((c, h))       # a twin column is a second column of the same segment
        aotp_rows = {label_key(cl): (r, vals) for r, cl, vals in aotp.rows}
        missing = [cl for _, cl, _ in comp.rows if label_key(cl) not in aotp_rows]
        if missing:
            ctx.warnings.append(f"cluster(s) of the comprehensive matrix that the AOTP matrix does not have: {missing}: their STP rows for the AOTP segments are not written")
        extra = [cl for _, cl, _ in aotp.rows if label_key(cl) not in {label_key(c) for _, c, _ in comp.rows}]
        if extra:
            ctx.warnings.append(f"cluster(s) of the AOTP matrix that the comprehensive matrix does not have: {extra}: not written")
    for r, cluster, vals in comp.rows:
        for c, header, seg in comp.cols:
            if twin_blank(vals, comp.twins, c):
                continue                                                     # a blank cell of a twin column is no payin: the other twin has one
            emit_cv_cell(ctx, cluster, header, seg, vals.get(c), "COMP", comp.name, f"{col_letter(c + 1)}{r}", comp.unit)
            if seg["key"] in aotp_cols:
                ar = aotp_rows.get(label_key(cluster))
                if ar is None:
                    ctx.skip("cv", aotp.name, f"{cluster} / {header}", header, "", "the AOTP matrix has no row for this cluster")
                    continue
                for ac, ah in aotp_cols[seg["key"]]:
                    if twin_blank(ar[1], aotp.twins, ac):
                        continue
                    emit_cv_cell(ctx, cluster, ah, seg, ar[1].get(ac), "AOTP", aotp.name, f"{col_letter(ac + 1)}{ar[0]}", aotp.unit)
            elif aotp is None and seg["family"] == "gcv4w" and seg["cc"][0] >= _kg(3.5):
                ctx.skip("cv", comp.name, f"{col_letter(c + 1)}{r} {cluster} / {header}", header, "", "no AOTP matrix in the workbook: the STP payout of this segment cannot be read")
            elif ctx.copy_off:
                ctx.skip("cv", comp.name, f"{col_letter(c + 1)}{r} {cluster} / {header}", header, "", "the AOTP matrix does not list this segment and the rule 'refer to the overall grid' is off")
            else:
                emit_cv_cell(ctx, cluster, header, seg, vals.get(c), "AOTP", comp.name, f"{col_letter(c + 1)}{r}", comp.unit)
    for c, h in comp.unknown_cols:
        ctx.unknown_cols.append((comp.name, h))
    if aotp is not None:
        for c, h in aotp.unknown_cols:
            ctx.unknown_cols.append((aotp.name, h))


# ---------------------------------------------------------------------------------------------------------------------------------- the notes
class Note:
    def __init__(self, role, sheet, row, text):
        self.role, self.sheet, self.row, self.text = role, sheet, row, text
        self.used = False
        self.rest = text


def read_cv_rules(notes, ask, warnings, log=say):
    """The rules the CV notes state in words: the makes declined across the system, the AOTP rule, the taxi seat limit; and a check that the notes behind the fixed T&C texts are still there."""
    rules = {"declined_makes": None, "aotp_copy": None, "taxi_seats": None}
    found_tnc = {"hazard": False, "bus": False, "taxi": False}
    for n in notes:
        if n.role != "cv":
            continue
        t = n.text
        m = re.search(r"([A-Za-z][A-Za-z ]+?)\s+(?:is|are)\s+declined\s+across\s+the\s+system", t, re.I)
        if m:
            tok = GROUP_TOKENS.get(re.sub(r"\s+", "_", m.group(1).strip().upper()))
            if tok:
                rules["declined_makes"] = (rules["declined_makes"] or set()) | {tok}
                n.used = True
            else:
                warnings.append(f"the note '{t[:90]}' declines '{m.group(1).strip()}', which this script does not know as a make (add it to GROUP_TOKENS): the default make groups are not changed for it")
        if re.search(r"AOTP", t) and re.search(r"refer\s+to\s+sheet|not\s+mentioned", t, re.I):
            rules["aotp_copy"] = bool(re.search(r"not\s+mentioned.*(?:refer|see)", t, re.I))
            n.used = True
        m = re.search(r"up\s*to\s*(\d+)\s*seater", t, re.I)
        if m:
            rules["taxi_seats"] = int(m.group(1))
        if re.search(r"hazardous\s+segment", t, re.I):
            found_tnc["hazard"] = True
        if re.search(r"contractor\s+owned\s+bus", t, re.I):
            found_tnc["bus"] = True
        if re.search(r"OLA\s*&\s*UBER", t, re.I):
            found_tnc["taxi"] = True
    for k, text in (("hazard", "hazardous segment (milk / water / chemical tanker)"), ("bus", "contractor owned buses"), ("taxi", "OLA & UBER / tourist taxi")):
        if not found_tnc[k]:
            warnings.append(f"the CV notes no longer mention {text}: the fixed T&C text of those rows is still written - check it with ICICI")
    if rules["declined_makes"] is None:
        pick = ask.choose("the CV note that declines a make across the system ('BHARAT BENZ is declined across the system') was not found",
                          [("Keep last month's: Bharat Benz declined (the default groups leave it out)", "last"), ("No make is declined: use the all-makes groups", "none"), ("Stop", "stop")],
                          question="I can't find the CV note that declines a make across the system. Keep last month's rule (Bharat Benz declined)?", short="declined makes", memo="rule declined_makes")
        if pick == "stop":
            raise GridError("stopped by you - the declined-makes note is missing")
        rules["declined_makes"] = set(LAST_KNOWN["declined_makes"]) if pick == "last" else set()
        warnings.append("the CV note that declines a make across the system was not found: " + ("last month's rule (Bharat Benz declined) is used" if pick == "last" else "no make is declined"))
    if rules["aotp_copy"] is None:
        pick = ask.choose("the CV note that says where the STP payout of a segment comes from was not found",
                          [("Keep last month's rule: the AOTP matrix for the segments it lists, the comprehensive payout for the others", "last"), ("Write STP rows only for the segments the AOTP matrix lists", "only"), ("Stop", "stop")],
                          question="I can't find the CV note about AOTP payouts ('for segments not mentioned in the AOTP grid refer to CV Overall'). What should the STP rows of the other segments use?", short="AOTP rule", memo="rule aotp_copy")
        if pick == "stop":
            raise GridError("stopped by you - the AOTP note is missing")
        rules["aotp_copy"] = pick == "last"
        warnings.append("the CV note about AOTP payouts was not found: " + ("last month's rule is used" if pick == "last" else "STP rows are written only for the segments the AOTP matrix lists"))
    return rules


NOTE_SEGMENTS = [      # words of a note -> the rows it concerns
    (r"tractor", lambda d: d["_seg"].str.startswith("tractor")),
    (r"school\s*bus|staff\s*bus|\bbus", lambda d: d["_fam"] == "bus"),
    (r"taxi", lambda d: d["_fam"] == "taxi"),
    (r"pcv\s*3\s*w|passenger", lambda d: d["_fam"] == "pcv3w"),
    (r"gcv\s*3\s*w|3\s*w", lambda d: d["_fam"] == "gcv3w"),
    (r"garbage|maruti|scv|bolero|supro|jeeto|3\.5\s*ton|ace\b", lambda d: (d["_fam"] == "gcv4w") & (pd.to_numeric(d["cc_max"], errors="coerce") <= 3500.02)),
    (r"crane|misd|misc|mixer|mobile plant", lambda d: d["_fam"] == "miscd"),
    (r"tanker|hazardous", lambda d: d["make_model_class"] == "Tanker"),
    (r"pcv\s*2\s*w", lambda d: d["_fam"] == "pcv2w"),
]


def note_mask(note, df):
    """the rows a note may affect (a superset): the segments it names, narrowed to STP rows when it speaks of AOTP; every row of its workbook when it names none"""
    mine = (df["role"] == note.role) if note.role != "tw" else df["role"].isin(["tw", "tw_new"])
    text = note.text.lower()
    hit = np.zeros(len(df), dtype=bool)
    named = False
    for rx, fn in NOTE_SEGMENTS:
        if re.search(rx, text):
            hit |= np.asarray(fn(df), dtype=bool)
            named = True
    if note.role == "cv" and re.search(r"aotp", text):
        hit = (hit if named else np.ones(len(df), dtype=bool)) & np.asarray(df["sub_product"] == "STP", dtype=bool)
        named = True
    return np.asarray(mine, dtype=bool) & (hit if named else True)


# ---------------------------------------------------------------------------------------------------------------------------------- loading the workbooks
ROLE_NAMES = {"cv": "CV grid", "pc": "Private Car grid", "tw": "Two-Wheeler grid", "tw_new": "new Two-Wheeler (1+5) grid"}


def load_grids(specs, ask, log=say):
    """specs = [(role hint or None, path)] -> {role: (path, workbook)}.  Each workbook is opened and recognised by what its sheets say; a file given for a role that it does not hold is refused."""
    found = {}
    for hint, path in specs:
        log(f"Reading {os.path.basename(str(path))}...")
        path, wb = open_file(ask, path, hint and ROLE_NAMES[hint] or "the ICICI workbook", read_workbook, EXCEL_EXTS)
        roles = classify_workbook(wb)
        if not roles:
            raise GridError(f"{os.path.basename(path)} holds none of the ICICI grids (no CV segment matrix, Private Car grid, Two-Wheeler grid or new Two-Wheeler sheets) - is it the right file?")
        if hint:
            if hint not in roles:
                raise GridError(f"{os.path.basename(path)} was given as the {ROLE_NAMES[hint]} but its sheets look like: {', '.join(ROLE_NAMES[r] for r in sorted(roles))}")
            roles = {hint}
        for role in sorted(roles):
            if role in found:
                pick = ask.choose(f"two workbooks hold the {ROLE_NAMES[role]}: {os.path.basename(found[role][0])} and {os.path.basename(path)}",
                                  [(f"Use {os.path.basename(found[role][0])}", "first"), (f"Use {os.path.basename(path)}", "second")], question=f"Two files hold the {ROLE_NAMES[role]}. Which one is this month's?", short=f"which {role} grid")
                if pick == "second":
                    found[role] = (path, wb)
            else:
                found[role] = (path, wb)
    if not found:
        raise GridError("no workbook given - pass --cv-grid, --pc-grid, --tw-grid, --tw-new-grid (or --grid <file>)")
    return found


VERTICAL_ARG = [None]       # --vertical: one insurer_vertical for every row
FROM_DATE_ARG = {}          # role -> its own from_date (--cv-from-date ...), DD/MM/YYYY


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
    fam = df["role"].map({"cv": "cv", "pc": "pc", "tw": "tw", "tw_new": "tw"})
    df["id"], df["insurance_company"] = ROW_ID, INSURANCE_COMPANY
    df["insurer_vertical"] = VERTICAL_ARG[0] or fam.map(INSURER_VERTICAL)
    for c in ("pi_tp_2", "pi_tp_3", "pi_tp_4", "pi_tp_5", "pi_flat_amount", "pi_vli", "po_flat_amount"):
        df[c] = 0.0
    df["po_type"], df["to_date"], df["user_id"], df["veh_use"], df["remarks"] = df["pi_type"], dates[1], USER_ID, "", ""
    df["from_date"] = [FROM_DATE_ARG.get(r, dates[0]) for r in df["role"]]
    df["tariff_min"] = df["tariff_min"].astype(float)
    return df


def _short(text):
    t = re.sub(r"\s+", " ", str(text)).strip()
    return t if len(t) <= MAX_NOTE_CHARS else t[:MAX_NOTE_CHARS - 3] + "..."


def build(grids, from_date, to_date, log=say, ask=None, loaded=None):
    """Read the workbooks and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    `grids` = [(role hint or None, path)]; from_date / to_date are datetime.date values, written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever a grid itself is missing something.  The notes are never asked about: what they say is applied, and what cannot be applied is flagged in the remarks column."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    warnings = []
    loaded = loaded or load_grids(grids, ask, log)
    masters = load_masters(ask)
    review = ManualReview()
    ctx = Ctx(masters, warnings, ask, dates)
    notes, extra_sheets = [], {}
    done = []
    if "cv" in loaded:
        path, wb = loaded["cv"]
        comp, aotp = pick_cv_matrices(wb, ask, warnings, log)
        skip = {m.name for m in (comp, aotp) if m is not None}
        cc_name, ctx.cc_entries = read_council_calls(wb, skip)
        ctx.places, ctx.subs, n_sources = read_places(wb)
        neg_name, negative = read_negative_list(wb)
        nn = read_notes(wb)
        notes += [Note("cv", s, r, t) for s, r, t in nn]
        if not nn:
            warnings.append("the CV workbook has no notes sheet: the declined-makes and AOTP rules are asked about")
        rules = read_cv_rules(notes, ask, warnings, log)
        ctx.makes = MakeNames(masters, rules["declined_makes"])
        if not n_sources:
            warnings.append("the CV workbook has no RTO list / Negative RTO list / Council Call sheet to name places from: a cell that names a district or a sub-cluster is listed, not converted")
        if aotp is None:
            warnings.append("no AOTP matrix was found in the CV workbook: the STP rows of the AOTP segments (LCV, MHCV) are not written; the other segments take the comprehensive payout")
        ctx.copy_off = not rules["aotp_copy"]
        emit_cv(ctx, comp, aotp, ask)
        for m in (comp, aotp):
            for j, h in (m.ignored if m is not None else ()):
                ask.decide(f"CV sheet '{m.name}': column {col_letter(j + 1)} '{h}' is no segment (your own RTO column?) - not used; new_rto_list carries the grid's cluster names")
        if not cc_name:
            warnings.append("the CV workbook has no Council Call list: the cells that say CC cannot be listed next to its entries")
        if not neg_name:
            warnings.append("the CV workbook has no Negative RTO list: no row is flagged for negative RTOs")
        if cc_name:
            log(f"  Council Call list '{cc_name}': {len(ctx.cc_entries)} entries (listed for the cells that say CC; never turned into rates)")
        if neg_name:
            log(f"  Negative RTO list '{neg_name}': {len(negative)} RTOs still negative (flagged on the rows of their segment and state)")
        ctx.negative = negative
        done.append("cv")
    if "pc" in loaded:
        t = find_pc_grid(loaded["pc"][1], ask, warnings, log)
        emit_pc(ctx, t, ask)
        notes += [Note("pc", s, r, tx) for s, r, tx in read_notes(loaded["pc"][1])]
        done.append("pc")
    if "tw" in loaded:
        wb = loaded["tw"][1]
        emit_tw_old(ctx, find_tw_old(wb, ask, warnings, log), ask)
        notes += [Note("tw", s, r, tx) for s, r, tx in read_notes(wb)]
        done.append("tw")
    if "tw_new" in loaded:
        emit_tw_new(ctx, find_tw_new(loaded["tw_new"][1], ask, warnings, log), ask)
        done.append("tw_new")
    if not ctx.rows:
        raise GridError("nothing to write - no table of the workbooks could be read")
    log(f"Converted: {', '.join(ROLE_NAMES[r] for r in done)}")
    for sheet, h in ctx.unknown_cols:
        warnings.append(f"column '{h}' of sheet '{sheet}' is not a segment this script knows: nothing is written for it")
        ctx.skip("cv", sheet, f"column {h}", h, "", "column header not recognised as a segment")

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
    # the Negative RTO list: the RTOs of a state that are still negative for a segment (the cluster rate does not apply there) - flagged on the rows of that segment's family in the cluster(s) of that state
    if "cv" in loaded and ctx.negative:
        flag_negative(ctx, df, review)
    # the notes nobody understood: standing ones are dropped, the rest flagged on the rows they may affect
    leftovers = []
    for n in notes:
        text = re.sub(r"\s+", " ", n.rest if n.used else n.text).strip(" *.-:")
        if not n.used:
            leftovers.append((n, text))
    unknown = unknown_notes([t for n, t in leftovers], STANDING_NOTES, min_len=MIN_NOTE_CHARS)
    for n, text in leftovers:
        if text in unknown:
            review.flag(_short(text), note_mask(n, df))
    zero_others = sum(1 for e in ctx.not_written if str(e["reason"]).startswith("OTHERS:0%"))
    if zero_others:
        ask.decide(f"{zero_others} CV cell(s) say OTHERS:0% (everything else declined): no zero card is written for it (WRITE_OTHERS_ZERO = False, as in your October upload; confirmed 4 Nov 2026) - listed in the sheet 'Not written'")
    if ctx.cc_cells:
        ask.decide(f"{len(ctx.cc_cells)} CV cell(s) say CC (council call): the rate they state (OTHERS:n%) is written and flagged; the Council Call entries are listed in the sheet 'Council call cells', never turned into rates")
    ask.decide("insurer_vertical = " + ", ".join(sorted(df["insurer_vertical"].unique())) + " on every row (your choice of 4 Nov 2026" + ("; --vertical" if VERTICAL_ARG[0] else "") + ")")
    add_sheet("Not written", pd.DataFrame(ctx.not_written).drop_duplicates() if ctx.not_written else None)
    add_sheet("Council call cells", council_sheet(ctx))
    global LAST_NOT_WRITTEN
    LAST_NOT_WRITTEN = list(ctx.not_written)
    globals()["MASTER_SHEET"] = "the grids"
    df = df.drop(columns=[c for c in df.columns if c.startswith("_") or c == "role"])
    check_names(df, review, masters, log)                                # every RTO / make group written must be in the master lists
    review.report(df, log)                                               # fills the remarks column
    return df[OUTPUT_COLUMNS], warnings


LAST_NOT_WRITTEN = []


NEG_FAMILY = [(r"tractor", lambda d: d["_seg"].str.startswith("tractor")), (r"gcv\s*3w", lambda d: d["_fam"] == "gcv3w"), (r"pcv\s*3w", lambda d: d["_fam"] == "pcv3w"),
              (r"taxi", lambda d: d["_fam"] == "taxi"), (r"bus", lambda d: d["_fam"] == "bus"),
              (r"hcv|mhcv|>\s*40|12|20|40", lambda d: (d["_fam"] == "gcv4w") & (pd.to_numeric(d["cc_min"], errors="coerce") >= 7500)),
              (r"lcv|scv|3\.5|7\.5", lambda d: (d["_fam"] == "gcv4w") & (pd.to_numeric(d["cc_max"], errors="coerce") <= 12000.02))]


def flag_negative(ctx, df, review):
    """Rows of the segments and clusters in which the Negative RTO list still names RTOs: the cluster rate does not apply in those RTOs (the portal cannot take single RTOs out of a cluster rate)."""
    by = {}
    for e in ctx.negative:
        fam = next((i for i, (rx, _) in enumerate(NEG_FAMILY) if re.search(rx, e["segment"].lower())), None)
        if fam is not None:
            by.setdefault((fam, e["state"]), []).append(e)
    clusters = sorted({c for c in df["_cl"].dropna().unique() if isinstance(c, str)})
    for (fam, state), entries in sorted(by.items(), key=lambda kv: str(kv[0])):
        mine = [c for c in clusters if cluster_in_state(c, state)]
        if not mine:
            continue
        mask = np.asarray(NEG_FAMILY[fam][1](df), dtype=bool) & np.asarray(df["_cl"].isin(mine), dtype=bool) & np.asarray(df["role"] == "cv", dtype=bool)
        seg = sorted({e["segment"] for e in entries})
        names = sorted({e["district"] or e["rto"] for e in entries if e["district"] or e["rto"]})
        review.flag(_short(f"Negative RTO list: {len(entries)} RTO(s) of {state} are still negative for {' / '.join(seg)} ({', '.join(names[:6])}{' ...' if len(names) > 6 else ''}): the cluster rate does not apply there"), mask)


def council_sheet(ctx):
    """the cells that say CC (cluster, segment, policy, the cell's text) next to the Active Council Call entries of the cluster's state - informational, an entry is never turned into a rate"""
    if not ctx.cc_cells:
        return None
    out = []
    for c in ctx.cc_cells:
        hits = [e for e in ctx.cc_entries if cluster_in_state(c["cluster"], e["state"]) and not re.search(r"revoked", e["status"], re.I)]
        for e in hits or [None]:
            row = dict(cluster=c["cluster"], segment=c["segment"], policy=c["policy"], cell=c["cell"])
            if e:
                row.update(cc_id=e["id"], cc_state=e["state"], cc_tonnage=e["tonnage"][:100], cc_age=e["age"], cc_transaction=e["txn"], cc_make=e["make"], cc_location=e["loc"], cc_policy=e["pol"],
                           cc_rate=e["coa"], cc_status=e["status"], cc_remark=e["remark"][:100])
            out.append(row)
    return pd.DataFrame(out)


# %% ---------------------------------------------------------------- OUTPUT AND CHECKS
def fmt(x):
    x = round(float(x), RATE_DECIMALS)
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
    d["add_tnc"] = d["add_tnc"].astype(str).str.replace(r"[^\x00-\x7f]", "-", regex=True).str.strip()
    return d[cols or KEY_COLUMNS].astype(str).agg("|".join, axis=1)


def diff_against_reference(df, ref_path, log=say):
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    mine = to_csv_frame(df)
    missing = [c for c in KEY_COLUMNS if c not in ref.columns]
    if missing:
        raise ValueError(f"the earlier upload has no column(s) {missing} - it does not look like an upload file")
    # an earlier upload covers some of the products only (CV, or Private Car + Two-Wheeler): compare the products it has
    mine = mine[mine["product"].isin(set(ref["product"]))].reset_index(drop=True)
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
    log(f"Rows that repeat a key (two grid cells, two payouts: only the lowest is written): {int(df.duplicated(KEY_COLUMNS).sum())}")
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
    po = df[["po_od_rate", "po_tp_rate", "po_net_rate"]].max(axis=1)
    log(f"PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {int(((pi > 0) & (po == 0)).sum())}")
    if LAST_NOT_WRITTEN:
        n = len(LAST_NOT_WRITTEN)
        log(f"Cells / columns NOT converted (listed in the sheet 'Not written' of the review workbook): {n}")
        why = {}
        for e in LAST_NOT_WRITTEN:
            key = re.sub(r"\(clause .*", "", e["reason"]).split(":")[0][:80]
            why[key] = why.get(key, 0) + 1
        for k, v in sorted(why.items(), key=lambda kv: -kv[1])[:8]:
            log(f"  {v:>5}  {k}")
    for w in unique_warnings(warnings):
        log(f"WARNING: {w}")
        RUN_WARNINGS.append(w)
    return df


def main(grids, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None):
    ask = ask or Asker()
    loaded = load_grids(grids, ask, log)
    df, warnings = build(grids, from_date, to_date, log, ask, loaded=loaded)
    df = report(df, warnings, log, ask)
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
    suffix = "" if set(loaded) >= {"cv", "pc", "tw", "tw_new"} else "_" + "_".join(r.upper() for r in ("cv", "pc", "tw", "tw_new") if r in loaded)
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"ICICI_UPLOAD_{from_date:%Y-%m}{suffix}.csv"), ask, log)
    log(f"\nWrote {out_path}")
    return out_path


# %% ---------------------------------------------------------------- ENTRY POINTS
def collect_uploads(ask):
    """Colab: upload the files; asks again if nothing usable came up. Returns ([workbooks], earlier upload or None)."""
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
        if not ask.confirm("No workbook (.xlsx) came with that upload. Upload again?", True):
            raise GridError("no workbook was uploaded")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return books, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the ICICI workbooks (CV, Private Car, Two-Wheeler, new Two-Wheeler - any of them) and, optionally, an earlier upload (.csv) to compare with")
    try:
        books, ref = collect_uploads(ask)
        say(f"\nFiles received: workbooks = {books}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main([(None, b) for b in books], from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("icici_grid_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="ICICI Lombard CV / Private Car / Two-Wheeler grids -> ICICI upload CSV")
    ap.add_argument("--grid", action="append", default=[], help="an ICICI workbook (.xlsx); repeat for several - each is recognised by what its sheets say")
    ap.add_argument("--cv-grid", help="the CV grid workbook")
    ap.add_argument("--pc-grid", help="the Private Car grid workbook")
    ap.add_argument("--tw-grid", help="the Two-Wheeler grid workbook (sheet 'TW Old')")
    ap.add_argument("--tw-new-grid", help="the new Two-Wheeler (1+5) workbook, one sheet per OEM")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    for role in ("cv", "pc", "tw", "tw-new"):
        ap.add_argument(f"--{role}-from-date", help=f"from_date of the {role} workbook's rows, when it differs from --from-date")
    ap.add_argument("--vertical", help="insurer_vertical for every row (default: ICICI-RB-Offline, the user's choice of 4 Nov 2026)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    VERTICAL_ARG[0] = a.vertical
    try:
        specs = [(role, clean_path(p)) for role, p in (("cv", a.cv_grid), ("pc", a.pc_grid), ("tw", a.tw_grid), ("tw_new", a.tw_new_grid)) if p]
        specs += [(None, clean_path(p)) for p in a.grid]
        if not specs and ask.interactive:
            while True:
                try:
                    p = ask_file(ask, "", "Path of an ICICI workbook (CV, Private Car, Two-Wheeler or new Two-Wheeler) - you can drag the file into this window; press Enter when you have given them all:",
                                 exts=EXCEL_EXTS, blank_ok=bool(specs))
                except NeedsAnswer:
                    break
                if not p:
                    break
                specs.append((None, clean_path(p)))
        if not specs:
            sys.exit("STOPPED - no workbook given: pass --cv-grid / --pc-grid / --tw-grid / --tw-new-grid (or --grid <file>)")
        for _, p in specs:
            if not os.path.isfile(p):
                raise GridError(f"workbook not found: {p}")
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
        for role in ("cv", "pc", "tw", "tw_new"):
            own = getattr(a, f"{role}_from_date")
            if own:
                FROM_DATE_ARG[role] = parse_ddmmyyyy(own).strftime("%d/%m/%Y")
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
        main(specs, from_date, to_date, a.out, ref, ask=ask)
    except (GridError, ValueError) as e:
        sys.exit(f"STOPPED - {e}")
    except KeyboardInterrupt:
        sys.exit("STOPPED - stopped by you. No file was written.")
    except Exception as e:
        saved = crash_report("icici_grid_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
