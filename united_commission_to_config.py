"""
United India Insurance (UIIC) "Revised commission structure" PDF -> United upload CSV
======================================================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (62).csv") from United's commission circular
("Revised commission structure.pdf", Motor Department: Sub Annexure 1-4).  The circular is a set of bordered tables, one per
product group; each has an "Existing Pay-out (w.e.f. 01-04-2026)" column pair and a "Pay-out (w.e.f. 01-07-2026)" column pair
(commission %, states).  The script reads the cell borders of every table, so merged cells (one class / band label spread over
several rows, one states cell over several payouts) are joined to the right payouts by position, not by guesswork.  Rates come only
from that PDF; the NEW column (the date in its heading) is what is written.  The "Existing" column is read for the log only.

How a table row becomes upload rows (checked against the July upload: see the regression case)
------------------------------------------------------------------------------------------------
* each payout (commission % + its states) of the new column becomes rows for every fuel the band applies to
    sub_product 1+1  (bundled / package)  pi_type On Net, pi_net_rate = the commission
    sub_product STP  (third party only)   pi_type On TP,  pi_tp_rate  = the commission
  (Two wheelers also: SAOD -> On OD, and 1+5 New = the bundled payouts; Taxis: the Package row is 1+1, the SATP row is STP.)
* po = pi - 7 points, never below 0 (zero stays zero).  A payout written "OD - 10%, TP - 0%" is pi_od 10 / pi_tp 0, On OD and TP.
* states: "All States" -> ALLINDIA; a list of states -> their codes (new_rto_list is a list of state codes, Uttarakhand = UK,UA);
  "Other than above States" -> every state code NOT named by the other payouts of the same band (STATE_CODES lists the 40 codes
  the portal knows).  "(Except RTOs listed in Sub Annexure 3)" -> the portal's RTO group of the other RTOs of those states, and the
  preferred RTOs of Sub Annexure 3 (read from the annexure pages, with their commission) -> the portal's RTO group of the preferred
  RTOs (ANNEX3_GROUPS).  A state the group does not cover (Pondicherry) gets a row with its own state code.
* goods carrying: the GVW band is cc_min / cc_max in kg ("2000 < GVW <= 3500" -> 2000.01-3500.01); E-Cart is GCV 3W / Goods Carrying
  Rickshaw / Electric.  Passenger carrying: the passenger band is sc_min / sc_max ("10 < PCC <= 20" -> 10.01-20.01) and the fuels of
  the band ("All Fuel" 5 fuels, "Except EV" 4, "Electric Vehicles" Electric only).  Classes, make groups and product names are the
  portal's (see CLASSES).
* the PRIVATE CAR table, Sub Annexure 2 (extra incentive by channel) and Sub Annexure 4 (40% RTOs) are NOT converted here: they
  depend on rules the circular leaves to the portal (segments by make / cc / fuel, channels), and are listed as warnings.
* Two Wheeler Long Term Package (5+5), Motor Trade and Standalone CPA have no sub_product / class in the portal's July upload and are
  not written: they are listed as warnings and the rows of that product carry them in the remarks column.

Manual review (the remarks column)
----------------------------------
Anything this script cannot turn into rows - a class or band label it does not know, a states text with a state it does not know, a
commission that is not a percentage, text outside the tables, an effective date that differs from the from_date you typed - is NEVER
guessed at and NEVER stops the run: the rows it may affect carry  Manual Review Required: [the text]  in the `remarks` column (the
column already in the upload format), several notes joined with " | ".  Filter that column, take it up with United and map those rates by hand.

Interactive
-----------
The script checks the PDF while it reads it.  Whenever something about the DOCUMENT it needs is missing or odd - a table whose layout
it does not know, a commission it cannot read, dates that look wrong, a file that will not open or write - it stops, asks you a specific
question (type the number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every answer is listed
at the end of the run, and you confirm before the file is written.  It never asks you for a rate.  With --no-prompt (or when there is no
keyboard) it asks nothing and stops with the question as the error message instead.

Needs the pdfplumber library (pip install pdfplumber; setup.bat installs it).

Google Colab
------------
Paste this whole file into one cell (after `!pip install pdfplumber`) and run it.  Upload the PDF (and, optionally, an earlier upload CSV
to get a diff report).  It then asks for from_date and to_date (DD/MM/YYYY), and downloads UNITED_UPLOAD_<yyyy-mm>.csv.

Command line
------------
    python united_commission_to_config.py [--grid <commission.pdf>] [--from-date 01/07/2026 --to-date 31/07/2026] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the PDF, the dates, the earlier upload - is asked for)
"""
import argparse
import datetime as dt
import difflib
import hashlib
import json
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
INSURANCE_COMPANY = "UNITED INDIA INSURANCE CO LTD"
USER_ID = ""                                     # the July upload carries an empty user_id
ROW_ID = 1                                       # the export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 2
AGE = (0, 99)                                    # vehicle_age_min / max of every row
TARIFF = (0, 100.01)
CC_OPEN_END = 99999                              # cc_max of the top cc / GVW band
SC_FULL = (0, 999)                               # sc_min / sc_max of a row that is not seat-limited
SC_OPEN_END = 999                                # sc_max of the top passenger band
ALL_INDIA = "ALLINDIA"
FUELS_ALL = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NON_EV = ["Petrol", "Diesel", "CNG", "LPG"]
RTO_STRIP = r"^UNITED_(?:[A-Za-z]+\d{2}_)?"                # new_rto_list carries the group name without the insurer / month prefix (the user, 30 Oct 2026)
def rto_condition(value, product, sub_product):
    """(condition, part) in plain English for a row whose RTO scope is a group of an annexure of the circular (the preferred RTOs, or all the other RTOs of the same states, or
    the RTOs of the identified cities); None for state codes and ALLINDIA.  The shared `apply_cluster_names` writes them after the name, with the product and the sub-product
    (the user's rule of 30 Oct 2026: a complex grid's RTO cell holds the cluster name and its condition).  The annexure number, the states and the 'other RTOs' mark are read from
    the group's name (July: Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP, ..._EX_...; August / April: UNITED_RJ_GCV4W_2T-3.5T_Annexure-2, UNITED_CAR_RTO_EX_Annexure3 ...)."""
    v = str(value)
    if "annexure" not in v.lower():
        return None
    no = (re.search(r"annexure[ _-]?(\d)", v, re.I) or [None, "?"])[1]
    other = bool(re.search(r"(?:^|[_ ])EX(?:[_ ]|$)", v))
    if re.search(r"CAR_RTO", v, re.I):
        city = f"the RTOs of the identified cities listed in Sub Annexure {no}"
        return (f"the RTOs of the identified cities have their own row; this row is all the other RTOs", f"all RTOs without {city}") if other else (
            f"the circular pays this commission only in {city}", f"only {city}")
    states = " ".join(t for t in re.split(r"[_ \-]", v) if re.fullmatch(r"[A-Z]{2}", t) and t != "EX") or "these states"
    band = {"UPTO_2T": "up to 2 tonnes", "2T-3.5T": "2 to 3.5 tonnes", "3.5T-7.5T": "3.5 to 7.5 tonnes"}
    gvw = next((f" ({txt} goods vehicles)" for key, txt in band.items() if key in v.upper()), "")
    if other:
        return (f"the preferred RTOs listed in Sub Annexure {no} have their own row; this row is all the other RTOs of the same states",
                f"{states} without the RTOs listed in Sub Annexure {no}{gvw}")
    return (f"the circular pays this commission only in the preferred RTOs listed in Sub Annexure {no}", f"only the RTOs listed in Sub Annexure {no} for {states}{gvw}")


RTO_PREFIX = "UNITED_JULY26_"                    # the portal's RTO groups of Sub Annexure 3 carry the month of the circular
# Sub Annexure 3: GVW band (kg) -> (the portal's RTO group of the PREFERRED RTOs, the group of the OTHER RTOs of the same states); RTO_PREFIX is put in front
ANNEX3_GROUPS = {
    (2000, 3500): ("Sub Annexure3_GCV4W_2T-3.5T_UP_TN_MP", "Sub Annexure3_GCV4W_2T-3.5T_EX_UP_TN_MP"),
    (3500, 7500): ("Sub Annexure3_GCV4W_3.5T-7.5T_HR_MP_RJ_TN_UP", "Sub Annexure3_GCV4W_3.5T-7.5T_EX_HR_MP_RJ_TN_UP"),
}
# the 40 state codes the portal's RTO lists are made of (new_rto_list), and the codes of every state name the circular may use
STATE_CODES = ["AN", "AP", "AR", "AS", "BR", "CG", "CH", "DD", "DL", "DN", "GA", "GJ", "HP", "HR", "JH", "JK", "KA", "KL", "LA", "LD", "MH", "ML", "MN",
               "MP", "MZ", "NL", "OD", "OR", "PB", "PY", "RJ", "SK", "TG", "TN", "TR", "TS", "UA", "UK", "UP", "WB"]
STATE_NAMES = {
    "andaman and nicobar": ["AN"], "andaman & nicobar": ["AN"], "andhra pradesh": ["AP"], "arunachal pradesh": ["AR"], "assam": ["AS"], "bihar": ["BR"],
    "chhattisgarh": ["CG"], "chattisgarh": ["CG"], "chandigarh": ["CH"], "daman and diu": ["DD"], "daman & diu": ["DD"], "delhi": ["DL"],
    "dadra and nagar haveli": ["DN"], "dadra & nagar haveli": ["DN"], "goa": ["GA"], "gujarat": ["GJ"], "himachal pradesh": ["HP"], "haryana": ["HR"],
    "jharkhand": ["JH"], "jammu & kashmir": ["JK"], "jammu and kashmir": ["JK"], "karnataka": ["KA"], "kerala": ["KL"], "ladakh": ["LA"],
    "lakshadweep": ["LD"], "maharashtra": ["MH"], "meghalaya": ["ML"], "manipur": ["MN"], "madhya pradesh": ["MP"], "mizoram": ["MZ"], "nagaland": ["NL"],
    "odisha": ["OD", "OR"], "orissa": ["OD", "OR"], "punjab": ["PB"], "pondicherry": ["PY"], "puducherry": ["PY"], "rajasthan": ["RJ"], "rajastan": ["RJ"], "sikkim": ["SK"],
    "telangana": ["TG", "TS"], "tamil nadu": ["TN"], "tripura": ["TR"], "uttarakhand": ["UA", "UK"], "uttar pradesh": ["UP"], "west bengal": ["WB"],
}
# class / make group / product names of the portal (new_vehicle_makes, make_model_class, product)
CLASSES = {
    "gcv4w": dict(product="GCV 4W", cls="NA", makes="GCV_4W_ALL"),
    "gcv3w": dict(product="GCV 3W", cls="Goods Carrying Rickshaw", makes="GCV_3W_ALL"),
    "bus": dict(product="PCV 4W", cls="Other Bus", makes="BUS_ALL"),
    "taxi": dict(product="PCV 4W", cls="Taxi", makes="TAXI_ALL"),
    "school": dict(product="PCV 4W", cls="School Bus", makes="BUS_ALL"),
    "pcv3w": dict(product="PCV 3W", cls="Passenger Carrying Rickshaw", makes="PCV_3W_ALL"),
    "pcv2w": dict(product="PCV 2W", cls="PCV 2W", makes="PCV_2W_ALL"),
    "tw": dict(product="TW", cls="NA", makes="TW_ALL"),
    "misc_others": dict(product="MISCD", cls="MISCD-Others", makes="MISCD_ALL"),
    "misc_tractor": dict(product="MISCD", cls="MISCD-Tractor", makes="MISCD_ALL"),
    "car": dict(product="Private Car", cls="Car", makes="Private_Car_ALL"),     # the make group of a Private Car row is set per segment (MAKES_CAR_*)
}
# Private Car (the segments of the circular's table): make groups and RTO groups, all spelt as the make / model and RTO masters spell them
MAKES_CAR_ALL = "Private_Car_ALL"
MAKES_CAR_SEVEN = "CAR_Tata_Maruti_Mahindra_Toyota_Hyundai_Honda_Kia"                           # "Tata, Maruti, Mahindra, Toyota, Hyundai, Honda and Kia"
MAKES_CAR_EX_SEVEN = "PRIVATECAR_ALL_EX_UNITED_MARUTI_HYUNDAI_HONDA_KIA_OTHERS"                 # all other makes
CAR_RTO_ANNEX4 = "UNITED_CAR_RTO_Annexure4_JULY26"                                              # the RTOs of Sub Annexure 4 (identified cities, 40%)
CAR_RTO_EX_ANNEX4 = "UNITED_CAR_RTO_EX_Annexure4_JULY26"                                        # the other RTOs
# The circular does not say which policy types the Sub Annexure 4 commission covers. The user decided (26 Oct 2026): Package 1+1 only, as in the July upload.
# The Private Car rows of every other sub_product are written for ALLINDIA at the table's payouts, without a remark. Add a sub_product here to extend it.
CAR_ANNEX4_POLICIES = ("1+1",)
# The portal's RTO groups that a circular's annexures need, by the date its new column takes effect (names exactly as the RTO master spells them). The
# constants above (RTO_PREFIX, ANNEX3_GROUPS, CAR_RTO_*) are the 01/07/2026 circular's. gcv: GVW band -> (the group of the preferred RTOs at the
# annexure's commission, the group of the other RTOs of the same states); car: (the 40% RTOs, all the other RTOs). A circular dated otherwise gets
# no groups: the rows that need one are not written and say so in remarks, until its entry is added here.
CIRCULAR_GROUPS = {
    dt.date(2026, 8, 1): dict(gcv={(2000, 3500): ("UNITED_RJ_GCV4W_2T-3.5T_Annexure-2", "UNITED_RJ_EX_Annexure2_AUG26")},
                              car=("UNITED_CAR_RTO_Annexure3", "UNITED_CAR_RTO_EX_Annexure3")),
    # April 2026 ("Commission structure.pdf", Sub Annexure 2 = preferred GCV RTOs, Sub Annexure 3 = the 40% city RTOs). The master has a group of the OTHER RTOs
    # only for UP up to 2 T; for 2-3.5 T and 3.5-7.5 T (UP, HR, TN, RJ preferred RTOs) it has none, so the second name is None: the payouts "(Except RTOs listed in
    # Sub Annexure-2)" of those two bands are not written and are flagged until the master has the group (the user's rule: a master group that is missing is flagged).
    dt.date(2026, 4, 1): dict(gcv={(0, 2000): ("UNITED_UP_GCV4W_UPTO_2T_Annexure-2", "UNITED_EX_UP_GCV4W_UPTO_2T_Annexure-2"),
                                   (2000, 3500): ("UNITED_UP_HR_TN_RJ_GCV4W_2T-3.5T_Annexure-2", None),
                                   (3500, 7500): ("UNITED_UP_HR_TN_RJ_GCV4W_3.5T-7.5T_Annexure-2", None)},
                              car=("UNITED_CAR_RTO_Annexure3", "UNITED_CAR_RTO_EX_Annexure3")),
}


def circular_groups(effective):
    """The groups of the circular that takes effect on `effective` (None = the PDF gives no date: the July constants, as before) - or None if it is unknown."""
    if effective is None or effective == dt.date(2026, 7, 1):
        return dict(gcv={band: (RTO_PREFIX + a, RTO_PREFIX + b) for band, (a, b) in ANNEX3_GROUPS.items()}, car=(CAR_RTO_ANNEX4, CAR_RTO_EX_ANNEX4))
    return CIRCULAR_GROUPS.get(effective)


TNC_SCHOOL = "Educational Institution, Staff & School Buses"
TNC_AMBULANCE, TNC_NOT_AMBULANCE = "Ambulances", "EX:Ambulances"
MAX_NOTE_CHARS = 500                             # a note longer than this is cut short in the remarks column
# Text outside the tables already known not to need action.  A text is matched when it CONTAINS one of these phrases (capitals, spaces and
# punctuation ignored); any other rule-sized text outside a table is flagged for manual review.
STANDING_NOTES = [
    "INTERMEDIARY COMMISSION/BROKERAGE",             # the title of every page / department (April 2026: "... :: HEALTH DEPARTMENT" above a table that is not part of this upload)
]

PDF_EXTS = (".pdf", ".json")                    # .json = a transcription of a scanned circular (see read_transcription)
EXCEL_EXTS = PDF_EXTS + (".xlsx", ".xlsm", ".xls", ".xlsb")        # the shared upload helpers look at this name

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


OUTPUT_FOLDER = "united"            # the folder of this insurer under output\ (default_output)


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


# %% ---------------------------------------------------------------- READING THE PDF
class Cell:
    """One cell of a bordered table: its box (points, from the top-left of the page) and its text."""
    __slots__ = ("page", "x0", "x1", "top", "bottom", "text")

    def __init__(self, page, x0, x1, top, bottom, text):
        self.page, self.x0, self.x1, self.top, self.bottom, self.text = page, x0, x1, top, bottom, text

    @property
    def mid(self):
        return (self.top + self.bottom) / 2

    def __repr__(self):
        return f"Cell(p{self.page} x{self.x0:.0f}-{self.x1:.0f} y{self.top:.0f}-{self.bottom:.0f} {self.text[:40]!r})"


class PageTable:
    """A bordered table of one page, plus any text lines outside it."""

    def __init__(self, page, bbox, cells, outside):
        self.page, self.bbox, self.cells, self.outside = page, bbox, cells, outside


class PdfError(GridError):
    """The PDF could not be read."""


class TableList(list):
    """The PageTables of a PDF; `.untabled` = (page, has_images) of every page that has no table but is not blank (a cover letter, a notes page);
    `.transcribed_from` = the transcription .json the tables come from when the PDF is a scan."""
    untabled = ()
    transcribed_from = None


def _clean(s):
    return re.sub(r"\s+", " ", str(s or "").replace(" ", " ")).strip()


def _label_key(s):
    return re.sub(r"[^a-z0-9<>=]+", "", _clean(s).lower())


def read_transcription(path):
    """The tables of a scanned circular that someone (or an agent reading the page images) transcribed into a JSON file:
        {"transcription_of": "04001051.pdf", "tables": [{"page": 3, "columns": 7, "rows": [[cell, cell, ...], ...]}]}
    A row lists the cells that START in that row, left to right; a cell is a string, or {"t": text, "rs": rowspan, "cs": colspan} (both default 1). Each cell goes
    into the first free column of its row (the columns taken by a rowspan above are skipped). The cells get synthetic boxes (100 wide, 20 high), so the readers
    below treat a transcription exactly like a table read from a PDF: merged cells are joined to payouts by position."""
    try:
        with open(path, encoding="utf-8") as fh:
            doc = json.load(fh)
        tables = TableList()
        for tno, t in enumerate(doc["tables"], 1):
            ncol, grid, cells = int(t["columns"]), {}, []
            page = int(t.get("page", tno))
            for r, row in enumerate(t["rows"]):
                c = 0
                for cell in row:
                    text, rs, cs = (cell, 1, 1) if isinstance(cell, str) else (cell["t"], int(cell.get("rs", 1)), int(cell.get("cs", 1)))
                    while grid.get((r, c)):
                        c += 1
                    if c + cs > ncol:
                        raise ValueError(f"table {tno}, row {r + 1}: the cell {text[:30]!r} runs past the {ncol} columns")
                    for rr in range(r, r + rs):
                        for cc in range(c, c + cs):
                            grid[(rr, cc)] = True
                    cells.append(Cell(page, 100.0 * c, 100.0 * (c + cs), 20.0 * r, 20.0 * (r + rs), _clean(text)))
                    c += cs
            tables.append(PageTable(page, (0.0, 0.0, 100.0 * ncol, 20.0 * len(t["rows"])), cells, [_clean(x) for x in t.get("outside", [])]))
    except (OSError, ValueError, KeyError, TypeError) as e:
        raise PdfError(f"cannot read the transcription {os.path.basename(str(path))}: {type(e).__name__}: {e}")
    if not tables:
        raise PdfError("the transcription has no tables")
    return tables


def find_transcription(pdf_path):
    """The transcription (.json) of THIS scanned PDF, if one exists: a JSON next to the PDF, in the transcriptions folder or in regression/inputs whose 'source_sha256' is the SHA-256 of
    the PDF's bytes (so a renamed copy is still found and a different PDF with the same name never is). None if there is none."""
    try:
        with open(pdf_path, "rb") as fh:
            digest = hashlib.sha256(fh.read()).hexdigest()
    except OSError:
        return None
    here = os.path.dirname(os.path.abspath(globals().get("__file__", "scripts")))
    for folder in (os.path.dirname(os.path.abspath(pdf_path)), os.path.join(here, "transcriptions"), os.path.join(here, "regression", "inputs")):
        try:
            names = sorted(n for n in os.listdir(folder) if n.lower().endswith(".json"))
        except OSError:
            continue
        for name in names:
            try:
                with open(os.path.join(folder, name), encoding="utf-8") as fh:
                    if json.load(fh).get("source_sha256") == digest:
                        return os.path.join(folder, name)
            except (OSError, ValueError, AttributeError):
                continue
    return None


def read_pdf(path):
    """Open the PDF and return a list of PageTable (one per bordered table found). A .json file is a transcription of a scanned PDF (read_transcription);
    a scanned PDF whose transcription exists (find_transcription) is read through it, and the TableList says so (`transcribed_from`)."""
    if str(path).lower().endswith(".json"):
        return read_transcription(path)
    try:
        import pdfplumber
    except ImportError:
        raise PdfError("this script reads PDFs with the pdfplumber library, which is not installed - run  .venv\\Scripts\\python -m pip install pdfplumber  "
                       "(setup.bat does it), or in Colab  !pip install pdfplumber")
    tables, untabled = TableList(), []
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            found = page.find_tables()
            if not found and (page.images or _clean(page.extract_text())):
                untabled.append((pno, bool(page.images)))
            outside_all = None                                      # the text outside EVERY table of the page (a page with two tables: neither table's text is "outside" the other)
            if len(found) > 1:
                try:
                    cut = page
                    for t in found:
                        cut = cut.outside_bbox(t.bbox)
                    outside_all = [_clean(ln) for ln in (cut.extract_text() or "").split("\n") if _clean(ln)]
                except Exception:
                    outside_all = None
            for k, t in enumerate(found):
                cells = []
                for c in t.cells:
                    x0, top, x1, bottom = c
                    cells.append(Cell(pno, x0, x1, top, bottom, _clean(page.crop(c).extract_text())))
                try:
                    if outside_all is not None:
                        outside = outside_all if k == 0 else []
                    else:
                        outside = [_clean(ln) for ln in (page.outside_bbox(t.bbox).extract_text() or "").split("\n") if _clean(ln)]
                except Exception:
                    outside = []
                tables.append(PageTable(pno, t.bbox, cells, outside))
    if not tables:
        scans = [pno for pno, has_images in untabled if has_images]
        found = find_transcription(path) if scans else None
        if found:
            tables = read_transcription(found)
            tables.transcribed_from = found
            return tables
        if scans:
            raise PdfError(f"this PDF is a scan: {len(scans)} of its pages are pictures with no text and no table the script can read. Render the pages "
                           f"(grid-converter skill: scripts/render_pages.py), read the tables off the images into a transcription .json file "
                           f"(read_transcription in this script explains the format) and run the script with that file instead of the PDF. If the tables of THIS PDF were already "
                           f"transcribed, give the script that .json instead of the PDF, or put it in the transcriptions folder with the PDF's SHA-256 in 'source_sha256' "
                           f"and it is found by itself")
        raise PdfError("no bordered table found in the PDF - is it a different document?")
    tables.untabled = tuple(untabled)
    return tables


# ---- what a cell says
def parse_date(text):
    """the first DD-MM-YYYY (or DD/MM/YYYY) in a text -> date, or None"""
    m = re.search(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", text or "")
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def read_commission(text, where, ask):
    """('num', x) | ('odtp', od, tp) | ('none',) for NA / - | ('unread', text).  An unreadable commission is put to the person (once per text)."""
    t = _clean(text)
    if t.upper() in ("", "NA", "N/A", "-", "--", "NIL"):
        return ("none",)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%", t)
    if m:
        return ("num", float(m.group(1)))
    m = re.fullmatch(r"OD\s*[-:–]\s*(\d+(?:\.\d+)?)\s*%\s*[,;]?\s*TP\s*[-:–]\s*(\d+(?:\.\d+)?)\s*%", t, re.I)
    if m:
        return ("odtp", float(m.group(1)), float(m.group(2)))
    ask.choose(f"{where}: the commission {t!r} is not a percentage",
               [("Leave this payout out of the upload (its product's rows are flagged for manual review)", "skip")],
               question=f"{where}: the commission cell says {t!r}, which I can't read as a percentage. What should I do?",
               short=f"commission {t!r}", memo=("commission", t))
    return ("unread", t)


def read_states(text):
    """(kind, codes, qualifier, leftover): kind 'all' / 'other' (every state not named by the other payouts of the band) / 'list'.
    qualifier 'annexure3' for "(Except RTOs listed in Sub Annexure 3)".  leftover = words that are no state name (the caller flags them)."""
    low = _clean(text).lower()
    if re.fullmatch(r"all(\s+states?)?", low):
        return "all", [], "", ""
    if re.fullmatch(r"other\s+than\s+above\s+states?", low):
        return "other", [], "", ""
    qual = ""
    m = re.search(r"\(\s*except\s+rtos?\s+listed\s+in\s+sub\s+annexure[-\s]*(\d+)[^)]*\)", low)         # "... Sub Annexure 2 where commission would be 56.5%)", April: "Sub Annexure-2"
    if m:
        qual = f"annexure{m.group(1)}"
        low = low.replace(m.group(0), " ")
    codes, rest = [], f" {low} "
    for name in sorted(STATE_NAMES, key=len, reverse=True):
        rest, n = re.subn(r"(?<![a-z])" + re.escape(name) + r"(?![a-z])", " ", rest)
        if n:
            codes += STATE_NAMES[name]
    left = re.sub(r"\band\b|[,&;]|\s+", " ", rest).strip()
    return "list", sorted(set(codes), key=STATE_CODES.index), qual, left


class Record:
    """One payout of the NEW column with the labels of the merged cells around it."""

    def __init__(self, section, page, serial, cls, band, comm, states, existing, date_new, date_old):
        self.section, self.page, self.serial, self.cls, self.band = section, page, serial, cls, band
        self.comm, self.states, self.existing, self.date_new, self.date_old = comm, states, existing, date_new, date_old

    def where(self):
        return f"page {self.page}, {self.cls} / {self.band}"

    def __repr__(self):
        return f"Record({self.section[:25]!r}, {self.cls[:25]!r}, {self.band[:30]!r}, {self.comm!r}, {self.states!r})"


def _inside(c, lo, hi):
    return c.x0 >= lo - 1 and c.x1 <= hi + 1


def _covering(cells, y, lo=None, hi=None):
    return [c for c in cells if c.top - 0.5 <= y <= c.bottom + 0.5 and (lo is None or _inside(c, lo, hi))]


def extract_sections(tables, ask, warnings):
    """Split every table into sections (a full-width title cell, a header with an 'Existing Pay-out' and a 'Pay-out (w.e.f. ...)' column pair, then
    the payout rows) and return (records, annexes).  annexes = [{'title', 'lo', 'hi', 'commission', 'states': {state: [rtos]}}] from Sub Annexure 3."""
    records, annexes = [], []
    for tb in tables:
        width = tb.bbox[2] - tb.bbox[0]
        cells = sorted(tb.cells, key=lambda c: (round(c.top), c.x0))
        full = [c for c in cells if (c.x1 - c.x0) >= 0.9 * width and c.text]
        if any(_label_key(c.text) == "rtos" for c in cells) and any("commission" in _label_key(c.text) for c in cells):
            annexes.append(read_annexure(tb, cells, full))
            continue
        if any(_label_key(c.text).startswith("preferredrtos") for c in full):         # April 2026 layout of the annexures (no 'RTOs' heading cell)
            annexes += read_annexure_grid(tb, cells, full)
            continue
        heads = [c for c in cells if _label_key(c.text).startswith("existingpayout")]
        if not heads:
            for c in full:
                if len(c.text) > 25:
                    warnings.append(f"page {tb.page}: '{c.text[:90]}' is a table this script does not convert (no Existing / new Pay-out columns)")
                    break
            continue
        heads.sort(key=lambda c: c.top)
        for hi_idx, head in enumerate(heads):
            nxt = heads[hi_idx + 1].top if hi_idx + 1 < len(heads) else tb.bbox[3] + 1
            new_head = next((c for c in cells if abs(c.top - head.top) < 1 and c.x0 >= head.x1 - 1 and _is_new_head(c.text)), None)
            titles = [c for c in full if c.bottom <= head.top + 0.5 and (hi_idx == 0 or c.top >= heads[hi_idx - 1].top)]
            title = titles[-1].text if titles else ""
            if "privatecar" in _label_key(title):
                continue                                    # read by extract_car: its columns are different (policy type / segment / business type / fuel / payout)
            if new_head is None:
                raise GridError(f"page {tb.page}: the table '{title[:60]}' has an 'Existing Pay-out' column but no heading of the NEW payout column in the same row to its right - expected "
                                f"'Pay-out (w.e.f. DD-MM-YYYY)', 'Revised Pay-out ...' or a similar heading; found {[c.text[:40] for c in cells if abs(c.top - head.top) < 1 and c.x0 >= head.x1 - 1] or 'nothing'}. "
                                f"The table would be left out of the upload: add its wording to _is_new_head and run again")
            sub = [c for c in cells if c.top >= head.bottom - 0.5 and c.bottom <= head.bottom + 60 and c.top < nxt - 1 and (_label_key(c.text).startswith("commission"))]
            ex_c = next((c for c in sub if _inside(c, head.x0, head.x1)), None)
            nw_c = next((c for c in sub if _inside(c, new_head.x0, new_head.x1)), None)
            if ex_c is None or nw_c is None:
                raise GridError(f"page {tb.page}: the table '{title[:60]}' has the Existing and the new payout columns but not the 'Commission' / 'States' sub-columns under them - expected a "
                                f"'Commission' cell under each (found {[c.text[:30] for c in cells if c.top >= head.bottom - 0.5 and c.bottom <= head.bottom + 60][:8]}). The table would be left out of the upload")
            ex_cols, nw_cols = (ex_c.x0, ex_c.x1, head.x1), (nw_c.x0, nw_c.x1, new_head.x1)
            body_top = max(c.bottom for c in sub)
            body = [c for c in cells if c.top >= body_top - 0.5 and c.bottom <= nxt + 0.5 and c.text != title]
            body = [c for c in body if not ((c.x1 - c.x0) >= 0.9 * width)]
            label_cells = [c for c in body if c.x1 <= ex_cols[0] + 1]
            date_new, date_old = parse_date(new_head.text), parse_date(head.text)
            last_cls = last_band = ""
            for c in sorted((c for c in body if _inside(c, nw_cols[0], nw_cols[1])), key=lambda c: c.top):
                if not c.text.strip():
                    continue                                # an empty payout cell (the blank row under a table) is not a payout
                states = _covering(body, c.mid, nw_cols[1], nw_cols[2])
                ex_comm = _covering(body, c.mid, ex_cols[0], ex_cols[1])
                ex_states = _covering(body, c.mid, ex_cols[1], ex_cols[2])
                labels = sorted(_covering(label_cells, c.mid), key=lambda x: x.x0)
                texts = [x.text for x in labels]
                serial = texts[0] if len(texts) >= 3 else ""
                cls = texts[-2] if len(texts) >= 2 else ""
                band = texts[-1] if texts else ""
                # April 2026 circular: a class / band label sits in one merged cell that does not reach every row it belongs to (the second payout of
                # "Upto 150 CC", the GCV rows between two "Good Carrying Vehicle" cells): a blank label is the one of the row above, in the same table
                new_cls = bool(cls) and cls != last_cls
                cls = cls or last_cls
                band = band if (band or new_cls) else last_band
                last_cls, last_band = cls, band
                existing = ((ex_comm[0].text if ex_comm else ""), (ex_states[0].text if ex_states else ""))
                records.append(Record(title, tb.page, serial, cls, band, c.text, states[0].text if states else "", existing, date_new, date_old))
    # April 2026: the class cell may reach only some of the payouts of one band (GCV "b. 2000<GVW<=3500": the first payout has no class text, the second has it);
    # payouts of one band share a block ("Other than above States" = the states the other payouts of the band do not name), so a blank class takes the band's one class
    classes = {}
    for r in records:
        classes.setdefault((r.page, r.section, r.band), set()).add(r.cls)
    for r in records:
        known = classes[(r.page, r.section, r.band)] - {""}
        if not r.cls and r.band and len(known) == 1:
            r.cls = next(iter(known))
    return records, annexes


def _is_new_head(text):
    """the heading of the new payout column: 'Pay-out (w.e.f. 01-07-2026)', 'Revised Pay-out (w.e.f. 01-08-2026)', and (April 2026 circular)
    'Maximum proposed Commission/Brokerage % w.e.f. 01/04/2026'"""
    k = _label_key(text)
    return k.startswith(("payout", "revisedpayout", "maximumproposedcommission")) or bool(
        re.match(r"(new|revised|proposed|updated|latest|maximum)(proposed)?(payout|commission|brokerage)", k))      # a reworded heading is READ, not skipped


def rto_code(text):
    """'TN 01' / 'TN1' / 'TN01' -> 'TN01' (the master's spelling: two letters, two digits)"""
    return re.sub(r"^([A-Z]{2})(\d)$", r"\g<1>0\2", re.sub(r"\s+", "", text))


def read_annexure(tb, cells, full):
    """An annexure of preferred RTOs. Sub Annexure 3: the GVW band in its title, the RTOs per state; Sub Annexure 4 (Private Car): the RTOs per city, and a
    city that is 'All RTOs except KL 15'. Both: the commission in the merged Commission % cell."""
    title = next((c.text for c in reversed(full) if re.search(r"\d+\s*-\s*\d+\s*kgs?", c.text, re.I)), "")
    m = re.search(r"(\d+)\s*-\s*(\d+)\s*kgs?", title, re.I)
    lo, hi = (int(m.group(1)), int(m.group(2))) if m else (None, None)
    head = next(c for c in cells if _label_key(c.text) == "rtos")
    com_head = next((c for c in cells if _label_key(c.text) == "commission"), None)          # the 'Commission %' column heading (not a title that mentions commission)
    state_head = next((c for c in cells if _label_key(c.text) == "state"), None)
    commission = None
    for c in cells if com_head is not None else []:          # no heading -> no commission; the callers flag/skip an annexure without one
        if _inside(c, com_head.x0, com_head.x1) and c.top >= com_head.bottom - 0.5:
            m2 = re.search(r"(\d+(?:\.\d+)?)\s*%", c.text)
            if m2:
                commission = float(m2.group(1))
                break
    rto_cells = [c for c in cells if _inside(c, head.x0, head.x1) and c.top >= head.bottom - 0.5 and re.fullmatch(r"[A-Z]{2}\s?\d+", c.text)]
    states = {}
    if state_head is not None:
        state_cells = [c for c in cells if _inside(c, state_head.x0, state_head.x1) and c.top >= state_head.bottom - 0.5 and c.text]
        for sc in state_cells:
            states[sc.text] = [rto_code(c.text) for c in rto_cells if sc.top - 0.5 <= c.mid <= sc.bottom + 0.5]
    all_except = []                                         # "All RTOs except KL 15" -> ('KL', {'KL15'})
    for c in cells:
        m3 = re.fullmatch(r"all\s+rtos?\s+except\s+(.+)", c.text, re.I)
        if m3 and _inside(c, head.x0, head.x1):
            gone = {rto_code(x) for x in re.findall(r"[A-Z]{2}\s?\d+", m3.group(1))}
            all_except.append(({g[:2] for g in gone}, gone))
    except_diesel = None
    for c in full:
        m4 = re.search(r"except\s+diesel\s+up\s?to\s+(\d+)\s*cc", c.text, re.I)
        if m4:
            except_diesel = int(m4.group(1))
    no = next((m5.group(1) for m5 in (re.search(r"sub\s+annexure\s*(\d+)", c.text, re.I) for c in full) if m5), "")
    return dict(title=title, lo=lo, hi=hi, commission=commission, states=states, page=tb.page, rtos={rto_code(c.text) for c in rto_cells}, no=no,
                all_except=all_except, except_diesel=except_diesel, city=any(_label_key(c.text) == "city" for c in cells))


def read_annexure_grid(tb, cells, full):
    """The preferred-RTO annexures of the April 2026 circular: Sub Annexure 2 (goods vehicles: one block per GVW band, 'A.Goods Carrying vehicle Upto 2000kgs',
    'B. ... 2000 to 3500 Kgs', 'C. ... 3500 -7500 Kgs', the RTO code cells of each state beside its 'Uttar Pradesh State RTO' cell) and Sub Annexure 3 (Private Car: the RTO
    codes of each identified city, and 'Kerala State  All RTOs except KL 15'). The commission is the merged '40%' / '57.50%' cell of the 'Commission %' column. Returns
    annexure dicts in the shape read_annexure gives (GVW blocks: lo / hi / states; Private Car: city=True, rtos, all_except)."""
    no = next((m.group(1) for m in (re.search(r"sub\s+annexure\s*(\d+)", c.text, re.I) for c in full) if m), "")
    squash = lambda c: re.sub(r"\s+", "", c.text)
    rto_cells = [c for c in cells if re.fullmatch(r"[A-Z]{2}\d{1,2}", squash(c))]
    width = tb.bbox[2] - tb.bbox[0]
    com_head = next((c for c in cells if _label_key(c.text) == "commission"), None)
    com_cells = []
    if com_head is not None:
        for c in cells:
            m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%", _clean(c.text))
            if m and _inside(c, com_head.x0, com_head.x1) and c.top >= com_head.bottom - 0.5:
                com_cells.append((c.top, float(m.group(1))))
    first_x = min((c.x0 for c in rto_cells), default=None)
    names = [c for c in cells if first_x is not None and c.text and c.x1 <= first_x + 1 and (c.x1 - c.x0) < 0.9 * width and not re.fullmatch(r"\(\w+\)", c.text.strip())]
    titles = []                                              # (cell, lo, hi): the GVW block titles
    for c in sorted((c for c in full if re.search(r"goods?\s+carrying", c.text, re.I)), key=lambda c: c.top):
        m = re.search(r"(\d+)\s*(?:to|-|–)\s*(\d+)\s*kgs?", c.text, re.I)
        if m:
            titles.append((c, int(m.group(1)), int(m.group(2))))
        else:
            m = re.search(r"up\s?to\s*(\d+)\s*kgs?", c.text, re.I)
            if m:
                titles.append((c, 0, int(m.group(1))))
    out = []
    if titles:
        for i, (t, lo, hi) in enumerate(titles):
            end = titles[i + 1][0].top if i + 1 < len(titles) else tb.bbox[3] + 1
            states = {}
            block_names = sorted((n for n in names if t.bottom - 0.5 <= n.top < end), key=lambda n: n.top)
            for rc in sorted((c for c in rto_cells if t.bottom - 0.5 <= c.top < end), key=lambda c: (round(c.top), c.x0)):
                owner = [n for n in block_names if n.top <= rc.top + 0.5]
                label = re.sub(r"\s*state\s+rtos?\s*$", "", owner[-1].text.strip(), flags=re.I) if owner else ""
                states.setdefault(label, []).append(rto_code(squash(rc)))
            com = next((v for top, v in sorted(com_cells) if t.bottom - 0.5 <= top < end), None)
            out.append(dict(title=_clean(t.text), lo=lo, hi=hi, commission=com, states=states, page=tb.page, no=no, all_except=[], except_diesel=None, city=False,
                            rtos={rto_code(squash(c)) for c in rto_cells if t.bottom - 0.5 <= c.top < end}))
        return out
    all_except = []                                         # "All RTOs except KL 15" -> ('KL', {'KL15'})
    for c in cells:
        m3 = re.fullmatch(r"all\s+rtos?\s+except\s+(.+)", _clean(c.text), re.I)
        if m3:
            gone = {rto_code(re.sub(r"\s+", "", x)) for x in re.findall(r"[A-Z]{2}\s?\d+", m3.group(1))}
            all_except.append(({g[:2] for g in gone}, gone))
    except_diesel = None
    for c in full:
        m4 = re.search(r"except\s+diesel\s+up\s?to\s+(\d+)\s*cc", c.text, re.I)
        if m4:
            except_diesel = int(m4.group(1))
    city = any(re.search(r"\bcity\b", n.text, re.I) for n in names)
    if not (rto_cells or all_except):
        return []
    return [dict(title=next((_clean(c.text) for c in full if re.search(r"private\s+car", c.text, re.I)), ""), lo=None, hi=None,
                 commission=min(com_cells)[1] if com_cells else None, states={}, page=tb.page, no=no, all_except=all_except, except_diesel=except_diesel,
                 city=city, rtos={rto_code(squash(c)) for c in rto_cells})]


class CarRecord:
    """One payout of the new column of the Private Car table with the labels of the merged cells around it."""

    def __init__(self, page, policy, segment, business, fuel, comm, existing, date_new, date_old):
        self.page, self.policy, self.segment, self.business, self.fuel = page, policy, segment, business, fuel
        self.comm, self.existing, self.date_new, self.date_old = comm, existing, date_new, date_old
        self.section, self.states = "Private Car", ""

    def where(self):
        return f"page {self.page}, Private Car {self.policy} / {self.segment[:40]} / {self.business} / {self.fuel}"


def extract_car(tables, warnings):
    """The Private Car table: policy type / segment / business type / fuel type down the side, 'Existing Pay-out' and 'Pay-out (w.e.f. ...)' across."""
    out = []
    for tb in tables:
        cells = sorted(tb.cells, key=lambda c: (round(c.top), c.x0))
        width = tb.bbox[2] - tb.bbox[0]
        if not any(_label_key(c.text) == "privatecar" and (c.x1 - c.x0) >= 0.9 * width for c in cells) or any(_label_key(c.text) == "rtos" for c in cells) \
                or any(_label_key(c.text).startswith("preferredrtos") for c in cells):
            continue                                        # not the Private Car table (Sub Annexure 4 / April's Sub Annexure 3 also have a 'Private Car' title, but list RTOs)
        if any(_label_key(c.text) == "states" for c in cells) and not any(_label_key(c.text) == "businesstype" for c in cells):
            continue                                        # the August layout (policy type / segment / states / revised commission): read by extract_car_revised
        hdr = {_label_key(c.text): c for c in cells if _label_key(c.text) in ("policytype", "segment", "businesstype", "fueltype")}
        ex = next((c for c in cells if _label_key(c.text).startswith("existingpayout")), None)
        nw = next((c for c in cells if ex is not None and _is_new_head(c.text) and abs(c.top - ex.top) < 1 and c.x0 >= ex.x1 - 1), None)
        if len(hdr) < 4 or ex is None or nw is None:
            raise GridError(f"page {tb.page}: the Private Car table is in a layout this script does not know - expected the columns Policy Type / Segment / Business Type / Fuel Type with an "
                            f"'Existing Pay-out' and a new payout column, or Policy Type / Segment / States with a 'Revised Commission'; found the headings {sorted(hdr)} and "
                            f"{'no Existing column' if ex is None else 'no new payout column'}. The Private Car rows would be left out of the upload")
        top = max(c.bottom for c in (*hdr.values(), ex, nw))
        body = [c for c in cells if c.top >= top - 0.5 and (c.x1 - c.x0) < 0.9 * width]
        for c in sorted((c for c in body if _inside(c, nw.x0, nw.x1)), key=lambda c: c.top):
            lab = {k: next(iter(_covering([b for b in body if _inside(b, h.x0, h.x1)], c.mid)), None) for k, h in hdr.items()}
            ex_c = next(iter(_covering(body, c.mid, ex.x0, ex.x1)), None)
            out.append(CarRecord(tb.page, *(lab[k].text if lab[k] else "" for k in ("policytype", "segment", "businesstype", "fueltype")), c.text,
                                 ex_c.text if ex_c else "", parse_date(nw.text), parse_date(ex.text)))
    return out


class CarRevised:
    """One payout of the Private Car table of the August layout: policy type / segment / states / 'Revised Commission' (no Existing column, no date)."""
    date_new = date_old = None
    section = "Private Car"

    def __init__(self, page, policy, segment, states, comm):
        self.page, self.policy, self.segment, self.states, self.comm = page, policy, segment, states, comm

    def where(self):
        return f"page {self.page}, Private Car {self.policy[:40]} / {self.segment} / {self.states[:30]}"


def extract_car_revised(tables, warnings):
    """The Private Car table that has only a 'Revised Commission' column: policy type, cc segment and states down the side. The heading cell of that last
    column has no border in the circular, so it is not a cell: the commission column is whatever lies to the right of the States column."""
    out = []
    for tb in tables:
        cells = sorted(tb.cells, key=lambda c: (round(c.top), c.x0))
        width = tb.bbox[2] - tb.bbox[0]
        if not any(_label_key(c.text) == "privatecar" and (c.x1 - c.x0) >= 0.9 * width for c in cells) or any(_label_key(c.text) == "businesstype" for c in cells):
            continue
        hdr = {_label_key(c.text): c for c in cells if _label_key(c.text) in ("policytype", "segment", "states")}
        if "states" not in hdr:
            continue                                        # not this layout (the July Private Car table, the RTO annexure)
        if len(hdr) < 3:
            raise GridError(f"page {tb.page}: the Private Car table has a 'States' column but not both of Policy Type / Segment - expected Policy Type / Segment / States / Revised Commission; "
                            f"found the headings {sorted(hdr)}. The Private Car rows would be left out of the upload")
        top = max(c.bottom for c in hdr.values())
        body = [c for c in cells if c.top >= top - 0.5 and (c.x1 - c.x0) < 0.9 * width]
        rev = next((c for c in cells if _label_key(c.text).startswith("revisedcommission") and abs(c.top - hdr["states"].top) < 1), None)
        com = (rev.x0, rev.x1) if rev is not None else (hdr["states"].x1, tb.bbox[2])         # October: the heading is a cell (beside an 'Existing Commission' one)
        for c in sorted((c for c in body if _inside(c, *com)), key=lambda c: c.top):
            lab = {k: next(iter(_covering([b for b in body if _inside(b, h.x0, h.x1)], c.mid)), None) for k, h in hdr.items()}
            out.append(CarRevised(tb.page, *(lab[k].text if lab[k] else "" for k in ("policytype", "segment", "states")), c.text))
    return out


# %% ---------------------------------------------------------------- BUILD
def _short(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text if len(text) <= MAX_NOTE_CHARS else text[:MAX_NOTE_CHARS].rstrip() + "..."


def rto_text(codes):
    """a list of state codes in the portal's order, as new_rto_list"""
    return ",".join(sorted(set(codes), key=STATE_CODES.index))


def cc_band(lo, hi):
    """(from, to) in kg / cc -> (cc_min, cc_max): '2000 < GVW <= 3500' -> (2000.01, 3500.01); no lower limit -> 0; no upper -> CC_OPEN_END"""
    return (0 if not lo else round(lo + 0.01, 2)), (round(hi + 0.01, 2) if hi else CC_OPEN_END)


def _row(spec, sub, policy, kind, fuel, dates, rate):
    """One upload row. rate = ('num', x) or ('odtp', od, tp); kind = 'net' / 'tp' / 'od' / 'odtp'.
    The payout is pi - PO_CUT, never below 0 (zero stays zero)."""
    def po(x):
        return round(max(x - PO_CUT, 0.0), RATE_DECIMALS) if x > 0 else 0.0
    od = tp = net = 0.0
    if rate[0] == "odtp":
        if kind == "odtp":
            od, tp = rate[1], rate[2]
        elif kind == "tp":
            tp = rate[2]
        elif kind == "od":
            od = rate[1]
    elif kind == "net":
        net = rate[1]
    elif kind == "tp":
        tp = rate[1]
    elif kind == "od":
        od = rate[1]
    od, tp, net = round(od, RATE_DECIMALS), round(tp, RATE_DECIMALS), round(net, RATE_DECIMALS)
    label = {"net": "On Net", "tp": "On TP", "od": "On OD", "odtp": "On OD and TP"}[kind]
    c = CLASSES[spec["class"]]
    return dict(
        insurance_company=INSURANCE_COMPANY, insurer_vertical="Not Applicable", product=c["product"], sub_product=sub, policy_type=policy, fuel_type=fuel,
        make_model_class=c["cls"], new_vehicle_makes=spec.get("makes", c["makes"]), new_rto_list=spec["rto"], vehicle_age_min=AGE[0], vehicle_age_max=AGE[1],
        cc_min=spec["cc"][0], cc_max=spec["cc"][1], sc_min=spec["sc"][0], sc_max=spec["sc"][1], pi_od_rate=od, pi_tp_rate=tp, pi_tp_2=0, pi_tp_3=0,
        pi_tp_4=0, pi_tp_5=0, pi_net_rate=net, pi_flat_amount=0, pi_vli=0, pi_type=label, tariff_min=TARIFF[0], tariff_max=TARIFF[1], is_ncb="NA",
        is_cpa="NA", is_zd="NA", from_date=dates[0], to_date=dates[1], user_id=USER_ID, veh_use="", add_tnc=spec.get("tnc", ""), remarks="",
        po_type=label, po_od_rate=po(od), po_tp_rate=po(tp), po_net_rate=po(net), po_flat_amount=0, _product=c["product"],
    )


# the sub_products of a class: (sub_product, policy_type, kind of pi): 'net' = On Net, 'tp' = On TP, 'od' = On OD
STD_SUBS = [("1+1", "Renew", "net"), ("STP", "Renew", "tp")]
TW_SUBS = STD_SUBS + [("1+5", "New", "net")]


def dispatch(rec):
    """What a payout is: (kind, details) - or (None, why) when its class / band is not one this script knows (the payout is then flagged, not guessed)."""
    sec, cls, band = _label_key(rec.section), _label_key(rec.cls), _label_key(rec.band)
    if "twowheelers" in sec:
        if "saod" in cls:
            return "tw_saod", {}
        if "longtermpackage" in cls:
            return "unmapped", f"{rec.cls}: the portal's July upload has no sub_product for it"
        if "bundled" in cls:
            known = {"upto150cc": (None, 150, False), "above150ccandlessthan350cc": (150, 350, False), "above350cc": (350, None, False), "electricvehicles": (None, None, True)}
            if band in known:
                lo, hi, ev = known[band]
                return "tw", dict(lo=lo, hi=hi, ev=ev)
    elif "passengercarrying" in sec:
        if "4wpcv" in cls:
            m = re.search(r"(?:(\d+)\s*<\s*)?(?:PCC|Seating\s+Capacity)\s*(<=|>)\s*(\d+)", rec.band, re.I)
            above = re.search(r">\s*(\d+)\s*passengers", rec.cls, re.I)
            fuels = FUELS_ALL if "allfuel" in band else (FUELS_NON_EV if "exceptev" in band else (["Electric"] if "electric" in band else None))
            if m and fuels:
                lo, op, n = (int(m.group(1)) if m.group(1) else None), m.group(2), int(m.group(3))
                if op == "<=":
                    lo = lo if lo is not None else (int(above.group(1)) if above else None)
                    if lo is not None:
                        return "bus", dict(lo=lo, hi=n, fuels=fuels)
                else:
                    return "bus", dict(lo=n, hi=None, fuels=fuels)
        elif "twowheel" in cls:                                 # "Two Wheeled (PCV)" (July / August), "Two Wheeler (PCV)" (October)
            return "pcv2w", {}
        elif "threewheeled" in cls:
            return "pcv3w", {}
        elif "educational" in cls:
            return "school", {}
        elif "taxi" in cls:
            subs = ([("1+1", "Renew", "net")] if "package" in band else []) + ([("STP", "Renew", "tp")] if ("satp" in band or "standalone" in band) else [])
            if not subs and "package" not in band and "satp" not in band and "standalone" not in band:
                subs = list(STD_SUBS)                       # April: "All CC Bands (including Electric)" names no policy type: like every other PCV row, Package and SATP
            if subs:                                        # October: one row says "Both Package policy & Stand-Alone Third Party policy (SATP)"
                return "taxi", dict(subs=subs)
    elif "goodscarrying" in sec or "goodcarrying" in sec:     # July / August: "GOODS CARRYING VEHICLE", April: "GOOD CARRYING VEHICLE"
        if "ecart" in band:
            return "gcv3w", {}
        m = re.search(r"(?:(\d+)\s*<\s*)?GVW\s*(<=|>)\s*(\d+)", rec.band, re.I)
        if m:
            lo, op, n = (int(m.group(1)) if m.group(1) else None), m.group(2), int(m.group(3))
            return "gcv4w", (dict(lo=lo or 0, hi=n) if op == "<=" else dict(lo=n, hi=None))
    elif "miscellaneous" in sec or "miscvehicles" in sec:     # July: "Miscellaneous ...", April: "Misc Vehicles, Motor Trade and Standalone CPA"
        if "motortrade" in cls or "standalonecpa" in cls:
            return "unmapped", f"{rec.cls}: the portal's July upload has no class for it"
        if "ambulance" in band:
            return "misc", dict(cls="misc_others", tnc=TNC_AMBULANCE)
        if "tractor" in band:
            return "misc", dict(cls="misc_tractor", tnc="")
        if "allothermiscellaneous" in band:
            return "misc", dict(cls="misc_others", tnc=TNC_NOT_AMBULANCE)
    return None, f"the class / band '{rec.cls}' / '{rec.band}' in '{rec.section[:50]}' is not one this script knows"


def specs_for(kind, det, skind, codes, qual, named, annex_by, why, groups=None, notes=None):
    """The row specs of one payout: class, rto list, cc, sc, fuels, sub_products. why = list the reason is appended to when nothing can be written.
    groups = circular_groups(...): the portal's RTO groups of the annexures. An annexure that is NOT in the document is not applied (user's rule, 26 Oct 2026;
    never carried over from an earlier circular): the payout 'except RTOs listed in Sub Annexure N' is then written for the whole state(s) and
    `notes` gets (text, spec) so build() can flag exactly those rows."""
    spec = dict(sc=SC_FULL, cc=(0, CC_OPEN_END), fuels=list(FUELS_ALL), subs=list(STD_SUBS), tnc="")
    if kind == "gcv4w":
        spec.update({"class": "gcv4w", "cc": cc_band(det["lo"], det["hi"])})
    elif kind == "gcv3w":
        spec.update({"class": "gcv3w", "fuels": ["Electric"]})
    elif kind == "tw":
        spec.update({"class": "tw", "subs": list(TW_SUBS), "fuels": ["Electric"] if det["ev"] else ["Petrol"],
                     "cc": (0, CC_OPEN_END) if det["ev"] else cc_band(det["lo"], det["hi"])})
    elif kind == "tw_saod":
        spec.update({"class": "tw", "subs": [("SAOD", "Renew", "od")], "fuels": ["Petrol", "Electric"]})
    elif kind == "bus":
        spec.update({"class": "bus", "fuels": list(det["fuels"]),
                     "sc": (round(det["lo"] + 0.01, 2), round(det["hi"] + 0.01, 2) if det["hi"] else SC_OPEN_END)})
    elif kind == "taxi":
        spec.update({"class": "taxi", "subs": list(det["subs"])})
    elif kind in ("school", "pcv2w", "pcv3w"):
        spec.update({"class": kind})
        if kind == "school":
            spec["tnc"] = TNC_SCHOOL
    elif kind == "misc":
        spec.update({"class": det["cls"], "tnc": det["tnc"]})
    rtos, plain_note = None, ""
    if qual.startswith("annexure"):
        no = qual[len("annexure"):]
        if kind != "gcv4w":
            why.append(f"'(Except RTOs listed in Sub Annexure {no})' on a class that has no annexure group")
            return []
        group = ((groups or {}).get("gcv") or {}).get((det["lo"], det["hi"]))
        annex = annex_by.get((det["lo"], det["hi"]))
        if annex is None:
            plain_note = (f"'(Except RTOs listed in Sub Annexure {no})' for {det['lo']}-{det['hi']} kg: that annexure is not in this document, so it is not applied - "
                          f"the payout is written for the whole state(s), the annexure's RTOs included")
        elif not group or not group[1]:
            why.append(f"'(Except RTOs listed in Sub Annexure {no})' for {det['lo']}-{det['hi']} kg, but no portal RTO group "
                       f"{'for the other RTOs of those states' if group else 'for that band'} (ANNEX3_GROUPS, CIRCULAR_GROUPS) for this circular")
            return []
        else:
            covered = {c for s in annex["states"] for c in STATE_NAMES.get(s.lower(), [])}
            rtos = [group[1]]
            extra = [c for c in codes if c not in covered]
            if extra:                                       # a state named in the payout that the annexure has no RTOs for (Pondicherry)
                rtos.append(rto_text(extra))
    if rtos is None:
        if skind == "all":
            rtos = [ALL_INDIA]
        elif skind == "other":
            rest = [c for c in STATE_CODES if c not in named]
            rtos = [rto_text(rest)] if rest else []
        else:
            rtos = [rto_text(codes)] if codes else []
    if not rtos:
        why.append("no state could be worked out for it")
    specs = [dict(spec, rto=r) for r in rtos]
    if plain_note and notes is not None:
        notes += [(plain_note, sp) for sp in specs]
    return specs


def emit(rows, spec, rate, dates):
    for sub, policy, pkind in spec["subs"]:
        kind = "odtp" if (rate[0] == "odtp" and pkind == "net") else pkind
        for fuel in spec["fuels"]:
            rows.append(_row(spec, sub, policy, kind, fuel, dates, rate))


# ---- Private Car: the table's segments become rows by cc / fuel / make group
class Gap:
    """Something about the Private Car rows that cannot be written or checked, in the shape build() flags (where / comm / states / section)."""

    def __init__(self, text, section="Private Car"):
        self.section, self.comm, self.states, self._where = section, "", "", text

    def where(self):
        return self._where


def _makes_in(text):
    """'Tata, Maruti, Mahindra, Toyota, Hyundai, Honda and Kia' -> {'tata', 'maruti', ...}"""
    return {k for k in (re.sub(r"[^a-z]", "", t.lower()) for t in re.split(r"[/,&]|\band\b", text, flags=re.I)) if k}


def car_segment(label):
    """('s1', n): diesel vehicles up to n cc | ('s2', m, makes): all vehicles above m cc except <makes> | ('s3',): other than above 2 segments |
    ('all',): all vehicles | None: not one this script knows"""
    m = re.search(r"diesel\s+vehicles\s+up\s?to\s+(\d+)\s*cc", label, re.I)
    if m:
        return ("s1", int(m.group(1)))
    m = re.search(r"above\s+(\d+)\s*cc\s+except\s+(.+)$", label, re.I | re.S)
    if m:
        return ("s2", int(m.group(1)), _makes_in(m.group(2)))
    k = _label_key(label)
    if k.startswith("otherthanabove"):
        return ("s3",)
    return ("all",) if k == "allvehicles" else None


# policy type of the table -> (sub_product, policy_type, kind of pi); fuel type -> fuels.  Row 7 (Long Term Package 3+3) has no sub_product in the portal's uploads.
CAR_POLICIES = {
    "bundled13": [("1+3", "New", "net")],
    "package": [("1+1", "Renew", "net")],
    "saod": [("SAOD", "Renew", "od")],
    "satplongtermliability": [("STP", "Renew", "tp")],
    "packagesaodsatplongtermliability": [("1+1", "Renew", "net"), ("SAOD", "Renew", "od"), ("STP", "Renew", "tp")],
    # April 2026 wording ("Brand New SAOD/ Brand New SATP" is left out on purpose: whether it is SAOD / STP with policy type New or Renew is the user's call)
    "satp": [("STP", "Renew", "tp")],
    "packagesaodsatp": [("1+1", "Renew", "net"), ("SAOD", "Renew", "od"), ("STP", "Renew", "tp")],
}
CAR_FUELS = {"nonelectric": FUELS_NON_EV, "electric": ["Electric"], "electricandnonelectric": FUELS_ALL}


def car_parts(seg, n, m, fuels):
    """The (fuels, cc, make group) pieces of a segment. The segments of the table partition every vehicle: diesel up to n cc (s1); above m cc except the
    7 makes (s2); and 'other than above 2 segments' (s3) = the 7 makes above m cc, every other non-diesel up to m cc, diesel from n to m cc. None if s3 lacks n or m."""
    nonev = [f for f in fuels if f != "Electric"]
    diesel = [f for f in fuels if f == "Diesel"]
    if seg[0] == "s1":
        return [dict(fuels=diesel, cc=(0, round(seg[1] + 0.01, 2)), makes=MAKES_CAR_ALL)]
    if seg[0] == "s2":
        return [dict(fuels=nonev, cc=(round(seg[1] + 0.01, 2), CC_OPEN_END), makes=MAKES_CAR_EX_SEVEN)]
    if seg[0] == "s3":
        if n is None or m is None:
            return None
        return [dict(fuels=nonev, cc=(round(m + 0.01, 2), CC_OPEN_END), makes=MAKES_CAR_SEVEN),
                dict(fuels=[f for f in nonev if f != "Diesel"], cc=(0, round(m + 0.01, 2)), makes=MAKES_CAR_ALL),
                dict(fuels=diesel, cc=(round(n + 0.01, 2), round(m + 0.01, 2)), makes=MAKES_CAR_ALL)]
    return [dict(fuels=list(fuels), cc=(0, CC_OPEN_END), makes=MAKES_CAR_ALL)]


def car_specs(cars, annex4, from_date, to_date, ask, why, groups=None):
    """[(spec, rate, dates)] for the Private Car table; `why` collects (record, reason) of what cannot be written.
    Dates: a payout that is the same in the Existing and the new column keeps the Existing column's date (it has been in force since then, as in the July
    upload); a changed or new payout, and the Sub Annexure 4 commission, take from_date. groups = circular_groups(...)."""
    out, fmt_d = [], (lambda d: d.strftime("%d/%m/%Y"))
    a_grp, ex_grp = ((groups or {}).get("car") or (None, None))
    new_dates = (fmt_d(from_date), fmt_d(to_date))
    blocks = {}
    for r in cars:
        blocks.setdefault((_label_key(r.policy), _label_key(r.business), _label_key(r.fuel)), []).append(r)
    for (pk, bk, fk), recs in blocks.items():
        subs, fuels = CAR_POLICIES.get(pk), CAR_FUELS.get(fk)
        if subs is None or fuels is None:
            why += [(r, f"{r.policy}: the portal's July upload has no sub_product for it" if subs is None else f"the fuel type {r.fuel!r} is not one this script knows") for r in recs]
            continue
        segs = [(r, car_segment(r.segment)) for r in recs]
        n = next((s[1] for _, s in segs if s and s[0] == "s1"), None)
        m = next((s[1] for _, s in segs if s and s[0] == "s2"), None)
        s1_rate = None
        for r, seg in segs:
            comm = read_commission(r.comm, r.where(), ask)
            if comm[0] == "unread":
                why.append((r, f"the commission {comm[1]!r} is not a percentage"))
                continue
            if comm[0] != "num":
                continue                                    # NA: no payout for that case
            if seg is None:
                why.append((r, f"the segment {r.segment!r} is not one this script knows"))
                continue
            parts = car_parts(seg, n, m, fuels)
            if parts is None:
                why.append((r, "'Other than above 2 segments' needs the 'diesel up to N cc' and 'above M cc except ...' segments of the same policy type, which were not found"))
                continue
            unchanged = bool(r.date_old) and _clean(r.existing) == _clean(r.comm) and _clean(r.existing) != ""
            dates = (fmt_d(r.date_old), new_dates[1]) if unchanged else new_dates
            if seg[0] in ("s1", "all"):                     # the payout of diesel up to n cc: its own segment, or 'all vehicles'
                s1_rate = (comm[1], dates)
            for sub, policy, kind in subs:
                rto = ex_grp if sub in CAR_ANNEX4_POLICIES else ALL_INDIA
                if rto is None:
                    why.append((r, f"no portal RTO group is known for this circular's Sub Annexure (CIRCULAR_GROUPS), so the {sub} rows cannot be written"))
                    continue
                for part in parts:
                    if part["fuels"]:
                        out.append((dict(part, **{"class": "car"}, rto=rto, sc=SC_FULL, subs=[(sub, policy, kind)], tnc=""), ("num", comm[1]), dates))
        for sub, policy, kind in subs:                      # Sub Annexure 4: the 40% RTOs, all segments except diesel up to n cc
            if sub not in CAR_ANNEX4_POLICIES:
                continue
            if annex4 is None or annex4["commission"] is None:
                why.append((Gap(f"Private Car {sub}"), "Sub Annexure 4 (the 40% RTOs) was not found or has no commission % - no 40% rows are written"))
                continue
            if a_grp is None:
                continue                                    # reported above, row by row
            cut = annex4["except_diesel"] or n
            if n is not None and annex4["except_diesel"] and annex4["except_diesel"] != n:
                why.append((Gap(f"Private Car {sub}"), f"Sub Annexure 4 leaves out diesel up to {annex4['except_diesel']} cc but the table's diesel segment is up to {n} cc"))
            base = dict({"class": "car"}, rto=a_grp, makes=MAKES_CAR_ALL, sc=SC_FULL, subs=[(sub, policy, kind)], tnc="")
            other = [f for f in fuels if f != "Diesel"]
            if other:
                out.append((dict(base, fuels=other, cc=(0, CC_OPEN_END)), ("num", annex4["commission"]), new_dates))
            if "Diesel" in fuels and cut:
                out.append((dict(base, fuels=["Diesel"], cc=(round(cut + 0.01, 2), CC_OPEN_END)), ("num", annex4["commission"]), new_dates))
                if s1_rate:
                    out.append((dict(base, fuels=["Diesel"], cc=(0, round(cut + 0.01, 2))), ("num", s1_rate[0]), s1_rate[1]))
                else:
                    why.append((Gap(f"Private Car {sub}"), f"no 'diesel up to {cut} cc' payout in the table, so diesel up to {cut} cc has no row in the Sub Annexure 4 RTOs"))
    return out


def car_policies(label):
    """The policy-type cell of the August / October Private Car table -> (fuels, [(sub_product, policy_type, kind of pi)], [phrases that have no sub_product]).
    'Non Electric - ...' = the four non-electric fuels, 'Electric Cars - ...' = Electric, 'Non Electric & Electric - SAOD' = all five."""
    low, unwritten = _clean(label).lower(), []
    plain = re.sub(r"non[\s-]*electric", " ", low)
    fuels = (list(FUELS_ALL) if "electric" in plain else list(FUELS_NON_EV)) if plain != low else (["Electric"] if "electric" in low else list(FUELS_NON_EV))
    for phrase, name in (("long term package", "Long Term Package (3+3)"), ("long term liability", "Long term Liability")):
        if phrase in low:
            unwritten.append(name)
            low = low.replace(phrase, " ")
    subs = [s for word, s in (("bundled", ("1+3", "New", "net")), ("package", ("1+1", "Renew", "net")), ("satp", ("STP", "Renew", "tp")),
                              ("saod", ("SAOD", "Renew", "od"))) if word in low]
    return fuels, subs, unwritten


def car_cc(label):
    """'Upto 1000 CC' / '>1000 CC and upto 1200 CC' / '>2000 CC' / 'All Vehicles' -> (cc_min, cc_max), or None"""
    low = _clean(label).lower()
    if re.fullmatch(r"all\s+vehicles?", low):
        return 0, CC_OPEN_END
    m = re.fullmatch(r"upto\s+(\d+)\s*cc", low)
    if m:
        return cc_band(None, int(m.group(1)))
    m = re.fullmatch(r">\s*(\d+)\s*cc\s+and\s+upto\s+(\d+)\s*cc", low)
    if m:
        return cc_band(int(m.group(1)), int(m.group(2)))
    m = re.fullmatch(r">\s*(\d+)\s*cc", low)
    return cc_band(int(m.group(1)), None) if m else None


def car_specs_revised(recs, annex, groups, from_date, to_date, ask, why, overlaps):
    """[(spec, rate, dates)] for the Private Car table of the August layout (cc segment x states, no make groups, no existing column; every row takes
    from_date). Package (1+1) is the only policy type that gets Sub Annexure 3's 40% (CAR_ANNEX4_POLICIES). In its RTOs the 40% replaces the table's
    payout for every fuel except diesel up to the annexure's cc, so the table's rows must not cover those RTOs too: an 'All states' row goes to the
    group of all the other RTOs, but a row for some states cannot - the master has no 'these states except the 40% RTOs' group - so it is written for the
    states and `overlaps` gets a note that build() flags on exactly those rows. `why` collects (record, reason) of what cannot be written."""
    out = []
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    a_grp, ex_grp = ((groups or {}).get("car") or (None, None))
    cut = (annex or {}).get("except_diesel")
    cut_cc = round(cut + 0.01, 2) if cut else None
    annex_states = set()
    if annex:
        annex_states = {c[:2] for c in annex["rtos"]} | {p for prefixes, _ in annex["all_except"] for p in prefixes}
    blocks, seen_unwritten = {}, set()
    for r in recs:
        blocks.setdefault(r.policy, {}).setdefault(r.segment, []).append(r)
    for policy, segs in blocks.items():
        fuels, subs, unwritten = car_policies(policy)
        first = next(iter(segs.values()))[0]
        for phrase in unwritten:
            if phrase not in seen_unwritten:
                seen_unwritten.add(phrase)
                why.append((Gap("Private Car"), f"{phrase} has no sub_product in the portal's upload, so it is not written"))
        if not subs:
            why.append((first, f"the policy type {policy!r} is not one this script knows"))
            continue
        for seg_label, rs in segs.items():
            cc = car_cc(seg_label)
            if cc is None:
                why += [(r, f"the segment {seg_label!r} is not one this script knows") for r in rs]
                continue
            parsed = [(r, read_commission(r.comm, r.where(), ask), read_states(r.states)) for r in rs]
            named = {c for _, _, (k, codes, q, left) in parsed if k == "list" for c in codes}
            for r, comm, (skind, codes, qual, left) in parsed:
                if comm[0] == "unread":
                    why.append((r, f"the commission {comm[1]!r} is not a percentage"))
                    continue
                if comm[0] != "num":
                    if comm[0] == "odtp":
                        why.append((r, "an OD / TP split payout in the Private Car table is not one this script knows"))
                    continue                                # NA: no payout for that case
                if left:
                    why.append((r, f"the states text has words that are no state name: {left!r}"))
                    continue
                if skind == "all":
                    plain, row_states = ALL_INDIA, set(STATE_CODES)
                elif skind == "other":
                    rest = [c for c in STATE_CODES if c not in named]
                    plain, row_states = (rto_text(rest) if rest else None), set(rest)
                else:
                    plain, row_states = (rto_text(codes) if codes else None), set(codes)
                if plain is None:
                    why.append((r, "no state could be worked out for it"))
                    continue
                for sub, ptype, kind in subs:
                    base = dict({"class": "car"}, makes=MAKES_CAR_ALL, sc=SC_FULL, subs=[(sub, ptype, kind)], tnc="", cc=cc)
                    over = []
                    if sub in CAR_ANNEX4_POLICIES and annex is not None and annex["commission"] is not None:
                        if "Diesel" in fuels and cut_cc and cc[0] < cut_cc < cc[1]:
                            why.append((r, f"the segment {seg_label!r} runs across the annexure's diesel limit of {cut} cc, so its diesel rows are written for the whole scope"))
                        over = [f for f in fuels if f != "Diesel" or (cut_cc and cc[0] >= cut_cc)]
                    keep = [f for f in fuels if f not in over]
                    if keep:
                        out.append((dict(base, fuels=keep, rto=plain), ("num", comm[1]), dates))
                    if not over:
                        continue
                    if skind == "all":
                        if ex_grp is None:
                            why.append((r, f"no portal RTO group is known for this circular's Sub Annexure (CIRCULAR_GROUPS), so the {sub} rows cannot be written"))
                            continue
                        out.append((dict(base, fuels=over, rto=ex_grp), ("num", comm[1]), dates))
                    else:
                        out.append((dict(base, fuels=over, rto=plain), ("num", comm[1]), dates))
                        shared = sorted(row_states & annex_states, key=STATE_CODES.index)
                        if shared:
                            overlaps.append(dict(rto=plain, sub=sub, fuels=list(over), cc=cc, states=shared, seg=seg_label))
        for sub, ptype, kind in subs:                       # Sub Annexure 3: the 40% RTOs, every fuel except diesel up to `cut` cc
            if sub not in CAR_ANNEX4_POLICIES:
                continue
            if annex is None or annex["commission"] is None:
                why.append((Gap(f"Private Car {sub}"), "the annexure of the 40% RTOs was not found in this document (or has no commission %), so it is not applied - no 40% rows are written"))
                continue
            if a_grp is None:
                continue                                    # reported above, row by row
            base = dict({"class": "car"}, rto=a_grp, makes=MAKES_CAR_ALL, sc=SC_FULL, subs=[(sub, ptype, kind)], tnc="")
            other = [f for f in fuels if f != "Diesel"]
            if other:
                out.append((dict(base, fuels=other, cc=(0, CC_OPEN_END)), ("num", annex["commission"]), dates))
            if "Diesel" in fuels and cut_cc:
                out.append((dict(base, fuels=["Diesel"], cc=(cut_cc, CC_OPEN_END)), ("num", annex["commission"]), dates))
    return out


def master_checks(df, review, masters, annex_by, annex4, seven, groups=None):
    """The master lists are the authority, so a master group that does not hold what the circular says is flagged on the rows that use it: the Sub Annexure 3
    and 4 RTO groups against the RTO lists printed in the circular, the groups of the other RTOs against 'all RTOs of the state minus those', the Private Car
    make groups against the makes the table names. Nothing is corrected here."""
    allc = masters.codes(ALL_INDIA)

    def flag(col, name, text):
        review.flag(_short(text), lambda d, c=col, n=name: np.array(d[c].astype(str) == n, dtype=bool))

    gcv_groups = ((groups or {}).get("gcv")) or {}
    a_grp, ex_grp = ((groups or {}).get("car")) or (None, None)
    for (lo, hi), a in sorted(annex_by.items()):
        group = gcv_groups.get((lo, hi))
        if not group:
            continue
        pdf = {c for v in a["states"].values() for c in v}
        for name, expected, what in ((group[0], pdf, f"the preferred RTOs listed in Sub Annexure {a.get('no') or 3}"),
                                     (group[1], {c for c in allc if c[:2] in {x for s in a["states"] for x in STATE_NAMES.get(s.lower(), [])}} - pdf,
                                      "the other RTOs of those states")):
            have = masters.codes(name) if name else set()
            if have and have != expected:
                flag("new_rto_list", name, f"master RTO group '{name}' does not hold {what} ({lo}-{hi} kg): only in the circular {sorted(expected - have)[:10]}, "
                                           f"only in the master {sorted(have - expected)[:10]}")
    if annex4 and a_grp:
        no = annex4.get("no") or 4
        pdf4 = set(annex4["rtos"])
        for prefixes, gone in annex4["all_except"]:
            pdf4 |= {c for c in allc if c[:2] in prefixes and c not in gone}
        have = masters.codes(a_grp)
        if have and have != pdf4:
            flag("new_rto_list", a_grp, f"master RTO group '{a_grp}' does not hold the RTOs of Sub Annexure {no}: only in the circular {sorted(pdf4 - have)[:10]}, "
                                        f"only in the master {sorted(have - pdf4)[:10]}")
        ex = masters.codes(ex_grp)
        gap = allc - have - ex
        if ex and gap:
            states = {}
            for c in gap:
                states[c[:2]] = states.get(c[:2], 0) + 1
            flag("new_rto_list", ex_grp, f"master RTO group '{ex_grp}' lacks {len(gap)} RTO(s) that {ALL_INDIA} has and no Annexure {no} group has "
                                         f"({', '.join(f'{s} {n}' for s, n in sorted(states.items()))}): policies registered there match no rate of this group")
    models = masters.makes.get(MAKES_CAR_ALL)
    if models and seven:
        allm = set(models)
        absent = sorted(mk for mk in seven if not any(e.split()[0].lower() == mk for e in allm))
        if absent:
            flag("new_vehicle_makes", MAKES_CAR_SEVEN, f"the make(s) {absent} named in the Private Car table have no model in the master's {MAKES_CAR_ALL}")
        sev = {e for e in allm if e.split()[0].lower() in seven}
        for name, expected, what in ((MAKES_CAR_SEVEN, sev, "the models of the makes the table names"), (MAKES_CAR_EX_SEVEN, allm - sev, "all other models")):
            have = set(masters.makes.get(name, ()))
            if have and have != expected:
                flag("new_vehicle_makes", name, f"master make group '{name}' does not hold {what} of {MAKES_CAR_ALL}: it lacks {len(expected - have)} "
                                                f"{sorted(expected - have)[:6]} and has {len(have - expected)} that do not belong {sorted(have - expected)[:6]}")


_SECTION_PRODUCTS = (("twowheelers", ["TW"]), ("passengercarrying", ["PCV 4W", "PCV 3W", "PCV 2W"]), ("goodscarrying", ["GCV 4W", "GCV 3W"]),
                     ("goodcarrying", ["GCV 4W", "GCV 3W"]), ("miscellaneous", ["MISCD"]), ("miscvehicles", ["MISCD"]), ("privatecar", ["Private Car"]))


def products_of(section):
    k = _label_key(section)
    return next((p for key, p in _SECTION_PRODUCTS if key in k), None)


def build(pdf_path, from_date, to_date, log=say, ask=None, tables=None):
    """Read the PDF and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    from_date / to_date are datetime.date values; they are written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever the document itself is missing something. Text rules that cannot be applied are never asked
    about: they are flagged in the remarks column (see ManualReview)."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    if tables is None:
        log(f"Reading {os.path.basename(str(pdf_path))}...")
        pdf_path, tables = open_file(ask, pdf_path, "the United PDF", read_pdf, PDF_EXTS)
        if getattr(tables, "transcribed_from", None):
            log(f"This PDF is a scan: using its transcription {tables.transcribed_from} (the tables were read off the page images, not parsed from the PDF)")
    warnings, rows, review = [], [], ManualReview()
    records, annexes = extract_sections(tables, ask, warnings)
    cars, annex4, annex_by = extract_car(tables, warnings), None, {}
    cars_rev = extract_car_revised(tables, warnings)       # the August layout of the Private Car table (revised commission only)
    for a in annexes:                                       # one Sub Annexure 3 band can run over several pages
        if a["lo"] is None:
            if a["city"] and annex4 is None:
                annex4 = a                                  # Sub Annexure 4: the RTOs of identified cities (Private Car)
            elif a["city"]:
                annex4["rtos"] |= a["rtos"]
                annex4["all_except"] += a["all_except"]
                annex4["commission"] = annex4["commission"] if annex4["commission"] is not None else a["commission"]
            else:
                warnings.append(f"page {a['page']}: '{a['title'][:70] or 'a table of RTOs'}' ({a['commission']}% for the listed RTOs) is not converted")
            continue
        got = annex_by.setdefault((a["lo"], a["hi"]), dict(lo=a["lo"], hi=a["hi"], commission=None, states={}, no=a["no"]))
        got["commission"] = got["commission"] if got["commission"] is not None else a["commission"]
        for s, r in a["states"].items():
            got["states"].setdefault(s, []).extend(r)
    log(f"Read {len(tables)} table(s): {len(records)} payout(s) of the new column, {len(cars) + len(cars_rev)} Private Car payout(s), {len(annex_by)} GCV annexure band(s)"
        + (f", Sub Annexure {annex4['no'] or 4} ({annex4['commission']}% for {len(annex4['rtos'])} listed RTOs and {len(annex4['all_except'])} whole state(s))" if annex4 else ""))
    if not records and not cars and not cars_rev:
        raise GridError("no payout found in the PDF - looked for tables with an 'Existing Pay-out' and a 'Pay-out (w.e.f. ...)' column")
    new_dates = sorted({r.date_new for r in records + cars if r.date_new})
    groups = circular_groups(new_dates[-1] if new_dates else None)         # the portal's RTO groups of THIS circular's annexures
    if groups is None:
        warnings.append(f"no portal RTO group is known for the circular effective {new_dates[-1]:%d/%m/%Y} (CIRCULAR_GROUPS): the rows that need an annexure group are not written")
    for pno, has_images in getattr(tables, "untabled", ()):
        what = "a scanned image whose text this script cannot read" if has_images else "text"
        warnings.append(f"page {pno} has no table (a cover letter or notes: {what}) - it is NOT read by this script: read it for rules that change the rows")
    if new_dates:
        old = sorted({r.date_old for r in records + cars if r.date_old})
        log(f"The new column is effective from {', '.join(d.strftime('%d/%m/%Y') for d in new_dates)}"
            + (f" (existing column: {', '.join(d.strftime('%d/%m/%Y') for d in old)})" if old else ""))
        if any(d != from_date for d in new_dates):
            review.flag(f"the PDF says the new payouts are effective from {', '.join(d.strftime('%d/%m/%Y') for d in new_dates)} "
                        f"but these rows carry from_date {from_date:%d/%m/%Y}")
    else:
        warnings.append("no 'w.e.f.' date found in the heading of the new column - from_date is the one you typed")
    for tb in tables:
        for line in tb.outside:
            if unknown_notes([line], STANDING_NOTES):                  # a rule printed under one table may concern every product: flagged on every row (a superset, never fewer)
                review.flag(_short(f"text outside the tables (page {tb.page}): {line}"))

    blocks = {}                                             # all the payouts of one class / band share a block ("Other than above States" = not named by the others)
    for rec in records:
        blocks.setdefault((rec.page, rec.section, rec.cls, rec.band), []).append(rec)
    not_written, applied_plain = [], []                     # (record, why): flagged in remarks, never guessed | (text, spec): a payout written without its annexure
    for recs in blocks.values():
        kind, det = dispatch(recs[0])
        if kind in (None, "unmapped"):
            not_written += [(r, det) for r in recs]
            continue
        parsed = [(r, read_commission(r.comm, r.where(), ask), read_states(r.states)) for r in recs]
        named = {c for _, _, (k, codes, q, left) in parsed if k == "list" for c in codes}
        for r, comm, (skind, codes, qual, left) in parsed:
            if comm[0] == "unread":
                not_written.append((r, f"the commission {comm[1]!r} is not a percentage"))
            elif comm[0] == "none":
                continue                                    # NA: the new column has no payout for that case
            elif left:
                not_written.append((r, f"the states text has words that are no state name: {left!r}"))
            else:
                why = []
                for spec in specs_for(kind, det, skind, codes, qual, named, annex_by, why, groups, applied_plain):
                    emit(rows, spec, comm if comm[0] == "odtp" else ("num", comm[1]), dates)
                not_written += [(r, w) for w in why]
    for (lo, hi), a in sorted(annex_by.items()):            # the GCV annexure: the preferred RTOs of excluded states, a payout of their own
        group = ((groups or {}).get("gcv") or {}).get((lo, hi))
        n = sum(len(v) for v in a["states"].values())
        log(f"Sub Annexure {a['no'] or 3}, {lo}-{hi} kg: {a['commission']}% for the preferred RTOs ({', '.join(f'{s} {len(v)}' for s, v in a['states'].items())}; {n} RTOs read)")
        if group is None:
            warnings.append(f"Sub Annexure {a['no'] or 3} has a table for {lo}-{hi} kg but no portal RTO group is known for it for this circular "
                            f"(ANNEX3_GROUPS, CIRCULAR_GROUPS) - its payout is NOT in the output")
        elif a["commission"] is None:
            warnings.append(f"Sub Annexure {a['no'] or 3} {lo}-{hi} kg: no commission % found - its payout is NOT in the output")
        else:
            emit(rows, dict(sc=SC_FULL, cc=cc_band(lo, hi), fuels=list(FUELS_ALL), subs=list(STD_SUBS), tnc="", **{"class": "gcv4w", "rto": group[0]}),
                 ("num", a["commission"]), dates)
    seven, overlaps = set(), []
    if cars:                                                # Private Car: the table's segments, then Sub Annexure 4
        car_why = []
        for spec, rate, car_dates in car_specs(cars, annex4, from_date, to_date, ask, car_why, groups):
            emit(rows, spec, rate, car_dates)
        not_written += car_why
        seven = next((car_segment(r.segment)[2] for r in cars if (car_segment(r.segment) or ("",))[0] == "s2"), set())
    if cars_rev:                                            # Private Car, August layout: cc segment x states, then the 40% RTOs of the annexure
        car_why = []
        for spec, rate, car_dates in car_specs_revised(cars_rev, annex4, groups, from_date, to_date, ask, car_why, overlaps):
            emit(rows, spec, rate, car_dates)
        not_written += car_why
    if not rows:
        raise GridError("nothing to write - no payout of the PDF could be turned into rows")
    df = pd.DataFrame(rows)
    df["id"] = ROW_ID
    for text, spec in applied_plain:                        # a payout whose annexure is not in the document: written for the whole state(s), flagged on exactly those rows
        warnings.append(f"applied without its annexure - {text}")
        review.flag(_short(text), lambda d, sp=spec: np.array((d["_product"] == CLASSES[sp["class"]]["product"]) & (d["new_rto_list"] == sp["rto"])
                                                              & (d["cc_min"].astype(float) == sp["cc"][0]) & (d["cc_max"].astype(float) == sp["cc"][1]), dtype=bool))
    for o in overlaps:                                      # a state-wide row that also covers some of the 40% RTOs: the master has no group to keep them apart
        a_no = (annex4 or {}).get("no") or 3
        review.flag(_short(f"Private Car {o['sub']} {o['seg']}, states {o['rto']}: this row covers the whole state(s), including the Sub Annexure {a_no} RTOs of {', '.join(o['states'])} "
                           f"that the 40% rows of {((groups or {}).get('car') or ('the annexure group',))[0]} also cover, so both match there - the portal needs a group of these states without those RTOs "
                           f"(the master has none)"),
                    lambda d, o=o: np.array((d["_product"] == "Private Car") & (d["sub_product"] == o["sub"]) & (d["new_rto_list"] == o["rto"]) & d["fuel_type"].isin(o["fuels"])
                                            & (d["cc_min"].astype(float) == o["cc"][0]) & (d["cc_max"].astype(float) == o["cc"][1]), dtype=bool))
    for r, why in not_written:
        prods = products_of(r.section)
        detail = "".join(f", {x}" for x in ((f"payout {r.comm!r}" if r.comm else ""), (f"states {r.states!r}" if r.states else "")) if x)
        review.flag(_short(f"{why} ({r.where()}{detail})"), (lambda d, p=prods: np.array(d["_product"].isin(p), dtype=bool)) if prods else None)
        warnings.append(f"not written - {why} ({r.where()}{', payout ' + repr(r.comm) if r.comm else ''})")
    masters = load_masters(ask)
    check_names(df, review, masters, log)                   # every RTO / make group written must be in the master lists
    master_checks(df, review, masters, annex_by, annex4, seven, groups)     # ... and hold what the circular says
    review.report(df, log)                                  # fills the remarks column
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


def _norm_rto(s):
    """'HR,KA,RJ' and 'RJ,HR,KA' are the same list of states; anything else (an RTO group name, ALLINDIA) is compared as it is"""
    parts = [p.strip() for p in str(s).split(",")]
    return ",".join(sorted(set(parts))) if parts and all(re.fullmatch(r"[A-Z]{2}", p) for p in parts) else str(s)


def diff_against_reference(df, ref_path, log=say):
    """Compare with an earlier upload CSV. Lists of states are compared as sets; the remarks column is ignored."""
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    mine = to_csv_frame(df)
    missing = [c for c in KEY_COLUMNS if c not in ref.columns]
    if missing:
        raise ValueError(f"the earlier upload has no column(s) {missing} - it does not look like an upload file")
    mine["new_rto_list"], ref["new_rto_list"] = mine["new_rto_list"].map(_norm_rto), ref["new_rto_list"].map(_norm_rto)
    case_only = 0                                           # a make name that differs from the reference only in capitals is the same group: it is counted, not reported as new + removed
    for col in ("new_vehicle_makes", "new_rto_list"):
        known = set(ref[col])
        case_only += int((~mine[col].isin(known) & mine[col].str.lower().isin({x.lower() for x in known})).sum())
        mine[col], ref[col] = mine[col].str.lower(), ref[col].str.lower()
    if case_only:
        log(f"  {case_only} row(s) spell a make or RTO group name differently from the reference only in capitals (compared as the same name; the masters' spelling is written)")

    def key(d):
        d = d.copy()
        for c in ("vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "sc_min", "sc_max", "tariff_min", "tariff_max"):
            d[c] = pd.to_numeric(d[c], errors="coerce").map(lambda x: f"{x:.4f}")
        return d[KEY_COLUMNS].astype(str).agg("|".join, axis=1)

    mine["_k"], ref["_k"] = key(mine), key(ref)
    if ref["_k"].duplicated().any():
        log(f"  note: reference has {int(ref['_k'].duplicated().sum())} duplicate keys")
    merged = mine.merge(ref.drop_duplicates("_k"), on="_k", how="outer", suffixes=("", "_ref"), indicator=True)
    only_new = merged[merged["_merge"] == "left_only"]
    only_ref = merged[merged["_merge"] == "right_only"]
    both = merged[merged["_merge"] == "both"]
    changed = pd.Series(False, index=both.index)
    detail, masks = {}, {}
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
            detail[c], masks[c] = int(d.sum()), d
        changed |= d
    log(f"  vs {os.path.basename(ref_path)}: {len(only_new)} new rows, {len(only_ref)} removed rows, "
        f"{int(changed.sum())} rows with changed values (of {len(both)} matching keys)")
    for c, n in detail.items():
        ex = both[masks[c]][["_k", c, c + "_ref"]].head(3).values.tolist()
        log(f"    column {c}: {n} differ, e.g. {ex}")
    for label, frame, col in (("new rows", only_new, "product"), ("removed rows", only_ref, "product_ref" if "product_ref" in only_ref else "product")):
        if len(frame):
            by = frame.groupby(col).size()
            log(f"    {label} by product: " + "; ".join(f"{n} x {p}" for p, n in by.items()))
    return len(only_new), len(only_ref), int(changed.sum())


def report(df, warnings, log=say, ask=None):
    """Print what is about to be written; rows that repeat a key keep the lowest payin (resolve_duplicates). Returns the frame to write."""
    ask = ask or Asker()
    df = resolve_duplicates(df, KEY_COLUMNS, ask)
    log(f"\nRows to write: {len(df):,}   ({df['product'].nunique()} products, fuels {', '.join(sorted(df['fuel_type'].unique()))})")
    for (prod, cls, sub), n in df.groupby(["product", "make_model_class", "sub_product"]).size().items():
        log(f"  {prod:<8} {cls:<28} {sub:<5} {n:>6,}")
    log(f"Duplicate keys: {int(df.duplicated(KEY_COLUMNS).sum())}")
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
    po = df[["po_od_rate", "po_tp_rate", "po_net_rate"]].max(axis=1)
    log(f"PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {int(((pi > 0) & (po == 0)).sum())}")
    for w in warnings:
        log(f"WARNING: {w}")
    return df


def main(pdf_path, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None):
    ask = ask or Asker()
    df, warnings = build(pdf_path, from_date, to_date, log, ask)
    df = report(df, warnings, log, ask)
    diff = compare_safely(ask, diff_against_reference, df, ref_path, log)
    ask.summary(log)
    if ask.interactive:
        notes = []
        if warnings:
            notes.append("some payouts of the PDF are not in the output (see WARNING above)")
        if diff and (diff[0] + diff[1]) > 0.25 * len(df):
            notes.append(f"{diff[0]} rows are new and {diff[1]} rows are gone compared with the earlier upload")
        question = ("Heads-up: " + "; ".join(notes) + ". Write the file anyway?" if notes
                    else "Everything look right? Write the file now?")
        if not ask.confirm(question, True):
            raise GridError("stopped by you - nothing was written")
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"UNITED_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
    log(f"\nWrote {out_path}")
    return out_path


# %% ---------------------------------------------------------------- ENTRY POINTS
def collect_uploads(ask):
    """Colab: upload the files; asks again if nothing usable came up. Returns (PDF, earlier upload or None)."""
    while True:
        uploaded = colab_upload(ask)
        pdfs, csvs, ignored = [], [], []
        for name, data in (uploaded or {}).items():
            try:
                with open(name, "wb") as fh:
                    fh.write(data)
            except OSError as e:
                raise GridError(f"cannot save the uploaded file {name}: {e}")
            low = name.lower()
            (csvs if low.endswith(".csv") else pdfs if low.endswith(PDF_EXTS) else ignored).append(name)
        if ignored:
            say(f"Ignored (neither a PDF nor a CSV): {ignored}")
        if pdfs:
            break
        if not ask.confirm("No PDF came with that upload. Upload again?", True):
            raise GridError("no PDF was uploaded")
    pdf = pdfs[0] if len(pdfs) == 1 else ask.choose(
        "more than one PDF was uploaded", [(n, n) for n in pdfs], question="You uploaded more than one PDF. Which one is the United commission structure?", short="PDF")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return pdf, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the United commission structure PDF and, optionally, an earlier upload (.csv) to compare with")
    try:
        pdf, ref = collect_uploads(ask)
        say(f"\nFiles received: PDF = {pdf}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(pdf, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("united_commission_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="United commission structure PDF -> United upload CSV")
    ap.add_argument("--grid", help="the United commission PDF - asked for if left out")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    try:
        pdf = a.grid or ask_file(ask, "no PDF given - pass --grid <commission.pdf>",
                                 "Path of the United commission PDF - you can drag the file into this window:", exts=PDF_EXTS)
        pdf = clean_path(pdf)
        if not os.path.isfile(pdf):
            raise GridError(f"PDF not found: {pdf}")
        ref = a.ref
        if ref is None and ask.interactive:
            try:
                ref = ask_file(ask, "", "Earlier upload CSV to compare with (path), or press Enter to skip:", exts=(".csv",), blank_ok=True) or None
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
        main(pdf, from_date, to_date, a.out, ref, ask=ask)
    except (GridError, ValueError) as e:
        sys.exit(f"STOPPED - {e}")
    except KeyboardInterrupt:
        sys.exit("STOPPED - stopped by you. No file was written.")
    except Exception as e:
        saved = crash_report("united_commission_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
    