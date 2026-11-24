"""
Shriram General Insurance broker grid -> Shriram upload CSV
===========================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (NN).csv") from the workbook Shriram circulates to brokers
every month ("Shriram - OCT26 BROKER GRID.xlsx").  Rates come only from that workbook.

What the workbook holds and how it becomes rows (checked against the portal export "rates_export (63).csv", the answer key of October 2026)
---------------------------------------------------------------------------------------------------------------------------------------
* One sheet of rate rows ("GRID OCT' 26"): STATE | PRODUCTS | DIS % | % PO | POLICY TYPE | AGE | UW REMARKS | REMARKS 1 | EFFECTIVE DATE.
  One grid row becomes one or more upload rows:
    PRODUCTS      -> product / make_model_class / make group / fuels / GVW (cc) or seating (sc) band   (see KINDS and gvw_band)
    POLICY TYPE   -> P = package (1+1)   L = STP   P & L = both   P & SA-OD = 1+1 and SAOD   SA-OD = SAOD   Bundle (1+3) = 1+3 (New)
    % PO          -> the PAYIN (pi).  "25" = plain, "15 OD" / "19.5% OD" = on OD, "16 Net", "15 OD + 2.5 TP" = On OD and TP, IRDA = On OD 15.
                     Plain numbers are On Net (On OD for Private Car) for the package, On TP for STP, On OD for SAOD.  po = pi - 7, never below 0.
    DIS %         -> tariff_max = DIS + 0.01 for package rows (STP is always 0 - 100.01).  A higher DIS slab of the SAME key starts where
                     the lower one ends (tariff_min = the lower slab's tariff_max), as in the answer key.
    AGE           -> vehicle_age_min / max: "upto 15 years" = 0 - 15.01, "4 to 20 years" = 3.01 - 20.01, "... & upto 25 Years for STP" = a
                     different band for the STP row.
    UW REMARKS    -> read by READ_RULES: "Except New Mahindra" (an all-makes row + a Mahindra row for old vehicles), NCB cases, tankers,
                     make lists, SMP / Nil-dep conditions, Bangalore / zone / RTO-code lists ...  Whatever a rule does not understand is NEVER
                     guessed at: the rows carry  Manual Review Required: [text]  in the remarks column (see ManualReview).
* RTO scope.  The STATE column names an area (a state, a state pair such as Odisha = OD + OR, Mumbai / ROM / Delhi & NCR ...).  The declined
  RTOs of the sheet "Declined-Excluded" (per product and state) and, where that sheet has nothing for the product and state, the declined
  RTOs of the sheet "PAN India" are taken out of the area (pan-India exemptions: e-rickshaw / e-cart, school and corporate buses).  The upload
  then names the RTO group of the portal's master that holds EXACTLY the remaining RTOs (resolve_rto); a whole state is written as its state
  code(s).  If the master has no such group the row is written for the whole area and flagged  (the group is named in the message so it can
  be created in the portal).
* ADDENDUM (--addendum <workbook>): Shriram sends a monthly addendum with the same columns as the rate sheet ("w.e.f 1st Sep" rows); its rows OVERRIDE grid rows.  A goods-vehicle addendum row replaces the grid rows
  of the same state whose GVW band it covers (together with the other addendum rows of that state and policy type - a band split in two is still a cover), only for the policy types it names (P, L or P & L):
  an addendum "L" row replaces just the STP part of a grid "P & L" row.  A band that overlaps only in part is NOT trimmed: both rows are written and both say so in the remarks.  Rows with no grid
  counterpart ("New Addition") are simply added.  Addendum rows of other products are added and flagged where a grid row of the same state, product and policy type exists.  Row numbers above 10000 in the
  messages are addendum rows (number - 10000).
* STP rows carry no NCB and no zero-dep (is_ncb and is_zd are NA): the grid's "NCB cases" / "Without NCB" (and "Zero dep" / "Without Zero Dep") pair of one P & L row therefore gives one STP entry two payins, and
  whenever grid rows give one entry (the same card key) different payins the LOWEST is written (the user, 27 Oct 2026); every such entry is listed in the sheet 'Lowest rate chosen' of the review workbook (next to the upload CSV).
* Not converted (warnings): the sheet "PCV 6+1 Short-term policy" (the portal has no short-term sub_product), hazardous-category payout
  reductions, declined vehicle lists (they live in the make groups) and every other rule in the text that the rules cannot read.

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a sheet or header
that is not where it should be, dates that look wrong, a file that will not open or write - it stops, asks you a specific question (type the
number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every answer is listed at the end of the run,
and you confirm before the file is written.  It never asks you for a rate.  With --no-prompt (or when there is no keyboard) it asks nothing
and stops with the question as the error message instead.  If something goes wrong that it did not foresee it says so in one line and saves
the technical details to a *_error_*.txt file.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run shriram_broker_grid_to_config.py`) and run it.  Upload the Shriram workbook (and,
optionally, an earlier upload CSV to get a diff report).  It then asks for from_date and to_date (DD/MM/YYYY), writes them into every row,
and downloads SHRIRAM_UPLOAD_<yyyy-mm>.csv.  The master workbooks (RTO / make-model) must be in a `master` folder next to the script.

Command line
------------
    python shriram_broker_grid_to_config.py [--grid <workbook.xlsx>] [--addendum <addendum.xlsx>] [--from-date 01/10/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the workbook, the dates, the earlier upload - is asked for)
"""
import argparse
import datetime as dt
import difflib
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
INSURANCE_COMPANY = "SHRIRAM GENERAL INSURANCE COMPANY LTD"
USER_ID = 2966
ROW_ID = 1                                       # the export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 2
IRDA_RATE = 15.0                                 # a % PO cell that says IRDA = the IRDA cap on OD
CC_OPEN_END = 99999
AGE_OPEN_END = 99
TARIFF_STP = (0, 100.01)                         # STP rows carry no discount band
HEADER_SCAN_ROWS = 12
MAX_NOTE_CHARS = 500

# sheet names (matched ignoring capitals, spaces and punctuation; the rate sheet is found by its header when the name changed)
GRID_SHEETS = ["GRID OCT' 26", "GRID"]
DECLINED_SHEET = "Declined-Excluded"
PAN_SHEET = "PAN India"
# the side tables are found by their column headers, not by the sheet's name or the columns' places (shared helpers `find_tables` / `pick_table`)
DECLINED_FIELDS = {"product": ["product"], "state": ["rto state", "state"], "code": ["rto code", "rto codes", "code"], "name": ["rto name", "city"], "remarks": ["remarks", "remark"]}
DECLINED_REQUIRED = ("product", "code")
PAN_FIELDS = {"code": ["declined rto", "decline rto", "declined rto code"], "city": ["decline rto city", "declined rto city", "city"], "state": ["state"], "product": ["product"]}
PAN_REQUIRED = ("code", "product")
METRO_SHEET = "Metro RTO Codes"
SHORT_TERM_SHEET = "PCV 6+1 Short-term policy"

# ---- areas: the STATE column -> the RTOs it stands for.  A piece is ("codes", letters) = every RTO of the master's ALLINDIA list whose code starts
# with one of the letters, or ("group", master RTO group).  A multi-state label is written as a comma list of its pieces (GJ,DD,DN).
def _L(*letters):
    return ("codes", tuple(letters))


def _G(name):
    return ("group", name)


STATE_AREAS = {
    "ANDHRA PRADESH": [_L("AP")],
    "ASSAM/TRIPURA/ARUNACHAL/MEGHALAYA/ MIZORAM": [_L("AS"), _L("TR"), _L("AR"), _L("ML"), _L("MZ")],
    "ASSAM/TRIPURA": [_L("AS"), _L("TR")],
    "ASSAM/ARUNACHAL/MEGHALAYA/ MIZORAM": [_L("AS"), _L("AR"), _L("ML"), _L("MZ")],
    "TRIPURA": [_L("TR")],
    "BIHAR": [_L("BR")],
    "CHATTISGARH": [_L("CG")],
    "DELHI&NCR": [_G("SHRIRAM_JUNE26_Delhi_NCR")],
    "GOA": [_L("GA")],
    "GUJARAT & DADRA NAGAR HAVELI & DAMAN & DIU": [_L("GJ"), _L("DD"), _L("DN")],
    "HARYANA": [_L("HR")],
    "HIMACHAL PRADESH": [_L("HP")],
    "J & K": [_L("JK")],
    "JHARKHAND": [_L("JH")],
    "KARNATAKA": [_L("KA")],
    "KERALA": [_L("KL")],
    "KERELA": [_L("KL")],
    "MADHYA PRADESH": [_L("MP")],
    "MUMBAI": [_G("SHRIRAM_JUNE26_Mumbai")],
    "MUMBAI & GOA": [_G("SHRIRAM_JUNE26_Mumbai"), _L("GA")],
    "ODISHA": [_L("OD", "OR")],
    "PUNJAB/CHANDIGARH": [_L("PB"), _L("CH")],
    "RAJASTHAN": [_L("RJ")],
    "ROM": [_G("SHRIRAM_JUNE26_ROM")],
    "TAMILNADU & PONDICHERRY": [_L("TN"), _L("PY")],
    "TELANGANA": [_L("TG", "TS")],
    "UTTAR PRADESH": [_L("UP")],
    "UTTARANCHAL-AA": [_G("SHRIRAM_JUNE26_UTTARANCHAL-AA")],
    "UTTARANCHAL-KS": [_G("SHRIRAM_JUNE26_UTTARANCHAL-RSD")],          # the grid says KS, the portal's group is called RSD (same RTOs, per the answer key)
    "WEST BENGAL": [_L("WB")],
}
ALL_INDIA = "ALLINDIA"                           # the master group of every RTO; the PAN INDIA rows use it
# partitions named in the remarks (the master's groups); a name that the master lacks is flagged by check_names
NCR_GROUP = "SHRIRAM_JUNE26_Delhi_NCR"
BANGALORE_GROUP = "SHRIRAM_JUNE26_KA_BANGLORE"
TN_METRO_GROUP = "SHRIRAM_JUNE26_TN_METRO"
TN_NON_METRO_GROUP = "SHRIRAM_JUNE26_TN_NON_METRO"
CHENNAI_GROUP = "SGIC_CHENNAI_OCT26"
RJ_ZONE_GROUPS = {1: "SHRIRAM_JUNE26_RJ_ZONE1", 2: "SHRIRAM_JUNE26_RJ_ZONE2", 3: "SHRIRAM_JUNE26_RJ_ZONE3"}
METRO_GROUP = "SHRIRAM_AUG26_METRO_RTO"          # corporate bus: metro / non-metro RTOs (the sheet "Metro RTO Codes" lists the metro ones)
NON_METRO_GROUP = "SHRIRAM_AUG26_NON-METRO_RTO"
# among master groups that hold exactly the same RTOs, the one whose name starts with the first matching prefix is written
RTO_NAME_PREFERENCE = ("SHRIRAM_JUNE26", "SHRIRAM", "SGIC")
RTO_NAME_AVOID = ("_TW", "HDFC", "ICICI", "LIBERTY", "MAGMA", "ROYAL", "RELIANCE", "CHOLA", "GODIGIT", "BAJAJ")

# ---- make groups of the portal (new_vehicle_makes)
MAKES_PC = "SHRIRAM_PC_EX_DECLINE_JUNE26"
MAKES_PC_BUNDLE = "HONDA,Hyundai,KIA"
MAKES_TAXI_06 = "SHRIRAM_PCV_4W_0-7_JUNE26"
MAKES_TAXI_710 = "SHRIRAM_PCV_4W_7-10_JUNE26"
MAKES_TAXI_MSHB = "SHRIRAM_Maruti_Suzuki_Hyundai_M&M_Bolero"
MAKES_TAXI_EX_MSHB = "SHRIRAM_TAXI_EX_Maruti_Suzuki_Hyundai_M&M_Bolero"
MAKES_BUS = "BUS_ALL"
MAKES_3W_GCV = "GCV_3W_ALL"
MAKES_3W_PCV = "PCV_3W_ALL"
MAKES_3W_BOTH = "GCV_3W_ALL,PCV_3W_ALL"
MAKES_MISCD = "SHRIRAM_MISD_EX_DECLINE_JUNE26"
GCV_BASE = "SHRIRAM_GCV_4W_JUNE26_EX_Bharat_Benz_MAN_Truck_AMW_Volvo"      # every GCV make except the ones the insurer declines
MAKES_BOLERO = "GCV_4W_BOLERO"                   # 'Bolero Model only' (the user's choice in November 2026; October used SGIC_CV_BOLERO, the master holds both)
MAKES_MAHINDRA_EX_BOLERO = "Mahindra_ex_Bolero"  # Mahindra without the Bolero models
# A row that prices ONE make on its own ('Only EECO' for Gujarat's cars, 'Bolero Model only' for Andhra Pradesh's goods vehicles) takes that make out of the other rows of the same state and
# product that would also cover it: the group those rows write -> the same group without the make.  The user did this by hand in the November 2026 upload (family, make) -> {group: group}.
CARVE_GROUPS = {("pc", "EECO"): {MAKES_PC: "SHRIRAM_PC_EX_Eecco_DECLINE_JUNE26"},
                ("gcv", "BOLERO"): {"Mahindra": MAKES_MAHINDRA_EX_BOLERO}}
# the groups that also leave out TATA / ASHOK / MAHINDRA: excluded set -> master name
GCV_EXCLUSIONS = {
    frozenset(): GCV_BASE,
    frozenset({"MAHINDRA"}): GCV_BASE + "_MAHINDRA",
    frozenset({"MAHINDRA", "ASHOK"}): GCV_BASE + "_MAHINDRA_ASHOK",
    frozenset({"TATA", "ASHOK"}): GCV_BASE + "_TATA_ASHOK",
    frozenset({"TATA", "ASHOK", "MAHINDRA"}): GCV_BASE + "_TATA_ASHOK_MAHINDRA",
}
# the grid's make -> the master's spelling (the master is the authority: 'Ashok', 'Mahindra', 'Hyundai', 'Maruti')
MAKE_TOKENS = {"TATA": "TATA", "ASHOK": "Ashok", "MAHINDRA": "Mahindra", "SML ISUZU": "SML_ISUZU", "BAJAJ": "BAJAJ", "TVS": "TVS",
               "HONDA": "HONDA", "HYUNDAI": "Hyundai", "KIA": "KIA", "CITROEN": "CITROEN", "MARUTI": "Maruti"}
GCV_CLASSES_OTHER = ["Dumper Tipper", "Goods Carrying Tractor", "Pick Up Van", "Trailer", "Truck"]      # every class except Tanker
TAXI_TNC = "Self Driven Taxi are decline"                              # the T&C of every taxi row ('Declined vehicle' sheet: Self Driven Car and Taxi are decline)

# ---- fuels
FUELS_ALL5 = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NON_DIESEL = ["Petrol", "CNG", "LPG", "Electric"]                # "non-diesel = petrol": Electric is written with them
FUELS_NO_EV = ["Petrol", "Diesel", "CNG", "LPG"]

# ---- notes already known not to change the rows of this upload (matched when a remark CONTAINS the phrase, capitals / spaces / punctuation ignored)
STANDING_NOTES = [
    "Additional Slab", "New Slab Addtion", "No change", "Battery watt condition removed", "GVW criteria improved", "uw's guideline revised",
    "Net PO allowed", "ALL (EXCEPT - COMMANDER",
    # the user cleared these in the November 2026 upload without changing a row: an underwriting instruction the rate card cannot hold; the age and the RTO area the row already says
    "theft excess in E-cart and E-Rickshaw", "(NCR)(OR BASED DL)",
]
# the change notes of the column REMARKS 1 ("PO reduced by 2.5%", "PO increased by 9%" ...): they describe last month's change, the cell already holds the new payout
CHANGE_NOTE = re.compile(r"^\s*PO\s+(?:was\s+)?(?:reduced|decreased|increased|incresed|increase|decrease)\b[^|]*$", re.I)
# the words of a note that only describes a change of the grid (is_change_note); the verbs include the typos seen in the grids
CHANGE_WORDS = frozenset("po payout disc dic discount tp od by and age gvw uw uws guideline guidelines criteria slab new".split())
CHANGE_VERBS = frozenset("increased increase incresed increaed inceased reduced reduce redcued decreased decrease improved improve changed change revised revise removed added addition addtion".split())
# "For Hazardous category PO will be 5% less" (141 grid rows, goods vehicles). The hazardous category of a policy is not in the MIS and add_tnc does not take part in
# the matching, so the only thing that tells the portal a hazardous vehicle apart is its class:
#   "tanker_rows" = the Tanker class gets the payin HAZARDOUS_CUT points lower on the rows whose own remark already separates tankers ("Tankers", "All Tankers ...");
#                   the other rows that mention it are left as listed and counted in a warning
#   "tanker"      = as above, and every other row that mentions it is split into a Tanker row (payin lower) and the five other classes (the user asked on 26 Oct 2026 to apply
#                   the note wherever it is mentioned: this takes the upload from about 6,700 to 16,400 rows and makes 195 keys contradict each other)
#   "flag"        = not applied, every row that mentions it carries it in the remarks column (the user's working sheet does not reduce it either; "need to apply this if mentioned")
#   "off"         = not applied, one warning
#   "tnc"         = the text 'Hazardous category PO N% less' (N from the grid's own note) goes into add_tnc of every row that carries the note, no flag, the payouts as listed
#                   (the user, 31 Oct 2026: "add it in add_tnc as a condition")
HAZARDOUS_MODE = "tnc"
HAZARDOUS_CUT = 5.0                              # "5% less" read as 5 percentage points of payin (like "PO reduced by 2.5%" elsewhere in the grid)
# a row identical (key and payin) to a row of the earlier upload given with --ref keeps that row's from_date ("the grids of September and October were the same")
KEEP_REF_DATES = True
ADDENDUM_ROW_BASE = 10000                        # addendum rows are numbered 10000 + their Excel row in messages and in the contradicting-payout notes

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")

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


OUTPUT_FOLDER = "shriram"            # the folder of this insurer under output\ (default_output)


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
class EmptySheet(GridError):
    """A sheet that cannot be used at all (no header, no rate rows): the caller leaves it out or asks."""


def _text(v):
    return "" if v is None else str(v).replace("\xa0", " ").replace("​", "").strip()


def norm(v):
    """A cell as one line of text: no line breaks, no double spaces, no non-breaking spaces."""
    return re.sub(r"\s+", " ", _text(v)).strip()


def _col_letter(idx):
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx + 1)


def read_workbook(path):
    """Open with values only. NOT read-only mode: the sheets have merged cells and only the normal mode keeps them intact."""
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def sheet_rows(wb, name):
    return [list(r) for r in wb[name].iter_rows(values_only=True)]


def label_key(s):
    """A label without spaces, punctuation and capitals: 'NCB Status' and 'ncb  status' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def num(v):
    """A cell as a float, or None when it is not a plain number ('APBG', 'WITHDRAW', blank)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*%?\s*", _text(v))
    return float(m.group(1)) if m else None


def round_half_up(x, places=RATE_DECIMALS):
    q = 10 ** places
    return math.floor(float(x) * q + 0.5 + 1e-9) / q


GRID_FIELDS = {          # field -> header label_keys that name it
    "state": ("state",), "product": ("products", "product"), "dis": ("dis", "dis%", "discount"), "po": ("po", "payout", "popercent"),
    "ptype": ("policytype",), "age": ("age",), "uw": ("uwremarks", "uwremark"), "rem1": ("remarks1", "remark1"), "eff": ("effectivedate", "effdate"),
    "rto": WORKING_HEADERS,          # the person's RTO working list (31 Oct 2026): optional, written exactly as typed
}
GRID_REQUIRED = ("state", "product", "po", "ptype", "age")


TWIN_PAYIN_FIELDS = ("po",)       # a payin column that is there twice under one heading: the LOWEST payin of the twins is written (the user, 12 Nov 2026); any other field twice stops the run
TWIN_ENTRIES = []                  # grid rows whose twin payin columns differ: the sheet 'Twin columns' of the review workbook (registered by build after ManualReview())


class TwinCols(dict):
    """{field: column} of a header row, plus `.twins` = {field: [every column with the SAME heading]} for the payin fields that are there more than once"""
    twins = {}


def find_header(rows, fields=GRID_FIELDS, required=GRID_REQUIRED):
    """(header row index, {field: column}) of the first row (within HEADER_SCAN_ROWS) that names every required field, else None."""
    for i, r in enumerate(rows[:HEADER_SCAN_ROWS]):
        keys = [label_key(c) for c in r]
        cols, twins = {}, {}
        for f, names in fields.items():
            for j, k in enumerate(keys):
                if k and k in names:
                    if f not in cols:
                        cols[f] = j
                    elif k == keys[cols[f]]:                # the SAME heading twice (a copy beside the original); a differently worded second column (the person's own mapping columns) is not
                        twins.setdefault(f, [cols[f]]).append(j)
        if all(f in cols for f in required):
            hard = {f: js for f, js in twins.items() if f not in TWIN_PAYIN_FIELDS}
            if hard:                                        # the same heading twice for something that is no payin (DIS %, STATE ...): there is no 'lowest' of two discount bands - which is meant is not in the sheet
                from openpyxl.utils import get_column_letter
                f, js = sorted(hard.items())[0]
                raise GridError(f"the header row {i + 1} has {len(js)} columns with the heading '{f}' ({', '.join(repr(_text(r[j])) + ' in column ' + get_column_letter(j + 1) for j in js)}) - expected one; "
                                f"which of them is meant is not in the sheet, so nothing is written (delete or rename the one that is not this month's, and run again)")
            cols = TwinCols(cols)
            cols.twins = twins                              # payin fields only (the others stopped above)
            return i, cols
    return None


def find_grid_sheet(wb, ask, log=say):
    """The rate sheet: a sheet called like GRID_SHEETS or "GRID <month> ' <yy>" (the name changes every month: GRID OCT' 26, GRID SEP' 26), else any sheet whose header
    names STATE / PRODUCTS / % PO / POLICY TYPE / AGE.  The short-term policy sheet has the same header and is never the rate sheet."""
    wanted = {label_key(n) for n in GRID_SHEETS}
    named = lambda n: label_key(n) in wanted or label_key(n).startswith("grid")
    order = sorted(wb.sheetnames, key=lambda n: (not named(n), wb.sheetnames.index(n)))
    found = []
    for name in order:
        if label_key(name) == label_key(SHORT_TERM_SHEET):
            continue
        h = find_header(sheet_rows(wb, name))
        if h:
            found.append((name, h))
    if not found:
        raise GridError("no sheet of the workbook has the rate-table header (STATE | PRODUCTS | DIS % | % PO | POLICY TYPE | AGE) - is this the Shriram broker grid?")
    if len(found) > 1 and not named(found[0][0]):
        pick = ask.choose("more than one sheet looks like the rate table", [(n, n) for n, _ in found],
                          question="More than one sheet has the rate-table header. Which one holds the Shriram rates?", short="rate sheet")
        found = [f for f in found if f[0] == pick]
    return found[0]


def read_grid(wb, ask, log=say, working=True):
    """-> (rows, warnings). One dict per rate row: r (Excel row), state, product, dis, po, ptype, age, uw, rem1, eff (all text, blanks = ''), rto (the person's working-list name, '' = none:
    `working` False = not read, the addendum)."""
    name, (hdr, cols) = find_grid_sheet(wb, ask, log)
    rows = sheet_rows(wb, name)
    out, warnings = [], []
    twins = getattr(cols, "twins", {})
    for f, js in twins.items():
        from openpyxl.utils import get_column_letter
        warnings.append(f"'{name}': the payout column '{_text(rows[hdr][js[0]])}' is there {len(js)} times (columns {', '.join(get_column_letter(j + 1) for j in js)}): "
                        f"the LOWEST payin of them is written for each entry (the rows where they differ are listed in the sheet 'Twin columns')")
    if not working:
        cols = {f: c for f, c in cols.items() if f != "rto"}
    for i in range(hdr + 1, len(rows)):
        r = rows[i]
        get = lambda f: r[cols[f]] if f in cols and cols[f] < len(r) else None
        pos = [get("po")]
        if twins.get("po"):                                 # a payin column there twice: every twin's value is a candidate; a blank twin cell is no payin
            pos = [r[j] if j < len(r) else None for j in twins["po"]]
            if any(_text(v) != "" for v in pos):
                pos = [v for v in pos if _text(v) != ""]
            plain = [v for v in pos if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if len(pos) > 1 and len(plain) == len(pos):     # plain numbers: the lowest of the twins is the payin of this grid row
                if len(set(plain)) > 1:
                    TWIN_ENTRIES.append(dict(sheet=name, excel_row=i + 1, payins=" / ".join(str(v) for v in plain), written=min(plain)))
                pos = [min(plain)]
        if all(_text(get(f)) == "" for f in ("state", "product", "ptype")) and _text(pos[0]) == "":
            continue
        if _text(get("state")) == "" or _text(get("product")) == "":
            warnings.append(f"'{name}' row {i + 1} has a payout but no state or no product - left out: {[_text(x) for x in r[:6]]}")
            continue
        for po in pos:
            out.append(dict(r=i + 1, sheet=name, state=norm(get("state")), product=norm(get("product")), dis=get("dis"), po=po,
                            ptype=norm(get("ptype")), age=norm(get("age")), uw=norm(get("uw")), rem1=norm(get("rem1")), eff=norm(get("eff")), rto=working_value(get("rto"))))
    if working:
        has = [g for g in out if g["rto"]]
        working_table(name, _text(rows[hdr][cols["rto"]]) if "rto" in cols else None, len(has), [g["r"] for g in out if not g["rto"]] if "rto" in cols else [])
    for f in ("dis", "uw", "rem1", "eff"):
        if f not in cols:
            warnings.append(f"'{name}' has no '{f}' column (looked for {GRID_FIELDS[f]}) - what it carried is not in the output")
    log(f"'{name}': {len(out)} rate rows (header on row {hdr + 1}, columns {sorted(cols)})")
    return out, warnings


def read_addendum(path, ask, log=say):
    """The addendum workbook: the same columns as the rate sheet (its extra RTO / MAKE mapping columns are ignored). -> (rows, warnings); the rows are numbered ADDENDUM_ROW_BASE + Excel row."""
    _, wb = open_file(ask, path, "the Shriram addendum", read_workbook, EXCEL_EXTS)
    rows, warnings = read_grid(wb, ask, log, working=False)
    for g in rows:
        g["r"] += ADDENDUM_ROW_BASE
        g["src"] = "addendum"
    return rows, warnings


def _overlap(a, b):
    return min(a[1], b[1]) - max(a[0], b[0]) > 0.005


def _covers(bands, target, eps=0.011):
    """True when the union of the GVW bands [(lo, hi)] covers the target band (adjacent bands count: 20000.01 meets 20000.01)"""
    cur = target[0]
    for lo, hi in sorted(bands):
        if lo > cur + eps:
            return False
        cur = max(cur, hi)
        if cur >= target[1] - eps:
            return True
    return cur >= target[1] - eps


def apply_addendum(grid, add, warnings, log=say):
    """The addendum rows override the grid rows (see the docstring of the script).  Returns the new row list: the grid rows (some with `drop` = the policy parts an addendum row replaced) + the addendum rows.
    Goods vehicles (kind gcv4w) are matched by state and GVW band; other products are only compared by state, product and policy type, and flagged, never replaced."""
    usable = []
    for a in add:
        kind, band = classify(a["product"])
        a["_kind"], a["_band"], a["_area"], a["_subs"] = kind, band, area_of(a["state"]), policy_subs(a["ptype"]) or []
        if kind is None or a["_area"] is None or not a["_subs"] or parse_po(a["po"]) is None:
            warnings.append(f"addendum row {a['r']} ({a['state']} | {a['product']} | {a['ptype']} | {_text(a['po'])}) is not readable - it replaces nothing and is converted only if the rules read it")
            a["_override"] = False
        else:
            a["_override"] = True
            usable.append(a)
    replaced, partial, touched = [], [], set()
    for g in grid:
        kind, band = classify(g["product"])
        subs, area = policy_subs(g["ptype"]) or [], area_of(g["state"])
        if kind is None or area is None:
            continue
        for part in subs:
            if kind == "gcv4w" and band:
                mates = [a for a in usable if a["_kind"] == "gcv4w" and a["_band"] and a["_area"] == area and part in a["_subs"] and _overlap(a["_band"], band)]
                if not mates:
                    continue
                touched.update(a["r"] for a in mates)
                if part in ("pkg", "stp") and _covers([a["_band"] for a in mates], band):
                    g.setdefault("drop", set()).add(part)
                    replaced.append((g["r"], part, sorted(a["r"] for a in mates)))
                else:
                    rows_txt = ", ".join(str(a["r"]) for a in mates)
                    g.setdefault("flags", []).append(f"the addendum row(s) {rows_txt} overlap this grid row ({part}) only in part - both are written, check the weight band")
                    for a in mates:
                        a.setdefault("flags", []).append(f"this addendum row overlaps grid row {g['r']} ({part}) only in part - both are written, check the weight band")
                    partial.append((g["r"], part, sorted(a["r"] for a in mates)))
            else:
                mates = [a for a in usable if a["_kind"] == kind and a["_area"] == area and part in a["_subs"] and label_key(a["product"]) == label_key(g["product"])]
                for a in mates:
                    touched.add(a["r"])
                    g.setdefault("flags", []).append(f"the addendum row {a['r']} covers the same state, product and policy type ({part}) - both are written, check which one applies")
                    a.setdefault("flags", []).append(f"this addendum row covers the same state, product and policy type ({part}) as grid row {g['r']} - both are written, check which one applies")
                    partial.append((g["r"], part, [a["r"]]))
    new_rows = [a["r"] for a in add if a["r"] not in touched]
    log(f"Addendum: {len(add)} rows; {len(replaced)} grid row part(s) replaced, {len(partial)} overlap(s) left as they are and flagged, {len(new_rows)} addendum row(s) with no grid counterpart (added).")
    if replaced:
        log("  replaced: " + "; ".join(f"grid row {r} ({part}) by addendum row {', '.join(map(str, ms))}" for r, part, ms in replaced[:12]) + (" ..." if len(replaced) > 12 else ""))
    if partial:
        warnings.append(f"{len(partial)} grid row part(s) overlap an addendum row only in part (not trimmed, both written and flagged): "
                        + "; ".join(f"grid {r} ({part}) / addendum {', '.join(map(str, ms))}" for r, part, ms in partial[:8]) + (" ..." if len(partial) > 8 else ""))
    return list(grid) + list(add)


# ---- the decline lists
CODE_RX = re.compile(r"\b([A-Z]{2})\s*[-_]?\s*(\d{1,2})\b")                # 'HR-19', 'SK01', 'TN_85'
PAIR_RX = re.compile(r"\b([A-Z]{2})\s*/\s*([A-Z]{2})\s*-?\s*(\d{1,2})\b")


def codes_in(text):
    """RTO codes named in a cell: 'HR-19,HR-84' -> {HR19, HR84}; 'TS/TG-05' -> {TS05, TG05}; 'SK01' -> {SK01}."""
    t = _text(text).upper()
    out = set()
    for a, b, n in PAIR_RX.findall(t):
        out |= {a + n.zfill(2), b + n.zfill(2)}
    t = PAIR_RX.sub(" ", t)
    for a, n in CODE_RX.findall(t):
        out.add(a + n.zfill(2))
    return out


def _has_codes(col):
    """check for `find_tables`: the column the header says holds RTO codes really holds them (half of its cells, at least)"""
    def check(body, cols):
        cells = [r[cols[col]] for r in body[:60] if len(r) > cols[col] and _text(r[cols[col]])]
        return bool(cells) and sum(1 for c in cells if codes_in(c)) * 2 >= len(cells)
    return check


class Declines:
    """The sheets 'Declined-Excluded' (product x state x RTO) and 'PAN India' (RTOs declined for everything except the exempt products).
    decline(labels, letters, exempt) -> the RTOs taken out of an area for a product: the sheet's entries for that product and state when it has
    any, else the pan-India RTOs of the state (unless the product is exempt); a whole state declined pan-India has no RTO left at all."""

    def __init__(self):
        self.entries = []            # (label_key, state text, codes, remark)
        self.pan = {}                # letters -> codes (the 'Pan India Except (...)' block)
        self.whole = set()           # letters of the states declined for everything ('Pan India')
        self.exempt_text = ""
        self.found = False
        self.declined_sheet, self.pan_sheet = DECLINED_SHEET, PAN_SHEET        # where the tables were really found (read)
        self.last_sheets = set()     # the sheets the last decline() took RTOs from (for the warning text)
        self.labels = {}             # label_key -> product text, to warn about a product the rules do not know

    def _declined_by_values(self, wb, warnings):
        """The declined-RTO table of a sheet called like 'Declined-Excluded' whose header does not say 'RTO Code' (a deleted or reworded header): the columns are inferred from the VALUES - the
        column whose cells are RTO codes, the text columns to its left (product, state), the remarks two columns after it - and the run says so in the warnings."""
        name = sheet_like(wb, DECLINED_SHEET, ("declined", "excluded"))
        if name is None:
            return None
        rows = sheet_rows(wb, name)
        best = None
        for j in range(max((len(r) for r in rows), default=0)):
            cells = [r[j] for r in rows if len(r) > j and _text(r[j])]
            n = sum(1 for c in cells if codes_in(c))
            if n and n * 2 >= len(cells) and (best is None or n > best[1]):
                best = (j, n)
        if best is None:
            return None
        code = best[0]
        texts = [j for j in range(code) if any(len(r) > j and _text(r[j]) and not codes_in(r[j]) for r in rows)]
        if not texts:
            return None
        cols = {"product": texts[0], "code": code, "remarks": code + 2}
        if code - 1 > texts[0]:
            cols["state"] = code - 1
        head = next((i for i, r in enumerate(rows[:15]) if any(head_match(c, "product") for c in r)), 0)
        warnings.append(f"the table of the declined RTOs on '{name}' has no 'RTO Code' header: its columns were inferred from the values (product, state, RTO code, RTO name, remarks)")
        return (name, head, cols, rows[head + 1:])

    def read(self, wb, warnings, ask=None, log=say):
        """Reads the two tables by their headers wherever they are (any sheet, any column, any title rows above): 'Declined-Excluded' says Product / RTO State / RTO Code /
        RTO Name / Remarks, 'PAN India' says Declined RTO / Decline RTO City / State / Product. A table that is not there is reported, never skipped silently."""
        hit = pick_table(find_tables(wb, DECLINED_FIELDS, DECLINED_REQUIRED, _has_codes("code")), DECLINED_SHEET, "the declined RTOs (product, state, RTO code)", ask, warnings, log)
        if hit is None:
            hit = self._declined_by_values(wb, warnings)
        used = set()
        if hit:
            self.found = True
            self.declined_sheet = hit[0]
            used.add(hit[0])
            cols = hit[2]
            at = lambda r, k: (r[cols[k]] if k in cols and len(r) > cols[k] else None)
            for r in hit[3]:
                if _text(at(r, "product")) and _text(at(r, "code")) and label_key(at(r, "product")) != "product":
                    codes = codes_in(at(r, "code"))
                    if codes:
                        extra = r[cols["remarks"]:cols["remarks"] + 2] if "remarks" in cols else []
                        self.entries.append((label_key(at(r, "product")), norm(at(r, "state") or ""), codes, " ".join(norm(x) for x in extra if _text(x))))
                        self.labels.setdefault(label_key(at(r, "product")), norm(at(r, "product")))
        else:
            warnings.append(f"no sheet with the declined RTOs (a table headed 'Product' and 'RTO Code'; usually '{DECLINED_SHEET}') - declined RTOs are not taken out of any area")
        hit = pick_table(find_tables(wb, PAN_FIELDS, PAN_REQUIRED, _has_codes("code"), exclude=used), PAN_SHEET, "the RTOs declined for every product (declined RTO, product)", ask, warnings, log)
        if hit:
            self.pan_sheet = hit[0]
            cols = hit[2]
            product = ""
            for r in hit[3]:
                if len(r) > cols["product"] and _text(r[cols["product"]]):
                    product = norm(r[cols["product"]])
                if not r or len(r) <= cols["code"] or not _text(r[cols["code"]]):
                    continue
                codes = codes_in(r[cols["code"]])
                if not codes:
                    continue
                if re.match(r"pan\s*india\s*except", product, re.I):
                    self.exempt_text = self.exempt_text or product
                    for c in codes:
                        self.pan.setdefault(c[:2], set()).add(c)
                elif re.fullmatch(r"pan\s*india", product, re.I):
                    self.whole |= {c[:2] for c in codes}
        else:
            warnings.append(f"no sheet with the pan-India declined RTOs (a table headed 'Declined RTO' and 'Product'; usually '{PAN_SHEET}') - they are not taken out of any area")

    def decline(self, labels, letters, exempt=False):
        out = set()
        self.last_sheets = set()
        for L in letters:
            if L in self.whole:
                out |= {"*" + L}                                    # marker: the whole state is declined
                self.last_sheets.add(self.pan_sheet)
                continue
            mine = set()
            hit = False
            for lab, _st, codes, _rem in self.entries:
                if lab in labels and any(c.startswith(L) for c in codes):
                    hit = True
                    mine |= {c for c in codes if c.startswith(L)}
            if hit:
                out |= mine
                self.last_sheets.add(self.declined_sheet)
            elif not exempt:
                out |= self.pan.get(L, set())
                if self.pan.get(L):
                    self.last_sheets.add(self.pan_sheet)
        return out

    def remarks_for(self, labels, letters):
        return [(rem, sorted(c for c in codes if c[:2] in letters)) for lab, _st, codes, rem in self.entries
                if lab in labels and rem and not re.fullmatch(r"(declined|excluded)\b[\s\w().,&-]*|\d+(st|nd|rd|th)\s+\w+\s+\d{4}", rem, re.I)]


def metro_codes(wb):
    """Every RTO code listed in the sheet 'Metro RTO Codes' (Bangalore, Chennai, Hyderabad, Coimbatore ... ), or None when the sheet is missing."""
    name = sheet_like(wb, METRO_SHEET, ("metro",))
    if name is None:                                        # renamed: the sheet that says 'metro' in its first rows and is full of RTO codes (every cell is read, so the columns' places do not matter)
        for n in wb.sheetnames:
            rows = sheet_rows(wb, n)
            if find_header(rows):
                continue                                    # the rate sheet itself
            if "metro" in _hwords(" ".join(str(c) for r in rows[:6] for c in r if isinstance(c, str))) and sum(1 for r in rows for c in r if isinstance(c, str) and codes_in(c)) >= 10:
                name = n
                break
    if name is None:
        return None
    out = set()
    for r in sheet_rows(wb, name):
        for c in r:
            if isinstance(c, str):
                out |= codes_in(c)
    return out


# 'Decline Product [MISC.D - OTHERS (BODY TYPE GARBAGE CARRIER OR GARBAGE COMPACTOR)' - a line of the sheet 'Declined vehicle' (new in the November 2026 grid; the first build did not read that sheet's
# text at all and said nothing).  A body type is not something the portal matches on, so it is written the way the person already records the bodies a MISC-D others row leaves out: into the exclusion
# text of add_tnc ('EX:Crane,Tractor,Ambulance' -> 'EX:Crane,Tractor,Ambulance,Garbage Carrier,Garbage Compactor'; the person's choice of 4 Nov 2026).  A line this script cannot write that way is flagged.
DECLINE_PRODUCT = re.compile(r"decline\s+product\s*\[\s*(?P<product>[^(\[\]]+?)\s*(?:\((?P<what>[^)\]]*)\)?)?\s*\]?\s*$", re.I)
DECLINE_PRODUCT_ROWS = {"miscd_others": ("MISCD", "MISCD-Others")}            # the kind of `classify` -> (product, make_model_class) of the rows it stands for
DECLINE_BODY = re.compile(r"body\s+type\s+(?P<bodies>.+)$", re.I)


def decline_product_notes(wb):
    """-> [(the line as written, the product label, the body types it names)] of every 'Decline Product [...]' line of the sheet 'Declined vehicle' (found by its name, spaces and capitals ignored);
    the body types are the words after 'BODY TYPE', split at 'OR' / ',' / '&' / 'AND' and written Like This ('GARBAGE  CARRIER OR GARBAGE COMPACTOR' -> Garbage Carrier, Garbage Compactor); [] when there are none"""
    name = sheet_like(wb, "Declined vehicle", ("declined", "vehicle"))
    out = []
    for r in (sheet_rows(wb, name) if name else []):
        for c in r:
            t = norm(c)
            m = DECLINE_PRODUCT.search(t) if t else None
            if m:
                b = DECLINE_BODY.match(norm(m.group("what") or ""))
                bodies = [x.strip().title() for x in re.split(r"\s+or\s+|\s*[,&]\s*|\s+and\s+", b.group("bodies"), flags=re.I) if x.strip()] if b else []
                out.append((t, m.group("product").strip(), bodies))
    return out


def add_declined_bodies(tnc, bodies):
    """The T&C of a MISC-D others row with the declined body types written into its exclusion list.  'EX:Crane,Tractor,Ambulance' gets them at the end of the list, 'EX:Crane,Tractor, Ambulance & other than
    Combined Harvester' before the 'other than' clause, a row with no T&C at all gets 'EX:<bodies>'.  Any other T&C is left as it is: a row for ONE body ('JCB /-', 'Combined Harvester /-') cannot have the
    declined body among its vehicles."""
    if not bodies:
        return tnc
    add = ",".join(bodies)
    if not tnc:
        return "EX:" + add
    m = re.match(r"(?P<head>EX:.*?)(?P<tail>\s+&\s+other\s+than\b.*)?$", tnc, re.I)
    return (m.group("head") + "," + add + (m.group("tail") or "")) if m else tnc


# ---- RTO areas -> master groups
class Rtos:
    """Turns an area (the STATE column) minus the declined RTOs into new_rto_list: the state code(s) when nothing is taken out, otherwise the master
    group that holds EXACTLY the remaining RTOs.  Nothing is invented: if the master has no such group the whole area is written and the caller
    flags the rows (the group is named in the message so it can be created in the portal)."""

    def __init__(self, masters):
        self.m = masters
        self.all = set(masters.rto.get(ALL_INDIA, ()))
        self.gaps = {}                                  # "what is missing in the master" -> the RTO codes the missing group would hold
        self.ncr = masters.codes(NCR_GROUP) - masters.codes("DL")          # the NCR RTOs outside Delhi (UP 14 / 16 / 37, HR 26 / 38 / 51 / 55)
        self.by_content = {}
        for name, codes in masters.rto.items():
            self.by_content.setdefault(frozenset(codes), []).append(name)

    def letters_codes(self, letters):
        """The RTOs of the states with these letters: the master's state group when it has one (ALLINDIA lacks OR and TG), else ALLINDIA's codes."""
        out = set()
        for L in letters:
            got = self.m.codes(L)
            out |= got if got else {c for c in self.all if c[:2] == L}
        return out

    def piece(self, p):
        """(codes, token, letters) of one piece of an area"""
        kind, v = p
        if kind == "codes":
            return self.letters_codes(v), ",".join(v), set(v)
        codes = self.m.codes(v)
        return codes, v, {c[:2] for c in codes}

    def pick(self, codes):
        names = self.by_content.get(frozenset(codes), [])
        names = [n for n in names if n != ALL_INDIA]
        if not names:
            return None

        def rank(n):
            pref = next((i for i, p in enumerate(RTO_NAME_PREFERENCE) if n.upper().startswith(p)), len(RTO_NAME_PREFERENCE))
            return (any(a in n.upper() for a in RTO_NAME_AVOID), pref, len(n), n)
        return sorted(names, key=rank)[0]

    def resolve(self, area, declines, labels, exempt=False, include=None, exclude=None, extra=frozenset(), include_names=(), sheet=""):
        """-> (new_rto_list, [what is missing in the master]).  `include` / `exclude` are sets of RTO codes from the remarks, `extra` more RTOs to take out.
        The name is an RtoScope: when the area is narrower than whole states it also carries, in plain English, the condition and the exact part (`rto_cell` writes
        them next to the cluster name in new_rto_list, with the product and the sub-product; the user's rule of 30 Oct 2026)."""
        if area == ALL_INDIA:
            return RtoScope(ALL_INDIA), []
        tokens, missing, parts, why, whole, warns = [], [], [], [], 0, []
        shown = lambda x: re.sub(RTO_STRIP, "", x)                  # a master group is shown without its insurer / month prefix
        for p in area:
            base, token, letters = self.piece(p)
            if not base:
                tokens.append(token)                                 # a group the master lacks: check_names flags it
                whole += 1
                continue
            gone = set(extra)
            marked = declines.decline(labels, letters, exempt)
            sources = set(getattr(declines, "last_sheets", ()))                # where this piece's RTOs were taken out (the sheets named in a new-mapping warning)
            whole_gone = {m[1:] for m in marked if m.startswith("*")}
            gone |= {m for m in marked if not m.startswith("*")}
            target = {c for c in base if c not in gone and c[:2] not in whole_gone}
            reasons = []
            if any(c in gone or c[:2] in whole_gone for c in base if c not in extra):
                reasons.append("the insurer declines some RTOs of this state for this product so they are left out")
            if extra and base & set(extra):
                reasons.append("the NCR RTOs belong to the Delhi NCR area so they are left out of the other states")
            if include is not None and letters & {c[:2] for c in include}:
                if target - include:
                    reasons.append("the grid row names only some RTOs of this state")
                    sources.add(sheet or "the grid")
                target &= include
            if exclude and letters & {c[:2] for c in exclude}:
                if target & exclude:
                    reasons.append("the grid row leaves some RTOs of this state out")
                    sources.add(sheet or "the grid")
                target -= exclude
            if not target:
                continue
            if target == base:
                tokens.append(token)
                whole += 1
                continue
            cut = sorted(base - target)
            part = f"{shown(token)} without {plain_list(cut)}" if len(cut) <= len(target) else f"only {plain_list(target)} of {shown(token)}"
            got = self.pick(target)
            if got and RTO_NAMES == "cluster" and _single_rtos(got):
                got = None                         # single RTOs (RJ41) are not a cluster name: the state is the name and the part says which RTOs (the user, 30 Oct 2026)
            if not got and self.ncr and target & self.ncr and target - self.ncr:
                got = self.pick(target - self.ncr)                 # the NCR RTOs belong to the Delhi & NCR area: a group without them is accepted
            names = [x for x in include_names if self.m.codes(x) & target]
            if not got and names and include is not None and letters & {c[:2] for c in include}:
                cover = set().union(*[self.m.codes(x) for x in names])      # the partition the remarks named (Bangalore, a zone, metro ...)
                tokens.append(",".join(names))
                parts.append(f"the {' and '.join(shown(n) for n in names)} part of {shown(token)}")
                reasons.append("the grid row names a part of this state")
                why.extend(reasons)
                if cover != target:
                    missing.append(f"{', '.join(names)} without {_codes_text(sorted(cover - target))}")
                    self.gaps[missing[-1]] = sorted(target)
                continue
            why.extend(reasons or ["this part of the state has its own rate"])
            if got:
                tokens.append(got)
                parts.append(part)
            elif RTO_NAMES == "master":
                # the master has no group with exactly these RTOs: the cell says so in plain text (master_builder.py reads it); nothing is guessed and nothing is widened to the whole state
                warns.append(master_text(" + ".join(sorted(sources)) or sheet or "the grid", sorted(target)))
                self.gaps[f"{token} {len(target)} RTOs"] = sorted(target)
            else:
                tokens.append(token)
                parts.append(part)
                cut = sorted(base - target)
                missing.append(f"{token} without {_codes_text(cut)}" if len(cut) <= len(target) else f"only {_codes_text(sorted(target))} of {token}")
                self.gaps[missing[-1]] = sorted(target)
        out = []
        for t in tokens + warns:
            if t not in out:
                out.append(t)
        if RTO_NAMES in ("cluster", "master"):
            missing = []                                            # the condition is written in the RTO cell itself: nothing is 'missing in the master' any more
        part_text = " and ".join(parts) + (" and the whole of every other state named in the cluster name" if parts and whole else "")
        return RtoScope(",".join(out), "; ".join(dict.fromkeys(why)), part_text if why else ""), missing


def _single_rtos(name):
    """True for a name that is only RTO codes ('RJ41', 'RJ14,RJ41'): a list of individual RTOs, not a cluster"""
    return bool(name) and all(re.fullmatch(r"[A-Z]{2}\d{1,3}", p.strip()) for p in str(name).split(","))


RTO_NAMES = os.environ.get("GRID_RTO_NAMES", "master")       # Shriram (31 Oct 2026): the master's group name when the RTO codes are exactly a group, else a plain-text warning for master_builder.py
RTO_STRIP = r"^(?:SHRIRAM|SGIC)_(?:[A-Za-z]+\d{2}_)?"       # the insurer (+ month) prefix of the master's cluster groups (new_rto_list carries the bare cluster, 30 Oct 2026)


def _codes_text(codes):
    """['HR16', 'HR17', 'HR19'] -> 'HR 16, 17, 19'"""
    by = {}
    for c in codes:
        by.setdefault(c[:2], []).append(c[2:])
    return "; ".join(f"{k} {', '.join(v)}" for k, v in by.items())


def parse_code_list(text, default_letters=()):
    """'JK-01,04,06,08' -> {JK01, JK04, JK06, JK08}; 'RJ04, RJ06' -> {RJ04, RJ06}; 'WB-02 & 06' -> {WB02, WB06}. A bare number takes the letters before it."""
    out, letters = set(), (list(default_letters)[:1] or [None])[0]
    for m in re.finditer(r"([A-Za-z]{2})?\s*[-/ ]?\s*(\d{1,2})(?!\d)", _text(text)):
        L, n = m.group(1), m.group(2)
        if L and L.upper() in ("TO", "OR", "IN", "OF", "AT", "ON", "AS"):
            if L.upper() in ("TO", "OF", "AT", "ON", "IN"):
                L = None
        if L:
            letters = L.upper()
        if letters:
            out.add(letters + n.zfill(2))
    return out


# ---- product labels
def gvw_band(label):
    """GVW band (kg) of a goods-vehicle label -> (cc_min, cc_max) in the +0.01 convention.
    'UPTO 2800 GVW' (0, 2800.01); 'ABOVE 2801 GVW TO 3500 GVW' (2800.01, 3500.01); '12001 TO 20000' (12000.01, 20000.01); 'Above 50000' (50000.01, 99999)."""
    t = _text(label).upper().replace(",", "")
    nums = [int(x) for x in re.findall(r"(?<!\d)\d{3,6}(?!\d)", t)]
    start = lambda n: round(n - 1 + 0.01, 2) if n % 100 == 1 else round(n + 0.01, 2)
    if len(nums) >= 2:
        a, b = sorted(nums[:2])
        return start(a), round(b + 0.01, 2)
    if len(nums) == 1:
        n = nums[0]
        if re.search(r"UP\s*TO|UPTO|BELOW", t):
            return 0, round(n + 0.01, 2)
        if re.search(r"ABOVE|OVER|MORE", t):
            return start(n), CC_OPEN_END
    return None


def classify(label):
    """(kind, extra) of a PRODUCTS label; kind None = not a product this script knows."""
    k = label_key(label)
    if k.startswith("privatecar"):
        return ("pc_diesel" if "diesel" in k else "pc_petrol"), None
    if "corporatebus" in k:
        return "corp", None
    if "schoolbus" in k:
        return "school", None
    if k.startswith("pccv4w"):
        if re.search(r"upto61|upto6\+?1|61", k) and "710" not in k:
            return "taxi06", None
        if re.search(r"7to10|710", k):
            return "taxi710", None
    if k.startswith("pccv3w"):
        if "erickshaw" in k or "ecart" in k:
            return "erick", None
        if "otherthandiesel" in k:
            return "p3w_other", None
        if "diesel" in k:
            return "p3w_diesel", None
    if k.startswith("gccv3w"):
        return "g3w", None
    if k.startswith("miscd"):
        return ("tractor" if "tractor" in k else "miscd_others" if "others" in k else None), None
    if "newmahindra" in k:
        return "newmah", None
    if k.startswith("gccv") and re.search(r"lcv|mcv|hcv|above42501", k):
        band = gvw_band(label)
        return ("gcv4w", band) if band else (None, None)
    return None, None


# declined-sheet product labels (label_key) that apply to a product kind
def decline_labels(kind, band=None):
    if kind in ("pc_petrol",):
        return ["privatecar"]
    if kind == "pc_diesel":
        return ["privatecar", "privatecardiesel"]
    if kind == "taxi06":
        return ["pccv4w61"]
    if kind == "taxi710":
        return ["pccv4w710"]
    if kind == "g3w":
        return ["gccv3wexceptecart"]
    if kind == "p3w_diesel":
        return ["pccv3wdiesel", "pccv3wotherthanelectric"]
    if kind == "p3w_other":
        return ["pccv3wotherthandiesel", "pccv3wotherthanelectric"]
    if kind == "tractor":
        return ["miscdtractor"]
    if kind == "gcv4w" and band:
        lo, hi = band
        if hi <= 7500.01:
            labs = ["gccvlcvupto7500gvw"]
            if lo >= 2800:
                labs.append("gccvlcv2801to7500gvw")
            if lo >= 3500:
                labs.append("gccvlcvabove3500gvw")
            return labs
        if hi <= 12000.01:
            return ["gccvmcv"]
        return ["gccvhcv"] + (["gccvhcvabove50000gvw"] if lo >= 50000 else [])
    return []


EXEMPT_KINDS = {"erick", "school", "corp"}       # 'Pan India Except (PCCV 3W E-Rickshaw, E Cart & E Bike, PCCV 4W School bus and PCCV 4W Corporate bus)'


# ---- policy type, payout, age
def policy_subs(text):
    k = label_key(text)
    table = {"p": ["pkg"], "l": ["stp"], "pl": ["pkg", "stp"], "psaod": ["pkg", "saod"], "saod": ["saod"], "bundle13": ["b13"], "": []}
    return table.get(k)


def parse_po(v):
    """(kind, od, tp) of a payout cell: plain 25 / 'od' '15 OD' / 'net' '16 Net' / 'odtp' '15 OD + 2.5 TP' / 'irda'; None if unreadable."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return ("plain", float(v), None)
    t = norm(v)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%?", t)
    if m:
        return ("plain", float(m.group(1)), None)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%?\s*OD", t, re.I)
    if m:
        return ("od", float(m.group(1)), None)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%?\s*Net", t, re.I)
    if m:
        return ("net", float(m.group(1)), None)
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*%?\s*OD\s*\+\s*(\d+(?:\.\d+)?)\s*%?\s*TP", t, re.I)
    if m:
        return ("odtp", float(m.group(1)), float(m.group(2)))
    if t.upper() == "IRDA":
        return ("irda", IRDA_RATE, None)
    return None


def parse_age(text):
    """-> dict(pkg=(lo, hi) or None, stp=(lo, hi) or None, new=bool, old=bool).  'upto 15 years' -> (0, 15.01); '4 to 20 years' -> (3.01, 20.01);
    'upto 15 years & upto 25 Years for STP' -> pkg (0, 15.01), stp (0, 25.01); 'NEW' -> new."""
    t = norm(text).lower().rstrip(".")
    out = dict(pkg=None, stp=None, new=False, old=False)
    if not t:
        return out
    if re.fullmatch(r"new", t):
        out["new"] = True
        return out
    if re.match(r"old\b", t):
        out["old"] = True
        t = re.sub(r"^old\s*[,(]*\s*(age\s*)?", "", t).rstrip(")")

    def span(s):
        m = re.search(r"(?:upto|up to|below|<=?)\s*(\d+)", s)
        if m:
            return 0, round(int(m.group(1)) + 0.01, 2)
        m = re.search(r"(\d+)\s*(?:years?)?\s*(?:to|-)\s*(\d+)", s)
        if m:
            return round(int(m.group(1)) - 1 + 0.01, 2), round(int(m.group(2)) + 0.01, 2)
        return None
    parts = [p.strip() for p in re.split(r"\s*(?:&|\band\b)\s*(?=(?:upto|up to|\d+\s*(?:years?)?\s*(?:to|-)))", t) if p.strip()]
    stp = [p for p in parts if "stp" in p]
    pkg = [p for p in parts if "stp" not in p]
    s_pkg = span(pkg[0]) if pkg else None
    s_stp = span(stp[0]) if stp else None
    if re.search(r"package\s+and", t):
        m = re.search(r"(\d+)\s*to\s*(\d+)\s*years\s*package\s+and\s*(\d+)\s*to\s*(\d+)\s*years?\s*for\s*stp", t)
        if m:
            s_pkg = (round(int(m.group(1)) - 1 + 0.01, 2), round(int(m.group(2)) + 0.01, 2))
            s_stp = (round(int(m.group(3)) - 1 + 0.01, 2), round(int(m.group(4)) + 0.01, 2))
    if s_pkg and not s_stp and not stp:
        s_stp = s_pkg
    if s_stp and not s_pkg and not pkg:
        s_pkg = None
    out["pkg"], out["stp"] = s_pkg, s_stp
    return out


# ---- the UW REMARKS / REMARKS 1 columns: rules that turn a phrase into an effect on the rows
class Spec:
    """What the remarks of one grid row ask for.  Every phrase a rule understands is used up; what is left is flagged, never guessed at."""

    def __init__(self):
        self.mahindra_split = False        # 'Except New Mahindra': an all-makes row + a Mahindra row for old vehicles
        self.exclude = set()               # makes (TATA / ASHOK / MAHINDRA) left out of the all-makes group
        self.include = None                # makes the payout is for
        self.makes_override = None         # a make group of the portal named by the phrase
        self.ncb = None                    # 'YES' / 'NO' (package rows)
        self.tnc = []                      # T&C texts of the add_tnc column
        self.classes = None                # 'tanker' / 'nontanker'
        self.fuels = None
        self.part_include = []             # master RTO groups the payout is for (union)
        self.part_exclude = []             # master RTO groups taken out
        self.code_include = None           # RTO codes the payout is for
        self.code_exclude = set()          # RTO codes taken out
        self.skip = set()                  # sub-products not allowed ('saod')
        self.flags = []                    # texts nothing could read
        self.old = False
        self.zd = "NA"                     # with / without zero (nil) depreciation
        self.make_key = None               # a make group named by the phrase (MAKE_TABLE)
        self.make_text = None
        self.hazard = False                # the row carries the hazardous-category note
        self.cpa = "NA"                    # 'With CPA policy'
        self.sc = None                     # seating limit named in the remarks ('SC upto (3+1)')
        self.new_or_ncb = False            # 'New & With NCB cases only': new vehicles of any make, and used vehicles with NCB
        self.class_only = None             # vehicle classes the payout is for ('Truck', 'Goods Carrying Tractor')
        self.tnc_text = None               # a T&C text that replaces the kind's own
        self.taxi = None                   # 'ex_kp' (other than kaali peeli) / 'kp' (only kaali peeli taxi)
        self.misc = None                   # MISC-D 'others' wording
        self.bolero_ex = False             # 'Except (New Mahindra &) Bolero Model': the Mahindra row leaves the Bolero models out
        self.zd_pay = None                 # 'for NIL dep cases payout is 40%': the payin of the zero-dep rows (the cell's payout is for the others)
        self.cc_excl = None                # 'EXCLUDING 1001 CC TO 1500 CC': the cc range left out (lo, hi)
        self.mm_gt = None                  # 'M & M (All CC) & Others makes - Greater than 1500 CC (EXCEPT MARUTI)': the cc limit
        self.no_fuels = set()              # fuels the row declines ('OTHER THAN DIESEL')


STOP_WORDS = set("years year and only cases case make makes manf manuf manufacture manufactures the all for of in is are with a rto rtos code codes upto to by be will "
                 "payout po only not allowed under private cars car also".split())


def _lead(m):
    return m.group(0)


def _r(pattern):
    return re.compile(pattern, re.I)


def _codes_after(prefix, tail=r""):
    """regex: <prefix> followed by a list of RTO codes ('JK-01,04,06', 'UK 08/UK 07', 'WB-02 & 06', 'RJ04, RJ06')"""
    # a code is one or two digits and nothing else: '1001 CC' (the cc range of 'EXCLUDING 1001 CC TO 1500 CC') used to be read as RTO code 10, and '10 years' as code 10
    item = r"(?:[A-Za-z]{2}\s*-?\s*\d{1,2}(?!\d)|(?<![\d.])\d{1,2}(?!\d|\.\d))"
    return _r(prefix + r"[\s:(\-]*(?P<codes>" + item + r"(?:\s*(?:,|/|&|\band\b)\s*" + item + r")*)(?!\s*(?:cc\b|years?\b|yrs?\b|lacs?\b|%))" + tail + r"\s*\)?")


def _rule_list():
    R = []

    def add(pattern, fn):
        R.append((pattern if hasattr(pattern, "search") else _r(pattern), fn))

    def flag(msg):
        def f(s, m, kind, row):
            s.flags.append(msg)
        return f

    # -- hazardous category (cannot be seen in the MIS)
    add(r"for\s+hazardous\s+category\s+po\s+will\s+be\s+(\d+(?:\.\d+)?)\s*%\s+less",
        lambda s, m, k, row: (s.flags.append("For Hazardous category PO will be 5% less - not applied (the hazardous category is not in the MIS)") if HAZARDOUS_MODE == "flag" else None,
                              s.tnc.append(f"Hazardous category PO {float(m.group(1)):g}% less") if HAZARDOUS_MODE == "tnc" else None,
                              setattr(s, "hazard", True)))
    add(r"new\s*&\s*with\s+ncb\s+cases\s+only", lambda s, m, k, row: setattr(s, "new_or_ncb", True))
    # 'NCB Cases and New vehicles excluding New Mahindra' (Gujarat): NCB yes, an all-makes row and a Mahindra row for old vehicles - how the October and November uploads wrote it
    add(r"ncb\s+cases?\s+and\s+new\s+vehicles?\s+(?:excluding|except)\s+new\s+mahindra", lambda s, m, k, row: (setattr(s, "ncb", "YES"), setattr(s, "mahindra_split", True)))
    # -- Mahindra / Tata / Ashok Leyland
    add(r"other\s+than\s+tata\s*&\s*ashok\s+leyland(?:\s+make)?", lambda s, m, k, row: s.exclude.update({"TATA", "ASHOK"}))
    add(r"other\s+than\s+tata\s*&\s*new\s+mahindra", lambda s, m, k, row: (s.exclude.add("TATA"), setattr(s, "mahindra_split", True)))
    add(r"other\s+than\s+tata\b", lambda s, m, k, row: s.exclude.add("TATA"))
    add(r"except\s+new\s+mahindra\s*&\s*ashok\s+leyland", lambda s, m, k, row: (s.exclude.add("ASHOK"), setattr(s, "mahindra_split", True)))
    # Bolero: the portal has a group of the Bolero models and one of Mahindra without them.  'Bolero Model only' is the first (the November 2026 upload of the user: GCV_4W_BOLERO);
    # 'Except (New Mahindra &) Bolero Model' leaves them out of the Mahindra row (Mahindra_ex_Bolero); the October upload ignored the Bolero wording, the November one did not
    add(r"bolero\s+models?\s+only", lambda s, m, k, row: setattr(s, "makes_override", MAKES_BOLERO))
    add(r"except\s+new\s+mahindra\s*&\s*bolero\s+models?\s*\(dl\s+rto\s+only\)", lambda s, m, k, row: (setattr(s, "mahindra_split", True), setattr(s, "bolero_ex", True), s.part_include.append("DL")))
    add(r"except\s+(?:new\s+mahindra\s*&\s*bolero\s+models?|bolero\s+models?\s*&\s*new\s+mahindra)", lambda s, m, k, row: (setattr(s, "mahindra_split", True), setattr(s, "bolero_ex", True)))
    add(r"except\s+ashok\s+leyland\s*&\s*m\s*&\s*m\s+manufactures", lambda s, m, k, row: s.exclude.update({"ASHOK", "MAHINDRA"}))
    add(r"except\s+ashok\s+leyland", lambda s, m, k, row: s.exclude.add("ASHOK"))
    add(r"except\s+tata\b", lambda s, m, k, row: s.exclude.add("TATA"))
    add(r"except\s+m\s*&\s*m\b", lambda s, m, k, row: s.exclude.add("MAHINDRA"))
    add(r"excluded\s+ashok\s+leyland", lambda s, m, k, row: s.exclude.add("ASHOK"))
    add(r"(?:except|excluding)\s+new\s+mahindra|except\s+new\b|all\s+cases\s*,\s*except\s+new\s+mahindra", lambda s, m, k, row: setattr(s, "mahindra_split", True))
    add(r"tata\s*&\s*ashok\s+leyland\s+make\s+only", lambda s, m, k, row: setattr(s, "include", ["TATA", "ASHOK"]))
    add(r"tata\s*&\s*sml\s+isuzu(?=\s*\(cng)", lambda s, m, k, row: setattr(s, "include", ["SML ISUZU", "TATA"]))
    add(r"only\s+tata\s+make|\(\s*tata\s+make\s+only\s*\)", lambda s, m, k, row: setattr(s, "include", ["TATA"]))
    add(r"\(\s*ashok\s+leyland\s*\)", lambda s, m, k, row: setattr(s, "include", ["ASHOK"]))
    add(r"bajaj\s*(?:manufacturer|make)\s+only|bajaj\s+and\s+tvs\s+make\s+only", lambda s, m, k, row: setattr(s, "include", ["BAJAJ", "TVS"] if "tvs" in m.group(0).lower() else ["BAJAJ"]))
    add(r"other\s+than\s+bajaj\s+manufacturer", lambda s, m, k, row: s.exclude.add("BAJAJ"))
    # -- NCB
    add(r"ncb\s*=\s*[<>]\s*=?\s*\d+\s*%", lambda s, m, k, row: (setattr(s, "ncb", "YES"), s.tnc.append(
        re.sub(r"\s+", " ", m.group(0)).upper() if k.startswith("pc") else re.sub(r"ncb\s*=\s*([<>])\s*=?\s*(\d+\s*%)", r"\1=\2", m.group(0), flags=re.I))))
    # 'NCB Cases 20% and Above' = the 'NCB =>20%' of the car rows: NCB yes and the condition as T&C (the user's November 2026 upload: 'NCB >=20%'; October wrote '>=20%')
    add(r"ncb\s+cases?\s+(\d+(?:\.\d+)?)\s*%\s+and\s+(above|more|below|less)",
        lambda s, m, k, row: (setattr(s, "ncb", "YES"), s.tnc.append(f"NCB {'>=' if m.group(2).lower() in ('above', 'more') else '<='}{float(m.group(1)):g}%")))
    add(r"(?:with)?out\s+ncb|non[\s-]*ncb", lambda s, m, k, row: setattr(s, "ncb", "NO"))
    add(r"(?:with\s+)?ncb\s+cases?(?:\s+only)?|with\s+ncb(?:\s+only)?|only\s+ncb\s+cases?|ncb\s+case(?:s)?\b", lambda s, m, k, row: setattr(s, "ncb", "YES"))
    # -- SA-OD, IDV, add-ons
    add(r"\(?\s*sa\s*-?\s*od\s+(?:is\s+)?not\s+allowed(?:\s+under\s+private\s+cars)?\.?\s*\)?", lambda s, m, k, row: s.skip.add("saod"))      # 'SA- OD not allowed' (a space after the hyphen) too
    add(r"\(?\s*except\s+sa\s*-?\s*od\s*\)?", lambda s, m, k, row: s.skip.add("saod"))
    add(r"idv\s+upto\s+10\s+lacs(?:\s+only)?\.?", lambda s, m, k, row: s.tnc.append("IDV UPTO 10 LACS"))
    add(r"without\s+smp\s+addon(?:\s+compulsory)?", lambda s, m, k, row: s.tnc.append("Without SMP Addon Compulsory"))
    add(r"smp\s+addon\s+compulsory", lambda s, m, k, row: s.tnc.append("SMP Addon Compulsory"))
    add(r"body\s+type\s+buker\s+truck\s+only", lambda s, m, k, row: (s.tnc.append("Buker Truck"), setattr(s, "class_only", ["Truck"])))
    add(r"gccv\s+lcv\s+upto\s+2800\s+gvw\s*-\s*tractor\s+only", lambda s, m, k, row: setattr(s, "class_only", ["Goods Carrying Tractor"]))
    add(r"3\s+add\s+on\s+cover\s+compulsory", lambda s, m, k, row: s.tnc.append("3 Add on cover compulsory"))
    add(r"with\s+cpa\s+policy", lambda s, m, k, row: setattr(s, "cpa", "YES"))
    add(r"sc\s+upto\s*\(\s*(\d+)\s*\+\s*(\d+)\s*\)", lambda s, m, k, row: setattr(s, "sc", (0, round(int(m.group(1)) + int(m.group(2)) + 0.01, 2))))
    # -- fuel
    add(r"\(?\s*cng\s+only\s*\)?", lambda s, m, k, row: setattr(s, "fuels", ["CNG"]))
    add(r"\(?\s*diesel\s+only\s*\)?", lambda s, m, k, row: setattr(s, "fuels", ["Diesel"]))
    add(r"except\s+cng\b", lambda s, m, k, row: setattr(s, "fuels", ["Petrol", "Diesel", "LPG", "Electric"]))
    add(r"electric\s+vehicles?\s+only", lambda s, m, k, row: setattr(s, "fuels", ["Electric"]))
    # 'GCCV 3W EXCEPT E-CART OTHER THAN DIESEL' (MP, RJ): no Diesel rows - the October upload had none for these two states either (the product words in front are the row's own label)
    add(r"(?:gccv\s+3w\s+except\s+e-?cart\s+)?(?:other\s+than|except|excluding)\s+diesel\b", lambda s, m, k, row: s.no_fuels.add("Diesel"))
    # -- tankers
    add(r"(?:excluding|other\s+than)\s+tankers?|tankers?\s+excluded", lambda s, m, k, row: setattr(s, "classes", "nontanker"))
    add(r"(?:all\s+)?tankers?(?:\s*\(new\s*&\s*claim\s+free\s+tankers\))?", lambda s, m, k, row: setattr(s, "classes", "tanker"))
    # -- Private Car make groups
    mk = lambda key: (lambda s, m, k, row: (setattr(s, "make_key", key), setattr(s, "make_text", re.sub(r"\s+", " ", m.group(0)).strip())))
    first_codes = lambda m, row: parse_code_list(m.group("codes"), _state_letters(row["state"]))
    # taxis: models the portal has groups for (the user's November 2026 upload: Dzire, Ambasador, Mahindra left out)
    add(r"(?:maruti\s+suzuki\s*-?\s*)?swift\s+dzire\s+models?\s+(?:is\s+)?(?:declined|excluded)(?:\s+for\s+(?P<codes>[A-Za-z]{2}\s*-?\s*\d{1,2}))?",
        lambda s, m, k, row: (mk("dzire_ex")(s, m, k, row), s.code_exclude.update(first_codes(m, row)) if m.group("codes") else None))
    dzire_only = lambda s, m, k, row: (mk("dzire_only")(s, m, k, row), setattr(s, "code_include", (s.code_include or set()) | first_codes(m, row)))
    add(r"for\s+(?P<codes>[A-Za-z]{2}\s*-?\s*\d{1,2})\s+and\s+for\s+maruti\s+suzuki\s*-?\s*swift\s+dzire", dzire_only)               # November: 'For RJ-14 and for Maruti Suzuki-Swift Dzire'
    add(r"for\s+maruti\s+suzuki\s*-?\s*swift\s+dzire\s+of\s+(?P<codes>[A-Za-z]{2}\s*-?\s*\d{1,2})", dzire_only)                        # October: 'For Maruti Suzuki-Swift Dzire of RJ-14'
    add(r"(?:6\s*\+\s*1\s+and\s+)?including\s+kaali\s+peeli\s+taxi\s+ambass?ador\s+models?\s+only", mk("ambasador"))
    add(r"excluding\s+m\s*&\s*m(?:\s+make)?\b(?!\s+bolero)", lambda s, m, k, row: s.exclude.add("MAHINDRA") if k == "gcv4w" else mk("mm_ex")(s, m, k, row))
    # Private Car: 'M & M (All CC) & Others makes - Greater than 1500 CC (EXCEPT MARUTI make)' = Mahindra at any cc + every other make but Maruti above the limit; the row 'OTHER THAN M &M'
    # next to it then stops at the limit (the user's November 2026 upload for Bihar)
    add(r"m\s*&\s*m\s*\(\s*all\s+cc\s*\)\s*&\s*others?\s+makes?\s*-?\s*greater\s+than\s+(\d+)\s*cc\s*\(\s*except\s+maruti\s+make\s*\)\.?", lambda s, m, k, row: setattr(s, "mm_gt", int(m.group(1))))
    # a cc range that is left out ('EXCLUDING 1001 CC TO 1500 CC'): the cc band is written in two pieces, below and above it
    add(r"excluding\s+(\d{3,5})\s*cc\s+to\s+(\d{3,5})\s*cc\b\.?", lambda s, m, k, row: setattr(s, "cc_excl", (int(m.group(1)), int(m.group(2)))))
    add(r"only\s+honda\s*&\s*hyundai\s*&\s*kia\s*&\s*citroen\s+manufacture\s+only", mk("hhk_citroen"))
    add(r"only\s+honda\s*&\s*hyundai\s*&\s*kia\s+manufacture\s+only", mk("hhk"))
    add(r"citroen\s+manufacture\s+only", mk("citroen"))
    add(r"hyundai\s*&\s*maruti\s+only\s*\(\s*alto\s+declined\s*\)", mk("hyundai_maruti_ex_alto"))
    add(r"(?:upto|above)\s+1000\s+cc\s*-\s*hyundai\s*&\s*maruti\s*\(\s*except\s+alto\s*\)\.?", mk("hyundai_maruti_ex_alto"))      # the cc split is not written (as in the earlier uploads)
    add(r"hyundai\s*,\s*honda\s*,\s*m\s*&\s*m\s*,\s*tata\s*&\s*toyota\s+only", mk("hhtmt"))
    add(r"other\s+than\s+maruti\s+suzuki\s*&\s*m\s*&\s*m", mk("ex_maruti_mm"))
    add(r"except\s+maruti\s+alto\s*/\s*swift", mk("ex_alto_swift"))
    add(r"maruti\s+suzuki\s*\(\s*excluding\s+xl-?6\s*&\s*swift\s*\)", mk("maruti_ex_swift_xl6"))
    add(r"except\s+maruti\s*/\s*hyundai", mk("ex_hyundai_maruti"))
    add(r"alto\s+decline[d]?", mk("alto_declined"))
    add(r"(?:eecco|ecco|eeco)\s+(?:is\s+)?(?:declined|delcined|delined)|with\s+ncb\s*/\s*eeco\s+is\s+declined", mk("eeco_declined"))
    add(r"excluding\s+renault\s+manuf(?:acture|\.)?", mk("renault_ex"))
    add(r"other\s+than\s+hyundai\s*&\s*maruti\.?", mk("ex_hyundai_maruti"))
    add(r"(?:except|excluding)\s+maruti(?:\s+suzuki)?(?!\s*/)", mk("ex_maruti"))
    add(r"other\s+than\s+m\s*&\s*m\b", mk("ex_mm"))
    add(r"m\s*&\s*m\s+(?:manf|make)\s+only", mk("mm_only"))
    add(r"maruti\s+suzuki(?:\s+manf\s+only)?(?=\s*(?:$|\(|[.-]\s*$))", mk("maruti_only"))
    add(r"only\s+eeco\s*\(\s*sa\s*-?\s*od\s+and\s+stp\s+not\s+allowed\s*\)", lambda s, m, k, row: (setattr(s, "make_key", "only_eeco"), s.skip.update({"saod", "stp"})))
    add(r"excluding\s+m\s*&\s*m\s+bolero", mk("bolero_ex"))
    add(r"allowed\s+manufactures\s*-\s*tata\s+magic[^.]*\.", mk("allowed_list"))
    # -- taxis
    add(r"other\s+than\s+maruti\s+suzuki\s*,\s*hyundai\s+and\s+m\s*&\s*m\s+bolero\s+make\s+only\.?", mk("ex_mshb"))
    add(r"maruti\s+suzuki\s*,\s*hyundai\s+and\s+m\s*&\s*m\s+bolero\s+make\s+only\.?", mk("mshb"))
    add(r"other\s+than\s+kaali\s+peeli\s+taxi(?:\s+and\s+mumbai\s+rto\s+codes)?\.?", lambda s, m, k, row: setattr(s, "taxi", "ex_kp"))
    add(r"only\s+kaali\s+peeli\s+taxi\s+and\s+blue\s+cabs(?:\s+and\s+mumbai\s+rto\s+codes)?", lambda s, m, k, row: setattr(s, "taxi", "kp"))
    add(r"all\s+\(except\s*-\s*commander[^)]*(?:\)|$)", lambda s, m, k, row: None)
    # -- MISC-D
    add(r"jcb\s*&\s*excavator\s*\(\s*all\s+make\s*\)\s+and\s+except\s+cranes?,?\s*ambulance,?\s*tractor\.?", lambda s, m, k, row: setattr(s, "tnc_text", "JCB & Excavator /-"))
    add(r"jcb\s+manufacture\s+only\s*-\s*excluding\s+cranes?,?\s*tractor\s+and\s+ambulance|\bjcb\s+only\b", lambda s, m, k, row: setattr(s, "tnc_text", "JCB /-"))
    add(r"back\s*hoe\s+loader\s*only\s*-\s*excluding\s+cranes?,?\s*tractor\s+and\s+ambulance", lambda s, m, k, row: setattr(s, "tnc_text", "Backhoe loader /-"))
    add(r"excluding\s+crane\s+and\s+ambulance", lambda s, m, k, row: setattr(s, "tnc_text", "EX:Crane & Ambulance"))
    # crane / tractor / ambulance in any order ('Upto 10 years and except cranes, ambulance, tractor.' left ', tractor.' over and was flagged)
    veh = r"(?:cranes?|tractors?|ambulances?)"
    add(r"(?:\(age\s*>0-10\)\s*all\s+except|upto\s+10\s+years\s+and\s+except|excluding)\s+" + veh + r"(?:\s*[,/]\s*(?:and\s+)?" + veh + r"|\s+and\s+" + veh + r"){1,2}"
        r"(?:\s+and\s+other\s+than\s+combined\s+harvester(?:\s+and\s+harvester)?)?\.?",
        lambda s, m, k, row: (setattr(s, "misc", "ex_other" if "other than combined" in m.group(0).lower() else "ex")
                              if all(w in m.group(0).lower() for w in ("crane", "tractor", "ambulance")) else s.flags.append(m.group(0))))      # only the three together are the known exclusion
    add(r"for\s+combined\s+harvester\s+and\s+harvester\s+only", lambda s, m, k, row: setattr(s, "misc", "harvester"))
    add(r"agriculture\s+tractors?\s+only\s+without\s+trolly/trailer\.?", lambda s, m, k, row: setattr(s, "tnc_text", "Agriculture tractors only without trolly,trailer"))
    add(r"tractor\s+is\s+only\s+new(?:\s+and\s+without\s+trailer)?\.?", lambda s, m, k, row: None)
    add(r"^\s*old\s*$", lambda s, m, k, row: setattr(s, "old", True))
    # -- areas named in the remarks that the STATE column already says
    add(r"\b(?:ga|mumbai|rom)\s+rto(?:'?s)?(?:\s+codes?)?(?:\s+only)?\b|all\s+rom\s+rto'?s", lambda s, m, k, row: None)
    add(r"whole\s+ka\s+rto\s+codes", lambda s, m, k, row: None)
    # -- partitions: Bangalore, Rajasthan zones
    add(r"other\s+than\s+bangalore\s+rto\s+codes?", lambda s, m, k, row: s.part_exclude.append(BANGALORE_GROUP))
    add(r"bangalore\s+rto\s+codes?", lambda s, m, k, row: s.part_include.append(BANGALORE_GROUP))
    add(r"zone\s*([123])\s*(?:&|and)\s*(?:zone\s*)?([123])", lambda s, m, k, row: s.part_include.extend(RJ_ZONE_GROUPS[int(g)] for g in m.groups()))
    add(r"\bzone\s*([123])\b", lambda s, m, k, row: s.part_include.append(RJ_ZONE_GROUPS[int(m.group(1))]))
    letters = lambda row: _state_letters(row["state"])

    # -- Delhi / Tamil Nadu metro partitions
    add(r"\(\s*dl\s*(?:rto\s+only)?\s*\)", lambda s, m, k, row: s.part_include.append("DL"))
    add(r"non[\s-]*metro\s+rto(?:'?s|\s+codes?)(?:\s+only)?", lambda s, m, k, row: s.part_include.append(TN_NON_METRO_GROUP) if "TN" in letters(row) else s.flags.append(m.group(0)))
    add(r"\bmetro\s+rto(?:'?s|\s+codes?)(?:\s+only)?", lambda s, m, k, row: s.part_include.append(TN_METRO_GROUP) if "TN" in letters(row) else s.flags.append(m.group(0)))
    add(r"chennai\s+rto\s+codes?", lambda s, m, k, row: s.part_include.append(CHENNAI_GROUP))
    add(r"tn\s*&\s*py\s+rto\s+codes?", lambda s, m, k, row: None)
    # -- zero / nil depreciation
    # 'for NIL dep cases payout is 40%': the payout of the cell is for the cases without nil dep, the zero-dep rows get the 40
    add(r"for\s+(?:nil|zero)\s+dep(?:\s+cases)?\s+(?:payout|po)\s+(?:is|will\s+be)\s+(\d+(?:\.\d+)?)\s*%?", lambda s, m, k, row: (setattr(s, "zd_pay", float(m.group(1))), setattr(s, "zd", "NO")))
    add(r"(?:without|except)\s+(?:nil|zero)\s+dep(?:\s+cases)?", lambda s, m, k, row: setattr(s, "zd", "NO"))      # 'Except NIL dep cases' is WITHOUT nil dep (it used to be read as with)
    add(r"(?:with\s+)?(?:nil|zero)\s+dep(?:\s+cases)?", lambda s, m, k, row: setattr(s, "zd", "YES"))
    # 'UPTO 15 YEARS (DL)' says the age of the AGE column again: nothing to apply; a limit that differs from the AGE column is flagged, not dropped
    def age_again(s, m, k, row):
        a = parse_age(row["age"])
        if round(int(m.group(1)) + 0.01, 2) not in {x[1] for x in (a["pkg"], a["stp"]) if x}:
            s.flags.append(m.group(0))
    add(r"upto\s+(\d{1,2})\s+years?\b", age_again)
    # -- explicit RTO codes
    def codes_of(m, row):
        return parse_code_list(m.group("codes"), letters(row))
    add(_r(r"(?P<codes>[A-Za-z]{2}\s*-?\s*\d{1,2}(?:\s*,\s*(?:[A-Za-z]{2}\s*-?\s*)?\d{1,2})*)\s+(?:is\s+)?(?:excluded|declined)(?:\s*\(part of [^)]*\))?"),
        lambda s, m, k, row: s.code_exclude.update(codes_of(m, row)))
    add(r"\(\s*declined\s+rto(?:\s+code)?\s+is\s+[A-Za-z]{2}-?\d{1,2}\s*(?:only)?\s*\)", lambda s, m, k, row: None)
    # 'excluded / declined for RJ-14' takes the RTO OUT (it used to be read as 'for RJ-14', i.e. as the only RTO); 'EXL RTO codes 03, 06', 'excluding RTO (WB-04, ...)' too
    add(_codes_after(r"(?:excluded|declined|decline)\s+for\s*(?:rtos?(?:\s+codes?)?)?"), lambda s, m, k, row: s.code_exclude.update(codes_of(m, row)))
    add(_codes_after(r"(?:other\s+than|excluding|excl?|exl)\s+rtos?(?:\s+codes?)?\s*[-:]?"), lambda s, m, k, row: s.code_exclude.update(codes_of(m, row)))
    add(_codes_after(r"other\s+than"), lambda s, m, k, row: s.code_exclude.update(codes_of(m, row)))                  # 'other than JK- 01, 03, 04 ... 22 RTO'S'
    add(_codes_after(r"(?:except|allow\s+except|excluding)\s+(?:rto\s+codes?\s*(?:rtos:)?)?"), lambda s, m, k, row: s.code_exclude.update(codes_of(m, row)))
    add(_codes_after(r"(?:allowed\s+for|for(?:\s+for)?)\s+(?:rtos?(?:\s+codes?)?\s*[-:]?)?\s*\(?\s*(?=[A-Za-z]{2}\s*-?\s*\d)"), lambda s, m, k, row: s.__setattr__(
        "code_include", (s.code_include or set()) | codes_of(m, row)))
    add(_r(r"^\W*\(?\s*(?P<codes>[A-Za-z]{2}\s*-?\s*\d{1,2}(?:\s*[,/&]\s*(?:[A-Za-z]{2}\s*-?\s*)?\d{1,2})*)\s*\)?\s*(?=[,;(]|$)"),
        lambda s, m, k, row: s.__setattr__("code_include", (s.code_include or set()) | codes_of(m, row)))
    return R


RULES = _rule_list()

# the sheets name each state by its letters; the STATE column of the grid by words
def _state_letters(state):
    area = area_of(state)
    if area == ALL_INDIA:
        return ()
    out = []
    for kind, v in area:
        if kind == "codes":
            out.extend(v)
    return tuple(out)


def area_of(state):
    k = label_key(state)
    if k in ("panindia", "allindia"):
        return ALL_INDIA
    for name, pieces in STATE_AREAS.items():
        if label_key(name) == k:
            return pieces
    return None


def is_change_note(text):
    """A note that only says what changed since the last grid ('Disc by 5% & PO by 2% increased', 'PO redcued by 1%', 'Dic by 5% & PO increased by 7.5%', 'Age criteria improved', "UW's guideline
    changed", 'New Addition'): every word is one of CHANGE_WORDS / CHANGE_VERBS and one of them is a verb.  A state, an RTO, a make or a product word makes it a rule, not a change note.  The
    user cleared all of these without touching a row in the November 2026 upload (and each month the typo is different: 'redcued', 'Dic', 'Addtion')."""
    words = re.findall(r"[a-z]+", re.sub(r"['’`]s\b", "s", str(text).lower()))
    return bool(words) and all(w in CHANGE_WORDS or w in CHANGE_VERBS for w in words) and any(w in CHANGE_VERBS for w in words)


def is_standing(text):
    k = note_key(text)
    return any(note_key(s) in k for s in STANDING_NOTES) or bool(CHANGE_NOTE.match(text)) or is_change_note(text)


def read_remarks(kind, row):
    """Spec of one grid row from its UW REMARKS and REMARKS 1 cells."""
    spec = Spec()
    for original in (row["uw"], row["rem1"]):
        if not original:
            continue
        t = original
        for rx, fn in RULES:
            guard = 0
            while guard < 20:
                m = rx.search(t)
                if not m:
                    break
                fn(spec, m, kind, row)
                t = t[:m.start()] + " " + t[m.end():]
                guard += 1
        words = [w for w in re.findall(r"[A-Za-z]{3,}", t.lower()) if w not in STOP_WORDS]
        numbers = re.findall(r"\d+(?:\.\d+)?", t)            # a number no rule used up ('for NIL dep cases payout is 40%', '1001 CC TO 1500 CC') is a rule nobody read: flagged, never dropped
        if (words or numbers) and not is_standing(original):
            spec.flags.append(original)
    return spec


# ---- make groups of goods vehicles
def gcv_group(excluded):
    """new_vehicle_makes of 'every GCV make except these (TATA / ASHOK / MAHINDRA)': the portal's group, or a group plus the makes it must give back."""
    excluded = frozenset(excluded)
    if excluded in GCV_EXCLUSIONS:
        return GCV_EXCLUSIONS[excluded], []
    bigger = [e for e in GCV_EXCLUSIONS if excluded <= e]
    if not bigger:
        return None, [f"no make group excludes exactly {sorted(excluded)}"]
    e = min(bigger, key=len)
    back = sorted(e - excluded)
    return ",".join([GCV_EXCLUSIONS[e]] + [MAKE_TOKENS.get(b, b) for b in back]), []


def gcv_variants(spec, floor_lo=0, carve=False):
    """[(new_vehicle_makes, add_tnc prefix, minimum vehicle age)] of a goods-vehicle row.  The Mahindra row (old vehicles of a row that says 'Except New Mahindra') leaves the Bolero models out
    when the row says so itself ('Except Bolero Model & New Mahindra') or when another row of the state prices the Bolero models on their own (`carve`)."""
    out, problems = [], []
    if spec.makes_override:
        return [(spec.makes_override, "", 0)], problems
    mahindra = MAKES_MAHINDRA_EX_BOLERO if (spec.bolero_ex or carve) else MAKE_TOKENS["MAHINDRA"]
    if spec.include is not None:
        names = ",".join(MAKE_TOKENS.get(x, x) for x in spec.include)
        out.append((names, "", 0))
        if spec.mahindra_split and "MAHINDRA" in spec.include:
            out.append((mahindra, "Old", 1))
        return out, problems
    excl = set(spec.exclude) | ({"MAHINDRA"} if spec.mahindra_split else set())
    name, probs = gcv_group(excl)
    problems += probs
    out.append((name or GCV_BASE, "", 0))
    if spec.mahindra_split and "MAHINDRA" not in spec.exclude:
        out.append((mahindra, "Old", 1))
    return out, problems


# ---- one grid row -> upload rows
SUBS = {"pkg": ("1+1", "Renew"), "stp": ("STP", "Renew"), "saod": ("SAOD", "Renew"), "b13": ("1+3", "New")}


def rate_for(po, sub, kind):
    """(pi_type, od, tp, net) of a payout cell for a sub-product; None when the cell gives none."""
    typ, a, b = po
    pc = kind in ("pc_petrol", "pc_diesel")
    if sub == "stp":
        if typ == "irda":
            return None
        return "On TP", 0.0, (b if typ == "odtp" else a), 0.0
    if sub in ("saod", "b13"):
        return "On OD", a, 0.0, 0.0
    if typ == "odtp":
        return "On OD and TP", a, b, 0.0
    if typ in ("od", "irda") or (pc and typ in ("plain", "net")):
        return "On OD", a, 0.0, 0.0
    return "On Net", 0.0, 0.0, a


def make_row(g, sub, fuel, product, mmc, makes, rto, age, cc, sc, rate, tariff_hi, ncb, tnc, dates, flags, need):
    pi_type, od, tp, net = rate
    pi_od, pi_tp, pi_net = round_half_up(od), round_half_up(tp), round_half_up(net)
    cut = lambda x: round_half_up(max(x - PO_CUT, 0.0)) if x > 0 else 0.0
    policy_sub, policy = SUBS[sub]
    return dict(
        id=ROW_ID, insurance_company=INSURANCE_COMPANY, insurer_vertical="Not Applicable", product=product, sub_product=policy_sub, policy_type=policy,
        fuel_type=fuel, make_model_class=mmc, new_vehicle_makes=makes, new_rto_list=rto_cell(rto), vehicle_age_min=age[0], vehicle_age_max=age[1],
        cc_min=cc[0], cc_max=cc[1], sc_min=sc[0], sc_max=sc[1], pi_od_rate=pi_od, pi_tp_rate=pi_tp, pi_tp_2=0, pi_tp_3=0, pi_tp_4=0, pi_tp_5=0,
        pi_net_rate=pi_net, pi_flat_amount=0, pi_vli=0, pi_type=pi_type, tariff_min=0, tariff_max=tariff_hi, is_ncb=ncb, is_cpa="NA", is_zd="NA",
        from_date=dates[0], to_date=dates[1], user_id=USER_ID, veh_use="", add_tnc=tnc, remarks="", po_type=pi_type,
        po_od_rate=cut(pi_od), po_tp_rate=cut(pi_tp), po_net_rate=cut(pi_net), po_flat_amount=0,
        _flags=tuple(flags), _need=need, _r=g["r"], _sub=sub,
    )


def tariff_hi(g, sub, warn):
    if sub == "stp":
        return TARIFF_STP[1]
    d = g["dis"]
    x = num(d)
    if x is None:
        if _text(d):
            warn(f"row {g['r']}: DIS % '{_text(d)}' is not a number - the row is written with no discount limit (tariff 0-100.01)")
        return TARIFF_STP[1]
    return TARIFF_STP[1] if x >= 99 else round(x + 0.01, 2)


def _old_tnc(tnc, keep):
    """'Old' / 'Old,SMP Addon Compulsory' lose their 'Old' when the age band already starts above new vehicles"""
    if keep or not tnc.startswith("Old"):
        return tnc
    return tnc[4:] if tnc.startswith("Old,") else ("" if tnc == "Old" else tnc)


def join_tnc(*parts):
    return ",".join(p for p in parts if p)


def hazard_variants(vs):
    """'For Hazardous category PO will be 5% less': the Tanker rows lose HAZARDOUS_CUT points of payin; in mode "tanker" a row for all classes becomes a Tanker row and the five other classes."""
    out = []
    for v in vs:
        if v["product"] != "GCV 4W" or v["mmc"] not in ("NA", "Tanker"):
            out.append(v)
        elif v["mmc"] == "Tanker":
            out.append(dict(v, cut=HAZARDOUS_CUT))
        elif HAZARDOUS_MODE != "tanker":
            out.append(v)
        else:
            out.append(dict(v, mmc="Tanker", cut=HAZARDOUS_CUT))
            out.extend(dict(v, mmc=c) for c in GCV_CLASSES_OTHER)
    return out


def cut_rate(rate, pts):
    """the payin of a (pi_type, od, tp, net) rate lowered by pts points on its main part (never below 0)"""
    typ, od, tp, net = rate
    if typ == "On Net":
        net = max(net - pts, 0.0)
    elif typ == "On TP":
        tp = max(tp - pts, 0.0)
    else:
        od = max(od - pts, 0.0)
    return typ, od, tp, net


def convert_row(g, ctx):
    """The upload rows of one grid row, or [] with a warning."""
    warn = ctx.warn
    kind, extra = classify(g["product"])
    if kind is None:
        ctx.skipped[f"product '{g['product']}'"] = ctx.skipped.get(f"product '{g['product']}'", 0) + 1
        return []
    work = g.get("rto", "")                                          # the person's working-list name: used as it is, instead of the lookup below
    area = area_of(g["state"])
    if area is None and not work:
        ctx.skipped[f"state '{g['state']}'"] = ctx.skipped.get(f"state '{g['state']}'", 0) + 1
        return []
    subs = policy_subs(g["ptype"])
    if subs is None:
        ctx.skipped[f"policy type '{g['ptype']}'"] = ctx.skipped.get(f"policy type '{g['ptype']}'", 0) + 1
        return []
    if kind in ("corp", "school"):
        return convert_bus(g, kind, subs, ctx)
    po = parse_po(g["po"])
    if po is None:
        ctx.skipped[f"payout '{_text(g['po'])}'"] = ctx.skipped.get(f"payout '{_text(g['po'])}'", 0) + 1
        return []
    age = parse_age(g["age"])
    spec = ctx.spec_for(g, kind)
    out = []
    labels = decline_labels(kind, extra)
    exempt = kind in EXEMPT_KINDS
    for sub in subs:
        if sub in spec.skip or sub in g.get("drop", ()):
            continue                                            # skipped by a rule, or replaced by an addendum row
        rate = rate_for(po, sub, kind)
        if rate is None:
            warn(f"row {g['r']}: an IRDA payout has no rate for {SUBS[sub][0]} - that part is not written")
            continue
        # 'for NIL dep cases payout is 40%': the cell's payout is for the cases without nil dep, the zero-dep rows get the 40 (on STP both meet on one key and the lowest wins)
        rates = [(spec.zd, rate)]
        if spec.zd_pay is not None:
            alt = with_main(rate, spec.zd_pay)
            if alt is None:
                spec.flags.append(f"payout for NIL dep cases {spec.zd_pay:g}% - the cell has two parts (OD and TP), not applied")
            else:
                rates = [("NO", rate), ("YES", alt)]
        a = age["stp"] if sub == "stp" else age["pkg"]
        variants = split_cc(kind_variants(g, kind, extra, spec, age, sub, ctx), spec.cc_excl)
        if spec.hazard and HAZARDOUS_MODE in ("tanker", "tanker_rows"):
            variants = hazard_variants(variants)
            if HAZARDOUS_MODE == "tanker_rows" and spec.classes is None:
                ctx.hazard_unsplit.add(g["r"])
        for v in variants:
            lo_floor = v.get("floor", 0)
            rng = v.get("age") or a
            if rng is None:
                rng = (0, AGE_OPEN_END) if sub == "b13" else (0, AGE_OPEN_END)
            old_ok = rng[0] == 0 and rng[1] >= 15                          # the 'Old' T&C only where the band starts at new vehicles
            rng = (max(rng[0], lo_floor), rng[1])
            if work:
                rto, missing = RtoScope(working_rto(work)), []
            else:
                rto, missing = ctx.rtos.resolve(area, ctx.declines, labels if not v.get("labels") else v["labels"], exempt,
                                                include=_include(spec, ctx), exclude=_exclude(spec, ctx), extra=frozenset(spec.code_exclude),
                                                include_names=tuple(spec.part_include), sheet=g["sheet"])
            need = tuple(f"No RTO group in the master for {m} - written for the whole area" for m in missing)
            for fuel in [f for f in v["fuels"] if f not in spec.no_fuels]:
                for zd, rt in rates:
                    row = make_row(g, sub, fuel, v["product"], v["mmc"], v["makes"], rto, rng, v["cc"], v["sc"], cut_rate(rt, v["cut"]) if v.get("cut") else rt, tariff_hi(g, sub, warn),
                                   (v.get("ncb") or spec.ncb or "NA") if sub in ("pkg", "saod") else "NA", _old_tnc(v["tnc"], old_ok), ctx.dates, spec.flags + v.get("flags", []) + g.get("flags", []), need)
                    row["is_cpa"] = spec.cpa if sub in ("pkg", "saod") else "NA"
                    row["is_zd"] = zd if sub != "stp" else "NA"      # STP has no OD cover: no NCB, no zero dep (the user, 27 Oct 2026) - the 'NCB cases' / 'Without NCB' pair of a P & L row then meets on one STP key and the lowest rate wins
                    out.append(row)
    return out


def with_main(rate, value):
    """the (pi_type, od, tp, net) rate with the payin of its main part replaced; None for 'On OD and TP', which has two parts"""
    typ, od, tp, net = rate
    if typ == "On Net":
        return typ, od, tp, value
    if typ == "On TP":
        return typ, od, value, net
    if typ == "On OD":
        return typ, value, tp, net
    return None


def split_cc(variants, excl):
    """'EXCLUDING 1001 CC TO 1500 CC': every variant's cc band becomes the piece below the range and the piece above it (0-1000.01 and 1500.01-99999); a band that does not reach the range stays as it is"""
    if not excl:
        return variants
    below, above = round(excl[0] - 1 + 0.01, 2), round(excl[1] + 0.01, 2)
    out = []
    for v in variants:
        lo, hi = v["cc"]
        if lo < below:
            out.append(dict(v, cc=(lo, min(hi, below))))
        if hi > above:
            out.append(dict(v, cc=(max(lo, above), hi)))
    return out


def _include(spec, ctx):
    inc = None
    if spec.part_include:
        inc = set()
        for n in spec.part_include:
            inc |= ctx.masters.codes(n)
    if spec.code_include:
        inc = set(spec.code_include) if inc is None else inc & set(spec.code_include)
    return inc


def _exclude(spec, ctx):
    exc = set()
    for n in spec.part_exclude:
        exc |= ctx.masters.codes(n)
    return exc or None


MAKE_TABLE = {            # (product kind or family, phrase key) -> make group of the portal
    ("pc", "hhk"): MAKES_PC_BUNDLE, ("pc", "hhk_citroen"): "HONDA,Hyundai,KIA,CITROEN", ("pc", "citroen"): "CITROEN",
    ("pc", "hyundai_maruti_ex_alto"): "Hyundai,Maruti_EX_ALTO", ("pc", "alto_declined"): "SHRIRAM_PC_EX_DECLINE_ALTO",
    ("pc", "eeco_declined"): "SHRIRAM_PC_EX_DECLINE_EECO", ("pc", "renault_ex"): "SHRIRAM_PC_EX_DECLINE_RENAULT",
    ("pc", "ex_hyundai_maruti"): "SHRIRAM_PC_EX_DECLINE_MARUTI_HYUNDAI", ("pc", "ex_maruti"): "SHRIRAM_PC_EX_DECLINE_MARUTI",
    ("pc", "mm_only"): "Mahindra", ("pc", "only_eeco"): "EECO", ("pc", "ex_alto_swift"): "SHRIRAM_PC_EX_DECLINE_ALTO_SWIFT",
    ("pc", "hhtmt"): "Hyundai,HONDA,TATA,Mahindra,Toyota", ("pc", "ex_maruti_mm"): "SHRIRAM_PC_EX_DECLINE_MAHINDRA_MARUTI",
    ("pc", "ex_mm"): "SHRIRAM_PC_EX_DECLINE_MAHINDRA", ("pc", "maruti_only"): "Maruti", ("pc", "maruti_ex_swift_xl6"): "MARUTI_EX_SWIFT_XL6",
    ("taxi", "mshb"): MAKES_TAXI_MSHB, ("taxi", "ex_mshb"): MAKES_TAXI_EX_MSHB,
    ("taxi06", "eeco_declined"): "SHRIRAM_PCV_4W_0-7_JUNE26_EX_EECO", ("taxi06", "bolero_ex"): "SHRIRAM_PCV_4W_0-7_JUNE26_EX_BOLERO",
    ("taxi710", "allowed_list"): "Magic_Maximo_Dost_Bolero_Ertiga_Innova_Supro_Carens_Marrazo_Invicto",
    # the groups the user chose by hand in the November 2026 upload (flagged rows of the first build)
    ("taxi06", "mm_ex"): "SHRIRAM_PCV_4W_0-7_EX_MAHINDRA_JUNE26", ("taxi06", "dzire_ex"): "SHRIRAM_PCV_4W_0-7_JUNE26_EX_DZIRE",
    ("taxi", "dzire_only"): "SWIFT_DZIRE", ("taxi", "ambasador"): "Ambasador",
    ("pc", "mm_ex"): "SHRIRAM_PC_EX_DECLINE_MAHINDRA",
}
MAKES_PC_EX_MARUTI_MAHINDRA = "SHRIRAM_PC_EX_DECLINE_MARUTI_MAHINDRA"       # 'Others makes (EXCEPT MARUTI make)' next to Mahindra's own row (the user's group; the master's older one holds 226 entries, this one every make but the two)


def table_makes(kind, spec, flags):
    """The make group a phrase of the remarks names for this product, or None (a phrase with no group is flagged)."""
    if not spec.make_key:
        return None
    fam = "pc" if kind.startswith("pc") else "taxi" if kind.startswith("taxi") else kind
    got = MAKE_TABLE.get((kind, spec.make_key)) or MAKE_TABLE.get((fam, spec.make_key))
    if not got:
        flags.append(f"{spec.make_text or spec.make_key} - no make group for it in this script, so it is not applied")
    return got


def kind_variants(g, kind, band, spec, age, sub, ctx):
    """The (makes, class, T&C, fuels, cc, sc ...) variants a grid row stands for."""
    flags = []
    tnc_extra = list(spec.tnc)
    if kind in ("pc_petrol", "pc_diesel"):
        fuels = ["Diesel"] if kind == "pc_diesel" else (FUELS_NON_DIESEL if ctx.write_ev else [f for f in FUELS_NON_DIESEL if f != "Electric"])
        if kind == "pc_diesel" and spec.fuels:
            fuels = spec.fuels
        car = dict(product="Private Car", mmc="Car", tnc=join_tnc(*tnc_extra), fuels=fuels, cc=(0, CC_OPEN_END), sc=(0, 999), age=(0, AGE_OPEN_END) if sub == "b13" else None, flags=flags)
        if spec.mm_gt:                                    # Mahindra at any cc, every other make but Maruti above the limit (the row 'OTHER THAN M &M' next to it stops at the limit)
            return [dict(car, makes=MAKE_TOKENS["MAHINDRA"]), dict(car, makes=MAKES_PC_EX_MARUTI_MAHINDRA, cc=(round(spec.mm_gt + 0.01, 2), CC_OPEN_END))]
        makes = table_makes(kind, spec, flags) or (MAKES_PC_BUNDLE if sub == "b13" else MAKES_PC)
        if spec.make_key == "ex_mm" and ctx.mm_limit(g, kind):
            car["cc"] = (0, round(ctx.mm_limit(g, kind) + 0.01, 2))
        makes = ctx.carve_pc(g, kind, sub, makes)         # a row of the state that prices EECO alone: this row leaves it out
        return [dict(car, makes=makes)]
    if kind in ("taxi06", "taxi710"):
        makes = table_makes(kind, spec, flags) or (MAKES_TAXI_06 if kind == "taxi06" else MAKES_TAXI_710)
        tnc_parts = []
        if spec.taxi == "ex_kp":
            tnc_parts.append("EX:KALLI PEELI")
        elif spec.taxi == "kp":
            tnc_parts.append("KALLI PEELI & Blue Cab")
        tnc_parts += tnc_extra
        tnc_parts.append(TAXI_TNC)
        fuels = spec.fuels or (FUELS_ALL5 if ctx.write_ev else FUELS_NO_EV)
        sc = (0, 7.01) if kind == "taxi06" else (7.01, 10.01)
        return [dict(product="PCV 4W", mmc="Taxi", makes=makes, tnc=", ".join(tnc_parts), fuels=fuels, cc=(0, CC_OPEN_END), sc=sc, flags=flags)]
    if kind == "erick":
        return [dict(product="PCV 3W", mmc="Passenger Carrying Rickshaw", makes=MAKES_3W_BOTH, tnc=join_tnc(*tnc_extra), fuels=["Electric"], cc=(0, CC_OPEN_END), sc=(0, 999)),
                dict(product="GCV 3W", mmc="Goods Carrying Rickshaw", makes=MAKES_3W_GCV, tnc=join_tnc(*tnc_extra), fuels=["Electric"], cc=(0, CC_OPEN_END), sc=(0, 999))]
    if kind in ("p3w_diesel", "p3w_other"):
        fuels = ["Diesel"] if kind == "p3w_diesel" else ["Petrol", "CNG", "LPG"]
        if spec.include:
            makes = ",".join(MAKE_TOKENS.get(x, x) for x in spec.include)
        elif "BAJAJ" in spec.exclude:
            makes = "PCV_3W_ALL_EX_BAJAJ"
        else:
            makes = MAKES_3W_PCV
        return [dict(product="PCV 3W", mmc="Passenger Carrying Rickshaw", makes=makes, tnc=join_tnc(*tnc_extra), fuels=spec.fuels or fuels,
                     cc=(0, CC_OPEN_END), sc=spec.sc or (0, 999))]
    if kind == "g3w":
        return [dict(product="GCV 3W", mmc="Goods Carrying Rickshaw", makes=MAKES_3W_GCV, tnc=join_tnc(*tnc_extra), fuels=spec.fuels or ["Petrol", "Diesel", "CNG", "LPG"],
                     cc=(0, CC_OPEN_END), sc=(0, 999))]
    if kind == "tractor":
        new = age["new"]
        tnc = spec.tnc_text or ("NEW" if new else ("Old" if spec.old else ""))
        rng = (0, 1) if new else ((1, (age["pkg"] or (0, 15.01))[1]) if spec.old else None)
        return [dict(product="MISCD", mmc="MISCD-Tractor", makes=MAKES_MISCD, tnc=join_tnc(tnc, *tnc_extra), fuels=spec.fuels or FUELS_ALL5,
                     cc=(0, CC_OPEN_END), sc=(0, 999), age=rng)]
    if kind == "miscd_others":
        tnc = spec.tnc_text or {"ex": "EX:Crane,Tractor,Ambulance", "ex_other": "EX:Crane,Tractor, Ambulance & other than Combined Harvester",
                                "harvester": "Combined Harvester /-"}.get(spec.misc, "")
        tnc = add_declined_bodies(tnc, ctx.declined_bodies.get("miscd_others", []))                # the body types the sheet 'Declined vehicle' declines for MISC-D others
        return [dict(product="MISCD", mmc="MISCD-Others", makes=MAKES_MISCD, tnc=join_tnc(tnc, *tnc_extra), fuels=spec.fuels or FUELS_ALL5, cc=(0, CC_OPEN_END), sc=(0, 999))]
    if kind == "newmah":
        return [dict(product="GCV 4W", mmc="NA", makes=MAKE_TOKENS["MAHINDRA"], tnc="New", fuels=FUELS_ALL5, cc=(0, CC_OPEN_END), sc=(0, 999), age=(0, 1))]
    if kind == "gcv4w":
        carve = "BOLERO" in ctx.alone(g, kind, sub, band)                    # a row of the state that prices the Bolero models alone: the Mahindra row leaves them out
        vs, problems = gcv_variants(spec, carve=carve)
        if carve and spec.mahindra_split and not spec.bolero_ex and not spec.makes_override and any(v[0] == MAKES_MAHINDRA_EX_BOLERO for v in vs):
            ctx.carved.setdefault("BOLERO", set()).add(g["r"])
        for p in problems:
            ctx.warn(f"row {g['r']}: {p} - written with the base group")
        classes = ["NA"]
        if spec.class_only:
            classes = list(spec.class_only)
        elif spec.classes == "tanker":
            classes = ["Tanker"]
        elif spec.classes == "nontanker":
            classes = list(GCV_CLASSES_OTHER)
        out = []
        for makes, prefix, floor in vs:
            for mmc in classes:
                if spec.new_or_ncb:                      # new vehicles (any make but new Mahindra) + used vehicles with NCB
                    if not prefix:
                        out.append(dict(product="GCV 4W", mmc=mmc, makes=makes, tnc=join_tnc("New", *tnc_extra), fuels=spec.fuels or FUELS_ALL5, cc=band,
                                        sc=(0, 999), age=(0, 1), ncb="NA"))
                    out.append(dict(product="GCV 4W", mmc=mmc, makes=makes, tnc=join_tnc("Old", *tnc_extra), fuels=spec.fuels or FUELS_ALL5, cc=band,
                                    sc=(0, 999), floor=1, ncb="YES"))
                    continue
                out.append(dict(product="GCV 4W", mmc=mmc, makes=makes, tnc=join_tnc(prefix, *tnc_extra), fuels=spec.fuels or FUELS_ALL5, cc=band, sc=(0, 999),
                                floor=floor))
        if age["new"]:
            for o in out:
                o["age"] = (0, 1)
                o["tnc"] = join_tnc("New", o["tnc"])
        return out
    return []


def convert_bus(g, kind, subs, ctx):
    """Pan-India school bus / corporate bus: the payout depends on the seating capacity (and, for corporate buses, on metro / non-metro RTOs)."""
    t = norm(g["po"])
    out = []
    age = parse_age(g["age"])
    if kind == "school":
        m = re.search(r"SC\s*>\s*15\s*PO\s*(\d+(?:\.\d+)?)\s*%\s*&\s*<\s*15\s*(\d+(?:\.\d+)?)\s*%", t, re.I)
        plan = [(ALL_INDIA, "School Bus", float(m.group(1)), float(m.group(2)))] if m else None
    else:
        m = re.search(r"metro\s*-\s*(\d+(?:\.\d+)?)\s*%\s*SC\s*>=\s*15\s*,\s*(\d+(?:\.\d+)?)\s*%\s*SC\s*<\s*15\s*:\s*non\s*metro\s*-\s*(\d+(?:\.\d+)?)\s*%\s*SC\s*>=\s*15\s*,\s*(\d+(?:\.\d+)?)\s*%\s*SC\s*<\s*15", t, re.I)
        plan = ([(METRO_GROUP, "Other Bus", float(m.group(1)), float(m.group(2))), (NON_METRO_GROUP, "Other Bus", float(m.group(3)), float(m.group(4)))] if m else None)
    if not plan:
        ctx.skipped[f"bus payout '{t[:40]}'"] = ctx.skipped.get(f"bus payout '{t[:40]}'", 0) + 1
        return []
    spec = read_remarks(kind, g)
    work = g.get("rto", "")
    if work and len(plan) > 1:                                       # metro and non-metro pay differently: one name cannot say both, the master groups stay
        working_unused(g["sheet"], g["r"], "corporate bus: two RTO scopes, metro and non-metro, with different payouts")
        work = ""
    for rto, mmc, hi, lo in plan:
        if work:
            rto = RtoScope(working_rto(work))
        for sub in subs:
            if sub not in ("pkg", "stp"):
                continue
            a = (age["stp"] if sub == "stp" else age["pkg"]) or (0, AGE_OPEN_END)
            for sc, rate_value in (((0, 15.01), lo), ((15.01, 999), hi)):
                rate = rate_for(("plain", rate_value, None), sub, kind)
                for fuel in FUELS_NO_EV:
                    out.append(make_row(g, sub, fuel, "PCV 4W", mmc, MAKES_BUS, rto, a, (0, CC_OPEN_END), sc, rate, tariff_hi(g, sub, ctx.warn), "NA", "",
                                        ctx.dates, spec.flags, ()))
    return out


# %% ---------------------------------------------------------------- BUILD
# Electric vehicles: the sheet 'Conditions' says "EV PRIVATE CAR / EV 4W PCV DECLINED - PAN INDIA".  False = follow the document: no Electric rows for
# Private Car and for taxis (the grid's own 'Electric Vehicles only' STP row is written).  True = write Electric with the other non-diesel fuels, as the
# earlier uploads did.  --write-ev switches it on for one run.
WRITE_DECLINED_EV = False
LAST_GAPS = {}            # RTO groups the master lacks for this run: message -> the RTO codes the group would hold (a sheet of the review workbook, see collect_gaps)


class Ctx:
    """What convert_row needs to know about this run."""

    def __init__(self, warnings, dates, declines, masters, write_ev):
        self.warn = warnings.append
        self.warnings = warnings
        self.skipped = {}
        self.dates = dates
        self.declines = declines
        self.masters = masters
        self.rtos = Rtos(masters)
        self.write_ev = write_ev
        self.hazard_unsplit = set()                       # grid rows that mention the hazardous reduction but do not separate tankers
        self.specs = {}                                   # grid row -> what its remarks ask for (read once, by prepare)
        self.only = {}                                    # (state, kind, make) -> [(sub-products, weight band)] of the rows that price that make on its own
        self.mm_limits = {}                               # (state, kind) -> the cc limit of Mahindra's own car row
        self.carved = {}                                  # make -> the grid rows that took it out of their make group (said in a warning)
        self.declined_bodies = {}                         # kind -> the body types the sheet 'Declined vehicle' declines for that product (written into the exclusion T&C)

    def prepare(self, grid):
        """Read the remarks of every row once and note what the rows of a state say about each other: a row that prices ONE make alone ('Only EECO', 'Bolero Model only') and the cc limit of
        'M & M (All CC) & Others makes - Greater than 1500 CC' (the user did both by hand in the November 2026 upload)."""
        for g in grid:
            kind, band = classify(g["product"])
            if kind is None or kind in ("corp", "school"):
                continue
            spec = self.specs[g["r"]] = read_remarks(kind, g)
            subs = {s for s in (policy_subs(g["ptype"]) or []) if s not in spec.skip and s not in g.get("drop", ())}
            area = label_key(g["state"])
            if kind.startswith("pc") and spec.make_key == "only_eeco":
                self.only.setdefault((area, kind, "EECO"), []).append((subs, None))
            if kind == "gcv4w" and spec.makes_override == MAKES_BOLERO:
                self.only.setdefault((area, kind, "BOLERO"), []).append((subs, band))
            if spec.mm_gt:
                self.mm_limits[(area, kind)] = spec.mm_gt

    def spec_for(self, g, kind):
        spec = self.specs.get(g["r"])
        if spec is None:
            spec = self.specs[g["r"]] = read_remarks(kind, g)
        return spec

    def alone(self, g, kind, sub, band):
        """the makes ('EECO', 'BOLERO') that another row of this state and product prices on its own for this sub-product (and a weight band that overlaps this one)"""
        out = set()
        for (area, k, make), entries in self.only.items():
            if area == label_key(g["state"]) and k == kind:
                for subs, b in entries:
                    if sub in subs and (b is None or band is None or (b[0] < band[1] and band[0] < b[1])):
                        out.add(make)
        return out

    def carve_pc(self, g, kind, sub, makes):
        """the make group of a car row without the make another row of the state prices on its own (CARVE_GROUPS); the group itself when there is no such make or no group for it"""
        for make in sorted(self.alone(g, kind, sub, None)):
            group = CARVE_GROUPS.get(("pc", make), {}).get(makes)
            if group:
                self.carved.setdefault(make, set()).add(g["r"])
                return group
        return makes

    def mm_limit(self, g, kind):
        return self.mm_limits.get((label_key(g["state"]), kind))


MONTHS = {m: i for i, m in enumerate(("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"), 1)}


def check_effective(grid, from_date, warnings):
    """The EFFECTIVE DATE column says 'w.e.f 1st Oct': say so if from_date is in another month (P-12)."""
    months = {}
    for g in grid:
        for m in re.findall(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b", g["eff"].lower()):
            months[MONTHS[m]] = months.get(MONTHS[m], 0) + 1
    for mon, n in sorted(months.items()):
        if mon != from_date.month:
            warnings.append(f"{n} grid row(s) say 'w.e.f' month {mon} but from_date is {from_date:%d/%m/%Y} - check the dates")


def build(grid_path, from_date, to_date, log=say, ask=None, wb=None, write_ev=None, addendum_path=None):
    """Read the workbook and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    from_date / to_date are datetime.date values; they are written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever the grid itself is missing something. The remarks are never asked about: what they say is applied, and
    what cannot be applied is flagged in the remarks column (see ManualReview)."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    working_reset()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    if wb is None:
        log(f"Reading {os.path.basename(str(grid_path))}...")
        grid_path, wb = open_file(ask, grid_path, "the Shriram workbook", read_workbook, EXCEL_EXTS)
    log(f"Sheets in the workbook: {wb.sheetnames}")
    TWIN_ENTRIES.clear()
    grid, warnings = read_grid(wb, ask, log)
    if not grid:
        raise GridError("the rate sheet has no rate rows")
    if addendum_path:
        add_rows, add_warnings = read_addendum(addendum_path, ask, log)
        warnings.extend(add_warnings)
        grid = apply_addendum(grid, add_rows, warnings, log)
        log(f"  (row numbers above {ADDENDUM_ROW_BASE} are rows of the addendum: number - {ADDENDUM_ROW_BASE})")
    else:
        log("No addendum given (--addendum <file>): if Shriram sent one for this month, its rows override grid rows.")
    declines = Declines()
    declines.read(wb, warnings, ask, log)
    masters = load_masters(ask)
    ctx = Ctx(warnings, dates, declines, masters, WRITE_DECLINED_EV if write_ev is None else write_ev)
    ctx.prepare(grid)
    decline_notes = decline_product_notes(wb)
    for text, label, bodies in decline_notes:                                  # 'Decline Product [MISC.D - OTHERS (BODY TYPE GARBAGE CARRIER OR GARBAGE COMPACTOR)]'
        kind = classify(label.replace(".", ""))[0]
        if kind in DECLINE_PRODUCT_ROWS and bodies:
            known = ctx.declined_bodies.setdefault(kind, [])
            known.extend(b for b in bodies if b not in known)
    review = ManualReview()
    if TWIN_ENTRIES:
        add_sheet("Twin columns", pd.DataFrame(TWIN_ENTRIES))
    check_effective(grid, from_date, warnings)
    haz = sum(1 for g in grid if re.search(r"hazardous\s+category", g["uw"] + " " + g["rem1"], re.I))
    if haz and HAZARDOUS_MODE == "tanker":
        warnings.append(f"{haz} grid rows say 'For Hazardous category PO will be 5% less': applied to the Tanker class (payin {HAZARDOUS_CUT:g} points lower; a row for all classes is "
                        f"split into Tanker and the other classes) because the hazardous category of a policy is not in the MIS - check the Tanker rows")
    elif haz and HAZARDOUS_MODE == "tanker_rows":
        pass                                            # the count of the rows left as listed is added after the rows are built
    elif haz and HAZARDOUS_MODE == "flag":
        warnings.append(f"{haz} grid rows say 'For Hazardous category PO will be 5% less': payouts are written as listed and those rows carry the note in the remarks column "
                        f"(the hazardous category of a policy is not in the MIS; HAZARDOUS_MODE = 'tanker_rows' / 'tanker' reduce the Tanker class by {HAZARDOUS_CUT:g} points)")
    elif haz and HAZARDOUS_MODE == "tnc":
        warnings.append(f"{haz} grid rows say 'For Hazardous category PO will be N% less': the text 'Hazardous category PO N% less' is written in add_tnc of those rows (payouts as listed, no flag; "
                        f"the hazardous category of a policy is not in the MIS)")
    elif haz and HAZARDOUS_MODE == "off":
        warnings.append(f"{haz} grid rows say 'For Hazardous category PO will be 5% less': not applied (HAZARDOUS_MODE = 'off')")
    log(f"Declined RTOs: {len(declines.entries)} entries in '{DECLINED_SHEET}'; pan-India: "
        + ", ".join(f"{k} {len(v)}" for k, v in sorted(declines.pan.items())) + f"; whole states declined: {sorted(declines.whole)}")

    rows = []
    for g in grid:
        rows.extend(convert_row(g, ctx))
    if HAZARDOUS_MODE == "tanker_rows" and haz:
        warnings.append(f"{haz} grid rows say 'For Hazardous category PO will be 5% less': the Tanker rows of the rows that name tankers are {HAZARDOUS_CUT:g} points lower; "
                        f"{len(ctx.hazard_unsplit)} rows mention it without naming tankers and are written as listed (the hazardous category is not in the MIS; "
                        f"HAZARDOUS_MODE = 'tanker' would split them into Tanker and the other classes)")
    for what, n in sorted(ctx.skipped.items()):
        warnings.append(f"{n} rate row(s) with {what}: not written (this script does not know it)")
    for make, rows_ in sorted(ctx.carved.items()):
        warnings.append(f"{len(rows_)} grid row(s) ({', '.join(str(r) for r in sorted(rows_)[:8])}{' ...' if len(rows_) > 8 else ''}) leave {make} out of their make group because another row of the same "
                        f"state prices {make} on its own - the way the November 2026 upload was corrected by hand (CARVE_GROUPS)")
    note = working_note()
    if note:
        warnings.append(note)
        log(note)
    if label_key(SHORT_TERM_SHEET) in {label_key(n) for n in wb.sheetnames}:
        warnings.append(f"the sheet '{SHORT_TERM_SHEET}' is not converted: the portal has no short-term sub_product")
    known = {x for k in ("pc_petrol", "pc_diesel", "taxi06", "taxi710", "g3w", "p3w_diesel", "p3w_other", "tractor") for x in decline_labels(k)}
    known |= {x for b in ((0, 2800.01), (2800.01, 3500.01), (3500.01, 7500.01), (7500.01, 12000.01), (12000.01, 20000.01), (50000.01, CC_OPEN_END)) for x in decline_labels("gcv4w", b)}
    for lab, text in sorted(declines.labels.items()):
        if lab not in known:
            warnings.append(f"the declined list for '{text}' is not used by any product of this script")
    mc = metro_codes(wb)
    if mc and masters.codes(METRO_GROUP):
        have = masters.codes(METRO_GROUP)
        if mc - have:
            warnings.append(f"the sheet '{METRO_SHEET}' lists {len(mc - have)} RTO(s) that the master group {METRO_GROUP} lacks (e.g. {_codes_text(sorted(mc - have)[:6])})")
        extra = {c for c in have - mc if c[:2] != "DL"}                    # the sheet says 'DL RTO codes' without listing them
        if extra:
            warnings.append(f"the master group {METRO_GROUP} holds {len(extra)} RTO(s) that the sheet '{METRO_SHEET}' does not list (e.g. {_codes_text(sorted(extra)[:6])})")
    if not rows:
        raise GridError("nothing to write - no grid row could be converted")

    df = pd.DataFrame(rows)
    for c in ("_flags", "_need"):
        df[c] = df[c].map(tuple)
    keys = [c for c in KEY_COLUMNS if c not in ("tariff_min", "tariff_max")]
    df["tariff_min"] = df["tariff_min"].astype(float)
    df["tariff_max"] = df["tariff_max"].astype(float)
    df = settle_duplicates(df, keys + ["tariff_max"], warnings)        # one grid row twice: once; two grid rows, one key, two payouts: the lowest
    pk = df["sub_product"] != "STP"
    ck = [c for c in keys if c != "add_tnc"]
    slabs = df.loc[pk, ck + ["tariff_max"]].drop_duplicates().sort_values(ck + ["tariff_max"], kind="stable")
    slabs["tariff_min"] = slabs.groupby(ck, sort=False)["tariff_max"].shift(1).fillna(0.0)    # a higher discount slab starts where the lower one ends
    df = df.drop(columns=["tariff_min"]).merge(
        pd.concat([slabs, df.loc[~pk, ck + ["tariff_max"]].drop_duplicates().assign(tariff_min=0.0)]), on=ck + ["tariff_max"], how="left")
    df = df.reset_index(drop=True)

    flag_texts = {}
    for i, t in enumerate(zip(df["_flags"], df["_need"])):
        for text in t[0] + t[1]:
            flag_texts.setdefault(text, []).append(i)
    for text, label, bodies in decline_notes:                             # the line is written into the exclusion T&C (above); one this script cannot write that way is flagged
        rows_of = DECLINE_PRODUCT_ROWS.get(classify(label.replace(".", ""))[0])
        if rows_of is None:
            warnings.append(f"the sheet 'Declined vehicle' says '{text}' - this script does not know that product, so no row carries the line")
            continue
        hit = df.index[(df["product"] == rows_of[0]) & (df["make_model_class"] == rows_of[1])].tolist()
        if bodies:
            n = int(df.loc[hit, "add_tnc"].str.contains(bodies[0], regex=False).sum())
            warnings.append(f"'Declined vehicle' says '{text}': {', '.join(bodies)} written into the exclusion T&C (EX:...) of {n} of the {len(hit)} {label.title()} rows; the {len(hit) - n} rows for one body type only are left as they are")
        elif hit:
            flag_texts.setdefault(f"{text} - a declined body type is not applied by this script", []).extend(hit)
    for text, idx in sorted(flag_texts.items(), key=lambda kv: -len(kv[1])):
        mask = np.zeros(len(df), dtype=bool)
        mask[idx] = True
        review.flag(_short(text), mask)
    df = df.drop(columns=["_flags", "_need", "_r", "_sub"])
    global LAST_GAPS
    LAST_GAPS = dict(ctx.rtos.gaps)
    globals()["MASTER_SHEET"] = " + ".join(sorted({g["sheet"] for g in grid if g.get("sheet")})) or "the grid"          # the sheet the make conditions come from, for the new-mapping text
    check_names(df, review, masters, log)                                # every RTO / make group written must be in the master lists
    review.report(df, log)                                               # fills the remarks column
    return df[OUTPUT_COLUMNS], warnings


def settle_duplicates(df, key_cols, warnings):
    """Rows with the same key: identical payouts are one row; if two grid rows give one entry different payins the LOWEST payin is written (the user's rule, 27 Oct 2026: "if the rate grid
    contains two different rates for the same entry, always select the lower rate").  The rows that lost are dropped and listed in LAST_LOWEST (the shared write_csv writes the list next to the CSV)."""
    global LAST_LOWEST
    LAST_LOWEST = []
    value = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate", "pi_flat_amount"]].astype(float).sum(axis=1)       # 'On OD and TP' adds its two parts; a flat card is its amount
    rate = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate", "pi_flat_amount"]].astype(float).apply(lambda r: "/".join(fmt(x) for x in r if x), axis=1) + "|" + df["pi_type"]
    dup = df.duplicated(key_cols, keep=False)
    if not dup.any():
        return df
    drop = []
    for _, idx in df[dup].groupby(key_cols, sort=False).groups.items():
        idx = list(idx)
        seen = {}
        for i in idx:
            seen.setdefault(rate[i], []).append(i)
        for ids in seen.values():
            drop.extend(ids[1:])                                         # the same payout twice
        if len(seen) > 1:
            firsts = [ids[0] for ids in seen.values()]
            best = min(firsts, key=lambda i: (value[i], idx.index(i)))
            drop.extend(i for i in firsts if i != best)
            row = df.loc[best]
            LAST_LOWEST.append(dict(sub_product=row["sub_product"], rto=row["new_rto_list"], makes=row["new_vehicle_makes"], fuel=row["fuel_type"], make_model_class=row["make_model_class"],
                                    age=f"{fmt(row['vehicle_age_min'])}-{fmt(row['vehicle_age_max'])}", cc=f"{fmt(row['cc_min'])}-{fmt(row['cc_max'])}", tariff_max=fmt(row["tariff_max"]),
                                    grid_rows=", ".join(str(r) for r in sorted({df.at[i, "_r"] for i in idx})), rates_found=" / ".join(f"{fmt(value[ids[0]])}" for ids in seen.values()),
                                    written=fmt(value[best])))
    if LAST_LOWEST:
        warnings.append(f"{len(LAST_LOWEST)} combinations got two different payins from different grid rows: the LOWEST is written (listed in the review workbook next to the upload)")
    return df.drop(index=drop)


def _short(text):
    t = re.sub(r"\s+", " ", str(text)).strip()
    return t if len(t) <= MAX_NOTE_CHARS else t[:MAX_NOTE_CHARS - 3] + "..."


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
    """the comparison key of a frame of upload rows: numbers to 4 decimals, make and RTO names without regard to capitals ('MAHINDRA' of an old upload = 'Mahindra' of the master)"""
    d = d.copy()
    for c in ("vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "sc_min", "sc_max", "tariff_min", "tariff_max"):
        d[c] = pd.to_numeric(d[c], errors="coerce").map(lambda x: f"{x:.4f}")
    for c in ("new_vehicle_makes", "new_rto_list"):
        d[c] = d[c].astype(str).str.lower()
    d["new_rto_list"] = d["new_rto_list"].str.replace(r" - condition: .*$", "", regex=True)      # the comparison is on the cluster name: an earlier upload has no condition text
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
    heads = lambda s: {x.split(" - Condition: ")[0] for x in s}               # the cluster name without its condition text
    cl_new = sorted(heads(mine["new_rto_list"]) - heads(ref["new_rto_list"]))
    cl_old = sorted(heads(ref["new_rto_list"]) - heads(mine["new_rto_list"]))
    if cl_new:
        log(f"    RTO lists not in the earlier upload ({len(cl_new)}): {cl_new[:6]}{' ...' if len(cl_new) > 6 else ''} - their RTO groups must exist in the portal")
    if cl_old:
        log(f"    RTO lists in the earlier upload but not in this grid ({len(cl_old)}): {cl_old[:6]}{' ...' if len(cl_old) > 6 else ''}")
    return len(only_new), len(only_ref), int(changed.sum())


def collect_gaps(df):
    """For the review workbook: the RTO groups that have to be created in the portal's master (one row each, with the RTO codes to put in them)."""
    if not LAST_GAPS or RTO_NAMES in ("cluster", "master"):         # cluster mode: the condition is written in the RTO cell itself, there is no group to create
        return
    rows = []
    for text, codes in LAST_GAPS.items():
        full = f"No RTO group in the master for {text} - written for the whole area"
        rows.append(dict(group_needed=text, rows_flagged=int(df["remarks"].str.contains(full, regex=False).sum()), items=",".join(codes)))
    add_sheet("RTO groups to create", pd.DataFrame(rows).sort_values("rows_flagged", ascending=False))


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


def main(grid_path, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None, write_ev=None, addendum_path=None):
    ask = ask or Asker()
    df, warnings = build(grid_path, from_date, to_date, log, ask, write_ev=write_ev, addendum_path=addendum_path)
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
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"SHRIRAM_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
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
        if not ask.confirm("No workbook (.xlsx) came with that upload. Upload again?", True):
            raise GridError("no workbook was uploaded")
    def looks_like_grid(path):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                return any(label_key(s).startswith("grid") or label_key(s) in ("declinedexcluded", "panindia") for s in wb.sheetnames)
            finally:
                wb.close()
        except Exception:
            return False
    mains = [b for b in books if looks_like_grid(b)] or books
    grid = mains[0] if len(mains) == 1 else ask.choose(
        "more than one workbook was uploaded", [(n, n) for n in mains],
        question="You uploaded more than one workbook. Which one is the Shriram broker grid?", short="workbook")
    others = [b for b in books if b != grid]
    addendum = others[0] if others else None                      # a second workbook is the addendum (its rows override grid rows)
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, addendum, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the Shriram broker grid (.xlsx), the month's addendum (.xlsx) if Shriram sent one, and, optionally, an earlier upload (.csv) to compare with")
    try:
        grid, addendum, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", addendum = {addendum}" if addendum else "") + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask, addendum_path=addendum)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("shriram_broker_grid_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="Shriram broker grid -> Shriram upload CSV")
    ap.add_argument("--grid", help="Shriram workbook (.xlsx) - asked for if left out")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--addendum", help="Shriram's addendum workbook for the month (the same columns as the rate sheet): its rows override grid rows")
    ap.add_argument("--write-ev", action="store_true", help="also write Electric rows for Private Car and taxis (the grid's Conditions sheet says EV is declined)")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    try:
        grid = a.grid or ask_file(ask, "no workbook given - pass --grid <workbook.xlsx>",
                                  "Path of the Shriram broker grid (.xlsx) - you can drag the file into this window:", exts=EXCEL_EXTS)
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
        addendum = clean_path(a.addendum) if a.addendum else None
        if addendum and not os.path.isfile(addendum):
            raise GridError(f"addendum not found: {addendum}")
        main(grid, from_date, to_date, a.out, ref, ask=ask, write_ev=True if a.write_ev else None, addendum_path=addendum)
    except (GridError, ValueError) as e:
        sys.exit(f"STOPPED - {e}")
    except KeyboardInterrupt:
        sys.exit("STOPPED - stopped by you. No file was written.")
    except Exception as e:
        saved = crash_report("shriram_broker_grid_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
