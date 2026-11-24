"""
Go Digit General Insurance agency grid -> Go Digit upload CSV
=============================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (NN).csv") from the workbook Go Digit circulates for the agency channel
("Delhi_PAN_India_Agency.xls", password protected).  Rates come only from that workbook.

What the workbook holds and how it becomes rows (checked against the user's October 2026 upload, 47,532 rows, from_date 15/10/2026)
-----------------------------------------------------------------------------------------------------------------------------------
Hidden sheets are older versions or helpers and are never read; the visible sheets are the grid of the month.

How the workbook is read (10 Nov 2026)
--------------------------------------
Nothing is found by a sheet's NAME or by counting cells from a header.  Go Digit renamed every sheet of the November workbook ("LATEST CV COMBINED GRID" -> "CV Grid", "Latest 2W 1+1 & SATP" ->
"New 2W Comp & SATP wef 1st Nov", "PVT Car Comp+SAOD" -> "New 4W Comp_SAOD", " PVT Car TP" -> "4W TP Consolidated" ...) and put new columns in the middle of tables, so every table is found by what its
header row SAYS (Spec / find_table: aliases per field, a typo is read as the field and said so, the sheet name only breaks a tie) and every column inside it by its own header (block_ranges: from a
'Cluster' header to the next one).  Titles above the columns ('1+1' / 'SATP', 'With CPA' / 'Without CPA', 'SAOD' / 'Comp', 'COMPREHENSIVE' / 'THIRD-PARTY') tell the groups apart, never their order.
Two payout columns ('Max CD2' beside 'Avg CD2') are decided by PAYOUT_MEASURE (the user's answer of 10 Nov 2026: 'Avg CD2'; None = a numbered question, choose_payout), a volume slab the sheet no longer has is a question, a table that is not there stops the run with what was expected and
what the workbook has (table_missing), a header nobody reads is listed in the warnings (warn_unread).  The sheet names in the text below are those of the October workbook.

* "LATEST CV COMBINED GRID"  (commercial vehicles; one grid row = one cluster x segment x body type x make)
    Cluster segment | Segment | Body Type | RTO Cluster | State | Make / Variant | CD1 (OD-disc) | Comprehensive CD2 by vehicle age (New, 1 Yr ... 11+ Yr) | Third-party / SATP CD2 by age
    CD1 is the OD discount the payout is valid for, CD2 is the payout in percent (a fraction in the cell: 0.325 = 32.5).
    Cluster segment -> product / RTO group / make group:
      CV HCV -> GCV 4W (12 t and above; class Truck / Dumper Tipper / Tanker / Trailer from the body type), CV LCV, CV MCV -> GCV 4W, CV GCV3 -> GCV 3W (E-Loaders = Electric, with/without CPA),
      CV MiscD / CV Tractor -> MISCD, CV PCV2W(All) 3W Non-electric -> PCV 2W and PCV 3W (the same rates for both), PCV 3W Electric -> PCV 3W, CV PCV4 (Upto 5 seating) -> PCV 4W Taxi,
      Staff Bus -> PCV 4W Other Bus.  (School Bus: see WRITE_SCHOOL_BUS.)
    The RTO group is  GO-DIGIT_OCT26_<segment>_<cluster>  (spaces -> '_'; capitals as the portal's master spells it); a cluster whose group the master lacks is flagged and listed with its RTO codes.
    One row per age column and fuel:  New = age 0-1 (add_tnc 'New'), 1 Yr = 1-2 ... 5 Yr = 5-6, 6-10 Yr = 6-10.01, 11+ Yr = 10.01-99.  Comprehensive = 1+1 / On Net, SATP = STP / On TP, po = pi - 7.
    Weight band of the segment text = cc_min / cc_max in kg ('GCV4 12 to 20T' = 12000.01-20000.01).  PO of a 'D' (decline), a blank cell or an error cell (#VALUE!) is 0: a zero card is written so the
    entry is declined rather than falling through to a wider card (a blank cell is a placeholder that never beats a real rate for the same entry).
    CD1 -> tariff band:  '0-60%' / '60%' / 0.85 = tariff_max CD1 + 0.01 (no CD1 = 100.01); 'New/1yr: D | 2-5yr: 80% | 6yr+: 85%/90%' = a cap per age group ('D' = declined = 100.01, an age group the text
    does not name continues the previous cap); '85%/90%' with a CD2 cell '29%/26%' = two bands (0-85.01 at 29, 85.01-90.01 at 26); two grid rows of one entry chain their bands (a higher cap starts where
    the lower one ends); SATP rows are always 0-100.01.  'With Addon: 85%/90% Without Addon: 60%' = is_zd YES / NO; 'with NCB/break in > 90 days : 90%, without NCB : 80%' = is_ncb YES (add_tnc 'Break in > 90 days') / NO.
    Notes of the sheet 'CV Remarks' that the script applies:  'Volvo, Scania to block' (make group without Volvo) and, for Non-Dumper/Tipper (Truck), 'Daimler, Bharat Benz and Mahindra to operate at 70% CD1' (their own rows, cap 70).
    Taxi: 'With Add on cover' / 'Without Add on cover' = is_zd YES / NO on comprehensive rows; SATP rows carry no NCB and no ZD (NA).
* Two-wheelers and Private Car: see the sections further down (TW / PC).
* Whatever text this script does not understand is NEVER guessed at: the rows carry  Manual Review Required: [text]  in the remarks column.
* Two grid rows that give one entry (the same card key) different payins: the LOWEST is written (the user, 27 Oct 2026); each such entry is listed in the sheet 'Lowest rate chosen' of the review workbook.

The workbook is password protected (the file is called .xls but is really an encrypted .xlsx): give the password with --password (or the environment variable GRID_PASSWORD, or type it when asked: the typing
is hidden); it is never stored.  Reading it needs the small package msoffcrypto-tool (a genuine old .xls also needs xlrd); an .xlsx saved without the password needs neither.

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a sheet or header that is not where it should be, dates that look wrong,
a file that will not open or write - it stops, asks you a specific question (type the number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every answer is
listed at the end of the run, and you confirm before the file is written.  It never asks you for a rate.  With --no-prompt (or when there is no keyboard) it asks nothing and stops with the
question as the error message instead.  If something goes wrong that it did not foresee it says so in one line and saves the technical details to a *_error_*.txt file.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run godigit_agency_to_config.py`) and run it.  Upload the workbook (and, optionally, an earlier upload CSV to get a diff report); it asks for
the password, from_date and to_date (DD/MM/YYYY), writes the dates into every row, and downloads GODIGIT_UPLOAD_<yyyy-mm>.csv.  The master workbooks (RTO / make-model) must be in a `master` folder next to the script.

Command line
------------
    python godigit_agency_to_config.py [--grid <workbook.xls>] [--password <pw>] [--from-date 15/10/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the workbook, the password, the dates, the earlier upload - is asked for)
"""
import argparse
import datetime as dt
import difflib
import getpass
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
INSURANCE_COMPANY = "GO DIGIT GENERAL INSURANCE LTD"
USER_ID = 2966
ROW_ID = 1                                       # the portal's export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 2
CC_OPEN_END = 99999
AGE_OPEN_END = 99
SC_RANGE = (0, 999)
TARIFF_FULL = 100.01                             # no discount band / a declined entry / SATP rows
MAX_NOTE_CHARS = 500
HEADER_SCAN_ROWS = 12

GROUP_PREFIX = "GO-DIGIT_OCT26_"                 # the month the portal's groups were created (like Liberty's APR26): change it when the master gets groups under a new prefix
CV_SHEET_WORDS = ("latest", "cv", "combined")    # the visible commercial-vehicle grid: a name with these words, else the sheet whose header looks like it
CV_REMARKS_SHEET_WORDS = ("cv", "remarks")
CV_MAPPING_SHEET_WORDS = ("latest", "cv", "rto", "mapping")
PAYOUT_MEASURE = "avg"                           # the user, 10 Nov 2026: where a table gives BOTH 'Max CD2' and 'Avg CD2' the payin is 'Avg CD2' (it holds the figures of the earlier grids). "max" = Max CD2; None = ask in every run.
                                                 # A table with ONE payout column uses it, whatever its heading (the commercial-vehicle grid has only 'Max CD2').

WRITE_SCHOOL_BUS = False                         # School Bus rows are not in the user's October upload and give two rates per card by premium volume (<= 3 L / > 3 L): listed in 'Not written'
STAFF_BUS_SC = (10.01, 999)                      # the staff bus table is written for MORE than 10 seats (the note 'up to 10-seater' has its own rates, see SEAT_LIMIT_NOTE)
SEAT_LIMIT_NOTE = re.compile(r"seating capacity up to (\d+)\s*-?\s*seater", re.I)
STAFF_BUS_FUELS = ["CNG", "Diesel", "LPG", "Petrol"]    # the sheet says 'Above Grid is applicable for Non-Electric Bus' (the October upload also had Electric rows)
NCB_BREAK_TNC = "Break in > 90 days"
HCV_BLOCKED_MAKES_GROUP = "GCV_4W_ALL_EX_VOLVO"  # 'Volvo, Scania to block' (CV Remarks) -> every HCV row uses the make group without Volvo

STANDING_NOTES = [                               # texts that need no action (the rule is already inside the grid's cells, or the portal handles it)
    "Short Term Policies - CD2 reduction",       # short-term policies are not uploaded
    "CD2 for Short-term policies will be 5% lesser",
    "Refer to the attached excel for RTO Cluster Definition",
    "Discount as per system",
    "System guidelines will be applicable",
    "Buses which are blocked in exception table",
    "Documents required",                        # what the customer must hand in, not a rate rule
    "Permit Copy",
    "Self-declaration from customer",
    "Contract Agreement copy",
    "permit copy will not be required",
    "vehicle invoice/ pre- invoice with confirmation of seating capacity",
    "Above Grid is applicable for Non-Electric Bus",   # applied: STAFF_BUS_FUELS has no Electric
]
STANDING_WHY = {                                 # why each standing note needs no action - shown in the review sheet 'Notes read', so no sentence of the remarks sheet is invisible
    "Short Term Policies - CD2 reduction": "short-term policies are not uploaded (standing decision of the earlier runs)",
    "CD2 for Short-term policies will be 5% lesser": "short-term policies are not uploaded (standing decision of the earlier runs)",
    "Refer to the attached excel for RTO Cluster Definition": "the RTO clusters are read from the workbook's mapping sheets",
    "Discount as per system": "a disclaimer: the discount is the system's",
    "System guidelines will be applicable": "a disclaimer: the system's own guidelines apply",
    "Buses which are blocked in exception table": "a decline the portal's own exception table handles",
    "Documents required": "what the customer must hand in, not a rate rule",
    "Permit Copy": "what the customer must hand in, not a rate rule",
    "Self-declaration from customer": "what the customer must hand in, not a rate rule",
    "Contract Agreement copy": "what the customer must hand in, not a rate rule",
    "permit copy will not be required": "a document rule, not a rate rule",
    "vehicle invoice/ pre- invoice with confirmation of seating capacity": "a document rule, not a rate rule",
    "Above Grid is applicable for Non-Electric Bus": "applied: the staff bus rows are written for the non-electric fuels only",
}
NOTES_READ = []                                  # the sentences of the remarks sheet / the CV sheet's notes and what became of each: the review sheet 'Notes read' (cleared by build)

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

# a row identical (key and payin) to a row of the earlier upload given with --ref keeps that row's from_date
KEEP_REF_DATES = True


OUTPUT_FOLDER = "godigit"            # the folder of this insurer under output\ (default_output)


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


class PasswordNeeded(GridError):
    """The workbook is encrypted and no (or a wrong) password was given."""


def _text(v):
    return "" if v is None else str(v).replace("\xa0", " ").replace("​", "").strip()


def norm(v):
    """A cell as one line of text: no line breaks, no double spaces, no non-breaking spaces."""
    return re.sub(r"\s+", " ", _text(v)).strip()


def label_key(s):
    """A label without spaces, punctuation and capitals: 'NCB Status' and 'ncb  status' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def num(v):
    """A cell as a float, or None when it is not a plain number ('D', blank, '#VALUE!')."""
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


# ---- the workbook as plain values (the same for .xlsx and the old .xls)
class Sheet:
    """One worksheet as plain values: rows[r][c] (0-based) and whether the sheet is visible (hidden sheets are older versions or helpers)."""

    def __init__(self, name, rows, visible=True, whole_pct=None):
        self.name, self.rows, self.visible = name, rows, visible
        self.whole_pct = whole_pct          # {(row, column) 1-based} of the cells formatted as a whole percent ('0%'); None = the formats are not known (an old .xls)

    def shows_whole(self, r, c):
        """True: the sheet SHOWS this cell as a whole percent ('0%': 0.499915 appears as 50%); False: it shows decimals ('0.00%': 0.675 appears as 67.50%); None: not known"""
        return None if self.whole_pct is None else (r, c) in self.whole_pct

    def cell(self, r, c):
        """1-based, like Excel"""
        try:
            return self.rows[r - 1][c - 1]
        except IndexError:
            return None


class Book:
    def __init__(self, sheets):
        self.sheets = list(sheets)

    @property
    def sheetnames(self):
        return [s.name for s in self.sheets]

    def __getitem__(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        raise KeyError(name)


NEED_LIBS = ("a password-protected workbook needs the small package msoffcrypto-tool (a genuine old .xls would also need xlrd): in the grid_tools folder run  "
             ".venv\\Scripts\\python.exe -m pip install msoffcrypto-tool  (or save the workbook as .xlsx without the password and use that)")


def _sniff(data):
    if data[:4] == b"PK\x03\x04":
        return "zip"
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "ole"
    return "unknown"


def _decrypt(data, password):
    """The decrypted bytes of an encrypted workbook (needs msoffcrypto-tool)."""
    try:
        import msoffcrypto
    except ImportError:
        raise GridError(NEED_LIBS)
    if not password:
        raise PasswordNeeded("the workbook is password protected")
    try:
        f = msoffcrypto.OfficeFile(io.BytesIO(data))
        f.load_key(password=password)
        out = io.BytesIO()
        f.decrypt(out)
    except Exception as e:
        if "password" in str(e).lower() or "key" in type(e).__name__.lower() or "verif" in str(e).lower():
            raise PasswordNeeded("the password does not open the workbook")
        raise
    return out.getvalue()


def _book_openpyxl(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)

    def whole(ws):
        if ws.sheet_state != "visible":
            return set()
        return {(c.row, c.column) for row in ws.iter_rows() for c in row if isinstance(c.value, (int, float)) and c.number_format in ("0%", "0 %")}
    return Book(Sheet(ws.title, [list(r) for r in ws.iter_rows(values_only=True)], ws.sheet_state == "visible", whole(ws)) for ws in wb.worksheets)


def _book_xlrd(data):
    import xlrd
    wb = xlrd.open_workbook(file_contents=data)
    sheets = []
    for i in range(wb.nsheets):
        sh = wb.sheet_by_index(i)
        rows = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                t = cell.ctype
                if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    v = None
                elif t == xlrd.XL_CELL_BOOLEAN:
                    v = bool(cell.value)
                elif t == xlrd.XL_CELL_ERROR:
                    v = xlrd.error_text_from_code.get(cell.value, "#ERR")
                else:
                    v = cell.value
                row.append(v)
            rows.append(row)
        sheets.append(Sheet(sh.name, rows, sh.visibility == 0))
    return Book(sheets)


def book_from_bytes(data, password=None):
    """A workbook from its bytes.  An .xlsx is a zip.  An OLE container is either an ENCRYPTED workbook (the file this script is made for is called .xls but is an encrypted .xlsx: it holds the streams
    EncryptionInfo / EncryptedPackage) - decrypted with msoffcrypto-tool and read as the .xlsx it is - or a genuine old .xls, read with xlrd."""
    kind = _sniff(data)
    if kind == "zip":
        return _book_openpyxl(data)
    if kind != "ole":
        raise GridError("this does not look like an Excel workbook (.xlsx or .xls)")
    try:
        import msoffcrypto
    except ImportError:
        msoffcrypto = None
    if msoffcrypto is not None:
        try:
            encrypted = bool(msoffcrypto.OfficeFile(io.BytesIO(data)).is_encrypted())
        except Exception:
            encrypted = False
        if encrypted:
            return book_from_bytes(_decrypt(data, password), password)
    try:
        return _book_xlrd(data)
    except ImportError:
        raise GridError(NEED_LIBS)
    except Exception as e:
        if msoffcrypto is None:
            raise GridError(f"{NEED_LIBS} ({type(e).__name__}: {e})")
        raise


def read_workbook(path, password=None):
    """Open a workbook (values only). An encrypted file needs the password (and the package msoffcrypto-tool)."""
    with open(path, "rb") as fh:
        data = fh.read()
    return book_from_bytes(data, password)


def make_opener(ask, password=None):
    """The function open_file() calls: it asks for the password (hidden typing) when the workbook is protected and none, or a wrong one, was given."""
    state = {"pw": password or os.environ.get("GRID_PASSWORD") or None}

    def opener(path):
        for attempt in range(3):
            try:
                return read_workbook(path, state["pw"])
            except PasswordNeeded as e:
                if not ask.interactive:
                    raise GridError(f"{e} - give it with --password <pw> (or set the environment variable GRID_PASSWORD)")
                try:
                    state["pw"] = getpass.getpass(f"{e}. Password (typing is hidden; q = stop): ")
                except (EOFError, KeyboardInterrupt):
                    raise GridError("stopped by you - no password given")
                if state["pw"].strip().lower() == "q":
                    raise GridError("stopped by you - no password given")
        raise GridError("the password was wrong three times")
    return opener


# ---- finding the sheets and the columns
def visible_sheets(book):
    return [s for s in book.sheets if s.visible]


def find_sheet(book, words, header_keys, what, ask, optional=False, skip=()):
    """The visible sheet whose name has all `words` (letters/digits only, capitals ignored); failing that the visible sheet whose header row has all `header_keys`; several = ask."""
    names = [s for s in visible_sheets(book) if s.name not in skip]
    hit = [s for s in names if all(w in label_key(s.name) for w in words)]
    if not hit and header_keys:
        hit = [s for s in names if find_header_row(s, header_keys)]
    if len(hit) == 1:
        return hit[0]
    if not hit:
        if optional:
            return None
        raise GridError(f"{what}: no visible sheet has a name with {', '.join(words)}"
                        f"{' or a header with ' + ', '.join(header_keys) if header_keys else ''} (visible sheets: {[s.name for s in visible_sheets(book)]})")
    pick = ask.choose(f"{what}: more than one sheet fits", [(s.name, s.name) for s in hit], question=f"More than one visible sheet looks like {what}. Which one is this month's?", short=what)
    return book[pick]


def find_header_row(sheet, keys, scan=HEADER_SCAN_ROWS):
    """1-based number of the first row (within `scan` rows) whose cells contain every label key in `keys`, else None"""
    for r in range(1, min(scan, len(sheet.rows)) + 1):
        have = {label_key(v) for v in sheet.rows[r - 1] if v is not None}
        if all(k in have for k in keys):
            return r
    return None


# ---- finding a table and its columns by what the headers SAY - never by a sheet's name or by counting cells from a header (the user, 10 Nov 2026).
# Go Digit renamed every sheet of the November workbook, put a new 'Avg CD2' column in front of 'Max CD2' and redesigned the car sheets: a reader that takes "the 5th cell after Cluster" read the wrong
# column without a word.  A table is described by a Spec: its fields, and for each field the headers it may carry (an alias is compared by its label key, exactly, or is a compiled pattern matched
# against the label key; a typo of an alias of 6+ letters is read as the alias and the log says so).  The finder looks at every VISIBLE sheet and takes the best header row; a sheet's NAME only
# breaks a tie.  A table that is not there, or is there twice, is never skipped in silence: it is asked about, or the run stops with what was expected and what the workbook has.
TABLE_SCAN_ROWS = 25                   # a table's header row is looked for in the first rows of a sheet (title rows and blank rows above it are fine)
FUZZY_HEADER = 0.88                    # similarity at which a header that is no alias of a field is read as a typo of one ('Segement' -> Segment, 'Volumn Slab' -> Volume Slab)
PAYOUT_HEAD = re.compile(r"^(max|avg|average)?cd2$")      # the headers of a payout column: 'CD2', 'Max CD2' (also 'MaxCD2'), 'Avg CD2'
MISSING_TABLES = "ask"                 # a table that is not in the workbook: "ask" = stop with what was expected / found unless the person says to write the file without it; "skip" = warn and go on (tests)
KNOWN_PAYOUT_LABELS = ("cd2", "maxcd2")  # the payout headers the accepted uploads were built from (October 2026: 'CD2'; the November commercial-vehicle grid says 'Max CD2'); any other label is reported


def header_hit(cell, aliases, typo=False):
    """True when a header cell says one of the aliases (same label key, or a pattern matching it); with typo=True also a near-miss of an alias of 6+ letters"""
    k = label_key(cell)
    if not k:
        return False
    for a in aliases:
        if hasattr(a, "match"):
            if a.match(k):
                return True
        elif k == label_key(a):
            return True
    if typo and len(k) >= 6:
        return any(not hasattr(a, "match") and len(label_key(a)) >= 6 and difflib.SequenceMatcher(None, k, label_key(a)).ratio() >= FUZZY_HEADER for a in aliases)
    return False


class Spec:
    """What a table says about itself.  fields = {field: [alias, ...]}; required = the fields its header row must say; forbid = fields that make a sheet ANOTHER table; many = fields that may sit
    in several columns (blocks side by side, Max / Avg payouts); show = {field: the header as the grid writes it} for the messages; usual = the sheet's name last month (a tie-breaker, and said when
    the table is found under another name)."""

    def __init__(self, what, usual, fields, required, forbid=(), many=(), show=None):
        self.what, self.usual, self.fields = what, usual, fields
        self.required, self.forbid, self.many = tuple(required), tuple(forbid), tuple(many)
        self.show = {f: (show or {}).get(f) or next((a for a in aliases if isinstance(a, str)), f) for f, aliases in fields.items()}


class Hit:
    """A table found on a sheet: its header row (1-based) and, for every field, the columns (0-based, left to right) whose header says it"""

    def __init__(self, sheet, hr, cols, typos=()):
        self.sheet, self.hr, self.cols, self.typos = sheet, hr, cols, list(typos)

    @property
    def head(self):
        return self.sheet.rows[self.hr - 1]

    def first(self, field):
        c = self.cols.get(field)
        return c[0] if c else None

    def within(self, field, lo, hi):
        """the columns of a field inside [lo, hi)"""
        return [c for c in self.cols.get(field, []) if lo <= c < hi]

    def ncols(self):
        return max([len(r) for r in self.sheet.rows[:80]] + [len(self.head)])

    def label(self, c):
        return norm(self.head[c]) if c < len(self.head) else ""


def cell_at(raw, c):
    return raw[c] if c is not None and 0 <= c < len(raw) else None


def col_name(c):
    """0-based column -> Excel letters"""
    s, c = "", c + 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def scan_header_rows(sheet, spec):
    """The best header row of a sheet for a spec: (row number 1-based, {field: [columns]}, [(typo, field)]), or None.  Every required field must be said, no forbidden one; the row that says
    most fields wins, the upper one when two say the same."""
    best = None
    for r in range(1, min(TABLE_SCAN_ROWS, len(sheet.rows)) + 1):
        row = sheet.rows[r - 1]
        cols, typos, taken = {}, [], set()
        for typo in (False, True):                       # exact labels first; a near-miss is read only for a header nothing else claimed
            for c, v in enumerate(row):
                if c in taken or not label_key(v):
                    continue
                for field, aliases in spec.fields.items():
                    if field in cols and field not in spec.many:
                        continue
                    if header_hit(v, aliases, typo):
                        cols.setdefault(field, []).append(c)
                        taken.add(c)
                        if typo:
                            typos.append((norm(v), field))
                        break
        for f in cols:
            cols[f].sort()
        if all(f in cols for f in spec.required) and not any(f in cols for f in spec.forbid):
            if best is None or len(cols) > len(best[1]):
                best = (r, cols, typos)
    return best


def name_ratio(a, b):
    a, b = label_key(a), label_key(b)
    return difflib.SequenceMatcher(None, a, b).ratio() if a and b else 0.0


def expected_found(book, spec):
    """The message of a table that is not there: what was expected, the closest sheet and what it lacks, the sheets there are"""
    need = ", ".join(spec.show[f] for f in spec.required)
    seen = None
    for s in visible_sheets(book):
        for r in range(1, min(TABLE_SCAN_ROWS, len(s.rows)) + 1):
            said = [f for f, al in spec.fields.items() if any(header_hit(v, al, True) for v in s.rows[r - 1] if v is not None)]
            named = [f for f in said if any(isinstance(a, str) for a in spec.fields[f])]            # a sheet that only has a 'CD2' somewhere is no near miss
            if len(named) >= 2 and (seen is None or len(named) > seen[3]):
                seen = (s, r, said, len(named))
    near = ""
    if seen:
        s, r, said, _ = seen
        lacks = [spec.show[f] for f in spec.required if f not in said]
        other = [spec.show[f] for f in spec.forbid if f in said]
        near = (f" The closest sheet is '{s.name}' (row {r} says {', '.join(spec.show[f] for f in said)}"
                + (f"; it lacks {', '.join(lacks)}" if lacks else "") + (f"; it also has {', '.join(other)}, which belongs to another table" if other else "") + ").")
    return (f"{spec.what}: expected a header row that says {need} in the first {TABLE_SCAN_ROWS} rows of a visible sheet - no visible sheet has one.{near} "
            f"Visible sheets: {[s.name for s in visible_sheets(book)]}")


def find_table(book, spec, ask, warnings=None, log=say, optional=False, skip=()):
    """The Hit of the visible sheet that holds the table (None when optional and absent; otherwise GridError with what was expected and what was found).  The sheet that says most of the
    table's fields wins, then the one whose name is closest to last month's; two equally good sheets are asked about.  A table found under another name is said on the log and in the warnings."""
    hits = []
    for s in visible_sheets(book):
        if s.name in skip:
            continue
        got = scan_header_rows(s, spec)
        if got:
            hits.append(Hit(s, got[0], got[1], got[2]))
    if not hits:
        if optional:
            return None
        raise GridError(expected_found(book, spec))
    hits.sort(key=lambda h: (-len(h.cols), -name_ratio(h.sheet.name, spec.usual)))
    top = hits[0]
    tie = [h for h in hits if len(h.cols) == len(top.cols) and abs(name_ratio(h.sheet.name, spec.usual) - name_ratio(top.sheet.name, spec.usual)) < 0.05]
    if len(tie) > 1:
        pick = ask.choose(f"{spec.what}: more than one sheet fits", [(h.sheet.name, h.sheet.name) for h in tie],
                          question=f"More than one visible sheet looks like {spec.what}. Which one is this month's?", short=spec.what)
        top = next(h for h in tie if h.sheet.name == pick)
    notes = []
    if label_key(top.sheet.name) != label_key(spec.usual):
        notes.append(f"{spec.what}: found by its column headers on sheet '{top.sheet.name}' (header on row {top.hr}); last month's sheet was called '{spec.usual}'")
    notes += [f"{spec.what}: the header '{typo}' on sheet '{top.sheet.name}' is read as '{spec.show[field]}' (a typo?)" for typo, field in top.typos]
    for msg in notes:
        log("  " + msg)
        if warnings is not None:
            warnings.append(msg)
    return top


def table_missing(book, spec, ask, ctx, lost):
    """A table the workbook does not have is a structural break, not a detail: the run stops with what was expected and what was found, unless the person says to write the file without it
    (`lost` = what the upload then lacks).  MISSING_TABLES = "skip" only warns (the tests build small workbooks without the optional sheets)."""
    msg = expected_found(book, spec)
    if MISSING_TABLES != "skip":
        go = ask.choose(msg, [("Stop - look at the workbook first", False), (f"Write the file without it ({lost})", True)], question=f"{msg} What should I do?",
                        short=f"{spec.what} not found", memo=("missing table", spec.what))
        if not go:
            raise GridError(f"stopped by you - {spec.what} was not found")
    ctx.warnings.append(f"NOT FOUND - {msg} {lost[:1].upper() + lost[1:]}.")


def title_texts(sheet, hr, c):
    """The titles above column c of a header row, nearest row first: in each row the nearest non-empty cell at or to the left of c (a title spans the columns to its right) -> [(row, column, text)]"""
    out = []
    for r in range(hr - 1, 0, -1):
        row = sheet.rows[r - 1]
        for cc in range(min(c, len(row) - 1), -1, -1):
            t = norm(row[cc])
            if t:
                out.append((r, cc, t))
                break
    return out


def window_titles(sheet, hr, lo, hi):
    """The titles above the columns lo..hi (inclusive) of a header row, nearest row first: the first non-empty cell of each row inside that window -> [text]"""
    out = []
    for r in range(hr - 1, 0, -1):
        row = sheet.rows[r - 1]
        for cc in range(lo, min(hi, len(row) - 1) + 1):
            t = norm(row[cc])
            if t:
                out.append(t)
                break
    return out


def classify_title(text):
    """What a title above payout columns says: ('section', 'comp' | 'satp') for '1+1' / 'Comp' / 'Comprehensive' / 'SATP' / 'Third Party', ('cpa', 'with' | 'without') for 'With CPA' / 'Without CPA'"""
    k = label_key(text)
    if "cpa" in k:
        return "cpa", ("without" if k.startswith(("without", "wo")) else "with")
    if "satp" in k or "thirdparty" in k or k in ("tp", "stp", "aotp"):
        return "section", "satp"
    if k == "11" or k.startswith("comp"):
        return "section", "comp"
    return None, None


def measure_of(label):
    k = label_key(label)
    return "max" if k.startswith("max") else "avg" if k.startswith(("avg", "average")) else "plain"


def choose_payout(ask, where, cands):
    """The payout column of one table / section.  cands = [(column, header)] of the payout columns that belong together.  One candidate: it.  A 'Max CD2' AND an 'Avg CD2' (the November grid
    has both in the two-wheeler and car third-party tables): the person says which one is the payin - asked ONCE per run and never guessed, because both are plausible and last month's number sits
    in 'Avg CD2' while 'Max CD2' is the higher figure.  -> (column, header)"""
    if len(cands) == 1:
        return cands[0]
    kinds = {}
    for c, lab in cands:
        kinds.setdefault(measure_of(lab), []).append((c, lab))
    dup = [k for k, v in kinds.items() if len(v) > 1]
    if dup:
        raise GridError(f"{where}: expected one payout column per measure; found {len(kinds[dup[0]])} columns headed {[lab for _, lab in kinds[dup[0]]]}")
    names = {"max": "Max CD2", "avg": "Avg CD2", "plain": "CD2"}
    order = [m for m in ("max", "avg", "plain") if m in kinds]
    remembered = ask.memory.get("payout measure")
    if remembered is not None and remembered not in kinds:
        raise GridError(f"{where}: you chose the '{names[remembered]}' column for the payout, but this table has only {[lab for _, lab in cands]}")
    if remembered is None and PAYOUT_MEASURE in kinds:               # the user's standing answer (10 Nov 2026): applied without asking, said once in the decisions of the run
        if not ask.memory.get("payout rule said"):
            ask.decide(f"payout column: '{names[PAYOUT_MEASURE]}' wherever a table gives both Max CD2 and Avg CD2 (your answer of 10 Nov 2026, PAYOUT_MEASURE in the script's settings)")
            ask.memory["payout rule said"] = True
        return kinds[PAYOUT_MEASURE][0]
    pick = remembered or ask.choose(
        "the grid has two payout columns, 'Max CD2' and 'Avg CD2' (earlier grids had one, 'CD2'): which one is the payin?",
        [(f"{names[m]} - the column headed '{kinds[m][0][1]}'", m) for m in order],
        question=("Go Digit's two-wheeler and private-car third-party tables give TWO payout columns, 'Max CD2' and 'Avg CD2' (the commercial-vehicle grid has only 'Max CD2'). "
                  "The grids before November 2026 had one column, and its figures are the ones the November grid keeps under 'Avg CD2'; 'Max CD2' is the higher figure. Which column is the payin you want in the upload?"),
        short="payout column (Max CD2 or Avg CD2)", memo="payout measure")
    ask.memory["payout measure"] = pick
    return kinds[pick][0]


def unread_headers(hit, used, lo=0, hi=None, ignore=()):
    """The header cells of a table that no rule read: [(column, text)].  Long texts are notes, not headers; `ignore` = patterns for columns this script leaves out on purpose."""
    out = []
    for c in range(lo, hi if hi is not None else len(hit.head)):
        t = hit.label(c)
        if not t or c in used or len(t) > 40:
            continue
        if any(header_hit(t, [p]) for p in ignore):
            continue
        out.append((c, t))
    return out


READ_COLUMNS = {}         # sheet name -> the columns some reader of this run used (filled by warn_unread; cleared by build)
LEFT_OUT = {}             # sheet name -> {column: header} of the columns this script leaves out on purpose (the `ignore` patterns of warn_unread)
WARNED_COLUMNS = {}       # sheet name -> {column: header} of the columns warn_unread already named in a warning
SHEET_ROLES = {}          # sheet name -> 'mapping' | 'remarks': the sheets read for what they are (the RTO codes of a cluster, the notes)
UNREAD_PARTS = []         # what nothing read, as dicts (sheet, where, holds, cells, status, why): the review sheet 'Unread parts' and the run's warnings
LEFT_OUT_REASONS = [      # (pattern on the label key of a header, why this script leaves the column out)
    (re.compile(r"^mincd1$"), "the lowest discount of the band (negative = a loading): your October upload has tariff_min 0 on every two-wheeler SAOD row (380 rows) although that grid had Min CD1 values too, so the band starts at 0"),
    (re.compile(r"^max\d+(st|nd|rd|th)?year"), "a SAOD policy is one year: only the 'Max 1 year SAOD CD2' payout is written"),
]


def warn_unread(hit, used, warnings, lo=0, hi=None, ignore=()):
    """A column nobody reads is a column nobody noticed: say it (a new column is often a new rule).  Also remembers which columns were used / left out on purpose / named, for audit_unread."""
    name = hit.sheet.name
    READ_COLUMNS.setdefault(name, set()).update(used)
    for c in range(lo, hi if hi is not None else len(hit.head)):
        t = hit.label(c)
        if t and c not in used and any(header_hit(t, [p]) for p in ignore):
            LEFT_OUT.setdefault(name, {})[c] = t
    left = unread_headers(hit, used, lo, hi, ignore)
    if left:
        WARNED_COLUMNS.setdefault(name, {}).update({c: t for c, t in left})
        warnings.append(f"sheet '{hit.sheet.name}': the column(s) {', '.join(repr(t) + ' (' + col_name(c) + ')' for c, t in left)} are not read by this script - check what they mean before you use the upload")


def _column_cells(sheet, c):
    return [(r, row[c]) for r, row in enumerate(sheet.rows, 1) if c < len(row) and row[c] is not None and norm(row[c]) != ""]


def _is_key_column(cells):
    """a column nobody labelled whose cells are all different texts ('AndamanBAJAJ3-7 KW'): a helper key the sheet's author concatenated, no data"""
    if len(cells) < 20 or not all(isinstance(v, str) for _, v in cells):
        return False
    texts = [norm(v) for _, v in cells]
    return len(set(texts)) == len(texts) and sorted(len(t) for t in texts)[len(texts) // 2] >= 8


def _describe(sheet, cols):
    """(the header-like texts at the top, the long texts) of some columns: what the person reading the review sheet needs to recognise the part"""
    top, notes = [], []
    cols = [c for c in cols if not _is_key_column(_column_cells(sheet, c))]
    for r, row in enumerate(sheet.rows[:8], 1):                                   # the first rows across the columns: the titles and the header row, not the first column's data
        for c in cols:
            v = row[c] if c < len(row) else None
            t = norm(v)
            if isinstance(v, str) and t and len(t) < 25 and t not in top and len(top) < 10:
                top.append(t[:48])
    for c in cols:
        for r, v in _column_cells(sheet, c):
            t = norm(v)
            if isinstance(v, str) and len(t) >= 25 and len(notes) < 4 and t not in notes:
                notes.append(t[:160])
    return top, notes


def audit_unread(book, warnings):
    """Everything in the visible sheets that no reader used (the user, 18 Nov 2026: conditions 'not applied and not mentioned anywhere'): a whole sheet no table was found on, columns beside or beyond a table that was
    read (the car sheet's new-business blocks), columns left out on purpose.  Mapping sheets and the remarks sheet are read for what they are.  -> UNREAD_PARTS + one warning per unread sheet / block."""
    UNREAD_PARTS.clear()
    for s in visible_sheets(book):
        if SHEET_ROLES.get(s.name):
            continue
        n_all = sum(1 for row in s.rows for v in row if v is not None and norm(v) != "")
        if not n_all:
            continue
        ncols = max((len(r) for r in s.rows), default=0)
        used = READ_COLUMNS.get(s.name)
        if used is None:
            top, notes = _describe(s, range(ncols))
            what = "; ".join(top[:10] + ([f"text: {n}" for n in notes[:2]] if notes else []))
            UNREAD_PARTS.append(dict(sheet=s.name, where="the whole sheet", holds=what, cells=n_all, status="NOT READ", why="no table of this script was found on this sheet: nothing from it is in the upload"))
            warnings.append(f"sheet '{s.name}' ({n_all:,} filled cells; it says {what[:160]}) is not read by this script: nothing from it is in the upload")
            continue
        left, skip = LEFT_OUT.get(s.name, {}), WARNED_COLUMNS.get(s.name, {})
        for c, t in sorted(left.items()):                                            # columns this script leaves out on purpose: listed with the reason, no warning
            if c not in used and _column_cells(s, c):
                why = next((w for p, w in LEFT_OUT_REASONS if p.match(label_key(t))), "left out on purpose")
                UNREAD_PARTS.append(dict(sheet=s.name, where=f"column {col_name(c)}", holds=t, cells=len(_column_cells(s, c)), status="left out on purpose", why=why))
        for c in sorted(skip):                                                       # columns warn_unread named already: in the review sheet too (the warning is not repeated)
            if c not in used and _column_cells(s, c):
                top, notes = _describe(s, [c])
                UNREAD_PARTS.append(dict(sheet=s.name, where=f"column {col_name(c)}", holds="; ".join([skip[c]] + [t for t in top[:3] if t != skip[c]]), cells=len(_column_cells(s, c)), status="NOT READ",
                                         why="a column of the table that was read, named in the run's warnings: nothing from it is in the upload"))
        cand = []
        for c in range(ncols):
            cells = _column_cells(s, c)
            if not cells or c in used or c in skip or c in left or _is_key_column(cells):
                continue
            cand.append(c)
        runs = []
        for c in cand:
            if runs and c == runs[-1][-1] + 1:
                runs[-1].append(c)
            else:
                runs.append([c])
        for cols in runs:
            top, notes = _describe(s, cols)
            where = f"column {col_name(cols[0])}" if len(cols) == 1 else f"columns {col_name(cols[0])}:{col_name(cols[-1])}"
            what = "; ".join(top[:10] + [f"text: {n}" for n in notes[:3]])
            n = sum(len(_column_cells(s, c)) for c in cols)
            UNREAD_PARTS.append(dict(sheet=s.name, where=where, holds=what, cells=n, status="NOT READ", why="beside or beyond the table that was read: nothing from it is in the upload"))
            warnings.append(f"sheet '{s.name}': {where} ({n:,} filled cells; they say {what[:200]}) are not read by this script: nothing from them is in the upload")


def block_ranges(hit, anchors):
    """{anchor column: (first, end)}: a table with ONE anchor header ('Cluster') owns the whole row, whatever order its columns are in; with several side by side, each owns the columns from its anchor
    to the next anchor"""
    anchors = sorted(anchors)
    if len(anchors) == 1:
        return {anchors[0]: (0, max(hit.ncols(), anchors[0] + 1))}
    return {s: (s, anchors[i + 1] if i + 1 < len(anchors) else max(hit.ncols(), s + 1)) for i, s in enumerate(anchors)}


class WorkingCol:
    """The person's RTO working list in one table of a sheet (the user, 31 Oct 2026): the column headed RTO (WORKING_HEADERS) of the table.  A sheet with one cluster column: the nearest RTO
    column (its right-hand side first).  A sheet with several side-by-side blocks (`starts` = their cluster columns): the first RTO column at or to the right of this block's cluster column and
    before the next block, so each block has its own at its right edge.  `span` = how many columns the block's own fields cover (the 1+5 and the private-car TP tables): an RTO column between
    its first and last field would sit in a column the table uses, so it is not used and a warning says to put it to the right of the table.  `name(raw, r)` = the name of a grid row ('' = none:
    the master lookup is used); the rows are counted for the run line and `done()` hands the count to working_table."""

    def __init__(self, sheet, cluster_col, part="", starts=(), span=0, warnings=None, scan=HEADER_SCAN_ROWS):
        self.sheet, self.part = sheet, part
        starts = sorted(set(starts) | {cluster_col})
        nxt = next((x for x in starts if x > cluster_col), None)
        found = []
        for r in range(1, min(max(scan, HEADER_SCAN_ROWS), len(sheet.rows)) + 1):
            for c, v in enumerate(sheet.rows[r - 1]):
                if v is not None and working_header(v):
                    if len(starts) > 1 and not (c >= cluster_col and (nxt is None or c < nxt)):
                        continue
                    found.append((abs(c - cluster_col), 0 if c > cluster_col else 1, r, c))
        best = min(found) if found else None
        if best and span and cluster_col < best[3] < cluster_col + span:
            if warnings is not None:
                warnings.append(f"sheet '{sheet.name}': the RTO column {norm(sheet.rows[best[2] - 1][best[3]])!r} sits inside the columns of a table block (between its 'Cluster' column and its last column, so it "
                                f"would take the place of one of the table's own columns) - it is not used: put it to the right of the table (after its last column) and run again")
            best = None
        self.col = best[3] if best else None
        self.header = norm(sheet.rows[best[2] - 1][best[3]]) if best else None
        self.with_value, self.blank = 0, []

    def name(self, raw, r):
        if self.col is None:
            return ""
        v = working_value(raw[self.col]) if self.col < len(raw) else ""
        if v:
            self.with_value += 1
        else:
            self.blank.append(r)
        return v

    def done(self):
        working_table(self.sheet.name + (f" ({self.part})" if self.part else ""), self.header, self.with_value, self.blank)


AGE_LABELS = ["new", "1yr", "2yr", "3yr", "4yr", "5yr", "610yr", "11yr"]
AGES = [("New", 0, 1), ("1 Yr", 1, 2), ("2 Yr", 2, 3), ("3 Yr", 3, 4), ("4 Yr", 4, 5), ("5 Yr", 5, 6), ("6-10 Yr", 6, 10.01), ("11+ Yr", 10.01, AGE_OPEN_END)]


class CvRow:
    """one row of the commercial-vehicle grid"""
    __slots__ = ("r", "seg", "body_seg", "body", "cluster", "state", "make", "cd1", "comp", "satp", "rto")


SPEC_CV = Spec("the commercial-vehicle grid", "LATEST CV COMBINED GRID",
               fields=dict(seg=["Cluster segment"], body_seg=["Segment"], body=["Body Type", "Body"], cluster=["RTO Cluster"], state=["State"],
                           make=["Make / Variant", "Make", "Variant"], cd1=[re.compile(r"^cd1")]),
               required=("seg", "cluster"), show=dict(cd1="CD1 (OD-disc)"))


def age_group_kind(sheet, hr, cols):
    """('comp' | 'satp' | None, the payout header or '') of a group of age columns, from the titles written above them: 'COMPREHENSIVE ...' / 'Third Party/SATP ...' and 'CD2' / 'Max CD2'"""
    kind, measure = None, ""
    for t in window_titles(sheet, hr, cols[0], cols[-1]):
        what, val = classify_title(t)
        if what == "section" and kind is None:
            kind = val
        if not measure and PAYOUT_HEAD.match(label_key(t)):
            measure = t
    return kind, measure


def read_cv_grid(book, ask, log=say, warnings=None):
    """-> (CvRow list, note texts, the sheet).  The sheet and its columns are found by their header text, never by the sheet's name or a column's place; the comprehensive and the third-party group of age
    columns by the titles above them (the order is only assumed when there are none)."""
    warnings = warnings if warnings is not None else []
    hit = find_table(book, SPEC_CV, ask, warnings, log)
    sheet, hr, head = hit.sheet, hit.hr, hit.head
    col = {f: c[0] for f, c in hit.cols.items()}
    missing = [f for f in ("seg", "body_seg", "body", "cluster", "make", "cd1") if f not in col]
    if missing:
        raise GridError(f"sheet '{sheet.name}': the header row {hr} has no column for {', '.join(SPEC_CV.show[f] for f in missing)} (the row says {[norm(v) for v in head if norm(v)][:10]})")
    ages = [c for c, v in enumerate(head) if label_key(v) in AGE_LABELS]
    labels = [label_key(head[c]) for c in ages]
    if labels != AGE_LABELS + AGE_LABELS:
        raise GridError(f"sheet '{sheet.name}': the age columns are not New, 1 Yr, 2 Yr, 3 Yr, 4 Yr, 5 Yr, 6-10 Yr, 11+ Yr twice (comprehensive, then third-party); "
                        f"found {[norm(head[c]) for c in ages]} - the age bands of this script need a look")
    comp_cols, satp_cols = ages[:8], ages[8:]
    (k1, m1), (k2, m2) = age_group_kind(sheet, hr, comp_cols), age_group_kind(sheet, hr, satp_cols)
    if k1 and k1 == k2:
        raise GridError(f"sheet '{sheet.name}': expected one comprehensive and one third-party group of age columns; the titles above both groups say '{k1}' - the age columns need a look")
    if k1 == "satp" or k2 == "comp":
        comp_cols, satp_cols = satp_cols, comp_cols
        warnings.append(f"sheet '{sheet.name}': the titles above the age columns put the third-party group first - the groups are read by their titles")
    named = sorted({m for m in (m1, m2) if m and label_key(m) != "cd2"})
    if named:
        warnings.append(f"sheet '{sheet.name}': the payout columns are headed {', '.join(repr(m) for m in named)} (earlier grids said 'CD2'): the figures are read as they are"
                        + ("" if all(label_key(m) in KNOWN_PAYOUT_LABELS for m in named) else " - but this script was built for 'CD2' / 'Max CD2': check what the heading means") + " - compare them with last month's upload before you use the file")
    rows, notes = [], []
    wc = WorkingCol(sheet, col["cluster"], scan=hr)
    for r in range(hr + 1, len(sheet.rows) + 1):
        raw = sheet.rows[r - 1]
        get = lambda c: raw[c] if c < len(raw) else None
        if all(get(c) is None or norm(get(c)) == "" for c in range(len(raw))):
            continue
        seg, cluster = norm(get(col["seg"])), norm(get(col["cluster"]))
        if not seg or not cluster:
            notes.extend(norm(get(c)) for c in range(len(raw)) if len(norm(get(c))) >= 25)
            continue
        g = CvRow()
        g.r, g.seg, g.body_seg, g.body, g.cluster = r, seg, norm(get(col["body_seg"])), norm(get(col["body"])), cluster
        g.state, g.make, g.cd1 = norm(get(col["state"])) if "state" in col else "", norm(get(col["make"])), get(col["cd1"])
        g.comp, g.satp = [get(c) for c in comp_cols], [get(c) for c in satp_cols]
        g.rto = wc.name(raw, r)
        rows.append(g)
    wc.done()
    if not rows:
        raise GridError(f"sheet '{sheet.name}': no rate rows below the header")
    for r in range(1, hr):
        notes.extend(norm(v) for v in sheet.rows[r - 1] if len(norm(v)) >= 25 and "COMPREHENSIVE" not in norm(v).upper() and "THIRD-PARTY" not in norm(v).upper())
    warn_unread(hit, set(col.values()) | set(ages) | ({wc.col} if wc.col is not None else set()), warnings)
    log(f"Commercial-vehicle grid: sheet '{sheet.name}', {len(rows):,} rows")
    return rows, notes, sheet


REMARKS_HEAD = re.compile(r"^REMARKS FOR (.+)$", re.I)
AGE_PHRASE = re.compile(r"^(?:up\s*to\s*(?P<hi0>\d+)\s*months?|(?:(?P<y>\d+)\s*years?\s*)?(?P<m>\d+)\s*months?\s*onwards(?:\s*to\s*(?:(?P<y2>\d+)\s*years?\s*)?(?P<m2>\d+)\s*months?)?)$", re.I)


def find_remarks_sheet(book, ask):
    """The visible sheet that holds the 'REMARKS FOR ...' headings (the one with most of them; a tie goes to the name closest to 'CV Remarks'); failing that the sheet whose name has the words cv + remarks"""
    found = []
    for s in visible_sheets(book):
        n = sum(1 for row in s.rows for v in row if REMARKS_HEAD.match(norm(v)))
        if n:
            found.append((n, name_ratio(s.name, "CV Remarks"), s))
    if not found:
        return find_sheet(book, CV_REMARKS_SHEET_WORDS, (), "the CV remarks", ask, optional=True)
    found.sort(key=lambda x: (-x[0], -x[1]))
    tie = [x for x in found if x[0] == found[0][0] and abs(x[1] - found[0][1]) < 0.05]
    if len(tie) > 1:
        pick = ask.choose("the CV remarks: more than one sheet fits", [(x[2].name, x[2].name) for x in tie],
                          question="More than one visible sheet looks like the CV remarks. Which one is this month's?", short="the CV remarks")
        return book[pick]
    return found[0][2]


def read_cv_remarks(book, ask):
    """the text cells of the remarks sheet by section ('REMARKS FOR SCHOOL BUS / STAFF BUS / HCV'): {'schoolbus': [...], 'staffbus': [...], 'hcv': [...]} - {} when the sheet is not there.  The sheet is found by its
    'REMARKS FOR' headings, not by its name.  The table that defines the age bands ('System Age | Age Band': 'Up to 11 months', '12 months onwards to 1 year 11 months' ...) is not a note: its lines are
    returned apart, as '_age_table' = [(system age, phrase)], to be checked against this script's own bands (check_age_table)."""
    sh = find_remarks_sheet(book, ask)
    out, cur, age_lines = {}, "other", []
    if sh is None:
        return out
    SHEET_ROLES[sh.name] = "remarks"
    for row in sh.rows:
        for ci, v in enumerate(row):
            t = norm(v)
            m = REMARKS_HEAD.match(t)
            if m:
                cur = label_key(m.group(1))
            elif AGE_PHRASE.match(t):
                age_lines.append((norm(row[ci - 1]) if ci else "", t))
            elif len(t) >= 20:
                out.setdefault(cur, []).append(t)
    if age_lines:
        out["_age_table"] = age_lines
    return out


def age_phrase_years(text):
    """'Up to 11 months' -> (0, 1.0) ; '12 months onwards to 1 year 11 months' -> (1.0, 2.0) ; '10 years 12 months onwards' -> (11.0, None): the band in years, its upper end exclusive"""
    m = AGE_PHRASE.match(norm(text))
    if not m:
        return None
    if m.group("hi0") is not None:
        return 0.0, (int(m.group("hi0")) + 1) / 12
    lo = (int(m.group("y") or 0) * 12 + int(m.group("m"))) / 12
    hi = (int(m.group("y2") or 0) * 12 + int(m.group("m2")) + 1) / 12 if m.group("m2") is not None else None
    return lo, hi


def check_age_table(table, warnings, log=say):
    """The sheet's own definition of the age bands against this script's AGES (whole years: 10.01 means 'up to 10 years 11 months').  A band that differs is a warning - the rows would carry the wrong ages."""
    if not table:
        return
    bad, seen = [], 0
    for label, phrase in table:
        mine = next((a for a in AGES if label_key(a[0]) == label_key(label)), None)
        years = age_phrase_years(phrase)
        if mine is None or years is None:
            continue
        seen += 1
        lo, hi = years
        if math.ceil(mine[1] - 1e-9) != round(lo) or (hi is not None and math.ceil(mine[2] - 1e-9) != round(hi)):
            bad.append(f"'{label}' = {phrase!r} (years {lo:g} to {'open' if hi is None else format(hi, 'g')}) but this script writes {mine[1]:g} to {mine[2]:g}")
    if bad:
        warnings.append("the sheet's own age-band table differs from the age bands of this script: " + "; ".join(bad) + " - the age ranges of the rows need a look")
    elif seen:
        log(f"  age bands: the sheet's own table ({seen} bands) agrees with the bands of this script")


HCV_NOTE_BLOCK = re.compile(r"\b(volvo|scania)\b.*\bblock", re.I | re.S)
HCV_NOTE_CD1 = re.compile(r"only for\s+([a-z/\- ]+?)\s*:\s*([a-z ,&/]+?)\s+to operate at\s+(\d+(?:\.\d+)?)\s*%\s*CD1", re.I)


def read_hcv_rule(remarks):
    """The CV Remarks note  'Volvo, Scania to block. Only for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1'
    -> dict(blocked=True/False, cap=70.0 or None, body='non-dumper/tipper')."""
    out = dict(blocked=False, cap=None, body=None)
    for t in remarks:
        if HCV_NOTE_BLOCK.search(t):
            out["blocked"] = True
        m = HCV_NOTE_CD1.search(t.replace("\n", " "))
        if m:
            out["body"], out["cap"] = label_key(m.group(1)), float(m.group(3))
    return out


def read_cv_mapping(book, ask):
    """RTO codes of every (cluster segment, cluster) of the CV RTO mapping sheet -> {(segment, cluster upper): [codes]}.  Only used to say which RTOs a missing group would hold.  The sheet is found by what
    it says (a column of RTO codes and a cluster per product), not by its name - see read_all_mappings."""
    return read_all_mappings(book, ask)


# ---- the cells of the grid: CD2 (the payout) and CD1 (the discount band it is valid for)
NUM = r"(\d+(?:\.\d+)?)"
PCT_RX = re.compile(rf"^{NUM}\s*%$")
PAIR_RX = re.compile(rf"^{NUM}\s*%\s*/\s*{NUM}\s*%$")
ADDON_RX = re.compile(r"^With\s*Addon:\s*([\d.%/\s]+?)\s*Without\s*Addon:\s*([\d.%/\s]+)$", re.I)
TATA_CELL_RX = re.compile(rf"^{NUM}\s*%\s*;\s*TATA\s*-?\s*(With\s*Addon:.+)$", re.I)
NCB_RX = re.compile(rf"^with NCB/break in > 90 days\s*:\s*{NUM}\s*%\s*,\s*without NCB\s*:\s*{NUM}\s*%$", re.I)
TATA_CLAUSE_RX = re.compile(r"[;,]\s*TATA age 0\s*-\s*(.+)$", re.I)


def pcts(s):
    return [rnd(x) for x in re.findall(r"\d+(?:\.\d+)?", s)]


def cd2(v):
    """a payout cell -> ('blank',) ('D',) ('err',) ('rate', [a, b?]) ('addon', with, without) ('tata', base, with, without) ('raw', text)"""
    if v is None:
        return ("blank",)
    if isinstance(v, bool):
        return ("raw", str(v))
    if isinstance(v, (int, float)):
        return ("rate", [rnd(v * 100)])
    t = norm(v)
    if t == "":
        return ("blank",)
    if t.upper() == "D":
        return ("D",)
    if t.startswith("#"):
        return ("err",)
    m = PCT_RX.match(t)
    if m:
        return ("rate", [rnd(m.group(1))])
    if re.fullmatch(r"0?\.\d+|0", t):                                         # '0.4' typed as text: the same fraction the number cells hold (0.4 = 40%); '40' / '1' / '1.5' are not guessed
        return ("rate", [rnd(float(t) * 100)])
    m = PAIR_RX.match(t)
    if m:
        return ("rate", [rnd(m.group(1)), rnd(m.group(2))])
    m = ADDON_RX.match(t)
    if m:
        return ("addon", pcts(m.group(1)), pcts(m.group(2)))
    m = TATA_CELL_RX.match(t)
    if m:
        a = ADDON_RX.match(m.group(2).strip())
        if a:
            return ("tata", rnd(m.group(1)), pcts(a.group(1)), pcts(a.group(2)))
    return ("raw", t)


def cap_value(s):
    """'D' | '85%' | '85%/90%' | '0-85%' -> ('D',) | ('cap', [85]) | ('cap', [85, 90]) ; anything else -> None"""
    t = norm(s).replace("–", "-").replace("—", "-")
    if t.upper() == "D":
        return ("D",)
    t = re.sub(r"^0\s*-\s*", "", t)
    m = re.fullmatch(rf"{NUM}\s*%?(?:\s*/\s*{NUM}\s*%?)?", t)
    if m:
        return ("cap", [rnd(m.group(1))] + ([rnd(m.group(2))] if m.group(2) else []))
    return None


def slot_of(s):
    """one age group's CD1 text -> ('cap', [..]) | ('D',) | ('addon', with_caps, without_caps) | ('ncb', yes_cap, no_cap) | ('band', low, high) | None"""
    t = norm(s).replace("–", "-").replace("—", "-")
    one = cap_value(t)
    if one:
        return one
    m = ADDON_RX.match(t)
    if m:
        return ("addon", pcts(m.group(1)), pcts(m.group(2)))
    m = NCB_RX.match(t)
    if m:
        return ("ncb", rnd(m.group(1)), rnd(m.group(2)))
    m = re.fullmatch(rf"{NUM}\s*%?\s*-\s*{NUM}\s*%?", t)                 # '65.01%-80%': a band with its own lower limit
    if m:
        return ("band", rnd(m.group(1)), rnd(m.group(2)))
    return None


# ---- CD1: the discount cap each age group's payout is valid for
# A CD1 text names, per age group, the highest OD discount at which the payout holds.  The grid writes it in many wordings, so this is a small grammar and not a list of known texts:
#   text   := cap                                   one cap for every age: '85%', '0-85%', 'D', '85%/90%', 'With Addon: 85/90%; Without Addon: 60%' (slot_of reads the plain ones)
#           | clause (separator clause)*            one clause per age group
#   clause := ages [separator] cap                  'new, 1 yr cd1 - 60%', '2-5yr: 70%', '0-5 Yr: 0-60%', 'Age 6+ 70%', 'New/1yr: D'
#           | cap 'for' ages                        '0-40% for age 6+ years'
#   ages   := age (separator age)*                  'New/1yr', 'new, 1-5yr'
# EVERY character of the text must belong to a token and every age must be settled, otherwise the text is NOT read (None) and emit_cv flags it: an unknown word, two groups that disagree
# about one age, a column split between two caps, a column the text leaves out although its payout cell HAS a payout, a repair the payout cells do not back - nothing is guessed.
CD1_AGE_TOP = 40                                                            # ages are whole years; 40 stands for 'and above'
CD1_COLUMN_AGES = [{0}, {1}, {2}, {3}, {4}, {5}, set(range(6, 11)), set(range(11, CD1_AGE_TOP + 1))]      # the ages of the eight age columns of the grid (AGES): New, 1 ... 5 Yr, 6-10 Yr, 11+ Yr
_CD1_NUM = r"\d+(?:\.\d+)?"
_CD1_UNIT = r"(?:yrs?|years?)"
CD1_CAPS = [                                                                # (kind, regex), tried in this order at the start of every token
    ("addon", re.compile(r"with\s*add-?on\s*:?\s*(?P<a>[\d.%/\s]+?)\s*[;|,]?\s*without\s*add-?on\s*:?\s*(?P<b>[\d.%/\s]+)", re.I)),
    ("ncb", re.compile(rf"with\s*ncb\s*/\s*break\s*in\s*>\s*90\s*days\s*:?\s*(?P<y>{_CD1_NUM})\s*%?\s*[,;|]?\s*without\s*ncb\s*:?\s*(?P<n>{_CD1_NUM})\s*%?", re.I)),
    ("band", re.compile(rf"(?P<lo>{_CD1_NUM})\s*%?\s*-\s*(?P<hi>{_CD1_NUM})\s*%", re.I)),
    ("pair", re.compile(rf"(?P<a>{_CD1_NUM})\s*%?\s*/\s*(?P<b>{_CD1_NUM})\s*%", re.I)),
    ("glued", re.compile(rf"(?P<a>{_CD1_NUM})\s*%\s*(?P<b>{_CD1_NUM})\s*%", re.I)),
    ("single", re.compile(rf"(?P<a>{_CD1_NUM})\s*%", re.I)),
    ("decline", re.compile(r"d\b", re.I)),
]
CD1_AGE_WORDS = [
    ("new", re.compile(r"(?:for\s+)?new\b", re.I)),
    ("range", re.compile(rf"(?:for\s+)?(?:age\s*)?(?P<lo>\d+)\s*-\s*(?P<hi>\d+)(?![\d.])(?!-\d)\s*(?:{_CD1_UNIT})?(?!\s*%)", re.I)),     # '1 - 0-60%' is the age 1, a dash and the cap 0-60%: never the range 1-0
    ("plus", re.compile(rf"(?:for\s+)?(?:age\s*)?(?:(?P<lo>\d+)\s*(?:{_CD1_UNIT})?\s*\+|>=\s*(?P<ge>\d+))\s*(?:{_CD1_UNIT})?", re.I)),
    ("exact", re.compile(rf"(?:for\s+)?(?:age\s*)?(?P<n>\d+)\s*{_CD1_UNIT}\b", re.I)),
    ("exact0", re.compile(r"(?:for\s+)?age\s*(?P<n>\d+)\b(?!\s*[%./+])", re.I)),
]
CD1_SEP = re.compile(r"\s*[:;|,/&-]\s*|\s+|\band\b", re.I)
CD1_BARE = re.compile(rf"(?P<a>{_CD1_NUM})(?![\d.]|\s*(?:%|-|/|\+|yrs?\b|years?\b))", re.I)       # a cap written without its % sign: only straight after 'age group:'
CD1_AUTO = {}                                                               # text -> [the eight slots, grid rows] of every text the grammar read in this run (the review workbook lists them)


def cd1_normalize(v):
    t = norm(v).replace("–", "-").replace("—", "-")
    t = re.sub(r"-{2,}", "-", t)                                            # '20.1%-–70%'
    t = re.sub(r"%{2,}", "%", t)                                            # '70%%'
    t = re.sub(r"\bcd1\b", " ", t, flags=re.I)                              # the word 'CD1' between an age group and its cap
    return re.sub(r"\s+", " ", t).strip()


def _cd1_cap(kind, m):
    """one cap token -> (slot, repair); (None, None) when the number is not a percent or the band runs backwards"""
    g = m.groupdict()
    if kind == "decline":
        return ("D",), None
    if kind in ("single", "pair", "glued"):
        vals = [rnd(g[k]) for k in ("a", "b") if g.get(k)]
        return (("cap", vals), "pair" if kind == "glued" else None) if max(vals) <= 100 else (None, None)
    if kind == "band":
        lo, hi = rnd(g["lo"]), rnd(g["hi"])
        if not lo < hi <= 100:
            return None, None
        return (("cap", [hi]) if lo == 0 else ("band", lo, hi)), None
    if kind == "ncb":
        y, n = rnd(g["y"]), rnd(g["n"])
        return (("ncb", y, n), None) if max(y, n) <= 100 else (None, None)
    a, b = pcts(g["a"]), pcts(g["b"])                                       # add-on: 'With Addon: 85/90%; Without Addon: 60%'
    return (("addon", a, b), None) if a and b and max(a + b) <= 100 else (None, None)


def _cd1_age(kind, m):
    """one age token -> (first age, last age or None = 'and above'); None for a range that runs backwards"""
    g = m.groupdict()
    if kind == "new":
        return (0, 0)
    if kind == "range":
        lo, hi = int(g["lo"]), int(g["hi"])
        return (lo, hi) if lo <= hi else None
    if kind == "plus":
        return (int(g["lo"] if g["lo"] is not None else g["ge"]), None)
    return (int(g["n"]), int(g["n"]))


def cd1_lex(t):
    """text -> [('age', (first, last), None) | ('cap', slot, repair) | ('sep', text, None)], or None when some part of the text is no token"""
    out, i = [], 0
    while i < len(t):
        tok = None
        for kind, rx in CD1_CAPS:
            m = rx.match(t, i)
            if m:
                slot, rep = _cd1_cap(kind, m)
                if slot is None:
                    return None
                tok = (("cap", slot, rep), m.end())
                break
        if tok is None:
            for kind, rx in CD1_AGE_WORDS:
                m = rx.match(t, i)
                if m:
                    age = _cd1_age(kind, m)
                    if age is None:
                        return None
                    tok = (("age", age, None), m.end())
                    break
        if tok is None and len(out) >= 2 and out[-1][0] == "sep" and ":" in out[-1][1] and out[-2][0] == "age":
            m = CD1_BARE.match(t, i)
            if m:
                tok = (("cap", ("cap", [rnd(m.group("a"))]), None), m.end())
        if tok is None:
            m = CD1_SEP.match(t, i)
            if not m or m.end() == i:
                return None
            tok = (("sep", m.group(), None), m.end())
        out.append(tok[0])
        i = tok[1]
    return out


def cd1_clauses(tokens):
    """tokens -> [(age ranges, (slot, repair))]: 'new, 1 yr - 60%' (ages, then the cap) or '0-40% for age 6+' (the cap, then 'for' its age); None when ages and caps do not pair up"""
    clauses, ages, pending = [], [], None
    for kind, val, rep in tokens:
        if kind == "age":
            if pending is not None and not ages:
                clauses.append(([val], pending))
                pending = None
            else:
                ages.append(val)
        elif kind == "cap":
            if ages:
                clauses.append((ages, (val, rep)))
                ages = []
            elif pending is None:
                pending = (val, rep)
            else:
                return None
    return None if (ages or pending is not None) else clauses


def cd1_columns(clauses):
    """[(age ranges, cap)] -> for each of the eight age columns (slot, repair), or None where the text names nothing for it.  Every age 0..CD1_AGE_TOP takes the cap of the group that names it; where
    a closed range ('6-10yr') and an open one ('10yr+') both name an age, the closed range keeps it (the groups of the text partition the age axis, and a group the next one overwrote entirely would
    be pointless: the payout cells of the grid split exactly there, F-50).  Two closed or two open groups that disagree about one age, a group that names no age the grid has, and a column
    whose named ages take two different caps are refusals (None for the whole text)."""
    claim = {}
    for ages, cap in clauses:
        named = 0
        for first, last in ages:
            for a in range(first, (CD1_AGE_TOP if last is None else min(last, CD1_AGE_TOP)) + 1):
                claim.setdefault(a, []).append((last is not None, cap))
                named += 1
        if not named:
            return None
    cols = []
    for ages in CD1_COLUMN_AGES:
        got = []
        for a in sorted(ages):
            c = claim.get(a)
            if not c:
                continue
            pool = [x for x in c if x[0]] or c
            if any(x[1][0] != pool[0][1][0] for x in pool):
                return None
            got.append(pool[0][1])
        if not got:
            cols.append(None)
        elif any(x[0] != got[0][0] for x in got):
            return None
        else:
            cols.append(got[0])
    return cols


def cd1_read(v, cells=None):
    """the eight slots (one per age column) of a CD1 text with age groups or an unusual wording, or None when the text is not understood.  `cells` are the eight payout cells of the same grid row:
    they back the repair ('85%90%' = '85%/90%' only when the payout cells are pairs) and mark an age group that the text leaves out although its payout cell HAS a payout (slot ('cap', [100.0], 'unnamed'): nothing is guessed for it)."""
    toks = cd1_lex(cd1_normalize(v))
    if not toks:
        return None
    caps = [x for x in toks if x[0] == "cap"]
    if not caps:
        return None
    if not any(x[0] == "age" for x in toks):
        if len(caps) != 1:
            return None
        cols = [(caps[0][1], caps[0][2])] * 8
    else:
        clauses = cd1_clauses(toks)
        cols = cd1_columns(clauses) if clauses else None
        if cols is None:
            return None
    slots, prev = [], None
    for i, c in enumerate(cols):
        if c is None:                                                       # an age group the text does not name: before the first named group = declined, after it = the previous cap continues
            if cells is not None and cd2(cells[i])[0] not in ("blank", "D", "err"):
                slots.append(("cap", [100.0], "unnamed"))                   # ... unless its payout cell HAS a payout: no limit is named for it, so it is written for the whole range and emit_cv says so
                continue
            slots.append(prev if prev is not None else ("D",))
        else:
            slots.append(c[0])
            prev = c[0]
    if any(c and c[1] == "pair" for c in cols):                             # '85%90%': two caps whose slash is missing - read as a pair only where the payout cells are pairs too
        if cells is None:
            return None
        pairs = 0
        for i, c in enumerate(cols):
            if c and c[1] == "pair":
                k = cd2(cells[i])
                if k[0] == "rate" and len(k[1]) == 2:
                    pairs += 1
                elif k[0] not in ("blank", "D", "err"):
                    return None
        if not pairs:
            return None
    return slots


def cd1_show(slot):
    """a slot as text for the review workbook: 'D', '85', '85/90', 'addon(85/90|60)', 'ncb(75|65)', '20.1-70', '100 (no limit named)'"""
    n = lambda xs: "/".join(f"{x:g}" for x in xs)
    if slot[0] == "D":
        return "D"
    if slot[0] == "cap":
        return n(slot[1]) + (" (no limit named)" if len(slot) > 2 else "")
    if slot[0] == "band":
        return f"{slot[1]:g}-{slot[2]:g}"
    if slot[0] == "addon":
        return f"addon({n(slot[1])}|{n(slot[2])})"
    return f"ncb({slot[1]:g}|{slot[2]:g})"


def cd1_readings_frame():
    """the CD1 wordings the age-group grammar read in this run, with the cap of each of the eight age groups (an audit list: nothing in it needs an answer)"""
    return pd.DataFrame([dict(text=t, grid_rows=n, **{AGES[i][0]: cd1_show(sl_) for i, sl_ in enumerate(slots)}) for t, (slots, n) in sorted(CD1_AUTO.items())])


def cd1_slots(v, cells=None):
    """-> (8 slots, the 'TATA age 0 - ...' clause or None); the slots are None when the text is not understood.  `cells` = the row's eight payout cells (see cd1_read)."""
    if v is None or norm(v) == "":
        return [("cap", [100.0])] * 8, None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return [("cap", [rnd(v * 100 if v <= 1 else v)])] * 8, None
    t = norm(v).replace("–", "-").replace("—", "-")
    tata = None
    m = TATA_CLAUSE_RX.search(t)
    if m:
        tata, t = m.group(1).strip(), t[:m.start()].strip()
    one = slot_of(t)
    if one:
        return [one] * 8, tata
    slots = cd1_read(t, cells)
    if slots is not None:
        seen = CD1_AUTO.setdefault(t, [slots, 0])
        seen[1] += 1
    return slots, tata


# ---- the segment texts: weight bands, body types, makes
GVW_BANDS = {
    "gcv412to20t": (12000.01, 20000.01), "gcv420to40t": (20000.01, 40000.01), "gcv440to44t": (40000.01, 44000.01), "gcv440t+": (40000.01, CC_OPEN_END), "gcv444t+": (44000.01, CC_OPEN_END),
    "gcv47.5to12t": (7500.01, 12000.01), "gcv43.5to7.5t": (3500.01, 7500.01), "gcv42.5to3.5t": (2500.01, 3500.01), "gcv41.6tto2.5t": (1600.01, 2500.01), "gcv4upto1.6t": (0, 1600.01),
    "gcv4upto2.5t": (0, 2500.01), "gcvupto2.5t": (0, 2500.01), "gcv42tto2.5t": (2000.01, 2500.01), "gcv41.6tto2t": (1600.01, 2000.01),
}


def gvw_of(seg):
    """a weight-band label -> (from kg, to kg): the table above for the labels it knows, else the same convention read from the words ('GCV4 upto 2T' = 0-2000.01, 'GCV4 40T-43T' = 40000.01-43000.01, 'GCV4 43T+' = 43000.01 and above);
    None for any other label (the row is listed in 'Not written')"""
    key = re.sub(r"\s+", "", seg.lower())
    if key in GVW_BANDS:
        return GVW_BANDS[key]
    s = re.sub(r"\s+", " ", seg.lower().strip())                              # the words need the space after 'GCV4' ('GCV 44T+' is not 'GCV4 4T+')
    kg = lambda x: round(float(x) * 1000 + 0.01, 2)
    m = re.fullmatch(r"gcv4? upto ?(\d+(?:\.\d+)?) ?t", s)
    if m:
        return (0, kg(m.group(1)))
    m = re.fullmatch(r"gcv4 (\d+(?:\.\d+)?) ?t? ?(?:to|-) ?(\d+(?:\.\d+)?) ?t", s)
    if m and float(m.group(1)) < float(m.group(2)):
        return (kg(m.group(1)), kg(m.group(2)))
    m = re.fullmatch(r"gcv4 (\d+(?:\.\d+)?) ?t ?\+", s)
    if m:
        return (kg(m.group(1)), CC_OPEN_END)
    return None


WEIGHT_CLAUSE = re.compile(rf"(?:(?P<lo>{_CD1_NUM})\s*(?:t\s*)?(?:to|-)\s*(?P<hi>{_CD1_NUM})\s*t|(?P<plus>{_CD1_NUM})\s*t\s*\+)\s*:\s*(?P<v>d\b|{_CD1_NUM}\s*%)", re.I)


def weight_cell(text, band):
    """A payout cell that names its payout per weight band ('40 to 43T : 2.5% 43T+: D') on the row of ONE band: the clause of that very band is the payout (cd2 of it); None when the text is anything else, a clause
    is not understood, or no clause (or more than one) is for this row's band"""
    t = norm(text)
    pos, found = 0, []
    while pos < len(t):
        m = WEIGHT_CLAUSE.match(t, pos)
        if not m:
            return None
        if m.group("plus") is not None:
            b = (round(float(m.group("plus")) * 1000 + 0.01, 2), CC_OPEN_END)
        else:
            b = (round(float(m.group("lo")) * 1000 + 0.01, 2), round(float(m.group("hi")) * 1000 + 0.01, 2))
        found.append((b, m.group("v")))
        pos = m.end()
        while pos < len(t) and t[pos] == " ":
            pos += 1
    mine = [v for b, v in found if band is not None and (round(b[0], 2), b[1]) == (round(band[0], 2), band[1])]
    return cd2(mine[0]) if len(mine) == 1 else None


def body_class(body, make):
    """body type (and, for medium vehicles, the make column) -> (make_model_class, add_tnc)"""
    b = body.lower()
    if "non-dumper" in b:
        return "Truck", ""
    if "dumper" in b:
        return "Dumper Tipper", ""
    if "oil tanker" in b:
        return "Tanker", "Oil Tanker"
    if "gas tanker" in b:
        return "Tanker", "Gas Tanker"
    if "port trailer" in b:
        return "Trailer", "Port Trailer"
    if make.lower() == "oil tanker":
        return "Tanker", "Oil Tanker"
    return "NA", ""


CLUSTER_NAME_EXCEPTIONS = {"MUMBAI2": "Mumbai_2"}                           # a grid cluster the portal's group spells differently (P-9): the grid's 'Mumbai2' is the group 'Mumbai_2'
FUELS_ALL = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NON_EV = ["Petrol", "Diesel", "CNG", "LPG"]
TAXI_MAKES = {                                                              # the make column of the taxi rows -> the master's make group
    "all": "TAXI_ALL", "kaali peeli": "TAXI_ALL", "maruti & hyundai (excl wagonr model)": "MARUTI_EX_WAGONR,Hyundai", "other than maruti & hyundai": "TAXI_ALL_EX_MARUTI_HYUNDAI",
    "other thanmaruti & hyundai": "TAXI_ALL_EX_MARUTI_HYUNDAI", "maruti & hyundai": "Maruti,Hyundai", "maruti & hyundai(wagonr included)": "Maruti,Hyundai",
    "maruti & hyundai (wagonr included)": "Maruti,Hyundai", "maruti": "Maruti", "other": "TAXI_ALL_EX_MARUTI_HYUNDAI,Hyundai"}
CNG_FUELS = ("CNG", "LPG")                                                  # a make / body text that says CNG: LPG takes the CNG rate (the user's convention since the October upload: 'PCV3 Petrol/CNG 6+1', car third party)
GOODS_MAKE_READINGS = {                                                     # wnorm(make column of a CV LCV / CV MCV row) -> (make group(s) in the master's spelling, the fuels the text limits it to or None = every fuel)
    "": ("GCV_4W_ALL", None), "all": ("GCV_4W_ALL", None), "all make": ("GCV_4W_ALL", None),
    "tata": ("TATA", None), "non-tata": ("GCV_4W_ALL_EX_TATA", None), "non tata": ("GCV_4W_ALL_EX_TATA", None),
    "tata ace": ("DIGIT_CV_TATAACE", None), "tata intra": ("DIGIT_CV_TATAINTRA", None),
    "mahindra": ("Mahindra", None), "mahindra and al": ("Mahindra,Ashok", None), "non mahindra": ("GCV_4W_ALL_EX_MAHINDRA", None),         # 'AL' = Ashok Leyland = the master's make Ashok
    "all(except tata,mahindra and al)": ("GCV_4W_ALL_EX_TATA_MAHINDRA_ASHOK", None),
    "others": ("GCV_4W_ALL_EX_TATA_MAHINDRA_ASHOK,Ashok", None),
    "all make petrol/cng": ("GCV_4W_ALL", ("Petrol",) + CNG_FUELS), "all make cng": ("GCV_4W_ALL", CNG_FUELS),
    "oil tanker": ("GCV_4W_ALL", None),                                       # October's MCV rows name the body in the make column: body_class makes it class Tanker / T&C Oil Tanker
}
HCV_MAKE_READINGS = {                                                       # wnorm(make column of a CV HCV row) -> 'all' | 'tata' | 'non_tata' (the Volvo / Scania block of the remarks sheet is applied to all of them)
    "": "all", "all": "all", "tata": "tata", "non tata": "non_tata", "non-tata": "non_tata", "all (except tata)": "non_tata", "other makes": "non_tata",
    "all excluding volvo and scania": "all", "all excluding volvo and scania/ non dumper/tipper": "all", "all excluding volvo and scania(non dumper/tipper)": "all",
    "all excluding volvo and scania/ non dumper tipper": "all",
}
HCV_BODIES = {"non-dumper/tipper", "dumper/tipper", "dumper tipper", "oil tanker", "gas tanker", "port trailer"}
PCV23_MAKE_READINGS = {                                                     # wnorm(make of a 'CV PCV2W(All) 3W Non-electric' row) -> {product: make group or None = no row for that product}
    "": {"PCV 2W": "PCV_2W_ALL", "PCV 3W": "PCV_3W_ALL"}, "all": {"PCV 2W": "PCV_2W_ALL", "PCV 3W": "PCV_3W_ALL"}, "all make": {"PCV 2W": "PCV_2W_ALL", "PCV 3W": "PCV_3W_ALL"},
    "all excl piaggio": {"PCV 2W": "PCV_2W_ALL_EX_PIAGGAO", "PCV 3W": "PCV_3W_ALL_EX_PIAGGAO"}, "all excluding piaggio": {"PCV 2W": "PCV_2W_ALL_EX_PIAGGAO", "PCV 3W": "PCV_3W_ALL_EX_PIAGGAO"},
    "other than bajaj": {"PCV 2W": "PCV_2W_ALL_EX_BAJAJ", "PCV 3W": "PCV_3W_ALL_EX_BAJAJ"},
    "bajaj and tvs": {"PCV 2W": "BAJAJ,TVS", "PCV 3W": "PCV_3W_BAJAJ_TVS"}, "bajaj and tvs only": {"PCV 2W": "BAJAJ,TVS", "PCV 3W": "PCV_3W_BAJAJ_TVS"},
    "bajaj re": {"PCV 2W": None, "PCV 3W": "PCV_3W_BAJAJ_RE"},                # the RE is Bajaj's three-wheeler: no two-wheeler row
}
PCV2_SEGMENTS = {"pcv2w", "pcv2 - without zd"}                                  # the two-wheeler rows of the combined segment: written for PCV 2W only
PCV23_SEGMENTS = PCV2_SEGMENTS | {"pcv3w non-diesel", "pcv3w diesel", "pcv3 petrol/cng 6+1"}
E3W_MAKE_LIST = "bajaj, tvs, piaggio, omega, hykon, montra electric, mahindra, atul auto, euler motors pvt ltd, greaves electric mobility"   # the list the make groups E3W_ALL / E3W_EX are made for
E3W_ALL = "GCV_3W_ALL_Bajaj_TVS_Piaggio_Omega_Montra_Mahindra_Atul_others"
E3W_EX = "GCV_3W_ALL_EX_Bajaj_TVS_Piaggio_Omega_Montra_Mahindra_Atul_others"
HCV_TRUCK_OTHERS = "GCV_4W_ALL_EX_VOLVO_TATA_MAHINDRA_BHARAT"             # every make but Volvo, Tata, Mahindra and Bharat Benz ('...,TATA' adds Tata back)
HCV_TRUCK_70 = "Mahindra,Bharat_Benz"


MAKE_DEFINITIONS = {"TAXI_ALL_EX_HINDUSTAN": ("TAXI_ALL", ["Hindustan"]),   # 'Others (Other than Hindustan make)': a taxi group the master may lack = every taxi make but Hindustan (how master_builder.py builds it; listed in 'Make groups to create')
                    "GCV_4W_ALL_EX_MAHINDRA": ("GCV_4W_ALL", ["Mahindra"]),                         # 'Non Mahindra' (LCV)
                    "PCV_2W_ALL_EX_BAJAJ": ("PCV_2W_ALL", ["BAJAJ"]),                                # 'Other than Bajaj' (PCV 3W diesel rows, their two-wheeler side)
                    "PCV_3W_BAJAJ_TVS": ("PCV_3W_ALL", ["PCV_3W_ALL_EX_BAJAJ_TVS"]),                  # 'Bajaj and TVS (only)': the three-wheeler models of Bajaj and TVS = PCV_3W_ALL without the group that excludes them
                    "PCV_3W_BAJAJ_RE": ("PCV_3W_ALL", ["PCV_3W_ALL_Ex_BajajRE"])}                    # 'Bajaj RE'
TAXI_MAKE_WORDS = {"maruti": "Maruti", "hyundai": "Hyundai", "hindustan": "Hindustan"}      # the makes the words of a taxi make text may name -> the master's make group
TAXI_INCLUDED_MODELS = {"wagonr", "ertiga", "xl6", "xi6", "eco", "eeco", "invicto"}          # Maruti models a text may say it includes ('Maruti & Hyundai (including Ertiga, Maruti XI6, ECO, Invicto)'): the group Maruti holds them already
_TAXI_MODIFIERS = [                                                           # (regex, what it does) of the words after a make: '(all age)', '(excl WagonR Model)', '(WagonR included)', 'including ...', 'age 0', 'age 1+', 'age 2-5'
    (re.compile(r"\(\s*all\s*ages?\s*\)"), "all"),
    (re.compile(r"\(\s*excl(?:uding|\.)?\s*wagon\s*-?\s*r(?:\s*models?)?\s*\)"), "exwagonr"),
    (re.compile(r"\(?\s*wagon\s*-?\s*r\s+included\s*\)?"), "nop"),
    (re.compile(r"\(?\s*including\s+(?P<models>[^)]*)\)?"), "including"),
    (re.compile(r"age\s*(?P<lo>\d+)\s*\+"), "plus"),
    (re.compile(r"age\s*(?P<lo>\d+)\s*-\s*(?P<hi>\d+)(?![\d.])"), "range"),
    (re.compile(r"age\s*(?P<lo>\d+)(?![\d.+-])"), "exact"),
]


def taxi_age_columns(first, last):
    """the ages first..last (None = and above) -> the set of age-column numbers they cover, or False when a column is only partly inside (None = every column)"""
    ages = set(range(first, (CD1_AGE_TOP if last is None else min(last, CD1_AGE_TOP)) + 1))
    cols = set()
    for i, col in enumerate(CD1_COLUMN_AGES):
        if col <= ages:
            cols.add(i)
        elif col & ages:
            return False
    if not cols:
        return False
    return None if len(cols) == len(CD1_COLUMN_AGES) else frozenset(cols)


def taxi_make_caps(v):
    """'Hyundai: 70%, Maruti 85%' (a cap per make) -> {'hyundai': 70.0, 'maruti': 85.0}; None for any other text"""
    t = cd1_normalize(v).lower() if v is not None else ""
    if not re.fullmatch(r"(?:[a-z]+\s*:?\s*\d+(?:\.\d+)?\s*%\s*[,;&]?\s*)+", t):
        return None
    out = {}
    for word, n in re.findall(r"([a-z]+)\s*:?\s*(\d+(?:\.\d+)?)\s*%", t):
        if word not in TAXI_MAKE_WORDS or word in out or float(n) > 100:
            return None
        out[word] = rnd(n)
    return out


def taxi_make_parts(make_text, cd1_text=None):
    """The make column of a taxi row in words -> {"entries": [(make group, age columns or None = every age)], "slots": {group: 8 slots} or None}, or None when a word is not understood (the row is then
    written for TAXI_ALL and flagged).  'Maruti & Hyundai age 1+' = Maruti and Hyundai from age 1; 'Hyundai Age 0' = new vehicles only; 'Maruti (all age) (excl WagonR Model) & Hyundai age 1+' = the master's
    group MARUTI_EX_WAGONR at every age and Hyundai from age 1; 'Maruti & Hyundai (including Ertiga, ...)' = Maruti and Hyundai; 'Others (Other than Hindustan make)' = TAXI_ALL_EX_HINDUSTAN (a group the
    master may lack: check_names then lists it).  A CD1 text that gives a cap per make ('Hyundai: 70%, Maruti 85%') makes one entry (and one cap) per make."""
    t = re.sub(r"\s+", " ", str(make_text).lower().replace("*", "")).strip()
    m = re.fullmatch(r"others?\s*\(\s*other\s+than\s+(?P<names>[a-z ,&]+?)\s*(?:makes?)?\s*\)", t)
    if m:
        words = [w for w in re.split(r"\s*&\s*|\s*,\s*|\s+and\s+", m.group("names")) if w]
        if not words or any(w not in TAXI_MAKE_WORDS for w in words):
            return None
        return dict(entries=[("TAXI_ALL_EX_" + "_".join(w.upper() for w in words), None)], slots=None)
    hidden = []

    def hide(mm):
        hidden.append(mm.group(0))
        return f"\x00{len(hidden) - 1}\x00"
    t2 = re.sub(r"\([^)]*\)", hide, t)
    pieces = [re.sub(r"\x00(\d+)\x00", lambda mm: hidden[int(mm.group(1))], p).strip() for p in re.split(r"\s*&\s*|\s+and\s+|\s*,\s*", t2) if p.strip()]
    parts = []
    for piece in pieces:
        m = re.match(r"(maruti|hyundai|hindustan)\b\s*(.*)$", piece)
        if not m:
            return None
        word, rest, group, window, explicit = m.group(1), m.group(2).strip(), TAXI_MAKE_WORDS[m.group(1)], None, False
        while rest:
            for rx, what in _TAXI_MODIFIERS:
                mm = rx.match(rest)
                if mm:
                    break
            else:
                return None
            if what == "all":
                explicit = True
            elif what == "exwagonr":
                if word != "maruti":
                    return None
                group = "MARUTI_EX_WAGONR"
            elif what == "including":
                models = [re.sub(r"[^a-z0-9]", "", x.replace("maruti", "")) for x in re.split(r"\s*,\s*|\s*&\s*|\s+and\s+", mm.group("models").strip())]
                if not models or any(x not in TAXI_INCLUDED_MODELS for x in models if x):
                    return None
            elif what in ("plus", "range", "exact"):
                lo = int(mm.group("lo"))
                hi = None if what == "plus" else (int(mm.group("hi")) if what == "range" else lo)
                if hi is not None and hi < lo:
                    return None
                window, explicit = taxi_age_columns(lo, hi), True
                if window is False:
                    return None
            rest = rest[mm.end():].strip()
        parts.append((word, group, window, explicit))
    if not parts or len({p[0] for p in parts}) != len(parts):
        return None
    if len(parts) > 1 and parts[-1][3]:                                          # 'Maruti & Hyundai age 1+': an age window after the last make is for every make of the list that names no ages of its own
        parts = [(w, g, win if ex else parts[-1][2], True) for w, g, win, ex in parts]
    parts = [(w, g, win) for w, g, win, _ in parts]
    caps = taxi_make_caps(cd1_text) if cd1_text is not None else None
    if caps is not None:
        if set(caps) != {w for w, _, _ in parts}:
            return None
        return dict(entries=[(g, win) for _, g, win in parts], slots={g: [("cap", [caps[w]])] * 8 for w, g, _ in parts})
    merged = {}
    for _, g, win in parts:
        merged.setdefault(win, []).append(g)
    return dict(entries=[(",".join(gs), win) for win, gs in merged.items()], slots=None)


# ---- the WORDS of the make / body / segment columns of the commercial-vehicle grid  (the user, 18 Nov 2026: "at some places there is some make or other conditions which is not applied, also it is not
# mentioned anywhere").  Until then a wording the code did not know fell through to a default - all makes, all fuels, no CPA, any cc - in silence.  Now every wording is READ by an exact table or a
# refusing pattern below; one that nothing reads is still written the way the default did, AND flagged on exactly its rows (`unread` of the spec) and listed in the review sheet 'Wording readings'.
# A new wording of a new month: add it to the tables (with a test in tests/godigit_wording_tests.py); until then it is flagged, never silent.
WORDINGS = {}                                    # (cluster segment, segment, body, make) -> dict(rows, written, unread) of this run: the review sheet 'Wording readings' (cleared by build)


def wnorm(text):
    """a make / body / segment wording in one comparable shape: lower case, single spaces, footnote stars dropped, dashes unified; a lone dash (= no text) is ''"""
    t = str(text if text is not None else "").lower().replace("*", "").replace("–", "-").replace("—", "-").replace("\xa0", " ")
    t = re.sub(r"\s+", " ", t).strip()
    return "" if t.strip("- ") == "" else t


CPA_NO = re.compile(r"\b(?:w/o|without|with no|no)\s+cpa\b")
CPA_YES = re.compile(r"\bwith\s+cpa\b")


def cpa_of(body):
    """the CPA a body text names: 'with CPA' / 'With CPA and ...' = YES, 'w/o CPA' / 'without CPA' / 'With No CPA and ...' = NO, None when it names none"""
    t = wnorm(body)
    if CPA_NO.search(t):
        return "NO"
    return "YES" if CPA_YES.search(t) else None


TAXI_SEGMENT = re.compile(r"taxi upto 5 seater(?: (?P<fuel>all|diesel|non-diesel|electric))?"
                          r"(?:\s*(?P<c1>[<>])\s*1000\s*cc(?:\s*&\s*(?P<c2><=)\s*1500\s*cc)?(?:\s*\(wagonr included\))?)?")      # '... > 1000 cc & <=1500 cc' with or without the spaces, '(WagonR Included)'
TAXI_FUEL_WORDS = {"all": list(FUELS_ALL), "diesel": ["Diesel"], "non-diesel": ["CNG", "LPG", "Petrol"], "electric": ["Electric"]}
TAXI_BODY_FUELS = {"diesel": ["Diesel"], "non diesel": ["CNG", "LPG", "Petrol"], "non diesel/ electric": ["CNG", "LPG", "Petrol", "Electric"]}      # the body column of a taxi row that names no fuel in its segment
TAXI_BODIES = {"", "with add on cover", "without add on cover", "with & without add on cover"} | set(TAXI_BODY_FUELS)


def read_wordings(g, sp, ctx):
    """Refine the spec of a grid row (sp, from segment_spec) from the WORDS of its segment / body / make columns, and collect in sp['unread'] what no table or pattern reads: the texts of the remarks column.
    What a wording means is decided here and nowhere else: makes -> sp['goods_make'] / sp['pcv23_makes'] / sp['makes'], fuels, CPA, cc band."""
    s, b, body, mk = g.seg, wnorm(g.body_seg), wnorm(g.body), wnorm(g.make)
    unread = sp.setdefault("unread", [])
    if sp.get("extra_flag"):
        unread.append(sp.pop("extra_flag"))

    def miss(what, text, written):
        unread.append(f"{what} '{norm(text)}' is not read by this script: {written}")

    if s == "CV HCV":
        if body not in HCV_BODIES:
            miss("body", g.body, "written without a body class")
        if mk not in HCV_MAKE_READINGS:
            miss("make", g.make, "written for all makes but Volvo / Scania")
    elif s in ("CV LCV", "CV MCV"):
        if body not in ({"", "lcv"} if s == "CV LCV" else {"", "oil tanker", "oil tanker cng"}):
            miss("body", g.body, "written for every fuel")
        got = GOODS_MAKE_READINGS.get(mk)
        if got is None:
            miss("make", g.make, "written for all makes (GCV_4W_ALL)")
        else:
            sp["goods_make"] = got[0]
        make_fuels, body_fuels = (got[1] if got else None), (list(CNG_FUELS) if body == "oil tanker cng" else None)
        if make_fuels and body_fuels and set(make_fuels) != set(body_fuels):
            miss("fuels in the make and the body text", f"{g.make} / {g.body}", "written for the fuels of the make text")
        pick = make_fuels or body_fuels
        if pick:
            sp["fuels"] = list(pick)
    elif s == "CV GCV3(EV/Non-EV)":
        if body not in {"", "with cpa", "w/o cpa", "without cpa"}:
            miss("body", g.body, "written without a CPA condition")
        if cpa_of(body):
            sp["is_cpa"] = cpa_of(body)
        if mk not in ("", "all", "all make"):
            miss("make", g.make, "written for all makes (GCV_3W_ALL)")
        if b not in ("gcv3", "e-loaders"):
            miss("segment", g.body_seg, "written for every fuel")
    elif s == "CV MiscD":
        known = {"": "MISCD_ALL", "all": "MISCD_ALL", "all make": "MISCD_ALL", "jcb": "JCB_ALL", "all make (except hyundai)": "MISDALL_EX_HYUNDAI"}
        if mk in known:
            sp["makes"] = [(known[mk], None)]
        else:
            miss("make", g.make, "written for all makes (MISCD_ALL)")
        if body:
            miss("body", g.body, "written without a body condition")
        if b not in ("misc d", "jcb", "backhoe loader, forklift, excavator, and loader"):
            miss("segment", g.body_seg, "written for every fuel")
    elif s == "CV Tractor":
        known = {"": "MISCD_ALL", "all": "MISCD_ALL", "all make": "MISCD_ALL", "sonalika make": "SONALIKA", "all - excl sonalika": "MISCD_ALL_EX_SONALIKA"}
        if mk in known:
            sp["makes"] = [(known[mk], None)]
        else:
            miss("make", g.make, "written for all makes (MISCD_ALL)")
        if body:
            miss("body", g.body, "written without a body condition")
        if b == "tractor electric and low hp":
            unread.append(f"segment '{norm(g.body_seg)}': the electric / low-horsepower condition is not applied (the upload has no horsepower field): written for every fuel and horsepower")
        elif b != "tractor":
            miss("segment", g.body_seg, "written for every fuel")
    elif s == "CV PCV2W(All) 3W Non-electric":
        if b not in PCV23_SEGMENTS:
            miss("segment", g.body_seg, "written for the non-electric fuels")
        elif b in PCV2_SEGMENTS:
            sp["products"] = ("PCV 2W",)                                             # a row that says PCV2 is for the two-wheeler: it must not pull the PCV 3W cards of its cluster down (Andaman: PCV2W 20% vs PCV3W 64%)
            if b == "pcv2 - without zd":
                sp["is_zd"] = "NO"                                                    # 'without ZD' = the comprehensive card without zero dep (the STP rows carry NA, the STP rule)
        elif g.cluster in ctx.pcv2_clusters:
            sp["products"] = ("PCV 3W",)                                              # the cluster has its own PCV2W row: the PCV3W rows are the three-wheeler's alone (they are also written for PCV 2W where the cluster has no row of its own, as in the October upload)
        if body not in {"", "with cpa", "without cpa"}:
            miss("body", g.body, "written without a CPA condition")
        if cpa_of(body):
            sp["is_cpa"] = cpa_of(body)
        got = PCV23_MAKE_READINGS.get(mk)
        if got is None:
            miss("make", g.make, "written for all makes (PCV_2W_ALL / PCV_3W_ALL)")
            got = PCV23_MAKE_READINGS[""]
        sp["pcv23_makes"] = got
    elif s == "PCV 3W Electric":
        if b not in ("e-auto", "e-rickshaw"):
            miss("segment", g.body_seg, "written for Electric")
        lst = E3W_MAKE_LIST
        if mk == lst:
            sp["makes"] = [(E3W_ALL, None)]
        elif mk == f"all make (except {lst})":
            sp["makes"] = [(E3W_EX, None)]
        elif mk == "":
            sp["makes"] = [("PCV_3W_ALL", None)]
        else:
            miss("make", g.make, "written for all makes (PCV_3W_ALL)")
            sp["makes"] = [("PCV_3W_ALL", None)]
        if cpa_of(body):
            sp["is_cpa"] = cpa_of(body)
        if "tppd" in body:
            unread.append(f"body '{norm(g.body)}': the 'TPPD reduction' part is not applied (the upload has no field for it): rows that differ only in it share one card and the lowest payin is written")
        if "consider all rtos" in body:
            unread.append(f"body '{norm(g.body)}': the sentence 'Consider all RTOs as Good UK for E-Rickshaw' is not applied (which RTOs get which cluster's rate is not in the upload)")
        elif body and not cpa_of(body) and "tppd" not in body:
            miss("body", g.body, "written without a CPA condition")
    elif s == "CV PCV4(Upto 5 seating)":
        m = TAXI_SEGMENT.fullmatch(b)
        if m is None:
            miss("segment", g.body_seg, "written with the cc and fuels the default reading gives")
        else:
            c1, c2 = m.group("c1"), m.group("c2")
            sp["cc"] = (0, CC_OPEN_END) if c1 is None else (0, 1000.01) if c1 == "<" else (1000.01, 1500.01) if c2 else (1000.01, CC_OPEN_END)
            fuel, body_fuels = m.group("fuel"), TAXI_BODY_FUELS.get(body)
            if fuel:
                sp["fuels"] = list(TAXI_FUEL_WORDS[fuel])
                if body_fuels and set(body_fuels) != set(sp["fuels"]):
                    miss("fuels in the segment and the body text", f"{g.body_seg} / {g.body}", "written for the fuels of the segment")
            elif body_fuels:
                sp["fuels"] = list(body_fuels)
            else:
                miss("segment", g.body_seg, "names no fuel: written for the non-diesel fuels (CNG, LPG, Petrol)")
        if body not in TAXI_BODIES:
            miss("body", g.body, "written without an add-on cover condition")
    elif s == "Staff Bus":
        if mk != ">10":
            miss("seating text", g.make, f"written for more than {STAFF_BUS_SC[0] - 0.01:g} seats")


def written_text(g, sp, ctx):
    """what the spec makes of a grid row, in words (the review sheet 'Wording readings')"""
    if sp.get("pcv23_makes") is not None:
        makes = "; ".join(f"{p}: {m or 'no row'}" for p, m in sp["pcv23_makes"].items() if p in sp["products"])
    elif sp.get("makes"):
        makes = ", ".join(m for m, _ in sp["makes"])
    else:
        makes = " / ".join(dict.fromkeys(m for sub in ("1+1", "STP") for m, _ in gcv4_makes(g, sp, sub, ctx)))
    parts = [f"makes {makes}"]
    if sorted(sp["fuels"]) != sorted(FUELS_ALL):
        parts.append("fuels " + ",".join(sp["fuels"]))
    for label, key in (("CPA", "is_cpa"), ("zero dep", "is_zd")):
        if sp.get(key, "NA") != "NA":
            parts.append(f"{label} {sp[key]}")
    if sp.get("klass") not in (None, "NA"):
        parts.append(f"class {sp['klass']}")
    if sp.get("tnc"):
        parts.append(f"T&C {sp['tnc']}")
    if sp.get("cc") and tuple(sp["cc"]) != (0, CC_OPEN_END):
        parts.append(f"cc {sp['cc'][0]:g}-{sp['cc'][1]:g}")
    return " | ".join(parts)


def note_wording(g, sp, ctx):
    """remember the reading of this grid row's wording (one entry per distinct wording)"""
    key = (g.seg, g.body_seg, g.body, g.make)
    w = WORDINGS.get(key)
    if w is None:
        w = WORDINGS[key] = dict(rows=0, written=written_text(g, sp, ctx), unread=list(dict.fromkeys(sp.get("unread", []))))
    w["rows"] += 1


def wording_readings_frame():
    """the review sheet 'Wording readings': every distinct (cluster segment, segment, body, make) wording of the commercial-vehicle grid, what it was written as, and - first - the ones nothing read"""
    rows = [dict(cluster_segment=k[0], segment=k[1], body=k[2], make=k[3], grid_rows=w["rows"], status=("NOT APPLIED - flagged on its rows" if w["unread"] else "read"),
                 written_as=w["written"], not_applied=" | ".join(w["unread"])) for k, w in WORDINGS.items()]
    df = pd.DataFrame(rows, columns=["cluster_segment", "segment", "body", "make", "grid_rows", "status", "written_as", "not_applied"])
    df["_read"] = df["status"] == "read"
    return df.sort_values(["_read", "cluster_segment", "segment", "body", "make"], kind="stable").drop(columns="_read").reset_index(drop=True)


def segment_spec(g, ctx):
    """How a grid row becomes upload rows: product(s), RTO-group token, class, weight band, seating, fuels, flags, makes.  None = not written (the reason is in ctx)."""
    s, b, body, mk = g.seg, g.body_seg, g.body, g.make
    low = mk.lower().replace("*", "").strip()
    sp = dict(sc=SC_RANGE, cc=None, fuels=list(FUELS_ALL), is_zd="NA", is_cpa="NA", tnc="", klass="NA", makes=None, products=None, token=None, kind="gcv4")
    if s in ("CV HCV", "CV LCV", "CV MCV"):
        sp.update(products=("GCV 4W",), token={"CV HCV": "HCV", "CV LCV": "LCV", "CV MCV": "MCV"}[s], cc=gvw_of(b))
        sp["klass"], sp["tnc"] = body_class(body, mk)
        if sp["cc"] is None:
            ctx.not_written.append(dict(sheet="CV grid", grid_row=g.r, cluster=g.cluster, segment=b, reason=f"weight band '{b}' is not one this script knows"))
            return None
    elif s == "CV GCV3(EV/Non-EV)":
        electric = b.lower().startswith("e-")
        sp.update(products=("GCV 3W",), token="GCV3W", cc=(0, CC_OPEN_END), klass="Goods Carrying Rickshaw", fuels=["Electric"] if electric else list(FUELS_NON_EV),
                  tnc="Loaders" if electric else "", makes=[("GCV_3W_ALL", None)], kind="plain")
        if body.lower().startswith("with cpa"):
            sp["is_cpa"] = "YES"
        elif body.lower().startswith("w/o cpa"):
            sp["is_cpa"] = "NO"
    elif s == "CV MiscD":
        sp.update(products=("MISCD",), token="MISD", cc=(0, CC_OPEN_END), klass="MISCD-Others", kind="plain")
        sp["makes"] = [("JCB_ALL" if low == "jcb" else "MISDALL_EX_HYUNDAI" if "except hyundai" in low else "MISCD_ALL", None)]
    elif s == "CV Tractor":
        sp.update(products=("MISCD",), token="MISD-TRACTOR", cc=(0, CC_OPEN_END), klass="MISCD-Tractor", kind="plain")
        sp["makes"] = [("SONALIKA" if low.startswith("sonalika") else "MISCD_ALL_EX_SONALIKA" if "excl sonalika" in low else "MISCD_ALL", None)]
    elif s == "CV PCV2W(All) 3W Non-electric":
        sp.update(products=("PCV 2W", "PCV 3W"), token="PCV2W_3W_NON-EV", cc=(0, CC_OPEN_END), fuels=list(FUELS_NON_EV), kind="pcv23")
        if "diesel" in b.lower() and "non" not in b.lower():
            sp["fuels"] = ["Diesel"]
        if "6+1" in b:
            sp.update(sc=(0, 7.01), fuels=["Petrol", "CNG", "LPG"])
    elif s == "PCV 3W Electric":
        sp.update(products=("PCV 3W",), token="PCV3W_EV", cc=(0, CC_OPEN_END), klass="Passenger Carrying Rickshaw", fuels=["Electric"], kind="plain")
        sp["makes"] = [(E3W_EX if low.startswith("all make (except") else E3W_ALL if low.startswith("bajaj") else "PCV_3W_ALL", None)]
    elif s == "CV PCV4(Upto 5 seating)":
        bl = b.lower()
        sp.update(products=("PCV 4W",), token="PCV4W_UPTO-5", klass="Taxi", sc=(0, 5.01), kind="plain")
        sp["cc"] = (1000.01, 1500.01) if "<= 1500" in bl else (1000.01, CC_OPEN_END) if "> 1000" in bl else (0, 1000.01) if "< 1000" in bl else (0, CC_OPEN_END)
        sp["fuels"] = ["Electric"] if "electric" in bl else ["Diesel"] if ("non-diesel" not in bl and "diesel" in bl) else ["CNG", "LPG", "Petrol"]
        c = body.lower()
        sp["is_zd"] = "YES" if c.startswith("with add on") else "NO" if c.startswith("without add on") else "NA"
        if low in TAXI_MAKES:
            sp["makes"] = [(TAXI_MAKES[low], None)]
        else:
            parts = taxi_make_parts(mk, g.cd1)                                   # the words of the make column: makes, 'age 0' / 'age 1+', '(excl WagonR Model)', a cap per make in the CD1 text
            if parts:
                sp.update(makes=[(grp, None) for grp, _ in parts["entries"]], part_ages=dict(parts["entries"]), part_slots=parts["slots"])
            else:
                sp["makes"] = [("TAXI_ALL", None)]
                sp["extra_flag"] = f"taxi make '{mk}' is not one this script knows: written for TAXI_ALL"
        if low == "kaali peeli":
            sp["tnc"] = "Kaali peeli"
    elif s == "Staff Bus":
        sp.update(products=("PCV 4W",), token="STAFF-BUS", cc=(0, CC_OPEN_END), klass="Other Bus", sc=(STAFF_BUS_SC[0], SC_RANGE[1]), makes=[("BUS_ALL", None)], tnc=body.replace("–", "-").replace("—", "-"),
                  fuels=list(STAFF_BUS_FUELS), kind="plain")
    elif s == "School Bus":
        ctx.not_written.append(dict(sheet="CV grid", grid_row=g.r, cluster=g.cluster, segment=f"School Bus - {body}", reason=(
            "School Bus rates depend on the premium volume (<= 3 L / > 3 L) and the portal has no field for it; they are not in the October upload (WRITE_SCHOOL_BUS)")))
        if not WRITE_SCHOOL_BUS:
            return None
        sp.update(products=("PCV 4W",), token="SCHOOL-BUS", cc=(0, CC_OPEN_END), klass="School Bus", sc=(7.01, SC_RANGE[1]), makes=[("BUS_ALL", None)], tnc=body, fuels=list(STAFF_BUS_FUELS), kind="plain")
    else:
        ctx.not_written.append(dict(sheet="CV grid", grid_row=g.r, cluster=g.cluster, segment=f"{s} / {b}", reason=f"cluster segment '{s}' is not one this script knows"))
        return None
    read_wordings(g, sp, ctx)
    return sp


def gcv4_makes(g, sp, sub, ctx):
    """(make group, discount cap override or None) per make column of a goods-vehicle row.  Heavy: 'Volvo, Scania to block' and, for the Truck body, Mahindra / Bharat Benz at 70% CD1 (CV Remarks)."""
    if g.seg in ("CV LCV", "CV MCV"):
        return [(sp.get("goods_make") or "GCV_4W_ALL", None)]                     # read_wordings: GOODS_MAKE_READINGS ('Mahindra and AL', 'Non Mahindra', 'TATA Intra' ...)
    if g.seg != "CV HCV":
        return [("GCV_4W_ALL", None)]
    rule = ctx.hcv
    kind = HCV_MAKE_READINGS.get(wnorm(g.make))
    if kind == "tata":
        return [("TATA", None)]
    non_tata = kind == "non_tata"
    truck = sp["klass"] == "Truck" and rule.get("cap") and label_key(rule.get("body") or "") == label_key("Non-Dumper/Tipper")
    if truck and sub == "1+1":
        return [(HCV_TRUCK_OTHERS if non_tata else HCV_TRUCK_OTHERS + ",TATA", None), (HCV_TRUCK_70, rule["cap"])]
    if non_tata:
        return [("GCV_4W_ALL_EX_TATA_VOLVO", None)]
    return [(HCV_BLOCKED_MAKES_GROUP if rule.get("blocked") else "GCV_4W_ALL", None)]


class Ctx:
    """what the converters share: the masters, the dates, the warnings and the lists for the review workbook"""

    def __init__(self, warnings, dates, masters, mapping, hcv):
        self.warnings, self.dates, self.masters, self.mapping, self.hcv = warnings, dates, masters, mapping, hcv
        self.gaps, self.not_written, self.counts = {}, [], {}
        self.pcv2_clusters = set()                                    # clusters whose grid has its own PCV2W row (read_wordings: the PCV3W rows of such a cluster are for PCV 3W only)
        self._rto = {}
        for n in masters.rto:
            self._rto.setdefault(Masters._fold(n), []).append(n)
        self._mk = {}
        for n in masters.makes:
            self._mk.setdefault(Masters._fold(n), []).append(n)

    def count(self, what, n=1):
        self.counts[what] = self.counts.get(what, 0) + n

    def _find(self, table, known, name):
        if name in known:
            return name
        hits = table.get(Masters._fold(name), [])
        return hits[0] if hits else None

    def rto(self, token, cluster, seg_header, prefix=None, working=""):
        """the master's spelling of  GO-DIGIT_OCT26_<token>_<cluster>  ; a group the master lacks is kept by the name it would have and remembered with the RTO codes it should hold.
        `working` = the name the person typed in the RTO working-list column of the grid row: used exactly as it is, no lookup"""
        if working:
            return working_rto(working)
        cand = (prefix or GROUP_PREFIX) + token + "_" + CLUSTER_NAME_EXCEPTIONS.get(re.sub(r"\s+", "_", cluster.strip()).upper(), re.sub(r"\s+", "_", cluster.strip()))
        hit = self._find(self._rto, self.masters.rto, cand)
        if hit:
            return rto_cluster(hit, cluster)                    # the CSV carries the bare cluster of the grid (apply_cluster_names)
        if cand not in self.gaps:
            self.gaps[cand] = sorted(set(self.mapping.get((label_key(seg_header), label_key(cluster)), [])))
        return rto_cluster(cand, cluster)

    def rto_pc(self, prefix, cluster, working=""):
        """the private-car groups are  GO-DIGIT_PC_1+1_OCT26_<cluster>  /  ..._PC_SAOD_OCT26_  /  ..._PC_TP_OCT26_"""
        return self.rto(prefix, cluster, "4W", prefix="GO-DIGIT_", working=working)

    def makes(self, text):
        """a make group or a comma list of them -> the master's spelling of each name (an unknown name stays as written; check_names flags it)"""
        return ",".join(self._find(self._mk, self.masters.makes, t.strip()) or t.strip() for t in text.split(","))


def emit_cv(g, ctx):
    """All upload rows of one grid row (before fuels are expanded, duplicates settled and the discount bands chained)."""
    sp = segment_spec(g, ctx)
    if sp is None:
        return []
    flags = []
    slots, tata = cd1_slots(g.cd1, g.comp)
    if slots is None and sp.get("part_slots"):
        slots = [("cap", [100.0])] * 8                                            # a cap per make in the CD1 text: taxi_make_parts read it, the caps are applied make by make below
    elif slots is None:
        slots = [("cap", [100.0])] * 8
        flags.append(f"CD1 '{norm(g.cd1)}' is not a discount text this script knows: written for the whole discount range (0-{TARIFF_FULL:g})")
    unnamed = [AGES[i][0] for i, s_ in enumerate(slots) if len(s_) > 2 and s_[2] == "unnamed"]
    if unnamed:
        flags.append(f"CD1 '{norm(g.cd1)}' does not name the age group(s) {', '.join(unnamed)}, whose payout is written for the whole discount range (0-{TARIFF_FULL:g})")
    flags.extend(sp.get("unread", []))                                              # the wordings of the segment / body / make columns that nothing reads (read_wordings)
    note_wording(g, sp, ctx)
    rows = []
    for sub, cells in (("1+1", g.comp), ("STP", g.satp)):
        for a, (lab, lo, hi) in enumerate(AGES):
            cell = cd2(cells[a])
            if cell[0] == "raw" and sp.get("cc"):                                   # '40 to 43T : 2.5% 43T+: D' on the row of the 43T+ band: the clause of this row's own weight band
                cell = weight_cell(cell[1], sp["cc"]) or cell
            sl = slots[a]
            if cell[0] == "raw":
                ctx.not_written.append(dict(sheet="CV grid", grid_row=g.r, cluster=g.cluster, segment=f"{g.body_seg} / {g.body} / {g.make}", reason=f"{sub} {lab}: the cell text '{cell[1]}' is not a payout this script knows"))
                flags.append(f"{sub} {lab} cell '{cell[1]}' not understood: no row written for that age")
                continue
            if cell[0] == "err":
                ctx.count("error cell")
            makes = sp["makes"] or gcv4_makes(g, sp, sub, ctx)
            if sp.get("part_ages"):                                                 # 'Hyundai Age 0', 'Maruti & Hyundai age 1+': a make text that names the ages it is for writes only those age columns
                makes = [(m_, c_) for m_, c_ in makes if sp["part_ages"].get(m_) is None or a in sp["part_ages"][m_]]
            tw = slot_of(tata) if tata else None
            if tata and a == 0 and sub == "1+1" and sp["klass"] == "Truck" and cell[0] == "tata" and g.seg == "CV HCV" and ctx.hcv.get("cap"):
                makes = [(HCV_TRUCK_OTHERS, None), (HCV_TRUCK_70, ctx.hcv["cap"]), ("TATA", "tata")]
            for mk, capov in makes:
                outs = []                                                       # (cap or None, rate, placeholder?, extra)
                sl = sp["part_slots"][mk][a] if sp.get("part_slots") and mk in sp["part_slots"] else slots[a]
                if capov == "tata" and tw and tw[0] == "addon":
                    for i, c in enumerate(tw[1]):
                        outs.append((c, cell[2][min(i, len(cell[2]) - 1)], False, dict(is_zd="YES")))
                    for i, c in enumerate(tw[2]):
                        outs.append((c, cell[3][min(i, len(cell[3]) - 1)], False, dict(is_zd="NO")))
                    capov = None
                elif sub == "STP":
                    if cell[0] == "rate":
                        outs.append((None, cell[1][0], False, {}))
                    elif cell[0] == "blank":
                        outs.append((None, 0.0, True, {}))
                    elif cell[0] in ("D", "err"):
                        outs.append((None, 0.0, False, {}))
                    elif cell[0] in ("addon", "tata"):
                        outs.append((None, (cell[2] if cell[0] == "addon" else cell[3])[0] if (cell[2] if cell[0] == "addon" else cell[3]) else 0.0, False, {}))
                else:
                    base = cell[1] if cell[0] == "tata" else None
                    if cell[0] == "err" or sl[0] == "D":
                        outs.append((100.0, 0.0, False, {}))
                    elif sl[0] == "cap":
                        caps = list(sl[1])
                        if cell[0] == "rate":
                            rs = list(cell[1])
                            if len(caps) >= 2 and len(rs) >= 2:
                                if rs[0] == rs[1]:
                                    outs.append((caps[1], rs[0], False, {}))              # '10%/10%': one band up to the higher cap
                                else:
                                    outs.extend((caps[i], rs[i], False, {}) for i in range(2))
                            else:
                                outs.append((caps[0], rs[0], False, {}))                  # one payout: the first (lower) band only
                        elif cell[0] == "tata":
                            outs.append((caps[0], base, False, {}))
                        elif cell[0] in ("D", "blank"):
                            if len(caps) >= 2:
                                outs.append((100.0, 0.0, False, {}))                       # two bands and no payout: declined for the whole discount range
                            else:
                                outs.append((caps[0], 0.0, cell[0] == "blank", {}))
                        elif cell[0] == "addon":
                            outs.append((caps[0], (cell[2] or cell[1])[0], False, {}))
                    elif sl[0] == "addon":
                        if cell[0] == "addon":                                   # the add-on split is in the cell itself: with / without add-on
                            outs.extend((c, cell[1][min(i, len(cell[1]) - 1)], False, dict(is_zd="YES")) for i, c in enumerate(sl[1]))
                            outs.extend((c, cell[2][min(i, len(cell[2]) - 1)], False, dict(is_zd="NO")) for i, c in enumerate(sl[2]))
                        elif cell[0] == "rate":
                            outs.append((100.0, cell[1][0], False, {}))              # a plain payout under an add-on text: no discount band is named for it
                        else:
                            outs.append((100.0, 0.0, False, {}))                    # no payout for this age: declined
                    elif sl[0] == "ncb":
                        r = cell[1][0] if cell[0] == "rate" else 0.0
                        z = cell[0] == "blank"
                        outs.append((sl[1], r, z, dict(is_ncb="YES", ncb_tnc=NCB_BREAK_TNC)))
                        outs.append((sl[2], r, z, dict(is_ncb="NO")))
                    elif sl[0] == "band":
                        outs.append((sl[2], cell[1][0] if cell[0] == "rate" else 0.0, False, dict(tmin=sl[1])))
                if capov is not None and sub == "1+1" and outs:
                    byflag = {}
                    for o in sorted(outs, key=lambda o: o[0]):                         # 'operate at 70% CD1': that cap, whatever the row says, with the rate of the lowest band (per NCB / add-on variant)
                        byflag.setdefault(tuple(sorted(o[3].items())), o)
                    outs = [(capov, o[1], o[2], o[3]) for o in byflag.values()]
                for cap, rate, z, extra in outs:
                    for prod in sp["products"]:
                        tnc = ",".join(x for x in (("New" if a == 0 else ""), sp["tnc"], extra.get("ncb_tnc", "")) if x)
                        makes_text = ctx.makes(mk)
                        if sp["kind"] == "pcv23":
                            pm = sp["pcv23_makes"].get(prod)
                            if pm is None:
                                continue                                            # a make of the other product only ('Bajaj RE' is a three-wheeler): no row for this product
                            makes_text = ctx.makes(pm)
                        klass = "PCV 2W" if prod == "PCV 2W" else ("Passenger Carrying Rickshaw" if prod == "PCV 3W" and sp["kind"] == "pcv23" else sp["klass"])
                        row = dict(product=prod, sub_product=sub, policy_type="Renew", make_model_class=klass, new_vehicle_makes=makes_text,
                                   new_rto_list=ctx.rto(sp["token"], g.cluster, g.seg, working=g.rto), _grp=sp["token"], vehicle_age_min=lo, vehicle_age_max=hi, cc_min=sp["cc"][0], cc_max=sp["cc"][1],
                                   sc_min=sp["sc"][0], sc_max=sp["sc"][1], is_ncb="NA" if sub == "STP" else extra.get("is_ncb", "NA"), is_cpa=sp["is_cpa"],
                                   is_zd="NA" if sub == "STP" else extra.get("is_zd", sp["is_zd"]), add_tnc=tnc, _r=g.r, _z=z, _fuels=tuple(sp["fuels"]), _flags=tuple(flags),
                                   _d=bool(rate == 0.0 and not z and (cell[0] in ("D", "err") or sl[0] == "D")),                     # a zero card that comes from a D (a decline) - or from a CD1 text that declines this age - not from a number
                                   pi_type="On TP" if sub == "STP" else "On Net", pi_od_rate=0.0, pi_tp_rate=0.0, pi_net_rate=0.0)
                        if sub == "STP":
                            row.update(pi_tp_rate=rate, tariff_max=TARIFF_FULL, _cap=None)
                        else:
                            row.update(pi_net_rate=rate, tariff_max=rnd(cap + 0.01), _cap=cap)
                            if "tmin" in extra:
                                row["_tmin"] = extra["tmin"]
                        rows.append(row)
    return rows


# ---- the other mapping sheets (RTO code -> cluster per product), only used to say which RTOs a missing group would hold
RTO_CODE_SHAPE = re.compile(r"^[A-Za-z]{2}[-\s]?\d{1,3}[A-Za-z]{0,2}$")          # AN01, MH12, DL1C
MAPPING_SKIP = {"city", "state", "district", "rto", "rtocode", "rtoname", "key"}   # columns of a mapping sheet that are not a cluster per product


def read_mapping(sheet):
    """{(column header key, cluster key): [RTO codes]} of a sheet that has a column of RTO codes (header 'RTO Code' / 'RTO_Code' / 'RTO' in its first rows, wherever it sits) and, in its other
    columns, a cluster per product.  A column headed RTO whose cells are not RTO codes (the person's own working-list names) is not a mapping."""
    out = {}
    if sheet is None:
        return out
    for hr in range(1, min(6, len(sheet.rows)) + 1):
        head = sheet.rows[hr - 1]
        codes = [c for c, v in enumerate(head) if label_key(v) in ("rto", "rtocode")]
        if not codes:
            continue
        code_col = codes[0]
        body = [norm(r[code_col]) for r in sheet.rows[hr:hr + 40] if code_col < len(r) and norm(r[code_col])]
        if not body or sum(1 for v in body if RTO_CODE_SHAPE.match(v)) * 2 < len(body):
            continue
        for row in sheet.rows[hr:]:
            code = norm(row[code_col]) if code_col < len(row) else ""
            if not code:
                continue
            for c in range(len(head)):
                key = label_key(head[c])
                if c == code_col or not key or key in MAPPING_SKIP:
                    continue
                cl = norm(row[c]) if c < len(row) else ""
                if cl and cl != "-":
                    out.setdefault((key, label_key(cl)), []).append(code)
        return out
    return out


def read_all_mappings(book, ask):
    """the RTO code -> cluster mapping sheets of the workbook (commercial vehicles, two-wheelers, cars), found by what they say and not by their names"""
    m = {}
    for s in visible_sheets(book):
        got = read_mapping(s)
        if got:
            SHEET_ROLES[s.name] = "mapping"
        m.update(got)
    return m


# ---- Two-wheelers
TW_FUELS = ["Petrol", "Electric"]
MISP_RATE = 17.5                                 # the grid's 'MISP' cells carry no number: the user's October upload wrote On OD 17.5 (po 10.5) and the user accepted the same in November: a standing decision, said in the run's warnings (not flagged per row)
TW_11_SEGMENTS = {                               # 'Agency/PB Seg' (label key) -> (make_model_class, make group, cc band)
    "scev": ("SCOOTER", "SCOOTERALL", (0, CC_OPEN_END)),
    "mc180herohonda": ("BIKE", "HERO,HONDA", (0, 180.01)),
    "mc180herohondatvs": ("BIKE", "HERO,HONDA,TVS", (0, 180.01)),
    "mc180others": ("BIKE", "BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,TVS,SUZUKI", (0, 180.01)),
    "mc180350re": ("BIKE", "ROYAL", (180.01, 350.01)),
    "mc180350hondajawaavenger": ("BIKE", "HONDA_JAVA_AVENGER", (180.01, 350.01)),
    "mc180350others": ("BIKE", "BIKEALL_EX_HONDA_JAVA_AVENGER_ROYAL", (180.01, 350.01)),
    "mc180350otherthanre": ("BIKE", "BIKEALL_EX_ROYAL", (180.01, 350.01)),
    "mc350": ("BIKE", "BIKEALL", (350.01, CC_OPEN_END)),
}
TW_11_OTHERS_EX_TVS = "BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,SUZUKI"
BIKES_OTHER = "BIKEALL_EX_HONDA_SUZUKI_HERO_TVS_BAJAJ_YAMAHA_ROYAL,HERO"
SCOOTERS_OTHER = "SCOOTERALL_EX_HERO_HONDA_TVS_SUZUKI_RE_YAMAHA,HERO"
EV_OTHER_3_7 = "SCOOTERALL_EX_YAMAHA_SUZUKI_BAJAJ,YAMAHA,SUZUKI"
TW_15_MAKES = {"bajaj": "BAJAJ", "honda": "HONDA", "suzuki": "SUZUKI", "tvs": "TVS", "yamaha": "YAMAHA", "royalenfield": "ROYAL"}
CC_LOW, CC_HIGH = (0, 155.01), (155.01, CC_OPEN_END)
TW_15_SEGMENTS = {                               # 'Agency/PB Seg' (label key) -> list of (class, cc band, fuels)
    "mc155": [("BIKE", CC_LOW, TW_FUELS)], "mc155_": [("BIKE", CC_HIGH, TW_FUELS)],
    "mc": [("BIKE", CC_LOW, TW_FUELS), ("BIKE", CC_HIGH, TW_FUELS)],
    "scootermc": [("BIKE", CC_LOW, TW_FUELS), ("BIKE", CC_HIGH, TW_FUELS), ("SCOOTER", (0, CC_OPEN_END), ["Petrol"])],
    "scooter": [("SCOOTER", (0, CC_OPEN_END), ["Petrol"])],
    "350": [("BIKE", (0, 350.01), TW_FUELS)], "350_": [("BIKE", (350.01, CC_OPEN_END), TW_FUELS)],
    "37kw": [("SCOOTER", (3.01, 7.01), ["Electric"])], "3kw": [("SCOOTER", (0, 3.01), ["Electric"])], "kw3": [("SCOOTER", (3.01, CC_OPEN_END), ["Electric"])],
    "kw7": [("SCOOTER", (7.01, CC_OPEN_END), ["Electric"])], "7kw": [("SCOOTER", (0, 7.01), ["Electric"])],
}


def tw15_seg(seg):
    """'MC <=155' -> 'mc155' ; 'MC >155' -> 'mc155_' ; '< =350' -> '350' ; '> 350' -> '350_' ; '< 3 KW' -> '3kw' ; '> 3 KW' -> 'kw3' ; '>7 KW' -> 'kw7' ; '< 7 KW' -> '7kw' ; '3-7 KW' -> '37kw'"""
    t = norm(seg).lower().replace(" ", "")
    if re.fullmatch(r"mc<=?155", t):
        return "mc155"
    if re.fullmatch(r"mc>=?155", t):
        return "mc155_"
    if re.fullmatch(r"<=?350", t):
        return "350"
    if re.fullmatch(r">=?350", t):
        return "350_"
    m = re.fullmatch(r"([<>])=?(\d+)kw", t)
    if m:
        return f"{m.group(2)}kw" if m.group(1) == "<" else f"kw{m.group(2)}"
    return label_key(t)


def tw15_targets(make, seg):
    """-> list of (class, make group, cc band, fuels) for a 1+5 grid row, or [] when the pair is not one this script knows (the row is then listed in 'Not written')"""
    mk, key = label_key(make), tw15_seg(seg)
    segs = TW_15_SEGMENTS.get(key)
    if not segs:
        return []
    out = []
    for klass, cc, fuels in segs:
        if mk in TW_15_MAKES:
            out.append((klass, TW_15_MAKES[mk], cc, fuels))
        elif mk == "others":
            if klass == "BIKE":
                out.append((klass, BIKES_OTHER, cc, fuels))
            elif fuels == ["Electric"]:
                out.append((klass, EV_OTHER_3_7 if cc == (3.01, 7.01) else "SCOOTERALL", cc, fuels))
            else:
                out.append((klass, SCOOTERS_OTHER, cc, fuels))
        else:
            return []
    return out


DISCOUNT_ROUNDING = "sheet"     # the user, 31 Oct 2026 ("it should go as it is written in the sheet"): a discount limit is written as the SHEET SHOWS it. A cell formatted as a whole percent ('0%') shows the
                                # nearest whole percent (49.9915 -> 50, 39.9951 -> 40, 60.2443 -> 60, 67.5 -> 68); a cell with decimals ('0.00%') is written as it is (67.5 stays 67.5).
                                # "nearest" = the first build (every limit to the nearest whole percent, 67.5 -> 68, flagged on the rows), "down" = always the whole percent below (49.79 -> 49)
CD1_ROUNDED = []                # (the grid's percent, the percent written) of every limit rounded in this run: one line in the run's warnings, not a flag on every row


def cd1_cap(v, shows_whole=None):
    """the TW / PC discount limit: a fraction (0.4 = 40%). 0 = no limit named (100.01); a negative value = no discount can be given = the entry is declined.
    A value that is not a whole percent within half a point (0.4979, 0.4000514 - formula noise) is written as the whole percent and said so.  -> (cap or None for declined, note)"""
    x = num(v)
    if x is None:
        return 100.0, ""
    if x < 0:
        return None, ""
    if x == 0:
        return 100.0, ""
    pct = x * 100
    whole = float(round(pct))
    if abs(pct - whole) <= 0.0005:                          # a whole percent (up to float noise: 0.29 * 100 = 28.999999999999996)
        return rnd(pct), ""
    if DISCOUNT_ROUNDING == "sheet":
        if shows_whole is None:                             # the cell's format is not known: a multiple of half a point is a real limit (67.5), anything else is formula noise that the sheet shows as a whole percent
            shows_whole = abs(pct * 2 - round(pct * 2)) > 0.001
        if shows_whole:
            shown = float(math.floor(pct + 0.5))             # Excel rounds half up when it shows a whole percent
            CD1_ROUNDED.append((rnd(pct), shown))
            return shown, ""
        return rnd(pct), ""
    if DISCOUNT_ROUNDING == "down":
        down = float(pct // 1)
        CD1_ROUNDED.append((rnd(pct), down))
        return down, ""
    if abs(pct - whole) <= 0.51:
        return whole, f"discount limit {rnd(pct):g}% in the grid written as {whole:g}%"
    return rnd(pct), ""


def tw_row(ctx, g_r, sub, klass, makes, cluster, cc, fuels, rate, cap, pi_type="On Net", is_cpa="NA", tnc="", z=False, flags=(), group_prefix="2W_1+1_TP_", rto=""):
    return dict(product="TW", sub_product=sub, policy_type="New" if sub == "1+5" else "Renew", make_model_class=klass, new_vehicle_makes=ctx.makes(makes),
                new_rto_list=ctx.rto(group_prefix.rstrip("_"), cluster, "2W SAOD CLUSTER" if "SAOD" in group_prefix else "2W COMP/SATP CLUSTER", working=rto), _grp=group_prefix.rstrip("_"), vehicle_age_min=0, vehicle_age_max=AGE_OPEN_END, cc_min=cc[0], cc_max=cc[1], sc_min=SC_RANGE[0], sc_max=SC_RANGE[1],
                is_ncb="NA", is_cpa=is_cpa, is_zd="NA", add_tnc=tnc, _r=g_r, _z=z, _fuels=tuple(fuels), _flags=tuple(flags), pi_type=pi_type,
                pi_od_rate=rate if pi_type == "On OD" else 0.0, pi_tp_rate=rate if pi_type == "On TP" else 0.0, pi_net_rate=rate if pi_type == "On Net" else 0.0,
                tariff_max=rnd(cap + 0.01) if cap is not None else TARIFF_FULL, _cap=cap)


SPEC_TW11 = Spec("the two-wheeler 1+1 / SATP grid", "Latest 2W 1+1 & SATP",
                 fields=dict(cluster=["Agency/PB Clusters", "Agency/PB Cluster"], seg=["Agency/PB Seg", "Agency/PB Segment"], cd1=["CD1"], cd2=[PAYOUT_HEAD], make=["Make"], formula=["Formula Type"]),
                 required=("cluster", "seg", "cd1", "cd2"), forbid=("make", "formula"), many=("cd2",), show=dict(cd2="CD2 (or Max CD2 / Avg CD2)"))


def emit_tw11(book, ask, ctx, log=say):
    """'Latest 2W 1+1 & SATP': cluster | segment | CD1 | the payouts of 1+1 (comprehensive) and of SATP.  The sheet is found by its headers (it is the one WITHOUT a Make / Formula Type column - those are the 1+5
    sheet's).  The payout columns are told apart by the titles above them ('1+1' / 'SATP', 'With CPA' / 'Without CPA'), never by their order: a second payout column ('Max CD2' beside 'Avg CD2') is asked about,
    and a SATP payout that is no longer split by CPA is one row per entry with is_cpa NA (the schema: separate rows only where the grid splits them)."""
    hit = find_table(book, SPEC_TW11, ask, ctx.warnings, log, optional=True)
    if hit is None:
        table_missing(book, SPEC_TW11, ask, ctx, "no two-wheeler 1+1 / STP rows")
        return []
    sh, hr = hit.sheet, hit.hr
    col = {f: c[0] for f, c in hit.cols.items() if f != "cd2"}
    pays = []                                                       # (column, header, section 'comp' | 'satp', cpa 'with' | 'without' | None)
    for c in hit.cols["cd2"]:
        section = cpa = None
        for _, _, t in title_texts(sh, hr, c):
            what, val = classify_title(t)
            if what == "section" and section is None:
                section = val
            elif what == "cpa" and cpa is None:
                cpa = val
        pays.append((c, hit.label(c), section, cpa))
    if any(s is None for _, _, s, _ in pays):
        raise GridError(f"sheet '{sh.name}': expected the titles '1+1' and 'SATP' above the payout columns, to tell the comprehensive ones from the third-party ones; "
                        f"found no such title above {[lab + ' (' + col_name(c) + ')' for c, lab, s, _ in pays if s is None]}")
    comp = [(c, lab) for c, lab, s, _ in pays if s == "comp"]
    satp = {}
    for c, lab, s, q in pays:
        if s == "satp":
            satp.setdefault(q, []).append((c, lab))
    if not comp or set(satp) not in ({"with", "without"}, {None}):
        raise GridError(f"sheet '{sh.name}': expected a payout column under '1+1' and, under 'SATP', either one column or one 'With CPA' and one 'Without CPA' column; "
                        f"found {[(lab + ' (' + col_name(c) + ')', s, q) for c, lab, s, q in pays]}")
    comp_col = choose_payout(ask, f"sheet '{sh.name}' (1+1)", comp)[0]
    split = set(satp) == {"with", "without"}
    satp_cols = [(choose_payout(ask, f"sheet '{sh.name}' (SATP{' ' + q + ' CPA' if q else ''})", satp[q])[0], yes) for q, yes in ((("with", "YES"), ("without", "NO")) if split else ((None, "NA"),))]
    if not split:
        ctx.warnings.append(f"sheet '{sh.name}': the SATP payout is not split into 'With CPA' / 'Without CPA' (earlier grids had both): one row per entry is written, is_cpa NA")
    rows = []
    wc = WorkingCol(sh, col["cluster"], scan=hr)
    warn_unread(hit, set(col.values()) | {c for c, _, _, _ in pays} | ({wc.col} if wc.col is not None else set()), ctx.warnings)
    tvs_with_hero_honda = {norm(cell_at(x, col["cluster"])) for x in sh.rows[hr:] if label_key(norm(cell_at(x, col["seg"]))) == "mc180herohondatvs"}
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        cluster, seg = norm(cell_at(raw, col["cluster"])), norm(cell_at(raw, col["seg"]))
        if not cluster or not seg:
            continue
        wk = wc.name(raw, r)
        tgt = TW_11_SEGMENTS.get(label_key(seg))
        if tgt is None:
            ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=seg, reason=f"segment '{seg}' is not one this script knows"))
            continue
        klass, makes, cc = tgt
        if label_key(seg) == "mc180others" and cluster in tvs_with_hero_honda:         # a cluster whose Hero / Honda segment also holds TVS: 'Others' there is every make but Hero, Honda and TVS (TVS is not added back)
            makes = TW_11_OTHERS_EX_TVS
        cap, note = cd1_cap(cell_at(raw, col["cd1"]), sh.shows_whole(r, col["cd1"] + 1))
        comp_v = num(cell_at(raw, comp_col))
        rate = 0.0 if (cap is None or comp_v is None) else rnd(comp_v * 100)
        rows.append(tw_row(ctx, r, "1+1", klass, makes, cluster, cc, TW_FUELS, rate, cap if cap is not None else 100.0, z=comp_v is None, flags=[note] if note else [], rto=wk))
        for cpa_col, yes in satp_cols:
            x = num(cell_at(raw, cpa_col))
            rows.append(tw_row(ctx, r, "STP", klass, makes, cluster, cc, TW_FUELS, 0.0 if x is None else rnd(x * 100), None, pi_type="On TP", is_cpa=yes, z=x is None, rto=wk))
    wc.done()
    log(f"Two-wheeler 1+1 / SATP: sheet '{sh.name}'")
    return rows


SPEC_TW15 = Spec("the two-wheeler 1+5 grid", "2W 1+5",
                 fields=dict(cluster=["Agency/PB Clusters", "Agency/PB Cluster"], make=["Make"], seg=["Agency/PB Seg", "Agency/PB Segment"], cd1=["CD1"], cd2=[PAYOUT_HEAD], formula=["Formula Type"]),
                 required=("cluster", "make", "seg", "cd1", "cd2", "formula"), many=("cluster", "make", "seg", "cd1", "cd2", "formula"), show=dict(cd2="Max CD2 (or Avg CD2)"))


def emit_tw15(book, ask, ctx, log=say):
    """'2W 1+5' / '2W Grid 1+5' (new two-wheelers): one or two side-by-side blocks of  cluster | make | segment | CD1 | payout | formula type.  Every column of a block is found by its header inside the
    block (from its 'Agency/PB Clusters' header to the next one), never by counting cells from the cluster column: the November grid put an 'Avg CD2' column in front of 'Max CD2', which a position
    reader took for the payout (and the Max figure for the formula type)."""
    hit = find_table(book, SPEC_TW15, ask, ctx.warnings, log, optional=True)
    if hit is None:
        table_missing(book, SPEC_TW15, ask, ctx, "no 1+5 two-wheeler rows")
        return []
    sh, hr = hit.sheet, hit.hr
    all_starts = sorted(hit.cols["cluster"])
    rng = block_ranges(hit, all_starts)
    starts = list(all_starts)
    titled = [s for s in starts if any(label_key(sh.cell(rr, c + 1)) == "15" for rr in range(1, hr) for c in range(*rng[s]))]       # the block with '1+5' written above it
    if len(starts) > 1 and len(titled) == 1:
        ctx.warnings.append(f"sheet '{sh.name}' holds two blocks of the same columns; only the one titled '1+5' is read (the other block is not titled and gives other rates for the same clusters)")
        starts = titled
    blocks = {}
    for s in starts:
        got = {f: hit.within(f, *rng[s]) for f in ("make", "seg", "cd1", "cd2", "formula")}
        lacking = [SPEC_TW15.show[f] for f, v in got.items() if not v]
        if lacking:
            raise GridError(f"sheet '{sh.name}': the block that starts at column {col_name(s)} has no {', '.join(lacking)} column - expected cluster, make, segment, CD1, payout (CD2) and formula type side by side; "
                            f"the block says {[hit.label(c) for c in range(*rng[s]) if hit.label(c)]}")
        pay = choose_payout(ask, f"sheet '{sh.name}' (1+5, block at {col_name(s)})", [(c, hit.label(c)) for c in got["cd2"]])[0]
        used = {s, got["make"][0], got["seg"][0], got["cd1"][0], got["formula"][0]} | set(got["cd2"])
        blocks[s] = dict(make=got["make"][0], seg=got["seg"][0], cd1=got["cd1"][0], pay=pay, formula=got["formula"][0], used=used)
    rows, misp = [], 0
    wcs = {s: WorkingCol(sh, s, f"block {all_starts.index(s) + 1}" if len(all_starts) > 1 else "", starts=all_starts, span=max(0, max(blocks[s]["used"]) - s + 1), warnings=ctx.warnings, scan=hr) for s in starts}
    for s in starts:
        warn_unread(hit, blocks[s]["used"] | ({wcs[s].col} if wcs[s].col is not None else set()), ctx.warnings, lo=rng[s][0], hi=rng[s][1])
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        for s in starts:
            b = blocks[s]
            cluster, make, seg = norm(cell_at(raw, s)), norm(cell_at(raw, b["make"])), norm(cell_at(raw, b["seg"]))
            if not cluster or not make:
                continue
            wk = wcs[s].name(raw, r)
            targets = tw15_targets(make, seg)
            val = cell_at(raw, b["pay"])
            if not targets:
                if label_key(make) in ("heromotocorp", "hero") and (norm(val).upper() == "D" or num(val) == 0):
                    why = ("the grid pays HERO MOTOCORP 0 (declined) for every cluster and segment: no card is written for it; the make group of the 'Others' rows still contains HERO, so Hero is paid the rate "
                           "of 'Others' (as in your October upload) - the decline of the grid is not applied")
                else:
                    why = "make / segment is not one this script knows"
                ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=f"{make} / {seg}", reason=why))
                continue
            flags = []
            formula = label_key(cell_at(raw, b["formula"]))
            pi_type, z = ("On OD" if "addon" in formula else "On Net"), False             # Formula Type 'Slab OD+Add on' = a payout on the OD premium, 'Slab Net 5' = on the net premium
            if isinstance(val, str) and val.strip().upper() == "MISP":
                rate, pi_type = MISP_RATE, "On OD"
                misp += 1
            elif isinstance(val, str) and val.strip().upper() == "D":
                rate = 0.0
            elif num(val) is None:
                rate, z = 0.0, True
            else:
                rate = rnd(num(val) * 100)
            cap, note = cd1_cap(cell_at(raw, b["cd1"]), sh.shows_whole(r, b["cd1"] + 1))
            if note:
                flags.append(note)
            if cap is None:
                rate, cap = 0.0, 100.0
            for klass, makes, cc, fuels in targets:
                rows.append(tw_row(ctx, r, "1+5", klass, makes, cluster, cc, fuels, rate, cap, pi_type=pi_type, z=z, flags=flags, rto=wk))
    if misp:
        ctx.warnings.append(f"{misp} two-wheeler 1+5 grid cell(s) say MISP and carry no number: written as On OD {MISP_RATE:g}, the value of your October upload (setting MISP_RATE; a standing decision, not a grid figure)")
    for w in wcs.values():
        w.done()
    log(f"Two-wheeler 1+5: sheet '{sh.name}'")
    return rows


TW_SAOD_SEGMENTS = {                             # (segment key, make key) -> (class, make group, cc band, fuels)
    ("mc155", "otherthanhondaandbajaj"): ("BIKE", "BIKEALL_EX_HONDA_BAJAJ", (0, 155.01), TW_FUELS),
    ("mc155_", "otherthanhondaandbajaj"): ("BIKE", "BIKEALL_EX_HONDA_BAJAJ", (155.01, CC_OPEN_END), TW_FUELS),
    ("re", "otherthanhondaandbajaj"): ("BIKE", "BIKEALL_EX_HONDA_BAJAJ", (0, CC_OPEN_END), TW_FUELS),
    ("sc", "otherthanhondaandbajaj"): ("SCOOTER", "SCOOTERALL_EX_HONDA_SUZUKI_HERO_TVS_BAJAJ_YAMAHA_ROYAL,SUZUKI,HERO,TVS,YAMAHA", (0, CC_OPEN_END), ["Petrol"]),
    ("scev", "otherthanhondaandbajaj"): ("SCOOTER", "SCOOTERALL_EX_HONDA_SUZUKI_HERO_TVS_BAJAJ_YAMAHA_ROYAL,SUZUKI,HERO,TVS,YAMAHA,HONDA", (0, CC_OPEN_END), ["Electric"]),
    ("scev", "bajaj"): ("SCOOTER", "BAJAJ", (0, CC_OPEN_END), ["Electric"]),
    ("sc", "bajaj"): ("SCOOTER", "BAJAJ", (0, CC_OPEN_END), ["Petrol"]), ("sc", "honda"): ("SCOOTER", "HONDA", (0, CC_OPEN_END), ["Petrol"]),
}
for _mk in ("bajaj", "honda"):
    TW_SAOD_SEGMENTS[("mc155", _mk)] = ("BIKE", _mk.upper(), (0, 155.01), TW_FUELS)
    TW_SAOD_SEGMENTS[("mc155_", _mk)] = ("BIKE", _mk.upper(), (155.01, CC_OPEN_END), TW_FUELS)


SPEC_SAOD = Spec("the two-wheeler SAOD grid", "Latest 2W SAOD",
                 fields=dict(cluster=["Cluster", "Agency/PB Clusters"], segment=["Segment", "Agency/PB Seg"], make=["Make"],
                             nobreak=[re.compile(r"^(?=.*cd1)(?=.*nobreakin)")], break_=[re.compile(r"^(?=.*cd1)(?=.*breakin)(?!.*nobreakin)")], year1=[re.compile(r"^(?=.*1year)(?=.*cd2)")]),
                 required=("cluster", "segment", "make", "nobreak", "break_", "year1"), show=dict(nobreak="Max CD1 for NoBreakIn", break_="Max CD1 for BreakIn", year1="Max 1 year SAOD CD2"))
SAOD_UNREAD_OK = [re.compile(r"^mincd1$"), re.compile(r"^max\d+(st|nd|rd|th)?year")]           # Min CD1 and the 2nd / 3rd / 4th year payouts: left out on purpose (a SAOD policy is one year)


def emit_tw_saod(book, ask, ctx, log=say):
    """'Latest 2W SAOD' / '2W SAOD': cluster | segment | make | Min CD1 | Max CD1 for NoBreakIn | Max CD1 for BreakIn | Max CD2 for the 1st / 2nd / 3rd / 4th SAOD year.  The 1-year payout is written (a SAOD
    policy is a year).  Columns that are in the sheet and that nothing reads (the November sheet's new 'CD2' and 'Status ...' columns) are listed in the warnings."""
    hit = find_table(book, SPEC_SAOD, ask, ctx.warnings, log, optional=True)
    if hit is None:
        table_missing(book, SPEC_SAOD, ask, ctx, "no two-wheeler SAOD rows")
        return []
    sh, hr = hit.sheet, hit.hr
    col = {("break" if f == "break_" else f): c[0] for f, c in hit.cols.items()}
    rows = []
    wc = WorkingCol(sh, col["cluster"], scan=hr)
    warn_unread(hit, set(col.values()) | ({wc.col} if wc.col is not None else set()), ctx.warnings, ignore=SAOD_UNREAD_OK)
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        cluster, seg, make = norm(cell_at(raw, col["cluster"])), norm(cell_at(raw, col["segment"])), norm(cell_at(raw, col["make"]))
        if not cluster or not seg:
            continue
        wk = wc.name(raw, r)
        sk = tw15_seg(seg) if re.match(r"(?i)^mc", seg) else label_key(seg)
        tgt = TW_SAOD_SEGMENTS.get((sk, label_key(make)))
        if tgt is None:
            ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=f"{seg} / {make}", reason="segment / make is not one this script knows"))
            continue
        klass, makes, cc, fuels = tgt
        pay = num(cell_at(raw, col["year1"]))
        for tnc, key in (("No BreakIn", "nobreak"), ("BreakIn", "break")):
            cap, note = cd1_cap(cell_at(raw, col[key]), sh.shows_whole(r, col[key] + 1))
            rate = 0.0 if (cap is None or pay is None) else rnd(pay * 100)
            rows.append(tw_row(ctx, r, "SAOD", klass, makes, re.sub(r"[,\s]+", "_", cluster), cc, fuels, rate, cap if cap is not None else 100.0, pi_type="On OD", tnc=tnc,
                               z=pay is None, group_prefix="2W_SAOD_", flags=[note] if note else [], rto=wk))
    wc.done()
    log(f"Two-wheeler SAOD: sheet '{sh.name}'")
    return rows


# ---- Private Car
PC_VOLUME_SLAB = "1-3 Lac"                       # the premium-volume slab of the user's deal (their October upload used it); the other slabs of the grid are not written
PC_FUELS = ["Petrol", "Diesel", "CNG", "LPG"]    # the comp / SAOD table names no fuel: all the non-electric fuels (the October upload had Petrol only)
PC_TP_FUELS = {"petrol": ["Petrol", "Electric"], "cng": ["CNG", "LPG"], "diesel": ["Diesel"]}      # SATP table: Electric takes the Petrol rate, LPG the CNG rate (as in the October upload)


def volume_key(s):
    return re.sub(r"[^a-z0-9]+", "", norm(s).lower())


def pc_row(ctx, r, sub, cluster, prefix, rate, ncb="NA", cc=(0, CC_OPEN_END), age=(0, AGE_OPEN_END), fuels=PC_FUELS, pi_type="On OD", z=False, flags=(), rto=""):
    return dict(product="Private Car", sub_product=sub, policy_type="Renew", make_model_class="Car", new_vehicle_makes=ctx.makes("Private_Car_ALL"),
                new_rto_list=ctx.rto_pc(prefix, cluster, rto), _grp=prefix, vehicle_age_min=age[0], vehicle_age_max=age[1], cc_min=cc[0], cc_max=cc[1], sc_min=SC_RANGE[0], sc_max=SC_RANGE[1], is_ncb=ncb,
                is_cpa="NA", is_zd="NA", add_tnc="", _r=r, _z=z, _fuels=tuple(fuels), _flags=tuple(flags), pi_type=pi_type,
                pi_od_rate=rate if pi_type == "On OD" else 0.0, pi_tp_rate=rate if pi_type == "On TP" else 0.0, pi_net_rate=rate if pi_type == "On Net" else 0.0,
                tariff_max=TARIFF_FULL, _cap=None)


SPEC_PCC = Spec("the private-car comprehensive / SAOD grid", "PVT Car Comp+SAOD",
                fields=dict(cluster=["Cluster"], volume=["Volume Slab", "Volume"], ncb=[re.compile(r"^(non)?ncb$")]),
                required=("cluster", "volume", "ncb"), many=("cluster", "ncb"), show=dict(ncb="Non NCB / NCB"))


def emit_pc_comp(book, ask, ctx, log=say):
    """'PVT Car Comp+SAOD' / 'New 4W Comp_SAOD': cluster | volume slab | SAOD Non NCB | SAOD NCB | Comp Non NCB | Comp NCB (payouts in percent already).  The cluster column is the one of the SAME block as the
    'Volume Slab' column - the nearest 'Cluster' header to its left - not the first 'Cluster' on the row: the November sheet has a new-business block on the left whose clusters run in another order.
    The slab of PC_VOLUME_SLAB is written; if the sheet no longer has it (November: one slab, '> 5 Lacs') the person is asked which slab to use - never guessed."""
    hit = find_table(book, SPEC_PCC, ask, ctx.warnings, log, optional=True)
    if hit is None:
        table_missing(book, SPEC_PCC, ask, ctx, "no private-car 1+1 / SAOD rows")
        return []
    sh, hr, head = hit.sheet, hit.hr, hit.head
    c_vol = hit.first("volume")
    left = [c for c in hit.cols["cluster"] if c <= c_vol]
    if not left:
        raise GridError(f"sheet '{sh.name}': expected a 'Cluster' header at or to the left of the 'Volume Slab' header (the same block); the header row {hr} says {[hit.label(c) for c in range(len(head)) if hit.label(c)]}")
    c_cl = max(left)
    right = [c for c in hit.cols["cluster"] if c > c_vol]
    sections = {}                                  # 'saod' / 'comp' -> columns of Non NCB, NCB
    for c in hit.within("ncb", c_vol + 1, min(right) if right else hit.ncols()):
        titles = title_texts(sh, hr, c)
        t = label_key(titles[0][2]) if titles else ""
        if t in ("saod", "comp"):
            sections.setdefault(t, {}).setdefault(label_key(head[c]), c)
    if not all(set(sections.get(k, {})) == {"nonncb", "ncb"} for k in ("saod", "comp")):
        raise GridError(f"sheet '{sh.name}': could not find the SAOD / Comp columns Non NCB and NCB under the headings (found {sections})")
    slabs = []
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        cl, vol = norm(cell_at(raw, c_cl)), norm(cell_at(raw, c_vol))
        if cl and vol and vol not in slabs:
            slabs.append(vol)
    if not slabs:
        raise GridError(f"sheet '{sh.name}': no row has both a cluster (column {col_name(c_cl)}) and a volume slab (column {col_name(c_vol)}) - the private-car comprehensive / SAOD grid is empty or laid out differently")
    use = next((s for s in slabs if volume_key(s) == volume_key(PC_VOLUME_SLAB)), None)
    if use is None:
        use = ask.choose(f"the private-car volume slab '{PC_VOLUME_SLAB}' (the slab your uploads used) is not in the sheet '{sh.name}': it has {slabs}",
                         [(f"the rows of the volume slab '{s}'", s) for s in slabs] + [("none - leave the private-car 1+1 / SAOD rows out of this file", None)],
                         question=(f"The private-car comprehensive / SAOD grid no longer has the volume slab '{PC_VOLUME_SLAB}' that your uploads used. The sheet has: {', '.join(repr(s) for s in slabs)}. "
                                   f"Which slab should the private-car 1+1 / SAOD rows be written from?"), short="private-car volume slab", memo="private-car volume slab")
        if use is None:
            ctx.warnings.append(f"sheet '{sh.name}': the volume slab '{PC_VOLUME_SLAB}' is not in the sheet (slabs: {slabs}) and you chose to leave the private-car 1+1 / SAOD rows out")
            return []
        ctx.warnings.append(f"sheet '{sh.name}': the volume slab '{PC_VOLUME_SLAB}' is not in the sheet; the rows of '{use}' are written instead (your choice)")
    want = volume_key(use)
    rows = []
    wc = WorkingCol(sh, c_cl, starts=hit.cols["cluster"], scan=hr)
    used = {c_cl, c_vol} | {c for sec in sections.values() for c in sec.values()} | ({wc.col} if wc.col is not None else set())
    warn_unread(hit, used, ctx.warnings, lo=c_cl, hi=max(used) + 1)
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        cluster, vol = norm(cell_at(raw, c_cl)), norm(cell_at(raw, c_vol))
        if not cluster or not vol:
            continue
        if volume_key(vol) != want:
            continue
        wk = wc.name(raw, r)
        for sub, key, prefix in (("1+1", "comp", "PC_1+1_OCT26"), ("SAOD", "saod", "PC_SAOD_OCT26")):
            for ncb, name in (("YES", "ncb"), ("NO", "nonncb")):
                x = num(cell_at(raw, sections[key][name]))
                rows.append(pc_row(ctx, r, sub, cluster, prefix, 0.0 if x is None else rnd(x), ncb=ncb, z=x is None, rto=wk))
    wc.done()
    log(f"Private car comprehensive / SAOD: sheet '{sh.name}', volume slab '{use}'")
    return rows


PC_TP_SEGMENT = re.compile(r"^(petrol|diesel|cng)\s*(<\s*1000|<\s*1500|>\s*1000|>\s*1500|1000\s*-\s*1500)?$", re.I)


def pc_cc(fuel, cmp_text):
    t = re.sub(r"\s+", "", cmp_text or "")
    return {"": (0, CC_OPEN_END), "<1000": (0, 1000.01), "<1500": (0, 1500.01), ">1000": (1000.01, CC_OPEN_END), ">1500": (1500.01, CC_OPEN_END), "1000-1500": (1000.01, 1500.01)}[t]


TP_DECLINE_NOTE = re.compile(r"^(?P<cc>more than (?P<over>\d+) ?cc|less than (?P<under>\d+) ?cc|(?P<lo>\d+) ?- ?(?P<hi>\d+) ?cc) and (?P<age>(?P<aplus>\d+) ?\+ ?age|< ?(?P<aunder>\d+) ?age) (?:is|are) declined$", re.I)


def tp_decline_slice(note):
    """'More than 1500CC and 10+age is declined' / '1000-1500CC and <10 age is declined' -> ((cc from, cc to), (age from, age to)) of the declined slice in the table's own numbers; None for any other note
    ('Few RTOs for 10+ age are declined' names no RTOs: not applied)"""
    m = TP_DECLINE_NOTE.match(norm(note))
    if not m:
        return None
    if m.group("over"):
        cc = (round(int(m.group("over")) + 0.01, 2), CC_OPEN_END)
    elif m.group("under"):
        cc = (0, round(int(m.group("under")) + 0.01, 2))
    else:
        cc = (round(int(m.group("lo")) + 0.01, 2), round(int(m.group("hi")) + 0.01, 2))
    age = (round(int(m.group("aplus")) + 0.01, 2), AGE_OPEN_END) if m.group("aplus") else (0, round(int(m.group("aunder")) + 0.01, 2))
    return cc, age


def split_declined(cc, age, dec):
    """the slice (cc band x age band) of one table row minus the declined slice of its note -> [(cc band, age band, pays?)] covering the row exactly once; None when the declined slice does not touch the row
    (the note is then left flagged).  The row's own payout goes to the parts that pay; the declined part gets no card."""
    (c0, c1), (a0, a1), ((d0, d1), (e0, e1)) = cc, age, dec
    if not (max(c0, d0) < min(c1, d1) and max(a0, e0) < min(a1, e1)):
        return None
    cuts = lambda lo, hi, x0, x1: sorted({lo, hi} | {x for x in (x0, x1) if lo < x < hi})
    cs, as_ = cuts(c0, c1, d0, d1), cuts(a0, a1, e0, e1)
    parts = []
    for i in range(len(cs) - 1):
        for j in range(len(as_) - 1):
            inside = d0 <= cs[i] and cs[i + 1] <= d1 and e0 <= as_[j] and as_[j + 1] <= e1
            if parts and parts[-1][0] == (cs[i], cs[i + 1]) and parts[-1][2] == (not inside) and parts[-1][1][1] == as_[j]:
                parts[-1] = (parts[-1][0], (parts[-1][1][0], as_[j + 1]), parts[-1][2])          # the same cc band and the same payout side by side in age: one row
            else:
                parts.append(((cs[i], cs[i + 1]), (as_[j], as_[j + 1]), not inside))
    return [(c, a, pays) for c, a, pays in parts]


SPEC_PCT = Spec("the private-car third-party grid", " PVT Car TP",
                fields=dict(cluster=["Cluster"], segment=["Segment"], age=["Age"], cd2=[PAYOUT_HEAD], note=["Notes", "Note"]),
                required=("cluster", "segment", "age", "cd2"), many=("cluster", "segment", "age", "cd2", "note"), show=dict(cd2="Max CD2 (or Avg CD2)"))


def emit_pc_tp(book, ask, ctx, log=say):
    """' PVT Car TP' / '4W TP Consolidated': cluster | segment (fuel and cc) | age (<10, >10, All) | payout (Max CD2, and from November also Avg CD2) | note.  Every column is found by its header inside its block
    (from a 'Cluster' header to the next one), never by counting cells from 'Cluster': the November sheet put 'Avg CD2' where the note used to be, which a position reader took for a grid note on every row.
    When the sheet holds the table twice side by side, the block of the first header row that names it is read and the copies are compared."""
    hit = find_table(book, SPEC_PCT, ask, ctx.warnings, log, optional=True)
    if hit is None:
        table_missing(book, SPEC_PCT, ask, ctx, "no private-car STP rows")
        return []
    sh, hr = hit.sheet, hit.hr
    starts = sorted(hit.cols["cluster"])
    rng = block_ranges(hit, starts)
    blocks = []
    for s in starts:
        got = {f: hit.within(f, *rng[s]) for f in ("segment", "age", "cd2", "note")}
        lacking = [SPEC_PCT.show[f] for f in ("segment", "age", "cd2") if not got[f]]
        if lacking:
            raise GridError(f"sheet '{sh.name}': the block that starts at column {col_name(s)} has no {', '.join(lacking)} column - expected Cluster, Segment, Age and a payout (Max CD2 / Avg CD2) side by side; "
                            f"the block says {[hit.label(c) for c in range(*rng[s]) if hit.label(c)]}")
        pay = choose_payout(ask, f"sheet '{sh.name}' (private-car third party, block at {col_name(s)})", [(c, hit.label(c)) for c in got["cd2"]])[0]
        used = {s, got["segment"][0], got["age"][0]} | set(got["cd2"]) | set(got["note"][:1])
        blocks.append(dict(start=s, seg=got["segment"][0], age=got["age"][0], pay=pay, note=(got["note"] or [None])[0], used=used))
    scanned = []
    for b in blocks:
        blk = {}
        for r in range(hr + 1, len(sh.rows) + 1):
            raw = sh.rows[r - 1]
            if norm(cell_at(raw, b["start"])) and norm(cell_at(raw, b["seg"])):
                blk[(norm(cell_at(raw, b["start"])).upper(), label_key(cell_at(raw, b["seg"])), label_key(cell_at(raw, b["age"])))] = (r, num(cell_at(raw, b["pay"])), norm(cell_at(raw, b["note"])))
        scanned.append(blk)
    if len(scanned) > 1 and {k: v[1] for k, v in scanned[0].items()} != {k: v[1] for k, v in scanned[1].items()}:
        ctx.warnings.append(f"sheet '{sh.name}': the two copies of the table differ in {sum(1 for k in scanned[0] if scanned[0][k][1] != scanned[1].get(k, (0, None))[1])} entries: the left copy is used")
    rows = []
    b0 = blocks[0]
    s = b0["start"]
    wc = WorkingCol(sh, s, starts=starts, span=max(0, max(b0["used"]) - s + 1), warnings=ctx.warnings, scan=hr)
    warn_unread(hit, b0["used"] | ({wc.col} if wc.col is not None else set()), ctx.warnings, lo=rng[s][0], hi=rng[s][1])
    for r in range(hr + 1, len(sh.rows) + 1):
        raw = sh.rows[r - 1]
        cluster, seg, age, x, note = norm(cell_at(raw, s)), norm(cell_at(raw, b0["seg"])), norm(cell_at(raw, b0["age"])), cell_at(raw, b0["pay"]), norm(cell_at(raw, b0["note"]))
        if not cluster or not seg:
            continue
        wk = wc.name(raw, r)
        m = PC_TP_SEGMENT.match(norm(seg))
        if not m:
            ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=seg, reason=f"segment '{seg}' is not a fuel + cc text this script knows"))
            continue
        fuel, ccx = m.group(1).lower(), m.group(2)
        ages = {"all": (0, AGE_OPEN_END), "<10": (0, 10.01), ">10": (10.01, AGE_OPEN_END)}.get(norm(age).replace(" ", "").lower())
        if ages is None:
            ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=f"{seg} / {age}", reason=f"age '{age}' is not <10, >10 or All"))
            continue
        flags = [f"grid note: {note}"] if note else []
        v = num(x)
        cc = pc_cc(fuel, ccx)
        parts = [(cc, ages, True)]
        dec = tp_decline_slice(note) if note else None
        split = split_declined(cc, ages, dec) if dec else None
        if split:                                                                  # 'More than 1500CC and 10+age is declined': that slice of THIS row gets no card (your October upload has none either), the rest keeps the row's payout (no flag)
            parts, flags = split, []
        for cc_, ages_, pays in parts:
            if not pays:
                ctx.not_written.append(dict(sheet=sh.name, grid_row=r, cluster=cluster, segment=f"{seg} / {age}", reason=f"declined by the grid's own note '{note}' (cc {cc_[0]:g}-{cc_[1]:g}, age {ages_[0]:g}-{ages_[1]:g}): no card is written for that slice"))
                continue
            rows.append(pc_row(ctx, r, "STP", cluster, "PC_TP_OCT26", 0.0 if v is None else rnd(v * 100), cc=cc_, age=ages_, fuels=PC_TP_FUELS[fuel], pi_type="On TP", z=v is None, flags=flags, rto=wk))
    wc.done()
    log(f"Private car third-party: sheet '{sh.name}'")
    return rows


# %% ---------------------------------------------------------------- BUILD
LAST_GAPS = {}            # RTO groups the master lacks for this run: group name -> the RTO codes the grid's mapping sheet puts in it (a sheet of the review workbook, see collect_gaps)
LAST_NOT_WRITTEN = []     # grid rows that gave no upload row: the sheet 'Not written' of the review workbook


def expand_fuels(df):
    """one row per fuel of the row's fuel list"""
    parts = []
    for f in FUELS_ALL:
        d = df[df["_fuels"].map(lambda fs, f=f: f in fs)].copy()
        d["fuel_type"] = f
        parts.append(d)
    return pd.concat(parts, ignore_index=True)


def _short(text):
    t = re.sub(r"\s+", " ", str(text)).strip()
    return t if len(t) <= MAX_NOTE_CHARS else t[:MAX_NOTE_CHARS - 3] + "..."


D_IGNORED = []            # the entries where a D (decline) of one grid row was ignored because another grid row gives a rate for the same entry (the review sheet 'D ignored'; cleared by build)


def drop_d_placeholders(df, key_cols, warnings):
    """A D in a payout cell (or a CD1 text that declines an age) of a row that covers only some ages / makes / discount bands means 'not on this row', not 'declined': Delhi NCR taxi - the row 'Hyundai' pays
    15% for age 0 and says D for age 1+, the row above says 'Hyundai age 1+' = 25%.  The user (18 Nov 2026, 'Yes, fix it'): a D must never beat a real rate of another grid row for the same card.
    So a zero card that comes from a D is dropped when another row gives a rate for the SAME ENTRY (the key without the discount band) at the same discount cap or a higher one - at the same cap it would
    beat the rate (the lowest-rate rule), below it it would turn the discounts under the rate's band into 'no payout' - or when its own band (a CD1 text '30.1-90%') starts below the real cap, i.e. overlaps the
    rate.  A D that lies ABOVE every real band stays (an explicit decline above the cap), an entry whose rows are all D keeps its zero card.  The ignored entries are listed in the review sheet 'D ignored'."""
    global D_IGNORED
    D_IGNORED = []
    if "_d" not in df.columns or not df["_d"].fillna(False).astype(bool).any():
        return df
    d = df["_d"].fillna(False).astype(bool)
    z = df["_z"].fillna(False).astype(bool)
    entry = df[[c for c in key_cols if c != "tariff_max"]].astype(str).agg("|".join, axis=1)
    real = ~d & ~z
    top = df.loc[real].groupby(entry[real])["tariff_max"].max()                      # the highest discount cap of a real card of each entry
    cap_real = entry.map(top).astype(float)
    start = df["_tmin"].astype(float) if "_tmin" in df.columns else pd.Series(np.nan, index=df.index)          # a D row of a CD1 band text ('30.1-90%') starts where its text says: it overlaps a real band when it starts below the real cap
    drop = d & cap_real.notna() & ((df["tariff_max"].astype(float) <= cap_real) | (start.notna() & (start < cap_real)))
    if not drop.any():
        return df
    value = payin_value(df)
    rows_of = df.loc[real].groupby(entry[real])["_r"].agg(lambda x: sorted({int(i) for i in x}))
    rates_of = df.loc[real].assign(_v=value[real], _e=entry[real]).groupby("_e").apply(lambda g: ", ".join(f"{fmt(v)} up to {fmt(c)}" for v, c in sorted(set(zip(g["_v"], g["tariff_max"].astype(float))), key=lambda t: t[1])), include_groups=False)
    out = {}
    for i in df.index[drop]:
        e = entry[i]
        row = df.loc[i]
        k = (row["product"], row["sub_product"], row["new_rto_list"], row["new_vehicle_makes"], row["make_model_class"], e)
        o = out.setdefault(k, dict(product=row["product"], sub_product=row["sub_product"], rto=row["new_rto_list"], makes=row["new_vehicle_makes"], make_model_class=row["make_model_class"], age=f"{fmt(row['vehicle_age_min'])}-{fmt(row['vehicle_age_max'])}",
                                   cc=f"{fmt(row['cc_min'])}-{fmt(row['cc_max'])}", fuels=set(), d_caps=set(), d_rows=set(), real_rows=rows_of.get(e, []), real_rates=rates_of.get(e, ""), flags=set()))
        o["fuels"].add(row["fuel_type"])
        o["d_caps"].add(fmt(row["tariff_max"]))
        o["d_rows"].add(int(row["_r"]))
        o["flags"].update(row["_flags"])
    for o in out.values():
        D_IGNORED.append(dict(product=o["product"], sub_product=o["sub_product"], rto=o["rto"], makes=o["makes"], make_model_class=o["make_model_class"], fuels=",".join(sorted(o["fuels"])), age=o["age"], cc=o["cc"],
                              d_up_to_discount=", ".join(sorted(o["d_caps"], key=float)), d_grid_rows=", ".join(str(r) for r in sorted(o["d_rows"])), rate_written=o["real_rates"],
                              rate_grid_rows=", ".join(str(r) for r in o["real_rows"]), flags_of_the_ignored_rows=" | ".join(sorted(o["flags"]))))
    warnings.append(f"{len(D_IGNORED)} entries: a D (decline) of one grid row was ignored because another grid row gives a rate for the same entry at the same or a higher discount cap (your answer of "
                    f"18 Nov 2026: a D never beats a real rate); listed in the review workbook next to the upload ('D ignored')")
    return df[~drop]


def settle_duplicates(df, key_cols, warnings):
    """Rows with the same key (the discount band included) are one entry given more than once.
    * A zero row that only stands for a BLANK cell never beats a real rate for the same entry (it is dropped when one exists).
    * The same payin twice is one row; if two grid rows give one entry different payins the LOWEST is written (the user's rule, 27 Oct 2026: "if the rate grid contains two different rates for the
      same entry, always select the lower rate").  Every such entry is listed in LAST_LOWEST (the shared write_csv writes it into the review workbook)."""
    global LAST_LOWEST
    LAST_LOWEST = []
    keys = df[key_cols].astype(str).agg("|".join, axis=1)
    z = df["_z"].fillna(False).astype(bool)
    real = set(keys[~z])
    extra = {}                                                      # key -> the flags of the rows that are dropped: the row that stays carries them (a flag is never lost with a duplicate; 18 Nov 2026)
    gone = z & keys.isin(real)
    for i in df.index[gone]:
        extra.setdefault(keys[i], []).extend(df.at[i, "_flags"])
    df = df[~gone].copy()
    keys = df[key_cols].astype(str).agg("|".join, axis=1)
    value = payin_value(df)
    dup = keys.duplicated(keep=False)
    drop = []
    for key, idx in (df[dup].groupby(keys[dup], sort=False).groups.items() if dup.any() else []):
        idx = list(idx)
        best = min(idx, key=lambda i: (value[i], idx.index(i)))
        drop.extend(i for i in idx if i != best)
        for i in idx:
            if i != best:
                extra.setdefault(key, []).extend(df.at[i, "_flags"])
        found = list(dict.fromkeys(fmt(value[i]) for i in idx))
        if len(found) > 1:
            row = df.loc[best]
            LAST_LOWEST.append(dict(product=row["product"], sub_product=row["sub_product"], rto=row["new_rto_list"], makes=row["new_vehicle_makes"], fuel=row["fuel_type"],
                                    make_model_class=row["make_model_class"], age=f"{fmt(row['vehicle_age_min'])}-{fmt(row['vehicle_age_max'])}", cc=f"{fmt(row['cc_min'])}-{fmt(row['cc_max'])}",
                                    tariff_max=fmt(row["tariff_max"]), grid_rows=", ".join(str(r) for r in sorted({df.at[i, "_r"] for i in idx})), rates_found=" / ".join(found), written=fmt(value[best])))
    if LAST_LOWEST:
        warnings.append(f"{len(LAST_LOWEST)} combinations got two different payins from different grid rows: the LOWEST is written (listed in the review workbook next to the upload)")
    out = df.drop(index=drop)
    if extra:                                                       # the row that stays carries the flags of the rows that were dropped for it
        positions = dict(zip(out[key_cols].astype(str).agg("|".join, axis=1), out.index))
        for k, fl in extra.items():
            i = positions.get(k)
            if i is None:
                continue
            mine = tuple(out.at[i, "_flags"])
            merged = tuple(dict.fromkeys(list(mine) + [f for f in fl if f not in mine]))
            if merged != mine:
                out.at[i, "_flags"] = merged
    return out


def chain_tariffs(df):
    """A higher discount slab of one entry starts where the lower one ends (0-50.01, then 50.01-85.01); SATP rows are always 0-100.01 (the grid names no discount for them)."""
    ck = [c for c in KEY_COLUMNS if c not in ("tariff_min", "tariff_max")]
    comp = df["sub_product"] != "STP"
    slabs = df.loc[comp, ck + ["tariff_max"]].drop_duplicates().sort_values(ck + ["tariff_max"], kind="stable")
    slabs["tariff_min"] = slabs.groupby(ck, sort=False)["tariff_max"].shift(1).fillna(0.0)
    stp = df.loc[~comp, ck + ["tariff_max"]].drop_duplicates().assign(tariff_min=0.0)
    out = df.merge(pd.concat([slabs, stp]), on=ck + ["tariff_max"], how="left")
    if "_tmin" in out.columns:
        out["tariff_min"] = out["_tmin"].where(out["_tmin"].notna(), out["tariff_min"])
    return out


NOTE_SECTIONS = [("schoolbus", "REMARKS FOR SCHOOL BUS"), ("staffbus", "REMARKS FOR STAFF BUS"), ("hcv", "REMARKS FOR HCV"), ("other", "remarks sheet (outside the headings)")]


def notes_table(sections, cv_notes, hcv_applied, seat_notes):
    """Every sentence of the remarks sheet and the notes of the CV sheet with what became of it - applied, not applied (listed), no action (a standing decision, with its reason) or flagged on the rows it
    may affect.  Nothing a sheet says about a rule is invisible any more (the user, 18 Nov 2026)."""
    out = []
    for key, label in NOTE_SECTIONS + [("_cv", "notes on the CV sheet")]:
        for t in (cv_notes if key == "_cv" else sections.get(key, [])):
            t = re.sub(r"\s+", " ", str(t)).strip()
            k = note_key(t)
            standing = next((s for s in STANDING_NOTES if note_key(s) in k), None)
            if key == "hcv" and hcv_applied and (HCV_NOTE_BLOCK.search(t) or HCV_NOTE_CD1.search(t)):
                status, why = "applied", "Volvo and Scania are blocked (the make group without Volvo); for the Non-Dumper/Tipper body Mahindra and Bharat Benz are written at the CD1 of the note"
            elif any(norm(t) == norm(x) for x in seat_notes):
                status, why = "not applied - listed in 'Not written'", "it gives its own rates for buses with up to 10 seats; the staff bus table is written for more than 10 seats only"
            elif standing:
                status, why = "no action (standing decision)", STANDING_WHY.get(standing, "")
            elif len(t) < NOTE_MIN_CHARS:
                status, why = "a label, not a rule", ""
            elif key == "schoolbus":
                status, why = "no rows affected", "School Bus rows are not written (see 'Not written')"
            else:
                status, why = "FLAGGED on the rows it may affect", "see 'Manual review notes'"
            out.append(dict(where=label, text=t[:300], status=status, why=why))
    return out


def build(grid_path, from_date, to_date, log=say, ask=None, book=None, password=None):
    """Read the workbook and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    from_date / to_date are datetime.date values; they are written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever the grid itself is missing something. The notes are never asked about: what they say is applied, and
    what cannot be applied is flagged in the remarks column (see ManualReview)."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    working_reset()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    if book is None:
        log(f"Reading {os.path.basename(str(grid_path))}...")
        grid_path, book = open_file(ask, grid_path, "the Go Digit workbook", make_opener(ask, password), EXCEL_EXTS)
    hidden = [s.name for s in book.sheets if not s.visible]
    log(f"Sheets: {len(book.sheets) - len(hidden)} visible, {len(hidden)} hidden (hidden sheets are older versions or helpers and are not read)")
    warnings = []
    masters = load_masters(ask)
    review = ManualReview()
    CD1_ROUNDED.clear()
    CD1_AUTO.clear()
    WORDINGS.clear()
    for reg in (READ_COLUMNS, LEFT_OUT, WARNED_COLUMNS, SHEET_ROLES):
        reg.clear()
    UNREAD_PARTS.clear()
    NOTES_READ.clear()
    D_IGNORED.clear()
    rows = []

    cv, cv_notes, _ = read_cv_grid(book, ask, log, warnings)
    sections = read_cv_remarks(book, ask)
    check_age_table(sections.get("_age_table"), warnings, log)
    hcv = read_hcv_rule(sections.get("hcv", []))
    if not hcv["cap"]:
        warnings.append("the CV remarks sheet has no note 'Only for Non-Dumper/Tipper: ... to operate at N% CD1': the Truck rows are written for every make at the grid's own discount band")
    if not hcv["blocked"]:
        warnings.append("the CV remarks sheet has no note 'Volvo, Scania to block': the heavy-vehicle rows use GCV_4W_ALL (Volvo / Scania included)")
    mapping = read_all_mappings(book, ask)
    ctx = Ctx(warnings, dates, masters, mapping, hcv)
    ctx.pcv2_clusters = {g.cluster for g in cv if g.seg == "CV PCV2W(All) 3W Non-electric" and wnorm(g.body_seg) in PCV2_SEGMENTS}
    for g in cv:
        rows.extend(emit_cv(g, ctx))
    if CD1_AUTO:
        log(f"{len(CD1_AUTO)} CD1 wordings (on {sum(n for _, n in CD1_AUTO.values())} grid rows) were read by the age-group grammar: the review workbook lists them (sheet 'CD1 readings')")
    rows.extend(emit_tw11(book, ask, ctx, log))
    rows.extend(emit_tw15(book, ask, ctx, log))
    rows.extend(emit_tw_saod(book, ask, ctx, log))
    rows.extend(emit_pc_comp(book, ask, ctx, log))
    rows.extend(emit_pc_tp(book, ask, ctx, log))
    not_applied = [k for k, w in WORDINGS.items() if w["unread"]]
    if not_applied:
        warnings.append(f"{len(not_applied)} of the {len(WORDINGS)} wordings of the CV grid's segment / body / make columns are not applied: they are flagged on their rows and listed first in the review sheet "
                        f"'Wording readings' ({'; '.join(' / '.join(x for x in k[1:] if x and x != '—') for k in not_applied[:4])}{' ...' if len(not_applied) > 4 else ''})")
    audit_unread(book, warnings)                                          # sheets / columns no reader used: named, never silent
    note = working_note()
    if note:
        warnings.append(note)
        log(note)
    n_err = ctx.counts.get("error cell", 0)
    if n_err:
        warnings.append(f"{n_err} payout cells of the CV grid show an Excel error (#VALUE!): treated as 'declined' (a zero card, discount range 0-{TARIFF_FULL:g}), as in the October upload - ask Go Digit for the corrected cells")
    if not rows:
        raise GridError("nothing to write - no grid row could be converted")

    df = pd.DataFrame(rows)
    df = expand_fuels(df)
    df["_flags"] = df["_flags"].map(tuple)
    keys = [c for c in KEY_COLUMNS if c != "tariff_min"]
    df = drop_d_placeholders(df, keys, warnings)                           # a D never beats a real rate of another grid row
    df = settle_duplicates(df, keys, warnings)
    df = chain_tariffs(df.reset_index(drop=True)).reset_index(drop=True)
    df = fill_columns(df, dates)

    flag_texts = {}
    for i, fl in enumerate(df["_flags"]):
        for text in fl:
            flag_texts.setdefault(text, []).append(i)
    for text, idx in sorted(flag_texts.items(), key=lambda kv: -len(kv[1])):
        mask = np.zeros(len(df), dtype=bool)
        mask[idx] = True
        review.flag(_short(text), mask)
    in_group = lambda token: (df["new_rto_list"].str.contains(f"_{token}_", regex=False) | df["_grp"].astype(str).map(lambda g: f"_{token}_" in f"_{g}_")).values      # also the rows named by the working list
    hcv_applied = bool(hcv["cap"] and hcv["blocked"])
    seat_notes = []
    for t in sections.get("staffbus", []):                                    # 'For Seating Capacity up to 10-seater, CD1_85% & CD2_20% in all open cohorts': rates for a seat range the staff bus table does not write
        m = SEAT_LIMIT_NOTE.search(t)
        if m and STAFF_BUS_SC[0] > int(m.group(1)):
            seat_notes.append(t)
            ctx.not_written.append(dict(sheet="CV remarks", grid_row="", cluster="all clusters", segment=f"Staff Bus up to {m.group(1)} seats",
                                        reason=f"the note '{norm(t)}' gives its own rates for buses with up to {m.group(1)} seats; the staff bus table of this grid is written for more than {m.group(1)} seats only, so those cards are not written"))
    scoped = [("staffbus", in_group("STAFF-BUS"), STANDING_NOTES + seat_notes), ("hcv", in_group("HCV"), STANDING_NOTES + (["Volvo, Scania to block"] if hcv_applied else []))]
    for section, mask, standing in scoped:                                   # a rule printed in one section of 'CV Remarks' concerns that product's rows
        for t in unknown_notes(sections.get(section, []), standing):
            if hcv_applied and section == "hcv":
                continue                                                      # the HCV note was read and applied
            review.flag(_short(t), mask)
    for t in unknown_notes(sections.get("other", []) + cv_notes, STANDING_NOTES):
        review.flag(_short(t), None)
    NOTES_READ.extend(notes_table(sections, cv_notes, hcv_applied, seat_notes))
    global LAST_GAPS, LAST_NOT_WRITTEN
    LAST_GAPS = dict(ctx.gaps)
    LAST_NOT_WRITTEN = list(ctx.not_written)
    df = df.drop(columns=[c for c in df.columns if c.startswith("_")])
    if CD1_ROUNDED:
        how = {"sheet": "written as the sheet shows them (a cell formatted as a whole percent shows the nearest whole percent)", "down": "rounded DOWN to the whole percent"}.get(DISCOUNT_ROUNDING, "rounded")
        warnings.append(f"{len(CD1_ROUNDED)} discount limit(s) that are not a whole percent in the file were {how} (your rule of 31 Oct 2026, DISCOUNT_ROUNDING = '{DISCOUNT_ROUNDING}'), "
                        f"for example {CD1_ROUNDED[0][0]:g}% -> {CD1_ROUNDED[0][1]:g}%")
    check_names(df, review, masters, log)                                # every RTO / make group written must be in the master lists
    review.report(df, log)                                               # fills the remarks column
    return df[OUTPUT_COLUMNS], warnings


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
    sheet = EXTRA_SHEETS.get("Make groups to create")
    if sheet is not None and len(sheet):                                       # a missing group this script knows how to make: its definition goes into the list (nothing for a person to look up)
        sheet = sheet.copy()
        sheet["items"] = [(f"made of {MAKE_DEFINITIONS[g][0]} except {', '.join(MAKE_DEFINITIONS[g][1])}" if g in MAKE_DEFINITIONS else it) for g, it in zip(sheet["group_needed"], sheet["items"])]
        EXTRA_SHEETS["Make groups to create"] = sheet
    if CD1_AUTO:
        add_sheet("CD1 readings", cd1_readings_frame())
    if WORDINGS:
        add_sheet("Wording readings", wording_readings_frame())
    if UNREAD_PARTS:
        add_sheet("Unread parts", pd.DataFrame(UNREAD_PARTS, columns=["sheet", "where", "holds", "cells", "status", "why"]))
    if NOTES_READ:
        add_sheet("Notes read", pd.DataFrame(NOTES_READ, columns=["where", "text", "status", "why"]))
    if D_IGNORED:
        add_sheet("D ignored", pd.DataFrame(D_IGNORED))


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


def main(grid_path, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None, password=None):
    ask = ask or Asker()
    df, warnings = build(grid_path, from_date, to_date, log, ask, password=password)
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
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"GODIGIT_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
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
        question="You uploaded more than one workbook. Which one is this month's Go Digit grid?", short="workbook")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the Go Digit agency grid (.xls / .xlsx) and, optionally, an earlier upload (.csv) to compare with")
    try:
        try:
            import msoffcrypto  # noqa: F401
            import xlrd  # noqa: F401
        except ImportError:
            say("Installing the two small packages that read the protected .xls (msoffcrypto-tool, xlrd) ...")
            import subprocess
            subprocess.run([sys.executable, "-m", "pip", "install", "-q", "msoffcrypto-tool", "xlrd"], check=False)
        grid, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("godigit_agency_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="Go Digit agency grid -> Go Digit upload CSV")
    ap.add_argument("--grid", help="Go Digit workbook (.xls / .xlsx) - asked for if left out")
    ap.add_argument("--password", help="the workbook's password (or set GRID_PASSWORD; asked for, hidden, if left out); it is never stored")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    try:
        grid = a.grid or ask_file(ask, "no workbook given - pass --grid <workbook.xls>",
                                  "Path of the Go Digit workbook (.xls / .xlsx) - you can drag the file into this window:", exts=EXCEL_EXTS)
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
        main(grid, from_date, to_date, a.out, ref, ask=ask, password=a.password)
    except (GridError, ValueError) as e:
        sys.exit(f"STOPPED - {e}")
    except KeyboardInterrupt:
        sys.exit("STOPPED - stopped by you. No file was written.")
    except Exception as e:
        saved = crash_report("godigit_agency_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
