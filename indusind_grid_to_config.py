"""
IndusInd General Insurance (Reliance General) commercial-vehicle grid -> IndusInd upload CSV
=============================================================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (NN).csv") from the workbook the insurer circulates for brokers
("Reliance_-_TC_Grid_for_Brokers_Premier-_Nov26.xlsx"; the insurer's name in the upload is INDUSIND GENERAL INSURANCE COMPANY LIMITED -
"Reliance General" and "IndusInd General" are one insurer).  Rates come only from that workbook.  The full description (grid structure, how a header
becomes rows, rules, edge cases, validations) is INDUSIND_GRID.md; the short memory of the insurer is .claude/skills/indusind-grid/SKILL.md.

What the workbook holds and how it becomes rows (checked against the user's own export of the November 2026 grid: 11,270 of its 11,304 rows are equal)
--------------------------------------------------------------------------------------------------------------------------------------------------
* ONE rate table per product family, found by what its header SAYS and never by a cell address: a header row with a region column ('RTO Region', often
  twice: a parent region and the cluster) and, above the payout columns, a header band of up to three merged rows - the make label ('TATA / Maruti /
  Mahindra ( Only Jeeto & Supro )'), the product segment ('GCV < 2.0 K', 'PCV TAXI <6 Str ( NND)', 'MISD CPM (Only JCB, L&T and Caterpillar)') and the
  leaf ('COMP', 'STP', 'Comp - Others ( Petrol + CNG + Battery )', 'Diesel - Comp & STP', 'Comp (Fresh)').  `read_chain` carries merged cells into every
  column they cover; `classify_chain` reads the words of the chain (every word is USED UP by a pattern; a word nobody understands refuses the column) and gives
  the product, the make group, the fuels, the age, the weight (cc) or seating (sc) band, the nil-dep flag, the T&C and the policy types.
* One row per region x column x fuel: Comp = 1+1 / On Net, STP = On TP, a cell = the payin as a fraction (0.65 = 65), po = pi - 7 (never below 0), a zero or
  blank cell is a ZERO card.  The cluster name is the region exactly as the grid writes it (the user's standing rule: new_rto_list carries the bare cluster).
* The sheets of the other products (Private Car, TW Scooter, Short Term Taxi) are NOT converted in this build (the user, 20 Nov 2026): they are listed in the
  review workbook's sheet 'Not written' with what they hold, and so is every column that is deliberately left out (a deal-based column).
* The text sheets (Guidelines, Contest) and any text above / under the table are NOTES: each block is flagged  Manual Review Required: [text]  on the rows it may
  affect (the product it names; a table keyed by region names keeps to those regions), never guessed, never asked about.

Conditional readings (only when the wording is in the grid - the user, 20 Nov 2026: no hard-coding, a reusable header-driven brain)
-----------------------------------------------------------------------------------------------------------------------------------
* a model-level qualifier in a make label ('Only Jeeto & Supro', 'Mahindra Below 2T other than Jeeto & Supro') cannot be written with make groups: the rows of
  that column are written for the makes and flagged;
* a make label gives the make group ('TATA / Maruti / Mahindra' = TATA,Maruti,Mahindra; 'Others Tata / Maruti / Mahindra' = GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA;
  no label = GCV_4W_ALL), a taxi make list gives the master's own group, 'Other makes' = the complement of the make list of the block before it;
* a column with no T&C that would pay every vehicle of its class (Car Carrier = MISCD-Others with no T&C) is flagged where its payout is NOT zero;
* an unknown word, an unreadable cell, a missing table or a table found twice STOPS the run (or refuses that column) with a message that says what was
  expected and what was found - never a guess.

Dates (the user, 20 Nov 2026): the workbook's own title is never read; from_date and to_date are given on EVERY run.

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a payout cell that is not a number,
two sheets that both hold the table, dates that look wrong - it stops, asks you a specific question (type the number of an option, or q to stop, or ? for the
details), remembers the answer and carries on.  Every answer is listed at the end of the run, and you confirm before the file is written.  It never asks you for a
rate.  With --no-prompt (or when there is no keyboard) it asks nothing and stops with the question as the error message instead.  If something goes wrong that it did
not foresee it says so in one line and saves the technical details to a *_error_*.txt file.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run indusind_grid_to_config.py`) and run it.  Upload the workbook (and, optionally, an earlier upload CSV to get a
diff report); it asks for from_date and to_date (DD/MM/YYYY), writes them into every row, and downloads INDUSIND_UPLOAD_<yyyy-mm>.csv.  The master workbooks (RTO /
make-model) must be in a `master` folder next to the script.

Command line
------------
    python indusind_grid_to_config.py [--grid <workbook.xlsx>] [--from-date 01/11/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--out file.csv] [--no-prompt]
(anything left out - the workbook, the dates, the earlier upload - is asked for)
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
INSURANCE_COMPANY = "INDUSIND GENERAL INSURANCE COMPANY LIMITED"      # Reliance General and IndusInd General are one insurer (the user, 20 Nov 2026)
USER_ID = 2966
ROW_ID = 1                                       # the portal's export carries id = 1 on every row
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 6                                # a cell x 100 carries float noise (0.47500000000000003 -> 47.50000000000001): cleaned here, the CSV rounds to 4
POLICY_TYPE = "Renew"
PI_TYPES = {"1+1": "On Net", "STP": "On TP"}     # the pi / po type of a policy type (the user's export)
CC_RANGE, SC_RANGE, AGE_RANGE = (0, 99999), (0, 999), (0, 99)
CC_OPEN_END = 99999
TARIFF = (0, 100.01)                             # no discount band in this grid: every row is 0-100.01
MAX_NOTE_CHARS = 500
ANCHOR_SCAN = 30                                 # the region header is looked for in the first rows
DATA_PROBE = 12                                  # ... and the first data row within this many rows below it
MIN_DATA_ROWS = 5                                # a table with fewer rows is not a rate table

# The portal's RTO groups are looked up ONLY to flag a region that has none (the user's decision, 20 Nov 2026: 'write all Indore rows and flag them as no RTO group');
# the name WRITTEN is the bare region as the grid spells it (the standing cluster rule).  RTO_GROUP_PREFIX is also the group name of GRID_RTO_NAMES=group.
RTO_GROUP_PREFIX = "RELIANCE_SEPT26_"
RTO_GROUP_ALIASES = {"trivandrum": "Thiruvananthapuram"}      # a region the master spells another way (checked by name only)
RTO_GROUP_MIN_SHARE = 0.2                        # when fewer regions than this have a group under the prefix the prefix is stale: the check is skipped with a warning

FUELS_ALL = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]
FUELS_NON_DIESEL = ["Petrol", "CNG", "LPG", "Electric"]
FUEL_WORDS = {"petrol": ["Petrol"], "diesel": ["Diesel"], "cng": ["CNG", "LPG"], "lpg": ["LPG"], "battery": ["Electric"], "electric": ["Electric"], "ev": ["Electric"]}
# 'CNG' in a header = CNG + LPG: the user's export carries LPG rows next to CNG although the grid says 'Petrol + CNG + Battery' (decided 20 Nov 2026, as in the Go Digit script)

ZERO_WORDS = ("NA", "N/A", "-", "--", "NIL", "DECLINE", "DECLINED", "BLOCK", "BLOCKED", "NOT APPLICABLE", "NO BIZ")     # a rate cell that says no rate = a ZERO card
NEW_TNC, OLD_TNC = "New", "Old"

# the product families: the master's 'all' group of each, and the prefix of its 'all except these makes' group (the master's own spelling)
FAMILY_ALL = {"bus": "BUS_ALL", "pcv3w": "PCV_3W_ALL", "taxi": "TAXI_ALL", "gcv3w": "GCV_3W_ALL", "gcv4w": "GCV_4W_ALL", "miscd": "MISCD_ALL"}
# a make word of a header -> (key, the master's make group of that make)
MAKE_WORDS = {"tata": ("tata", "TATA"), "tata motors": ("tata", "TATA"), "maruti": ("maruti", "Maruti"), "maruti suzuki": ("maruti", "Maruti"), "msil": ("maruti", "Maruti"),
              "mahindra": ("mahindra", "Mahindra"), "m&m": ("mahindra", "Mahindra"), "mahindra & mahindra": ("mahindra", "Mahindra"), "toyota": ("toyota", "Toyota"),
              "kia": ("kia", "KIA"), "mg": ("mg", "MG"), "mg motor": ("mg", "MG"), "ashok leyland": ("ashok", "Ashok"), "ashok": ("ashok", "Ashok"), "al": ("ashok", "Ashok"),
              "eicher": ("eicher", "EICHER"), "force": ("force", "Force"), "isuzu": ("isuzu", "Isuzu"), "hyundai": ("hyundai", "Hyundai"), "honda": ("honda", "Honda")}
# a make set the master names by its own group (the user, 20 Nov 2026: 'use the exact names from make_model_master'): (family, sorted make keys, complement) -> the group.
# A set that is not here is written as the single-make groups (a list) or as <family group>_EX_<makes> (a complement) and checked against the master like every name.
MAKE_SETS = {("taxi", ("kia", "mahindra", "maruti", "mg", "toyota"), False): "TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG",
             ("taxi", ("kia", "mahindra", "maruti", "mg", "toyota"), True): "TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG"}
# the T&C words of a body-type column: (words of the segment, the T&C of its rows, the T&C after New / Old, the make_model_class); a segment whose body words are not here is not written
BODY_TYPES = [
    ("car_carrier", r"car\s*carrier", None, "", "", "MISCD-Others"),
    ("flat_bed", r"flat\s*bed", None, "Flat Bed", "Flat Bed", "MISCD-Others"),
    ("cpm", r"mis[cd]*\s*cpm", frozenset({"jcb", "l&t", "caterpillar"}), "JCB, L&T,Caterpillar", "JCB, L&T,Caterpillar", "MISCD-Others"),
    ("tractor", r"tractor", None, "", "Agriculture Tractor", "MISCD-Tractor"),
]

# the user's old make names -> the master's, for the comparison with an earlier upload only (the upload itself writes the master's names)
REF_MAKE_ALIASES = {"maruti,toyota,mahindra,kia,mg": "taxiall_mahi_kia_toyota_maruti_mg", "taxi_all_ex_maruti_toyota_mahindra_kia_mg": "taxiall_ex_mahi_kia_toyota_maruti_mg"}

# a note that is a label or a title of the grid and needs no action (the standing notes of this script; the web app adds the person's own)
STANDING_STRICT = True                           # a standing phrase covers its own words only (the shared unknown_notes, 18 Nov 2026)
STANDING_NOTES = [
    re.compile(r"^\s*PO\s+Grid\s+for\b", re.I),                  # the (hidden) title 'PO Grid for Jun'26 FY 26-27': the dates are typed on every run, never read from it
    "PO at ODD as per UW",
]
STANDING_WHY = {STANDING_NOTES[0]: "the title of the workbook: from_date and to_date are given on every run, the title is never read (the user, 20 Nov 2026)",
                "PO at ODD as per UW": "the header over the zone / region columns: nothing to apply"}

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

OUTPUT_FOLDER = "indusind"           # the folder of this insurer under output\ (default_output)


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
def _text(v):
    return "" if v is None else str(v).replace("\xa0", " ").replace("​", "").strip()


def col_letter(c):
    """1-based column number -> letters (1 = A)"""
    from openpyxl.utils import get_column_letter
    return get_column_letter(c)


def read_workbook(path):
    """Open with values only. NOT read-only mode: the header cells are merged and only the normal mode keeps the merged ranges."""
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def sheet_rows(wb, name):
    return [list(r) for r in wb[name].iter_rows(values_only=True)]


def label_key(s):
    """A label without spaces, punctuation and capitals: 'RTO Region' and 'rto  region ' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def norm(v):
    """a cell as one line of text: line breaks and runs of spaces become one space"""
    return re.sub(r"\s+", " ", _text(v)).strip()


def low(v):
    return norm(v).lower()


def num(v):
    """the number a cell holds (a payout fraction), None for text, blanks, booleans and errors"""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v) if math.isfinite(v) else None


def hidden_sheets(wb):
    return {n for n in wb.sheetnames if wb[n].sheet_state != "visible"}


def hidden_columns(ws):
    """the numbers of the columns hidden in Excel (one entry of the sheet's column dimensions may span several columns)"""
    from openpyxl.utils import column_index_from_string
    out = set()
    for key, dim in ws.column_dimensions.items():
        if dim.hidden:
            lo = dim.min or column_index_from_string(key)
            out.update(range(lo, (dim.max or lo) + 1))
    return out


class Refuse(Exception):
    """a header (or a part of one) the reader cannot place: the column is not written, the reason is listed in the review workbook and warned about - never a guess"""


class Sheet:
    """A worksheet with its merged ranges carried into every cell they cover (the value of the top-left cell), so a merged header is read in every column it spans."""

    def __init__(self, ws):
        self.ws, self.name = ws, ws.title
        self.merged = {}
        for mr in ws.merged_cells.ranges:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    self.merged[(r, c)] = (mr.min_row, mr.min_col)
        self.max_row, self.max_col = ws.max_row, ws.max_column

    def raw(self, r, c):
        return self.ws.cell(r, c).value

    def val(self, r, c):
        return self.ws.cell(*self.merged.get((r, c), (r, c))).value

    def rows(self):
        return [list(r) for r in self.ws.iter_rows(values_only=True)]


# ---- finding the rate table by what its header SAYS
REGION_HEADS = ["rto region", "region", "rto cluster", "cluster"]
REGION_WORDS = ("region", "cluster")
REGION_HEAD_MAX_WORDS = 8                        # a header cell of at most this many words that holds 'region' / 'cluster' says a region, whatever else it says ('RTO Region FY 26-27 (New)')


def region_header(v):
    """True when a cell says a region: the words of REGION_HEADS (head_match), or a SHORT cell that holds the word 'region' / 'cluster' among a few others (a year, 'FY 26-27', '(new)'
    - October 2026 headed its clusters 'RTO Region FY 26-27', which the exact match only just allowed); a code column ('Region_Code') is not one"""
    if not isinstance(v, str) or not v.strip():
        return False
    w = _hwords(v)
    if any(x in ("code", "codes") for x in w):                               # a code column is never the column that names the regions
        return False
    if any(head_match(v, s) for s in REGION_HEADS):
        return True
    return len(w) <= REGION_HEAD_MAX_WORDS and any(x in REGION_WORDS for x in w) and not any(x in ("no", "count") for x in w)


def head_text(sh, r, c):
    """the header text of a cell (merged cells carried into every cell they cover): '' when nothing readable is in it - a stray backtick, dash or bracket is noise, not a header"""
    t = norm(sh.val(r, c))
    return t if re.search(r"[A-Za-z0-9]", t) else ""


def find_anchor(sh):
    """(header row, [the columns that say a region]) of the first row, within the first ANCHOR_SCAN rows, that says a region header and has the header band of a table to its right
    (at least three columns with text in the rows around it); None when no row does.  The region header is usually there twice (a parent region and the cluster)."""
    for r in range(1, min(sh.max_row, ANCHOR_SCAN) + 1):
        regs = [c for c in range(1, sh.max_col + 1) if region_header(sh.val(r, c))]
        if not regs:
            continue
        right = [c for c in range(max(regs) + 1, sh.max_col + 1) if any(norm(sh.val(rr, c)) for rr in range(max(1, r - 3), min(sh.max_row, r + 1) + 1))]
        if len(right) >= 3:
            return r, regs
    return None


class RateTable:
    """One rate table of a sheet: the header band above the first data row (merged cells carried into every column they cover), the region column(s), the data rows.
    The chain of a column = the distinct texts of its header cells from the top of the band to the leaf; empty rows, a merged cell repeated over several rows and one more
    header row change nothing.  Nothing is found by a cell address."""

    def __init__(self, sh, anchor, regs):
        self.sh, self.anchor, self.regs = sh, anchor, regs
        self.c0 = max(regs) + 1                                              # the first column that can hold a payout
        self.first = next((r for r in range(anchor + 1, min(sh.max_row, anchor + DATA_PROBE) + 1) if self.is_data(r)), None)
        self.ok = self.first is not None
        self.rcol = self.parent = None
        self.banner = []
        if not self.ok:
            return
        top, r, gap = anchor, anchor - 1, 0                                 # the header band reaches up while the rows above have text over the payout columns; a blank row (up to two)
        while r >= 1 and gap <= 2:                                          # inside the band does not end it when the text above it still reads like a header (a product, a make, Comp / STP)
            texts = [x for x in (head_text(sh, r, c) for c in range(self.c0, sh.max_col + 1)) if x]
            if texts:
                if gap == 0 or any(header_like(x) for x in texts):
                    top, gap = r, 0
                else:
                    break
            else:
                gap += 1
            r -= 1
        self.top = top
        self.chain = {}
        for c in range(self.c0, sh.max_col + 1):
            ch = []
            for r in range(top, self.first):
                t = head_text(sh, r, c)
                if t and (not ch or ch[-1] != t):
                    ch.append(t)
            self.chain[c] = ch
        heads = [ch[0] for ch in self.chain.values() if len(ch) >= 2]
        for text in set(heads):                                              # a banner over the whole table (one text on every column) is a note, not a header level
            if not header_like(text) and heads.count(text) >= max(3, int(0.9 * sum(1 for ch in self.chain.values() if ch))):
                self.banner.append(text)
                for c, ch in self.chain.items():
                    if len(ch) >= 2 and ch[0] == text:
                        self.chain[c] = ch[1:]
        self.data, self.text_rows = [], []
        for r in range(self.first, sh.max_row + 1):
            if self.is_data(r):
                self.data.append(r)
            elif any(norm(sh.raw(r, c)) for c in range(1, sh.max_col + 1)):
                self.text_rows.append(r)

    def region_text(self, r):
        for c in self.regs:
            t = norm(self.sh.raw(r, c))
            if t and num(self.sh.raw(r, c)) is None:
                return t
        return ""

    def numeric_count(self, r):
        return sum(1 for c in range(self.c0, self.sh.max_col + 1) if num(self.sh.raw(r, c)) is not None)

    def is_data(self, r):
        return bool(self.region_text(r)) and self.numeric_count(r) >= 3

    def pick_region(self, log, ask):
        """The column that names the clusters.  When two columns say 'RTO Region' (a parent region and the cluster) the one with MORE distinct names is the cluster (every row of it is
        a different name; the parent repeats); two columns with the same number of distinct names that differ stop the run - which one is meant is not in the document."""
        sh = self.sh
        vals = {c: [norm(sh.raw(r, c)) for r in self.data] for c in self.regs}
        distinct = {c: len({v for v in vs if v}) for c, vs in vals.items()}
        best = max(distinct.values())
        tops = [c for c in self.regs if distinct[c] == best]
        if len(tops) > 1 and any(vals[c] != vals[tops[0]] for c in tops):
            raise GridError(f"expected ONE column that names the regions of '{sh.name}', found {len(tops)} columns that say 'RTO Region' ({', '.join(col_letter(c) for c in tops)}) and hold "
                            f"different names with the same number of distinct values - which one names the clusters is not in the document")
        self.rcol = tops[0]
        others = [c for c in self.regs if c != self.rcol]
        self.parent = others[0] if others else None
        if others:
            ask.decide(f"'{sh.name}': {len(self.regs)} columns say 'RTO Region' ({', '.join(col_letter(c) for c in self.regs)}): {col_letter(self.rcol)} names the clusters "
                       f"({distinct[self.rcol]} distinct names, one per row); {col_letter(others[0])} holds {distinct[others[0]]} (the parent region) and is not used")
        else:
            cand = self.parent_column(distinct[self.rcol])
            if cand:
                self.parent = cand
                ask.decide(f"'{sh.name}': {col_letter(self.rcol)} names the clusters ({distinct[self.rcol]} distinct names, one per row); {col_letter(cand)} "
                           f"('{head_text(sh, self.anchor, cand)}') holds {len({norm(sh.raw(r, cand)) for r in self.data})} (the parent region) and is not used")

    def parent_column(self, n_cluster):
        """the parent region's column when its header does not say 'region' at all ('RTO Fy 26 -27'): the nearest column left of the cluster column whose header (anywhere in the band) says
        RTO / region / cluster, that holds text on every data row and fewer distinct names than the cluster column; None when there is none"""
        sh = self.sh
        for c in range(self.rcol - 1, 0, -1):
            if not any(set(_hwords(head_text(sh, r, c))) & {"rto", "region", "cluster"} for r in range(self.top, self.first)):
                continue
            vals = [norm(sh.raw(r, c)) for r in self.data]
            if all(vals) and all(num(sh.raw(r, c)) is None for r in self.data) and 1 < len(set(vals)) < n_cluster:
                return c
        return None

    def name_of(self, r):
        return norm(self.sh.raw(r, self.rcol))

    def parent_of(self, r):
        return norm(self.sh.raw(r, self.parent)) if self.parent else ""

    def zone_of(self, r):
        z = [c for c in range(1, min(self.regs) ) if norm(self.sh.raw(r, c))]
        return norm(self.sh.raw(r, z[0])) if z else ""

    def columns(self):
        """every column right of the region columns that has a header or a number: dict(c, letter, chain, numbers, texts)"""
        out = []
        for c in range(self.c0, self.sh.max_col + 1):
            nums = [r for r in self.data if num(self.sh.raw(r, c)) is not None]
            texts = [r for r in self.data if num(self.sh.raw(r, c)) is None and norm(self.sh.raw(r, c))]
            if not self.chain[c] and not nums and not texts:
                continue
            out.append(dict(c=c, letter=col_letter(c), chain=self.chain[c], numbers=len(nums), texts=len(texts)))
        return out

    def path(self, col):
        return " / ".join(col["chain"]) if col["chain"] else "(no header)"


# ---- reading the words of a header chain (every word is USED UP by a pattern; a word nobody understands refuses the column)
KIND_INFO = {"school_bus": ("PCV 4W", "School Bus", "bus"), "other_bus": ("PCV 4W", "Other Bus", "bus"), "pcv3w": ("PCV 3W", "Passenger Carrying Rickshaw", "pcv3w"),
             "taxi": ("PCV 4W", "Taxi", "taxi"), "gcv3w": ("GCV 3W", "Goods Carrying Rickshaw", "gcv3w"), "gcv4w": ("GCV 4W", "NA", "gcv4w")}
# (kind, the regex that FINDS the product words, the regex that takes them out of the text): 'PCV 7+1' is a taxi too, and its seats stay for seats_of
SEG_PATTERNS = [("gcv3w", r"gcv\s*3\s*w", r"gcv\s*3\s*w"), ("pcv3w", r"pcv\s*3\s*w", r"pcv\s*3\s*w"), ("school_bus", r"school\s*bus", r"school\s*bus"),
                ("other_bus", r"other\s*bus", r"other\s*bus"), ("taxi", r"taxi|pcv\s*\d+\s*\+\s*\d+", r"taxi"),
                ("gcv4w", r"\b(?:gcv|lcv|mcv|hcv)\b(?!\s*3\s*w)", r"\b(?:gcv|lcv|mcv|hcv)\b"), ("pickup", r"employee\s*pick\s*up", r"employee\s*pick\s*up")] \
    + [(k, rx, rx) for k, rx, *_ in BODY_TYPES] \
    + [("private_car", r"\bpc\b|\bpvt\.?\s*car\b|private\s*car", r"\bpc\b|\bpvt\.?\s*car\b|private\s*car"),
       ("two_wheeler", r"\btw\b|two[\s-]*wheel|scooter|\bbikes?\b", r"\btw\b|two[\s-]*wheel|scooter|\bbikes?\b")]
BODY_KIND = {k: (words, tnc, aged, klass) for k, _rx, words, tnc, aged, klass in BODY_TYPES}
# a product of another sheet that shows up as columns of this one (October 2026: 'PC STP Non-Diesel' x 3 cc bands inside the commercial-vehicle sheet): understood, NOT converted in this build
# (the user, 20 Nov 2026: Private Car / TW are 'Not written') - listed in the sheet 'Not written', never refused as an unknown word and never written
NOT_CONVERTED = {"private_car": "Private Car", "two_wheeler": "Two Wheeler"}
# a STYLE of a taxi (October 2026: the column 'PCV TAXI KAALI PEELI / Kaali Peeli'): (name, the words, the T&C of its rows, the products it can belong to).  The upload has no vehicle-type field,
# so the style is written as a T&C (the user, 23 Nov 2026: 'T&C Kaali Peeli') and the rows are flagged: the ordinary cards of the product also cover these vehicles
STYLES = [("Kaali Peeli", r"ka+l+i+\s*p[ei]+l+i+", "Kaali Peeli", ("taxi",))]
IGNORABLE = {"pcv", "gcv", "str", "strs", "seater", "seats", "k", "t"}                    # words of a title that carry no rule (the units of a band, the product letters)
KIND_EXTRA = {"tractor": {"agricultural", "agriculture", "trailer", "trailers", "without"}}   # more words a product title may carry: 'Agricultural Tractor W/o Trailer'
QUALIFIER_RX = re.compile(r"\b(?:only|other\s+than|except|excluding|below|above|upto|up\s+to)\b", re.I)


def kinds_in(text):
    t = low(text)
    return [k for k, rx, _rm in SEG_PATTERNS if re.search(rx, t)]


LEAF_RX = re.compile(r"\b(?:comp\w*|stp|fresh|all|diesel|petrol|cng|years?|yrs?)\b")


def header_like(text):
    """True when a header text reads like a header of this grid: it names a product, a Comp / STP / fuel word or a make"""
    t = low(text)
    return bool(kinds_in(t) or LEAF_RX.search(t) or any(re.search(r"\b" + re.escape(w) + r"\b", t) for w in MAKE_WORDS) or re.match(r"\s*others?\b", t))


def split_makes(text):
    """([make keys in the order written], [the parts that are no make]) of a list like 'Maruti, Toyota & Mahindra'"""
    keys, rest = [], []
    for p in (x.strip() for x in re.split(r"\s+&\s*|\s*&\s+|\s*[,/+]\s*|\s+and\s+", low(text))):          # 'Kia& MG' splits too; 'm&m' and 'l&t' (no spaces around the &) do not
        if not p:
            continue
        if p in MAKE_WORDS:
            if MAKE_WORDS[p][0] not in keys:
                keys.append(MAKE_WORDS[p][0])
        else:
            rest.append(p)
    return keys, rest


MAKE_NAME = {k: name for k, name in MAKE_WORDS.values()}


def make_group(family, makes, complement):
    """The make group a make list gives: no list = the family's 'all' group; a list the master names by its own group (MAKE_SETS); else the single-make groups (a list) or
    '<family group>_EX_<makes>' (all except them), which `check_names` looks up in the master like every name."""
    if not makes:
        return FAMILY_ALL[family]
    if (family, tuple(sorted(makes)), bool(complement)) in MAKE_SETS:
        return MAKE_SETS[(family, tuple(sorted(makes)), bool(complement))]
    names = [MAKE_NAME[m] for m in makes]
    if not complement:
        return ",".join(names)
    name = f"{FAMILY_ALL[family]}_EX_" + "_".join(n.upper() for n in names)
    MAKE_DEFINITIONS[name] = (FAMILY_ALL[family], names)
    return name


def fuels_of(text):
    """the fuels a header text names ('Petrol + CNG + Battery', 'Diesel', 'All Fuel Type', 'Non diesel'), in the grid's order; None when it holds anything else"""
    q = low(text)
    if re.fullmatch(r"all\s+fuels?(?:\s+types?)?", q):
        return list(FUELS_ALL)
    if re.fullmatch(r"non[\s-]*diesel", q):
        return list(FUELS_NON_DIESEL)
    out = []
    for p in (x.strip() for x in re.split(r"\s*(?:\+|,|/|&|\band\b)\s*", q)):
        if not p:
            continue
        if p not in FUEL_WORDS:
            return None
        out += [f for f in FUEL_WORDS[p] if f not in out]
    return sorted(out, key=FUELS_ALL.index) if out else None


def band_of(w):
    """((cc_min, cc_max) in kg, w without the band) of a GCV title's weight words ('< 2.0 K', '2 - 2.5 K', '2.5K - 3.5K', 'Above 3.5K', '7.5 T - 12 T'; K or T = a thousand kg); None when there is none"""
    n, unit = r"(\d+(?:\.\d+)?)", r"(?:k|tn|tons?|tonnes?|t)\b"
    m = re.search(rf"{n}\s*(?:{unit})?\s*[-–]\s*{n}\s*{unit}", w)
    if m:
        return (round(float(m.group(1)) * 1000 + 0.01, 2), round(float(m.group(2)) * 1000 + 0.01, 2)), w[:m.start()] + " " + w[m.end():]
    m = re.search(rf"(?:<=?|below|under|upto|up\s+to)\s*{n}\s*{unit}", w)
    if m:
        return (0, round(float(m.group(1)) * 1000 + 0.01, 2)), w[:m.start()] + " " + w[m.end():]
    m = re.search(rf"(?:>=?|above|over|more\s+than)\s*{n}\s*{unit}", w)
    if m:
        return (round(float(m.group(1)) * 1000 + 0.01, 2), CC_OPEN_END), w[:m.start()] + " " + w[m.end():]
    return None


def seats_of(w):
    """((sc_min, sc_max), w without the words) of a title's seating words: '<6 Str' (0, 6.01), '7+1' (6.01, 8.01), '> 8 STR' (8, 999); None when there are none"""
    m = re.search(r"(\d+)\s*\+\s*(\d+)", w)
    if m:
        n = int(m.group(1))
        return (round(n - 1 + 0.01, 2), round(n + int(m.group(2)) + 0.01, 2)), w[:m.start()] + " " + w[m.end():]
    m = re.search(r"<=?\s*(\d+)\s*(?:str\w*|seat\w*)?", w)
    if m:
        return (0, round(int(m.group(1)) + 0.01, 2)), w[:m.start()] + " " + w[m.end():]
    m = re.search(r">\s*(\d+)\s*(?:str\w*|seat\w*)?", w)
    if m:
        return (int(m.group(1)), SC_RANGE[1]), w[:m.start()] + " " + w[m.end():]
    return None


def eat_paren(p, seg):
    """True when the bracketed text `p` of a segment title is understood (and has been applied to `seg`)"""
    q = low(p)
    m = re.fullmatch(r"(?:only\s+)?(?:(nnd|nd)|(without\s+|with\s+)?nil\s*dep(?:reciation)?)", q)
    if m:                                                                    # NND / 'without nil dep' = no nil dep (is_zd NO), ND / 'nil dep' = nil dep (YES)
        set_zd(seg, "NO" if (m.group(1) == "nnd" or (m.group(2) or "").startswith("without")) else "YES")
        return True
    if re.fullmatch(r"(?:only\s+)?deals?\s*(?:based|basis)", q):             # a negotiated deal, not a rate card ('Only Deal Based', 'Only deal basis')
        seg["deal"], seg["deal_words"] = True, norm(p)
        return True
    if re.fullmatch(r"other\s+makes?", q):
        seg["other_makes"] = True
        return True
    if re.fullmatch(r"others?", q):                                          # '(Others)' beside '(Diesel Only)': the fuels other than diesel - the Comp / STP header lists them
        seg["fuel_other"] = True
        return True
    if re.fullmatch(r"(?:only\s+)?diesel(?:\s+only)?", q):
        seg["fuels"] = ["Diesel"]
        return True
    f = fuels_of(q)
    if f:
        seg["fuels"] = f
        return True
    if seg["kind"] == "tractor" and re.fullmatch(r"agricultur\w*\s+without\s+trailers?", q):
        return True
    m = re.fullmatch(r"only\s+(.+)", q)
    if m and seg["kind"] in BODY_KIND and BODY_KIND[seg["kind"]][0] is not None:
        seg["body_words"] = frozenset(x.strip() for x in re.split(r"\s*,\s*|\s+and\s+", m.group(1)) if x.strip())
        return True
    keys, rest = split_makes(q)
    if keys and not rest:
        seg["makes"] = keys
        return True
    return False


def set_zd(seg, value):
    """nil dep of a segment: NO / YES; two texts of one header that say different things are a conflict (a refusal), never a pick"""
    if seg["zd"] not in (None, value):
        raise Refuse(f"'{seg['text']}' says both 'nil dep' and 'without nil dep'")
    seg["zd"] = value


def eat_zd(w):
    """(the words without the nil-dep wording, NO / YES / None) of a segment title outside brackets: 'Without Nil Dep', 'Nil Dep', NND, ND"""
    for rx, value in ((r"\bwithout\s+nil\s*dep(?:reciation)?\b", "NO"), (r"\bnnd\b", "NO"), (r"\b(?:with\s+)?nil\s*dep(?:reciation)?\b", "YES"), (r"\bnd\b", "YES")):
        m = re.search(rx, w)
        if m:
            return w[:m.start()] + " " + w[m.end():], value
    return w, None


def parse_clause(text, final=True):
    """One product clause of a segment title -> dict(kind, product, klass, family, cc, sc, zd, fuels, fuel_other, makes, other_makes, qualifiers, deal, style, tnc, tnc_aged, text); Refuse when a
    word is not understood.  `final=False` leaves the check of the body words to the caller: a header that names its product twice (a group title and a sub-title) is merged first."""
    t = norm(text)                                                           # a stray bracket ('Agricultural Tractor W/o Trailer)') is punctuation: only words are read
    parens = [norm(p) for p in re.findall(r"\(([^)]*)\)", t)]
    base = low(re.sub(r"\([^)]*\)", " ", t))
    kinds = kinds_in(base)
    if len(kinds) != 1:
        raise Refuse(f"the product of '{t}' is " + ("not one this script knows" if not kinds else f"ambiguous ({', '.join(kinds)})"))
    kind = kinds[0]
    if kind in NOT_CONVERTED:
        raise Refuse(f"'{t}' is a {NOT_CONVERTED[kind]} column: not converted in this build")
    seg = dict(kind=kind, fuels=None, fuel_other=False, makes=[], other_makes=False, qualifiers=[], deal=False, zd=None, cc=CC_RANGE, sc=SC_RANGE, tnc="", tnc_aged="", text=t, body_words=None, style=None, deal_words="")
    w = next(re.sub(rm, " ", base, count=1) for k, _rx, rm in SEG_PATTERNS if k == kind)
    for name, rx, style_tnc, ok_kinds in STYLES:                             # 'Kaali Peeli' in the words of a taxi title
        m = re.search(rx, w)
        if m:
            if kind not in ok_kinds:
                raise Refuse(f"'{name}' in '{t}': this script knows that style for {', '.join(ok_kinds)} only")
            seg["style"] = (name, style_tnc)
            w = w[:m.start()] + " " + w[m.end():]
    if kind == "pickup":
        pass
    elif kind in KIND_INFO:
        seg["product"], seg["klass"], seg["family"] = KIND_INFO[kind]
    else:
        words, tnc, aged, klass = BODY_KIND[kind]
        seg.update(product="MISCD", klass=klass, family="miscd", tnc=tnc, tnc_aged=aged)
    m = re.search(r"\ball\s+fuels?(?:\s+types?)?\b", w)
    if m:
        seg["fuels"] = list(FUELS_ALL)
        w = w[:m.start()] + " " + w[m.end():]
    if kind == "gcv4w":
        got = band_of(w)
        if not got:
            raise Refuse(f"no weight band in '{t}'")
        seg["cc"], w = got
    elif kind in ("taxi", "other_bus"):
        got = seats_of(w)
        if got:
            seg["sc"], w = got
    w, zd = eat_zd(w)                                                        # 'Without Nil Dep' / 'Nil Dep' outside brackets
    if zd:
        set_zd(seg, zd)
    m = re.search(r"\b(?:only\s+)?deals?\s*(?:based|basis)\b", w)
    if m:
        seg["deal"], seg["deal_words"] = True, norm(m.group(0))
        w = w[:m.start()] + " " + w[m.end():]
    for p in parens:
        if not eat_paren(p, seg):
            if QUALIFIER_RX.search(p):
                seg["qualifiers"].append(p)
            else:
                raise Refuse(f"the bracketed text '({p})' of '{t}' is not understood")
    left = [x for x in re.findall(r"[a-z]{2,}", w) if x not in IGNORABLE and x not in KIND_EXTRA.get(kind, ())]
    digits = re.findall(r"\d+(?:\.\d+)?", w)
    if left or digits:
        raise Refuse(f"words or numbers of '{t}' that nothing understands: {', '.join(left + digits)}")
    if final:
        finish_clause(seg)
    return seg


def finish_clause(seg):
    """the checks of a product clause that need ALL the texts of its header (merged first): the body words of a body-type column (MISD CPM: JCB, L&T and Caterpillar) - a deal column has no rows,
    so its body words do not matter"""
    kind = seg["kind"]
    if kind in BODY_KIND and BODY_KIND[kind][0] is not None and seg["body_words"] != BODY_KIND[kind][0] and not seg["deal"]:
        raise Refuse(f"the body words of '{seg['text']}' are {sorted(seg['body_words'] or [])}, not {sorted(BODY_KIND[kind][0])}: no T&C is known for them")


def merge_segs(parts):
    """ONE product clause from the texts of a header that name the same product twice (October 2026: a group title 'PCV TAXI <6 Seater' above a sub-title 'PCV TAXI < 6Str Without Nil Dep
    (Others)').  Texts that agree - or one of which says nothing about a property - are merged; texts that disagree (another product, another seat band, nil dep and without nil dep ...)
    REFUSE the column: which one applies is not in the document."""
    out = dict(parts[0])
    out["qualifiers"], out["makes"] = list(parts[0]["qualifiers"]), list(parts[0]["makes"])
    for p in parts[1:]:
        if p["kind"] != out["kind"]:
            raise Refuse(f"two product texts of one header name different products: '{out['text']}' and '{p['text']}'")
        for key, unset in (("cc", CC_RANGE), ("sc", SC_RANGE)):
            if tuple(p[key]) != tuple(unset):
                if tuple(out[key]) not in (tuple(unset), tuple(p[key])):
                    raise Refuse(f"'{out['text']}' and '{p['text']}' give different {'weight' if key == 'cc' else 'seat'} bands")
                out[key] = p[key]
        if p["zd"] is not None:
            set_zd(out, p["zd"]) if out["zd"] is not None else out.update(zd=p["zd"])
        for key in ("fuels", "style"):
            if p[key] is not None:
                if out[key] not in (None, p[key]):
                    raise Refuse(f"'{out['text']}' and '{p['text']}' give different {'fuels' if key == 'fuels' else 'styles'}")
                out[key] = p[key]
        if p["makes"]:
            if out["makes"] and out["makes"] != p["makes"]:
                raise Refuse(f"'{out['text']}' and '{p['text']}' give different make lists")
            out["makes"] = list(p["makes"])
        out["other_makes"] = out["other_makes"] or p["other_makes"]
        out["fuel_other"] = out["fuel_other"] or p["fuel_other"]
        out["deal"], out["deal_words"] = out["deal"] or p["deal"], out["deal_words"] or p["deal_words"]
        out["qualifiers"] += [q for q in p["qualifiers"] if q not in out["qualifiers"]]
        if p["body_words"] is not None:
            out["body_words"] = frozenset(out["body_words"] or ()) | frozenset(p["body_words"])
        if len(p["text"]) > len(out["text"]):                                # the most specific text names the segment in messages and flags
            out["text"] = p["text"]
    return out


LEAF_WORDS = {"comp": "1+1", "comprehensive": "1+1", "stp": "STP"}
AGE_COND_RX = r"(<=?|>=?|below|above|under|over|upto|up\s+to|less\s+than|more\s+than)\s*(\d+(?:\.\d+)?)\s*(?:years?|yrs?)\b"


def parse_leaf(text):
    """The leaf of a header chain ('COMP', 'STP', 'All', 'Comp/STP', 'Diesel - Comp & STP', 'Comp - Others ( Petrol + CNG + Battery )', 'Comp (Fresh)') -> dict(subs, fuels, age); Refuse otherwise"""
    t = norm(text)
    w = low(re.sub(r"[()]", " ", t))                                         # the brackets carry no structure in a leaf: the words are used up one by one
    age = None
    if re.search(r"non[\s-]*fresh", w):
        age, w = "nonfresh", re.sub(r"non[\s-]*fresh", " ", w)
    elif re.search(r"\b(?:fresh|new)\b", w):                                 # 'Comp (Fresh)' and 'Comp (New)' are the same: age 0-1 (the user, 23 Nov 2026)
        age, w = "fresh", re.sub(r"\b(?:fresh|new)\b", " ", w)
    lo = hi = None                                                           # an age condition of the leaf: '< 5 Years' (up to 5.01), '> 5 Years' (from 5.01) - the user's standard convention of 23 Nov 2026
    for m in list(re.finditer(AGE_COND_RX, w)):
        op, n = re.sub(r"\s+", " ", m.group(1)), float(m.group(2))
        if op in (">", "above", "over", "more than", ">="):
            if lo is not None:
                raise Refuse(f"two lower age limits in the column header '{t}'")
            lo = n if op == ">=" else round(n + 0.01, 2)
        else:
            if hi is not None:
                raise Refuse(f"two upper age limits in the column header '{t}'")
            hi = round(n + 0.01, 2)
    w = re.sub(AGE_COND_RX, " ", w)
    if age == "fresh" and (lo is not None or hi is not None):
        raise Refuse(f"a fresh / new column with an age limit in the column header '{t}': which one applies is not in the document")
    subs, fuels, left = [], [], []
    for word in re.findall(r"[a-z]+", w):
        if word in LEAF_WORDS:
            if LEAF_WORDS[word] not in subs:
                subs.append(LEAF_WORDS[word])
        elif word == "all":
            subs = ["1+1", "STP"]
        elif word in ("others", "other"):
            continue
        elif word in FUEL_WORDS:
            fuels += [f for f in FUEL_WORDS[word] if f not in fuels]
        else:
            left.append(word)
    if left:
        raise Refuse(f"words of the column header '{t}' that nothing understands: {', '.join(left)}")
    if not subs:
        raise Refuse(f"no Comp / STP / All wording in the column header '{t}'")
    return dict(subs=sorted(subs, key=("1+1", "STP").index), fuels=sorted(fuels, key=FUELS_ALL.index) or None, age=age, age_lo=lo, age_hi=hi)


def parse_make_label(text):
    """A make label above a segment ('TATA / Maruti / Mahindra ( Only Jeeto & Supro )', 'Others Tata / Maruti / Mahindra ( ... )') -> dict(makes, complement, qualifiers, text)"""
    t = norm(text)
    qualifiers = []
    for p in (norm(x) for x in re.findall(r"\(([^)]*)\)", t)):
        if QUALIFIER_RX.search(p):
            qualifiers.append(p)
        else:
            raise Refuse(f"the bracketed text '({p})' of the make label '{t}' is not understood")
    w = low(re.sub(r"\([^)]*\)", " ", t))
    complement = bool(re.match(r"\s*(?:all\s+)?others?\b", w))
    if complement:
        w = re.sub(r"^\s*(?:all\s+)?others?\b", " ", w)
    keys, rest = split_makes(w)
    if rest or not keys:
        raise Refuse(f"the make label '{t}' holds words that are no make: {', '.join(rest) or '(none)'}")
    return dict(makes=keys, complement=complement, qualifiers=qualifiers, text=t)


def style_only(text):
    """the style (a STYLES entry) a header text is made of entirely ('Kaali Peeli'), or None"""
    t = low(text)
    for style in STYLES:
        if re.fullmatch(r"\W*" + style[1] + r"\W*", t):
            return style
    return None


SCOPE_PART = re.compile(r"all\s+(?:makes?|models?|vehicles?|ages?)")


def scope_only(text):
    """True for a header text that only says 'no restriction' ('All Make/All Ages', 'All Makes'): it applies nothing, and it must not sit beside a make list"""
    parts = [p.strip() for p in re.split(r"\s*(?:/|,|&|\+|\band\b)\s*", low(text)) if p.strip()]
    return bool(parts) and all(SCOPE_PART.fullmatch(p) for p in parts)


def split_clauses(text):
    """the product clauses of a segment title: 'PCV Other bus > 8 STR & PCV 7+1 ( Other makes )' holds two"""
    return [c for c in re.split(r"\s*&\s*(?=(?:pcv|gcv|lcv|misc?d?)\b)", norm(text), flags=re.I) if c.strip()]


def classify_chain(chain, ctx):
    """The specs a column's header chain gives: [dict(product, klass, family, makes, fuels, subs, zd, age, cc, sc, tnc, qualifiers, label, deal, skip, style, guard)] - or Refuse.
    A chain text is a SEGMENT (it names a product), a LEAF (Comp / STP / All / fuels / fresh / an age limit), a MAKE LABEL (makes, 'Others'), a STYLE ('Kaali Peeli') or a SCOPE ('All Make/All
    Ages'); anything else refuses the column.  A header that names its product twice (a group title above a sub-title) is merged when the texts agree and refused when they do not.
    A column of another product (Private Car) is understood and NOT written: its spec carries `skip` (the reason) - nothing is guessed about its words."""
    segs, leaves, labels, styles, scopes, unknown = [], [], [], [], [], []
    for t in chain:
        if kinds_in(t):
            segs.append(t)
        elif style_only(t):
            styles.append(style_only(t))
        elif scope_only(t):
            scopes.append(t)
        elif LEAF_RX.search(low(t)):
            leaves.append(t)
        elif any(re.search(r"\b" + re.escape(w) + r"\b", low(t)) for w in MAKE_WORDS) or re.match(r"\s*others?\b", low(t)):
            labels.append(t)
        else:
            unknown.append(t)
    seg_kinds = sorted({k for t in segs for k in kinds_in(t)})
    if seg_kinds and all(k in NOT_CONVERTED for k in seg_kinds):             # a Private Car / two-wheeler column inside this sheet: its other words are not for this reader
        return [dict(deal=False, skip=f"another product ({NOT_CONVERTED[seg_kinds[0]]}): this build converts the commercial-vehicle products only (the user, 20 Nov 2026); its mapping rules "
                                     f"are not given", other=NOT_CONVERTED[seg_kinds[0]], label=" / ".join(chain), chain=chain, subs=[], seg=None)]
    if unknown:
        raise Refuse(f"no rule reads the header text '{'; '.join(unknown)}'")
    if not segs:
        raise Refuse("no product wording in the header")
    if not leaves:
        raise Refuse(f"no Comp / STP / All wording under '{segs[0]}'")
    leaf = parse_leaf(" ".join(leaves))
    label = parse_make_label(" ".join(labels)) if labels else None
    if len(segs) == 1:
        parsed = [parse_clause(c) for c in split_clauses(segs[0])]
    else:                                                                    # one product named at two levels: merged when the texts agree
        parts = []
        for s in segs:
            cl = split_clauses(s)
            if len(cl) != 1:
                raise Refuse(f"several product texts in one header ({'; '.join(segs)}) and one of them holds more than one product: which belongs to which is not in the document")
            parts.append(parse_clause(cl[0], final=False))
        merged = merge_segs(parts)
        finish_clause(merged)
        parsed = [merged]
    specs = []
    for seg in parsed:
        for style in styles:                                                 # a style text under / over the product ('Kaali Peeli')
            if seg["kind"] not in style[3]:
                raise Refuse(f"'{style[0]}' in the header of '{seg['text']}': this script knows that style for {', '.join(style[3])} only")
            if seg["style"] not in (None, (style[0], style[2])):
                raise Refuse(f"'{seg['text']}' and the header text '{style[0]}' name different styles")
            seg["style"] = (style[0], style[2])
        if seg["deal"]:                                                      # a negotiated deal, not a rate card: whatever the product ('Employee Pick up', 'Car Carrier (Only deal basis)')
            specs.append(dict(deal=True, skip=None, label=seg["text"], chain=chain, subs=leaf["subs"], seg=seg))
            continue
        if seg["kind"] == "pickup":
            raise Refuse(f"'{seg['text']}' is a product this script has no mapping for")
        makes, complement, qualifiers, label_text = list(seg["makes"]), False, list(seg["qualifiers"]), ""
        if label:
            if makes:
                raise Refuse(f"'{seg['text']}' names makes and a make label '{label['text']}' sits above it: which one applies is not in the document")
            makes, complement, qualifiers, label_text = list(label["makes"]), label["complement"], qualifiers + label["qualifiers"], label["text"]
        if scopes and (makes or seg["other_makes"]):
            raise Refuse(f"'{'; '.join(scopes)}' (all makes) and a make list ({', '.join(makes) or 'other makes'}) in one header: which one applies is not in the document")
        key = (seg["family"], seg["sc"])
        if seg["other_makes"]:
            if key not in ctx["makes"]:
                raise Refuse(f"'{seg['text']}' says 'Other makes' but no make list comes before it for the same product")
            makes, complement = list(ctx["makes"][key]), True
        elif makes and not complement:
            ctx["makes"][key] = list(makes)
        if seg["fuel_other"]:                                                # '(Others)' = the fuels other than diesel: the Comp / STP header must list them
            if not leaf["fuels"]:
                raise Refuse(f"'(Others)' in '{seg['text']}' needs the fuels in the Comp / STP header ('Other -(Battery,CNG, Petrol)'): which fuels are meant is not in the document")
            if "Diesel" in leaf["fuels"]:
                raise Refuse(f"'(Others)' in '{seg['text']}' means the fuels other than diesel, but the header lists diesel")
        fuels = leaf["fuels"] or seg["fuels"] or list(FUELS_ALL)
        if leaf["fuels"] and seg["fuels"] and not set(leaf["fuels"]) <= set(seg["fuels"]):
            raise Refuse(f"the fuels of '{' / '.join(chain)}' ({', '.join(leaf['fuels'])}) are not among the fuels of its segment ({', '.join(seg['fuels'])})")
        if leaf["age"] == "fresh":
            age, tnc = (0, 1), ",".join(x for x in (NEW_TNC, seg["tnc_aged"]) if x)
        else:                                                                # the standard age-slab convention (the user, 23 Nov 2026): non fresh = from 1; '< N years' = up to N + 0.01, '> N years' = from N + 0.01
            lo = 1 if leaf["age"] == "nonfresh" else AGE_RANGE[0]
            hi = AGE_RANGE[1]
            lo, hi = (leaf["age_lo"] if leaf["age_lo"] is not None else lo), (leaf["age_hi"] if leaf["age_hi"] is not None else hi)
            if lo >= hi:
                raise Refuse(f"the age band of the column header '{' / '.join(chain)}' is empty ({lo} to {hi})")
            age = (lo, hi)
            tnc = ",".join(x for x in (OLD_TNC, seg["tnc_aged"]) if x) if leaf["age"] == "nonfresh" else seg["tnc"]
        if seg["style"]:                                                     # the style goes into the T&C first ('Kaali Peeli')
            tnc = ",".join(x for x in (seg["style"][1], tnc) if x)
        specs.append(dict(deal=False, skip=None, product=seg["product"], klass=seg["klass"], family=seg["family"], makes=make_group(seg["family"], makes, complement), fuels=fuels,
                          subs=leaf["subs"], zd=seg["zd"] or "NA", age=age, cc=seg["cc"], sc=seg["sc"], tnc=tnc, qualifiers=qualifiers, label=label_text or seg["text"],
                          segment=seg["text"], chain=chain, style=(seg["style"][0] if seg["style"] else None), guard=(seg["klass"] == "MISCD-Others" and tnc == "")))
    return specs


def classify_columns(table):
    """Every column of the table read in order: ([(column, [specs])], [(column, why it was refused)], [columns without a number])"""
    ctx = {"makes": {}}
    good, refused, empty = [], [], []
    for col in table.columns():
        if col["numbers"] == 0:
            empty.append(col)
            continue
        try:
            good.append((col, classify_chain(col["chain"], ctx)))
        except Refuse as e:
            refused.append((col, str(e)))
    return good, refused, empty


# ---- the notes (text blocks of the sheets that hold no table, and the text above / under the rate table)
PAYOUT_WORDS = re.compile(r"payout|\bpo\b|\btc\b|\bvli\b|slab|volume|contest|commission|incentive|enabler|recover|%|\bgrid\b|\bdeal|brokerage|\bgwp\b", re.I)
SCOPE_WORDS = [("school bus", r"school\s*bus|\bsb\b"), ("taxi", r"\btaxi\b|kaali|pili"), ("gcv", r"\bgcv\b|\blcv\b|\bmcv\b|\bhcv\b|goods|pick\s*up"), ("pcv", r"\bpcv\b|passenger"),
               ("miscd", r"\bmisc\w*|tractor|\bjcb\b|flat\s*bed|car\s*carrier|\bcpm\b"), ("not written", r"\btw\b|scooter|two[\s-]*wheel|bike|private\s*car|\bpvt\b|\bhba\b")]


def fmt_cell(v):
    """a cell of a text block as text: a fraction under 1 is shown as a percent (0.015 = 1.5%)"""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        x = float(v)
        if 0 < x < 1:
            return f"{round(x * 100, 4):g}%"
        return f"{x:g}"
    return norm(v)


def text_blocks(sh):
    """[dict(row, text, keys, lines)]: the sheet's rows grouped in blocks of consecutive non-empty rows (a blank row ends a block); a block's text is its rows, cells joined by ' ' and rows
    by '; '; `keys` = the first cell of every row that is a table row (a text followed by a number): the keys of a table such as 'Location | TC %': Mumbai, Delhi, Kolkata"""
    blocks, cur, keys, start = [], [], [], None
    for r in range(1, sh.max_row + 1):
        raws = [v for v in (sh.raw(r, c) for c in range(1, sh.max_col + 1)) if v is not None and norm(v)]
        if raws:
            if start is None:
                start = r
            cur.append(" ".join(fmt_cell(v) for v in raws))
            if num(raws[0]) is None and any(num(v) is not None for v in raws[1:]):
                keys.append(norm(raws[0]))
        elif cur:
            blocks.append(dict(row=start, text="; ".join(cur), keys=keys, lines=cur))
            cur, keys, start = [], [], None
    if cur:
        blocks.append(dict(row=start, text="; ".join(cur), keys=keys, lines=cur))
    return blocks


def split_block(lines, limit=MAX_NOTE_CHARS - 80):
    """a block's rows as texts of at most `limit` characters: the rows are packed in order; every text after the first starts with the block's first row (its title), so a part still says
    what it is about (and its words still scope it)"""
    parts, cur = [], ""
    for ln in lines:
        if cur and len(cur) + len(ln) + 2 > limit:
            parts.append(cur)
            cur = f"{lines[0]} (continued)"
        cur = f"{cur}; {ln}" if cur else ln
    if cur:
        parts.append(cur)
    return parts


def keyed_regions(keys, region_names):
    """the regions a block is keyed by: when it is a table (two rows or more) whose EVERY key is the name of a region of the grid (a table 'Location | TC %': Mumbai, Delhi, Kolkata) -> those
    regions; when one key is no region of the grid it cannot be placed, and the note keeps to the products it names (never fewer rows than it may affect)"""
    low_keys = {k.lower() for k in keys}
    return low_keys if len(keys) >= 2 and low_keys <= region_names else None


# %% ---------------------------------------------------------------- BUILD
LAST_NOT_WRITTEN = []        # what was read but left out of the upload (a column, a sheet): the sheet 'Not written' of the review workbook
LAST_CHANGES = []            # the rows that differ from the earlier upload: the sheet 'Changed vs earlier upload'


def rto_name(region):
    """new_rto_list of a region: the portal's group name RTO_GROUP_PREFIX + the region (group mode only); the CSV carries the region as the grid writes it (apply_cluster_names)"""
    return rto_cluster(RTO_GROUP_PREFIX + re.sub(r"[^A-Za-z0-9]+", "_", region).strip("_"), region)


def _row(rto, spec, sub, fuel, region, col, pi, dates):
    return dict(
        insurance_company=INSURANCE_COMPANY, insurer_vertical="Not Applicable", product=spec["product"], sub_product=sub, policy_type=POLICY_TYPE, fuel_type=fuel,
        make_model_class=spec["klass"], new_vehicle_makes=spec["makes"], new_rto_list=rto, vehicle_age_min=spec["age"][0], vehicle_age_max=spec["age"][1],
        cc_min=spec["cc"][0], cc_max=spec["cc"][1], sc_min=spec["sc"][0], sc_max=spec["sc"][1], pi_od_rate=0.0, pi_tp_2=0, pi_tp_3=0, pi_tp_4=0, pi_tp_5=0, pi_flat_amount=0, pi_vli=0,
        tariff_min=TARIFF[0], tariff_max=TARIFF[1], is_ncb="NA", is_cpa="NA", is_zd=("NA" if sub == "STP" else spec["zd"]), from_date=dates[0], to_date=dates[1], user_id=USER_ID,
        veh_use="", add_tnc=spec["tnc"], remarks="", po_od_rate=0.0, po_flat_amount=0, _pi=pi, _region=region, _col=col, _klass=spec["klass"], _product=spec["product"],
        _guard=bool(spec["guard"]) and pi > 0)


def _short(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text if len(text) <= MAX_NOTE_CHARS else text[:MAX_NOTE_CHARS].rstrip() + "..."


def note_rows(text, region_names=None):
    """The rows a note may affect, as a function of the frame: the products it names (a note that names none, or that names a product nobody wrote, never fewer: every row), and
    only the regions it is keyed by (`keyed_regions`)."""
    t = low(text)
    named = [k for k, rx in SCOPE_WORDS if re.search(rx, t)]
    written = [k for k in named if k != "not written"]

    def mask(d):
        m = np.ones(len(d), dtype=bool) if not written else np.zeros(len(d), dtype=bool)
        for k in written:
            if k == "school bus":
                m |= (d["_klass"] == "School Bus").to_numpy()
            elif k == "taxi":
                m |= (d["_klass"] == "Taxi").to_numpy()
            elif k == "gcv":
                m |= d["_product"].isin(["GCV 4W", "GCV 3W"]).to_numpy()
            elif k == "pcv":
                m |= d["_product"].isin(["PCV 4W", "PCV 3W"]).to_numpy()
            elif k == "miscd":
                m |= (d["_product"] == "MISCD").to_numpy()
        if region_names:
            m &= d["_region"].str.lower().isin(region_names).to_numpy()
        return m
    return mask, named, written


def build(grid_path, from_date, to_date, log=say, ask=None, wb=None):
    """Read the workbook and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).  from_date / to_date are datetime.date values (given on every run:
    the workbook's own title is never read); they are written as DD/MM/YYYY into every row.  `ask` (an Asker) is consulted whenever the grid itself is unclear."""
    global LAST_NOT_WRITTEN, LAST_CHANGES
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    review = ManualReview()                                                  # a new run: the lists of the review workbook start empty
    LAST_NOT_WRITTEN, LAST_CHANGES = [], []
    MAKE_DEFINITIONS.clear()
    REGION_HINTS.clear()
    if wb is None:
        log(f"Reading {os.path.basename(str(grid_path))}...")
        grid_path, wb = open_file(ask, grid_path, "the IndusInd / Reliance workbook", read_workbook, EXCEL_EXTS)
    log(f"Sheets in the workbook: {wb.sheetnames}")
    warnings, hidden = [], hidden_sheets(wb)

    # ---- every visible sheet is a rate table (found by its header) or a sheet of text; a hidden sheet is an old version and is never read
    tables, plain = [], []
    for name in wb.sheetnames:
        if name in hidden:
            continue
        sh = Sheet(wb[name])
        hit = find_anchor(sh)
        t = RateTable(sh, *hit) if hit else None
        if t is not None and t.ok and len(t.data) >= MIN_DATA_ROWS:
            t.good, t.refused, t.empty = classify_columns(t)
            tables.append(t)
        else:
            plain.append(sh)
    mine = [t for t in tables if any(not s["skip"] for _c, specs in t.good for s in specs)]      # a sheet whose every column is another product (Private Car, TW Scooter) is a table of another product
    if not mine:
        found = ", ".join(f"'{t.sh.name}' (header on row {t.anchor}, {len(t.data)} rows" + (f"; no column read, e.g. column {t.refused[0][0]['letter']}: {t.refused[0][1][:140]}" if t.refused else "") + ")"
                          for t in tables) or "no sheet with a region header"
        raise GridError("expected a visible sheet whose header row says 'RTO Region' above payout columns that name a product (School Bus, PCV 3W, PCV Taxi, GCV ..., Car Carrier, "
                        f"Flat Bed, MISD CPM, Tractor) with Comp / STP columns; found: {found}. Visible sheets: {[n for n in wb.sheetnames if n not in hidden]}")
    shape = {id(t): {(s["product"], s["klass"], s["makes"], tuple(s["cc"]), tuple(s["sc"])) for _c, specs in t.good for s in specs if not (s["deal"] or s["skip"])} for t in mine}
    for i, a in enumerate(mine):
        clash = next((b for b in mine[i + 1:] if shape[id(a)] & shape[id(b)]), None)
        if clash is not None:                                                # two sheets that say the same thing: which one is this month's is not in the document
            pick = ask.choose(f"more than one sheet holds the same table: '{a.sh.name}' and '{clash.sh.name}'",
                              [(f"Use the sheet '{a.sh.name}'", a.sh.name), (f"Use the sheet '{clash.sh.name}'", clash.sh.name)],
                              question=f"The sheets '{a.sh.name}' and '{clash.sh.name}' both hold the commercial-vehicle table. Which one is this month's?", short="sheet of the table",
                              memo=("table", a.sh.name, clash.sh.name))
            mine = [t for t in mine if t.sh.name == pick]
            break
    others = [t for t in tables if t not in mine]
    for t in mine:
        t.pick_region(log, ask)

    # ---- the rows: one per region x payout column x policy type x fuel
    rows, flags, guards, blanks, cells_total, note_blocks = [], {}, {}, 0, 0, []
    for t in mine:
        sh = t.sh
        log(f"Rate table: sheet '{sh.name}', header band rows {t.top}-{t.first - 1}, regions in column {col_letter(t.rcol)} ({len(t.data)} rows), payout columns {col_letter(t.c0)}-{col_letter(sh.max_col)}: "
            f"{len(t.good)} read, {len(t.refused)} refused, {len(t.empty)} without a number")
        for text in t.banner:                                                # one text over every column of the table: a note, not a header level
            note_blocks.append((f"{sh.name}: over the table", text, None))
        for col, why in t.refused:
            warnings.append(f"'{sh.name}' column {col['letter']} ('{t.path(col)[:80]}', {col['numbers']} numbers) is NOT in the upload: {why}")
            note_unread(sh.name, f"column {col['letter']}", t.path(col)[:120], col["numbers"], "NOT READ", why)
        for col in t.empty:
            if col["chain"] or col["texts"]:
                warnings.append(f"'{sh.name}' column {col['letter']} ('{t.path(col)[:80]}') has a header or text but no number: not read")
                note_unread(sh.name, f"column {col['letter']}", t.path(col)[:120], col["texts"], "NOT READ", "a header or text with no number under it: nothing to write")
        live = []                                                            # the columns that give rows
        other_products = {}
        for col, specs in t.good:
            for s in specs:
                if s["deal"] or s["skip"]:                                   # a deal, or a column of another product (Private Car): read, understood, NOT written - listed in 'Not written'
                    nz = sum(1 for r in t.data if (num(sh.raw(r, col["c"])) or 0) > 0)
                    why = s["skip"] or f"deal based ('{s['seg']['deal_words'] or 'Only Deal Based'}'): a negotiated deal, not a rate card - the user's export has no rows for it"
                    LAST_NOT_WRITTEN.append(dict(sheet=sh.name, where=f"column {col['letter']}", holds=" / ".join(col["chain"])[:160], cells=f"{len(t.data)} regions ({nz} with a payout)", why=why))
                    if s["skip"]:
                        other_products.setdefault(s["other"], []).append(col["letter"])
                    continue
                for q in s["qualifiers"]:                                    # a model-level condition no make group can say: written for the makes, flagged
                    text = (f"{s['segment']}: '{s['label']}' holds a model-level condition ('{q}') that make groups cannot express - the rows are written for every model of the make group "
                            f"{s['makes']}")
                    flags.setdefault(_short(text), set()).add(col["letter"])
                if s["style"]:                                               # a style (Kaali Peeli) has no field in the upload: written as a T&C, flagged (the user, 23 Nov 2026)
                    text = (f"{s['segment']}: the style '{s['style']}' has no field in the upload - the rows are written as {s['product']} {s['klass']} cards with the T&C '{s['tnc']}'; the ordinary "
                            f"{s['klass']} cards of the same regions also cover these vehicles, so check how the portal tells them apart")
                    flags.setdefault(_short(text), set()).add(col["letter"])
                if s["guard"]:
                    guards[col["letter"]] = s
            if any(not (s["deal"] or s["skip"]) for s in specs):
                live.append((col, [s for s in specs if not (s["deal"] or s["skip"])]))
        for product, letters in other_products.items():
            ask.decide(f"'{sh.name}': column(s) {', '.join(letters)} are {product} columns inside the commercial-vehicle sheet: not converted in this build (listed in 'Not written')")
        hidden_rows = [r for r in t.data if sh.ws.row_dimensions[r].hidden]
        hidden_cols = [col["letter"] for col, _s in live if col["c"] in hidden_columns(sh.ws)]
        if hidden_rows:                                                      # rows / columns hidden in Excel are read like the others: said, never a silent surprise
            warnings.append(f"'{sh.name}': {len(hidden_rows)} row(s) with payouts are hidden in Excel (Excel row {', '.join(str(r) for r in hidden_rows[:8])}{' ...' if len(hidden_rows) > 8 else ''}): "
                            f"they are written like the others - unhide or delete them if they are not meant to count")
        if hidden_cols:
            warnings.append(f"'{sh.name}': column(s) {', '.join(hidden_cols)} with payouts are hidden in Excel: they are written like the others - unhide or delete them if they are not meant to count")
        for r in t.data:
            region = t.name_of(r)
            rto = rto_name(region)
            for col, specs in live:
                v = sh.raw(r, col["c"])
                cells_total += 1
                if v is None or (isinstance(v, str) and not v.strip()):
                    blanks += 1
                kind, x = read_number(ask, f"'{sh.name}'!{col['letter']}{r}", v, "fraction", "a payout fraction like 0.65 (= 65%)", zero_words=ZERO_WORDS, blank_zero=True)
                pi = round(float(x) * 100, RATE_DECIMALS) if kind == "num" else 0.0
                for s in specs:
                    for sub in s["subs"]:
                        for fuel in s["fuels"]:
                            rows.append(_row(rto, s, sub, fuel, region, col["letter"], pi, dates))
        # the text above, beside and under the table
        for r in range(1, t.top):
            for c in range(1, sh.max_col + 1):
                if norm(sh.raw(r, c)):
                    note_blocks.append((f"{sh.name}!{col_letter(c)}{r}", norm(sh.raw(r, c)), None))
        for r in range(t.top, t.first):
            for c in range(1, t.c0):
                if norm(sh.raw(r, c)) and not region_header(sh.raw(r, c)):
                    note_blocks.append((f"{sh.name}!{col_letter(c)}{r}", norm(sh.raw(r, c)), None))
        for r in t.text_rows:
            for c in range(1, sh.max_col + 1):
                if norm(sh.raw(r, c)) and num(sh.raw(r, c)) is None:
                    note_blocks.append((f"{sh.name}!{col_letter(c)}{r}", norm(sh.raw(r, c)), None))
    n_refused = sum(len(t.refused) for t in mine)
    if n_refused:                                                            # a half-read grid must never pass for a whole one (the user, 23 Nov 2026)
        n_cols = sum(len(t.good) + len(t.refused) for t in mine)
        where = "; ".join(f"'{t.sh.name}' column(s) {', '.join(c['letter'] for c, _w in t.refused[:12])}{' ...' if len(t.refused) > 12 else ''}" for t in mine if t.refused)
        warnings.insert(0, f"INCOMPLETE: {n_refused} of {n_cols} payout columns were not read ({where}): the upload does not hold their rows - read the warnings below and the review sheet 'Unread parts'")
    if not rows:
        raise GridError("nothing to write - no payout column of the table could be read")
    if blanks:
        ask.decide(f"{blanks} blank payout cell(s) were written as zero cards (a blank cell is a declined entry)")
    if blanks > 0.1 * max(cells_total, 1):
        warnings.append(f"{blanks:,} of {cells_total:,} payout cells are blank and were written as zero cards: if the workbook holds formulas whose saved values are missing, open it in Excel, "
                        f"save it and run again")
    if any("cng" in low(x) for t in mine for col, _s in t.good for x in col["chain"]):
        ask.decide("CNG in a header means CNG + LPG: the same payout is written for both (as in your export)")

    df = pd.DataFrame(rows)
    df["id"] = ROW_ID
    df["pi_type"] = df["sub_product"].map(PI_TYPES)
    df["po_type"] = df["pi_type"]
    net = (df["sub_product"] == "1+1").to_numpy()
    po = np.where(df["_pi"] > 0, (df["_pi"] - PO_CUT).clip(lower=0), 0.0).round(RATE_DECIMALS)
    df["pi_net_rate"], df["pi_tp_rate"] = np.where(net, df["_pi"], 0.0), np.where(net, 0.0, df["_pi"])
    df["po_net_rate"], df["po_tp_rate"] = np.where(net, po, 0.0), np.where(net, 0.0, po)

    # ---- what the grid says that the rows cannot carry: flagged on the rows it may affect, never guessed
    for text, letters in flags.items():
        review.flag(text, lambda d, letters=frozenset(letters): d["_col"].isin(letters).to_numpy())
    for letter, s in guards.items():
        regions_nz = sorted({x["_region"] for x in rows if x["_col"] == letter and x["_guard"]})
        if regions_nz:                                                       # only where the payout is NOT zero: a zero card pays nothing, whatever it covers
            review.flag(_short(f"{s['segment']} is written as a {s['klass']} card with no T&C, so it would pay EVERY {s['klass']} vehicle (flat bed, JCB ... too): its payout is not zero in "
                               f"{len(regions_nz)} region(s): {', '.join(regions_nz[:12])}{' ...' if len(regions_nz) > 12 else ''}"),
                        lambda d, letter=letter: ((d["_col"] == letter) & d["_guard"]).to_numpy())
    region_names = {t.name_of(r).lower() for t in mine for r in t.data}
    rto_list = None
    for sh in plain:                                                         # sheets of text: Guidelines, Contest ...
        listing = read_rto_list(sh)
        if listing:                                                          # the workbook's own RTO codes per region: 'did you mean' hints and the codes of 'RTO groups to create' only
            rto_list = listing
            n_codes = sum(len(v) for v in listing.values())
            note_unread(sh.name, "the whole sheet", f"{n_codes:,} RTO codes under {len(listing)} regions", n_codes, "left out on purpose",
                        "the workbook's RTO codes per region: used only for 'did you mean' hints and the RTO codes in 'RTO groups to create'; the names written are the grid's regions (the user's cluster rule)")
            log(f"RTO List: sheet '{sh.name}' holds {n_codes:,} RTO codes under {len(listing)} regions (used for hints and the codes of the RTO groups to create)")
            continue
        blocks = text_blocks(sh)
        if not any(PAYOUT_WORDS.search(b["text"]) for b in blocks):
            unread_sheet(sh.name, sh.rows(), "NOT READ", "no table and no sentence about payouts, slabs, contests or deals: nothing for this script to read", warnings)
            continue
        for b in blocks:
            keyed = keyed_regions(b["keys"], region_names)
            for part in split_block(b["lines"]):
                note_blocks.append((f"{sh.name}!row {b['row']}", part, keyed))
    for t in others:                                                         # the tables of the other products: not written in this build
        top, notes = describe_rows(t.sh.rows())
        LAST_NOT_WRITTEN.append(dict(sheet=t.sh.name, where="the whole sheet", holds="; ".join(top[:8] + [f"text: {n}" for n in notes[:2]])[:300],
                                     cells=f"{len(t.data)} regions x {len(t.columns())} columns",
                                     why="a table of another product: this build converts the commercial-vehicle sheet only (the user, 20 Nov 2026); its mapping rules are not given yet"))
        ask.decide(f"'{t.sh.name}': a table of another product - not converted in this build")
        for r in range(1, t.sh.max_row + 1):                                 # its sentences: listed in 'Notes read' (not applied), so the rules of a product to come are not lost
            for c in range(1, t.sh.max_col + 1):
                x = norm(t.sh.raw(r, c))
                if len(x) >= NOTE_MIN_CHARS and num(t.sh.raw(r, c)) is None:
                    NOTES_SEEN.setdefault(note_key(x), dict(text=x[:400], status="not applied (the sheet is not converted in this build)",
                                                            why=f"text of the sheet '{t.sh.name}' ({col_letter(c)}{r}): the rule is for a product that is not written yet"))
    flagged = set(unknown_notes([txt for _a, txt, _k in note_blocks], STANDING_NOTES))
    for addr, txt, keyed in note_blocks:
        if re.sub(r"\s+", " ", txt).strip() not in flagged:
            continue
        mask, named, written = note_rows(txt, keyed)
        if named and not written:                                            # a text about two-wheelers / Private Car only: no such row is written in this build
            NOTES_SEEN[note_key(txt)] = dict(text=txt[:400], status="no rows affected", why=f"it names {', '.join(named)} only: no row of that product is written in this build ({addr})")
            continue
        review.flag(_short(txt), mask)
    if LAST_NOT_WRITTEN:
        add_sheet("Not written", pd.DataFrame(LAST_NOT_WRITTEN))

    masters = load_masters(ask)
    check_regions(df, {t.name_of(r): (t.parent_of(r), t.zone_of(r)) for t in mine for r in t.data}, masters, review, warnings, log, ask, rto_list)
    check_names(df, review, masters, log)                                    # every make group written (and, in group mode, every RTO group) must be in the master lists
    fill_definitions()
    review.report(df, log)                                                   # fills the remarks column
    return df[OUTPUT_COLUMNS], warnings


RTO_LIST_MIN_CODES = 20                          # a sheet that lists at least this many RTO codes under a region is the workbook's RTO list
HINT_MIN_RATIO = 0.45                            # how alike two region names must be for a 'did you mean' hint, and by how much the closest free name must beat the next one ...
HINT_MARGIN = 0.15
HINT_ELIMINATION_RATIO = 0.3                     # ... or, when exactly one name is left on each side, how alike the pair left over must be


def read_rto_list(sh):
    """{region: [RTO codes]} of a sheet that lists the RTO codes under a region (a header row in its first rows with a code column and a city / cluster / region column - October 2026:
    'Region_Code | Column1 | Region_City | Market_Type | Zone'); None when the sheet is not one"""
    for r in range(1, min(sh.max_row, 10) + 1):
        heads = {c: _hwords(sh.raw(r, c)) for c in range(1, sh.max_col + 1) if isinstance(sh.raw(r, c), str)}
        code_c = next((c for c, w in heads.items() if "code" in w), None)
        city_c = next((c for c, w in heads.items() if c != code_c and ("city" in w or "cluster" in w or ("region" in w and "code" not in w))), None)
        if code_c and city_c:
            out, spelled = {}, {}                                            # a region spelt in capitals in one row and in title case in another is one region
            for rr in range(r + 1, sh.max_row + 1):
                code, city = norm(sh.raw(rr, code_c)), norm(sh.raw(rr, city_c))
                if code and city and num(sh.raw(rr, city_c)) is None:
                    out.setdefault(spelled.setdefault(city.upper(), city), []).append(code)
            return out if sum(len(v) for v in out.values()) >= RTO_LIST_MIN_CODES else None
    return None


def alike(a, b):
    """how alike two region names are: the best ratio of the whole names and of their words ('LEH & LAHDHK' has the word LAHDHK, which is nearly LADAKH); words of one or two letters do not count"""
    from difflib import SequenceMatcher
    a, b = a.upper(), b.upper()
    wa, wb = re.findall(r"[A-Z0-9]+", a), re.findall(r"[A-Z0-9]+", b)
    return max([SequenceMatcher(None, a, b).ratio()] + [SequenceMatcher(None, x, y).ratio() for x in wa for y in wb if len(x) > 2 and len(y) > 2])


def region_hints(regions, listing):
    """{region of the grid: the RTO List's own spelling} for the grid regions that are no region of the RTO List ('LEH & LAHDHK' = 'Ladakh', 'TRIVANDRUM' = 'Thiruvananthapuram'): matched to the
    regions of the RTO List that no grid region uses, by how alike the names are, and by elimination (a lower bar) when one is left on each side.  A HINT for the person - the names written stay
    the grid's."""
    known = {norm(k).upper(): k for k in listing}
    mine = {norm(r).upper(): r for r in regions}
    lonely, spare = [r for u, r in mine.items() if u not in known], [k for u, k in known.items() if u not in mine]
    best = {}                                                                # the closest free name of the list for each lonely region, when it is clearly the closest
    for r in lonely:
        cands = sorted(((alike(r, s), s) for s in spare), reverse=True)
        if cands and cands[0][0] >= HINT_MIN_RATIO and (len(cands) == 1 or cands[0][0] - cands[1][0] >= HINT_MARGIN):
            best[r] = cands[0]
    out, used = {}, set()
    for r, (ratio, s) in sorted(best.items(), key=lambda kv: -kv[1][0]):
        if s not in used:
            out[r] = s
            used.add(s)
    left_r, left_s = [r for r in lonely if r not in out], [s for s in spare if s not in used]
    if len(left_r) == 1 and len(left_s) == 1 and alike(left_r[0], left_s[0]) >= HINT_ELIMINATION_RATIO:        # one left on each side: the pair by elimination
        out[left_r[0]] = left_s[0]
    return out


def check_regions(df, parents, masters, review, warnings, log, ask=None, rto_list=None):
    """The user's decision (20 Nov 2026): a region with no RTO group in the master is written all the same and flagged.  The group is looked up by the name the portal's groups have
    (RTO_GROUP_PREFIX + the region, spaces -> '_', capitals ignored); the name WRITTEN stays the bare region.  In group mode check_names does this itself.  When the workbook has an RTO list
    (`rto_list`: region -> RTO codes) a region it spells differently gets a 'did you mean' hint and the sheet 'RTO groups to create' carries the RTO codes (the user, 23 Nov 2026)."""
    if RTO_NAMES != "cluster":
        return
    known = {k.lower(): k for k in masters.rto}
    regions = sorted(df["_region"].unique())
    hints = region_hints(regions, rto_list) if rto_list else {}
    REGION_HINTS.clear()
    REGION_HINTS.update({re.sub(r"[^a-z0-9]+", "_", k.lower()).strip("_"): re.sub(r"[^a-z0-9]+", "_", v.lower()).strip("_") for k, v in hints.items()})

    def group_of(reg):
        name = RTO_GROUP_ALIASES.get(reg.lower(), reg)
        return known.get((RTO_GROUP_PREFIX + re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")).lower())
    have = {reg: group_of(reg) for reg in regions}
    share = sum(1 for g in have.values() if g) / max(len(regions), 1)
    if share < RTO_GROUP_MIN_SHARE:
        warnings.append(f"only {sum(1 for g in have.values() if g)} of {len(regions)} regions have an RTO group named {RTO_GROUP_PREFIX}<region> in the master: the prefix is probably out of date, "
                        f"so the check for regions with no RTO group was skipped (RTO_GROUP_PREFIX in the settings)")
        return
    missing = [reg for reg in regions if not have[reg]]
    spelled = {norm(k).upper(): k for k in (rto_list or {})}
    if hints and ask is not None:
        ask.decide("regions the workbook's RTO List spells differently (the rows keep the grid's spelling): " + ", ".join(f"{k} = {v}" for k, v in sorted(hints.items())))
    rows = []
    for reg in missing:
        n = int((df["_region"] == reg).sum())
        parent, zone = parents.get(reg, ("", ""))
        want = RTO_GROUP_PREFIX + re.sub(r"[^A-Za-z0-9]+", "_", reg).strip("_")
        hint = hints.get(reg)
        codes = rto_list[spelled[norm(hint or reg).upper()]] if rto_list and norm(hint or reg).upper() in spelled else []
        mean = ""
        if hint:                                                             # the RTO List spells the region another way: a hint, never applied
            theirs = group_of(hint)
            mean = f"; did you mean '{hint}' - the workbook's RTO List spells it so" + (f", and the master has the group {theirs}" if theirs else "") + "?"
        review.flag(f"no RTO group in the master for the region '{reg}' (looked for {want}{mean}): create the group in the portal - the rows carry the region name the grid writes",
                    lambda d, reg=reg: (d["_region"] == reg).to_numpy())
        row = dict(group_needed=want, rows_flagged=n,
                   items=((f"the grid lists no RTO codes; it puts '{reg}' under '{parent}'" if parent and parent.lower() != reg.lower() else "the grid lists no RTO codes")
                          + (f" (zone {zone})" if zone else "") if not codes else
                          (f"RTO codes from the workbook's RTO List (under '{hint or reg}'); the grid puts '{reg}' under '{parent}'" if parent and parent.lower() != reg.lower()
                           else f"RTO codes from the workbook's RTO List (under '{hint or reg}')") + (f" (zone {zone})" if zone else "")))
        if rto_list:
            row["rto_codes"] = ", ".join(codes)
        rows.append(row)
    add_sheet("RTO groups to create", pd.DataFrame(rows))
    log(f"Regions: {len(regions) - len(missing)} of {len(regions)} have an RTO group in the master" + (f"; no group for {', '.join(missing)} (flagged)" if missing else ""))


def fill_definitions():
    """the sheet 'Make groups to create' says how each missing group is made (MAKE_DEFINITIONS: the family group except the makes)"""
    sheet = EXTRA_SHEETS.get("Make groups to create")
    if sheet is None or not len(sheet):
        return
    sheet = sheet.copy()
    sheet["items"] = [(f"= {MAKE_DEFINITIONS[g][0]} except {'|'.join(MAKE_DEFINITIONS[g][1])}" if g in MAKE_DEFINITIONS else it) for g, it in zip(sheet["group_needed"], sheet["items"])]
    EXTRA_SHEETS["Make groups to create"] = sheet


# %% ---------------------------------------------------------------- OUTPUT AND CHECKS
def fmt(x):
    x = round(float(x), 4)
    return str(int(x)) if x == int(x) else repr(x)


def to_csv_frame(df):
    out = df.copy()
    for c in NUMERIC_COLUMNS:
        out[c] = out[c].map(fmt)
    return out


REGION_HINTS = {}        # the regions of THIS grid the workbook's RTO List spells differently, as folded keys ({'leh_lahdhk': 'ladakh'}): used only to pair a region with the earlier upload's spelling


def rto_key(name):
    """a region as the earlier upload spells it (RELIANCE_SEPT26_Mumbai, capitals, underscores) and as the grid spells it, folded to one key (and, when the workbook's RTO List spells the
    region another way, that spelling)"""
    t = re.sub(r"^reliance_[a-z0-9]+_", "", str(name).strip().lower())
    t = re.sub(r"[^a-z0-9]+", "_", t).strip("_")
    back = {re.sub(r"[^a-z0-9]+", "_", v.lower()).strip("_"): k for k, v in RTO_GROUP_ALIASES.items()}
    t = back.get(t, t)
    return REGION_HINTS.get(t, t)


def make_key(name):
    t = str(name).strip().lower()
    return REF_MAKE_ALIASES.get(t, t)


def key_of(d, cols=None):
    """the comparison key of a frame of upload rows: numbers to 4 decimals, regions and make names folded (the user's earlier names for the same groups are the same key)"""
    d = d.copy()
    for c in ("vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "sc_min", "sc_max", "tariff_min", "tariff_max"):
        d[c] = pd.to_numeric(d[c], errors="coerce").map(lambda x: f"{x:.4f}")
    d["new_rto_list"] = d["new_rto_list"].astype(str).map(rto_key)
    d["new_vehicle_makes"] = d["new_vehicle_makes"].astype(str).map(make_key)
    return d[cols or KEY_COLUMNS].astype(str).agg("|".join, axis=1)


def pair_leftovers(mine, ref1, new_keys, ref_keys):
    """The rows no key matched, paired by the CARD (every key column but the make group and the T&C): the earlier upload named the make group or wrote the T&C another way.  A group of equal
    cards is paired only when both sides hold the same number of rows, in the order of their payins.  Returns (this grid's paired rows, the earlier upload's paired rows), row for row."""
    soft = [c for c in KEY_COLUMNS if c not in ("new_vehicle_makes", "add_tnc")]
    nl = mine[mine["_k"].isin(new_keys)].copy().reset_index(drop=True)
    rl = ref1[ref1["_k"].isin(ref_keys)].copy().reset_index(drop=True)
    if not len(nl) or not len(rl):
        return nl.iloc[0:0], rl.iloc[0:0]
    for d in (nl, rl):
        d["_s"] = key_of(d, soft)
        net = (d["sub_product"] == "1+1").to_numpy()
        d["_pi"] = pd.to_numeric(pd.Series(np.where(net, d["pi_net_rate"], d["pi_tp_rate"]), index=d.index), errors="coerce").fillna(0.0)
        d["_po"] = pd.to_numeric(pd.Series(np.where(net, d["po_net_rate"], d["po_tp_rate"]), index=d.index), errors="coerce").fillna(0.0)
    rg = rl.groupby("_s").indices
    take_a, take_b = [], []

    def order(d, idx):
        return sorted(idx, key=lambda i: (round(float(d["_pi"].iat[i]), 4), round(float(d["_po"].iat[i]), 4), make_key(d["new_vehicle_makes"].iat[i]), str(d["add_tnc"].iat[i])))
    for s_key, ia in nl.groupby("_s").indices.items():
        ib = rg.get(s_key)
        if ib is None or len(ib) != len(ia):
            continue
        take_a += order(nl, ia)
        take_b += order(rl, ib)
    return nl.iloc[take_a].reset_index(drop=True), rl.iloc[take_b].reset_index(drop=True)


def diff_against_reference(df, ref_path, log=say):
    """Compare with the user's earlier upload (an export of the portal).  The regions and make groups are folded to one key (RELIANCE_SEPT26_Mumbai = MUMBAI; the user's two taxi make names =
    the master's groups they chose on 20 Nov 2026; a region the workbook's RTO List spells another way = its spelling).  The rows no key matches are then paired by the CARD (everything but the
    make group and the T&C: the earlier upload may have named a make group or written a T&C another way) and counted by what differs (the sheet 'Other names vs earlier upload').
    Every row whose payin or payout differs is LISTED, here and in the sheet 'Changed vs earlier upload'."""
    global LAST_CHANGES
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    mine = to_csv_frame(df)
    missing = [c for c in KEY_COLUMNS if c not in ref.columns]
    if missing:
        raise ValueError(f"the earlier upload has no column(s) {missing} - it does not look like an upload file")
    mine["_k"], ref["_k"] = key_of(mine), key_of(ref)
    if ref["_k"].duplicated().any():
        log(f"  note: reference has {int(ref['_k'].duplicated().sum())} duplicate keys")
    ref1 = ref.drop_duplicates("_k")
    merged = mine.merge(ref1, on="_k", how="outer", suffixes=("", "_ref"), indicator=True)
    only_new = merged[merged["_merge"] == "left_only"]
    only_ref = merged[merged["_merge"] == "right_only"]
    both = merged[merged["_merge"] == "both"]
    pa, pb = pair_leftovers(mine, ref1, set(only_new["_k"]), set(only_ref["_k"]))          # the rows only the make group / the T&C tells apart
    n_pair = len(pa)
    cols = [c for c in OUTPUT_COLUMNS if c not in ("id", "remarks") and c in ref.columns and c + "_ref" in both.columns]

    def differs(c, x, y):
        if c in NUMERIC_COLUMNS + ["user_id"]:
            return (pd.to_numeric(x, errors="coerce") - pd.to_numeric(y, errors="coerce")).abs() > 1e-6
        if c in ("from_date", "to_date"):
            return x.astype(str).str.replace("-", "/") != y.astype(str).str.replace("-", "/")
        return x.astype(str) != y.astype(str)
    changed, changed2 = pd.Series(False, index=both.index), pd.Series(False, index=pa.index)
    detail, per_col, per_col2 = {}, {}, {}
    for c in cols:
        if c in KEY_COLUMNS and c not in NUMERIC_COLUMNS:
            continue                                   # part of the key: equal by construction
        d = differs(c, both[c], both[c + "_ref"])
        d2 = differs(c, pa[c], pb[c]) if n_pair else pd.Series(False, index=pa.index)
        if d.any() or d2.any():
            detail[c] = int(d.sum() + d2.sum())
            per_col[c], per_col2[c] = d, d2
        changed |= d
        changed2 |= d2
    n_new, n_removed, n_changed = len(only_new) - n_pair, len(only_ref) - n_pair, int(changed.sum() + changed2.sum())
    log(f"  vs {os.path.basename(ref_path)}: {n_new} new rows, {n_removed} removed rows, "
        f"{n_changed} rows with changed values (of {len(both) + n_pair} matching keys)")
    for c, n in detail.items():
        ex = both[per_col[c]][["_k", c, c + "_ref"]].head(3).values.tolist() or [[pa["_k"].iat[i], pa[c].iat[i], pb[c].iat[i]] for i in list(per_col2[c][per_col2[c]].index[:3])]
        log(f"    column {c}: {n} differ, e.g. {ex}")
    rows = []
    for i in both.index[changed.to_numpy()]:
        r = both.loc[i]
        diff = {c: (r[c], r[c + "_ref"]) for c, d in per_col.items() if bool(d.loc[i]) and c not in ("from_date", "to_date")}
        rows.append(dict(region=CLUSTER_OF.get(r["new_rto_list"], r["new_rto_list"]), product=r["product"], sub_product=r["sub_product"], fuel=r["fuel_type"], make_model_class=r["make_model_class"], makes=r["new_vehicle_makes"],
                         is_zd=r["is_zd"], columns=", ".join(diff), this_grid=" | ".join(str(v[0]) for v in diff.values()), earlier_upload=" | ".join(str(v[1]) for v in diff.values())))
    for i in pa.index[changed2.to_numpy()]:
        diff = {c: (pa[c].iat[i], pb[c].iat[i]) for c, d in per_col2.items() if bool(d.iat[i]) and c not in ("from_date", "to_date")}
        rows.append(dict(region=CLUSTER_OF.get(pa["new_rto_list"].iat[i], pa["new_rto_list"].iat[i]), product=pa["product"].iat[i], sub_product=pa["sub_product"].iat[i], fuel=pa["fuel_type"].iat[i],
                         make_model_class=pa["make_model_class"].iat[i], makes=pa["new_vehicle_makes"].iat[i], is_zd=pa["is_zd"].iat[i], columns=", ".join(diff),
                         this_grid=" | ".join(str(v[0]) for v in diff.values()), earlier_upload=" | ".join(str(v[1]) for v in diff.values())))
    LAST_CHANGES = rows
    if rows:
        add_sheet("Changed vs earlier upload", pd.DataFrame(rows[:3000]))
        log(f"    the {len(rows)} rows whose value differs from the earlier upload (the grid wins; listed in the review workbook, sheet 'Changed vs earlier upload'):")
        for r in rows[:60]:
            log(f"      {r['region']:<16} {r['product']} {r['sub_product']:<3} {r['fuel']:<8} {r['make_model_class']:<12} {r['columns']}: grid {r['this_grid']} | earlier upload {r['earlier_upload']}")
        if len(rows) > 60:
            log(f"      ... and {len(rows) - 60} more (all in the review workbook)")
    if n_pair:                                                               # the same card, named another way: counted by what differs
        names = {}
        for i in range(n_pair):
            ma, mb = make_key(pa["new_vehicle_makes"].iat[i]), make_key(pb["new_vehicle_makes"].iat[i])
            ta, tb = str(pa["add_tnc"].iat[i]).strip(), str(pb["add_tnc"].iat[i]).strip()
            if ma != mb or ta != tb:
                key = (pa["product"].iat[i], pa["make_model_class"].iat[i], pa["new_vehicle_makes"].iat[i], pb["new_vehicle_makes"].iat[i], ta, tb)
                names[key] = names.get(key, 0) + 1
        if names:
            add_sheet("Other names vs earlier upload", pd.DataFrame([dict(rows=n, product=k[0], make_model_class=k[1], makes_this_grid=k[2], makes_earlier_upload=k[3], tnc_this_grid=k[4], tnc_earlier_upload=k[5])
                                                                     for k, n in sorted(names.items(), key=lambda kv: -kv[1])]))
            log(f"    {sum(names.values())} rows pair with a row of the earlier upload that is the same card but names the make group and / or the T&C another way (this grid's names are written; "
                f"listed in the review sheet 'Other names vs earlier upload'):")
            for k, n in sorted(names.items(), key=lambda kv: -kv[1])[:14]:
                bits = ([f"make group {k[2]} / earlier {k[3]}"] if make_key(k[2]) != make_key(k[3]) else []) + ([f"T&C '{k[4]}' / earlier '{k[5]}'"] if k[4] != k[5] else [])
                log(f"      {n:>5} x {k[0]} {k[1]}: " + "; ".join(bits))
    only_ref_left = only_ref[~only_ref["_k"].isin(set(pb["_k"]))] if n_pair else only_ref
    only_new_left = only_new[~only_new["_k"].isin(set(pa["_k"]))] if n_pair else only_new
    if len(only_ref_left):
        by = only_ref_left.groupby(["product_ref" if "product_ref" in only_ref_left else "product", "sub_product_ref" if "sub_product_ref" in only_ref_left else "sub_product"]).size()
        log("    removed rows are: " + "; ".join(f"{n} x {p} {s}" for (p, s), n in by.head(8).items()) + "  (rows the earlier upload has and this grid does not)")
    if len(only_new_left):
        by = only_new_left.groupby(["product", "sub_product"]).size()
        regions = sorted({CLUSTER_OF.get(x, x) for x in only_new_left["new_rto_list"].astype(str).unique()})
        log("    new rows are: " + "; ".join(f"{n} x {p} {s}" for (p, s), n in by.head(8).items()) + f"  (regions: {', '.join(regions[:6])}{' ...' if len(regions) > 6 else ''})")
    theirs = {rto_key(t) for x in ref["new_rto_list"] for t in str(x).split(",") if t.strip()}          # a list of groups (the earlier Private Car rows) is read name by name
    ours = {rto_key(t) for x in mine["new_rto_list"] for t in str(x).split(",") if t.strip()}
    cl_new, cl_old = sorted(ours - theirs), sorted(theirs - ours)
    if cl_new:
        log(f"    regions not in the earlier upload ({len(cl_new)}): {cl_new[:6]}{' ...' if len(cl_new) > 6 else ''}")
    if cl_old:
        log(f"    regions in the earlier upload but not in this grid ({len(cl_old)}): {cl_old[:6]}{' ...' if len(cl_old) > 6 else ''}")
    return n_new, n_removed, n_changed


def unique_warnings(warnings):
    """each warning once, with how many times it came up"""
    seen, order = {}, []
    for w in warnings:
        if w not in seen:
            order.append(w)
        seen[w] = seen.get(w, 0) + 1
    return [w if seen[w] == 1 else f"{w} (x{seen[w]})" for w in order]


def report(df, warnings, log=say, ask=None):
    """Print what is about to be written; rows that repeat a key keep the lowest payin (resolve_duplicates). Returns the frame to write."""
    ask = ask or Asker()
    df = resolve_duplicates(df, KEY_COLUMNS, ask)
    log(f"\nRows to write: {len(df):,}   ({df['new_rto_list'].nunique()} regions)")
    counts = df.groupby(["product", "make_model_class", "sub_product"]).size()
    for (prod, klass, sub), n in counts.items():
        log(f"  {prod:<8} {klass:<28} {sub:<4} {n:>7,}")
    log(f"Rows that repeat a key (two columns, two payouts: only the lowest is written): {len(LAST_LOWEST)} entries")
    pi = df[["pi_od_rate", "pi_tp_rate", "pi_net_rate"]].max(axis=1)
    log(f"Zero cards: {int((pi == 0).sum()):,}   PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {int(((pi > 0) & (df[['po_od_rate', 'po_tp_rate', 'po_net_rate']].max(axis=1) == 0)).sum())}")
    for w in unique_warnings(warnings):
        log(f"WARNING: {w}")
        RUN_WARNINGS.append(w)
    return df


def main(grid_path, from_date, to_date, out_path=None, ref_path=None, log=say, ask=None):
    ask = ask or Asker()
    df, warnings = build(grid_path, from_date, to_date, log, ask)
    df = report(df, warnings, log, ask)
    diff = compare_safely(ask, diff_against_reference, df, ref_path, log)
    ask.summary(log)
    if ask.interactive:
        notes = []
        if warnings:
            notes.append("some columns or sheets are not in the output (see WARNING above)")
        if diff and (diff[0] + diff[1]) > 0.25 * len(df):
            notes.append(f"{diff[0]} rows are new and {diff[1]} rows are gone compared with the earlier upload")
        question = ("Heads-up: " + "; ".join(notes) + ". Write the file anyway?" if notes
                    else "Everything look right? Write the file now?")
        if not ask.confirm(question, True):
            raise GridError("stopped by you - nothing was written")
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"INDUSIND_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
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
            low_name = name.lower()
            (csvs if low_name.endswith(".csv") else books if low_name.endswith(EXCEL_EXTS) else ignored).append(name)
        if ignored:
            say(f"Ignored (neither a workbook nor a CSV): {ignored}")
        if books:
            break
        if not ask.confirm("No workbook (.xlsx) came with that upload. Upload again?", True):
            raise GridError("no workbook was uploaded")
    grid = books[0] if len(books) == 1 else ask.choose(
        "more than one workbook was uploaded", [(n, n) for n in books],
        question="You uploaded more than one workbook. Which one is this month's IndusInd / Reliance grid?", short="workbook")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the IndusInd / Reliance grid (.xlsx) and, optionally, an earlier upload (.csv) to compare with")
    try:
        grid, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("indusind_grid_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    ap = argparse.ArgumentParser(description="IndusInd General (Reliance) commercial-vehicle grid -> upload CSV")
    ap.add_argument("--grid", help="the workbook (.xlsx) - asked for if left out")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    try:
        grid = a.grid or ask_file(ask, "no workbook given - pass --grid <workbook.xlsx>",
                                  "Path of the IndusInd / Reliance workbook (.xlsx) - you can drag the file into this window:", exts=EXCEL_EXTS)
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
        saved = crash_report("indusind_grid_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
