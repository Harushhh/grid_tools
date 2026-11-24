"""
Royal Sundaram (RSA) Private Car grid -> RSA upload CSV
=======================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (58).csv") from the RSA Private Car workbook
circulated every month ("RSA - Pvt Car Grid Agency RB Oct26 V1.xlsx").  Rates come only from that workbook.

What the workbook holds and how it becomes rows (checked against the October upload: every one of its 17,430 rows matches;
the October upload had no New rows for other makes, they are written on top of it)
---------------------------------------------------------------------------------------------------------------------------
* One sheet per region (West, North, East, South).  Down the side: Comp (1+1), SAOD and New car, each by engine size
  (Upto 1000 / 1000-1500 / >1500 cc) and, for Comp and SAOD, With NCB / Without NCB.  Across: the clusters.  A cell is a
  payout as a fraction (0.26 = 26%).
    Comp   -> sub_product 1+1,  policy_type Renew, is_ncb YES / NO
    SAOD   -> sub_product SAOD, policy_type Renew, is_ncb YES / NO
    New    -> sub_product 1+3,  policy_type New,   is_ncb NA
  Every row is written for each of the five fuels (Petrol, Diesel, CNG, LPG, Electric) with the same rate, age 0-99.
* Two make groups.  The "Top End Variant List" sheet names the top variants of Maruti, Tata, Mahindra, Hyundai, Toyota, Kia,
  MG and Skoda; the portal keeps them as one make group (MAKES_TOP).  The notes under each sheet say what that group gets:
    Comp and SAOD:  the make group gets the cell PLUS the "Additional N% payout for the Top Variant Models" of the note
                    (when the cell is above zero) - or, in the September layout, the base-payout extra of the "Top Variant
                    Models Drive" block, which differs with / without NCB ("Base payout  4% additional" / "2% additional");
                    every other make (MAKES_OTHER) gets the cell as it stands.  The business-share steps of that block
                    ("25% business share ... Base + 1% additional", "50% ... Base + 2% additional") are NOT in the upload
    New car:        MAKES_TOP gets the cell as it stands (the note says the New car payout is for those makes); every other
                    make (MAKES_OTHER) gets the flat rate of the note "Any other make New car payout will be N% flat"
                    (OTHER_MAKES_NEW)
* The "Pvt Car STP" sheet: third-party only, by RTO division and engine size, one set of columns for Diesel and one for
  Non Diesel (Petrol, CNG, LPG, Electric).  Rows: sub_product STP, make group Private_Car_ALL, pi_type / po_type On TP.
* Payout (po) is the payin (pi) minus 7 points, never below 0 (zero stays zero).  Rates are rounded to 2 decimals from the
  full-precision cell (0.3312502 -> 33.13, 0.0972496 -> 9.72, as in the September upload); a true half rounds up.
* new_rto_list = RTO_PREFIX + "PVT_CAR_" + cluster (STP: "PVT_STP_CAR_" + division).  A name is the grid's name in capitals
  with spaces and commas turned into underscores and any "(...)" dropped - except where the portal's group is spelt
  differently (CLUSTER_NAME_EXCEPTIONS: Baroda -> VADODARA, Delhi NCR -> Delhi_NCR, Mysore, Vijayawada ...).

Notes are READ, not just skipped
--------------------------------
The text under each sheet is run through patterns that pick out the rules this script applies ("Additional 2% payout for the
Top Variant Models of ...", "Any other make New car payout will be 28% flat", the makes the New car payout is for).  Text
about volume-linked payouts (VLI) and the other notes in STANDING_NOTES is known not to change the rows of this upload.  Everything
else that is a rule-sized piece of text is NEVER guessed at and NEVER stops the run: the rows it may affect carry
Manual Review Required: [the note text]  in the `remarks` column (the column already in the upload format), several notes joined
with " | ".  Filter that column, take the notes up with Royal and map those rates by hand.  Which rows: the clusters the note
names (a state brings in all its clusters), or the whole sheet when it names none.  Also flagged:
  * a rule the notes do not give (no "Additional N%" note -> the top-variant rows get no extra)
  * the makes a note names differ from the make group (the group in the portal needs checking)
  * "Any other make New car payout will be 28% flat": New rows are written for the other makes (MAKES_OTHER, every engine size,
    flat rate from the note) - the October upload had none, you asked for them.  If that note is missing or unreadable in a sheet,
    no such rows are written for it and its New rows are flagged.  Set OTHER_MAKES_NEW = "flag" for the October behaviour.

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a
sheet or header that is not where it should be, an engine-size band it cannot read, a business type or NCB label it does not
know, a cell that is not a rate, dates that look wrong, a file that will not open or write - it stops, asks you a specific
question (type the number of an option, or q to stop, or ? for the details), remembers the answer and carries on.  Every
answer is listed at the end of the run, and you confirm before the file is written.  It never asks you for a rate.  With
--no-prompt (or when there is no keyboard) it asks nothing and stops with the question as the error message instead.  If
something goes wrong that it did not foresee it says so in one line and saves the technical details to a *_error_*.txt file.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run rsa_pvtcar_to_config.py`) and run it.
Upload the RSA workbook (and, optionally, an earlier upload CSV to get a diff report).  It then asks
for from_date and to_date (DD/MM/YYYY), writes them into every row, and downloads RSA_PVTCAR_UPLOAD_<yyyy-mm>.csv.

Command line
------------
    python rsa_pvtcar_to_config.py [--grid <workbook.xlsx>] [--from-date 01/10/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--prefix ROYAL_SEPT26_] [--out file.csv] [--no-prompt]
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
INSURANCE_COMPANY = "ROYAL SUNDARAM GENERAL INSURANCE COMPANY LIMITED"
PRODUCT = "Private Car"
FUELS = ["Petrol", "Diesel", "CNG", "LPG", "Electric"]   # the same rate goes to each fuel (the grid has one set of columns)
# new_rto_list = RTO_PREFIX + COMP_RTO_PART + cluster (Comp / SAOD / New) or RTO_PREFIX + STP_RTO_PART + division (STP).
# The portal's RTO groups were created with the September prefix and are still used; change it here (or pass --prefix) if
# new groups are ever created.
RTO_PREFIX = "ROYAL_SEPT26_"
COMP_RTO_PART = "PVT_CAR_"
STP_RTO_PART = "PVT_STP_CAR_"
# the grid's name (without any "(...)") -> the part after the prefix, where the portal's group is not just "capitals, spaces -> underscores"
CLUSTER_NAME_EXCEPTIONS = {
    "Baroda": "VADODARA", "Delhi NCR": "Delhi_NCR", "DELHI": "Delhi_NCR", "Bhuvaneshwar": "BHUBANESWAR",
    "ANDAMAN and NICOBAR": "ANDAMAN_and_NICOBAR", "Mysore": "Mysore", "Vijayawada": "Vijayawada",
    "Vishakapatnam": "VISAKHAPATANAM", "ERNAKULAM": "ERNAKULAM/KOCHI",
}
# make groups of the portal (new_vehicle_makes)
MAKES_TOP = "ROYAL_SEP26_Skoda_TATA_MAHINDRA_KIA_Maruti_MG_Toyota_Hyundai_SPE"        # the top variants of the listed makes
MAKES_OTHER = "ROYAL_SEP26_EX_Skoda_TATA_MAHINDRA_KIA_Maruti_MG_Toyota_Hyundai_SPE"   # everything else
MAKES_STP = "Private_Car_ALL"
LISTED_MAKES = {"maruti", "tata", "mahindra", "hyundai", "toyota", "kia", "mg", "skoda"}   # the makes of MAKES_TOP
# "Any other make New car payout will be 28% flat": "write" = write New rows for the other makes (MAKES_OTHER, the flat rate of
# the note; chosen by the user on 24 Oct 2026); "flag" = write none, as the October upload did, and flag the New rows instead.
OTHER_MAKES_NEW = "write"
USER_ID = 2966
ROW_ID = 1                                       # the export carries id = 1 on every row
MAKE_MODEL_CLASS = "Car"
PO_CUT = 7.0                                     # po rate = pi rate - PO_CUT, never below 0
RATE_DECIMALS = 2                                # the grid has cells such as 0.2598599; the upload carries 25.99
AGE = (0, 99)                                    # vehicle_age_min / max of every row
CC_OPEN_END = 99999                              # cc_max of the top engine-size band
TARIFF = (0, 100.01)                             # tariff_min / tariff_max of every row
HEADER_SCAN_ROWS = 12                            # the table headers are looked for in these rows
ZERO_WORDS = ("NA", "N/A", "-", "--", "NIL", "BLOCK", "BLOCKED", "NOT APPLICABLE")
MAX_NOTE_CHARS = 500                             # a note longer than this is cut short in the remarks column

# Notes (text under a table, or above it) already known not to change the rows of this upload: titles, the volume-linked
# payout (VLI) tables and their exemptions - the upload carries the base payout only (pi_vli = 0).  A text is matched when
# it CONTAINS one of these phrases (capitals, spaces and punctuation ignored); a note that is not here and not read by a
# rule is flagged for manual review.  Add a phrase only after checking that the note really needs no action.
STANDING_NOTES = [
    "Pvt Car New/Renewal & Roll-over",
    "volume linked payouts applicable",
    "Base (Payout shown as per Grid above)",
    "exempted from the Volume linked payout slab",
    "exempted from VLI payouts",
    "VLI payouts are applicable only for Comp",
    # the September "Top Variant Models Drive" block: its base-payout extra is READ (see read_rules); these lines are its other text -
    # the business-share steps ("Base + 1% additional" at 25%, "Base + 2% additional" at 50%) are conditions on the agent's business mix,
    # not rate-card rows, and the upload does not carry them. A step with another percentage is a new line and gets flagged.
    "This payout is over & above the Grid payout",
    "List of Top End Variants attached in a separate sheet",
    "25% business share from Top variant models",
    "50% business share from Top variant models",
]

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
    "vehicle_age_max", "make_model_class", "is_ncb", "is_cpa", "is_zd", "cc_min", "cc_max", "add_tnc",
    "tariff_min", "tariff_max",
]


OUTPUT_FOLDER = "rsa"            # the folder of this insurer under output\ (default_output)


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
    """A sheet that cannot be used at all (no header, no clusters, no rate rows): the caller leaves it out or asks."""


def _text(v):
    return "" if v is None else str(v).replace(" ", " ").replace("​", "").strip()


def _col_letter(idx):
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx + 1)


def read_workbook(path):
    """Open with values only. NOT read-only mode: the sheets have merged note cells and only the normal mode keeps them intact."""
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def sheet_rows(wb, name):
    return [list(r) for r in wb[name].iter_rows(values_only=True)]


def label_key(s):
    """A label without spaces, punctuation and capitals: 'NCB Status' and 'ncb  status' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def is_rate_cell(v):
    return (isinstance(v, (int, float)) and not isinstance(v, bool)) or (isinstance(v, str) and v.strip().upper() in ZERO_WORDS)


def rto_name(cluster, part):
    """new_rto_list of a cluster: the prefix + part + the grid's name in capitals with spaces and commas -> underscores
    (see CLUSTER_NAME_EXCEPTIONS).  A name broken over two lines is joined, text in brackets is dropped.  A comma becomes an
    underscore: new_rto_list is itself a comma-separated list, so a comma inside a name would split it."""
    t = re.sub(r"\s*[\r\n]+\s*", " ", _text(cluster))
    t = re.sub(r"\s*\([^)]*\)", "", t).strip()
    special = {label_key(k): v for k, v in CLUSTER_NAME_EXCEPTIONS.items()}
    return rto_cluster(RTO_PREFIX + part + special.get(label_key(t), t.upper().replace(",", "_").replace(" ", "_")), t)      # the CSV carries the grid's cluster name (apply_cluster_names)


# ---- engine-size bands ("Upto 1000" -> cc_min / cc_max)
def cc_limits(label):
    """'Upto 1000' / 'Below 1000 CC' -> (0, 1000); '1000-1500' / '1000 to 1500 CC' -> (1000, 1500); '>1500' / 'Above 1500 CC'
    -> (1500, None); None if unreadable."""
    t = _text(label).lower().replace(",", "")
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", t)]
    if len(nums) >= 2 and re.search(r"\d\s*(?:cc)?\s*(?:to|-|–|—)\s*\d", t) and nums[1] > nums[0]:
        return nums[0], nums[1]
    if len(nums) == 1 and re.search(r"above|over|more|greater|>|\+|onwards", t):
        return nums[0], None
    if len(nums) == 1 and re.search(r"upto|up to|below|under|less|<", t):
        return 0.0, nums[0]
    return None


def resolve_cc(label, ask):
    """(cc_min, cc_max) of an engine-size label; asks once if the label cannot be read."""
    got = cc_limits(label)
    if got is None:
        def check(a):
            g = cc_limits(a)
            if g is None:
                raise ValueError("Type it like 1000-1500 (cc), upto 1000, or >1500 for no upper limit.")
            return g
        key = ("cc", label_key(label))
        if key not in ask.memory:
            ask.memory[key] = ask.text(f"cannot read the engine size '{label}'",
                                       question=f"I can't read the engine-size limits in '{_text(label)}'. "
                                                f"Type them like 1000-1500 (or >1500 for no upper limit):", validate=check)
            ask.decide(f"engine size '{_text(label)}' -> limits typed by you")
        got = ask.memory[key]
    lo, hi = got
    return (0 if lo == 0 else round(lo + 0.01, 2)), (round(hi + 0.01, 2) if hi is not None else CC_OPEN_END)


def check_bands(sheet, bands, ask):
    """The engine-size bands of a sheet should start at 0, join up and end open; anything else is put to the person."""
    ordered = sorted(set(bands))
    problems = []
    if ordered[0][0] != 0:
        problems.append(f"the lowest band starts at {ordered[0][0]:g} cc, not at 0")
    for a, b in zip(ordered, ordered[1:]):
        if abs(a[1] - b[0]) > 0.5:
            problems.append(f"a band ends at {a[1]:g} cc but the next starts at {b[0]:g} cc")
    if ordered[-1][1] != CC_OPEN_END:
        problems.append(f"the highest band stops at {ordered[-1][1]:g} cc (no 'above' band)")
    if problems and not ask.choose(
            f"'{sheet}': the engine-size bands do not join up: {'; '.join(problems)}",
            [("Carry on with the bands as they are", True), ("Stop so I can look at the grid", False)],
            question=f"In '{sheet}' the engine-size bands do not join up ({'; '.join(problems)}). What should I do?",
            short=f"{sheet}: engine-size bands", memo=("bandjoin", sheet)):
        raise GridError(f"stopped by you - '{sheet}' engine-size bands do not join up")


# ---- business type / NCB labels
def business_kind(label, ask):
    """'Renewal & Roll Over Comp (1+1)' -> 'comp'; '... SAOD' -> 'saod'; 'New Car(including EV)' -> 'new'; None = leave out."""
    k = label_key(label)
    if "saod" in k:
        return "saod"
    if "comp" in k:
        return "comp"
    if k.startswith("new") and "renew" not in k:
        return "new"
    return ask.choose(
        f"unknown business type {_text(label)!r}",
        [("Comprehensive (1+1)", "comp"), ("SAOD (own damage only)", "saod"), ("New car (1+3)", "new"), ("Leave these rows out of the upload", None)],
        question=f"I don't know the business type {_text(label)!r}. What is it?", short=f"business type {_text(label)!r}",
        memo=("business", k))


def ncb_value(label, ask):
    """'With NCB' -> 'YES', 'Without NCB' -> 'NO', 'NA' / blank -> 'NA'."""
    k = label_key(label)
    if "without" in k or k in ("nonncb", "noncb"):
        return "NO"
    if "with" in k:
        return "YES"
    if k in ("", "na", "nil", "notapplicable"):
        return "NA"
    return ask.choose(
        f"unknown NCB label {_text(label)!r}",
        [("With NCB (is_ncb = YES)", "YES"), ("Without NCB (is_ncb = NO)", "NO"), ("Not applicable (is_ncb = NA)", "NA")],
        question=f"I don't know the NCB label {_text(label)!r}. What is it?", short=f"NCB label {_text(label)!r}", memo=("ncb", k))


def fuels_for(label, ask):
    """The fuels a column group of the STP sheet is for: 'DIESEL' -> Diesel, 'Non Diesel' -> every other fuel."""
    k = label_key(label)
    if "non" in k and "diesel" in k:
        return [f for f in FUELS if f != "Diesel"]
    named = [f for f in FUELS if label_key(f) in k]
    if named:
        return named
    if "ev" in k or "electric" in k:
        return ["Electric"]
    return ask.choose(
        f"unknown fuel group {_text(label)!r} in the STP sheet",
        [("Diesel only", ["Diesel"]), ("Every fuel except Diesel", [f for f in FUELS if f != "Diesel"]),
         ("Every fuel", list(FUELS)), ("Leave these columns out", None)],
        question=f"I don't know which fuels the STP columns headed {_text(label)!r} are for. Which are they?",
        short=f"STP fuel group {_text(label)!r}", memo=("fuelgroup", k))


# ---- the rules written in the notes under each sheet
_N = r"\d+(?:\.\d+)?"
_PCT = rf"(?P<n>{_N})\s*(?:%|percent|pts|points)"
_WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}
P_TOP = re.compile(rf"(?:additional|extra|plus|more)\s+{_PCT}[^.]*?top[\s-]*(?:end\s+)?variants?", re.I)
P_OTHER_NEW = re.compile(rf"other\s+makes?\s+new\s+(?:car\s+)?payout\s+(?:will\s+be|is|shall\s+be)\s+{_PCT}", re.I)
P_NEW_MAKES = re.compile(r"new\s+car\s+payout\s+is\s+applicable\s+only\s+for\s+the\s+following\s+makes?\s*[-–:]*\s*(?P<who>.+)", re.I | re.S)
P_TOP_MAKES = re.compile(r"top[\s-]*(?:end\s+)?variants?\s+models?\s+of\s+(?P<who>[^()]+)", re.I)
P_ELIGIBLE = re.compile(r"eligible\s+makes?\s*[-–:]*\s*(?P<who>.+)", re.I | re.S)
# the "Top Variant Models Drive" block: "With NCB:" / "Without NCB:" headings, then a "Base payout" cell followed by the "N% additional" cell
P_NCB_HEAD = re.compile(r"^\s*(?P<w>with|without)\s+ncb\s*:?\s*$", re.I)
P_BASE_TIER = re.compile(r"^\s*base\s+payout\s*:?\s*$", re.I)
P_TIER_VALUE = re.compile(rf"^\s*(?P<n>{_N})\s*(?:%|percent)\s*additional\s*$", re.I)


def _normalise(text):
    """'two percent' -> '2%' so the patterns see digits"""
    text = re.sub(r"\b(" + "|".join(_WORDS) + r")\b(?=\s*(?:%|percent|points?|pts))", lambda m: str(_WORDS[m.group(1).lower()]), text, flags=re.I)
    return re.sub(r"\s*percent\b", "%", text, flags=re.I)


def parse_makes(text):
    """'Maruti/Tata/Mahindra/Hyundai/Toyota/Kia/MG & Skoda.' -> {'maruti', 'tata', ...} (a mangled S in Skoda is understood)"""
    out = set()
    for tok in re.split(r"[/,&]|\band\b", text, flags=re.I):
        k = re.sub(r"[^a-z]", "", tok.lower())
        if k.endswith("koda"):
            k = "skoda"
        if k:
            out.add(k)
    return out


class Reading:
    """What the notes of one sheet say."""

    def __init__(self):
        self.top_extra, self.top_text = None, ""            # "Additional N% payout for the Top Variant Models" (the same for every NCB)
        self.top_ncb = {}                                   # the "Drive" block: {"YES": N, "NO": N} = the base-payout extra With / Without NCB
        self.other_new, self.other_text = None, ""          # "Any other make New car payout will be N% flat"
        self.new_makes, self.new_makes_text = None, ""      # makes the New car payout is for
        self.top_makes, self.top_makes_text = None, ""      # makes the top-variant extra is for
        self.conflicts = []                                 # notes that give two different values for one rule
        self.conflict_rules = set()                         # which rules ("top", "other") have such a conflict
        self.leftover = []                                  # notes no rule understood

    def extra_for(self, ncb):
        """The top-variant extra in points for rows with this is_ncb (YES / NO), or None when the notes do not give it."""
        return self.top_ncb[ncb] if ncb in self.top_ncb else self.top_extra


def read_rules(items):
    """Run every note (one cell) through the patterns; the notes none of them understood are left in .leftover."""
    rd = Reading()
    found = {"top": [], "other": []}
    drive = {}                                              # NCB ("YES" / "NO") -> [(N, the cell text)] from the "Drive" block
    ncb, expect = None, False
    for raw in items:
        text = _normalise(re.sub(r"\s+", " ", raw)).strip()
        used = False
        h = P_NCB_HEAD.match(text)
        if h:
            ncb, expect = ("NO" if h.group("w").lower() == "without" else "YES"), False
            continue
        if P_BASE_TIER.match(text):
            expect = ncb is not None
            continue
        m = P_TIER_VALUE.match(text)
        if m and expect:
            drive.setdefault(ncb, []).append((float(m.group("n")), raw))
            expect = False
            continue
        if "business share" in text.lower():
            expect = False
        m = P_ELIGIBLE.search(text)
        if m:
            rd.top_makes, rd.top_makes_text, used = parse_makes(m.group("who")), raw, True
        m = P_TOP.search(text)
        if m:
            found["top"].append((float(m.group("n")), raw))
            used = True
        m = P_OTHER_NEW.search(text)
        if m:
            found["other"].append((float(m.group("n")), raw))
            used = True
        m = P_NEW_MAKES.search(text)
        if m:
            rd.new_makes, rd.new_makes_text, used = parse_makes(m.group("who")), raw, True
        m = P_TOP_MAKES.search(text)
        if m:
            rd.top_makes, rd.top_makes_text, used = parse_makes(m.group("who")), raw, True
        if not used:
            rd.leftover.append(raw)
    for key, attr in (("top", "top"), ("other", "other")):
        hits = found[key]
        if not hits:
            continue
        if len({v for v, _ in hits}) > 1:
            rd.conflicts += [t for _, t in hits]
            rd.conflict_rules.add(key)
            continue
        if attr == "top":
            rd.top_extra, rd.top_text = hits[0]
        else:
            rd.other_new, rd.other_text = hits[0]
    for w, hits in drive.items():                          # the "Drive" block: one base-payout extra per NCB heading
        if len({v for v, _ in hits}) > 1:
            rd.conflicts += [t for _, t in hits]
            rd.conflict_rules.add("top")
            rd.top_ncb.pop(w, None)
        else:
            rd.top_ncb[w] = hits[0][0]
    if rd.top_extra is not None and any(v != rd.top_extra for v in rd.top_ncb.values()):
        rd.conflicts += [rd.top_text] + [t for hits in drive.values() for _, t in hits]    # a flat note AND a "Drive" block that disagree
        rd.conflict_rules.add("top")
        rd.top_extra, rd.top_ncb = None, {}
    if "top" in rd.conflict_rules:
        rd.top_extra, rd.top_ncb = None, {}
    return rd


def rules_line(rd):
    if rd.top_ncb:
        top = "top-variant extra " + " / ".join(f"+{rd.extra_for(w):g} pts {'With' if w == 'YES' else 'Without'} NCB" for w in ("YES", "NO") if rd.extra_for(w) is not None)
        if any(rd.extra_for(w) is None for w in ("YES", "NO")):
            top += " (the other NCB NOT FOUND, none added)"
    elif rd.top_extra is not None:
        top = f"top-variant extra +{rd.top_extra:g} pts"
    else:
        top = "top-variant extra NOT FOUND (none added)"
    if rd.other_new is not None:
        other = f"other makes New {rd.other_new:g}% flat ({'written' if OTHER_MAKES_NEW == 'write' else 'NOT written, rows flagged'})"
    else:
        other = "no flat New rate for other makes in the notes"
    return f"{top}; {other}"


def _short(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text if len(text) <= MAX_NOTE_CHARS else text[:MAX_NOTE_CHARS].rstrip() + "..."


# ---- one comprehensive / SAOD / New sheet
def find_comp_header(rows):
    """Index of the row that holds 'Business Type', 'CC' and 'NCB Status', or None."""
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        keys = [label_key(v) for v in row]
        if "businesstype" in keys and "ncbstatus" in keys and "cc" in keys:
            return i
    return None


class CompSheet:
    """One region sheet: business type / engine size / NCB down the side, clusters across."""

    def __init__(self, sheet, rows, head, ask):
        self.sheet, self.rows, self.ask = sheet, rows, ask
        header = rows[head]
        keys = [label_key(v) for v in header]
        self.bcol, self.ccol, self.ncol = keys.index("businesstype"), keys.index("cc"), keys.index("ncbstatus")
        first = max(self.bcol, self.ccol, self.ncol) + 1
        self.region = next((_text(v) for r in range(head) for v in rows[r][first:] if _text(v)), "") or sheet
        self.table, seen, label = [], {}, ""
        for r in range(head + 1, len(rows)):
            row = rows[r]
            get = lambda c: _text(row[c]) if c < len(row) else ""
            cc_text = get(self.ccol)
            if not cc_text or not any(is_rate_cell(v) for v in row[first:]):
                if self.table:
                    if self._table_follows(r + 1, first):
                        continue                                            # a blank row inside the table: its rows go on below it
                    break                                                   # the first blank or non-rate row that no table row follows ends the table
                continue
            label = get(self.bcol) or label                                 # a business type merged over several rows
            kind = business_kind(label, ask)
            if kind is None:
                continue
            ncb = "NA" if kind == "new" and not get(self.ncol) else ncb_value(get(self.ncol), ask)
            band = resolve_cc(cc_text, ask)
            key = (kind, ncb, band)
            if key in seen:
                keep = ask.choose(f"'{sheet}': {kind.upper()} / NCB {ncb} / {band[0]:g}-{band[1]:g} cc appears twice (rows {seen[key] + 1} and {r + 1})",
                                  [("Use the first of them", "first"), ("Use the last of them", "last")],
                                  question=f"'{sheet}' has the same row twice ({kind.upper()}, NCB {ncb}, {cc_text}: rows {seen[key] + 1} and "
                                           f"{r + 1}). Which one should I use?", short=f"{sheet}: repeated row {cc_text}",
                                  memo=("duprow", sheet, key))
                if keep == "first":
                    continue
                self.table = [t for t in self.table if t["key"] != key]
            seen[key] = r
            self.table.append(dict(r=r, kind=kind, ncb=ncb, band=band, key=key))
        if not self.table:
            raise EmptySheet(f"'{sheet}': no Comp / SAOD / New rate rows found under the header row {head + 1}")
        check_bands(sheet, [t["band"] for t in self.table], ask)
        self.last_row = max(t["r"] for t in self.table)
        self.clusters, self.no_data = [], []
        for j, v in enumerate(header):
            if j < first or not _text(v):
                continue
            has = any(j < len(rows[t["r"]]) and is_rate_cell(rows[t["r"]][j]) for t in self.table)
            (self.clusters if has else self.no_data).append((j, _text(v)))
        if not self.clusters:
            raise EmptySheet(f"'{sheet}': no cluster column with rates under the header row {head + 1}")
        named = {j for j, _ in self.clusters} | {j for j, _ in self.no_data}
        self.unnamed = [j for j in range(first, max(len(rows[t["r"]]) for t in self.table))
                        if j not in named and any(j < len(rows[t["r"]]) and isinstance(rows[t["r"]][j], (int, float))
                                                  and not isinstance(rows[t["r"]][j], bool) and rows[t["r"]][j] != 0
                                                  for t in self.table)]
        above = rows[head - 1] if head > 0 else []
        self.states = {name: (_text(above[j]) if j < len(above) else "") for j, name in self.clusters}   # cluster -> state label
        notes = []                                                          # one item per cell: a note can run over several lines
        for r in list(range(head)) + list(range(self.last_row + 1, min(len(rows), self.last_row + 80))):
            for v in rows[r]:
                if isinstance(v, str) and _text(v):
                    notes.append(_text(v))
        self.notes = notes

    def _table_follows(self, start, first):
        """True when the next non-blank row from `start` is a row of this table (a CC text and rates); notes under the table are not"""
        for i in range(start, len(self.rows)):
            row = self.rows[i]
            if not any(_text(v) for v in (row or [])):
                continue
            cc = _text(row[self.ccol]) if self.ccol < len(row) else ""
            return bool(cc and any(is_rate_cell(v) for v in row[first:]))
        return False

    def rate(self, r, j):
        """The cell as a payout in percent (0.26 in the grid -> 26)."""
        row = self.rows[r]
        v = row[j] if j < len(row) else None
        addr = f"'{self.sheet}'!{_col_letter(j)}{r + 1}"
        kind, x = read_number(self.ask, addr, v, "fraction", "a payout such as 0.26 (= 26%)", zero_words=ZERO_WORDS)
        return x * 100 if kind == "num" else 0.0                # full precision: rounding to 2 decimals is done once, in _row


# ---- the STP sheet
def find_stp_header(rows):
    """Index of the row that holds 'RTO DIVISION', or None."""
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        if any(label_key(v) == "rtodivision" for v in row):
            return i
    return None


class StpSheet:
    """Third-party-only rates by RTO division: a column group per fuel kind (Diesel / Non Diesel), engine-size columns in each."""

    def __init__(self, sheet, rows, head, ask):
        self.sheet, self.rows, self.ask = sheet, rows, ask
        header = rows[head]
        self.dcol = next(j for j, v in enumerate(header) if label_key(v) == "rtodivision")
        sub = rows[head + 1] if head + 1 < len(rows) else []
        groups = [(j, _text(v)) for j, v in enumerate(header) if j > self.dcol and _text(v)]
        self.cols = []                                                      # (column, (cc_min, cc_max), fuels)
        for j in range(self.dcol + 1, len(sub)):
            if not _text(sub[j]) or cc_limits(sub[j]) is None:
                continue
            owner = [g for g in groups if g[0] <= j]
            if owner:
                fuels = fuels_for(owner[-1][1], ask)
            else:
                fuels = ask.choose(f"'{sheet}': no fuel heading above column {_col_letter(j)}",
                                   [("Every fuel", list(FUELS)), ("Leave these columns out", None)],
                                   question=f"'{sheet}' has no fuel heading (Diesel / Non Diesel) above the column '{_text(sub[j])}'. Which fuels is it for?",
                                   short=f"{sheet}: no fuel heading", memo=("stpnofuel", sheet))
            if fuels:
                self.cols.append((j, resolve_cc(sub[j], ask), fuels))
        if not self.cols:
            raise EmptySheet(f"'{sheet}': no engine-size rate columns found under the header row {head + 1}")
        check_bands(sheet, [c[1] for c in self.cols], ask)
        self.divisions = []                                                 # (row index, division name)
        for r in range(head + 2, len(rows)):
            row = rows[r]
            name = _text(row[self.dcol]) if self.dcol < len(row) else ""
            if not name or not any(c < len(row) and is_rate_cell(row[c]) for c, _, _ in self.cols):
                if self.divisions:
                    if self._division_follows(r + 1):
                        continue                                            # a blank row inside the table: the divisions go on below it
                    break
                continue
            self.divisions.append((r, name))
        if not self.divisions:
            raise EmptySheet(f"'{sheet}': no RTO division rows found under the header row {head + 1}")
        self.last_row = self.divisions[-1][0]
        notes = []
        for r in list(range(head)) + list(range(self.last_row + 1, min(len(rows), self.last_row + 80))):
            for v in rows[r]:
                if isinstance(v, str) and _text(v):
                    notes.append(_text(v))
        self.notes = notes
        covered = {f for _, _, fs in self.cols for f in fs}
        self.uncovered = [f for f in FUELS if f not in covered]

    def _division_follows(self, start):
        """True when the next non-blank row from `start` is a division row of this table (a name and rates); notes under the table are not"""
        for i in range(start, len(self.rows)):
            row = self.rows[i]
            if not any(_text(v) for v in (row or [])):
                continue
            name = _text(row[self.dcol]) if self.dcol < len(row) else ""
            return bool(name and any(c < len(row) and is_rate_cell(row[c]) for c, _, _ in self.cols))
        return False

    def rate(self, r, j):
        row = self.rows[r]
        v = row[j] if j < len(row) else None
        addr = f"'{self.sheet}'!{_col_letter(j)}{r + 1}"
        kind, x = read_number(self.ask, addr, v, "fraction", "a payout such as 0.35 (= 35%)", zero_words=ZERO_WORDS)
        return x * 100 if kind == "num" else 0.0                # full precision: rounding to 2 decimals is done once, in _row


def top_variant_counts(rows):
    """{make: number of variant rows} from the 'Top End Variant List' sheet: a block starts at a cell such as 'Skoda - Top-end Variants',
    two rows below it the make / model / variant rows follow until a blank row."""
    out = {}
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if isinstance(v, str) and label_key(v).endswith("topendvariants"):
                n, i = 0, r + 2
                while i < len(rows) and any(_text(rows[i][k]) for k in range(c, min(c + 3, len(rows[i])))):
                    n, i = n + 1, i + 1
                out[_text(v).split(" - ")[0].strip() or _text(v)] = n
    return out


def find_sheets(wb, ask, log):
    """(comp sheets, STP sheets, variant counts). Sheets are found by their headers, not their names; hidden sheets are not read."""
    comp, stp, top, hidden, left_out = [], [], {}, [], []
    for name in wb.sheetnames:
        if getattr(wb[name], "sheet_state", "visible") != "visible":
            hidden.append(name)
            continue
        rows = sheet_rows(wb, name)
        key = label_key(name)
        ch, sh = find_comp_header(rows), find_stp_header(rows)
        if "topend" in key and ch is None and sh is None:
            top = top_variant_counts(rows)
            continue
        if ch is None and sh is None:
            if any(w in key for w in ("comp", "saod", "stp")):
                ask.choose(f"sheet '{name}' looks like a rate sheet but has no 'Business Type / NCB Status' or 'RTO DIVISION' header",
                           [("Leave this sheet out", None)],
                           question=f"The sheet '{name}' looks like a rate sheet, but I can't find its header row in the first "
                                    f"{HEADER_SCAN_ROWS} rows. Leave it out?", short=f"sheet '{name}'", memo=("nohead", name))
            left_out.append(name)
            continue
        title = " ".join(_text(v) for r in rows[:3] for v in r if _text(v))
        if ch is not None and not any(w in label_key(name) + label_key(title) for w in ("pvtcar", "privatecar")):
            if not ask.choose(f"sheet '{name}' (title {title[:60]!r}) does not say Pvt Car / Private Car",
                              [("Use it anyway - it is a Private Car sheet", True), ("Leave this sheet out", False)],
                              question=f"The sheet '{name}' has the title {title[:60]!r}, which does not say Pvt Car. Is it a Private Car sheet?",
                              short=f"sheet '{name}' title", memo=("title", name)):
                left_out.append(name)
                continue
        try:
            if ch is not None:
                comp.append(CompSheet(name, rows, ch, ask))
            else:
                stp.append(StpSheet(name, rows, sh, ask))
        except EmptySheet as e:
            ask.choose(str(e), [("Leave this sheet out", None)], question=f"{e}. I can't use this sheet. Leave it out?",
                       short=f"sheet '{name}' unusable")
            left_out.append(name)
    for group, names_of in ((comp, lambda x: {n for _, n in x.clusters}), (stp, lambda x: {n for _, n in x.divisions})):
        for i, a in enumerate(group):
            for b in group[i + 1:]:
                both = names_of(a) & names_of(b)
                if both:
                    raise GridError(f"the sheets '{a.sheet}' and '{b.sheet}' both hold {sorted(both)[:4]}{' ...' if len(both) > 4 else ''} - expected one sheet per region; "
                                    f"hide or delete the one that is not this month's grid")
    if not comp:
        raise GridError("no Private Car Comp / SAOD sheet found - looked for the header cells 'Business Type', 'CC', 'NCB Status' "
                        f"in the first {HEADER_SCAN_ROWS} rows of every visible sheet ({wb.sheetnames})")
    if not stp:
        ask.choose("no STP sheet found (a sheet with the header 'RTO DIVISION')",
                   [("Carry on without STP rows", None)],
                   question="I can't find the STP sheet (the one with the header 'RTO DIVISION'). Carry on without STP rows?",
                   short="no STP sheet", memo=("nostp",))
    log(f"Sheets used: {[s.sheet for s in comp + stp]}" + (f" | not used: {left_out}" if left_out else "")
        + (f" | hidden, not read: {hidden}" if hidden else ""))
    return comp, stp, top


def note_scope(text, sh):
    """The clusters a note may concern: the ones it names (a state name brings in all of that state's clusters, and also every
    "Priority Cities" column, because the sheet does not say which state those cities are in); None = the whole sheet (it names
    none the script can match). A superset is flagged, never a subset."""
    compact = note_key(text)
    hit = set()
    for name in sh.states:
        k = note_key(re.sub(r"\s*\([^)]*\)", "", name))
        if len(k) >= 4 and k in compact:
            hit.add(name)
    for label in set(sh.states.values()):
        k = note_key(label)
        if len(k) >= 4 and not k.startswith("priority") and k in compact:
            hit |= {c for c, s in sh.states.items() if s == label}
            hit |= {c for c, s in sh.states.items() if not note_key(s) or note_key(s).startswith("priority")}
    return hit or None


def rows_where(sheet, rto_names=None, makes=None, subs=None, ncbs=None):
    """A function of the output frame: the rows of one sheet, narrowed by cluster, make group, sub_product, is_ncb."""
    def fn(df):
        m = np.array(df["_sheet"] == sheet, dtype=bool)
        if rto_names is not None:
            m = m & np.array(df["new_rto_list"].isin(rto_names), dtype=bool)
        if makes is not None:
            m = m & np.array(df["new_vehicle_makes"].isin(makes), dtype=bool)
        if subs is not None:
            m = m & np.array(df["sub_product"].isin(subs), dtype=bool)
        if ncbs is not None:
            m = m & np.array(df["is_ncb"].isin(ncbs), dtype=bool)
        return m
    return fn


def round_half_up(x, places=RATE_DECIMALS):
    """Round to `places` decimals with a true half going up, as Excel does (Python's round() sends 33.125 to 33.12). The value is
    first cleaned of binary noise (6 decimals), so 0.285 * 100 = 28.499999999999996 is the half it was meant to be."""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(repr(round(float(x), 6))).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP))


# %% ---------------------------------------------------------------- BUILD
SUBS = {"comp": ("1+1", "Renew"), "saod": ("SAOD", "Renew"), "new": ("1+3", "New")}


def _row(rto, make, sub, policy, fuel, cc, ncb, rate, dates, sheet, stp=False):
    """One upload row. `rate` is the payin in percent; the payout is rate - PO_CUT, never below 0 (zero stays zero)."""
    pi = round_half_up(rate)
    po = round_half_up(max(pi - PO_CUT, 0.0)) if pi > 0 else 0.0
    kind = "On TP" if stp else "On OD"
    return dict(
        insurance_company=INSURANCE_COMPANY, insurer_vertical="Not Applicable", product=PRODUCT, sub_product=sub,
        policy_type=policy, fuel_type=fuel, make_model_class=MAKE_MODEL_CLASS, new_vehicle_makes=make, new_rto_list=rto,
        vehicle_age_min=AGE[0], vehicle_age_max=AGE[1], cc_min=cc[0], cc_max=cc[1], sc_min=0, sc_max=999,
        pi_od_rate=0.0 if stp else pi, pi_tp_rate=pi if stp else 0.0, pi_tp_2=0, pi_tp_3=0, pi_tp_4=0, pi_tp_5=0,
        pi_net_rate=0.0, pi_flat_amount=0, pi_vli=0, pi_type=kind, tariff_min=TARIFF[0], tariff_max=TARIFF[1],
        is_ncb=ncb, is_cpa="NA", is_zd="NA", from_date=dates[0], to_date=dates[1], user_id=USER_ID, veh_use="", add_tnc="",
        remarks="", po_type=kind, po_od_rate=0.0 if stp else po, po_tp_rate=po if stp else 0.0, po_net_rate=0.0,
        po_flat_amount=0, _sheet=sheet,
    )


def build(grid_path, from_date, to_date, log=say, ask=None, wb=None):
    """Read the workbook and return (DataFrame in OUTPUT_COLUMNS order with numeric columns as floats, warnings).
    from_date / to_date are datetime.date values; they are written as DD/MM/YYYY into every row.
    `ask` (an Asker) is consulted whenever the grid itself is missing something. The NOTES are never asked about: what they
    say is applied, and what cannot be applied is flagged in the remarks column (see ManualReview)."""
    if to_date < from_date:
        raise GridError(f"to_date {to_date:%d/%m/%Y} is before from_date {from_date:%d/%m/%Y}")
    ask = ask or Asker()
    dates = (from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
    if wb is None:
        log(f"Reading {os.path.basename(str(grid_path))}...")
        grid_path, wb = open_file(ask, grid_path, "the RSA workbook", read_workbook, EXCEL_EXTS)
    log(f"Sheets in the workbook: {wb.sheetnames}")
    comp, stps, top = find_sheets(wb, ask, log)
    log(f"new_rto_list = {RTO_PREFIX}{COMP_RTO_PART}<cluster>  /  {RTO_PREFIX}{STP_RTO_PART}<division>   (RTO_PREFIX in the settings, or --prefix)")
    if top:
        log("Top End Variant List: " + ", ".join(f"{k} {v}" for k, v in top.items()) + f"  (total {sum(top.values())} variants) - if this "
            f"list changed since last month, the make group {MAKES_TOP} in the portal needs the same change")
    warnings, rows, review = [], [], ManualReview()
    rto_of = lambda clusters, part: None if clusters is None else {rto_name(c, part) for c in clusters}

    for sh in comp:
        rd = read_rules(sh.notes)
        log(f"'{sh.sheet}' ({sh.region}): {len(sh.clusters)} clusters x {len(sh.table)} rate rows")
        log(f"    rules from the notes: {rules_line(rd)}")
        for j, name in sh.no_data:
            warnings.append(f"'{sh.sheet}' column {_col_letter(j)} ('{name}') has a header but no rates - left out")
        for j in sh.unnamed:
            warnings.append(f"'{sh.sheet}' column {_col_letter(j)} has rates but no cluster name in the header - its rates are NOT in the output")
        for kind in SUBS:
            if not any(t["kind"] == kind for t in sh.table):
                warnings.append(f"'{sh.sheet}' has no {kind.upper()} rows")
        write_other = OTHER_MAKES_NEW == "write" and rd.other_new is not None
        for t in sh.table:
            sub, policy = SUBS[t["kind"]]
            for j, cluster in sh.clusters:
                base = sh.rate(t["r"], j)
                rto = rto_name(cluster, COMP_RTO_PART)
                for fuel in FUELS:
                    if t["kind"] == "new":
                        rows.append(_row(rto, MAKES_TOP, sub, policy, fuel, t["band"], t["ncb"], base, dates, sh.sheet))
                        if write_other:
                            rows.append(_row(rto, MAKES_OTHER, sub, policy, fuel, t["band"], t["ncb"], rd.other_new, dates, sh.sheet))
                    else:
                        extra = rd.extra_for(t["ncb"]) or 0.0
                        rows.append(_row(rto, MAKES_TOP, sub, policy, fuel, t["band"], t["ncb"], base + extra if base > 0 else 0.0, dates, sh.sheet))
                        rows.append(_row(rto, MAKES_OTHER, sub, policy, fuel, t["band"], t["ncb"], base, dates, sh.sheet))

        # ---- what in the notes could not be applied: flagged on the rows it may affect, never guessed
        for text in unknown_notes(rd.leftover, STANDING_NOTES):
            review.flag(_short(text), rows_where(sh.sheet, rto_of(note_scope(text, sh), COMP_RTO_PART)))
        for text in rd.conflicts:
            review.flag(_short(f"{sh.sheet}: two notes give different values for one rule: {text}"), rows_where(sh.sheet))
        if "top" not in rd.conflict_rules:
            missing = [w for w in ("YES", "NO") if rd.extra_for(w) is None]
            if missing:
                which = "" if len(missing) == 2 else f" (rows with is_ncb {missing[0]} only)"
                review.flag(f"{sh.sheet}: the 'Additional N% payout for the Top Variant Models' note was not found or not readable, "
                            f"so no extra was added to the top-variant rows{which}",
                            rows_where(sh.sheet, makes=[MAKES_TOP], subs=["1+1", "SAOD"], ncbs=missing))
        if rd.other_text and OTHER_MAKES_NEW != "write":
            review.flag(_short(rd.other_text) + " - not applied: no New rows are written for other makes", rows_where(sh.sheet, subs=["1+3"]))
        if OTHER_MAKES_NEW == "write" and rd.other_new is None and "other" not in rd.conflict_rules:
            review.flag(f"{sh.sheet}: the flat New car payout for other makes was not found in the notes, so no rows were written for them",
                        rows_where(sh.sheet, subs=["1+3"]))
        for names, text, what in ((rd.new_makes, rd.new_makes_text, "New car payout"), (rd.top_makes, rd.top_makes_text, "top-variant extra")):
            if names is not None and names != LISTED_MAKES:
                review.flag(_short(f"{sh.sheet}: the {what} note names the makes {sorted(names)} but the make group {MAKES_TOP} is for "
                                   f"{sorted(LISTED_MAKES)} - check the group in the portal. Note: {text}"),
                            rows_where(sh.sheet, makes=[MAKES_TOP]))

    for st in stps:
        rd = read_rules(st.notes)
        log(f"'{st.sheet}': {len(st.divisions)} RTO divisions x {len(st.cols)} rate columns")
        for f in st.uncovered:
            warnings.append(f"'{st.sheet}' has no columns for {f} - no STP rows are written for it")
        for r, name in st.divisions:
            rto = rto_name(name, STP_RTO_PART)
            for j, cc, fuels in st.cols:
                base = st.rate(r, j)
                for fuel in fuels:
                    rows.append(_row(rto, MAKES_STP, "STP", "Renew", fuel, cc, "NA", base, dates, st.sheet, stp=True))
        for text in unknown_notes(st.notes, STANDING_NOTES):
            review.flag(_short(text), rows_where(st.sheet))

    first = comp[0]
    for sh in comp[1:]:
        if {t["band"] for t in sh.table} != {t["band"] for t in first.table}:
            warnings.append(f"'{sh.sheet}' has different engine-size bands from '{first.sheet}' - check both sheets")
    if not rows:
        raise GridError("nothing to write - every part of the workbook was left out")
    df = pd.DataFrame(rows)
    df["id"] = ROW_ID
    check_names(df, review, load_masters(ask), log)                      # every RTO / make group written must be in the master lists
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


def diff_against_reference(df, ref_path, log=say):
    ref = pd.read_csv(ref_path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig", encoding_errors="replace")
    mine = to_csv_frame(df)
    missing = [c for c in KEY_COLUMNS if c not in ref.columns]
    if missing:
        raise ValueError(f"the earlier upload has no column(s) {missing} - it does not look like an upload file")

    def key(d):
        d = d.copy()
        for c in ("vehicle_age_min", "vehicle_age_max", "cc_min", "cc_max", "tariff_min", "tariff_max"):
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
        by = only_ref.groupby(["fuel_type_ref" if "fuel_type_ref" in only_ref else "fuel_type",
                               "new_vehicle_makes_ref" if "new_vehicle_makes_ref" in only_ref else "new_vehicle_makes"]).size()
        log("    removed rows are: " + "; ".join(f"{n} x {f} / {m}" for (f, m), n in by.head(4).items())
            + "  (rows the earlier upload has and this grid does not, for example Electric rows from another grid)")
    cl_new = sorted(set(mine["new_rto_list"]) - set(ref["new_rto_list"]))
    cl_old = sorted(set(ref["new_rto_list"]) - set(mine["new_rto_list"]))
    if cl_new:
        log(f"    clusters not in the earlier upload ({len(cl_new)}): {cl_new[:8]}{' ...' if len(cl_new) > 8 else ''}"
            f" - their RTO groups must exist in the portal")
    if cl_old:
        log(f"    clusters in the earlier upload but not in this grid ({len(cl_old)}): {cl_old[:8]}{' ...' if len(cl_old) > 8 else ''}")
    return len(only_new), len(only_ref), int(changed.sum())




def report(df, warnings, log=say, ask=None):
    """Print what is about to be written; rows that repeat a key keep the lowest payin (resolve_duplicates). Returns the frame to write."""
    ask = ask or Asker()
    df = resolve_duplicates(df, KEY_COLUMNS, ask)
    log(f"\nRows to write: {len(df):,}   ({df['new_rto_list'].nunique()} clusters, fuels {', '.join(sorted(df['fuel_type'].unique()))})")
    counts = df.groupby(["sub_product", "new_vehicle_makes"]).size()
    for (sub, mk), n in counts.items():
        log(f"  {sub:<5} {mk:<68} {n:>7,}")
    log(f"Duplicate keys: {int(df.duplicated(KEY_COLUMNS).sum())}")
    pi = df[["pi_od_rate", "pi_tp_rate"]].max(axis=1)
    po = df[["po_od_rate", "po_tp_rate"]].max(axis=1)
    log(f"PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {int(((pi > 0) & (po == 0)).sum())}")
    for w in warnings:
        log(f"WARNING: {w}")
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
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"RSA_PVTCAR_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
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
    grid = books[0] if len(books) == 1 else ask.choose(
        "more than one workbook was uploaded", [(n, n) for n in books],
        question="You uploaded more than one workbook. Which one is the RSA Private Car grid?", short="workbook")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the RSA Private Car workbook (.xlsx) and, optionally, an earlier upload (.csv) to compare with")
    try:
        grid, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("rsa_pvtcar_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    global RTO_PREFIX
    ap = argparse.ArgumentParser(description="RSA Private Car grid -> RSA upload CSV")
    ap.add_argument("--grid", help="RSA workbook (.xlsx) - asked for if left out")
    ap.add_argument("--from-date", help="from_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--to-date", help="to_date as DD/MM/YYYY (asked for if left out)")
    ap.add_argument("--ref", help="earlier upload CSV to diff against (asked for if left out)")
    ap.add_argument("--prefix", help=f"RTO group prefix for new_rto_list (default {RTO_PREFIX})")
    ap.add_argument("--out", help="output CSV path")
    ap.add_argument("--no-prompt", action="store_true", help="never ask questions: stop with the question as the error")
    a = ap.parse_args()
    ask = Asker(interactive=not a.no_prompt)      # a closed keyboard (EOF) is reported when a question cannot be asked
    try:
        if a.prefix:
            if re.search(r"\s", a.prefix):
                raise GridError(f"the prefix {a.prefix!r} contains a space - RTO group names use underscores")
            RTO_PREFIX = a.prefix
        grid = a.grid or ask_file(ask, "no workbook given - pass --grid <workbook.xlsx>",
                                  "Path of the RSA Private Car workbook (.xlsx) - you can drag the file into this window:",
                                  exts=EXCEL_EXTS)
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
        saved = crash_report("rsa_pvtcar_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
