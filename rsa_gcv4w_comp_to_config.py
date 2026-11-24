"""
Royal Sundaram (RSA) GCV 4W comprehensive grid -> RSA upload CSV
================================================================

Builds the rate-config upload file (the same 42 columns as "rates_export (56).csv") from the RSA workbook
circulated every month ("RSA - GCV Comp Grid Oct 26 V1 ...xlsx": one sheet per region - WEST, EAST, NORTH, SOUTH -
with weight bands down the side and clusters across).  Rates come only from that workbook.

How a grid cell becomes rows (checked against the October upload: every one of its 16,848 rate rows per fuel matches)
---------------------------------------------------------------------------------------------------------------------
Each cell is the payout for TATA & Ashok Leyland, for Old business (age 1+), without Nil Dep.  The notes under each
sheet say how the other cases move; the script READS those notes (they differ by region and Royal rewrites them from month
to month) and applies them.  Each rule is recognised in several wordings ("3% lower", "reduced by 3", "deduction of three",
"w.e.f. 22 Oct 2026" ...).  The words a rule understands are used up; any text that still holds a number afterwards is either a
region-specific discount note or something this script could not read.  Those are NEVER guessed at and NEVER stop the run (see
"Manual review" below):
  * Eicher is a few points lower, every other make lower still        ("For Eicher Make - payout will be 3% lower", ...)
  * Nil Dep business (is_zd = YES) is 2 points lower than without     ("Business sourced with Nil Dep - payout will be 2% lower")
  * New CV (age 0-1) is lower than Old by the note's cut - which can differ by weight ("upto 12 ton 2% lower, above 12 ton
    1% lower" / "same for old & new") - but a cell of 10% or less stays as it is
  * TATA & AL get 2 points extra when the discount is below 81% (85% in EAST): two discount bands, tariff 0-81.01 and
    81.01-86.01 (EAST 0-85.01 and 85.01-90.01); Eicher and the other makes have one band, up to 81% (EAST 85%)
  * a zero cell stays zero everywhere; nothing goes below zero; po_net_rate = pi_net_rate - 7 (zero stays zero)
All four fuels (Petrol, Diesel, CNG, LPG) get the same rates.  The grid says "NON EV"; Electric is not in it.
If a grid ever carries explicit "New" rows next to the "Old" ones, those are used for New instead of the New CV rule.
"This is effective from <date>" is compared with the from_date you give (rows are flagged for manual review if they differ);
"The existing grid continues" is read too - the UNCHANGED_SHEETS setting says whether such a sheet is written (include / skip / ask).

Manual review (the remarks column)
----------------------------------
A note or text rule this script cannot turn into logic - a region-specific discount, a sentence with a number it does not
understand, a rule the notes do not give, a New CV reduction whose weight limits do not clearly cover a band - is not
applied and not guessed at.  Instead the rows it may affect carry  Manual Review Required: [the note text]  in the `remarks`
column (the column already in the upload format), several notes joined with " | ".  Which rows: the clusters the note names
(a state brings in all its clusters; a place the script cannot place, such as NCR, or no place at all, means the whole
sheet) and the make groups it names.  Filter that column, take the notes up with Royal, and map those rates by hand.
Where a rule is missing the neutral value is used (no reduction, no extra, one discount band up to TARIFF_MAX).

Interactive
-----------
The script checks the workbook while it reads it.  Whenever something about the GRID it needs is missing, unclear or odd - a
sheet or header that is not where it should be, a weight band it cannot read, an age label it does not know, a cell that is
not a rate, dates that look wrong, a file that will not open or write - it stops, asks
you a specific question (type the number of an option, or q to stop, or ? for the details), remembers the answer and
carries on.  Every answer is listed at the end of the run, and you confirm before the file is written.  It never asks you
for a rate.  With --no-prompt (or when there is no keyboard) it asks nothing and stops with the question as the error
message instead.  If something goes wrong that it did not foresee it says so in one line and saves the technical details
to a *_error_*.txt file instead of printing a traceback.

Google Colab
------------
Paste this whole file into one cell (or upload it and `%run rsa_gcv4w_comp_to_config.py`) and run it.
Upload the RSA workbook (and, optionally, an earlier upload CSV to get a diff report).  It then asks
for from_date and to_date (DD/MM/YYYY), writes them into every row, and downloads RSA_GCV4W_COMP_UPLOAD_<yyyy-mm>.csv.

Command line
------------
    python rsa_gcv4w_comp_to_config.py [--grid <workbook.xlsx>] [--from-date 01/10/2026 --to-date 31/03/2030] [--ref <old upload.csv>] [--prefix ROYAL_SEPT26_] [--out file.csv] [--no-prompt]
(anything left out - the workbook, the dates, the earlier upload - is asked for)

new_rto_list is RTO_PREFIX + the cluster name with spaces turned into underscores (see the settings).
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
PRODUCT, SUB_PRODUCT = "GCV 4W", "1+1"
FUELS = ["Petrol", "Diesel", "CNG", "LPG"]      # the grid is "NON EV"; the same rate goes to each of these
# new_rto_list = RTO_PREFIX + cluster name (spaces and commas -> underscores). The portal's RTO groups were created with the
# September prefix and are still used; change it here (or pass --prefix) if new groups are ever created.
RTO_PREFIX = "ROYAL_SEPT26_"
# cluster name in the grid -> the part after the prefix, where the portal's group is not just "spaces -> underscores"
CLUSTER_NAME_EXCEPTIONS = {"DADRA & NAGAR HAVELI": "DADRA_AND_NAGAR_HAVELI"}
USER_ID = 2966
ROW_ID = 1                                       # the export carries id = 1 on every row
PI_TYPE = "On Net"
PO_CUT = 7.0                                     # po_net_rate = pi_net_rate - PO_CUT, never below 0

MAKE_TATA_AL = "TATA,Ashok"                      # the make / model master spells it "Ashok" (the October upload had "ASHOK")
MAKE_EICHER = "EICHER"
MAKE_OTHERS = "GCV_4W_ALL_EX_ASHOK_TATA_EICHER"
OLD_AGE, NEW_AGE = (1, 99), (0, 1)               # vehicle_age_min / max of Old and New business
KG_OPEN_END = 99999                              # cc_max of the top weight band
HEADER_SCAN_ROWS = 20                            # the "Tonnage Band" header is looked for in these rows
ZERO_WORDS = ("NA", "N/A", "-", "--", "NIL", "BLOCK", "BLOCKED", "NOT APPLICABLE")
# A sheet whose note says "the existing grid continues" (no new rates): "include" it anyway with the dates you typed,
# "skip" it, or "ask" each time.
UNCHANGED_SHEETS = "include"

TARIFF_MAX = 100.01                              # upper discount limit used when the notes give none
MAX_NOTE_CHARS = 500                             # a note longer than this is cut short in the remarks column

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
    """A sheet that cannot be used at all (no header, no clusters, no weight bands): the caller leaves it out or asks."""


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
    """A label without spaces, punctuation and capitals: 'Tonnage Band' and 'tonnage  band' are the same."""
    return re.sub(r"[^a-z0-9]+", "", _text(s).lower())


def is_rate_cell(v):
    return (isinstance(v, (int, float)) and not isinstance(v, bool)) or (isinstance(v, str) and v.strip().upper() in ZERO_WORDS)


# ---- weight bands ("A. 0 to 2.3 Ton" -> cc_min / cc_max in kg)
def _limits(label):
    """'A. 0 to 2.3 Ton' -> (0, 2.3, 1000); '48+ Ton' -> (48, None, 1000); 'Upto 2 T' -> (0, 2, 1000); None if unreadable.
    The last number is the multiplier to kg (1000 for tonnes, 1 when the label says kg)."""
    t = re.sub(r"^\s*[A-Za-z]\s*[.)]\s+", "", str(label))
    scale = 1.0 if re.search(r"\bkgs?\b", t, re.I) else 1000.0
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:to|-|–|—)\s*(\d+(?:\.\d+)?)", t, re.I)
    if m:
        lo, hi = float(m.group(1)), float(m.group(2))
        return (lo, hi, scale) if hi > lo else None
    m = (re.search(r"(\d+(?:\.\d+)?)\s*\+", t)
         or re.search(r"(\d+(?:\.\d+)?)\s*(?:tons?|tonnes?|t|kgs?)?\s*(?:and\s+|&\s+|or\s+)?(?:above|over|more|plus|onwards?|upwards?)\b", t, re.I)
         or re.search(r"(?:above|over|more than|greater than|>=|>)\s*(\d+(?:\.\d+)?)", t, re.I))
    if m:
        return float(m.group(1)), None, scale
    m = re.search(r"(?:upto|up to|below|under|less than|<=|<)\s*(\d+(?:\.\d+)?)", t, re.I)
    if m:
        return 0.0, float(m.group(1)), scale
    return None


def resolve_band(label, ask):
    """(cc_min, cc_max, low_kg, high_kg or None) of a weight band label; asks once if the label cannot be read."""
    got = _limits(label)
    if got is None:
        def check(a):
            g = _limits(a)
            if g is None:
                raise ValueError("Type it like 7.5-12 (tonnes), 0-2.3, or 48+ for no upper limit.")
            return g
        key = ("band", label_key(label))
        if key not in ask.memory:
            ask.memory[key] = ask.text(f"cannot read the weight band '{label}'",
                                       question=f"I can't read the weight limits in the band '{_text(label)}'. "
                                                f"Type its limits in tonnes, like 7.5-12 (or 48+ for no upper limit):",
                                       validate=check)
            ask.decide(f"weight band '{_text(label)}' -> limits typed by you")
        got = ask.memory[key]
    lo, hi, scale = got
    lo_kg, hi_kg = lo * scale, (hi * scale if hi is not None else None)
    cc_min = 0 if lo_kg == 0 else round(lo_kg + 0.01, 2)
    cc_max = round(hi_kg + 0.01, 2) if hi_kg is not None else KG_OPEN_END
    return cc_min, cc_max, lo_kg, hi_kg


# ---- the rules written in the notes under each sheet
# The notes are plain text that Royal rewrites from month to month. Every note (one cell) is run through the patterns
# below; the words a pattern understands are used up. Whatever is left that still holds a number is either a
# region-specific discount note or something this script did not understand. Neither is guessed at and neither stops the
# run: the rows they may affect get "Manual Review Required: [the note]" in the remarks column, so they can be taken up
# with Royal and mapped by hand afterwards.
# Notes (or what is left of a note after the patterns below used up their words) that need no action: they say how the grid is built and the script builds its rows the same way.  Anything else
# that holds words and no number is flagged like a note with a number (the user, 18 Nov 2026: 'Eicher make is not allowed in Gujarat' was dropped in silence).
STANDING_STRICT = True      # a standing phrase covers its own words only (the shared unknown_notes, 18 Nov 2026)
STANDING_NOTES = [
    "Note",
    "The above payout grid is applicable for TATA & AL makes",
    "The above payout grid (old) is applicable for business sourced without Nil Dep",
    "Business sourced",
    "No other payout will provided apart from the above mentioned grid",
]
STANDING_WHY = {
    "Note": "the heading of the notes",
    "The above payout grid is applicable for TATA & AL makes": "applied by the layout: the grid's cell is the TATA & AL payout, the other makes are worked out from it with the notes' cuts",
    "The above payout grid (old) is applicable for business sourced without Nil Dep": "applied by the layout: the grid's Old cell is the payout without Nil Dep, the Nil Dep rows take the note's cut",
    "Business sourced": "what is left of 'Business sourced with Nil Dep - payout will be N% lower' after that rule was applied",
    "No other payout will provided apart from the above mentioned grid": "a statement that the grid is all there is: nothing to apply",
}
_N = r"\d+(?:\.\d+)?"
_TON = r"(?:tons?|tonnes?|tonnages?|t)\b"
_LOWERS = r"lower|less|lesser|reduced|reduction|deducted|deduction|lowered|decreased|decrease"
_HIGHERS = r"higher|more|extra|increased|increase|additional|added|plus"
# an adjustment, however it is worded: "2% lower", "lower by 2%", "reduced by 2", "deduction of 2%", "extra 2%"
_ADJ = (rf"(?:(?P<n1>{_N})\s*%\s*(?P<d1>{_LOWERS}|{_HIGHERS})\b"
        rf"|(?P<d2>{_LOWERS}|{_HIGHERS})\b\s*(?:by|of)?\s*(?P<n2>{_N})\s*%?)")
_THAN = r"(?:\s+than\s+(?:the\s+)?(?:current|old|existing|above)\s+grid)?"
_DATE = r"(?P<d>\d{1,2})\s*(?:st|nd|rd|th)?(?:\s+|\s*[-/.]\s*)(?P<m>[a-z]{3,9}|\d{1,2})(?:\.?\s*,?\s*|\s*[-/.]\s*)(?P<y>\d{4})"
_WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
          "eleven": 11, "twelve": 12, "fifteen": 15, "twenty": 20}

# "All new CV [upto 12 ton] - payout will be 2% lower than the current grid"
P_NEW = re.compile(rf"all\s+new\s+cv(?:\s+(?:up\s*to|upto|below|under|less\s+than)\s*(?P<x>{_N})\s*{_TON})?\s*[-–:,]*\s*"
                   rf"(?:the\s+)?(?:payout\s+)?(?:will\s+be\s+|is\s+|shall\s+be\s+|would\s+be\s+)?{_ADJ}{_THAN}", re.I)
# "For GVW above 12 ton - payout will be 1% lower than the current grid"  /  "... the payout will be same for old & new"
P_NEW_ABOVE = re.compile(rf"for\s+(?:gvw\s+|weight\s+)?(?:above|over|more\s+than|greater\s+than)\s*(?P<x>{_N})\s*{_TON}\s*[-–:,]*\s*"
                         rf"(?:the\s+)?(?:payout\s+)?(?:will\s+be\s+|is\s+)?(?:{_ADJ}{_THAN}"
                         rf"|(?:the\s+)?(?P<same>same)(?:\s+for\s+old\s*(?:&|and)\s*new)?)", re.I)
# "(If the current payout is less than 10% then same payout as current grid will be maintained)"
P_KEEP = re.compile(rf"\(?\s*(?:if\s+)?(?:the\s+)?current\s+payout\s+is\s+(?:less\s+than|below|under|upto|up\s+to)\s*(?P<t>{_N})\s*%\s*,?\s*"
                    rf"(?:then\s+)?(?:the\s+)?same\s+payout(?:\s+as\s+(?:the\s+)?(?:current|old|existing)\s+grid)?"
                    rf"(?:\s+will\s+be\s+maintained)?\s*\)?", re.I)
# "For Eicher Make - payout will be 3% lower"  /  "For other makes - payout will be 5% lower"  /  "For Eicher make - deduction of 3%"
P_MAKE = re.compile(rf"for\s+(?P<who>[a-z][a-z &,/]*?)\s+makes?\s*[-–:,]*\s*(?:the\s+)?(?:payout\s+)?(?:will\s+be\s+|is\s+)?{_ADJ}{_THAN}", re.I)
# "Deduction of 3% for Eicher make"
P_MAKE2 = re.compile(rf"(?:payout\s+)?(?:deduction|reduction)\s+of\s+(?P<n>{_N})\s*%?\s+(?:for|on)\s+(?P<who>[a-z][a-z &,/]*?)\s+makes?", re.I)
# "Business sourced with Nil Dep - payout will be 2% lower"
P_ZD = re.compile(rf"\bwith\s+nil\s*dep\w*\s*[-–:,]*\s*(?:the\s+)?(?:payout\s+)?(?:will\s+be\s+|is\s+)?{_ADJ}{_THAN}", re.I)
# "Discounts for Tata & AL makes will be at max 86% for all tonnages, for discount less than 81% additional 2% payout"
P_TATA = re.compile(rf"discounts?\s+for\s+tata[^.%]*?at\s+max(?:imum)?\s*(?P<mx>{_N})\s*%[^.%]*?(?:less\s+than|below|under)\s*(?P<lo>{_N})\s*%\s*,?\s*"
                    rf"additional\s*(?P<ex>{_N})\s*%\s*(?:payout)?", re.I)
# "Discounts for Eicher and others will be at max 81%, no additional payouts applicable"
P_OTHER_MAX = re.compile(rf"discounts?\s+for\s+eicher\s*(?:&|and)\s*others?(?:\s+makes?)?\s+will\s+be\s+at\s+max(?:imum)?\s*(?P<mx>{_N})\s*%"
                         rf"(?:\s*,?\s*no\s+additional\s+payouts?\s+(?:applicable|will\s+be\s+(?:given|provided)))?", re.I)
# "This is effective from 22nd Oct 2026"  /  "The existing 1st Oct 2026 grid continues"  /  "valid till 31 Mar 2027"
P_EFF = re.compile(rf"(?:this\s+(?:is\s+)?)?(?:(?:effective|with\s+effect|w\.?e\.?f\.?)\s*(?:from\s+|date\s*[-:]?\s*)?|(?:applicable|valid)\s+from\s+)(?:the\s+)?{_DATE}", re.I)
P_CONT = re.compile(rf"(?:the\s+)?(?:existing|current|previous|old|earlier)\s+(?:{_DATE}\s+)?grid\s+(?:continues?|remains?|will\s+continue|stays?|is\s+continued)", re.I)
P_VALID = re.compile(rf"valid\s+(?:till|until|upto|up\s+to|through)\s+(?:the\s+)?{_DATE}", re.I)

_MAKE_KEY = {"eicher": "eicher_cut", "tata": "tata_cut", "al": "tata_cut", "ashok": "tata_cut", "leyland": "tata_cut",
             "other": "other_cut", "others": "other_cut"}
_FILLER = {"and", "the", "all", "of", "make", "makes", "for"}


def _who_keys(who):
    """{'eicher_cut'} for 'Eicher', {'other_cut'} for 'other', {'tata_cut'} for 'Tata & AL' ... None if a word is not a known make group."""
    words = [w for w in re.findall(r"[a-z]+", who.lower()) if w not in _FILLER]
    keys = {_MAKE_KEY.get(w) for w in words}
    return None if (not words or None in keys) else keys


def _adj(m):
    """The signed cut of a P_... match: positive = points taken off the grid cell, negative = added."""
    first = m.group("n1") is not None
    n, d = (m.group("n1"), m.group("d1")) if first else (m.group("n2"), m.group("d2"))
    return float(n) if re.fullmatch(_LOWERS, d, re.I) else -float(n)


def _normalise(item):
    """One note cell as a single line, with "three percent" / "3 points" / "of three" turned into "3%" / "of 3"."""
    t = re.sub(r"\s+", " ", _text(item)).strip()
    words = "|".join(_WORDS)
    t = re.sub(rf"\b({words})\b(?=\s*(?:%|percent|per\s*cent|points?|pts?)\b)", lambda m: str(_WORDS[m.group(1).lower()]), t, flags=re.I)
    t = re.sub(rf"\b(of|by)\s+({words})\b", lambda m: f"{m.group(1)} {_WORDS[m.group(2).lower()]}", t, flags=re.I)
    return re.sub(rf"({_N})\s*(?:percent|per\s*cent|pct|ppt|points?|pts?)\b", r"\1%", t, flags=re.I)


def _note_date(m):
    sep = "/" if m.group("m").isdigit() else " "
    try:
        return parse_ddmmyyyy(f"{m.group('d')}{sep}{m.group('m')}{sep}{m.group('y')}")
    except ValueError:
        return None


def _take(pattern, state, accept=None):
    """The matches of `pattern` in the part of a note not yet understood; those characters are then used up."""
    hits = []
    for m in pattern.finditer("".join(state)):
        if accept is not None and not accept(m):
            continue
        hits.append(m)
        for i in range(m.start(), m.end()):
            state[i] = " "
    return hits


RULE_TEXT = {
    "keep_below": "the payout at or below which New CV stays unreduced",
    "eicher_cut": "the payout reduction for Eicher make",
    "other_cut": "the payout reduction for other makes",
    "zd_cut": "the payout reduction for Nil Dep business",
    "tata_discount": "the maximum discount and the extra payout for Tata & AL makes",
    "other_max": "the maximum discount for Eicher and other makes",
}
# which output rows go without a rule the notes did not give: the arguments of rows_where()
MISSING_SCOPE = {
    "keep_below": {"new_only": True}, "eicher_cut": {"makes": {MAKE_EICHER}}, "other_cut": {"makes": {MAKE_OTHERS}},
    "zd_cut": {"zd": {"YES"}}, "tata_discount": {"makes": {MAKE_TATA_AL}}, "other_max": {"makes": {MAKE_EICHER, MAKE_OTHERS}},
}
_STATE_ABBR = {"MP": ("madhyapradesh",), "UP": ("uttarpradesh", "eastup", "westup"), "RJ": ("rajasthan",), "HR": ("haryana",),
               "PB": ("punjab",), "TN": ("tamilnadu",), "AP": ("andhrapradesh",), "KA": ("karnataka",), "KL": ("kerala",),
               "WB": ("westbengal",), "MH": ("maharashtra", "maharastra"), "GJ": ("gujarat",), "OD": ("odisha",),
               "JH": ("jharkhand",), "BR": ("bihar",), "CG": ("chattisgarh", "chhattisgarh"), "TS": ("telangana",)}


class Reading:
    """What one sheet's notes say - and what in them could not be turned into logic."""

    def __init__(self):
        self.rules, self.effective, self.continues = {}, None, False
        self.region_notes, self.unread, self.missing = [], [], []
        self.words = []                    # (note text, what is left of it) of the notes whose leftover holds words but no number: build() asks unknown_notes whether they are standing
        self.new_spec = {"main": None, "above": None, "text": ""}


def _believable(key, v):
    return -50 <= v <= 100 if key.endswith("_cut") else 0 <= v <= 100


def read_rules(items, geo_words=()):
    """Read the notes of one sheet (a list of note texts, one per cell) into a Reading. Nothing is asked and nothing is guessed:
    a rule the notes do not give - or give in words that are not understood - gets a neutral value (no reduction, no extra,
    no discount band) and its name goes to .missing; text with numbers that matches no rule goes to .unread, and a note
    about maximum discounts for particular places to .region_notes. The caller flags, in the remarks column, the rows each
    of these may affect."""
    out, got, new_texts = Reading(), {}, []
    new_main = new_above = None
    for item in items:
        text = _normalise(item)
        if not text:
            continue
        state = list(text)
        lead = re.match(r"\s*\d+\s*[.)]\s*", text)                       # the note's own number: "7."
        for i in range(lead.end() if lead else 0):
            state[i] = " "
        for m in _take(P_TATA, state):
            got.update(tata_max=float(m["mx"]), tata_below=float(m["lo"]), tata_extra=float(m["ex"]))
        for m in _take(P_OTHER_MAX, state):
            got["other_max"] = float(m["mx"])
        for m in _take(P_NEW, state):
            new_main = (float(m["x"]) if m["x"] else None, _adj(m))
            new_texts.append(text)
        for m in _take(P_NEW_ABOVE, state):
            new_above = (float(m["x"]), 0.0 if m["same"] else _adj(m))
            new_texts.append(text)
        for m in _take(P_KEEP, state):
            got["keep_below"] = float(m["t"])
        for m in _take(P_MAKE, state, lambda m: _who_keys(m["who"]) is not None):
            for key in _who_keys(m["who"]):
                got[key] = _adj(m)
        for m in _take(P_MAKE2, state, lambda m: _who_keys(m["who"]) is not None):
            for key in _who_keys(m["who"]):
                got[key] = float(m["n"])
        for m in _take(P_ZD, state):
            got["zd_cut"] = _adj(m)
        for m in _take(P_CONT, state):
            out.continues = True
        for m in _take(P_EFF, state, lambda m: _note_date(m) is not None):          # an impossible date stays unread
            out.effective = _note_date(m)
        _take(P_VALID, state)                                            # understood, not used
        left = re.sub(r"\s+", " ", "".join(state)).strip(" .,;:-–()")
        if left and not re.search(r"\d", left) and re.search(r"[A-Za-z]{3,}", left):
            out.words.append((text, left))                                   # 'Eicher make is not allowed in Gujarat': words, no number - it used to be dropped in silence (18 Nov 2026)
        if re.search(r"\d", left):
            words = set(re.findall(r"[a-z]+", left.lower()))
            geography = re.search(r"\b(?:rtos?|rto's|states?|clusters?|mp|ncr)\b", left, re.I) or words & set(geo_words)
            if geography and re.search(r"\bdis?count|\bdicount", left, re.I):    # "Discounts ... in the RTOs of MP & Rajasthan ..."
                out.region_notes.append(text)
            else:
                out.unread.append(text)
    got = {k: v for k, v in got.items() if _believable(k, v)}
    main = new_main if new_main and _believable("new_cut", new_main[1]) else None
    above = new_above if new_above and _believable("new_cut", new_above[1]) else None
    out.new_spec = {"main": main, "above": above, "text": " ".join(dict.fromkeys(new_texts))}
    rules = {"tata_cut": got.get("tata_cut", 0.0)}
    for k in ("eicher_cut", "other_cut", "zd_cut"):
        rules[k] = got.get(k, 0.0)
        if k not in got:
            out.missing.append(k)
    rules["keep_below"] = got.get("keep_below", 0.0)
    if "keep_below" not in got and (main or above):
        out.missing.append("keep_below")
    if all(k in got for k in ("tata_max", "tata_below", "tata_extra")) and got["tata_below"] < got["tata_max"]:
        rules.update(tata_max=got["tata_max"], tata_below=got["tata_below"], tata_extra=got["tata_extra"])
    else:
        rules.update(tata_max=None, tata_below=None, tata_extra=0.0)
        out.missing.append("tata_discount")
    rules["other_max"] = got.get("other_max")
    if rules["other_max"] is None:
        out.missing.append("other_max")
    out.rules = rules
    return out


def new_cut_for(spec, lo_kg, hi_kg):
    """(New CV cut of one weight band, why that band is not reduced - or None). Only what the note says is applied: a band the
    note does not clearly cover gets no cut and a reason, which the caller turns into a manual-review flag."""
    main, above = spec["main"], spec["above"]
    if not main and not above:
        return 0.0, "the New CV payout reduction was not found in the notes, so New is not reduced"
    if main and main[0] is None and not above:
        return main[1], None

    def within(limit_t):                                                 # the band lies at or below the limit
        return hi_kg is not None and hi_kg <= limit_t * 1000 + 0.5

    def beyond(limit_t):                                                 # the band lies at or above the limit
        return lo_kg >= limit_t * 1000 - 0.5

    if main and above:
        a, b = sorted((main[0] if main[0] is not None else above[0], above[0]))
        if within(a):
            return main[1], None
        if beyond(b):
            return above[1], None
        return 0.0, f"this weight band crosses or falls between the weight limits of the New CV note ({a:g} / {b:g} ton), so New is not reduced"
    if main:
        if within(main[0]):
            return main[1], None
        return 0.0, f"the New CV note only gives a reduction for up to {main[0]:g} ton, so heavier bands are not reduced"
    if beyond(above[0]):
        return above[1], None
    return 0.0, f"the New CV note only gives a reduction for above {above[0]:g} ton, so lighter bands are not reduced"


def make_groups(rules):
    """[(make, points added to the grid cell, [(tariff_min, tariff_max, extra points)])] for one sheet's rules. Discount limits
    the notes did not give become one band up to TARIFF_MAX."""
    o_hi = round(rules["other_max"] + 0.01, 2) if rules["other_max"] is not None else TARIFF_MAX
    if rules["tata_max"] is not None:
        t_lo, t_hi = round(rules["tata_below"] + 0.01, 2), round(rules["tata_max"] + 0.01, 2)
        tata_tiers = [(t_lo, t_hi, 0.0), (0.0, t_lo, rules["tata_extra"])]
    else:
        tata_tiers = [(0.0, TARIFF_MAX, 0.0)]
    return [(MAKE_TATA_AL, -rules["tata_cut"], tata_tiers),
            (MAKE_EICHER, -rules["eicher_cut"], [(0.0, o_hi, 0.0)]),
            (MAKE_OTHERS, -rules["other_cut"], [(0.0, o_hi, 0.0)])]


def _pm(cut):
    return "0" if cut == 0 else (f"-{cut:g}" if cut > 0 else f"+{-cut:g}")


def rules_line(rd):
    rules, spec = rd.rules, rd.new_spec
    main, above = spec["main"], spec["above"]
    if main and above:
        a = main[0] if main[0] is not None else above[0]
        new = f"{_pm(main[1])} up to {a:g} T, {_pm(above[1])} above {above[0]:g} T"
    elif main:
        new = _pm(main[1]) if main[0] is None else f"{_pm(main[1])} up to {main[0]:g} T only"
    elif above:
        new = f"{_pm(above[1])} above {above[0]:g} T only"
    else:
        new = "not found"
    tata = f"TATA & AL {_pm(rules['tata_cut'])} | " if rules["tata_cut"] else ""
    disc = (f"TATA & AL: discount up to {rules['tata_max']:g}, +{rules['tata_extra']:g} below {rules['tata_below']:g}"
            if rules["tata_max"] is not None else "TATA & AL: discount limits not found")
    other = f"Eicher & others: discount up to {rules['other_max']:g}" if rules["other_max"] is not None else "Eicher & others: discount limit not found"
    when = f" | effective {rd.effective:%d/%m/%Y}" if rd.effective else ""
    when += " | note says the existing grid continues" if rd.continues else ""
    return (f"new CV {new} (cells of {rules['keep_below']:g} or less stay) | Nil Dep {_pm(rules['zd_cut'])} | {tata}"
            f"Eicher {_pm(rules['eicher_cut'])} | other makes {_pm(rules['other_cut'])} | {disc} | {other}{when}")


def note_scope(text, reg):
    """(clusters, make groups) a note may concern. Clusters: the ones it names (a state name or abbreviation brings in all of that
    state's clusters); if it names none the script can match - or names a place it cannot place, such as NCR - None = the
    whole sheet. Make groups: the ones it names, None = all. A superset is flagged, never a subset."""
    compact = note_key(text)
    hit = set()
    for name in reg.states:
        k = note_key(name)
        if len(k) >= 4 and k in compact:
            hit.add(name)
    for label in set(reg.states.values()):
        k = note_key(label)
        if len(k) >= 4 and k in compact:
            hit |= {c for c, s in reg.states.items() if s == label}
    for m in re.finditer(r"\b(" + "|".join(_STATE_ABBR) + r")\b", text):         # capitals only: "up to" is not UP
        hit |= {c for c, s in reg.states.items() if note_key(s) in _STATE_ABBR[m.group(1)]}
    clusters = None if (not hit or re.search(r"\b(?:NCR|NCT)\b", text)) else hit
    low = text.lower()
    makes = set()
    if re.search(r"\btata\b|\bal\b|ashok|leyland", low):
        makes.add(MAKE_TATA_AL)
    if re.search(r"\beicher\b", low):
        makes.add(MAKE_EICHER)
    if re.search(r"\bothers?\b", low):
        makes.add(MAKE_OTHERS)
    if re.search(r"\ball\s+makes?\b", low):
        makes = set()
    return clusters, (makes or None)


def rows_where(sheet, rto_names=None, makes=None, new_only=False, zd=None, bands=None):
    """A function of the output frame: the rows of one sheet, narrowed by cluster, make group, New business, Nil Dep, weight band."""
    def fn(df):
        m = np.array(df["_sheet"] == sheet, dtype=bool)
        if rto_names is not None:
            m = m & np.array(df["new_rto_list"].isin(rto_names), dtype=bool)
        if makes is not None:
            m = m & np.array(df["new_vehicle_makes"].isin(makes), dtype=bool)
        if new_only:
            m = m & np.array(df["add_tnc"] == "New", dtype=bool)
        if zd is not None:
            m = m & np.array(df["is_zd"].isin(zd), dtype=bool)
        if bands is not None:
            m = m & np.array(df["cc_min"].isin(bands), dtype=bool)
        return m
    return fn


# ---- one sheet
class Region:
    """One sheet: weight bands down the side (an Old row, sometimes a New row, per band) and clusters across."""

    def __init__(self, sheet, rows, head, ask, fuels, title_row=0):
        self.sheet, self.rows, self.ask, self.head, self.fuels = sheet, rows, ask, head, fuels
        header = rows[head]
        self.age_col = next((j for j, v in enumerate(header) if j > 0 and label_key(v) == "age"), None)
        self.default_kind = "old"
        if self.age_col is None:
            self.default_kind = ask.choose(
                f"'{sheet}': no 'Age' column in the header row {head + 1}",
                [("Every row is Old business (age 1+); New is worked out from the notes", "old"),
                 ("Every row is New business (age 0-1)", "new")],
                question=f"'{sheet}' has no 'Age' column. What do its rows hold?", short=f"{sheet}: no Age column",
                memo=("noage", sheet))
        first = (self.age_col if self.age_col is not None else 0) + 1
        self.region = (_text(rows[title_row][first]) if len(rows) > title_row and len(rows[title_row]) > first else "") or sheet
        self.bands, self.order = {}, []
        for r in range(head + 1, len(rows)):
            row = rows[r]
            label = _text(row[0]) if row else ""
            age_text = _text(row[self.age_col]) if (self.age_col is not None and row and self.age_col < len(row)) else ""
            if not label and age_text and self.order:                       # a band name merged over its Old / New rows
                label = self.order[-1]
            has_numbers = any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in row[first:])
            if not label or not (re.search(r"\d", label) or has_numbers):   # a band is named with numbers, or at least has some
                if self.order:
                    if self._band_follows(r + 1, first):
                        continue                                            # a blank row inside the table: the weight bands go on below it
                    break                                                   # the first blank or non-band row that no band follows ends the table
                continue
            kind = self._kind(age_text, r, label)
            if kind is None:
                continue
            slot = self.bands.setdefault(label, {})
            if kind in slot:
                keep = ask.choose(f"'{sheet}': the band '{label}' has two {kind.upper()} rows (rows {slot[kind] + 1} and {r + 1})",
                                  [("Use the first of them", "first"), ("Use the last of them", "last")],
                                  question=f"'{sheet}' has the band '{label}' twice for {kind.upper()} business (rows {slot[kind] + 1} "
                                           f"and {r + 1}). Which one should I use?",
                                  short=f"{sheet}: repeated '{label}' {kind.upper()} row", memo=("dupband", sheet, label, kind))
                if keep == "first":
                    continue
            slot[kind] = r
            if label not in self.order:
                self.order.append(label)
        if not self.order:
            raise EmptySheet(f"'{sheet}': no weight-band rows found under the header row {head + 1}")
        self.last_band_row = max(r for slot in self.bands.values() for r in slot.values())
        self.clusters, self.no_data = [], []
        for j, v in enumerate(header):
            if j < first or not _text(v):
                continue
            has_rates = any(j < len(rows[r]) and is_rate_cell(rows[r][j]) for slot in self.bands.values() for r in slot.values())
            (self.clusters if has_rates else self.no_data).append((j, _text(v)))
        if not self.clusters:
            raise EmptySheet(f"'{sheet}': no cluster column with rates under the header row {head + 1}")
        named = {j for j, _ in self.clusters} | {j for j, _ in self.no_data}
        self.unnamed = [j for j in range(first, max(len(rows[r]) for r in range(head, self.last_band_row + 1)))
                        if j not in named and any(j < len(rows[r]) and isinstance(rows[r][j], (int, float))
                                                  and not isinstance(rows[r][j], bool) and rows[r][j] != 0
                                                  for slot in self.bands.values() for r in slot.values())]
        above = rows[head - 1] if head > 0 else []
        self.states = {name: (_text(above[j]) if j < len(above) else "") for j, name in self.clusters}   # cluster -> state label
        notes = []                                                          # one item per cell: a note can run over several lines
        for r in range(self.last_band_row + 1, min(len(rows), self.last_band_row + 80)):
            for v in rows[r]:
                if isinstance(v, str) and _text(v):
                    notes.append(_text(v))
        self.notes = notes
        words = set()                                                       # state / cluster names: they tell a region note from other text
        for row in (rows[head - 1] if head > 0 else [], header):
            for v in row[first:]:
                words |= {w for w in re.findall(r"[a-z]+", _text(v).lower()) if len(w) >= 4}
        self.geo_words = words - {"rest", "india", "other", "others"}

    def _band_follows(self, start, first):
        """True when the next non-blank row from `start` is a weight-band row of this table (a band label and at least three numbers in the cluster columns): notes under the table are not."""
        for i in range(start, len(self.rows)):
            row = self.rows[i]
            if not any(_text(v) for v in (row or [])):
                continue
            label = _text(row[0]) if row else ""
            age_text = _text(row[self.age_col]) if (self.age_col is not None and self.age_col < len(row)) else ""
            numbers = sum(1 for v in row[first:] if isinstance(v, (int, float)) and not isinstance(v, bool))
            return bool((label or age_text) and numbers >= 3 and (label_key(age_text).startswith(("old", "new")) or self.age_col is None))
        return False

    def _kind(self, age_text, r, label):
        if self.age_col is None:
            return self.default_kind
        t = label_key(age_text)
        if t.startswith("old"):
            return "old"
        if t.startswith("new"):
            return "new"
        return self.ask.choose(
            f"'{self.sheet}' row {r + 1}: the age label {age_text!r} is neither Old nor New",
            [("Treat these rows as Old business (age 1+)", "old"), ("Treat these rows as New business (age 0-1)", "new"),
             ("Leave these rows out of the upload", None)],
            question=f"'{self.sheet}' row {r + 1} ('{label}') has the age label {age_text!r}, which is neither Old nor New. What is it?",
            short=f"age label {age_text!r}", memo=("age", t))

    def rate(self, r, j):
        """The cell as a payout in percent (0.34 in the grid -> 34)."""
        row = self.rows[r]
        v = row[j] if j < len(row) else None
        addr = f"'{self.sheet}'!{_col_letter(j)}{r + 1}"
        kind, x = read_number(self.ask, addr, v, "fraction", "a payout such as 0.34 (= 34%)", zero_words=ZERO_WORDS)
        return round(x * 100, 4) if kind == "num" else 0.0


def find_header(rows):
    """(row index, column index) of the weight-band header cell ('Tonnage Band'), wherever it sits in the first rows, or (None, 0)."""
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        for c0, v in enumerate(row or []):
            t = _text(v)
            k = label_key(t)
            short = len(t) <= 30                                                 # a note that mentions "tonnages" is not the header
            if t and short and ("tonnage" in k or "gvw" in k or "weightband" in k) and sum(1 for x in row[c0 + 1:] if _text(x)) >= 3:
                return i, c0
    return None, 0


UNUSED_SHEETS = []        # (sheet, status, why) of the sheets of this run that no region was read from: build() lists them in 'Unread parts' (find_regions fills it)


def find_regions(wb, ask, log):
    """Every sheet that holds a GCV 4W comprehensive table: its header row is found by the label 'Tonnage Band'."""
    regions, left_out, hidden = [], [], []
    for name in wb.sheetnames:
        if getattr(wb[name], "sheet_state", "visible") != "visible":
            hidden.append(name)                                                  # an old version nobody deleted is not this month's grid
            continue
        rows = sheet_rows(wb, name)
        head, c0 = find_header(rows)
        if head is not None and c0:
            rows = [list(r)[c0:] if r else r for r in rows]                      # the table is read from its own first column on
        looks = "gcv" in label_key(name) and "4w" in label_key(name)
        if head is None:
            if looks:
                pick = ask.choose(f"sheet '{name}' looks like a GCV 4W sheet but has no 'Tonnage Band' header",
                                  [("Tell me which row holds the header", "row"), ("Leave this sheet out", None)],
                                  question=f"The sheet '{name}' looks like a GCV 4W sheet, but I can't find a 'Tonnage Band' header "
                                           f"in its first {HEADER_SCAN_ROWS} rows. What should I do?",
                                  short=f"sheet '{name}'", memo=("nohead", name))
                if pick == "row":
                    head = ask.text(f"no header row given for '{name}'", question=f"Which row of '{name}' holds the header "
                                    f"(the one with the cluster names)? Type its row number:",
                                    validate=lambda a: row_number(a, len(rows)) - 1)
                    c0 = 0
            if head is None:
                left_out.append(name)
                continue
        above = [(i, _text(r[0])) for i, r in enumerate(rows[:head]) if r and _text(r[0])]      # the texts in the table's first column above its header: the title, then the fuel label
        title_row = above[-2][0] if len(above) >= 2 else (above[0][0] if above else 0)
        title = above[-2][1] if len(above) >= 2 else (above[0][1] if above else "")
        tk = label_key(title) + label_key(name)
        odd = [w for w, word in (("4W", "4w"), ("GCV", "gcv"), ("COMP (comprehensive)", "comp")) if word not in tk]
        if odd:
            use = ask.choose(f"sheet '{name}' (title {title!r}) does not say {', '.join(odd)}",
                             [("Use it anyway - it is a GCV 4W comprehensive sheet", True), ("Leave this sheet out", False)],
                             question=f"The sheet '{name}' has the title {title!r}, which does not say {', '.join(odd)}. "
                                      f"Is it a GCV 4W comprehensive sheet?", short=f"sheet '{name}' title", memo=("title", name))
            if not use:
                left_out.append(name)
                continue
        fuel_label = above[-1][1] if len(above) >= 2 else ""
        fk = label_key(fuel_label)
        fuels = FUELS
        if fk and "nonev" not in fk:
            fuels = ask.choose(f"sheet '{name}' is labelled {fuel_label!r}, not NON EV",
                               [(f"Write the usual fuels ({', '.join(FUELS)})", FUELS), ("Write Electric rows only", ["Electric"]),
                                ("Leave this sheet out", None)],
                               question=f"The sheet '{name}' says {fuel_label!r} where the other sheets say NON EV. Which fuels do its rows get?",
                               short=f"sheet '{name}' fuel", memo=("fuel", name))
            if fuels is None:
                left_out.append(name)
                continue
        try:
            regions.append(Region(name, rows, head, ask, fuels, title_row))
        except EmptySheet as e:
            pick = ask.choose(str(e), [("Leave this sheet out", None)],
                              question=f"{e}. I can't use this sheet. Leave it out?", short=f"sheet '{name}' unusable")
            left_out.append(name)
    if not regions:
        raise GridError("no GCV 4W comprehensive sheet found - looked for a header cell 'Tonnage Band' in the first "
                        f"{HEADER_SCAN_ROWS} rows of every visible sheet ({wb.sheetnames})")
    for i, a in enumerate(regions):                                              # two sheets with the same clusters = this month's grid and a copy of another month: which one is current is not for a script to guess
        for b in regions[i + 1:]:
            both = {n for _, n in a.clusters} & {n for _, n in b.clusters}
            if both:
                raise GridError(f"the sheets '{a.sheet}' and '{b.sheet}' both hold the clusters {sorted(both)[:4]}{' ...' if len(both) > 4 else ''} - expected one sheet per region; "
                                f"hide or delete the one that is not this month's grid")
    log(f"Sheets used: {[r.sheet for r in regions]}" + (f" | not used: {left_out}" if left_out else "") + (f" | hidden, not read: {hidden}" if hidden else ""))
    UNUSED_SHEETS[:] = [(n, "NOT READ", "no GCV 4W comprehensive table (a 'Tonnage Band' header) was found on this sheet, or you left it out") for n in left_out] + \
                       [(n, "left out on purpose", "a hidden sheet (an older version nobody deleted): never read") for n in hidden]
    return regions


# %% ---------------------------------------------------------------- BUILD
def rto_name(cluster):
    """new_rto_list of a cluster: the prefix + the grid's name with spaces -> underscores (see CLUSTER_NAME_EXCEPTIONS).
    A comma becomes an underscore too: new_rto_list is itself a comma-separated list, so a comma inside a name would split it."""
    special = {label_key(k): v for k, v in CLUSTER_NAME_EXCEPTIONS.items()}
    return rto_cluster(RTO_PREFIX + special.get(label_key(cluster), cluster.replace(",", "_").replace(" ", "_")), cluster)      # the CSV carries the grid's cluster name (apply_cluster_names)


def _row(rto, make, fuel, age, cc, zd, tnc, tariff, rate, dates, sheet=""):
    return dict(
        insurance_company=INSURANCE_COMPANY, insurer_vertical="Not Applicable", product=PRODUCT, sub_product=SUB_PRODUCT,
        policy_type="Renew", fuel_type=fuel, make_model_class="NA", new_vehicle_makes=make, new_rto_list=rto,
        vehicle_age_min=age[0], vehicle_age_max=age[1], cc_min=cc[0], cc_max=cc[1], sc_min=0, sc_max=999,
        pi_od_rate=0.0, pi_tp_rate=0.0, pi_tp_2=0, pi_tp_3=0, pi_tp_4=0, pi_tp_5=0, pi_net_rate=rate, pi_flat_amount=0,
        pi_vli=0, pi_type=PI_TYPE, tariff_min=tariff[0], tariff_max=tariff[1], is_ncb="NA", is_cpa="NA", is_zd=zd,
        from_date=dates[0], to_date=dates[1], user_id=USER_ID, veh_use="", add_tnc=tnc, remarks="", _sheet=sheet,
    )


def cell_variants(old_base, new_base, adj, extra, rules, new_cut):
    """(Old with Nil Dep, Old without, New) payouts in percent for one cell, one make and one discount band.
    The make's points are added to the grid cell, the extra points only when the cell is above zero; New is the Old
    payout minus new_cut (the cut of this weight band) unless the cell is at or below keep_below; Nil Dep takes zd_cut
    off; nothing goes below zero."""
    old = max(old_base + adj, 0.0) + (extra if old_base > 0 else 0.0)
    zd = max(old - rules["zd_cut"], 0.0)
    if new_base is None:
        core = max(old_base + adj, 0.0)
        core = max(core - new_cut, 0.0) if old_base > rules["keep_below"] else core
        new = core + (extra if old_base > 0 else 0.0)
    else:
        new = max(new_base + adj, 0.0) + (extra if new_base > 0 else 0.0)
    return zd, old, new


def _short(text):
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text if len(text) <= MAX_NOTE_CHARS else text[:MAX_NOTE_CHARS].rstrip() + "..."


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
    regions = find_regions(wb, ask, log)
    log(f"new_rto_list = {RTO_PREFIX}<cluster name>   (RTO_PREFIX in the settings, or --prefix)")
    warnings, rows, band_sets, review = [], [], {}, ManualReview()
    for name, status, why in UNUSED_SHEETS:                              # sheets of the workbook that no region was read from: in the review workbook, and a warning for a visible one
        unread_sheet(name, sheet_rows(wb, name), status, why, warnings)

    # ---- what the notes of every sheet say
    readings = {reg.sheet: read_rules(reg.notes, reg.geo_words) for reg in regions}
    for reg in list(regions):
        rd = readings[reg.sheet]
        if rd.continues:
            mode = UNCHANGED_SHEETS
            if mode == "ask":
                mode = ask.choose(f"'{reg.sheet}': its note says the existing grid continues",
                                  [("Write this sheet anyway, with the dates you gave", "include"), ("Leave this sheet out (its rows already exist)", "skip")],
                                  question=f"The note of '{reg.sheet}' says the existing grid continues, so its rates may not be new. What should I do?",
                                  short=f"{reg.sheet}: 'existing grid continues'", memo=("continues", reg.sheet))
            if mode == "skip":
                log(f"'{reg.sheet}': its note says the existing grid continues - left out")
                regions.remove(reg)
            else:
                log(f"'{reg.sheet}': its note says the existing grid continues - written anyway (UNCHANGED_SHEETS = '{UNCHANGED_SHEETS}')")
    if not regions:
        raise GridError("nothing to write - every sheet was left out")

    for reg in regions:
        rd = readings[reg.sheet]
        rules = rd.rules
        groups = make_groups(rules)
        bands = [(label, resolve_band(label, ask)) for label in reg.order]
        ordered = sorted((b for _, b in bands), key=lambda b: b[2])
        problems = []
        if ordered[0][2] != 0:
            problems.append(f"the lowest band starts at {ordered[0][2] / 1000:g} T, not at 0")
        for a, b in zip(ordered, ordered[1:]):
            if a[3] is None or abs(a[3] - b[2]) > 0.5:
                problems.append(f"a band ends at {('open' if a[3] is None else f'{a[3] / 1000:g} T')} but the next starts at {b[2] / 1000:g} T")
        if ordered[-1][3] is not None:
            problems.append(f"the highest band stops at {ordered[-1][3] / 1000:g} T (no 'and above' band)")
        if problems and not ask.choose(
                f"'{reg.sheet}': the weight bands do not join up: {'; '.join(problems)}",
                [("Carry on with the bands as they are", True), ("Stop so I can look at the grid", False)],
                question=f"In '{reg.sheet}' the weight bands do not join up ({'; '.join(problems)}). What should I do?",
                short=f"{reg.sheet}: weight bands", memo=("bandjoin", reg.sheet)):
            raise GridError(f"stopped by you - '{reg.sheet}' weight bands do not join up")
        band_sets[reg.sheet] = {(b[0], b[1]) for _, b in bands}
        n_new = sum(1 for label in reg.order if "new" in reg.bands[label])
        log(f"'{reg.sheet}' ({reg.region}): {len(reg.clusters)} clusters x {len(reg.order)} weight bands, fuels {', '.join(reg.fuels)}"
            + (f"; New rows in the grid for {n_new} band(s), New worked out from the notes for the rest" if n_new else
               "; Old rows only - New worked out from the notes"))
        log(f"    rules from the notes: {rules_line(rd)}")
        for j, name in reg.no_data:
            warnings.append(f"'{reg.sheet}' column {_col_letter(j)} ('{name}') has a header but no rates - left out")
        for j in reg.unnamed:
            warnings.append(f"'{reg.sheet}' column {_col_letter(j)} has rates but no cluster name in the header - its rates are NOT in the output")

        # ---- what in the notes could not be applied: flagged on the rows it may affect, never guessed
        rto_of = lambda clusters: None if clusters is None else {rto_name(c) for c in clusters}
        for text in rd.region_notes:                                     # a different maximum discount for particular places / weights
            clusters, makes = note_scope(text, reg)
            review.flag(_short(text), rows_where(reg.sheet, rto_of(clusters), makes))
        for text in rd.unread:                                           # text with a number that matches no rule
            clusters, _ = note_scope(text, reg)
            review.flag(_short(text), rows_where(reg.sheet, rto_of(clusters), None))
        for text, left in rd.words:                                      # words but no number: standing (the layout of the grid) or flagged - never dropped in silence
            if unknown_notes([left], STANDING_NOTES, min_len=4):
                clusters, makes = note_scope(text, reg)
                review.flag(_short(text), rows_where(reg.sheet, rto_of(clusters), makes))
        for key in rd.missing:                                           # a rule the notes did not give
            review.flag(f"{reg.sheet}: {RULE_TEXT[key]} not found or not readable in the notes, so it was not applied",
                        rows_where(reg.sheet, **MISSING_SCOPE[key]))
        if rd.effective and rd.effective != from_date and not rd.continues:
            review.flag(f"{reg.sheet}: the note says it is effective from {rd.effective:%d/%m/%Y} but these rows carry from_date {from_date:%d/%m/%Y}",
                        rows_where(reg.sheet))
        band_reasons = {}
        for label, (cc_min, cc_max, lo_kg, hi_kg) in bands:
            slot = reg.bands[label]
            if "old" not in slot:
                ask.choose(f"'{reg.sheet}': the band '{label}' has no Old row, so its other cases cannot be worked out",
                           [("Leave this weight band out of the upload", None)],
                           question=f"'{reg.sheet}': the band '{label}' has only a New row. I can't work out the Old and Nil Dep rows "
                                    f"from it. Leave this weight band out?", short=f"{reg.sheet}: band '{label}' without Old row",
                           memo=("noold", reg.sheet, label))
                continue
            new_cut, reason = new_cut_for(rd.new_spec, lo_kg, hi_kg)
            if reason:
                band_reasons.setdefault(reason, set()).add(cc_min)
            for j, cluster in reg.clusters:
                rto = rto_name(cluster)
                old_base = reg.rate(slot["old"], j)
                new_base = reg.rate(slot["new"], j) if "new" in slot else None
                for make, adj, tiers in groups:
                    for t_lo, t_hi, extra in tiers:
                        zd, old, new = cell_variants(old_base, new_base, adj, extra, rules, new_cut)
                        for fuel in reg.fuels:
                            for age, z, tnc, rate in ((OLD_AGE, "YES", "Old", zd), (OLD_AGE, "NO", "Old", old), (NEW_AGE, "NA", "New", new)):
                                rows.append(_row(rto, make, fuel, age, (cc_min, cc_max), z, tnc, (t_lo, t_hi), rate, dates, reg.sheet))
        for reason, cc_mins in band_reasons.items():                     # weight bands the New CV note does not clearly cover
            note = f" Note: {rd.new_spec['text']}" if rd.new_spec["text"] else ""
            review.flag(_short(f"{reg.sheet}: {reason}.{note}"), rows_where(reg.sheet, new_only=True, bands=cc_mins))
    first = next(iter(band_sets))
    for sheet, bs in band_sets.items():
        if bs != band_sets[first]:
            warnings.append(f"'{sheet}' has different weight bands from '{first}' - check both sheets")
    if not rows:
        raise GridError("nothing to write - every part of the workbook was left out")
    df = pd.DataFrame(rows)
    df["id"] = ROW_ID
    df["po_type"] = df["pi_type"]
    df["po_net_rate"] = np.where(df["pi_net_rate"] > 0, (df["pi_net_rate"] - PO_CUT).clip(lower=0), 0.0)
    df["po_od_rate"] = 0.0
    df["po_tp_rate"] = 0.0
    df["po_flat_amount"] = 0
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
    counts = df.groupby(["new_vehicle_makes", "add_tnc", "is_zd"]).size()
    for (mk, tnc, zd), n in counts.items():
        log(f"  {mk:<33} {tnc:<4} Nil Dep {zd:<3} {n:>7,}")
    log(f"Duplicate keys: {int(df.duplicated(KEY_COLUMNS).sum())}")
    floored = int(((df['pi_net_rate'] > 0) & (df['po_net_rate'] == 0)).sum())
    log(f"PO floored to 0 (pi > 0 but pi <= {PO_CUT:g}): {floored}")
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
            notes.append("some notes or columns are not in the output (see WARNING above)")
        if diff and (diff[0] + diff[1]) > 0.25 * len(df):
            notes.append(f"{diff[0]} rows are new and {diff[1]} rows are gone compared with the earlier upload")
        question = ("Heads-up: " + "; ".join(notes) + ". Write the file anyway?" if notes
                    else "Everything look right? Write the file now?")
        if not ask.confirm(question, True):
            raise GridError("stopped by you - nothing was written")
    out_path = write_csv(to_csv_frame(df), out_path or default_output(f"RSA_GCV4W_COMP_UPLOAD_{from_date:%Y-%m}.csv"), ask, log)
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
        question="You uploaded more than one workbook. Which one is the RSA GCV 4W grid?", short="workbook")
    ref = None
    if csvs:
        ref = csvs[0] if len(csvs) == 1 else ask.choose(
            "more than one CSV was uploaded", [(n, n) for n in csvs] + [("None of them - skip the comparison", None)],
            question="You uploaded more than one CSV. Which one is the earlier upload to compare with?", short="earlier upload")
    return grid, ref


def colab_main():
    ask = Asker(interactive=True)
    say("Upload: the RSA GCV 4W workbook (.xlsx) and, optionally, an earlier upload (.csv) to compare with")
    try:
        grid, ref = collect_uploads(ask)
        say(f"\nFiles received: workbook = {grid}" + (f", reference = {ref}" if ref else ""))
        from_date, to_date = ask_date_range(ask=ask)
        out = main(grid, from_date, to_date, ref_path=ref, ask=ask)
    except (GridError, KeyboardInterrupt) as e:
        say(f"\nSTOPPED - {e if isinstance(e, GridError) else 'stopped by you'}\nNo file was written.")      # a plain message: raising SystemExit makes Colab print a long traceback
        return
    except Exception as e:
        saved = crash_report("rsa_gcv4w_comp_to_config")
        say(f"\nSTOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
            + (f"\nThe technical details are in {saved}" if saved else ""))
        return
    colab_download(out)


def cli():
    global RTO_PREFIX
    ap = argparse.ArgumentParser(description="RSA GCV 4W comprehensive grid -> RSA upload CSV")
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
                                  "Path of the RSA GCV 4W workbook (.xlsx) - you can drag the file into this window:",
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
        saved = crash_report("rsa_gcv4w_comp_to_config")
        sys.exit(f"STOPPED - something unexpected went wrong ({type(e).__name__}: {e}). No file was written."
                 + (f" The technical details are in {saved}" if saved else ""))


if __name__ == "__main__":
    colab_main() if IN_COLAB else cli()
