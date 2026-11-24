"""Checks of the CD1 age-group grammar and of the readers added on 10 Nov 2026 to godigit_agency_to_config.py (the user's rule: new wordings must be read automatically, never flagged for a person, and never guessed):
  * the CD1 grammar: every wording of the October and November grids read as written in a hand-checked table; the same readings as the earlier reader wherever the earlier reader understood a text (the only
    differences are the wordings in which a closed range and an open one name the same age); generated wordings in many spellings; everything that must be refused is refused;
  * the taxi make / age grammar, a cap per make, the text-typed fraction, the TATA cell without its dash, MISP as a recorded decision, the staff-bus seat note.
Run:  .venv\\Scripts\\python.exe tests\\godigit_cd1_tests.py"""
import datetime as dt
import os
import random
import re
import sys

import pandas as pd

os.environ["GRID_RTO_NAMES"] = "cluster"                 # the default of the script (new_rto_list = the cluster name): the flagged-row counts below are those of a real run
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

mod = helpers.load_script("godigit_agency_to_config.py")
F, T = dt.date(2026, 11, 1), dt.date(2030, 3, 31)
quiet = lambda *a, **k: None
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


show = lambda slots: " ".join(mod.cd1_show(s) for s in slots)


def read(t, cells=None):
    r = mod.cd1_read(t, cells)
    return None if r is None else show(r)


# ---- every non-trivial CD1 wording of the two real workbooks -> the cap of each age column (New, 1, 2, 3, 4, 5, 6-10, 11+), checked by hand against the text and the grid's payout cells
GOLDEN = {   # CD1 text -> the cap of each age column (New, 1, 2, 3, 4, 5, 6-10, 11+) as cd1_show writes it; checked by hand against the text and the grid's payout cells
    'New/1yr: 60% | 2-4yr: 85% | 5yr+: 85%': '60 60 85 85 85 85 85 85',
    'New/1yr: D | 2-5yr: 80% | 6yr+: 80%': 'D D 80 80 80 80 80 80',
    'New/1yr: D | 2-5yr: 85% | 6yr+: 85%': 'D D 85 85 85 85 85 85',
    'New/1yr: D | 2-5yr: D | 6yr+: 90%': 'D D D D D D 90 90',
    'New/1yr: D | 2-5yr: 80% | 6yr+: 90%': 'D D 80 80 80 80 90 90',
    'New/1yr: D | 2-5yr: 85% | 6yr+: 90%': 'D D 85 85 85 85 90 90',
    'New/1yr: D | 2-5yr: D | 6yr+: 80%': 'D D D D D D 80 80',
    'New/1yr: 60% | 2-5yr: 85% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': '60 60 85 85 85 85 ncb(90|80) ncb(90|80)',
    'New/1yr: 60% | 2-4yr: 85% | 5yr+: 90%': '60 60 85 85 85 90 90 90',
    'New/1yr: 60% | 2-5yr: 85% | 6yr+: 85%': '60 60 85 85 85 85 85 85',
    'New/1yr: D | 2-4yr: D | 5yr+: 85%': 'D D D D D 85 85 85',
    'New/1yr: D | 2-4yr: 80% | 5yr+: 85%': 'D D 80 80 80 85 85 85',
    'New/1yr: D | 2-4yr: 85% | 5yr+: 85%': 'D D 85 85 85 85 85 85',
    'New/1yr: D | 2-5yr: 40% | 6yr+: 85%': 'D D 40 40 40 40 85 85',
    'New/1yr: 60% | 2-4yr: 85% | 5yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': '60 60 85 85 85 ncb(90|80) ncb(90|80) ncb(90|80)',
    'New/1yr: 60% | 2-4yr: D | 5yr+: 90%': '60 60 D D D 90 90 90',
    'New/1yr: 60% | 2-4yr: 80% | 5yr+: 90%': '60 60 80 80 80 90 90 90',
    'New/1yr: 60% | 2-5yr: 60% | 6yr+: 85%': '60 60 60 60 60 60 85 85',
    'Age 0: 60% Age>=1: 75%': '60 75 75 75 75 75 75 75',
    'New: 60% | 1yr: 0%| 2-5yr: 0% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': '60 0 0 0 0 0 ncb(90|80) ncb(90|80)',
    'Age 0-55%; 1+ 80%': '55 80 80 80 80 80 80 80',
    'New/1yr: 60% | 2-5yr: 70% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': '60 60 70 70 70 70 ncb(90|80) ncb(90|80)',
    'New/1yr: D | 2-5yr: D | 6yr+: 85%': 'D D D D D D 85 85',
    'New/1yr: D | 2-5yr: 80% | 6yr+: 85%': 'D D 80 80 80 80 85 85',
    'New: With Addon: 85%/90% Without Addon: 60% | 1yr: 60% | 2-4yr: 85% | 5yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': 'addon(85/90|60) 60 85 85 85 ncb(90|80) ncb(90|80) ncb(90|80)',
    '60% for age 0-1, 85% for age 2-5, 90% for age 6+': '60 60 85 85 85 85 90 90',
    'New/1yr: 60% | 2-5yr: 80% | 6yr+: with NCB/break in > 90 days : 85%, without NCB : 75%': '60 60 80 80 80 80 ncb(85|75) ncb(85|75)',
    'New/1yr: 60% | 2-5yr: 80% | 6yr+: 80%': '60 60 80 80 80 80 80 80',
    'Age 0-5:0-60%, Age 6+:70%': '60 60 60 60 60 60 70 70',
    'New/1yr: 40% | 2-5yr: 40% | 6yr+: 70%': '40 40 40 40 40 40 70 70',
    'New/1yr: 40% | 2-5yr: 80% | 6yr+: 80%': '40 40 80 80 80 80 80 80',
    'New/1yr: 40% | 2-5yr: 85% | 6yr+: 85%': '40 40 85 85 85 85 85 85',
    'New/1yr: D | 2-5yr: 80% | 6yr+: D': 'D D 80 80 80 80 D D',
    'New/1yr: D | 2-5yr: 85% | 6yr+: D': 'D D 85 85 85 85 D D',
    'New/1yr: 60% | 2-5yr: 75% | 6yr+: with NCB/break in > 90 days : 80%, without NCB : 70%': '60 60 75 75 75 75 ncb(80|70) ncb(80|70)',
    'New/1yr: 60% | 2-5yr: 60% | 6yr+: 80%': '60 60 60 60 60 60 80 80',
    'New/1yr: D | 2-5yr: 85% | 6yr+: 80%': 'D D 85 85 85 85 80 80',
    'Age 0-5: 0-70% Age 6+ 60%': '70 70 70 70 70 70 60 60',
    'New/1yr: D | 2-5yr: 60%/85% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': 'D D 60/85 60/85 60/85 60/85 ncb(90|80) ncb(90|80)',
    'New/1yr: D | 2-5yr: 85% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': 'D D 85 85 85 85 ncb(90|80) ncb(90|80)',
    'New/1yr: D | 2-5yr: 0% | 6yr+: 85%/90%': 'D D 0 0 0 0 85/90 85/90',
    'New/1yr: D | 2-5yr: 80% | 6-10yr: 0% | 11yr+: 0%': 'D D 80 80 80 80 0 0',
    'New/1yr: 60% | 2-5yr: 92% | 6yr+: 92%': '60 60 92 92 92 92 92 92',
    'New/1yr: D | 2-5yr: D | 6yr+: with NCB/break in > 90 days : 85%, without NCB : 75%': 'D D D D D D ncb(85|75) ncb(85|75)',
    'New/1yr: D | 2-5yr: 80% | 6yr+: with NCB/break in > 90 days : 85%, without NCB : 75%': 'D D 80 80 80 80 ncb(85|75) ncb(85|75)',
    'New/1yr: D | 2-5yr: 85% | 6yr+: with NCB/break in > 90 days : 85%, without NCB : 75%': 'D D 85 85 85 85 ncb(85|75) ncb(85|75)',
    'New/1yr: 60% | 2-5yr: 70% | 6yr+: 85%/90%': '60 60 70 70 70 70 85/90 85/90',
    'New/1yr: 60% | 2-5yr: 70% | 6yr+: 85%': '60 60 70 70 70 70 85 85',
    'New/1yr:0| 2-5yr: 0 | 6yr+: 70%': '0 0 0 0 0 0 70 70',
    '0-5 yr cd1- 0-80%; 6 yr + cd1 60%': '80 80 80 80 80 80 60 60',
    'new, 1 yr cd1- 60%, 2+ cd1 85%': '60 60 85 85 85 85 85 85',
    'New/1yr: 60% | 2-5yr: 70%% | 6yr+: with NCB/break in > 90 days : 75%, without NCB : 65%': '60 60 70 70 70 70 ncb(75|65) ncb(75|65)',
    '0-5 yr - 0-60%; 6 yr+ cd1 - 0-85%': '60 60 60 60 60 60 85 85',
    '0-5yr: 0-80% | 6yr+ : 0-70%': '80 80 80 80 80 80 70 70',
    '2-5yr: D | 6yr+: 90%': 'D D D D D D 90 90',
    '2-5yr: 85% | 6yr+: 85%/90%': 'D D 85 85 85 85 85/90 85/90',
    'New/1yr: 60% | 2-4yr: 85% | 5-10yr:90%/92%|10yr+:85%/90%': '60 60 85 85 85 90/92 90/92 85/90',
    'New/1yr: D | 2-4yr: 85% |5-10yr:90%/92%|10yr+:85%/90%': 'D D 85 85 85 90/92 90/92 85/90',
    '2-5yr: D | 6yr+: 85%/90%': 'D D D D D D 85/90 85/90',
    'New/1yr: D | 2-4yr: D |5-10yr:90%/92%|10yr+:85%/90%': 'D D D D D 90/92 90/92 85/90',
    'New/1yr: D | 2-4yr: 80% |5-10yr:90%/92%|10yr+:85%/90%': 'D D 80 80 80 90/92 90/92 85/90',
    '2-5yr: 75% | 6yr+: 90%': 'D D 75 75 75 75 90 90',
    'new, 1 yr cd1 - 60%; 2-5yr: 75% | 6yr+: 85%/90%': '60 60 75 75 75 75 85/90 85/90',
    '0-5yr : 0-60% |6yr+ : 0-70%': '60 60 60 60 60 60 70 70',
    'new, 1 yr cd1 - 60%; 2+ yr cd1 - 85%': '60 60 85 85 85 85 85 85',
    '0-2yr - 0-70%; 3yr+ cd1 - 0-85%': '70 70 70 85 85 85 85 85',
    '0yr - 0-70%; 1 yr+ cd1 - 0-85%': '70 85 85 85 85 85 85 85',
    '0-2 yr - 0-70%; 3 yr+ cd1 - 0-85%': '70 70 70 85 85 85 85 85',
    '20.1%--70%': '20.1-70 20.1-70 20.1-70 20.1-70 20.1-70 20.1-70 20.1-70 20.1-70',
    '0-2 yr - 0-60%; 3yr+ cd1 - 0-90%': '60 60 60 90 90 90 90 90',
    '0-5yr:0-70% | 6yr+: 0-80%': '70 70 70 70 70 70 80 80',
    '0-5yr:0-80%| 6yr+ : 0-70%': '80 80 80 80 80 80 70 70',
    'New/1yr: D | 2-5yr: 85%': 'D D 85 85 85 85 85 85',
    'New/1yr: D | 2-5yr: 80%': 'D D 80 80 80 80 80 80',
    'New/1yr: 60% | 2-4yr: 85% | 5yr: 85%/90%|6-10yr:90%/92%|10yr+:85%/90%': '60 60 85 85 85 85/90 90/92 85/90',
    'New/1yr: D | 2-4yr: 85% | 5yr: 85%/90%|6-10yr:90%/92%|10yr+:85%/90%': 'D D 85 85 85 85/90 90/92 85/90',
    'New/1yr: D | 2-4yr: D | 5yr: 85%/90%|6-10yr:90%/92%|10yr+:85%/90%': 'D D D D D 85/90 90/92 85/90',
    'New/1yr: D | 2-4yr: 80% | 5yr: 85%/90%|6-10yr:90%/92%|10yr+:85%/90%': 'D D 80 80 80 85/90 90/92 85/90',
    'new, 1 yr - 60%, 2-5yr: 75% | 6yr+: 85%/90%': '60 60 75 75 75 75 85/90 85/90',
    'new, 1 yr cd1 -60%, 2+ yr cd1 - 85%': '60 60 85 85 85 85 85 85',
    'New cd1-60%; 1+yr cd1 0-85%': '60 85 85 85 85 85 85 85',
    'new cd1 - 0-60%; 1yr+ cd1 - 0-85%': '60 85 85 85 85 85 85 85',
    '2-5yr: 65% | 6yr+: 85%': 'D D 65 65 65 65 85 85',
    'New/1yr: 60% | 2-5yr: 80% | 6yr+: 85%/90%': '60 60 80 80 80 80 85/90 85/90',
    'New/1yr: 75% | 2-5yr: 75% | 6yr+: 85%/90%': '75 75 75 75 75 75 85/90 85/90',
    'new, 1 yr cd1 - 60%; 2-5yr: 65% | 6yr+: 85%': '60 60 65 65 65 65 85 85',
    'New: With Addon: 85%/90% Without Addon: 60% | 1yr: 60% | 2-5yr: 85% | 6yr+: 85/90%': 'addon(85/90|60) 60 85 85 85 85 85/90 85/90',
    '2-5yr: 70% | 6yr+: 90%': 'D D 70 70 70 70 90 90',
    '2-5yr: 70% | 6yr+: 85%': 'D D 70 70 70 70 85 85',
    'New/1yr: 60% | 2-5yr: 85% | 6yr+: 85%/90%': '60 60 85 85 85 85 85/90 85/90',
    'new, 1 yr cd1 - 60%; 2-5yr: 70% | 6yr+: 85%': '60 60 70 70 70 70 85 85',
    'With Addon: 85%/90%; Without Addon: 60%': 'addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60)',
    'new, 1 yr cd1 - 60%; 2yr+ cd1 - 80%': '60 60 80 80 80 80 80 80',
    'new, 1 yr cd1 - 0-70%, 2 yr+ cd1 - 0-60%': '70 70 60 60 60 60 60 60',
    'age 0-1 - 70%; age 2+ 0-60%': '70 70 60 60 60 60 60 60',
    '2-5yr: 60% | 6yr+: 80%': 'D D 60 60 60 60 80 80',
    'new, 1-5yr: 60% | 6yr+: 80%': '60 60 60 60 60 60 80 80',
    '0-5 yr 0-50%; 6yr+ 0-70%': '50 50 50 50 50 50 70 70',
    '0-5 yr 0-45%; 6 yr+ 0-85%': '45 45 45 45 45 45 85 85',
    '2-5yr: 0% | 6yr+: 50%': 'D D 0 0 0 0 50 50',
    'new, 1-5yr: 0% | 6yr+: 50%': '0 0 0 0 0 0 50 50',
    '0-5 yr cd1 - 0-45%; 6 yr+ cd1 - 0-85%': '45 45 45 45 45 45 85 85',
    'new, 1-5yr: 45% | 6yr+: 85%': '45 45 45 45 45 45 85 85',
    '2-5yr: 45% | 6yr+: 85%': 'D D 45 45 45 45 85 85',
    '0-5 yr cd1 - 0-45%; 6 yr+ cd1 - 0-70%': '45 45 45 45 45 45 70 70',
    'New/1yr: D | 2-5yr: 20% | 6yr+: with NCB/break in > 90 days : 85%, without NCB : 75%': 'D D 20 20 20 20 ncb(85|75) ncb(85|75)',
    'new/1yr/2-5yr: 20% | 6-10yr: 85% | 10-20yr : 85% | 20+yr: 85%': '20 20 20 20 20 20 85 85',
    'New/1yr: 85% | 2-5yr: 0% | 6-10yr : 50% | 10-20yr : 50%': '85 85 0 0 0 0 50 50',
    '0-5 yr cd1 - 0-45%; 6 yr+ cd1 - 0-80%': '45 45 45 45 45 45 80 80',
    'new, 1 yr cd1 - 85%; 2 yr+ cd1 - 90%': '85 85 90 90 90 90 90 90',
    'New/1yr: 60% | 2-5yr: 90% | 6yr+: 90%': '60 60 90 90 90 90 90 90',
    '2-5yr: 90% | 6yr+: 85%/90%': 'D D 90 90 90 90 85/90 85/90',
    'New/1yr: 60% | 2-5yr: 90% | 6yr+: 85%/90%': '60 60 90 90 90 90 85/90 85/90',
    'new cd1 - 0-70%; 1 yr+ cd1 - 0-50%': '70 50 50 50 50 50 50 50',
    '2-5yr: 80% | 6-10yr: 50% | 11yr+: 82%': 'D D 80 80 80 80 50 82',
    'new, 1yr cd1 - 60%; 2-5yr: 80% | 6-10yr: 50% | 11yr+: 92%': '60 60 80 80 80 80 50 92',
    'New/1yr: 60% | 2-5yr: 80% | 6yr+: 92%': '60 60 80 80 80 80 92 92',
    '2-5yr: 80% | 6yr+: 92%': 'D D 80 80 80 80 92 92',
    'New/1yr: 60% | 2-5yr: 85%/92% | 6yr+: 92%': '60 60 85/92 85/92 85/92 85/92 92 92',
    'new, 1 yr cd1 - 60%; 2-5yr: 80% | 6-10yr: 50% | 11yr+: 92%': '60 60 80 80 80 80 50 92',
    'new, 1-5 yr cd1 - 0-80%; 6 yr+ cd1 - 0-60%': '80 80 80 80 80 80 60 60',
    '0-5yr:0-65% | 6yr+: 0-75%': '65 65 65 65 65 65 75 75',
    'new, 1-5 yr cd1 - 0-40%; 6 yr+ cd1 - 0-70%': '40 40 40 40 40 40 70 70',
    '2-5yr:80% | 6-10yr:80% |10yr+: 85%': 'D D 80 80 80 80 80 85',
    '0-5yr:80% | 6-10yr:80% |10yr+: 85%': '80 80 80 80 80 80 80 85',
    '2-5yr: 70% | 6-10yr:80% |10yr+: 85%': 'D D 70 70 70 70 80 85',
    'new, 1 yr cd1 - 60%; 2-5yr: 70% | 6-10yr: 80% | 10yr+:85%': '60 60 70 70 70 70 80 85',
    'With Addon: 85/90%; Without Addon: 60%': 'addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60) addon(85/90|60)',
    'new, 1 yr cd1 - 60%, 2 yr+ cd1 - 85%': '60 60 85 85 85 85 85 85',
    '1 yr+ cd1 - 0-85%': 'D 85 85 85 85 85 85 85',
    'new, 1 yr cd1 - 60%; 2 yr+ cd1 - 85%': '60 60 85 85 85 85 85 85',
    'new, 1 yr cd1 - 60%; 2yr+ cd1 - 85%': '60 60 85 85 85 85 85 85',
    'new, 1 yr cd1 - 60%; 2yr+ cd1 - 85%/90%': '60 60 85/90 85/90 85/90 85/90 85/90 85/90',
    'New/1yr: 60% | 2-5yr: 85%/90% | 6yr+: 85%/90%': '60 60 85/90 85/90 85/90 85/90 85/90 85/90',
    'new, 1 yr cd1 - 60%; 2-5yr: 60% | 6yr+: 85%': '60 60 60 60 60 60 85 85',
    '2-5yr: 60% | 6yr+: 85%': 'D D 60 60 60 60 85 85',
    'New/1yr: D | 2-5yr: 85% | 6yr+: 85%/90%': 'D D 85 85 85 85 85/90 85/90',
    '0-5 yr - 0-60%; 6 yr+ cd1 - 0-70%': '60 60 60 60 60 60 70 70',
    '0-5 yr - 0-40%; 6 yr+ cd1 - 0-70%': '40 40 40 40 40 40 70 70',
    'new, 1 yr cd1 - 60%; 2 yr+ cd1 - 75%': '60 60 75 75 75 75 75 75',
    '0-5yr:0-70% | 6yr+ : 0-80%': '70 70 70 70 70 70 80 80',
    'New/1yr: D | 2-5yr: 80% | 6-10yr: 85%/90%|10yr+: 80%': 'D D 80 80 80 80 85/90 80',
    'New/1yr:0%|2-5yr: 85% | 6yr+: 90%': '0 0 85 85 85 85 90 90',
    'New/1yr:0%|2-5yr: D | 6yr+: 85%': '0 0 D D D D 85 85',
    'New/1yr: D | 2-5yr: 75%/80% | 6-10yr: 85%/90% | 11yr+: 75%/80%': 'D D 75/80 75/80 75/80 75/80 85/90 75/80',
    'New/1yr:0%|2yr+: 85%': '0 0 85 85 85 85 85 85',
    'New/1yr: 0%| 2-5yr: 85% | 6yr+: 90%': '0 0 85 85 85 85 90 90',
    'New/1yr: 0% | 2-5yr: D | 6yr+: 85%': '0 0 D D D D 85 85',
    'New/1yr: D | 2-5yr: 0% | 6yr+: with NCB/break in > 90 days : 90%, without NCB : 80%': 'D D 0 0 0 0 ncb(90|80) ncb(90|80)',
    'New/1yr: 0% | 1+: 85%': '0 0 85 85 85 85 85 85',
    '0-40% for age 6+ years , 0-70% for age 0-5 years': '70 70 70 70 70 70 40 40',
    '0-5 Yr: 0-70% | 6+ Yr: 0-80%': '70 70 70 70 70 70 80 80',
    '0-5 Yr: 0-70% | 6+ Yr: 0-40%': '70 70 70 70 70 70 40 40',
    '0-5 Yr: 0-75% | 6+ Yr: 0-40%': '75 75 75 75 75 75 40 40',
    '0-5 Yr: 0-75% | 6+ Yr: 0-70%': '75 75 75 75 75 75 70 70',
    'New: 0-85% | 1+ Yr: 0-80%': '85 80 80 80 80 80 80 80',
    'Age 0-5: 0-70% Age 5+: 0-20%': '70 70 70 70 70 70 20 20',
    '0-5 Yr: 0-60% | 6+ Yr: 0-65%': '60 60 60 60 60 60 65 65',
    'New: 0-85% | 1+ Yr: 0-65%': '85 65 65 65 65 65 65 65',
    'New: 0-85% | 1+ Yr: 0-60%': '85 60 60 60 60 60 60 60',
    '0-5 Yr: 0-75% | 6+ Yr: 0-60%': '75 75 75 75 75 75 60 60',
    '0-5 Yr: 0-85% | 6+ Yr: 0-60%': '85 85 85 85 85 85 60 60',
    '0-5 Yr: 0-20% | 6+ Yr: 0-65%': '20 20 20 20 20 20 65 65',
    '0-5 Yr: 0-70% | 6+ Yr: 0-50%': '70 70 70 70 70 70 50 50',
    '0-5 Yr: 0-80% | 6+ Yr: 0-20%': '80 80 80 80 80 80 20 20',
    '0-5 Yr: 0-80% | 6+ Yr: 0-80%': '80 80 80 80 80 80 80 80',
    '0-5 Yr: 0-30% | 6+ Yr: 0-65%': '30 30 30 30 30 30 65 65',
    'New/1yr: D | 2-4yr: D | 5yr+: 85%/90%': 'D D D D D 85/90 85/90 85/90',
    'New/1yr: 60% | 2-4yr: 85% | 5yr+: 85%/90%': '60 60 85 85 85 85/90 85/90 85/90',
    '2-5yr: 75% | 6yr+: 85%/90%': 'D D 75 75 75 75 85/90 85/90',
    'New/1yr: D | 2-4yr: 80% | 5yr+: 85%/90%': 'D D 80 80 80 85/90 85/90 85/90',
    'New/1yr: D | 2-4yr: 85% | 5yr+: 85%/90%': 'D D 85 85 85 85/90 85/90 85/90',
    'New: With Addon: 85%/90% Without Addon: 60% | 1yr: 60% | 2-5yr: 85% | 6yr+: 90%': 'addon(85/90|60) 60 85 85 85 85 90 90',
    '2-5yr: 75% | 6yr+: 85%': 'D D 75 75 75 75 85 85',
    '2-5yr: 80% | 6-10yr: 50% | 11yr+: 92%': 'D D 80 80 80 80 50 92',
}
REFUSED_WITHOUT_CELLS = [  # read only with the payout cells (or not at all)
    'Hyundai: 70%, Maruti 85%',   # ['0.4', '0.4', '0.4', '0.4', '0.4', '0.4', '0.4', '0.4']
    '85%90%',   # ['D', 'D', '17.5%/10%', '17.5%/10%', '17.5%/10%', '17.5%/10%', '20%/12.5%', '20%/12.5%']
    'New: With Addon: 85/90%|Without Addon: -100%',   # ['With Addon: 10% Without Addon: 2.5%', 'D', '18%/9%', '18%/9%', '18%/9%', '18%/9%', '32%/29%', '18%/15%']
]

# ================================================================ 1. the table
bad = [(t, read(t), want) for t, want in GOLDEN.items() if read(t) != want]
check(f"all {len(GOLDEN)} non-trivial CD1 wordings of the October and November grids are read as in the table (hand-checked)", not bad, bad[:3])
check("cd1_slots gives the same slots (the plain-slot and TATA-clause steps come first)", all(show(mod.cd1_slots(t)[0]) == want for t, want in list(GOLDEN.items())[:60]))
check("a wording that is only read with the payout cells (or not at all) is refused without them", all(read(t) is None for t in REFUSED_WITHOUT_CELLS), [t for t in REFUSED_WITHOUT_CELLS if read(t) is not None])

# ================================================================ 2. what the payout cells of the row back, and what they forbid
PAIRS = ["D", "D", "17.5%/10%", "17.5%/10%", "17.5%/10%", "17.5%/10%", "20%/12.5%", "20%/12.5%"]
check("'85%90%' (the slash is missing) is read as the pair 85/90 only where the payout cells are pairs", read("85%90%", PAIRS) == " ".join(["85/90"] * 8) and read("85%90%") is None and read("85%90%", ["10%"] * 8) is None
      and read("85%90%", ["D"] * 8) is None, read("85%90%", PAIRS))
x = mod.cd1_read("2-5yr: D | 6yr+: 90%", ["2.5%", "2.5%", "D", "D", "D", "D", "25%", "25%"])
check("a text that leaves out an age group whose payout cell HAS a payout: that group gets the whole range (marked 'unnamed', flagged by emit_cv), the groups it names follow the text",
      x[:2] == [("cap", [100.0], "unnamed")] * 2 and x[2:6] == [("D",)] * 4 and x[6:] == [("cap", [90.0])] * 2, x)
x = mod.cd1_read("2-5yr: D | 6yr+: 90%", ["", "", "D", "D", "D", "D", "25%", "25%"])
check("... but where the payout cells of the left-out groups are blank or D the groups are declined as before", x[:2] == [("D",)] * 2, x)
check("'Without Addon: -100%' is no cap: the wording is refused whatever the cells say", read("New: With Addon: 85/90%|Without Addon: -100%", ["With Addon: 10% Without Addon: 2.5%"] + ["18%/9%"] * 7) is None)

# ================================================================ 3. the earlier reader, frozen: the new grammar must agree with it wherever it understood a text
def _l_age_cols(label):
    l = label.lower().replace(" ", "")
    if l in ("new/1yr", "age0-1"):
        return [0, 1]
    if l in ("new", "age0"):
        return [0]
    if l == "1yr":
        return [1]
    if l in ("age>=1", "1+"):
        return list(range(1, 8))

    def col(a):
        return a if a <= 5 else (6 if a <= 10 else 7)
    m = re.fullmatch(r"(\d+)-(\d+)yr", l)
    if m:
        return sorted({col(x) for x in range(int(m.group(1)), int(m.group(2)) + 1)})
    m = re.fullmatch(r"(\d+)yr\+", l)
    if m:
        return sorted({col(x) for x in range(int(m.group(1)), 12)})
    return None


_L_SYNTAX = [
    (re.compile(r"^Age 0:\s*(.+?)\s+Age>=1:\s*(.+)$", re.I), lambda m: [("new", m.group(1)), ("1+", m.group(2))]),
    (re.compile(r"^(.+?) for age 0-1,\s*(.+?) for age 2-5,\s*(.+?) for age 6\+$", re.I), lambda m: [("new/1yr", m.group(1)), ("2-5yr", m.group(2)), ("6yr+", m.group(3))]),
    (re.compile(r"^Age 0-?\s*(\d+(?:\.\d+)?%?);\s*1\+\s*(\d+(?:\.\d+)?%?)$", re.I), lambda m: [("new", m.group(1)), ("1+", m.group(2))]),
]


def legacy(v):
    """the reader of 29 Oct - 10 Nov 2026 (cd1_slots with its pipe and 'Age 0: ...' branches), copied here as it was"""
    t = mod.norm(v).replace("–", "-").replace("—", "-")
    m = mod.TATA_CLAUSE_RX.search(t)
    if m:
        t = t[:m.start()].strip()
    one = mod.slot_of(t)
    if one:
        return [one] * 8
    out = None
    for rx, fn in _L_SYNTAX:
        m = rx.match(t)
        if m:
            out = [None] * 8
            for lab, val in fn(m):
                sl = mod.slot_of(val if ("%" in val or val.upper() == "D") else val + "%")
                if sl is None:
                    return None
                for c in _l_age_cols(lab) or []:
                    out[c] = sl
            break
    if out is None and "|" in t:
        out = [None] * 8
        for part in [p.strip() for p in t.split("|")]:
            m = re.match(r"^([^:]+):\s*(.+)$", part)
            cols = _l_age_cols(m.group(1).strip()) if m else None
            sl = mod.slot_of(m.group(2)) if m else None
            if cols is None or sl is None:
                return None
            for c in cols:
                out[c] = sl
    if out is None:
        return None
    res = []
    for i, o in enumerate(out):
        res.append(o if o is not None else (res[i - 1] if i else ("D",)))
    return res


OVERLAP = {                                           # the wordings in which a closed range ('6-10yr') and an open one ('10yr+', '1+') name the same age: the earlier reader let the LATER group overwrite the closed one
    "New/1yr: 60% | 2-4yr: 85% | 5-10yr:90%/92%|10yr+:85%/90%", "New/1yr: D | 2-4yr: 85% |5-10yr:90%/92%|10yr+:85%/90%", "New/1yr: D | 2-4yr: D |5-10yr:90%/92%|10yr+:85%/90%",
    "New/1yr: D | 2-4yr: 80% |5-10yr:90%/92%|10yr+:85%/90%", "2-5yr:80% | 6-10yr:80% |10yr+: 85%", "0-5yr:80% | 6-10yr:80% |10yr+: 85%", "2-5yr: 70% | 6-10yr:80% |10yr+: 85%",
    "New/1yr: D | 2-5yr: 80% | 6-10yr: 85%/90%|10yr+: 80%", "New/1yr: 0% | 1+: 85%"}
diff = [t for t in GOLDEN if legacy(t) is not None and legacy(t) != mod.cd1_read(t)]
check("the grammar reads every wording the earlier reader understood the same way - the only differences are the nine overlap wordings", set(diff) == OVERLAP, sorted(set(diff) ^ OVERLAP)[:3])
check("... and every wording the earlier reader refused as a whole is read now (except the three that need the payout cells)", all(mod.cd1_read(t) is not None for t in GOLDEN if legacy(t) is None))
check("overlap rule, 6-10yr / 10yr+: the closed range keeps the 6-10 Yr column, '10yr+' is the 11+ Yr column", read("2-5yr:80% | 6-10yr:80% |10yr+: 85%") == "D D 80 80 80 80 80 85"
      and read("New/1yr: D | 2-5yr: 80% | 6-10yr: 85%/90%|10yr+: 80%") == "D D 80 80 80 80 85/90 80")
check("overlap rule, 'New/1yr' / '1+' and 'Age 0-5' / 'Age 5+': the closed range keeps the age both name", read("New/1yr: 0% | 1+: 85%") == "0 0 85 85 85 85 85 85" and read("Age 0-5: 0-70% Age 5+: 0-20%") == "70 70 70 70 70 70 20 20")

CAPS = ["D", "60%", "85%", "85%/90%", "0%", "70%"]
CAP_SLOT = {"D": ("D",), "60%": ("cap", [60.0]), "85%": ("cap", [85.0]), "85%/90%": ("cap", [85.0, 90.0]), "0%": ("cap", [0.0]), "70%": ("cap", [70.0])}
SEGMENTATIONS = [[(0, 1), (2, 3, 4, 5), (6, 7)], [(0,), (1, 2, 3, 4, 5, 6, 7)], [(0, 1, 2, 3, 4, 5), (6, 7)], [(0, 1), (2, 3, 4), (5, 6), (7,)], [(0, 1), (2, 3, 4, 5), (6,), (7,)],
                 [(2, 3, 4, 5), (6, 7)], [(0, 1), (2, 3, 4, 5)], [(0,), (1,), (2, 3, 4, 5), (6, 7)], [(0, 1, 2), (3, 4, 5, 6, 7)], [(0, 1), (2, 3, 4), (5, 6, 7)]]


def expected(groups, caps):
    out = [None] * 8
    for cols, cap in zip(groups, caps):
        for c in cols:
            out[c] = CAP_SLOT[cap]
    res = []
    for i, o in enumerate(out):
        res.append(o if o is not None else (res[i - 1] if i else ("D",)))
    return res


L_LABELS = {(0,): ["new", "New", "age0", "Age 0"], (0, 1): ["new/1yr", "New/1yr", "age0-1", "Age 0-1"], (1,): ["1yr", "1 yr"], (1, 2, 3, 4, 5, 6, 7): ["1+", "age>=1", "Age>=1"], (2, 3, 4, 5): ["2-5yr", "2-5 yr"],
            (2, 3, 4): ["2-4yr"], (3, 4, 5, 6, 7): ["3yr+"], (6,): ["6-10yr"], (6, 7): ["6yr+", "6 yr+"], (7,): ["11yr+"], (0, 1, 2, 3, 4, 5): ["0-5yr"], (5, 6): ["5-10yr"], (0, 1, 2): ["0-2yr"], (5, 6, 7): ["5yr+"]}
rng = random.Random(int(os.environ.get("CD1_SEED", "20261110")))               # CD1_SEED / CD1_FUZZ (a multiplier of the sample sizes) are for a bigger one-off fuzz run
FUZZ = int(os.environ.get("CD1_FUZZ", "1"))
n_bad, first_bad = 0, None
for _ in range(2500 * FUZZ):
    groups = rng.choice(SEGMENTATIONS)
    caps = [rng.choice(CAPS) for _ in groups]
    text = rng.choice([" | ", "|", " |"]).join(f"{rng.choice(L_LABELS[g])}{rng.choice([': ', ':'])}{c}" for g, c in zip(groups, caps))
    want = expected(groups, caps)
    got_old, got_new = legacy(text), mod.cd1_read(text)
    if (got_old is not None and got_old != want) or got_new != want:                  # the earlier reader refuses a few spellings (its 'Age 0: .. Age>=1: ..' pattern swallowed a '|'); where it reads, it must agree
        n_bad += 1
        first_bad = first_bad or (text, got_old, got_new, want)
check("2,500 generated 'label: cap | label: cap' wordings: the new grammar, the earlier reader and the construction agree", n_bad == 0, first_bad)
n_bad, first_bad = 0, None
for _ in range(300 * FUZZ):
    a, b, c = (rng.choice(CAPS) for _ in range(3))
    for text, groups, caps in ((f"Age 0: {a} Age>=1: {b}", [(0,), (1, 2, 3, 4, 5, 6, 7)], [a, b]), (f"{a} for age 0-1, {b} for age 2-5, {c} for age 6+", [(0, 1), (2, 3, 4, 5), (6, 7)], [a, b, c])):
        if mod.cd1_read(text) != expected(groups, caps) or legacy(text) != expected(groups, caps):
            n_bad += 1
            first_bad = first_bad or (text, legacy(text), mod.cd1_read(text))
check("generated 'Age 0: .. Age>=1: ..' and '.. for age 0-1, .. for age 2-5, .. for age 6+' wordings", n_bad == 0, first_bad)

# ================================================================ 4. spellings the earlier reader did not know
N_LABELS = {(0,): ["new", "New", "NEW", "age 0", "Age0", "0yr", "0 yr"], (0, 1): ["new/1yr", "New / 1 yr", "new, 1 yr", "new & 1yr", "age 0-1", "0-1yr", "0-1 yr"], (1,): ["1yr", "1 yr", "1 Yr", "age 1"],
            (1, 2, 3, 4, 5, 6, 7): ["1+", "1+ yr", "1 yr+", "1yr+", "age>=1", "age 1+", ">=1"], (2, 3, 4, 5): ["2-5yr", "2-5 yr", "2-5 Yr", "age 2-5", "2 - 5 yrs", "2-5years"],
            (6, 7): ["6yr+", "6+ yr", "6 yr +", "6+ Yr", "age 6+", "6 years+", ">=6", "6+"], (6,): ["6-10yr", "6-10 Yr", "age 6-10"], (7,): ["11yr+", "11+ yr", "11+ Yr", "age 11+"],
            (0, 1, 2, 3, 4, 5): ["0-5yr", "0-5 yr", "0-5 Yr", "age 0-5"], (2, 3, 4): ["2-4yr", "2-4 yr"], (5, 6): ["5-10yr"], (5, 6, 7): ["5yr+", "5+ yr"], (0, 1, 2): ["0-2yr", "0-2 yr"], (3, 4, 5, 6, 7): ["3yr+", "3 yr+"]}
N_CAPS = [("D", ("D",)), ("d", ("D",)), ("60%", ("cap", [60.0])), ("0-60%", ("cap", [60.0])), ("0–60%", ("cap", [60.0])), ("60 %", ("cap", [60.0])), ("85%/90%", ("cap", [85.0, 90.0])),
          ("85 % / 90 %", ("cap", [85.0, 90.0])), ("0%", ("cap", [0.0])), ("70%%", ("cap", [70.0])), ("20.1%-70%", ("band", 20.1, 70.0))]
n_bad, first_bad = 0, None
for _ in range(3000 * FUZZ):
    groups = rng.choice(SEGMENTATIONS)
    picks = [rng.choice(N_CAPS) for _ in groups]
    parts = []
    for g, (ctext, _) in zip(groups, picks):
        label = rng.choice(N_LABELS[g])
        cd1w = rng.choice(["", " cd1", " CD1"]) if rng.random() < 0.3 else ""
        parts.append(f"{label}{cd1w}{rng.choice([': ', ':', ' - ', ' '])}{ctext}")
    text = rng.choice([" | ", "|", "; ", ", ", " ; "]).join(parts)
    out = [None] * 8
    for g, (_, slot) in zip(groups, picks):
        for c in g:
            out[c] = slot
    want = []
    for i, o in enumerate(out):
        want.append(o if o is not None else (want[i - 1] if i else ("D",)))
    got = mod.cd1_read(text)
    if got != want:
        n_bad += 1
        first_bad = first_bad or (text, got, want)
check("3,000 generated wordings in many spellings (labels, units, 'cd1', separators, caps) are read as constructed", n_bad == 0, first_bad)
FOR_LABELS = {(0, 1): ["age 0-1", "0-1yr"], (2, 3, 4, 5): ["age 2-5", "2-5yr", "2-5 yrs"], (6, 7): ["age 6+ years", "age 6+", "6yr+", "6+ yr"], (0,): ["age 0", "new"], (1, 2, 3, 4, 5, 6, 7): ["age 1+", "1+", "1yr+"],
              (0, 1, 2, 3, 4, 5): ["age 0-5", "0-5 years", "0-5yr"]}
n_bad, first_bad = 0, None
for _ in range(400 * FUZZ):
    groups = rng.choice([[(0, 1), (2, 3, 4, 5), (6, 7)], [(0, 1, 2, 3, 4, 5), (6, 7)], [(0,), (1, 2, 3, 4, 5, 6, 7)]])
    picks = [rng.choice(N_CAPS[:7]) for _ in groups]
    text = ", ".join(f"{c} for {rng.choice(FOR_LABELS[g])}" for g, (c, _) in zip(groups, picks))
    out = [None] * 8
    for g, (_, slot) in zip(groups, picks):
        for c in g:
            out[c] = slot
    want = [o if o is not None else ("D",) for o in out]
    got = mod.cd1_read(text)
    if got != want:
        n_bad += 1
        first_bad = first_bad or (text, got, want)
check("400 generated 'cap for age group' wordings ('0-40% for age 6+ years, 0-70% for age 0-5 years')", n_bad == 0, first_bad)

# ================================================================ 5. what must be refused (nothing guessed)
REFUSE = ["only for hilly regions", "New/1yr: 60% | 2-5yr: 80% | 6yr+: 85% only Tata", "0-5yr: 60% | 5-10yr: 70%", "0-5yr: 60% | 6-8yr: 70% | 9yr+: 80%", "60% | 70%", "2-5yr:", "150%", "New/1yr: 90-85%",
          "0-5yr: 60% | 3-1yr: 70%", "Hyundai: 70%, Maruti 85%", "45yr+: 50%", "#VALUE!", "New/1yr: 60% | 2-5yr: 80% | 6yr+: ", "New/1yr: 60% | 2-5yr: 80% | 6yr+: 85% | 1+: 70%", "1+: 70% | 3+: 80%",
          "New/1yr: 60% | 2-5yr: 80% | 6yr+: 85% (Delhi only)", "0-5yr: 60% | 6yr+: 70% | Tata: 80%"]
check("refused: an unknown word, a group of two ages that disagrees with another group, a column split between two caps, two caps without an age, a cap above 100, a backwards range, an age the grid has not, an error",
      all(read(t) is None for t in REFUSE), [t for t in REFUSE if read(t) is not None])
n_bad, first_bad = 0, None
for _ in range(500 * FUZZ):
    groups = rng.choice(SEGMENTATIONS)
    caps = [rng.choice(CAPS) for _ in groups]
    parts = [f"{rng.choice(L_LABELS[g])}: {c}" for g, c in zip(groups, caps)]
    i = rng.randrange(len(parts) + 1)
    parts.insert(i, rng.choice(["only", "except", "hilly", "xyz", "Tata", "RTO"]))
    text = " | ".join(parts)
    if mod.cd1_read(text) is not None:
        n_bad += 1
        first_bad = first_bad or text
check("500 valid wordings with one unknown word put in between: every one is refused", n_bad == 0, first_bad)

# ================================================================ 6. the payout cells that were not read
check("CD2: '0.4' typed as text is the fraction the number cells hold (40%); '40' and '1.5' are not guessed", mod.cd2("0.4") == ("rate", [40.0]) and mod.cd2(0.4) == ("rate", [40.0]) and mod.cd2(".25") == ("rate", [25.0])
      and mod.cd2("40")[0] == "raw" and mod.cd2("1.5")[0] == "raw")
check("CD2: '2.5%; TATA With Addon: 15%/10% Without Addon: 10%' (no dash after TATA) is the TATA cell like October's 'TATA-With Addon'",
      mod.cd2("2.5%; TATA With Addon: 15%/10% Without Addon: 10%") == ("tata", 2.5, [15.0, 10.0], [10.0]) and mod.cd2("2.5%; TATA-With Addon: 10%/7.5%Without Addon: 7.5%") == ("tata", 2.5, [10.0, 7.5], [7.5]))

# ================================================================ 7. the taxi make column
P = mod.taxi_make_parts
ALL7 = frozenset(range(1, 8))
check("taxi make words: 'Hyundai Age 0' / 'Maruti Age 0' are the new-vehicle column only", P("Hyundai Age 0", "0-40%") == dict(entries=[("Hyundai", frozenset({0}))], slots=None) and P("Maruti Age 0")["entries"] == [("Maruti", frozenset({0}))])
check("taxi make words: an age window after the last make is for every make of the list ('Maruti & Hyundai age 1+')", P("Maruti & Hyundai age 1+", "0-70%")["entries"] == [("Maruti,Hyundai", ALL7)])
check("taxi make words: 'Maruti (all age) (excl WagonR Model) & Hyundai age 1+' = MARUTI_EX_WAGONR at every age + Hyundai from age 1",
      P("Maruti (all age) (excl WagonR Model) & Hyundai age 1+", "0-55%")["entries"] == [("MARUTI_EX_WAGONR", None), ("Hyundai", ALL7)])
check("taxi make words: 'including Wagon R' / '(including Ertiga, Maruti XI6, ECO, Invicto)' keep the group Maruti,Hyundai; 'Hyundai' and 'Hindustan' are the master's groups",
      P("Maruti & Hyundai including Wagon R")["entries"] == [("Maruti,Hyundai", None)] and P("Maruti & Hyundai (including Ertiga, Maruti XI6, ECO, Invicto)")["entries"] == [("Maruti,Hyundai", None)]
      and P("Hyundai")["entries"] == [("Hyundai", None)] and P("Hindustan")["entries"] == [("Hindustan", None)])
check("taxi make words: 'Others (Other than Hindustan make)' = TAXI_ALL_EX_HINDUSTAN (a group the master lacks: check_names lists it)", P("Others (Other than Hindustan make)")["entries"] == [("TAXI_ALL_EX_HINDUSTAN", None)])
x = P("Maruti & Hyundai (including Ertiga, Maruti XI6, ECO, Invicto)", "Hyundai: 70%, Maruti 85%")
check("a CD1 text with a cap per make ('Hyundai: 70%, Maruti 85%') gives one entry and one cap per make", x["entries"] == [("Maruti", None), ("Hyundai", None)]
      and x["slots"]["Maruti"] == [("cap", [85.0])] * 8 and x["slots"]["Hyundai"] == [("cap", [70.0])] * 8, x)
check("taxi make words not understood are refused: an unknown make, an unknown model, a window that splits an age column, two entries of one make, a cap for a make the text does not name",
      all(P(t, c) is None for t, c in (("Kia", None), ("Maruti & Hyundai (including Tata Tigor)", None), ("Hyundai age 8", None), ("Maruti & Maruti", None), ("Hyundai", "Hyundai: 70%, Maruti 85%"), ("Maruti (excl WagonR) & Hyundai age 0-8", None))))

C_ = lambda *v: list(v)
taxi_head = ["Cluster segment", "Segment", "Body Type", "RTO Cluster", "State", "Make / Variant", "CD1 (OD-disc)"] + ["New", "1 Yr", "2 Yr", "3 Yr", "4 Yr", "5 Yr", "6-10 Yr", "11+ Yr"] * 2
HCV_NOTE = "Note: Volvo, Scania to block. \nOnly for Non-Dumper/Tipper: Diamler, Bharat Benz and Mahindra to operate at 70% CD1"


def book(cv_rows, extra=(), remarks=None):
    sheets = [mod.Sheet("LATEST CV COMBINED GRID", [[], ["", "", "", "", "", "", "", "COMPREHENSIVE"], ["", "", "", "", "", "", "", "CD2"], taxi_head] + [list(r) for r in cv_rows])]
    sheets.append(mod.Sheet("CV Remarks", [["REMARKS FOR SCHOOL BUS"], ["1. Short Term Policies - CD2 reduction of 25% on the above Rate"], ["REMARKS FOR STAFF BUS"], ["Note:"]]
                            + [[r] for r in (remarks if remarks is not None else ["For Seating Capacity up to 10-seater, CD1_85% & CD2_20% in all open cohorts."])] + [["REMARKS FOR HCV"], [HCV_NOTE]]))
    sheets.extend(extra)
    sheets.append(mod.Sheet("Hidden old grid", [["x"]], visible=False))
    return mod.Book(sheets)


def run(b):
    mod.reset_run_lists()
    return mod.build(None, F, T, quiet, mod.Asker(interactive=False), book=b)


mod.MISSING_TABLES = "skip"


def taxi(make, cd1, body="With Add on cover", seg="Taxi upto 5 seater non-diesel > 1000 cc", comp=0.15, satp=0.4):
    return ("CV PCV4(Upto 5 seating)", seg, body, "DELHI-NCR", "Delhi", make, cd1) + (comp,) * 8 + (satp,) * 8


df, w = run(book([taxi("Hyundai Age 0", "0-40%"), taxi("Maruti & Hyundai age 1+", "0-70%", comp=0.3), taxi("Maruti (all age) (excl WagonR Model) & Hyundai age 1+", "0-55%", comp=0.25)]))
d = df[df["product"] == "PCV 4W"]
g1 = d[d.new_vehicle_makes == "Hyundai"]
g2 = d[d.new_vehicle_makes.str.lower() == "maruti,hyundai"]
g3a, g3b = d[d.new_vehicle_makes.str.upper() == "MARUTI_EX_WAGONR"], d[(d.new_vehicle_makes == "Hyundai") & (d.vehicle_age_min > 0)]
check("taxi end to end: 'Hyundai Age 0' writes only the new-vehicle age band (both 1+1 and STP, three non-diesel fuels)", len(g1[g1.vehicle_age_min == 0]) == 6 and set(g1[g1.pi_net_rate == 15.0].vehicle_age_min) == {0.0}, len(g1))
check("taxi end to end: 'Maruti & Hyundai age 1+' = the group Maruti,Hyundai for every age band from 1 year (7 bands x 3 fuels x 1+1 / STP), no New row",
      len(g2) == 7 * 3 * 2 and 0 not in set(g2.vehicle_age_min) and set(g2[g2.sub_product == "1+1"].tariff_max) == {70.01}, len(g2))
check("taxi end to end: 'Maruti (all age) (excl WagonR Model) & Hyundai age 1+' = MARUTI_EX_WAGONR at all 8 age bands and Hyundai from age 1 (7 bands)", len(g3a) == 8 * 3 * 2 and len(g3b[g3b.tariff_max == 55.01]) == 7 * 3, (len(g3a), len(g3b)))
check("taxi end to end: none of these rows is flagged (the make texts are read, nothing is left for a person)", (d.remarks == "").all() and not any("taxi make" in x for x in w), d.remarks.unique()[:2])
df, w = run(book([taxi("Maruti & Hyundai (including Ertiga, Maruti XI6, ECO, Invicto)", "Hyundai: 70%, Maruti 85%", body="Without Add on cover", seg="Taxi upto 5 seater non-diesel < 1000 cc", comp=0.4)]))
d = df[df["product"] == "PCV 4W"]
check("a cap per make: Maruti at 85.01, Hyundai at 70.01 (comprehensive rows), every age, no flag", set(d[(d.new_vehicle_makes == "Maruti") & (d.sub_product == "1+1")].tariff_max) == {85.01} and set(d[(d.new_vehicle_makes == "Hyundai") & (d.sub_product == "1+1")].tariff_max) == {70.01}
      and (d.remarks == "").all() and len(d) == 2 * 8 * 3 * 2, (d.new_vehicle_makes.unique(), d.remarks.unique()[:1]))
df, w = run(book([taxi("Kia", "0-40%")]))
d = df[df["product"] == "PCV 4W"]
check("a taxi make the grammar does not know is still written for TAXI_ALL and flagged (as before)", set(d.new_vehicle_makes) == {"TAXI_ALL"} and d.remarks.str.contains("taxi make 'Kia'").all())

df, w = run(book([taxi("Others (Other than Hindustan make)", "0-30%", seg="Taxi upto 5 seater diesel > 1000 cc & <=1500 cc", comp=0.1, satp=0.1)]))
d = df[df["product"] == "PCV 4W"]
mod.collect_gaps(df)
mg = mod.EXTRA_SHEETS.get("Make groups to create")
check("'Others (Other than Hindustan make)': written for TAXI_ALL_EX_HINDUSTAN (the master lacks it: the rows say so) and the review workbook lists the group WITH its definition ('made of TAXI_ALL except Hindustan', never a text that starts with '=': Excel would read it as a formula)",
      set(d.new_vehicle_makes) == {"TAXI_ALL_EX_HINDUSTAN"} and mg is not None and "TAXI_ALL_EX_HINDUSTAN" in set(mg.group_needed) and mg[mg.group_needed == "TAXI_ALL_EX_HINDUSTAN"]["items"].iloc[0] == "made of TAXI_ALL except Hindustan",
      None if mg is None else mg.values.tolist())

# ================================================================ 8. a CD1 text that leaves out an age group which has a payout; the staff-bus seat note
hcv = ("CV HCV", "GCV4 44T+", "Gas Tanker", "GOOD GJ SOUTH", "Gujarat", "TATA", "2-5yr: D | 6yr+: 90%") + (0.025, 0.025, "D", "D", "D", "D", 0.25, 0.25) + (0.125, 0.125, "D", "D", "D", "D", 0.055, 0.055)
df, w = run(book([hcv]))
d = df[(df["product"] == "GCV 4W") & (df.sub_product == "1+1")]
check("a text that does not name New / 1 Yr although they carry a payout: those rows are written at 2.5% for the whole range and flagged by name; the 6+ rows keep the cap 90.01 and the same flag lists the groups",
      set(d[d.vehicle_age_min < 2].tariff_max) == {100.01} and set(d[d.vehicle_age_min < 2].pi_net_rate) == {2.5} and set(d[d.vehicle_age_min >= 6].tariff_max) == {90.01}
      and d.remarks.str.contains("does not name the age group").all(), (sorted(set(d.tariff_max)), d.remarks.iloc[0][:150]))
ok = ("CV HCV", "GCV4 12 to 20T", "Dumper/Tipper", "BAD GJ", "Gujarat", "All", "New/1yr: 60% | 2-5yr: 80% | 6-10yr: 85%/90%|10yr+: 80%") + (0.25,) * 6 + ("33%/30%", "19%/16%") + (0.3,) * 8
df, w = run(book([ok]))
d = df[(df["product"] == "GCV 4W") & (df.sub_product == "1+1")]
check("a wording read by the grammar writes rows with no flag; the pair payout of '6-10yr: 85%/90%' gets both caps (85.01 at 33, 90.01 at 30)",
      (d.remarks == "").all() and sorted(set(d[(d.vehicle_age_min == 6) & (d.fuel_type == "Diesel")].tariff_max)) == [85.01, 90.01]
      and sorted(set(d[(d.vehicle_age_min == 6) & (d.fuel_type == "Diesel")].pi_net_rate)) == [30.0, 33.0], d.remarks.unique()[:1])
check("the grammar's readings are listed for the review workbook (sheet 'CD1 readings'), nothing in them needs an answer", len(mod.CD1_AUTO) == 1 and mod.cd1_readings_frame().shape == (1, 10), mod.cd1_readings_frame().columns.tolist())
mod.collect_gaps(df)
check("... and the sheet is added", "CD1 readings" in mod.EXTRA_SHEETS)
df, w = run(book([ok]))
check("the staff-bus note 'up to 10-seater' is no flag on any row: its seat range is not written, it is listed once in 'Not written'",
      (df.remarks == "").all() and any(x["segment"] == "Staff Bus up to 10 seats" for x in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN[:2])
staff = ("Staff Bus", "Staff Bus", "CORPORATE (SELF-USE) ", "ROM", "MAHARASHTRA", ">10", "60%") + (0.4,) * 16
df, w = run(book([staff]))
check("... and the staff-bus rows (more than 10 seats) are written with no flag", len(df) > 0 and (df.remarks == "").all(), df.remarks.unique()[:1])
df, w = run(book([staff], remarks=["For Seating Capacity up to 12-seater, CD1_85% & CD2_20% in all open cohorts."]))
check("a seat limit ABOVE where the staff-bus table starts (up to 12-seater) is not that case: the note stays flagged on the staff-bus rows", df.remarks.str.contains("12-seater").all() and len(df) > 0, df.remarks.unique()[:1])

# ================================================================ 9. MISP
def sheet(name, above, head, rows):
    return mod.Sheet(name, [list(r) for r in above] + [list(head)] + [list(r) for r in rows])


lcv = ("CV LCV", "GCV4 upto 1.6T", "LCV", "Delhi", "Delhi", "All", "0-60%") + (0.25,) * 16
s15 = sheet("2W Grid 1+5", [[None, None, None, None, "1+5"]], ["Key", "Agency/PB Clusters", "Make", "Agency/PB Seg", "CD1", "Avg CD2", "Max CD2", "Formula Type"],
            [["GJ_ABSHONDAMC <=155", "GJ_ABS", "HONDA", "MC <=155", 0.4, "MISP", "MISP", "Slab OD+Add on"]])
df, w = run(book([lcv], extra=[s15]))
m = df[(df["product"] == "TW") & (df.sub_product == "1+5")]
check("MISP: written On OD 17.5 (your October upload's value) with NO per-row flag; the run's warnings say so once", len(m) > 0 and set(m.pi_type) == {"On OD"} and set(m.pi_od_rate) == {17.5} and (m.remarks == "").all()
      and any("MISP" in x and "17.5" in x for x in w), (m.remarks.unique()[:1], [x for x in w if "MISP" in x]))

# ================================================================ 9b. weight bands read from their words, weight-keyed payout cells, the 180 cc Hero / Honda / TVS segment
REAL_BANDS = ["GCV4 12 to 20T", "GCV4 20 to 40T", "GCV4 40 to 44T", "GCV4 40T+", "GCV4 44T+", "GCV Upto 2.5T", "GCV4 1.6T To 2.5T", "GCV4 1.6T To 2T", "GCV4 2.5 To 3.5T", "GCV4 2T to 2.5T", "GCV4 Upto 2.5T",
              "GCV4 upto 1.6T", "GCV4 upto 2.5T", "GCV4 3.5 To 7.5T", "GCV4 7.5 to 12T"]
table = dict(mod.GVW_BANDS)
by_table = {b: mod.gvw_of(b) for b in REAL_BANDS}
mod.GVW_BANDS.clear()
try:
    by_words = {b: mod.gvw_of(b) for b in REAL_BANDS}
finally:
    mod.GVW_BANDS.update(table)
check("weight bands: the rule that reads the words gives the table's numbers for every label of the two workbooks", by_words == by_table and all(v is not None for v in by_table.values()), [b for b in REAL_BANDS if by_words[b] != by_table[b]])
check("weight bands: 'GCV4 upto 2T' = 0-2000.01, 'GCV4 40T-43T' = 40000.01-43000.01, 'GCV4 43T+' = 43000.01 and above; a backwards band, 'GCV 44T+' (no 4) and other words are not guessed",
      mod.gvw_of("GCV4 upto 2T") == (0, 2000.01) and mod.gvw_of("GCV4 40T-43T") == (40000.01, 43000.01) and mod.gvw_of("GCV4 43T+") == (43000.01, 99999)
      and mod.gvw_of("GCV4 43T-40T") is None and mod.gvw_of("GCV 44T+") is None and mod.gvw_of("Bus 40T") is None)
W = mod.weight_cell
check("weight-keyed payout cell: on the 43T+ row the clause '43T+: D' is the payout, on the 40T-43T row the clause '40 to 43T : 2.5%' is; no clause for the row's band, or text that is not a clause, is refused",
      W("40 to 43T : 2.5% 43T+: D", (43000.01, 99999)) == ("D",) and W("40 to 43T : 2.5% 43T+: D", (40000.01, 43000.01)) == ("rate", [2.5]) and W("40 to 43T : 2.5% 43T+: D", (44000.01, 99999)) is None
      and W("40 to 43T : 2.5% and more", (40000.01, 43000.01)) is None and W("see grid", (43000.01, 99999)) is None)
wb_rows = [("CV HCV", "GCV4 40T-43T", "Non-Dumper/Tipper", "WB_REF", "West Bengal", "All excluding Volvo and Scania*/ Non Dumper Tipper", "0") + ("D", "D", 0.025, 0.025, 0.025, 0.025, 0.05, 0.05) + ("D", "D", 0.05, 0.05, 0.05, 0.05, 0.05, 0.05),
           ("CV HCV", "GCV4 43T+", "Non-Dumper/Tipper", "WB_REF", "West Bengal", "All excluding Volvo and Scania*/ Non Dumper Tipper", "0") + ("D", "D") + ("40 to 43T : 2.5% 43T+: D",) * 4 + (0.025, 0.025) + ("D", "D", "D", "D", "D", "D", 0.025, 0.025)]
df, w = run(book(wb_rows))
d = df[(df["product"] == "GCV 4W") & (df.sub_product == "1+1") & (df.fuel_type == "Diesel")]
a, b = d[d.cc_min == 40000.01], d[d.cc_min == 43000.01]
check("the 40T-43T and 43T+ rows are written (no 'band not known', no unread cell): 40T-43T pays 2.5% for ages 2-5 and 5% from 6; 43T+ declines ages 2-5 (the cell's own '43T+: D') and pays 2.5% from 6",
      len(a) > 0 and len(b) > 0 and set(a[a.vehicle_age_min.between(2, 5)].pi_net_rate) == {2.5} and set(a[a.vehicle_age_min >= 6].pi_net_rate) == {5.0}
      and set(b[b.vehicle_age_min.between(2, 5)].pi_net_rate) == {0.0} and set(b[b.vehicle_age_min >= 6].pi_net_rate) == {2.5} and set(b.cc_max) == {99999.0} and set(a.cc_max) == {43000.01},
      (len(a), len(b), sorted(set(b.pi_net_rate))))
check("... and nothing about them is flagged or listed as not written", (df.remarks == "").all() and not any("43T" in x["segment"] or "40T" in x["segment"] for x in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN[:2])

hh = [None, "Agency/PB Clusters", "Agency/PB Seg", "CD1", "Max CD2", "Avg CD2", "Max CD2", "Avg CD2"]
s11 = sheet("New 2W Comp & SATP wef 1st Nov", [[None, None, None, "1+1"], [None, None, None, None, None, None, "SATP"]], hh,
            [["x", "WB_Bad", "MC <= 180 Hero/Honda/TVS", 0.4, 0.325, 0.3, 0.26, 0.235], ["x", "WB_Bad", "MC <= 180 Others", 0.2, 0.29, 0.265, 0.2, 0.185],
             ["x", "Goa", "MC <= 180 Hero/Honda", 0.4, 0.47, 0.435, 0.52, 0.485], ["x", "Goa", "MC <= 180 Others", 0.4, 0.4, 0.35, 0.3, 0.25]])
df, w = run(book([lcv], extra=[s11]))
t = df[(df["product"] == "TW") & (df.sub_product == "1+1")]
mk = lambda c: sorted(set(t[t.new_rto_list.str.upper().str.endswith("_" + c.upper())].new_vehicle_makes.str.upper()))          # build() writes the portal group names GO-DIGIT_OCT26_2W_1+1_TP_<cluster>
check("two-wheeler 180 cc: a cluster whose Hero / Honda segment also holds TVS writes HERO,HONDA,TVS and its 'Others' without TVS; a cluster with the plain Hero / Honda segment is unchanged (Others = TVS added back)",
      mk("WB_Bad") == sorted(["HERO,HONDA,TVS", "BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,SUZUKI"]) and mk("Goa") == sorted(["HERO,HONDA", "BIKEALL_EX_HERO_HONDA_TVS_SUZUKI,TVS,SUZUKI"]), (mk("WB_Bad"), mk("Goa")))
check("... and the Hero/Honda/TVS row is not listed as 'not written'", not any("Hero/Honda/TVS" in x["segment"] for x in mod.LAST_NOT_WRITTEN), mod.LAST_NOT_WRITTEN[:2])

# ================================================================ 9c. the car third-party notes that state their own slice ('More than 1500CC and 10+age is declined')
OPEN, AGE_OPEN = mod.CC_OPEN_END, mod.AGE_OPEN_END
check("TP notes: the declined slice is read from the words (cc band, age band); a note that names no slice ('Few RTOs for 10+ age are declined') is not a slice",
      mod.tp_decline_slice("More than 1500CC and 10+age is declined") == ((1500.01, OPEN), (10.01, AGE_OPEN)) and mod.tp_decline_slice("1000-1500CC and <10 age is declined") == ((1000.01, 1500.01), (0, 10.01))
      and mod.tp_decline_slice("Few RTOs for 10+ age are declined") is None and mod.tp_decline_slice("More than 1500CC is declined") is None)
SD = mod.split_declined
check("TP notes: a row of age 'All' loses its 10+ part, a row 'Petrol>1000' loses its 1000-1500 / <10 corner, a row that is the declined slice gets no card at all, a row the slice does not touch is not split",
      SD((1500.01, OPEN), (0, AGE_OPEN), ((1500.01, OPEN), (10.01, AGE_OPEN))) == [((1500.01, OPEN), (0, 10.01), True), ((1500.01, OPEN), (10.01, AGE_OPEN), False)]
      and SD((1000.01, OPEN), (0, AGE_OPEN), ((1000.01, 1500.01), (0, 10.01))) == [((1000.01, 1500.01), (0, 10.01), False), ((1000.01, 1500.01), (10.01, AGE_OPEN), True), ((1500.01, OPEN), (0, AGE_OPEN), True)]
      and SD((1500.01, OPEN), (10.01, AGE_OPEN), ((1500.01, OPEN), (10.01, AGE_OPEN))) == [((1500.01, OPEN), (10.01, AGE_OPEN), False)] and SD((0, 1000.01), (10.01, AGE_OPEN), ((1500.01, OPEN), (10.01, AGE_OPEN))) is None)
tp = sheet(" PVT Car TP", [], ["Cluster", "Segment", "Age", "Max CD2", "Note"],
           [["AP TS_Good2", "Petrol>1500", "All", 0.265, "More than 1500CC and 10+age is declined"], ["HR_Bad", "Petrol>1000", "All", 0.05, "1000-1500CC and <10 age is declined"],
            ["Delhi", "Petrol<1000", "All", 0.145, "Few RTOs for 10+ age are declined"], ["GJ_Good", "Diesel<1500", "All", 0.1, None]])
df, w = run(book([lcv], extra=[tp]))
t = df[(df["product"] == "Private Car") & (df.sub_product == "STP") & (df.fuel_type == "Petrol")]
ap = t[t.new_rto_list.str.upper().str.endswith("AP_TS_GOOD2")].sort_values("vehicle_age_min")
hr_ = t[t.new_rto_list.str.upper().str.endswith("HR_BAD")].sort_values(["cc_min", "vehicle_age_min"])
dl = t[t.new_rto_list.str.upper().str.endswith("DELHI")]
check("TP notes end to end: AP TS_Good2 pays 26.5 under 10 years; from 10 years (> 1500 cc) no card is written (listed in 'Not written'), no flag",
      [(r.cc_min, r.vehicle_age_min, r.vehicle_age_max, r.pi_tp_rate) for r in ap.itertuples()] == [(1500.01, 0.0, 10.01, 26.5)] and any("AP TS_Good2" in x["cluster"] and "declined by the grid" in x["reason"] for x in mod.LAST_NOT_WRITTEN) and (ap.remarks == "").all(), ap[["cc_min", "vehicle_age_min", "pi_tp_rate", "remarks"]].values.tolist())
check("TP notes end to end: HR_Bad 'Petrol>1000' = no card for 1000-1500 cc under 10, 1000-1500 cc from 10 pays 5, above 1500 cc pays 5 at every age (the very rows of the user's October upload), no flag",
      [(r.cc_min, r.cc_max, r.vehicle_age_min, r.vehicle_age_max, r.pi_tp_rate) for r in hr_.itertuples()] == [(1000.01, 1500.01, 10.01, 99.0, 5.0), (1500.01, 99999.0, 0.0, 99.0, 5.0)] and (hr_.remarks == "").all(),
      hr_[["cc_min", "cc_max", "vehicle_age_min", "vehicle_age_max", "pi_tp_rate", "remarks"]].values.tolist())
check("TP notes end to end: 'Few RTOs for 10+ age are declined' names no RTOs: the row is written as it is and stays flagged", len(dl) > 0 and dl.remarks.str.contains("Few RTOs for 10+ age", regex=False).all(), dl.remarks.tolist())

# ================================================================ 10. the real November workbook: which CD1 wordings are still not read?
nov = os.path.join(helpers.INPUTS, "godigit_nov26.xlsx")
if os.path.exists(nov):
    rows, _, _ = mod.read_cv_grid(mod.read_workbook(nov), mod.Asker(interactive=False), quiet)
    refused = sorted({mod.norm(g.cd1) for g in rows if g.cd1 is not None and mod.cd1_slots(g.cd1, g.comp)[0] is None})
    check("November workbook: the only CD1 wordings no reader understands are 'Hyundai: 70%, Maruti 85%' (a cap per make: read by the taxi make grammar) and the TATA New-only wording with '-100%'",
          refused == ["Hyundai: 70%, Maruti 85%", "New: With Addon: 85/90%|Without Addon: -100%"], refused)
    mod.MISSING_TABLES = "ask"
    seen_prompts = iter(["1"])                                    # the one question of the November workbook: the private-car slab ('> 5 Lacs' is the only one)

    def _answer(prompt):
        try:
            return next(seen_prompts)
        except StopIteration:
            raise EOFError
    mod.reset_run_lists()
    df, w = mod.build(nov, F, T, quiet, mod.Asker(interactive=True, input_fn=_answer, out=lambda *x: None))
    flagged = int((df.remarks != "").sum())
    check("November workbook: 828 rows are flagged (was 16,690 on 10 Nov, 436 on 11 Nov) - the 436 of 11 Nov (the TATA New-only wording 320, New / 1 Yr left out of a text 80, TAXI_ALL_EX_HINDUSTAN missing from the master 24, "
          "the 'Few RTOs declined' notes 12) + the 18 Nov wording flags (Tractor Electric and Low HP 160, the E-Rickshaw TPPD bodies 64) + four more make groups the master lacks "
          "(GCV_4W_ALL_EX_MAHINDRA 80, PCV_3W_BAJAJ_TVS 56, PCV_2W_ALL_EX_BAJAJ 16, PCV_3W_BAJAJ_RE 16)",
          flagged == 828, flagged)
else:
    print("note: regression/inputs/godigit_nov26.xlsx is missing - the real-workbook checks were skipped")

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
