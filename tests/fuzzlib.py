"""Shared pieces of the fuzzers: workbook mutations, a user who answers every question at random, output checks."""
import os as _os
ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))      # the grid_tools folder
import copy
import datetime as dt
import math
import random
import re

import pandas as pd

WEIRD = [None, "", " ", "  ", "NA", "N/A", "na", "n/a", "-", "--", "TBD", "#N/A", "#REF!", "#DIV/0!", "#VALUE!", "Block",
         "BLOCK", "IRDA", "IRDA*", "irda", "25%", "0.25", "0,25", "25", "0.5.5", "'0.5", "x", "Ω", "x" * 300, "a\u00a0b",
         "\n", "\t0.2", " 12 ", 25, 0.25, 1.5, 100, 101, 0, 1, -0.1, -5, 1e9, 1e-9, float("nan"), float("inf"), float("-inf"),
         True, False, dt.datetime(2026, 1, 1), dt.date(2026, 1, 1), 10 ** 20, b"bytes"]


def width(rows):
    return max((len(r) for r in rows), default=0)


def pick_sheet(S, rng):
    return rng.choice(list(S)) if S else None


# ---------------------------------------------------------------- mutations: each takes (S, rng) and edits S in place
def m_rename_sheet(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        new = rng.choice([n + " v2", n.upper(), n.replace(" ", "_"), " " + n, n[:5], "Sheet1", "", n.lower(), n + " "])
        S[new] = S.pop(n)


def m_delete_sheet(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        del S[n]


def m_empty_sheet(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        S[n] = rng.choice([[], [[]], [[None] * 5] * 10, [["", " "]] * 3])


def m_header_only(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        S[n] = S[n][: rng.randint(1, 8)]


def m_swap_sheets(S, rng):
    if len(S) >= 2:
        a, b = rng.sample(list(S), 2)
        S[a], S[b] = S[b], S[a]


def m_extra_sheet(S, rng):
    S[rng.choice(["Notes", "Sheet2", "Misc"])] = [["note", "x"], [1, 2], [None, None]]


def m_delete_rows(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        a = rng.randrange(len(S[n]))
        del S[n][a: a + rng.randint(1, 30)]


def m_blank_rows(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        a = rng.randrange(len(S[n]))
        S[n][a:a] = [[None] * width(S[n]) for _ in range(rng.randint(1, 4))]


def m_dup_rows(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        a = rng.randrange(len(S[n]))
        S[n][a:a] = [list(r) for r in S[n][a: a + rng.randint(1, 6)]]


def m_shuffle_rows(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and len(S[n]) > 8:
        for _ in range(rng.randint(1, 6)):
            i, j = rng.randrange(len(S[n])), rng.randrange(len(S[n]))
            S[n][i], S[n][j] = S[n][j], S[n][i]


def m_ragged(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        for _ in range(rng.randint(1, 40)):
            if S[n]:
                r = rng.randrange(len(S[n]))
                S[n][r] = S[n][r][: rng.randint(0, max(1, len(S[n][r])))]


def m_blank_col0(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        for _ in range(rng.randint(1, 20)):
            r = rng.randrange(len(S[n]))
            if S[n][r]:
                S[n][r][0] = rng.choice([None, "", "  "])


def m_blank_col(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        c = rng.randrange(max(1, width(S[n])))
        for r in S[n][rng.randint(0, 5):]:
            if c < len(r):
                r[c] = None


def m_delete_col(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        c = rng.randrange(max(1, width(S[n])))
        for r in S[n]:
            if c < len(r):
                del r[c]


def m_insert_col(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        c = rng.randrange(max(1, width(S[n])))
        filler = rng.choice([None, "SPOC", "Name", 0.1, "x"])
        for r in S[n]:
            r.insert(min(c, len(r)), filler)


def m_swap_cols(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        w = width(S[n])
        if w >= 2:
            a, b = rng.sample(range(w), 2)
            for r in S[n]:
                r.extend([None] * (max(a, b) + 1 - len(r)))
                r[a], r[b] = r[b], r[a]


def m_dup_col(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        c = rng.randrange(max(1, width(S[n])))
        for r in S[n]:
            if c < len(r):
                r.insert(c + 1, r[c])


def m_header_weird(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        for _ in range(rng.randint(1, 8)):
            r = rng.randrange(min(8, len(S[n])))
            if S[n][r]:
                S[n][r][rng.randrange(len(S[n][r]))] = rng.choice(WEIRD)


def m_cells(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        for _ in range(rng.randint(1, 15)):
            r = rng.randrange(len(S[n]))
            if S[n][r]:
                S[n][r][rng.randrange(len(S[n][r]))] = rng.choice(WEIRD)


def m_cells_data(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and len(S[n]) > 8:
        for _ in range(rng.randint(1, 60)):
            r = rng.randrange(6, len(S[n]))
            if len(S[n][r]) > 3:
                S[n][r][rng.randrange(2, len(S[n][r]))] = rng.choice(WEIRD)


def m_scale(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        k = rng.choice([100, 0.01, 10])
        for r in S[n]:
            for i, v in enumerate(r):
                if i > 1 and isinstance(v, (int, float)) and not isinstance(v, bool):
                    r[i] = v * k


def m_text_numbers(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        form = rng.choice(["{}", "{:g}%", "{:.2f}", " {} "])
        for r in S[n]:
            for i, v in enumerate(r):
                if i > 1 and isinstance(v, (int, float)) and not isinstance(v, bool) and rng.random() < 0.5:
                    r[i] = form.format(v * 100 if "%" in form else v)


def m_negatives(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        for r in S[n]:
            for i, v in enumerate(r):
                if i > 1 and isinstance(v, (int, float)) and not isinstance(v, bool) and v and rng.random() < 0.05:
                    r[i] = -v


def m_geo_names(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        for r in S[n][5:]:
            if r and isinstance(r[0], str) and rng.random() < 0.3:
                r[0] = rng.choice([r[0] + "\u00a0", r[0].lower(), r[0] + "  X", "  " + r[0], r[0] * 5, 12345, r[0].replace(" ", "\u00a0"),
                                   "Ünïcode ✓ " + r[0], "", None])


def m_tiny(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        S[n] = S[n][: rng.randint(2, 14)]


def m_wide(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        for r in S[n]:
            r.extend([None] * 150)


def m_tile(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and len(S[n]) > 8:
        body = S[n][6:]
        S[n] = S[n] + [list(r) for r in body[: len(body) // 2]]


def m_header_blank(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and S[n]:
        for r in S[n][: min(8, len(S[n]))]:
            for c in range(len(r)):
                if rng.random() < 0.15:
                    r[c] = None


def m_slab_names(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        for r in S[n][5:]:
            for i in range(1, min(len(r), 5)):
                if isinstance(r[i], str) and ("L" in r[i] or "k" in r[i]) and rng.random() < 0.5:
                    r[i] = r[i] + " "


def m_cells_dense(S, rng):
    """turn a big share of the rate cells into odd values - exercises the 'same kind of problem' memory"""
    n = pick_sheet(S, rng)
    if n is not None and len(S[n]) > 8:
        share = rng.choice([0.02, 0.1, 0.3, 0.8])
        pool = rng.sample(WEIRD, rng.randint(1, 6))
        for r in S[n][5:]:
            for i in range(2, len(r)):
                if isinstance(r[i], (int, float)) and not isinstance(r[i], bool) and rng.random() < share:
                    r[i] = rng.choice(pool)


def m_label_noise(S, rng):
    """header labels the way people actually mangle them: case, stray spaces, NBSP, trailing dots, numbering"""
    n = pick_sheet(S, rng)
    if n is not None:
        for r in S[n][: min(8, len(S[n]))]:
            for i, v in enumerate(r):
                if isinstance(v, str) and v and rng.random() < 0.3:
                    r[i] = rng.choice([v.upper(), v.lower(), v + " ", " " + v, v.replace(" ", "  "), v.replace(" ", " "), v + ".", "01. " + v, v + "\t"])


def m_prefix_rows(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        junk = [[rng.choice(["Note", "Circular", None, "Effective 1st", 2026])] + [None] * 3 for _ in range(rng.randint(1, 5))]
        S[n] = junk + S[n]


def m_footnotes(S, rng):
    n = pick_sheet(S, rng)
    if n is not None:
        S[n] = S[n] + [[None] * 3] * rng.randint(0, 2) + [["* Rates are subject to the notes below " + "x" * rng.randint(0, 80)] + [None] * 3
                                                              for _ in range(rng.randint(1, 4))]


def m_blank_gap(S, rng):
    n = pick_sheet(S, rng)
    if n is not None and len(S[n]) > 12:
        a = rng.randrange(6, len(S[n]) - 1)
        S[n].insert(a, [None] * width(S[n]))


MUTATIONS = [m_cells_dense, m_label_noise, m_prefix_rows, m_footnotes, m_blank_gap, m_rename_sheet, m_delete_sheet, m_empty_sheet, m_header_only, m_swap_sheets, m_extra_sheet, m_delete_rows, m_blank_rows,
             m_dup_rows, m_shuffle_rows, m_ragged, m_blank_col0, m_blank_col, m_delete_col, m_insert_col, m_swap_cols, m_dup_col,
             m_header_weird, m_cells, m_cells_data, m_scale, m_text_numbers, m_negatives, m_geo_names, m_tiny, m_wide, m_tile,
             m_header_blank, m_slab_names]


def mutate(sheets, rng, how_many=None):
    """A mutated deep copy of {sheet: rows} and the names of the mutations applied"""
    S = copy.deepcopy(sheets)
    chosen = [rng.choice(MUTATIONS) for _ in range(how_many or rng.randint(1, 5))]
    for m in chosen:
        m(S, rng)
    return S, [m.__name__ for m in chosen]


# ---------------------------------------------------------------- a user who answers every question at random
class Overload(Exception):
    pass


class RandomUser:
    def __init__(self, rng, max_prompts=1500, stop_prob=0.0):
        self.rng, self.max, self.stop_prob = rng, max_prompts, stop_prob
        self.menu, self.question, self.prompts, self.same, self.last = [], "", 0, 0, None
        self.transcript, self.labels = [], {}

    def out(self, s):
        s = str(s)
        self.transcript.append(s)
        if s.startswith("\n? "):
            self.menu, self.question, self.labels = [], s[3:], {}
        else:
            m = re.match(r"\s+(\d+)\) (.*)", s)
            if m:
                self.menu.append(int(m.group(1)))
                self.labels[int(m.group(1))] = m.group(2)

    def inp(self, prompt=""):
        self.prompts += 1
        if self.prompts > self.max:
            raise Overload(f"more than {self.max} prompts; last question: {self.question[:200]}")
        key = (prompt, self.question)
        self.same = self.same + 1 if key == self.last else 0
        self.last = key
        r = self.rng
        if self.same > 25:                   # a loop of refused answers: give up like a person would
            return "q"
        if self.same > 8 and prompt.startswith("   Type 1-"):
            return "1"
        if r.random() < self.stop_prob:
            return "q"
        if prompt.startswith("   Type 1-"):
            n = int(re.match(r"   Type 1-(\d+)", prompt).group(1))
            if r.random() < 0.02:
                return "?"
            if r.random() < 0.03:
                return r.choice(["0", str(n + 1), "abc", "", "1,,", "-3"])
            pool = list(range(1, n + 1))
            keep_going = [i for i in pool if not self.labels.get(i, "").lower().startswith(("stop", "no, stop"))]
            if keep_going and r.random() < 0.97:              # mostly avoid the "stop here" options
                pool = keep_going
            if "several like" in prompt:
                k = r.choice([1, 1, 1, 2, 3])
                ids = sorted(r.sample(pool, min(k, len(pool))))
                if len(ids) >= 2 and ids == list(range(ids[0], ids[-1] + 1)) and r.random() < 0.5:
                    return f"{ids[0]}-{ids[-1]}"
                return ",".join(map(str, ids))
            return str(r.choice(pool))
        if prompt.strip() == ">":
            q = self.question.lower()
            if "row number" in q:
                return r.choice([str(r.randint(1, 60)), str(r.randint(1, 9)), "abc", "0", "99999"])
            return r.choice(["nofile.xlsx", "C:\\no\\such\\file.xlsx", "out2.csv"])
        if "[Y/n]" in prompt or "[y/N]" in prompt:
            return r.choice(["", "", "", "y"]) if r.random() < 0.93 else r.choice(["n", "maybe", "?"])
        return ""


# ---------------------------------------------------------------- checks on a written upload file
def check_csv(path, columns, numeric):
    d = pd.read_csv(path, dtype=str, keep_default_na=False)
    problems = []
    if list(d.columns) != list(columns):
        problems.append("header differs")
    if len(d) == 0:
        problems.append("no rows")
    low = d.drop(columns=["new_rto_list"], errors="ignore").astype(str).apply(lambda s: s.str.lower())      # a geo may really be called "inf"
    for bad in ("nan", "inf", "-inf", "none"):
        if (low == bad).any().any():
            problems.append(f"'{bad}' text in the file")
    for c in numeric:
        if c in d:
            v = pd.to_numeric(d[c], errors="coerce")
            if v.isna().any() or not v.map(math.isfinite).all():
                problems.append(f"column {c} has non-numbers")
    for c in ("from_date", "to_date"):
        if c in d and not d[c].str.fullmatch(r"\d{2}/\d{2}/\d{4}").all():
            problems.append(f"{c} not DD/MM/YYYY")
    return problems
