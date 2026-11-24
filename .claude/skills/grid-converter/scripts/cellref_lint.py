"""Excel cell references in the documentation and the skills (the user, 23 Nov 2026: "no hardcoded Excel cell references (specific row numbers or
column letters) anywhere in these files ... all instructions, documentation and skills must be purely rule-based and logic-based so they remain
reliable even if the Excel layout changes").

A place in a grid is named by its LABELS - the words of a sheet's title, of a header chain, of a row's label cell - and by the RULE that finds it,
never by an address, a range, a row number or a column letter, because the next month's grid moves every one of them. This module finds the
forbidden forms in a text:

    address         a cell address (a column letter group + a row number, with or without dollar signs, e.g. in a sheet-qualified reference)
    column range    letters on both sides of a colon (a block of columns)
    row number      the word row / rows + a number (or a list or range of numbers)
    column letter   the word column / columns + a column letter (or a list or range of letters)
    column number   the word column + a number, or an ordinal + column / row ('the 5th column')
    index           the position idioms of code in a text: worksheet indexing, a cell call, an iloc index, a fixed scan depth
    letter table    a markdown table whose first column runs through column letters (a 'Col' column that lists D, E-F, G ...)

Not forbidden (they are names or rules, not positions): RTO codes (state code + two digits), month-year tokens, the insurers' own header words that
look like an address (VOCABULARY: Go Digit's discount columns, quoted segment words), the function key F5, a rule that says 'the first column that
says a field' (it is a rule about HOW a reader picks a column, not a place in a grid).

Standard library only; learn.py, check_regressions.py and the tests load it by path.

    python cellref_lint.py [file ...]       # default: every markdown file of the project that is documentation or a skill; exit 1 when one holds a reference
"""
import os
import re
import sys
from typing import NamedTuple

DEFAULT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), *([os.pardir] * 4)))

STATE_CODES = frozenset("AN AP AR AS BR CG CH DD DL DN GA GJ HP HR JH JK KA KL LA LD MH ML MN MP MZ NL OD OR PB PY RJ SK TN TR TS TG UK UP WB".split())
MONTH_CODES = frozenset("JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC".split())
# words of the insurers' own grids and of the project that merely LOOK like an address; each is a name, never a place in a workbook
VOCABULARY = frozenset({
    "CD1", "CD2",                    # Go Digit: the OD-discount band and the age-wise discount (header words)
    "GCV4", "HCV3", "PCV2",          # segment words quoted from grids ('GCV4 upto 2T', 'HCV3 upto 43.5T', 'PCV2 - without ZD')
    "XI6", "OLE2", "RE15",           # a car model in a quoted note, the OLE2 container error, the 'before 15th' text once read as an RTO code
    "HP1", "CG1",                    # cluster names (BIHAR without PURNIA / HP1, the CG1 geography)
    "V2",                            # the 'V2 grid' of Royal Sundaram
})
# two-letter capitals that follow the word 'column' as header words or plain words, not as column letters
HEADER_WORDS = frozenset("ID NA OD TP PO PI TW PC CV PV EV NB CC CD ZD UW TC SC GV SA OR NO BY IN ON TO OF IF IS AS AT SO UP".split())

ADDRESS = re.compile(r"(?<![\w$.])\$?([A-Z]{1,3})\$?([0-9]{1,7})(?!\w)")
COLUMN_RANGE = re.compile(r"(?<![\w$])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![\w:])")
ROWS = re.compile(r"\b[Rr]ows?\s+\d+(?![\d/.,%:])(?:\s*(?:-|–|to|and|,)\s*\d+(?![\d/.,%:]))?")
COLUMNS = re.compile(r"\b[Cc]olumns?\s*\(?\s*([A-Z]{1,2}\b(?!-\d)(?:\s*(?:-|–|:|/|,|&|and|to)\s*[A-Z]{1,2}\b)*)")     # '(P-35)' / '(F-45)' after the word column is a pattern / failure id, not a letter
COLUMN_NUMBER = re.compile(r"\b[Cc]ol(?:umn)?\s+\d+(?![\d/.,%])|\b[Cc]olumns\s+\d+\s*(?:-|–|to|and)\s*\d+")
ORDINAL_NUMBER = re.compile(r"\b\d+(?:st|nd|rd|th)\s+(?:column|row)s?\b", re.I)
INDEX = re.compile(r"\bws\[|\.cell\(|\.iloc\[|\.loc\[|\brows?\[\d|\br\[\d|\bscan=\d|\[\d+\]\[\d+\]|\brow=\d|\bcolumn=\d|\bmin_row\b|\bmax_col\b")
TABLE_FIRST = re.compile(r"^[A-Z]{1,3}(?:\s*[,/–-]\s*[A-Z]{1,3})*$")


class Hit(NamedTuple):
    line: int
    kind: str
    token: str
    text: str


def _is_vocabulary(token, line):
    """A token that looks like an address but is a name: an RTO code, a month-year, a fiscal year, one of VOCABULARY, the function key F5 in a debugging sentence."""
    if "$" in token:
        return False
    m = re.fullmatch(r"([A-Z]{1,3})(\d+)", token)
    if not m:
        return False
    letters, digits = m.groups()
    if token in VOCABULARY:
        return True
    if len(digits) == 2 and (letters in STATE_CODES or letters in MONTH_CODES or letters == "FY"):
        return True
    return letters == "F" and bool(re.search(r"(?i)debug|press|launch\.json|run and", line))


def _column_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _letter_tables(lines):
    """A markdown table whose first column runs through column letters (three rows in a row with growing letters) - a 'Col' column."""
    hits, run = [], []
    for no, line in enumerate(lines, 1):
        s = line.strip()
        if not s.startswith("|"):
            run = []
            continue
        first = s.strip("|").split("|")[0].strip().strip("`*_ ")
        if not first or set(first) <= set("-: "):
            continue                                      # the separator row does not break a run
        if TABLE_FIRST.fullmatch(first):
            idx = _column_index(re.match(r"[A-Z]{1,3}", first).group(0))
            run = run + [(no, idx, first)] if run and idx > run[-1][1] else [(no, idx, first)]
            if len(run) == 3:
                hits.append(Hit(run[0][0], "letter table", ", ".join(r[2] for r in run) + " ...", lines[run[0][0] - 1]))
        else:
            run = []
    return hits


def find(text):
    """Every forbidden reference of a text, as Hit(line, kind, token, text), in order."""
    lines = text.splitlines()
    hits = []
    for no, line in enumerate(lines, 1):
        for m in ADDRESS.finditer(line):
            if not _is_vocabulary(m.group(0), line):
                hits.append(Hit(no, "address", m.group(0), line))
        for m in COLUMN_RANGE.finditer(line):
            hits.append(Hit(no, "column range", m.group(0), line))
        for m in ROWS.finditer(line):
            hits.append(Hit(no, "row number", m.group(0), line))
        for m in COLUMNS.finditer(line):
            letters = re.findall(r"[A-Z]{1,2}\b", m.group(1))
            if any(t not in HEADER_WORDS for t in letters):
                hits.append(Hit(no, "column letter", m.group(0), line))
        for m in COLUMN_NUMBER.finditer(line):
            hits.append(Hit(no, "column number", m.group(0), line))
        for m in ORDINAL_NUMBER.finditer(line):
            hits.append(Hit(no, "column number", m.group(0), line))
        for m in INDEX.finditer(line):
            hits.append(Hit(no, "index", m.group(0), line))
    hits += _letter_tables(lines)
    return sorted(hits, key=lambda h: (h.line, h.kind, h.token))


def find_in_file(path):
    with open(path, encoding="utf-8") as fh:
        return find(fh.read())


def doc_files(root=DEFAULT_ROOT):
    """The documentation and the skills of the project: every markdown file in the root, the skills, the commands and the web app's README (never the vendored folders, the backups or the outputs)."""
    out = [os.path.join(root, n) for n in sorted(os.listdir(root)) if n.lower().endswith(".md") and os.path.isfile(os.path.join(root, n))]
    for sub in (os.path.join(".claude", "skills"), os.path.join(".claude", "commands")):
        for dirpath, dirs, files in os.walk(os.path.join(root, sub)):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            out += [os.path.join(dirpath, f) for f in sorted(files) if f.lower().endswith(".md")]
    readme = os.path.join(root, "webapp", "README.md")
    if os.path.isfile(readme):
        out.append(readme)
    return out


def describe(hits, limit=6):
    """One short sentence per hit for a refusal or a failure line."""
    shown = [f"line {h.line}: {h.kind} '{h.token}'" for h in hits[:limit]]
    return "; ".join(shown) + (f"; ... {len(hits) - limit} more" if len(hits) > limit else "")


ADVICE = ("Documentation and skills hold RULES, never positions: name a place in a grid by its labels (the words of the sheet's title, of the header chain, "
          "of the row's label) and by the rule that finds it, because the next month's grid moves every address, row number and column letter.")


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    files = args or doc_files()
    bad = 0
    for path in files:
        hits = find_in_file(path)
        for h in hits:
            print(f"{os.path.relpath(path, DEFAULT_ROOT) if not args else path}:{h.line}: {h.kind} '{h.token}'  | {h.text.strip()[:150]}")
        bad += len(hits)
    print(f"\n{bad} cell reference(s) in {len(files)} file(s)" + ("" if not bad else "\n" + ADVICE))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
