"""The insurer reference files and the rule 'no Excel cell references in documentation' (the user, 23 Nov 2026: "each individual insurer has its own dedicated .md file and its
own dedicated Skill file ... no hardcoded Excel cell references (specific row numbers or column letters) anywhere in these files ... purely rule-based and logic-based so they
remain reliable even if the Excel layout changes").

Every converter script has its insurer's dedicated reference <INSURER>_GRID.md next to its skill; the skill, grid-converter and CLAUDE.md point at it; no document or skill holds a
cell address, a column range, a row number or a column letter (.claude/skills/grid-converter/scripts/cellref_lint.py); learn.py refuses such text.
The examples of forbidden text below are in this test file on purpose: it is code, not documentation.
Run:  .venv\\Scripts\\python.exe tests\\insurer_docs_tests.py"""
import glob
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

ROOT = helpers.ROOT
SKILLS = os.path.join(ROOT, ".claude", "skills")
SCRIPTS = os.path.join(SKILLS, "grid-converter", "scripts")
results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def load(name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(SCRIPTS, name + ".py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def read(path):
    with open(path, encoding="utf-8") as fh:
        return fh.read()


lint = load("cellref_lint")
learn = load("learn")

scripts = sorted(os.path.basename(p) for p in glob.glob(os.path.join(ROOT, "*_to_config.py")))
insurers = sorted({s.split("_")[0] for s in scripts})
check("there are converter scripts to cover", len(scripts) >= 9 and {"shriram", "sbi", "rsa", "liberty", "magma", "united", "godigit", "icici", "indusind"} <= set(insurers), insurers)

# ---------------------------------------------------------------- one dedicated reference per insurer
NEEDED = {"at a glance": r"at a glance", "checking a run": r"checking a run", "decisions": r"\bdecisions\b", "open": r"\bopen\b", "files": r"\bfiles\b"}
for ins in insurers:
    md = f"{ins.upper()}_GRID.md"
    path = os.path.join(ROOT, md)
    if not os.path.isfile(path):
        check(f"{md}: every converter script has its insurer's dedicated reference in the project root", False, f"create {md} (copy the nearest insurer's file)")
        continue
    text = read(path)
    low = text.lower()
    heads = [ln.lower() for ln in text.splitlines() if ln.startswith("#")]
    check(f"{md}: exists and is a real reference (not a stub)", len(text) > 3000 and text.startswith("# "), len(text))
    missing = [k for k, rx in NEEDED.items() if not any(re.search(rx, h) for h in heads)]
    check(f"{md}: has the sections At a glance / Checking a run / Decisions / Open / Files", not missing, missing)
    check(f"{md}: names its script(s) and its skill", all(s in text for s in scripts if s.split("_")[0] == ins) and f".claude/skills/{ins}-grid/SKILL.md" in text,
          [s for s in scripts if s.split("_")[0] == ins and s not in text])
    check(f"{md}: says it is rule-based and tells how to check a run BY LABELS", "rule-based" in low and re.search(r"spot-check[^.\n]*by labels", low) is not None)
    skill = read(os.path.join(SKILLS, f"{ins}-grid", "SKILL.md"))
    check(f"{ins}-grid skill points at {md}", md in skill)

conv = read(os.path.join(SKILLS, "grid-converter", "SKILL.md"))
claude = read(os.path.join(ROOT, "CLAUDE.md"))
mds = [f"{i.upper()}_GRID.md" for i in insurers]
check("grid-converter points at every insurer's dedicated reference and at the lint", all(m in conv for m in mds) and "cellref_lint" in conv, [m for m in mds if m not in conv])
check("CLAUDE.md lists every insurer's dedicated reference and states the rule-based rule", all(m in claude for m in mds) and "cellref_lint.py" in claude and "rule-based" in claude, [m for m in mds if m not in claude])

# ---------------------------------------------------------------- no cell references anywhere in the documentation and the skills
docs = lint.doc_files(ROOT)
bad = {p: lint.find_in_file(p) for p in docs}
bad = {p: h for p, h in bad.items() if h}
check(f"no document or skill names a place in a grid by a cell address, a column range, a row number or a column letter ({len(docs)} files)", not bad,
      "; ".join(f"{os.path.relpath(p, ROOT)}: {lint.describe(h, 3)}" for p, h in list(bad.items())[:4]))
check("the lint reads the insurer files, the skills, CLAUDE.md and AGENT_CONTEXT.md (it is not scanning nothing)",
      all(any(os.path.basename(p) == m for p in docs) for m in mds + ["CLAUDE.md", "AGENT_CONTEXT.md"]) and sum(1 for p in docs if p.endswith("SKILL.md")) >= 10, len(docs))


# ---------------------------------------------------------------- the lint itself: every kind it must find, and the names it must leave alone
def kinds(text):
    return {h.kind for h in lint.find(text)}


MUST_FIND = [
    ("an address", "the payout is in D5 of the sheet", "address"),
    ("an absolute address", "use $B$3 as the anchor", "address"),
    ("a sheet-qualified address", "see Sheet1!K12 for the rate", "address"),
    ("a column range", "the block K:L has the make label", "column range"),
    ("a row number", "the title sits in row 1", "row number"),
    ("a list of rows", "the header rows 2-4 hold the band", "row number"),
    ("a column letter", "the payouts start in column D", "column letter"),
    ("a range of column letters", "columns A-C hold the region", "column letter"),
    ("a column letter in brackets", "the qualifier columns (K-N) are flagged", "column letter"),
    ("a column number", "read column 3 of the table", "column number"),
    ("an ordinal column", "the 5th column is the payout", "column number"),
    ("a code index", "never use rows[1] here", "index"),
    ("a fixed scan depth", "the reader used scan=4 once", "index"),
    ("worksheet indexing", 'ws["D5"].value gives the payout', "index"),
]
for what, text, kind in MUST_FIND:
    check(f"the lint finds {what}", kind in kinds(text), (text, kinds(text)))

MUST_PASS = [
    "RTO codes MH01, HR38 and UP14 are state codes",
    "the groups were made in APR26 and the grid of OCT26 for FY26",
    "Max CD2 and Avg CD2 are headers and CD1 is the OD-discount band",
    "the master export has columns ID, RTO NAME and RTO CLUSTER",
    "the flat card has rate columns 0, then the type",
    "the user's own upload dated its rows 01/11",
    "3,650 rows are flagged and 2.5 rows per cluster",
    "a pattern id after the word column (P-35) and a failure id (F-45) are no letters",
    "press F5 to start debugging in VS Code",
    "the first column that says a field is the one taken",
    "the upload has 42 columns. A note is a flag",
    "COLUMN BY COLUMN, never by sums",
    "the V2 grid of 22 Oct rewrote the notes",
    "one row per cluster and one column per segment",
]
for text in MUST_PASS:
    check(f"the lint leaves alone: {text[:60]}", not lint.find(text), [(h.kind, h.token) for h in lint.find(text)])

table = "| Col | Header |\n|---|---|\n| D | School Bus |\n| E, F | PCV 3W |\n| G | Taxi |\n"
check("the lint finds a table whose first column runs through column letters", "letter table" in kinds(table), kinds(table))
check("... and leaves a table of product abbreviations alone", not lint.find("| Product | Rows |\n|---|---|\n| TW | 10 |\n| PC | 20 |\n| CV | 30 |\n"))
sample = read(os.path.join(ROOT, "INDUSIND_GRID.md"))
planted = sample + "\nthe payout is in D5 of the sheet\n"
check("put-back bug: an address added to a real insurer file is found (and only that line)",
      [h.line for h in lint.find(planted)] == [planted.splitlines().index("the payout is in D5 of the sheet") + 1])

# ---------------------------------------------------------------- learn.py refuses positions and writes nothing
tmp = tempfile.mkdtemp(prefix="insurer_docs_test_")
try:
    shutil.copytree(SKILLS, os.path.join(tmp, ".claude", "skills"), ignore=shutil.ignore_patterns("__pycache__"))
    skill_file = os.path.join(tmp, ".claude", "skills", "sbi-grid", "SKILL.md")
    before = read(skill_file)
    for text in ("the payin is in cell D5", "the header sits in row 4", "the payouts start in column G", "the block K:L is the qualifier"):
        try:
            learn.add_skill_note(tmp, "sbi", "traps", text, "17 Feb 2026")
            check(f"learn.py skill refuses: {text}", False, "it wrote the bullet")
        except SystemExit as e:
            check(f"learn.py skill refuses: {text}", "refused" in str(e) and "labels" in str(e), str(e)[:160])
    check("a refused text writes nothing", read(skill_file) == before)
    path = learn.add_skill_note(tmp, "sbi", "traps", "the Mumbai row under the School Bus header chain was misread", "17 Feb 2026")
    check("a text that names its place by labels is written", "- (17 Feb 2026) the Mumbai row under the School Bus header chain was misread" in read(path))
    for what, call in (("lesson", lambda: learn.add_lesson(tmp, "sbi", "payout read from row 5", "cause", "fix", "covered")),
                       ("note", lambda: learn.add_note(tmp, "failure", "symptom in column D -> cause -> rule")),
                       ("lesson (cause)", lambda: learn.add_lesson(tmp, "sbi", "title", "the cell D5 was misread", "fix", "covered"))):
        try:
            call()
            check(f"learn.py {what} refuses a position", False, "it did not refuse")
        except SystemExit as e:
            check(f"learn.py {what} refuses a position", "refused" in str(e), str(e)[:160])
    LEARN = os.path.join(SCRIPTS, "learn.py")
    p = subprocess.run([sys.executable, "-I", "-X", "utf8", LEARN, "--root", tmp, "skill", "--insurer", "sbi", "--section", "traps", "--text", "see cell D5"],
                       capture_output=True, text=True, encoding="utf-8")
    check("the command line refuses too (non-zero exit, the reason on stderr)", p.returncode != 0 and "refused" in p.stderr and "cell D5" not in read(skill_file), (p.returncode, p.stderr[-160:]))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---------------------------------------------------------------- the gates are wired
check("check_regressions.py checks the dedicated references and the lint", "cellref_lint" in read(os.path.join(ROOT, "check_regressions.py")) and "_GRID.md" in read(os.path.join(ROOT, "check_regressions.py")))
check("learn.py calls the refusal in lesson, note and skill", read(os.path.join(SCRIPTS, "learn.py")).count("refuse_cell_references(") >= 4)
check("run_tests.py lists this suite", "insurer_docs_tests.py" in read(os.path.join(ROOT, "tests", "run_tests.py")))

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
