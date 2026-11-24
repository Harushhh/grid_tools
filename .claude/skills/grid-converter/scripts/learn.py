"""The file mechanics of the self-learning loop (see AGENT_CONTEXT.md section 12). Standard library only.

    learn.py lesson --insurer united --title "..." --cause "..." --fix "..." --covered "..." [--rule "..."] [--when 2026-10]
        adds an entry at the top of lessons_learned.md (newest first)

    learn.py case --name united_sep26 --script united_commission_to_config.py --from-date 01/09/2026 --to-date 30/09/2026
                  --input grid=C:\\path\\circular.pdf [--input satp=C:\\path\\second_file.xlsb] [--input ref=C:\\path\\their_upload.csv]
                  [--answers "" y y] [--note "what is special about this grid"] [--bless]
        copies the files into regression\\inputs (grid, other inputs) / regression\\references (ref), runs the script once the way
        check_regressions.py will, and records the measured rows / questions / flagged / new / removed / changed as the case's expectations
        in regression\\cases.json. Check the numbers with the user, then --bless (or run check_regressions.py --bless --only NAME).

    learn.py note --section pattern|failure|decision|open --text "..." [--title "..."] [--date 2026-10-26] [--reason "..."]
        adds a numbered pattern (P-n) / failure (F-n) / decision row / open item to AGENT_CONTEXT.md and refreshes its "Last updated" line

    learn.py skill --insurer shriram --section status|rules|traps|open --text "..." [--date "4 Nov 2026"]
        adds a dated bullet to the matching section of that insurer's skill (.claude/skills/<insurer>-grid/SKILL.md): the per-insurer knowledge
        the agent loads with grid-converter (liberty, magma, rsa, united, shriram, godigit, sbi, icici); a missing section is created

    learn.py stamp --reason "..."          refresh only the "Last updated" line of AGENT_CONTEXT.md

--root <folder> points at another copy of grid_tools (the default is the folder this skill lives in). Nothing here touches the portal repository.

Every command that writes prose (lesson, note, skill) REFUSES a text that names a place in a grid by a cell address, a column range, a row number or a column letter
(cellref_lint.py; the user's rule of 23 Nov 2026: documentation and skills are rule-based) - say it by labels and rules."""
import argparse
import datetime as dt
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys

DEFAULT_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), *([os.pardir] * 4)))


def read(path):
    with open(path, encoding="utf-8", newline="") as fh:
        return fh.read()


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)


def eol(text):
    return "\r\n" if "\r\n" in text else "\n"


# ---------------------------------------------------------------- no cell references in documentation (the user, 23 Nov 2026)
def _lint_module():
    """cellref_lint.py next to this file, loaded by path (so it works however learn.py is started); None when it is not there."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cellref_lint.py")
    if not os.path.isfile(path):
        return None
    spec = importlib.util.spec_from_file_location("cellref_lint", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def refuse_cell_references(what, *texts):
    """Documentation and skills hold RULES, never positions: stop (SystemExit, nothing is written) when a text names a place in a grid by a cell address, a column range,
    a row number or a column letter."""
    lint = _lint_module()
    if lint is None:
        return
    hits = lint.find("\n".join(t for t in texts if t))
    if hits:
        raise SystemExit(f"refused: the {what} names a place in a grid by its position ({lint.describe(hits)}). {lint.ADVICE} Reword it by labels and rules, then run the command again.")


# ---------------------------------------------------------------- lesson
def add_lesson(root, insurer, title, cause, fix, covered, rule="", when=None):
    refuse_cell_references("lesson", title, cause, fix, covered, rule)
    path = os.path.join(root, "lessons_learned.md")
    text = read(path)
    nl = eol(text)
    when = when or dt.date.today().strftime("%Y-%m")
    lines = [f"## {when}  {insurer}  - {title}", f"**Cause:** {cause}", f"**Fix:** {fix}", f"**Covered by:** {covered}"]
    if rule:
        lines.append(f"**Rule:** {rule}")
    entry = nl.join(lines) + nl + nl
    m = re.search(r"^## \d{4}-\d{2}  ", text, re.M)               # the newest entry: insert before it
    if not m:
        raise SystemExit("lessons_learned.md has no '## YYYY-MM  insurer  - ...' entry to put the new one before")
    write(path, text[:m.start()] + entry + text[m.start():])
    return path


# ---------------------------------------------------------------- regression case
def load_checker(root):
    spec = importlib.util.spec_from_file_location("check_regressions_dyn", os.path.join(root, "check_regressions.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def add_case(root, name, script, from_date, to_date, inputs, answers, note, bless):
    inputs_dir, refs_dir = os.path.join(root, "regression", "inputs"), os.path.join(root, "regression", "references")
    os.makedirs(inputs_dir, exist_ok=True)
    os.makedirs(refs_dir, exist_ok=True)
    args = []
    for spec in inputs:
        flag, _, src = spec.partition("=")
        if not flag or not src or not os.path.isfile(src):
            raise SystemExit(f"--input {spec!r}: use flag=<path of an existing file> (flags: grid, satp, ref, ...)")
        ext = os.path.splitext(src)[1].lower()
        if flag == "ref":
            dest = os.path.join(refs_dir, f"{name}_rates_export{ext}")
        else:
            dest = os.path.join(inputs_dir, f"{name}{'' if flag == 'grid' else '_' + flag}{ext}")
        if os.path.abspath(dest) != os.path.abspath(src):
            shutil.copyfile(src, dest)
        args += [f"--{flag}", os.path.relpath(dest, root).replace(os.sep, "/")]
    args += ["--from-date", from_date, "--to-date", to_date]
    case = {"name": name, "script": script, "args": args, "answers": list(answers), "note": note, "expect": {}}
    if not any(a == "--grid" for a in args):
        raise SystemExit("one input must be --input grid=<path>")
    checker = load_checker(root)
    got = checker.run_case(case)
    if got["exit"] != 0 or got["sha256"] is None:
        tail = [ln for ln in got["text"].splitlines() if ln.strip()][-4:]
        raise SystemExit("the script did not finish with these answers - fix that first (answers go to its questions in order):\n  " + "\n  ".join(tail))
    case["expect"] = {k: got[k] for k in ("rows", "questions", "new", "removed", "changed", "flagged") if got.get(k) is not None}
    cases_path = os.path.join(root, "regression", "cases.json")
    raw = read(cases_path)
    data = json.loads(raw)
    data["cases"] = [c for c in data["cases"] if c["name"] != name] + [case]
    write(cases_path, json.dumps(data, indent=2, ensure_ascii=False) + (eol(raw) if raw.endswith(("\n", "\r\n")) else ""))
    print(f"case '{name}' written to regression/cases.json with expectations {case['expect']}")
    if bless:
        subprocess.run([sys.executable, "-I", os.path.join(root, "check_regressions.py"), "--bless", "--only", name], cwd=root, check=True)
    else:
        print(f"next: show the user the numbers, then  python check_regressions.py --bless --only {name}")
    return cases_path


# ---------------------------------------------------------------- insurer skills
SKILL_SECTIONS = {"status": ("Status", "status"), "rules": ("Rules the user has set (do not re-ask)", "rules the user"),
                  "traps": ("Traps seen", "traps seen"), "open": ("Open items", "open items")}


def skill_path(root, insurer):
    """.claude/skills/<insurer>-grid/SKILL.md ('shriram' and 'shriram-grid' both work); stops with the list of the skills that exist"""
    name = insurer if insurer.endswith("-grid") else insurer + "-grid"
    path = os.path.join(root, ".claude", "skills", name, "SKILL.md")
    if not os.path.isfile(path):
        known = sorted(d[:-5] for d in os.listdir(os.path.join(root, ".claude", "skills")) if d.endswith("-grid") and d != "grid-converter")
        raise SystemExit(f"no skill {name!r} ({path} is missing). Insurer skills that exist: {', '.join(known)}")
    return path


def add_skill_note(root, insurer, section, text, date=None):
    """Append '- (<date>) <text>' to the end of the matching ## section of the insurer's skill (the section is created at the end when it is not there)."""
    refuse_cell_references("skill note", text)
    path = skill_path(root, insurer)
    title, key = SKILL_SECTIONS[section]
    raw = read(path)
    nl = eol(raw)
    lines = raw.split(nl)
    when = date or f"{dt.date.today().day} {dt.date.today():%b %Y}"
    bullet = f"- ({when}) {' '.join(text.split())}"
    head = next((i for i, ln in enumerate(lines) if ln.startswith("## ") and key in ln.lower()), None)
    if head is None:
        while lines and not lines[-1].strip():
            lines.pop()
        lines += ["", f"## {title}", bullet, ""]
    else:
        end = next((i for i in range(head + 1, len(lines)) if lines[i].startswith("## ")), len(lines))
        at = end
        while at > head + 1 and not lines[at - 1].strip():
            at -= 1
        lines.insert(at, bullet)
        if end == len(lines) - 1 and lines[-1].strip():
            lines.append("")
    write(path, nl.join(lines))
    return path


# ---------------------------------------------------------------- AGENT_CONTEXT.md
SECTION_TITLES = {"pattern": "## 7. ", "failure": "## 9. ", "open": "## 11. "}      # the heading AFTER the section being extended


def context_path(root):
    return os.path.join(root, "AGENT_CONTEXT.md")


def _insert_before_rule(lines, heading_prefix):
    """index of the blank/rule lines in front of the next section's heading: new content goes there"""
    h = next((i for i, ln in enumerate(lines) if ln.startswith(heading_prefix)), None)
    if h is None:
        raise SystemExit(f"AGENT_CONTEXT.md has no heading starting {heading_prefix!r}")
    i = h
    while i > 0 and (lines[i - 1].strip() == "" or lines[i - 1].startswith("-----")):
        i -= 1
    return i


def stamp(root, reason=""):
    path = context_path(root)
    text = read(path)
    today = dt.date.today().strftime("%d %b %Y").lstrip("0")
    kept = re.search(r"^Last updated: [^(]*\((.*)\)\.?$", text, re.M)               # no reason given: keep the earlier one
    reason = reason or (kept.group(1) if kept else "")
    new = f"Last updated: {today}" + (f" ({reason})" if reason else "") + "."
    text2, n = re.subn(r"^Last updated:.*$", lambda m: new, text, count=1, flags=re.M)
    if n:
        write(path, text2)
    return n > 0


def add_note(root, section, text, title="", date=None, reason=""):
    refuse_cell_references("note", text, title, reason)
    path = context_path(root)
    raw = read(path)
    nl = eol(raw)
    lines = raw.split(nl)
    text = re.sub(r"\s+", " ", text).strip()
    if section in ("pattern", "failure"):
        letter = "P" if section == "pattern" else "F"
        nums = [int(m.group(1)) for ln in lines if (m := re.match(rf"^{letter}-(\d+)\b", ln))]
        n = max(nums or [0]) + 1
        body = f"{letter}-{n}" + ("  " if section == "failure" else " ") + (f"**{title.rstrip('.')}.** " if title else "") + text
        at = _insert_before_rule(lines, SECTION_TITLES[section])
        lines[at:at] = [body]
        label = f"{letter}-{n}"
    elif section == "decision":
        row = f"| {date or dt.date.today().strftime('%d %b %Y').lstrip('0')} | {text} | {title or 'all'} |"
        start = next((i for i, ln in enumerate(lines) if ln.startswith("| date | decision")), None)
        if start is None:
            raise SystemExit("AGENT_CONTEXT.md has no decision table")
        end = start + 1
        while end < len(lines) and lines[end].startswith("|"):
            end += 1
        lines[end:end] = [row]
        label = "decision row"
    elif section == "open":
        at = _insert_before_rule(lines, SECTION_TITLES[section])
        lines[at:at] = [f"- {text}"]
        label = "open item"
    else:
        raise SystemExit("--section must be pattern, failure, decision or open")
    write(path, nl.join(lines))
    stamp(root, reason)
    return label


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=DEFAULT_ROOT, help="the grid_tools folder (default: the one this skill lives in)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("lesson")
    p.add_argument("--insurer", required=True)
    p.add_argument("--title", required=True)
    p.add_argument("--cause", required=True)
    p.add_argument("--fix", required=True)
    p.add_argument("--covered", required=True)
    p.add_argument("--rule", default="")
    p.add_argument("--when", default=None, help="YYYY-MM (default: this month)")
    p = sub.add_parser("case")
    p.add_argument("--name", required=True)
    p.add_argument("--script", required=True)
    p.add_argument("--from-date", required=True)
    p.add_argument("--to-date", required=True)
    p.add_argument("--input", action="append", default=[], help="flag=path, repeatable (grid=..., satp=..., ref=...)")
    p.add_argument("--answers", nargs="*", default=[])
    p.add_argument("--note", default="")
    p.add_argument("--bless", action="store_true")
    p = sub.add_parser("note")
    p.add_argument("--section", required=True, choices=["pattern", "failure", "decision", "open"])
    p.add_argument("--text", required=True)
    p.add_argument("--title", default="", help="pattern/failure: the bold lead-in; decision: the scope column")
    p.add_argument("--date", default=None)
    p.add_argument("--reason", default="")
    p = sub.add_parser("skill")
    p.add_argument("--insurer", required=True, help="liberty, magma, rsa, united, shriram, godigit, sbi, icici")
    p.add_argument("--section", required=True, choices=sorted(SKILL_SECTIONS))
    p.add_argument("--text", required=True)
    p.add_argument("--date", default=None)
    p = sub.add_parser("stamp")
    p.add_argument("--reason", default="")
    a = ap.parse_args(argv)
    root = os.path.abspath(a.root)
    if a.cmd == "lesson":
        print("lesson added to", add_lesson(root, a.insurer, a.title, a.cause, a.fix, a.covered, a.rule, a.when))
    elif a.cmd == "case":
        add_case(root, a.name, a.script, a.from_date, a.to_date, a.input, a.answers, a.note, a.bless)
    elif a.cmd == "note":
        print("added", add_note(root, a.section, a.text, a.title, a.date, a.reason), "to AGENT_CONTEXT.md")
    elif a.cmd == "skill":
        print("added to", add_skill_note(root, a.insurer, a.section, a.text, a.date))
    elif a.cmd == "stamp":
        print("stamped" if stamp(root, a.reason) else "no 'Last updated' line found")


if __name__ == "__main__":
    main()
