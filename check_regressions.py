"""Re-run the known-good grids and prove the scripts still give exactly the same output.

    python check_regressions.py              # all cases
    python check_regressions.py --only magma_oct26
    python check_regressions.py --bless      # after an INTENDED change: record the new outputs as the baseline

A case passes when the script finishes, asks the expected number of questions, matches the earlier upload CSV where one
is given (new / removed / changed rows), and writes a CSV byte-identical to the recorded baseline.
Exit code 0 = all passed, 1 = something differs.
"""
import argparse
import glob
import hashlib
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CASES = os.path.join(HERE, "regression", "cases.json")
BASELINE = os.path.join(HERE, "regression", "baseline.json")


def shared_block(script):
    """The question layer every script must carry identically (between the QUESTIONS and READING headings)."""
    text = open(os.path.join(HERE, script), encoding="utf-8").read()
    m = re.search(r"^# %% -+ QUESTIONS\n(.*?)^# %% -+ READING THE ", text, re.S | re.M)
    return m.group(1) if m else None


def values_hash(path):
    """sha256 of the CSV with the remarks column blanked: tells 'only the manual-review notes changed' from 'the rates moved'."""
    import io
    import pandas as pd
    frame = pd.read_csv(path, dtype=str, keep_default_na=False)
    if "remarks" in frame.columns:
        frame["remarks"] = ""
    buf = io.StringIO()
    frame.to_csv(buf, index=False, lineterminator="\n")
    return hashlib.sha256(buf.getvalue().encode("utf-8")).hexdigest()


def stp_violations(path):
    """Rows of an upload file whose sub_product is STP but whose is_ncb or is_zd is not NA.  The user's rule (27 Oct 2026, repeated 4 Nov 2026: "there should be always NA in NCB and ZD if it is STP"):
    STP has no own-damage cover, so both columns are NA on EVERY STP row, in every script.  A file without those columns (a side file) has none."""
    import pandas as pd
    frame = pd.read_csv(path, dtype=str, keep_default_na=False)
    if not {"sub_product", "is_ncb", "is_zd"} <= set(frame.columns):
        return 0
    stp = frame[frame["sub_product"].str.upper() == "STP"]
    return int(((stp["is_ncb"] != "NA") | (stp["is_zd"] != "NA")).sum())


def run_case(case):
    out = os.path.join("regression", "out", case["name"] + ".csv")
    if os.path.exists(os.path.join(HERE, out)):
        os.remove(os.path.join(HERE, out))
    cmd = [sys.executable, "-I", case["script"], *case["args"], "--out", out]
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run(cmd, cwd=HERE, env=env, input="\n".join(case["answers"]) + "\n",
                       capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
    text = p.stdout + "\n" + p.stderr
    got = {"exit": p.returncode, "text": text}
    m = re.search(r"Rows to write: ([\d,]+)", text)
    got["rows"] = int(m.group(1).replace(",", "")) if m else None
    m = re.search(r"(\d+) new rows, (\d+) removed rows, (\d+) rows with changed values", text)
    if m:
        got["new"], got["removed"], got["changed"] = map(int, m.groups())
    m = re.search(r"MANUAL REVIEW - ([\d,]+) row", text)
    got["flagged"] = int(m.group(1).replace(",", "")) if m else 0
    got["questions"] = sum(1 for ln in text.splitlines() if ln.startswith("? ") and "Earlier upload CSV" not in ln)
    full = os.path.join(HERE, out)
    got["sha256"] = hashlib.sha256(open(full, "rb").read()).hexdigest() if os.path.exists(full) else None
    got["values"] = values_hash(full) if os.path.exists(full) else None
    got["stp_bad"] = stp_violations(full) if os.path.exists(full) else 0
    return got


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--only", help="run just this case name")
    ap.add_argument("--bless", action="store_true", help="record the current outputs as the new baseline")
    a = ap.parse_args()
    cases = json.load(open(CASES, encoding="utf-8"))["cases"]
    if a.only:
        cases = [c for c in cases if c["name"] == a.only]
        if not cases:
            sys.exit(f"no case named {a.only}")
    baseline = json.load(open(BASELINE)) if os.path.exists(BASELINE) else {}
    failed = 0
    scripts = sorted(os.path.basename(p) for p in glob.glob(os.path.join(HERE, "*_to_config.py")))
    blocks = {s: shared_block(s) for s in scripts}
    if None in blocks.values() or len(set(blocks.values())) != 1:
        failed += 1
        print("FAIL  shared question block differs between the scripts " + str([s for s, b in blocks.items() if b != next(iter(blocks.values()))] or scripts)
              + " - make the same edit in all of them (the text between '# %% --- QUESTIONS' and '# %% --- READING THE ...')")
    else:
        print(f"ok    shared question block identical in all {len(scripts)} scripts")
    # every script must carry the manual-review safeguard: text rules it cannot apply go to the remarks column, never a guess or a stop
    for name in scripts:
        text = open(os.path.join(HERE, name), encoding="utf-8").read()
        if not re.search(r"\breview\b[^\n=]*=[^\n]*ManualReview\(\)", text) or "review.report(df" not in text:
            failed += 1
            print(f"FAIL  {name} has no manual-review safeguard: it must create `review = ManualReview()`, flag every note or text rule it "
                  f"cannot apply with review.flag(text, rows), and call review.report(df, log) before returning the frame")
        else:
            print(f"ok    {name} uses the manual-review safeguard (remarks column)")
    # every script must be startable from VS Code, where the person runs it: a launch entry (Run and Debug, F5) and a task (Run Task).
    # Go Digit and both Shriram scripts were once built without them and were invisible in the Run and Debug list.
    vscode = {}
    for fname in ("launch.json", "tasks.json"):
        try:
            vscode[fname] = open(os.path.join(HERE, ".vscode", fname), encoding="utf-8").read()
        except OSError:
            vscode = {}
            break
    if vscode:
        for name in scripts:
            missing = [fname for fname, text in vscode.items() if name not in text]
            if missing:
                failed += 1
                print(f"FAIL  {name} cannot be started from VS Code: it is not in .vscode/{' or .vscode/'.join(missing)} "
                      f"(add an entry named like the others: 'Run and Debug' = launch.json, 'Run Task' = tasks.json)")
            else:
                print(f"ok    {name} can be started from VS Code (launch entry and task)")
    # every script has its insurer's skill, .claude/skills/<insurer>-grid/SKILL.md (the insurer = the first word of the script's file name): the knowledge the agent
    # loads for that insurance company (the user, 4 Nov 2026). tests/skills_tests.py checks what the skills say.
    skills_dir = os.path.join(HERE, ".claude", "skills")
    if os.path.isdir(skills_dir):
        for insurer in sorted({s.split("_")[0] for s in scripts}):
            if os.path.isfile(os.path.join(skills_dir, f"{insurer}-grid", "SKILL.md")):
                print(f"ok    {insurer}-grid: the insurer's skill exists")
            else:
                failed += 1
                print(f"FAIL  no insurer skill for {insurer}: create .claude/skills/{insurer}-grid/SKILL.md (copy the nearest insurer's skill; grid-converter Job C) - the agent loads it with grid-converter")
    # every insurer also has its own dedicated reference, <INSURER>_GRID.md in this folder (the user, 23 Nov 2026), and no document or skill names a place in a grid by a cell
    # address, a column range, a row number or a column letter: documentation is rule-based (.claude/skills/grid-converter/scripts/cellref_lint.py; tests/insurer_docs_tests.py checks what the files say)
    for insurer in sorted({s.split("_")[0] for s in scripts}):
        md = f"{insurer.upper()}_GRID.md"
        if os.path.isfile(os.path.join(HERE, md)):
            print(f"ok    {md}: the insurer's dedicated reference exists")
        else:
            failed += 1
            print(f"FAIL  no dedicated reference for {insurer}: create {md} (copy the nearest insurer's file; grid-converter Job C) - the whole grid and script of that insurer, rule-based, next to its skill")
    lint_path = os.path.join(HERE, ".claude", "skills", "grid-converter", "scripts", "cellref_lint.py")
    if os.path.isfile(lint_path):
        import importlib.util
        spec = importlib.util.spec_from_file_location("cellref_lint", lint_path)
        lint = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(lint)
        docs = lint.doc_files(HERE)
        bad = []
        for doc in docs:
            hits = lint.find_in_file(doc)
            if hits:
                bad.append((doc, hits))
        if bad:
            failed += 1
            for doc, hits in bad:
                print(f"FAIL  {os.path.relpath(doc, HERE)} names a place in a grid by its position: {lint.describe(hits, 4)}")
            print("      " + lint.ADVICE)
        else:
            print(f"ok    no document or skill names a cell address, column range, row number or column letter ({len(docs)} files)")
    for case in cases:
        name = case["name"]
        got = run_case(case)
        problems = []
        if got["exit"] != 0 or got["sha256"] is None:
            tail = [ln for ln in got["text"].splitlines() if ln.strip()][-3:]
            problems.append("did not finish: " + " | ".join(tail))
        for k, want in case["expect"].items():
            if got.get(k) != want:
                problems.append(f"{k}: expected {want}, got {got.get(k)}")
        if got.get("stp_bad"):
            problems.append(f"{got['stp_bad']} STP row(s) carry is_ncb or is_zd other than NA - STP rows must ALWAYS have NA in both (the user's rule)")
        if a.bless:
            if problems:
                print(f"NOT BLESSED {name}: " + "; ".join(problems))
                failed += 1
            else:
                baseline[name] = {"sha256": got["sha256"], "values": got["values"]}
                print(f"blessed     {name}  ({got['rows']:,} rows)")
            continue
        if name not in baseline:
            problems.append("no baseline yet - run with --bless once the output is confirmed good")
        else:
            base = baseline[name] if isinstance(baseline[name], dict) else {"sha256": baseline[name], "values": baseline[name]}
            if got["sha256"] != base["sha256"]:
                if got["values"] == base["values"]:
                    problems.append("ONLY the remarks column differs from the baseline (rates and every other column are identical): "
                                    "the manual-review flags changed - check they are intended, then --bless")
                else:
                    problems.append("output differs from the recorded baseline (the rates or other columns moved)")
        if problems:
            failed += 1
            print(f"FAIL  {name}")
            for pr in problems:
                print("        " + pr)
        else:
            flagged = f", {got['flagged']:,} rows flagged for manual review" if got.get("flagged") else ""
            print(f"ok    {name}  ({got['rows']:,} rows, identical to baseline{flagged})")
    if a.bless:
        json.dump(baseline, open(BASELINE, "w"), indent=1)
    print("\nALL OK" if not failed else f"\n{failed} case(s) need attention")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
