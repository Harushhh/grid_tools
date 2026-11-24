"""Run the test suites of this folder and print one line per suite.

    .venv\\Scripts\\python.exe tests\\run_tests.py            # RSA + manual-review suites (about 3 minutes)
    .venv\\Scripts\\python.exe tests\\run_tests.py --all      # also the Liberty / Magma / environment suites (about 15 minutes)
    .venv\\Scripts\\python.exe tests\\run_tests.py --fuzz     # also the random-corruption fuzzers (about 15 minutes)
    .venv\\Scripts\\python.exe tests\\run_tests.py --shift    # also the structure-shift suite on EVERY script x every operator (about an hour); --all runs its quick subset

A suite prints PASS / FAIL per check; it passes when no check says FAIL, nothing crashed and at least one check ran.
check_regressions.py (the known-good grids must give identical output) is separate and quicker - run both before you finish a change."""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
QUICK = ["cluster_names_tests.py", "resolved_notes_tests.py", "layout_shift_tests.py", "master_pipeline_tests.py", "rsa_tests.py", "rsa_tests_notes.py", "pvtcar_tests.py", "united_tests.py", "united_aug_tests.py", "united_oct_tests.py", "united_apr_tests.py", "shriram_tests.py", "shriram_2w_tests.py", "godigit_tests.py", "godigit_cd1_tests.py", "godigit_wording_tests.py", "godigit_d_tests.py", "sbi_tests.py", "icici_tests.py", "indusind_tests.py", "working_list_tests.py", "lowest_rate_tests.py", "stp_rule_tests.py", "skills_tests.py", "insurer_docs_tests.py", "review_tests.py", "service_tests.py", "united_shift_tests.py", "twin_column_tests.py", "structure_helpers_tests.py"]
SLOW = ["magma_interactive_tests.py", "liberty_tests.py", "liberty_interactive_tests.py", "env_tests.py"]
SHIFT = ["structure_shift_tests.py"]        # the structure-shift suite (11 Nov 2026): --shift = every script x every operator; --all = the small grids only (SHIFT_QUICK=1)
FUZZ = [("fuzz_rsa.py", ["300", "1", "5"]), ("fuzz_pvtcar.py", ["200", "1", "4"]), ("fuzz_united.py", ["200", "1", "4"]), ("fuzz_united.py", ["200", "1", "4", "apr"]),
        ("fuzz_run.py", ["lib", "120", "1", "3"]),("fuzz_run.py", ["mag", "100", "1", "3"])]


CLUSTER_SUITES = {"cluster_names_tests.py"}       # the suites of the RTO name rule of 30 Oct 2026 (new_rto_list = the bare cluster name); every other suite tests the group-name mode


def run(script, args=()):
    env = dict(os.environ, GRID_RTO_NAMES="cluster" if script in CLUSTER_SUITES else "group")
    p = subprocess.run([sys.executable, "-I", "-u", os.path.join(HERE, script), *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", cwd=os.path.dirname(HERE), timeout=3600, env=env)
    return p.returncode, p.stdout + p.stderr


def tidy():
    """The older suites write their scratch workbooks next to themselves; remove them (the .py files and the .pkl bases stay)."""
    import shutil
    for name in os.listdir(HERE):
        path = os.path.join(HERE, name)
        if os.path.isdir(path) and name != "__pycache__":
            shutil.rmtree(path, ignore_errors=True)
        elif name.lower().endswith((".xlsx", ".xlsb", ".csv", ".json", ".txt")):
            os.remove(path)


def main():
    failed = 0
    if not all(os.path.exists(os.path.join(HERE, f)) for f in ("fuzz_lib_bases.pkl", "fuzz_mag_bases.pkl", "fuzz_rsa_bases.pkl")):
        run("make_bases.py")
    suites = QUICK + (SLOW if "--all" in sys.argv else []) + (SHIFT if ("--all" in sys.argv or "--shift" in sys.argv) else [])
    if "--all" in sys.argv and "--shift" not in sys.argv:
        os.environ.setdefault("SHIFT_QUICK", "1")
    for s in suites:
        code, out = run(s)
        lines = out.splitlines()
        passed = sum(1 for ln in lines if ln.startswith(("PASS", "ok ")))
        bad = [ln for ln in lines if ln.startswith(("FAIL ", "NOT OK"))]
        crashed = "Traceback (most recent call last)" in out
        good = passed > 0 and not bad and not crashed and (code == 0 or not re.search(r"(\d+) passed, (\d+) failed", out))
        print(f"{'ok  ' if good else 'FAIL'}  {s:<30} {passed} passed, {len(bad)} failed" + (", CRASHED" if crashed else ""))
        for ln in (bad + [ln for ln in lines if ln.startswith(("AssertionError", "TypeError", "KeyError", "NameError"))])[:6]:
            print("        " + ln[:200])
        failed += not good
    if "--fuzz" in sys.argv:
        for script, args in FUZZ:
            code, out = run(script, args)
            m = re.search(r"FAILURES: (\d+) of (\d+)", out)
            print(f"{'ok  ' if m and m.group(1) == '0' else 'FAIL'}  {script + ' ' + ' '.join(args[:2]):<30} {m.group(0) if m else 'did not finish'}")
            failed += not (m and m.group(1) == "0")
    tidy()
    print("\nALL OK" if not failed else f"\n{failed} suite(s) need attention")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
