"""A note the person resolved (explained, or said needs no action) is never flagged again - in any of the eight converters, whatever number it carries (the user, 31 Oct 2026).
`fold_note` / `resolved_keys` / `ManualReview.flag` of the shared block, against a GRID_STANDING_NOTES_FILE like the web app writes it.
Run:  .venv\\Scripts\\python.exe tests\\resolved_notes_tests.py"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


SCRIPTS = ["liberty_grid_to_config.py", "magma_grid_to_config.py", "rsa_gcv4w_comp_to_config.py", "rsa_pvtcar_to_config.py", "united_commission_to_config.py",
           "shriram_broker_grid_to_config.py", "shriram_2w_to_config.py", "godigit_agency_to_config.py", "sbi_grid_to_config.py", "icici_grid_to_config.py", "indusind_grid_to_config.py"]
tmp = tempfile.mkdtemp(prefix="resolved-")
path = os.path.join(tmp, "standing.json")


def write(items):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(items, fh)


M = helpers.load_script("liberty_grid_to_config.py")
same = [("For Hazardous category PO will be 5% less", "For Hazardous category PO will be 7% less"),
        ("discount limit 50.0072% in the grid written as 50%", "discount limit 49.7903% in the grid written as 49%"),
        ("grid note: Few RTOs for 1000-1500CC segment are declined", "grid note: Few RTOs for 1500-2000CC segment are declined"),
        ("TP PO increased by 12%  & UW's guidelines revised", "tp po increased by 9% & uw's guidelines revised.")]
different = [("swift dzire model declined for RJ-14", "swift dzire model declined for MH-12"), ("EXL RTO codes 03, 06, 18, 24, 26", "EXL RTO codes 10, 11, 12, 13"),
             ("For Hazardous category PO will be 5% less", "For Hazardous category PO will be 5% more"), ("Petrol,CNG,LPG Other than TR RTO codes only", "Petrol,CNG,LPG Other than TS RTO codes only")]
for a, b in same:
    check(f"fold_key: '{a[:50]}' and the same note with other numbers are one note", M.fold_key(a) == M.fold_key(b))
for a, b in different:
    check(f"fold_key: '{a[:50]}' and '{b[-22:]}' are different notes (an RTO code, a list of codes, or a different wording)", M.fold_key(a) != M.fold_key(b))


class Rows:
    pass


for s in SCRIPTS:
    m = helpers.load_script(s)
    os.environ.pop("GRID_STANDING_NOTES_FILE", None)
    rv = m.ManualReview()
    rv.flag("For Hazardous category PO will be 5% less", None)
    check(f"{s}: without the web app's list the note is flagged", len(rv.items) == 1)
    write(["For Hazardous category PO will be 7% less", "discount limit 50.0072% in the grid written as 50%"])
    os.environ["GRID_STANDING_NOTES_FILE"] = path
    rv = m.ManualReview()
    rv.flag("For Hazardous category PO will be 5% less", None)
    rv.flag("discount limit 49.7903% in the grid written as 49%", None)
    rv.flag("a note nobody explained, with a number 12", None)
    rv.flag("For Hazardous category PO will be 5% more", None)
    check(f"{s}: a resolved note is not flagged (also with another number); a new note and a reworded one are", [t for t, _ in rv.items] == ["a note nobody explained, with a number 12", "For Hazardous category PO will be 5% more"], rv.items)
    check(f"{s}: what was kept from the flags is kept for the report (the page shows it as 'from memory')", len(m.NOTES_SILENCED) == 2, m.NOTES_SILENCED)
    write(["a note nobody explained, with a number 99"])                      # the file changes between two runs of one process: it is read again
    rv = m.ManualReview()
    rv.flag("a note nobody explained, with a number 12", None)
    check(f"{s}: a changed list is read again (a long-lived process, the tests)", rv.items == [])
    write(["For vehicles registered after the cut-off date the payout is reduced by 2 points."])
    long_same, long_other = "For vehicles registered after the cut-off date the payout is reduced by 3 points.", "For vehicles registered after the cut-off date the payout is increased by 3 points."
    check(f"{s}: unknown_notes (a grid note read from a cell) is silenced by a resolved note with another number, not by a reworded one", m.unknown_notes([long_same], []) == [] and m.unknown_notes([long_other], []) == [long_other])
os.environ.pop("GRID_STANDING_NOTES_FILE", None)

bad = [r for r in results if not r[1]]
print(f"\n{len(results) - len(bad)} passed, {len(bad)} failed")
sys.exit(1 if bad else 0)
