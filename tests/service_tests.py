"""The service mode of the shared question layer (the web app drives the scripts through environment variables): the same checks against all eight scripts.
Without the variables nothing changes; with GRID_SERVICE=1 a question is written to a file instead of being asked, an answer comes from a file, a name mapping and extra
standing notes can be given as files, and the run's report is written as JSON.
Run:  .venv\\Scripts\\python.exe tests\\service_tests.py"""
import json
import os
import shutil
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

SCRIPTS = ["liberty_grid_to_config.py", "magma_grid_to_config.py", "rsa_gcv4w_comp_to_config.py", "rsa_pvtcar_to_config.py", "united_commission_to_config.py",
           "shriram_broker_grid_to_config.py", "shriram_2w_to_config.py", "godigit_agency_to_config.py", "sbi_grid_to_config.py", "icici_grid_to_config.py", "indusind_grid_to_config.py"]
VARS = ("GRID_SERVICE", "GRID_QUESTIONS_FILE", "GRID_ANSWERS_FILE", "GRID_REPORT_FILE", "GRID_ALIASES_FILE", "GRID_STANDING_NOTES_FILE")
results = []
quiet = lambda *a, **k: None
tmp = tempfile.mkdtemp(prefix="service_")


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def env(**kw):
    for v in VARS:
        os.environ.pop(v, None)
    for k, v in kw.items():
        os.environ[k] = v


def write(name, data):
    p = os.path.join(tmp, name)
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(data, fh)
    return p


def frame(rows):
    return pd.DataFrame([dict(k=k, product="P", sub_product="1+1", new_rto_list="GJ", new_vehicle_makes="X", fuel_type="Petrol", make_model_class="NA", tariff_max=100.01,
                              pi_type=t, pi_od_rate=0, pi_tp_rate=0, pi_net_rate=net, pi_flat_amount=0, from_date="01/10/2026", to_date="31/03/2030", remarks="") for k, t, net in rows])


for file in SCRIPTS:
    m = helpers.load_script(file)
    tag = file.replace("_to_config.py", "")
    OPTS = [("Use the first sheet", "first"), ("Use the second sheet", "second"), ("Stop", None)]

    # ---- without the variables nothing changes
    env()
    try:
        m.Asker(interactive=False).choose("sheet is unclear", OPTS, question="Which sheet?")
        ok = False
    except m.NeedsAnswer as e:
        ok = str(e) == "sheet is unclear"
    check(f"{tag}: without GRID_SERVICE a question that cannot be asked stops the run, as before", ok)
    check(f"{tag}: service_on() is off and apply_aliases / unknown_notes change nothing", not m.service_on() and m.unknown_notes(["A long enough rule text that nobody has seen before."], []) != [])

    # ---- a question is written to the questions file and the run stops
    qf = os.path.join(tmp, f"q_{tag}.json")
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf)
    ask = m.Asker(interactive=False)
    try:
        ask.choose("sheet is unclear", OPTS, question="Which sheet?", short="sheet")
        ok = False
    except m.NeedsAnswer:
        ok = True
    q = json.load(open(qf, encoding="utf-8"))["questions"]
    check(f"{tag}: with GRID_SERVICE=1 the question is written to the questions file (id, fingerprint, options) and the run stops",
          ok and len(q) == 1 and q[0]["status"] == "pending" and [o["label"] for o in q[0]["options"]] == [l for l, _ in OPTS] and len(q[0]["id"]) == 16 and len(q[0]["fp"]) == 16, q)
    qid = q[0]["id"]
    try:
        ask.choose("sheet is unclear", OPTS, question="Which sheet?", short="sheet")
    except m.NeedsAnswer:
        pass
    check(f"{tag}: the same question asked twice is one entry", len(json.load(open(qf, encoding="utf-8"))["questions"]) == 1)

    # ---- its answer comes from the answers file
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf, GRID_ANSWERS_FILE=write(f"a_{tag}.json", {"answers": [{"id": qid, "pick": [2]}]}))
    ask = m.Asker(interactive=False)
    v = ask.choose("sheet is unclear", OPTS, question="Which sheet?", short="sheet")
    q = json.load(open(qf, encoding="utf-8"))["questions"]
    check(f"{tag}: the answer 'pick 2' returns the second option's value, is marked answered and noted among the decisions", v == "second" and q[0]["status"] == "answered"
          and any("Use the second sheet" in d for d in ask.decisions), (v, q, ask.decisions))
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf, GRID_ANSWERS_FILE=write(f"a2_{tag}.json", {"answers": [{"id": qid, "pick": [9]}]}))
    try:
        m.Asker(interactive=False).choose("sheet is unclear", OPTS, question="Which sheet?", short="sheet")
        ok = False
    except m.NeedsAnswer:
        ok = json.load(open(qf, encoding="utf-8"))["questions"][0]["invalid"] is True
    check(f"{tag}: an answer that does not fit the options is asked again (marked invalid)", ok)

    # ---- several picks and a check
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf)
    try:
        m.Asker(interactive=False).choose("pick several", [("A", 1), ("B", 2), ("C", 3)], question="Which?", multi=True)
    except m.NeedsAnswer:
        pass
    mid = [x for x in json.load(open(qf, encoding="utf-8"))["questions"] if x["problem"] == "pick several"][0]["id"]
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf, GRID_ANSWERS_FILE=write(f"m_{tag}.json", {"answers": [{"id": mid, "pick": [1, 3]}]}))
    check(f"{tag}: multi-choice answers return the list of values", m.Asker(interactive=False).choose("pick several", [("A", 1), ("B", 2), ("C", 3)], question="Which?", multi=True) == [1, 3])

    def refuse(values):
        raise ValueError("no")
    try:
        m.Asker(interactive=False).choose("pick several", [("A", 1), ("B", 2), ("C", 3)], question="Which?", multi=True, check=refuse)
        ok = False
    except m.NeedsAnswer:
        ok = True
    check(f"{tag}: a multi-choice answer the script's own check refuses is asked again", ok)

    # ---- free text
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf)
    try:
        m.Asker(interactive=False).text("folder missing", question="Which folder?")
    except m.NeedsAnswer:
        pass
    tid = [x for x in json.load(open(qf, encoding="utf-8"))["questions"] if x["problem"] == "folder missing"][0]["id"]
    env(GRID_SERVICE="1", GRID_QUESTIONS_FILE=qf, GRID_ANSWERS_FILE=write(f"t_{tag}.json", {"answers": [{"id": tid, "text": " D:/x "}]}))
    check(f"{tag}: a text answer is returned (trimmed) and goes through the script's validator",
          m.Asker(interactive=False).text("folder missing", question="Which folder?") == "D:/x"
          and m.Asker(interactive=False).text("folder missing", question="Which folder?", validate=lambda a: a.upper()) == "D:/X")

    def bad(a):
        raise ValueError("not a folder")
    try:
        m.Asker(interactive=False).text("folder missing", question="Which folder?", validate=bad)
        ok = False
    except m.NeedsAnswer:
        ok = True
    check(f"{tag}: a text answer the validator refuses is asked again", ok)
    env(GRID_SERVICE="1")
    check(f"{tag}: confirm() still returns its default (the web app reviews the result afterwards)", m.Asker(interactive=False).confirm("Write the file?", True) is True)

    # ---- the fingerprint ignores what changes every month
    a1 = m.question_fp("choose", "cannot open the workbook Jul26_Grid_v2.xlsx: BadZipFile 12", ["Retry", "Use another file"])
    a2 = m.question_fp("choose", "cannot open the workbook Aug26_Grid_v3.xlsx: BadZipFile 34", ["Retry", "Use another file"])
    a3 = m.question_fp("choose", "cannot open the workbook Aug26_Grid_v3.xlsx: OSError", ["Retry", "Use another file"])
    check(f"{tag}: the fingerprint ignores digits and file names, but not the wording", a1 == a2 and a1 != a3, (a1, a2, a3))
    check(f"{tag}: the id is exact (it changes with the wording)", m.question_id("choose", "a 1", ["x"]) != m.question_id("choose", "a 2", ["x"]))

    # ---- name mappings and standing notes
    env(GRID_ALIASES_FILE=write(f"al_{tag}.json", {"rto": {"NO_SUCH_GROUP_X": "ALLINDIA", "no_such_other": "GJ,ALLINDIA"}, "makes": {"NO_SUCH_MAKE": "HERO"}}))
    m.reset_run_lists()
    masters = m.load_masters(m.Asker(interactive=False))
    df = pd.DataFrame([dict(new_rto_list="NO_SUCH_GROUP_X", new_vehicle_makes="NO_SUCH_MAKE"), dict(new_rto_list="GJ,No_Such_Other", new_vehicle_makes="HERO,NO_SUCH_MAKE"),
                       dict(new_rto_list="STILL_UNKNOWN", new_vehicle_makes="HERO")])
    rv = m.ManualReview()
    m.check_names(df, rv, masters, quiet)
    check(f"{tag}: the person's name mapping is applied before the master check (whole value and name by name, capitals ignored)",
          list(df.new_rto_list) == ["ALLINDIA", "GJ,GJ,ALLINDIA", "STILL_UNKNOWN"] and list(df.new_vehicle_makes) == ["HERO", "HERO,HERO", "HERO"], df.to_dict("records"))
    check(f"{tag}: ... and a name that is still unknown is still flagged and listed; the mappings used are kept for the report",
          list(m.EXTRA_SHEETS["RTO groups to create"].group_needed) == ["STILL_UNKNOWN"] and {a["name"] for a in m.ALIAS_USED} >= {"NO_SUCH_GROUP_X", "NO_SUCH_MAKE"}, (m.EXTRA_SHEETS, m.ALIAS_USED))
    env()
    note = "For vehicles registered after the cut-off date the payout is reduced by two points."
    check(f"{tag}: a long note is unknown ...", m.unknown_notes([note], []) == [note])
    env(GRID_STANDING_NOTES_FILE=write(f"sn_{tag}.json", [note.upper().replace(".", " ")]))
    check(f"{tag}: ... and a standing note given as a file (capitals, spaces and punctuation ignored) needs no action; a reworded one still does",
          m.unknown_notes([note], []) == [] and m.unknown_notes([note.replace("two", "three")], []) == [note.replace("two", "three")])
    longer = note[:-1] + ", and by five points for diesel vehicles."
    check(f"{tag}: ... and an approved note never silences a longer rule that merely contains it", m.unknown_notes([longer], []) == [longer], m.unknown_notes([longer], []))
    m.reset_run_lists()
    m.unknown_notes([note, note.lower(), longer], [])
    check(f"{tag}: ... and the notes the approved list silenced are kept for the report (once each)", list(m.NOTES_SILENCED.values()) == [note], m.NOTES_SILENCED)

    # ---- the report
    env(GRID_REPORT_FILE=os.path.join(tmp, f"r_{tag}.json"))
    m.reset_run_lists()
    m.resolve_duplicates(frame([("a", "On Net", 30), ("a", "On Net", 25)]), ["k"], m.Asker(interactive=False))
    m.RUN_WARNINGS.append("a warning of this run")
    out = os.path.join(tmp, f"o_{tag}", f"{tag}_UPLOAD_2026-10.csv")
    m.write_csv(frame([("a", "On Net", 25)]), out, m.Asker(interactive=False), quiet)
    rep = json.load(open(os.path.join(tmp, f"r_{tag}.json"), encoding="utf-8"))
    check(f"{tag}: the report holds the counts, warnings, dates and the sheets of the review workbook", rep["rows"] == 1 and rep["warnings"] == ["a warning of this run"] and rep["from_date"] == "01/10/2026" and rep["notes_silenced"] == []
          and "Lowest rate chosen" in rep["sheets"] and rep["sheets"]["Lowest rate chosen"][0]["written"] == "25" and rep["csv"].endswith(".csv") and rep["review"].endswith(".xlsx"), rep)

    # ---- the report says which flagged notes are grid text (a "no action needed" entry can silence those) and which are the script's own checks
    env(GRID_REPORT_FILE=os.path.join(tmp, f"n_{tag}.json"))
    m.reset_run_lists()
    long_note = "For vehicles registered after the cut-off date the payout is reduced by two points, as per the insurer's circular dated the first of the month."
    m.unknown_notes([long_note], [])
    m.LAST_NOTES = [(long_note[:70] + "...", 10), (long_note, 4), ("the script's own check - not a text of the grid", 3)]
    m.service_report(frame([("a", "On Net", 25)]), os.path.join(tmp, f"o_{tag}", f"{tag}_UPLOAD_2026-10.csv"), {}, m.Asker(interactive=False))
    rep = json.load(open(os.path.join(tmp, f"n_{tag}.json"), encoding="utf-8"))
    check(f"{tag}: the report gives the full grid text behind a flagged note (also when the flag is cut short) and nothing for the script's own checks",
          [(n["rows_flagged"], n["grid_note"]) for n in rep["notes"]] == [(10, long_note), (4, long_note), (3, "")], rep["notes"])
    m.LAST_NOTES = []
    env()

# the closest master names are offered for a name the master lacks (only when a report is asked for)
m = helpers.load_script("godigit_agency_to_config.py")
env(GRID_REPORT_FILE=os.path.join(tmp, "s.json"))
m.reset_run_lists()
masters = m.load_masters(m.Asker(interactive=False))
df = pd.DataFrame([dict(new_rto_list="GO-DIGIT_OCT26_HCV_Bad_Gj", new_vehicle_makes="HERO"), dict(new_rto_list="NOT_A_THING_AT_ALL", new_vehicle_makes="HERO")])
m.check_names(df, m.ManualReview(), masters, quiet)
sheet = m.EXTRA_SHEETS.get("RTO groups to create")
check("the report offers the closest master names for an unknown group", sheet is not None and "suggestions" in sheet.columns and "GO-DIGIT_OCT26_HCV_BAD_GJ" in " ".join(sheet.suggestions) if sheet is not None else False, sheet)
env()
m.reset_run_lists()
df = pd.DataFrame([dict(new_rto_list="NOT_A_THING_AT_ALL", new_vehicle_makes="HERO")])
m.check_names(df, m.ManualReview(), masters, quiet)
check("without GRID_REPORT_FILE the review workbook's columns are exactly what they were (no suggestions column)", list(m.EXTRA_SHEETS["RTO groups to create"].columns) == ["group_needed", "rows_flagged", "items"])

shutil.rmtree(tmp, ignore_errors=True)
passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
