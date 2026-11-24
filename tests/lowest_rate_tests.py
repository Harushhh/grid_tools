"""The shared duplicate rule (the user, 27 Oct 2026: "if the rate grid contains two different rates for the same entry, always select the lower rate"): the same checks against the
shared block of all seven scripts - nothing is asked, identical payins are one row, different payins keep the lowest, the choices are listed next to the CSV.
Run:  .venv\\Scripts\\python.exe tests\\lowest_rate_tests.py"""
import os
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

SCRIPTS = ["liberty_grid_to_config.py", "magma_grid_to_config.py", "rsa_gcv4w_comp_to_config.py", "rsa_pvtcar_to_config.py", "united_commission_to_config.py",
           "shriram_broker_grid_to_config.py", "shriram_2w_to_config.py", "godigit_agency_to_config.py", "sbi_grid_to_config.py", "icici_grid_to_config.py", "indusind_grid_to_config.py"]
FOLDERS = {"liberty_grid_to_config.py": "liberty", "magma_grid_to_config.py": "magma", "rsa_gcv4w_comp_to_config.py": "rsa", "rsa_pvtcar_to_config.py": "rsa",
           "united_commission_to_config.py": "united", "shriram_broker_grid_to_config.py": "shriram", "shriram_2w_to_config.py": "shriram", "godigit_agency_to_config.py": "godigit", "sbi_grid_to_config.py": "sbi", "icici_grid_to_config.py": "icici", "indusind_grid_to_config.py": "indusind"}
results = []
quiet = lambda *a, **k: None


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


def frame(rows):
    """rows = (entry key, pi_type, od, tp, net, flat)"""
    return pd.DataFrame([dict(k=k, product="P", sub_product="1+1", new_rto_list="GJ", new_vehicle_makes="X", fuel_type="Petrol", make_model_class="NA", tariff_max=100.01,
                              pi_type=t, pi_od_rate=od, pi_tp_rate=tp, pi_net_rate=net, pi_flat_amount=fl) for k, t, od, tp, net, fl in rows])


for file in SCRIPTS:
    m = helpers.load_script(file)
    tag = file.replace("_to_config.py", "")
    ask = m.Asker(interactive=False)                         # any question would raise: the rule must not ask
    d = frame([("a", "On Net", 0, 0, 25, 0), ("a", "On Net", 0, 0, 25, 0), ("b", "On Net", 0, 0, 10, 0)])
    out = m.resolve_duplicates(d, ["k"], ask)
    check(f"{tag}: the same payin twice is one row, nothing is asked, nothing is listed", len(out) == 2 and not m.LAST_LOWEST and not ask.decisions, (len(out), m.LAST_LOWEST))
    d = frame([("a", "On Net", 0, 0, 30, 0), ("a", "On Net", 0, 0, 25, 0), ("a", "On Net", 0, 0, 27.5, 0), ("b", "On Net", 0, 0, 10, 0)])
    out = m.resolve_duplicates(d, ["k"], ask)
    got = out.set_index("k")["pi_net_rate"].to_dict()
    check(f"{tag}: different payins keep the LOWEST of three (25), other entries untouched", got == {"a": 25.0, "b": 10.0} and len(out) == 2, got)
    e = m.LAST_LOWEST[0] if m.LAST_LOWEST else {}
    check(f"{tag}: the choice is listed (rates found, the one written) and noted among the answers", len(m.LAST_LOWEST) == 1 and e.get("written") == "25" and e.get("rates_found") == "30 / 25 / 27.5"
          and e.get("rto") == "GJ" and any("lowest" in x for x in ask.decisions), (m.LAST_LOWEST, ask.decisions))
    d = frame([("a", "On OD and TP", 10, 5, 0, 0), ("a", "On Net", 0, 0, 20, 0)])
    out = m.resolve_duplicates(d, ["k"], m.Asker(interactive=False))
    check(f"{tag}: 'On OD and TP' counts as the sum of its two parts (10 + 5 = 15 beats 20)", len(out) == 1 and out.iloc[0]["pi_type"] == "On OD and TP", out.to_dict("records"))
    d = frame([("a", "On Net", 0, 0, 20, 0), ("a", "On OD", 20, 0, 0, 0)])
    out = m.resolve_duplicates(d, ["k"], m.Asker(interactive=False))
    check(f"{tag}: equal payins of two types: the first row stays and nothing is listed", len(out) == 1 and out.iloc[0]["pi_type"] == "On Net" and not m.LAST_LOWEST)
    d = frame([("a", "On Net", 0, 0, 0, 250), ("a", "On Net", 0, 0, 0, 300)])
    out = m.resolve_duplicates(d, ["k"], m.Asker(interactive=False))
    check(f"{tag}: a flat amount counts too (250 beats 300)", len(out) == 1 and out.iloc[0]["pi_flat_amount"] == 250)
    d = frame([("a", "On Net", 0, 0, 5, 0), ("b", "On Net", 0, 0, 6, 0)])
    check(f"{tag}: a frame without repeated keys is returned as it is", m.resolve_duplicates(d, ["k"], m.Asker(interactive=False)) is d and not m.LAST_LOWEST)
    d = frame([("a", "On Net", 0, 0, 30, 0), ("a", "On Net", 0, 0, 25, 0)])
    out = m.resolve_duplicates(d, ["k"], m.Asker(interactive=False))
    path = os.path.join(tempfile.mkdtemp(prefix="lowest_"), "o.csv")
    m.write_csv(out, path, m.Asker(interactive=False), quiet)
    rev = m.review_path(path)
    book = pd.read_excel(rev, sheet_name=None, dtype=str, keep_default_na=False) if os.path.exists(rev) else {}
    low = book.get("Lowest rate chosen")
    check(f"{tag}: write_csv writes the CSV and ONE review workbook next to it, with the sheet 'Lowest rate chosen' (and a Summary that names the file to upload)",
          os.path.exists(path) and low is not None and len(low) == 1 and low.iloc[0]["written"] == "25" and low.iloc[0]["rates_found"] == "30 / 25" and "Summary" in book
          and "Upload THIS file" in " ".join(book["Summary"]["item"]), sorted(book))
    m.resolve_duplicates(frame([("a", "On Net", 0, 0, 5, 0)]), ["k"], m.Asker(interactive=False))
    path2 = os.path.join(tempfile.mkdtemp(prefix="lowest_"), "o.csv")
    m.write_csv(out, path2, m.Asker(interactive=False), quiet)
    m.reset_run_lists()
    m.write_csv(out, path2, m.Asker(interactive=False), quiet)
    check(f"{tag}: with nothing to review only the CSV is written, no workbook", os.path.exists(path2) and not os.path.exists(m.review_path(path2)))
    check(f"{tag}: review_path names the workbook (X_UPLOAD_2026-10.csv -> X_REVIEW_2026-10.xlsx, any other name + _review)", m.review_path("a/X_UPLOAD_2026-10.csv") == "a/X_REVIEW_2026-10.xlsx" and m.review_path("a/b.csv") == "a/b_review.xlsx")
    o = m.default_output("X_UPLOAD_2026-10.csv").replace("\\", "/")
    check(f"{tag}: the default output goes to output/{FOLDERS[file]}/ next to the script, one folder per insurer", o.endswith(f"/output/{FOLDERS[file]}/X_UPLOAD_2026-10.csv") and m.OUTPUT_FOLDER == FOLDERS[file], o)
    rv = m.ManualReview()
    masters = m.load_masters(m.Asker(interactive=False))
    probe = pd.DataFrame([dict(new_rto_list="NO_SUCH_RTO_GROUP", new_vehicle_makes="NO_SUCH_MAKE"), dict(new_rto_list="NO_SUCH_RTO_GROUP", new_vehicle_makes="NO_SUCH_MAKE"), dict(new_rto_list="GJ", new_vehicle_makes="HERO")])
    m.check_names(probe, rv, masters, quiet)
    rto, mk = m.EXTRA_SHEETS.get("RTO groups to create"), m.EXTRA_SHEETS.get("Make groups to create")
    check(f"{tag}: names that are not in the master become the sheets 'RTO groups to create' / 'Make groups to create' (with the rows that carry them)",
          rto is not None and mk is not None and rto.iloc[0]["group_needed"] == "NO_SUCH_RTO_GROUP" and int(rto.iloc[0]["rows_flagged"]) == 2 and mk.iloc[0]["group_needed"] == "NO_SUCH_MAKE", m.EXTRA_SHEETS)
    m.add_sheet("RTO groups to create", pd.DataFrame([dict(group_needed="X", rows_flagged=1, items="")]))
    check(f"{tag}: add_sheet stacks lists added under one name", len(m.EXTRA_SHEETS["RTO groups to create"]) == 2)
    rv2 = m.ManualReview()
    check(f"{tag}: a new run (ManualReview()) starts with empty lists", m.EXTRA_SHEETS == {} and m.RUN_WARNINGS == [] and m.LAST_NOTES == [])

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
