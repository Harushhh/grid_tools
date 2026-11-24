"""The STP rule of the user (27 Oct 2026, repeated 4 Nov 2026: "there should be always NA in NCB and ZD if it is STP"): every STP row of every script has is_ncb = NA and is_zd = NA.
check_regressions.py enforces it on the output of every known-good case (`stp_violations`); this suite tests that check itself and the Shriram broker script on a grid that tries to break the rule.
Run:  .venv\\Scripts\\python.exe tests\\stp_rule_tests.py"""
import datetime as dt
import importlib.util
import os
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

results = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + name + ("" if cond else f"  -> {detail}"))


spec = importlib.util.spec_from_file_location("check_regressions", os.path.join(helpers.ROOT, "check_regressions.py"))
cr = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cr)
tmp = tempfile.mkdtemp(prefix="stp_rule_")


def csv_of(rows, name):
    path = os.path.join(tmp, name)
    pd.DataFrame(rows, columns=["sub_product", "is_ncb", "is_zd"]).to_csv(path, index=False)
    return path


check("a file whose STP rows are all NA / NA has no violation", cr.stp_violations(csv_of([("STP", "NA", "NA"), ("1+1", "YES", "NO"), ("STP", "NA", "NA")], "ok.csv")) == 0)
check("an STP row with is_ncb YES is a violation (the 10 rows typed into the November 2026 portal export)", cr.stp_violations(csv_of([("STP", "YES", "NA"), ("STP", "NA", "NA")], "ncb.csv")) == 1)
check("an STP row with is_zd NO / YES is a violation, a row with both counts once", cr.stp_violations(csv_of([("STP", "NA", "NO"), ("STP", "NA", "YES"), ("STP", "YES", "YES")], "zd.csv")) == 3)
check("the sub-product is matched ignoring capitals, and a package row with NCB / ZD is fine", cr.stp_violations(csv_of([("stp", "YES", "NA"), ("1+1", "YES", "YES"), ("SAOD", "YES", "NO")], "case.csv")) == 1)
side = os.path.join(tmp, "side.csv")
pd.DataFrame([("a", "b")], columns=["rto", "makes"]).to_csv(side, index=False)
check("a side file without the three columns has no violation", cr.stp_violations(side) == 0)

# the Shriram broker script on a grid that asks for NCB and zero dep on a P & L row: the package rows keep them, the STP rows have NA / NA
B = helpers.load_script("shriram_broker_grid_to_config.py")
HDR = ["STATE", "PRODUCTS", "DIS %", "% PO", "POLICY TYPE", "AGE", "UW REMARKS", "REMARKS 1", "EFFECTIVE DATE"]
rows = [["GOA", "GCCV LCV UPTO 2800 GVW", 80, 45, "P & L", "upto 15 years & upto 20 Years for STP", "NCB Cases 20% and Above, Except New Mahindra"],
        ["GOA", "PCCV 4W Upto 6+1", 75, 42.5, "P & L", "upto 15 years & upto 20 Years for STP", "for NIL dep cases payout is 40%"],
        ["GOA", "PRIVATE CAR PETROL", 85, "20 OD", "P & SA-OD", "upto 15 years.", "NCB Cases"]]
wb = helpers.FakeWB({"GRID NOV' 26": [HDR] + [r + [None] * (9 - len(r)) for r in rows]})
df, _ = B.build(None, dt.date(2026, 11, 1), dt.date(2030, 3, 31), lambda *a, **k: None, B.Asker(interactive=False), wb=wb)
stp, pkg = df[df.sub_product == "STP"], df[df.sub_product != "STP"]
out = os.path.join(tmp, "broker.csv")
df.to_csv(out, index=False)
check("Shriram broker: NCB and zero-dep asked for on P & L rows land on the package rows only - every STP row is NA / NA", len(stp) > 0 and set(stp.is_ncb) == {"NA"} and set(stp.is_zd) == {"NA"} and set(pkg.is_ncb) >= {"YES"} and set(pkg.is_zd) >= {"YES", "NO"}
      and cr.stp_violations(out) == 0, (set(stp.is_ncb), set(stp.is_zd), set(pkg.is_ncb), set(pkg.is_zd)))

passed = sum(1 for _, ok, _ in results if ok)
print(f"\n{passed} passed, {len(results) - passed} failed")
sys.exit(0 if passed == len(results) else 1)
