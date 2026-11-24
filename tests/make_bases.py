"""Rebuild the in-memory starting points of the fuzzers (pickles) from the known-good grids in regression\\inputs.
Run by tests\\run_tests.py when a pickle is missing:  .venv\\Scripts\\python.exe tests\\make_bases.py"""
import os
import pickle
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers  # noqa: E402

HERE = helpers.HERE
pickle.dump({"lib_sep": helpers.liberty_base("sep"), "lib_jul_b": helpers.liberty_base("jul"),
             "lib_jul_a": helpers.sheets_of(os.path.join(helpers.INPUTS, "liberty_jul26_v1.xlsx"))},
            open(os.path.join(HERE, "fuzz_lib_bases.pkl"), "wb"))
motor, satp = helpers.magma_base()
pickle.dump({"motor": motor, "satp": satp}, open(os.path.join(HERE, "fuzz_mag_bases.pkl"), "wb"))
pickle.dump({"rsa_oct": helpers.rsa_base("oct"), "rsa_v2": helpers.rsa_base("v2")}, open(os.path.join(HERE, "fuzz_rsa_bases.pkl"), "wb"))
print("bases written to", HERE)
