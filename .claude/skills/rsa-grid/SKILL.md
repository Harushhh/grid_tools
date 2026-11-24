---
name: rsa-grid
description: Royal Sundaram (RSA) knowledge for grid_tools - the GCV 4W comprehensive grid (.xlsx, one sheet per region WEST / EAST / NORTH / SOUTH, rules in the notes under each sheet) and the Private Car grid (.xlsx: four region sheets, Top End Variant List, Pvt Car STP) converted by rsa_gcv4w_comp_to_config.py and rsa_pvtcar_to_config.py. Load it with grid-converter whenever the user sends a Royal Sundaram / RSA grid, a ROYAL_* rates_export to review, or asks about Eicher / other makes / Nil Dep / New adjustments, the NORTH discount notes, top-variant makes, the 28% New rate, or what the RSA scripts do.
---

# Royal Sundaram (RSA) - GCV 4W comprehensive and Private Car

Dedicated reference: `RSA_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` sections 7.3 and 7.4, `CLAUDE.md` (the two RSA bullets), F-12 (the V2 grid), `tests/rsa_tests.py`, `tests/rsa_tests_notes.py`, `tests/pvtcar_tests.py`. RSA is built product by product: other RSA products will be separate scripts with the same shared question block.

## Run it
```
.venv/Scripts/python.exe rsa_gcv4w_comp_to_config.py --grid "<GCV grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their export>.csv" [--prefix ROYAL_<MON><YY>_] [--no-prompt]
.venv/Scripts/python.exe rsa_pvtcar_to_config.py     --grid "<Private Car grid>.xlsx" --from-date ... --to-date ... --ref ... [--prefix ...]
```
Outputs `output\rsa\RSA_GCV4W_COMP_UPLOAD_<yyyy-mm>.csv` and `RSA_PVTCAR_UPLOAD_<yyyy-mm>.csv`.

## GCV 4W comprehensive
- The cell is the payout for TATA & Ashok Leyland, Old business (age 1+), without Nil Dep. **Everything else comes from the NOTES under each sheet**, read by `read_rules` / `interpret` patterns (several wordings each: "3% lower", "reduced by 3", "deduction of three", "w.e.f. 22 Oct 2026"): Eicher -3, other makes -5, Nil Dep -2, New CV -2 (EAST -1; per weight when the note says so; only when the cell is above 10), TATA & AL +2 in the lower discount band (tariff 0-81.01 / 81.01-86.01; EAST 0-85.01 / 85.01-90.01; only when the cell is above 0); floor 0; four non-EV fuels identical.
- Every pattern USES UP the words it understands; any text still holding a number is a region-specific discount note or unread - BOTH go to `remarks` as manual-review flags. **When Royal rewords a rule, add a pattern and a test in `tests/rsa_tests_notes.py` - never hard-code the month's number.** `grid_doctor.py check` diffs the NOTES between months: run it first, it shows which sentence changed.
- The **NORTH region-specific discount notes** (different maximum discount per cluster and weight; MP / Rajasthan / Haryana / NCR / UP clusters) are NOT applied and are flagged on the clusters they name - the user will brief a skill for them; do not guess their encoding.
- `new_rto_list`: the cluster name by default; with `GRID_RTO_NAMES=group` `RTO_PREFIX` (`ROYAL_SEPT26_`) + cluster (spaces and commas -> `_`; `DADRA & NAGAR HAVELI` is an exception). Make group `TATA,Ashok` (the master's spelling; the October upload had `ASHOK`).
- Regression `rsa_gcv4w_oct26` and `rsa_gcv4w_v2_oct22` (67,392 rows each, 21,168 flagged).

## Private Car
- Cell = payout fraction by cluster x engine band (Upto 1000, 1000-1500, > 1500) x With / Without NCB: Comp -> 1+1, SAOD -> SAOD, New -> 1+3; five fuels, same rate. Top-variant make group (`MAKES_TOP`) = cell + the extra (October: the note's flat "Additional N%"; September: the base-payout lines of the "Top Variant Models Drive" block, +4 With NCB / +2 Without NCB; only when the cell is above 0); every other make (`MAKES_OTHER`) = the cell; New car: top group = cell, other makes = the note's flat 28% (`OTHER_MAKES_NEW = "write"`, the user's choice of 24 Oct; `"flag"` gives the October upload's rows). The STP sheet gives `STP` rows (Diesel / Non Diesel columns, make `Private_Car_ALL`, On TP).
- **Rates are rounded to 2 decimals FROM THE FULL-PRECISION CELL** (never from an already-rounded one: a 4-decimal step once turned 9.72496 into a false tie and 9.73). Names: `RTO_PREFIX` + `PVT_CAR_` / `PVT_STP_CAR_` + name with `CLUSTER_NAME_EXCEPTIONS` (Baroda -> VADODARA, Delhi NCR -> Delhi_NCR, Mysore ...): a NEW header spelling is a portal-group risk - check the diff against the last upload. The October upload lacked 3 North-East clusters the grid has; the script writes them.
- Notes: `read_rules` reads "Additional N%", "other make New car payout N% flat" and "makes"; VLI texts and titles are `STANDING_NOTES`; anything else rule-sized is flagged (`note_scope` names the clusters). Regression `rsa_pvtcar_oct26`, `rsa_pvtcar_sep26` (19,620 rows each).

## Monthly procedure (on top of grid-converter section 2)
`grid_doctor.py check` (notes diff) -> run with `--ref` -> read every flagged note -> compare with their upload; report the NORTH notes as not applied.

## Status
GCV 4W: regression `rsa_gcv4w_oct26` and `rsa_gcv4w_v2_oct22`. Private Car: `rsa_pvtcar_oct26`, `rsa_pvtcar_sep26` (checked against the user's uploads; the `OTHER_MAKES_NEW = "write"` rows are new against the October upload by the user's choice of 24 Oct 2026).

## Rules the user has set (do not re-ask)
- The NORTH region discount notes are NOT applied (flagged); the user will brief a skill - do not guess their encoding. `OTHER_MAKES_NEW = "write"`: New rows for the other makes at the note's flat rate (24 Oct). Rates are rounded from the full-precision cell. When a note is reworded, add a pattern and a test - never a hard-coded number.

## Traps seen
F-12 (the V2 grid of 22 Oct rewrote the notes and the old script mis-read them silently); the 4-decimal intermediate rounding that turned 9.72496 into 9.73; a new cluster header spelling is a portal-group risk (`CLUSTER_NAME_EXCEPTIONS`).
- (11 Nov 2026) Structure shifts (11 Nov 2026): both RSA scripts: a blank row inside a table does not end it (the next non-blank row decides: `_band_follows`, `_table_follows`, `_division_follows`); two sheets that hold the same clusters / divisions (last month's left in) STOP the run ('expected one sheet per region'); the GCV script finds 'Tonnage Band' in any column, takes the title / fuel label from the texts above the header (not a fixed number of lines from the top) and does not read hidden sheets.

## Open items
The NORTH region discount notes (different maximum discount per cluster and weight).
