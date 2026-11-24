# ROYAL SUNDARAM (RSA) - the GCV 4W comprehensive grid and the Private Car grid -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/rsa-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` sections 7.3 and 7.4 and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. The script finds every table by what its header row says, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).
RSA is built **product by product**: one script per product, the same shared question block; other RSA products will be separate scripts under this insurer.

## 1. At a glance

| | GCV 4W comprehensive | Private Car |
|---|---|---|
| Script | `rsa_gcv4w_comp_to_config.py` | `rsa_pvtcar_to_config.py` |
| Input | ONE .xlsx, one sheet per region (WEST, EAST, NORTH, SOUTH); the rules are in the NOTES under each sheet | ONE .xlsx: four region sheets (West / North / East / South) with Comp + SAOD + New tables, the `Top End Variant List`, the `Pvt Car STP` sheet |
| Cell means | the payout for TATA and Ashok Leyland, Old business (age 1+), without Nil Dep | the payout **fraction** by cluster x engine band x With / Without NCB |
| Output | `output\rsa\RSA_GCV4W_COMP_UPLOAD_<yyyy-mm>.csv` | `output\rsa\RSA_PVTCAR_UPLOAD_<yyyy-mm>.csv` |
| Regression | `rsa_gcv4w_oct26`, `rsa_gcv4w_v2_oct22` (67,392 rows each, 21,168 flagged) | `rsa_pvtcar_oct26`, `rsa_pvtcar_sep26` (19,620 rows each) |
| Tests | `tests/rsa_tests.py`, `tests/rsa_tests_notes.py` | `tests/pvtcar_tests.py` |
| Layouts | `rsa_gcv4w_oct26`, `rsa_gcv4w_v2_oct22` | `rsa_pvtcar_oct26`, `rsa_pvtcar_sep26` |

```
.venv/Scripts/python.exe rsa_gcv4w_comp_to_config.py --grid "<GCV grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their export>.csv" [--prefix ROYAL_<MON><YY>_] [--no-prompt]
.venv/Scripts/python.exe rsa_pvtcar_to_config.py     --grid "<Private Car grid>.xlsx" --from-date ... --to-date ... --ref ... [--prefix ...]
```
Dates are given on every run. Without `--no-prompt` the scripts ask numbered questions about anything ambiguous and list every answer at the end.

## 2. GCV 4W comprehensive

### 2.1 The workbook (found by words)
* A region sheet is the sheet that says `Tonnage Band` in a header cell (found wherever it is, in any column); the title and the fuel label are the TEXTS ABOVE the header, not a fixed number of lines from the top. Hidden sheets are never read. **Two sheets that hold the same clusters (last month's left in) STOP the run** (`expected one sheet per region`); a repeated cluster column inside a sheet is settled by the lowest-rate rule, not stopped. A blank row inside a table does not end it: the next non-blank row decides (`_band_follows`).
* The cell holds only the payout for TATA and Ashok Leyland, Old business, without Nil Dep. **Everything else is in the NOTES under each sheet**, read by `read_rules` / `interpret` patterns with several wordings each (`3% lower`, `reduced by 3`, `deduction of three`, `w.e.f. 22 Oct 2026`).

### 2.2 What the notes do (October 2026 numbers; they are READ from the notes each month, never hard-coded)
| Note | Rows |
|---|---|
| Eicher | payout 3 points lower |
| other makes | 5 points lower |
| Nil Dep | 2 points lower |
| New CV | 2 points lower (EAST 1; per weight when the note says so; only when the cell is above 10) |
| TATA and Ashok Leyland in the lower discount band | 2 points higher (tariff 0 to 81.01 and 81.01 to 86.01; EAST 0 to 85.01 and 85.01 to 90.01); only when the cell is above 0 |
| floor | a payout never goes below 0; the four non-EV fuels (Petrol, Diesel, CNG, LPG) are identical |

* **Every pattern USES UP the words it understands.** Any text still holding a number afterwards is either a region-specific discount note or unread - BOTH go to `remarks` as `Manual Review Required`, never a guess, never a question, never a stop.
* **When Royal rewords a rule, add a pattern and a test in `tests/rsa_tests_notes.py` - never hard-code the month's number.** `grid_doctor.py check` diffs the NOTES between months: run it first and it shows which sentence changed.
* **The NORTH region-specific discount notes** (a different maximum discount per cluster and weight: MP, Rajasthan, Haryana, NCR, UP clusters) are NOT applied and are flagged on the clusters they name - the user will brief a skill for them; do not guess their encoding.

### 2.3 A row and names
`PRODUCT` GCV 4W, sub-product 1+1, `Renew`, fuels Petrol / Diesel / CNG / LPG, make group `TATA,Ashok` (the master's spelling: the October upload had `ASHOK`), `insurance_company` `ROYAL SUNDARAM GENERAL INSURANCE COMPANY LIMITED`, `insurer_vertical` `Not Applicable`, `user_id` 2966, `tariff_max` 100.01 where the notes give no limit. `new_rto_list` = the cluster as the grid writes it; with `GRID_RTO_NAMES=group`: `RTO_PREFIX` (`ROYAL_SEPT26_`) + the cluster, spaces and commas -> `_` (`DADRA & NAGAR HAVELI` -> `DADRA_AND_NAGAR_HAVELI`).

## 3. Private Car

### 3.1 The workbook and the rows
* The comprehensive tables give: **Comp -> 1+1, SAOD -> SAOD, New -> 1+3**, five fuels (Petrol, Diesel, CNG, LPG, Electric) at the same rate, by cluster x engine band (up to 1000, 1000-1500, above 1500 cc) x With / Without NCB.
* **Make groups**: the top-variant group (`MAKES_TOP`, `ROYAL_SEP26_Skoda_TATA_MAHINDRA_KIA_Maruti_MG_Toyota_Hyundai_SPE`) = the cell + the extra; every other make (`MAKES_OTHER`, `ROYAL_SEP26_EX_...`) = the cell. The extra is read from the notes: in October the flat `Additional N%`; in September the base-payout lines of the `Top Variant Models Drive` block (+4 With NCB, +2 Without NCB); only when the cell is above 0.
* **New car**: the top group = the cell; the other makes = the note's flat 28% (`OTHER_MAKES_NEW = "write"`, the user's choice of 24 Oct 2026 - the October upload had no such rows; `"flag"` gives the October upload's rows and flags the New rows instead).
* **The STP sheet** gives `STP` rows (Diesel / Non Diesel columns, make group `Private_Car_ALL`, `On TP`). Payout = payin - 7, floor 0.
* **Rounding**: rates are rounded to 2 decimals FROM THE FULL-PRECISION CELL, never from an already-rounded one (a 4-decimal step once turned 9.72496 into a false tie and 9.73).
* Notes: `read_rules` reads `Additional N%`, `other make New car payout N% flat` and the `makes` note; the VLI texts and titles are `STANDING_NOTES`; anything else rule-sized is flagged (`note_scope` names the clusters).

### 3.2 Names
`new_rto_list` = the cluster as the grid writes it; with `GRID_RTO_NAMES=group`: `RTO_PREFIX` + `PVT_CAR_` / `PVT_STP_CAR_` + the name, following `CLUSTER_NAME_EXCEPTIONS` (Baroda -> VADODARA, Delhi NCR -> Delhi_NCR, Mysore ...); a header spelling that is NOT in that list is just capitalised with spaces -> `_`, so **a new header spelling is a portal-group risk**: check the diff against the last upload. The October upload lacked three North-East clusters the grid has; the script writes them.

## 4. Edge cases and what the scripts do

| Situation | What happens |
|---|---|
| A note reworded (`3% lower` -> `reduced by 3`) | read by another pattern of the same rule; a wording nobody reads keeps its number in the leftover text -> flagged whole |
| A note with a number nothing uses up | flagged `Manual Review Required` on the rows it may affect (the clusters it names) |
| A region sheet appears twice | stops: `expected one sheet per region` |
| Blank row inside a table | does not end the table |
| A hidden sheet | never read |
| The title / fuel label sit in different places | read from the texts above the header |
| A new cluster header spelling (Private Car) | capitalised with spaces -> `_`; a portal-group risk the person or agent catches in the diff against the last upload |
| NORTH discount notes | flagged, not applied |

## 5. Checking a run

* `python grid_doctor.py check` first - it diffs the NOTES, so it names the sentence that changed. Run with `--ref`, read every flagged note, compare with the user's upload, and report the NORTH notes as not applied.
* **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a cluster by its label cell and a tonnage band (GCV) or an engine band and NCB column (Private Car) by its header texts, read the crossing cell with openpyxl by searching for those texts (never by typing an address), apply the notes' numbers and compare with the matching upload row; payout = payin - 7 (Private Car) or the cell and its note adjustments (GCV).
* `python check_regressions.py --only rsa_gcv4w_oct26` (and the other three cases): byte-identical output.

## 6. Decisions (the user) - do not re-ask

1. The NORTH discount notes are NOT applied (flagged); the user will brief a skill - do not guess their encoding.
2. `OTHER_MAKES_NEW = "write"`: New rows for the other makes at the note's flat rate (24 Oct 2026).
3. Rates are rounded from the full-precision cell.
4. When a note is reworded, add a pattern and a test - never a hard-coded number.
5. Structure shifts (11 Nov 2026): the converters give the SAME upload for a re-laid-out grid or a clear stop.

## 7. Traps seen

The V2 grid of 22 Oct 2026 rewrote the notes and the old script mis-read them silently (F-12); the 4-decimal intermediate rounding that turned 9.72496 into 9.73; a new cluster header spelling is a portal-group risk (`CLUSTER_NAME_EXCEPTIONS`).

## 8. Open

The NORTH region discount notes (a different maximum discount per cluster and weight).

## 9. Files

`rsa_gcv4w_comp_to_config.py`, `rsa_pvtcar_to_config.py` - `tests/rsa_tests.py`, `tests/rsa_tests_notes.py`, `tests/pvtcar_tests.py`, `tests/fuzz_rsa.py`, `tests/fuzz_pvtcar.py`, `tests/structure_shift_tests.py` (subjects `rsa_gcv4w`, `rsa_pvtcar`) - `regression/inputs/rsa_gcv4w_oct26.xlsx`, `rsa_gcv4w_v2_oct22.xlsx`, `rsa_pvtcar_oct26.xlsx`, `rsa_pvtcar_sep26.xlsx` and their cases - `layouts/rsa_*.json` - `.claude/skills/rsa-grid/SKILL.md` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
