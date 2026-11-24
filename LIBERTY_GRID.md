# LIBERTY GENERAL - the Branch Grid -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/liberty-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` section 7.1 and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. The script finds every table by what its header row says, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `liberty_grid_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | ONE workbook, the Liberty "Branch Grid" (.xlsx): a Private Car + Two-Wheeler comprehensive / SAOD sheet, a Two-Wheeler SATP sheet, a Private Car SATP sheet, a CV below-7.5-tonne sheet and a CV above-7.5-tonne sheet |
| Cells | **fractions** (0.25 = 25%): the output is x100 - Liberty is the only converter that multiplies |
| Output | `output\liberty\LIBERTY_UPLOAD_<yyyy-mm>.csv` (42 columns) + the review workbook when there is anything to look at |
| Dates | given on every run (`--from-date`, `--to-date`) |
| Answer key | the user's own portal export of the same grid (`--ref`); regression cases `liberty_sep26` (23,690 rows) and `liberty_jul26` (23,154 rows); layout snapshots `liberty_jul26`, `liberty_sep26` |
| Tests | `tests/liberty_tests.py`, `tests/liberty_interactive_tests.py` (both in `run_tests.py --all`) |

```
.venv/Scripts/python.exe liberty_grid_to_config.py --grid "<Branch Grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--no-prompt]
```
Without `--no-prompt` the script asks numbered questions about anything ambiguous (a missing column, a header that appears twice) and lists every answer at the end.

## 2. The workbook (found by words)

* **The sheets** are found by their ROLE: Private Car + Two-Wheeler comprehensive and SAOD (the Two-Wheeler columns are used, the Private Car columns on that sheet are NOT), Two-Wheeler SATP, Private Car SATP, CV below 7.5 T, CV above 7.5 T. A sheet that is not called as the settings say is the ONE sheet whose columns are the role's columns (`sheet_by_columns`: at least 80% of the role's columns and clearly more than any other sheet); the run says which sheet it used. The RTO prefix follows the CONFIGURED sheet name (the key of `RTO_PREFIX`), not the sheet the person redirected the run to.
* **The geo table**: the header cell `Geo Segment` is found wherever it sits; every geo is a data row. Each geo row carries a volume-slab column found by the header `Volume Slab`; slab names are compared with `slab_key` (spaces and capitals ignored), so `<1.5L` and `< 1.5L` are one slab. The slab of the Private Car / Two-Wheeler sheet is `50k-1.5L`, the slab of the CV sheets is `< 1.5L`.
* **The CV sheets** hold the comprehensive and the SATP table side by side: a block is the columns under its own header (`Block`); a heading that appears twice inside a block is ASKED about (`Block.dups`), never taken by position.
* Text above a table that is not a known standing note is a note (section 6).

## 3. How the grid becomes rows

| Part of the grid | Rows written |
|---|---|
| Two-Wheeler comprehensive (Bike / Scooter) | 1+1, `On Net`; the SAOD columns: SAOD, `On OD` |
| Two-Wheeler SATP | STP, `On TP`, by engine band (below 75, 75-150, 150-350, above 350 cc; Scooter has two bands) |
| Private Car SATP | STP, `On TP`, by cc band and fuel (Petrol, Diesel, CNG; LPG takes the CNG column; Electric from the EV column) |
| CV comprehensive | 1+1, `On Net`, by vehicle age band (NEW = 0-1.01, 1-5, 5-10, 10+) x weight column (up to 2.5 T ... 20-40 T; the 20-40 T column runs to 43.5 T by the grid's own note); PCV 3W, GCV 3W, Tractor, MISCD columns |
| CV SATP | STP, `On TP`, the same age x weight cards |

Payout = payin - 7, floor 0 (the standing rule). Every cell is a fraction: payin = cell x 100.
**A column the grid does not have is a QUESTION, never a guess and never a silent zero**: when the Private Car SATP sheet has no EV column the person chooses (`Electric same as Petrol` = the three Petrol cc columns, each keeping its own cc band).

## 4. A row

| Column | Value |
|---|---|
| id | `1` on every row (Liberty's convention; Magma counts 1..N) |
| insurance_company / insurer_vertical | `LIBERTY GENERAL INSURANCE LTD` / `Not Applicable` |
| policy_type | `Renew` |
| sub_product | by the part of the grid above (Comp = 1+1, SATP = STP, SAOD = SAOD) |
| pi_type / po_type | `On Net` for comprehensive, `On TP` for STP, `On OD` for SAOD |
| payin / payout | cell x 100; payout = payin - 7 floor 0 |
| is_zd | `NA` |
| from_date / to_date | the dates given for the run, `DD/MM/YYYY` |
| user_id | `2966` |
| new_rto_list | the geo / cluster name as the grid writes it (the cluster rule of 30 Oct 2026); the Pan India Private Car rows say `ALLINDIA` |

## 5. Names: geos and groups

* `new_rto_list` carries the bare geo name by default. With `GRID_RTO_NAMES=group` it is the master's group: `RTO_PREFIX[role]` + the geo, spaces -> `_` - `LIBERTY_PC_TW_APR26_` (Private Car / Two-Wheeler sheets), `LIBERTY_GCV_4W_<7.5T_APR26_` (CV below 7.5 T), `LIBERTY_GCV_4W>7.5T_APR26_` (CV above 7.5 T). The `APR26` in every name is the month the groups were made, not the grid's month: change `RTO_PREFIX` only when the master gets new Liberty groups.
* In the group-name mode a geo the master lacks is flagged by `check_names` and never mapped: the July and September grids flag ASSAM (the master has `...>7.5T_APR26_Assam`) and TELANGANA (the master has `...>7.5T_APR26_TELANGANA_` with a trailing underscore) on the CV above-7.5-T sheet. (In the default cluster mode RTO names are not checked against the master: the person maps and groups the RTOs in the portal.)
* **The fixed Pan India rows**: Private Car comprehensive / SAOD = 27 / 25 on `ALLINDIA`, ten rows, written from a setting - they are NOT in the grid and must never be dropped.

## 6. Notes (text the script cannot apply)

Text above or beside a table that is not in `STANDING_NOTES` is flagged `Manual Review Required: [text]` on the rows of that sheet, never guessed, never asked about, never a stop. Example: the grid note `HCV3 upto 43.5T for the GEOs highlighted in green` is a colour rule the script cannot see - it is flagged on the rows of the CV above-7.5-T sheet (5,880 and 5,600 flagged rows in the two regression cases). A note the person resolved in the web app is never flagged again.

## 7. Edge cases and what the script does

| Situation | What happens |
|---|---|
| The volume-slab column moves (it did when contact columns were added) | found by its header: same rows |
| `<1.5L` vs `< 1.5L` spelling (GOA and the MAHARASHTRA geos) | one slab (`slab_key`) |
| No EV column on the Private Car SATP sheet | a question (above) |
| A sheet renamed | found by its columns; the run names the sheet it used |
| `Geo Segment` in another column | found by its words |
| A header appears twice in a block | asked (`Block.dups`) |
| A banner / title text above a table | flagged in `remarks`; the rates are the same |
| A geo the master lacks (group mode) | written and flagged by `check_names`, never mapped |

## 8. Checking a run

* `python grid_doctor.py check "<grid>"` first: columns move from month to month. Then run with `--ref` and explain every number that differs from their export (rows new / removed / changed).
* **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a geo by its label cell and a weight column by its header text, read the crossing cell with openpyxl by searching for those texts (never by typing an address), multiply by 100 and compare with the matching upload row - e.g. the GOA geo under the 7.5-12 T band; payout = payin - 7.
* Read every `Manual Review Required` note and, in the group-name mode, the geos the master lacks.
* `python check_regressions.py --only liberty_sep26` (and `liberty_jul26`): the output must be byte-identical.

## 9. Decisions (the user) - do not re-ask

1. Pan India Private Car comprehensive / SAOD = 27 / 25 on `ALLINDIA` (a setting, not in the grid).
2. Liberty `id` = 1 on every row. Slab names: `50k-1.5L` for Private Car / Two-Wheeler, `< 1.5L` for CV.
3. A missing column is a question, never a guess and never a silent zero.
4. Text above the tables that is not a standing note is flagged on that sheet's rows.
5. Structure shifts (11 Nov 2026): the converter gives the SAME upload for a re-laid-out grid or a clear stop; a banner above a table only adds a remark.

## 10. Open

Nothing is recorded as open in `AGENT_CONTEXT.md` 7.1. The acceptance of the July and September uploads is not written down there: ask the user for their upload as the answer key each month.

## 11. Files

`liberty_grid_to_config.py` - `tests/liberty_tests.py`, `tests/liberty_interactive_tests.py`, `tests/structure_shift_tests.py` (subject `liberty`) - `regression/inputs/liberty_sep26.xlsx`, `liberty_jul26.xlsx`, cases `liberty_sep26` / `liberty_jul26` in `regression/cases.json` - `layouts/liberty_sep26.json`, `layouts/liberty_jul26.json` - `.claude/skills/liberty-grid/SKILL.md` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
