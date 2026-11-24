---
name: liberty-grid
description: Liberty General knowledge for grid_tools - the Liberty "Branch Grid" workbook (.xlsx: PC TW_Comp and SAOD, TW SATP, PC SATP, CV <7.5T and CV >7.5T sheets with a Volume Slab column) converted to LIBERTY_UPLOAD_<yyyy-mm>.csv by liberty_grid_to_config.py. Load it with grid-converter whenever the user sends a Liberty grid or a Liberty rates_export to review, or asks about volume slabs, the fixed Pan India Private Car 27 / 25 rows, the EV column question, or what the Liberty script does.
---

# Liberty - Branch Grid

Dedicated reference: `LIBERTY_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` section 7.1, `CLAUDE.md` (the Liberty bullet), `lessons_learned.md` (2026-07 liberty entries).

## Status and checks
Regression cases `liberty_sep26` (23,690 rows) and `liberty_jul26` (23,154 rows); tests `tests/liberty_tests.py`, `tests/liberty_interactive_tests.py` (run in `--all`). Master flags on the CV >7.5T sheet: ASSAM and TELANGANA (the master's spelling differs) - flagged by `check_names`, never mapped.

## Run it
```
.venv/Scripts/python.exe liberty_grid_to_config.py --grid "<Branch Grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--no-prompt]
```
Output `output\liberty\LIBERTY_UPLOAD_<yyyy-mm>.csv` (+ review workbook).

## How the grid becomes rows
- Sheets: `PC TW_Comp and SAOD` (TW 1+1 and SAOD, slab `50k-1.5L`; the PC columns on it are NOT used), `TW SATP`, `PC SATP`, `CV <7.5T Comp & SATP`, `CV >7.5T Comp & SATP` (comprehensive and SATP tables side by side, slab `< 1.5L`). Geo rows carry a volume-slab column found by the header `Volume Slab`; slab names are compared with `slab_key` (ignoring spaces and capitals: GOA and the MAHARASHTRA geos write `<1.5L`).
- **Cells are fractions (0.25 = 25%): the output is x100.** TW comp 1+1 (Bike / Scooter, On Net) + SAOD (On OD), TW STP by cc, PC STP by cc and fuel (LPG uses the CNG column; Electric from the EV column), CV comp (On Net) and STP (On TP) by age (NEW 0-1.01, 1-5, 5-10, 10+) x column (Upto 2.5T ... 20-40T), PCV 3W, GCV 3W, Tractor, MISCD.
- `new_rto_list`: the cluster / geo name by default; with `GRID_RTO_NAMES=group` it is the master's group `RTO_PREFIX[sheet]` + geo (`LIBERTY_PC_TW_APR26_`, `LIBERTY_GCV_4W_<7.5T_APR26_`, `LIBERTY_GCV_4W>7.5T_APR26_`; the `APR26` stays whatever the grid's month - change `RTO_PREFIX` only when the master gets new Liberty groups). The Pan India Private Car rows stay `ALLINDIA`.

## Rules the user has set (do not re-ask)
- **Pan India Private Car comprehensive / SAOD = 27 / 25 on `ALLINDIA`** (a setting, 10 rows, not in the grid; must never be dropped). Liberty `id` = 1 on every row. Liberty slabs: `50k-1.5L` for PC / TW, `< 1.5L` for CV.
- A missing column is a question, never a guess and never a silent 0: when the PC SATP sheet has no EV column the person chooses ("Electric same as Petrol" = the three Petrol cc columns, each keeping its cc band).
- Text above the tables that is not in `STANDING_NOTES` is flagged on that sheet's rows (e.g. "HCV3 upto 43.5T for the GEOs highlighted in green" - a colour rule - on the CV >7.5T rows).

## Monthly procedure (on top of grid-converter section 2)
`grid_doctor.py check` first (columns move: the slab column moved in July). Run with `--ref`. Spot-check 3-5 cells **x100** (e.g. GOA 7.5-12T bands). Read the geos the master lacks.

## Traps seen
Slab column moved when SPOC / ZH columns were added (find columns by header, never by position); `<1.5L` vs `< 1.5L`; no Electric column in PC SATP. Liberty cells are fractions - the only converter where the output is x100.
- (11 Nov 2026) Structure shifts (11 Nov 2026): the 'Geo Segment' header is found in any column (not only the leftmost column), a sheet that is not called 'PC TW_Comp and SAOD' / 'TW SATP' / 'PC SATP' / 'CV <7.5T ...' is found by its COLUMNS (`sheet_by_columns`: at least 80% of the role's columns and clearly more than any other sheet; the run says which sheet it used), a twin header is asked about (`Block.dups`). Harness: `tests/structure_shift_tests.py` (every operator IDENTICAL, a banner in title_rows only adds a remark).

## Open items
None recorded in `AGENT_CONTEXT.md` 7.1. The acceptance of the July / September uploads is not written down there: ask the user for their upload as the answer key each month.
