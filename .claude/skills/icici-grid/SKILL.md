---
name: icici-grid
description: ICICI Lombard knowledge for grid_tools - FOUR workbooks (CV grid, Private Car grid, Two-Wheeler grid, new Two-Wheeler 1+5 grid) recognised by what their sheets say and converted to ICICI_UPLOAD_<yyyy-mm>.csv by icici_grid_to_config.py. Load it with grid-converter whenever the user sends an ICICI / ICICI Lombard grid, an ICICI rates_export or upload to review, or asks about Council Call cells, OTHERS:0%, ICICI-RB-Offline, sub-cluster rates, or what the ICICI script does.
---

# ICICI Lombard - CV, Private Car, Two-Wheeler, new Two-Wheeler

Dedicated reference: `ICICI_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` section 7.10, `CLAUDE.md` (the ICICI bullet), P-43 / P-53, F-44 / F-45 / F-46, memory `icici-script`. (This skill summarises what another session recorded; re-read 7.10 before changing the script.)

## Status
First build 4 Nov 2026, **NOT yet accepted by the user**; they confirmed the assumptions the same day (below). Regression: `icici_cv_oct26`, `icici_pc_tw_nov26`, `icici_cv_nov26` (no answer key yet: the November CV upload was not made); `tests/icici_tests.py`.
- (5 Nov 2026) November 2026 outputs written on 4 Nov to `output\icici\`: `ICICI_UPLOAD_2026-11_PC_TW_TW_NEW.csv` (3,342 rows) and `ICICI_UPLOAD_2026-11_CV.csv` (20,317 rows, 7,586 flagged, 307 cells / columns in `Not written`), dates 01/11/2026 - 31/03/2030 ASSUMED: the CV workbook states no date, and the new Two-Wheeler workbook is called 'NEW 15 Nov' although the user's own upload dated its rows 01/11 (`--tw-new-from-date` changes it). The user restarted their web app on 5 Nov 2026 and the dropdown lists ICICI; offered a restart by me, they answered 'No, I'll do it' - do not kill their server.

## Run it
```
.venv/Scripts/python.exe icici_grid_to_config.py [--cv-grid cv.xlsx] [--pc-grid pc.xlsx] [--tw-grid tw.xlsx] [--tw-new-grid twnew.xlsx] (or --grid once per file) --from-date DD/MM/YYYY --to-date DD/MM/YYYY [--cv-from-date .. --pc-from-date .. --tw-from-date .. --tw-new-from-date ..] [--vertical ..] --ref "<their export>.csv" --no-prompt
```
Any subset of the four workbooks works (a run with only some adds `_CV` / `_PC_TW...` to the file name). The web app offers four OPTIONAL inputs; the server refuses a run with none.

## How it reads the workbooks
Nothing by cell address or exact sheet name (P-43): `classify_workbook` recognises each file by its sheets; the CV matrices are the sheets whose header row holds >= 6 known segments (`find_cv_matrices`); the cluster column is the one whose header says cluster / segment and whose values are text - NOT the user's own `RTO` column (F-45). CV: COMP = 1+1 On Net, AOTP / TP = STP On TP; one row per cluster, one column per segment (`classify_segment`); cells are small sentences read as boxes by `parse_cell` / `regions` (P-53: clauses by commas, words by `|`, rate after `:`; policy, NEW / OLD, age bands, ZD / NZD, makes, vehicle types -> add_tnc, places, OTHERS, CC). Private Car: one sheet with a two-row header (New 1+3 / SAOD / Comprehensive with NCB per fuel, Non-NCB one rate, Act only = STP On TP); comprehensive is **On OD** (F-44). Two-Wheeler `TW Old` (Bike / Scooter x Comp / SAOD / AOTP) and the new TW (one sheet per OEM, 1+5 New On Net age 0-99).

## Rules the user has set / confirmed (4 Nov 2026 - do not re-ask)
- `insurer_vertical` = `ICICI-RB-Offline` on EVERY row (also Private Car and Two-Wheeler; `--vertical` overrides).
- `OTHERS:0%` writes NO zero card (`WRITE_OTHERS_ZERO = False`; listed in `Not written`). A **Council Call** (`CC`) cell has no fixed rate: its entries are listed in the sheet `Council call cells` and NEVER turned into rates; `CC,OTHERS:n%` writes the OTHERS rate, flagged.
- Rates are written to 8 decimals (`RATE_DECIMALS = 8`, their own precision). A plain percentage, `No Biz`, `0%` or blank = a ZERO card. Bharat Benz is left out of the default make groups because CV NOTES say it is declined (read from the note; asked about when the note has gone).
- STP rows: a segment the AOTP matrix lists takes its payout from there; every other segment takes the comprehensive cell unless the cell says `AOTP:n%`. STP rows carry NCB / ZD `NA`.

## Monthly procedure (on top of grid-converter section 2)
`grid_doctor.py check` each workbook; run all that came; compare Private Car + TW with their upload (the November one reproduced all 3,318 rows, only `insurer_vertical` differing by their decision); for CV compare by the portal's group names (`GRID_RTO_NAMES=group`); read `Not written`, `Council call cells` and the Negative RTO list; report cells/columns the script did not understand.

## Open items
Council Call integration (which cluster an entry applies to is hand work), the sub-cluster rows' portal groups, `Used car`, the `Nov'26 CHNAGES` sheet as a cross-check (idea), the 5+5 two-wheeler grid. The user's running web server did not list the script until it was restarted (F-46): they restarted it themselves on 5 Nov 2026 and it lists ICICI (resolved).
- (5 Nov 2026) Jharkhand taxi up to 1000 cc, Old vehicles: the grid says 40, the user's October upload 35 (8 rows; the script follows the grid) - ask which is right. 40 November CV cells whose make set EECO + MAGIC + OMNI has no group in the make master are not written (listed in `Not written`): the user creates the group if they want them. Lakshadweep (24 Two-Wheeler rows, in the grid but not in their upload) and the sub-cluster / district rows (`<cluster> - Condition: ... - Part: ...`) need RTO mappings in the portal - the user maps RTOs themselves (cluster-name mode).

## Traps seen
F-44 (comprehensive is On OD for Private Car), F-45 (the user's own `RTO` column is not the cluster column), F-46 (the user's running web server did not list the script until restarted), P-43 (find workbooks and sheets by what they say), P-53 (cells are small sentences read as boxes).
- (11 Nov 2026) Structure shifts (11 Nov 2026): blank rows inside the CV matrix / Private Car / Two-Wheeler sheets are skipped (a text row ends the table: 20,317 CV rows used to become 11,264); the reference sheets (RTO list, Council Call) are found within the first 25 rows (were 4); a column heading that is there TWICE gives the LOWEST payin (the user, 12 Nov 2026: `twin_groups` / `twin_blank` - the rows of both columns are written, the lowest-rate rule keeps the lowest and lists it in 'Lowest rate chosen', a blank twin cell is no payin; it used to replace the first column in silence); two payout columns with DIFFERENT headings on a new Two-Wheeler sheet still stop; the new Two-Wheeler OEM is read from the words of the sheet name (`tw_oem_key`: 'HMC v2', 'Hero (Sept)'), two sheets that say one OEM stop unless one is called exactly the OEM, and a workbook of which no sheet could be read stops. Every sheet called 'SheetN' stops for the CV and the new Two-Wheeler workbooks (the notes sheet / the OEM is the name) - by design.
