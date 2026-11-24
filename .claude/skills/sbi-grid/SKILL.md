---
name: sbi-grid
description: SBI General knowledge for grid_tools - the "Agent, C & D Cat. Broker Grid" workbook (ONE wide rate sheet of 54 RTO clusters x ~111 labelled columns under merged headers, plus 'PV SAOD- HIGHEND Enabler', 'RTO Cluster Mapping' and the declined make / model lists) converted to SBI_UPLOAD_<yyyy-mm>.csv by sbi_grid_to_config.py. Load it with grid-converter whenever the user sends an SBI grid, an SBI rates_export to review, or asks about the premium slab (Pvt Car Slab / CV Slab), High-End models, declined clusters, the DL +2% note, EV declines, or what the SBI script got wrong.
---

# SBI General - broker grid (Agency, C & D)

Dedicated reference: `SBI_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` section 7.9, `CLAUDE.md` (the SBI bullet), P-46..P-50, F-39, F-41, `lessons_learned.md` (2026-11 sbi entries).

## Status
First build 2 Nov 2026; the **November 2026 output was ACCEPTED by the user on 3 Nov 2026** (cases `sbi_nov26` - 106,428 rows, answer key `rates_export (74).csv` - and `sbi_oct26`, layout snapshot `sbi_nov26`, `tests/sbi_tests.py` 108 checks).
- (11 Nov 2026) Structure shifts (11 Nov 2026): audited with 17 operators (renamed / reordered / extra / hidden sheets, rows and columns inserted, blank rows, header noise ...): identical output for all; a banner above the table is flagged in remarks. `Ctx.one` already stops when two columns match one header chain. No change was needed.

## Run it
```
.venv/Scripts/python.exe sbi_grid_to_config.py --grid "<grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--pc-slab "1L-3L"|none] [--cv-slab "3L-4.99L"|none] --no-prompt
```
Output `output\sbi\SBI_UPLOAD_<yyyy-mm>.csv` + `SBI_REVIEW_<yyyy-mm>.xlsx`. The premium slabs are ASKED every month (the user's rule, 1 Nov): without `--pc-slab` / `--cv-slab` the run stops with the question.

## How it reads the grid - NOTHING BY CELL ADDRESS (the user, 1 Nov: SBI changes this layout all the time)
`RateTable` turns the merged header of every column into a chain of texts; every block is found by header regexes (`GCV_BLOCKS`, `find_cols`, `Ctx.pair` / `Ctx.one`). A block that is missing, a labelled column no rule reads (even an empty one) and a cluster only one sheet knows are WARNED about and listed in `Not written`. Rules come from the TEXT of the notes (`read_rules`, `RULE_TABLE`, `collect_notes`): ages per product, non-NCB cuts, High-End non-NCB = 0, taxi / car SATP cuts, TW cc 500, fuel declines (`read_fuel_declines`), the DL uplift (`read_uplifts`), the non-doable AP RTOs, the slab notes, the validity dates. `take()` uses up the words a rule understood; a number left over is flagged. A rule whose note has gone is ASKED about (last month's value first, `LAST_KNOWN`), never silently kept; a vanished fuel decline / uplift only warns (`LAST_FUEL_DECLINES`, `LAST_UPLIFTS`).

## Rules the user has set (do not re-ask)
- Slab: 'Pvt Car Slab' (1L-3L -1 / 50K-1L -2 / 15K-50K -3) applies to **Private Car ONLY**; 'CV Slab' to everything else; **two-wheelers have NO slab** (3 Nov). An adjustment applies only to the family a note names - never by analogy (F-41). The `<15K` IRDA slab is not built. High-End rates are as printed (the enabler sheet's 20).
- Rows: Comp = 1+1 On Net, SATP = STP On TP, SAOD = On OD; blank / zero cell = a ZERO card; payout = payin - 7 floor 0; unchanged rows do NOT keep an old from_date.
- `In PCV 3W EV Decline` (any product, any fuel decline note) = no rows of that fuel. `additional 2% - DL Registration Vehical 2.5 to 3.5 Segment Only` is APPLIED: GCV 4W weight bands inside 2500.01-3500.01, Comp and SATP, the whole cluster `DL - NCR` (as the user did by hand).
- The 340 High-End 'MB Class' rows (no make group in the master) stay written and FLAGGED; two-wheeler rows ALWAYS carry `user_id` 2966; the School Bus 'Decline Cluster' list stays flagged only (no cross-check entry); STP rows have NCB / ZD `NA`.
- The `Declined cluster` notes are CROSS-CHECKED with the grid (P-48): a declined cluster must pay nothing and a cluster that pays nothing must be named; each gap is flagged on its rows only.

## Monthly procedure (on top of grid-converter section 2)
1. `grid_doctor.py check`; read each changed note. 2. Run; answer the two slab questions from the grid's own notes (October = -1 / -1). 3. **Print every note of the sheet with its fate** (rule / standing / flagged) - a note with no fate was dropped (F-39). 4. Compare with their export (`--ref`); explain every number: differences from their upload are listed in the case note (document wins, 25 Oct). 5. Spot-check 3-5 cells against the workbook (cell minus slab, payout - 7). 6. Review against their corrections if they send an export (grid-converter Job E).

## Traps seen
F-39 (a 20-character decline note dropped by the shared 25-char floor; SBI flags a whole unread note from 8 chars, `MIN_NOTE_CHARS`, and keeps 25 for the leftover of a read note), F-41 (slab extended by analogy), P-47 (text scanner), P-48 (declined clusters), P-49 (fuel declines), P-50 (uplift note, read as a WHOLE sentence). Scratch scripts: `df["product"]`, not `df.product` (pandas 3).

## Open items
The November PCV 3W block HEADER also says "(EV Decline)" - the script reads only the note (idea: read the header too). Their portal data oddities (TW row id 7 payin 30 / payout 0, their own MISCD `MH - Rest` group) are theirs, not the script's.
