# SBI GENERAL - the broker grid (Agency, C & D) -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/sbi-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` section 7.9, patterns P-46 .. P-50, failures F-39 and F-41, and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header chain, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. SBI changes this layout all the time (the user, 1 Nov 2026), so the script finds every table and every column by what its headers SAY, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `sbi_grid_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | ONE workbook, the "Agent, C & D Cat. Broker Grid" (.xlsx): ONE wide rate sheet (54 RTO clusters x about 111 labelled columns under a merged multi-row header) + `PV SAOD- HIGHEND Enabler`, `RTO Cluster Mapping` and the two declined make / model lists. A hidden `Sheet1` (an old version) is never read |
| Output | `output\sbi\SBI_UPLOAD_<yyyy-mm>.csv` (42 columns) + `SBI_REVIEW_<yyyy-mm>.xlsx` |
| Dates | given on every run (`--from-date`, `--to-date`) |
| Questions | the two premium slabs are ASKED every month (`--pc-slab`, `--cv-slab` answer them); without them a run with `--no-prompt` stops with the question |
| Status | first build 2 Nov 2026; the **November 2026 output was ACCEPTED by the user on 3 Nov 2026** |
| Answer keys | November: the user's corrected portal export (`rates_export (74).csv`, case `sbi_nov26`, 106,428 rows); October: their October upload (`rates_export (69).csv`, 102,604 rows, case `sbi_oct26`); layout snapshot `sbi_nov26` |
| Tests | `tests/sbi_tests.py` (108 checks: the scanner, reworded and missing notes, the slabs, the fuel declines, the DL uplift, every block against cells of the workbook, a shifted + renamed + extended layout, the real grid against the upload) |

```
.venv/Scripts/python.exe sbi_grid_to_config.py --grid "<grid>.xlsx" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--pc-slab "1L-3L"|none] [--cv-slab "3L-4.99L"|none] --no-prompt
```

## 2. The grid (found by what its headers say)

* **`RateTable`** turns the merged header of every column into a CHAIN of its distinct texts (the top text ... the text above the leaf ... the leaf). Empty rows above the header, a merged cell repeated over several rows, or one more header row change nothing. Every block of columns is found by header regexes (`GCV_BLOCKS`, `find_cols`, `Ctx.pair` / `Ctx.one`).
* **Missing and unread parts are WARNED about, never guessed**: a block that is missing, a labelled column that no rule reads (even an empty one) and a cluster that only one of the sheets knows are warned about and listed in the sheet `Not written`. `Ctx.one` STOPS when two columns match one header chain.
* **The other sheets** are found by their headers too (`find_tables`): the High-End enabler sheet (make / model rows with NCB columns), the cluster mapping, the declined make / model lists. The declined lists are fingerprinted (`DECLINE_FINGERPRINT`): a change is warned about, because the portal's make groups then need a look.

## 3. The text scanner: rules come from the notes

`read_rules` + `RULE_TABLE` (P-47): each rule of the notes is found by its words (a label such as `Pvt Car`, `PCV Taxi`, `Tw` + keywords) and its number is taken from the note; `take()` USES UP the words a rule understood and any number left over is flagged. What the scanner reads:

* vehicle age limits per product (comprehensive / SAOD and SATP); Private Car non-NCB cut; High-End non-NCB; taxi non-NCB cut; car SATP cut by age; Two-Wheeler cc limit; the Two-Wheeler make list; the non-doable AP RTOs; the two slab notes; the grid's validity dates;
* **fuel declines** (`read_fuel_declines`, P-49): any note that declines a fuel for any product (`In PCV 3W EV Decline`, `Electric declined for all GCV` ...), clause by clause and by full match only - `Mahindra EV declined` or a cluster list in brackets is never read; `Ctx.add` leaves the fuel out, and a decline that has gone from the new grid raises a warning (`LAST_FUEL_DECLINES`);
* **uplifts** (`read_uplifts`, P-50): the WHOLE sentence is read - a qualifier after it is not guessed: the note stays flagged (`LAST_UPLIFTS` warns when the note has gone);
* **declined clusters** (`read_declines`, P-48): the `Declined cluster` notes are CROSS-CHECKED with the grid - a cluster the note declines must pay nothing, and a cluster that pays nothing must be named; each gap is flagged on its rows only (October had 17 such gaps: taxi 7, car SATP 7, PCV 3W, GCV 3W, tractor).

**A rule whose note has gone is ASKED about** (last month's value first, `LAST_KNOWN`), never silently kept; a slab note that cannot be read stops with a question. What is neither understood nor in `STANDING_NOTES` is flagged `Manual Review Required` on the rows it may affect (`note_scope`: the segments and states the text names - weight bands, state codes - else every row). A whole unread note is flagged from 8 characters (`MIN_NOTE_CHARS`); the 25-character floor of the shared `unknown_notes` is kept only for the leftover of a note a rule has read (F-39). Examples flagged in October, none of them applied (the RTOs they name are not in the portal's cluster groups): the TN-C RTOs on 12-40 t, `Rest of TN = Chennai II for Tata Ace`, UP78 / HR68, the J&K 12-20 t SATP rates, the WB-K RTOs for taxis, models 3718 / 3723.

## 4. How the grid becomes rows

| Part | Rows |
|---|---|
| Policy / type | Comp = 1+1 `On Net`; SATP = STP `On TP`; SAOD and Two-Wheeler SAOD = `On OD`; `Renew`; payout = payin - 7, floor 0 |
| A blank or zero cell | a ZERO card (the October upload has one per cluster) |
| **Premium slab** | payin = cell - the slab. `Pvt Car Slab` (1L-3L -1, 50K-1L -2, 15K-50K -3) applies to **Private Car ONLY**; `CV Slab` (3L-4.99L -1, 1L-2.99L -2, 15K-99K -3) to everything else; **two-wheelers have NO slab** (the user, 3 Nov 2026, F-41); the `<15K` IRDA slab is not built; the High-End rates are as printed. October = -1 and -1 |
| GCV 4W | five weight classes with the same payout, four non-electric fuels; make groups `SBI_Below_3.5T_GCV4W_MAKES` (up to 2.0 t), `TATA` (2.0-3.5 t), `..._EX_TATA` (2.0-2.5 t other makes), `Mahindra`, `Ashok`, `TATA,Ashok` (12 t and above: "allowed make are Tata, AL & Mahindra only"), `GCV_4W_ALL` zero cards for the declined 3.5-12 t; the tractor column becomes class `Goods Carrying Tractor` |
| Private Car (not High-End) | NCB yes = cell - slab; NCB no = the note's cut below that; age to 15.01 |
| Private Car **High-End** | the enabler sheet, rates as printed (20): make group `SBI_MAY26_<make>_<model>`; NCB yes = 20 and **NCB no = 0** (the note); `Petrol / Hybrid` = Petrol + Electric, `Diesel / CNG / LPG` = those three |
| Car SATP | cc bands from the labels (contiguous, 0-1000.01 ...); ages 10+ = the cell, ages 1-10 = the note's cut, **below 1 year = ZERO** |
| Taxi | three cc bands x NCB yes / no (the note's cut) x nil-dep yes / no; SATP to age 20.01 |
| School bus | by seats (12-17 = 12.01-18.01, 18+ = 18.01-999) |
| Two-Wheeler | bike / scooter (scooters without ROYAL), cc 0-500.01, `add_tnc` = the races exclusion |
| PCV 3W, Tractor / Harvester | PCV 3W new / old; Tractor / Harvester as MISCD |
| **AP - Rest** in the 2.5-3.5 t segment | the cluster minus the six non-doable RTOs: the portal group that holds exactly those RTOs (`SBI_OCT26_AP_-_Rest_2.5_3.5T`); in cluster mode the cell says it in plain English (`rto_condition` + `SPECIAL_GROUPS`) |
| **The DL +2% note** (P-50) | applied ONLY to GCV 4W weight bands inside 2500.01-3500.01 (the user's scope, 3 Nov 2026), Comp and SATP, for the whole `DL - NCR` cluster (it also holds HR26 / 38 / 51 / 55 / 72 / 87 / 98 and UP14 / 16 / 18; the whole cluster gets the points, as the user did by hand), added after the CV slab; a blank cell stays a zero card |

## 5. A row

| Column | Value |
|---|---|
| insurance_company / insurer_vertical | `SBI GENERAL INSURANCE COMPANY LIMITED` / `Not Applicable` |
| policy_type | `Renew` |
| user_id | `2966` - Two-Wheeler rows ALWAYS carry 2966 (the user, 3 Nov 2026) |
| new_rto_list | the cluster as the grid writes it (`MH - M`; the cluster rule of 30 Oct 2026); with `GRID_RTO_NAMES=group` the portal's group `SBI_OCT26_<cluster, spaces -> _>` (`GROUP_PREFIX`) |
| STP rows | `is_ncb` and `is_zd` always `NA` |
| Unchanged rows | do NOT keep an old `from_date` |

## 6. Edge cases and what the script does

| Situation | What happens |
|---|---|
| A header row added / a merged cell repeated / empty rows above | the chain is the same: identical output (audited with 17 structure operators, 11 Nov 2026) |
| A block of columns missing | warned + `Not written`; never guessed |
| A labelled column no rule reads | warned + `Not written`, even when empty |
| Two columns match one header chain | stops (`Ctx.one`) |
| A rule's note is gone | ASKED, last month's value first (`LAST_KNOWN`); a vanished fuel decline / uplift only warns |
| A slab note that cannot be read | stops with a question |
| A note neither read nor standing | flagged on the rows it may affect |
| A fuel decline note for any product | no rows of that fuel |
| A cluster named as declined that pays, or one that pays nothing and is not named | flagged on its rows (cross-check, P-48) |
| A banner above the table | flagged in `remarks`; the rates are the same |
| `--no-prompt` and a slab question | `STOPPED - <the question>` |

## 7. Differences from the user's own uploads (the document wins; listed in the case notes)

* **From the October upload** (the first answer key; 94,176 keys equal with not one payin or payout differing until 3 Nov): GCV 2.5-3.5 t SATP above 5 years goes to age 25.01 like every other SATP row (the upload had 20.01; 4,320 rows); the High-End rows (3,298 with NCB `NA`) become NCB yes / no with the sheet's fuels (6,528, of which 340 are `MB Class` - no make group in the master, written and flagged); car SATP 0-10 years (810) becomes 0-1 (zero) + 1-10 (1,620).
* **Corrections of 3 Nov 2026** (the November upload): (1) the grid says `In PCV 3W EV Decline`, so the 216 PCV 3W Electric rows must not exist - the only rule was written for one product, `read_fuel_declines` now reads any fuel decline for any product (P-49); that note had also been dropped in silence by the shared 25-character floor (F-39). (2) Two-wheelers have NO slab (F-41). (3) The DL +2% (section 4), which the user had added by hand.
* November: 106,428 rows, 10,540 flagged, equal to the user's corrected export except what is theirs: the 340 `MB Class` rows, their `user_id` 7436 on the Two-Wheeler rows (and one Two-Wheeler row of theirs: AP - Rest, bike, 1+1, Petrol, payin 30 with payout 0), their own MISCD `MH - Rest` group. Regression `sbi_oct26` has `changed: 648` since then (248 Two-Wheeler rows +1 and 400 `DL - NCR` rows +2; the other 80 rows of the DL band have no key in the October upload) and 9,792 flagged.

## 8. Checking a run

1. `python grid_doctor.py check "<grid>"`; read each changed note.
2. Run; answer the two slab questions from the grid's own notes (October = -1 / -1).
3. **Print every note of the sheet with its fate** (rule / standing / flagged): a note with no fate was dropped (F-39).
4. Compare with their export (`--ref`) and explain every number; the document wins (25 Oct 2026).
5. **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a cluster by its label cell and a column by its header chain, read the crossing cell with openpyxl by searching for those texts (never by typing an address), subtract the slab, compare with the matching upload row (payout = payin - 7).
6. If the user sends their corrected export, review against it (`grid-converter` Job E).

## 9. Decisions (the user) - do not re-ask

1. Slab: `Pvt Car Slab` = Private Car only; `CV Slab` = everything else; two-wheelers have none (3 Nov). An adjustment applies only to the family a note names - never by analogy (F-41).
2. Comp = 1+1 On Net, SATP = STP On TP, SAOD = On OD; a blank / zero cell is a ZERO card; payout = payin - 7 floor 0; STP rows `NA` for NCB / ZD.
3. `In PCV 3W EV Decline` and any fuel-decline note = no rows of that fuel; the DL +2% is APPLIED within the scope of section 4.
4. The 340 High-End `MB Class` rows stay written and FLAGGED; two-wheeler rows always carry `user_id` 2966; the School Bus `Decline Cluster` list stays flagged only (no cross-check entry).
5. The short-note floor of the OTHER scripts is for the Shriram session ("keep focus on SBI").

## 10. Traps seen

F-39 (a 20-character decline note dropped by the shared floor), F-41 (the slab extended by analogy), P-47 (the text scanner), P-48 (declined clusters), P-49 (fuel declines), P-50 (the uplift note read as a WHOLE sentence). In scratch scripts write `df["product"]`, not `df.product` (pandas 3).

## 11. Open

The November PCV 3W block HEADER also says "(EV Decline)": the script reads only the note (idea: read the header too). Their portal data oddities (the Two-Wheeler row with payin 30 and payout 0, their own MISCD `MH - Rest` group) are theirs, not the script's.

## 12. Files

`sbi_grid_to_config.py` - `tests/sbi_tests.py`, `tests/structure_shift_tests.py` (subject `sbi`) - `regression/inputs/sbi_nov26.xlsx`, `sbi_oct26.xlsx`, `regression/references/sbi_nov26_rates_export.csv`, `sbi_oct26_rates_export.csv`, cases `sbi_nov26` / `sbi_oct26` - `layouts/sbi_nov26.json` - `.claude/skills/sbi-grid/SKILL.md` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
