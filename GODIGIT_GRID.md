# GO DIGIT GENERAL - the agency grid -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/godigit-grid/SKILL.md` (status, the user's rules, traps, open items - a dated log); the long record is `AGENT_CONTEXT.md` section 7.8, patterns P-39, P-56 .. P-58, P-65, failures F-47 .. F-52, F-59 and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a workbook is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. **Since 10 Nov 2026 nothing is found by a sheet's name or a column's place (P-56): every table is found by what its header row says** - Go Digit renames or redesigns its sheets every month - and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `godigit_agency_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | the agency workbook. October 2026: `Delhi_PAN_India_Agency.xls`, really an **encrypted .xlsx** (password-protected; `msoffcrypto-tool` decrypts it, `xlrd` is only for a genuine old .xls). November 2026: `Nov26_communication_grid.xlsx`, 13 visible sheets, NOT password-protected (the `.xls` sent with it was October's workbook again) |
| Password | `--password`, the environment variable `GRID_PASSWORD` (the web app sets it for the child process) or a hidden prompt - **never stored in a file or a document** |
| Output | `output\godigit\GODIGIT_UPLOAD_<yyyy-mm>.csv` (42 columns) + `GODIGIT_REVIEW_<yyyy-mm>.xlsx` |
| Dates | given on every run (October assumed 15/10/2026 - 31/03/2030 from the user's upload; November 01/11/2026 - 31/03/2030, confirmed by the user's acceptance) |
| Status | October: first build 28 Oct 2026, never formally accepted (then 53,210 rows, 46,854 keys equal with the same rates against the user's October upload of 47,532 rows; today's October output is 53,212 rows). **November: ACCEPTED by the user on 10 Nov 2026**; the file was replaced by the automatic-reading version (v2: 134,179 rows, 436 flagged) on 17 Nov 2026; a v3 with the wording readings (134,310 rows, 828 flagged) is NOT blessed |
| Regression | `godigit_oct26` (the workbook saved WITHOUT the password: the case needs neither), `godigit_nov26` (answers `''`, `1`, `y`; no `--ref`); layout snapshots `godigit_oct26`, `godigit_nov26` |
| Tests | `tests/godigit_tests.py` (158 checks: every put-back position reader or silent guess is caught), `tests/godigit_cd1_tests.py` (55), `tests/godigit_wording_tests.py` (44), `tests/godigit_d_tests.py` (17) |

```
.venv/Scripts/python.exe godigit_agency_to_config.py --grid "<workbook>" [--password "<password>"] --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their export>.csv" [--no-prompt]
```
Run WITHOUT `--no-prompt` to answer the questions (the first prompt asks for the earlier upload: Enter skips it); `--no-prompt` stops at the first question and prints it.

## 2. The workbook (found by what its headers say)

* **Only VISIBLE sheets are read** (in October 14 of 26 were hidden older versions). A table is found by what its header row says: `Spec` + `find_table` (aliases, typo-tolerant; the sheet's name only breaks a tie), every column by its own header inside its block (`block_ranges`), every group by the titles above it (`title_texts`, `classify_title`). A table that is not there STOPS the run with `expected a header row that says ...; the closest sheet is ...; visible sheets ...` unless the person says to write the file without it (`MISSING_TABLES`). Columns nobody reads are warned about (`warn_unread`); the sheet's own age-band table is checked against `AGES`.
* **The tables** (October names in brackets; November renamed all of them): the CV grid (`LATEST CV COMBINED GRID`: GCV 4W / 3W, taxi, staff bus, PCV 2W / 3W, MISCD, tractors; **CD1 = the OD-discount band, CD2 by vehicle age** for comprehensive and SATP); the Two-Wheeler 1+1 and SATP table; the Two-Wheeler 1+5 table (only the block whose title says `1+5`; two blocks sit side by side); the Two-Wheeler SAOD table; the Private Car comprehensive + SAOD table (with a volume slab); the Private Car TP table (it appears twice, and its sheet name has a leading space in October); the remarks sheet (`CV Remarks`: Volvo / Scania blocked; Truck Mahindra + Bharat Benz at 70% CD1); the RTO mapping sheets (only to put RTO codes next to a group the master lacks).
* **A new-business block can sit beside the car table with its OWN cluster column in another row order**: a row's cluster is the one of the volume-slab block (F-49).
* **Two payout columns** (`Max CD2` + `Avg CD2`): the user ANSWERED `Avg CD2` on 10 Nov 2026 (`PAYOUT_MEASURE = "avg"`, applied without asking and said once in the run's decisions; `None` = a numbered question). October's figures are the ones now under `Avg CD2`; `Max CD2` is higher. A table with ONE payout column (the CV grid has only `Max CD2`, about 5 points above October) uses it as it is.
* **The car volume slab**: October's `1-3 Lac` (`PC_VOLUME_SLAB`); November's sheet has only `> 5 Lacs` - the user chose it for that month (1 to 2 points above `1-3 Lac` in October's data); a slab the sheet lost is a question.

## 3. How the workbook becomes rows

* **Comprehensive = 1+1 `On Net`, SATP = STP `On TP`** (tariff 0-100.01, NCB / ZD `NA`); ages New 0-1 (add_tnc `New`), 1 Yr 1-2 ... 6-10 Yr, 11+ Yr. Cells are fractions (0.325 = 32.5). A SATP payout no longer split by CPA is one row per entry with `is_cpa` `NA`.
* **CD1 -> `tariff_max = cap + 0.01`**. Its text forms are many (`0-60%`, `New/1yr: D | 2-5yr: 80% | 6yr+: 85%/90%`, `With Addon: 85%/90% Without Addon: 60%` -> `is_zd` YES / NO ...); a NEGATIVE CD1 = declined (a zero card at 100.01); formula noise in a CD1 is written as the whole percent and flagged. **Discount limits are written as the sheet SHOWS them** (`DISCOUNT_ROUNDING = "sheet"`: a cell formatted as a whole percent is the nearest whole percent; a two-decimal cell is written as it is, so 67.5 stays 67.5).
* `D` / `#VALUE!` = a zero card; blank = a placeholder; **`MISP` = `On OD` 17.5** (not in the grid: the user's October upload value, accepted again for November - a standing decision, said once in the warnings, no flag per row). The staff-bus `up to 10-seater` note: not written (the table starts above 10 seats), one line in `Not written`.
* **A `D` placeholder never beats a real rate** (the user, 18 Nov 2026, `drop_d_placeholders`): a `D` (or a CD1 text that declines an age) in a row that covers only some ages / makes / discount bands means "not on this row"; it is dropped when another row gives a rate for the SAME ENTRY at the same or a higher cap, or when its own band starts below that cap; a `D` ABOVE every real band stays (an explicit decline); an entry that is `D` everywhere keeps its zero card; a NUMBER 0% is a real rate and follows the lowest-rate rule. Every ignored entry is listed in the review sheet `D ignored`.
* **Private Car**: comprehensive / SAOD at the deal slab, four non-electric fuels; a car third-party note that states its own cc and age slice is applied to the row it sits on (no card for that slice, as in the user's October upload); `Few RTOs ... declined` stays flagged.

## 4. Automatic reading (the user, 10 and 18 Nov 2026: "100% automation with absolutely zero human intervention ... nothing about a condition may be silent")

* A new wording is READ by a grammar - `cd1_read` (age groups + a cap each, every character a token; an unknown word, a conflict or a split column is a REFUSAL; the closed range keeps an age an open one also names; a group the text leaves out WITH a payout is written for the whole range and flagged by name), `taxi_make_parts` (taxi make words with age windows and a cap per make: `TAXI_ALL_EX_<X>` + `MAKE_DEFINITIONS`), `tp_decline_slice`, `gvw_of` / `weight_cell`, text-typed fractions - never just flagged; **what a grammar does not understand is refused and keeps the old flag, never guessed**. Every automatic reading is listed in the review sheet `CD1 readings` (one line per wording).
* **`read_wordings`** reads every make / body / segment wording of the CV grid by an exact table or refusing pattern (`GOODS_MAKE_READINGS`, `HCV_MAKE_READINGS`, `PCV23_MAKE_READINGS`, `TAXI_SEGMENT`, `cpa_of`; CNG in a text = CNG + LPG; taxi `All` = all five fuels; a PCV 2W row = PCV 2W only and `without ZD` = `is_zd` NO; the PCV 3W rows of a cluster that has its own PCV 2W row = PCV 3W only) and FLAGS what no table reads on exactly its rows (`sp["unread"]` -> remarks). **A wording that falls through to a default (all makes, all fuels, no CPA) without a flag is a bug.** New wordings go into the tables with a test in `tests/godigit_wording_tests.py`.
* **Review sheets**: `Wording readings` (every wording, the ones not applied first), `Unread parts` (`audit_unread`: whole sheets and columns no reader used, with the reason), `Notes read` (`notes_table`: every sentence of the remarks sheet and what became of it), `CD1 readings`, `D ignored`, `Lowest rate chosen`, `Make groups to create`. `settle_duplicates` carries the flags of a dropped duplicate onto the row that stays. A long `Lowest rate chosen` list can mean UNREAD WORDINGS: look at the wordings first.
* **Five wordings nothing can apply** (no field in the upload; flagged): Tractor `Electric and Low HP` and the four E-Rickshaw TPPD-reduction bodies. Not converted and listed (never silent): the Two-Wheeler 5+5 sheet, the car sheet's new-business blocks (1+3 `90:10` by make, 3+3, HEV; one block each at the left and right edge of the car sheet) and their notes, School Bus, `CV PCV 6+Seating`, the Two-Wheeler SAOD columns `CD2` / `Status No Break in` / `Status with Break in`.

## 5. Names

* **RTO**: `new_rto_list` carries the bare cluster by default. With `GRID_RTO_NAMES=group` the portal's group `GO-DIGIT_OCT26_<segment>_<cluster>` (Two-Wheeler `GO-DIGIT_OCT26_2W_1+1_TP_` / `_2W_SAOD_`, cars `GO-DIGIT_PC_1+1_OCT26_` / `_PC_SAOD_OCT26_` / `_PC_TP_OCT26_`), resolved to the master's spelling; a missing group is flagged and listed with its RTO codes (18 in October). Car TP clusters are named like the 4W TP mapping (`All_India_Decline` is a real cluster of 80 RTOs).
* **The person's RTO working list** (the user, 31 Oct 2026): `WorkingCol` finds the person's `RTO` column per table by header (six tables). Put the column at the **RIGHT EDGE** of side-by-side blocks: an RTO column between a block's first and last column would take the place of one of the table's own columns, so it is refused with a warning. The hidden `_grp` column keeps the group token because the remarks sheet is scoped by it - a working name must not lose those notes.
* **Make groups**: a group the master lacks is written by name and defined in `MAKE_DEFINITIONS`, and listed in `Make groups to create` (on 18 Nov 2026 five: `GCV_4W_ALL_EX_MAHINDRA`, `PCV_3W_BAJAJ_TVS`, `PCV_3W_BAJAJ_RE`, `PCV_2W_ALL_EX_BAJAJ`, `TAXI_ALL_EX_HINDUSTAN`; the checklist of that day left only `TAXI_ALL_EX_HINDUSTAN` for the portal side).

## 6. Edge cases and what the script does

| Situation | What happens |
|---|---|
| The file is called `.xls` but holds an encrypted `.xlsx` | the bytes are sniffed, then decrypted |
| The `.xls` the user sends is LAST month's workbook | compare its content with the regression input before deciding which month it is |
| A sheet renamed / moved / hidden, columns inserted | found by header words; hidden sheets are never read |
| Two payout columns | the user's answer (`Avg CD2`) is applied without asking |
| A volume slab the sheet lost | a numbered question |
| A table that is not there | stops with expected / found unless the person says to write the file without it |
| An unknown CD1 / make / body wording | refused by its grammar and flagged on exactly its rows, listed in the review sheets |
| `Max CD2` and `Avg CD2` present, one table only has `Max` | each table uses what it has |
| An RTO column inside a block | refused with a warning |
| A wording that names a fuel / make / CPA the code did not know | a bug unless flagged (F-59) |

## 7. Results

* **October** (the first build): 53,212 rows; rates identical in every later version; 90 flagged (was 10: the Tractor `Electric and Low HP` wording).
* **November**: the accepted file of 10 Nov (133,446 rows, 16,690 flagged) -> v2 after the automation pass (134,179 rows, 436 flagged; the user's `Replace with v2` of 17 Nov made it the file in `output\godigit\`; the accepted original is in `backups\godigit_accepted_20261110`) -> v3 with the wording readings (134,310 rows, 828 flagged after the `D` rule). `godigit_oct26` / `godigit_nov26` were BLESSED on 11 Nov for v2; the v3 changes are not blessed yet.

## 8. Checking a run

1. FIRST find out which month a file is: compare its content with `regression/inputs/godigit_oct26.xlsx` (decrypt into a scratch folder; the password goes only into `GRID_PASSWORD` / `--password` at run time, never into a file or a stored message).
2. `python grid_doctor.py check` on the decrypted copy if needed; run with `--ref`; read the RTO groups the master lacks in the review workbook; read the remarks sheet's text changes; report what was assumed (dates, `MISP`).
3. **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a cluster by its label cell, a segment by its header chain and a vehicle age by its band heading, read the crossing cell with openpyxl by searching for those texts (never by typing an address), then check `tariff_max = cap + 0.01` and the payin on the matching rows.
4. Look at `Lowest rate chosen`, `Wording readings`, `Unread parts`, `D ignored` BEFORE doubting a rate.
5. `python check_regressions.py --only godigit_oct26` and `--only godigit_nov26`.

## 9. Decisions (the user) - do not re-ask

1. Only visible sheets. Discount limits as the sheet shows them. The reader NEVER guesses a new layout: a fundamental break stops with `expected ... found ...` and the 42-column CSV stays (10 Nov 2026). The JSON configuration-engine idea is NOT approved or built.
2. The payin is `Avg CD2` wherever a table gives both (10 Nov 2026).
3. Private Car slab `1-3 Lac` (`PC_VOLUME_SLAB`); for November 2026 the sheet's only slab `> 5 Lacs` (a choice for that month).
4. 100% automation, no mistakes (10 Nov 2026): new wordings are READ by grammars; every automatic reading is listed; a decision that changes accepted rows is reported and waits for the user's OK before a baseline is blessed.
5. `MISP` = `On OD` 17.5 is a standing decision. A car TP note with its own slice is applied to its row.
6. Nothing about a condition may be silent (18 Nov 2026): read by a table or pattern, or FLAGGED on exactly its rows and listed.
7. A `D` never beats a real rate of another row for the same card (18 Nov 2026, `drop_d_placeholders`).
8. The RTO working-list column goes at the right edge of side-by-side blocks.

## 10. Traps seen

The file is called `.xls` but holds an encrypted `.xlsx`; older versions hidden among the sheets; two side-by-side blocks on the 1+5 sheet; a leading space in a sheet name; CD1 has many text forms and formula noise; sheet names change every month; F-50 (`6-10yr` + `10yr+` name the same age: the closed range keeps it), F-51 (a CD1 that leaves out New / 1 Yr while the payout cells hold a rate: those groups are NOT declined, they get the whole range and a flag by name), F-52 (`age 1 - 0-60%` is the age 1 with the cap 0-60%, not the range 1-0); a comprehensive cell can be text (`0.4`, `2.5%; TATA With Addon: ...`, `40 to 43T : 2.5% 43T+: D`); the lowest-rate rule HID unread wordings (F-59); in `build()` the RTO names are the portal's group names - the cluster names appear when the file is written (tests that look rows up by cluster use `endswith`).

## 11. Open

What is left for a person on the November file is one checklist (`output\godigit\GODIGIT_CHECKLIST_2026-11.xlsx`). **Portal side**: the make group `TAXI_ALL_EX_HINDUSTAN` and the RTO groups for the cluster names. **Ask Go Digit**: the TATA `-100%` wording, the grid line whose CD1 reads `2-5yr: D | 6yr+: 90%` (New / 1 Yr come out as 100 `no limit named`, flagged), the `Few RTOs ... declined` rows, the Two-Wheeler SAOD columns (do the STP rows pay CD2 22.5% instead of the written `Max 1 year SAOD CD2`, which is higher on 83 rows by up to 32.5 points?), the Hero 1+5 row `HERO MOTOCORP / All` = 0 (not written; 260 `Others` rows end in `,HERO`). In the grid but not in the CSV: School Bus (107 rows), `CV PCV 6+Seating` (25), the car new-business blocks and the 5+5 sheet, the staff-bus note. The E-Rickshaw TPPD reduction: four rates in the grid, CPA YES / NO is written, the TPPD variants share one card (lowest). Not tested: the Colab path and the hidden password prompt.

## 12. Files

`godigit_agency_to_config.py` - `tests/godigit_tests.py`, `godigit_cd1_tests.py`, `godigit_wording_tests.py`, `godigit_d_tests.py`, `tests/structure_shift_tests.py` (subject `godigit`), `tests/unread_audit.py` - `regression/inputs/godigit_oct26.xlsx`, `godigit_nov26.xlsx`, `regression/references/godigit_oct26_rates_export.csv`, cases `godigit_oct26` / `godigit_nov26` - `layouts/godigit_oct26.json`, `layouts/godigit_nov26.json` - `.claude/skills/godigit-grid/SKILL.md` - `webapp/src/registry.js` (secret `GRID_PASSWORD`) - `.vscode/launch.json` / `tasks.json`.
