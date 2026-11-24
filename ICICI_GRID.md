# ICICI LOMBARD - four workbooks (CV, Private Car, Two-Wheeler, new Two-Wheeler) -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/icici-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` section 7.10, patterns P-43 and P-53, failures F-44 .. F-46, and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. Nothing is found by a cell address or an exact sheet name (P-43): every workbook is recognised by what its sheets SAY, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `icici_grid_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | up to FOUR workbooks, each recognised by `classify_workbook` from what its sheets say: `--cv-grid`, `--pc-grid`, `--tw-grid`, `--tw-new-grid` (or `--grid` once per file). Any subset works; a run with only some adds `_CV` / `_PC_TW...` to the file name. The web app offers four OPTIONAL inputs (the server refuses a run with none) |
| Output | `output\icici\ICICI_UPLOAD_<yyyy-mm>.csv` (42 columns) + `ICICI_REVIEW_<yyyy-mm>.xlsx` |
| Dates | per workbook: `--from-date` / `--to-date`, or `--cv-from-date`, `--pc-from-date`, `--tw-from-date`, `--tw-new-from-date` |
| Status | first build 4 Nov 2026, **NOT yet accepted by the user**; they confirmed the assumptions of section 9 the same day |
| Answer keys | the user's November Private Car + Two-Wheeler upload (reproduced in full); their October CV upload (`icici_cv_oct26`); the November CV upload was not made |
| Regression | `icici_cv_oct26`, `icici_pc_tw_nov26`, `icici_cv_nov26`; layout snapshots `icici_cv_oct26`, `icici_cv_nov26`, `icici_pc_nov26`, `icici_tw_nov26`, `icici_tw_new_nov26` |
| Tests | `tests/icici_tests.py` (105 checks, in the quick suite; nine put-back bugs are caught); the web app: `webapp\test\icici.e2e.test.js` (the download = the command line's bytes) |

```
.venv/Scripts/python.exe icici_grid_to_config.py [--cv-grid cv.xlsx] [--pc-grid pc.xlsx] [--tw-grid tw.xlsx] [--tw-new-grid twnew.xlsx] --from-date DD/MM/YYYY --to-date DD/MM/YYYY [--vertical ..] --ref "<their export>.csv" --no-prompt
```

## 2. The CV grid

* **Sheets and columns by words**: the CV matrices are the sheets whose header row holds at least six known segments (`find_cv_matrices`); the **cluster column** is the one whose header says cluster / segment AND whose values are text - NEVER the person's own `RTO` column (F-45). A COMP matrix and an AOTP matrix; one row per cluster x segment (`classify_segment`: GCV 3W, SCV, LCV, MHCV by tonnage and body, MISCD, tractor, PCV 3W, school + staff bus, taxi, PCV 2W; tonnages and seats are read from the header).
* **Rows**: COMP = 1+1 `On Net`; AOTP / TP = STP `On TP` (a segment that is not in the AOTP matrix, or a cell without an `AOTP:` clause, takes the comprehensive cell: CV NOTES rule 9); payout = payin - 7; `No Biz`, 0% or blank = a ZERO card.
* **A cell is a small language** (P-53), e.g. `OLD:10%,NEW:0%`, `TATA|M&M:55%,OTHERS:35%`, `COMP|NEW|NZD:25%`: `parse_cell` gives clauses of typed words (clauses by commas, words by `|`, the rate after the colon); `regions` makes boxes over policy / zero-dep / places / body / makes / age. `OTHERS` = the complement of the named clauses; a 0% clause is a decline that is also taken out of the other clauses; `a-b yrs` = (a-1)+0.01 to b+0.01; `NEW` = age 0-1 (T&C `New`), `OLD` = 1-99 (T&C `Old`; the SCV segments keep their tag, `tag_always`).
* **Make groups** are composed from the make master BY CONTENT (`MakeNames`, P-54: `GCV_4W_ALL_EX_BHARAT_BENZ`, `TaxiALL_EX_EECO_OMNI_MAGIC` ...); vehicle-type words (cranes, ACE, tankers ...) go to `add_tnc`; **a place is accepted only when the workbook's own lists know it** (RTO MASTER, Negative list, Council Call, RTO CODES).
* **An unknown word, a make set the master cannot express, an Excel error or an unread column** -> the cell or column is listed in the sheet `Not written`, never guessed.
* **Council Call** (`CC`): a cell with no fixed rate is NEVER turned into a rate; its active entries are listed beside it in the sheet `Council call cells` (`CC,OTHERS:n%` writes the OTHERS rate, flagged).
* The Negative RTO list and the CV NOTES the script cannot apply are flagged on the rows they name (`note_mask`); the T&C texts `TNC_*` are the user's wording of CV NOTES 4, 13 and 14. Bharat Benz is left out of the default make groups because the CV NOTES say it is declined (read from the note; asked about when the note has gone).

## 3. Private Car, Two-Wheeler and the new Two-Wheeler

| Workbook | Rows |
|---|---|
| **Private Car** (one sheet, two header rows) | New = 1+3 (all five fuels, T&C `Applicable for 3+3`); SAOD and 1+1 with NCB per fuel (LPG takes the CNG column) and without NCB one column for all fuels; Act only = STP `On TP`. **1+1 is `On OD`** (`PI_TYPE_PC`, F-44). `Used car` and `Broker` are listed, not written |
| **Two-Wheeler** | Bike / Scooter x Petrol / Electric: 1+1 with the T&C of the column title, SAOD, STP |
| **New Two-Wheeler** (one sheet per OEM) | HERO, HONDA, TVS, Suzuki, Yamaha, ROYAL (bikes only), EV = Electric `TW_ALL`; 1+5 / New / `On Net`, age 0-99. The payout column is the first numeric one right of the `Bike/ Scooter` header; its title names a month and is only quoted. The OEM is read from the WORDS of the sheet name (`tw_oem_key`: `HMC v2`, `Hero (Sept)`) |

## 4. A row

| Column | Value |
|---|---|
| insurance_company | `ICICI LOMBARD GENERAL INSURANCE CO LTD` |
| insurer_vertical | `ICICI-RB-Offline` on EVERY row, also Private Car and Two-Wheeler (the user, 4 Nov 2026; `--vertical` overrides) |
| policy_type | `Renew`; `New` for the 1+3 and the 1+5 rows |
| rates | written to **8 decimals** (`RATE_DECIMALS`, the user's own precision) |
| STP rows | `is_ncb` and `is_zd` always `NA` |
| user_id | `2966` |
| new_rto_list | the grid's cluster name (the cluster rule of 30 Oct 2026); `GRID_RTO_NAMES=group` writes the portal groups of the user's uploads (`ICICI_SEPT26_CV_`, `ICICI-RB_PC_JUL26_`, `ICICI-RB_JUL26_`) |

## 5. Notes

The Negative RTO list, the CV NOTES the script cannot apply and a Council Call cell are flagged / listed as above. A note the person resolved in the web app is never flagged again. A text the script cannot apply is `Manual Review Required: [text]` on the rows it may affect - never guessed, never a stop.

## 6. Edge cases and what the script does

| Situation | What happens |
|---|---|
| A sheet renamed or moved, rows / columns inserted | recognised by its words: the same upload |
| Blank rows inside the CV matrix / Private Car / Two-Wheeler sheets | skipped; a TEXT row ends the table (20,317 CV rows used to become 11,264 before this was fixed) |
| The reference sheets (RTO list, Council Call) start lower down | the scan window that looks for them was widened (11 Nov 2026) so a shifted layout still finds them |
| A payout heading that appears TWICE | both columns are read; the lowest-rate rule keeps the LOWEST payin and lists it in `Lowest rate chosen` (the user, 12 Nov 2026); a blank twin cell is no payin (`twin_groups` / `twin_blank`) |
| Two payout columns with DIFFERENT headings on a new Two-Wheeler sheet | stops (which month?) |
| Two sheets that say one OEM | stops, unless one is called exactly the OEM; a workbook of which no sheet could be read stops |
| Every sheet called `Sheet` + a number (CV and new Two-Wheeler workbooks) | stops, by design: the notes sheet / the OEM is the name |
| `OTHERS:0%` | no zero card is written (listed in `Not written`) |
| A make set with no group in the master (e.g. EECO + MAGIC + OMNI) | the cell is not written, listed in `Not written`; the user creates the group if they want it |
| The user's running web server | reads `registry.js` only when it starts: restart it to list a new script (F-46) |

## 7. Results and differences (4 Nov 2026)

* **Private Car + Two-Wheeler + new Two-Wheeler** = 3,342 rows; EVERY one of the 3,318 rows of the user's November upload is reproduced (only `insurer_vertical` differs, by their decision; 24 Lakshadweep rows are new).
* **October CV** = 20,327 rows, 11,219 flagged, 8 rows with a rate other than their upload's (Jharkhand taxi up to 1000 cc, Old vehicles: grid 40, upload 35 - the script follows the grid; ask which is right). The other differences are their hand work (Council Call entries, their `NCR_ONLY` / `DELHI_ONLY` / `BR11` / `JK02` groups) or rows from cells they left out.
* **November CV** converts too: 20,317 rows, 7,586 flagged, 307 cells / columns in `Not written` (it has no Negative RTO list and dropped CV NOTE 18). Written to `output\icici\` on 4 Nov with the dates 01/11/2026 - 31/03/2030 ASSUMED: the CV workbook states no date, and the new Two-Wheeler workbook is named for a mid-month date although the user's own upload dated its rows 01/11 (`--tw-new-from-date` changes it).

## 8. Checking a run

1. `python grid_doctor.py check` on each workbook; run every workbook that came.
2. Compare Private Car + Two-Wheeler with their upload (the November one reproduces all 3,318 rows); for CV compare by the portal's group names (`GRID_RTO_NAMES=group`).
3. Read `Not written`, `Council call cells` and the Negative RTO list; report the cells / columns the script did not understand.
4. **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a cluster by its label cell and a segment by its header text, read the crossing cell with openpyxl by searching for those texts (never by typing an address), parse the cell's clauses by hand and compare with the matching upload rows.

## 9. Decisions (the user, 4 Nov 2026 unless dated) - do not re-ask

1. `insurer_vertical` = `ICICI-RB-Offline` on EVERY row (also Private Car and Two-Wheeler).
2. `OTHERS:0%` writes NO zero card (`WRITE_OTHERS_ZERO = False`); a Council Call cell has no fixed rate: listed, never rated (`CC,OTHERS:n%` writes the OTHERS rate, flagged).
3. Rates are written to 8 decimals; a plain percentage, `No Biz`, `0%` or blank = a ZERO card.
4. STP rows: a segment the AOTP matrix lists takes its payout from there; every other segment takes the comprehensive cell unless the cell says `AOTP:n%`; STP rows carry NCB / ZD `NA`.
5. A banner / twin heading: the lowest payin is written (12 Nov 2026).

## 10. Open

Council Call integration (which cluster an entry applies to is hand work); the portal groups of the sub-cluster rows (`<cluster> - Condition: ... - Part: ...`); `Used car`; the `Nov'26` changes sheet as a cross-check (idea); the 5+5 Two-Wheeler grid; the 40 November CV cells whose make set has no group in the make master; Lakshadweep (24 Two-Wheeler rows, in the grid but not in their upload) and the sub-cluster / district rows need RTO mappings in the portal (the user maps RTOs themselves in cluster-name mode); the first build is not accepted.

## 11. Traps seen

F-44 (comprehensive is `On OD` for Private Car), F-45 (the user's own `RTO` column is not the cluster column), F-46 (the running web server did not list the script until restarted), P-43 (find workbooks and sheets by what they say), P-53 (cells are small sentences read as boxes).

## 12. Files

`icici_grid_to_config.py` - `tests/icici_tests.py`, `tests/structure_shift_tests.py` (subjects `icici_cv`, `icici_pc`, `icici_tw`, `icici_tw_new`), `tests/twin_column_tests.py` - `regression/inputs/icici_cv_oct26.xlsx`, `icici_cv_nov26.xlsx`, `icici_pc_nov26.xlsx`, `icici_tw_nov26.xlsx`, `icici_tw_new_nov26.xlsx`, cases `icici_cv_oct26` / `icici_pc_tw_nov26` / `icici_cv_nov26` - `layouts/icici_*.json` - `.claude/skills/icici-grid/SKILL.md` - `webapp/src/registry.js`, `webapp/test/icici.e2e.test.js` - `.vscode/launch.json` / `tasks.json`.
