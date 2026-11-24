# SHRIRAM GENERAL (SGIC) - the broker grid and the two-wheeler grid -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/shriram-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` sections 7.6 (broker) and 7.7 (two-wheeler), patterns P-24, P-25, P-51, P-52, failures F-42, F-43 and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. Both scripts find every table by what its header row says, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | Broker grid | Two-wheeler grid |
|---|---|---|
| Script | `shriram_broker_grid_to_config.py` | `shriram_2w_to_config.py` |
| Input | the .xlsx **BROKER GRID**: the rate sheet (named `GRID <MON>' 26`, the month changes) + the sheets `Declined-Excluded`, `PAN India`, `Conditions`, `Declined vehicle`, `Zone`, `Hazardous Category`; the monthly **ADDENDUM** workbook (`--addendum`) is a separate file | the .xlsx **2W NEWROLLOVER**: `New Business - 2W`, `ROLLOVER 2W (1+1)`, `TW SAOD` (not converted), `TW decline make model`; `--broker-grid` supplies the PAN India sheet |
| Output | `output\shriram\SHRIRAM_UPLOAD_<yyyy-mm>.csv` + `SHRIRAM_REVIEW_<yyyy-mm>.xlsx` | `output\shriram\SHRIRAM_TW_UPLOAD_<yyyy-mm>.csv` + the review workbook |
| Status | first run 26 Oct 2026; the **November 2026 output was ACCEPTED by the user on 4 Nov 2026** (after their hand corrections became rules); the October and September cases are first builds | built 27 Oct 2026, **NOT yet accepted** |
| Regression | `shriram_nov26` (6,397 rows, 15 flagged), `shriram_oct26` (6,659; 343 flagged), `shriram_sep26` (6,564; 367), `shriram_sep26_addendum` (6,612); layout snapshots `shriram_nov26`, `shriram_oct26` | `shriram_2w_oct26`, `shriram_2w_sep26` |
| Answer keys | the user's portal exports (`rates_export (63).csv`: 6,911 rows, no Two-Wheeler; 84.8% of its keys reproduced - the rest is in `AGENT_CONTEXT.md` 7.6; and the corrected November export of 4 Nov) | the August 2026 Two-Wheeler upload (148 rows) and the user's September working file |
| Tests | `tests/shriram_tests.py` (209 checks), `tests/shriram_working_compare.py` (their own per-grid-row mapping: 90% of their 796 rows agree) | `tests/shriram_2w_tests.py` (73 checks), `tests/shriram_2w_working_compare.py` |

```
.venv/Scripts/python.exe shriram_broker_grid_to_config.py --grid "<grid>.xlsx" [--addendum "<addendum>.xlsx"] --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last portal export>.csv" --no-prompt
.venv/Scripts/python.exe shriram_2w_to_config.py --grid "<2W grid>.xlsx" [--broker-grid "<broker grid>.xlsx"] --from-date ... --to-date ... --ref ...
```
Ask for the monthly ADDENDUM every month: its rows override grid rows. Its own RTO / MAKE mapping columns are IGNORED (the user, 1 Nov 2026 - do not ask again).

## 2. The broker grid (found by what its headers say)

* **One grid row** - a state, the products, the discount (`DIS %`), the payout percentage (`% PO`), the policy type, the age, the underwriting remarks (`UW REMARKS`), a second remarks column, the effective date, and the person's own `RTO` column - becomes a product / class / make group / fuels / weight (cc) or seat band by `classify`. The columns are found by their HEADERS (`read_grid`, `GRID_FIELDS`), in any position.
* **Policy type**: `P` = 1+1, `L` = STP, `P & L` = both, `P & SA-OD`, `SA-OD` = SAOD, `Bundle (1+3)` = 1+3 New.
* **`% PO` is the payin** (`parse_po`): plain = `On Net`, Private Car = `On OD`, `N OD`, `N Net`, `N OD + M TP`, `IRDA` = `On OD` 15; **payout = payin - 7, floor 0**. `DIS %` -> `tariff_max = DIS + 0.01`, and a higher slab of the same key starts where the lower one ends (`tariff_min`). The age (`parse_age`) reads `upto 15 years & upto 25 Years for STP`.
* **UW REMARKS are read by `RULES`** (`read_remarks`): `Except New Mahindra` = an all-makes row + a Mahindra row for old vehicles, NCB, tankers, fuel, zones, RTO-code lists, SMP ... **Every phrase a rule understands is USED UP; whatever is left - words OR NUMBERS (P-52) - is flagged whole as `Manual Review Required`, never guessed.** A note made only of change words (`Dic by 5% & PO increased by 7.5%`, typos included) is standing (`is_change_note`). A restriction written in the remarks is APPLIED automatically when its scope is clear; only what cannot be read is flagged.
* **Duplicates**: the same card from two grid rows with different payins - the LOWEST payin is written (the user, 27 Oct 2026; `settle_duplicates`) and listed in the review sheet `Lowest rate chosen`. The same payin twice is one row. The same rule is in all converters.
* **STP rows carry NO NCB and NO zero dep**: `is_ncb` and `is_zd` are `NA` on every STP row (the user, 27 Oct and 4 Nov 2026); `check_regressions.py` enforces it.
* **Settings**: `WRITE_DECLINED_EV = False` (the `Conditions` sheet declines EV Private Car / EV taxis, so no Electric rows; `--write-ev` writes them only on request); `HAZARDOUS_MODE = "tnc"` (the grid's own percentage of `For Hazardous category PO will be N% less` goes into `add_tnc` as `Hazardous category PO N% less`; payouts as listed); `KEEP_REF_DATES = True` (with `--ref`, a row identical in key and payin to a row of that upload keeps its `from_date`).
* **Not converted**: the sheet `PCV 6+1 Short-term policy`.

### 2.1 RTO scope and names (the master mode, since 31 Oct 2026)

* `Rtos.resolve`: the STATE area minus the RTOs of the sheet `Declined-Excluded` for that product (else the `PAN India` list, except e-rickshaw / school / corporate bus) -> **the master RTO group that holds EXACTLY those RTOs** (found by content, `Rtos.pick`; state code(s) when nothing is taken out).
* **A scope the master has no group for is NEVER guessed, never widened to the whole state, never stops the run**: the RTO cell holds the plain text `New changes found, new mapping needs to be created: <Sheet Name> - <CODE/CODE/...>`; a make or model name the master lacks becomes the same kind of text in `new_vehicle_makes`; the sheet `New mappings to create` of the review workbook lists them and the rows are flagged. **Phase 2 is `master_builder.py`**: it reads those warnings from one or more upload CSVs and writes the new RTO groups (and make groups) as two-column Excel files; once a group is in the master the warning disappears.
* **The person's RTO working list** (the user, 31 Oct 2026): ONE column headed `RTO` (also `RTO working list`, `RTO list`, `RTO name`, `Working RTO`) added to the rate table, with the RTO name of each grid row, is written into `new_rto_list` EXACTLY as typed (no master lookup, no condition text). A blank, `NA`, `N/A`, `nil`, `-` or Excel-error cell (`=#REF!`, 52 cells in the November grid) has no record and keeps today's lookup; one run line says how many rows carry a name. The **addendum** file is NOT read for it. A corporate-bus row has two scopes (metro / non-metro, different payouts) that one name cannot say: the master groups stay and the run line says so.
* Make names are written as the master spells them (`Mahindra`, `Ashok`, `Hyundai`; the user's old uploads have capitals).

### 2.2 The addendum

`--addendum <workbook>` (Colab: upload it with the grid): the monthly addendum has the grid's columns and rows with `w.e.f 1st <month>` that OVERRIDE grid rows (`apply_addendum`). A goods-vehicle addendum row replaces the grid rows of the same state whose weight band it - together with the other addendum rows of that state and policy type - covers, only for the policy types it names (an addendum `L` row replaces just the STP part of a `P & L` row); a band that overlaps only in part is not trimmed - both rows are written and flagged; rows with no counterpart are added; other products are added and flagged when a grid row has the same state, product and policy type. Addendum rows are numbered 10000 + the row's own position in the addendum sheet, so they never collide with a grid row. Without the flag the run says so. Case `shriram_sep26_addendum`: 23 grid-row parts replaced; the 95 payin differences against the user's September GCV 4W upload are gone.

### 2.3 The user's hand corrections of 4 Nov 2026 are rules now (with THEIR group names)

`Bolero Model only` = `GCV_4W_BOLERO`; `Except (New Mahindra &) Bolero Model` = the Mahindra row `Mahindra_ex_Bolero`; a make-only row (`Only EECO`, `Bolero Model only`) is carved out of the other rows of its state (`CARVE_GROUPS`, said in a warning); taxi Dzire excluded / declined = `..._EX_DZIRE`, Dzire only = `SWIFT_DZIRE`, kaali-peeli Ambassador = `Ambasador` (without nil dep), `EXCLUDING M &M make` = `SHRIRAM_PCV_4W_0-7_EX_MAHINDRA_JUNE26`; `NCB Cases 20% and Above` = NCB yes + T&C `NCB >=20%`; `for NIL dep cases payout is 40%` = zero-dep rows at 40, without-nil-dep rows at the cell's payout; `EXCLUDING 1001 CC TO 1500 CC` splits the cc band; `OTHER THAN DIESEL` = no Diesel rows; Bihar's `M & M(All CC) & Others > 1500 CC` split; and `Decline Product [MISC.D - OTHERS (BODY TYPE ...)]` of the sheet `Declined vehicle` is written into the `EX:` T&C of the MISC-D others rows (the user's choice, 4 Nov 2026; F-43). Where the grid beats their upload (Delhi `Mahindra_ex_Bolero`, no Diesel for MP / RJ goods 3W, without-nil-dep taxi rows at 42.5, STP `NA`) the grid wins. Their October mapping used other group names (`SGIC_CV_BOLERO`, `>=20%`): the LATEST month wins.

## 3. The two-wheeler grid

* **New Business** -> 1+5 / New, `On Net`, payin = `Average Net PO (Converted)`, `tariff_max = DIS + 0.01` (DIS 0 with payout 0 = 100.01), age 0-99; `BIKE "<=180CC"` -> cc 0-180.01; EV -> Electric once per class (BIKE, SCOOTER). The headers `BODY TYPE` / `POLICY TYPE` of that sheet are SWAPPED: `body_of` classifies by content, not by header.
* **Rollover**: `P` -> 1+1, `L` -> STP (both Renew); age `6 to 15` -> 5.01-15.01; the column `NPO(%) / FIXED`: plain = `On Net` (STP: `On TP`), `35 ON OD` = `On OD`, `IRDA` = `On OD` 15 (none on STP), and **`Rs. 250` = a fixed rupee amount -> a FLAT CARD**: `pi_flat_amount` = the amount, the rate columns 0, `pi_type` by policy (`FLAT_PI_TYPES`: STP `On TP`, 1+1 and 1+3 `On Net`, SAOD `On OD`), `po_flat_amount` = 70% of the payin amount (`FLAT_PO_SHARE`, the user's rule of 27 Oct 2026: 100 pays out 70), flagged in `remarks`. IRDA on STP and WITHDRAW rows are NOT written (listed in the sheet `Not written`).
* **Make groups**: `ALL EXCL. TVS` -> `TW_ALL_EX_TVS` (the September working file's name; not in the master export of 25 Oct, so flagged and listed in the sheets `RTO groups to create` / `Make groups to create` with the entries it would hold), `TVS` -> `TVS_EX_APACHE` (the decline sheet declines Bajaj Pulsar and Apache); bike rollover rows carry the T&C `EX:Super bikes`.
* **RTO scope** = the STATE area minus the RTOs of the `PAN India` sheet of the broker grid (`--broker-grid`; without it a built-in October 2026 copy and a warning) minus the row's own condition (`scope_of`: `HR38/HR51/HR55 are excluded`, `Excluding OD-01,13,...`, a bare code list = only these RTOs), resolved to a master group BY CONTENT with the broker script's `Rtos`. Telangana is written `TG,TS`, Odisha `OD,OR`. The person's RTO working-list column of the sheets `New Business` and `Rollover` is used the same way as in the broker grid.
* **Not converted**: the sheet `TW SAOD` (684 rows; a warning).
* October 2026: 543 rows (235 x 1+1, 99 x 1+5, 209 x STP; 14 flat cards), 229 flagged (missing master groups / makes + the flat cards), 8 grid rows not written. Open: `scope_of` drops text beside RTO codes without a flag.

## 4. A row

| Column | Value |
|---|---|
| id | `1` on every row (the portal's export carries 1) |
| insurance_company / insurer_vertical | `SHRIRAM GENERAL INSURANCE COMPANY LTD` / `Not Applicable` |
| user_id | `2966` |
| policy_type | by the grid row (Renew / New) |
| payin / payout | `% PO` as parsed; payout = payin - 7 floor 0 |
| is_ncb / is_zd on STP rows | `NA` always |
| new_rto_list | the master group's name, the working-list text, or the `New changes found ...` text (section 2.1) |

## 5. Edge cases and what the scripts do

| Situation | What happens |
|---|---|
| The same payin heading twice (`% PO`, `Average Net PO (Converted)`, `NPO(%) / FIXED`) | the LOWEST of the plain numbers is the grid row's payin (the user, 12 Nov 2026); rows where the twins differ are listed in the review sheet `Twin columns`; a blank twin cell is no payin |
| The same heading twice for `DIS %` or any other field | stops (there is no lowest of two discount bands) |
| Every sheet called `Sheet` + a number | stops: two sheets look like the rate table (by design) |
| An Excel error in the RTO working-list column | read as no record; counted in the run line (`working_value`) |
| A new line in a text sheet (`Conditions`, `Declined vehicle`) | read; a line nothing applies is flagged, never dropped (F-43) |
| A note with a number left over | flagged whole (P-52) |
| The master export is older than the user's groups | `New changes found` rows: ask for a fresh `make_model_master` / `rto_master` export into `master\` |
| Notes in the sheets the script does not read | `grid_doctor.py check` shows every new or edited line; resolve each as applied, flagged or standing - never just "seen" |

## 6. Checking a run

1. `python grid_doctor.py check "<grid>"` and resolve EVERY `new note` / `note edited` line, including the sheets the script does not read (`Conditions`, `Declined vehicle`, `Hazardous Category`): applied, flagged or standing (F-43).
2. Run with `--ref` = their latest portal export. Read the Summary warnings (the working-list run line, the carve-out lines, the `Decline Product` line).
3. If the review shows `New changes found` rows: ask for a fresh master export.
4. **Review against their corrections** (`grid-converter` Job E): the first output of a month is in the web job folder `webapp\data\jobs\<id>\run-1\out` with its input in the same job's `in` folder (the user edits their grid afterwards); the file they really uploaded is in their month's UPLOAD folder; compare by key columns, group the differences by grid row, turn each hand edit into a rule with a test and a mutation check.
5. **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a grid row by its state and product labels and read its `% PO` / `DIS %` cells with openpyxl by searching for those texts (never by typing an address); compare with the matching upload rows (payout = payin - 7).
6. Report: rows, flagged, what differs from their export and why, what the grid decided against their upload, what they must do.

## 7. Decisions (the user) - do not re-ask

1. STP rows: `is_ncb` and `is_zd` ALWAYS `NA`. Two grid rows, one key, two payins: the LOWEST is written and listed.
2. No Electric rows for Private Car / taxi (the `Conditions` sheet declines EV); unchanged rows keep the earlier `from_date` with `--ref`.
3. Hazardous note: `HAZARDOUS_MODE = "tnc"`.
4. Their hand corrections of 4 Nov 2026 are the rules, with THEIR group names (section 2.3). Where the grid beats their upload, the grid wins. The 1001-1500 cc split was not answered and stays.
5. The addendum's own RTO / MAKE columns are ignored (1 Nov 2026).
6. A restriction written in the remarks is applied automatically when its scope is clear.
7. Two-wheeler: the flat-card payout is 70% of the payin amount (27 Oct 2026); `TW SAOD` is not converted; `Rs. 250`-type cells are flat cards.

## 8. Traps seen

F-42 (`=#REF!` cells written as RTO names; a number dropped by a words-only check), F-43 (new lines in text sheets dropped), P-51 (carve-out of a make-only row), P-52 (a leftover NUMBER is flagged), P-24 / P-25 (RTO groups by content, discount slabs chained). The shared 25-character `unknown_notes` floor is not used by the broker script (it has its own leftover check). Structure shifts: the shared `locate_table` of the declined / PAN India sheets still takes the first column that says a field.

## 9. Open

The `Conditions` general rules nobody applies (Private Car SAOD 15% where there is no SAOD row; SMP mandatory in Mumbai); the person's merged 1+3 RTO list is not built; ten STP rows typed with NCB YES in their portal (they correct them); the Two-Wheeler first build is not accepted.

## 10. Files

`shriram_broker_grid_to_config.py`, `shriram_2w_to_config.py`, `master_builder.py` - `tests/shriram_tests.py`, `tests/shriram_2w_tests.py`, `tests/shriram_working_compare.py`, `tests/shriram_2w_working_compare.py`, `tests/master_pipeline_tests.py`, `tests/working_list_tests.py`, `tests/lowest_rate_tests.py`, `tests/layout_shift_tests.py` - `regression/inputs/shriram_*.xlsx` and their cases - `layouts/shriram_nov26.json`, `layouts/shriram_oct26.json` - `.claude/skills/shriram-grid/SKILL.md` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
