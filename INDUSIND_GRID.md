# INDUSIND GENERAL (RELIANCE GENERAL) - the TC grid for brokers -> the portal's rate-config upload

Reference for the next person (or agent) who converts this insurer's grid. One page of facts, then the mapping, the rules, the edge cases and how to check a run.
Written 20 Nov 2026 with the first build of `indusind_grid_to_config.py`; the short memory of the insurer is `.claude/skills/indusind-grid/SKILL.md`, the long record is `AGENT_CONTEXT.md` section 7.11 and the decision log. **This file is rule-based:** a place in the grid is named by its labels and by the rule that finds it - never by a row number, a column letter or a cell address (the user, 23 Nov 2026; `cellref_lint.py` fails the build when one slips in).
**Reliance General and IndusInd General are ONE insurer: write `INDUSIND GENERAL INSURANCE COMPANY LIMITED` everywhere** (the user, 20 Nov 2026).

## 1. At a glance

| | |
|---|---|
| Script | `indusind_grid_to_config.py` (standalone, the shared question / safeguard block of all the converters) |
| Input | the workbook `Reliance_-_TC_Grid_for_Brokers_Premier-_Nov26.xlsx` (.xlsx / .xlsm; six visible sheets in November) |
| Converted | the commercial-vehicle sheet (`Nov'26`): 73 regions x 31 payout columns -> **157 rows per region** |
| Not converted (listed in the sheet `Not written`) | `Private Car`, `TW Scooter`, `Short Term Taxi` (another product each), the two deal-based `Employee Pick Up` columns |
| Notes | the sheets `Guidelines` and `Contest` and any text above / under the table: flagged `Manual Review Required` on the rows they may affect |
| Output | `output\indusind\INDUSIND_UPLOAD_<yyyy-mm>.csv` (42 columns) + `INDUSIND_REVIEW_<yyyy-mm>.xlsx` |
| Dates | **given on EVERY run** (`--from-date`, `--to-date`); the workbook's title is never read |
| Answer key | the user's own portal export `rates_export (91).csv` (case `indusind_nov26`) |
| November 2026 | 11,461 rows, 7,970 zero cards, 11,461 flagged (13 notes); 11,270 of the export's 11,304 CV rows equal in every column |

```
.venv/Scripts/python.exe indusind_grid_to_config.py --grid "<grid>.xlsx" --from-date 01/11/2026 --to-date 31/03/2030 --ref "<their last export>.csv" --no-prompt
```
Without `--no-prompt` the script asks numbered questions about the grid (two sheets with the same table, a payout cell that is not a number, dates that look odd) and lists every answer at the end.

## 2. The workbook

**The commercial-vehicle sheet** (`Nov'26` in November; found by its header words, not its name - the visible sheet whose header row says a region header above payout columns that name a product). The title text above the table (in November it still said `PO Grid for Jun'26 FY 26-27`) is never read. To the left of the payouts sit the label columns:

* a `Zone` column (West / North / South / East); a first `RTO Region` column = the PARENT region (36 distinct names: `ROM` for five Maharashtra cities, `GUJARAT`, `UP West` ...); a second `RTO Region` column = the CLUSTER (73 names, one per row, e.g. `MUMBAI`, `AHMEDNAGAR`, `GUJ 1`, `Indore`, `Ladakh`). **The cluster column is the one with more distinct names** (decided by the values, so a swapped pair works too); two columns with the same count and different names stop the run.
* A header band of three merged levels: the top level = a MAKE LABEL (only over the two GCV blocks that carry a make qualifier, the `< 2.0 K` block and the `2 - 2.5 K` block), the middle level = the PRODUCT SEGMENT, the lowest level = the LEAF (`COMP`, `STP`, `All`, `Comp - Others ( Petrol + CNG + Battery )`, `Diesel - Comp & STP`, `Comp/STP`, `Comp (Fresh)` ...). Every payout is a FRACTION (`0.725` = 72.5%).
* One data row per region (73 in November); nothing under the table in November.

**Other sheets**: `Private Car` (24 regions x 5 columns: Fresh 1+3 / 3+3, Comp non-diesel / diesel, SAOD, STP; notes under the table: STP only for non-diesel, EW flat 20%, 10% lower for non-NCB Maruti, 2.5% extra PO for secure premium), `TW Scooter` (47 x 5: Fresh 1+5 / 5+5, Comp, SA OD, STP, notes), `Short Term Taxi` (72 x 4: Non-ND / ND x petrol + CNG + battery / diesel, volume enabler notes) - tables of other products, listed in `Not written` with what they hold; `Guidelines` (Kaali Pili taxi 40% in Mumbai / Delhi / Kolkata, School Bus volume slabs and their conditions, hygiene recoveries, specific deals) and `Contest` (LCV contest, TW scooter contest): text, read as notes (section 7).

## 3. How a header becomes rows (the brain)

Nothing is found by a cell address. `find_anchor` finds the header row by its words (a cell that says a region - `RTO Region`, `Region`, `RTO Cluster`, `Cluster` - with at least three header columns to its right); `RateTable` carries the merged cells into every column they cover and builds each payout column's CHAIN of header texts (the band reaches up while the rows above have text over the payout columns; one blank row inside the band does not end it; a text shared by every payout column that reads like no header - no product, no Comp / STP, no make - is a banner = a note). `classify_chain` then gives each chain a meaning, **by the words, not by the level**:

* a text that names a product is the SEGMENT (`SEG_PATTERNS`: school bus, other bus, PCV 3W, taxi / PCV 7+1, GCV by weight band, GCV 3W, car carrier, flat bed, MISD CPM, tractor, employee pick up);
* a text with Comp / STP / All / Fresh / a fuel is the LEAF (`parse_leaf`: Comp = 1+1, STP = STP, All and Comp/STP = both, `Fresh` = age 0-1 + T&C `New`, `Non fresh` = age 1-99 + `Old`, fuels in the text);
* a text of makes or 'Others' is a MAKE LABEL (`parse_make_label`).
Every word is USED UP by a pattern (`parse_clause`, `eat_paren`: `NND` = no nil dep = is_zd NO, `ND` = nil dep = YES, `Non diesel`, `All Fuel Type`, a make list, `Other makes`, a body-type list, `Only Deal Based`); **a word nobody understands refuses that column** - it is warned, listed in the sheet `Unread parts` (status NOT READ) and NOT written, never guessed. A title with two products (`PCV Other bus > 8 STR & PCV 7+1 ( Other makes )`) gives two specs.
Weight words (`< 2.0 K`, `2 - 2.5 K`, `2.5K - 3.5K`, `Above 3.5K`, `7.5 T - 12 T`; K or T = a thousand kg) become kg bands (`band_of`); seat words (`<6 Str`, `7+1`, `> 8 STR`) become seat bands (`seats_of`).

### The columns of November (hand-checked against the export)

| Header (make label / segment / leaf) | product / class | make group | fuels | policy | nil dep | age | cc (kg) | seats | T&C |
|---|---|---|---|---|---|---|---|---|---|
| School Bus / All | PCV 4W / School Bus | `BUS_ALL` | all 5 | 1+1 + STP | NA | 0-99 | - | - | - |
| PCV 3W (Non diesel) / COMP, STP | PCV 3W / Passenger Carrying Rickshaw | `PCV_3W_ALL` | Petrol CNG LPG Electric | 1+1 / STP | NA | | | | |
| PCV TAXI <6 Str ( NND) / Comp - Others (Petrol + CNG + Battery) | PCV 4W / Taxi | `TAXI_ALL` | Petrol CNG LPG Electric | 1+1 | NO | | | 0-6.01 | |
| ... / Diesel - Comp & STP | same | `TAXI_ALL` | Diesel | 1+1 + STP | NO (STP NA) | | | 0-6.01 | |
| ... / STP - Others | same | `TAXI_ALL` | Petrol CNG LPG Electric | STP | NA | | | 0-6.01 | |
| PCV TAXI <6 Str ( ND) / Comp - Others | same | `TAXI_ALL` | Petrol CNG LPG Electric | 1+1 | YES | | | 0-6.01 | |
| **[TATA / Maruti / Mahindra ( Only Jeeto & Supro )]** GCV < 2.0 K / COMP, STP | GCV 4W / NA | `TATA,Maruti,Mahindra` | all 5 | 1+1 / STP | NA | | 0-2000.01 | | |
| **[Others Tata / Maruti / Mahindra ( Mahindra Below 2T other than Jeeto & Supro )]** GCV 2 - 2.5 K | GCV 4W / NA | `GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA` | all 5 | 1+1 / STP | NA | | 2000.01-2500.01 | | |
| GCV 2.5K - 3.5K | GCV 4W / NA | `GCV_4W_ALL` | all 5 | 1+1 / STP | NA | | 2500.01-3500.01 | | |
| GCV Above 3.5K | GCV 4W / NA | `GCV_4W_ALL` | all 5 | 1+1 / STP | NA | | 3500.01-99999 | | |
| GCV 3W | GCV 3W / Goods Carrying Rickshaw | `GCV_3W_ALL` | all 5 | 1+1 / STP | NA | | | | |
| PCV Taxi 7+1 ( Only NND ) ( Maruti, Toyota, Mahindra, Kia, MG ) All Fuel Type | PCV 4W / Taxi | `TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG` | all 5 | 1+1 / STP | NO (STP NA) | | | 6.01-8.01 | |
| PCV Other bus > 8 STR & PCV 7+1 ( Other makes ) | (a) PCV 4W / Other Bus; (b) PCV 4W / Taxi | (a) `BUS_ALL`; (b) `TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG` | all 5 | 1+1 / STP | NA | | | (a) 8-999; (b) 6.01-8.01 | |
| Car Carrier / Comp/STP | MISCD / MISCD-Others | `MISCD_ALL` | all 5 | 1+1 + STP | NA | | | | none (**guard**, section 6) |
| Flat Bed | MISCD / MISCD-Others | `MISCD_ALL` | all 5 | 1+1 / STP | NA | | | | `Flat Bed` |
| MISD CPM (Only JCB, L&T and Caterpillar) | MISCD / MISCD-Others | `MISCD_ALL` | all 5 | 1+1 / STP | NA | | | | `JCB, L&T,Caterpillar` |
| Tractor - (Agricultural without Trailer) / Comp (Fresh) | MISCD / MISCD-Tractor | `MISCD_ALL` | all 5 | 1+1 | NA | 0-1 | | | `New,Agriculture Tractor` |
| ... / Comp (Non fresh) | MISCD-Tractor | `MISCD_ALL` | all 5 | 1+1 | NA | 1-99 | | | `Old,Agriculture Tractor` |
| ... / STP | MISCD-Tractor | `MISCD_ALL` | all 5 | STP | NA | 0-99 | | | none |
| Employee Pick Up ( Only Deal Based ) / COMP, STP | **not written** (a deal, not a rate card; 4 regions carry a payout in it) | | | | | | | | |

Rows per region by header group: School Bus 10, PCV 3W 8, PCV taxi < 6 seats 14, GCV 4W 40 (four blocks of 10), GCV 3W 10, PCV taxi 7+1 10, other bus + 7+1 other makes 20, Car Carrier 10, Flat Bed 10, MISD CPM 10, Tractor 15 = **157** (all columns: age 0-99, cc 0-99999, seats 0-999 where the cell is empty; `LPG` rides with `CNG`).

## 4. A row

| Column | Value |
|---|---|
| id | `1` (the portal's export carries 1 on every row) |
| insurance_company | `INDUSIND GENERAL INSURANCE COMPANY LIMITED` |
| insurer_vertical | `Not Applicable` |
| policy_type | `Renew` |
| sub_product | `1+1` (Comp) / `STP`; pi_type and po_type `On Net` for 1+1 (the rate in `pi_net_rate` / `po_net_rate`), `On TP` for STP (`pi_tp_rate` / `po_tp_rate`) |
| pi | the cell x 100, cleaned of float noise (`RATE_DECIMALS = 6`: 0.47500000000000003 -> 47.5; the CSV rounds to 4) |
| po | `max(pi - 7, 0)`; a zero or blank cell (or NA / Declined / `-`) is a ZERO card; 39 rows have a payin of 7 or less and a payout floored to 0 |
| is_ncb / is_cpa | `NA`; is_zd `NA` except the taxi ND / NND columns (1+1 only; **STP rows are always `NA`** - the user's rule) |
| tariff_min / tariff_max | `0` / `100.01` |
| from_date / to_date | the dates given for the run, `DD/MM/YYYY` |
| user_id | `2966`; veh_use empty; remarks = the `Manual Review Required: [...]` flags |
| new_rto_list | the cluster exactly as the grid writes it (`MUMBAI`, `Indore`, `GUJ 1`): the user's standing rule - **never** `RELIANCE_SEPT26_<region>` (that name exists only in `GRID_RTO_NAMES=group`, section 5) |

A payout cell that is not a number is put to the person (a stop with `--no-prompt`, naming the cell); `NA` / blank / `Declined` = zero card; a number above 1 (65 instead of 0.65) is asked about, never taken as 6,500%.

## 5. Names: makes and regions

* **Make groups** are the master's own names (the user's decision of 20 Nov 2026): `TATA,Maruti,Mahindra` (a list of the master's single-make groups, in the order the label gives), `GCV_4W_ALL_EX_TATA_MARUTI_MAHINDRA`, `GCV_4W_ALL`, `TAXI_ALL`, `TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG` (7+1 taxis of Maruti, Toyota, Mahindra, Kia, MG), `TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG` (the other makes), `BUS_ALL`, `PCV_3W_ALL`, `GCV_3W_ALL`, `MISCD_ALL`. No name has a leading or trailing space (the user's export had two with one). A make set not in `MAKE_SETS` is written as the single-make groups (a list) or `<family group>_EX_<MAKES>` (all except) and checked against the master like every name (`check_names`; a missing one is flagged and listed in `Make groups to create` with how it is made). All names written in November are in the master of 25 Oct.
* **Regions**: `new_rto_list` carries the bare region. The master is used for ONE check (the user's decision of 20 Nov 2026): a region with no RTO group in the master is written all the same and **flagged `no RTO group`** (November: Indore, 157 rows). The group is looked up by the portal's naming - `RTO_GROUP_PREFIX` + the region (`RELIANCE_SEPT26_Indore`), `RTO_GROUP_ALIASES` for a spelling the master changed (`TRIVANDRUM` = `Thiruvananthapuram`) - by name only; the sheet `RTO groups to create` names the group, the rows flagged and the parent region the grid puts the region under (the grid lists no RTO codes). If fewer than 20% of the regions have a group under the prefix the prefix is out of date: the check is skipped with a WARNING.
* `GRID_RTO_NAMES=group` writes `RELIANCE_SEPT26_<region>` and `check_names` checks them against the master (the older test suites run in it); the default is `cluster`.

## 6. Conditional readings - only when the wording is in the grid

The user (20 Nov 2026): "keep the model-level qualifier flags, the GCV 4W tonnage / make pairings and the Car Carrier non-zero guard strictly conditional - trigger them dynamically only when those exact wordings or conditions appear in a grid. Build this as a reusable, header-driven brain for future grids."

* **Model-level qualifier**: a make label with a bracket that says `Only ...`, `... other than ...`, `except`, `excluding`, `below`, `above`, `upto` (November: `Only Jeeto & Supro` over the GCV `< 2.0 K` block, `Mahindra Below 2T other than Jeeto & Supro` over the GCV `2 - 2.5 K` block). A make group cannot say a model, so the rows are written for every model of the make group and the rows of THAT column are flagged: *`GCV < 2.0 K: 'TATA / Maruti / Mahindra ( Only Jeeto & Supro )' holds a model-level condition ('Only Jeeto & Supro') that make groups cannot express - the rows are written for every model of the make group TATA,Maruti,Mahindra`*. Take the label away and nothing is flagged and both columns become `GCV_4W_ALL`; reword it (`( Only Jeeto )`) and the new wording is flagged too.
* **The make pairing**: a make list gives the single-make groups, `Others` + a list the complement, no label `GCV_4W_ALL`; `( Other makes )` in a 7+1 taxi title is the complement of the make list of the 7+1 taxi block before it (a refusal when there is none).
* **Car Carrier guard**: a MISCD-Others column with no T&C (Car Carrier) pays EVERY MISCD-Others vehicle (flat bed, JCB ... too) the moment its payout is above 0. It is written as a zero card in November and flagged nowhere; a region where its cell is above 0 gets the rows of that region flagged, naming the region.
* **Deal based**: `( Only Deal Based )` = not a rate card: not written, listed in `Not written` with how many regions carry a payout (4 in November). `Employee Pick Up` without that wording is a product with no mapping: refused.

## 7. The notes (text the script cannot apply)

Every text block of a sheet that holds no table, and every text above / beside / under the rate table, goes through `unknown_notes`: a label (under 25 characters) or a standing note (`STANDING_NOTES`: the hidden title `PO Grid for ...`, `PO at ODD as per UW`) is no action; anything else is flagged `Manual Review Required: [text]` on the rows it may affect, **never guessed, never asked about, never a stop**.
* A BLOCK is consecutive non-empty rows (a blank row ends it), rendered `cells joined by ' '`, `rows by '; '`, a fraction under 1 shown as a percent (`Mumbai 40%`); a block longer than 420 characters is split in parts of at most 420 (a flag is cut at 500 anyway), each part after the first starting with the block's first row (so it still says what it is about, and its words still scope it).
* SCOPE = the products the text names (`SCOPE_WORDS`: school bus, taxi / kaali pili, LCV / GCV / goods, PCV, MISCD words); a text that names none (hygiene recoveries, specific deals) affects EVERY row - never fewer rows than it may affect; a text about two-wheelers / Private Car only affects no row (listed in `Notes read` as `no rows affected`). A table whose every row key is a region of the grid (`Location | TC %`: Mumbai, Delhi, Kolkata) keeps to those regions (102 taxi rows in November); if a key is not a region of the grid it cannot be placed and the note keeps to the product.
* A text sheet with no sentence about payouts, slabs, contests or deals is not a rule sheet: warned and listed in `Unread parts`, nothing flagged (an unrelated sheet added to the workbook must not change the upload).
* A note the person resolved in the web app (`GRID_STANDING_NOTES_FILE`) is never flagged again.
* The sentences of the sheets of other products (`TW Scooter`, `Private Car`, `Short Term Taxi`: e.g. 'Grid not applicable for Bike Comp ...', the short-term-taxi volume enabler) are listed in `Notes read` as `not applied (the sheet is not converted in this build)` - the rules of a product to come are not lost - and flag no row.
November's 13 flags: the two qualifiers (730 rows each), Kaali Pili (102), the School Bus volume block (4 parts, 730 rows each), the hygiene guideline (2 parts) and `Specific Deals` (all 11,461 rows), the LCV contest (3,650: GCV 4W and 3W), Indore (157). The TW scooter contest affects no written row.

## 8. Edge cases and what the script does

| Situation | What happens |
|---|---|
| Header with a word nobody understands (`COMP + RSA`, `GCV Above 3.5K Special Permit`, a new product) | that column is REFUSED: warning with the column and the words, `Unread parts` (NOT READ), the other columns are written |
| Two products in one header cell | read as two specs (`&` before `PCV` / `GCV` / `MISC`) |
| `Other makes` with no make list before it | the column is refused |
| Two visible sheets with the same table (last month's left in) | asks which is this month's; a stop (`more than one sheet holds the same table`) with `--no-prompt` |
| A hidden sheet (an old version) | never read |
| The table split in two sheets (different products on each) | both read, every product once |
| No sheet with a region header over product columns | stops: `expected a visible sheet whose header row says 'RTO Region' above payout columns that name a product ... found: ... Visible sheets: [...]` |
| Another insurer's workbook | the same stop |
| A payout column there twice under one heading | the LOWEST payin is written (the user's rule of 12 Nov 2026), listed in `Lowest rate chosen` |
| A region row twice | the same rule |
| A blank row inside the header band or the data | does not end the table |
| A banner over every payout column (a text that is no product, Comp / STP or make) | read as a note (flagged on every row), the rows are the same; a product title shared by every column of a small table is a header, not a banner |
| An extra column at the right edge (`Comments`) | warned and listed in `Unread parts`; the rows are the same |
| A cell that is NA / blank / Declined / `-` | a zero card (blank cells are counted in the run's decisions; more than 10% blank is WARNED about: formulas without saved values? open, save in Excel, run again) |
| Rows or columns with payouts hidden in Excel | written like the others and WARNED about (the warning lists them so the person can unhide and check them) |
| Every payout column refused | stops, naming the first refusals (`no column read, e.g. column <letter>: words of the column header ... that nothing understands: ...`) |
| A cell that is other text, a number above 1 or negative | asked about (a stop with `--no-prompt`, naming the cell) |
| `--from-date` / `--to-date` missing in a run that cannot ask | `STOPPED - pass --from-date and --to-date (DD/MM/YYYY)` |
| `to_date` before `from_date` | stops before anything is read |
| The workbook's title says another month | ignored |

## 9. Differences from the user's own export of the same grid (the grid wins; listed here and, with `--ref`, in the run's log and the sheet `Changed vs earlier upload`)

1. **34 diesel-taxi rows** (1+1 and STP of 17 regions): the user's export has, in each of these 17 regions, the diesel value of the NEXT grid row (checked for all 17; after the last row = blank = 0): a one-row offset in their file. The script follows the grid's `Diesel - Comp & STP` column. Payin in %, grid / export: ANDHRA PRADESH 0 / 10, CHHATTISGARH 32.5 / 52.5, CUTTACK 37.5 / 32.5, DELHI 15 / 12.5, GOA 40 / 15, GUJ 1 15 / 0, HYDERABAD 10 / 0, JHARKHAND 32.5 / 37.5, KOLKATA 52.5 / 30, LUDHIANA 12.5 / 0, MUMBAI 25 / 15, NORTH EAST 21.25 / 0, PATNA 37.5 / 32.5, PONDICHERRY 0 / 15, PUNE 25 / 10, UTTARAKHAND 0 / 37.5, WEST BENGAL 30 / 21.25 (the payout is the payin - 7, floor 0). The other regions agree, including Nagpur and `ROM`.
2. **Indore**: 157 rows the export lacks (same payouts as Madhya Pradesh), flagged `no RTO group`.
3. **The two 7+1 taxi make names**: the export has `Maruti,Toyota,Mahindra,Kia,MG ` and `Taxi_all_Ex_Maruti_Toyota_Mahindra_Kia_MG ` (a trailing space, 1,440 rows); the upload has the master's `TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG` and `TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG` (the user's choice). The comparison folds the two spellings (`REF_MAKE_ALIASES`), so these rows count as equal.
4. **79 Private Car rows** of the export (from_date 13-10-2026: 64 x 1+1, 5 x 1+3, 10 x SAOD) are an EARLIER set; they do not derive from this workbook's Private Car sheet (e.g. Kerala diesel comprehensive is 0 in the sheet and 5 / 10 in those rows) and are not produced.
5. Everything else - regions, products, make groups, fuels, ages, weights, seats, T&C, nil dep, payins, payouts, types, dates, user_id - is equal: 11,270 rows.

## 10. Checking a run

* The run prints rows per product / class / policy type, zero cards, the PO floor, every WARNING and what was decided. `INDUSIND_REVIEW` lists `Not written`, `RTO groups to create`, `Make groups to create`, `Changed vs earlier upload`, `Lowest rate chosen`, `Manual review notes`, `Notes read`, `Unread parts`.
* Spot-check BY LABELS with openpyxl, independently of the script (never by a typed address): find the data row whose cluster cell reads `MUMBAI` and the payout column whose header chain reads `School Bus / All`, read the crossing cell (0.725) -> 72.5 in `pi_net_rate` (1+1) / `pi_tp_rate` (STP) of the Mumbai `School Bus` rows, `po` 65.5; the same row under the make label `TATA / Maruti / Mahindra ( Only Jeeto & Supro )` / `GCV < 2.0 K` / `COMP` (0.675) -> 67.5 on `TATA,Maruti,Mahindra`; the same row under `PCV TAXI <6 Str ( NND)` / `Diesel - Comp & STP` (0.25) -> 25 on the Mumbai diesel taxi rows (the export says 15).
* `.venv/Scripts/python.exe tests/indusind_tests.py` (167 checks: every header of November read by its words against a hand-checked table, cells against rows, the export, the conditional readings, unseen wordings, cells, structure, notes, groups, dates, the command line); `python check_regressions.py --only indusind_nov26`; `SHIFT_ONLY=indusind python tests/structure_shift_tests.py` (14 operators apply: renamed / reordered / extra / hidden sheets, rows and columns inserted, header noise, a blank row, an extra column: the same upload; the three that need two adjacent columns outside a merged header - two swapped columns, a payout column twice - do not apply to this grid, the twin-column rule is tested by hand in `tests/indusind_tests.py`; a banner or title text above the table: only `remarks` differ; a visible copy of the table: a stop).

## 11. Decisions (the user, 20 Nov 2026) - do not ask again

1. Diesel taxi: follow the grid; fix the 34 shifted rows; list them in the output log.
2. Indore: write all 157 rows, flag `no RTO group`.
3. Make groups: the master's names (`TAXIALL_MAHI_KIA_TOYOTA_MARUTI_MG`, `TAXIALL_EX_MAHI_KIA_TOYOTA_MARUTI_MG`), no trailing / extra spaces; the existing `master\` files are used for this build (a follow-up: API keys in a `.env` to keep `rto_master` and `make_model_master` in sync).
4. RTO names: the standing rule - the bare grid region, never `RELIANCE_SEPT26_<region>`.
5. Scope: the `Nov'26` commercial-vehicle sheet only; Private Car, TW Scooter and Short Term Taxi go to `Not written`; the sheet-level notes (School Bus volume slabs, LCV contest, Kaali Pili 40%) are flagged for manual review.
6. Dynamic rules, no hard-coding: the qualifier flags, the tonnage / make pairings and the Car Carrier guard fire only when those wordings or conditions are in the grid.
7. Dates: the workbook's title is not read, `to_date` is not hard-coded: `from_date` and `to_date` on every run.
8. (23 Nov 2026) The hygiene recoveries and `Specific Deals` notes name no product and flag every row: they STAY flagged (the user: 'Keep flagging'); they are not standing notes - do not ask again.

## 12. Open

Private Car / TW Scooter / Short Term Taxi (mapping rules and an answer key are needed; the export's Private Car rows are an earlier set); hygiene recoveries and `Specific Deals` flag every row and STAY flagged (decided 23 Nov 2026: 'Keep flagging' - section 11, item 8); the master has make groups for the Jeeto / Supro qualifier (`JEETO_SUPRO`, `MAHINDRA_EX_JEETO_SUPRO`, `GCV4W_ALL_EX_TATA_MARUTI_MAHINDRA_JEETO_SUPRO` ...) - apply them if the user wants; the master sync by API; the first build is not accepted yet.

## 13. Files

`indusind_grid_to_config.py` - `tests/indusind_tests.py` - `regression/inputs/indusind_nov26.xlsx`, `regression/references/indusind_nov26_rates_export.csv`, case `indusind_nov26` in `regression/cases.json` - `layouts/indusind_nov26.json` - `.claude/skills/indusind-grid/SKILL.md` - `webapp/src/registry.js` (id `indusind`), `webapp/test/indusind.e2e.test.js` - `.vscode/launch.json` / `tasks.json` - subject `indusind` in `tests/structure_shift_tests.py` - the script is in the shared-block suites (`service_tests`, `cluster_names_tests`, `resolved_notes_tests`, `lowest_rate_tests`).

## 14. October 2026 - a second layout generation of the same grid (added 23 Nov 2026)

Workbook `Reliance_-_PO_Grid_Broking_Premier_11-Oct_26.xlsx` (case `indusind_oct26`, dates 13/10/2026 - 31/10/2026, answer key the user's export `rates_export (92).csv`). The SAME script reads it: 8,784 rows = 72 regions x 122 rows, 5,369 zero cards, 8,784 flagged; against the export every payin and payout is equal (0 changed), 0 new, 181 removed (their 30 M & HCV rows, 72 nil-dep diesel taxi rows, 79 earlier Private Car rows). November is unchanged (byte-identical).

| | November | October |
|---|---|---|
| Rate sheet | `Nov'26`, 73 regions x 31 payout columns | `Oct 26`, 72 regions x 30 payout columns, visible title row |
| Region headers | `RTO Region` twice | `RTO Fy 26 -27` (parent, no 'region') and `RTO Region FY 26-27` |
| Header band | make label / segment / leaf | a top row of GROUP TITLES (`GCV < 2.0 K`, `PCV TAXI <6 Seater`, `MISD CPM` ...) above sub-title and leaf; a stray backtick |
| Regions | Indore, Ladakh | no Indore; `LEH & LAHDHK` (Ladakh) |
| New / gone columns | GCV above 3.5 K, GCV 3W, other bus / 7+1 other makes | gone; new: Kaali Peeli, tractor non-fresh split by age, diesel taxi Comp and STP in two columns, 3 `PC STP Non-Diesel` columns, Car Carrier and Employee Pick up are deal columns |
| Sheets | Private Car, TW Scooter, Short Term Taxi, Guidelines, Contest | renamed product sheets, `Enabler, PO Guidelines & T&C` (a new M & HCV Grid table), Contest, new `RTO List` (1,501 RTO codes) and `Make Model wise segment` |

**What the reader learned** (no month-specific code): a region header is any short cell with 'region' / 'cluster' (never a code column); noise text (a backtick) is not a header; two product texts of one header are MERGED when they agree and the column is refused when they conflict; nil dep in words (`Without Nil Dep`, `Nil Dep`); fuel lists in the leaf (`Comp- Other -(Battery,CNG, Petrol)`, `STP- Diesel`); `All Make/All Ages`; age words (`< 5 Years` = up to 5.01, `> 5 Years` = from 5.01, Fresh / New = 0-1, Non fresh from 1); `(Only deal basis)` on any product = deal, not written; Private Car / two-wheeler columns inside the CV sheet = another product, listed in `Not written`; Kaali Peeli = a taxi style written as the T&C `Kaali Peeli` and flagged.

**The user's decisions (23 Nov 2026)**: Kaali Peeli cards (all fuels / makes, Comp + STP, T&C `Kaali Peeli`, flagged); the standard age slabs above; the M & HCV grid is NOT written (it stays a flagged note on the GCV 4W rows); the 'Others' band is what the header says (`GCV < 2.5 K` = 0-2500.01, confirmed by their export); Private Car / TW / Short Term Taxi and the 3 PC STP columns stay `Not written`; dates only from the run.

**The workbook's RTO List** gives 'did you mean' hints (`LEH & LAHDHK` = Ladakh, `TRIVANDRUM` = Thiruvananthapuram; a hint needs a clearly closest free name) and the RTO codes in `RTO groups to create`; the names written stay the grid's. A refused column raises `INCOMPLETE: n of m payout columns were not read` as the FIRST warning (run log and review Summary).

**Differences from their export (listed in the sheet `Other names vs earlier upload`, never applied - November must stay identical)**: the master's Jeeto / Supro make groups instead of `TATA,Maruti,Mahindra` / its EX group (1,440 rows); `PCV_4W_ALL` instead of `BUS_ALL` / `TAXI_ALL` (720 + 1,728); `JCB,Caterpillar` instead of `JCB, L&T,Caterpillar` (720); the T&C `Non,Kalli Pilli` on the ordinary <6-seat taxi cards (1,008) and a nil-dep DIESEL taxi row (72) that the grid's header does not give.

Tests: `tests/indusind_tests.py` (272 checks: the 30 October columns against a hand-checked table, all 1,800 payout cells re-read with openpyxl, the export, the new wordings and refusals, RTO List hints, INCOMPLETE, the command line), shift subject `indusind_oct` (18 results; a column / blank row inside the M & HCV text table changes only the quoted note), `layouts/indusind_oct26.json`. Open: see the skill (apply the Jeeto / Supro groups? the 'Non,Kalli Pilli' T&C? 12 SILENT columns in other-product sheets in `tests/unread_audit.py indusind_oct26`).
