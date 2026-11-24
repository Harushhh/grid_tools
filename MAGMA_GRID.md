# MAGMA HDI GENERAL - the comprehensive grid + the SATP grid -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/magma-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` section 7.2 and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A place in a grid is named by its labels (a sheet's title words, a header's words, a row's label) and by the rule that finds it - never by a row number, a column letter or a cell address. The script finds every table by what its header row says, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `magma_grid_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | **TWO files**: the comprehensive workbook (.xlsx; its premium slab is `< =3 lac`) and the SATP workbook (.xlsb or .xlsx; slab `<=2 Lacs`; `.xlsb` needs `pyxlsb`) |
| Cells | already **percent** (unlike Liberty): never multiply |
| Output | `output\magma\MAGMA_UPLOAD_<yyyy-mm>.csv` (42 columns) + the review workbook when there is anything to look at |
| Dates | given on every run (`--from-date`, `--to-date`) |
| Answer key | the user's re-upload (`--ref`); regression case `magma_oct26` (39,102 rows, 0 flagged); all 114 clusters of the October grid exist in the master and equal the user's re-upload; layout snapshots `magma_oct26_comprehensive`, `magma_oct26_satp` |
| Tests | `tests/magma_interactive_tests.py` (in `run_tests.py --all`) |

```
.venv/Scripts/python.exe magma_grid_to_config.py --grid "<comprehensive>.xlsx" --satp "<SATP>.xlsb" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their re-upload>.csv" [--no-prompt]
```
Without `--no-prompt` the script asks numbered questions about anything ambiguous (a new or odd cell value, a column with rates that no segment reads) and lists every answer at the end.

## 2. The workbooks (found by words)

* **The table** starts at the cell whose text says `UW Cluster` (the header `UW Cluster/Outgo on`), wherever it is (`find_uw_header`; `Table.col0` keeps where the table starts, so every other column is read relative to it). The clusters run down the side; the segments run across a two-row header. An alphabetical cluster list that goes on after a blank row is read to its end (`_continues`), and the run says so in its decisions.
* **The premium slab** of a sheet is read from the WORDS of its name (`slab_of`): `<=3 Lacs v2` is the same slab as `< =3 lac`. A sheet whose name is only `Sheet` + a number STOPS the run - the slab is only in the name - by design.
* **Rule sheets**: `Generic Guidelines`, `Notes`, `Note` hold text rules; a rule that is new is flagged (section 6).
* The comprehensive sheet and the SATP sheet are different files: ask for BOTH every month and run `grid_doctor.py check` on each.

## 3. How the grid becomes rows

* Payout = payin - 7, floor 0. **`IRDA`** in a cell = `On OD` 15; a Two-Wheeler with CPA = Yes gets `TW_CPA_CUT` = 2 points off (IRDA 15 becomes 13). **Private Car** is always `On OD`, with NCB Yes / No, and "non-diesel = petrol": Petrol, Electric, CNG and LPG share the Petrol column.
* The segments (`motor_segments()` for the comprehensive file, `satp_segments()` for the SATP file):

| Segment | Rows |
|---|---|
| GCV 4W | by gross vehicle weight; the 12-20 T and 20-40 T bands also by age (below 5 years / 5 years and above); GCV above 40 T; the **"3718" make group** (`Magma_3718_APR26` = the makes the grid's guideline allows up to 43.5 T, and `Magma_EX_3718_APR26` = the rest) |
| GCV 3W | make group `GCV_3W_ALL`, class Goods Carrying Rickshaw, all fuels, age 0-99 |
| PCV 3W | Electric / New / Old, each with its T&C |
| PCV bus / taxi | by the grid's own columns |
| Misc-D | tractor, harvester, construction equipment, garbage |
| Two-Wheeler | 1+1, 1+5, SAOD x CPA No / Yes x bike cc bands + scooter |
| Private Car | 1+1, 1+3, SAOD |

* **A column no segment reads is a MISSING SEGMENT, not noise**: `unused_labels` warns when such a column carries non-zero rates. The SATP columns known to carry no upload rows are in `SATP_NOT_UPLOADED` (`TP_Taxi`, `TP-PCV-Bus_Other`, `TP-PCV-Bus_School`).

## 4. A row

| Column | Value |
|---|---|
| id | 1, 2, 3 ... N (Magma counts rows; Liberty writes 1 on every row) |
| insurance_company / insurer_vertical | `MAGMA HDI GENERAL INSURANCE CO LTD` / `Not Applicable` |
| policy_type / sub_product | by segment; the SATP file gives `STP` / `Renew` |
| pi_type | `On Net` for a net segment, `On OD` for an OD segment (Private Car, IRDA), `On TP` for the SATP rows |
| payin / payout | the cell as it is (already percent); payout = payin - 7 floor 0 |
| tariff_min / tariff_max | `0` / the script's `TARIFF_MAX` |
| from_date / to_date | the dates given for the run, `DD/MM/YYYY` |
| user_id | `2966` |
| new_rto_list | the cluster as the grid writes it (the cluster rule of 30 Oct 2026) |

## 5. Names

`new_rto_list` carries the bare cluster by default. With `GRID_RTO_NAMES=group` it is `MAGMA_APR26_` + the cluster, spaces -> `_` (the `APR26` is the month the groups were made; change `RTO_PREFIX` only when the master gets new Magma groups). The "3718" make groups are dated the same way (`Magma_3718_APR26`, `Magma_EX_3718_APR26`). All 114 clusters of the October grid are in the master and equal the user's re-upload.

## 6. Notes (text the script cannot apply)

A text rule the script cannot turn into logic is flagged `Manual Review Required: [text]` on the rows it may affect - never guessed, never asked about, never a stop. A note the person resolved in the web app is never flagged again.

## 7. Edge cases and what the script does

| Situation | What happens |
|---|---|
| A cluster added or removed between months | the run does NOT stop |
| A new or odd cell value | a numbered question with clear options - never a crash, never a guess |
| A column with non-zero rates that no segment reads | warned (`unused_labels`): a missing segment |
| The table moved, or rows / columns were inserted around it | found by its `UW Cluster` words: the same upload |
| A sheet renamed (`<=3 Lacs v2`) | the slab is read from the name's words: the same upload |
| Every sheet called `Sheet` + a number | stops, by design |
| A banner above the table | flagged in `remarks`; the rates are the same |
| An alphabetical cluster list that goes on after a blank row | read to its end, said in the decisions |
| `.xlsb` SATP file | read with `pyxlsb` |

## 8. Checking a run

* Ask for BOTH files; `python grid_doctor.py check` on each; run with `--ref` = their last re-upload; read every `unused_labels` warning (a new column with rates is a missing segment).
* **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a cluster by its label cell and a segment by its header texts, read the crossing cell with openpyxl by searching for those texts (never by typing an address), and compare it with the matching upload row: the cell is already percent, payout = payin - 7.
* `python check_regressions.py --only magma_oct26`: byte-identical output.

## 9. Decisions (the user) - do not re-ask

1. Clusters added or removed between months must NOT stop the run.
2. A new or odd cell value is a question with clear options, never a crash and never a guess.
3. Text rules the script cannot apply go to `remarks` (`Manual Review Required`).
4. Magma `id` = 1..N.
5. Structure shifts (11 Nov 2026): the converter gives the SAME upload for a re-laid-out grid or a clear stop (`SheetN` names stop by design).

## 10. Open

Nothing is recorded as open in `AGENT_CONTEXT.md` 7.2.

## 11. Files

`magma_grid_to_config.py` - `tests/magma_interactive_tests.py`, `tests/structure_shift_tests.py` (subject `magma`) - `regression/inputs/magma_oct26_comprehensive.xlsx`, `magma_oct26_satp.xlsb`, case `magma_oct26` - `layouts/magma_oct26_comprehensive.json`, `layouts/magma_oct26_satp.json` - `.claude/skills/magma-grid/SKILL.md` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
