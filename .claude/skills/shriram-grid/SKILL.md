---
name: shriram-grid
description: Shriram General (SGIC) knowledge for grid_tools - the broker grid (.xlsx "BROKER GRID": rate sheet "GRID <MON>' 26" + Declined-Excluded, PAN India, Conditions, Declined vehicle, Zone, Hazardous sheets) and the two-wheeler grid ("2W NEWROLLOVER") converted to SHRIRAM_UPLOAD / SHRIRAM_TW_UPLOAD CSVs. Load it with grid-converter whenever the user sends a Shriram or SGIC grid, an addendum, a SHRIRAM_UPLOAD or rates_export to review, mentions the RTO working-list column, Bolero / Dzire / Ambasador / EECO / nil dep / hazardous / New Mahindra remarks, or asks what the Shriram script does or got wrong.
---

# Shriram General - broker grid and two-wheeler grid

Dedicated reference: `SHRIRAM_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Source of truth for detail: `AGENT_CONTEXT.md` section 7.6 (broker) and 7.7 (two-wheeler), `CLAUDE.md` (the Shriram bullets), `lessons_learned.md`. This skill is the short version an agent needs before touching a Shriram grid.

## Status
- Broker script `shriram_broker_grid_to_config.py`: the **November 2026 output was ACCEPTED by the user on 4 Nov 2026** (case `shriram_nov26`, layout snapshot `shriram_nov26`, 6,397 rows, 15 flagged). October / September cases are first builds (not formally accepted).
- Two-wheeler script `shriram_2w_to_config.py`: built 27 Oct 2026, NOT accepted. The TW SAOD sheet is not converted.

## Run it
```
.venv/Scripts/python.exe shriram_broker_grid_to_config.py --grid "<grid>.xlsx" [--addendum "<addendum>.xlsx"] --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last portal export>.csv" --no-prompt
.venv/Scripts/python.exe shriram_2w_to_config.py --grid "<2W grid>.xlsx" [--broker-grid "<broker grid>.xlsx"] --from-date ... --to-date ... --ref ...
```
Output: `output\shriram\SHRIRAM_UPLOAD_<yyyy-mm>.csv` + `SHRIRAM_REVIEW_<yyyy-mm>.xlsx` (Summary, New mappings to create, Lowest rate chosen, Manual review notes). Ask for the monthly ADDENDUM every month (its rows override grid rows; its own RTO / MAKE columns are ignored - decided 1 Nov).

## How the broker grid becomes rows
One grid row (STATE | PRODUCTS | DIS % | % PO | POLICY TYPE | AGE | UW REMARKS | REMARKS 1 | EFFECTIVE DATE [+ the user's RTO column]) -> product / class / make group / fuels / GVW (cc) or seats band (`classify`). P = 1+1, L = STP, P & L both, P & SA-OD, SA-OD, Bundle (1+3) = 1+3 New. `% PO` is the payin (plain = On Net, Private Car = On OD, `N OD + M TP`, IRDA = On OD 15); payout = payin - 7 floor 0; DIS % -> `tariff_max = DIS + 0.01`, a higher slab starts where the lower ends. UW REMARKS / REMARKS 1 are read by `RULES` (`read_remarks`): every phrase a rule understands is USED UP; leftover words OR NUMBERS are flagged whole as `Manual Review Required`, never guessed.
RTO scope: the user's `RTO` working-list column is written EXACTLY as typed (blank / NA / `#REF!` = no record -> master lookup); without it `Rtos.resolve` = state minus the declined RTOs -> the master group that holds exactly those RTOs, else the plain-text `New changes found, new mapping needs to be created: ...` (phase 1; `master_builder.py` is phase 2).

## Rules the user has set (do not re-ask)
- STP rows: `is_ncb` and `is_zd` ALWAYS `NA` (27 Oct and 4 Nov); `check_regressions.py` enforces it. Two grid rows, one key, two payins: the LOWEST is written and listed (`Lowest rate chosen`).
- No Electric rows for Private Car / taxi (the Conditions sheet declines EV); `--write-ev` only on request. Unchanged rows keep the earlier `from_date` with `--ref`.
- Hazardous note: `HAZARDOUS_MODE = "tnc"` - the grid's own percentage goes into add_tnc (`Hazardous category PO N% less`), payouts as listed.
- **Their hand corrections are the rules (4 Nov 2026), with THEIR group names**: `Bolero Model only` = `GCV_4W_BOLERO`; `Except (New Mahindra &) Bolero Model` = Mahindra row `Mahindra_ex_Bolero`; a make-only row (`Only EECO`, `Bolero Model only`) is carved out of the other rows of its state (`CARVE_GROUPS`, no duplicate rates); taxi Dzire excluded / declined = `..._EX_DZIRE`, Dzire only = `SWIFT_DZIRE`, Kaali-peeli ambassador = `Ambasador` (without nil dep), `EXCLUDING M &M make` = `SHRIRAM_PCV_4W_0-7_EX_MAHINDRA_JUNE26`; `NCB Cases 20% and Above` = NCB yes + T&C `NCB >=20%`; `for NIL dep cases payout is 40%` = zero-dep rows at 40, without-nil-dep rows at the cell's rate; `EXCLUDING 1001 CC TO 1500 CC` splits the cc band; `OTHER THAN DIESEL` = no Diesel rows; Bihar `M & M(All CC) & Others > 1500 CC` split; `Decline Product [MISC.D - OTHERS (BODY TYPE ...)]` (sheet Declined vehicle) is written into the EX: T&C of the MISC-D others rows ("Write it as T&C").
- Confirmed by the user: where the grid beats their upload (Delhi `Mahindra_ex_Bolero`, no Diesel for MP / RJ goods 3W, without-nil-dep at 42.5, STP NA) the grid wins. Notes that only say what changed ("Dic by 5% & PO increased by 7.5%", typos included) are standing (`is_change_note`).
- A restriction written in the remarks is APPLIED automatically when its scope is clear; flag only what cannot be read (feedback memory `feedback-grid-remarks`).

## Monthly procedure (on top of grid-converter section 2)
1. `grid_doctor.py check "<grid>"` and resolve EVERY "new note" / "note edited" line - including the sheets the script does not read (`Conditions`, `Declined vehicle`, `Hazardous Category`): applied, flagged or standing, never just "seen" (F-43).
2. Run with `--ref` = their latest portal export. Read the Summary warnings (working-list run line, carve-out lines, `Decline Product` line).
3. If the review shows `New changes found` rows: the master export is older than their groups - ask for a fresh `make_model_master` / `rto_master` export into `master\`.
4. **Review against their corrections** (grid-converter Job E): the first output of a month is in the web job folder `webapp\data\jobs\<id>\run-1\out` with its input `in\grid.xlsx` (the user edits their grid afterwards); the file they really uploaded is in `D:\harsh\grids\<MONTH>\UPLOAD`; compare by key columns, group differences by grid row, turn each hand edit into a rule with a test and a mutation check.
5. Report: rows, flagged, what differs from their export and why, what the grid decided against their upload, what they must do.

## Traps seen
F-42 (`=#REF!` cells written as RTO names; a number dropped by a words-only check), F-43 (new lines in text sheets dropped), P-51 (carve-out of a make-only row), P-52 (a leftover NUMBER is flagged), P-24 / P-25 (RTO groups by content, discount slabs chained). The shared 25-character `unknown_notes` floor is not used by this script (it has its own leftover check). Their October mapping used other group names (SGIC_CV_BOLERO, '>=20%'): the LATEST month wins.
- (11 Nov 2026) Structure shifts (11 Nov 2026): `find_header` (both scripts): the SAME heading twice for the PAYIN ('% PO', 'Average Net PO (Converted)', 'NPO(%) / FIXED') gives the LOWEST of the plain numbers as the grid row's payin (the user, 12 Nov 2026: chosen when the row is read; the rows where the twins differ are the review sheet 'Twin columns'; a blank twin cell is no payin) - it used to take the first column and ignore the other, so a copy left of the original silently changed the rates; the same heading twice for 'DIS %' or any other field still stops (no lowest of two discount bands). Every sheet called 'SheetN' stops ('more than one sheet looks like the rate table') - by design. The shared `locate_table` of the declined / PAN India sheets still takes the first matching column.

## Open items
`Conditions` general rules not applied by anyone (Private Car SAOD 15% where no SAOD row; SMP mandatory Mumbai); 2W `scope_of` drops text beside RTO codes without a flag; their merged 1+3 RTO list not built; 10 STP rows typed with NCB YES in their portal (they correct them).
