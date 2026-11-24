# UNITED INDIA INSURANCE - the "Revised commission structure" circular (PDF) -> the portal's rate-config upload

Dedicated reference for ONE insurer (the user, 23 Nov 2026: every insurer has its own `.md` and its own skill). The short memory the agent loads first is `.claude/skills/united-grid/SKILL.md` (status, the user's rules, traps, open items); the long record is `AGENT_CONTEXT.md` section 7.5, patterns P-23 and P-35, rule 11 (an annexure that is not in the document is not applied) and the decision log; the rules shared by every converter are in `CLAUDE.md` and `MASTER_SCHEMA.md`.
**This file is rule-based.** A part of a circular is named by its title, its heading words and its labels - never by a page number, a row number or a column letter - because the next circular moves them all. The script finds every section by what its headings say, and every check below is done the same way (`cellref_lint.py` fails the build when a position slips into a document).

## 1. At a glance

| | |
|---|---|
| Script | `united_commission_to_config.py` (standalone; the shared question / safeguard block of all the converters) |
| Input | the Motor Department circular **PDF** (Sub Annexures included) - or, for a fully SCANNED circular, a transcription `.json` of its tables. `pdfplumber` (in `requirements.txt`) is the only PDF tool |
| Output | `output\united\UNITED_UPLOAD_<yyyy-mm>.csv` (42 columns) + the review workbook |
| Dates | given on every run; a payout equal in the Existing and the new column keeps the Existing date (01/04/2026), changed payouts take `from_date` |
| `user_id` | EMPTY (as in the July upload) |
| Regression | `united_jul26` (457 rows, 109 flagged), `united_aug26` (511, 163), `united_oct26` (496, 168; dates 26/10/2026 - 31/03/2030), `united_apr26` (451 rows; first build NOT accepted: its only answer key is the user's May upload, the same rates on 333 shared keys) |
| Tests | `tests/united_tests.py` (74 checks, July), `united_aug_tests.py` (25), `united_oct_tests.py` (23: October + the transcription reader), `united_apr_tests.py`, `united_shift_tests.py` (12, mutations of the October transcription); fuzzer `tests/fuzz_united.py N seed workers [aug|apr]` |
| Limits | `grid_doctor.py` cannot read PDFs yet |

```
.venv/Scripts/python.exe united_commission_to_config.py --grid "<circular>.pdf" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--no-prompt]
```
`--grid` may be a transcription `.json` instead; a scanned PDF is found BY ITSELF from its SHA-256 in the `source_sha256` of the transcription files (pass a `.json` as `--grid` to force one).

## 2. The circular (found by what it says)

* **`read_pdf`** returns every bordered table as cells with boxes; **`extract_sections`** finds each section by its title cell and its pair of payout columns - an "Existing Pay-out" column and the NEW column, whose heading is recognised in many wordings (`_is_new_head`: New / Revised / Proposed / Updated / Latest / Maximum + Pay-out / Payout / Commission / Brokerage; July `Pay-out (w.e.f. ...)`, August `Revised Pay-out`, April `Maximum proposed Commission/Brokerage % w.e.f. ...`). Each payout of the NEW column is joined to its states cell and its class / band labels **by position on the page** (a merged cell covers the payout's vertical centre), never by text order. Labels in merged cells that reach only some rows are carried down (P-35).
* **A payout table the script finds but cannot read STOPS the run** (expected / found): a section that has an Existing Pay-out column, or the Private Car table, whose columns it cannot find is never left out with a warning.
* **The cover letter is a scanned image**: no table, an unusable text layer - the script can only warn. **READ IT YOURSELF EVERY MONTH** (render it with `.claude/skills/grid-converter/scripts/render_pages.py <pdf> [--crop-images]` and look). It holds rules: all policy types apply unless stated otherwise, commission on Net Premium, the 40% for the preferred-city RTOs (no exception named), what happens to the additional incentives.
* **A fully scanned circular** (October 2026 = seven pages of pictures): `read_pdf` stops with `this PDF is a scan`; the Motor tables are read through a transcription (`read_transcription` explains the format: rows of cells, `{t, rs, cs}` for spans; example `regression/inputs/united_oct26.json`, the scan itself `united_oct26.pdf`). Save a transcription with `source_sha256` (the SHA-256 of the PDF's bytes) and say in the report that the numbers were read from images. Do not transcribe what you cannot see.

## 3. How the circular becomes rows

* `dispatch` maps a class / band label to a kind (TW, bus, taxi, GCV 4W ...); `specs_for` builds the rows. **States -> state codes** (`STATE_CODES`, 40 codes): `Other than above States` = the codes not named by the other payouts of the band; `All States` = `ALLINDIA`. `(Except RTOs listed in Sub Annexure ...)` -> the portal's RTO group (`ANNEX3_GROUPS`, `RTO_PREFIX`); the preferred RTOs and their commission are read from the annexure pages of the same document.
* **1+1 = `On Net`, STP = `On TP`; Two-Wheeler also SAOD (`On OD`) and 1+5 New; payout = payin - 7.** New wordings handled in October: `Seating Capacity` bands, `Two Wheeler (PCV)`, a Taxis row for Package + SATP, Private Car with Existing + Revised Commission and `Non Electric & Electric - SAOD`.
* **Private Car**: the July table is turned into rows by cc / fuel / make group as a **PARTITION of every non-electric vehicle** (diesel up to 1500 cc; above 2500 cc except the seven makes `MAKES_CAR_EX_SEVEN`; "other" = the seven makes above 2500 cc `MAKES_CAR_SEVEN` + other non-diesel up to 2500 cc + diesel 1500-2500 cc; Electric / SAOD = all vehicles). August and October have only a revised commission by cc band x states and NO make groups (`extract_car_revised`, `car_specs_revised`; the heading cell has no border, so the commission column is read as "right of the States column"): `Private_Car_ALL`.
* **The 40% preferred-city annexure** (12 cities' RTOs) is applied to the sub-products in `CAR_ANNEX4_POLICIES` (default `("1+1",)`: **Package 1+1 only**, the user's decision of 26 Oct 2026, CONFIRMED for July and August after being shown that the cover letter could be read as covering more - do not ask again, see `lessons_learned.md`). Those rows use the EX-annexure group + the annexure group (diesel up to 1500 cc keeps the table's rate there); every other Private Car sub-product uses `ALLINDIA` with no scope remark. A Package (1+1) row for SOME states cannot leave the 40% RTOs out (the master has no "these states except the annexure" group): it is written for the states and flagged (14 rows in August); an 'All states' row goes to the EX group.
* **Long Term Package (3+3)** has no sub-product: not written, flagged.

## 4. Circular-specific groups and dates

The portal's RTO groups differ per circular, so they are looked up by the date the new column takes effect: `circular_groups(date)` (July = the constants `RTO_PREFIX`, `ANNEX3_GROUPS`, `CAR_RTO_*`; August onward = `CIRCULAR_GROUPS`). **A circular with a date the script does not know gets no groups, the rows that need one are not written and a warning says so: ADD ITS ENTRY to `CIRCULAR_GROUPS` (names exactly as the RTO master spells them) when a new circular arrives.** In April the master has no group for the "other RTOs" of the 15% (2000-3500 kg) and 10% (3500-7500 kg) payouts (`CIRCULAR_GROUPS` holds `None`): those rows are not written and flagged.

## 5. What each circular taught

* **July**: the base case; the user's July upload differs from the PDF for School Bus (nine states vs All States), PCV 3W (the MP 25% missing), GCV 3500-7500 kg (Pondicherry), the STP rows of the preferred-RTO groups and the Private Car structure - the user chose "follow the PDF, list the differences" (case `united_jul26`).
* **August**: the new column is headed `Revised Pay-out`; GCV 2000-3500 Rajasthan is split by Sub Annexure 2 (31 RTOs at 56.5%); the 40% annexure is now Sub Annexure 3.
* **October** (a scan): the PDF has NO Sub Annexure pages, and the user's rule (26 Oct 2026) is that **an annexure that is not in the document is NOT applied** (never carried over from August): the GCV 2000-3500 kg MP / TN / UP / HR 15% and Rajasthan 35% payouts are written for the whole states and flagged (20 rows, `specs_for(..., notes)`), and no 40% Private Car rows are written.
* **April**: the new column is headed `Maximum proposed Commission/Brokerage % w.e.f. 01/04/2026`; GCV Sub Annexure 2 and Private Car Sub Annexure 3 are in a new layout (`read_annexure_grid`); the 15% / 10% "except Annexure-2" payouts and `Brand New SAOD / SATP` (policy type New or Renew is the user's call) are NOT written and flagged.

## 6. Names and the master

Every make group written goes through `check_names`, and `master_checks` flags a master group that does not hold what the circular says (the annexure RTO lists, the groups of the other RTOs, the Private Car make groups). In the default cluster mode a row whose RTO scope is narrower than whole states carries its condition in plain English after the cluster name (`rto_condition`, built from the annexure group names).

## 7. NOT converted (warned, and flagged on their product's rows)

Sub Annexure 2 (extra incentive by channel), TW Long Term 5+5, Motor Trade, Standalone CPA, Private Car 3+3 / Long Term Liability, April's 15% / 10% "except Annexure-2" payouts and `Brand New SAOD / SATP`.

## 8. Edge cases and what the script does

| Situation | What happens |
|---|---|
| A scanned PDF | `read_pdf` stops with `this PDF is a scan`; use a transcription |
| A circular with a date `CIRCULAR_GROUPS` does not know | no RTO groups: rows that need one are not written, a warning says so |
| A payout table with an Existing column whose columns cannot be found | the run STOPS (expected / found) |
| The new column's heading in another wording | recognised in many wordings; otherwise the stop above |
| Labels in merged cells that reach only some rows | carried down |
| An annexure that is not in the document | not applied; the payouts that cite it are written for the whole states and flagged |
| A table order change, an extra table, an extra column at either edge | the same upload (`tests/united_shift_tests.py`) |

## 9. Checking a run

1. **Render and look at EVERY page** (a scan has no text); if `read_pdf` stops with `this PDF is a scan`, transcribe the Motor tables into a `.json` and say so in the report.
2. Add the `CIRCULAR_GROUPS` entry; run with `--ref`; read every warning; report what the annexures did and did not cover.
3. **Spot-check by labels** (rule 7 of `CLAUDE.md`): find a section by its title, a row by its states / class labels and the payout by the new column's heading, read the printed number and compare it with the matching upload rows (payout = payin - 7).
4. `python check_regressions.py --only united_oct26` (and the other three cases): byte-identical output.

## 10. Decisions (the user) - do not re-ask

1. Read the cover letter yourself every month; the script cannot.
2. The 40% annexure applies to Package 1+1 only (26 Oct 2026, confirmed for July and August).
3. An annexure that is not in the document is NOT applied (26 Oct 2026, rule 11).
4. Follow the PDF and LIST the differences from the user's upload.
5. Group names change per circular: add the `CIRCULAR_GROUPS` entry for a new circular.

## 11. Traps seen

P-35 (labels in merged cells that reach only some rows are carried down); a scanned PDF has no text at all (use a transcription); a circular with a date `CIRCULAR_GROUPS` does not know gets no RTO groups; the Private Car layout differs per circular (July partition, August / October cc band x states).

## 12. Open

April 2026: the master has no group of the "other RTOs" for the 15% and 10% payouts; `Brand New SAOD / SATP` (policy type New or Renew is the user's call); the first April build is not accepted.

## 13. Files

`united_commission_to_config.py` - `tests/united_*.py`, `tests/fuzz_united.py` - `regression/inputs/united_jul26.pdf`, `united_aug26.pdf`, `united_oct26.pdf` + `united_oct26.json`, `united_apr26.pdf`, `regression/references/united_jul26_rates_export.csv`, `united_may26_rates_export.csv`, the four cases - `.claude/skills/united-grid/SKILL.md`, `.claude/skills/grid-converter/scripts/render_pages.py` - `webapp/src/registry.js` - `.vscode/launch.json` / `tasks.json`.
