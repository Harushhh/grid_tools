---
name: united-grid
description: United India Insurance knowledge for grid_tools - the "Revised commission structure" PDF circular (Motor Department, Sub Annexures; sometimes a fully SCANNED PDF read through a transcription .json) converted to UNITED_UPLOAD_<yyyy-mm>.csv by united_commission_to_config.py. Load it with grid-converter whenever the user sends a United India circular / commission PDF, a United rates_export to review, or asks about Sub Annexure 3 / 4, the 40% preferred-city RTOs, CIRCULAR_GROUPS, the scanned cover letter, or what the United script does.
---

# United India - commission circular (PDF)

Dedicated reference: `UNITED_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` section 7.5, `CLAUDE.md` (the United bullet), `lessons_learned.md`, P-23 / P-35, rule 11 (an annexure not in the document is not applied). Tests: `tests/united_tests.py` (July), `united_aug_tests.py`, `united_oct_tests.py`, `united_apr_tests.py`; fuzzer `tests/fuzz_united.py N seed workers [aug|apr]`.

## Run it
```
.venv/Scripts/python.exe united_commission_to_config.py --grid "<circular>.pdf" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their last export>.csv" [--no-prompt]
```
`--grid` may be a transcription `.json` instead (a fully scanned PDF is found BY ITSELF from its SHA-256 in `source_sha256` of `regression/inputs/united_oct26.json`-style files). Output `output\united\UNITED_UPLOAD_<yyyy-mm>.csv`. `user_id` is empty (as in the July upload). `grid_doctor.py` cannot read PDFs.

## Status
Regression: `united_jul26` (457 rows, 109 flagged), `united_aug26` (511, 163), `united_oct26` (496, 168; dates 26/10/2026 - 31/03/2030), `united_apr26` (451 rows; first build NOT accepted - its only answer key is the user's May upload, same rates on 333 shared keys).

## How it reads the circular
First PDF input; needs `pdfplumber`. `read_pdf` returns every bordered table with boxes; `extract_sections` finds each section (title cell, an "Existing Pay-out" and a new-payout column pair - headed `Pay-out (w.e.f. ...)` in July, `Revised Pay-out` in August, `Maximum proposed Commission/Brokerage % w.e.f. ...` in April) and joins each payout to its states cell and class / band labels by POSITION (a merged cell covers the payout's vertical centre) - never by text order. `dispatch` maps labels to a kind (TW, bus, taxi, GCV 4W ...); `specs_for` builds rows: states -> codes (`STATE_CODES`; "Other than above States" = the codes not named by the other payouts of the band; "All States" = `ALLINDIA`), 1+1 On Net, STP On TP, TW also SAOD (On OD) and 1+5; payout = payin - 7. Private Car: July = segments by cc / fuel / make group as a PARTITION of every non-electric vehicle; August / October = cc band x states x policy group with `Private_Car_ALL`.

## Rules the user has set (do not re-ask)
- **READ PAGE 1 YOURSELF EVERY MONTH.** It is a scanned cover letter (an image); the script can only warn. Render it (`.claude/skills/grid-converter/scripts/render_pages.py <pdf> [--crop-images]`) and read: all policy types apply unless stated (item 2), commission on Net Premium (item 3), the 40% for preferred-city RTOs (item 5), the fate of the additional incentives (item 6).
- The 40% preferred-city annexure applies to **Package 1+1 only** (decided 26 Oct 2026 and CONFIRMED for July and August after being shown that the cover letter could be read wider - do not ask again). Those rows use the EX-annexure group + the annexure group; diesel up to 1500 cc keeps the table's rate there.
- **An annexure that is not in the document is NOT applied** (never carried over from an earlier circular): the payout that cites it is written for the whole state(s) and flagged, rows that exist only through the annexure are not written (October: 20 flagged rows, no 40% Private Car rows).
- Group names change per circular: `circular_groups(date)` (`CIRCULAR_GROUPS`) - a circular with another date gets no groups and a warning, so **add its entry** (names exactly as the RTO master spells them). A payout equal in the Existing and the new column keeps the Existing date (01/04/2026); changed payouts take `from_date`.
- Follow the PDF and LIST the differences from the user's upload (July upload differs for School Bus, PCV 3W MP 25%, GCV 3500-7500 Pondicherry, STP rows of the preferred-RTO groups, Private Car structure).
- Not converted (warned, flagged on their product's rows): Sub Annexure 2 extra incentive by channel, TW Long Term 5+5, Motor Trade, Standalone CPA, Private Car 3+3 / Long Term Liability, April's 15% / 10% "except Annexure-2" payouts and "Brand New SAOD / SATP".

## Monthly procedure (on top of grid-converter section 2)
Render and look at EVERY page (a scan has no text; if `read_pdf` stops with "this PDF is a scan", transcribe the Motor tables into a `.json` - format in `read_transcription`, example `united_oct26.json` - and say in the report that the numbers were read from images). Add the `CIRCULAR_GROUPS` entry. Run with `--ref`; read every warning; report what the annexures did and did not cover.

## Traps seen
P-35 (class / band labels in merged cells that reach only some rows are carried down); a scanned PDF has no text at all (`read_pdf` stops: use a transcription); a circular with a date `CIRCULAR_GROUPS` does not know gets no RTO groups (rows that need one are not written, a warning says so); Private Car layouts differ per circular (July partition, August / October cc band x states).
- (11 Nov 2026) Structure shifts (11 Nov 2026): the new column's heading is recognised in more wordings (`_is_new_head`: New / Revised / Proposed / Updated / Latest / Maximum + Pay-out / Payout / Commission / Brokerage), and a payout table that has an 'Existing Pay-out' column (or the Private Car table) but whose columns the script cannot find STOPS with expected / found - it used to be left out with a warning. `tests/united_shift_tests.py` mutates the October transcription (headings, table order, an extra table, an extra column at either edge). The four real circulars are byte-identical.

## Open items
April 2026: the master has no group of the "other RTOs" for the 15% (2000-3500 kg) and 10% (3500-7500 kg) payouts (`CIRCULAR_GROUPS` holds `None`); `Brand New SAOD / SATP` (policy type New or Renew is the user's call); the first April build is not accepted.
