---
name: magma-grid
description: Magma knowledge for grid_tools - the Magma comprehensive grid (.xlsx, sheet "< =3 lac") plus the SATP grid (.xlsb or .xlsx, sheet "<=2 Lacs") converted to MAGMA_UPLOAD_<yyyy-mm>.csv by magma_grid_to_config.py. Load it with grid-converter whenever the user sends a Magma grid or SATP file, a Magma re-upload / rates_export to review, or asks about UW clusters, the TW CPA cut, the 3718 make group, or what the Magma script does.
---

# Magma - comprehensive + SATP grids

Dedicated reference: `MAGMA_GRID.md` (this insurer's whole grid and script in one rule-based file - read it when you need more than this short memory; keep it true in the same job, then `learn.py skill`).

Detail: `AGENT_CONTEXT.md` section 7.2, `CLAUDE.md` (the Magma bullet).

## Status and checks
Regression `magma_oct26` (39,102 rows, 0 flagged); all 114 clusters of the October grid exist in the master and equal the user's re-upload. Tests: `tests/magma_interactive_tests.py` (in `--all`).

## Run it
```
.venv/Scripts/python.exe magma_grid_to_config.py --grid "<comprehensive>.xlsx" --satp "<SATP>.xlsb" --from-date DD/MM/YYYY --to-date DD/MM/YYYY --ref "<their re-upload>.csv" [--no-prompt]
```
Two input files: the comprehensive workbook and the SATP workbook (`.xlsb` needs pyxlsb). Output `output\magma\MAGMA_UPLOAD_<yyyy-mm>.csv`.

## How the grid becomes rows
- Sheets: `< =3 lac` (comprehensive; clusters down the side under "UW Cluster/Outgo on", segments across under a two-row header), `<=2 Lacs` (SATP); `Generic Guidelines` / `Notes` / `Note` sheets hold rules (flagged if new).
- **Cells are already percent** (unlike Liberty). `IRDA` -> On OD 15; a 2W with CPA = Yes gets `TW_CPA_CUT` = 2 (IRDA 15 -> 13); Private Car is always On OD with NCB Yes / No and "non-diesel = petrol" (Petrol, Electric, CNG, LPG share the Petrol column).
- Segments in `motor_segments()` / `satp_segments()`: GCV 4W by GVW with age bands for 12-20T and 20-40T (age < 5 / >= 5), GCV > 40T and the "3718" make group (`Magma_3718_APR26`, `Magma_EX_3718_APR26`), GCV 3W, PCV 3W (Electric / New / Old with T&C), PCV bus / taxi, Misc-D (tractor, harvester, CE, garbage), TW (1+1, 1+5, SAOD x CPA No / Yes x bike cc bands + scooter), Private Car (1+1, 1+3, SAOD).
- `new_rto_list`: the cluster name by default; with `GRID_RTO_NAMES=group` it is `MAGMA_APR26_` + cluster (spaces -> `_`). Columns the script does not map (`unused_labels`) warn when they carry non-zero rates; SATP columns known to carry no upload rows are in `SATP_NOT_UPLOADED`.

## Rules the user has set (do not re-ask)
- Clusters added or removed between months must NOT stop the run. A new or odd cell value = a question with clear options, never a crash and never a guess. Text rules the script cannot apply go to `remarks` (`Manual Review Required`).
- Magma `id` = 1..N (Liberty's is 1 on every row).

## Monthly procedure (on top of grid-converter section 2)
Ask for BOTH files; `grid_doctor.py check` each; run with `--ref` = their last re-upload; read the `unused_labels` warnings (a new column with rates is a missing segment); spot-check cells (already percent, payout = payin - 7).

## Traps seen
A column with non-zero rates that no segment reads is a MISSING SEGMENT, not noise (`unused_labels` warns; `SATP_NOT_UPLOADED` lists the SATP columns known to carry no upload rows). The "3718" make groups are dated (`Magma_3718_APR26`, `Magma_EX_3718_APR26`). Magma cells are already percent - do not multiply (Liberty's are fractions).
- (11 Nov 2026) Structure shifts (11 Nov 2026): the table starts at its 'UW Cluster' cell wherever it is (`find_uw_header`, `Table.col0` keeps the Excel column in cell addresses); the premium slab of a sheet is read from its name's WORDS (`slab_of`: '<=3 Lacs v2' is the '< =3 lac' slab), never from its spelling; an alphabetical cluster list that goes on after a blank row is read to its end (`_continues`, said in the decisions). Every sheet called 'SheetN' stops (the slab is only in the name) - by design.

## Open items
None recorded in `AGENT_CONTEXT.md` 7.2.
