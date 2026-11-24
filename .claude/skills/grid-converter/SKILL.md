---
name: grid-converter
description: Convert an insurer's payout grid or commission circular (Excel .xlsx/.xlsb or PDF) into the portal's 42-column rate-config upload CSV, fix a converter when an insurer changes its grid format, build a new converter for a new insurer or product, and teach the system what each run taught. Use this skill whenever the user sends an insurer grid, circular, rates_export/upload CSV or "this month's grid", names Liberty, Magma, Royal Sundaram/RSA, United India or any other insurer in a payout/rate/commission context, mentions grid_tools, RTO or make/model master mapping, "Manual Review Required", the remarks column, /fix-grid, or asks to automate, document, or self-improve this workflow - even if they never say "skill" or "convert". For an insurer we already convert, ALSO load that insurer's own skill (shriram-grid, sbi-grid, icici-grid, liberty-grid, magma-grid, rsa-grid, united-grid, godigit-grid, indusind-grid): it holds what the user has already told us about that insurer.
---

# grid-converter

You turn an insurer's monthly grid into a CSV the portal's Rate Master accepts, and you leave the system a little smarter than you found it.
The scripts and the rules live in the `grid_tools` folder (`D:\harsh\grid_tools`, outside the Django portal repo). This skill is the procedure; the knowledge
is in `AGENT_CONTEXT.md` (in that folder; a packaged copy of this skill carries it as `references/agent_context.md`). Read it - do not work from memory,
because most of what matters here (the user's decisions, the traps that cost a day once) is not guessable.

## 0. Orient (cheap, always)

1. Work in `D:\harsh\grid_tools`. Use `.venv/Scripts/python.exe` for every Python command (forward slashes work in bash and PowerShell).
2. Read `MASTER_SCHEMA.md` (packaged copy: `references/master_schema.md`; the logic of every upload column, what the scripts write today and worked examples) and `AGENT_CONTEXT.md` sections 1-4 (rules, output contract, master lists, remarks protocol) and the playbook of the insurer you are dealing with (section 7).
   Skim section 8 (failure catalogue) if anything looks odd, and section 9 (decisions already made) before asking the user anything.
3. Read the tail of `lessons_learned.md` only if the task is a fix. Do not read a 1,500-line script top to bottom: grep for the function in the error and read that.
   This runs on a Pro plan, so tokens matter: use `grid_doctor.py dump` instead of opening whole sheets.
4. **Load the insurer's own skill** with the Skill tool: `shriram-grid`, `sbi-grid`, `icici-grid`, `liberty-grid`, `magma-grid`, `rsa-grid` (both RSA scripts), `united-grid`, `godigit-grid`, `indusind-grid`
   (`<insurer>` = the first word of the script's file name; `.claude/skills/<insurer>-grid/SKILL.md`). Its Status, **Rules the user has set (do not re-ask)**, Traps seen and Open items are what the user
   has already told us about THAT insurer - read them before you ask the user anything or touch the script. They are short; `AGENT_CONTEXT.md` (section 7 and the decision log) is the long record behind them.
   **Each insurer also has a dedicated reference file in the project root, `<INSURER>_GRID.md`** (`SHRIRAM_GRID.md`, `SBI_GRID.md`, `ICICI_GRID.md`, `LIBERTY_GRID.md`, `MAGMA_GRID.md`, `RSA_GRID.md`, `UNITED_GRID.md`, `GODIGIT_GRID.md`, `INDUSIND_GRID.md`):
   the whole grid and script of that insurer - how the grid is found, how it becomes rows, names, notes, edge cases, how to check a run, the user's decisions, what is open. Read it when the skill is not enough, and keep it true in the same job.

## 1. Decide which job this is

| The user... | Job | Where |
|---|---|---|
| sends this month's grid of an insurer we already convert | **A. new month** | section 2 |
| says it stopped, converts wrongly, or an insurer changed the format | **B. fix** | section 3 (or `/fix-grid`) |
| sends a grid of an insurer or product with no script yet | **C. new script** | section 4 |
| asks what changed, compares with their upload, asks about a number | **D. explain** | section 5 |
| sends the grid + your output + THEIR corrected export ("I fixed these by hand") and wants the script to make those fixes itself | **E. review my output** | section 5b |

Known scripts: Liberty, Magma, RSA GCV 4W comprehensive, RSA Private Car, United India (PDF), Shriram General broker grid (everything except Two-Wheeler), Shriram General two-wheeler grid (New Business + Rollover, fixed rupee payouts as flat cards; SAOD not converted), Go Digit General agency grid (a password-protected workbook: commercial vehicles, two-wheelers, Private Car; needs --password and the package msoffcrypto-tool), SBI General broker grid (one wide rate sheet found by header words, rules read from its notes; the premium slab is asked every month), ICICI Lombard (four workbooks - CV, Private Car, Two-Wheeler, new Two-Wheeler - recognised by their sheets; CV cells are small sentences read as boxes; Council Call cells are listed, never rated), IndusInd General / Reliance General (the commercial-vehicle sheet of the TC grid for brokers: a three-row merged header read by its words; the other product sheets are listed, not converted). Anything else is job C.

## 2. Job A - a new month for a known insurer

1. `python grid_doctor.py check "<grid>"` (Excel). For a PDF there is no doctor: render **every page** to an image and look at it
   with `scripts/render_pages.py <pdf> [--crop-images] [--resolution 200]`. Scanned pages (a cover letter) have no readable text and carry rules; the script only warns about them.
   **If the script stops with "this PDF is a scan"** (every page a picture - the whole October 2026 United circular): read the tables off the page images into a transcription `.json`
   (format: `read_transcription` in `united_commission_to_config.py`; example `regression/inputs/united_oct26.json`), save it with `"source_sha256"` (the SHA-256 of the PDF's bytes) in `transcriptions/` - the script then finds it by itself whenever it is run on that PDF, which is how the user runs it - or pass the `.json` as `--grid`, and say in the report that the numbers were read from images.
   Do not transcribe what you cannot see (annexure pages that are not in the PDF): an annexure that is not in the document is not applied (rule 11 in `AGENT_CONTEXT.md`) - the script writes the payouts that cite it for the whole state(s), flagged - and you tell the user.
2. Ask for the dates if the user did not give them (DD/MM/YYYY), and say which you assumed. Ask for their own upload CSV of the same grid if they have one:
   it is the answer key, and the only way to know the output is right rather than merely plausible.
3. Run the script (`--grid`, `--from-date`, `--to-date`, `--ref` if you have the CSV, `--out output/<NAME>_UPLOAD_<yyyy-mm>.csv`).
   With a keyboard it asks numbered questions about the GRID; without one use `--no-prompt` and the question becomes the error message.
   Never overwrite a file in `output\` the user may have uploaded.
4. Read the summary: rows per product, `MANUAL REVIEW` flags, every `WARNING` (a warning is a table or payout the file does not contain), duplicate keys.
5. Verify independently of the script: pick 3-5 cells in the source and confirm the matching output rows (Liberty x100, Magma already percent, PO = PI - 7).
6. If the circular carries group names that change per month (United), check `CIRCULAR_GROUPS` has the new circular's entry.
7. Report (section 6), then **learn** (section 7).

## 3. Job B - the format changed or the script stops

Follow `/fix-grid` (`.claude/commands/fix-grid.md`): reproduce with `--no-prompt`, copy the `STOPPED - ...` line, diagnose in one sentence, show the diagnosis,
back up the script to `backups\`, make the smallest edit, prove it (re-run, spot-check, `check_regressions.py` = `ALL OK`), report, then learn.
Prefer making the script ask a question over hard-coding this month's answer; a note that cannot be turned into logic needs no code - it is flagged in `remarks`.
If you edit the shared block (between `# %% ---- QUESTIONS` and `# %% ---- READING THE ...`), edit it in ONE script and run `sync_shared.py`.

## 4. Job C - a new insurer or product

Follow `AGENT_CONTEXT.md` section 5.3. In short: look at every page/sheet, get the user's answer-key CSV, copy the nearest script as the template, keep the shared block
identical, reverse-engineer until you reproduce the answer key row for row, list real differences between their upload and the document, then add tests, a regression case,
a layout snapshot and docs - and the entries the user starts it from (`.vscode\launch.json`, `.vscode\tasks.json`, `webapp\src\registry.js` - the id also goes into the expected list of `webapp\test\api.test.js` ('the scripts the page can offer are exactly the registry' fails without it: it did for ICICI, 4 Nov 2026); a script without them is invisible in VS Code). Every new script must create `review = ManualReview()`, call `check_names(df, review, load_masters(ask), log)` and `review.report(df, log)`
before returning - `check_regressions.py` fails a script without the safeguard. One isolated script per insurer/product is intentional: never merge or refactor scripts into a common module.
**A new script also gets its insurer's dedicated reference `<INSURER>_GRID.md` in the project root (copy the nearest insurer's file; `check_regressions.py` fails without it) and its insurer's skill**: `.claude/skills/<insurer>-grid/SKILL.md` (`<insurer>` = the first word of the script's file name; copy the nearest insurer's skill): frontmatter `name` = the folder name and a `description` that names the insurer, the file types and when to load it; then
the sections Status, Rules the user has set (do not re-ask), Traps seen, Open items and Monthly procedure, the run command, and a pointer to its `AGENT_CONTEXT.md` section. Add its name to this skill's description, to section 0 step 4 and to the list in section 7.
`tests/skills_tests.py` and `check_regressions.py` fail a script that has none.

## 5. Job D - explain

Use `--ref <their CSV>`: the comparison reports new / removed / changed rows and 3 examples per differing column. Explain every non-zero number.
The comparison keys on the key columns (state lists as sets, names on the bare cluster) and ignores `remarks` and `id`. Where their upload and the document disagree, the
document wins and you list the differences (decision of 25 Oct 2026).

## 5b. Job E - review my output against their corrections

The user sends the grid, the output you made and THEIR corrected portal export (`rates_export (N).csv`): they fixed rows by hand and want the script to make those fixes by itself next time. The corrections are the user's
rules shown instead of said - this is where the system learns most (Shriram, 4 Nov 2026: 302 flagged rows had been fixed or cleared by hand, and every one became a rule). Load the insurer's skill first (section 0 step 4).

1. **Find out what they really uploaded.** The web app keeps each job's input and output in `webapp\data\jobs\<id>\run-1\in` and `...\out` (the user edits their grid afterwards, so `in\` is the grid as it was run); what they uploaded to the portal is in their month folder `D:\harsh\grids\<MONTH>\UPLOAD`; the files they give you may be copies (`... (2).csv`). Rebuild your old output from the SAME grid, dates and options
   and confirm it equals the file they were given (row count, keys). Until it does you are comparing against a guess.
2. **Diff by key** (`--ref` reports new / removed / changed). A hand edit of ONE column shows up as a removed row plus a new row, so relax one key column at a time (rto, make group, T&C, age, cc ...) to pair them:
   "N rows changed X" is the finding, not "N removed, N new".
3. **Trace each difference to its grid row.** Keep the grid row number on every output row (a throw-away tracer in the scratchpad that wraps `settle_duplicates`), group the differences by grid row and read that row's text
   (UW remarks, REMARKS 1, the notes beside it). The cause is nearly always a sentence the script flagged, or read differently from the user.
4. **Fate table of the flags.** For every `Manual Review Required` note: was its row kept as written, edited, or removed in their file? Kept = a standing note or a rule that changes nothing; edited / removed = a rule to build.
5. **Classify every difference**: (a) a note I flagged and they resolved -> a RULE; (b) a note I read wrongly -> fix the rule; (c) a choice the grid does not show (their group name, their working column) -> use the group they chose when it is in the master, else ask;
   (d) their mistake, or something the document contradicts -> the document wins: list it and ask once (their answer is then a decision row). Build each rule from the WORDS of the note (patterns that use up the words they understand;
   leftover words or numbers stay flagged) - never from this month's row numbers, so next month's rewording is caught too.
6. `grid_doctor.py check` lists new and edited notes: resolve EVERY one, including those on sheets the script never reads.
7. **Prove it**: the rebuilt output equals their file except what is listed as theirs or as the document's; a test for every rule (change the code and watch the test fail); after the user accepts, the month becomes a regression case (`learn.py case`, `--bless` only on their yes) and `grid_doctor.py snapshot`.
8. Report the differences by cause, say plainly which are the document's against theirs, ask the ONE question you need - then learn (section 7).

## 6. The rules that decide most cases (why they exist is in AGENT_CONTEXT.md)

- **Rates only from the document.** Never invent one, never ask the user for one. If a number cannot be read, ask which cell/column to use.
- **Notes you cannot turn into logic are flagged, never guessed, asked about, or allowed to stop the run**: `Manual Review Required: [text]` in `remarks` on every row it may affect.
  The user maps those by hand after clarifying with the insurer. No fallback logic for combinations a note does not mention.
- **The two master workbooks (`master\`) are the authority for RTO group and make/model group names.** A name not in them is flagged with the master's spelling as a hint; never mapped to a lookalike.
  If a master group is incomplete or wrong, flag it; do not work around it.
- **An annexure or exception list that is not in the document is not applied** - never carried over from an earlier circular. The payout it qualifies is written for the whole state(s) and flagged; rows that exist only through the missing annexure are not written.
- `po_* = max(pi_* - 7, 0)`; dates `DD/MM/YYYY`; `+0.01` band convention; `IRDA` = On OD 15.
- **Document problems become numbered questions; document rules you cannot read become flags.** Stop and ask the user (one question) when a column's meaning is unclear,
  a sheet looks like another product, row counts move by more than ~5% unexplained, or the only way forward is a guess.
- **Report the report**: where the file is, rows, dates (and which you assumed), what changed vs last month / their upload, what is flagged and why, what was NOT written, what the
  user must do, and any mistake of your own - plainly. One question at the end, if you need one. Never say "done" before `check_regressions.py` and the tests pass; say what you could not verify.
- Never touch the portal repo, never push, never `--bless` to hide a difference.
- **Documentation and skills are rule-based: no Excel cell references** (the user, 23 Nov 2026). A place in a grid is named by its LABELS (a sheet's title words, a header chain, a row's label) and by the rule that finds it - never by a cell address, a column range,
  a row number or a column letter: the next month's grid moves every one of them. `scripts/cellref_lint.py` lists every such form; `learn.py` refuses text that holds one; `check_regressions.py` and `tests/insurer_docs_tests.py` fail on one.
  **Spot-check by labels**: find the data row by its label cell and the payout column by its header chain (search the sheet for those texts with openpyxl), read the crossing cell and compare it with the upload row (x100 for Liberty, payout = payin - 7).

## 7. Learn (do this before you say "done" - it is how the system improves)

The point of this skill is that run N+1 is easier than run N. After every grid, record what the run taught, using
`python .claude/skills/grid-converter/scripts/learn.py <command>` (from `D:\harsh\grid_tools`; `--help` shows all options):

| What happened | Record it with |
|---|---|
| anything surprising, any fix, any new layout | `learn.py lesson --insurer X --title "symptom" --cause ... --fix ... --covered "test/case" [--rule ...]` -> `lessons_learned.md` |
| a new, verified grid (the user accepted the output) | `learn.py case --name <insurer_monYY> --script <script>.py --from-date .. --to-date .. --input grid=<path> [--input ref=<their csv>] --answers ... --note "..."`, then `--bless` only after the user confirmed |
| a reusable parsing technique | `learn.py note --section pattern --title "..." --text "..."` -> next `P-n` in `AGENT_CONTEXT.md` |
| a trap with symptom -> cause -> rule | `learn.py note --section failure --title "..." --text "symptom -> cause -> rule"` -> next `F-n` |
| a decision the user made | `learn.py note --section decision --title "<scope>" --text "<decision and why>" --date "26 Oct 2026"` |
| something still open | `learn.py note --section open --text "..."` |
| the playbook of an insurer changed (new sheets, new rules, new numbers) | edit section 7 of `AGENT_CONTEXT.md` directly, keep it true, then `learn.py stamp --reason "..."` |
| anything that is true of ONE insurer from now on: a rule the user gave, a trap, a status change (first build / accepted), something still open | `learn.py skill --insurer <shriram/sbi/icici/liberty/magma/rsa/united/godigit/indusind> --section <status/rules/traps/open> --text "..." [--date "4 Nov 2026"]` -> a dated bullet in `.claude/skills/<insurer>-grid/SKILL.md` |
| the grid, the mapping or the rules of an insurer changed | edit that insurer's `<INSURER>_GRID.md` (the full reference; rule-based, no cell addresses) in the same job, then `learn.py skill` for the short memory |

**The insurer skills are the short memory** - the first thing the next run reads (status, the user's rules, the traps, what is open); `AGENT_CONTEXT.md` is the long record. Every run that taught something about an insurer ends with a `learn.py skill` line
for it: a rule the user gives goes into BOTH (`learn.py note --section decision` and `learn.py skill --section rules`), an acceptance or a first build into Status, a still-open question into Open items. When a bullet turns out wrong, edit the skill's text - do not append a contradiction. `learn.py lesson`, `note` and `skill` REFUSE a text that names a place in a grid by a cell address, a row number or a column letter: say it by labels.

Also: `grid_doctor.py snapshot "<grid>" --name <insurer_monYY>` (Excel, after the user accepts), tests in `tests\` for every behaviour you added (assert a selection is **not empty**
before asserting something about it - an empty selection passes vacuously), and the numbers in `CLAUDE.md` / `README.md` kept true.
If what you learned contradicts `AGENT_CONTEXT.md`, correct the old statement - do not append a contradiction. Do not record rates or one-off dates; record patterns, rules, decisions and traps.
Finally run `python check_regressions.py` (ALL OK) and `python tests/run_tests.py` (ALL OK; `--all` if you touched the shared block).
If you touched how a script FINDS a table, sheet or column (anything that reads a header, a sheet name, a row or column position), also run the structure-shift suite for that script: `SHIFT_ONLY=<subject> python tests/structure_shift_tests.py`
(subjects: liberty, rsa_gcv4w, rsa_pvtcar, magma, shriram_broker, shriram_2w, sbi, icici_cv, icici_pc, icici_tw, icici_tw_new, godigit, indusind; `run_tests.py --shift` runs all of them). Its rule (CLAUDE.md "Structure shifts"): a re-laid-out grid gives the SAME upload or
a `STOPPED - expected ... found ...`, never a different file in silence. A new reader: find its anchor by its words, crop at it, read relative to it; a blank row does not end a table; two columns with one heading stop the run.

## 8. Files this skill uses

- `AGENT_CONTEXT.md` - the knowledge base (read it; extend it through `learn.py`).
- `MASTER_SCHEMA.md` - the logic of every column of the upload CSV (the user's schema), what the scripts write today, worked examples and the rules for the agent.
- `lessons_learned.md`, `CLAUDE.md`, `README.md` - history, current state, how a person runs it.
- `scripts/learn.py` - the file mechanics of section 7 (standard library only).
- `.claude/skills/<insurer>-grid/SKILL.md` - one skill per insurance company (shriram, sbi, icici, liberty, magma, rsa, united, godigit, indusind): its status, the user's rules for it, the traps seen, what is open, and the monthly procedure. Load it with this skill; keep it current with `learn.py skill`.
- `<INSURER>_GRID.md` (project root) - one dedicated reference per insurance company: the whole grid and script in one rule-based file; read it for the detail the skill leaves out and edit it when the grid, the mapping or the rules change.
- `scripts/cellref_lint.py` - finds cell addresses, column ranges, row numbers and column letters in the documentation and the skills (they are rule-based); used by `learn.py`, `check_regressions.py` and `tests/insurer_docs_tests.py`.
- `scripts/render_pages.py` - PDF pages -> PNG (with `--crop-images` a scanned table page is cropped to its picture); needs pdfplumber, i.e. the grid_tools venv.
- `*_to_config.py`, `check_regressions.py`, `grid_doctor.py`, `sync_shared.py`, `tests\`, `regression\`, `master\`, `layouts\` - the tools.
