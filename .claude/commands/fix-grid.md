---
description: Diagnose and fix a grid-converter failure after an insurer changed the grid format
argument-hint: <grid path(s)> [what went wrong / pasted error]
---

An insurer grid no longer converts, or converts differently than expected. Arguments from the user: $ARGUMENTS

Work in this order and keep every step short. `CLAUDE.md` rules apply; they are not repeated here.

1. **Context.** Read `AGENT_CONTEXT.md` (sections 1-4, the insurer's playbook in 7, the failure catalogue in 8) and `lessons_learned.md`. Work out which script applies (Liberty = one .xlsx with sheets such as
   `PC TW_Comp and SAOD` / `CV <7.5T Comp & SATP`; Magma = comprehensive .xlsx with `< =3 lac` plus the SATP .xlsb/.xlsx).
   If the paths or the insurer are not clear from the arguments, ask the user - one question.

2. **What changed in the file.** For each grid path run
   `.venv/Scripts/python.exe grid_doctor.py check "<grid>"`.
   If it reports changes, `dump` the sheets it names. If it says nothing looks like a remembered layout, treat this as a
   new format and say so. (No need to read whole sheets.)

3. **Reproduce.** Run the matching script with `--no-prompt` and the grid path(s), plus `--from-date`/`--to-date` (use the
   dates the user gave; otherwise the first of the current month and 31/03/2030, and say which you used).
   Copy the exact `STOPPED - ...` line into your notes. If it did not stop but the user says the result looks wrong,
   say so and compare the output with the previous month's upload via `--ref`.

4. **Diagnose in one sentence.** Which of these is it?
   - a header was renamed or reworded -> labels are matched by `label_key`; widen the match or add the alias
   - columns or rows moved or were inserted -> lookup must be by label, not position
   - a new kind of cell value (text, symbol, blank pattern) -> new `read_number` handling or a new question
   - a NOTE worded differently, or a new kind of note -> add a pattern for the new wording (and a test); a note that really cannot be
     turned into logic needs NO code: it is flagged in `remarks` ("Manual Review Required: [...]") by `ManualReview` - never guess it,
     never ask about it, never stop for it. Check the flag appears on the right rows instead.
   - a sheet or table appeared, vanished or was renamed -> `resolve_sheet` options / segment table
   - a column was added or removed -> the "missing column" menu, or a new entry in the segment table
   - the *meaning* changed with no visible change -> cannot be seen from the file: ask the user
   Show the user the diagnosis (and the doctor's diff) before editing.

5. **Back up and fix.** Copy the script into `backups\`, then make the smallest edit (grep for the function in the error;
   do not read the whole file). Prefer making the script ask a question over hard-coding this month's answer.
   Shared block edits go into both scripts.

6. **Prove it.**
   - Re-run the failing command; if the script now asks a question, answer it with piped input and record the answers.
   - Spot-check 3-5 rates straight from the Excel against the output CSV (openpyxl; Liberty x100; PO = PI - 7, floor 0).
   - Compare with last month's upload using `--ref` when the user has it; explain every added/removed/changed row count.
   - `.venv/Scripts/python.exe check_regressions.py` must print `ALL OK`. If a known-good case changed, the fix is wrong
     or too broad - revert to the backup and narrow it. Never `--bless` to hide a difference. "ONLY the remarks column differs"
     means the rates are identical and the manual-review flags changed: re-bless only if that is what the fix is meant to do.

7. **Report to the user** in a few lines: the diagnosis, the edit (function + what it does now), the questions the script
   now asks, row counts, the spot-check table, regression result. Say clearly if anything could not be verified.

8. **Only after the user says it is good, learn** (AGENT_CONTEXT.md section 12): `grid_doctor.py snapshot "<grid>" --name <insurer_monYY>`; `python .claude/skills/grid-converter/scripts/learn.py lesson ...` (symptom, cause, fix, covered by); offer the new grid as a regression case with `learn.py case ...` then `--bless`; and if you found a reusable technique, a trap or a decision, add it with `learn.py note --section pattern|failure|decision|open`.

Never write to the portal repository, never push anything, and stop to ask rather than guess a rate.
