# Grid agent - the converters as a web application

Upload an insurer's grid, run its converter, download the upload CSV. When the converter meets something it cannot decide, you get a **card** instead of a failure.
What you answer is stored in a **memory** (a registry of rules). The next time the same or a very similar situation comes up, the agent **uses your earlier answer by itself** and
tells you it did.

```
 browser ──►  Express (auth, uploads, API)  ──►  Agent loop  ──►  python <insurer>_to_config.py   (service mode)
   ▲                 │                              │  ▲                  │
   │  cards          ▼                              ▼  │                  ▼
   └──────────  SQLite: jobs, cards, timeline    Memory (rules)     questions.json / report.json / CSV / review workbook
```

## What it does, against the four requirements

| Requirement | How |
|---|---|
| 1. Upload, run, download | `New run` takes the converter, `from_date`/`to_date` and the files; the selected job offers the upload **CSV** and the review workbook. |
| 2. Cards instead of failures | A question the converter cannot ask (which sheet? which spelling? a password?) is written to a file by the converter (service mode) and shown as a card. The job waits; nothing fails. After a run, **names the master lacks** and **notes the converter could not apply** become cards too. A real stop (an unreadable file) gets a card as well, asking what is different - honestly not a fix, since only an updated converter can get past it. |
| 3. A memory (its "brain") | Every answer becomes a rule: `answer`, `alias` (name mapping), `standing_note` (no action needed), `note_explanation`. Rules have a scope (this converter / this insurer / everything), your reason, use counts, an on/off switch, export and import. |
| 4. Autonomous resolution | **Memory first.** For each question the agent looks for a rule: the same question (fingerprint) or a very similar one (word and trigram similarity >= 0.9) *whose chosen option is still there with the same label* is applied and the converter is run again. Between 0.65 and 0.9 it is shown as a suggestion. Everything applied by itself is listed on the job (rule, match, reason) and has an "Ask me instead" button. |

## The page: one screen, three columns

Like a feed with a rail on each side, so nothing has to be left to look at something else:

| Column | What is in it |
|---|---|
| **Jobs** (left) | Every job, newest first: converter, period, status, rows / flagged and how many things need you (the browser tab's title shows that count too). It refreshes by itself while a job is running. **+ New run** opens the form. |
| **New run**, then the **selected job** (middle) | "New run" is a bar that opens in place; under it the job you picked: its cards, the result and downloads, what the agent did. The form is kept while you look at a job, so a half-filled form is not lost; it is cleared after Run and when you sign out. |
| **Memory** (right) | The rules the agent knows: search, filter by kind, switch off / delete, export / import. It refreshes when you answer a card. |

The address only says what is selected, so the old links still work: `#/` (nothing selected, New run open), `#/jobs/<id>` (that job), `#/jobs` and `#/memory` (the same page; on a narrow screen they pick
the Jobs / Memory tab). Below 1200 px (a phone, a small window) one column is shown at a time and a bar at the bottom (Jobs / Run / Memory) switches between them. Only the sign-in screen is separate.
`public/app.js` has no framework and no build step; the three columns are built once (`buildShell`) and then updated in place, and a slow answer for a job you have already left can never paint over the one you chose after it.

## Why the converters were not rewritten in JavaScript

They are the verified part: every converter reproduces its baseline byte for byte (`check_regressions.py`). The web app runs them as they are, as child processes
(`python -I -X utf8 <script> --no-prompt ...`), so the web path and the command line cannot drift apart - the end-to-end tests compare the downloaded CSV with the command line's
baseline hash. The converters get a *service mode* (environment variables only; without them they behave exactly as before): a question is written to `GRID_QUESTIONS_FILE`
instead of being asked, answers come from `GRID_ANSWERS_FILE`, name mappings from `GRID_ALIASES_FILE`, extra "no action needed" notes from `GRID_STANDING_NOTES_FILE`, and the
run's report is written to `GRID_REPORT_FILE`. That block is shared and identical in every script (`sync_shared.py`, `tests/service_tests.py`).

## Boundaries: a project of its own

The Grid agent is separate from the Django insurance portal. Treat them as two projects: they share no code, no database, no settings file, no Python environment and no server process.

* Its own folder (`webapp/`), its own `package.json` and `node_modules`, its own `.env` (never copy the portal's into it), its own database (SQLite in `DATA_DIR`).
* Its engine is the converter scripts in the folder above (`GRID_TOOLS_DIR`) with *their* `.venv` - never the portal's `venv`. The converters are a third piece of the picture: the app starts them as programs and exchanges files with them, nothing more.
* It reads only its own environment variables (the table below) and removes the portal's settings (`PORTAL_ENV` in `src/config.js`) from the environment of every converter it starts, even if they are in the shell it was started from.
* It makes no outgoing calls of its own except the optional advisor, and it never imports or opens anything of the portal.

Today there are exactly two hand-offs, and both are a person moving a file: the master lists exported from the portal into `../master/`, and the upload CSV that a person uploads in the portal.
If the two ever need to exchange data directly, that is built on purpose as one API client module in this app (with its own settings, for example an address and a key), against a key-protected
endpoint of the portal (it already has some, under `/api/v1/`) - and nothing else in either project reaches into the other.
`test/isolation.test.js` fails the moment this app starts to depend on the portal (a mention in code, a setting read, an outgoing call), so a link cannot appear by accident.

## Run it

Needs Node 24+ and the grid_tools Python environment (`.venv`, see the parent README). The converters, `master/` (RTO and make-model exports) and `.venv` are found in the folder above `webapp/`.

```bash
cd webapp
npm install
set GRID_APP_PASSWORD=choose-a-password-of-8-or-more-characters      # PowerShell: $env:GRID_APP_PASSWORD="..."
npm start                                                            # http://localhost:3000
```

On Windows you can instead copy `.env.example` to `.env`, put your password in it, and double-click `start.bat`.
For local development only, `GRID_APP_INSECURE=1` removes the password and listens on 127.0.0.1 only.
The list of converters is read when the server starts: after a converter is added to `src/registry.js`, stop the app (Ctrl+C in its window) and start it again, then reload the page.
Finished jobs (their uploads, outputs and logs) stay in `DATA_DIR\jobs` until you delete them (the Delete button of the job).

| Variable | Default | |
|---|---|---|
| `GRID_APP_PASSWORD` | - | the one password of the app (>= 8 characters); the server refuses to start without it |
| `GRID_APP_INSECURE` | - | `1` = no password, 127.0.0.1 only (development) |
| `PORT`, `HOST` | `3000`, `0.0.0.0` | |
| `DATA_DIR` | `webapp/data` | database, uploads, outputs - **put this on a persistent volume** |
| `GRID_TOOLS_DIR` | the folder above | where the `*_to_config.py`, `master/` and `.venv` are |
| `PYTHON` | `.venv`'s python | |
| `SESSION_SECRET`, `SESSION_HOURS` | random per start, 12 | set the secret to stay signed in across restarts |
| `COOKIE_SECURE`, `TRUST_PROXY` | - | `1` behind https / a reverse proxy |
| `MAX_UPLOAD_MB`, `MAX_PARALLEL_RUNS`, `RUN_TIMEOUT_MS` | 150, 2, 20 min | |
| `MEMORY_AUTO_APPLY`, `MEMORY_SUGGEST` | 0.9, 0.65 | the similarity thresholds |

## A job, step by step

`queued` -> `running` -> one of
* `needs_input` - the converter is blocked by a question or a password. Answer the card; when nothing blocking is left the job continues by itself.
* `review` - the upload is ready (download it already), and there are names / notes to look at. Answer them and press **Apply and run again**.
* `done` - nothing needs you.
* `failed` - the converter said why (e.g. "the workbook is not the layout this script expects"); the technical output is under the job.

**Several cards at once.** When a job has two or more note cards, or two or more "name not in the master" cards, a bar above them lets you tick the cards (one by one, "Select all N like this", or
by part of a name) and give them all one answer. Notes that say the same apart from numbers are one issue (`issueKey`): they sit together, and an explanation you give for one is shown for the same
issue next month with other numbers in it. "No action needed" still matches the exact text only. The answer is all-or-none (`POST /api/jobs/:id/notes/answer`, `.../aliases/answer`).

**RTO names are not asked about.** The converters write only the cluster name in `new_rto_list` (the user's rule of 30 Oct 2026, see `CLAUDE.md` in the folder above), so there are no "RTO group is not in the
master" cards; only make / model names can still produce a name card. `GRID_RTO_NAMES=group` in the server's environment brings the earlier behaviour (the master's group names, and the RTO name cards) back.

The agent runs the converter up to 14 times per press (each question memory answers is one more run). Passwords are only ever held in the server's memory, for that job,
and forgotten 30 minutes after their last use or when the server stops; they are handed to the converter in an environment variable and are never written to the database, a file or the log (tested).

## The memory, precisely

* **Answers** are found by the question's *fingerprint* (its wording without numbers and file names) or by similarity. A rule is applied by itself only if the option you chose
  is still there under the same label (capitals and spaces aside; a month written in the label - `OCT26` / `NOV26` - is not a difference, unless two options then look alike).
  A reordered list still works. A label that merely looks close (one number or letter differs: `< 1000 CC` / `< 1500 CC` may be another column) is **not** applied: you are asked,
  with the closest option offered as a suggestion to check.
* **Name mappings** (`GO-DIGIT_OCT26_HCV_Mumbai` -> master groups) are applied for the same name apart from capitals, stray underscores or the **month in the name**
  (`OCT26` -> `NOV26`), and only if every target group exists in the current master. A merely *similar* name is only suggested.
* **"No action needed" notes** silence exactly that text (spaces, capitals and punctuation ignored) - never a longer rule that happens to contain it. Flags that a converter makes
  about its own checks cannot be silenced, only explained; the page does not offer a button that would do nothing.
* **Stops**: when a converter stops on something it cannot decide (an unreadable file, a layout it does not know), the job still gives you a card: say what is different about this
  file. The note is kept and shown on the card the next time the converter stops the same way (>= 0.8 similar message). It never changes a run - nothing can get past such a stop
  except an updated converter - but it is exported with the memory for whoever updates it.
* Scope: a rule made on a card is for **this converter** unless you widen it to the insurer or to everything.
* Every use is recorded (`applications`): the job shows which rule, how close the match was, and your own reason. "Stop using this rule" switches it off; the Memory column switches rules on and off, deletes,
  exports (JSON) and imports them.

## Security

One shared password (constant-time check, 8 attempts per 15 minutes per address), a signed HttpOnly `SameSite=Strict` cookie, a custom header + origin check on every state-changing
request, a content-security-policy without inline script, no request bodies in the log. Uploads go to a temporary folder, are checked against the converter's accepted extensions and
stored under fixed names (never the browser's file name). Only the registered scripts (the list in `src/registry.js`) can be started, with arguments built from the registry - nothing from the browser is part
of a command line. Job ids are 16 hex characters and looked up in the database before they touch a path. This is an internal tool for a trusted team; there is one password, not users.

## Tests

```bash
npm run test:fast     # memory, HTTP layer, isolation from the portal: seconds
npm test              # everything, including the end-to-end tests that run the real converters on the regression grids (a few minutes)
```

The end-to-end tests need the regression grids in `../regression/inputs` and the Python environment; without them they are skipped.
They check, among other things: the downloaded CSV is byte-identical to the baseline of `check_regressions.py`; a second job is answered from memory; a wrong, then the right password
of an encrypted workbook; the password appears in no file or response; a name mapping and a "no action needed" note are applied by themselves on the next upload.

## Docker

`../Dockerfile` builds Node 24 + Python with the converters. **It has not been built or run** (there is no Docker in the environment this was written in) - treat it as a starting point.
Mount a volume at `/data`, and the master exports at `/app/master` so a new export does not need a new image.

## Known limits

* One shared password, not user accounts: the timeline says what the agent did, not who clicked. One server process (SQLite, jobs run on the same machine; `MAX_PARALLEL_RUNS` at a time).
* The database is Node's built-in `node:sqlite`, which Node 24 still calls experimental (the app uses only prepared statements; `db.js` is the one place to swap in another driver).
* The master lists are the two workbooks in `master/` (the portal's exports), read at start and whenever they change on disk; the app does not fetch them from the portal. Update them
  when the portal's groups change, or name mappings are checked against stale lists.
* A converter that stops on an unknown layout still needs a developer to update it (the stop card only records what you saw). The app makes the existing scripts' questions answerable
  in a browser; it does not write new converter code.
* Not verified here: the Docker image (no Docker where this was written), https behind a proxy, the advisor against the real API, many simultaneous people.

## Adding or changing a converter

1. The script needs the shared *service mode* block (it comes with `sync_shared.py`) and the `--no-prompt`, `--from-date`, `--to-date`, `--out` arguments.
2. Add it to `src/registry.js` (files it takes, options, secrets). That list is the only way a script can be started.
3. Run `python check_regressions.py` and `python tests/service_tests.py`.

## Optional: a language-model advisor

Off by default, and nothing depends on it. With `ADVISOR=1` and `ANTHROPIC_API_KEY` set, a card that memory knows nothing about may carry a *suggested* answer from a model.
It uses the paid API (billed per call). **Put the key only in `webapp\.env`** (read by `npm start` for that one process), never in the PC's environment variables: with
`ANTHROPIC_API_KEY` set system-wide, Claude Code bills that key instead of the subscription.
It is labelled as such, is never applied by itself, and is never given a password or the grid file - only the question text and its options. (Written against the public Messages API and
tested with a stand-in server only; it has not been run against the real API.)
