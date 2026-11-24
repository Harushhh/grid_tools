/**
 * The agent: runs a converter for a job and, between the runs, does what a person would otherwise do at the keyboard.
 *
 *   1. Run the converter in service mode.
 *   2. It stopped with a question  -> memory first. A past answer to the same or a very similar question is applied by itself and the converter is run again.
 *                                     Nothing known -> a card asks the person (the closest past answer is offered as a suggestion) and the job waits.
 *   3. It finished                 -> the names the master lacks and the notes it could not apply become review cards. Memory first again: a mapping made before is
 *                                     applied by itself (the run is repeated), a note explained before is shown with the explanation.
 * What the person answers is stored as rules, so the next job needs fewer answers. Everything applied by itself is listed with the rule behind it and can be taken back.
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { now } from './db.js';
import { Memory, foldKey, foldName, foldNote, issueKey, noteKey, normText } from './memory.js';
import { COMMON_INPUTS, getScript, validDate } from './registry.js';
import { runConverter, tail } from './runner.js';

const RTO_SHEET = 'RTO groups to create';
const MAKE_SHEET = 'Make groups to create';
const NOTES_SHEET = 'Manual review notes';
// new_rto_list holds only the cluster names of the grid (the user, 30 Oct 2026): the person maps and groups the RTOs in the portal, so no RTO name is asked about here.
// GRID_RTO_NAMES=group (the converters' earlier behaviour: the master's group names) brings the RTO cards back.
const nameKinds = (config) => (config.rtoNames === 'group' ? ['rto', 'makes'] : ['makes']);
const KIND_LABEL = { rto: 'RTO group', makes: 'make / model group' };
const BLOCKING = new Set(['question', 'password']);   // cards without which the converter cannot run
const MAX_BULK = 500;                                  // notes answered together in one request
const NAME_FLAG = /^(new_rto_list|new_vehicle_makes)\b.*\bnot in the\b/i;      // the converter's flag for a group name the master lacks

export const httpError = (status, message) => Object.assign(new Error(message), { status });

const splitNames = (to) => String(to ?? '').split(',').map((x) => x.trim()).filter(Boolean);
const aliasKey = (kind, name) => `${kind}:${foldName(name)}`;
const noteId = (text) => crypto.createHash('sha1').update(noteKey(text)).digest('hex').slice(0, 20);
const pct = (x) => `${Math.round(x * 100)}%`;
const clip = (s, n = 150) => (String(s).length > n ? `${String(s).slice(0, n - 1).trimEnd()}...` : String(s));
const count = (n) => Number(n || 0).toLocaleString('en-US');
const describe = (answer) => (answer?.labels?.length ? answer.labels.join('; ') : String(answer?.text ?? ''));
const toConverter = (card) => (card.answer?.pick ? { id: card.payload.qid, pick: card.answer.pick } : { id: card.payload.qid, text: card.answer?.text ?? '' });

function parseDate(text) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text);
  return Date.UTC(+m[3], +m[2] - 1, +m[1]);
}

/** The last 'STOPPED - ...' message of a run (what the converter says when it cannot go on). */
function lastStop(res) {
  for (const text of [res.stderr, res.stdout]) {
    const hits = [...String(text || '').matchAll(/STOPPED - ([^\r\n]+)/g)];
    if (hits.length) return hits[hits.length - 1][1].trim();
  }
  return '';
}

function lastLine(text) {
  const lines = String(text || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  return lines.length ? lines[lines.length - 1].slice(0, 400) : '';
}

/** A report small enough to keep in the database; the review workbook has everything. */
function trimReport(report) {
  const sheets = {};
  for (const [name, rows] of Object.entries(report.sheets || {})) sheets[name] = rows.slice(0, 200);
  return { ...report, sheets };
}

/** Passwords the person typed: in memory only, for one job, forgotten after a while and when the server stops. */
export class SecretStore {
  constructor(ttlMs = 30 * 60 * 1000) {
    this.ttl = ttlMs;
    this.map = new Map();
  }

  #sweep() {
    const limit = Date.now() - this.ttl;
    for (const [id, e] of this.map) if (e.at < limit) this.map.delete(id);
  }

  set(jobId, name, value) {
    this.#sweep();
    const e = this.map.get(jobId) || { values: {}, at: 0 };
    e.values[name] = value;
    e.at = Date.now();
    this.map.set(jobId, e);
  }

  get(jobId) {
    this.#sweep();
    const e = this.map.get(jobId);
    if (!e) return {};
    e.at = Date.now();
    return { ...e.values };
  }

  drop(jobId, name) {
    const e = this.map.get(jobId);
    if (e) delete e.values[name];
  }

  clear(jobId) {
    this.map.delete(jobId);
  }
}

class Gate {
  constructor(n) {
    this.free = Math.max(1, n);
    this.waiting = [];
  }

  async acquire() {
    if (this.free > 0) {
      this.free -= 1;
      return;
    }
    await new Promise((resolve) => this.waiting.push(resolve));
  }

  release() {
    const next = this.waiting.shift();
    if (next) next();
    else this.free += 1;
  }
}

export class Agent {
  constructor({ store, memory, masters, config, advisor = null, secrets = new SecretStore() }) {
    this.store = store;
    this.memory = memory;
    this.masters = masters;
    this.config = config;
    this.advisor = advisor;
    this.secrets = secrets;
    this.jobsDir = path.join(config.dataDir, 'jobs');
    this.gate = new Gate(config.maxParallelRuns);
    this.active = new Map();       // job id -> promise of the running loop
    this.again = new Set();        // jobs that must run once more when the current loop ends (something changed while it ran)
  }

  dir(jobId, ...parts) {
    return path.join(this.jobsDir, jobId, ...parts);
  }

  /** Jobs that were running when the server stopped cannot be continued: say so, and let the person run them again. */
  recover() {
    for (const j of this.store.listJobs(1000)) {
      if (j.status === 'running' || j.status === 'queued') {
        this.store.updateJob(j.id, { status: 'failed', error: 'The server was restarted while this job was running. Press Re-run to run it again.' });
        this.store.addEvent(j.id, 'error', 'The server was restarted while this job was running.');
      }
    }
  }

  /** Resolves when no job is running (for tests and for a clean shutdown). */
  async idle() {
    while (this.active.size) await Promise.allSettled([...this.active.values()]);
  }

  // ------------------------------------------------------------------ creating a job
  /** @param files [{ role, path, orig, size }] already on disk (the upload folder); they are moved into the job's folder. */
  createJob({ scriptId, params = {}, files = [], secrets = {} }) {
    const script = getScript(scriptId);
    if (!script) throw httpError(400, 'Unknown converter.');
    const fromDate = String(params.fromDate ?? '').trim();
    const toDate = String(params.toDate ?? '').trim();
    if (!validDate(fromDate)) throw httpError(400, 'from_date must be a real date as DD/MM/YYYY.');
    if (!validDate(toDate)) throw httpError(400, 'to_date must be a real date as DD/MM/YYYY.');
    if (parseDate(toDate) < parseDate(fromDate)) throw httpError(400, 'to_date is before from_date.');
    const options = {};
    for (const o of script.options) {
      const v = params.options?.[o.name];
      if (o.type === 'flag') {
        options[o.name] = v === true || v === 'true' || v === 'on' || v === '1';
      } else if (v !== undefined && v !== null && String(v).trim() !== '') {
        const s = String(v).trim();
        if (o.pattern && !new RegExp(o.pattern).test(s)) throw httpError(400, `${o.label}: that value is not allowed.`);
        options[o.name] = s;
      }
    }
    const allowed = new Map([...script.inputs, ...COMMON_INPUTS].map((i) => [i.role, i]));
    const byRole = new Map();
    for (const f of files) {
      const spec = allowed.get(f.role);
      if (!spec) throw httpError(400, `Unexpected file "${f.role}".`);
      if (byRole.has(f.role)) throw httpError(400, `${spec.label}: only one file, please.`);
      const ext = path.extname(f.orig).toLowerCase();
      if (!spec.accept.includes(ext)) throw httpError(400, `${spec.label}: expected ${spec.accept.join(' or ')}, got ${ext || 'a file without an extension'}.`);
      if (!f.size) throw httpError(400, `${spec.label}: the file is empty.`);
      byRole.set(f.role, f);
    }
    for (const spec of allowed.values()) {
      if (spec.required && !byRole.has(spec.role)) throw httpError(400, `${spec.label}: please choose a file.`);
    }
    // a converter whose grids are all optional (ICICI: any of four workbooks) still needs one of them - a lone "earlier upload" does not count
    if (script.inputs.length && !script.inputs.some((i) => byRole.has(i.role))) throw httpError(400, 'Please choose at least one grid file.');
    const id = this.store.createJob({ script: script.id, params: { fromDate, toDate, options } });
    const inDir = this.dir(id, 'in');
    fs.mkdirSync(inDir, { recursive: true });
    for (const f of byRole.values()) {
      const stored = `${f.role}${path.extname(f.orig).toLowerCase()}`;      // never a name the browser chose
      const target = path.join(inDir, stored);
      try {
        fs.renameSync(f.path, target);
      } catch {
        fs.copyFileSync(f.path, target);
        fs.rmSync(f.path, { force: true });
      }
      this.store.addFile(id, { role: f.role, orig: path.basename(f.orig).slice(0, 200), stored, size: f.size });
    }
    for (const s of script.secrets) if (secrets[s.name]) this.secrets.set(id, s.name, String(secrets[s.name]));
    this.store.addEvent(id, 'info', `Job created for ${script.label}, ${fromDate} to ${toDate}.`);
    return id;
  }

  #fileMap(jobId) {
    const map = {};
    for (const f of this.store.files(jobId)) map[f.role] = this.dir(jobId, 'in', f.stored_name);
    return map;
  }

  // ------------------------------------------------------------------ running
  /** Run (or re-run) the job in the background and return at once. A job that is already running is run once more when it ends. */
  start(jobId) {
    if (this.active.has(jobId)) {
      this.again.add(jobId);
      return 'again';
    }
    this.store.updateJob(jobId, { status: 'queued', error: null });
    const loop = (async () => {
      try {
        do {
          this.again.delete(jobId);
          await this.gate.acquire();
          try {
            await this.#execute(jobId);
          } finally {
            this.gate.release();
          }
        } while (this.again.has(jobId));
      } catch (e) {
        this.#fail(jobId, `Internal error: ${e.message}`);
      } finally {
        this.active.delete(jobId);
      }
    })();
    this.active.set(jobId, loop);
    return 'started';
  }

  #fail(jobId, error) {
    this.store.updateJob(jobId, { status: 'failed', error });
    this.store.addEvent(jobId, 'error', error);
  }

  async #execute(jobId) {
    const { store } = this;
    const first = store.getJob(jobId);
    if (!first) return;
    const script = getScript(first.script);
    if (!script) {
      this.#fail(jobId, `Unknown converter "${first.script}".`);
      return;
    }
    const fileMap = this.#fileMap(jobId);
    for (let round = 1; round <= this.config.maxLoops; round += 1) {
      const job = store.getJob(jobId);
      const run = job.run_count + 1;
      const masters = await this.masters.ensure();
      const cards = store.cards(jobId);
      const answers = cards.filter((c) => c.type === 'question' && (c.status === 'answered' || c.status === 'auto') && c.answer).map(toConverter);
      store.updateJob(jobId, { status: 'running', run_count: run, error: null });
      store.addEvent(jobId, 'run', `Run ${run} started.`);
      const res = await runConverter({
        script, params: job.params, fileMap, secrets: this.secrets.get(jobId), answers,
        aliases: this.#aliasesFor(script, masters, cards), standing: this.#standingFor(script, cards),
        runDir: this.dir(jobId, `run-${run}`), config: this.config,
      });
      const log = tail(`${res.stdout}\n${res.stderr}`.trim(), 8000);
      const asked = res.questions?.questions || [];
      const pending = asked.filter((q) => q.status === 'pending');
      const ok = res.exitCode === 0 && !!res.report && !!res.upload;

      if (!ok && pending.length) {                         // stopped at a question
        const open = await this.#handleQuestions(job, script, pending, run);
        store.updateJob(jobId, { log });
        if (open) {
          store.updateJob(jobId, { status: 'needs_input' });
          return;
        }
        continue;                                          // memory answered it: run again
      }
      if (!ok && script.secrets.length) {                  // a workbook that needs its password
        const said = lastStop(res);
        const secret = script.secrets[0];
        if (/password/i.test(said) && /protected|does not open|wrong/i.test(said)) {
          const had = !!this.secrets.get(jobId)[secret.name];
          this.secrets.drop(jobId, secret.name);
          const payload = { secret: secret.name, label: secret.label, reason: had ? 'The workbook did not accept that password.' : 'The workbook is password protected.' };
          const existing = store.getCard(jobId, 'password', secret.name);
          if (existing) store.updateCard(existing.id, { status: 'open', answer: null, payload, answered_at: null });
          else store.addCard(jobId, { type: 'password', key: secret.name, payload });
          store.addEvent(jobId, 'ask', `Waiting for you: ${payload.reason}`);
          store.updateJob(jobId, { status: 'needs_input', log });
          return;
        }
      }
      if (!ok) {
        const error = res.timedOut ? `The converter ran for more than ${Math.round(this.config.runTimeoutMs / 60000)} minutes and was stopped.`
          : lastStop(res) || lastLine(res.stderr) || `The converter stopped (exit code ${res.exitCode}) without a report.`;
        store.updateJob(jobId, { status: 'failed', error, log });
        store.addEvent(jobId, 'error', error);
        if (!res.timedOut) this.#failureCard(job, script, error);
        return;
      }

      const report = res.report;
      for (const q of pending) store.addEvent(jobId, 'info', `The converter skipped an optional question: ${q.short || q.problem}`);
      store.updateJob(jobId, {
        report: trimReport(report), log,
        outputs: { run, upload: path.basename(res.upload), review: res.review ? path.basename(res.review) : null, rows: report.rows, flagged: report.flagged, at: now() },
      });
      store.addEvent(jobId, 'run', `Run ${run} finished: ${count(report.rows)} rows, ${count(report.flagged)} flagged for manual review.`);
      const { rerun } = this.#review(job, script, report, masters, run);
      if (rerun) continue;                                 // names mapped from memory: write the upload again with them
      this.#settle(jobId);
      this.#cleanRuns(jobId, run);
      return;
    }
    this.#fail(jobId, `The agent stopped after ${this.config.maxLoops} rounds without finishing. Press Re-run to go on.`);
  }

  // ------------------------------------------------------------------ what a run is given
  #aliasesFor(script, masters, cards) {
    const out = { rto: {}, makes: {} };
    const mem = this.memory.aliasFile(script.id);
    for (const kind of nameKinds(this.config)) {
      for (const [from, to] of Object.entries(mem[kind])) {
        const names = splitNames(to);
        if (names.length && names.every((n) => masters.has(kind, n))) out[kind][from] = to;          // a mapping to a group the master no longer has is never applied
      }
    }
    for (const c of cards) {
      if (c.type !== 'alias') continue;
      const { kind, name } = c.payload;
      if (!out[kind]) continue;
      const drop = () => {
        for (const k of Object.keys(out[kind])) if (foldName(k) === foldName(name)) delete out[kind][k];
      };
      if (c.status === 'open' || c.answer?.skip) drop();
      else if (c.answer?.to) {
        drop();
        out[kind][name] = c.answer.to;
      }
    }
    return out;
  }

  /** The notes the converters will not flag in the next run (for the tests and the page). */
  standingList(script, cards) {
    return this.#standingFor(script, cards);
  }

  /** The notes the converter must not flag: everything the person explained or said needs no action - except a note whose card is open (asked again) or was left flagged on purpose. */
  #standingFor(script, cards) {
    const off = new Set(cards.filter((c) => c.type === 'note' && (c.status === 'open' || (c.status === 'answered' && c.answer?.action === 'leave'))).map((c) => foldKey(c.payload.gridNote || c.payload.text)));
    return this.memory.resolved(script.id).map((x) => x.text).filter((t) => !off.has(foldKey(t)));
  }

  // ------------------------------------------------------------------ questions
  async #handleQuestions(job, script, pending, run) {
    const { store } = this;
    let open = 0;
    for (const q of pending) {
      const existing = store.getCard(job.id, 'question', q.id);
      if (existing && existing.status === 'open') {
        open += 1;
        continue;
      }
      if (existing) {                                     // the converter did not take the stored answer: ask again, and not from memory
        store.updateCard(existing.id, { status: 'open', answer: null, rule_id: null, answered_at: null, payload: { ...existing.payload, invalid: true, noAuto: true } });
        store.addEvent(job.id, 'ask', `The converter did not accept the answer to "${clip(q.short || q.problem)}". Asking again.`);
        open += 1;
        continue;
      }
      const payload = { qid: q.id, fp: q.fp, kind: q.kind, problem: q.problem, question: q.question, short: q.short, multi: q.multi, blank_ok: q.blank_ok, options: q.options };
      const m = this.memory.matchQuestion(script.id, q);
      if (m.auto) {
        const a = m.auto.answer;
        const answer = a.pick ? { pick: a.pick, labels: a.pick.map((n) => q.options.find((o) => o.n === n).label) } : { text: a.text };
        store.addCard(job.id, { type: 'question', key: q.id, payload, status: 'auto', answer, ruleId: m.auto.rule.id });
        this.memory.recordUse(m.auto.rule.id, { jobId: job.id, key: q.id, mode: m.auto.mode, score: m.auto.score, detail: describe(answer), run });
        store.addEvent(job.id, 'memory', `"${clip(q.short || q.problem)}" answered from memory (rule #${m.auto.rule.id}, ${m.auto.mode === 'exact' ? 'the same question' : `${pct(m.auto.score)} similar`}): ${clip(describe(answer))}`);
        continue;
      }
      let suggestion = null;
      if (m.suggestion) {
        const a = m.suggestion.answer;
        suggestion = {
          ruleId: m.suggestion.rule.id, score: m.suggestion.score, explanation: m.suggestion.rule.explanation || '', source: 'memory', close: !!m.suggestion.close,
          answer: a ? (a.pick ? { pick: a.pick, labels: a.pick.map((n) => q.options.find((o) => o.n === n).label) } : { text: a.text }) : null,
          was: m.suggestion.rule.data.kind === 'text' ? m.suggestion.rule.data.text : (m.suggestion.rule.data.labels || []).join('; '),
        };
      }
      const id = store.addCard(job.id, { type: 'question', key: q.id, payload, status: 'open', suggestion });
      if (!suggestion) await this.#advise(job, script, store.getCardById(id));
      store.addEvent(job.id, 'ask', `Waiting for you: ${clip(q.short || q.problem)}`);
      open += 1;
    }
    return open;
  }

  /** The optional advisor (a language model) may propose an answer for something memory knows nothing about. It only proposes: the person decides. */
  async #advise(job, script, card) {
    if (!this.advisor) return;
    try {
      const s = await this.advisor.suggest({ script, type: card.type, payload: card.payload });
      if (s) this.store.updateCard(card.id, { suggestion: { ...s, source: 'advisor' } });
    } catch (e) {
      this.store.addEvent(job.id, 'info', `The advisor could not help: ${String(e.message).slice(0, 200)}`);
    }
  }

  /** A stop that is not a question: the person is asked what is different about this file, and what they say is kept for the next time (and for whoever updates the converter). */
  #failureCard(job, script, error) {
    const key = noteId(normText(error));
    if (this.store.getCard(job.id, 'failure', key)) return;
    const m = this.memory.matchFailure(script.id, error);
    this.store.addCard(job.id, {
      type: 'failure', key, payload: { error },
      suggestion: m ? { ruleId: m.rule.id, score: m.score, note: m.rule.explanation || '', source: 'memory' } : null,
    });
    this.store.addEvent(job.id, 'ask', m ? `The converter stopped the way it did before (rule #${m.rule.id}). Your earlier note is on the card.` : 'The converter stopped. Tell the agent what is different about this file.');
  }

  // ------------------------------------------------------------------ after a successful run
  /** Names the master lacks and notes that need a decision become cards; mappings made before are applied by themselves. Returns { rerun } when the upload must be written again. */
  #review(job, script, report, masters, run) {
    const { store, memory } = this;
    let rerun = false;
    const sheet = (name) => report.sheets?.[name] || [];

    for (const u of report.aliases_used || []) {                  // mappings this run applied: keep the trail
      const key = aliasKey(u.kind, u.name);
      if (store.getCard(job.id, 'alias', key)) continue;
      const m = memory.matchAlias(script.id, u.kind, u.name, masters.set(u.kind));
      if (!m.auto) continue;
      store.addCard(job.id, { type: 'alias', key, payload: { kind: u.kind, name: u.name, rows: u.rows }, status: 'auto', answer: { to: u.written }, ruleId: m.auto.rule.id });
      memory.recordUse(m.auto.rule.id, { jobId: job.id, key, mode: 'alias', score: m.auto.score, detail: `${u.name} -> ${u.written}`, run });
      store.addEvent(job.id, 'memory', `Mapped the ${KIND_LABEL[u.kind]} "${u.name}" to "${u.written}" from memory (rule #${m.auto.rule.id}).`);
    }

    const wanted = new Set();
    for (const kind of nameKinds(this.config)) {
      for (const r of sheet(kind === 'rto' ? RTO_SHEET : MAKE_SHEET)) {
        const name = String(r.group_needed ?? '').trim();
        if (!name) continue;
        const key = aliasKey(kind, name);
        wanted.add(key);
        const closest = [...new Set([...masters.search(kind, name, 5), ...String(r.suggestions || '').split(',').map((x) => x.trim()).filter(Boolean)])].slice(0, 8);     // the master's own spelling first
        const payload = { kind, name, rows: Number(r.rows_flagged || 0), hint: String(r.items || ''), closest };
        const card = store.getCard(job.id, 'alias', key);
        if (card) {
          if (card.status === 'open') {
            store.updateCard(card.id, { payload: { ...card.payload, ...payload } });
          } else if (card.answer?.to) {                           // applied, and still unknown: look at it again
            store.updateCard(card.id, { status: 'open', answer: null, rule_id: null, answered_at: null, payload: { ...payload, stale: true, noAuto: true } });
            store.addEvent(job.id, 'info', `The ${KIND_LABEL[kind]} "${name}" is still not in the master after its mapping. Please look at it again.`);
          }
          continue;
        }
        const m = memory.matchAlias(script.id, kind, name, masters.set(kind));
        if (m.auto) {
          store.addCard(job.id, { type: 'alias', key, payload, status: 'auto', answer: { to: m.auto.to }, ruleId: m.auto.rule.id });
          memory.recordUse(m.auto.rule.id, { jobId: job.id, key, mode: 'alias', score: m.auto.score, detail: `${name} -> ${m.auto.to}`, run });
          store.addEvent(job.id, 'memory', `Mapped the ${KIND_LABEL[kind]} "${name}" to "${m.auto.to}" from memory (rule #${m.auto.rule.id}, the same mapping with this month's name).`);
          rerun = true;
          continue;
        }
        const suggestion = m.suggestion ? { ruleId: m.suggestion.rule.id, score: m.suggestion.score, to: m.suggestion.to, source: 'memory' } : null;
        store.addCard(job.id, { type: 'alias', key, payload, status: 'open', suggestion });
        store.addEvent(job.id, 'ask', `The ${KIND_LABEL[kind]} "${name}" is not in the master (${count(payload.rows)} rows).`);
      }
    }

    const noted = new Set();
    const flags = new Map();                                      // one card per grid note, however many flags were made from it
    for (const n of report.notes || sheet(NOTES_SHEET).map((x) => ({ note: x.note, rows_flagged: x.rows_flagged, grid_note: '' }))) {
      const text = String(n.note ?? '').trim();
      if (!text || NAME_FLAG.test(text)) continue;                // "name not in the master": the name cards deal with it
      const gridNote = String(n.grid_note ?? '').trim();
      const key = noteId(gridNote || text);
      const seen = flags.get(key);
      if (seen) seen.rows += Number(n.rows_flagged || 0);
      else flags.set(key, { text, gridNote: gridNote || null, silenceable: !!gridNote, rows: Number(n.rows_flagged || 0) });
    }
    for (const [key, payload] of flags) {
      noted.add(key);
      const card = store.getCard(job.id, 'note', key);
      if (card) {
        if (card.status === 'open') store.updateCard(card.id, { payload: { ...card.payload, ...payload } });
        continue;
      }
      const exact = memory.noteExplanation(script.id, payload.text);
      const similar = exact ? null : memory.noteExplanationForIssue(script.id, payload.text);     // the same issue with other numbers in it
      const known = exact || similar;
      if (known) {
        store.addCard(job.id, { type: 'note', key, payload, status: 'auto', answer: { action: 'explain', explanation: known.explanation || '', ...(similar ? { similarTo: similar.data.text } : {}) }, ruleId: known.id });
        memory.recordUse(known.id, { jobId: job.id, key, mode: 'note', score: similar ? 0.95 : 1, detail: similar ? 'explained before, other numbers' : 'explained before', run });
        store.addEvent(job.id, 'memory', similar
          ? `A note like one you explained before (only the numbers differ) is still flagged (rule #${known.id}).`
          : `A note you explained before is still flagged (rule #${known.id}).`);
        continue;
      }
      store.addCard(job.id, { type: 'note', key, payload, status: 'open' });
      store.addEvent(job.id, 'ask', `A note the converter could not apply (${count(payload.rows)} rows): ${payload.text.slice(0, 120)}`);
    }
    for (const text of report.notes_silenced || []) {                    // notes the person resolved: the converter did not flag them (one card per issue, whatever number it carries)
      const key = noteId(foldNote(text));
      noted.add(key);
      const same = store.cards(job.id).find((c) => c.type === 'note' && foldKey(c.payload.gridNote || c.payload.text) === foldKey(text));
      if (same) {
        noted.add(same.key);
        continue;
      }
      const hit = memory.resolvedFor(script.id, text);
      const action = hit ? hit.action : 'ignore';
      store.addCard(job.id, { type: 'note', key, payload: { text, gridNote: text, silenceable: true, rows: 0 }, status: 'auto', answer: { action, explanation: hit?.rule.explanation || '' }, ruleId: hit ? hit.rule.id : null });
      if (hit) memory.recordUse(hit.rule.id, { jobId: job.id, key, mode: 'note', score: 1, detail: action === 'explain' ? 'explained' : 'no action needed', run });
      store.addEvent(job.id, 'memory', `A note was not flagged because you ${action === 'explain' ? 'explained it' : 'said it needs no action'}${hit ? ` (rule #${hit.rule.id})` : ''}: ${text.slice(0, 120)}`);
    }

    for (const c of store.cards(job.id)) {                           // what is no longer in the report is no longer a question; a stop that no longer happens needs no explaining
      if (c.status !== 'open') continue;
      if ((c.type === 'alias' && !wanted.has(c.key)) || (c.type === 'note' && !noted.has(c.key)) || c.type === 'failure') store.deleteCard(c.id);
    }
    return { rerun };
  }

  /** The status of a job that is not running, from what is still open. */
  #settle(jobId) {
    const job = this.store.getJob(jobId);
    const open = this.store.cards(jobId).filter((c) => c.status === 'open');
    let status;
    if (open.some((c) => BLOCKING.has(c.type))) status = 'needs_input';
    else if (job.outputs) status = open.length ? 'review' : 'done';
    else status = job.status;
    if (status !== job.status) this.store.updateJob(jobId, { status });
  }

  #cleanRuns(jobId, keep) {
    const base = this.dir(jobId);
    try {
      for (const name of fs.readdirSync(base)) {
        if (/^run-\d+$/.test(name) && name !== `run-${keep}`) fs.rmSync(path.join(base, name), { recursive: true, force: true });
      }
    } catch {
      // a file that is still open (Windows) is removed with the next run
    }
  }

  // ------------------------------------------------------------------ the person's answers
  #idle(jobId) {
    const job = this.store.getJob(jobId);
    if (!job) throw httpError(404, 'No such job.');
    if (job.status === 'running' || job.status === 'queued') throw httpError(409, 'The job is running. Wait until it stops.');
    return job;
  }

  #card(job, cardId) {
    const card = this.store.getCardById(cardId);
    if (!card || card.job_id !== job.id) throw httpError(404, 'No such card.');
    return card;
  }

  #scope(script, scope) {
    if (scope === 'global') return 'global';
    if (scope === 'insurer') return `insurer:${script.insurer}`;
    return `script:${script.id}`;
  }

  /** Answer a card. Returns { started, open }: started = the job runs again by itself (it was waiting for exactly this). */
  answer(jobId, cardId, body = {}) {
    const job = this.#idle(jobId);
    const card = this.#card(job, cardId);
    const script = getScript(job.script);
    const b = body || {};
    if (card.type === 'question') this.#answerQuestion(job, script, card, b);
    else if (card.type === 'alias') this.#answerAlias(job, script, card, b);
    else if (card.type === 'note') this.#answerNote(job, script, card, b);
    else if (card.type === 'password') this.#answerPassword(job, card, b);
    else if (card.type === 'failure') this.#answerFailure(job, script, card, b);
    return this.#after(jobId);
  }

  #answerFailure(job, script, card, body) {
    if (body.action === 'dismiss') {
      this.store.updateCard(card.id, { status: 'answered', answer: { action: 'dismiss' }, rule_id: null, answered_at: now() });
      this.store.addEvent(job.id, 'ask', 'You dismissed the note about the stop.');
      return;
    }
    const note = String(body.note ?? '').trim().slice(0, 2000);
    if (!note) throw httpError(400, 'Please write what is different about this file.');
    const ruleId = this.memory.addRule({
      kind: 'failure_note', scope: this.#scope(script, body.scope), matchText: normText(card.payload.error),
      data: { error: clip(card.payload.error, 300) }, explanation: note, sourceJob: job.id,
    });
    this.store.updateCard(card.id, { status: 'answered', answer: { action: 'save', note }, rule_id: ruleId, answered_at: now() });
    this.store.addEvent(job.id, 'ask', `You noted what is different about this file. Remembered as rule #${ruleId}.`);
  }

  #after(jobId) {
    const job = this.store.getJob(jobId);
    const open = this.store.cards(jobId).filter((c) => c.status === 'open');
    if (job.status === 'needs_input' && !open.some((c) => BLOCKING.has(c.type))) {
      this.start(jobId);
      return { started: true, open: open.length };
    }
    this.#settle(jobId);
    return { started: false, open: open.length };
  }

  #answerQuestion(job, script, card, body) {
    const q = card.payload;
    let answer;
    if (q.kind === 'text') {
      const text = String(body.text ?? '').trim();
      if (!text && !q.blank_ok) throw httpError(400, 'Please type an answer.');
      if (text.length > 1000) throw httpError(400, 'That answer is too long.');
      answer = { text };
    } else {
      const pick = (Array.isArray(body.pick) ? body.pick : [body.pick]).map(Number);
      const fits = pick.length > 0 && pick.every((n) => Number.isInteger(n) && q.options.some((o) => o.n === n)) && new Set(pick).size === pick.length && (q.multi || pick.length === 1);
      if (!fits) throw httpError(400, 'Please choose from the options.');
      answer = { pick, labels: pick.map((n) => q.options.find((o) => o.n === n).label) };
    }
    let ruleId = null;
    if (body.remember !== false) {
      ruleId = this.memory.addRule({
        kind: 'answer', scope: this.#scope(script, body.scope), fp: q.fp,
        matchText: normText(`${q.problem} ${q.options.map((o) => o.label).join(' ')}`),
        data: q.kind === 'text' ? { kind: 'text', text: answer.text, problem: clip(q.short || q.problem, 300) } : { kind: 'choose', labels: answer.labels, problem: clip(q.short || q.problem, 300) },
        explanation: String(body.explanation ?? '').trim().slice(0, 1000), sourceJob: job.id,
      });
    }
    this.store.updateCard(card.id, { status: 'answered', answer, rule_id: ruleId, answered_at: now(), payload: { ...q, invalid: false } });
    this.store.addEvent(job.id, 'ask', `You answered "${clip(q.short || q.problem)}": ${clip(describe(answer)) || '(nothing)'}${ruleId ? `. Remembered as rule #${ruleId}.` : '.'}`);
  }

  #answerAlias(job, script, card, body) {
    const { kind, name } = card.payload;
    if (body.skip) {
      this.store.updateCard(card.id, { status: 'answered', answer: { skip: true }, rule_id: null, answered_at: now() });
      this.store.addEvent(job.id, 'ask', `You left the ${KIND_LABEL[kind]} "${name}" as it is.`);
      return;
    }
    const names = splitNames(body.to);
    if (!names.length) throw httpError(400, 'Choose the master group this name stands for.');
    if (names.length > 60) throw httpError(400, 'Too many groups.');
    if (!this.masters.loaded()) throw httpError(409, 'The master lists could not be read, so the name cannot be checked.');
    const bad = names.filter((n) => !this.masters.has(kind, n));
    if (bad.length) throw httpError(400, `Not in the master: ${bad.slice(0, 3).join(', ')}`);
    const to = names.join(',');
    let ruleId = null;
    if (body.remember !== false) {
      ruleId = this.memory.addRule({
        kind: 'alias', scope: this.#scope(script, body.scope), matchText: foldName(name), data: Memory.aliasData(kind, name, to),
        explanation: String(body.explanation ?? '').trim().slice(0, 1000), sourceJob: job.id,
      });
    }
    this.store.updateCard(card.id, { status: 'answered', answer: { to }, rule_id: ruleId, answered_at: now() });
    this.store.addEvent(job.id, 'ask', `You mapped the ${KIND_LABEL[kind]} "${name}" to "${to}"${ruleId ? `. Remembered as rule #${ruleId}.` : ' for this job only.'}`);
  }

  #answerNote(job, script, card, body) {
    const { text, gridNote } = card.payload;
    const action = String(body.action ?? '');
    if (!['ignore', 'explain', 'leave'].includes(action)) throw httpError(400, 'Choose what to do with this note.');
    const explanation = String(body.explanation ?? '').trim().slice(0, 2000);
    let ruleId = null;
    if (action === 'explain' && !explanation) throw httpError(400, 'Please write your explanation.');
    if (action === 'ignore') {                                       // explained or dismissed: the converters stop flagging it (also when only a number in it changes)
      ruleId = this.memory.addRule({ kind: 'standing_note', scope: this.#scope(script, body.scope), matchText: foldKey(gridNote || text), data: { text: gridNote || text }, explanation, sourceJob: job.id });
    } else if (action === 'explain') {
      ruleId = this.memory.addRule({ kind: 'note_explanation', scope: this.#scope(script, body.scope), matchText: foldKey(text), data: { text }, explanation, sourceJob: job.id });
    }
    this.store.updateCard(card.id, { status: 'answered', answer: { action, explanation }, rule_id: ruleId, answered_at: now() });
    const said = { ignore: 'needs no action', explain: 'explained', leave: 'left flagged' }[action];
    this.store.addEvent(job.id, 'ask', `You said a note ${said}${ruleId ? ` (rule #${ruleId})` : ''}: ${text.slice(0, 120)}`);
  }

  /**
   * Answer several note cards of one job with one decision (the same three choices as a single card). All or none: every card is checked before anything is written.
   * 'explain' keeps ONE rule per issue (the same wording apart from numbers - found again by `noteExplanationForIssue`); 'ignore' keeps one rule per exact text, because a
   * standing note silences only that text, never a note with other numbers in it.
   */
  answerNotes(jobId, body = {}) {
    const job = this.#idle(jobId);
    const script = getScript(job.script);
    const b = body || {};
    const ids = [...new Set((Array.isArray(b.cardIds) ? b.cardIds : []).map(Number))];
    if (!ids.length) throw httpError(400, 'Select at least one note.');
    if (ids.length > MAX_BULK) throw httpError(400, `Too many notes at once (at most ${MAX_BULK}).`);
    const action = String(b.action ?? '');
    if (!['ignore', 'explain', 'leave'].includes(action)) throw httpError(400, 'Choose what to do with these notes.');
    const explanation = String(b.explanation ?? '').trim().slice(0, 2000);
    if (action === 'explain' && !explanation) throw httpError(400, 'Please write your explanation.');
    const cards = ids.map((id) => this.#card(job, id));
    for (const c of cards) {
      if (c.type !== 'note') throw httpError(400, 'Only notes can be answered together.');
      if (c.status !== 'open') throw httpError(409, 'One of the notes was already answered. Reload the page.');
    }
    const scope = this.#scope(script, b.scope);
    const ruleOf = new Map();                                    // one rule per issue ('explain') or per exact text ('ignore')
    const db = this.store.db;
    db.exec('BEGIN');
    try {
      for (const c of cards) {
        const { text, gridNote } = c.payload;
        let ruleId = null;
        if (action !== 'leave') {
          const k = foldKey(gridNote || text);                   // one rule per issue: the notes that differ only in a number are one note
          if (!ruleOf.has(k)) {
            ruleOf.set(k, action === 'ignore'
              ? this.memory.addRule({ kind: 'standing_note', scope, matchText: k, data: { text: gridNote || text }, explanation, sourceJob: job.id })
              : this.memory.addRule({ kind: 'note_explanation', scope, matchText: k, data: { text }, explanation, sourceJob: job.id }));
          }
          ruleId = ruleOf.get(k);
        }
        this.store.updateCard(c.id, { status: 'answered', answer: { action, explanation, together: true }, rule_id: ruleId, answered_at: now() });
      }
      const said = { ignore: 'need no action', explain: 'explained', leave: 'left flagged' }[action];
      const rules = [...new Set(ruleOf.values())];
      this.store.addEvent(job.id, 'ask', `You said ${cards.length} note${cards.length === 1 ? '' : 's'} ${said}, all at once${rules.length ? ` (rule${rules.length === 1 ? '' : 's'} ${rules.map((r) => `#${r}`).join(', ')})` : ''}.`);
      db.exec('COMMIT');
    } catch (e) {
      db.exec('ROLLBACK');
      throw e;
    }
    return { answered: cards.length, ...this.#after(jobId) };
  }

  /**
   * Answer several "name not in the master" cards of one job with one decision: leave them all as they are, or map them all to the same master group(s).
   * All or none, like `answerNotes`. A mapping needs names of ONE kind (an RTO group is not a make / model group) and the target must exist in the master.
   * Each name keeps its own rule when it is remembered (a rule maps one name), so next month's cards for these names are answered by themselves.
   */
  answerAliases(jobId, body = {}) {
    const job = this.#idle(jobId);
    const script = getScript(job.script);
    const b = body || {};
    const ids = [...new Set((Array.isArray(b.cardIds) ? b.cardIds : []).map(Number))];
    if (!ids.length) throw httpError(400, 'Select at least one name.');
    if (ids.length > MAX_BULK) throw httpError(400, `Too many names at once (at most ${MAX_BULK}).`);
    const cards = ids.map((id) => this.#card(job, id));
    for (const c of cards) {
      if (c.type !== 'alias') throw httpError(400, 'Only names that are not in the master can be answered together.');
      if (c.status !== 'open') throw httpError(409, 'One of the names was already answered. Reload the page.');
    }
    const skip = !!b.skip;
    let to = '';
    if (!skip) {
      if (new Set(cards.map((c) => c.payload.kind)).size > 1) throw httpError(400, 'RTO groups and make / model groups cannot be mapped to the same group. Select names of one kind.');
      const kind = cards[0].payload.kind;
      const names = splitNames(b.to);
      if (!names.length) throw httpError(400, 'Choose the master group these names stand for.');
      if (names.length > 60) throw httpError(400, 'Too many groups.');
      if (!this.masters.loaded()) throw httpError(409, 'The master lists could not be read, so the names cannot be checked.');
      const bad = names.filter((n) => !this.masters.has(kind, n));
      if (bad.length) throw httpError(400, `Not in the master: ${bad.slice(0, 3).join(', ')}`);
      to = names.join(',');
    }
    const scope = this.#scope(script, b.scope);
    const explanation = String(b.explanation ?? '').trim().slice(0, 1000);
    const remember = !skip && b.remember !== false;
    const db = this.store.db;
    db.exec('BEGIN');
    try {
      for (const c of cards) {
        const { kind, name } = c.payload;
        const ruleId = remember ? this.memory.addRule({ kind: 'alias', scope, matchText: foldName(name), data: Memory.aliasData(kind, name, to), explanation, sourceJob: job.id }) : null;
        this.store.updateCard(c.id, { status: 'answered', answer: skip ? { skip: true, together: true } : { to, together: true }, rule_id: ruleId, answered_at: now() });
      }
      this.store.addEvent(job.id, 'ask', skip
        ? `You left ${cards.length} name${cards.length === 1 ? '' : 's'} as they are, all at once.`
        : `You mapped ${cards.length} name${cards.length === 1 ? '' : 's'} to "${clip(to, 100)}", all at once${remember ? '. Each one is remembered.' : ' (for this job only).'}`);
      db.exec('COMMIT');
    } catch (e) {
      db.exec('ROLLBACK');
      throw e;
    }
    return { answered: cards.length, ...this.#after(jobId) };
  }

  #answerPassword(job, card, body) {
    const value = String(body.value ?? '');
    if (!value) throw httpError(400, 'Please type the password.');
    if (value.length > 200) throw httpError(400, 'That password is too long.');
    this.secrets.set(job.id, card.payload.secret, value);
    this.store.updateCard(card.id, { status: 'answered', answer: { provided: true }, answered_at: now() });
    this.store.addEvent(job.id, 'ask', 'You gave the password. It is kept in memory for this job only and never saved.');
  }

  /** Take an answer back (or an answer the agent gave by itself): the card is asked again, and never answered from memory for this job. */
  reopen(jobId, cardId) {
    const job = this.#idle(jobId);
    const card = this.#card(job, cardId);
    if (card.status === 'open') return this.#after(jobId);
    if (card.type === 'password') this.secrets.drop(jobId, card.payload.secret);
    this.store.updateCard(card.id, { status: 'open', answer: null, rule_id: null, answered_at: null, payload: { ...card.payload, noAuto: true } });
    this.store.addEvent(jobId, 'ask', `You took back the answer to: ${clip(card.payload.short || card.payload.problem || card.payload.name || card.payload.text || card.type)}`);
    this.#settle(jobId);
    return { started: false, open: this.store.cards(jobId).filter((c) => c.status === 'open').length };
  }

  rerun(jobId) {
    const job = this.store.getJob(jobId);
    if (!job) throw httpError(404, 'No such job.');
    if (job.status === 'running' || job.status === 'queued') throw httpError(409, 'The job is already running.');
    if (this.store.cards(jobId).some((c) => c.status === 'open' && BLOCKING.has(c.type))) throw httpError(409, 'Answer the open question first.');
    this.store.addEvent(jobId, 'info', 'Run again, with your answers.');
    return this.start(jobId);
  }

  remove(jobId) {
    const job = this.#idle(jobId);
    this.store.deleteJob(job.id);
    this.secrets.clear(job.id);
    fs.rmSync(this.dir(job.id), { recursive: true, force: true });
  }

  /** Where a finished job's file is (null when it does not have it). */
  outputPath(jobId, which) {
    const job = this.store.getJob(jobId);
    const o = job?.outputs;
    if (!o || !o[which]) return null;
    const file = this.dir(jobId, `run-${o.run}`, 'out', path.basename(o[which]));
    return fs.existsSync(file) ? file : null;
  }
}
