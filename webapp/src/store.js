/** Data access for jobs, their files, cards and timeline. Only the columns named here can be updated, so the SQL built below never contains anything from a request. */
import crypto from 'node:crypto';
import { fromJson, now, toJson } from './db.js';

const JOB_FIELDS = {
  status: (v) => v,
  error: (v) => v ?? null,
  run_count: (v) => v,
  log: (v) => v ?? null,
  report: toJson,
  outputs: toJson,
  params: toJson,
};

const CARD_FIELDS = {
  status: (v) => v,
  payload: toJson,
  answer: (v) => (v == null ? null : toJson(v)),
  suggestion: (v) => (v == null ? null : toJson(v)),
  rule_id: (v) => v ?? null,
  answered_at: (v) => v ?? null,
};

function update(db, table, fields, id, patch, stamp = false) {
  const sets = [];
  const vals = [];
  for (const [k, v] of Object.entries(patch)) {
    if (!(k in fields)) throw new Error(`unknown ${table} field ${k}`);
    sets.push(`${k} = ?`);
    vals.push(fields[k](v));
  }
  if (stamp) {
    sets.push('updated_at = ?');
    vals.push(now());
  }
  if (!sets.length) return;
  db.prepare(`UPDATE ${table} SET ${sets.join(', ')} WHERE id = ?`).run(...vals, id);
}

export class Store {
  constructor(db) {
    this.db = db;
  }

  // ------------------------------------------------------------------ jobs
  createJob({ script, params }) {
    const id = crypto.randomBytes(8).toString('hex');
    const t = now();
    this.db.prepare('INSERT INTO jobs (id, script, status, params, created_at, updated_at, run_count) VALUES (?, ?, \'queued\', ?, ?, ?, 0)').run(id, script, toJson(params), t, t);
    return id;
  }

  #job(r) {
    return r && { ...r, params: fromJson(r.params, {}), report: fromJson(r.report), outputs: fromJson(r.outputs) };
  }

  getJob(id) {
    return this.#job(this.db.prepare('SELECT * FROM jobs WHERE id = ?').get(String(id)));
  }

  /** The jobs, newest first, without their reports and logs. */
  listJobs(limit = 200) {
    const open = new Map(this.db.prepare('SELECT job_id, COUNT(*) AS n FROM cards WHERE status = \'open\' GROUP BY job_id').all().map((r) => [r.job_id, Number(r.n)]));
    return this.db.prepare('SELECT id, script, status, params, created_at, updated_at, run_count, error, outputs FROM jobs ORDER BY created_at DESC, id DESC LIMIT ?').all(limit)
      .map((r) => ({ ...r, params: fromJson(r.params, {}), outputs: fromJson(r.outputs), open_cards: open.get(r.id) || 0 }));
  }

  updateJob(id, patch) {
    update(this.db, 'jobs', JOB_FIELDS, id, patch, true);
  }

  deleteJob(id) {
    this.db.prepare('DELETE FROM jobs WHERE id = ?').run(id);
  }

  // ------------------------------------------------------------------ the uploaded files
  addFile(jobId, { role, orig, stored, size }) {
    this.db.prepare('INSERT INTO job_files (job_id, role, orig_name, stored_name, size) VALUES (?, ?, ?, ?, ?)').run(jobId, role, orig, stored, size);
  }

  files(jobId) {
    return this.db.prepare('SELECT role, orig_name, stored_name, size FROM job_files WHERE job_id = ? ORDER BY id').all(jobId);
  }

  // ------------------------------------------------------------------ cards
  #card(r) {
    return r && { ...r, payload: fromJson(r.payload, {}), answer: fromJson(r.answer), suggestion: fromJson(r.suggestion) };
  }

  cards(jobId) {
    return this.db.prepare('SELECT * FROM cards WHERE job_id = ? ORDER BY id').all(jobId).map((r) => this.#card(r));
  }

  getCard(jobId, type, key) {
    return this.#card(this.db.prepare('SELECT * FROM cards WHERE job_id = ? AND type = ? AND key = ?').get(jobId, type, String(key)));
  }

  getCardById(id) {
    return this.#card(this.db.prepare('SELECT * FROM cards WHERE id = ?').get(Number(id)));
  }

  addCard(jobId, { type, key, payload, status = 'open', answer = null, ruleId = null, suggestion = null }) {
    const info = this.db.prepare('INSERT INTO cards (job_id, type, key, payload, status, answer, rule_id, suggestion, created_at, answered_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)')
      .run(jobId, type, String(key), toJson(payload), status, answer == null ? null : toJson(answer), ruleId, suggestion == null ? null : toJson(suggestion), now(), status === 'open' ? null : now());
    return Number(info.lastInsertRowid);
  }

  updateCard(id, patch) {
    update(this.db, 'cards', CARD_FIELDS, id, patch);
  }

  deleteCard(id) {
    this.db.prepare('DELETE FROM cards WHERE id = ?').run(id);
  }

  // ------------------------------------------------------------------ the timeline and the audit trail
  addEvent(jobId, kind, text) {
    this.db.prepare('INSERT INTO events (job_id, at, kind, text) VALUES (?, ?, ?, ?)').run(jobId, now(), kind, String(text).slice(0, 2000));
  }

  events(jobId) {
    return this.db.prepare('SELECT at, kind, text FROM events WHERE job_id = ? ORDER BY id').all(jobId);
  }

  applications(jobId) {
    return this.db.prepare(`SELECT a.id, a.rule_id, a.key, a.mode, a.score, a.detail, a.run, a.applied_at,
                                   r.kind AS rule_kind, r.scope AS rule_scope, r.explanation AS rule_explanation, r.enabled AS rule_enabled
                            FROM applications a LEFT JOIN rules r ON r.id = a.rule_id WHERE a.job_id = ? ORDER BY a.id`).all(jobId);
  }
}
