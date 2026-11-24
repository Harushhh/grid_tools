/** The HTTP side: routes for jobs, cards, the memory (rules) and downloads, plus the static page. */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import multer from 'multer';
import { Agent, SecretStore, httpError } from './agent.js';
import { createAuth } from './auth.js';
import { openDb } from './db.js';
import { MasterNames } from './masters.js';
import { Memory, issueKey } from './memory.js';
import { SCRIPTS, getScript, publicScripts } from './registry.js';
import { Store } from './store.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC = path.resolve(here, '..', 'public');
const JOB_ID = /^[0-9a-f]{16}$/;

const wrap = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

export function createApp({ config, advisor = null, quiet = false }) {
  fs.mkdirSync(config.dataDir, { recursive: true });
  const tmpDir = path.join(config.dataDir, 'tmp');
  fs.rmSync(tmpDir, { recursive: true, force: true });             // half-finished uploads of an earlier run
  fs.mkdirSync(tmpDir, { recursive: true });

  const db = openDb(path.join(config.dataDir, 'app.db'));
  const store = new Store(db);
  const memory = new Memory(db, { autoApply: config.autoApply, suggest: config.suggest });
  const masters = new MasterNames(config);
  const agent = new Agent({ store, memory, masters, config, advisor, secrets: new SecretStore() });
  agent.recover();
  const auth = createAuth({ password: config.appPassword, secret: config.sessionSecret, hours: config.sessionHours, secure: config.secureCookie, insecure: config.insecure });

  const upload = multer({
    storage: multer.diskStorage({ destination: tmpDir, filename: (req, file, cb) => cb(null, crypto.randomBytes(12).toString('hex')) }),
    limits: { fileSize: config.maxUploadMb * 1024 * 1024, files: 8, fields: 40, fieldSize: 20 * 1024 },
  });

  const app = express();
  app.disable('x-powered-by');
  if (process.env.TRUST_PROXY === '1') app.set('trust proxy', 1);

  app.use((req, res, next) => {
    res.set({
      'X-Content-Type-Options': 'nosniff',
      'X-Frame-Options': 'DENY',
      'Referrer-Policy': 'no-referrer',
      'Content-Security-Policy': "default-src 'self'; img-src 'self' data:; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
    });
    if (req.path.startsWith('/api/')) res.set('Cache-Control', 'no-store');
    next();
  });
  if (!quiet) {
    app.use((req, res, next) => {
      const t0 = Date.now();
      res.on('finish', () => console.log(`${new Date().toISOString()} ${req.method} ${req.path} ${res.statusCode} ${Date.now() - t0}ms`));      // never the query, the body or cookies
      next();
    });
  }
  app.use(express.json({ limit: '5mb' }));

  app.get('/healthz', (req, res) => res.json({ ok: true }));
  app.post('/api/login', auth.requireSameSite, auth.login);
  app.post('/api/logout', auth.requireSameSite, auth.logout);
  app.get('/api/me', (req, res) => res.json({ authed: auth.isAuthed(req), required: auth.required }));

  const api = express.Router();
  api.use(auth.requireSession, auth.requireSameSite);
  app.use('/api', api);

  // ------------------------------------------------------------------ what the page needs
  api.get('/scripts', (req, res) => res.json({ scripts: publicScripts(), maxUploadMb: config.maxUploadMb }));

  api.get('/stats', (req, res) => {
    const rules = memory.listRules();
    res.json({
      jobs: store.listJobs(10000).length,
      rules: rules.length,
      enabledRules: rules.filter((r) => r.enabled).length,
      uses: rules.reduce((n, r) => n + r.uses, 0),
      mastersLoaded: masters.loaded(),
    });
  });

  api.get('/masters/search', wrap(async (req, res) => {
    const kind = req.query.kind === 'makes' ? 'makes' : 'rto';
    await masters.ensure();
    res.json({ loaded: masters.loaded(), names: masters.search(kind, req.query.q, 25) });
  }));

  // ------------------------------------------------------------------ jobs
  function cardView(c) {
    const rule = c.rule_id ? memory.getRule(c.rule_id) : null;
    return {
      id: c.id, type: c.type, status: c.status, payload: c.payload, answer: c.answer, suggestion: c.suggestion, created_at: c.created_at, answered_at: c.answered_at,
      issue: c.type === 'note' ? issueKey(c.payload.text) : null,      // notes that differ only in numbers share it: the page groups them
      rule: rule && { id: rule.id, kind: rule.kind, scope: rule.scope, explanation: rule.explanation, enabled: rule.enabled, uses: rule.uses },
    };
  }

  function jobView(id) {
    const job = store.getJob(id);
    if (!job) return null;
    const script = getScript(job.script);
    const cards = store.cards(id).map(cardView);
    const o = job.outputs;
    const r = job.report;
    return {
      id, status: job.status, error: job.error, params: job.params, run_count: job.run_count, created_at: job.created_at, updated_at: job.updated_at,
      script: script ? { id: script.id, label: script.label, insurer: script.insurer } : { id: job.script, label: job.script, insurer: '' },
      files: store.files(id).map((f) => ({ role: f.role, name: f.orig_name, size: f.size })),
      outputs: o && { rows: o.rows, flagged: o.flagged, upload: o.upload, review: o.review, at: o.at },
      report: r && { from_date: r.from_date, to_date: r.to_date, warnings: r.warnings, decisions: r.decisions, products: r.products, sheet_rows: r.sheet_rows },
      cards,
      events: store.events(id),
      applications: store.applications(id),
      pending_changes: o ? cards.filter((c) => c.status === 'answered' && c.answered_at && c.answered_at > o.at).length : 0,
      log: job.status === 'failed' || job.status === 'needs_input' ? job.log : null,
    };
  }

  function jobId(req) {
    const id = String(req.params.id || '');
    if (!JOB_ID.test(id) || !store.getJob(id)) throw httpError(404, 'No such job.');
    return id;
  }

  api.get('/jobs', (req, res) => {
    res.json({
      jobs: store.listJobs(200).map((j) => ({
        id: j.id, status: j.status, script: getScript(j.script)?.label || j.script, params: j.params, created_at: j.created_at, updated_at: j.updated_at,
        rows: j.outputs?.rows ?? null, flagged: j.outputs?.flagged ?? null, open_cards: j.open_cards, error: j.error,
      })),
    });
  });

  api.post('/jobs', upload.any(), wrap(async (req, res) => {
    const files = (req.files || []).map((f) => ({ role: f.fieldname, path: f.path, orig: f.originalname, size: f.size }));
    try {
      const body = req.body || {};
      const options = Object.create(null);
      const secrets = Object.create(null);
      for (const [k, v] of Object.entries(body)) {
        if (k.startsWith('opt_')) options[k.slice(4)] = v;
        else if (k.startsWith('secret_')) secrets[k.slice(7)] = v;
      }
      const id = agent.createJob({ scriptId: String(body.script ?? ''), params: { fromDate: body.fromDate, toDate: body.toDate, options }, files, secrets });
      agent.start(id);
      res.status(201).json({ id });
    } catch (e) {
      for (const f of files) fs.rmSync(f.path, { force: true });
      throw e;
    }
  }));

  api.get('/jobs/:id', (req, res) => res.json(jobView(jobId(req))));

  api.post('/jobs/:id/cards/:cardId/answer', (req, res) => {
    const id = jobId(req);
    res.json({ ...agent.answer(id, Number(req.params.cardId), req.body), job: jobView(id) });
  });

  api.post('/jobs/:id/notes/answer', (req, res) => {
    const id = jobId(req);
    res.json({ ...agent.answerNotes(id, req.body), job: jobView(id) });
  });

  api.post('/jobs/:id/aliases/answer', (req, res) => {
    const id = jobId(req);
    res.json({ ...agent.answerAliases(id, req.body), job: jobView(id) });
  });

  api.post('/jobs/:id/cards/:cardId/reopen', (req, res) => {
    const id = jobId(req);
    res.json({ ...agent.reopen(id, Number(req.params.cardId)), job: jobView(id) });
  });

  api.post('/jobs/:id/rerun', (req, res) => {
    const id = jobId(req);
    agent.rerun(id);
    res.json({ job: jobView(id) });
  });

  api.delete('/jobs/:id', (req, res) => {
    agent.remove(jobId(req));
    res.json({ ok: true });
  });

  api.get('/jobs/:id/download/:which', (req, res) => {
    const id = jobId(req);
    const which = req.params.which === 'review' ? 'review' : req.params.which === 'upload' ? 'upload' : null;
    const file = which && agent.outputPath(id, which);
    if (!file) throw httpError(404, 'That file is not available.');
    res.download(file, path.basename(file));
  });

  // ------------------------------------------------------------------ the memory
  const scopeOk = (scope) => scope === 'global' || SCRIPTS.some((s) => scope === `script:${s.id}` || scope === `insurer:${s.insurer}`);

  api.get('/rules', (req, res) => {
    const enabled = req.query.enabled === '1' ? true : req.query.enabled === '0' ? false : undefined;
    res.json({ rules: memory.listRules({ kind: req.query.kind || undefined, q: req.query.q || undefined, enabled }) });
  });

  api.get('/rules/export', (req, res) => {
    res.set('Content-Disposition', 'attachment; filename="grid-agent-memory.json"');
    res.json({ version: 1, exported_at: new Date().toISOString(), rules: memory.export() });
  });

  api.post('/rules/import', (req, res) => {
    const rules = Array.isArray(req.body) ? req.body : req.body?.rules;
    if (!Array.isArray(rules) || rules.length > 5000) throw httpError(400, 'Expected a file exported from this page.');
    const bad = rules.find((r) => r && r.scope && !scopeOk(r.scope));
    if (bad) throw httpError(400, `Unknown scope in the file: ${String(bad.scope).slice(0, 60)}`);
    res.json({ imported: memory.import(rules) });
  });

  api.patch('/rules/:id', (req, res) => {
    const rule = memory.getRule(Number(req.params.id));
    if (!rule) throw httpError(404, 'No such rule.');
    const b = req.body || {};
    if (b.scope !== undefined && !scopeOk(b.scope)) throw httpError(400, 'Unknown scope.');
    if (b.explanation !== undefined || b.scope !== undefined) memory.updateRule(rule.id, { explanation: b.explanation === undefined ? undefined : String(b.explanation).slice(0, 2000), scope: b.scope });
    if (b.enabled !== undefined) memory.setEnabled(rule.id, !!b.enabled);
    res.json({ rule: memory.getRule(rule.id) });
  });

  api.delete('/rules/:id', (req, res) => {
    if (!memory.getRule(Number(req.params.id))) throw httpError(404, 'No such rule.');
    memory.deleteRule(Number(req.params.id));
    res.json({ ok: true });
  });

  api.get('/rules/:id/uses', (req, res) => {
    const rows = db.prepare('SELECT job_id, key, mode, score, detail, applied_at FROM applications WHERE rule_id = ? ORDER BY id DESC LIMIT 50').all(Number(req.params.id));
    res.json({ uses: rows });
  });

  api.use((req, res) => res.status(404).json({ error: 'Not found.' }));

  // ------------------------------------------------------------------ the page
  app.use(express.static(PUBLIC, { index: 'index.html', maxAge: 0 }));

  app.use((err, req, res, next) => {          // eslint-disable-line no-unused-vars
    if (err?.code === 'LIMIT_FILE_SIZE') return res.status(413).json({ error: `A file is larger than ${config.maxUploadMb} MB.` });
    if (err instanceof multer.MulterError) return res.status(400).json({ error: `Upload problem: ${err.message}` });
    if (err?.type === 'entity.parse.failed') return res.status(400).json({ error: 'That is not valid JSON.' });
    if (err?.type === 'entity.too.large') return res.status(413).json({ error: 'That is too large.' });
    const status = err?.status && err.status < 500 ? err.status : 500;
    if (status === 500) console.error(`${new Date().toISOString()} ${req.method} ${req.path}`, err);
    return res.status(status).json({ error: status === 500 ? 'Something went wrong on the server.' : err.message });
  });

  return { app, agent, store, memory, masters, db, auth };
}
