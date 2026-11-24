/** Shared set-up of the tests that run the real converters on the regression grids. */
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Agent } from '../src/agent.js';
import { config as base, GRID_TOOLS } from '../src/config.js';
import { openDb } from '../src/db.js';
import { MasterNames } from '../src/masters.js';
import { Memory } from '../src/memory.js';
import { Store } from '../src/store.js';

export const INPUTS = path.join(GRID_TOOLS, 'regression', 'inputs');
export const BASELINE = (() => {
  try {
    return JSON.parse(fs.readFileSync(path.join(GRID_TOOLS, 'regression', 'baseline.json'), 'utf8'));
  } catch {
    return {};
  }
})();

/** The converters can only be run where the grids and a Python with the packages are. */
export const CAN_RUN = fs.existsSync(path.join(INPUTS, 'liberty_jul26.xlsx')) && fs.existsSync(path.join(GRID_TOOLS, 'liberty_grid_to_config.py'))
  && (!path.isAbsolute(base.python) || fs.existsSync(base.python));

export const sha256 = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

/** A scratch folder that is removed when the test process ends (the runs write tens of MB each). */
const made = [];
process.on('exit', () => {
  for (const d of made) {
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      // still open: the operating system's temp cleaner gets it
    }
  }
});
export function tempDir(prefix) {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  made.push(d);
  return d;
}

export function setup(over = {}) {
  const dataDir = tempDir('gridapp-');
  const config = { ...base, dataDir, ...over };
  const db = openDb(':memory:');
  const store = new Store(db);
  const memory = new Memory(db);
  const masters = new MasterNames(config);
  const agent = new Agent({ store, memory, masters, config, advisor: over.advisor || null });
  return { dataDir, config, db, store, memory, masters, agent };
}

/** A job for a grid file (the file is copied: creating a job moves it). */
export function newJob(env, scriptId, file, params, extra = {}) {
  const tmp = path.join(env.dataDir, 'tmp');
  fs.mkdirSync(tmp, { recursive: true });
  const copy = path.join(tmp, `${crypto.randomBytes(4).toString('hex')}-${path.basename(file)}`);
  fs.copyFileSync(file, copy);
  return env.agent.createJob({ scriptId, params, files: [{ role: 'grid', path: copy, orig: path.basename(file), size: fs.statSync(copy).size }], ...extra });
}

export async function runJob(env, id) {
  env.agent.start(id);
  await env.agent.idle();
  return env.store.getJob(id);
}

export const openCards = (env, id, type) => env.store.cards(id).filter((c) => c.status === 'open' && (!type || c.type === type));
