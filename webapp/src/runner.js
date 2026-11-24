/**
 * Starts one converter script as a child process, in service mode (GRID_SERVICE=1), and collects what it leaves behind:
 * the questions it could not ask (questions.json), its report (report.json), the upload CSV and the review workbook.
 * The command line is built only from the registry; file paths are the app's own storage paths; the password goes in an environment variable.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { PORTAL_ENV, SECRET_ENV } from './config.js';
import { COMMON_INPUTS, SCRIPTS, outputName } from './registry.js';

const KEEP = 600 * 1024;        // the most output kept per stream
const CONVERTER_SECRET_ENV = new Set(SCRIPTS.flatMap((s) => s.secrets.map((x) => x.env)));

function tail(text, n = 4000) {
  return text.length > n ? text.slice(-n) : text;
}

function readJson(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
}

/** The arguments of one run (exported for the tests). */
export function buildArgs(script, params, fileMap, outCsv, config) {
  const args = ['-I', '-X', 'utf8', path.join(config.gridTools, script.file), '--no-prompt', '--from-date', params.fromDate, '--to-date', params.toDate, '--out', outCsv];
  for (const input of [...script.inputs, ...COMMON_INPUTS]) {
    if (fileMap[input.role]) args.push(input.flag, fileMap[input.role]);
  }
  for (const opt of script.options) {
    const v = params.options?.[opt.name];
    if (opt.type === 'flag') {
      if (v === true) args.push(opt.flag);
    } else if (v !== undefined && v !== null && String(v) !== '') {
      args.push(opt.flag, String(v));
    }
  }
  return args;
}

/**
 * @returns {Promise<{exitCode, timedOut, stdout, stderr, questions, report, upload, review, runDir}>}
 */
export function runConverter({ script, params, fileMap, secrets = {}, answers = [], aliases = { rto: {}, makes: {} }, standing = [], runDir, config }) {
  fs.mkdirSync(path.join(runDir, 'out'), { recursive: true });
  const at = (name) => path.join(runDir, name);
  fs.writeFileSync(at('answers.json'), JSON.stringify({ answers }));
  fs.writeFileSync(at('aliases.json'), JSON.stringify(aliases));
  fs.writeFileSync(at('standing.json'), JSON.stringify(standing));
  for (const f of ['questions.json', 'report.json']) fs.rmSync(at(f), { force: true });
  const outCsv = path.join(runDir, 'out', outputName(script, params.fromDate));
  const args = buildArgs(script, params, fileMap, outCsv, config);
  const env = {
    ...process.env,
    PYTHONIOENCODING: 'utf-8',
    GRID_SERVICE: '1',
    ...(config.rtoNames ? { GRID_RTO_NAMES: config.rtoNames } : {}),
    GRID_QUESTIONS_FILE: at('questions.json'),
    GRID_ANSWERS_FILE: at('answers.json'),
    GRID_REPORT_FILE: at('report.json'),
    GRID_ALIASES_FILE: at('aliases.json'),
    GRID_STANDING_NOTES_FILE: at('standing.json'),
  };
  for (const k of SECRET_ENV) delete env[k];
  for (const k of PORTAL_ENV) delete env[k];                       // the Django portal is another project: its settings are not a converter's business
  for (const k of CONVERTER_SECRET_ENV) delete env[k];            // never a password inherited from the server's own environment ...
  for (const s of script.secrets) if (secrets[s.name]) env[s.env] = secrets[s.name];     // ... only the one the person typed for this job
  return new Promise((resolve) => {
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    let child;
    try {
      child = spawn(config.python, args, { cwd: runDir, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });      // cwd = the run's folder: a crash report lands there
    } catch (e) {
      resolve({ exitCode: -1, timedOut: false, stdout: '', stderr: `cannot start ${config.python}: ${e.message}`, questions: null, report: null, runDir });
      return;
    }
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, config.runTimeoutMs);
    child.stdout.on('data', (d) => {
      stdout = tail(stdout + d.toString('utf8'), KEEP);
    });
    child.stderr.on('data', (d) => {
      stderr = tail(stderr + d.toString('utf8'), KEEP);
    });
    const finish = (exitCode) => {
      clearTimeout(timer);
      const report = readJson(at('report.json'));
      const upload = report?.csv && fs.existsSync(path.join(runDir, 'out', report.csv)) ? path.join(runDir, 'out', report.csv) : null;
      const review = report?.review && fs.existsSync(path.join(runDir, 'out', report.review)) ? path.join(runDir, 'out', report.review) : null;
      resolve({ exitCode, timedOut, stdout, stderr, questions: readJson(at('questions.json')), report, upload, review, runDir });
    };
    child.on('error', (e) => {
      stderr += `\n${e.message}`;
      finish(-1);
    });
    child.on('close', (code) => finish(code));
  });
}

export { tail };
