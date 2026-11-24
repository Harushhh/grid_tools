/** The runner: what a converter is started with (arguments from the registry only; secrets only where a script declares them), checked with a throw-away script. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { config as base } from '../src/config.js';
import { SCRIPTS, getScript } from '../src/registry.js';
import { buildArgs, runConverter } from '../src/runner.js';
import { tempDir } from './helpers.js';

const pythonThere = !path.isAbsolute(base.python) || fs.existsSync(base.python);

test('the command line is built from the registry only, one argument per value (no shell, no splitting)', () => {
  const rsa = getScript('rsa_pvtcar');
  const args = buildArgs(rsa, { fromDate: '01/10/2026', toDate: '31/03/2030', options: { prefix: 'A B; rm -rf /', unknown: 'x' } }, { grid: 'C:/jobs/x/in/grid.xlsx', evil: '/etc/passwd' }, 'C:/out/RSA.csv', { gridTools: 'C:/tools' });
  assert.deepEqual(args.slice(0, 4), ['-I', '-X', 'utf8', path.join('C:/tools', 'rsa_pvtcar_to_config.py')]);
  assert.ok(args.includes('--no-prompt'));
  assert.deepEqual(args.slice(args.indexOf('--grid'), args.indexOf('--grid') + 2), ['--grid', 'C:/jobs/x/in/grid.xlsx']);
  assert.deepEqual(args.slice(args.indexOf('--prefix'), args.indexOf('--prefix') + 2), ['--prefix', 'A B; rm -rf /'], 'a value stays one argument');
  assert.ok(!args.includes('/etc/passwd') && !args.includes('--unknown') && !args.includes('x'), 'nothing that the registry does not list');
  const shriram = buildArgs(getScript('shriram_broker'), { fromDate: '01/10/2026', toDate: '31/03/2030', options: { writeEv: true } }, { grid: 'g.xlsx', addendum: 'a.xlsx', ref: 'r.csv' }, 'o.csv', { gridTools: 'T' });
  for (const flag of ['--write-ev', '--addendum', '--ref']) assert.ok(shriram.includes(flag), flag);
  assert.ok(!buildArgs(getScript('shriram_broker'), { fromDate: '01/10/2026', toDate: '31/03/2030', options: { writeEv: false } }, { grid: 'g.xlsx' }, 'o.csv', { gridTools: 'T' }).includes('--write-ev'));
});

test('every registered script file exists and takes the arguments the runner passes', () => {
  const root = path.resolve(base.gridTools);
  for (const s of SCRIPTS) {
    const file = path.join(root, s.file);
    if (!fs.existsSync(file)) continue;                      // the web app may be deployed without some converters
    const text = fs.readFileSync(file, 'utf8');
    for (const flag of ['--no-prompt', '--from-date', '--to-date', '--out', '--ref', ...s.inputs.map((i) => i.flag), ...s.options.map((o) => o.flag)]) {
      assert.ok(text.includes(`"${flag}"`), `${s.file} has ${flag}`);
    }
    assert.ok(text.includes('GRID_SERVICE'), `${s.file} has the service mode`);
  }
});

test('a started script sees the service variables, its own working folder and only the secrets it declares', { skip: !pythonThere && 'no Python here' }, async () => {
  const tools = tempDir('tools-');
  fs.writeFileSync(path.join(tools, 'fake_to_config.py'), [
    'import json, os, sys',
    'out = sys.argv[sys.argv.index("--out") + 1]',
    'os.makedirs(os.path.dirname(out), exist_ok=True)',
    'open(out, "w").write("a,b\\n1,2\\n")',
    'keys = ["GRID_APP_PASSWORD", "SESSION_SECRET", "ANTHROPIC_API_KEY", "GRID_PASSWORD"]',
    'rep = dict(csv=os.path.basename(out), review="", rows=1, flagged=0, seen={k: os.environ.get(k) for k in keys}, service=os.environ.get("GRID_SERVICE"),',
    '           cwd=os.getcwd(), argv=sys.argv[1:], files={k: bool(os.environ.get(k)) for k in ["GRID_QUESTIONS_FILE", "GRID_ANSWERS_FILE", "GRID_ALIASES_FILE", "GRID_STANDING_NOTES_FILE"]},',
    '           answers=json.load(open(os.environ["GRID_ANSWERS_FILE"])), aliases=json.load(open(os.environ["GRID_ALIASES_FILE"])), standing=json.load(open(os.environ["GRID_STANDING_NOTES_FILE"])))',
    'json.dump(rep, open(os.environ["GRID_REPORT_FILE"], "w"))',
  ].join('\n'));
  const script = { id: 'fake', file: 'fake_to_config.py', outPrefix: 'FAKE', inputs: [{ role: 'grid', flag: '--grid' }], options: [], secrets: [{ name: 'password', env: 'GRID_PASSWORD' }] };
  const config = { ...base, gridTools: tools, runTimeoutMs: 60000 };
  const saved = { ...process.env };
  Object.assign(process.env, { GRID_APP_PASSWORD: 'app-pw', SESSION_SECRET: 'sess', ANTHROPIC_API_KEY: 'api-key', GRID_PASSWORD: 'inherited-from-the-server-environment' });
  try {
    const params = { fromDate: '01/10/2026', toDate: '31/03/2030', options: {} };
    const common = { script, params, fileMap: { grid: 'g.xlsx' }, answers: [{ id: 'q1', pick: [2] }], aliases: { rto: { A: 'B' }, makes: {} }, standing: ['a note'], config };
    const dirA = tempDir('run-');
    const withSecret = await runConverter({ ...common, secrets: { password: 'typed-by-the-person' }, runDir: dirA });
    assert.equal(withSecret.exitCode, 0, withSecret.stderr);
    const r = withSecret.report;
    assert.equal(r.service, '1');
    assert.equal(r.seen.GRID_PASSWORD, 'typed-by-the-person');
    assert.equal(r.seen.GRID_APP_PASSWORD, null, 'the app\'s own secrets never reach a converter');
    assert.equal(r.seen.SESSION_SECRET, null);
    assert.equal(r.seen.ANTHROPIC_API_KEY, null);
    assert.equal(path.resolve(r.cwd), path.resolve(dirA), 'a converter works in the run\'s own folder');
    assert.deepEqual(r.files, { GRID_QUESTIONS_FILE: true, GRID_ANSWERS_FILE: true, GRID_ALIASES_FILE: true, GRID_STANDING_NOTES_FILE: true });
    assert.deepEqual(r.answers, { answers: [{ id: 'q1', pick: [2] }] });
    assert.deepEqual(r.aliases, { rto: { A: 'B' }, makes: {} });
    assert.deepEqual(r.standing, ['a note']);
    assert.equal(withSecret.upload, path.join(dirA, 'out', 'FAKE_UPLOAD_2026-10.csv'));
    assert.ok(!r.argv.join(' ').includes('typed-by-the-person'), 'the password is not on the command line');

    const dirB = tempDir('run-');
    const without = await runConverter({ ...common, secrets: {}, runDir: dirB });
    assert.equal(without.report.seen.GRID_PASSWORD, null, 'a password in the server\'s environment is not handed on');
  } finally {
    for (const k of Object.keys(process.env)) if (!(k in saved)) delete process.env[k];
    Object.assign(process.env, saved);
  }
});

test('a script that runs too long is stopped; one that cannot be started is reported', { skip: !pythonThere && 'no Python here' }, async () => {
  const tools = tempDir('tools-');
  fs.writeFileSync(path.join(tools, 'slow_to_config.py'), 'import time\ntime.sleep(60)\n');
  const script = { id: 'slow', file: 'slow_to_config.py', outPrefix: 'SLOW', inputs: [], options: [], secrets: [] };
  const params = { fromDate: '01/10/2026', toDate: '31/03/2030', options: {} };
  const t0 = Date.now();
  const slow = await runConverter({ script, params, fileMap: {}, runDir: tempDir('run-'), config: { ...base, gridTools: tools, runTimeoutMs: 1500 } });
  assert.equal(slow.timedOut, true);
  assert.ok(Date.now() - t0 < 20000, 'it did not wait for the script to finish');
  const missing = await runConverter({ script, params, fileMap: {}, runDir: tempDir('run-'), config: { ...base, python: path.join(tools, 'no-such-python'), gridTools: tools, runTimeoutMs: 5000 } });
  assert.notEqual(missing.exitCode, 0);
  assert.match(missing.stderr, /no-such-python|ENOENT/);
});
