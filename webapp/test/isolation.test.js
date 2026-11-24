/**
 * The Grid agent is a standalone project. Nothing in it may import, copy, call or point at the Django insurance portal, and it must neither read the portal's settings nor hand them
 * on. If a link to the portal is ever wanted, it is built on purpose - one explicit API client, its own settings - and THIS TEST IS CHANGED TOGETHER WITH IT: that is its job.
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { PORTAL_ENV, ROOT, config } from '../src/config.js';
import { runConverter } from '../src/runner.js';
import { tempDir } from './helpers.js';

const SELF = fileURLToPath(import.meta.url);
const pythonThere = !path.isAbsolute(config.python) || fs.existsSync(config.python);

function walk(dir, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (['node_modules', 'data', '.git'].includes(e.name)) continue;
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, out);
    else out.push(p);
  }
  return out;
}

const text = (f) => fs.readFileSync(f, 'utf8');
const rel = (f) => path.relative(ROOT, f).replaceAll('\\', '/');

/** Code and configuration of the app (the README may talk about the portal; the code may not). Its Dockerfile, which sits one folder up, counts as part of it. */
const dockerfile = path.join(ROOT, '..', 'Dockerfile');
const CODE = walk(ROOT)
  .concat(fs.existsSync(dockerfile) ? [dockerfile] : [])
  .filter((f) => /\.(js|mjs|json|html|css|py|bat)$/.test(f) || /^\.env\.example$|^Dockerfile$/.test(path.basename(f)))
  .filter((f) => f !== SELF && path.basename(f) !== 'package-lock.json');
const SOURCES = CODE.filter((f) => /[\\/]src[\\/]/.test(f) || path.basename(f) === 'server.js');

test('no code or configuration of the app mentions the Django portal: its folder, its manage.py, its modules, its port, its tools', () => {
  const markers = [
    /insurance[-_ ]broker[-_ ]portal/i, /manage\.py/, /\bdjango/i, /DJANGO_SETTINGS_MODULE/, /project\.settings/, /\bcelery\b/i, /\bgunicorn\b/i,
    /\b(from|import)\s+insurance\b/, /require\(['"]insurance/, /(localhost|127\.0\.0\.1):8000/,
  ];
  assert.ok(CODE.length > 20, `scanned ${CODE.length} files`);
  const barrier = new Set(['src/config.js', 'src/runner.js']);        // the two places that keep the portal's settings out of a converter may say why
  for (const f of CODE) {
    if (barrier.has(rel(f))) continue;
    const body = text(f);
    for (const rx of markers) assert.doesNotMatch(body, rx, `${rel(f)} mentions the portal (${rx})`);
  }
});

test("the portal's settings are not named anywhere in the app, except in the one list that removes them", () => {
  for (const f of CODE) {
    if (rel(f) === 'src/config.js') continue;                      // the list itself
    const body = text(f);
    for (const name of PORTAL_ENV) assert.doesNotMatch(body, new RegExp(`\\b${name}\\b`), `${rel(f)} names the portal setting ${name}`);
  }
});

test('the app reads only its own environment variables - a fixed list, no dynamic lookups', () => {
  const OWN = new Set([
    'GRID_TOOLS_DIR', 'PYTHON', 'PORT', 'HOST', 'DATA_DIR', 'GRID_APP_PASSWORD', 'GRID_APP_INSECURE', 'SESSION_SECRET', 'SESSION_HOURS', 'COOKIE_SECURE', 'TRUST_PROXY',
    'MAX_UPLOAD_MB', 'RUN_TIMEOUT_MS', 'MAX_PARALLEL_RUNS', 'MEMORY_AUTO_APPLY', 'MEMORY_SUGGEST', 'ADVISOR', 'ADVISOR_MODEL', 'ANTHROPIC_API_KEY', 'ANTHROPIC_BASE_URL',
    'GRID_RTO_NAMES',      // the converters' RTO name mode (cluster / group): the app mirrors it and hands it on
  ]);
  const read = new Set();
  for (const f of SOURCES) {
    const body = text(f);
    assert.doesNotMatch(body, /process\.env\s*\[/, `${rel(f)} looks a variable up by a computed name`);
    for (const m of body.matchAll(/process\.env\.([A-Za-z0-9_]+)/g)) read.add(m[1]);
  }
  assert.ok(read.size >= 10);
  assert.deepEqual([...read].filter((n) => !OWN.has(n)), [], 'a variable that is not one of this app\'s own');
  for (const name of PORTAL_ENV) assert.ok(!read.has(name), `${name} is the portal's`);
});

test('the app makes no outgoing calls of its own except the optional advisor: nothing can reach the portal by accident', () => {
  const outgoing = /\bfetch\(|from 'node:https?'|require\(['"]https?['"]\)|\baxios\b|\bundici\b|\bgot\(|XMLHttpRequest|WebSocket\(/;
  for (const f of SOURCES) {
    if (path.basename(f) === 'advisor.js') continue;
    assert.doesNotMatch(text(f), outgoing, `${rel(f)} makes an outgoing call`);
  }
  assert.match(text(path.join(ROOT, 'src', 'advisor.js')), /api\.anthropic\.com|baseUrl/, 'the one place that does is the advisor');
});

test('its own places: data under the app folder, converters in the folder above, nothing that goes through the portal', () => {
  const home = path.resolve(ROOT);
  assert.ok(path.resolve(config.dataDir).startsWith(home) || process.env.DATA_DIR, 'the database and the job files live under the app unless DATA_DIR says otherwise');
  assert.equal(path.resolve(config.gridTools), process.env.GRID_TOOLS_DIR ? path.resolve(process.env.GRID_TOOLS_DIR) : path.resolve(home, '..'));
  assert.doesNotMatch(JSON.stringify(config), /insurance[-_ ]broker[-_ ]portal/i);
  assert.ok(fs.existsSync(path.join(ROOT, 'package.json')), 'its own package.json');
  assert.ok(!fs.existsSync(path.join(ROOT, 'manage.py')) && !fs.existsSync(path.join(ROOT, 'venv')), 'nothing of the portal inside it');
  const pkg = JSON.parse(text(path.join(ROOT, 'package.json')));
  assert.deepEqual(Object.keys(pkg.dependencies).sort(), ['express', 'multer'], 'two dependencies, none of them a bridge to another system');
});

test("a converter is started without the portal's settings, even when the server's own shell has them", { skip: !pythonThere && 'no Python here' }, async () => {
  const tools = tempDir('tools-');
  fs.writeFileSync(path.join(tools, 'fake_to_config.py'), [
    'import json, os, sys',
    'out = sys.argv[sys.argv.index("--out") + 1]',
    'os.makedirs(os.path.dirname(out), exist_ok=True)',
    'open(out, "w").write("a\\n1\\n")',
    `names = json.loads(${JSON.stringify(JSON.stringify(PORTAL_ENV))})`,
    'rep = dict(csv=os.path.basename(out), review="", rows=1, flagged=0, seen={k: os.environ.get(k) for k in names})',
    'json.dump(rep, open(os.environ["GRID_REPORT_FILE"], "w"))',
  ].join('\n'));
  const script = { id: 'fake', file: 'fake_to_config.py', outPrefix: 'FAKE', inputs: [], options: [], secrets: [] };
  const saved = { ...process.env };
  for (const name of PORTAL_ENV) process.env[name] = `portal-value-of-${name}`;
  try {
    const res = await runConverter({
      script, params: { fromDate: '01/10/2026', toDate: '31/03/2030', options: {} }, fileMap: {}, runDir: tempDir('run-'),
      config: { ...config, gridTools: tools, runTimeoutMs: 60000 },
    });
    assert.equal(res.exitCode, 0, res.stderr);
    assert.deepEqual(Object.entries(res.report.seen).filter(([, v]) => v !== null), [], 'a portal setting reached the converter');
  } finally {
    for (const k of Object.keys(process.env)) if (!(k in saved)) delete process.env[k];
    Object.assign(process.env, saved);
  }
});
