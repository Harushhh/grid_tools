/** The HTTP layer: sign-in, the guards, validation of uploads, the memory (rules) endpoints. No converter is run here. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { PASSWORD, boot, client, jobForm } from './http.js';

const XLSX = { data: Buffer.from('PK\u0003\u0004 not really a workbook'), name: 'grid.xlsx' };

test('the page and the health check are open; the API is not', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  assert.equal((await call('GET', '/healthz')).status, 200);
  const page = await call('GET', '/');
  assert.equal(page.status, 200);
  assert.match(page.buf.toString(), /Grid agent/);
  assert.match(page.headers.get('content-security-policy'), /default-src 'self'/);
  assert.equal(page.headers.get('x-content-type-options'), 'nosniff');
  assert.equal(page.headers.get('x-powered-by'), null);
  for (const [m, p] of [['GET', '/api/jobs'], ['GET', '/api/rules'], ['GET', '/api/scripts'], ['POST', '/api/jobs'], ['GET', '/api/rules/export'], ['GET', '/api/masters/search?q=a']]) {
    assert.equal((await call(m, p)).status, 401, `${m} ${p} needs a session`);
  }
  const me = await call('GET', '/api/me');
  assert.deepEqual(me.data, { authed: false, required: true });
});

test('sign-in: wrong password refused, right one gives a HttpOnly SameSite=Strict cookie, too many attempts are throttled, sign-out ends the session', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  assert.equal((await call('POST', '/api/login', { password: 'nope' })).status, 401);
  const ok = await call('POST', '/api/login', { password: PASSWORD }, { raw: true });
  assert.equal(ok.status, 200);
  const cookie = ok.headers.get('set-cookie');
  assert.match(cookie, /HttpOnly/);
  assert.match(cookie, /SameSite=Strict/);
  assert.equal((await call('GET', '/api/jobs')).status, 200);
  assert.equal((await call('GET', '/api/me')).data.authed, true);
  await call('POST', '/api/logout', {});
  assert.equal((await call('GET', '/api/jobs')).status, 401);

  const other = client(app.url);
  let last;
  for (let i = 0; i < 9; i += 1) last = await other('POST', '/api/login', { password: `wrong-${i}` });
  assert.equal(last.status, 429);
  assert.equal((await other('POST', '/api/login', { password: PASSWORD })).status, 429, 'the right password does not get round the throttle');
});

test('a forged cookie is refused', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const forged = Buffer.from(JSON.stringify({ exp: Date.now() + 1e9 })).toString('base64url');
  const r = await fetch(`${app.url}/api/jobs`, { headers: { cookie: `gridsid=${forged}.AAAA` } });
  assert.equal(r.status, 401);
});

test('state-changing requests must come from the app itself (header and origin)', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const noHeader = await fetch(`${app.url}/api/rules/import`, { method: 'POST', headers: { cookie: call.cookie(), 'content-type': 'application/json' }, body: '{"rules":[]}' });
  assert.equal(noHeader.status, 403);
  const foreign = await call('POST', '/api/rules/import', { rules: [] }, { headers: { Origin: 'https://evil.example' } });
  assert.equal(foreign.status, 403);
  assert.equal((await call('POST', '/api/rules/import', { rules: [] })).status, 200);
});

test('uploads are validated before anything is created', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const base = { script: 'liberty', fromDate: '01/09/2026', toDate: '31/03/2030', files: { grid: XLSX } };
  const bad = async (form, pattern) => {
    const r = await call('POST', '/api/jobs', form);
    assert.equal(r.status, 400, JSON.stringify(r.data));
    assert.match(r.data.error, pattern);
  };
  await bad(jobForm({ ...base, script: 'rm -rf' }), /Unknown converter/);
  await bad(jobForm({ ...base, fromDate: '31/02/2026' }), /from_date/);
  await bad(jobForm({ ...base, toDate: '2030-03-31' }), /to_date/);
  await bad(jobForm({ ...base, fromDate: '01/04/2030' }), /before from_date/);
  await bad(jobForm({ ...base, files: {} }), /please choose a file/);
  await bad(jobForm({ ...base, files: { grid: { data: Buffer.from('x'), name: 'grid.exe' } } }), /expected \.xlsx or \.xlsm/);
  await bad(jobForm({ ...base, files: { grid: { data: Buffer.alloc(0), name: 'grid.xlsx' } } }), /empty/);
  await bad(jobForm({ ...base, files: { grid: XLSX, extra: XLSX } }), /Unexpected file/);
  await bad(jobForm({ ...base, script: 'rsa_pvtcar', fields: { opt_prefix: 'a b; rm -rf /' } }), /not allowed/);
  assert.equal((await call('GET', '/api/jobs')).data.jobs.length, 0, 'nothing was created');
  assert.equal(app.db.prepare('SELECT COUNT(*) AS n FROM job_files').get().n, 0);
  assert.deepEqual(fs.readdirSync(path.join(app.config.dataDir, 'tmp')), [], 'no half-finished upload is left behind');
});

test('job ids and file names cannot be used to reach other files', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  for (const id of ['..', '../../etc/passwd', '%2e%2e%2f', 'zzzzzzzzzzzzzzzz', '0123456789abcdef']) {
    assert.equal((await call('GET', `/api/jobs/${id}`)).status, 404, id);
    assert.equal((await call('GET', `/api/jobs/${id}/download/upload`)).status, 404, id);
  }
  assert.equal((await call('POST', '/api/jobs/0123456789abcdef/rerun', {})).status, 404);
  assert.equal((await call('DELETE', '/api/jobs/0123456789abcdef')).status, 404);
});

test('the scripts the page can offer are exactly the registry, with no paths in them', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const { scripts } = (await call('GET', '/api/scripts')).data;
  assert.deepEqual(scripts.map((s) => s.id), ['liberty', 'magma', 'rsa_gcv4w', 'rsa_pvtcar', 'united', 'shriram_broker', 'shriram_2w', 'godigit', 'sbi', 'icici', 'indusind']);
  assert.ok(!JSON.stringify(scripts).includes('.py'), 'no script file names');
  const digit = scripts.find((s) => s.id === 'godigit');
  assert.deepEqual(digit.secrets.map((s) => s.name), ['password']);
});

test('the memory: import, list, switch off, edit, export, delete; a bad file is refused', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const rules = [
    { kind: 'standing_note', scope: 'script:liberty', match_text: 'n', data: { text: 'A note that needs no action at all, written out in full.' }, explanation: 'ok' },
    { kind: 'alias', scope: 'global', match_text: 'a', data: { kind: 'rto', from: 'OLD', to: 'NEW' } },
    { kind: 'bogus', scope: 'global', data: {} },
  ];
  assert.equal((await call('POST', '/api/rules/import', { rules })).data.imported, 2, 'the unknown kind is skipped');
  assert.equal((await call('POST', '/api/rules/import', { rules: [{ kind: 'alias', scope: 'script:../../x', data: { kind: 'rto', from: 'a', to: 'b' } }] })).status, 400);
  assert.equal((await call('POST', '/api/rules/import', { nothing: 1 })).status, 400);
  const list = (await call('GET', '/api/rules')).data.rules;
  assert.equal(list.length, 2);
  assert.equal((await call('GET', '/api/rules?kind=alias')).data.rules.length, 1);
  assert.equal((await call('GET', '/api/rules?q=needs%20no%20action')).data.rules.length, 1);
  const note = list.find((r) => r.kind === 'standing_note');
  assert.equal((await call('PATCH', `/api/rules/${note.id}`, { enabled: false })).data.rule.enabled, false);
  assert.equal((await call('GET', '/api/rules?enabled=1')).data.rules.length, 1);
  assert.equal((await call('PATCH', `/api/rules/${note.id}`, { explanation: 'changed', scope: 'insurer:Liberty General' })).data.rule.scope, 'insurer:Liberty General');
  assert.equal((await call('PATCH', `/api/rules/${note.id}`, { scope: 'nonsense' })).status, 400);
  const exported = await call('GET', '/api/rules/export');
  assert.match(exported.headers.get('content-disposition'), /attachment/);
  assert.equal(exported.data.rules.length, 2);
  assert.equal((await call('DELETE', `/api/rules/${note.id}`)).status, 200);
  assert.equal((await call('DELETE', `/api/rules/${note.id}`)).status, 404);
  assert.equal((await call('GET', '/api/stats')).data.rules, 1);
});

test('oversized and malformed bodies get a plain error, not a stack trace', async (t) => {
  const app = await boot({ maxUploadMb: 1 });
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const broken = await fetch(`${app.url}/api/rules/import`, { method: 'POST', headers: { cookie: call.cookie(), 'x-requested-with': 'gridapp', 'content-type': 'application/json' }, body: '{not json' });
  assert.equal(broken.status, 400);
  assert.doesNotMatch(await broken.text(), /at .*\.js/);
  const big = jobForm({ script: 'liberty', fromDate: '01/09/2026', toDate: '31/03/2030', files: { grid: { data: Buffer.alloc(2 * 1024 * 1024, 1), name: 'grid.xlsx' } } });
  const r = await call('POST', '/api/jobs', big);
  assert.equal(r.status, 413);
  assert.match(r.data.error, /larger than 1 MB/);
});
