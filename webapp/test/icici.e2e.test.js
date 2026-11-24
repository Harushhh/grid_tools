/**
 * ICICI Lombard through HTTP, with the real converter: the page offers it; one run takes any of the four workbooks (all optional, at least one needed) and the download is
 * what the command line writes (regression baseline); a run with no grid, with only an "earlier upload", or with a file of the wrong type is refused.
 */
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { BASELINE, CAN_RUN, INPUTS } from './helpers.js';
import { boot, client, jobForm, settled, waitFor } from './http.js';

const skip = !CAN_RUN && 'the converters or the regression grids are not available here';
const read = (name) => ({ data: fs.readFileSync(path.join(INPUTS, name)), name });
const job = (call, id) => call('GET', `/api/jobs/${id}`).then((r) => r.data);
const finish = (call, id) => waitFor(() => job(call, id), settled, 240000);
const sha = (buf) => crypto.createHash('sha256').update(buf).digest('hex');

test('ICICI over HTTP: offered by the page; Private Car + Two-Wheeler + new Two-Wheeler in one run, then the CV grid - the downloads are the command line\'s bytes', { skip, timeout: 900000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const icici = (await call('GET', '/api/scripts')).data.scripts.find((s) => s.id === 'icici');
  assert.ok(icici, 'the page offers ICICI Lombard');
  assert.deepEqual(icici.inputs.filter((i) => i.role !== 'ref').map((i) => i.role), ['cvGrid', 'pcGrid', 'twGrid', 'twNewGrid']);
  assert.ok(icici.inputs.every((i) => !i.required), 'every grid is optional (any of the four), at least one is checked by the server');

  const made = await call('POST', '/api/jobs', jobForm({
    script: 'icici', fromDate: '01/11/2026', toDate: '31/03/2030',
    files: { pcGrid: read('icici_pc_nov26.xlsx'), twGrid: read('icici_tw_nov26.xlsx'), twNewGrid: read('icici_tw_new_nov26.xlsx') },
  }));
  assert.ok(made.data.id, JSON.stringify(made.data));
  const view = await finish(call, made.data.id);
  assert.ok(['done', 'review'].includes(view.status), `status ${view.status}`);
  assert.equal(view.outputs.rows, 3342);
  const csv = await call('GET', `/api/jobs/${view.id}/download/upload`);
  assert.equal(sha(csv.buf), BASELINE.icici_pc_tw_nov26.sha256, 'the web path writes the same bytes as the command line');
  const lines = csv.buf.toString('utf8').split(/\r?\n/).filter(Boolean);
  assert.equal(lines.length, 3343, 'header + 3,342 rows');
  assert.ok(lines.slice(1).every((l) => l.split(',')[2] === 'ICICI-RB-Offline'), 'insurer_vertical is ICICI-RB-Offline on every row');

  const cv = await finish(call, (await call('POST', '/api/jobs', jobForm({ script: 'icici', fromDate: '01/11/2026', toDate: '31/03/2030', files: { cvGrid: read('icici_cv_nov26.xlsx') } }))).data.id);
  assert.ok(['done', 'review'].includes(cv.status), `status ${cv.status}`);
  assert.equal(cv.outputs.rows, 20317);
  assert.equal(sha((await call('GET', `/api/jobs/${cv.id}/download/upload`)).buf), BASELINE.icici_cv_nov26.sha256, 'the CV grid: the same bytes as the command line');
});

test('ICICI over HTTP: no grid, only an earlier upload, or a file of the wrong type is refused - nothing is started', { timeout: 120000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const base = { script: 'icici', fromDate: '01/11/2026', toDate: '31/03/2030' };
  const none = await call('POST', '/api/jobs', jobForm(base));
  assert.equal(none.status, 400);
  assert.match(JSON.stringify(none.data), /at least one grid file/);
  const refOnly = await call('POST', '/api/jobs', jobForm({ ...base, files: { ref: { data: Buffer.from('id,product\n1,TW\n'), name: 'old.csv' } } }));
  assert.equal(refOnly.status, 400);
  assert.match(JSON.stringify(refOnly.data), /at least one grid file/);
  const wrong = await call('POST', '/api/jobs', jobForm({ ...base, files: { cvGrid: { data: Buffer.from('a,b\n'), name: 'grid.csv' } } }));
  assert.equal(wrong.status, 400);
  assert.match(JSON.stringify(wrong.data), /expected \.xlsx or \.xlsm/);
  assert.equal((await call('GET', '/api/jobs')).data.jobs.length, 0, 'no job was created');
});
