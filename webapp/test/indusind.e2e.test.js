/**
 * IndusInd General (Reliance) through HTTP, with the real converter: the page offers it, the dates are the ones typed for the run, and the download is exactly what the command line writes
 * (regression baseline indusind_nov26); the one grid is required.
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

test('IndusInd over HTTP: offered by the page; one run of the November grid - the download is the command line\'s bytes, 11,461 rows, the standardised insurer name', { skip, timeout: 600000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const indusind = (await call('GET', '/api/scripts')).data.scripts.find((s) => s.id === 'indusind');
  assert.ok(indusind, 'the page offers IndusInd General (Reliance)');
  assert.deepEqual(indusind.inputs.filter((i) => i.role !== 'ref').map((i) => i.role), ['grid']);
  assert.ok(indusind.inputs.find((i) => i.role === 'grid').required, 'the grid is required');

  const made = await call('POST', '/api/jobs', jobForm({ script: 'indusind', fromDate: '01/11/2026', toDate: '31/03/2030', files: { grid: read('indusind_nov26.xlsx') } }));
  assert.ok(made.data.id, JSON.stringify(made.data));
  const view = await finish(call, made.data.id);
  assert.ok(['done', 'review'].includes(view.status), `status ${view.status}`);
  assert.equal(view.outputs.rows, 11461);
  const csv = await call('GET', `/api/jobs/${view.id}/download/upload`);
  assert.equal(sha(csv.buf), BASELINE.indusind_nov26.sha256, 'the web path writes the same bytes as the command line');
  const lines = csv.buf.toString('utf8').split(/\r?\n/).filter(Boolean);
  assert.equal(lines.length, 11462, 'header + 11,461 rows');
  assert.ok(lines.slice(1).every((l) => l.split(',')[1] === 'INDUSIND GENERAL INSURANCE COMPANY LIMITED'), 'the insurer is INDUSIND GENERAL INSURANCE COMPANY LIMITED on every row');
  assert.ok(lines.slice(1).every((l) => /,01\/11\/2026,31\/03\/2030,2966,/.test(l)), 'every row carries the dates typed for the run');
});

test('IndusInd over HTTP: a run without the grid is refused - nothing is started', { timeout: 120000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const none = await call('POST', '/api/jobs', jobForm({ script: 'indusind', fromDate: '01/11/2026', toDate: '31/03/2030' }));
  assert.equal(none.status, 400);
});
