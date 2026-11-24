/**
 * Through HTTP, with the real converters: a question card answered over the API and remembered; a password-protected workbook (password card, wrong password,
 * right password) with a check that the password is never written anywhere; name and note cards with their rules.
 */
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { config as base } from '../src/config.js';
import { BASELINE, CAN_RUN, INPUTS, sha256 } from './helpers.js';
import { boot, client, jobForm, settled, waitFor } from './http.js';

const skip = !CAN_RUN && 'the converters or the regression grids are not available here';
const read = (name) => ({ data: fs.readFileSync(path.join(INPUTS, name)), name });
const job = (call, id) => call('GET', `/api/jobs/${id}`).then((r) => r.data);
const finish = (call, id) => waitFor(() => job(call, id), settled, 240000);

function bytesOf(dir) {                                   // every file under a folder, for "is this text anywhere?"
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...bytesOf(p));
    else out.push(fs.readFileSync(p));
  }
  return out;
}

test('over HTTP: the question card, the answer, the download - and the next upload needs nobody', { skip, timeout: 900000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const upload = () => call('POST', '/api/jobs', jobForm({ script: 'liberty', fromDate: '01/07/2026', toDate: '20/03/2030', files: { grid: read('liberty_jul26.xlsx') } }));

  const created = await upload();
  assert.equal(created.status, 201, JSON.stringify(created.data));
  let view = await finish(call, created.data.id);
  assert.equal(view.status, 'needs_input');
  const card = view.cards.find((c) => c.status === 'open');
  assert.equal(card.type, 'question');
  assert.equal(view.outputs, null, 'nothing to download yet');
  assert.equal((await call('GET', `/api/jobs/${view.id}/download/upload`)).status, 404);
  assert.equal((await call('POST', `/api/jobs/${view.id}/rerun`, {})).status, 409, 'cannot run past an open question');

  const pick = card.payload.options.find((o) => /electric/i.test(o.label) && /petrol/i.test(o.label));
  const answered = await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { pick: [pick.n], scope: 'insurer', explanation: 'no Electric column in July' });
  assert.equal(answered.status, 200, JSON.stringify(answered.data));
  assert.equal(answered.data.started, true);
  view = await finish(call, view.id);
  assert.ok(['review', 'done'].includes(view.status), `${view.status} ${view.error}`);
  const csv = await call('GET', `/api/jobs/${view.id}/download/upload`);
  assert.equal(csv.status, 200);
  assert.match(csv.headers.get('content-disposition'), /LIBERTY_UPLOAD_2026-07\.csv/);
  assert.equal(crypto.createHash('sha256').update(csv.buf).digest('hex'), BASELINE.liberty_jul26.sha256, 'the download is what the command line writes');
  const review = await call('GET', `/api/jobs/${view.id}/download/review`);
  assert.equal(review.buf.subarray(0, 2).toString(), 'PK', 'the review workbook is an xlsx');

  const rules = (await call('GET', '/api/rules?kind=answer')).data.rules;
  assert.equal(rules.length, 1);
  assert.equal(rules[0].scope, 'insurer:Liberty General');

  const again = await upload();
  view = await finish(call, again.data.id);
  assert.ok(['review', 'done'].includes(view.status));
  assert.equal(view.cards.filter((c) => c.status === 'open' && c.type === 'question').length, 0);
  const auto = view.cards.find((c) => c.type === 'question' && c.status === 'auto');
  assert.ok(auto, 'answered from memory');
  assert.equal(auto.rule.id, rules[0].id);
  assert.equal(view.applications.length, 1);
  assert.ok(view.events.some((e) => e.kind === 'memory' && /rule #/.test(e.text)));

  // switching the rule off makes the agent ask again
  await call('PATCH', `/api/rules/${rules[0].id}`, { enabled: false });
  const third = await finish(call, (await upload()).data.id);
  assert.equal(third.status, 'needs_input');

  // deleting a job removes its files and keeps the memory
  assert.equal((await call('DELETE', `/api/jobs/${third.id}`)).status, 200);
  assert.equal((await call('GET', `/api/jobs/${third.id}`)).status, 404);
  assert.equal(fs.existsSync(path.join(app.config.dataDir, 'jobs', third.id)), false);
  assert.equal((await call('GET', '/api/rules')).data.rules.length, 1);
});

test('a stop the converter cannot get past becomes a card; the note is remembered and shown the next time it stops that way', { skip, timeout: 300000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const garbage = () => call('POST', '/api/jobs', jobForm({ script: 'godigit', fromDate: '15/10/2026', toDate: '31/03/2030', files: { grid: { data: Buffer.alloc(3000, 7), name: 'x.xls' } } }));

  let view = await finish(call, (await garbage()).data.id);
  assert.equal(view.status, 'failed');
  assert.match(view.error, /does not look like an Excel workbook/);
  const card = view.cards.find((c) => c.type === 'failure' && c.status === 'open');
  assert.ok(card, 'a card, not a dead end');
  assert.equal(card.suggestion, null, 'nothing is known yet');
  assert.equal((await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { action: 'save', note: '  ' })).status, 400);
  const saved = await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { action: 'save', note: 'Digit sent a PDF this month', scope: 'script' });
  assert.equal(saved.status, 200, JSON.stringify(saved.data));
  assert.equal(saved.data.started, false, 'a stop is not run again by itself');
  assert.equal(saved.data.job.status, 'failed');
  const rules = (await call('GET', '/api/rules?kind=failure_note')).data.rules;
  assert.equal(rules.length, 1);
  assert.equal(rules[0].explanation, 'Digit sent a PDF this month');

  view = await finish(call, (await garbage()).data.id);
  const again = view.cards.find((c) => c.type === 'failure' && c.status === 'open');
  assert.equal(again.suggestion.note, 'Digit sent a PDF this month', 'the earlier note comes with the card');
  assert.equal(again.suggestion.ruleId, rules[0].id);

  assert.equal((await call('POST', `/api/jobs/${view.id}/rerun`, {})).status, 200, 'a stopped job can be run again');
  view = await finish(call, view.id);
  assert.equal(view.status, 'failed');
  assert.equal(view.cards.filter((c) => c.type === 'failure').length, 1, 'the same stop is one card, not two');
  assert.equal((await call('POST', `/api/jobs/${view.id}/cards/${view.cards[0].id}/answer`, { action: 'dismiss' })).status, 200);
});

test('a password-protected workbook: password card, a wrong password, the right one - and the password is never stored', { skip, timeout: 900000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const secret = `Tst-${crypto.randomBytes(9).toString('hex')}`;            // a throw-away password for this test only
  const wrong = `Bad-${crypto.randomBytes(9).toString('hex')}`;
  const enc = path.join(os.tmpdir(), `enc-${crypto.randomBytes(4).toString('hex')}.xls`);
  execFileSync(base.python, ['-I', '-c', 'import os, sys, msoffcrypto; f = msoffcrypto.OfficeFile(open(sys.argv[1], "rb")); f.encrypt(os.environ["PW"], open(sys.argv[2], "wb"))',
    path.join(INPUTS, 'godigit_oct26.xlsx'), enc], { env: { ...process.env, PW: secret } });
  t.after(() => fs.rmSync(enc, { force: true }));

  const created = await call('POST', '/api/jobs', jobForm({ script: 'godigit', fromDate: '15/10/2026', toDate: '31/03/2030', files: { grid: { data: fs.readFileSync(enc), name: 'Delhi_PAN_India_Agency.xls' } } }));
  assert.equal(created.status, 201, JSON.stringify(created.data));
  let view = await finish(call, created.data.id);
  assert.equal(view.status, 'needs_input');
  let card = view.cards.find((c) => c.status === 'open');
  assert.equal(card.type, 'password');
  assert.match(card.payload.reason, /password protected/);

  const bad = await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { value: wrong });
  assert.equal(bad.data.started, true);
  view = await finish(call, view.id);
  assert.equal(view.status, 'needs_input');
  card = view.cards.find((c) => c.status === 'open');
  assert.equal(card.type, 'password');
  assert.match(card.payload.reason, /did not accept/);

  assert.equal((await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { value: '' })).status, 400);
  await call('POST', `/api/jobs/${view.id}/cards/${card.id}/answer`, { value: secret });
  view = await finish(call, view.id);
  assert.ok(['review', 'done'].includes(view.status), `${view.status}: ${view.error}`);
  assert.equal(view.outputs.rows, 53212);        // the Go Digit baseline blessed on 11 Nov 2026 (was 53210)
  const csv = await call('GET', `/api/jobs/${view.id}/download/upload`);
  assert.equal(crypto.createHash('sha256').update(csv.buf).digest('hex'), BASELINE.godigit_oct26.sha256, 'the same bytes as the command line');

  // the password is nowhere: not in the database (or its log), not in any file of the job, not in anything the API says
  const needles = [Buffer.from(secret), Buffer.from(wrong)];
  const where = [];
  for (const f of ['app.db', 'app.db-wal', 'app.db-shm']) {
    const p = path.join(app.config.dataDir, f);
    if (fs.existsSync(p)) where.push(['database', fs.readFileSync(p)]);
  }
  for (const b of bytesOf(path.join(app.config.dataDir, 'jobs', view.id))) where.push(['job file', b]);
  where.push(['api view', Buffer.from(JSON.stringify(await job(call, view.id)))], ['job list', Buffer.from(JSON.stringify((await call('GET', '/api/jobs')).data))]);
  for (const [label, buf] of where) for (const n of needles) assert.ok(!buf.includes(n), `a password was found in the ${label}`);

  // taking the password back forgets it and asks again
  const given = view.cards.find((c) => c.type === 'password');
  assert.equal(given.status, 'answered');
  assert.deepEqual(given.answer, { provided: true });
  const reopened = await call('POST', `/api/jobs/${view.id}/cards/${given.id}/reopen`, {});
  assert.equal(reopened.data.job.status, 'needs_input');
  assert.equal((await call('POST', `/api/jobs/${view.id}/rerun`, {})).status, 409);
});

test('name and note cards: a name the master lacks is mapped and remembered; a note is switched off for good; a flag of the converter itself cannot be', { skip, timeout: 900000 }, async (t) => {
  // this flow is the group-name mode (new_rto_list = the master's group names, RTO name cards); the default mode writes bare cluster names and asks no RTO name
  const app = await boot({ rtoNames: 'group' });
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const upload = () => call('POST', '/api/jobs', jobForm({ script: 'liberty', fromDate: '01/09/2026', toDate: '31/03/2030', files: { grid: read('liberty_sep26.xlsx') } }));

  let view = await finish(call, (await upload()).data.id);
  assert.equal(view.status, 'review');
  assert.equal(view.outputs.rows, 23690);
  const names = view.cards.filter((c) => c.type === 'alias' && c.status === 'open');
  const notes = view.cards.filter((c) => c.type === 'note' && c.status === 'open');
  assert.equal(names.length, 2, 'two group names the master spells differently');
  assert.equal(notes.length, 1, 'the converter\'s own "not in the master" flags are not offered a second time as notes');
  assert.equal(notes[0].payload.silenceable, true);
  const flaggedBefore = view.outputs.flagged;

  // a name: the closest master name is offered, and only a name that is in the master is accepted
  const assam = names.find((c) => /ASSAM/.test(c.payload.name));
  const master = (await call('GET', `/api/masters/search?kind=rto&q=${encodeURIComponent('LIBERTY_GCV_4W>7.5T_APR26_Assam')}`)).data;
  assert.equal(master.loaded, true);
  assert.ok(master.names.includes('LIBERTY_GCV_4W>7.5T_APR26_Assam'));
  assert.equal((await call('POST', `/api/jobs/${view.id}/cards/${assam.id}/answer`, { to: 'NOT_A_GROUP_IN_THE_MASTER' })).status, 400);
  const mapped = await call('POST', `/api/jobs/${view.id}/cards/${assam.id}/answer`, { to: 'LIBERTY_GCV_4W>7.5T_APR26_Assam', scope: 'script', explanation: 'capital letters only' });
  assert.equal(mapped.status, 200, JSON.stringify(mapped.data));
  assert.equal(mapped.data.started, false, 'review answers wait for "run again"');

  // a note that the grid carries: "no action needed"
  const answeredNote = await call('POST', `/api/jobs/${view.id}/cards/${notes[0].id}/answer`, { action: 'ignore', scope: 'insurer' });
  assert.equal(answeredNote.status, 200, JSON.stringify(answeredNote.data));
  assert.equal(answeredNote.data.job.pending_changes, 2);

  assert.equal((await call('POST', `/api/jobs/${view.id}/rerun`, {})).status, 200);
  view = await finish(call, view.id);
  assert.equal(view.outputs.rows, 23690);
  assert.ok(view.outputs.flagged < flaggedBefore, `fewer rows flagged: ${flaggedBefore} -> ${view.outputs.flagged}`);
  assert.equal(view.cards.filter((c) => c.type === 'alias' && c.status === 'open').length, 1, 'the mapped name is no longer open');
  assert.equal(view.pending_changes, 0);

  // the next upload does both by itself, and says so
  const next = await finish(call, (await upload()).data.id);
  const fromMemory = next.cards.filter((c) => c.status === 'auto');
  assert.ok(fromMemory.some((c) => c.type === 'alias' && /ASSAM/.test(c.payload.name)), 'the name is mapped from memory');
  assert.ok(fromMemory.some((c) => c.type === 'note' && c.answer.action === 'ignore'), 'the note is not flagged, and the page says so');
  assert.equal(next.outputs.flagged, view.outputs.flagged, 'the same upload as the one you finished by hand');
  assert.equal(next.cards.filter((c) => c.status === 'open').length, 1, 'only the Telangana name is still open');
  assert.ok(next.applications.length >= 2);
  const bytes = async (id) => crypto.createHash('sha256').update((await call('GET', `/api/jobs/${id}/download/upload`)).buf).digest('hex');
  assert.equal(await bytes(next.id), await bytes(view.id), 'by itself or by hand, the same upload');
});

test('cluster names: the default mode writes only the cluster name, asks about no RTO name, and says nothing about the RTO master', { skip, timeout: 900000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const view = await finish(call, (await call('POST', '/api/jobs', jobForm({ script: 'liberty', fromDate: '01/09/2026', toDate: '31/03/2030', files: { grid: read('liberty_sep26.xlsx') } }))).data.id);
  assert.equal(view.status, 'review');
  assert.equal(view.outputs.rows, 23690);
  assert.equal(view.cards.filter((c) => c.type === 'alias').length, 0, 'no "RTO group is not in the master" card');
  const lines = (await call('GET', `/api/jobs/${view.id}/download/upload`)).buf.toString('utf8').split(/\r?\n/).filter(Boolean);
  const cells = (line) => [...line.matchAll(/(?:^|,)("(?:[^"]|"")*"|[^,]*)/g)].map((m) => m[1].replace(/^"|"$/g, ''));      // a quote-aware split: a make list holds commas
  const col = cells(lines[0]).indexOf('new_rto_list');
  assert.ok(col >= 0);
  const names = new Set(lines.slice(1).map((l) => cells(l)[col]));
  assert.ok(names.has('ALLINDIA'));
  assert.ok(![...names].some((n) => /^LIBERTY_/.test(n)), 'no portal group name in the CSV');
});

test('the memory acts at once: every note of a Go Digit run is explained or dismissed, and the run again - and a fresh upload - flag none of them', { skip, timeout: 1200000 }, async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const upload = () => call('POST', '/api/jobs', jobForm({ script: 'godigit', fromDate: '01/10/2026', toDate: '31/03/2030', files: { grid: read('godigit_oct26.xlsx') } }));
  let view = await finish(call, (await upload()).data.id);
  const open = () => view.cards.filter((c) => c.type === 'note' && c.status === 'open');
  const notes = open();
  assert.ok(notes.length >= 3 && view.outputs.flagged > 0, `notes ${notes.length}, flagged ${view.outputs.flagged}`);
  const half = Math.ceil(notes.length / 2);
  const explained = await call('POST', `/api/jobs/${view.id}/notes/answer`, { cardIds: notes.slice(0, half).map((c) => c.id), action: 'explain', explanation: 'checked by hand', scope: 'script' });
  assert.equal(explained.status, 200, JSON.stringify(explained.data));
  const dismissed = await call('POST', `/api/jobs/${view.id}/notes/answer`, { cardIds: notes.slice(half).map((c) => c.id), action: 'ignore', scope: 'script' });
  assert.equal(dismissed.status, 200, JSON.stringify(dismissed.data));
  assert.equal((await call('POST', `/api/jobs/${view.id}/rerun`, {})).status, 200);
  view = await finish(call, view.id);
  assert.equal(open().length, 0, 'no note is asked again');
  assert.equal(view.outputs.flagged, 0, 'no row says Manual Review Required any more');
  const csv = (await call('GET', `/api/jobs/${view.id}/download/upload`)).buf.toString('utf8');
  assert.ok(!csv.includes('Manual Review Required'));
  // a new upload of the same grid, minutes later: the memory is used at once, nothing is asked, nothing is flagged
  const next = await finish(call, (await upload()).data.id);
  assert.equal(next.cards.filter((c) => c.type === 'note' && c.status === 'open').length, 0);
  assert.equal(next.outputs.flagged, 0);
  assert.ok(next.cards.some((c) => c.type === 'note' && c.status === 'auto' && ['explain', 'ignore'].includes(c.answer.action)), 'the page says so: from memory');
});
