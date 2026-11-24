/** Notes answered together: the notes of one issue (the same wording apart from numbers) are grouped, and one decision clears the ticked ones - all or none. No converter is run here. */
import assert from 'node:assert/strict';
import test from 'node:test';
import { boot, client } from './http.js';
import { setup } from './helpers.js';
import { getScript as getScriptForTest } from '../src/registry.js';
import { issueKey, noteKey } from '../src/memory.js';

const N = (range) => `grid note: Few RTOs for ${range} segment are declined`;
const OTHER = 'grid note: Hazardous category PO will be paid in full';

/** A finished job with note cards (text, silenceable) and, optionally, a question card. */
function jobWithNotes(env, notes, { status = 'review' } = {}) {
  const id = env.store.createJob({ script: 'liberty', params: { fromDate: '01/09/2026', toDate: '31/03/2030', options: {} } });
  env.store.updateJob(id, { status });
  return {
    id,
    cards: notes.map(([text, silenceable = false], i) => env.store.addCard(id, {
      type: 'note', key: `n${i}`, payload: { text, gridNote: silenceable ? text : null, silenceable, rows: i + 1 }, status: 'open',
    })),
  };
}

test('notes that differ only in numbers are one issue; another wording is another issue', () => {
  assert.equal(issueKey(N('1000-1500CC')), issueKey(N('1000-1250CC')));
  assert.equal(issueKey(N('1000-1500CC')), issueKey(N('1600-2500CC')));
  assert.notEqual(issueKey(N('1000-1500CC')), issueKey(OTHER));
  assert.notEqual(issueKey(N('1000-1500CC')), issueKey('grid note: Many RTOs for 1000-1500CC segment are declined'));
});

test('the job view gives every note card its issue, so the page can group them', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const { id } = jobWithNotes(app, [[N('1000-1500CC')], [N('1000-1250CC')], [OTHER]]);
  const view = (await call('GET', `/api/jobs/${id}`)).data;
  const issues = view.cards.map((c) => c.issue);
  assert.equal(issues[0], issues[1]);
  assert.notEqual(issues[0], issues[2]);
});

test('leave: the ticked notes are answered with one request; the others stay open', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC')], [N('1000-1250CC')], [OTHER]]);
  const r = env.agent.answerNotes(id, { cardIds: [cards[0], cards[1]], action: 'leave' });
  assert.equal(r.answered, 2);
  assert.equal(r.open, 1);
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['answered', 'answered', 'open']);
  assert.equal(env.store.getCardById(cards[0]).answer.action, 'leave');
  assert.equal(env.memory.listRules().length, 0, 'leaving a note flagged remembers nothing');
  assert.equal(env.store.events(id).filter((e) => /all at once/.test(e.text)).length, 1, 'one line in the timeline');
});

test('explain: ONE rule for the notes of one issue, and the same issue with other numbers finds it next time', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC')], [N('1000-1250CC')], [N('1250-1600CC')], [OTHER]]);
  env.agent.answerNotes(id, { cardIds: cards, action: 'explain', explanation: 'Handled by hand with the insurer', scope: 'script' });
  const rules = env.memory.listRules({ kind: 'note_explanation' });
  assert.equal(rules.length, 2, 'one for the RTO issue, one for the other issue');
  const rule = env.store.getCardById(cards[0]).rule_id;
  assert.equal(env.store.getCardById(cards[1]).rule_id, rule);
  assert.equal(env.store.getCardById(cards[2]).rule_id, rule);
  assert.notEqual(env.store.getCardById(cards[3]).rule_id, rule);
  assert.equal(env.store.getCardById(cards[0]).answer.explanation, 'Handled by hand with the insurer');
  const again = env.memory.noteExplanationForIssue('liberty', N('3000-3500CC'));
  assert.equal(again?.id, rule, 'the same issue with other numbers');
  assert.equal(env.memory.noteExplanation('liberty', N('3000-3500CC')), null, 'not the same TEXT');
  assert.equal(env.memory.noteExplanationForIssue('liberty', 'grid note: Many RTOs for 3000-3500CC segment are declined'), null, 'another wording is not the same issue');
  assert.equal(env.memory.noteExplanationForIssue('magma', N('3000-3500CC')), null, 'a rule of another converter does not apply');
});

test('ignore: ONE standing note for the notes that differ only in a number - and another number is the same note', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC'), true], [N('1000-1250CC'), true]]);
  env.agent.answerNotes(id, { cardIds: cards, action: 'ignore' });
  assert.equal(env.memory.listRules({ kind: 'standing_note' }).length, 1);
  assert.ok(env.memory.isStandingNote('liberty', N('1000-1500CC')));
  assert.ok(env.memory.isStandingNote('liberty', N('1000-1600CC')), 'a third number: the same note');
  assert.equal(env.memory.isStandingNote('liberty', 'grid note: Many RTOs for 1000-1600CC segment are declined'), null, 'another wording is another note');
  assert.equal(noteKey(N('1000-1500CC')) === noteKey(N('1000-1250CC')), false);
});

test('explained and dismissed notes are resolved: the converter is told not to flag them, a note left flagged is not', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC'), false], [OTHER, false], ['grid note: PO will be 5% less on tankers only', false]]);
  env.agent.answerNotes(id, { cardIds: [cards[0]], action: 'explain', explanation: 'insurer gave no RTO codes: keep as it is' });
  env.agent.answerNotes(id, { cardIds: [cards[1]], action: 'ignore' });
  env.agent.answerNotes(id, { cardIds: [cards[2]], action: 'leave' });
  const job = env.store.getJob(id);
  const list = env.agent.standingList(getScriptForTest(job.script), env.store.cards(id));
  assert.equal(list.length, 2, 'the explained and the dismissed note; not the one left flagged');
  assert.ok(list.some((t) => /1000-1500CC/.test(t)) && list.includes(OTHER));
  assert.ok(env.memory.resolvedFor('liberty', N('3000-3500CC')), 'the explanation also covers the same note with other numbers');
});

test('a flag of the converter itself can be dismissed too (it used to be impossible); all or none: a card of another job refuses the whole request', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC'), true], [N('1000-1250CC'), false]]);
  const other = jobWithNotes(env, [[OTHER, false]]);
  assert.throws(() => env.agent.answerNotes(id, { cardIds: [...cards, other.cards[0]], action: 'ignore' }), (e) => e.status === 404);
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['open', 'open']);
  assert.equal(env.memory.listRules().length, 0);
  env.agent.answerNotes(id, { cardIds: cards, action: 'ignore' });
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['answered', 'answered']);
  assert.equal(env.memory.listRules({ kind: 'standing_note' }).length, 1);
});

test('bad requests are refused with a plain message', () => {
  const env = setup();
  const { id, cards } = jobWithNotes(env, [[N('1000-1500CC')], [N('1000-1250CC')]]);
  const other = jobWithNotes(env, [[OTHER]]);
  const question = env.store.addCard(id, { type: 'question', key: 'q', payload: { problem: 'which sheet?', options: [] }, status: 'open' });
  const refused = (body, status, pattern) => assert.throws(() => env.agent.answerNotes(id, body), (e) => e.status === status && pattern.test(e.message), JSON.stringify(body));
  refused({ cardIds: [], action: 'leave' }, 400, /at least one/);
  refused({ action: 'leave' }, 400, /at least one/);
  refused({ cardIds: cards, action: 'delete' }, 400, /Choose what to do/);
  refused({ cardIds: cards, action: 'explain', explanation: '  ' }, 400, /write your explanation/);
  refused({ cardIds: [...cards, question], action: 'leave' }, 400, /Only notes/);
  refused({ cardIds: [...cards, other.cards[0]], action: 'leave' }, 404, /No such card/);
  refused({ cardIds: Array.from({ length: 501 }, (_, i) => i + 1), action: 'leave' }, 400, /Too many/);
  env.agent.answerNotes(id, { cardIds: [cards[0]], action: 'leave' });
  refused({ cardIds: cards, action: 'leave' }, 409, /already answered/);
  env.store.updateJob(id, { status: 'running' });
  refused({ cardIds: [cards[1]], action: 'leave' }, 409, /running/);
  assert.equal(env.store.getCardById(cards[1]).status, 'open', 'nothing was written by the refused requests');
});

test('over HTTP: the endpoint answers the ticked notes and returns the job; a take-back opens one again', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  const call = client(app.url);
  await call.signIn();
  const { id, cards } = jobWithNotes(app, [[N('1000-1500CC')], [N('1000-1250CC')], [OTHER]]);
  assert.equal((await call('POST', `/api/jobs/${id}/notes/answer`, { cardIds: cards.slice(0, 2), action: 'maybe' })).status, 400);
  const r = await call('POST', `/api/jobs/${id}/notes/answer`, { cardIds: cards.slice(0, 2), action: 'explain', explanation: 'Done by hand', scope: 'insurer' });
  assert.equal(r.status, 200, JSON.stringify(r.data));
  assert.equal(r.data.answered, 2);
  assert.deepEqual(r.data.job.cards.map((c) => c.status), ['answered', 'answered', 'open']);
  assert.equal(r.data.job.cards[0].rule.scope, 'insurer:Liberty General');
  assert.equal((await call('POST', `/api/jobs/${id}/notes/answer`, { cardIds: [cards[0]], action: 'leave' })).status, 409, 'already answered');
  const back = await call('POST', `/api/jobs/${id}/cards/${cards[0]}/reopen`, {});
  assert.equal(back.data.job.cards[0].status, 'open');
  assert.equal((await call('POST', '/api/jobs/0123456789abcdef/notes/answer', { cardIds: [1], action: 'leave' })).status, 404);
});
