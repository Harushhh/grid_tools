/** "Name not in the master" cards answered together: leave them all as they are, or map them all to one master group - all or none. No converter is run here. */
import assert from 'node:assert/strict';
import test from 'node:test';
import { boot, client } from './http.js';
import { setup } from './helpers.js';

const MASTER = { rto: new Set(['GO-DIGIT_OCT26_HCV_ROM3', 'GO-DIGIT_OCT26_HCV_GOA']), makes: new Set(['TATA']) };
const fakeMasters = { loaded: () => true, has: (kind, name) => MASTER[kind].has(name), set: (kind) => MASTER[kind] };

function jobWithNames(env, names, { status = 'review' } = {}) {
  const id = env.store.createJob({ script: 'godigit', params: { fromDate: '01/10/2026', toDate: '31/03/2030', options: {} } });
  env.store.updateJob(id, { status });
  return {
    id,
    cards: names.map(([kind, name], i) => env.store.addCard(id, { type: 'alias', key: `${kind}:${name}`, payload: { kind, name, rows: 100 * (i + 1), hint: '', closest: [] }, status: 'open' })),
  };
}

const withMasters = (env) => {
  env.agent.masters = fakeMasters;
  return env;
};

test('leave: the ticked names are answered with one request and remember nothing', () => {
  const env = withMasters(setup());
  const { id, cards } = jobWithNames(env, [['rto', 'GO-DIGIT_OCT26_HCV_Mumbai'], ['rto', 'GO-DIGIT_OCT26_HCV_HR_Ref_B'], ['rto', 'GO-DIGIT_OCT26_HCV_Pune']]);
  const r = env.agent.answerAliases(id, { cardIds: cards.slice(0, 2), skip: true });
  assert.equal(r.answered, 2);
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['answered', 'answered', 'open']);
  assert.equal(env.store.getCardById(cards[0]).answer.skip, true);
  assert.equal(env.memory.listRules().length, 0);
  assert.equal(env.store.events(id).filter((e) => /left 2 names as they are, all at once/.test(e.text)).length, 1);
});

test('map: every ticked name gets the chosen group and its own rule, which answers next month by itself', () => {
  const env = withMasters(setup());
  const { id, cards } = jobWithNames(env, [['rto', 'GO-DIGIT_OCT26_HCV_Mumbai'], ['rto', 'GO-DIGIT_OCT26_HCV_Pune'], ['rto', 'GO-DIGIT_OCT26_HCV_Surat']]);
  env.agent.answerAliases(id, { cardIds: cards.slice(0, 2), to: 'GO-DIGIT_OCT26_HCV_ROM3', scope: 'script', explanation: 'the rest of Maharashtra' });
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['answered', 'answered', 'open']);
  assert.equal(env.store.getCardById(cards[0]).answer.to, 'GO-DIGIT_OCT26_HCV_ROM3');
  const rules = env.memory.listRules({ kind: 'alias' });
  assert.equal(rules.length, 2, 'a rule maps one name');
  assert.ok(rules.every((r) => r.data.to === 'GO-DIGIT_OCT26_HCV_ROM3' && r.explanation === 'the rest of Maharashtra' && r.scope === 'script:godigit'));
  assert.notEqual(env.store.getCardById(cards[0]).rule_id, env.store.getCardById(cards[1]).rule_id);
  assert.equal(env.memory.matchAlias('godigit', 'rto', 'GO-DIGIT_NOV26_HCV_Mumbai', MASTER.rto).suggestion?.to, 'GO-DIGIT_NOV26_HCV_ROM3', 'another month: the rule carries the month, and the group must exist in the master to be applied by itself');
  assert.equal(env.memory.matchAlias('godigit', 'rto', 'GO-DIGIT_OCT26_HCV_Mumbai', MASTER.rto).auto?.to, 'GO-DIGIT_OCT26_HCV_ROM3');
  assert.deepEqual(env.memory.matchAlias('godigit', 'rto', 'GO-DIGIT_OCT26_HCV_Surat', MASTER.rto), {}, 'the name that was not ticked is still unknown');
});

test('map with "remember" switched off: the names are mapped for this job only', () => {
  const env = withMasters(setup());
  const { id, cards } = jobWithNames(env, [['rto', 'A'], ['rto', 'B']]);
  env.agent.answerAliases(id, { cardIds: cards, to: 'GO-DIGIT_OCT26_HCV_GOA', remember: false });
  assert.equal(env.memory.listRules().length, 0);
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).answer.to), ['GO-DIGIT_OCT26_HCV_GOA', 'GO-DIGIT_OCT26_HCV_GOA']);
});

test('all or none: a group the master lacks, or two kinds in one mapping, refuses the whole request', () => {
  const env = withMasters(setup());
  const { id, cards } = jobWithNames(env, [['rto', 'A'], ['rto', 'B'], ['makes', 'Tata']]);
  const refused = (body, status, pattern) => assert.throws(() => env.agent.answerAliases(id, body), (e) => e.status === status && pattern.test(e.message), JSON.stringify(body));
  refused({ cardIds: cards.slice(0, 2), to: 'NOT-IN-THE-MASTER' }, 400, /Not in the master/);
  refused({ cardIds: cards, to: 'GO-DIGIT_OCT26_HCV_ROM3' }, 400, /different kinds|one kind/);
  refused({ cardIds: cards.slice(0, 2), to: '' }, 400, /Choose the master group/);
  assert.deepEqual(cards.map((c) => env.store.getCardById(c).status), ['open', 'open', 'open']);
  assert.equal(env.memory.listRules().length, 0);
  assert.equal(env.agent.answerAliases(id, { cardIds: cards, skip: true }).answered, 3, 'leaving them as they are works across kinds');
});

test('bad requests are refused with a plain message', () => {
  const env = withMasters(setup());
  const { id, cards } = jobWithNames(env, [['rto', 'A'], ['rto', 'B']]);
  const other = jobWithNames(env, [['rto', 'C']]);
  const note = env.store.addCard(id, { type: 'note', key: 'n', payload: { text: 'x', rows: 1 }, status: 'open' });
  const refused = (body, status, pattern) => assert.throws(() => env.agent.answerAliases(id, body), (e) => e.status === status && pattern.test(e.message), JSON.stringify(body));
  refused({ cardIds: [], skip: true }, 400, /at least one/);
  refused({ skip: true }, 400, /at least one/);
  refused({ cardIds: [...cards, note], skip: true }, 400, /Only names/);
  refused({ cardIds: [...cards, other.cards[0]], skip: true }, 404, /No such card/);
  refused({ cardIds: Array.from({ length: 501 }, (_, i) => i + 1), skip: true }, 400, /Too many/);
  env.agent.masters = { loaded: () => false, has: () => false };
  refused({ cardIds: cards, to: 'GO-DIGIT_OCT26_HCV_ROM3' }, 409, /could not be read/);
  env.agent.masters = fakeMasters;
  env.agent.answerAliases(id, { cardIds: [cards[0]], skip: true });
  refused({ cardIds: cards, skip: true }, 409, /already answered/);
  env.store.updateJob(id, { status: 'running' });
  refused({ cardIds: [cards[1]], skip: true }, 409, /running/);
  assert.equal(env.store.getCardById(cards[1]).status, 'open');
});

test('over HTTP: the endpoint answers the ticked names and returns the job', async (t) => {
  const app = await boot();
  t.after(() => app.close());
  app.agent.masters = fakeMasters;
  const call = client(app.url);
  await call.signIn();
  const { id, cards } = jobWithNames(app, [['rto', 'GO-DIGIT_OCT26_HCV_Mumbai'], ['rto', 'GO-DIGIT_OCT26_HCV_Pune'], ['rto', 'GO-DIGIT_OCT26_HCV_Surat']]);
  assert.equal((await call('POST', `/api/jobs/${id}/aliases/answer`, { cardIds: cards.slice(0, 2), to: 'NOPE' })).status, 400);
  const r = await call('POST', `/api/jobs/${id}/aliases/answer`, { cardIds: cards.slice(0, 2), to: 'GO-DIGIT_OCT26_HCV_ROM3', scope: 'insurer' });
  assert.equal(r.status, 200, JSON.stringify(r.data));
  assert.equal(r.data.answered, 2);
  assert.deepEqual(r.data.job.cards.map((c) => c.status), ['answered', 'answered', 'open']);
  assert.equal(r.data.job.cards[0].rule.scope, 'insurer:Go Digit');
  const back = await call('POST', `/api/jobs/${id}/cards/${cards[0]}/reopen`, {});
  assert.equal(back.data.job.cards[0].status, 'open', 'one of them can be taken back on its own');
  assert.equal((await call('POST', '/api/jobs/0123456789abcdef/aliases/answer', { cardIds: [1], skip: true })).status, 404);
});
