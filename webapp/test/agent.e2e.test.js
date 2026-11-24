/**
 * The agent against the real converters and the regression grids: a question becomes a card, the answer is remembered, the next job answers itself,
 * and what the web path writes is byte-for-byte what the command line writes.
 */
import assert from 'node:assert/strict';
import test from 'node:test';
import path from 'node:path';
import { BASELINE, CAN_RUN, INPUTS, newJob, openCards, runJob, setup, sha256 } from './helpers.js';

const skip = !CAN_RUN && 'the converters or the regression grids are not available here';
const JULY = { fromDate: '01/07/2026', toDate: '20/03/2030' };
const julyFile = path.join(INPUTS, 'liberty_jul26.xlsx');

test('a question of the converter becomes a card; the answer is remembered; the next job answers itself', { skip, timeout: 900000 }, async () => {
  const env = setup();

  // ---- first job: nothing is known, so the person is asked
  const id1 = newJob(env, 'liberty', julyFile, JULY);
  let job = await runJob(env, id1);
  assert.equal(job.status, 'needs_input', job.error || 'waiting for the answer');
  const asked = openCards(env, id1, 'question');
  assert.equal(asked.length, 1);
  const options = asked[0].payload.options;
  const electric = options.find((o) => /electric/i.test(o.label) && /petrol/i.test(o.label));
  assert.ok(electric, `an option "Electric same as Petrol" among: ${options.map((o) => o.label).join(' | ')}`);

  const r = env.agent.answer(id1, asked[0].id, { pick: [electric.n], scope: 'script', explanation: 'The July grid has no Electric column' });
  assert.equal(r.started, true, 'answering the only open question starts the run again');
  await env.agent.idle();
  job = env.store.getJob(id1);
  assert.ok(['review', 'done'].includes(job.status), `${job.status}: ${job.error}`);
  assert.equal(job.outputs.rows, BASELINE.liberty_jul26 ? 23154 : job.outputs.rows);
  assert.equal(sha256(env.agent.outputPath(id1, 'upload')), BASELINE.liberty_jul26.sha256, 'the web path writes the same bytes as the command line');
  assert.ok(env.agent.outputPath(id1, 'review'), 'the review workbook is there too');

  const rules = env.memory.listRules({ kind: 'answer' });
  assert.equal(rules.length, 1);
  assert.equal(rules[0].scope, 'script:liberty');
  assert.deepEqual(rules[0].data.labels, [electric.label]);
  assert.equal(rules[0].explanation, 'The July grid has no Electric column');

  // ---- second job, the same situation: memory answers, nobody is asked
  const id2 = newJob(env, 'liberty', julyFile, JULY);
  job = await runJob(env, id2);
  assert.ok(['review', 'done'].includes(job.status), `${job.status}: ${job.error}`);
  assert.equal(openCards(env, id2, 'question').length, 0, 'no question is open');
  const auto = env.store.cards(id2).filter((c) => c.type === 'question' && c.status === 'auto');
  assert.equal(auto.length, 1);
  assert.equal(auto[0].rule_id, rules[0].id);
  assert.equal(sha256(env.agent.outputPath(id2, 'upload')), BASELINE.liberty_jul26.sha256, 'the same bytes again');
  assert.equal(env.store.applications(id2).length, 1, 'the audit trail names the rule');
  assert.equal(env.memory.getRule(rules[0].id).uses, 1);
  assert.ok(env.store.events(id2).some((e) => e.kind === 'memory'), 'the timeline says it was answered from memory');

  // ---- taking the answer back asks again, and not from memory
  env.agent.reopen(id2, auto[0].id);
  assert.equal(env.store.getJob(id2).status, 'needs_input');
  const again = env.store.getCardById(auto[0].id);
  assert.equal(again.status, 'open');
  assert.equal(again.payload.noAuto, true);
});

test('an advisor only suggests: the card stays open and the person decides; one that fails does not stop the job', { skip, timeout: 600000 }, async () => {
  const calls = [];
  const advisor = {
    async suggest({ payload }) {
      calls.push(payload.problem);
      const o = payload.options[1];
      return { answer: { pick: [o.n], labels: [o.label] }, reason: 'because' };
    },
  };
  const env = setup({ advisor });
  const id = newJob(env, 'liberty', julyFile, JULY);
  const job = await runJob(env, id);
  assert.equal(job.status, 'needs_input');
  assert.equal(calls.length, 1);
  const card = openCards(env, id, 'question')[0];
  assert.equal(card.suggestion.source, 'advisor');
  assert.equal(card.suggestion.reason, 'because');
  assert.equal(env.store.cards(id).filter((c) => c.status !== 'open').length, 0, 'nothing was answered by the advisor');

  const failing = setup({ advisor: { async suggest() { throw new Error('offline'); } } });
  const id2 = newJob(failing, 'liberty', julyFile, JULY);
  assert.equal((await runJob(failing, id2)).status, 'needs_input');
  assert.ok(failing.store.events(id2).some((e) => /advisor could not help: offline/.test(e.text)));
});

test('a run is refused while the job is running, and a bad answer is refused', { skip, timeout: 600000 }, async () => {
  const env = setup();
  const id = newJob(env, 'liberty', julyFile, JULY);
  env.agent.start(id);
  assert.throws(() => env.agent.rerun(id), /already running/);
  await env.agent.idle();
  const card = openCards(env, id, 'question')[0];
  assert.throws(() => env.agent.answer(id, card.id, { pick: [99] }), /choose from the options/);
  assert.throws(() => env.agent.answer(id, card.id, {}), /choose from the options/);
  assert.throws(() => env.agent.answer(id, 999999, { pick: [1] }), /No such card/);
  assert.equal(env.store.getCardById(card.id).status, 'open', 'a refused answer changes nothing');
});
