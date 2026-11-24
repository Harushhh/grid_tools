/** The optional advisor, against a stand-in for the API (the real service is not called from the tests). */
import assert from 'node:assert/strict';
import http from 'node:http';
import test from 'node:test';
import { createAdvisor } from '../src/advisor.js';

function standIn(reply) {
  const seen = [];
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', (d) => { body += d; });
    req.on('end', () => {
      seen.push({ url: req.url, headers: req.headers, body: JSON.parse(body || '{}') });
      const r = typeof reply === 'function' ? reply(seen.length) : reply;
      res.writeHead(r.status || 200, { 'content-type': 'application/json' });
      res.end(JSON.stringify(r.json ?? {}));
    });
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ seen, url: `http://127.0.0.1:${server.address().port}`, close: () => new Promise((r) => server.close(r)) })));
}

const QUESTION = {
  type: 'question',
  payload: { kind: 'choose', multi: false, problem: 'two sheets could be this month\'s grid', question: 'Which sheet is this month\'s?', options: [{ n: 1, label: 'GRID NOV' }, { n: 2, label: 'GRID NOV (2)' }, { n: 3, label: 'Stop' }] },
};
const say = (text) => ({ json: { content: [{ type: 'text', text }] } });

test('it sends the question and its options only, and returns a suggestion that names an existing option', async (t) => {
  const api = await standIn(say('Here: {"pick": 2, "reason": "The second sheet is the newer one."}'));
  t.after(api.close);
  const advisor = createAdvisor({ apiKey: 'test-key', model: 'some-model', baseUrl: api.url });
  const s = await advisor.suggest({ ...QUESTION, script: { id: 'x' } });
  assert.deepEqual(s.answer, { pick: [2], labels: ['GRID NOV (2)'] });
  assert.match(s.reason, /newer/);
  assert.equal(api.seen.length, 1);
  const { url, headers, body } = api.seen[0];
  assert.equal(url, '/v1/messages');
  assert.equal(headers['x-api-key'], 'test-key');
  assert.equal(headers['anthropic-version'], '2023-06-01');
  assert.equal(body.model, 'some-model');
  assert.equal(body.messages.length, 1);
  const sent = body.messages[0].content;
  assert.match(sent, /Which sheet is this month's\?/);
  assert.match(sent, /2\. GRID NOV \(2\)/);
  assert.doesNotMatch(JSON.stringify(body), /password|GRID_|\.xlsx/i, 'nothing but the question');
});

test('an answer that is not one of the options, or not JSON, is no suggestion; an error is an error', async (t) => {
  const outcomes = [say('{"pick": 9, "reason": "x"}'), say('I think the second one.'), say('{"pick": "two"}'), { status: 500, json: { error: 'down' } }];
  const api = await standIn((n) => outcomes[n - 1]);
  t.after(api.close);
  const advisor = createAdvisor({ apiKey: 'k', baseUrl: api.url });
  assert.equal(await advisor.suggest(QUESTION), null);
  assert.equal(await advisor.suggest(QUESTION), null);
  assert.equal(await advisor.suggest(QUESTION), null);
  await assert.rejects(() => advisor.suggest(QUESTION), /500/);
});

test('it never guesses a typed answer, a multiple choice, a name or a password', async (t) => {
  const api = await standIn(say('{"pick": 1}'));
  t.after(api.close);
  const advisor = createAdvisor({ apiKey: 'k', baseUrl: api.url });
  assert.equal(await advisor.suggest({ type: 'question', payload: { ...QUESTION.payload, kind: 'text', options: [] } }), null);
  assert.equal(await advisor.suggest({ type: 'question', payload: { ...QUESTION.payload, multi: true } }), null);
  assert.equal(await advisor.suggest({ type: 'alias', payload: { kind: 'rto', name: 'X' } }), null);
  assert.equal(await advisor.suggest({ type: 'password', payload: {} }), null);
  assert.equal(api.seen.length, 0, 'no request was made for any of them');
});
