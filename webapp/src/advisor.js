/**
 * Optional advisor: a language model that may SUGGEST an answer for a question memory knows nothing about.
 * Off unless ADVISOR=1 and ANTHROPIC_API_KEY are set. It only suggests - the card stays open and the person decides; a suggestion is never applied by itself.
 * It is given the text of the question and its options, nothing else (no grid file, no password, no rates), and only for multiple-choice questions:
 * a typed answer (a folder, a name) is never guessed.
 */
const SYSTEM = [
  'You help an insurance broker turn an insurer\'s rate grid into an upload file with a converter script.',
  'The script has stopped with a question that it cannot decide by itself. You are shown the question and numbered options.',
  'Choose the option that is the most careful reading of the situation: never invent rates, never pick an option that silently drops or guesses data when another option keeps it visible for a person to check.',
  'If the question cannot be answered from its own text, choose the option that stops the run, if there is one.',
  'Reply with JSON only: {"pick": <the option number>, "reason": "<one short sentence>"}',
].join(' ');

export function createAdvisor({ apiKey, model = 'claude-sonnet-5-5', baseUrl = 'https://api.anthropic.com', timeoutMs = 30000, fetchImpl = fetch }) {
  return {
    async suggest({ type, payload }) {
      if (type !== 'question' || payload.kind !== 'choose' || payload.multi || !Array.isArray(payload.options) || payload.options.length < 2) return null;
      const prompt = `Problem: ${String(payload.problem).slice(0, 1500)}\nQuestion: ${String(payload.question).slice(0, 1500)}\nOptions:\n${payload.options.map((o) => `${o.n}. ${String(o.label).slice(0, 300)}`).join('\n')}`;
      const res = await fetchImpl(`${baseUrl}/v1/messages`, {
        method: 'POST',
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
        body: JSON.stringify({ model, max_tokens: 300, system: SYSTEM, messages: [{ role: 'user', content: prompt }] }),
        signal: AbortSignal.timeout(timeoutMs),
      });
      if (!res.ok) throw new Error(`the advisor service answered ${res.status}`);
      const data = await res.json();
      const text = (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('\n');
      const json = /\{[\s\S]*\}/.exec(text);
      if (!json) return null;
      let parsed;
      try {
        parsed = JSON.parse(json[0]);
      } catch {
        return null;
      }
      const pick = Number(parsed.pick);
      const option = payload.options.find((o) => o.n === pick);
      if (!Number.isInteger(pick) || !option) return null;
      return { answer: { pick: [pick], labels: [option.label] }, reason: String(parsed.reason || '').slice(0, 300) };
    },
  };
}
