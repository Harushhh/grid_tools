/** Test client: the real app on an ephemeral port, a cookie jar, and a few shortcuts. */
import { createApp } from '../src/app.js';
import { config as base } from '../src/config.js';
import { tempDir } from './helpers.js';

export const PASSWORD = 'test-password-123';

export async function boot(over = {}) {
  const dataDir = tempDir('gridapi-');
  const config = { ...base, dataDir, appPassword: PASSWORD, insecure: false, sessionSecret: 's'.repeat(48), ...over };
  const made = createApp({ config, quiet: true });
  const server = await new Promise((resolve) => {
    const s = made.app.listen(0, '127.0.0.1', () => resolve(s));
  });
  const url = `http://127.0.0.1:${server.address().port}`;
  return {
    ...made, config, url,
    async close() {
      await made.agent.idle();
      await new Promise((resolve) => server.close(resolve));
      made.db.close();
    },
  };
}

export function client(url) {
  let cookie = '';
  const call = async (method, p, body, { headers = {}, raw = false } = {}) => {
    const h = { 'X-Requested-With': 'gridapp', ...headers };
    if (cookie) h.cookie = cookie;
    let payload;
    if (body instanceof FormData) payload = body;
    else if (body !== undefined) {
      h['content-type'] = 'application/json';
      payload = JSON.stringify(body);
    }
    const res = await fetch(url + p, { method, headers: h, body: payload });
    const set = res.headers.get('set-cookie');
    if (set) cookie = set.split(';')[0];
    if (raw) return res;
    const buf = Buffer.from(await res.arrayBuffer());
    let data = null;
    try {
      data = JSON.parse(buf.toString('utf8'));
    } catch {
      data = null;
    }
    return { status: res.status, data, buf, headers: res.headers };
  };
  call.signIn = async () => {
    const r = await call('POST', '/api/login', { password: PASSWORD });
    if (r.status !== 200) throw new Error(`sign-in failed: ${r.status}`);
  };
  call.cookie = () => cookie;
  return call;
}

export async function waitFor(read, done, timeoutMs = 240000, everyMs = 400) {
  const t0 = Date.now();
  for (;;) {
    const v = await read();
    if (done(v)) return v;
    if (Date.now() - t0 > timeoutMs) throw new Error(`timed out waiting; last: ${JSON.stringify(v).slice(0, 300)}`);
    await new Promise((r) => setTimeout(r, everyMs));
  }
}

export const settled = (job) => !['queued', 'running'].includes(job.status);

/** multipart form for POST /api/jobs */
export function jobForm({ script, fromDate, toDate, files = {}, fields = {} }) {
  const form = new FormData();
  form.append('script', script);
  form.append('fromDate', fromDate);
  form.append('toDate', toDate);
  for (const [role, spec] of Object.entries(files)) form.append(role, new Blob([spec.data]), spec.name);
  for (const [k, v] of Object.entries(fields)) form.append(k, v);
  return form;
}
