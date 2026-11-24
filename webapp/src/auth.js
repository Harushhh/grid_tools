/**
 * One shared password, a signed session cookie, and the guards around state-changing requests.
 * The cookie is HttpOnly + SameSite=Strict (+ Secure behind https); a state-changing request must also carry a header that a page of another site cannot send.
 */
import crypto from 'node:crypto';

export const COOKIE = 'gridsid';
export const CSRF_HEADER = 'x-requested-with';
export const CSRF_VALUE = 'gridapp';

const sign = (secret, payload) => crypto.createHmac('sha256', secret).update(payload).digest('base64url');
const sha = (s) => crypto.createHash('sha256').update(String(s)).digest();

function cookiesOf(req) {
  const out = Object.create(null);
  for (const part of String(req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = part.slice(i + 1).trim();
  }
  return out;
}

export function createAuth({ password, secret, hours = 12, secure = false, insecure = false, maxAttempts = 8, windowMs = 15 * 60 * 1000 }) {
  const attempts = new Map();                         // ip -> { n, since }
  const expected = sha(password || '');

  const issue = () => {
    const payload = Buffer.from(JSON.stringify({ exp: Date.now() + hours * 3600 * 1000, n: crypto.randomBytes(8).toString('hex') })).toString('base64url');
    return `${payload}.${sign(secret, payload)}`;
  };

  const valid = (token) => {
    const [payload, sig] = String(token || '').split('.');
    if (!payload || !sig) return false;
    const a = Buffer.from(sig);
    const b = Buffer.from(sign(secret, payload));
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return false;
    try {
      return JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')).exp > Date.now();
    } catch {
      return false;
    }
  };

  const cookie = (value, maxAge) => `${COOKIE}=${value}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${maxAge}${secure ? '; Secure' : ''}`;
  const isAuthed = (req) => insecure || valid(cookiesOf(req)[COOKIE]);

  return {
    required: !insecure,
    isAuthed,

    /** Everything under /api except login / me needs a session. */
    requireSession(req, res, next) {
      if (isAuthed(req)) return next();
      return res.status(401).json({ error: 'Please sign in.' });
    },

    /** A state-changing request must come from this site's own page. */
    requireSameSite(req, res, next) {
      if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) return next();
      const origin = req.headers.origin;
      let sameOrigin = true;
      if (origin) {
        try {
          sameOrigin = new URL(origin).host === req.headers.host;
        } catch {
          sameOrigin = false;
        }
      }
      if (!sameOrigin || req.headers[CSRF_HEADER] !== CSRF_VALUE) return res.status(403).json({ error: 'This request did not come from the app.' });
      return next();
    },

    login(req, res) {
      const ip = req.ip || 'unknown';
      const now = Date.now();
      if (attempts.size > 5000) for (const [k, v] of attempts) if (now - v.since > windowMs) attempts.delete(k);       // the table cannot grow without limit
      const rec = attempts.get(ip);
      if (rec && now - rec.since > windowMs) attempts.delete(ip);
      const cur = attempts.get(ip) || { n: 0, since: now };
      if (cur.n >= maxAttempts) {
        res.set('Retry-After', String(Math.ceil((cur.since + windowMs - now) / 1000)));
        return res.status(429).json({ error: 'Too many attempts. Try again in a few minutes.' });
      }
      const given = sha(req.body?.password ?? '');
      if (!password || !crypto.timingSafeEqual(given, expected)) {
        cur.n += 1;
        attempts.set(ip, cur);
        return res.status(401).json({ error: 'That password is not right.' });
      }
      attempts.delete(ip);
      res.set('Set-Cookie', cookie(issue(), Math.floor(hours * 3600)));
      return res.json({ ok: true });
    },

    logout(req, res) {
      res.set('Set-Cookie', cookie('', 0));
      res.json({ ok: true });
    },
  };
}
