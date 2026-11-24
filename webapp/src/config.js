import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

export const ROOT = path.resolve(here, '..');
/** The folder with the converter scripts (*_to_config.py), master/ and the .venv. */
export const GRID_TOOLS = path.resolve(process.env.GRID_TOOLS_DIR || path.join(ROOT, '..'));

function findPython() {
  if (process.env.PYTHON) return process.env.PYTHON;
  for (const p of [path.join(GRID_TOOLS, '.venv', 'Scripts', 'python.exe'), path.join(GRID_TOOLS, '.venv', 'bin', 'python')]) {
    if (fs.existsSync(p)) return p;
  }
  return process.platform === 'win32' ? 'python' : 'python3';
}

export const config = {
  port: Number(process.env.PORT || 3000),
  gridTools: GRID_TOOLS,
  dataDir: path.resolve(process.env.DATA_DIR || path.join(ROOT, 'data')),
  python: findPython(),
  /** The one password of the app. Without it the server refuses to start (unless GRID_APP_INSECURE=1, for local development). */
  appPassword: process.env.GRID_APP_PASSWORD || '',
  insecure: process.env.GRID_APP_INSECURE === '1',
  sessionSecret: process.env.SESSION_SECRET || crypto.randomBytes(32).toString('hex'),
  sessionHours: Number(process.env.SESSION_HOURS || 12),
  secureCookie: process.env.COOKIE_SECURE === '1',
  maxUploadMb: Number(process.env.MAX_UPLOAD_MB || 150),
  runTimeoutMs: Number(process.env.RUN_TIMEOUT_MS || 20 * 60 * 1000),
  maxParallelRuns: Number(process.env.MAX_PARALLEL_RUNS || 2),
  /** '' (default): every converter uses its own mode - 'cluster' (only the cluster name in new_rto_list, no RTO name asked) for most, 'master' (the master's name, or a plain-text new-mapping warning) for Shriram; 'group' is the earlier mode (the master's group names, RTO name cards). Handed on to the converters when set. */
  rtoNames: ['group', 'cluster', 'master'].includes(process.env.GRID_RTO_NAMES) ? process.env.GRID_RTO_NAMES : '',
  maxLoops: 14,
  /** Memory thresholds: a past answer is applied by itself from `autoApply`, offered as a suggestion from `suggest`. */
  autoApply: Number(process.env.MEMORY_AUTO_APPLY || 0.9),
  suggest: Number(process.env.MEMORY_SUGGEST || 0.65),
};

/** Variables of this app that must never reach the converter scripts. */
export const SECRET_ENV = ['GRID_APP_PASSWORD', 'SESSION_SECRET', 'ANTHROPIC_API_KEY'];

/**
 * The Django insurance portal's own settings (the keys of its .env). This app is a separate project: it never reads them and never hands them on, even if they happen to be in
 * the shell it was started from. (test/isolation.test.js keeps it that way.)
 */
export const PORTAL_ENV = [
  'DJANGO_SETTINGS_MODULE', 'SECRET_KEY', 'DEBUG', 'ALLOWED_HOSTS', 'CSRF_TRUSTED_ORIGINS', 'PARTNER_SSO_TICKET_SECRET',
  'DATABASE_URL', 'DB_HOST', 'DB_PORT', 'DB_NAME', 'DB_USER', 'DB_PASSWORD', 'REDIS_URL',
  'AWS_STORAGE_BUCKET_NAME', 'AWS_ACCESS_KEY_ID', 'AWS_SECRET_ACCESS_KEY', 'AWS_S3_ENDPOINT_URL', 'GRID_ARCHIVE_BUCKET_NAME',
  'EMAIL_HOST', 'EMAIL_PORT', 'EMAIL_USE_TLS', 'EMAIL_USE_SSL', 'EMAIL_HOST_USER', 'EMAIL_HOST_PASSWORD', 'DEFAULT_FROM_EMAIL', 'GEMINI_API_KEY',
];
