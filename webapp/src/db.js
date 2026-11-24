import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';

const SCHEMA = `
CREATE TABLE IF NOT EXISTS jobs (
  id TEXT PRIMARY KEY,
  script TEXT NOT NULL,
  status TEXT NOT NULL,                -- queued | running | needs_input | done | failed
  params TEXT NOT NULL,                -- dates, options (never a password)
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  run_count INTEGER NOT NULL DEFAULT 0,
  error TEXT,
  report TEXT,                         -- the last successful run's report (JSON)
  outputs TEXT,                        -- {upload, review} file names of the last successful run
  log TEXT                             -- tail of the last run's output
);
CREATE TABLE IF NOT EXISTS job_files (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  job_id TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
  role TEXT NOT NULL,
  orig_name TEXT NOT NULL,
  stored_name TEXT NOT NULL,
  size INTEGER NOT NULL
);
-- A card is one thing the app needs from the person: a question of the converter, a name to map, a note to judge, a password.
CREATE TABLE IF NOT EXISTS cards (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  job_id TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
  type TEXT NOT NULL,                  -- question | alias | note | password
  key TEXT NOT NULL,                   -- the question id / the name / the note text
  payload TEXT NOT NULL,
  status TEXT NOT NULL,                -- open | answered | auto
  answer TEXT,
  rule_id INTEGER,
  suggestion TEXT,
  created_at TEXT NOT NULL,
  answered_at TEXT,
  UNIQUE (job_id, type, key)
);
-- What the agent did and why, in order: the timeline of a job.
CREATE TABLE IF NOT EXISTS events (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  job_id TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
  at TEXT NOT NULL,
  kind TEXT NOT NULL,                  -- run | memory | ask | info | error
  text TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS events_job ON events (job_id, id);
-- The memory ("brain"): what the person explained, to be applied the next time the same or a similar thing comes up.
CREATE TABLE IF NOT EXISTS rules (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL,                  -- answer | alias | standing_note | note_explanation
  scope TEXT NOT NULL,                 -- global | insurer:<name> | script:<id>
  fp TEXT,
  match_text TEXT NOT NULL,
  data TEXT NOT NULL,
  explanation TEXT,
  enabled INTEGER NOT NULL DEFAULT 1,
  uses INTEGER NOT NULL DEFAULT 0,
  last_used_at TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  source_job TEXT
);
CREATE INDEX IF NOT EXISTS rules_kind ON rules (kind, enabled);
-- Every time a rule was applied by itself: the audit trail (and the way to take it back).
CREATE TABLE IF NOT EXISTS applications (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  job_id TEXT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
  rule_id INTEGER NOT NULL,
  key TEXT,
  mode TEXT NOT NULL,                  -- exact | similar | alias | note
  score REAL NOT NULL,
  detail TEXT,
  run INTEGER,
  applied_at TEXT NOT NULL,
  UNIQUE (job_id, rule_id, key)
);
`;

export function openDb(file) {
  if (file !== ':memory:') fs.mkdirSync(path.dirname(file), { recursive: true });
  const db = new DatabaseSync(file);
  db.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON; PRAGMA busy_timeout = 5000;');
  db.exec(SCHEMA);
  return db;
}

export const now = () => new Date().toISOString();
export const toJson = (v) => JSON.stringify(v ?? null);
export const fromJson = (s, fallback = null) => {
  try {
    return s == null ? fallback : JSON.parse(s);
  } catch {
    return fallback;
  }
};
