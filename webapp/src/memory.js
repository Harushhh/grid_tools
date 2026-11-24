/**
 * The memory ("brain") of the pipeline: a registry of what the person explained, and the matching that decides, for a new situation,
 * whether a past explanation applies by itself, is only worth suggesting, or does not apply.
 *
 * Rule kinds
 *   answer           the answer to a question of a converter (choose an option / type a text), found again by the question's fingerprint or by similarity
 *   alias            a name the master lacks -> the master group(s) it stands for
 *   standing_note    a note of the grid that needs no action (the converter stops flagging it)
 *   note_explanation what the person said about a note that stays flagged (shown, never changes a run; found again for the same note, or for the same issue with other numbers in it)
 *   failure_note     what the person said about a grid the converter stopped on (shown the next time it stops the same way, never changes a run)
 *
 * Safety: an answer is applied by itself only when the question is the same (fingerprint) or very similar AND the option that was chosen is in the new
 * question under the same label (or the same apart from a month written in it); a name mapping only when the name is the same apart from capitals / underscores / the month in it and the target exists in the master; a standing
 * note only when the note is the same text apart from spaces, capitals, punctuation and the numbers in it (so a reworded rule is never waved through; an RTO code or a list of codes still has to match).
 */
import { fromJson, now, toJson } from './db.js';
import { getScript } from './registry.js';

const MONTH = '(?:jan|feb|mar|apr|may|jun|jul|aug|sept|sep|oct|nov|dec)';
const PERIOD_RX = new RegExp(`(?<![a-z0-9])${MONTH}[-_ ]?\\d{2}(?:\\d{2})?(?![a-z0-9])`, 'i');

/** Wording without what changes from month to month (file names, numbers) and without punctuation. */
export function normText(s, keepDigits = false) {
  let t = String(s ?? '').toLowerCase().replace(/[\w\-./:()\\]*\.(?:xlsx|xlsm|xlsb|xls|csv|pdf|json)\b/g, '<file>');
  if (!keepDigits) t = t.replace(/\d+/g, '#');
  return t.replace(/[^a-z0-9#<>]+/g, ' ').replace(/\s+/g, ' ').trim();
}

const tokenSet = (s) => new Set(s.split(' ').filter(Boolean));

function jaccard(a, b) {
  if (!a.size && !b.size) return 1;
  let inter = 0;
  for (const x of a) if (b.has(x)) inter += 1;
  return inter / (a.size + b.size - inter);
}

function trigrams(s) {
  const t = `  ${s} `;
  const m = new Map();
  for (let i = 0; i < t.length - 2; i += 1) {
    const g = t.slice(i, i + 3);
    m.set(g, (m.get(g) || 0) + 1);
  }
  return m;
}

function dice(a, b) {
  let inter = 0;
  let na = 0;
  let nb = 0;
  for (const v of a.values()) na += v;
  for (const v of b.values()) nb += v;
  for (const [g, v] of a) inter += Math.min(v, b.get(g) || 0);
  return na + nb === 0 ? 1 : (2 * inter) / (na + nb);
}

/** 0..1: 1 = the same wording (apart from case, punctuation, and - unless keepDigits - numbers and file names). */
export function similarity(a, b, keepDigits = false) {
  const x = normText(a, keepDigits);
  const y = normText(b, keepDigits);
  if (x === y) return 1;
  return 0.5 * jaccard(tokenSet(x), tokenSet(y)) + 0.5 * dice(trigrams(x), trigrams(y));
}

/** An option label for comparing: spaces and capitals ignored. */
const labelKey = (s) => String(s ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
const PERIOD_RX_ALL = new RegExp(PERIOD_RX.source, 'gi');
/** ... and with a month written in it ('OCT26', 'Oct 2026') taken out: the same label of another month. */
const periodFold = (s) => labelKey(s).replace(PERIOD_RX_ALL, '{p}');

/** The master's comparison of names: capitals, stray leading/trailing '_' or spaces ignored. */
export const foldName = (s) => String(s ?? '').trim().replace(/^[ _]+|[ _]+$/g, '').toLowerCase();
const alnum = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '');
/** A note without spaces, punctuation and capitals - the converters compare notes this way. */
export const noteKey = (s) => alnum(s);
/** The issue a note is about: the same wording with the numbers folded away ('Few RTOs for 1000-1500CC are declined' and '... 1000-1250CC ...' are one issue). Groups notes, never silences one. */
export const issueKey = (s) => normText(s);

const STATE_CODE = /\b(?:AN|AP|AR|AS|BR|CG|CH|DD|DL|DN|GA|GJ|HP|HR|JH|JK|KA|KL|LA|LD|MH|ML|MN|MP|MZ|NL|OD|OR|PB|PY|RJ|SK|TG|TN|TR|TS|UA|UK|UP|WB)\s*-?\s*\d{1,3}\b/g;
const CODE_LIST = /\b\d{1,3}(?:\s*[,/&]\s*\d{1,3}){2,}\b/g;
/**
 * A note with its numbers folded to '#': 'PO will be 5% less' and 'PO will be 7% less' are ONE note (the user, 31 Oct 2026: the condition is the same, only the percentage changes).
 * The numbers that ARE the condition stay: an RTO code ('RJ-14') and a list of three or more numbers ('EXL RTO codes 03, 06, 18'), so another RTO is another note. Mirrors fold_note of the converters.
 */
export function foldNote(text) {
  const keep = [];
  const hold = (m) => {
    keep.push(m);
    return `\u0001${'k'.repeat(keep.length)}\u0001`;
  };
  const t = String(text ?? '').replace(/\s+/g, ' ').replace(CODE_LIST, hold).replace(STATE_CODE, hold).replace(/\d+(?:\.\d+)?/g, '#');
  return t.replace(/\u0001(k+)\u0001/g, (_, ks) => keep[ks.length - 1]);
}
export const foldKey = (s) => noteKey(foldNote(s));

/** 'GO-DIGIT_OCT26_HCV_Mumbai' -> { template: 'GO-DIGIT_{P}_HCV_Mumbai', token: 'OCT26' } */
export function periodTemplate(name) {
  const m = PERIOD_RX.exec(String(name));
  if (!m) return null;
  return { template: String(name).slice(0, m.index) + '{P}' + String(name).slice(m.index + m[0].length), token: m[0] };
}

const splitNames = (to) => String(to).split(',').map((x) => x.trim()).filter(Boolean);

export class Memory {
  constructor(db, { autoApply = 0.9, suggest = 0.65 } = {}) {
    this.db = db;
    this.autoApply = autoApply;
    this.suggestAt = suggest;
  }

  // ------------------------------------------------------------------ the registry
  #row(r) {
    return r && { ...r, data: fromJson(r.data, {}), enabled: !!r.enabled };
  }

  getRule(id) {
    return this.#row(this.db.prepare('SELECT * FROM rules WHERE id = ?').get(id));
  }

  addRule({ kind, scope = 'global', fp = null, matchText = '', data, explanation = '', sourceJob = null }) {
    const stamp = now();
    const same = this.db.prepare('SELECT id FROM rules WHERE kind = ? AND scope = ? AND COALESCE(fp, \'\') = ? AND match_text = ? AND (kind NOT IN (\'alias\', \'standing_note\') OR data = ?)')
      .get(kind, scope, fp || '', matchText, toJson(data));
    if (same) {
      this.db.prepare('UPDATE rules SET data = ?, explanation = CASE WHEN ? <> \'\' THEN ? ELSE explanation END, enabled = 1, updated_at = ?, source_job = COALESCE(?, source_job) WHERE id = ?')
        .run(toJson(data), explanation, explanation, stamp, sourceJob, same.id);
      return Number(same.id);
    }
    if (kind === 'alias') {      // one mapping per name and kind in a scope: a new one replaces the old
      this.db.prepare('DELETE FROM rules WHERE kind = \'alias\' AND scope = ? AND match_text = ?').run(scope, matchText);
    }
    const info = this.db.prepare('INSERT INTO rules (kind, scope, fp, match_text, data, explanation, enabled, uses, created_at, updated_at, source_job) VALUES (?, ?, ?, ?, ?, ?, 1, 0, ?, ?, ?)')
      .run(kind, scope, fp, matchText, toJson(data), explanation || '', stamp, stamp, sourceJob);
    return Number(info.lastInsertRowid);
  }

  listRules({ kind, q, enabled } = {}) {
    let rows = this.db.prepare('SELECT * FROM rules ORDER BY updated_at DESC, id DESC').all().map((r) => this.#row(r));
    if (kind) rows = rows.filter((r) => r.kind === kind);
    if (enabled !== undefined) rows = rows.filter((r) => r.enabled === enabled);
    if (q) {
      const needle = String(q).toLowerCase();
      rows = rows.filter((r) => `${r.match_text} ${JSON.stringify(r.data)} ${r.explanation} ${r.scope}`.toLowerCase().includes(needle));
    }
    return rows;
  }

  setEnabled(id, on) {
    this.db.prepare('UPDATE rules SET enabled = ?, updated_at = ? WHERE id = ?').run(on ? 1 : 0, now(), id);
  }

  updateRule(id, { explanation, scope, data }) {
    const r = this.getRule(id);
    if (!r) return null;
    this.db.prepare('UPDATE rules SET explanation = ?, scope = ?, data = ?, updated_at = ? WHERE id = ?')
      .run(explanation ?? r.explanation, scope ?? r.scope, toJson(data ?? r.data), now(), id);
    return this.getRule(id);
  }

  deleteRule(id) {
    this.db.prepare('DELETE FROM rules WHERE id = ?').run(id);
  }

  recordUse(ruleId, { jobId, key = '', mode, score, detail = '', run = null }) {
    this.db.prepare('UPDATE rules SET uses = uses + 1, last_used_at = ? WHERE id = ?').run(now(), ruleId);
    this.db.prepare('INSERT OR IGNORE INTO applications (job_id, rule_id, key, mode, score, detail, run, applied_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
      .run(jobId, ruleId, key, mode, score, detail, run, now());
  }

  export() {
    return this.listRules().map(({ id, uses, last_used_at, source_job, ...r }) => ({ ...r, enabled: r.enabled }));
  }

  /** A rule from a file: only the shapes the matching code relies on are accepted. */
  static valid(r) {
    const d = r?.data;
    const text = (x) => typeof x === 'string' && x.length > 0 && x.length < 5000;
    if (!d || typeof d !== 'object') return false;
    switch (r.kind) {
      case 'answer': return (d.kind === 'text' && typeof d.text === 'string') || (d.kind === 'choose' && Array.isArray(d.labels) && d.labels.length > 0 && d.labels.every(text));
      case 'alias': return ['rto', 'makes'].includes(d.kind) && text(d.from) && text(d.to);
      case 'standing_note':
      case 'note_explanation': return text(d.text);
      case 'failure_note': return text(d.error);
      default: return false;
    }
  }

  import(rules) {
    let n = 0;
    for (const r of rules || []) {
      if (!Memory.valid(r)) continue;
      const id = this.addRule({ kind: r.kind, scope: r.scope || 'global', fp: r.fp || null, matchText: r.match_text || '', data: r.data, explanation: r.explanation || '' });
      if (r.enabled === false) this.setEnabled(id, false);
      n += 1;
    }
    return n;
  }

  // ------------------------------------------------------------------ which rules may apply to a script
  #scopeOk(scope, scriptId) {
    if (scope === 'global') return true;
    if (scope === `script:${scriptId}`) return true;
    const s = getScript(scriptId);
    return !!s && scope === `insurer:${s.insurer}`;
  }

  #candidates(kind, scriptId) {
    return this.db.prepare('SELECT * FROM rules WHERE kind = ? AND enabled = 1 ORDER BY updated_at ASC, id ASC').all(kind)
      .map((r) => this.#row(r)).filter((r) => this.#scopeOk(r.scope, scriptId));
  }

  // ------------------------------------------------------------------ answers to the converters' questions
  /**
   * The stored answer of a rule translated to this question: { answer: { pick: [n] } | { text }, sure }, or null when the option that was chosen is not here.
   * `sure` = every chosen option was found by its label - the same label, or the same label apart from a month written in it ('OCT26' -> 'NOV26') that picks exactly one
   * option. A label that is merely CLOSE (a number or a letter differs: '< 1000 CC' / '< 1500 CC' may be another column) is returned with sure = false: a suggestion, never applied by itself.
   */
  applyAnswer(rule, q) {
    if (rule.data.kind === 'text') return q.kind === 'text' && rule.data.text !== undefined ? { answer: { text: rule.data.text }, sure: true } : null;
    if (q.kind !== 'choose') return null;
    const options = q.options || [];
    const picks = [];
    let sure = true;
    for (const label of rule.data.labels || []) {
      let hit = options.filter((o) => labelKey(o.label) === labelKey(label));
      if (hit.length > 1) return null;                                   // two options with the same label: nothing says which
      if (!hit.length) hit = options.filter((o) => periodFold(o.label) === periodFold(label));
      if (hit.length > 1) return null;                                   // 'OCT26 grid' and 'NOV26 grid' in one question: nothing says which
      if (!hit.length) {
        let best = 0;
        for (const o of options) {
          const s = similarity(o.label, label, true);
          if (s > best) {
            best = s;
            hit = [o];
          }
        }
        if (best < 0.9) return null;
        sure = false;
      }
      if (!picks.includes(hit[0].n)) picks.push(hit[0].n);
    }
    if (!picks.length || (!q.multi && picks.length !== 1)) return null;
    return { answer: { pick: picks }, sure };
  }

  /**
   * Memory first: is there a past answer for this question?
   * -> { auto: { rule, mode: 'exact'|'similar', score, answer } }   apply it, say so, allow undo
   *  | { suggestion: { rule, score, answer|null } }                  show it on the card, one click to use
   *  | {}                                                            nothing known
   */
  matchQuestion(scriptId, q) {
    const matchText = normText(`${q.problem} ${(q.options || []).map((o) => o.label).join(' ')}`);
    const found = [];
    for (const rule of this.#candidates('answer', scriptId)) {
      if (rule.data.kind !== q.kind) continue;
      const exact = !!rule.fp && rule.fp === q.fp;
      const score = exact ? 1 : similarity(matchText, rule.match_text);
      if (score < this.suggestAt) continue;
      const applied = this.applyAnswer(rule, q);
      found.push({ rule, exact, score, answer: applied ? applied.answer : null, sure: !!applied?.sure });
    }
    found.sort((a, b) => b.score - a.score || (a.rule.scope === 'global') - (b.rule.scope === 'global') || b.rule.id - a.rule.id);
    const usable = found.find((f) => f.answer && f.sure && f.score >= this.autoApply);
    if (usable) return { auto: { rule: usable.rule, mode: usable.exact ? 'exact' : 'similar', score: usable.score, answer: usable.answer } };
    const first = found[0];
    return first ? { suggestion: { rule: first.rule, score: first.score, answer: first.answer, close: !!first.answer && !first.sure } } : {};
  }

  // ------------------------------------------------------------------ stops the converter cannot get past
  /** What the person said the last time the converter stopped with (about) this message - shown on the card, never acted on. */
  matchFailure(scriptId, error) {
    const text = normText(error);
    let best = null;
    for (const rule of this.#candidates('failure_note', scriptId)) {
      const score = similarity(text, rule.match_text);
      if (score >= 0.8 && (!best || score > best.score)) best = { rule, score };
    }
    return best;
  }

  // ------------------------------------------------------------------ names the master lacks
  /** Is there a mapping for this name?  `masterNames` (a Set) is used to check that the target exists; without it a mapping is never applied by itself. */
  matchAlias(scriptId, kind, name, masterNames) {
    const rules = this.#candidates('alias', scriptId).filter((r) => r.data.kind === kind);
    const known = (to) => !!masterNames && splitNames(to).length > 0 && splitNames(to).every((n) => masterNames.has(n));
    const f = foldName(name);
    const a = alnum(name);
    for (const r of rules) {
      if (foldName(r.data.from) === f) return known(r.data.to) ? { auto: { rule: r, mode: 'exact', score: 1, to: r.data.to } } : { suggestion: { rule: r, score: 1, to: r.data.to } };
    }
    for (const r of rules) {
      if (alnum(r.data.from) === a) return known(r.data.to) ? { auto: { rule: r, mode: 'similar', score: 0.98, to: r.data.to } } : { suggestion: { rule: r, score: 0.98, to: r.data.to } };
    }
    const t = periodTemplate(name);
    if (t) {
      for (const r of rules) {
        const tpl = r.data.template;
        if (tpl && foldName(tpl.from) === foldName(t.template)) {
          const to = tpl.to.replaceAll('{P}', t.token);
          return known(to) ? { auto: { rule: r, mode: 'similar', score: 0.95, to } } : { suggestion: { rule: r, score: 0.95, to } };
        }
      }
    }
    let best = null;
    for (const r of rules) {
      const s = similarity(r.data.from, name, true);
      if (s >= 0.9 && (!best || s > best.score)) best = { rule: r, score: s, to: r.data.to };
    }
    return best ? { suggestion: best } : {};
  }

  /** { rto: { from: to }, makes: {...} } of every enabled mapping that applies to the script (the converters apply the exact ones themselves). */
  aliasFile(scriptId) {
    const out = { rto: {}, makes: {} };
    for (const r of this.#candidates('alias', scriptId)) {
      if (out[r.data.kind]) out[r.data.kind][r.data.from] = r.data.to;
    }
    return out;
  }

  static aliasData(kind, from, to) {
    const data = { kind, from, to };
    const t = periodTemplate(from);
    if (t && String(to).includes(t.token)) data.template = { from: t.template, to: String(to).replaceAll(t.token, '{P}') };
    return data;
  }

  // ------------------------------------------------------------------ notes
  standingNotes(scriptId) {
    return this.#candidates('standing_note', scriptId).map((r) => r.data.text);
  }

  isStandingNote(scriptId, text) {
    const k = foldKey(text);
    return this.#candidates('standing_note', scriptId).find((r) => foldKey(r.data.text) === k) || null;
  }

  /** What the person resolved for this converter - explained, or said needs no action: [{ text, rule, action: 'explain' | 'ignore' }]. The converters never flag these again (numbers folded). */
  resolved(scriptId) {
    return [
      ...this.#candidates('standing_note', scriptId).map((rule) => ({ text: rule.data.text, rule, action: 'ignore' })),
      ...this.#candidates('note_explanation', scriptId).map((rule) => ({ text: rule.data.text, rule, action: 'explain' })),
    ];
  }

  resolvedFor(scriptId, text) {
    const k = foldKey(text);
    return this.resolved(scriptId).find((x) => foldKey(x.text) === k) || null;
  }

  noteExplanation(scriptId, text) {
    const k = noteKey(text);
    return this.#candidates('note_explanation', scriptId).find((r) => noteKey(r.data.text) === k) || null;
  }

  /** An explanation given for the same issue with other numbers in it, or null. It is only shown (an explanation never changes a run), so this is safe where a standing note is not. */
  noteExplanationForIssue(scriptId, text) {
    const k = issueKey(text);
    return this.#candidates('note_explanation', scriptId).find((r) => issueKey(r.data.text) === k) || null;
  }
}
