import assert from 'node:assert/strict';
import test from 'node:test';
import { openDb } from '../src/db.js';
import { Memory, foldKey, foldName, foldNote, noteKey, normText, periodTemplate, similarity } from '../src/memory.js';

const mem = () => new Memory(openDb(':memory:'));

const Q = (over = {}) => ({
  id: 'q1', fp: 'fp-sheet', kind: 'choose', problem: 'the commercial-vehicle grid: more than one sheet fits', question: 'Which sheet is this month\'s?', multi: false,
  options: [{ n: 1, label: 'LATEST CV COMBINED GRID' }, { n: 2, label: 'LATEST CV COMBINED GRID (2)' }, { n: 3, label: 'Stop' }], ...over,
});

test('normText drops numbers, file names and punctuation; similarity is 1 for the same wording', () => {
  assert.equal(normText('Cannot open Jul26_Grid_v2.xlsx: BadZipFile 12!'), 'cannot open <file> badzipfile #');
  assert.equal(similarity('cannot open Jul26.xlsx 1', 'Cannot open Aug26_v3.xlsx 2'), 1);
  assert.ok(similarity('which sheet is the rate table', 'which sheet holds the rate table') > 0.7);
  assert.ok(similarity('which sheet is the rate table', 'dates look wrong, to_date is in the past') < 0.4);
});

test('foldName / noteKey / periodTemplate', () => {
  assert.equal(foldName(' _GO-DIGIT_oct26_HCV_Bad_KA_ '), 'go-digit_oct26_hcv_bad_ka');
  assert.equal(noteKey('For Hazardous category, PO will be 5% less.'), noteKey('for hazardous category PO will be 5%  less'));
  assert.notEqual(noteKey('payout cannot be reduced'), noteKey('payout can be reduced'));
  assert.deepEqual(periodTemplate('GO-DIGIT_OCT26_HCV_Mumbai'), { template: 'GO-DIGIT_{P}_HCV_Mumbai', token: 'OCT26' });
  assert.equal(periodTemplate('GO-DIGIT_PC_TP_Bangalore'), null);
});

test('a stored answer is applied by itself to the same question, whatever the order of the options', () => {
  const m = mem();
  const id = m.addRule({ kind: 'answer', scope: 'script:godigit', fp: 'fp-sheet', matchText: normText('x'), data: { kind: 'choose', labels: ['LATEST CV COMBINED GRID (2)'] }, explanation: 'always the newer one' });
  const reordered = Q({ options: [{ n: 1, label: 'Stop' }, { n: 2, label: 'LATEST CV COMBINED GRID (2)' }, { n: 3, label: 'LATEST CV COMBINED GRID' }] });
  const r = m.matchQuestion('godigit', reordered);
  assert.equal(r.auto.rule.id, id);
  assert.equal(r.auto.mode, 'exact');
  assert.deepEqual(r.auto.answer, { pick: [2] });
  assert.deepEqual(m.matchQuestion('liberty', Q()), {}, 'a rule of another script does not apply');
});

test('a similar question is applied by itself only when the chosen option exists; otherwise it is only a suggestion', () => {
  const m = mem();
  const text = normText('the commercial-vehicle grid: more than one sheet fits LATEST CV COMBINED GRID LATEST CV COMBINED GRID (2) Stop');
  m.addRule({ kind: 'answer', scope: 'global', fp: 'old-fp', matchText: text, data: { kind: 'choose', labels: ['LATEST CV COMBINED GRID (2)'] } });
  const same = m.matchQuestion('godigit', Q({ fp: 'new-fp' }));
  assert.equal(same.auto.mode, 'similar');
  assert.deepEqual(same.auto.answer, { pick: [2] });
  const renamed = m.matchQuestion('godigit', Q({ fp: 'new-fp', options: [{ n: 1, label: 'NOV CV GRID' }, { n: 2, label: 'NOV CV GRID OLD' }, { n: 3, label: 'Stop' }] }));
  assert.equal(renamed.auto, undefined);
  assert.ok(renamed.suggestion, 'the question is similar enough to suggest');
  assert.equal(renamed.suggestion.answer, null, 'but the option that was chosen does not exist, so there is nothing to click');
});

test('only the same label - or the same label of another month - is applied by itself; a label that differs by a number is a suggestion to check', () => {
  const m = mem();
  const col = (cc, letter) => `Use the column 'Petrol / PC < ${cc} CC' (${letter}) of the table called Branch Grid for the private car own damage cover`;
  const ask = (labels, fp = 'fp-col') => ({ id: 'q', fp, kind: 'choose', problem: 'the column is not in the header rows', question: 'Which column?', options: labels.map((label, i) => ({ n: i + 1, label })) });
  m.addRule({ kind: 'answer', scope: 'global', fp: 'fp-col', matchText: normText(`the column is not in the header rows ${col(1000, 'D')} ${col(1500, 'F')}`), data: { kind: 'choose', labels: [col(1000, 'D')] } });
  // the same question, the same labels in another order: by itself
  assert.deepEqual(m.matchQuestion('liberty', ask([col(1500, 'F'), col(1000, 'D')])).auto.answer, { pick: [2] });
  // the same question, but the chosen label now says 1500: another column, maybe. Not applied; offered, marked as only close
  const changed = m.matchQuestion('liberty', ask([col(1500, 'D'), col(2000, 'F')]));
  assert.equal(changed.auto, undefined);
  assert.deepEqual(changed.suggestion.answer, { pick: [1] });
  assert.equal(changed.suggestion.close, true);
  // a month in the label is not a difference ...
  const sheets = mem();
  sheets.addRule({ kind: 'answer', scope: 'global', fp: 'fp-sheet2', matchText: normText('which sheet Use sheet OCT26 GRID Stop'), data: { kind: 'choose', labels: ['Use sheet OCT26 GRID'] } });
  const sheetQ = (labels) => ({ id: 'q', fp: 'fp-sheet2', kind: 'choose', problem: 'which sheet', question: 'Which sheet?', options: labels.map((label, i) => ({ n: i + 1, label })) });
  assert.deepEqual(sheets.matchQuestion('liberty', sheetQ(['Stop', 'Use sheet NOV26 GRID'])).auto.answer, { pick: [2] });
  assert.deepEqual(sheets.matchQuestion('liberty', sheetQ(['Stop', 'use  sheet Oct 2026 grid'])).auto.answer, { pick: [2] });
  // ... unless it leaves two options that cannot be told apart
  const both = sheets.matchQuestion('liberty', sheetQ(['Use sheet NOV26 GRID', 'Use sheet DEC26 GRID', 'Stop']));
  assert.equal(both.auto, undefined);
  assert.equal(both.suggestion.answer, null);
});

test('a text answer is remembered; a script-specific rule beats a global one of the same score; a disabled rule is ignored', () => {
  const m = mem();
  const q = { id: 't', fp: 'fp-folder', kind: 'text', problem: 'master folder not found in D:/x', question: 'Which folder?', options: [] };
  const g = m.addRule({ kind: 'answer', scope: 'global', fp: 'fp-folder', matchText: 'm', data: { kind: 'text', text: 'D:/global' } });
  const s = m.addRule({ kind: 'answer', scope: 'script:liberty', fp: 'fp-folder', matchText: 'm2', data: { kind: 'text', text: 'D:/liberty' } });
  assert.deepEqual(m.matchQuestion('liberty', q).auto.answer, { text: 'D:/liberty' });
  assert.deepEqual(m.matchQuestion('magma', q).auto.answer, { text: 'D:/global' });
  m.setEnabled(s, false);
  assert.deepEqual(m.matchQuestion('liberty', q).auto.answer, { text: 'D:/global' });
  m.setEnabled(g, false);
  assert.deepEqual(m.matchQuestion('liberty', q), {});
});

test('name mappings: exact, capitals / underscores, the month in the name, never an unknown target', () => {
  const m = mem();
  const masters = new Set(['GO-DIGIT_OCT26_HCV_MH_MUMBAI1', 'GO-DIGIT_NOV26_HCV_MH_MUMBAI1', 'GO-DIGIT_OCT26_HCV_MH_MUMBAI2']);
  const to = 'GO-DIGIT_OCT26_HCV_MH_MUMBAI1,GO-DIGIT_OCT26_HCV_MH_MUMBAI2';
  m.addRule({ kind: 'alias', scope: 'script:godigit', matchText: foldName('GO-DIGIT_OCT26_HCV_Mumbai'), data: Memory.aliasData('rto', 'GO-DIGIT_OCT26_HCV_Mumbai', to) });
  assert.equal(m.matchAlias('godigit', 'rto', 'GO-DIGIT_OCT26_HCV_Mumbai', masters).auto.to, to);
  assert.equal(m.matchAlias('godigit', 'rto', 'go-digit_oct26_hcv_mumbai_', masters).auto.to, to, 'capitals and stray underscores');
  assert.equal(m.matchAlias('godigit', 'makes', 'GO-DIGIT_OCT26_HCV_Mumbai', masters).auto, undefined, 'the other kind');
  const next = m.matchAlias('godigit', 'rto', 'GO-DIGIT_NOV26_HCV_Mumbai', new Set(['GO-DIGIT_NOV26_HCV_MH_MUMBAI1', 'GO-DIGIT_NOV26_HCV_MH_MUMBAI2']));
  assert.equal(next.auto.to, 'GO-DIGIT_NOV26_HCV_MH_MUMBAI1,GO-DIGIT_NOV26_HCV_MH_MUMBAI2', 'next month: the same mapping with the new month');
  const missing = m.matchAlias('godigit', 'rto', 'GO-DIGIT_NOV26_HCV_Mumbai', new Set(['GO-DIGIT_NOV26_HCV_MH_MUMBAI1']));
  assert.equal(missing.auto, undefined);
  assert.ok(missing.suggestion, 'a target that is not in the master is only suggested');
  assert.equal(m.matchAlias('godigit', 'rto', 'GO-DIGIT_OCT26_HCV_Mumbai', null).auto, undefined, 'without the master nothing is applied by itself');
  assert.deepEqual(m.aliasFile('godigit').rto, { 'GO-DIGIT_OCT26_HCV_Mumbai': to });
  assert.deepEqual(m.aliasFile('liberty').rto, {}, 'another script does not get it');
});

test('a new mapping for the same name replaces the old one', () => {
  const m = mem();
  m.addRule({ kind: 'alias', scope: 'global', matchText: 'a', data: Memory.aliasData('rto', 'A', 'X') });
  m.addRule({ kind: 'alias', scope: 'global', matchText: 'a', data: Memory.aliasData('rto', 'A', 'Y') });
  assert.equal(m.listRules({ kind: 'alias' }).length, 1);
  assert.equal(m.aliasFile('liberty').rto.A, 'Y');
});

test('a resolved note matches the same text apart from spaces, capitals, punctuation AND the numbers in it - but not another RTO code, and not a reworded rule', () => {
  const m = mem();
  m.addRule({ kind: 'standing_note', scope: 'global', matchText: 'n', data: { text: 'For Hazardous category PO will be 5% less.' } });
  assert.ok(m.isStandingNote('shriram_broker', 'for hazardous category  PO will be 5% less'));
  assert.ok(m.isStandingNote('shriram_broker', 'For Hazardous category PO will be 7% less.'), 'only the percentage changes: the same note');
  assert.equal(m.isStandingNote('shriram_broker', 'For Hazardous category PO will be 5% more.'), null, 'a reworded rule is a new note');
  assert.deepEqual(m.standingNotes('shriram_broker'), ['For Hazardous category PO will be 5% less.']);
  m.addRule({ kind: 'note_explanation', scope: 'global', matchText: 'x', data: { text: 'swift dzire model declined for RJ-14' }, explanation: 'ok' });
  m.addRule({ kind: 'note_explanation', scope: 'global', matchText: 'y', data: { text: 'EXL RTO codes 03, 06, 18, 24' }, explanation: 'ok' });
  assert.ok(m.resolvedFor('shriram_broker', 'swift dzire model declined for RJ-14'));
  assert.equal(m.resolvedFor('shriram_broker', 'swift dzire model declined for MH-12'), null, 'another RTO is another note');
  assert.equal(m.resolvedFor('shriram_broker', 'EXL RTO codes 10, 11, 12, 13'), null, 'another list of RTO codes is another note');
  assert.deepEqual(m.resolved('shriram_broker').map((x) => x.action).sort(), ['explain', 'explain', 'ignore']);
});

test('foldNote: numbers fold to #, RTO codes and lists of codes stay (parity with fold_note of the converters)', () => {
  assert.equal(foldKey('discount limit 50.0072% in the grid written as 50%'), foldKey('discount limit 49.7903% in the grid written as 49%'));
  assert.equal(foldKey('grid note: Few RTOs for 1000-1500CC segment are declined'), foldKey('grid note: Few RTOs for 1500-2000CC segment are declined'));
  assert.notEqual(foldKey('declined for RJ-14'), foldKey('declined for MH-12'));
  assert.notEqual(foldKey('EXL RTO codes 03, 06, 18'), foldKey('EXL RTO codes 10, 11, 12'));
  assert.equal(foldNote('PO will be 5% less'), 'PO will be #% less');
});

test('a note about a stop is offered only for a message that is almost the same, in the same scope', () => {
  const m = mem();
  m.addRule({ kind: 'failure_note', scope: 'script:godigit', matchText: normText('the sheet Rates 2026 was not found in Agency.xlsx'), data: { error: 'x' }, explanation: 'renamed this month' });
  const hit = m.matchFailure('godigit', 'the sheet Rates 2027 was not found in Other_Agency.xlsx');
  assert.equal(hit.rule.explanation, 'renamed this month');
  assert.ok(hit.score >= 0.8);
  assert.equal(m.matchFailure('godigit', 'the password does not open the workbook'), null);
  assert.equal(m.matchFailure('liberty', 'the sheet Rates 2026 was not found in Agency.xlsx'), null, 'another converter');
});

test('a rule file with the wrong shape is skipped, not stored', () => {
  const m = mem();
  const bad = [
    { kind: 'answer', data: { kind: 'choose', labels: 'not a list' } },
    { kind: 'answer', data: { kind: 'choose', labels: [] } },
    { kind: 'alias', data: { kind: 'zip', from: 'A', to: 'B' } },
    { kind: 'alias', data: { kind: 'rto', from: 'A', to: '' } },
    { kind: 'standing_note', data: { text: 42 } },
    { kind: 'failure_note', data: {} },
    { kind: 'answer' },
  ];
  assert.equal(m.import([...bad, { kind: 'answer', scope: 'global', data: { kind: 'choose', labels: ['Stop'] } }]), 1);
  assert.equal(m.listRules().length, 1);
});

test('use is counted and audited; export / import round-trips the rules', () => {
  const m = mem();
  const id = m.addRule({ kind: 'standing_note', scope: 'global', matchText: 'n', data: { text: 'a note long enough to matter here' }, explanation: 'ok' });
  m.db.prepare('INSERT INTO jobs (id, script, status, params, created_at, updated_at) VALUES (\'j1\', \'liberty\', \'done\', \'{}\', \'t\', \'t\')').run();
  m.recordUse(id, { jobId: 'j1', key: 'k', mode: 'note', score: 1 });
  m.recordUse(id, { jobId: 'j1', key: 'k', mode: 'note', score: 1 });
  assert.equal(m.getRule(id).uses, 2);
  assert.equal(m.db.prepare('SELECT COUNT(*) AS n FROM applications').get().n, 1, 'one audit line per job, rule and key');
  const dump = m.export();
  const other = mem();
  assert.equal(other.import(dump), 1);
  assert.equal(other.standingNotes('liberty')[0], 'a note long enough to matter here');
  assert.equal(other.import([{ kind: 'bogus', data: {} }, null]), 0);
});
