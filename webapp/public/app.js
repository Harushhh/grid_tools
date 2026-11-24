// Grid agent - the page. No framework, no build step. Everything shown is built with DOM calls (never innerHTML),
// so nothing typed by a person or read from a grid file can turn into markup.

const main = document.querySelector('#main');
const state = { authed: false, required: true, scripts: null, maxUploadMb: 150 };
let pollTimer = null;

// ------------------------------------------------------------------ small helpers
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (k === 'value') el.value = v;
    else if (['checked', 'disabled', 'selected', 'required', 'multiple', 'hidden'].includes(k)) el[k] = !!v;
    else el.setAttribute(k, v === true ? '' : String(v));
  }
  for (const kid of kids.flat(Infinity)) {
    if (kid === null || kid === undefined || kid === false) continue;
    el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return el;
}

const num = (n) => (n === null || n === undefined ? '' : Number(n).toLocaleString('en-US'));
const when = (iso) => {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleString(undefined, { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
};
const clock = (iso) => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', second: '2-digit' });
};
const ddmmyyyy = (iso) => (/^\d{4}-\d{2}-\d{2}$/.test(iso) ? iso.split('-').reverse().join('/') : '');
const pct = (x) => `${Math.round(x * 100)}%`;
const clip = (s, n) => (String(s).length > n ? `${String(s).slice(0, n - 1).trimEnd()}...` : String(s));
const sentence = (s) => (s ? `${s.charAt(0).toUpperCase()}${s.slice(1)}${/[.!?]$/.test(s) ? '' : '.'}` : '');
const KIND = { rto: 'RTO group', makes: 'make / model group' };
const STATUS = { queued: 'Waiting to start', running: 'Running', needs_input: 'Needs your answer', review: 'Ready - please review', done: 'Done', failed: 'Stopped' };
const BUSY = new Set(['queued', 'running']);

function toast(message, bad = false) {
  const box = document.querySelector('#toast');
  const t = h('div', { class: `t${bad ? ' bad' : ''}` }, message);
  box.append(t);
  setTimeout(() => t.remove(), bad ? 7000 : 3500);
}

async function api(path, { method = 'GET', body } = {}) {
  const opts = { method, headers: { 'X-Requested-With': 'gridapp' }, credentials: 'same-origin' };
  if (body !== undefined) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const res = await fetch(`/api${path}`, opts);
  let data = null;
  try {
    data = await res.json();
  } catch {
    data = null;
  }
  if (res.status === 401 && !path.startsWith('/login')) {
    state.authed = false;
    location.hash = '#/login';
    throw new Error('Please sign in.');
  }
  if (!res.ok) throw new Error((data && data.error) || `The request failed (${res.status}).`);
  return data;
}

function postForm(path, form, onProgress) {
  return new Promise((resolve, reject) => {
    const x = new XMLHttpRequest();
    x.open('POST', `/api${path}`);
    x.setRequestHeader('X-Requested-With', 'gridapp');
    x.upload.onprogress = (e) => e.lengthComputable && onProgress(e.loaded / e.total);
    x.onload = () => {
      let d = null;
      try {
        d = JSON.parse(x.responseText);
      } catch {
        d = null;
      }
      if (x.status >= 200 && x.status < 300) resolve(d);
      else reject(new Error((d && d.error) || `The upload failed (${x.status}).`));
    };
    x.onerror = () => reject(new Error('The upload was interrupted. Check the connection and try again.'));
    x.send(form);
  });
}

function stopPolling() {
  clearTimeout(pollTimer);
  pollTimer = null;
}

/** Replace what an element shows with these nodes (arrays are flattened, null / false are left out). */
function fill(el, ...nodes) {
  el.replaceChildren(...nodes.flat(Infinity).filter((n) => n !== null && n !== undefined && n !== false));
}

async function busy(button, fn) {
  const label = button.textContent;
  button.disabled = true;
  button.textContent = 'Working...';
  try {
    return await fn();
  } catch (e) {
    toast(e.message, true);
    return null;
  } finally {
    if (button.isConnected) {
      button.disabled = false;
      button.textContent = label;
    }
  }
}

const chip = (text, cls = '') => h('span', { class: `chip ${cls}` }, text);

function scopeLabel(scope) {
  if (scope === 'global') return 'Every converter';
  if (scope.startsWith('script:')) {
    const s = (state.scripts || []).find((x) => x.id === scope.slice(7));
    return s ? s.label : `Converter ${scope.slice(7)}`;
  }
  if (scope.startsWith('insurer:')) return `All ${scope.slice(8)} converters`;
  return scope;
}

function scopeSelect(script) {
  return h('select', { name: 'scope', 'aria-label': 'Remember this for' },
    h('option', { value: 'script' }, `This converter only (${script.label})`),
    h('option', { value: 'insurer' }, `All ${script.insurer} converters`),
    h('option', { value: 'global' }, 'Every converter'));
}

/** "Remember this" controls shared by the cards. */
function rememberBox(script, { what, explainLabel = 'Why (optional) - shown next to the rule' }) {
  const remember = h('input', { type: 'checkbox', name: 'remember', checked: true });
  const scope = scopeSelect(script);
  const explanation = h('textarea', { name: 'explanation', placeholder: 'For example: the grid has no Electric column in July' });
  const body = h('div', { class: 'grid2' },
    h('div', null, h('label', null, 'Remember it for'), scope),
    h('div', null, h('label', null, explainLabel), explanation));
  remember.addEventListener('change', () => body.classList.toggle('hidden', !remember.checked));
  return {
    node: h('div', null, h('label', { class: 'inline' }, remember, `Remember ${what} for next time`), body),
    value: () => ({ remember: remember.checked, scope: scope.value, explanation: explanation.value.trim() }),
  };
}

// ------------------------------------------------------------------ the cards
function ruleLine(c) {
  const r = c.rule;
  if (!r) return null;
  return h('p', { class: 'small muted' }, `Rule #${r.id} - ${scopeLabel(r.scope)} - used ${r.uses}x${r.enabled ? '' : ' - switched off'}`, r.explanation ? ` - "${r.explanation}"` : '');
}

function describeCard(c) {
  const p = c.payload;
  const a = c.answer || {};
  if (c.type === 'question') return [h('strong', null, clip(p.short || p.problem, 130)), ' -> ', clip(a.labels ? a.labels.join('; ') : a.text || '', 160)];
  if (c.type === 'alias') return [h('strong', null, `${KIND[p.kind]} ${p.name}`), a.skip ? ' left as it is' : ` -> ${a.to}`];
  if (c.type === 'note') {
    const verdict = { ignore: 'needs no action - not flagged again', explain: `explained - not flagged again: ${a.explanation || ''}${a.similarTo ? ' (from a note that differed only in numbers)' : ''}`, leave: 'left flagged' }[a.action] || '';
    return [h('strong', null, p.text.length > 140 ? `${p.text.slice(0, 140)}...` : p.text), ` - ${verdict}`];
  }
  if (c.type === 'failure') return [h('strong', null, 'The converter stopped'), a.action === 'save' ? ` - you noted: ${clip(a.note || '', 160)}` : ' - dismissed'];
  return [h('strong', null, 'Password'), ' given (kept in memory for this job only)'];
}

function doneCard(job, c, refresh) {
  const fromMemory = c.status === 'auto';
  const change = h('button', { type: 'button', class: 'small' }, fromMemory ? 'Ask me instead' : 'Change');
  change.addEventListener('click', () => busy(change, async () => {
    const r = await api(`/jobs/${job.id}/cards/${c.id}/reopen`, { method: 'POST', body: {} });
    refresh(r.job);
  }));
  const buttons = [change];
  if (fromMemory && c.rule && c.rule.enabled) {
    const off = h('button', { type: 'button', class: 'small' }, 'Stop using this rule');
    off.addEventListener('click', () => busy(off, async () => {
      await api(`/rules/${c.rule.id}`, { method: 'PATCH', body: { enabled: false } });
      toast(`Rule #${c.rule.id} is switched off. It will not be used again.`);
      refresh((await api(`/jobs/${job.id}`)));
    }));
    buttons.push(off);
  }
  return h('div', { class: `panel ${fromMemory ? 'auto' : 'done'}` },
    h('div', { class: 'row spread' },
      h('div', { class: 'grow' }, fromMemory ? chip('From memory', 'mem') : chip(c.status === 'answered' ? 'You answered' : c.status, 'ok'), ' ', describeCard(c)),
      h('div', { class: 'row' }, buttons)),
    ruleLine(c));
}

function questionCard(job, c, refresh) {
  const q = c.payload;
  const form = h('form', { class: 'panel need' });
  form.append(h('h3', null, clip(q.question || q.problem, 220)));
  const detail = q.problem && q.problem !== q.question ? q.problem : (q.question || '').length > 220 ? q.question : '';
  if (detail) form.append(h('details', null, h('summary', null, 'Details from the converter'), h('p', { class: 'small muted' }, detail)));
  if (q.invalid) form.append(h('div', { class: 'banner warn' }, 'The converter did not accept the earlier answer. Please answer again.'));
  const sg = c.suggestion;
  let apply = null;
  if (sg) {
    const box = h('div', { class: 'suggest' });
    if (sg.source === 'advisor') {
      box.append(chip('Advisor', 'mem'), `Suggestion (not checked against your past answers): ${sg.answer ? (sg.answer.labels ? sg.answer.labels.join('; ') : sg.answer.text) : ''}`, sg.reason ? ` - ${sg.reason}` : '');
    } else {
      box.append(chip(`${pct(sg.score)} similar`, 'mem'), `Last time (rule #${sg.ruleId}) you answered: "${sg.was}"`, sg.explanation ? ` - ${sg.explanation}` : '',
        sg.close ? ' - the option below is the closest one here, but it is not exactly the same: please check it before using it' : '');
    }
    if (sg.answer) {
      apply = () => {
        if (sg.answer.pick) form.querySelectorAll('input[name=pick]').forEach((i) => { i.checked = sg.answer.pick.includes(Number(i.value)); i.dispatchEvent(new Event('change')); });
        else form.querySelector('input[name=text]').value = sg.answer.text;
      };
      box.append(' ', h('button', { type: 'button', class: 'small', onclick: apply }, 'Use this answer'));
    } else {
      box.append(' ', h('span', { class: 'muted' }, '(that option is not in this question)'));
    }
    form.append(box);
  }
  let text = null;
  if (q.kind === 'text') {
    text = h('input', { type: 'text', name: 'text', autocomplete: 'off', 'aria-label': 'Your answer' });
    form.append(text);
  } else {
    for (const o of q.options) {
      const input = h('input', { type: q.multi ? 'checkbox' : 'radio', name: 'pick', value: o.n });
      const row = h('label', { class: 'opt' }, input, h('span', null, o.label));
      input.addEventListener('change', () => form.querySelectorAll('.opt').forEach((l) => l.classList.toggle('pick', l.querySelector('input').checked)));
      form.append(row);
    }
    if (q.multi) form.append(h('p', { class: 'small muted' }, 'You can choose more than one.'));
  }
  const mem = rememberBox(job.script, { what: 'this answer' });
  form.append(mem.node);
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Answer');
  form.append(h('div', { class: 'row', style: null }, submit));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const body = { ...mem.value() };
    if (q.kind === 'text') body.text = text.value;
    else body.pick = [...form.querySelectorAll('input[name=pick]:checked')].map((i) => Number(i.value));
    busy(submit, async () => {
      const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body });
      refresh(r.job);
    });
  });
  return form;
}

// ------------------------------------------------------------------ several names at once
let bulkAlias = null;      // set by paintJob while the name cards are built (like `bulk` for the notes)

/** The "name not in the master" cards of the open list: one decision for the ticked ones - leave them all as they are, or map them all to one master group. */
function aliasesTogether(job, aliases, refresh) {
  const picked = new Set();
  const boxes = new Map();                                    // card id -> its tick box
  const chosen = [];
  const byKind = new Map();
  for (const a of aliases) byKind.set(a.payload.kind, (byKind.get(a.payload.kind) || 0) + 1);
  const kindsPicked = () => new Set([...picked].map((id) => aliases.find((a) => a.id === id).payload.kind));

  const action = h('select', { 'aria-label': 'What to do with the selected names' },
    h('option', { value: 'map' }, 'Map them all to one master group'),
    h('option', { value: 'skip' }, 'Leave them all as they are'));
  const filter = h('input', { type: 'search', placeholder: 'Select the names containing ... (for example HCV)', autocomplete: 'off', 'aria-label': 'Select the names containing' });
  const filterGo = h('button', { type: 'button' }, 'Select matching');
  const chosenBox = h('div', { class: 'pills' });
  const results = h('div', { class: 'pills' });
  const hint = h('p', { class: 'small muted' });
  const search = h('input', { type: 'search', placeholder: 'Search the master (RTO or make / model groups)', autocomplete: 'off', 'aria-label': 'Search the master' });
  const mem = rememberBox(job.script, { what: 'these mappings' });
  const mapOnly = h('div', null, h('label', null, 'Which master group do they stand for?'), chosenBox, search, results, hint, mem.node);

  const all = h('input', { type: 'checkbox', 'aria-label': 'Select all names' });
  const count = h('span', { class: 'muted' });
  const apply = h('button', { type: 'button', class: 'primary' }, 'Apply to selected');

  const paintChosen = () => {
    chosenBox.replaceChildren(...chosen.map((n, i) => h('button', { type: 'button', class: 'small', title: 'Remove', onclick: () => { chosen.splice(i, 1); sync(); } }, `${n}  x`)));
    chosenBox.classList.toggle('hidden', !chosen.length);
  };
  const sync = () => {
    const n = picked.size;
    const kinds = kindsPicked();
    const map = action.value === 'map';
    count.textContent = n ? `${n} of ${aliases.length} selected` : 'None selected';
    all.checked = n === aliases.length;
    all.indeterminate = n > 0 && n < aliases.length;
    mapOnly.classList.toggle('hidden', !map);
    hint.textContent = map && kinds.size > 1 ? 'The selection holds RTO groups and make / model groups. A group can only stand for names of its own kind: select one kind.' : '';
    paintChosen();
    apply.disabled = !n || (map && (!chosen.length || kinds.size > 1));
    apply.textContent = !n ? 'Apply to selected' : map ? `Map ${n} selected` : `Leave ${n} selected as they are`;
    for (const [id, box] of boxes) box.checked = picked.has(id);
  };
  const set = (ids, on) => {
    for (const id of ids) {
      if (on) picked.add(id);
      else picked.delete(id);
    }
    sync();
  };
  all.addEventListener('change', () => set(aliases.map((a) => a.id), all.checked));
  action.addEventListener('change', sync);
  filterGo.addEventListener('click', () => {
    const q = filter.value.trim().toLowerCase();
    if (!q) return toast('Type part of a name first.', true);
    const hit = aliases.filter((a) => a.payload.name.toLowerCase().includes(q));
    if (!hit.length) return toast('No name contains that.', true);
    return set(hit.map((a) => a.id), true);
  });
  let timer = null;
  search.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      if (!search.value.trim()) return results.replaceChildren();
      const kinds = kindsPicked();
      const kind = kinds.size === 1 ? [...kinds][0] : kinds.size === 0 ? aliases[0].payload.kind : null;
      if (!kind) return results.replaceChildren(h('span', { class: 'muted' }, 'Select names of one kind first.'));
      try {
        const r = await api(`/masters/search?kind=${kind}&q=${encodeURIComponent(search.value)}`);
        if (!r.loaded) results.replaceChildren(h('span', { class: 'muted' }, 'The master lists could not be read on the server.'));
        else results.replaceChildren(...r.names.map((n) => h('button', { type: 'button', onclick: () => { if (!chosen.includes(n)) chosen.push(n); sync(); } }, n)), r.names.length ? '' : h('span', { class: 'muted' }, 'No such group in the master.'));
      } catch (e) {
        toast(e.message, true);
      }
      return null;
    }, 220);
  });
  apply.addEventListener('click', () => busy(apply, async () => {
    const cardIds = [...picked];
    const body = action.value === 'skip' ? { cardIds, skip: true } : { cardIds, to: chosen.join(','), ...mem.value() };
    const r = await api(`/jobs/${job.id}/aliases/answer`, { method: 'POST', body });
    toast(`${r.answered} name${r.answered === 1 ? '' : 's'} ${body.skip ? 'left as they are' : 'mapped'}.`);
    refresh(r.job);
  }));

  const setup = h('div', { class: 'panel' },
    h('h3', null, 'Several names at once'),
    h('p', { class: 'muted' }, `${aliases.length} names are not in the master. Tick the ones that should get the same answer (below, or by part of the name), choose what to do with all of them, then press the button at the top of the list.`),
    h('div', { class: 'row' }, h('div', { class: 'grow' }, filter), filterGo),
    h('label', null, 'What to do with the selected names'), action,
    mapOnly);
  const bar = h('div', { class: 'panel bulk' }, h('label', { class: 'inline' }, all, 'Select all'), count, apply);

  /** The tick box (and "select all of this kind") at the top of one name card. */
  const picker = (c) => {
    const box = h('input', { type: 'checkbox', 'aria-label': 'Select this name' });
    boxes.set(c.id, box);
    box.addEventListener('change', () => set([c.id], box.checked));
    const same = byKind.get(c.payload.kind);
    return h('div', { class: 'pick-row' },
      h('label', { class: 'inline' }, box, 'Select'),
      same > 1 ? h('button', { type: 'button', class: 'small', onclick: () => set(aliases.filter((a) => a.payload.kind === c.payload.kind).map((a) => a.id), true) }, `Select all ${same} ${KIND[c.payload.kind]}s`) : null);
  };
  sync();
  return { setup, bar, picker };
}

function aliasCard(job, c, refresh) {
  const p = c.payload;
  const chosen = [];
  const form = h('form', { class: 'panel need' });
  if (bulkAlias) form.append(bulkAlias.picker(c));
  form.append(h('h3', null, `The ${KIND[p.kind]} "${p.name}" is not in the master`), h('p', { class: 'muted' }, `${num(p.rows)} rows of the upload use it.${p.hint ? ` ${sentence(p.hint)}` : ''}`));
  if (p.stale) form.append(h('div', { class: 'banner warn' }, 'This name was mapped, but it is still not in the master. Please look at it again.'));
  const chosenBox = h('div', { class: 'pills' });
  const results = h('div', { class: 'pills' });
  const submit = h('button', { type: 'submit', class: 'primary', disabled: true }, 'Map it');
  const paint = () => {
    chosenBox.replaceChildren(...chosen.map((n, i) => h('button', { type: 'button', class: 'small', title: 'Remove', onclick: () => { chosen.splice(i, 1); paint(); } }, `${n}  x`)));
    chosenBox.classList.toggle('hidden', !chosen.length);
    submit.disabled = !chosen.length;
  };
  const add = (name, replace = false) => {
    if (replace) chosen.length = 0;
    if (!chosen.includes(name)) chosen.push(name);
    paint();
  };
  const sg = c.suggestion;
  if (sg) {
    const box = h('div', { class: 'suggest' }, chip(`${pct(sg.score)} similar`, 'mem'), `Last time (rule #${sg.ruleId}) this name stood for: ${sg.to} `);
    box.append(h('button', { type: 'button', class: 'small', onclick: () => { chosen.length = 0; sg.to.split(',').forEach((n) => add(n.trim())); } }, 'Use it'));
    form.append(box);
  }
  if (p.closest && p.closest.length) {
    form.append(h('label', null, 'Closest names in the master'), h('div', { class: 'pills' }, p.closest.map((n) => h('button', { type: 'button', onclick: () => add(n, true) }, n))));
  }
  const search = h('input', { type: 'search', placeholder: 'Search the master (RTO or make / model groups)', autocomplete: 'off', 'aria-label': 'Search the master' });
  let timer = null;
  search.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      if (!search.value.trim()) return results.replaceChildren();
      try {
        const r = await api(`/masters/search?kind=${p.kind}&q=${encodeURIComponent(search.value)}`);
        if (!r.loaded) results.replaceChildren(h('span', { class: 'muted' }, 'The master lists could not be read on the server.'));
        else results.replaceChildren(...r.names.map((n) => h('button', { type: 'button', onclick: () => add(n) }, n)), r.names.length ? '' : h('span', { class: 'muted' }, 'No such group in the master.'));
      } catch (e) {
        toast(e.message, true);
      }
    }, 220);
  });
  form.append(h('label', null, 'Which master group does it stand for?'), chosenBox, search, results);
  paint();
  const mem = rememberBox(job.script, { what: 'this mapping' });
  form.append(mem.node);
  const skip = h('button', { type: 'button' }, 'Leave it as it is');
  skip.addEventListener('click', () => busy(skip, async () => {
    const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body: { skip: true } });
    refresh(r.job);
  }));
  form.append(h('div', { class: 'row' }, submit, skip));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    busy(submit, async () => {
      const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body: { to: chosen.join(','), ...mem.value() } });
      refresh(r.job);
    });
  });
  return form;
}

// ------------------------------------------------------------------ several notes at once
let bulk = null;      // set by paintJob while the note cards are built: the tick boxes and the bar of the open notes
/** The issue of a note card; a server that does not send it yet (not restarted after an update) gives every note its own issue, so nothing is grouped wrongly. */
const issueOf = (c) => c.issue ?? `#${c.id}`;

const ACTIONS = [
  ['explain', 'Explained - remember what they mean and do not flag them again'],
  ['ignore', 'No action needed - do not flag them again'],
  ['leave', 'Leave them flagged for now'],
];

/** Notes that say the same apart from numbers share `issue`: they can be ticked together and cleared with one answer. */
function notesTogether(job, notes, refresh) {
  const picked = new Set();
  const boxes = new Map();                                    // card id -> its tick box
  const size = new Map();
  for (const n of notes) size.set(issueOf(n), (size.get(issueOf(n)) || 0) + 1);

  const all = h('input', { type: 'checkbox', 'aria-label': 'Select all notes' });
  const count = h('span', { class: 'muted' });
  const action = h('select', { 'aria-label': 'What to do with the selected notes' }, ACTIONS.map(([v, label]) => h('option', { value: v }, label)));
  const why = h('input', { type: 'text', placeholder: 'What it means, or what you do about it', autocomplete: 'off', 'aria-label': 'Explanation' });
  const scope = scopeSelect(job.script);
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Clear selected');

  const sync = () => {
    const n = picked.size;
    count.textContent = n ? `${n} of ${notes.length} selected` : 'None selected';
    all.checked = n === notes.length;
    all.indeterminate = n > 0 && n < notes.length;
    why.classList.toggle('hidden', action.value !== 'explain');
    scope.classList.toggle('hidden', action.value === 'leave');
    submit.disabled = !n;
    submit.textContent = n ? `Clear ${n} selected` : 'Clear selected';
    for (const [id, box] of boxes) box.checked = picked.has(id);
  };
  const set = (ids, on) => {
    for (const id of ids) {
      if (on) picked.add(id);
      else picked.delete(id);
    }
    sync();
  };
  all.addEventListener('change', () => set(notes.map((n) => n.id), all.checked));
  action.addEventListener('change', sync);

  const bar = h('form', { class: 'panel bulk' },
    h('label', { class: 'inline' }, all, 'Select all'), count, action, why, scope, submit);
  bar.addEventListener('submit', (e) => {
    e.preventDefault();
    busy(submit, async () => {
      const r = await api(`/jobs/${job.id}/notes/answer`, { method: 'POST', body: { cardIds: [...picked], action: action.value, explanation: why.value, scope: scope.value } });
      toast(`${r.answered} note${r.answered === 1 ? '' : 's'} saved.${action.value === 'leave' ? '' : ' The next run does not flag them (also when only a number in them changes).'}`);
      refresh(r.job);
    });
  });

  /** The tick box (and "select the others like this") that goes at the top of one note card. */
  const picker = (c) => {
    const box = h('input', { type: 'checkbox', 'aria-label': 'Select this note' });
    boxes.set(c.id, box);
    box.addEventListener('change', () => set([c.id], box.checked));
    const like = size.get(issueOf(c));
    return h('div', { class: 'pick-row' },
      h('label', { class: 'inline' }, box, 'Select'),
      like > 1 ? chip(`${like} like this (only the numbers differ)`, 'mem') : null,
      like > 1 ? h('button', { type: 'button', class: 'small', onclick: () => set(notes.filter((n) => issueOf(n) === issueOf(c)).map((n) => n.id), true) }, `Select all ${like} like this`) : null);
  };
  sync();
  return { bar, picker };
}

function noteCard(job, c, refresh) {
  const p = c.payload;
  const form = h('form', { class: 'panel need' });
  if (bulk) form.append(bulk.picker(c));
  form.append(
    h('h3', null, p.silenceable ? 'A note in the grid that the converter could not apply' : 'The converter flagged this for a manual check'),
    h('div', { class: 'note-text' }, p.text),
    h('p', { class: 'small muted' }, p.rows ? `${num(p.rows)} rows of the upload carry "Manual Review Required" because of it.` : ''));
  const why = h('textarea', { name: 'explanation', placeholder: 'What it means, or what you do about it. Shown next time this note comes up.' });
  const choice = (value, label, extra, disabled = false) => {
    const input = h('input', { type: 'radio', name: 'action', value, disabled });
    input.addEventListener('change', () => form.querySelectorAll('.opt').forEach((l) => l.classList.toggle('pick', l.querySelector('input').checked)));
    return h('label', { class: `opt${disabled ? ' off' : ''}` }, input, h('span', null, label, extra));
  };
  form.append(
    choice('explain', 'Explained - remember what it means and do not flag it again'),
    choice('ignore', 'No action needed - do not flag this note again'),
    choice('leave', 'Leave it flagged for now'),
    why);
  const scope = scopeSelect(job.script);
  form.append(h('label', null, 'Remember it for'), scope);
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Save');
  form.append(h('div', { class: 'row' }, submit));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const action = (form.querySelector('input[name=action]:checked') || {}).value;
    if (!action) return toast('Please choose one of the three.', true);
    return busy(submit, async () => {
      const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body: { action, explanation: why.value, scope: scope.value } });
      if (action !== 'leave') toast('Saved. The next run does not flag this note (also when only a number in it changes).');
      refresh(r.job);
    });
  });
  return form;
}

function passwordCard(job, c, refresh) {
  const input = h('input', { type: 'password', autocomplete: 'off', 'aria-label': c.payload.label });
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Continue');
  const form = h('form', { class: 'panel need' },
    h('h3', null, c.payload.reason),
    h('p', { class: 'muted' }, c.payload.label),
    input, h('div', { class: 'row' }, submit));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    busy(submit, async () => {
      const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body: { value: input.value } });
      input.value = '';
      refresh(r.job);
    });
  });
  return form;
}

function failureCard(job, c, refresh) {
  const p = c.payload;
  const sg = c.suggestion;
  const note = h('textarea', { name: 'note', placeholder: 'For example: this month the sheet is called "Rates" instead of "RTO Rates".' });
  const scope = scopeSelect(job.script);
  const save = h('button', { type: 'submit', class: 'primary' }, 'Save my note');
  const dismiss = h('button', { type: 'button' }, 'Dismiss');
  const send = (body, button) => busy(button, async () => {
    const r = await api(`/jobs/${job.id}/cards/${c.id}/answer`, { method: 'POST', body });
    refresh(r.job);
  });
  dismiss.addEventListener('click', () => send({ action: 'dismiss' }, dismiss));
  const form = h('form', { class: 'panel need' },
    h('h3', null, 'The converter stopped on this file'),
    h('div', { class: 'note-text' }, p.error),
    sg ? h('div', { class: 'suggest' }, chip(`${pct(sg.score)} similar`, 'mem'), `It stopped like this before (rule #${sg.ruleId}). You noted: "${sg.note}"`) : null,
    h('p', { class: 'muted' }, 'The converter cannot get past this by itself. Tell it what is different about this file: the note is kept in the memory and shown when it stops this way again, and the Memory page can export it for whoever updates the converter.'),
    h('label', null, 'What is different about this file?'), note,
    h('label', null, 'Remember it for'), scope,
    h('div', { class: 'row' }, save, dismiss));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    send({ action: 'save', note: note.value, scope: scope.value }, save);
  });
  return form;
}

function cardView(job, c, refresh) {
  if (c.status !== 'open') return doneCard(job, c, refresh);
  return { question: questionCard, alias: aliasCard, note: noteCard, password: passwordCard, failure: failureCard }[c.type](job, c, refresh);
}

// ------------------------------------------------------------------ the page
// ONE page with three columns (a feed with a rail on each side): the JOBS on the left, "New run" and the selected job in the middle, the MEMORY on the right.
// Nothing is a separate screen any more, so nothing has to be left to look at something else (a half-filled New run form stays where it is while a job is read).
// The address only says what is selected, so every old link still works: #/ = nothing selected and New run open, #/jobs/<id> = that job,
// #/jobs and #/memory = the same page (on a narrow screen they pick the Jobs / Memory tab). #/login is the only other screen.
const $ = (sel) => document.querySelector(sel);
const ui = { login: $('#login'), shell: $('#shell'), nav: $('#nav'), jobsCol: $('#col-jobs'), memCol: $('#col-memory'), tabBadge: $('#tab-badge') };
const view = { built: false, active: false, loaded: false, jobId: null, token: 0, jobs: [], jobsLoaded: false, jobsSig: '', jobsSeq: 0, jobsTimer: null, memSeq: 0, composerOpen: true };
const narrowMq = window.matchMedia('(max-width: 1199px)');
const narrow = () => narrowMq.matches;

function errorBox(message, retry) {
  return h('div', { class: 'banner bad' }, h('div', null, h('strong', null, 'This could not be shown'), message,
    retry ? h('div', null, h('button', { type: 'button', class: 'small', onclick: () => retry() }, 'Try again')) : null));
}

function showView(name) {
  ui.login.hidden = name !== 'login';
  ui.shell.hidden = name !== 'shell';
  ui.nav.hidden = name !== 'shell';
}

/** Which column a narrow screen shows (a wide screen shows all three whatever this says). */
function setPane(pane) {
  document.body.dataset.pane = pane;
  ui.nav.querySelectorAll('a').forEach((a) => {
    const on = a.dataset.nav === pane;
    a.classList.toggle('on', on);
    if (on) a.setAttribute('aria-current', 'page');
    else a.removeAttribute('aria-current');
  });
}

/** Go to an address. When it is the one we are on no hashchange event comes, so apply it ourselves. */
function go(hash) {
  const here = location.hash || '#/';
  if (here === hash || (hash === '#/' && /^#\/?$/.test(here))) route();
  else location.hash = hash;
}

function renderLogin() {
  const input = h('input', { type: 'password', autocomplete: 'current-password', 'aria-label': 'Password' });
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Sign in');
  const form = h('form', { class: 'panel' }, h('h1', null, 'Sign in'), h('p', { class: 'sub' }, 'The password of this app.'), input, h('div', { class: 'row' }, submit));
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    busy(submit, async () => {
      await api('/login', { method: 'POST', body: { password: input.value } });
      state.authed = true;
      location.hash = '#/';
    });
  });
  ui.login.replaceChildren(h('div', { class: 'login-card' }, form));
  input.focus();
}

async function loadScripts() {
  if (!state.scripts) {
    const r = await api('/scripts');
    state.scripts = r.scripts;
    state.maxUploadMb = r.maxUploadMb;
  }
  return state.scripts;
}

// ------------------------------------------------------------------ the middle: New run, then the job that is selected
function setComposer(open) {
  view.composerOpen = open;
  const c = ui.composer;
  if (!c) return;
  c.body.hidden = !open;
  c.toggle.setAttribute('aria-expanded', String(open));
}

function focusComposer() {
  if (ui.composer && ui.composer.select) ui.composer.select.focus({ preventScroll: true });
}

function buildWork() {
  const toggle = h('button', { type: 'button', class: 'composer-toggle', 'aria-expanded': 'true', 'aria-controls': 'composer-body' },
    h('span', { class: 'composer-title' }, 'New run'),
    h('span', { class: 'composer-hint muted' }, 'Upload an insurer\'s grid'),
    h('span', { class: 'chev', 'aria-hidden': 'true' }));
  const body = h('div', { class: 'composer-body', id: 'composer-body' }, h('p', { class: 'muted' }, 'Loading...'));
  toggle.addEventListener('click', () => setComposer(!view.composerOpen));
  ui.composer = { toggle, body, form: null, select: null, reset: null };
  ui.detail = h('div', { class: 'detail', 'aria-live': 'polite' });
  main.append(h('section', { class: 'panel composer', 'aria-label': 'New run' }, toggle, body), ui.detail);
  setComposer(view.composerOpen);
}

/** The New run form. Built once and kept, so what was typed or chosen in it survives a look at a job; cleared after a run is sent and when signing out. */
function paintComposer(scripts) {
  const select = h('select', { name: 'script', 'aria-label': 'Converter' }, scripts.map((s) => h('option', { value: s.id }, `${s.insurer} - ${s.label}`)));
  const fields = h('div');
  const from = h('input', { type: 'date', required: true, 'aria-label': 'from_date' });
  const to = h('input', { type: 'date', required: true, 'aria-label': 'to_date' });
  const bar = h('progress', { max: 1, value: 0, class: 'hidden' });
  const submit = h('button', { type: 'submit', class: 'primary' }, 'Run');
  const paintFields = () => {
    const s = scripts.find((x) => x.id === select.value);
    fields.replaceChildren(
      ...s.inputs.map((i) => h('div', null, h('label', null, i.label, i.required ? '' : ' '), h('input', { type: 'file', name: `file:${i.role}`, accept: i.accept.join(','), required: i.required }))),
      ...s.options.map((o) => (o.type === 'flag'
        ? h('label', { class: 'inline' }, h('input', { type: 'checkbox', name: `opt:${o.name}` }), o.label)
        : h('div', null, h('label', null, o.label), h('input', { type: 'text', name: `opt:${o.name}`, autocomplete: 'off' })))),
      ...s.secrets.map((x) => h('div', null, h('label', null, x.label), h('input', { type: 'password', name: `secret:${x.name}`, autocomplete: 'off' }))));
  };
  select.addEventListener('change', paintFields);
  paintFields();
  const form = h('form', null,
    h('label', null, 'Converter'), select,
    h('div', { class: 'grid2' }, h('div', null, h('label', null, 'from_date'), from), h('div', null, h('label', null, 'to_date'), to)),
    fields, bar, h('div', { class: 'row' }, submit),
    h('p', { class: 'small muted' }, `Files up to ${state.maxUploadMb} MB. Nothing is sent anywhere else: the converter runs on this server.`));
  const reset = () => {
    form.reset();
    bar.value = 0;
    bar.classList.add('hidden');
    paintFields();
  };
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const data = new FormData();
    data.append('script', select.value);
    data.append('fromDate', ddmmyyyy(from.value));
    data.append('toDate', ddmmyyyy(to.value));
    for (const el of form.querySelectorAll('[name]')) {
      const [kind, name] = el.name.split(':');
      if (kind === 'file' && el.files[0]) data.append(name, el.files[0]);
      else if (kind === 'opt') data.append(`opt_${name}`, el.type === 'checkbox' ? String(el.checked) : el.value);
      else if (kind === 'secret' && el.value) data.append(`secret_${name}`, el.value);
    }
    bar.classList.remove('hidden');
    busy(submit, async () => {
      try {
        const r = await postForm('/jobs', data, (f) => { bar.value = f; });
        reset();
        loadJobs();
        go(`#/jobs/${r.id}`);
      } finally {
        bar.classList.add('hidden');
      }
    });
  });
  Object.assign(ui.composer, { form, select, reset });
  ui.composer.body.replaceChildren(h('p', { class: 'sub' }, 'Upload the insurer\'s grid. The converter runs, and you are asked only about what it cannot decide by itself or from what you explained before.'), form);
}

async function loadComposer() {
  try {
    paintComposer(await loadScripts());
  } catch (e) {
    if (view.active) ui.composer.body.replaceChildren(errorBox(e.message, loadComposer));
  }
}

/** The middle, when no job is selected: a line that says what to do next. */
function paintEmpty() {
  if (view.jobId) return;
  const text = !view.jobsLoaded ? '' : view.jobs.length ? 'Pick a job from the list to see it here.' : 'No jobs yet. Open "New run" above, choose the converter and the files, and press Run.';
  fill(ui.detail, text ? h('div', { class: 'panel empty' }, text) : null);
}

function timeline(events) {
  const names = { run: 'Run', memory: 'Memory', ask: 'You', info: 'Info', error: 'Problem' };
  return h('ul', { class: 'timeline' }, events.map((e) => h('li', null, h('span', { class: 'muted' }, clock(e.at)), h('span', { class: `k ${e.kind}` }, names[e.kind] || e.kind), h('span', null, e.text))));
}

function banner(job) {
  const open = job.cards.filter((c) => c.status === 'open');
  if (BUSY.has(job.status)) return h('div', { class: 'banner info' }, h('div', { class: 'spinner' }), h('div', null, h('strong', null, job.status === 'queued' ? 'Waiting for a free slot' : `The converter is running (run ${job.run_count})`), 'This page updates by itself.'));
  if (job.status === 'needs_input') return h('div', { class: 'banner warn' }, h('div', null, h('strong', null, 'I need an answer from you to go on'), 'Answer below. The job continues by itself, and what you answer is remembered.'));
  if (job.status === 'review') return h('div', { class: 'banner warn' }, h('div', null, h('strong', null, `The upload is ready - ${open.length} thing${open.length === 1 ? '' : 's'} to look at`), 'You can download it now. Answer the cards below and press "Apply and run again" to write it again with your answers.'));
  if (job.status === 'done') return h('div', { class: 'banner ok' }, h('div', null, h('strong', null, 'Done'), 'Nothing needs your attention.'));
  return h('div', { class: 'banner bad' }, h('div', null, h('strong', null, 'The job stopped'), job.error || 'Something went wrong.'));
}

function resultPanel(job) {
  const o = job.outputs;
  if (!o) return null;
  const dl = (which, label, name) => h('a', { class: 'btn primary', href: `/api/jobs/${job.id}/download/${which}`, download: name }, label);
  const products = (job.report && job.report.products) || [];
  return h('div', { class: 'panel' },
    h('div', { class: 'row spread' },
      h('div', null, h('h3', null, 'Result'), h('p', { class: 'muted' }, `${num(o.rows)} rows, ${num(o.flagged)} flagged for manual review${job.report ? ` - ${job.report.from_date} to ${job.report.to_date}` : ''}`)),
      h('div', { class: 'row' }, dl('upload', 'Download the upload (CSV)', o.upload), o.review ? h('a', { class: 'btn', href: `/api/jobs/${job.id}/download/review`, download: o.review }, 'Review workbook') : null)),
    job.report && job.report.warnings.length ? h('div', { class: 'banner warn' }, h('div', null, h('strong', null, 'Warnings'), job.report.warnings.map((w) => h('div', null, w)))) : null,
    products.length ? h('details', null, h('summary', null, `What is in it (${products.length} product lines)`),
      h('table', null, h('thead', null, h('tr', null, h('th', null, 'Product'), h('th', null, 'Sub product'), h('th', { class: 'num' }, 'Rows'))),
        h('tbody', null, products.map((p) => h('tr', null, h('td', null, p.product), h('td', null, p.sub_product), h('td', { class: 'num' }, num(p.rows))))))) : null,
    job.report && job.report.decisions.length ? h('details', null, h('summary', null, 'What was decided in this run'), h('ul', null, job.report.decisions.map((d) => h('li', null, d)))) : null);
}

/** Select a job: it is shown in the middle, under New run. `token` says which selection a late answer belongs to, so a slow reply can never paint over the job chosen after it. */
async function selectJob(id) {
  const token = ++view.token;
  view.jobId = id;
  stopPolling();
  markSelected();
  window.scrollTo(0, 0);
  fill(ui.detail, h('div', { class: 'panel muted' }, 'Loading...'));
  try {
    await loadScripts();
    const job = await api(`/jobs/${id}`);
    if (token === view.token) paintJob(job, token);
  } catch (e) {
    if (token === view.token && view.active) fill(ui.detail, errorBox(e.message, () => selectJob(id)));
  }
}

function clearSelection() {
  view.token += 1;                // whatever the job that was open is still waiting for is dropped
  view.jobId = null;
  stopPolling();
  markSelected();
  paintEmpty();
}

function paintJob(job, token) {
  stopPolling();
  const refresh = (next) => {     // after an answer: the job, and also the lists on both sides (a new rule, a changed count)
    if (token === view.token) paintJob(next, token);
    loadJobs();
    loadMemory();
  };
  const open = job.cards.filter((c) => c.status === 'open');
  const order = { password: 0, question: 1, failure: 2, alias: 3, note: 4 };
  open.sort((a, b) => order[a.type] - order[b.type] || (a.type === 'note' ? issueOf(a).localeCompare(issueOf(b)) : a.type === 'alias' ? a.payload.kind.localeCompare(b.payload.kind) : 0));      // notes about one issue sit together
  const notes = open.filter((c) => c.type === 'note');
  bulk = notes.length > 1 ? notesTogether(job, notes, refresh) : null;
  const aliases = open.filter((c) => c.type === 'alias');
  bulkAlias = aliases.length > 1 ? aliasesTogether(job, aliases, refresh) : null;
  const handled = job.cards.filter((c) => c.status !== 'open');
  const running = BUSY.has(job.status);
  const again = h('button', { type: 'button', class: job.pending_changes ? 'primary' : '' }, job.pending_changes ? `Apply ${job.pending_changes} answer${job.pending_changes === 1 ? '' : 's'} and run again` : 'Run again');
  again.disabled = running || open.some((c) => c.type === 'question' || c.type === 'password');
  again.addEventListener('click', () => busy(again, async () => {
    const r = await api(`/jobs/${job.id}/rerun`, { method: 'POST', body: {} });
    refresh(r.job);
  }));
  const del = h('button', { type: 'button', class: 'danger' }, 'Delete');
  del.disabled = running;
  del.addEventListener('click', () => {
    if (!confirm('Delete this job and its files? The memory (rules) is kept.')) return;
    busy(del, async () => {
      await api(`/jobs/${job.id}`, { method: 'DELETE' });
      clearSelection();
      loadJobs();
      go(narrow() ? '#/jobs' : '#/');
    });
  });
  fill(ui.detail,
    h('div', { class: 'panel job-head' },
      h('div', { class: 'row spread' },
        h('div', null, h('h1', null, job.script.label), h('p', { class: 'sub' }, `${job.params.fromDate} to ${job.params.toDate} - started ${when(job.created_at)} `, chip(STATUS[job.status] || job.status, job.status))),
        h('div', { class: 'row' }, again, del)),
      banner(job)),
    open.length ? h('h2', null, 'Needs you') : null,
    open.filter((c) => c.type !== 'note' && c.type !== 'alias').map((c) => cardView(job, c, refresh)),
    aliases.length ? h('div', { class: 'names-block' }, bulkAlias ? [bulkAlias.setup, bulkAlias.bar] : null, aliases.map((c) => cardView(job, c, refresh))) : null,
    notes.length ? h('div', { class: 'notes-block' }, bulk ? bulk.bar : null, notes.map((c) => cardView(job, c, refresh))) : null,
    resultPanel(job),
    handled.length ? h('h2', null, 'Handled') : null,
    handled.map((c) => cardView(job, c, refresh)),
    h('h2', null, 'What the agent did'),
    h('div', { class: 'panel' }, timeline(job.events)),
    h('h2', null, 'Files'),
    h('div', { class: 'panel' }, h('dl', { class: 'kv' }, job.files.map((f) => [h('dt', null, f.role), h('dd', null, `${f.name} (${(f.size / 1048576).toFixed(1)} MB)`)]))),
    job.log ? h('details', null, h('summary', null, 'Technical details (the converter\'s output)'), h('pre', { class: 'log' }, job.log)) : null);
  bulk = null;
  bulkAlias = null;
  if (running) {
    pollTimer = setTimeout(async () => {
      try {
        const next = await api(`/jobs/${job.id}`);
        if (token === view.token) paintJob(next, token);
      } catch {
        if (token === view.token) pollTimer = setTimeout(() => selectJob(job.id), 3000);
      }
    }, 1500);
  }
}

// ------------------------------------------------------------------ the left: the jobs
function keepInView(box, el) {
  const top = el.offsetTop;
  const bottom = top + el.offsetHeight;
  if (top < box.scrollTop) box.scrollTop = top;
  else if (bottom > box.scrollTop + box.clientHeight) box.scrollTop = bottom - box.clientHeight;
}

/** Highlight the selected job in the list (no rebuild, so the list keeps its place). */
function markSelected() {
  if (!ui.jobsList) return;
  let current = null;
  ui.jobsList.querySelectorAll('a.job').forEach((a) => {
    const on = a.dataset.id === view.jobId;
    a.classList.toggle('on', on);
    if (on) {
      a.setAttribute('aria-current', 'true');
      current = a;
    } else {
      a.removeAttribute('aria-current');
    }
  });
  if (current) keepInView(ui.jobsList, current);
}

function jobItem(j) {
  const on = j.id === view.jobId;
  const open = j.open_cards || 0;
  const period = j.params && j.params.fromDate ? `${j.params.fromDate} - ${j.params.toDate}` : '';
  const result = j.rows === null || j.rows === undefined ? '' : ` - ${num(j.rows)} rows${j.flagged ? `, ${num(j.flagged)} flagged` : ''}`;
  return h('li', null, h('a', { class: `job${on ? ' on' : ''}`, href: `#/jobs/${j.id}`, 'data-id': j.id, 'aria-current': on ? 'true' : null },
    h('span', { class: 'job-name' }, j.script),
    h('span', { class: 'job-line' }, BUSY.has(j.status) ? h('span', { class: 'spinner sm', 'aria-hidden': 'true' }) : null, chip(STATUS[j.status] || j.status, j.status), open ? h('span', { class: 'job-open' }, `${open} to do`) : null),
    period ? h('span', { class: 'job-meta' }, period) : null,
    h('span', { class: 'job-meta' }, `Started ${when(j.created_at)}${result}`)));
}

function buildJobsRail() {
  ui.jobsBadge = h('span', { class: 'badge', hidden: true });
  const add = h('button', { type: 'button', class: 'small primary' }, '+ New run');
  add.addEventListener('click', () => {
    setPane('work');
    setComposer(true);
    go('#/');
    focusComposer();
  });
  ui.jobsList = h('div', { class: 'rail-body' }, h('div', { class: 'empty small' }, 'Loading...'));
  ui.jobsCol.append(h('div', { class: 'rail-head row spread' }, h('h2', { class: 'col-title' }, 'Jobs', ui.jobsBadge), add), ui.jobsList);
}

function paintJobs() {
  const need = view.jobs.filter((j) => j.open_cards > 0).length;
  const text = `${need} need${need === 1 ? 's' : ''} you`;
  ui.jobsBadge.hidden = !need;
  ui.jobsBadge.textContent = text;
  ui.tabBadge.hidden = !need;
  ui.tabBadge.textContent = String(need);
  document.title = need ? `(${need}) Grid agent` : 'Grid agent';
  const sig = JSON.stringify(view.jobs);
  if (sig !== view.jobsSig) {            // only when something changed: a rebuild would drop the keyboard focus and the scroll place
    view.jobsSig = sig;
    const focused = ui.jobsList.contains(document.activeElement) && document.activeElement.dataset ? document.activeElement.dataset.id : null;
    const scroll = ui.jobsList.scrollTop;
    ui.jobsList.replaceChildren(view.jobs.length ? h('ul', { class: 'joblist' }, view.jobs.map(jobItem)) : h('div', { class: 'empty small' }, 'No jobs yet.'));
    ui.jobsList.scrollTop = scroll;
    if (focused) {
      const again = ui.jobsList.querySelector(`a.job[data-id="${focused}"]`);
      if (again) again.focus({ preventScroll: true });
    }
  }
  paintEmpty();
}

/** Fetch the list; while a job is queued or running keep fetching, so its chip and counts move by themselves. */
async function loadJobs() {
  if (!view.active) return;
  clearTimeout(view.jobsTimer);
  const seq = ++view.jobsSeq;
  try {
    const r = await api('/jobs');
    if (seq !== view.jobsSeq || !view.active) return;
    const was = new Map(view.jobs.map((j) => [j.id, j.status]));
    const finished = r.jobs.some((j) => BUSY.has(was.get(j.id)) && !BUSY.has(j.status));
    view.jobs = r.jobs;
    view.jobsLoaded = true;
    paintJobs();
    if (finished) loadMemory();          // a run that has just ended may have used rules: their counts moved
    markSelected();
  } catch (e) {
    if (seq !== view.jobsSeq || !view.active) return;
    if (!view.jobsLoaded) ui.jobsList.replaceChildren(errorBox(e.message, loadJobs));
  }
  if (view.active && view.jobs.some((j) => BUSY.has(j.status))) view.jobsTimer = setTimeout(loadJobs, 3000);
}

// ------------------------------------------------------------------ the right: the memory
function ruleSummary(r) {
  const d = r.data;
  if (r.kind === 'answer') return [h('strong', null, d.problem || r.match_text.slice(0, 120)), ' -> ', d.kind === 'text' ? d.text : (d.labels || []).join('; ')];
  if (r.kind === 'alias') return [h('strong', null, `${KIND[d.kind] || d.kind} ${d.from}`), ` -> ${d.to}`, d.template ? h('span', { class: 'small muted' }, '  (also with another month in the name)') : null];
  if (r.kind === 'standing_note') return [h('strong', null, 'No action needed - never flagged again: '), d.text];
  if (r.kind === 'failure_note') return [h('strong', null, 'When it stops with: '), d.error];
  return [h('strong', null, 'Explained - never flagged again: '), d.text];
}

function ruleCard(rule) {
  const text = h('div', { class: 'rule-text' }, ruleSummary(rule), rule.enabled ? null : [' ', chip('Off')]);
  let more = null;
  if (text.textContent.length > 280) {          // a long note: five lines, the rest on a click
    text.classList.add('clamp');
    more = h('button', { type: 'button', class: 'more' }, 'Show more');
    more.addEventListener('click', () => {
      const clamped = text.classList.toggle('clamp');
      more.textContent = clamped ? 'Show more' : 'Show less';
    });
  }
  const toggle = h('button', { type: 'button', class: 'small' }, rule.enabled ? 'Switch off' : 'Switch on');
  toggle.addEventListener('click', () => busy(toggle, async () => {
    await api(`/rules/${rule.id}`, { method: 'PATCH', body: { enabled: !rule.enabled } });
    await loadMemory();
  }));
  const del = h('button', { type: 'button', class: 'small danger' }, 'Delete');
  del.addEventListener('click', () => {
    if (!confirm('Delete this rule for good? Past jobs keep their record of it.')) return;
    busy(del, async () => {
      await api(`/rules/${rule.id}`, { method: 'DELETE' });
      await loadMemory();
    });
  });
  const used = rule.uses ? `Used ${rule.uses}x${rule.last_used_at ? `, last ${when(rule.last_used_at)}` : ''}` : 'Not used yet';
  return h('article', { class: `rule${rule.enabled ? '' : ' off'}` },
    text, more,
    h('p', { class: 'rule-scope small muted' }, h('strong', null, 'Applies to: '), scopeLabel(rule.scope), rule.explanation ? [' - ', h('strong', null, 'Why: '), rule.explanation] : null),
    h('div', { class: 'rule-foot' }, h('span', { class: 'small muted' }, used), h('div', { class: 'row tight' }, toggle, del)));
}

function buildMemoryRail() {
  ui.memStats = h('p', { class: 'small muted rail-sub' }, ' ');
  ui.memKind = h('select', { 'aria-label': 'Kind' },
    h('option', { value: '' }, 'All kinds'), h('option', { value: 'answer' }, 'Answers to questions'), h('option', { value: 'alias' }, 'Name mappings'),
    h('option', { value: 'standing_note' }, 'Notes needing no action'), h('option', { value: 'note_explanation' }, 'Explained notes'), h('option', { value: 'failure_note' }, 'Notes about stops'));
  ui.memQ = h('input', { type: 'search', placeholder: 'Search the memory', 'aria-label': 'Search the memory', autocomplete: 'off' });
  ui.memList = h('div', { class: 'rail-body' }, h('div', { class: 'empty small' }, 'Loading...'));
  ui.memKind.addEventListener('change', () => loadMemory());
  let timer = null;
  ui.memQ.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(loadMemory, 250); });
  const importer = h('input', { type: 'file', accept: '.json,application/json', class: 'hidden' });
  importer.addEventListener('change', async () => {
    try {
      const text = await importer.files[0].text();
      const r = await api('/rules/import', { method: 'POST', body: JSON.parse(text) });
      toast(`${r.imported} rules imported.`);
      await loadMemory();
    } catch (e) {
      toast(e.message, true);
    } finally {
      importer.value = '';
    }
  });
  ui.memCol.append(
    h('div', { class: 'rail-head' },
      h('div', { class: 'row spread' }, h('h2', { class: 'col-title' }, 'Memory'),
        h('div', { class: 'row tight' }, h('a', { class: 'btn small', href: '/api/rules/export', download: 'grid-agent-memory.json' }, 'Export'), h('button', { type: 'button', class: 'small', onclick: () => importer.click() }, 'Import'), importer)),
      ui.memStats,
      h('div', { class: 'filters' }, ui.memKind, ui.memQ)),
    ui.memList);
}

async function loadMemory() {
  if (!view.active) return;
  const seq = ++view.memSeq;          // typing fast sends several searches: only the newest may paint
  try {
    await loadScripts();
    const [r, s] = await Promise.all([api(`/rules?kind=${encodeURIComponent(ui.memKind.value)}&q=${encodeURIComponent(ui.memQ.value)}`), api('/stats')]);
    if (seq !== view.memSeq || !view.active) return;
    ui.memStats.textContent = `${s.enabledRules} of ${s.rules} rules switched on - applied by themselves ${s.uses} times so far.`;
    const note = h('p', { class: 'rail-note small muted' }, 'Everything you explained on a card. Next time the same or a very similar situation comes up, the agent uses it by itself and says so on the job. Switch a rule off when it should not be used any more.');
    ui.memList.replaceChildren(r.rules.length ? h('div', null, r.rules.map(ruleCard)) : h('div', { class: 'empty small' }, 'Nothing here yet. Answer a card on a job and it is remembered.'), note);
  } catch (e) {
    if (seq === view.memSeq && view.active) ui.memList.replaceChildren(errorBox(e.message, loadMemory));
  }
}

// ------------------------------------------------------------------ putting it together
function buildShell() {
  if (view.built) return;
  view.built = true;
  buildJobsRail();
  buildWork();
  buildMemoryRail();
}

function loadAll() {
  view.loaded = true;
  ui.jobsList.replaceChildren(h('div', { class: 'empty small' }, 'Loading...'));
  ui.memList.replaceChildren(h('div', { class: 'empty small' }, 'Loading...'));
  loadJobs();
  loadMemory();
  if (!ui.composer.form) loadComposer();
}

/** Leaving the page (sign-in screen): drop the timers and everything shown, so a later sign-in starts clean and nothing of the last session stays in the page. */
function resetSession() {
  Object.assign(view, { active: false, loaded: false, jobId: null, jobs: [], jobsLoaded: false, jobsSig: '' });
  view.token += 1;
  view.jobsSeq += 1;
  view.memSeq += 1;
  clearTimeout(view.jobsTimer);
  stopPolling();
  document.title = 'Grid agent';
  if (!view.built) return;
  ui.jobsList.replaceChildren();
  ui.memList.replaceChildren();
  ui.detail.replaceChildren();
  ui.memStats.textContent = ' ';
  ui.memQ.value = '';
  ui.memKind.value = '';
  if (ui.composer.reset) ui.composer.reset();          // a password typed for a converter must not outlive the session
}

function applyRoute(hash) {
  const job = /^#\/jobs\/([0-9a-f]{16})$/.exec(hash);
  if (job) {
    setPane('work');
    setComposer(false);
    if (job[1] !== view.jobId) selectJob(job[1]);
    return;
  }
  if (/^#\/?$/.test(hash)) {
    setPane('work');
    if (view.jobId) clearSelection();
    else paintEmpty();
    setComposer(true);
    return;
  }
  if (hash === '#/jobs') {
    setPane('jobs');
    if (narrow()) ui.jobsCol.focus({ preventScroll: true });
    return;
  }
  if (hash === '#/memory') {
    setPane('memory');
    if (narrow()) ui.memCol.focus({ preventScroll: true });
    else ui.memQ.focus({ preventScroll: true });
    return;
  }
  location.hash = '#/';
}

function route() {
  const hash = location.hash || '#/';
  if (state.required && !state.authed && hash !== '#/login') {
    location.hash = '#/login';
    return;
  }
  if (hash === '#/login') {
    resetSession();
    showView('login');
    renderLogin();
    return;
  }
  showView('shell');
  buildShell();
  view.active = true;
  if (!view.loaded) loadAll();
  try {
    applyRoute(hash);
  } catch (e) {
    fill(ui.detail, errorBox(e.message));
  }
}

async function start() {
  try {
    const me = await (await fetch('/api/me', { credentials: 'same-origin' })).json();
    state.authed = me.authed;
    state.required = me.required;
  } catch {
    state.required = true;
  }
  const out = $('#signout');
  out.hidden = !state.required;
  out.addEventListener('click', async () => {
    await api('/logout', { method: 'POST', body: {} });
    state.authed = false;
    location.hash = '#/login';
  });
  ui.nav.addEventListener('click', (e) => {          // the Run tab goes back to the job that is open, if there is one
    const tab = e.target.closest('a[data-nav="work"]');
    if (tab && view.jobId) {
      e.preventDefault();
      go(`#/jobs/${view.jobId}`);
    }
  });
  window.addEventListener('hashchange', route);
  route();
}

start();
