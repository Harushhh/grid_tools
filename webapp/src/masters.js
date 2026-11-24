/** The portal's master lists (RTO groups, make / model groups): the names a mapping may point to. Read through a small Python helper, cached until the files change. */
import { execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { similarity } from './memory.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const HELPER = path.resolve(here, '..', 'py', 'dump_masters.py');

const tokensOf = (s) => new Set(s.split(/[^a-z0-9]+/).filter((t) => t.length >= 2));

function jaccard(a, b) {
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  let inter = 0;
  for (const t of small) if (large.has(t)) inter += 1;
  return inter / (a.size + b.size - inter || 1);
}

export class MasterNames {
  constructor({ python, gridTools }) {
    this.python = python;
    this.folder = path.join(gridTools, 'master');
    this.stamp = '';
    this.#set({ rto: [], makes: [] });
  }

  #set(names) {
    this.names = names;
    this.sets = { rto: new Set(names.rto), makes: new Set(names.makes) };
    this.entries = {};            // lower-case name + its words, so a search does not redo that for every name
    for (const kind of ['rto', 'makes']) this.entries[kind] = names[kind].map((n) => ({ n, low: n.toLowerCase(), tok: tokensOf(n.toLowerCase()) }));
  }

  #stamp() {
    try {
      return fs.readdirSync(this.folder).filter((n) => /\.xlsx?$/i.test(n)).map((n) => `${n}:${fs.statSync(path.join(this.folder, n)).mtimeMs}`).join('|');
    } catch {
      return '';
    }
  }

  /** (Re)load when the master files changed. Never throws: without masters nothing is applied by itself. */
  async ensure() {
    const stamp = this.#stamp();
    if (stamp && stamp === this.stamp) return this;
    const out = await new Promise((resolve) => {
      execFile(this.python, ['-I', '-X', 'utf8', HELPER, this.folder], { maxBuffer: 64 * 1024 * 1024, windowsHide: true }, (err, stdout) => resolve(err ? null : stdout));
    });
    try {
      const data = JSON.parse(out);
      this.#set({ rto: data.rto, makes: data.makes });
      this.stamp = stamp;
    } catch {
      this.#set({ rto: [], makes: [] });
    }
    return this;
  }

  set(kind) {
    return this.sets[kind] || new Set();
  }

  has(kind, name) {
    return this.set(kind).has(name);
  }

  loaded() {
    return this.sets.rto.size > 0;
  }

  /**
   * The master names that match a typed text: the same name (capitals aside), starts with, contains - then the ones that share the most words,
   * re-ranked by text similarity. (The master has thousands of names and a grid can lack hundreds of them: the expensive comparison is only made for the closest few.)
   */
  search(kind, q, limit = 20) {
    const entries = this.entries[kind] || [];
    const needle = String(q || '').trim().toLowerCase();
    if (!needle) return entries.slice(0, limit).map((e) => e.n);
    const words = tokensOf(needle);
    const hits = [];
    const near = [];
    for (const e of entries) {
      let score = 0;
      if (e.low === needle) score = 100;
      else if (e.low.startsWith(needle)) score = 80;
      else if (e.low.includes(needle)) score = 60;
      if (score) hits.push([score, e.n]);
      else if (words.size) {
        const j = jaccard(words, e.tok);
        if (j > 0) near.push([j, e.n]);
      }
    }
    near.sort((a, b) => b[0] - a[0]);
    for (const [, n] of near.slice(0, 30)) {
      const s = similarity(n, q, true);
      if (s >= 0.45) hits.push([40 * s, n]);
    }
    return hits.sort((a, b) => b[0] - a[0] || a[1].length - b[1].length).slice(0, limit).map((x) => x[1]);
  }
}
