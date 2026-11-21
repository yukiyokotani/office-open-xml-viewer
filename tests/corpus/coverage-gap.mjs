// Renderer code-coverage gap between the public specimen corpus and a local
// private corpus.
//
//   node tests/corpus/coverage-gap.mjs --format docx [--dir SUBDIR] [--limit N]
//        [--workers N] [--max-units N] [--top N] [--port N] [--exclude REGEX]
//
// Each document is rendered in Chromium through the package's VRT fixture
// (every page / slide / sheet, up to --max-units) with V8 precise block
// coverage on. Coverage of the served TypeScript is mapped back through the
// dev server's inline source maps to source lines of packages/*/src. The report
// lists the source lines that some private document executes but no public
// specimen does, grouped into hunks and ranked by the number of private
// documents that reach them: the renderer behaviour the corpus does not
// exercise yet.
//
// Only code on the page's main thread is measured: parsing runs in a worker
// (Rust/WASM) and is covered by the structural report of gap.mjs instead.
//
// Private files are read recursively from packages/<fmt>/public/private/<fmt>/
// (--dir restricts to one subdirectory, '.' to the top level). The full report,
// which names private files, is written inside that ignored private directory
// (coverage-gap.json); stdout shows hunks with source text and counts only.

import { spawn } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { REPO_ROOT, outputDir } from './lib/cases.mjs';

const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : fallback;
};
const format = opt('--format');
const ADAPTERS = {
  docx: {
    url: (path, isPublic) => `/tests/visual/fixture.html?file=${encodeURIComponent(path)}${isPublic ? '&width=595&fonts=corpus' : '&width=612'}`,
    count: 'pageCount',
    render: 'renderDocxVrtPage',
  },
  pptx: {
    url: (path, isPublic) => `/tests/visual/fixture.html?pptx=${encodeURIComponent(path.replace(/\.pptx$/, ''))}${isPublic ? '&fonts=corpus' : ''}`,
    count: 'slideCount',
    render: 'renderPptxVrtSlide',
  },
  xlsx: {
    url: (path, isPublic) => `/tests/visual/fixture.html?file=${encodeURIComponent(path)}${isPublic ? '&rows=40&cols=10&width=1280&height=960&fonts=corpus' : ''}`,
    count: 'sheetCount',
    render: 'renderXlsxVrtSheet',
  },
};
const adapter = ADAPTERS[format];
if (!adapter) throw new Error('--format docx|pptx|xlsx is required');
const workers = Number(opt('--workers', 4));
const maxUnits = Number(opt('--max-units', 40));
const limit = Number(opt('--limit', Infinity));
const top = Number(opt('--top', 40));
const port = Number(opt('--port', 5197));
const subdir = opt('--dir');
// Font resolution differs by environment, not by document: the public corpus
// loads its own two families, private files fall back to host fonts. Those
// modules are left out unless --exclude '' is given.
const exclude = opt('--exclude', 'font') ? new RegExp(opt('--exclude', 'font')) : null;

const pkgDir = join(REPO_ROOT, 'packages', format);
const privateRoot = join(pkgDir, 'public', 'private', format);
if (!existsSync(privateRoot)) {
  console.log(`${format}: no private corpus at ${relative(REPO_ROOT, privateRoot)}/ — nothing to compare.`);
  process.exit(0);
}

function walk(dir, recursive) {
  const out = [];
  for (const name of readdirSync(dir).sort()) {
    const full = join(dir, name);
    const st = statSync(full);
    if (st.isDirectory()) { if (recursive) out.push(...walk(full, true)); } else if (name.endsWith(`.${format}`) && !name.startsWith('~$')) out.push(full);
  }
  return out;
}
const privateFiles = (subdir === '.' ? walk(privateRoot, false) : walk(subdir ? join(privateRoot, subdir) : privateRoot, true)).slice(0, limit);
const publicFiles = walk(outputDir(format), false);
const servedPath = (full) => relative(join(pkgDir, 'public'), full).split('\\').join('/');

// ── Source maps ──────────────────────────────────────────────────────────────
const B64 = Object.fromEntries([...'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'].map((c, i) => [c, i]));
function decodeMappings(mappings) {
  // → per generated line: [[genCol, srcIdx, srcLine]] (srcLine 0-based)
  const lines = [];
  let src = 0; let srcLine = 0; let srcCol = 0;
  for (const lineText of mappings.split(';')) {
    const segs = [];
    let genCol = 0;
    for (const seg of lineText.split(',')) {
      if (!seg) continue;
      const vals = [];
      let value = 0; let shift = 0;
      for (const ch of seg) {
        const d = B64[ch];
        value += (d & 31) << shift;
        if (d & 32) { shift += 5; } else { vals.push(value & 1 ? -(value >>> 1) : value >>> 1); value = 0; shift = 0; }
      }
      genCol += vals[0];
      if (vals.length >= 4) {
        src += vals[1]; srcLine += vals[2]; srcCol += vals[3];
        segs.push([genCol, src, srcLine]);
      }
    }
    lines.push(segs);
  }
  return lines;
}

// Source lines are interned as integer ids: thousands of documents each
// executing tens of thousands of lines do not fit in memory as strings.
const keyIds = new Map(); const keys = [];
const keyId = (key) => { let id = keyIds.get(key); if (id === undefined) { id = keys.length; keys.push(key); keyIds.set(key, id); } return id; };
const seenAll = new Set();

const mapCache = new Map(); // url → { lineStarts, mapped: [[offset, file, line]] } | null
function sourceMapFor(url, source) {
  if (mapCache.has(url)) return mapCache.get(url);
  let result = null;
  const m = /\/\/# sourceMappingURL=data:application\/json;(?:charset=utf-8;)?base64,([A-Za-z0-9+/=]+)\s*$/.exec(source);
  if (m) {
    const map = JSON.parse(Buffer.from(m[1], 'base64').toString('utf8'));
    // Sources are relative to the served module's directory: /src/x.ts lives
    // under the package, /@fs/<abs> outside it.
    const path = decodeURIComponent(new URL(url).pathname);
    const servedFile = path.startsWith('/@fs/') ? path.slice(4) : join(pkgDir, path);
    const files = map.sources.map((s) => {
      const abs = s.startsWith('/') && existsSync(s) ? s : resolve(servedFile, '..', map.sourceRoot ?? '', s);
      return relative(REPO_ROOT, abs).split('\\').join('/');
    });
    const keep = files.map((f) => /^packages\/[^/]+\/src\//.test(f) && !/\.(test|stories)\.ts$/.test(f) && !f.includes('/wasm/'));
    if (keep.some(Boolean)) {
      const lineStarts = [0];
      for (let i = 0; i < source.length; i++) if (source.charCodeAt(i) === 10) lineStarts.push(i + 1);
      const mapped = [];
      decodeMappings(map.mappings).forEach((segs, genLine) => {
        for (const [genCol, src, srcLine] of segs) if (keep[src]) mapped.push([lineStarts[genLine] + genCol, keyId(`${files[src]}:${srcLine + 1}`)]);
      });
      mapped.sort((a, b) => a[0] - b[0]);
      result = { mapped };
    }
  }
  mapCache.set(url, result);
  return result;
}

/** Lines executed (count > 0) and lines seen but not executed, per entry. */
function linesOf(entry, executed) {
  const map = sourceMapFor(entry.url, entry.source);
  if (!map) return;
  // Innermost range wins: V8 nests block ranges inside function ranges and
  // lists them outer-first; paint counts in that order over the mapped offsets.
  const ranges = entry.functions.flatMap((f) => f.ranges)
    .sort((a, b) => a.startOffset - b.startOffset || b.endOffset - a.endOffset);
  const counts = new Int32Array(map.mapped.length).fill(-1);
  const offsets = map.mapped.map((m) => m[0]);
  const lowerBound = (x) => { let lo = 0; let hi = offsets.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (offsets[mid] < x) lo = mid + 1; else hi = mid; } return lo; };
  for (const r of ranges) {
    for (let i = lowerBound(r.startOffset); i < offsets.length && offsets[i] < r.endOffset; i++) counts[i] = r.count;
  }
  map.mapped.forEach(([, id], i) => {
    if (counts[i] < 0) return;
    seenAll.add(id);
    if (counts[i] > 0) executed.add(id);
  });
}

// ── Rendering ────────────────────────────────────────────────────────────────
async function startServer() {
  const child = spawn('pnpm', ['exec', 'vite', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: pkgDir, stdio: 'ignore' });
  const base = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 120; i++) {
    try { if ((await fetch(base + '/tests/visual/fixture.html')).ok) return { child, base }; } catch {}
    await new Promise((r) => setTimeout(r, 500));
  }
  child.kill();
  throw new Error('vite did not start');
}

async function measure(context, base, full, isPublic) {
  const page = await context.newPage();
  const executed = new Set();
  let units = 0; let error = null;
  try {
    await page.coverage.startJSCoverage({ resetOnNavigation: false });
    await page.goto(base + adapter.url(servedPath(full), isPublic), { timeout: 60000 });
    await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status), null, { timeout: 120000 });
    const status = await page.evaluate(() => [document.body.dataset.status, document.body.dataset.errorMessage]);
    if (status[0] === 'error') throw new Error(status[1]);
    const total = Number(await page.evaluate((k) => document.body.dataset[k], adapter.count)) || 1;
    units = Math.min(total, maxUnits);
    for (let i = 1; i < units; i++) {
      await page.evaluate(([fn, idx]) => window[fn](idx), [adapter.render, i]);
    }
  } catch (e) {
    error = String(e?.message ?? e).split('\n')[0];
  }
  try {
    for (const entry of await page.coverage.stopJSCoverage()) linesOf(entry, executed);
  } catch {}
  await page.close();
  return { executed: Int32Array.from(executed), units, error };
}

async function run(files, isPublic, context, base, label) {
  const results = new Array(files.length);
  let next = 0; let done = 0;
  const t0 = Date.now();
  await Promise.all(Array.from({ length: workers }, async () => {
    while (next < files.length) {
      const i = next++;
      results[i] = await measure(context, base, files[i], isPublic);
      done++;
      if (done % 25 === 0 || done === files.length) {
        process.stderr.write(`  ${label}: ${done}/${files.length} (${Math.round((Date.now() - t0) / 1000)} s)\n`);
      }
    }
  }));
  return results;
}

// ── Main ─────────────────────────────────────────────────────────────────────
const { child, base } = await startServer();
const browser = await chromium.launch();
try {
  const context = await browser.newContext({ viewport: { width: 1280, height: 960 } });
  console.log(`== ${format}: ${publicFiles.length} public specimens, ${privateFiles.length} private files`
    + `${subdir ? ` (dir ${subdir})` : ''}; up to ${maxUnits} units each`);
  const pub = await run(publicFiles, true, context, base, 'public');
  const priv = await run(privateFiles, false, context, base, 'private');

  const pubExec = new Set(); for (const r of pub) for (const id of r.executed) pubExec.add(id);
  const allSeen = seenAll;
  const privHits = new Map(); // line id → [private file indices]
  const privExec = new Set();
  priv.forEach((r, i) => {
    for (const id of r.executed) {
      privExec.add(id);
      if (!pubExec.has(id)) (privHits.get(id) ?? privHits.set(id, []).get(id)).push(i);
    }
  });

  // Hunks: runs of gap lines in one file, at most two lines apart.
  const byFile = new Map();
  for (const id of privHits.keys()) {
    const k = keys[id];
    const at = k.lastIndexOf(':');
    const f = k.slice(0, at); const line = Number(k.slice(at + 1));
    if (exclude?.test(f)) continue;
    (byFile.get(f) ?? byFile.set(f, []).get(f)).push(line);
  }
  const hunks = [];
  for (const [file, lines] of byFile) {
    lines.sort((a, b) => a - b);
    let cur = null;
    for (const line of lines) {
      if (cur && line - cur.end <= 2) { cur.end = line; cur.lines.push(line); } else { cur = { file, start: line, end: line, lines: [line] }; hunks.push(cur); }
    }
  }
  for (const h of hunks) {
    const docs = new Set(h.lines.flatMap((l) => privHits.get(keyIds.get(`${h.file}:${l}`))));
    h.docs = [...docs];
  }
  hunks.sort((a, b) => b.docs.length - a.docs.length || b.lines.length - a.lines.length);

  const failed = [...pub.map((r, i) => [publicFiles[i], r]), ...priv.map((r, i) => [privateFiles[i], r])].filter(([, r]) => r.error);
  const pct = (a, b) => (b ? (100 * a / b).toFixed(1) : '0.0');
  console.log(`   measured lines ${allSeen.size}; executed by public ${pubExec.size} (${pct(pubExec.size, allSeen.size)} %), `
    + `by private ${privExec.size}; private-only ${privHits.size} lines in ${hunks.length} hunks`);
  if (failed.length) console.log(`   ${failed.length} files failed to render (listed in the JSON report)`);

  const src = new Map();
  const text = (file, line) => {
    if (!src.has(file)) src.set(file, readFileSync(join(REPO_ROOT, file), 'utf8').split('\n'));
    return src.get(file)[line - 1]?.trim() ?? '';
  };
  for (const h of hunks.slice(0, top)) {
    console.log(`\n${String(h.docs.length).padStart(5)} docs  ${h.file}:${h.start}${h.end > h.start ? `-${h.end}` : ''}  (${h.lines.length} lines)`);
    for (const l of h.lines.slice(0, 4)) console.log(`         ${String(l).padStart(5)}  ${text(h.file, l).slice(0, 110)}`);
    if (h.lines.length > 4) console.log('         …');
  }

  const reportPath = join(privateRoot, '..', 'coverage-gap.json');
  writeFileSync(reportPath, JSON.stringify({
    format,
    generatedAt: new Date().toISOString(),
    publicFiles: publicFiles.map(servedPath),
    privateFiles: privateFiles.map(servedPath),
    failed: failed.map(([f, r]) => ({ file: servedPath(f), error: r.error })),
    hunks: hunks.map((h) => ({
      file: h.file, start: h.start, end: h.end, lines: h.lines,
      docs: h.docs.length, examples: h.docs.slice(0, 10).map((i) => servedPath(privateFiles[i])),
    })),
  }, null, 1));
  console.log(`\nFull report (names private files): ${relative(REPO_ROOT, reportPath)}`);
} finally {
  await browser.close();
  child.kill();
}
