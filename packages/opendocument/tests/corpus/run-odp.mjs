#!/usr/bin/env node
// One-command ODP corpus harness (#1934), run from the checkout under test:
//
//   node packages/opendocument/tests/corpus/run-odp.mjs [--build] [--jobs 4]
//     [--filter substr] [--limit N] [--label name] [--fresh] [--worst 40]
//     [--port 5192] [--skip-browser]
//
// 1. Lists every ODP with a LibreOffice PDF in the private corpus that
//    convert.mjs built (main checkout, packages/opendocument/public/private/odp).
// 2. Classifies each file natively with the reader's `inspect_odp` example
//    (admission, and which slides are hidden: LibreOffice's PDF export skips
//    them).
// 3. Loads each file in Chrome through this checkout's `odp` model source and
//    the PPTX renderer, and diffs every shown slide against its PDF page
//    (odp-corpus.spec.ts).
// 4. Writes report.md / summary.json and copies side-by-side images of the
//    worst slides into worst/.
//
// Output: <corpus>/packages/opendocument/public/private/harness/odp-<label>/
// (git-ignored). A rerun with the same label resumes; --fresh starts over.
// The run holds /private/tmp/ooxml-vrt.lock (shared with the VRT runs).
import { execFileSync, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join, relative } from 'node:path';
import { parseArgs } from 'node:util';
import { checkoutRoot, corpusRepo } from './paths.mjs';

const { values: args } = parseArgs({
  options: {
    build: { type: 'boolean', default: false },
    jobs: { type: 'string', default: '4' },
    filter: { type: 'string' },
    limit: { type: 'string' },
    label: { type: 'string' },
    fresh: { type: 'boolean', default: false },
    worst: { type: 'string', default: '40' },
    port: { type: 'string', default: '5192' },
    'skip-browser': { type: 'boolean', default: false },
    'no-lock': { type: 'boolean', default: false },
  },
});

const LOCK = '/private/tmp/ooxml-vrt.lock';
let locked = false;
if (!args['no-lock']) {
  try {
    mkdirSync(LOCK);
    locked = true;
    writeFileSync(join(LOCK, 'owner'), `odp-harness ${process.pid}\n`);
  } catch {
    throw new Error(`${LOCK} is held by another run; wait for it or pass --no-lock for a small subset`);
  }
}
const release = () => { if (locked) { rmSync(LOCK, { recursive: true, force: true }); locked = false; } };
process.on('exit', release);
process.on('SIGINT', () => { release(); process.exit(130); });

const corpus = join(corpusRepo(), 'packages/opendocument/public/private');
const odpRoot = join(corpus, 'odp');
const parserDir = join(checkoutRoot, 'packages/opendocument/parser');
const run = (cmd, argv, opts = {}) => execFileSync(cmd, argv, { stdio: 'inherit', ...opts });
const git = (...argv) => execFileSync('git', argv, { cwd: checkoutRoot, encoding: 'utf8' }).trim();

const sha = git('rev-parse', '--short=10', 'HEAD');
const dirty = git('status', '--porcelain', '--untracked-files=no').length > 0;
const label = args.label ?? `${sha}${dirty ? '-dirty' : ''}`;
if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/u.test(label)) throw new Error(`invalid --label ${JSON.stringify(label)}`);
const runDir = join(corpus, 'harness', `odp-${label}`);
if (args.fresh) rmSync(runDir, { recursive: true, force: true });
mkdirSync(runDir, { recursive: true });

// ---- builds ----------------------------------------------------------------
const pptxWasm = join(checkoutRoot, 'packages/pptx/src/wasm/pptx_parser_bg.wasm');
const odpWasm = join(checkoutRoot, 'packages/opendocument/src/wasm-odp/opendocument_odp_bg.wasm');
if (args.build || !existsSync(pptxWasm)) run('pnpm', ['wasm'], { cwd: join(checkoutRoot, 'packages/pptx') });
if (args.build || !existsSync(odpWasm)) run('pnpm', ['wasm:odp'], { cwd: join(checkoutRoot, 'packages/opendocument') });
run('cargo', ['build', '--release', '--example', 'inspect_odp', '--features', 'odp'], { cwd: parserDir });
const nativeBinary = join(checkoutRoot, 'target/release/examples/inspect_odp');

// Results are reused only by a run of the same code and WASM.
const sha256 = (data) => createHash('sha256').update(data).digest('hex');
const fingerprint = sha256(JSON.stringify({
  head: git('rev-parse', 'HEAD'),
  diff: sha256(execFileSync('git', ['diff', 'HEAD'], { cwd: checkoutRoot, maxBuffer: 256 << 20 })),
  pptxWasm: sha256(readFileSync(pptxWasm)),
  odpWasm: sha256(readFileSync(odpWasm)),
}));
const fingerprintPath = join(runDir, 'fingerprint.txt');
if (existsSync(join(runDir, 'results')) && readdirSync(join(runDir, 'results')).length > 0) {
  const previous = existsSync(fingerprintPath) ? readFileSync(fingerprintPath, 'utf8').trim() : '';
  if (previous !== fingerprint) throw new Error(`${runDir} holds results from different code; rerun with --fresh or another --label`);
}
writeFileSync(fingerprintPath, `${fingerprint}\n`);

// ---- file list and native classification -----------------------------------
function walk(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) walk(path, out);
    else if (entry.name.endsWith('.odp')) out.push(path);
  }
  return out;
}
if (!existsSync(odpRoot)) throw new Error(`No ODP corpus at ${odpRoot}; run convert.mjs first`);
let files = walk(odpRoot).sort().map((odp) => {
  const rel = relative(odpRoot, odp);
  const hash = createHash('sha256').update(rel).digest('hex').slice(0, 12);
  const readable = rel.replace(/\.odp$/u, '').replace(/[^A-Za-z0-9._-]+/gu, '_').slice(-80);
  return { id: `${readable}-${hash}`, rel, odp, pdf: odp.replace(/\.odp$/u, '.pdf') };
}).filter((file) => existsSync(file.pdf));
if (args.filter) files = files.filter((file) => file.rel.includes(args.filter));
if (args.limit) files = files.slice(0, Number(args.limit));

const native = {};
for (let i = 0; i < files.length; i += 50) {
  const batch = files.slice(i, i + 50);
  const out = spawnSync(nativeBinary, batch.map((f) => f.odp), { encoding: 'utf8', maxBuffer: 256 << 20 });
  for (const file of batch) {
    const line = out.stdout.split('\n').find((l) => l.startsWith(`${file.odp}: `));
    native[file.id] = line ? line.slice(file.odp.length + 2) : `panic: ${(out.stderr.split('\n').find((l) => l.includes('panicked')) ?? '').trim()}`;
  }
}
writeFileSync(join(runDir, 'native.json'), `${JSON.stringify(native, null, 2)}\n`);
files = files.map((file) => {
  const hidden = /hidden=\[([^\]]*)\]/u.exec(native[file.id] ?? '')?.[1];
  return { ...file, hidden: hidden ? hidden.split(',').map((v) => Number(v.trim())).filter(Number.isFinite) : [] };
});
writeFileSync(join(runDir, 'files.json'), `${JSON.stringify(files, null, 2)}\n`);
console.log(`${files.length} ODP files with LibreOffice PDFs -> ${runDir}`);

// ---- browser pass ---------------------------------------------------------------
if (!args['skip-browser']) {
  const config = join(checkoutRoot, 'packages/opendocument/tests/corpus/playwright-odp.config.ts');
  const browser = spawnSync('pnpm', ['exec', 'playwright', 'test', '--config', config], {
    cwd: join(checkoutRoot, 'packages/opendocument'),
    stdio: 'inherit',
    env: {
      ...process.env,
      ODF_HARNESS_RUN_DIR: runDir,
      ODF_HARNESS_CHECKOUT: checkoutRoot,
      ODF_HARNESS_JOBS: args.jobs,
      ODF_HARNESS_PORT: args.port,
    },
  });
  if (browser.status !== 0) console.warn(`playwright exited with ${browser.status}`);
}

// ---- report ---------------------------------------------------------------------
const results = [];
const resultsDir = join(runDir, 'results');
if (existsSync(resultsDir)) {
  for (const name of readdirSync(resultsDir)) results.push(JSON.parse(readFileSync(join(resultsDir, name), 'utf8')));
}
const byStatus = (s) => results.filter((r) => r.status === s).length;
const opened = byStatus('opened');
const matched = results.filter((r) => r.slideCountMatch).length;
const slides = results.flatMap((r) => (r.slides ?? []).map((s) => ({ ...s, rel: r.rel })));
const scored = slides.filter((s) => typeof s.layoutDiffPct === 'number').sort((a, b) => a.layoutDiffPct - b.layoutDiffPct);
const pct = (q) => scored.length ? scored[Math.min(scored.length - 1, Math.floor(q * scored.length))].layoutDiffPct : NaN;
const mean = scored.length ? scored.reduce((sum, s) => sum + s.layoutDiffPct, 0) / scored.length : NaN;
const worst = scored.slice(-Number(args.worst)).reverse();
const summary = {
  label, sha, dirty,
  files: files.length,
  results: results.length,
  missing: files.length - results.length,
  opened, rejected: byStatus('rejected'), errors: byStatus('error'),
  openedPct: files.length ? opened / files.length * 100 : 0,
  slideCountMatch: matched,
  slideCountMatchPct: opened ? matched / opened * 100 : 0,
  slidesScored: scored.length,
  slideErrors: slides.length - scored.length,
  layoutDiff: { p50: pct(0.5), p90: pct(0.9), mean },
  worst: worst.map((s) => ({ rel: s.rel, slide: s.slide, layoutDiffPct: s.layoutDiffPct, diffPct: s.diffPct })),
};
writeFileSync(join(runDir, 'summary.json'), `${JSON.stringify(summary, null, 2)}\n`);
const worstDir = join(runDir, 'worst');
rmSync(worstDir, { recursive: true, force: true });
mkdirSync(worstDir, { recursive: true });
for (const [rank, s] of worst.entries()) {
  if (s.image && existsSync(s.image)) copyFileSync(s.image, join(worstDir, `${String(rank + 1).padStart(3, '0')}-${basename(s.image)}`));
}
const fmt = (v) => (Number.isFinite(v) ? v.toFixed(2) : 'n/a');
const report = [
  `# ODP corpus harness — ${label}`,
  '',
  `- Files: ${files.length} (results ${results.length}, missing ${summary.missing})`,
  `- Opened: ${opened} (${fmt(summary.openedPct)}%), rejected ${summary.rejected}, errors ${summary.errors}`,
  `- Shown-slide count equals PDF pages: ${matched} / ${opened} (${fmt(summary.slideCountMatchPct)}%)`,
  `- Slides scored: ${scored.length} (render errors / unmatched: ${summary.slideErrors})`,
  `- Layout diff (edge structure, %): p50 ${fmt(summary.layoutDiff.p50)}, p90 ${fmt(summary.layoutDiff.p90)}, mean ${fmt(mean)}`,
  '',
  '## Worst slides',
  '',
  '| layout % | pixel % | file | slide |',
  '|---:|---:|---|---:|',
  ...worst.map((s) => `| ${fmt(s.layoutDiffPct)} | ${fmt(s.diffPct)} | ${s.rel} | ${s.slide} |`),
  '',
  '## Errors',
  '',
  ...results.filter((r) => r.status !== 'opened').map((r) => `- ${r.rel}: ${r.status} ${r.message ?? ''}`),
].join('\n');
writeFileSync(join(runDir, 'report.md'), `${report}\n`);
console.log(report.split('\n').slice(0, 8).join('\n'));
release();
