#!/usr/bin/env node
// One-command ODT corpus harness (#1934), run from the checkout under test:
//
//   node packages/opendocument/tests/corpus/run.mjs [--rules off|on] [--build]
//     [--jobs 4] [--filter substr] [--limit N] [--label name] [--fresh]
//     [--worst 40] [--port 5191]
//
// 1. Lists every ODT with a LibreOffice PDF in the private corpus that
//    convert.mjs built (main checkout, packages/opendocument/public/private/odt).
// 2. Classifies each file natively with the reader's `inspect_odt` example,
//    which reports the exact `ODF_UNSUPPORTED:<code>` the browser path hides.
// 3. Loads each file in Chrome through this checkout's `odt` model source and
//    the DOCX renderer, and diffs every page against the LibreOffice PDF
//    (odt-corpus.spec.ts).
// 4. Writes report.md / summary.json and side-by-side images of the worst
//    pages (report.mjs).
//
// --rules on is a dry run with the pending LibreOffice compatibility rules
// (CompatRules) enabled: the harness flips LIBREOFFICE_RULES in an isolated
// copy of the reader's crates under target/ and builds a separate WASM and
// native binary from it; the checkout itself is never edited.
//
// --unmask-settings renders the browser pass through that diagnostic build too
// (DIAGNOSTIC: acceptance then is not the shipped reader's). It exercises the
// page diff on files the real reader rejects only for those flags; the flags
// themselves are Word-compatibility layout modes, so part of each diff is
// theirs.
//
// Every run also classifies the files natively with the settings.xml layout
// flag check disabled (diagnostic only): LibreOffice writes Word-compatibility
// flags into every DOCX-derived ODT, which would otherwise hide all content
// rejections behind `compatibility-setting`.
//
// Output: <corpus>/packages/opendocument/public/private/harness/<label>/
// (git-ignored). A rerun with the same label resumes; --fresh starts over.
import { execFileSync, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cpSync, existsSync, mkdirSync, unlinkSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { parseArgs } from 'node:util';
import { checkoutRoot, corpusRepo, inputsHash } from './paths.mjs';

const { values: args } = parseArgs({
  options: {
    rules: { type: 'string', default: 'off' },
    build: { type: 'boolean', default: false },
    jobs: { type: 'string', default: '4' },
    filter: { type: 'string' },
    limit: { type: 'string' },
    label: { type: 'string' },
    fresh: { type: 'boolean', default: false },
    worst: { type: 'string', default: '40' },
    port: { type: 'string', default: '5191' },
    'skip-browser': { type: 'boolean', default: false },
    'unmask-settings': { type: 'boolean', default: false },
  },
});
if (!['off', 'on'].includes(args.rules)) throw new Error('--rules must be off or on');

const corpus = join(corpusRepo(), 'packages/opendocument/public/private');
const odtRoot = join(corpus, 'odt');
const parserDir = join(checkoutRoot, 'packages/opendocument/parser');
const run = (cmd, argv, opts = {}) => execFileSync(cmd, argv, { stdio: 'inherit', ...opts });
const git = (...argv) => execFileSync('git', argv, { cwd: checkoutRoot, encoding: 'utf8' }).trim();

const sha = git('rev-parse', '--short=10', 'HEAD');
const dirty = git('status', '--porcelain', '--untracked-files=no').length > 0;
const label = args.label
  ?? `${sha}${dirty ? '-dirty' : ''}-rules-${args.rules}${args['unmask-settings'] ? '-settings-unmasked' : ''}`;
// The label names a directory under the ignored harness folder; nothing else.
if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/u.test(label)) throw new Error(`invalid --label ${JSON.stringify(label)}`);
const runDir = join(corpus, 'harness', label);
if (args.fresh) rmSync(runDir, { recursive: true, force: true });
mkdirSync(runDir, { recursive: true });

// ---- builds ----------------------------------------------------------------
const docxWasm = join(checkoutRoot, 'packages/docx/src/wasm/docx_parser_bg.wasm');
const odtWasm = join(checkoutRoot, 'packages/opendocument/src/wasm-odt/opendocument_odt_bg.wasm');
if (args.build || !existsSync(docxWasm)) run('pnpm', ['wasm'], { cwd: join(checkoutRoot, 'packages/docx') });
if (args.build || !existsSync(odtWasm)) run('pnpm', ['wasm:odt'], { cwd: join(checkoutRoot, 'packages/opendocument') });

const exampleArgs = ['build', '--release', '--example', 'inspect_odt', '--features', 'odt'];
const odtSource = join(parserDir, 'src/odt.rs');

/** Pending LibreOffice compatibility rules all switched on. */
function rulesOn(text) {
  const block = /(pub const LIBREOFFICE_RULES: CompatRules = CompatRules \{)([^}]*)(\})/u;
  const match = block.exec(text);
  if (!match || !/:\s*false/u.test(match[2])) throw new Error('LIBREOFFICE_RULES block not found or already all on');
  return text.replace(block, (_, open, body, close) => `${open}${body.replace(/:\s*false/gu, ': true')}${close}`);
}

/** Diagnostic only: settings.xml layout flags no longer reject the file. */
function settingsUnmasked(text) {
  const check = 'return Err(unsupported("compatibility-setting"));';
  if (text.split(check).length !== 2) throw new Error('compatibility-setting check not found exactly once');
  return text.replace(check, 'continue;');
}

/** Crates the ODT reader builds from; the isolated workspace holds only these. */
const ISOLATED_CRATES = ['packages/ooxml-common', 'packages/docx/model', 'packages/opendocument/parser'];

/**
 * Build the native classifier (and optionally the WASM) from a patched COPY
 * of the reader's crates under target/, never touching the checkout: other
 * builds and concurrent harness runs always see the real source. Returns the
 * native binary path.
 */
function patchedBuild(name, patches, wasmOut) {
  const workspace = join(checkoutRoot, `target/odf-harness-src/${name}-${process.pid}`);
  rmSync(workspace, { recursive: true, force: true });
  try {
    for (const crate of ISOLATED_CRATES) {
      cpSync(join(checkoutRoot, crate), join(workspace, crate), {
        recursive: true,
        filter: (path) => !/[\\/](target|pkg)$/u.test(path),
      });
    }
    const manifest = readFileSync(join(checkoutRoot, 'Cargo.toml'), 'utf8')
      .replace(/members\s*=\s*\[[^\]]*\]/u, `members = ${JSON.stringify(ISOLATED_CRATES)}`);
    writeFileSync(join(workspace, 'Cargo.toml'), manifest);
    cpSync(join(checkoutRoot, 'Cargo.lock'), join(workspace, 'Cargo.lock'));
    const source = join(workspace, 'packages/opendocument/parser/src/odt.rs');
    writeFileSync(source, patches.reduce((text, patch) => patch(text), readFileSync(source, 'utf8')));
    const env = { ...process.env, CARGO_TARGET_DIR: join(checkoutRoot, `target/odf-harness-${name}`) };
    const crateDir = join(workspace, 'packages/opendocument/parser');
    run('cargo', exampleArgs, { cwd: crateDir, env });
    if (wasmOut) {
      run('wasm-pack', ['build', '--target', 'web', '--out-dir', wasmOut, '--out-name', 'opendocument_odt', '--features', 'odt'],
        { cwd: crateDir, env });
    }
    return join(env.CARGO_TARGET_DIR, 'release/examples/inspect_odt');
  } finally {
    rmSync(workspace, { recursive: true, force: true });
  }
}

const rulePatches = args.rules === 'on' ? [rulesOn] : [];
let nativeBinary;
let customWasm;
if (args.rules === 'off') {
  run('cargo', exampleArgs, { cwd: parserDir });
  nativeBinary = join(checkoutRoot, 'target/release/examples/inspect_odt');
} else {
  const wasmOut = join(runDir, 'wasm-rules-on');
  nativeBinary = patchedBuild('rules-on', rulePatches, wasmOut);
  customWasm = join(wasmOut, 'opendocument_odt_bg.wasm');
}
// With --unmask-settings the browser pass also renders through the
// diagnostic build, so files rejected only for settings.xml flags open.
const unmaskedWasmOut = args['unmask-settings'] ? join(runDir, 'wasm-settings-unmasked') : undefined;
const unmaskedBinary = patchedBuild(`settings-unmasked-rules-${args.rules}`, [...rulePatches, settingsUnmasked], unmaskedWasmOut);
if (unmaskedWasmOut) customWasm = join(unmaskedWasmOut, 'opendocument_odt_bg.wasm');

// ---- resume guard ------------------------------------------------------------
// Results are reused only by a run of the same code: the fingerprint covers
// HEAD, the uncommitted diff, the WASM the browser loads and the page settings.
const sha256 = (data) => createHash('sha256').update(data).digest('hex');
const fingerprint = sha256(JSON.stringify({
  head: git('rev-parse', 'HEAD'),
  diff: sha256(execFileSync('git', ['diff', 'HEAD'], { cwd: checkoutRoot, maxBuffer: 256 << 20 })),
  // Untracked, non-ignored files (new helpers) by path and content.
  untracked: sha256(JSON.stringify(git('ls-files', '--others', '--exclude-standard', '-z').split('\0').filter(Boolean).sort()
    .map((path) => [path, sha256(readFileSync(join(checkoutRoot, path)))]))),
  docxWasm: sha256(readFileSync(docxWasm)),
  odtWasm: sha256(readFileSync(customWasm ?? odtWasm)),
  native: sha256(readFileSync(nativeBinary)),
  rules: args.rules,
  unmaskSettings: args['unmask-settings'],
  maxPages: process.env.ODF_HARNESS_MAX_PAGES ?? '0',
}));
const fingerprintPath = join(runDir, 'fingerprint.txt');
if (existsSync(join(runDir, 'results')) && readdirSync(join(runDir, 'results')).length > 0) {
  const previous = existsSync(fingerprintPath) ? readFileSync(fingerprintPath, 'utf8').trim() : '';
  if (previous !== fingerprint) {
    throw new Error(`${runDir} holds results from different code or settings; rerun with --fresh or another --label`);
  }
}
writeFileSync(fingerprintPath, `${fingerprint}\n`);
// Recorded once the resume guard has accepted this run, and before any
// classification or result is written, so report.mjs keeps the run's mode
// even for an interrupted run.
writeFileSync(join(runDir, 'meta.json'), `${JSON.stringify({ sha, dirty, rules: args.rules, unmaskSettings: args['unmask-settings'], label }, null, 2)}\n`);

// ---- file list ---------------------------------------------------------------
function walk(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) walk(path, out);
    else if (entry.name.endsWith('.odt')) out.push(path);
  }
  return out;
}
if (!existsSync(odtRoot)) throw new Error(`No ODT corpus at ${odtRoot}; run convert.mjs first`);
let files = walk(odtRoot).sort()
  .map((odt) => {
    const rel = relative(odtRoot, odt);
    // Readable prefix plus a hash of the relative path: unique per file.
    const hash = createHash('sha256').update(rel).digest('hex').slice(0, 12);
    const readable = rel.replace(/\.odt$/u, '').replace(/[^A-Za-z0-9._-]+/gu, '_').slice(-80);
    return { id: `${readable}-${hash}`, rel, odt, pdf: odt.replace(/\.odt$/u, '.pdf') };
  })
  .filter((file) => existsSync(file.pdf));
if (args.filter) files = files.filter((file) => file.rel.includes(args.filter));
if (args.limit) files = files.slice(0, Number(args.limit));
writeFileSync(join(runDir, 'files.json'), `${JSON.stringify(files, null, 2)}\n`);
console.log(`${files.length} ODT files with LibreOffice PDFs -> ${runDir}`);

// ---- native classification ---------------------------------------------------
function classify(binary, batch) {
  const result = spawnSync(binary, batch.map((file) => file.odt), { encoding: 'utf8', maxBuffer: 64 << 20 });
  const lines = result.stdout.split('\n');
  const out = {};
  for (const file of batch) {
    const line = lines.find((candidate) => candidate.startsWith(`${file.odt}: `));
    if (line) out[file.id] = line.slice(file.odt.length + 2);
  }
  if (result.status !== 0 && batch.length > 1) {
    for (const file of batch) if (!(file.id in out)) Object.assign(out, classify(binary, [file]));
  } else if (result.status !== 0) {
    out[batch[0].id] = `panic: ${(result.stderr.split('\n').find((l) => l.includes('panicked')) ?? '').trim()}`;
  }
  return out;
}
function classifyAll(binary) {
  const out = {};
  for (let i = 0; i < files.length; i += 100) Object.assign(out, classify(binary, files.slice(i, i + 100)));
  return out;
}
writeFileSync(join(runDir, 'native.json'), `${JSON.stringify(classifyAll(nativeBinary), null, 2)}\n`);
// The settings.xml check runs before the content is read, so it masks every
// later rejection; this pass shows what each file would hit next.
writeFileSync(join(runDir, 'native-settings-unmasked.json'), `${JSON.stringify(classifyAll(unmaskedBinary), null, 2)}\n`);

// Which layout flags each `compatibility-setting` rejection carries, against
// the reader's LAYOUT_SETTINGS table in this checkout.
const layoutTable = /const LAYOUT_SETTINGS[^=]*=\s*\[([\s\S]*?)\];/u.exec(readFileSync(odtSource, 'utf8'));
if (!layoutTable) throw new Error('LAYOUT_SETTINGS not found in odt.rs');
const expected = new Map([...layoutTable[1].matchAll(/\(\s*"([^"]+)",\s*"([^"]*)",?\s*\)/gu)].map((m) => [m[1], m[2]]));
const native = JSON.parse(readFileSync(join(runDir, 'native.json'), 'utf8'));
const settingsFlags = {};
for (const file of files) {
  if (!native[file.id]?.startsWith('ODF_UNSUPPORTED:compatibility-setting')) continue;
  const xml = spawnSync('unzip', ['-p', file.odt, 'settings.xml'], { encoding: 'utf8', maxBuffer: 64 << 20 }).stdout;
  settingsFlags[file.id] = [...xml.matchAll(/<config:config-item config:name="([^"]+)"[^>]*>([^<]*)<\/config:config-item>/gu)]
    .filter(([, name, value]) => expected.has(name) && expected.get(name) !== value.trim())
    .map(([, name, value]) => `${name}=${value.trim()}`);
}
writeFileSync(join(runDir, 'settings-flags.json'), `${JSON.stringify(settingsFlags, null, 2)}\n`);

// Drop cached results whose ODT or PDF changed since they were written, so a
// failed retry cannot leave an old score standing (the spec does the same).
const resultsDir = join(runDir, 'results');
if (existsSync(resultsDir)) {
  const byId = new Map(files.map((file) => [file.id, file]));
  for (const name of readdirSync(resultsDir)) {
    const path = join(resultsDir, name);
    const file = byId.get(name.replace(/\.json$/u, ''));
    if (!file) continue;
    let stored;
    try { stored = JSON.parse(readFileSync(path, 'utf8')).inputs; } catch { /* torn */ }
    if (stored !== inputsHash(file)) unlinkSync(path);
  }
}

// ---- browser pass ---------------------------------------------------------------
if (!args['skip-browser']) {
  const config = join(checkoutRoot, 'packages/opendocument/tests/corpus/playwright.config.ts');
  const browser = spawnSync('pnpm', ['exec', 'playwright', 'test', '--config', config], {
    cwd: join(checkoutRoot, 'packages/opendocument'),
    stdio: 'inherit',
    env: {
      ...process.env,
      ODF_HARNESS_RUN_DIR: runDir,
      ODF_HARNESS_CHECKOUT: checkoutRoot,
      ODF_HARNESS_JOBS: args.jobs,
      ODF_HARNESS_PORT: args.port,
      ...(customWasm ? { ODF_HARNESS_WASM: customWasm } : {}),
    },
  });
  // Failing tests (timeouts, crashes) are recorded by their missing result;
  // the report counts them, so a nonzero exit does not stop it.
  if (browser.status !== 0) console.warn(`playwright exited with ${browser.status}`);
}

run('node', [join(checkoutRoot, 'packages/opendocument/tests/corpus/report.mjs'), runDir, '--worst', args.worst]);
