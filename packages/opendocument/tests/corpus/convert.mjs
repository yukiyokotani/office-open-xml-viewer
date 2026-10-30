#!/usr/bin/env node
// Builds the local-only OpenDocument validation corpus (#1934): every private
// DOCX/XLSX/PPTX sample is converted to ODT/ODS/ODP with LibreOffice headless,
// and each converted ODF file is exported to PDF by LibreOffice. That PDF is
// the ground truth the ODF harness compares against.
//
//   node packages/opendocument/tests/corpus/convert.mjs [--formats odt,ods,odp]
//     [--jobs 4] [--timeout 120] [--retry-failed] [--limit N]
//
// Inputs:  <repo>/packages/{docx,xlsx,pptx}/public/private/**/*.{docx,xlsx,pptx}
// Outputs: <repo>/packages/opendocument/public/private/<odt|ods|odp>/<relative
//          path>.<ext> plus <relative path>.pdf next to it (git-ignored).
// <repo> is the main checkout (the private corpus lives there, not in
// worktrees); override with ODF_CORPUS_REPO.
//
// Resumable: an existing non-empty output is never redone. Failures are
// appended to convert-failures.local.jsonl and skipped on later runs unless
// --retry-failed is given.
//
// Headless LibreOffice on macOS enumerates fonts through fontconfig; without a
// config it sees no system CJK fonts and paints CJK as blank glyphs. The
// Homebrew fontconfig configuration scans the system font folders.
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, appendFileSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, relative, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { parseArgs } from 'node:util';
import { corpusRepo } from './paths.mjs';

const SOFFICE = process.env.SOFFICE ?? '/Applications/LibreOffice.app/Contents/MacOS/soffice';
const FONTCONFIG_FILE = process.env.FONTCONFIG_FILE ?? '/opt/homebrew/etc/fonts/fonts.conf';

const FORMATS = {
  odt: { source: 'docx', ext: 'docx' },
  ods: { source: 'xlsx', ext: 'xlsx' },
  odp: { source: 'pptx', ext: 'pptx' },
};

const { values: args } = parseArgs({
  options: {
    formats: { type: 'string', default: 'odt,ods,odp' },
    jobs: { type: 'string', default: '4' },
    timeout: { type: 'string', default: '120' },
    'retry-failed': { type: 'boolean', default: false },
    limit: { type: 'string' },
  },
});

const repo = corpusRepo();
const outRoot = join(repo, 'packages/opendocument/public/private');
const failuresPath = join(outRoot, 'convert-failures.local.jsonl');
const timeoutMs = Number(args.timeout) * 1000;
const jobs = Math.max(1, Number(args.jobs));
if (!existsSync(SOFFICE)) throw new Error(`LibreOffice not found: ${SOFFICE}`);
if (!existsSync(FONTCONFIG_FILE)) throw new Error(`fontconfig configuration not found: ${FONTCONFIG_FILE} (brew install fontconfig)`);

function walk(dir, ext, out = []) {
  if (!existsSync(dir)) return out;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) walk(path, ext, out);
    else if (entry.isFile() && entry.name.toLowerCase().endsWith(`.${ext}`) && !entry.name.startsWith('~$')) out.push(path);
  }
  return out;
}

function nonEmpty(path) {
  return existsSync(path) && statSync(path).size > 0;
}

function knownFailures() {
  const failed = new Set();
  if (args['retry-failed'] || !existsSync(failuresPath)) return failed;
  for (const line of readFileSync(failuresPath, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    try { failed.add(JSON.parse(line).output); } catch { /* torn line */ }
  }
  return failed;
}

/** Run soffice once, converting `input` into `outDir` with `filter`. */
function soffice(profile, filter, input, outDir) {
  return new Promise((resolvePromise) => {
    const child = spawn(SOFFICE, [
      `-env:UserInstallation=file://${profile}`,
      '--headless', '--norestore', '--nolockcheck', '--nodefault',
      '--convert-to', filter, '--outdir', outDir, input,
    ], {
      env: { ...process.env, FONTCONFIG_FILE },
      stdio: ['ignore', 'pipe', 'pipe'],
      detached: true,
    });
    let output = '';
    child.stdout.on('data', (chunk) => { output += chunk; });
    child.stderr.on('data', (chunk) => { output += chunk; });
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      try { process.kill(-child.pid, 'SIGKILL'); } catch { /* already gone */ }
    }, timeoutMs);
    child.on('close', (code) => {
      clearTimeout(timer);
      resolvePromise({ code, timedOut, output: output.slice(-2000) });
    });
  });
}

/**
 * Convert into a per-worker staging directory and move the result into place,
 * so an interrupted run never leaves a truncated output that looks complete.
 */
async function convertOne(profile, staging, filter, input, output) {
  rmSync(staging, { recursive: true, force: true });
  mkdirSync(staging, { recursive: true });
  const result = await soffice(profile, filter, input, staging);
  const produced = join(staging, basename(output));
  // Only a clean exit counts: a timed-out or failed export can leave a
  // partial file that would later look complete.
  if (result.timedOut || result.code !== 0 || !nonEmpty(produced)) {
    return { ok: false, reason: result.timedOut ? 'timeout' : `exit ${result.code}`, detail: result.output.trim() };
  }
  mkdirSync(dirname(output), { recursive: true });
  renameSync(produced, output);
  return { ok: true };
}

function recordFailure(entry) {
  appendFileSync(failuresPath, `${JSON.stringify({ ...entry, at: new Date().toISOString() })}\n`);
}

const tasks = [];
const skippedFailures = knownFailures();
for (const format of args.formats.split(',').map((f) => f.trim()).filter(Boolean)) {
  const spec = FORMATS[format];
  if (!spec) throw new Error(`Unknown format ${format}`);
  const sourceRoot = join(repo, `packages/${spec.source}/public/private`);
  for (const input of walk(sourceRoot, spec.ext).sort()) {
    const rel = relative(sourceRoot, input).replace(/\.[^.]+$/u, '');
    const odf = join(outRoot, format, `${rel}.${format}`);
    const pdf = join(outRoot, format, `${rel}.pdf`);
    tasks.push({ format, input, odf, pdf });
  }
}
const limited = args.limit ? tasks.slice(0, Number(args.limit)) : tasks;
mkdirSync(outRoot, { recursive: true });

const counts = {};
const bump = (format, key) => {
  counts[format] ??= { total: 0, done: 0, converted: 0, odfFailed: 0, pdfFailed: 0, skippedFailed: 0 };
  counts[format][key] += 1;
};
for (const task of limited) bump(task.format, 'total');

let cursor = 0;
let finished = 0;
const started = Date.now();
async function worker(index) {
  const base = join(tmpdir(), `odf-corpus-${process.pid}-${index}`);
  const profile = join(base, 'profile');
  // Staged next to the outputs (same filesystem) so promotion is a rename.
  const staging = join(outRoot, `.staging-${process.pid}-${index}`);
  mkdirSync(profile, { recursive: true });
  try {
    while (cursor < limited.length) {
      const task = limited[cursor++];
      const { format, input, odf, pdf } = task;
      if (nonEmpty(odf) && nonEmpty(pdf)) {
        bump(format, 'done');
      } else if (skippedFailures.has(odf) || skippedFailures.has(pdf)) {
        bump(format, 'skippedFailed');
      } else {
        let ok = true;
        if (!nonEmpty(odf)) {
          const result = await convertOne(profile, staging, format, input, odf);
          if (!result.ok) {
            ok = false;
            bump(format, 'odfFailed');
            recordFailure({ stage: 'odf', format, input, output: odf, reason: result.reason, detail: result.detail });
          }
        }
        if (ok) {
          const result = await convertOne(profile, staging, 'pdf', odf, pdf);
          if (result.ok) bump(format, 'converted');
          else {
            bump(format, 'pdfFailed');
            recordFailure({ stage: 'pdf', format, input: odf, output: pdf, reason: result.reason, detail: result.detail });
          }
        }
      }
      finished += 1;
      if (finished % 25 === 0 || finished === limited.length) {
        const rate = finished / ((Date.now() - started) / 1000);
        console.log(`[${finished}/${limited.length}] ${rate.toFixed(2)} files/s ${JSON.stringify(counts)}`);
      }
    }
  } finally {
    rmSync(base, { recursive: true, force: true });
    rmSync(staging, { recursive: true, force: true });
  }
}

await Promise.all(Array.from({ length: jobs }, (_, index) => worker(index)));
writeFileSync(join(outRoot, 'convert-summary.local.json'), `${JSON.stringify({ at: new Date().toISOString(), counts }, null, 2)}\n`);
console.log(JSON.stringify(counts, null, 2));
