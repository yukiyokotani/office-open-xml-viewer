#!/usr/bin/env node
// Explicit local-only Office oracle. Not part of the browser/library runtime.
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createHash } from 'node:crypto';
import { copyFile, mkdir, mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import init, { convert_legacy_office } from '../packages/legacy-converter/src/wasm/legacy_office_converter.js';

const exec = promisify(execFile);
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const families = { doc: ['docx', 'com.microsoft.Word'], xls: ['xlsx', 'com.microsoft.Excel'], ppt: ['pptx', 'com.microsoft.Powerpoint'] };
const options = new Map(process.argv.slice(2).map(arg => {
  const match = /^--(format|limit|python)=(.+)$/.exec(arg);
  if (!match) throw new Error('usage: node scripts/legacy-office-fidelity.mjs [--format=doc|xls|ppt] [--limit=N] [--python=PATH]');
  return [match[1], match[2]];
}));
if (process.platform !== 'darwin') throw new Error('The Office oracle requires locally installed Office for macOS.');
const selected = options.has('format') ? [options.get('format')] : Object.keys(families);
if (selected.some(f => !families[f])) throw new Error('Unknown Office format');
const limit = Number(options.get('limit') ?? Number.MAX_SAFE_INTEGER);
if (!Number.isSafeInteger(limit) || limit <= 0) throw new Error('limit must be a positive integer');
const python = options.get('python') ?? 'python3';
await exec(python, ['-c', 'import PIL, pypdf']);
await exec('pdftoppm', ['-v']);
const runDirectory = await mkdtemp(join(tmpdir(), 'legacy-office-fidelity-'));
const wasm = await readFile(join(root, 'packages/legacy-converter/src/wasm/legacy_office_converter_bg.wasm'));
await init({ module_or_path: wasm });
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const { stdout: revision } = await exec('git', ['rev-parse', 'HEAD'], { cwd: root });
const { stdout: patch } = await exec('git', ['diff', 'HEAD'], { cwd: root, maxBuffer: 16 * 1024 * 1024 });
const report = { oracle: 'local-office-pdf', comparison: 'exact-raster-at-96-dpi', revision: revision.trim(), trackedPatchSha256: hash(patch), converterWasmSha256: hash(wasm), cases: [] };
console.log(`Local report directory: ${runDirectory}`);
for (const family of selected) {
  const [modern, container] = families[family];
  const sourceDirectory = join(root, `packages/${modern}/public/private`);
  const names = (await readdir(sourceDirectory)).filter(n => !n.startsWith('~$') && n.toLowerCase().endsWith(`.${family}`)).sort().slice(0, limit);
  if (!names.length) throw new Error(`No local ${family} corpus found`);
  // Office's sandbox can access its own private container without granting
  // access to the original corpus. Only disposable copies are opened.
  const staging = await mkdtemp(join(homedir(), 'Library', 'Containers', container, 'Data', 'tmp', 'legacy-fidelity-'));
  for (const [index, name] of names.entries()) {
    const id = `${family}-${String(index + 1).padStart(4, '0')}`;
    const directory = join(runDirectory, id);
    await mkdir(directory);
    const source = await readFile(join(sourceDirectory, name));
    const entry = { id, sourceSha256: hash(source), status: 'unverified' };
    report.cases.push(entry);
    let output;
    try {
      output = convert_legacy_office(source, family, 256 * 1024 * 1024);
      const bytes = output.take_bytes();
      entry.warnings = output.warnings().split('\n').filter(Boolean);
      entry.outputSha256 = hash(bytes);
      const original = join(staging, `${id}-source.${family}`);
      const converted = join(staging, `${id}-converted.${modern}`);
      await writeFile(original, source, { flag: 'wx' });
      await writeFile(converted, bytes, { flag: 'wx' });
      for (const [label, file] of [['source', original], ['converted', converted]]) {
        const pdf = join(staging, `${id}-${label}.pdf`);
        await exec('osascript', [join(root, 'scripts/legacy-office-export.applescript'), family, file, pdf], { timeout: 270_000, maxBuffer: 1024 * 1024 });
        await copyFile(pdf, join(directory, `${label}.pdf`));
      }
      const { stdout } = await exec(python, [join(root, 'scripts/legacy-office-compare.py'), directory], { timeout: 270_000, maxBuffer: 8 * 1024 * 1024 });
      entry.comparison = JSON.parse(stdout);
      entry.status = entry.comparison.equal ? 'equal' : 'different';
    } catch (error) {
      entry.status = 'error';
      // Diagnostic details are local-only; never copy this report into public docs.
      entry.error = String(error);
    } finally {
      output?.free();
      await writeFile(join(runDirectory, 'report.json'), JSON.stringify(report, null, 2));
    }
    console.log(`${id}: ${entry.status}${entry.comparison ? ` (${entry.comparison.sourcePages}/${entry.comparison.convertedPages} pages)` : ''}`);
    // Stop after an Office/export failure: do not accumulate dialogs or continue
    // sending events to an application whose cleanup state is uncertain.
    if (entry.status === 'error') { process.exitCode = 1; break; }
  }
  if (process.exitCode) break;
}
const counts = report.cases.reduce((acc, entry) => ({ ...acc, [entry.status]: (acc[entry.status] ?? 0) + 1 }), {});
console.log(JSON.stringify(counts));
if (report.cases.some(entry => entry.status !== 'equal')) process.exitCode = 1;
