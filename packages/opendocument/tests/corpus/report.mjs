#!/usr/bin/env node
// Aggregates an ODT corpus harness run directory into summary.json and
// report.md, and copies the side-by-side images (LibreOffice | ours | diff) of
// the worst pages into <run>/worst/.
//
//   node packages/opendocument/tests/corpus/report.mjs <run dir> [--worst 40]
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { parseArgs } from 'node:util';

const { values: args, positionals } = parseArgs({
  allowPositionals: true,
  options: { worst: { type: 'string', default: '40' }, meta: { type: 'string', default: '{}' } },
});
const runDir = positionals[0];
if (!runDir) throw new Error('usage: report.mjs <run dir>');
const worstN = Number(args.worst);
const meta = JSON.parse(args.meta);

const files = JSON.parse(readFileSync(join(runDir, 'files.json'), 'utf8'));
const native = existsSync(join(runDir, 'native.json')) ? JSON.parse(readFileSync(join(runDir, 'native.json'), 'utf8')) : {};
const readJson = (name) => (existsSync(join(runDir, name)) ? JSON.parse(readFileSync(join(runDir, name), 'utf8')) : {});
const unmasked = readJson('native-settings-unmasked.json');
const settingsFlags = readJson('settings-flags.json');
const resultsDir = join(runDir, 'results');
const results = new Map();
if (existsSync(resultsDir)) {
  for (const name of readdirSync(resultsDir)) {
    const result = JSON.parse(readFileSync(join(resultsDir, name), 'utf8'));
    results.set(result.id, result);
  }
}

/** The reader's reason for a rejection, from the native classification. */
// In a --unmask-settings run the browser loaded the diagnostic build, so
// compare it with the unmasked native classification.
const browserNative = meta.unmaskSettings ? unmasked : native;
function nativeReason(id, table = browserNative) {
  const line = table[id];
  if (line === undefined) return 'unclassified';
  if (line.startsWith('admitted')) return 'admitted';
  const unsupported = /^ODF_UNSUPPORTED:([\w-]+)/u.exec(line);
  if (unsupported) return unsupported[1];
  if (line.startsWith('ODF_UNSUPPORTED_ENCRYPTION')) return 'encryption';
  if (line.startsWith('panic')) return 'native-panic';
  // Other failures (malformed input, limits): keep the leading token only.
  return `error:${line.split(/[\s:(]/u)[0]}`;
}

const status = { opened: 0, rejected: 0, error: 0, missing: 0 };
const reasons = new Map();
const disagreements = [];
const errors = [];
const pages = [];
let pageCountMatch = 0;
let pageCountMismatch = [];
for (const file of files) {
  const result = results.get(file.id);
  const reason = nativeReason(file.id);
  if (!result) {
    status.missing += 1;
    continue;
  }
  status[result.status] += 1;
  if (result.status === 'rejected') {
    const list = reasons.get(reason) ?? [];
    list.push(file.rel);
    reasons.set(reason, list);
  }
  if (result.status === 'error') errors.push({ rel: file.rel, message: result.message, native: reason });
  if ((result.status === 'opened') !== (reason === 'admitted')) disagreements.push({ rel: file.rel, browser: result.status, native: reason });
  if (result.status === 'opened') {
    if (result.pageCountMatch) pageCountMatch += 1;
    else pageCountMismatch.push({ rel: file.rel, ours: result.pageCount, libreoffice: result.pdfPageCount });
    for (const page of result.pages ?? []) {
      if (page.layoutDiffPct !== undefined) pages.push({ rel: file.rel, ...page });
    }
  }
}

const total = files.length;
const pct = (n, d = total) => (d === 0 ? '0.0' : (n / d * 100).toFixed(1));
const buckets = [0, 1, 2, 5, 10, 20, 40, 60, 80, 100.0001];
function histogram(key) {
  return buckets.slice(0, -1).map((low, i) => {
    const high = buckets[i + 1];
    const count = pages.filter((page) => page[key] >= low && page[key] < high).length;
    return { range: `${low}–${Math.min(high, 100).toFixed(0)}%`, count };
  });
}
function quantiles(key) {
  const sorted = pages.map((page) => page[key]).sort((a, b) => a - b);
  const at = (q) => (sorted.length ? sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))] : 0);
  return { p50: at(0.5), p90: at(0.9), p99: at(0.99), max: sorted.at(-1) ?? 0 };
}

const worst = [...pages].sort((a, b) => b.layoutDiffPct - a.layoutDiffPct).slice(0, worstN);
const worstDir = join(runDir, 'worst');
rmSync(worstDir, { recursive: true, force: true });
mkdirSync(worstDir, { recursive: true });
worst.forEach((page, index) => {
  if (page.image && existsSync(page.image)) {
    const target = join(worstDir, `${String(index + 1).padStart(3, '0')}-${basename(page.image)}`);
    copyFileSync(page.image, target);
    page.worstImage = target;
  }
});

const rankedReasons = [...reasons.entries()].sort((a, b) => b[1].length - a[1].length)
  .map(([reason, list]) => ({ reason, files: list.length, examples: list.slice(0, 3) }));
const rank = (map) => [...map.entries()].sort((a, b) => b[1].length - a[1].length)
  .map(([key, list]) => ({ key, files: list.length, examples: list.slice(0, 3) }));
const nextReasons = new Map();
const flagCounts = new Map();
for (const file of files) {
  if (nativeReason(file.id, native) !== 'compatibility-setting') continue;
  const next = nativeReason(file.id, unmasked);
  nextReasons.set(next, [...(nextReasons.get(next) ?? []), file.rel]);
  for (const flag of settingsFlags[file.id] ?? []) flagCounts.set(flag, [...(flagCounts.get(flag) ?? []), file.rel]);
}
const unmaskedRanking = rank(nextReasons);
const flagRanking = rank(flagCounts);

const summary = {
  meta,
  generatedAt: new Date().toISOString(),
  total,
  status,
  acceptancePct: Number(pct(status.opened)),
  rejectionReasons: rankedReasons,
  settingsUnmasked: { nextReasons: unmaskedRanking, layoutFlags: flagRanking },
  errors,
  disagreements,
  pageCount: { match: pageCountMatch, mismatch: pageCountMismatch },
  pages: {
    compared: pages.length,
    layoutDiffPct: { ...quantiles('layoutDiffPct'), histogram: histogram('layoutDiffPct') },
    inkDiffPct: { ...quantiles('inkDiffPct'), histogram: histogram('inkDiffPct') },
    diffPct: { ...quantiles('diffPct'), histogram: histogram('diffPct') },
  },
  worst,
};
writeFileSync(join(runDir, 'summary.json'), `${JSON.stringify(summary, null, 2)}\n`);

const md = [];
md.push(`# ODT corpus harness — ${meta.label ?? basename(runDir)}`, '');
md.push(`Checkout ${meta.sha ?? '?'}${meta.dirty ? ' (dirty)' : ''}, compatibility rules ${meta.rules ?? '?'}. ${summary.generatedAt}`, '');
if (meta.unmaskSettings) md.push('**DIAGNOSTIC: rendered with the settings.xml layout-flag check disabled; acceptance below is not the shipped reader\'s.**', '');
md.push(`- Files: ${total}`);
md.push(`- Opened: ${status.opened} (${pct(status.opened)}%)`);
md.push(`- Rejected: ${status.rejected} (${pct(status.rejected)}%)`);
md.push(`- Load errors (not a subset rejection): ${status.error}`);
md.push(`- No result (timeout or crash): ${status.missing}`);
md.push(`- Browser/native disagreements: ${disagreements.length}`);
md.push(`- Page count matches LibreOffice: ${pageCountMatch}/${status.opened}`, '');
md.push('## Rejection reasons (reader code, by file count)', '', '| Code | Files | % of corpus | Example |', '| --- | ---: | ---: | --- |');
for (const { reason, files: count, examples } of rankedReasons) md.push(`| \`${reason}\` | ${count} | ${pct(count)} | ${examples[0]} |`);
if (unmaskedRanking.length) {
  const masked = nextReasons.size ? [...nextReasons.values()].reduce((n, l) => n + l.length, 0) : 0;
  md.push('', `## Behind \`compatibility-setting\` (${masked} files, diagnostic build without the settings.xml check)`, '');
  md.push('First rejection each file hits once its settings.xml layout flags are ignored; `admitted` would open.', '');
  md.push('| Next code | Files | Example |', '| --- | ---: | --- |');
  for (const { key, files: count, examples } of unmaskedRanking) md.push(`| \`${key}\` | ${count} | ${examples[0]} |`);
  md.push('', 'Layout flags that differ from the reader\'s accepted values (files carrying each):', '');
  md.push('| Flag | Files |', '| --- | ---: |');
  for (const { key, files: count } of flagRanking) md.push(`| \`${key}\` | ${count} |`);
}
md.push('', `## Page diff distribution (${pages.length} pages compared)`, '');
md.push('Layout diff (primary) = inked pixels, either side, with no ink on the other side within 2 px, / inked pixels on either side: tolerant of antialiasing and 1–2 px offsets. Ink diff = pixelmatch-differing pixels (threshold 0.2, AA ignored) / inked pixels on either side. Page diff = pixelmatch-differing pixels / page area.', '');
const layout = summary.pages.layoutDiffPct;
const ink = summary.pages.inkDiffPct;
const area = summary.pages.diffPct;
md.push(`| Metric | p50 | p90 | p99 | max |`, '| --- | ---: | ---: | ---: | ---: |');
md.push(`| Layout diff % | ${layout.p50.toFixed(1)} | ${layout.p90.toFixed(1)} | ${layout.p99.toFixed(1)} | ${layout.max.toFixed(1)} |`);
md.push(`| Ink diff % | ${ink.p50.toFixed(1)} | ${ink.p90.toFixed(1)} | ${ink.p99.toFixed(1)} | ${ink.max.toFixed(1)} |`);
md.push(`| Page diff % | ${area.p50.toFixed(2)} | ${area.p90.toFixed(2)} | ${area.p99.toFixed(2)} | ${area.max.toFixed(2)} |`, '');
md.push('| Range | Pages (layout diff) | Pages (ink diff) |', '| --- | ---: | ---: |');
layout.histogram.forEach(({ range, count }, i) => md.push(`| ${range} | ${count} | ${ink.histogram[i].count} |`));
md.push('', `## Worst ${worst.length} pages`, '', '| # | File | Page | Layout diff % | Ink diff % | Page diff % | Image |', '| ---: | --- | ---: | ---: | ---: | ---: | --- |');
worst.forEach((page, i) => md.push(`| ${i + 1} | ${page.rel} | ${page.page} | ${page.layoutDiffPct.toFixed(1)} | ${page.inkDiffPct.toFixed(1)} | ${page.diffPct.toFixed(2)} | ${page.worstImage ? basename(page.worstImage) : '—'} |`));
if (pageCountMismatch.length) {
  md.push('', '## Page count mismatches', '', '| File | Ours | LibreOffice |', '| --- | ---: | ---: |');
  for (const m of pageCountMismatch) md.push(`| ${m.rel} | ${m.ours} | ${m.libreoffice} |`);
}
if (errors.length) {
  md.push('', '## Load errors', '');
  for (const e of errors) md.push(`- ${e.rel}: \`${String(e.message).slice(0, 200)}\` (native: ${e.native})`);
}
if (disagreements.length) {
  md.push('', '## Browser/native disagreements', '');
  for (const d of disagreements) md.push(`- ${d.rel}: browser ${d.browser}, native ${d.native}`);
}
writeFileSync(join(runDir, 'report.md'), `${md.join('\n')}\n`);
console.log(md.slice(0, 12).join('\n'));
console.log(`\nReport: ${join(runDir, 'report.md')}`);
