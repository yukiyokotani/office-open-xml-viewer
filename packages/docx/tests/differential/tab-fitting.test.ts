/// <reference types="node" />
/// <reference path="./baseline-modules.d.ts" />
import { execFileSync } from 'node:child_process';
import { readFile, writeFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from '../../src/wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from '../../src/parser-model.js';
import { layoutDocument } from '../../src/document-layout.js';
import { layoutDocument as baselineLayout } from '@docx-tab-baseline/document-layout';
import { createLayoutServices as baselineServices } from '@docx-tab-baseline/layout-runtime';
import { createLayoutServices } from '../../src/layout-runtime.js';
import type { ParagraphLayout } from '../../src/layout/types.js';
import { documentBytes, measureContext, matrix, fittingMatrix } from '../../src/test-support/tab-fitting.test-support.js';

// Run with DOCX_TAB_BASELINE_CHECKOUT pointing to a clean detached origin/main
// checkout: vitest run --config packages/docx/tests/differential/tab-fitting.config.ts.
// Compare observable line partitions and placement geometry, never private
// breaker state. The two layout graphs use the same parsed facts and metrics.
// Install dependencies and build WASM in both checkouts first. This gate is not
// part of `pnpm test` because it needs the second checkout. Set
// DOCX_TAB_DIFFERENTIAL_REPORT to a path outside the checkout to save counts.
beforeAll(async () => {
  const base = process.env.DOCX_TAB_BASELINE_CHECKOUT;
  if (!base) throw new Error('DOCX_TAB_BASELINE_CHECKOUT must name a clean origin/main checkout');
  const gitEnv = Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key)));
  const git = (cwd: string, args: string[]) => execFileSync('git', ['-C', cwd, ...args], { encoding: 'utf8', env: gitEnv }).trim();
  expect(git(base, ['rev-parse', 'HEAD'])).toBe(git(process.cwd(), ['rev-parse', 'origin/main']));
  expect(git(base, ['status', '--porcelain', '--untracked-files=no'])).toBe('');
  await init({ module_or_path: await readFile(new URL('../../src/wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function compare(bytes: Uint8Array) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const paragraph = (layout: ReturnType<typeof layoutDocument>) => {
    const node = layout.pages[0]?.layers.body.find((item) => item.kind === 'paragraph');
    if (node?.kind !== 'paragraph') throw new Error('Missing paragraph');
    return node;
  };
  const candidate = paragraph(layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 }));
  const baseline = paragraph(baselineLayout(model, baselineServices(model, { measureContext: measureContext() }), { currentDateMs: 0 }));
  return { candidate, baseline };
}

function outOfBand(paragraph: ParagraphLayout, indent: string, rtl: boolean) {
  const twips = (name: string) => Number(new RegExp(`w:${name}="(-?\\d+)"`).exec(indent)?.[1] ?? 0) / 20;
  return paragraph.lines.some((line, index) => line.placements.some((node) => {
    if (!node.bounds || (node.kind === 'text' && !node.text.trim())) return false;
    // §17.3.1.12: an authored hanging value (including zero) wins over firstLine.
    const first = index === 0 ? /w:hanging=/.test(indent) ? -twips('hanging') : twips('firstLine') : 0;
    const start = 72 + (rtl ? twips('right') : twips('left') + first);
    const end = 540 - (rtl ? twips('left') + first : twips('right'));
    const trailing = node.kind === 'text' ? (node.text.length - node.text.trimEnd().length) * 5 : 0;
    return node.bounds.xPt < start - 1e-9 || node.bounds.xPt + node.bounds.widthPt - trailing > end + 1e-9;
  }));
}

const geometry = (paragraph: ParagraphLayout) => paragraph.lines.map((line) => ({
  bounds: line.bounds,
  placements: line.placements.map((node) => ({ kind: node.kind, bounds: node.bounds, ...('text' in node ? { text: node.text } : {}) })),
}));

const matrixCases = [...matrix, ...fittingMatrix].filter((entry) => entry.float === 'none').map((entry) => ({
  indent: 'w:left="720" w:hanging="720"', rtl: entry.rtl,
  alignment: entry.alignment, count: entry.count,
  text: 'content' in entry ? entry.content : entry.alignment === 'decimal' ? '12.3' : 'word',
  positional: 'relativeTo' in entry, relativeTo: 'relativeTo' in entry ? entry.relativeTo : undefined,
  automatic: entry.kind === 'automatic',
}));
const generatedCases = ['', 'w:left="720" w:hanging="720"', 'w:left="720" w:right="1440"'].flatMap((indent) =>
  [false, true].flatMap((rtl) => [1, 2, 3].flatMap((count) =>
    ['left', 'start', 'right', 'end', 'center', 'decimal', 'bar', 'clear', 'num'].flatMap((alignment) =>
      [720, 2400, 5600, 11000].flatMap((stop) => ['', 'prefix '].flatMap((prefix) =>
        ['word', 'word '.repeat(15).trim(), '漢'.repeat(50), 'ภาษาไทย'.repeat(10)].map((text) =>
          ({ indent, rtl, count, alignment, stop, prefix, text, positional: false, relativeTo: undefined, automatic: false }))))))));

// Cartesian input space includes omitted firstLine/hanging so precedence does
// not mask either axis. Negative firstLine/hanging are robustness inputs;
// §17.3.1.12 defines unsigned measures for those two XML attributes. Signed
// left/right values are normative. mirrorIndents is varied in the XML, but
// neither parser retains it: this checks regression parity only, not support
// for the inside/outside mapping required by §17.3.1.18.
function* indentCases() {
  for (const left of [-720, 0, 720]) for (const right of [-720, 0, 720])
    for (const firstLine of [undefined, -720, 0, 720]) for (const hanging of [undefined, -720, 0, 720]) {
      const indent = `w:left="${left}" w:right="${right}"`
        + (firstLine === undefined ? '' : ` w:firstLine="${firstLine}"`)
        + (hanging === undefined ? '' : ` w:hanging="${hanging}"`);
      for (const rtl of [false, true]) for (const mirrorIndents of [false, true]) {
        const first = hanging === undefined ? (firstLine ?? 0) : -hanging;
        // Stops are signed distances from the reading-leading text margin.
        const start = left + Math.min(0, first);
        const end = 9360 - right;
        for (const stop of [start - 720, (start + end) / 2, end + 720])
          for (const count of [1, 2, 3])
            for (const alignment of ['left', 'start', 'right', 'end', 'center', 'decimal', 'bar', 'clear', 'num'])
              for (const text of ['12.3', 'word '.repeat(15).trim(), '漢'.repeat(50), '漢'.repeat(90), 'ภาษาไทย'.repeat(10)]) {
                const cells = count === 1 ? ['title', text] : count === 2 ? ['aaa', text, 'b'] : ['A', 'B', text, 'tail'];
                yield { indent, rtl, mirrorIndents, alignment, count, stop, text, cells,
                  positional: false, relativeTo: undefined, automatic: false };
              }
      }
    }
}

// Keep the reviewer's entire interleaved matrix, including the two isolated
// negative-trailing-indent regressions and its multi-cell counterexamples.
function* reviewerCases() {
  for (const indent of ['', 'w:left="720" w:hanging="720"', 'w:left="720" w:right="1440"',
    'w:left="720" w:firstLine="720"', 'w:left="-720"', 'w:right="-720"'])
    for (const rtl of [false, true]) for (const alignment of ['left', 'right', 'center', 'decimal'])
      for (const stop of [720, 2400, 5600, 11000])
        for (const cells of [['title', '12.3'], ['A', 'B', '漢'.repeat(50)],
          ['word ', 'word ', 'word '.repeat(15).trim()], ['', 'abc', '漢'.repeat(80)],
          ['pref ', '12.3', 'tail'], ['aaa', '漢'.repeat(90), 'b'], ['', '漢'.repeat(100)],
          ['', 'word '.repeat(40).trim()], ['a', 'b', 'c', 'd', 'e']]) {
          yield { indent, rtl, alignment, stop, cells, count: cells.length - 1, text: cells.join(''),
            positional: false, relativeTo: undefined, automatic: false };
        }
}

it('preserves every in-band no-float matrix and generated tab layout from origin/main', async () => {
  let exceptions = 0;
  let identical = 0;
  let overflow = 0;
  let total = 0;
  let containmentViolations = 0;
  let contentViolations = 0;
  let unexplained = 0;
  const examples: unknown[] = [];
  const geometryKey = (paragraph: ParagraphLayout) => JSON.stringify(geometry(paragraph));
  const inventory = (text: string) => [...text.replace(/\s/g, '')].sort().join('');
  function* cases(): Generator<NonNullable<Parameters<typeof documentBytes>[3]> & { indent: string; rtl: boolean }> {
    yield* matrixCases; yield* generatedCases; yield* reviewerCases(); yield* indentCases();
  }
  for (const entry of cases()) {
    const { candidate, baseline } = compare(documentBytes(entry.indent, entry.rtl, 0, { ...entry, noFloat: true }));
    total += 1;
    // Exactly five short Word-backed RTL positional controls; long cells are not exempt.
    const exception = entry.positional && entry.rtl && entry.text === 'word'
      && (entry.alignment !== 'left' || entry.relativeTo === 'margin');
    const content = candidate.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text').map((node) => node.text).join('');
    const expected = entry.cells?.join('') ?? `${entry.prefix ?? ''}${entry.text}`;
    const badContent = inventory(content) !== inventory(expected);
    const badBand = outOfBand(candidate, entry.indent, entry.rtl);
    contentViolations += Number(badContent);
    containmentViolations += Number(badBand);
    const mainOverflow = outOfBand(baseline, entry.indent, entry.rtl);
    const different = geometryKey(candidate) !== geometryKey(baseline);
    const regression = different && !exception && !mainOverflow;
    unexplained += Number(regression);
    if ((badContent || badBand || regression) && examples.length < 100) examples.push({ entry, badContent, badBand, regression,
      baseline: geometry(baseline), candidate: geometry(candidate) });
    if (exception) exceptions += 1;
    else if (mainOverflow) overflow += 1;
    else identical += Number(!different);
    if (total % 10_000 === 0 && process.env.DOCX_TAB_DIFFERENTIAL_REPORT) {
      await writeFile(`${process.env.DOCX_TAB_DIFFERENTIAL_REPORT}.progress`, JSON.stringify({ total, unexplained, containmentViolations, contentViolations }));
    }
  }
  const report = { matrix: matrixCases.length, generated: total - matrixCases.length, total,
    identical, overflow, exceptions, unexplained, containmentViolations, contentViolations, examples };
  if (process.env.DOCX_TAB_DIFFERENTIAL_REPORT) await writeFile(process.env.DOCX_TAB_DIFFERENTIAL_REPORT, JSON.stringify(report, null, 2));
  console.log({ ...report, examples: undefined });
  expect(matrixCases).toHaveLength(178);
  expect(exceptions).toBe(5);
  expect(contentViolations, JSON.stringify(examples.slice(0, 2))).toBe(0);
  expect(containmentViolations, JSON.stringify(examples.slice(0, 2))).toBe(0);
  expect(unexplained, JSON.stringify(examples.slice(0, 2))).toBe(0);
}, 1_800_000);
