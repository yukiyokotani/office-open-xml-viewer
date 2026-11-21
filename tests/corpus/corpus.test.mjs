// The public specimen corpus: every committed file is the reproducible build
// of its case module, is a structurally valid package, and contains every XML
// feature its case declares in `covers`.
//
// Run: node --test tests/corpus/corpus.test.mjs   (part of `pnpm test`)
// Fix a stale file: node tests/corpus/build.mjs

import assert from 'node:assert/strict';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import test from 'node:test';
import { FORMATS, buildCase, fileName, loadCases, outputDir, outputPath } from './lib/cases.mjs';
import { canonicalFeature, packageFeatures } from './lib/features.mjs';
import { checkPackage, parseXmlStrict } from './lib/package-check.mjs';
import { VARIANTS, isReferencePdf, needsReference, pdfName, readManifest, sha256 } from './lib/references.mjs';
import { readZip } from '../visual/corpus-features.mjs';

const cases = await loadCases();

test('case ids are unique and every format has cases', () => {
  const ids = cases.map((c) => c.id);
  assert.deepEqual(ids, [...new Set(ids)]);
  for (const format of FORMATS) assert.ok(cases.some((c) => c.format === format), `no ${format} case`);
});

test('the corpus directories hold exactly the built cases', () => {
  for (const format of FORMATS) {
    const dir = outputDir(format);
    const committed = existsSync(dir) ? readdirSync(dir).filter((f) => !f.startsWith('.') && !isReferencePdf(f)).sort() : [];
    const expected = cases.filter((c) => c.format === format).map(fileName).sort();
    assert.deepEqual(committed, expected, `packages/${format}/public/corpus`);
  }
});

for (const c of cases) {
  test(`${fileName(c)}`, async (t) => {
    const built = buildCase(c);

    await t.test('rebuilds byte for byte', () => {
      assert.ok(existsSync(outputPath(c)), 'not built; run node tests/corpus/build.mjs');
      assert.ok(readFileSync(outputPath(c)).equals(built), 'stale; run node tests/corpus/build.mjs');
      assert.ok(buildCase(c).equals(built), 'the build is not deterministic');
    });

    await t.test('is a well-formed package', () => {
      checkPackage(built);
    });

    await t.test('contains every declared feature', () => {
      assert.ok(c.covers.length > 0, 'covers is empty');
      const present = packageFeatures(c.format, built);
      const missing = c.covers.filter((f) => !present.has(canonicalFeature(f)));
      assert.deepEqual(missing, []);
    });
  });
}

// Office reference PDFs: each present PDF belongs to a specimen that needs
// one and was exported from the specimen as committed (references.json).
// Missing PDFs are listed by tests/corpus/references.mjs, not failed here.
test('Office reference PDFs are recorded and current', () => {
  const manifest = readManifest();
  const expected = new Map();
  for (const c of cases) {
    const file = `${c.id}-${c.slug}.${c.format}`;
    const source = readFileSync(outputPath(c));
    if (!needsReference(c.format, source, readZip)) continue;
    for (const variant of VARIANTS) expected.set(pdfName(file, variant), { format: c.format, source: sha256(source) });
  }
  const problems = [];
  for (const format of FORMATS) {
    const dir = outputDir(format);
    if (!existsSync(dir)) continue;
    for (const name of readdirSync(dir).filter(isReferencePdf)) {
      const want = expected.get(name);
      const entry = manifest[name];
      if (!want || want.format !== format) problems.push(`${name}: no specimen needs this PDF`);
      else if (!entry) problems.push(`${name}: not in references.json; run node tests/corpus/references.mjs --record`);
      else if (entry.source !== want.source) problems.push(`${name}: stale, the specimen changed; export it again`);
      else if (entry.pdf !== sha256(readFileSync(`${dir}/${name}`))) problems.push(`${name}: the PDF differs from references.json`);
    }
  }
  for (const name of Object.keys(manifest)) {
    const want = expected.get(name);
    if (!want || !existsSync(`${outputDir(want.format)}/${name}`)) problems.push(`${name}: in references.json but no PDF`);
  }
  assert.deepEqual(problems, []);
});

// The checker itself must reject what it exists to catch.
test('the package checker rejects malformed XML', () => {
  for (const bad of ['<a><b></a>', '<a x="1" x="2"/>', '<p:a/>', '<a>&nbsp;</a>', '<a/><b/>', '<a b="1"c="2"/>']) {
    assert.throws(() => parseXmlStrict(bad), undefined, bad);
  }
  assert.doesNotThrow(() => parseXmlStrict('<?xml version="1.0"?><a xmlns:p="u"><p:b c="&amp;"/></a>'));
});

// Grid discipline: every length that places text or an object comes from
// lib/grid.mjs. Checked mechanically so a new case cannot slip in an ad-hoc
// offset. Intentional exceptions are listed with their reason.
// DrawingML coordinates are integers (ST_Coordinate, ST_PositionOffset);
// Word refuses to open a document with a fractional one, so a length
// divided off the grid must be rounded before it is written.
test('DrawingML offsets and extents are integers', () => {
  const INT = /^-?\d+$/;
  const problems = [];
  for (const c of cases) {
    const zip = readZip(buildCase(c));
    for (const name of zip.names.filter((n) => n.endsWith('.xml'))) {
      const xml = zip.read(name).toString('utf8');
      for (const [, v] of xml.matchAll(/<wp:posOffset>([^<]*)</g)) if (!INT.test(v)) problems.push(`${c.id} ${name}: posOffset ${v}`);
      for (const [, tag, attrs] of xml.matchAll(/<(wp:extent|a:off|a:ext|a:chOff|a:chExt)\s([^>]*)>/g)) {
        for (const [, k, v] of attrs.matchAll(/\b(cx|cy|x|y)="([^"]*)"/g)) if (!INT.test(v)) problems.push(`${c.id} ${name}: ${tag}@${k} ${v}`);
      }
    }
  }
  assert.deepEqual(problems, []);
});

const WORD_CHARS_SENTINEL = 31680;

test('positions and lengths lie on the grid', async () => {
  const { DOCX, OUTLINED_INSET, PPTX, XLSX } = await import('./lib/grid.mjs');
  const { readZip } = await import('../visual/corpus-features.mjs');
  const edgesW = new Set([0, DOCX.LIVE_W, DOCX.GUTTER, DOCX.gutterAfter(1), DOCX.gutterAfter(2), DOCX.gutterAfter(3)]);
  for (let a = 1; a <= 4; a += 1) {
    for (const v of [DOCX.colStart(a), DOCX.colEnd(a), DOCX.LIVE_W - DOCX.colEnd(a), DOCX.LIVE_W - DOCX.colStart(a)]) edgesW.add(v);
    for (let b = a; b <= 4; b += 1) { edgesW.add(DOCX.span(a, b)); edgesW.add(DOCX.colStart(b) - DOCX.colStart(a)); }
  }
  const xs = new Set();
  const ys = new Set();
  for (let c = 1; c <= PPTX.COLUMNS; c += 1) { xs.add(PPTX.x(c)); xs.add(PPTX.x(c) + PPTX.COL); }
  for (let y = PPTX.MARGIN; y <= PPTX.SLIDE_H; y += PPTX.LINE) ys.add(y);
  const widths = new Set([PPTX.STEP * 3]);
  for (let a = 1; a <= PPTX.COLUMNS; a += 1) for (let b = a; b <= PPTX.COLUMNS; b += 1) widths.add(PPTX.w(a, b));

  const problems = [];
  let checked = 0;
  for (const c of cases) {
    const zip = readZip(buildCase(c));
    for (const name of zip.names.filter((n) => n.endsWith('.xml'))) {
      const elements = parseXmlStrict(zip.read(name).toString('utf8'));
      const bad = (what) => problems.push(`${c.id} ${name}: ${what}`);
      if (c.format === 'docx') {
        for (const el of elements) {
          const a = (k) => (el.attrs.has(k) ? Number(el.attrs.get(k)) : undefined);
          // docx-050 tests adjacent sections at one exact page geometry
          // (margins 851, column space 425) that is not the corpus grid.
          const offGrid = c.id === 'docx-050';
          if (el.name === 'w:pgMar' && !offGrid) {
            const m = DOCX.MARGIN;
            if (a('w:top') !== m.top || a('w:bottom') !== m.bottom || a('w:left') !== m.left || a('w:right') !== m.right) bad('page margins');
            if (![0, DOCX.HEADER_DIST].includes(a('w:header')) || ![0, DOCX.FOOTER_DIST].includes(a('w:footer'))) bad('header/footer distance');
          }
          if (el.name === 'w:tab' && el.attrs.has('w:pos') && !edgesW.has(a('w:pos'))) bad(`tab stop ${a('w:pos')}`);
          if (el.name === 'w:ind') {
            for (const k of ['w:left', 'w:right', 'w:hanging', 'w:firstLine']) {
              // 31680 twips (22 in) is the value Word writes beside *Chars
              // indents; Word ignores it, and docx-039 tests exactly that.
              if (a(k) === WORD_CHARS_SENTINEL && c.id === 'docx-039') continue;
              if (el.attrs.has(k) && !edgesW.has(a(k))) bad(`indent ${k} ${a(k)}`);
            }
          }
          if (el.name === 'w:spacing' && !name.endsWith('styles.xml')) {
            for (const k of ['w:before', 'w:after']) if (el.attrs.has(k) && a(k) % DOCX.LINE) bad(`spacing ${k} ${a(k)}`);
          }
          if ((el.name === 'w:cols' || el.name === 'w:col') && el.attrs.has('w:space') && !edgesW.has(a('w:space')) && !offGrid) bad('column space');
          if (el.name === 'w:col' && !edgesW.has(a('w:w'))) bad(`column width ${a('w:w')}`);
          if (el.name === 'w:gridCol' && !edgesW.has(a('w:w'))) bad(`table column ${a('w:w')}`);
        }
      }
      if (/slides\/slide|slidelayouts|slidemasters|drawings\/drawing/.test(name)) {
        // Text insets: zero, or one gutter when the shape's outline is drawn.
        let outlined = false;
        let lineDepth = -1;
        for (const el of elements) {
          if (el.name === 'p:sp' || el.name === 'xdr:sp') { outlined = false; lineDepth = -1; }
          if (el.name === 'a:ln') lineDepth = el.depth;
          else if (el.depth <= lineDepth) lineDepth = -1;
          if (lineDepth >= 0 && el.depth === lineDepth + 1 && el.name !== 'a:noFill' && /Fill$/.test(el.name)) outlined = true;
          if (el.name === 'a:bodyPr' && el.attrs.size) {
            const expected = String(outlined ? OUTLINED_INSET : 0);
            if (['lIns', 'tIns', 'rIns', 'bIns'].some((k) => el.attrs.get(k) !== expected)) {
              bad(`text body insets (${outlined ? 'outlined' : 'unoutlined'} shape)`);
            }
          }
        }
      }
      if (c.format === 'pptx' && /slides\/slide|slidelayouts|slidemasters/.test(name)) {
        // Children of a group use the group's child coordinate space;
        // connectors follow the shapes they join (their ends sit mid-edge).
        let skipBelow = Infinity;
        for (const el of elements) {
          if (el.depth <= skipBelow) skipBelow = Infinity;
          if (el.name === 'p:grpSp' || el.name === 'p:cxnSp') skipBelow = el.depth;
          if (skipBelow !== Infinity) continue;
          if (el.name === 'a:off') checked += 1;
          if (el.name === 'a:off' && (!xs.has(Number(el.attrs.get('x'))) || !ys.has(Number(el.attrs.get('y'))))) bad(`offset ${el.attrs.get('x')},${el.attrs.get('y')}`);
          if (el.name === 'a:ext' && el.attrs.has('cx') && Number(el.attrs.get('cx')) && !widths.has(Number(el.attrs.get('cx')))) bad(`extent cx ${el.attrs.get('cx')}`);
          if (el.name === 'a:ext' && el.attrs.has('cy') && Number(el.attrs.get('cy')) % PPTX.LINE) bad(`extent cy ${el.attrs.get('cy')}`);
        }
      }
      if (c.format === 'xlsx' && name.includes('worksheets/')) {
        for (const el of elements) {
          if (el.name === 'row' && el.attrs.has('ht') && Number(el.attrs.get('ht')) % XLSX.ROW_PT) bad(`row height ${el.attrs.get('ht')}`);
          if (el.name === 'sheetFormatPr' && Number(el.attrs.get('defaultRowHeight')) !== XLSX.ROW_PT) bad('default row height');
          if (el.name === 'col') {
            const w = Number(el.attrs.get('width'));
            if (w !== XLSX.MARGIN_COL && (2 * w) % XLSX.MODULE) bad(`column width ${w}`);
          }
        }
      }
    }
  }
  assert.ok(checked > 20, 'the slide offset check saw too few offsets');
  assert.deepEqual(problems, []);
});

test('the VRT specs enumerate exactly the committed specimens', async () => {
  const { resolve } = await import('node:path');
  const { selfVrtCorpusFiles } = await import('../visual/private-corpus.mjs');
  const previous = process.cwd();
  try {
    for (const format of FORMATS) {
      // The specs list the corpus from their package directory.
      process.chdir(resolve(outputDir(format), '../..'));
      assert.deepEqual(
        selfVrtCorpusFiles({ corpus: 'corpus', format }),
        cases.filter((c) => c.format === format).map((c) => `corpus/${fileName(c)}`)
          .sort((left, right) => left.localeCompare(right, undefined, { numeric: true })),
      );
    }
  } finally {
    process.chdir(previous);
  }
});
