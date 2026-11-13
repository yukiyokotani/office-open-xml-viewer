import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { findReferenceFontMetrics, parseOpenTypeResourceMetrics } from '@silurus/ooxml-core';
import { parseVariableOpenTypeResourceMetrics } from '@silurus/ooxml-core/internal/font-sources';
import { sourceFontLayoutOptions } from './testing/font-sources.js';
import { createLayoutServices } from './layout-runtime.js';
import { buildSegments, layoutLines } from './line-layout.js';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';
import { fontFiles } from '@silurus/ooxml-core/internal/font-sources';
import type { DocRun, DocxDocumentModel } from './types.js';
import { sourceFontLineMetrics } from './source-font-metrics.js';
import type { PreparedServedFaceLines } from './served-face-lines.js';

// A face served through the `fonts` option paints the text, so its own tables
// size the line, exactly as an installed or embedded face with the same tables
// (Word: hhea ascender + descender + lineGap for a non-Far-East face). The
// Canvas font box below omits the line gap, as every browser's does.

const FIXTURE = readFileSync(new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap.ttf', import.meta.url));
const TABLES = parseOpenTypeResourceMetrics(new Uint8Array(FIXTURE))!;
const VARIABLE = parseVariableOpenTypeResourceMetrics(new Uint8Array(readFileSync(
  new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap-var.ttf', import.meta.url))))!;

const ASCENT = 0.8;
const DESCENT = 0.2;

function canvasBoxContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => {
      const px = Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? 10);
      return {
        width: [...text].length * px * 0.6,
        actualBoundingBoxAscent: px * ASCENT,
        actualBoundingBoxDescent: px * DESCENT,
        fontBoundingBoxAscent: px * ASCENT,
        fontBoundingBoxDescent: px * DESCENT,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

/** The document names `family` (theme minor font), so the font plan demands it. */
function model(family: string): DocxDocumentModel {
  return {
    minorFont: family,
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
}

function run(family: string, text = 'Hello'): DocRun {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: family,
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocRun;
}

type Face = Parameters<typeof sourceFontLayoutOptions>[1][number];

/** Single line of one 10pt run, in em. */
function lineEm(
  family: string,
  faces: readonly Face[],
  sources = [fontFiles({ faces: faces.map((face) => ({ family: face.family, url: `https://app.example/${face.family}.ttf` })) })],
  text = 'Hello',
): number {
  const doc = model(family);
  const ctx = canvasBoxContext();
  const services = createLayoutServices(doc, {
    measureContext: ctx,
    ...sourceFontLayoutOptions(doc, faces, { sources }),
  });
  const segments = buildSegments([run(family, text)], { pageIndex: 0, totalPages: 1, layoutServices: services });
  const [line] = layoutLines(ctx, segments, 400, 0, 1);
  return Math.round((line!.ascent + line!.descent) / 10 * 1e6) / 1e6;
}

describe('served font line metrics', () => {
  const word = wordOpenTypeAutoLineRatios({ ...TABLES, farEastCodePage: false })!;

  it('the fixture has a line gap the Canvas box cannot show', () => {
    expect(TABLES).toMatchObject({ unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 300, farEastCodePage: false });
    expect(word.lineHeightRatio).toBeCloseTo(1.3, 9);
  });

  it('sizes a served face from its own tables, line gap included', () => {
    const served = { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES };
    expect(lineEm('Ooxml Fixture Gap', [served])).toBeCloseTo(word.lineHeightRatio, 6);
  });

  it('uses the served bytes, not the catalogue entry of the authored name', () => {
    // Metric-compatible substitutes are commonly served under the Office name.
    const served = { family: 'Calibri', weight: '400', style: 'normal', resourceMetrics: TABLES };
    expect(referenceFontLineMetrics('Calibri')!.lineHeightRatio).not.toBeCloseTo(1.3, 3);
    expect(lineEm('Calibri', [served])).toBeCloseTo(word.lineHeightRatio, 6);
  });

  it('sizes a served substitute from the substitute face tables', () => {
    const served = { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES };
    const sources = [fontFiles({
      faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/gap.ttf' }],
      substitutes: [{ for: 'Authored Missing', use: 'Ooxml Fixture Gap' }],
    })];
    expect(lineEm('Authored Missing', [served], sources)).toBeCloseTo(word.lineHeightRatio, 6);
  });

  it('applies to synthesized bold and italic of a regular-only served family', () => {
    const served = { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES };
    const doc = model('Ooxml Fixture Gap');
    const ctx = canvasBoxContext();
    const services = createLayoutServices(doc, {
      measureContext: ctx,
      ...sourceFontLayoutOptions(doc, [served], { sources: [fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/gap.ttf' }] })] }),
    });
    const segments = buildSegments([{ ...run('Ooxml Fixture Gap'), bold: true, italic: true } as DocRun],
      { pageIndex: 0, totalPages: 1, layoutServices: services });
    const [line] = layoutLines(ctx, segments, 400, 0, 1);
    expect((line!.ascent + line!.descent) / 10).toBeCloseTo(word.lineHeightRatio, 6);
  });

  it('keeps its served metrics whatever another open document serves under the same family name', () => {
    // Another document's source serves "Ooxml Fixture Gap" with other faces.
    // Its registration (under that document's private family, or the real
    // name as before) neither joins this document's matched set nor declines
    // its served metrics.
    const served = { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES };
    const doc = model('Ooxml Fixture Gap');
    const options = sourceFontLayoutOptions(doc, [served], {
      sources: [fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/gap.ttf' }] })],
    });
    const other = [{ family: 'Ooxml Fixture Gap', status: 'loaded', weight: '400', style: 'normal', unicodeRange: 'U+0-10FFFF' }];
    const alone = sourceFontLineMetrics(options.sourceFonts, null);
    expect(Object.keys(alone).length).toBeGreaterThan(0);
    expect(sourceFontLineMetrics(options.sourceFonts, other as unknown as FontFaceSet)).toEqual(alone);
  });

  it('keeps the Canvas box for a served face whose tables are unavailable', () => {
    // A face the browser fetched itself (or WOFF/WOFF2): no bytes to read.
    const served = { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal' };
    expect(lineEm('Ooxml Fixture Gap', [served])).toBeCloseTo(ASCENT + DESCENT, 6);
  });

  it('declines a matched face set when one of its subsets has no tables', () => {
    const faces = [
      { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', unicodeRange: 'U+0-7F', resourceMetrics: TABLES },
      { family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', unicodeRange: 'U+80-FF' },
    ];
    expect(lineEm('Ooxml Fixture Gap', faces)).toBeCloseTo(ASCENT + DESCENT, 6);
  });

  it('lends a subset geometry only to text inside its declared range', () => {
    const faces = [{ family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', unicodeRange: 'U+0-40', resourceMetrics: TABLES }];
    expect(lineEm('Ooxml Fixture Gap', faces, undefined, '123')).toBeCloseTo(word.lineHeightRatio, 6);
    expect(lineEm('Ooxml Fixture Gap', faces, undefined, 'Hello')).toBeCloseTo(ASCENT + DESCENT, 6);
  });

  it('sizes a variable face at the instance CSS selects for each tuple', () => {
    // MVAR has no tag for the hhea values the Latin projection reads, so every
    // instance keeps the default hhea line: (800 + 200 + 300) / 1000.
    const served = { family: 'Ooxml Fixture Gap', weight: '100 900', style: 'normal', resourceMetrics: VARIABLE };
    const doc = model('Ooxml Fixture Gap');
    const ctx = canvasBoxContext();
    const services = createLayoutServices(doc, {
      measureContext: ctx,
      ...sourceFontLayoutOptions(doc, [served], { sources: [fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/var.ttf' }] })] }),
    });
    const em = (bold: boolean) => {
      const segments = buildSegments([{ ...run('Ooxml Fixture Gap'), bold } as DocRun], { pageIndex: 0, totalPages: 1, layoutServices: services });
      const [line] = layoutLines(ctx, segments, 400, 0, 1);
      return (line!.ascent + line!.descent) / 10;
    };
    expect(em(false)).toBeCloseTo(1.3, 6);
    expect(em(true)).toBeCloseTo(1.3, 6);
  });

  it('applies the usWin deltas (hcla/hcld) of the instance to a Far East variable face', () => {
    // The fixture's usWinAscent masters are 800/900 (wght 400/900); wght 700
    // is 0.6 of the way: 860. The class is the parsed one, set Far East here.
    const farEast = { ...VARIABLE, farEastCodePage: true };
    const loaded = sourceFontLayoutOptions(model('V'), [{ family: 'V', weight: '100 900', style: 'normal', resourceMetrics: farEast }],
      { sources: [fontFiles({ faces: [{ family: 'V', url: 'https://app.example/v.ttf' }] })] }).sourceFonts;
    const metrics = Object.values(sourceFontLineMetrics(loaded, null));
    const at = (winAscent: number) => wordOpenTypeAutoLineRatios({ ...TABLES, winAscent, winDescent: 200, farEastCodePage: true })!.lineHeightRatio;
    const bold = metrics.find((metric) => metric.weight === 700 && metric.style === 'normal')!;
    const regular = metrics.find((metric) => metric.weight === 400 && metric.style === 'normal')!;
    expect(regular.lineHeightRatio).toBeCloseTo(at(800), 9);
    expect(bold.lineHeightRatio).toBeCloseTo(at(860), 9);
    expect(bold.farEastCodePage).toBe(true);
  });

  it('carries the parsed code-page class of a table-backed face', () => {
    const loaded = sourceFontLayoutOptions(model('Ooxml Fixture Gap'), [{ family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES }],
      { sources: [fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/gap.ttf' }] })] }).sourceFonts;
    // Grid allocation (segment-builder.ts) reads an explicit false.
    expect(Object.values(sourceFontLineMetrics(loaded, null)).every((metric) => metric.farEastCodePage === false)).toBe(true);
    const doc = model('Ooxml Fixture Gap');
    const services = createLayoutServices(doc, {
      measureContext: canvasBoxContext(),
      ...sourceFontLayoutOptions(doc, [{ family: 'Ooxml Fixture Gap', weight: '400', style: 'normal', resourceMetrics: TABLES }],
        { sources: [fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: 'https://app.example/gap.ttf' }] })] }),
    });
    const [segment] = buildSegments([run('Ooxml Fixture Gap')], { pageIndex: 0, totalPages: 1, layoutServices: services });
    expect((segment as { resolvedLatinGridCellAllocation?: boolean }).resolvedLatinGridCellAllocation).toBe(true);
  });

  it('clamps a variable instance to the weight descriptor (a 400-only face paints bold synthesized)', () => {
    const loaded = sourceFontLayoutOptions(model('V'), [{ family: 'V', weight: '400', style: 'normal', resourceMetrics: VARIABLE }],
      { sources: [fontFiles({ faces: [{ family: 'V', url: 'https://app.example/v.ttf' }] })] }).sourceFonts;
    const metrics = Object.values(sourceFontLineMetrics(loaded, null));
    expect(metrics.find((metric) => metric.weight === 700 && metric.style === 'normal')!.lineHeightRatio).toBeCloseTo(1.3, 9);
  });

  it('projects a measured normal line for a face without tables, owned by its private family', () => {
    const loaded = sourceFontLayoutOptions(model('Measured Face'), [{ family: 'Measured Face', weight: '400', style: 'normal', unicodeRange: 'U+0-FF' }],
      { sources: [fontFiles({ faces: [{ family: 'Measured Face', url: 'https://app.example/m.woff2' }] })] }).sourceFonts;
    const line = { ascentRatio: 0.8, descentRatio: 0.2, lineGapRatio: 0.3 };
    const prepared: PreparedServedFaceLines = {
      faceLines: new Map([[0, { ownershipFamily: '__probe', lines: { '400:normal': line, '700:normal': line, '400:italic': null, '700:italic': line } }]]),
      familyLines: new Map(),
    };
    const metrics = Object.values(sourceFontLineMetrics(loaded, null, prepared));
    const regular = metrics.find((metric) => metric.weight === 400 && metric.style === 'normal')!;
    expect(regular).toMatchObject({
      family: 'Measured Face', requiresGlyphCoverage: true, nativeRoutePrimaryFamily: '"__probe"',
      unicodeRanges: [[0, 0xff]],
    });
    // A CSS normal line cannot show the code pages: the class stays unknown.
    expect(regular.farEastCodePage).toBeUndefined();
    expect(regular.lineHeightRatio).toBeCloseTo(1.3, 9);
    expect(regular.designAscentRatio).toBeCloseTo(1.1, 9);
    // An unmeasurable tuple is declined (its spans keep the Canvas box).
    expect(metrics.some((metric) => metric.weight === 400 && metric.style === 'italic')).toBe(false);
    // Without its ownership sentinels (no DOM here) the line keeps the Canvas box.
    expect(lineEm('Measured Face', [{ family: 'Measured Face', weight: '400', style: 'normal' }])).toBeCloseTo(ASCENT + DESCENT, 6);
  });

  it('declines a measured face of a family the catalogue lists as Far East', () => {
    // A CSS normal line shows neither the code pages nor usWin, and two fonts
    // can share an hhea box with different usWin: nothing is inferred from it,
    // not even for a box equal to the catalogued face's.
    const loaded = sourceFontLayoutOptions(model('Yu Gothic'), [{ family: 'Yu Gothic', weight: '400', style: 'normal' }],
      { sources: [fontFiles({ faces: [{ family: 'Yu Gothic', url: 'https://app.example/yu.woff2' }] })] }).sourceFonts;
    const prepared = (line: { ascentRatio: number; descentRatio: number; lineGapRatio: number }): PreparedServedFaceLines => ({
      faceLines: new Map([[0, { ownershipFamily: '__probe', lines: { '400:normal': line } }]]), familyLines: new Map(),
    });
    const profiles = findReferenceFontMetrics('Yu Gothic', { weight: 400, style: 'normal' });
    expect(profiles.length).toBeGreaterThan(0);
    const [ascender, descender, gap] = profiles[0]!.hhea;
    const em = profiles[0]!.unitsPerEm;
    expect(Object.keys(sourceFontLineMetrics(loaded, null,
      prepared({ ascentRatio: ascender / em, descentRatio: -descender / em, lineGapRatio: gap / em })))).toEqual([]);
    expect(Object.keys(sourceFontLineMetrics(loaded, null,
      prepared({ ascentRatio: 0.8, descentRatio: 0.2, lineGapRatio: 0.3 })))).toEqual([]);
  });
});
