import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseOpenTypeResourceMetrics } from '@silurus/ooxml-core';
import { sourceFontLayoutOptions } from './testing/font-sources.js';
import { createLayoutServices } from './layout-runtime.js';
import { buildSegments, layoutLines } from './line-layout.js';
import { wordOpenTypeAutoLineRatios } from './layout/line-compatibility.js';
import { referenceFontLineMetrics } from './reference-font-line-metrics.js';
import { fontFiles } from '@silurus/ooxml-core/internal/font-sources';
import type { DocRun, DocxDocumentModel } from './types.js';

// A face served through the `fonts` option paints the text, so its own tables
// size the line, exactly as an installed or embedded face with the same tables
// (Word: hhea ascender + descender + lineGap for a non-Far-East face). The
// Canvas font box below omits the line gap, as every browser's does.

const FIXTURE = readFileSync(new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap.ttf', import.meta.url));
const TABLES = parseOpenTypeResourceMetrics(new Uint8Array(FIXTURE))!;

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
});
