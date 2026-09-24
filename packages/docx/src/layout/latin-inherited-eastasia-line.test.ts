import { expect, it } from 'vitest';
import { createLayoutServices } from '../layout-runtime.js';
import { testFontSnapshot } from './test-font-snapshot.js';
import { buildSegments, layoutLines, lineBoxHeight } from '../line-layout.js';
import type { DocRun, DocxDocumentModel } from '../types.js';

const context = {
  font: '10px serif', letterSpacing: '0px', fontKerning: 'auto',
  measureText: (text: string) => ({
    width: text.length * 5,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
  }),
} as unknown as CanvasRenderingContext2D;

function measuredLine(text: string, eastAsia: string | null, lineGridActive: boolean) {
  const document = {
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36,
      footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as DocxDocumentModel;
  const exactLocal: ReturnType<typeof testFontSnapshot> = eastAsia
    ? testFontSnapshot([{ family: eastAsia }]) : {};
  if (eastAsia) {
    const key = eastAsia.toLowerCase();
    exactLocal[key] = {
      ...exactLocal[key], sourceIdentity: `office-local:local("${eastAsia}")`,
    };
  }
  const services = createLayoutServices(document, {
    measureContext: context, localMetrics: exactLocal,
  });
  const runs = [{
    type: 'text', text,
    fontFamily: 'Calibri', fontFamilyEastAsia: eastAsia,
    fontSize: 10, bold: false, italic: false, underline: false,
    strikethrough: false,
  }] as DocRun[];
  const segments = buildSegments(runs, {
    pageIndex: 0, totalPages: 1, layoutServices: services,
    useFeLayout: true, lineGridActive,
    lineSpacing: { rule: 'auto', value: 1 },
  });
  const line = layoutLines(context, segments, 400, 0, 1)[0]!;
  const grid = lineGridActive
    ? { type: 'lines', linePitchPt: 18 } as const
    : undefined;
  return { line, advance: lineBoxHeight({ rule: 'auto', value: 1 },
    line.ascent, line.descent, 1, grid, false, line.intendedSingle,
    line.eastAsian, line.gridCountSingle) };
}

it('does not inflate a Latin-only non-grid line from an unused inherited eastAsia font', () => {
  // A paragraph style can supply an East Asian slot even when the visible run
  // contains only Latin text and selects its own ASCII theme face. With no
  // active §17.6.5 line grid, that unused slot cannot set the auto-line floor.
  const latinOnly = measuredLine('The consortium develops panels.', null, false);
  const inheritedEastAsia = measuredLine('The consortium develops panels.',
    'Arial Unicode MS', false);
  expect(inheritedEastAsia.advance).toBeCloseTo(latinOnly.advance, 8);
  expect(inheritedEastAsia.line.intendedSingle)
    .toBeCloseTo(latinOnly.line.intendedSingle, 8);
});

it('reserves the inherited East Asian axis for an active line grid only', () => {
  const latin = measuredLine('A', 'Arial Unicode MS', false);
  const gridded = measuredLine('A', 'Arial Unicode MS', true);
  expect(gridded.line.intendedSingle).toBeGreaterThan(latin.line.intendedSingle);
  expect(gridded.line.eastAsian).toBe(true);
  expect(gridded.advance).toBe(18);
});

it('keeps an actual East Asian glyph on its selected face without a line grid', () => {
  const literalEastAsian = measuredLine('あ', 'Arial Unicode MS', false);
  expect(literalEastAsian.line.eastAsian).toBe(true);
  expect(literalEastAsian.line.intendedSingle).toBeGreaterThan(10);
});
