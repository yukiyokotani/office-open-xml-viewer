import { describe, expect, it, vi } from 'vitest';

// Counts every ruby-cell lookup the layout makes, to bound the work of
// placing a line of ruby cells.
const lookups = vi.hoisted(() => ({ count: 0 }));
vi.mock('./line-breaker/ruby-cell.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./line-breaker/ruby-cell.js')>();
  return {
    ...actual,
    rubyCellOf: (segment: Parameters<typeof actual.rubyCellOf>[0]) => {
      lookups.count += 1;
      return actual.rubyCellOf(segment);
    },
  };
});

import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DocxDocumentModel, SectionProps } from './types.js';

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

/** Ruby-cell lookups made to lay out one vertical column of `cells` ruby cells. */
function lookupsFor(cells: number): number {
  const run = (index: number) => ({
    type: 'text', text: String.fromCodePoint(0x4e00 + index), bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 5, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ruby: { text: 'か', fontSizePt: 2, hpsRaisePt: 5 },
  });
  const section = {
    pageWidth: 600, pageHeight: 2000, marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false, textDirection: 'tbRl',
  } as SectionProps;
  const document = {
    section,
    body: [{
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0,
      spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: Array.from({ length: cells }, (_, index) => run(index)),
      defaultFontSize: 5, defaultFontFamily: 'NotInMetrics', widowControl: false,
    }],
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  lookups.count = 0;
  const page = layoutDocument(document, createLayoutServices(document, { measureContext: makeCtx() }), { currentDateMs: 0 }).pages[0]!;
  const lines = (page.layers.body.find((entry) => entry.kind === 'paragraph') as { lines: readonly unknown[] }).lines;
  expect(lines).toHaveLength(1);
  return lookups.count;
}

describe('ruby placement work (WORD_RUBY_ALIGN_PLACEMENT)', () => {
  it('places a line of ruby cells in work linear in its cells', () => {
    const small = lookupsFor(32);
    expect(small).toBeGreaterThan(0);
    // Twice the cells, about twice the lookups: no cell searches the line.
    expect(lookupsFor(64)).toBeLessThan(small * 3);
  });
});
