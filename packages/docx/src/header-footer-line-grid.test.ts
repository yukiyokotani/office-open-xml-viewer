import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { ParagraphLayout, PaintNode } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// `word-header-footer-off-line-grid`: a §17.6.5 docGrid line pitch applies
// to body lines only; header and footer lines keep their own height.

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

function para(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
}

function lineHeights(gridKind: string, textDirection?: string) {
  const section = {
    pageWidth: 600, pageHeight: 800,
    marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    docGridType: gridKind, docGridLinePitch: 18,
    ...(textDirection ? { textDirection } : {}),
  } as SectionProps;
  const document = {
    section, body: [para('Body')],
    headers: { default: { body: [para('Head')] }, first: null, even: null },
    footers: { default: { body: [para('Foot')] }, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const page = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  ).pages[0]!;
  const firstLine = (nodes: readonly PaintNode[]) =>
    nodes.find((node): node is ParagraphLayout => node.kind === 'paragraph')!.lines[0]!.advancePt;
  return {
    body: firstLine(page.layers.body),
    header: firstLine(page.layers.header),
    footer: firstLine(page.layers.footer),
  };
}

describe('header and footer on a §17.6.5 line grid', () => {
  it.each([undefined, 'tbRl', 'btLr', 'tbLrV'])('keeps %s header and footer lines off a "lines" grid', (direction) => {
    const heights = lineHeights('lines', direction);
    expect(heights.body).toBe(18);
    expect(heights.header).toBeLessThan(18);
    expect(heights.footer).toBe(heights.header);
  });

  it.each(['linesAndChars', 'snapToChars'])('keeps header and footer lines off a %s grid too', (kind) => {
    for (const direction of [undefined, 'tbRl']) {
      const heights = lineHeights(kind, direction);
      expect(heights.body).toBe(18);
      expect(heights.header).toBeLessThan(18);
      expect(heights.footer).toBe(heights.header);
    }
  });
});
