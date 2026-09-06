import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Issue #981 (dense-page bottom fill).
//
// Word's page-break test is baseline-based: a line whose baseline sits within the
// text area may let its below-baseline whitespace (descent + half of any leading)
// extend into the bottom margin. An empty paragraph's mark line paints no ink, so an
// empty paragraph that would otherwise be the last line on a page is KEPT on the page
// when only that invisible whitespace overflows the bottom content edge — rather than
// pushed to the next page's top, which would cascade the following visible content
// down by ~one line (the Thai reference put a formula paragraph and a table row one
// page late on three boundaries). This is Word runtime behaviour, reconstructed from
// its output; ECMA-376 §17.3.1.29 requires the mark line box to exist but neither it
// nor §17.3.1.33 specifies baseline-based pagination.
//
// The allowance is tightly scoped — it is taken ONLY when it is both observable and
// invisible. These synthetic cases pin each gate (canvas mock with glyph metrics
// exactly linear in the font px size; four single-line paragraphs fill the content
// band and the fifth element grazes the bottom edge — its box overflows the band but
// its baseline stays within it). Confirmed to fail before the fix: the grazing case
// packs 4 elements on page 1 and pushes the empty, so pages[0].length is 4 not 5.

interface Call { text: string; x: number; y: number; }

/** Recording canvas: glyph width linear in font px; fontBoundingBox a fixed
 *  0.8 / 0.2 em split so the below-baseline extent is a clean fraction of the box. */
function makeLinearCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeLinearCanvas().canvas.getContext('2d'); }
};

function para(text: string, extra: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
        }]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...extra,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight: number): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

const B = (...ps: DocParagraph[]): BodyElement[] => ps.map((p) => p as unknown as BodyElement);
const bodyPages = (model: DocxDocumentModel) => layoutDocument(
  model,
  createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]) }),
  { currentDateMs: 0 },
).pages.map((page) => page.layers.body);

// Content band = pageHeight − 2·margin = 57pt holds four single-line paragraphs
// (~12pt each → y = 48) exactly; the fifth element's ~12pt box then spans [48, 60],
// overflowing the band bottom (57) while its baseline (57) is still within it.
const PAGE_HEIGHT = 77;

describe('canonical body layout — trailing empty-paragraph mark grazes the bottom margin (issue #981)', () => {
  it('KEEPS an inkless empty paragraph on the page when ink-bearing content follows and only its below-baseline whitespace overflows', () => {
    // a,b,c,d fill page 1; the empty grazes the bottom (baseline within the band) and
    // is KEPT (page 1 = a,b,c,d,empty = 5); the following visible "e" flows to page 2.
    const pages = bodyPages(doc(B(para('a'), para('b'), para('c'), para('d'), para(''), para('e')), PAGE_HEIGHT));
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5); // was 4 before the fix (empty pushed to page 2 top)
    expect(pages[1].length).toBe(1);
  });

  it('does NOT graze a document-terminal empty run (no following ink) — Word keeps the trailing blank page', () => {
    // No visible content follows the empty, so pushing it changes nothing observable;
    // the empty is paginated normally (pushed to page 2). This preserves Word's
    // trailing blank page — the allowance must not silently drop terminal pages.
    const pages = bodyPages(doc(B(para('a'), para('b'), para('c'), para('d'), para('')), PAGE_HEIGHT));
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(4);
  });

  it('does NOT graze an empty paragraph that paints shading (its box paints ink)', () => {
    // A shaded empty paragraph fills its whole mark box, so grazing would push visible
    // fill into the bottom margin. It is treated as a normal box: pushed to page 2.
    const pages = bodyPages(doc(B(para('a'), para('b'), para('c'), para('d'), para('', { shading: 'ff0000' }), para('e')), PAGE_HEIGHT));
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(4);
  });

  it('does NOT graze a VISIBLE last line — its glyphs must fit the full box', () => {
    // The fifth element carries text, so it is not a mark line; it must fit its full
    // box and is pushed to page 2 (proving the allowance is scoped to empty marks).
    const pages = bodyPages(doc(B(para('a'), para('b'), para('c'), para('d'), para('e'), para('f')), PAGE_HEIGHT));
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(4);
  });

  it('does NOT graze an empty when a hard page break separates it from the following ink', () => {
    // A forced page break starts "e" on a fresh page regardless of whether the empty
    // is kept, so there is no cascade to justify grazing — the empty is paginated
    // normally (pushed), leaving page 1 = a,b,c,d. Without the forced-boundary guard
    // the look-ahead would find "e" past the break and wrongly graze the empty
    // (page 1 = a,b,c,d,empty). (Codex review of the #981 fix.)
    const pageBreak = { type: 'pageBreak' } as unknown as BodyElement;
    const pages = bodyPages(
      doc([...B(para('a'), para('b'), para('c'), para('d'), para('')), pageBreak, ...B(para('e'))], PAGE_HEIGHT),
    );
    expect(pages[0].length).toBe(4);
  });

  it('does not create an intermediate blank page solely for an inkless section mark', () => {
    const sectionBreak = {
      type: 'sectionBreak',
      kind: 'nextPage',
      geom: {
        pageWidth: 200,
        pageHeight: PAGE_HEIGHT,
        marginTop: 10,
        marginRight: 10,
        marginBottom: 10,
        marginLeft: 10,
        headerDistance: 4,
        footerDistance: 4,
      },
    } as unknown as BodyElement;
    const pages = bodyPages(doc([
      ...B(para('a'), para('b'), para('c'), para('d'), para('')),
      sectionBreak,
      ...B(para('e')),
    ], PAGE_HEIGHT));

    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5);
    expect(pages[1].length).toBe(1);
  });
});
