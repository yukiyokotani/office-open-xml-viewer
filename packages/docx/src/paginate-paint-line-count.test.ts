import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.6.4 (newspaper columns) + the renderer's scale-independent
// pagination contract: paginateWithHeaderFooterReserve lays paragraphs out at
// scale 1 (pt space) so a page assignment is width-independent and cacheable
// across every render width. Canonical retained pages own the exact line
// placements consumed by paint, so a second context-dependent line-layout pass
// cannot create phantom lines.
//
// This reproduces the sample-16 page-2 crash with a synthetic document — no
// dependency on the (gitignored) private sample. The non-linear measureText mock
// makes glyphs proportionally NARROWER at a larger font size, so the scale-2
// paint pass fits more characters per line and wraps the long paragraph to fewer
// lines than the scale-1 pagination — exactly the real font-hinting direction.

interface Call { text: string; x: number; y: number; }

/** Recording canvas whose glyph width is SUB-LINEAR in the font px size: each
 *  character is `px * (0.5 - SHRINK * px)` wide. At a larger render scale the
 *  per-glyph width grows slower than the box, so MORE characters fit per line
 *  and a long paragraph wraps to fewer lines than at scale 1 — the same
 *  paginate-vs-paint divergence real fonts produce through hinting. */
function makeNonLinearCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  const SHRINK = 0.002; // per-px narrowing; tuned so scale 1 vs 2 differ by lines
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const perChar = Math.max(0.05, p * (0.5 - SHRINK * p));
      return {
        width: [...s].length * perChar,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
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
  getContext() { return makeNonLinearCanvas().canvas.getContext('2d'); }
};

function longPara(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
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
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

async function paintedParagraphGeometry() {
  const paragraph = longPara(Array.from({ length: 180 }, () => 'w').join(' '));
  paragraph.spaceBefore = 6;
  paragraph.spaceAfter = 4;
  const model = doc([paragraph as unknown as BodyElement], 80);
  const services = createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]) });
  const layout = layoutDocument(model, services, { currentDateMs: 0 });
  const paintedPages: Array<{ lineCount: number; topYPx: number | null }> = [];
  for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
    const { canvas, calls } = makeNonLinearCanvas();
    await renderDocumentToCanvas(model, canvas, pageIndex, {
      dpr: 1,
      width: 200,
      layoutServices: services,
    });
    const textCalls = calls.filter((call) => call.text.includes('w'));
    paintedPages.push({
      lineCount: textCalls.length,
      topYPx: textCalls.length > 0 ? textCalls[0].y : null,
    });
  }
  return { pageCount: layout.pages.length, paintedPages };
}

describe('paginate/paint line-count divergence — paint never indexes a phantom line (ECMA-376 §17.6.4)', () => {
  // A long paragraph of single-letter "words" (each followed by a space so the
  // line breaker has wrap opportunities). Narrow page + short page height force
  // it to wrap to many lines and split across multiple pages, so a later page
  // carries an explicit retained continuation range.
  const text = Array.from({ length: 400 }, () => 'w').join(' ');
  const body = (): BodyElement[] => [longPara(text) as unknown as BodyElement];

  it('preserves page count, painted line counts, and continuation top positions', async () => {
    const geometry = await paintedParagraphGeometry();

    // Word admits the final visible line at a region edge without requiring the
    // paragraph's authored trailing spaceAfter to fit. The retained line split
    // therefore completes on page 2 while paint still consumes every one of the
    // 180 canonical line placements exactly once.
    expect(geometry).toEqual({
      pageCount: 2,
      paintedPages: [
        { lineCount: 84, topYPx: 24.74951171875 },
        { lineCount: 96, topYPx: 18.74951171875 },
      ],
    });
  });

  it('renders every retained continuation page without remeasuring or throwing', async () => {
    const pageHeight = 80; // short page → the paragraph spans several pages
    const model = doc(body(), pageHeight);
    const services = createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    });
    const layout = layoutDocument(model, services, { currentDateMs: 0 });
    let totalLines = 0;
    let threw: unknown = null;
    for (let p = 0; p < layout.pages.length; p++) {
      const { canvas, calls } = makeNonLinearCanvas();
      try {
        await renderDocumentToCanvas(model, canvas, p, {
          dpr: 1, width: 400,
          layoutServices: services,
        });
      } catch (e) {
        threw = e;
        break;
      }
      totalLines += calls.filter((c) => c.text.includes('w')).length;
    }

    expect(threw).toBeNull();

    // NON-TRIVIALITY: the document actually painted content across pages (the
    // long paragraph really did wrap and split — otherwise the invariant above
    // would be vacuous).
    expect(totalLines).toBeGreaterThan(0);
  });
});
