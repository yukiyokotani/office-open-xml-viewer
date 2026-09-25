import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
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

describe('paginate/paint line-count divergence — paint never indexes a phantom line (ECMA-376 §17.6.4)', () => {
  it.each(['Latin words', 'East Asian grid'] as const)(
    'retains and paints every source token exactly once across pages: %s',
    async (route) => {
      // Unique ordered labels detect missing, duplicated, or reordered content;
      // counting repeated glyphs cannot distinguish those failures. The grid
      // case exercises 20pt East Asian text at a 20pt grid-cell boundary, where
      // changing a single-line metric can move the paragraph continuation cursor.
      const eastAsian = route === 'East Asian grid';
      const tokens = Array.from({ length: eastAsian ? 36 : 60 }, (_, index) => {
        const number = String(index + 1).padStart(3, '0');
        return eastAsian
          ? `項目${number.replace(/\d/g, (digit) => String.fromCharCode(0xff10 + Number(digit)))}`
          : `item${number}`;
      });
      const text = tokens.join(eastAsian ? '' : ' ');
      const paragraph = longPara(text);
      if (eastAsian) {
        paragraph.runs = paragraph.runs.map((run) => ({
          ...run, fontSize: 20, fontFamily: 'Unresolved Grid Face',
          fontFamilyEastAsia: 'Unresolved Grid Face',
        }));
        paragraph.defaultFontSize = 20;
        paragraph.defaultFontFamily = 'Unresolved Grid Face';
        paragraph.defaultFontFamilyEastAsia = 'Unresolved Grid Face';
      }
      const model = doc([{ type: 'paragraph', ...paragraph }], eastAsian ? 140 : 80);
      if (eastAsian) {
        model.section.docGridType = 'lines';
        model.section.docGridLinePitch = 20;
        model.settings = { useFeLayout: true };
      }
      const services = createLayoutServices(model);
      const layout = layoutDocument(model, services, { currentDateMs: 0 });

      // Soft wrapping may suppress separator spaces at line ends. Compare all
      // visible source characters in order, retaining the unique token identity.
      const visibleText = (value: string) => value.replace(/ /g, '');
      const retainedPages = layout.pages.map((page) => page.layers.body
        .filter((node) => node.kind === 'paragraph')
        .flatMap((node) => node.lines)
        .flatMap((line) => line.placements)
        .filter((placement) => placement.kind === 'text')
        .map((placement) => placement.text).join(''));
      expect(retainedPages.length).toBeGreaterThan(1);
      expect(retainedPages.every((page) => visibleText(page).length > 0)).toBe(true);
      expect(visibleText(retainedPages.join(''))).toBe(visibleText(text));

      const paintedPages: string[] = [];
      for (let p = 0; p < layout.pages.length; p++) {
        const { canvas, calls } = makeNonLinearCanvas();
        // Paint at twice the layout width: the deliberately nonlinear measurer
        // would choose different lines if paint reacquired paragraph geometry.
        await renderDocumentToCanvas(model, canvas, p, {
          dpr: 1, width: 400, layoutServices: services,
        });
        paintedPages.push(calls.map((call) => call.text).join(''));
      }
      expect(paintedPages.map(visibleText)).toEqual(retainedPages.map(visibleText));
      expect(visibleText(paintedPages.join(''))).toBe(visibleText(text));
    },
  );
});
