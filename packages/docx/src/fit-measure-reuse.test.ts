import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { ParagraphLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// M-1 — avoid the double measurement per non-split body paragraph.
//
// The paginator measures each body paragraph once for the fit decision
// (measureBodyParagraphAtCursor) and, before M-1, measured it AGAIN at its final
// placement to build the fragment. When a paragraph is not relocated after the
// estimate its final placement is identical, so the fragment now REUSES the fit
// measurement. This suite pins the production result directly: acquisition is
// deterministic and a relocated paragraph owns geometry for its final placement.
// ─────────────────────────────────────────────────────────────────────────────

let measureCount = 0;

/** OffscreenCanvas polyfill whose measureText increments a shared counter (linear
 *  glyph metric = fontPx * 0.5, matching the other renderer suites). */
function makeCountingCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measureCount++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeCountingCtx(); }
  };
});

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
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
    ...over,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
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

/** A serialisable projection of a layout's fragment partitions + placements — enough
 *  to prove the reuse produced the SAME layout as a fresh measurement. */
function projection(model: DocxDocumentModel) {
  const services = createLayoutServices(model, {
    localMetrics: testFontSnapshot([{
      family: 'Times New Roman',
      lineHeightRatio: 2355 / 2048,
    }]),
  });
  return layoutDocument(model, services, { currentDateMs: 0 }).pages.map((page) =>
    page.layers.body.filter((node): node is ParagraphLayout => node.kind === 'paragraph').map((f) => {
      return {
        continuation: f.continuation,
        sourcePath: f.source.path,
        lines: f.lines.length,
        availW: f.flowBounds.widthPt,
        startY: f.flowBounds.yPt,
        heightPt: f.advancePt,
        advances: f.lines.map((l) => l.advancePt),
      };
    }),
  );
}

describe('M-1 fit-check measurement reuse for non-split body fragments', () => {
  it('reuses the fit measurement and produces a deterministic retained layout', () => {
    // Several short, non-splitting, float-free paragraphs on a tall page — each is
    // measured for the fit decision and then attached at the SAME placement, so the
    // reuse fires for every one of them.
    const model = doc([
      para('alpha beta gamma delta') as unknown as BodyElement,
      para('one two three four five six') as unknown as BodyElement,
      para('lorem ipsum dolor sit amet consectetur') as unknown as BodyElement,
      para('the quick brown fox jumps over') as unknown as BodyElement,
    ]);

    measureCount = 0;
    const firstProjection = projection(model);
    const firstCount = measureCount;
    measureCount = 0;
    const secondProjection = projection(model);
    const secondCount = measureCount;

    expect(firstCount).toBeGreaterThan(0);
    expect(secondCount).toBe(firstCount);
    expect(secondProjection).toEqual(firstProjection);
  });

  it('does not affect a RELOCATED paragraph (placement changed → remeasured, layout still correct)', () => {
    // A tall paragraph with keepLines that must move to a fresh page: its fit estimate
    // is taken at a mid-page cursor, then the paragraph relocates, so the placement no
    // longer matches and the retained node must be acquired at the destination-page
    // origin rather than keeping the rejected cursor geometry.
    const filler = para('filler');
    const keep = para(Array.from({ length: 20 }, () => 'w').join(' '), { keepLines: true });
    const model = doc([filler as unknown as BodyElement, keep as unknown as BodyElement], 40);

    const retained = projection(model);
    expect(retained).toHaveLength(2);
    expect(retained[1]).toEqual([
      expect.objectContaining({ sourcePath: [1], startY: 10 }),
    ]);
  });
});
