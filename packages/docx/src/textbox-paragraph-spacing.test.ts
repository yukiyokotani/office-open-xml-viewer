import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { createLayoutServices } from './layout-runtime.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, ShapeRun, ShapeText } from './types';

// ECMA-376 §17.3.1.33 — inside a text box, consecutive paragraphs collapse their
// inter-paragraph spacing to max(prev.after, this.before) (NOT the sum), exactly
// like the body flow. The first paragraph reserves only its own spaceBefore.

interface Call { text: string; y: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
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
    fillText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    strokeText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function block(text: string, spaceBefore: number, spaceAfter: number): ShapeText {
  return { text, fontSizePt: 10, fontFamily: 'Times New Roman', alignment: 'left', spaceBefore, spaceAfter } as unknown as ShapeText;
}

// A body paragraph carrying a wrapNone text box with two paragraphs A and B.
function docWithTwoBlockTextbox(): DocxDocumentModel {
  const shape = {
    type: 'shape',
    widthPt: 300, heightPt: 200,
    anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: false, anchorYFromPara: true,
    anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [block('AAA', 0, 20), block('BBB', 10, 0)],
  } as unknown as ShapeRun;
  const para = {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [shape as unknown as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section,
    body: [para as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('text-box paragraph spacing (§17.3.1.33)', () => {
  it("collapses the inter-paragraph gap to max(prev.after, this.before), not the sum", async () => {
    const { canvas, calls } = makeRecordingCanvas();
    const model = docWithTwoBlockTextbox();
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1,
      width: 400,
      layoutServices: createLayoutServices(model, {
        measureContext: canvas.getContext('2d'),
        localMetrics: testFontSnapshot([{
          family: 'Times New Roman',
          lineHeightRatio: 2355 / 2048,
        }]),
      }),
    });
    const a = calls.find((c) => c.text === 'AAA');
    const b = calls.find((c) => c.text === 'BBB');
    expect(a).toBeDefined();
    expect(b).toBeDefined();
    // The explicit fixture resource has a 2355/2048 design ratio, so one 10pt
    // line is 11.499px—larger than the
    // mock's naive 0.8 + 0.2 = 1.0 em / 10 px). Both lines share font/size and
    // carry no line-spacing rule, so no box expansion ⇒ their half-leading
    // baseline offsets are equal and cancel in the delta. Gap A→B baseline =
    // blockHeight + max(after=20, before=10) = 11.499 + 20 = 31.499 (COLLAPSE),
    // NOT 11.499 + (20 + 10) = 41.499 (sum). The collapse-vs-sum distinction — the
    // point of this test—is independent of the explicit resource metric.
    const gap = (b as Call).y - (a as Call).y;
    expect(gap).toBeCloseTo(31.5, 1);
  });
});
