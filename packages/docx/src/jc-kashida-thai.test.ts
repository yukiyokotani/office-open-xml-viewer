import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps, DocRun } from './types.js';

// ECMA-376 §17.18.44 — the Arabic kashida variants (lowKashida / mediumKashida /
// highKashida) and thaiDistribute are full-justification jc values. They were
// previously unmapped, so a paragraph carrying them fell through to left
// alignment and its lines were NOT justified. These end-to-end tests prove each
// value now enters the same slack-distribution path as `both` / `distribute`:
// on a two-word line that under-fills the column, the words are pushed apart so
// the second word's right edge reaches the right margin.

interface FillCall {
  text: string;
  x: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number) { fills.push({ text, x }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

const PAGE_W = 300;
const FS = 10; // glyph width px; scale = 1 px/pt

function textRun(text: string): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FS, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocRun;
}

// The paragraph "aa bb <long word>" wraps: line 1 = "aa bb" (short, under-fills
// the 300px column), and the long third word cannot fit on line 1 so it wraps to
// line 2. Line 1 is therefore a NON-LAST line, which every full-justification jc
// value (both / justify / kashida) stretches by expanding its inter-word space.
// A manual <w:br/> is intentionally NOT used: a break-terminated line is treated
// as a last line (§17.18.44) and left un-stretched, which would defeat the test.
const LONG = 'w'.repeat(28); // 280px — cannot share line 1 with "aa bb ".
function jcPara(alignment: string): DocParagraph {
  return {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [textRun(`aa bb ${LONG}`)],
    defaultFontSize: FS, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(p: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_W, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(alignment: string) {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(jcPara(alignment)), canvas, 0, { dpr: 1, width: PAGE_W });
  return fills;
}

describe('jc kashida / thaiDistribute enter the justification path (§17.18.44)', () => {
  // Words are drawn as space-delimited tokens, so the second word reaches the
  // draw path as "bb " (with its trailing space).
  const findBB = (f: FillCall[]) => f.find((c) => c.text.trimEnd() === 'bb');

  // Baseline: an unjustified (left) paragraph leaves the words at natural
  // positions — "bb " starts right after "aa " at x = 3·FS = 30.
  it('left alignment does NOT justify (baseline)', async () => {
    const f = await render('left');
    const bb = findBB(f);
    expect(bb, '"bb" must be drawn').toBeDefined();
    expect(bb!.x).toBeCloseTo(3 * FS, 3);
  });

  for (const jc of ['both', 'lowKashida', 'mediumKashida', 'highKashida', 'thaiDistribute']) {
    it(`${jc} justifies the first (non-last) line, pushing "bb" toward the right margin`, async () => {
      const f = await render(jc);
      const bb = findBB(f);
      expect(bb, '"bb" must be drawn').toBeDefined();
      // Justified: the inter-word space is expanded so the visible "bb" ends at
      // the right margin (PAGE_W). It therefore starts far right of its natural x=30.
      expect(bb!.x).toBeGreaterThan(3 * FS + 1);
      // "bb" is 2 glyphs; the line-end space hangs past the margin
      // (WORD_LINE_END_SPACE_HANGS).
      expect(bb!.x + 2 * FS).toBeCloseTo(PAGE_W, 2);
    });
  }
});
