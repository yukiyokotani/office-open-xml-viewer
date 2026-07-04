import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  SectionProps,
} from './types';

// Regression: inter-word space advance must survive a `measureText` engine that
// TRIMS trailing whitespace from `TextMetrics.width` (Firefox / WebKit / Safari,
// per the HTML canvas history). The docx layout stores a text segment's advance
// as `measureText(segmentTextIncludingTrailingSpace).width`. On a trimming engine
// that omits the space, so the pen under-advances and the next word collides
// ("Other countries" → "Othercountries"). The fix must recover the trailing-space
// advance so measure==draw holds on EVERY engine (all weights, all alignments) —
// not a bold-only patch.
//
// This mock context models a TRIMMING engine: `measureText(s)` returns the width
// of `s` with trailing spaces removed (interior spaces still counted). It also
// records fillText so we can confirm the drawn advance too.

const FONT_PX = 20;
const CHAR_W = 10; // per non-space glyph advance in the stub (scale 1)
const SPACE_W = 5; // interior-space advance in the stub

/** Width of `s` under a TRAILING-SPACE-TRIMMING engine: interior spaces count,
 *  trailing spaces do NOT (the exact Firefox/WebKit measureText divergence). */
function trimmingMeasure(s: string): number {
  const trimmed = s.replace(/ +$/, ''); // drop ONLY trailing spaces
  let w = 0;
  for (const ch of trimmed) w += ch === ' ' ? SPACE_W : CHAR_W;
  return w;
}

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: { text: string; x: number }[];
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fillTextCalls: { text: string; x: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      const scale = p / FONT_PX;
      const w = trimmingMeasure(s) * scale;
      return {
        width: w,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number) {
      fillTextCalls.push({ text, x });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

function textRun(text: string, bold = false): DocxTextRun {
  return {
    text, bold, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

type DocRun = DocParagraph['runs'][number];

function para(
  runs: DocxTextRun[],
  opts: { alignment?: DocParagraph['alignment'] } = {},
): BodyElement {
  const p: DocParagraph = {
    alignment: opts.alignment ?? 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics',
    widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 1000, pageHeight: 400,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  } as SectionProps;
}

function doc(body: BodyElement[], sec: SectionProps): DocxDocumentModel {
  return {
    section: sec, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(
  body: BodyElement[],
  sec: SectionProps,
): Promise<{ runs: DocxTextRunInfo[]; fillTextCalls: { text: string; x: number }[] }> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc(body, sec), canvas, 0, {
    dpr: 1,
    width: sec.pageWidth, // scale = 1 (px per pt)
    onTextRun: (r) => runs.push(r),
  });
  return { runs, fillTextCalls };
}

describe('inter-word space survives a trailing-space-trimming measureText engine', () => {
  it('bold left-aligned single run: "Other countries" keeps the space (no collision)', async () => {
    const { runs } = await render([para([textRun('Other countries', true)])], section());
    const other = runs.find((r) => /^Other/.test(r.text));
    const countries = runs.find((r) => /countries/.test(r.text));
    expect(other).toBeDefined();
    expect(countries).toBeDefined();
    // "Other" = 5 glyphs × 10 = 50; trailing space = 5 ⇒ next word starts at 55,
    // not 50. If the space advance is dropped, countries.x === other.x + 50.
    const otherGlyphs = 5 * CHAR_W; // 50
    expect(countries!.x).toBeGreaterThan(other!.x + otherGlyphs + SPACE_W - 0.01);
  });

  it('regular (non-bold) left run keeps the space too — fix is weight-independent', async () => {
    const { runs } = await render([para([textRun('Other countries', false)])], section());
    const other = runs.find((r) => /^Other/.test(r.text))!;
    const countries = runs.find((r) => /countries/.test(r.text))!;
    expect(countries.x).toBeGreaterThan(other.x + 5 * CHAR_W + SPACE_W - 0.01);
  });

  it('three words keep both spaces', async () => {
    const { runs } = await render([para([textRun('one two three', false)])], section());
    const one = runs.find((r) => /^one/.test(r.text))!;
    const two = runs.find((r) => /^two/.test(r.text))!;
    const three = runs.find((r) => /three/.test(r.text))!;
    expect(two.x).toBeGreaterThan(one.x + 3 * CHAR_W + SPACE_W - 0.01);
    expect(three.x).toBeGreaterThan(two.x + 3 * CHAR_W + SPACE_W - 0.01);
  });
});
