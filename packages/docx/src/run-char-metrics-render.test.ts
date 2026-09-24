import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxTextRun,
  DocxDocumentModel,
  DocSettings,
  SectionProps,
} from './types';

// WD4 end-to-end draw tests: a run carrying w:spacing / w:w / w:position / w:kern
// must reach the glyph-draw with the corresponding ctx state (letterSpacing,
// horizontal scale transform, baseline y-offset, fontKerning) so that paint
// matches the widened / shifted layout the measure pass produced (measure==paint).

const FONT_PX = 20;

interface FillCall {
  text: string;
  x: number;
  y: number;
  letterSpacing: string;
  fontKerning: string;
  scaleX: number;
  translateX: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fontKerning = 'auto';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const fills: FillCall[] = [];
  // Track a simple x-only transform stack so w:w's ctx.scale/translate is visible.
  let scaleX = 1;
  let translateX = 0;
  const stack: { scaleX: number; translateX: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    get fontKerning() { return fontKerning; },
    set fontKerning(v: string) { fontKerning = v; },
    measureText: (s: string) => {
      const p = px();
      const w = [...s].length * p;
      return {
        width: w,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ scaleX, translateX }); },
    restore() { const s = stack.pop(); if (s) { scaleX = s.scaleX; translateX = s.translateX; } },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    scale(sx: number) { scaleX *= sx; },
    translate(tx: number) { translateX += tx; },
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, letterSpacing, fontKerning, scaleX, translateX });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FONT_PX, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...extra,
  };
}

type DocRun = DocParagraph['runs'][number];

function para(runs: DocxTextRun[]): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: runs.map((r) => ({ type: 'text', ...r }) as DocRun),
    defaultFontSize: FONT_PX, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function section(): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400, marginTop: 0, marginRight: 0, marginBottom: 0,
    marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
    evenAndOddHeaders: false, docGridCharSpace: undefined,
  } as SectionProps;
}

function doc(body: BodyElement[], settings?: DocSettings): DocxDocumentModel {
  return {
    section: section(), body, settings,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
}

async function render(runs: DocxTextRun[], settings?: DocSettings): Promise<{ runs: DocxTextRunInfo[]; fills: FillCall[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const info: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc([para(runs)], settings), canvas, 0, {
    dpr: 1, width: 600, onTextRun: (r) => info.push(r),
  });
  return { runs: info, fills };
}

/** The fillText call that drew a given text (first match). */
function drawOf(fills: FillCall[], text: string): FillCall {
  const f = fills.find((c) => c.text === text);
  expect(f, `a fillText drew ${JSON.stringify(text)}`).toBeDefined();
  return f as FillCall;
}

describe('WD4 run character metrics reach the glyph draw (measure==paint)', () => {
  it('w:spacing (§17.3.2.35) draws the run with ctx.letterSpacing = charSpacing px', async () => {
    // 1 pt char spacing at scale 1 (600px page over 600pt) = 1px per glyph.
    const { fills } = await render([textRun('WORD', { charSpacing: 1 })]);
    const d = drawOf(fills, 'WORD');
    expect(d.letterSpacing).toBe('1px');
    expect(d.scaleX).toBe(1); // no horizontal scale
  });

  it('a plain run (no w:spacing) draws with letterSpacing 0', async () => {
    const { fills } = await render([textRun('WORD')]);
    expect(drawOf(fills, 'WORD').letterSpacing).toBe('0px');
  });

  it('w:spacing widens the reported run box vs an identical run without it', async () => {
    const withSpacing = await render([textRun('WORD', { charSpacing: 2 })]);
    const without = await render([textRun('WORD')]);
    const w1 = withSpacing.runs[0].w;
    const w0 = without.runs[0].w;
    // 4 glyphs × 2px = 8px wider.
    expect(w1 - w0).toBeCloseTo(8, 5);
  });

  it('w:w (§17.3.2.43) draws under a horizontal ctx.scale and narrows the run box', async () => {
    const { runs, fills } = await render([textRun('WORD', { charScale: 0.5 })]);
    const d = drawOf(fills, 'WORD');
    expect(d.scaleX).toBeCloseTo(0.5, 5); // drawn at 50% width
    // Reported width is the natural width × 0.5 (4 × 20 × 0.5 = 40).
    expect(runs[0].w).toBeCloseTo(40, 5);
  });

  it('w:position (§17.3.2.24) shifts the baseline up for a positive (raised) value', async () => {
    const raised = await render([
      textRun('N'),
      textRun('X', { position: 6 }),
    ]); // +6 pt raised relative to the surrounding normal run
    const yRaised = drawOf(raised.fills, 'X').y;
    const yPlain = drawOf(raised.fills, 'N').y;
    // Raised text sits HIGHER ⇒ smaller y (canvas y grows downward). 6 pt × scale 1.
    expect(yPlain - yRaised).toBeCloseTo(6, 5);
    // The raised ink participates in the line extent, preventing a cell border
    // or the following line from crossing the painted glyph.
    expect(raised.runs[0].h).toBeCloseTo(FONT_PX + 6, 5);
  });

  it('w:position lowers the baseline for a negative value (mirrors raised)', async () => {
    const lowered = await render([
      textRun('N'),
      textRun('X', { position: -6 }),
    ]);
    expect(drawOf(lowered.fills, 'X').y - drawOf(lowered.fills, 'N').y).toBeCloseTo(6, 5);
    expect(lowered.runs[0].h).toBeCloseTo(FONT_PX + 6, 5);
  });

  it('centers uniformly positioned runs within their enlarged line box', async () => {
    // §17.3.2.24 defines position relative to surrounding non-positioned text.
    // With no differently-positioned peer on the line, nothing pins the extra
    // leading to one side, so the surplus is shared above and below the glyphs.
    const plain = await render([textRun('N')]);
    const uniformlyRaised = await render([
      textRun('A', { position: 6 }),
      textRun('B', { position: 6 }),
    ]);

    expect(uniformlyRaised.runs[0].h).toBeCloseTo(FONT_PX + 6, 5);
    expect(drawOf(uniformlyRaised.fills, 'A').y - drawOf(plain.fills, 'N').y)
      .toBeCloseTo(3, 5);
    expect(drawOf(uniformlyRaised.fills, 'B').y - drawOf(plain.fills, 'N').y)
      .toBeCloseTo(3, 5);
  });

  it('w:vertAlign raises superscript, lowers subscript, and leaves ordinary baselines unchanged', async () => {
    const { fills } = await render([
      textRun('N'),
      textRun('S', { vertAlign: 'super' }),
      textRun('B', { vertAlign: 'sub' }),
      textRun('Z'),
    ]);
    const normal = drawOf(fills, 'N');
    const superscript = drawOf(fills, 'S');
    const subscript = drawOf(fills, 'B');
    const trailingNormal = drawOf(fills, 'Z');

    expect(normal.y - superscript.y).toBeCloseTo(FONT_PX * 0.35, 5);
    expect(subscript.y - normal.y).toBeCloseTo(FONT_PX * 0.15, 5);
    expect(trailingNormal.y).toBeCloseTo(normal.y, 5);
  });

  it('w:vertAlign baseline shift composes with an authored w:position', async () => {
    const composed = await render([
      textRun('S', { vertAlign: 'super' }),
      textRun('P', { vertAlign: 'super', position: 4 }),
    ]);

    expect(
      drawOf(composed.fills, 'S').y - drawOf(composed.fills, 'P').y,
    ).toBeCloseTo(4, 5);
  });

  it('w:kern (§17.3.2.19) enables ctx.fontKerning when the run size ≥ the threshold', async () => {
    // fontSize FONT_PX=20pt, threshold 14pt ⇒ 20 ≥ 14 ⇒ kerning normal.
    const { fills } = await render([textRun('WORD', { kerning: 14 })]);
    expect(drawOf(fills, 'WORD').fontKerning).toBe('normal');
  });

  it('w:kern disables kerning when the run size is below the threshold', async () => {
    // threshold 28pt > 20pt run ⇒ kerning none.
    const { fills } = await render([textRun('WORD', { kerning: 28 })]);
    expect(drawOf(fills, 'WORD').fontKerning).toBe('none');
  });

  it('keeps authored w:kern thresholds authoritative for complex-script runs', async () => {
    const { fills } = await render([
      textRun('نص', { rtl: true, cs: true, fontSizeCs: FONT_PX, kerning: 14 }),
      textRun('عنوان', { rtl: true, cs: true, fontSizeCs: FONT_PX, kerning: 28 }),
    ]);

    expect(drawOf(fills, 'نص').fontKerning).toBe('normal');
    expect(drawOf(fills, 'عنوان').fontKerning).toBe('none');
  });

  it('disables kerning when w:kern is absent from the resolved style hierarchy', async () => {
    const { fills } = await render([textRun('WORD')]);
    expect(drawOf(fills, 'WORD').fontKerning).toBe('none');
  });

  it('enables absent-threshold kerning only under enableOpenTypeFeatures', async () => {
    const enabled = await render([textRun('WORD')], { enableOpenTypeFeatures: true });
    const disabled = await render([textRun('WORD')], { enableOpenTypeFeatures: false });
    expect(drawOf(enabled.fills, 'WORD').fontKerning).toBe('normal');
    expect(drawOf(disabled.fills, 'WORD').fontKerning).toBe('none');
  });

  it('keeps an authored w:kern threshold authoritative over enableOpenTypeFeatures', async () => {
    const { fills } = await render([textRun('WORD', { kerning: 28 })], {
      enableOpenTypeFeatures: true,
    });
    expect(drawOf(fills, 'WORD').fontKerning).toBe('none');
  });
});
