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

function makeRecordingCanvas(advance?: (text: string, size: number) => number): { canvas: HTMLCanvasElement; fills: FillCall[] } {
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
      const w = advance?.(s, p) ?? [...s].length * p - (fontKerning === 'normal' ?
        (s.includes('AV') ? 2 : 0) + (s.includes('T ') ? 1 : 0) + (s.includes(' A') ? 1.5 : 0) : 0);
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

function para(runs: DocxTextRun[], alignment: DocParagraph['alignment'] = 'left'): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
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

async function render(runs: DocxTextRun[], settings?: DocSettings, alignment?: DocParagraph['alignment']): Promise<{ runs: DocxTextRunInfo[]; fills: FillCall[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const info: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(doc([para(runs, alignment)], settings), canvas, 0, {
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
    expect(drawOf(uniformlyRaised.fills, 'AB').y - drawOf(plain.fills, 'N').y)
      .toBeCloseTo(3, 5);
    expect(drawOf(uniformlyRaised.fills, 'AB').y - drawOf(plain.fills, 'N').y)
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

  it('kerns Arabic complex-script runs whatever their w:kern threshold', async () => {
    // Windows Word 16.0 #1850 controls: WORD_ARABIC_PAIR_KERNING.
    const { fills } = await render([
      textRun('نص', { rtl: true, cs: true, fontSizeCs: FONT_PX, kerning: 14 }),
      textRun('عنوان', { rtl: true, cs: true, fontSizeCs: FONT_PX, kerning: 28 }),
    ]);

    expect(drawOf(fills, 'نص').fontKerning).toBe('normal');
    expect(drawOf(fills, 'عنوان').fontKerning).toBe('normal');
  });

  // CAL-K measured sizes/thresholds plus K1/K4 boundaries. Both flag values
  // accept positive qualifying thresholds and reject absent/zero/too-large.
  it.each([
    { size: 8, threshold: undefined, expected: 'none' },
    { size: 8, threshold: 0, expected: 'none' },
    { size: 8, threshold: 8, expected: 'normal' },
    { size: 12, threshold: 8, expected: 'normal' },
    { size: 18, threshold: 18, expected: 'normal' },
    { size: 18, threshold: 20, expected: 'none' },
    { size: 20, threshold: 20, expected: 'normal' },
    { size: 18, threshold: 8, expected: 'normal' },
    { size: 18, threshold: undefined, expected: 'none' },
  ])('keeps size $size threshold $threshold authoritative for measurement and paint', async ({ size, threshold, expected }) => {
    for (const enableOpenTypeFeatures of [false, true]) {
      const { fills, runs } = await render([textRun('AV', { fontSize: size, kerning: threshold })], {
        compatibilityMode: 15, enableOpenTypeFeatures,
      });
      expect(drawOf(fills, 'AV').fontKerning).toBe(expected);
      expect(runs[0].w).toBe(expected === 'normal' ? 2 * size - 2 : 2 * size);
    }
  });

  it('paints same-format letter and space seams as their concatenated sequence', async () => {
    const settings = { compatibilityMode: 15 };
    for (const parts of [['T', ' beyond'], ['A', 'V']]) {
      const split = await render(parts.map(text => textRun(text, { kerning: 8 })), settings);
      const whole = await render([textRun(parts.join(''), { kerning: 8 })], settings);
      expect(split.fills).toEqual(whole.fills);
    }
    const changed = await render([textRun('T', { kerning: 8 }), textRun(' beyond', { kerning: 8, charSpacing: 1 })], settings);
    expect(drawOf(changed.fills, ' ').x - drawOf(changed.fills, 'T').x).toBe(FONT_PX);
    for (const kerning of [undefined, 0, 28]) {
      const split = await render([textRun('T', { kerning }), textRun(' beyond', { kerning })], settings);
      expect(split.fills).toEqual((await render([textRun('T beyond', { kerning })], settings)).fills);
    }
  });

  it.each(['right', 'center'] as const)('compares source formatting by value for %s alignment', async (alignment) => {
    const first = textRun('T', { kerning: 8 });
    const second = textRun(' X', { kerning: 8 });
    const reversed = Object.fromEntries(Object.entries(second).reverse()) as unknown as DocxTextRun;
    const settings = { compatibilityMode: 15 };
    const ordinary = await render([first, second], settings, alignment);
    const reordered = await render([first, reversed], settings, alignment);
    expect(ordinary.fills).toEqual((await render([textRun('T X', { kerning: 8 })], settings, alignment)).fills);
    expect(reordered.runs).toEqual(ordinary.runs);
    expect(reordered.fills).toEqual(ordinary.fills);
    const changed = await render([first, { ...reversed, italic: true }], settings, alignment);
    expect(drawOf(changed.fills, ' ').x - drawOf(changed.fills, 'T').x).toBe(FONT_PX);
  });

  it.each([14, undefined, 16, 15])('disables a zero threshold while source splits stay transparent (mode %s)', async (compatibilityMode) => {
    const settings = { compatibilityMode, enableOpenTypeFeatures: true };
    // WORD_KERN_THRESHOLD_AUTHORITY: measured with the mode omitted and 15.
    const zero = await render([textRun('AV', { kerning: 0 })], settings);
    expect(drawOf(zero.fills, 'AV').fontKerning).toBe('none');
    expect(zero.runs[0].w).toBe(2 * FONT_PX);
    const seam = await render([textRun('T', { kerning: 8 }), textRun(' X', { kerning: 8 })], settings);
    expect(seam.fills).toEqual((await render([textRun('T X', { kerning: 8 })], settings)).fills);
    // Absence and positive size boundaries are normative in every mode.
    for (const kerning of [undefined, FONT_PX, FONT_PX + 1]) {
      const result = await render([textRun('AV', { kerning })], settings);
      expect(drawOf(result.fills, 'AV').fontKerning).toBe(kerning === FONT_PX ? 'normal' : 'none');
      expect(result.runs[0].w).toBe(kerning === FONT_PX ? 2 * FONT_PX - 2 : 2 * FONT_PX);
    }
  });
});

describe('kerning authority and justified fitting share measured advances', () => {
  // Word 16.113.3 mode-15 K3/K4 controls, with the independently pinned
  // Times New Roman cmap/hmtx advances (2048 upm). These scalar metrics isolate
  // zero/absent kerning from the fitting decision; they are not fitted widths.
  const advances: Record<string, number> = {
    ' ': 512, A: 1479, E: 1251, L: 1251, R: 1366, T: 1251, U: 1479,
    V: 1479, W: 1933, b: 1024, d: 1024, e: 909, n: 1024, o: 1024, y: 1024,
  };
  async function lines(prefix: string, kerning: number | undefined, width: number): Promise<string[]> {
    const { canvas } = makeRecordingCanvas((text, size) =>
      [...text].reduce((sum, ch) => sum + advances[ch]!, 0) * size / 2048);
    const model = doc([para([
      textRun(prefix, { fontSize: 18, kerning }),
      textRun(' beyond', { fontSize: 18, kerning }),
    ], 'both')], {
      compatibilityMode: 15, enableOpenTypeFeatures: true,
      characterSpacingControl: 'compressPunctuation',
    });
    model.section.pageWidth = width;
    const result = new Map<number, string>();
    await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width, onTextRun(run) {
      result.set(run.y, (result.get(run.y) ?? '') + run.text);
    } });
    return [...result.values()].map(text => text.trim());
  }
  const controls = [
    { prefix: 'AVATAR To WAVE VAULT', kerning: 0, rejectWidth: 217.10, acceptWidth: 220.35,
      rejected: ['AVATAR To WAVE', 'VAULT beyond'] },
    { prefix: 'AVATAR To WAVE VAULT AVATAR To WAVE VAULT To', kerning: undefined, rejectWidth: 464.10, acceptWidth: 467.35,
      rejected: ['AVATAR To WAVE VAULT AVATAR To WAVE VAULT', 'To beyond'] },
  ];
  it.each(controls)('preserves the Word rejection below the fitting boundary (threshold $kerning)', async ({ prefix, kerning, rejectWidth, rejected }) => {
    expect(await lines(prefix, kerning, rejectWidth)).toEqual(rejected);
  });
  it.each(controls)('accepts the Word prefix above the compression boundary (threshold $kerning)', async ({ prefix, kerning, acceptWidth }) => {
    expect(await lines(prefix, kerning, acceptWidth)).toEqual([prefix, 'beyond']);
  });
});
