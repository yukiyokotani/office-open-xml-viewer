import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocxDocumentModel,
  SectionProps,
} from './types.js';

interface StrokeLine {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

interface FilledText {
  readonly text: string;
  readonly x: number;
  readonly y: number;
}

function recordingCanvas(inkFactor = 1): Readonly<{
  canvas: HTMLCanvasElement;
  strokes: StrokeLine[];
  texts: FilledText[];
}> {
  let font = '10px serif';
  const fontPx = () => Number(/([\d.]+)px/u.exec(font)?.[1] ?? 10);
  const strokes: StrokeLine[] = [];
  const texts: FilledText[] = [];
  let path: Array<Readonly<{ x: number; y: number }>> = [];
  const context = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const glyphs = [...text];
      const width = glyphs.length * fontPx();
      let lastInkIndex = -1;
      for (let index = glyphs.length - 1; index >= 0; index -= 1) {
        if (!/\s/u.test(glyphs[index]!)) {
          lastInkIndex = index;
          break;
        }
      }
      // Canvas advances through authored spaces, but their glyphs do not add
      // ink. Keep the preceding advance when a later visible glyph exists.
      const inkRight = lastInkIndex < 0
        ? 0
        : lastInkIndex * fontPx() + fontPx() * inkFactor;
      return {
        width,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: inkRight,
        actualBoundingBoxAscent: fontPx() * 0.8,
        actualBoundingBoxDescent: fontPx() * 0.2,
        fontBoundingBoxAscent: fontPx() * 0.8,
        fontBoundingBoxDescent: fontPx() * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() { path = []; }, closePath() {},
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() {
      for (let index = 1; index < path.length; index += 1) {
        strokes.push({
          x1: path[index - 1]!.x,
          y1: path[index - 1]!.y,
          x2: path[index]!.x,
          y2: path[index]!.y,
        });
      }
    },
    fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, drawImage() {},
    clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) { texts.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  return {
    canvas: {
      width: 0,
      height: 0,
      style: {} as CSSStyleDeclaration,
      getContext: () => context,
    } as unknown as HTMLCanvasElement,
    strokes,
    texts,
  };
}

function textRun(text: string, underline = true): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Test Mincho',
    fontFamilyEastAsia: 'Test Mincho',
    isLink: false,
    background: null,
    vertAlign: null,
    hyperlink: null,
  } as DocRun;
}

function document(
  text = 'あいう',
  options: Readonly<{
    runs?: DocRun[];
    bidi?: boolean;
    textDirection?: string;
  }> = {},
): DocxDocumentModel {
  const paragraph: DocParagraph = {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    widowControl: false,
    ...(options.bidi === undefined ? {} : { bidi: options.bidi }),
    defaultFontSize: 10,
    defaultFontFamily: 'Test Mincho',
    runs: options.runs ?? [textRun(text)],
  } as DocParagraph;
  return {
    section: {
      pageWidth: 40,
      pageHeight: 100,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      headerDistance: 0,
      footerDistance: 0,
      titlePage: false,
      evenAndOddHeaders: false,
      textDirection: options.textDirection ?? 'lrTb',
      docGridType: 'snapToChars',
      docGridLinePitch: 20,
      // The default 10pt font plus 10pt character spacing produces a 20pt cell.
      docGridCharSpace: 40960,
    } as SectionProps,
    body: [paragraph as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Test Mincho': 'roman' },
  } as unknown as DocxDocumentModel;
}

// Synthetic fixture matrix registered by WORD_SNAP_TO_CHARS_TERMINAL_UNDERLINE.
describe('snap-to-chars-terminal-underline-boundaries', () => {
  it('ends a soft-wrapped underline at the final glyph instead of the unused cell edge', async () => {
    const { canvas, strokes, texts } = recordingCanvas();
    await renderDocumentToCanvas(document(), canvas, 0, { dpr: 1, width: 40 });

    const horizontal = strokes.filter((line) => Math.abs(line.y1 - line.y2) < 1e-6);
    expect(texts.map(({ text, x }) => ({ text, x }))).toEqual([
      { text: 'あ', x: 5 },
      { text: 'い', x: 25 },
      { text: 'う', x: 5 },
    ]);
    expect(horizontal).toHaveLength(2);
    // Word retains the leading half-cell but ends the underline at the final
    // glyph instead of extending it through the trailing half-cell.
    expect(horizontal[0]!.x1).toBeCloseTo(0, 6);
    expect(horizontal[0]!.x2).toBeCloseTo(35, 6);
    expect(horizontal[1]!.x1).toBeCloseTo(0, 6);
    expect(horizontal[1]!.x2).toBeCloseTo(15, 6);
  });

  it('keeps one continuous rule across script segments before trimming the terminal cell', async () => {
    const { canvas, strokes, texts } = recordingCanvas();
    await renderDocumentToCanvas(document('Aあ'), canvas, 0, { dpr: 1, width: 40 });

    expect(texts.map(({ text, x }) => ({ text, x }))).toEqual([
      { text: 'A', x: 5 },
      { text: 'あ', x: 25 },
    ]);
    const horizontal = strokes.filter((line) => Math.abs(line.y1 - line.y2) < 1e-6);
    expect(horizontal).toHaveLength(1);
    expect(horizontal[0]!.x1).toBeCloseTo(0, 6);
    expect(horizontal[0]!.x2).toBeCloseTo(35, 6);
  });

  it('uses retained glyph ink instead of advance when the two metrics differ', async () => {
    const { canvas, strokes, texts } = recordingCanvas(0.6);
    await renderDocumentToCanvas(document('あ'), canvas, 0, { dpr: 1, width: 40 });

    expect(texts).toHaveLength(1);
    expect(texts[0]).toMatchObject({ text: 'あ', x: 5 });
    const horizontal = strokes.filter((line) => Math.abs(line.y1 - line.y2) < 1e-6);
    expect(horizontal).toHaveLength(1);
    expect(horizontal[0]!.x2).toBeGreaterThan(5);
    expect(horizontal[0]!.x2).toBeLessThan(15);
  });

  it('does not discard an explicitly authored terminal space', async () => {
    const plain = recordingCanvas();
    const spaced = recordingCanvas();
    await renderDocumentToCanvas(document('あ'), plain.canvas, 0, { dpr: 1, width: 80 });
    await renderDocumentToCanvas(document('あ '), spaced.canvas, 0, { dpr: 1, width: 80 });

    const end = (lines: readonly StrokeLine[]) => lines
      .filter((line) => Math.abs(line.y1 - line.y2) < 1e-6)[0]!.x2;
    expect(end(spaced.strokes)).toBeGreaterThan(end(plain.strokes));
  });

  it('trims each authored hard-break line at its final glyph', async () => {
    const { canvas, strokes } = recordingCanvas();
    await renderDocumentToCanvas(document('', {
      runs: [
        textRun('あい'),
        { type: 'break', breakType: 'line' } as DocRun,
        textRun('う'),
      ],
    }), canvas, 0, { dpr: 1, width: 80 });

    const horizontal = strokes.filter((line) => Math.abs(line.y1 - line.y2) < 1e-6);
    expect(horizontal).toHaveLength(2);
    expect(horizontal.map(line => line.x2).sort((a, b) => a - b)).toEqual([15, 35]);
  });

  it('trims only the terminal run of a continuous underline', async () => {
    const continuous = recordingCanvas(0.8);
    const stopped = recordingCanvas(0.8);
    await renderDocumentToCanvas(document('', {
      runs: [textRun('あ'), textRun('い')],
    }), continuous.canvas, 0, { dpr: 1, width: 80 });
    await renderDocumentToCanvas(document('', {
      runs: [textRun('あ'), textRun('い', false)],
    }), stopped.canvas, 0, { dpr: 1, width: 80 });

    const horizontal = (lines: readonly StrokeLine[]) => lines
      .filter((line) => Math.abs(line.y1 - line.y2) < 1e-6);
    expect(horizontal(continuous.strokes)).toHaveLength(1);
    expect(horizontal(continuous.strokes)[0]!.x2).toBeCloseTo(33, 6);
    expect(horizontal(stopped.strokes)).toHaveLength(1);
    expect(horizontal(stopped.strokes)[0]!.x2).toBeCloseTo(13, 6);
  });

  it('does not apply the LTR terminal trim to a bidi paragraph', async () => {
    const ltr = recordingCanvas(0.8);
    const rtl = recordingCanvas(0.8);
    await renderDocumentToCanvas(document('あ'), ltr.canvas, 0, { dpr: 1, width: 40 });
    await renderDocumentToCanvas(
      document('あ', { bidi: true }), rtl.canvas, 0, { dpr: 1, width: 40 },
    );

    const underlineLength = (lines: readonly StrokeLine[]) => {
      const line = lines.find(candidate => Math.abs(candidate.y1 - candidate.y2) < 1e-6)!;
      return Math.abs(line.x2 - line.x1);
    };
    expect(underlineLength(rtl.strokes)).toBeGreaterThan(underlineLength(ltr.strokes));
  });

  it('keeps vertical snap-to-grid decoration outside the horizontal trim rule', async () => {
    const horizontal = recordingCanvas(0.8);
    const vertical = recordingCanvas(0.8);
    await renderDocumentToCanvas(
      document('あ'), horizontal.canvas, 0, { dpr: 1, width: 40 },
    );
    await renderDocumentToCanvas(
      document('あ', { textDirection: 'tbRl' }), vertical.canvas, 0, { dpr: 1, width: 40 },
    );

    const maxLength = (lines: readonly StrokeLine[]) => Math.max(...lines.map(line =>
      Math.hypot(line.x2 - line.x1, line.y2 - line.y1)));
    expect(maxLength(vertical.strokes)).toBeGreaterThan(maxLength(horizontal.strokes));
  });
});
