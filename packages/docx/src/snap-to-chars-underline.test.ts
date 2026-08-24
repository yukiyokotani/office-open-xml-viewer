import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
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

function recordingCanvas(): Readonly<{
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
      const width = [...text].length * fontPx();
      return {
        width,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: width,
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

function document(text = 'あいう'): DocxDocumentModel {
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
    defaultFontSize: 10,
    defaultFontFamily: 'Test Mincho',
    runs: [{
      type: 'text',
      text,
      bold: false,
      italic: false,
      underline: true,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'Test Mincho',
      fontFamilyEastAsia: 'Test Mincho',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    }],
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

describe('snapToChars underline geometry', () => {
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
});
