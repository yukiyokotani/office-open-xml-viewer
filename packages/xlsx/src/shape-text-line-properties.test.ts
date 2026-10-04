import { describe, expect, it, vi } from 'vitest';
import type { MathNode } from '@silurus/ooxml-core';

vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  rasterizeMathSvg: async () => ({}),
  tintMathRaster: () => ({}),
}));

const { drawShapeText, prepareWorksheetMath } = await import('./renderer.js');
type Worksheet = Parameters<typeof prepareWorksheetMath>[0];
type ShapeText = Parameters<typeof drawShapeText>[1];

const EMU_PER_PX = 9525;
const equation = (): MathNode[] => [{ kind: 'run', text: 'x', style: {} } as MathNode];

function recordingContext(): { ctx: CanvasRenderingContext2D; images: number[]; imageTops: number[]; texts: [string, number][] } {
  let font = '11px sans-serif';
  const images: number[] = [];
  const imageTops: number[] = [];
  const texts: [string, number][] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: 8 }) as TextMetrics,
    fillText(text: string, _x: number, y: number) { texts.push([text, y]); },
    drawImage(_image: unknown, x: number, y: number) { images.push(x); imageTops.push(y); },
    fillStyle: '#000',
    textBaseline: 'alphabetic',
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, images, imageTops, texts };
}

async function prepare(text: ShapeText): Promise<void> {
  await prepareWorksheetMath({ shapeGroups: [{ shapes: [{ text }] }] } as unknown as Worksheet, {
    loadMathJax: async () => {},
    mathMLToSvg: async () => ({ svg: '<svg/>', widthEm: 1, ascentEm: 0.8, descentEm: 0.2 }),
  } as unknown as Parameters<typeof prepareWorksheetMath>[1]);
}
const PT = 96 / 72;

describe('xlsx shape text line properties (parity with the pre-shared renderer)', () => {
  it('keeps a grapheme across authored font styles on one overwide line', () => {
    const text: ShapeText = {
      anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
      paragraphs: [{ align: 'l', runs: [
        { type: 'text', text: '\u1000', bold: false, italic: false, size: 20 },
        { type: 'text', text: '\uaa7c', bold: true, italic: false, size: 20 },
      ] }],
    } as ShapeText;
    const { ctx, texts } = recordingContext();
    ctx.measureText = (value) => ({ width: [...value].reduce((sum, ch) => sum + (ch === '\u1000' ? 20 : 0), 0),
      actualBoundingBoxAscent: 8 }) as TextMetrics;
    drawShapeText(ctx, text, 10, 100, 1);
    expect(texts.map(([value]) => value)).toEqual(['\u1000', '\uaa7c']);
    expect(texts[0][1]).toBe(texts[1][1]);
  });

  it('keeps the first-line indent on an inline equation when a later equation is display math', async () => {
    const inline = equation();
    const display = equation();
    const text: ShapeText = {
      anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
      paragraphs: [{
        align: 'l', indent: 20 * EMU_PER_PX,
        runs: [
          { type: 'math', nodes: inline, display: false },
          { type: 'math', nodes: display, display: true },
        ],
      }],
    } as ShapeText;
    const worksheet = { shapeGroups: [{ shapes: [{ text }] }] } as unknown as Worksheet;
    await prepareWorksheetMath(worksheet, {
      loadMathJax: async () => {},
      mathMLToSvg: async () => ({ svg: '<svg/>', widthEm: 1, ascentEm: 0.8, descentEm: 0.2 }),
    } as unknown as Parameters<typeof prepareWorksheetMath>[1]);

    const { ctx, images } = recordingContext();
    drawShapeText(ctx, text, 300, 300, 1);
    // The inline equation's line is the paragraph's first line (indent 20 px);
    // the display equation sits on its own line at the paragraph margin.
    expect(images).toEqual([20, 0]);
  });

  it('reserves an empty line before a display equation that opens a paragraph', async () => {
    const text: ShapeText = {
      anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
      paragraphs: [
        { align: 'l', runs: [{ type: 'text', text: 'ab', bold: false, italic: false, size: 20 }] },
        { align: 'l', runs: [{ type: 'math', nodes: equation(), display: true, fontSize: 10 }] },
      ],
    } as ShapeText;
    await prepare(text);
    const { ctx, imageTops } = recordingContext();
    drawShapeText(ctx, text, 300, 300, 1);
    // 'ab' (20pt) line, then the pending empty line of the equation's
    // paragraph (no preceding text in that paragraph: the 11pt default), then
    // the equation seated top-flush on its own line.
    expect(imageTops[0]).toBeCloseTo(20 * PT * 1.2 + 11 * PT * 1.2, 5);
  });

  it('sizes a line by a run that paints nothing, such as a trimmed terminal space', () => {
    const text: ShapeText = {
      anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
      paragraphs: [
        { align: 'l', runs: [
          { type: 'text', text: 'ab', bold: false, italic: false, size: 10 },
          { type: 'text', text: ' ', bold: false, italic: false, size: 30 },
        ] },
        { align: 'l', runs: [{ type: 'text', text: 'cd', bold: false, italic: false, size: 10 }] },
      ],
    } as ShapeText;
    const { ctx, texts } = recordingContext();
    drawShapeText(ctx, text, 300, 300, 1);
    const first = texts.find(([value]) => value === 'ab')![1];
    const second = texts.find(([value]) => value === 'cd')![1];
    // The first line is 30pt tall (the space run), the second 10pt.
    expect(second - first).toBeCloseTo((30 * PT * 1.2) / 2 + (10 * PT * 1.2) / 2, 5);
  });

  it('gives a leading empty line the default size, not a later run size', () => {
    const text: ShapeText = {
      anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
      paragraphs: [{ align: 'l', runs: [
        { type: 'break' },
        { type: 'text', text: 'ab', bold: false, italic: false, size: 30 },
      ] }],
    } as ShapeText;
    const { ctx, texts } = recordingContext();
    drawShapeText(ctx, text, 300, 300, 1);
    // Line 1 is empty at the 11pt default; 'ab' is centred in the 30pt line.
    expect(texts.find(([value]) => value === 'ab')![1]).toBeCloseTo(11 * PT * 1.2 + (30 * PT * 1.2) / 2, 5);
  });

  it('sizes the line after a line feed by the runs that still have text on it (L00, L02)', () => {
    // Excel controls L00 and L02 (issue #1562): the run containing the LF sizes
    // the next line only when it has text after the LF.
    const pitch = (runs: [string, number][]): number => {
      const text: ShapeText = {
        anchor: 't', wrap: 'square', lIns: 0, tIns: 0, rIns: 0, bIns: 0,
        paragraphs: [{ align: 'l', runs: runs.map(([value, size]) =>
          ({ type: 'text', text: value, bold: false, italic: false, size })) }],
      } as ShapeText;
      const { ctx, texts } = recordingContext();
      drawShapeText(ctx, text, 300, 300, 1);
      return texts[texts.length - 1][1] - texts[0][1];
    };
    // L02: a 40 pt run ending at the LF, then 14 pt text: the second line is 14 pt.
    expect(pitch([['Ab\n', 40], ['Cd', 14]])).toBeCloseTo((40 * PT * 1.2) / 2 + (14 * PT * 1.2) / 2, 5);
    // L00: text after the LF in the 40 pt run keeps the second line at 40 pt.
    expect(pitch([['Ab\nCd', 40], ['Ef', 14]])).toBeCloseTo(40 * PT * 1.2, 5);
  });
});
