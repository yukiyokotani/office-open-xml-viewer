import { describe, expect, it } from 'vitest';
import type { FloatRect } from './float-layout.js';
import { layoutLines, type LayoutSeg, type WrapLayoutCtx } from './line-layout.js';

function linearContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const size = () => Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1] ?? 10);
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * size() * 0.5,
      fontBoundingBoxAscent: size() * 0.8, fontBoundingBoxDescent: size() * 0.2,
      actualBoundingBoxAscent: size() * 0.8, actualBoundingBoxDescent: size() * 0.2,
    }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

const text = (value: string): LayoutSeg => ({
  text: value, bold: false, italic: false, underline: false, strikethrough: false,
  fontSize: 10, color: null, fontFamily: 'Times New Roman', vertAlign: null,
} as LayoutSeg);

/** A left square float over [0, 100) x [0, 50). */
const float = {
  kind: 'shape', mode: 'square', imageKey: 'f',
  imageX: 0, imageY: 0, imageW: 100, imageH: 50,
  xLeft: 0, xRight: 100, yTop: 0, yBottom: 50,
  side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 1,
} as FloatRect;

describe('producer indents beside floats', () => {
  it('measures a LibreOffice paragraph indent from the float edge, a DOCX one from the column', () => {
    // A 50pt left indent: the paragraph band is [50, 300) in a [0, 300) column.
    const wrap = (indentsFromFloats: boolean): WrapLayoutCtx => ({
      startPageY: 0, paraX: 50, floats: [float], columnXPt: 0, columnWidthPt: 300,
      lineBoxH: (ascent: number, descent: number) => ascent + descent, pageH: 100000,
      ...(indentsFromFloats ? { indentsFromFloats: true } : {}),
    } as WrapLayoutCtx);
    const start = (indentsFromFloats: boolean) => {
      const [line] = layoutLines(linearContext(), [text('word')], 250, 0, 1, [], wrap(indentsFromFloats), {}, 0);
      return 50 + line!.xOffset;
    };
    expect(start(false)).toBe(100);
    expect(start(true)).toBe(150);
  });
});
