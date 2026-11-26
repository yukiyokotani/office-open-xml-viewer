import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import type { FloatRect } from '../layout/float-wrap.js';

// A line is float-conditioned (LayoutLine.wrapAffected) only when a float
// narrows, offsets or moves it; an active wrap context alone does not make
// it so (the exact-line object rule keys on this, WORD_INLINE_OBJECT_EXACT_LINE).

const context = { measureText: (text: string) => ({ width: text.length * 5,
  fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
const token = (text: string): LayoutTextSeg => ({ text, fontSize: 10, fontFamily: 'Arial', bold: false,
  italic: false, underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0 });
function band(left: number, right: number, top: number, bottom: number): FloatRect {
  return { kind: 'shape', mode: 'square', authoredWrap: 'square', imageKey: 'test',
    imageX: left, imageY: top, imageW: right - left, imageH: bottom - top,
    xLeft: left, xRight: right, yTop: top, yBottom: bottom, side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 };
}
const wrap = (floats: FloatRect[]): WrapLayoutCtx => ({ floats, paraX: 0, startPageY: 0, columnXPt: 0,
  columnWidthPt: 360, readingDirection: 'ltr', pageH: 800, lineBoxH: () => 10 });

describe('LayoutLine.wrapAffected', () => {
  it('leaves a line a distant float does not touch unconditioned', () => {
    const [line] = layoutLines(context, [token('Hello')], 360, 0, 1, [], wrap([band(0, 200, 500, 560)]));
    expect(line!.topY).toBeDefined();
    expect(line!.wrapAffected).toBeUndefined();
  });
  it('marks a line a float narrows', () => {
    const [line] = layoutLines(context, [token('Hello')], 360, 0, 1, [], wrap([band(0, 200, 0, 30)]));
    expect(line!.wrapAffected).toBe(true);
  });
});
