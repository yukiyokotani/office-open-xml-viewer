import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import type { FloatRect } from '../layout/float-wrap.js';

// WORD_GRID_LINE_CONTENT_FIT_ADMISSION (compatibility mode 15): a snapped grid
// line meets floats with its content box; a topAndBottom float that reaches
// only the line's trailing grid leading leaves the line where it is.

const context = { measureText: (text: string) => ({ width: text.length * 5,
  fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
const token = (text: string): LayoutTextSeg => ({ text, fontSize: 10, fontFamily: 'Arial', bold: false,
  italic: false, underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0 });
const float = (top: number): FloatRect => ({ kind: 'shape', mode: 'topAndBottom', authoredWrap: 'topAndBottom',
  imageKey: 'test', imageX: 0, imageY: top, imageW: 100, imageH: 20, xLeft: 0, xRight: 100,
  yTop: top, yBottom: top + 20, side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 });
// A 20pt cell around a 10pt line: 5pt of grid leading above and below.
const wrap = (top: number, excludeLeading: boolean): WrapLayoutCtx => ({
  floats: [float(top)], paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 360,
  readingDirection: 'ltr', pageH: 800, lineBoxH: () => 20,
  ...(excludeLeading ? { probeTrailingLeading: (_line, advance) => (advance - 10) / 2 } : {}),
});

describe('mode-15 grid line wrap band', () => {
  it('keeps a line whose trailing leading alone reaches a topAndBottom float', () => {
    expect(layoutLines(context, [token('Hello')], 360, 0, 1, [], wrap(17, true))[0]!.topY).toBe(0);
  });
  it('pushes the same line below the float when the band is the whole cell', () => {
    expect(layoutLines(context, [token('Hello')], 360, 0, 1, [], wrap(17, false))[0]!.topY).toBe(37);
  });
  it('still pushes a line whose content box reaches the float', () => {
    expect(layoutLines(context, [token('Hello')], 360, 0, 1, [], wrap(14, true))[0]!.topY).toBe(34);
  });
});
