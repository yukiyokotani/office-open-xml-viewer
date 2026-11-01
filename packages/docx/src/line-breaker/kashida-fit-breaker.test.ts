import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutLine, type LayoutSeg, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import type { FloatRect } from '../layout/float-wrap.js';
import type { LineFitWidth } from './pass-driver.js';

// WORD_KASHIDA_LINE_FIT inside the line breaker: the reserve narrows the fit
// width only, so every other width the breaker compares with a fitted advance
// (float-gap retries, right-to-left tab cells, aligned margin cells) must use
// the same fit width, while bands and stops keep their raw geometry.

/** Every character is 5 pt wide. */
function context(): CanvasRenderingContext2D {
  return { measureText: (text: string) => ({ width: text.length * 5,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
}
function token(text: string, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg {
  return { text, fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0, ...extra };
}
const tab = (): LayoutSeg =>
  ({ isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false }) as unknown as LayoutSeg;
function obstacle(left: number, right: number): FloatRect {
  return { kind: 'shape', mode: 'square', authoredWrap: 'square', imageKey: 'test',
    imageX: left, imageY: 0, imageW: right - left, imageH: 60,
    xLeft: left, xRight: right, yTop: 0, yBottom: 60, side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 };
}
function wrap(floats: FloatRect[]): WrapLayoutCtx {
  return { floats, paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 360,
    readingDirection: 'ltr', pageH: 800, lineBoxH: () => 10 };
}
const fit = (fraction: number, reserveShare: number): LineFitWidth => ({ fraction, reserveShare, rtlVisiblePrefix: false });

function run(segs: LayoutSeg[], options: {
  maxWidth?: number; marginRightPx?: number; wrapCtx?: WrapLayoutCtx; baseRtl?: boolean;
  tabStops?: { pos: number; alignment: 'left' | 'right' | 'start' | 'end'; leader: 'none' }[];
  lineFit?: LineFitWidth;
}): LayoutLine[] {
  const maxWidth = options.maxWidth ?? 360;
  return layoutLines(context(), segs, maxWidth, 0, 1, options.tabStops ?? [], options.wrapCtx, {}, 0,
    undefined, undefined, undefined, options.marginRightPx ?? maxWidth, options.baseRtl ?? false,
    true, false, undefined, 'bounded', undefined, true, false, options.lineFit);
}

describe('WORD_KASHIDA_LINE_FIT line breaker', () => {
  it('retries a rejected float gap with the band width the fitted advance needs', () => {
    // Gaps of 100 pt ([0, 100]) and 120 pt ([160, 280]) beside two floats.
    const floats = wrap([obstacle(100, 160), obstacle(280, 360)]);
    // A 10% reserve of the gap before mode 15 (blk-0898-kfit7 P1/P2: the
    // float edge is the far edge) and 0.9 of it from mode 15: the 95 pt word
    // misses the first gap (90 pt) and fits the second (108 pt).
    for (const lineFit of [fit(1, 0.1), fit(0.9, 0)]) {
      const lines = run([token('W'.repeat(19))], { wrapCtx: floats, lineFit });
      expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[0, 160]]);
    }
  });

  it('reserves a share of a float-narrowed window, not of the column, before mode 15', () => {
    // One float leaves [0, 200] of a 360 pt column: 0.1 × 200 = 20 pt
    // reserved (180 pt fit), not 36 pt (164 pt): a 170 pt word fits.
    const lines = run([token('W'.repeat(34))], { wrapCtx: wrap([obstacle(200, 360)]), lineFit: fit(1, 0.1) });
    expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[0, 0]]);
  });

  it('ends a right-to-left tab cell within the fit width', () => {
    // A start stop 300 pt from the right margin, then 50 pt of text, in a
    // 360 pt highKashida line (fit 288 pt): the gap shrinks to 238 pt.
    const [line] = run([tab(), token('B'.repeat(10))], {
      baseRtl: true, tabStops: [{ pos: 300, alignment: 'start', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    const gap = line!.segments.find((segment) => 'isTab' in segment)!;
    expect(gap.measuredWidth).toBeCloseTo(238);
    // Without a reserve the stop is reached.
    const [plain] = run([tab(), token('B'.repeat(10))], {
      baseRtl: true, tabStops: [{ pos: 300, alignment: 'start', leader: 'none' }],
    });
    expect(plain!.segments.find((segment) => 'isTab' in segment)!.measuredWidth).toBeCloseTo(300);
  });

  it('admits an aligned tab cell against the band, not the fit width', () => {
    // blk-0898-kfit7 T2: a right tab cell ending at 295 pt stays on a 360 pt
    // highKashida line whose text fits 288 pt.
    const cell = run([token('A'), tab(), token('B'.repeat(8))], {
      tabStops: [{ pos: 295, alignment: 'right', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    expect(cell).toHaveLength(1);
    // Text before the tab keeps the fit width (kfit7 T4): 305 pt of words do
    // not fit 288 pt, so the second word and the cell move to the next line.
    const lead = run([token(`${'A'.repeat(30)} `), token(`${'B'.repeat(30)} `), tab(), token('C')], {
      tabStops: [{ pos: 355, alignment: 'right', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    expect(lead.map((line) => line.segments.filter((segment) => 'text' in segment).map((segment) => (segment as LayoutTextSeg).text.trim()).join('|')))
      .toEqual(['A'.repeat(30), `${'B'.repeat(30)}|C`]);
    // A right-to-left end-aligned cell keeps its stop as well: 300 - 50 - 5.
    const [rtl] = run([token('A'), tab(), token('B'.repeat(10))], {
      baseRtl: true, tabStops: [{ pos: 300, alignment: 'end', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    expect(rtl!.segments.find((segment) => 'isTab' in segment)!.measuredWidth).toBeCloseTo(245);
  });
});
