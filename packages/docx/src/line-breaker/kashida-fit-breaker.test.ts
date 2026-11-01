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
function band(left: number, right: number, top: number, bottom: number): FloatRect {
  return { ...obstacle(left, right), imageY: top, imageH: bottom - top, yTop: top, yBottom: bottom };
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
    // kfit7 T4: a cell ending at a stop on the margin, whose twip arithmetic
    // differs from the column width by an ulp, still fits its band.
    const margin = run([token('A'.repeat(85)), tab(), token('B'.repeat(5))], {
      maxWidth: 11906 / 20 - 1134 / 20 - 1134 / 20, tabStops: [{ pos: 481.9, alignment: 'right', leader: 'none' }],
      lineFit: fit(0.9, 0),
    });
    expect(margin).toHaveLength(1);
    // A right-to-left end-aligned cell keeps its stop as well: 300 - 50 - 5.
    const [rtl] = run([token('A'), tab(), token('B'.repeat(10))], {
      baseRtl: true, tabStops: [{ pos: 300, alignment: 'end', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    expect(rtl!.segments.find((segment) => 'isTab' in segment)!.measuredWidth).toBeCloseTo(245);
  });
  it('retries with the exact inverse of a reserve that runs past a far-edge window', () => {
    // A 300 pt paragraph with a 60 pt trailing indent; floats at the start
    // side leave 100, 110 and 120 pt windows reaching the far edge on three
    // 10 pt lines. Before mode 15 each reserves 0.1 × (window + 60): fits of
    // 84, 93 and 102 pt. A 90 pt word enters the 110 pt window.
    const floats = wrap([band(0, 200, 0, 10), band(0, 190, 10, 20), band(0, 180, 20, 30)]);
    const lines = run([token('W'.repeat(18))], {
      maxWidth: 300, marginRightPx: 360, wrapCtx: floats, lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[10, 190]]);
  });

  it('admits a right-to-left aligned cell against the band during fitting', () => {
    // A 360 pt highKashida band (fit 288 pt): a 5 pt prefix and a 300 pt cell
    // ending at a 355 pt end stop stay on one line, the gap 355 - 300 - 5.
    const lines = run([token('A'), tab(), token('B'.repeat(60))], {
      baseRtl: true, tabStops: [{ pos: 355, alignment: 'end', leader: 'none' }], lineFit: fit(0.8, 0),
    });
    expect(lines).toHaveLength(1);
    expect(lines[0]!.segments.find((segment) => 'isTab' in segment)!.measuredWidth).toBeCloseTo(50);
  });

  it('admits a right positional tab and its cell against the band', () => {
    // A 5 pt prefix, an indent-relative right ptab and a 40 pt cell in a
    // 360 pt highKashida band: the cell ends at 360 on the same line.
    const ptab = { isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false,
      ptab: { alignment: 'right', relativeTo: 'indent' } } as unknown as LayoutSeg;
    const lines = run([token('A'), ptab, token('B'.repeat(8))], { lineFit: fit(0.8, 0) });
    expect(lines).toHaveLength(1);
    expect(lines[0]!.segments.find((segment) => 'isTab' in segment)!.measuredWidth).toBeCloseTo(315);
  });
});
