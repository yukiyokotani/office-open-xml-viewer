import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutLine, type LayoutSeg, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import { computePreparedLineFloatWindow, prepareFloatWrap } from '../layout/float-wrap.js';
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
function wrap(floats: FloatRect[], columnWidthPt = 360, rtl = false, columnXPt = 0): WrapLayoutCtx {
  return { floats, paraX: 0, startPageY: 0, columnXPt, columnWidthPt,
    readingDirection: rtl ? 'rtl' : 'ltr', pageH: 800, lineBoxH: () => 10 };
}
const fit = (fraction: number, reserveShare: number): LineFitWidth => ({ fraction, reserveShare, rtlVisiblePrefix: false });

function run(segs: LayoutSeg[], options: {
  maxWidth?: number; marginRightPx?: number; wrapCtx?: WrapLayoutCtx; baseRtl?: boolean; firstIndent?: number;
  tabStops?: { pos: number; alignment: 'left' | 'right' | 'start' | 'end'; leader: 'none' }[];
  lineFit?: LineFitWidth;
}): LayoutLine[] {
  const maxWidth = options.maxWidth ?? 360;
  return layoutLines(context(), segs, maxWidth, options.firstIndent ?? 0, 1, options.tabStops ?? [], options.wrapCtx, {}, 0,
    undefined, undefined, undefined, options.marginRightPx ?? maxWidth, options.baseRtl ?? false,
    true, false, undefined, 'bounded', undefined, true, false, options.lineFit);
}

describe('WORD_KASHIDA_LINE_FIT line breaker', () => {
  it('retries a rejected float gap until a window fits the advance', () => {
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
  it('retries past a far-edge window whose reserve runs past its edge', () => {
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
  it('searches every window with its own reserve after a far-edge rejection', () => {
    // A 300 pt paragraph with a 60 pt trailing indent and a 10% reserve: a
    // 100 pt far-edge window fits 84 pt; the next line's 105 pt window,
    // narrowed on its far side too, fits 94.5 pt and takes the 90 pt word.
    const floats = wrap([band(0, 200, 0, 10), band(0, 190, 10, 20), band(295, 300, 10, 20)]);
    const lines = run([token('W'.repeat(18))], {
      maxWidth: 300, marginRightPx: 360, wrapCtx: floats, lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => [line.topY, line.xOffset, line.availWidth])).toEqual([[10, 190, 105]]);
  });

  it('tests a reserve whose base ends before the band (negative trailing indent) forward', () => {
    // maxWidth 420 with the text margin at 360: a far-edge window's reserve
    // base is 60 pt shorter than the window. 270 pt fits 249 pt; 308 pt fits
    // 283.2 pt, which holds the 280 pt word.
    const floats = wrap([band(0, 150, 0, 10), band(0, 112, 10, 20)], 420);
    const lines = run([token('W'.repeat(56))], {
      maxWidth: 420, marginRightPx: 360, wrapCtx: floats, lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[10, 112]]);
  });

  it('admits a right-to-left right positional tab cell against the band', () => {
    // A 5 pt prefix, an indent-relative right ptab and a 300 pt cell in a
    // 360 pt highKashida band (fit 288 pt) share one line within the band.
    const ptab = { isTab: true, fontSize: 10, measuredWidth: 0, bold: false, italic: false,
      ptab: { alignment: 'right', relativeTo: 'indent' } } as unknown as LayoutSeg;
    const lines = run([token('A'), ptab, token('B'.repeat(60))], { baseRtl: true, lineFit: fit(0.8, 0) });
    expect(lines).toHaveLength(1);
    const total = lines[0]!.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0);
    expect(total).toBeLessThanOrEqual(360 + 1e-6);
    expect(total).toBeGreaterThan(300);
  });
  it('keeps a narrower later window that fits after a far-edge rejection', () => {
    // 300 pt paragraph, 60 pt trailing indent, 10% reserve: the 100 pt
    // far-edge window fits 84 pt; the next line's 95 pt window, narrowed at
    // its far edge, fits 85.5 pt and takes the 85 pt word.
    const floats = wrap([band(0, 200, 0, 10), band(0, 190, 10, 20), band(285, 300, 10, 20)]);
    const lines = run([token('W'.repeat(17))], {
      maxWidth: 300, marginRightPx: 360, wrapCtx: floats, lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => [line.topY, line.xOffset, line.availWidth])).toEqual([[10, 190, 95]]);
  });

  it('tests a hanging first line forward with its own window', () => {
    // A -50 pt hanging indent widens the first line to [-50, 300]. Floats
    // leave 90 pt and then 108 pt there: fits of 86 and 102.2 pt (0.1 of the
    // 40 and 58 pt windows). The 100 pt word takes the second.
    const floats = wrap([band(40, 400, 0, 10), band(58, 400, 10, 20)], 450, false, -50);
    const lines = run([token('W'.repeat(20))], {
      maxWidth: 300, marginRightPx: 300, firstIndent: -50, wrapCtx: floats, lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => line.topY)).toEqual([10]);
    expect(lines[0]!.availWidth).toBeCloseTo(58);
  });

  it('measures a separately formatted line-end space by the visible prefix in a gap', () => {
    // Right-to-left both: a 90 pt word and its glued 5 pt space fit a 92 pt
    // gap by the visible prefix; a separately formatted glued space follows.
    const visible: LineFitWidth = { fraction: 1, reserveShare: 0, rtlVisiblePrefix: true };
    const lines = run([token('W'.repeat(18)), token(' ', { joinPrev: true }), token(' ', { joinPrev: true, color: 'FF0000' }), token('next')], {
      baseRtl: true, wrapCtx: wrap([band(0, 268, 0, 10)], 360, true), lineFit: visible,
    });
    expect([lines[0]!.topY, lines[0]!.xOffset]).toEqual([0, 268]);
    expect(lines[0]!.segments.map((segment) => 'text' in segment ? segment.text : '').join('').trim()).toBe('W'.repeat(18));
  });

  it('finds the same window as a forward scan of every candidate (randomized)', () => {
    // Four 10 pt lines, each with one float leaving a single window at the
    // start side (interior: its far edge is the float) or at the far side
    // (reaching the band's far edge); a fifth free line. The retry must land
    // on the first window whose forward fit holds the word.
    let seed = 20081;
    const random = (): number => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    for (let trial = 0; trial < 1200; trial += 1) {
      // A third of the trials cut the reserve base short (negative trailing
      // indents beyond the window), where the reserve is zero.
      const trailing = trial % 3 === 0 ? -Math.round(100 + random() * 250) : Math.round(random() * 100) - 40;
      const indent = [0, 0, 20, 35][Math.floor(random() * 4)]!;
      const [fraction, share] = [[1, 0.1], [1, 0.2], [0.9, 0], [0.8, 0]][Math.floor(random() * 4)]!;
      // Window widths cluster around the word so that reserves of either kind
      // decide, and later windows are often narrower than rejected ones.
      const word = 5 * (12 + Math.floor(random() * 26));
      const windows = Array.from({ length: 4 }, () => ({
        width: Math.min(290, Math.round(word + indent + random() * 0.35 * (word + 60))), farEdge: random() < 0.5,
      }));
      const floats = windows.map(({ width, farEdge }, line) => farEdge
        ? band(0, 300 - width, line * 10, line * 10 + 10)
        : band(width, 300, line * 10, line * 10 + 10));
      const lines = run([token('W'.repeat(word / 5))], {
        maxWidth: 300, marginRightPx: 300 + trailing, firstIndent: indent, wrapCtx: wrap(floats), lineFit: fit(fraction, share),
      });
      const expected = windows.findIndex(({ width, farEdge }) =>
        (width - indent) * fraction - share * Math.max(0, width + (farEdge ? trailing : 0)) >= word - 1e-9);
      const label = JSON.stringify({ trial, trailing, indent, fraction, share, windows, word });
      if (expected < 0) {
        expect(lines[0]!.topY, label).toBe(40);
      } else {
        expect([lines[0]!.topY, lines[0]!.xOffset], label)
          .toEqual([expected * 10, windows[expected]!.farEdge ? 300 - windows[expected]!.width : 0]);
      }
    }
  }, 30_000);
  it('never lets a reserve base cut short by a negative indent widen the fit past the window', () => {
    // maxWidth 420, text margin 340: a 50 pt far-edge window [370, 420] has a
    // reserve base of 50 - 80 < 0, so it fits 50 pt, not 56: the 55 pt word
    // goes to the next line's 120 pt window [300, 420] (base 40, fit 112).
    const floats = wrap([band(0, 370, 0, 10), band(0, 300, 10, 20)], 420);
    const lines = run([token('W'.repeat(11))], {
      maxWidth: 420, marginRightPx: 340, wrapCtx: floats, lineFit: fit(1, 0.2),
    });
    expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[10, 300]]);
  });

  it('resumes excluded-window retries from the rejected window (bounded work)', () => {
    // 100 successive 10 pt bands each leave a 100 pt window; a 95 pt word
    // fits none of them at 0.8 and lands below them. The oracle counts the
    // band tops it scans: resuming keeps it linear, a restart is quadratic.
    let work = 0;
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      for (let top = input.topYPt; ; top += 10) {
        work += 1;
        if (top >= 1000) return { topYPt: top, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
        const excluded = input.excludedWindows?.some((window) =>
          window.topY === top && window.leftPt === input.paragraphXPt && window.widthPt === 100);
        if (100 >= input.minimumStartWidthPt && !excluded) return { topYPt: top, xOffsetPt: 0, maximumWidthPt: 100 };
      }
    };
    const lines = run([token('W'.repeat(19))], {
      wrapCtx: { ...wrap([]), floats: [], lineWindow }, lineFit: fit(0.8, 0),
    });
    expect(lines[0]!.topY).toBe(1000);
    expect(work).toBeLessThan(1000);
  });
  it('ends the retry budget on the band placement expands, never re-requesting a rejected window', () => {
    // A numbered first line with a -50 pt authored hanging indent expands the
    // 360 pt band to 410 pt. Every top offers a narrowed 390 pt window whose
    // 0.8 fit cannot hold the 320 pt word; past the retry budget the search
    // must leave the windows behind and place the word.
    let calls = 0;
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      calls += 1;
      if (calls > 5000) throw new Error('retry search did not terminate');
      if (input.minimumStartWidthPt > 390) {
        return { topYPt: 100000, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
      }
      for (let top = input.topYPt; ; top += 10) {
        const excluded = input.excludedWindows?.some((window) => window.topY === top);
        if (!excluded) return { topYPt: top, xOffsetPt: 20, maximumWidthPt: 390 };
      }
    };
    const wrapCtx: WrapLayoutCtx = { ...wrap([]), floats: [], lineWindow,
      numberingFirstLine: { authoredFirstIndentPx: -50, bodyIndentAt: (edge) => edge } };
    const lines = run([token('W'.repeat(64))], { wrapCtx, lineFit: fit(0.8, 0) });
    expect(calls).toBeLessThan(5000);
    expect(lines[0]!.topY).toBe(100000);
  });

  it('keeps the original line grid when a retry resumes below the request top', () => {
    // The breaker hands the solver the original request top as grid origin.
    const origins: Array<number | undefined> = [];
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      origins.push(input.gridOriginYPt);
      for (let top = input.topYPt; ; top += 5) {
        if (top >= 30) return { topYPt: top, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
        if (!input.excludedWindows?.some((window) => window.topY === top)) {
          return { topYPt: top, xOffsetPt: 0, maximumWidthPt: 100 };
        }
      }
    };
    run([token('W'.repeat(19))], { wrapCtx: { ...wrap([]), floats: [], lineWindow }, lineFit: fit(0.8, 0) });
    expect(origins.filter((origin) => origin !== undefined).every((origin) => origin === 0)).toBe(true);
    expect(origins.some((origin) => origin === 0)).toBe(true);
    // The solver steps a tight band on that grid: resumed at 5 with origin 0
    // it reaches the fitting window at 10, not 15.
    const tight: FloatRect = { ...obstacle(120, 360), authoredWrap: 'tight', imageY: 0, imageH: 1000, yTop: 0, yBottom: 1000,
      wrapPolygon: [{ xPt: 120, yPt: 0 }, { xPt: 360, yPt: 0 }, { xPt: 360, yPt: 1000 }, { xPt: 150, yPt: 1000 },
        { xPt: 150, yPt: 10 }, { xPt: 120, yPt: 10 }] };
    const prepared = prepareFloatWrap([tight]);
    const reference = { xLeftPt: 0, xRightPt: 360, readingDirection: 'ltr' as const };
    expect(computePreparedLineFloatWindow(5, 140, 10, 0, 360, prepared, 0, 360, reference, 140, [], 0).topY).toBe(10);
    expect(computePreparedLineFloatWindow(5, 140, 10, 0, 360, prepared, 0, 360, reference, 140, [], 5).topY).toBe(15);
  });

  it('measures a run of separately formatted line-end spaces with linear work', () => {
    const reads = (count: number): number => {
      let calls = 0;
      const ctx = { measureText: (text: string) => { calls += 1; return { width: text.length * 5,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }; } } as unknown as CanvasRenderingContext2D;
      const spaces = Array.from({ length: count }, (_, index) => token(' ', { joinPrev: true, color: index % 2 ? 'FF0000' : '0000FF' }));
      layoutLines(ctx, [token('W'.repeat(18)), ...spaces, token('next')], 360, 0, 1, [], wrap([band(0, 268, 0, 10)], 360, true), {}, 0,
        undefined, undefined, undefined, 360, true, true, false, undefined, 'bounded', undefined, true, false,
        { fraction: 1, reserveShare: 0, rtlVisiblePrefix: true });
      return calls;
    };
    const small = reads(400);
    const large = reads(800);
    expect(large).toBeLessThan(small * 2.5);
  });
});
