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

  it('asks the line-window oracle once per rejection with each kind\'s requirement', () => {
    // A numbered first line with a -50 pt authored hanging indent: every top
    // offers a narrowed 390 pt window, too small for the 320 pt word at 0.8
    // (it needs 450 pt). The forward-tested search leaves them in one query.
    let calls = 0;
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      calls += 1;
      if (calls > 50) throw new Error('search did not terminate');
      // Like the production oracle, it skips windows the forward test refuses.
      for (let top = input.topYPt; top < 1000; top += 10) {
        if (input.minimumStartWidthPt <= 390
          && (input.acceptWindow?.({ topY: top, xOffset: 20, maxWidth: 390 }) ?? true)) {
          return { topYPt: top, xOffsetPt: 20, maximumWidthPt: 390 };
        }
      }
      return { topYPt: 100000, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
    };
    const wrapCtx: WrapLayoutCtx = { ...wrap([]), floats: [], lineWindow,
      numberingFirstLine: { authoredFirstIndentPx: -50, bodyIndentAt: (edge) => edge } };
    const lines = run([token('W'.repeat(64))], { wrapCtx, lineFit: fit(0.8, 0) });
    expect(lines[0]!.topY).toBe(100000);
    expect(calls).toBeLessThan(10);
  });

  it('searches the active baseline with each window\'s own requirement', () => {
    // A 600 pt band, three floats leaving gaps [0, 100], [150, 250] and
    // [300, 450] on the first lines. A short word fills the first gap; the
    // continuation's 95 pt word at 0.8 needs 118.75 pt, misses the second
    // gap and takes the 150 pt third gap on the same baseline.
    const floats = wrap([band(100, 150, 0, 60), band(250, 300, 0, 60), band(450, 600, 0, 60)], 600);
    const lines = run([token('AAAA '), token('W'.repeat(19))], { maxWidth: 600, wrapCtx: floats, lineFit: fit(0.8, 0) });
    expect(lines.map((line) => [line.topY, line.xOffset, line.availWidth])).toEqual([[0, 0, 100], [0, 300, 150]]);
  });

  it('finds a sloped contour\'s fitting window at its exact threshold', () => {
    // A through-wrapped contour on the far side widens the gap from 100 pt at
    // y = 0 to 200 pt at y = 60; a 95 pt word at 0.8 needs 118.75 pt.
    const sloped: FloatRect = { ...obstacle(100, 360), authoredWrap: 'through', imageY: 0, imageH: 60, yTop: 0, yBottom: 60,
      wrapPolygon: [{ xPt: 100, yPt: 0 }, { xPt: 360, yPt: 0 }, { xPt: 360, yPt: 60 }, { xPt: 200, yPt: 60 }] };
    const lines = run([token('W'.repeat(19))], { wrapCtx: wrap([sloped]), lineFit: fit(0.8, 0) });
    const raw = computePreparedLineFloatWindow(0, 118.75, 10, 0, 360, prepareFloatWrap([sloped]), 0, 360,
      { xLeftPt: 0, xRightPt: 360, readingDirection: 'ltr' }, 118.75);
    expect(lines[0]!.topY).toBeCloseTo(raw.topY, 9);
    expect(lines[0]!.topY).toBeLessThan(20);
    expect(lines[0]!.availWidth * 0.8).toBeGreaterThanOrEqual(95 - 1e-9);
  });

  it('matches a fine scan of sloped contours, inside and past the band (randomized)', () => {
    let seed = 7717;
    const random = (): number => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    for (let trial = 0; trial < 80; trial += 1) {
      // Every other trial uses a 300 pt band in a 400 pt column, so contours
      // and their openings run past the band's far edge.
      const band = trial % 2 === 0 ? 360 : 300;
      const span = band === 360 ? 360 : 400;
      const top = 60 + Math.round(random() * 120);
      const bottom = 140 + Math.round(random() * 160);
      const height = 40 + Math.round(random() * 80);
      const startSide = random() < 0.5;
      // A contour whose free side widens linearly from `top` to `bottom`.
      const polygon = startSide
        ? [{ xPt: 0, yPt: 0 }, { xPt: span - top, yPt: 0 }, { xPt: span - bottom, yPt: height }, { xPt: 0, yPt: height }]
        : [{ xPt: top, yPt: 0 }, { xPt: span, yPt: 0 }, { xPt: span, yPt: height }, { xPt: bottom, yPt: height }];
      const float: FloatRect = { ...obstacle(0, span), authoredWrap: 'through', imageY: 0, imageH: height, yTop: 0, yBottom: height,
        wrapPolygon: polygon };
      const trailing = Math.round(random() * 80) - 20;
      const [fraction, share] = [[1, 0.1], [1, 0.2], [0.9, 0], [0.8, 0]][Math.floor(random() * 4)]!;
      const word = 5 * (12 + Math.floor(random() * 20));
      const lines = run([token('W'.repeat(word / 5))], {
        maxWidth: band, marginRightPx: band + trailing, wrapCtx: wrap([float], span), lineFit: fit(fraction, share),
      });
      const reference = { xLeftPt: 0, xRightPt: band, readingDirection: 'ltr' as const };
      const fitted = (window: { xOffset: number; maxWidth: number }): number => window.maxWidth * fraction
        - share * Math.max(0, window.maxWidth + (window.xOffset + window.maxWidth >= band - 1e-9 ? trailing : 0));
      // Fine scan: the first top (0.01 pt steps) whose window fits the word.
      const prepared = prepareFloatWrap([float]);
      let expected = height;
      for (let y = 0; y < height; y += 0.01) {
        const window = computePreparedLineFloatWindow(y, 1, 10, 0, band, prepared, 0, span, reference, 1);
        if (window.topY === y && fitted(window) >= word) { expected = y; break; }
      }
      const label = JSON.stringify({ trial, band, top, bottom, height, startSide, trailing, fraction, share, word });
      const got = lines[0]?.topY ?? Number.NaN;
      expect(Math.abs(got - expected), label).toBeLessThanOrEqual(0.011);
    }
  }, 60_000);

  it('tests a numbered first line\'s windows with the body indent each resolves', () => {
    // Band [0, 600], floats at 100-150, 250-300 and 450-600; authored indent
    // -50, resolved body indent 0, a 5 pt marker and default tabs every 36.
    // In [300, 450] the body starts at 324: (150 - 24) x 0.8 = 100.8 >= 95.
    const floats = wrap([band(100, 150, 0, 60), band(250, 300, 0, 60), band(450, 600, 0, 60)], 600);
    const wrapCtx: WrapLayoutCtx = { ...floats, numberingFirstLine: {
      authoredFirstIndentPx: -50, bodyIndentAt: (edge) => Math.ceil((edge + 5) / 36) * 36 } };
    const lines = run([token('W'.repeat(19))], { maxWidth: 600, marginRightPx: 600, wrapCtx, lineFit: fit(0.8, 0) });
    expect([lines[0]!.topY, lines[0]!.xOffset]).toEqual([0, 300]);
  });

  it('gives a through-wrap opening that runs past the band its far-edge root', () => {
    // Band [0, 300] in a 360 pt column, text margin at 360; one through float
    // whose notch opens from x = 200 - y to 350: the clipped far-edge window
    // is 100 + y wide and fits 18 W (90 pt) at a 10% reserve of (window + 60)
    // from y = 6.667.
    const through: FloatRect = { ...obstacle(0, 400), authoredWrap: 'through', imageY: 0, imageH: 60, yTop: 0, yBottom: 60,
      wrapPolygon: [{ xPt: 0, yPt: 0 }, { xPt: 200, yPt: 0 }, { xPt: 150, yPt: 50 }, { xPt: 350, yPt: 50 },
        { xPt: 350, yPt: 0 }, { xPt: 400, yPt: 0 }, { xPt: 400, yPt: 60 }, { xPt: 0, yPt: 60 }] };
    const lines = run([token('W'.repeat(18))], { maxWidth: 300, marginRightPx: 360, wrapCtx: wrap([through]), lineFit: fit(1, 0.1) });
    expect(lines[0]!.topY).toBeGreaterThan(6.6);
    expect(lines[0]!.topY).toBeLessThanOrEqual(7);
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
  // The resolved fit of a numbered first line in the solver's window at y:
  // window [left, 360] (paraX 0), body at max(left, bodyIndentAt(left)).
  const numberedScan = (contour: FloatRect, hanging: number, height: number,
    bodyIndentAt: (edge: number) => number, fraction: number, share: number, word: number): number => {
    const prepared = prepareFloatWrap([contour]);
    const reference = { xLeftPt: hanging, xRightPt: 360, readingDirection: 'ltr' as const };
    for (let y = 0; y < height; y += 0.01) {
      const window = computePreparedLineFloatWindow(y, 1, 10, hanging, 360 - hanging, prepared, hanging, 360, reference, 1);
      if (window.topY !== y) continue;
      const left = hanging + window.xOffset;
      if (left < 0) continue;
      const indent = Math.max(left, bodyIndentAt(left)) - left;
      if ((window.maxWidth - indent) * fraction - share * window.maxWidth >= word) return y;
    }
    return height;
  };

  it('stops a numbered first line on a sloped contour where its resolved fit first holds', () => {
    // A through contour on the start side recedes as y grows; the body of
    // the numbered first line starts at the next 36 pt stop after the
    // window's start + 5, so the 95 pt word at 0.8 fits only once that stop
    // is low enough, between plain-width roots.
    const contour: FloatRect = { ...obstacle(-50, 360), authoredWrap: 'through', imageY: 0, imageH: 60, yTop: 0, yBottom: 60,
      xLeft: -50, imageX: -50,
      wrapPolygon: [{ xPt: -50, yPt: 0 }, { xPt: 300, yPt: 0 }, { xPt: 100, yPt: 60 }, { xPt: -50, yPt: 60 }] };
    const bodyIndentAt = (edge: number): number => Math.ceil((edge + 5) / 36) * 36;
    const wrapCtx: WrapLayoutCtx = { ...wrap([contour], 410, false, -50), numberingFirstLine: {
      authoredFirstIndentPx: -50, bodyIndentAt } };
    const lines = run([token('W'.repeat(19))], { wrapCtx, lineFit: fit(0.8, 0) });
    const expected = numberedScan(contour, -50, 60, bodyIndentAt, 0.8, 0, 95);
    expect(expected).toBeLessThan(50);
    expect(Math.abs((lines[0]?.topY ?? Number.NaN) - expected)).toBeLessThanOrEqual(0.02);
  });

  it('matches a fine scan of numbered first lines on sloped contours (randomized)', () => {
    let seed = 90210;
    const random = (): number => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    for (let trial = 0; trial < 60; trial += 1) {
      const marker = 3 + Math.round(random() * 12);
      const hanging = -20 - Math.round(random() * 40);
      const tab = 24 + Math.round(random() * 24);
      const from = 220 + Math.round(random() * 100);
      const to = 40 + Math.round(random() * 120);
      const height = 40 + Math.round(random() * 60);
      const [fraction, share] = [[1, 0.1], [1, 0.2], [0.9, 0], [0.8, 0]][Math.floor(random() * 4)]!;
      const word = 5 * (10 + Math.floor(random() * 20));
      const bodyIndentAt = (edge: number): number => Math.ceil((edge + marker) / tab) * tab;
      const contour: FloatRect = { ...obstacle(hanging, 360), authoredWrap: 'through', imageY: 0, imageH: height,
        yTop: 0, yBottom: height, xLeft: hanging, imageX: hanging,
        wrapPolygon: [{ xPt: hanging, yPt: 0 }, { xPt: from, yPt: 0 }, { xPt: to, yPt: height }, { xPt: hanging, yPt: height }] };
      const wrapCtx: WrapLayoutCtx = { ...wrap([contour], 360 - hanging, false, hanging),
        numberingFirstLine: { authoredFirstIndentPx: hanging, bodyIndentAt } };
      const lines = run([token('W'.repeat(word / 5))], { wrapCtx, lineFit: fit(fraction, share) });
      const expected = numberedScan(contour, hanging, height, bodyIndentAt, fraction, share, word);
      const label = JSON.stringify({ trial, marker, hanging, tab, from, to, height, fraction, share, word, expected });
      const got = lines[0]?.topY ?? Number.NaN;
      if (expected >= height) continue; // fits only below the contour
      expect(Math.abs(got - expected), label).toBeLessThanOrEqual(0.02);
    }
  }, 60_000);
});
