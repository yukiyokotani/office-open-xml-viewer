import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutLine, type LayoutSeg, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import { computePreparedLineFloatWindow, prepareFloatWrap } from '../layout/float-wrap.js';
import type { FloatRect } from '../layout/float-wrap.js';
import type { LineFitWidth } from './pass-driver.js';
import { KASHIDA_FLOAT_RETRY_STEP_LIMIT } from './pass-operations.js';

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
type NumberingFirstLine = NonNullable<WrapLayoutCtx['numberingFirstLine']>;
/** A numbered first line whose body starts at the next `tab` grid stop at or
 * after the marker end (edge + marker). */
function gridNumbering(authored: number, marker: number, tab: number): NumberingFirstLine {
  return { authoredFirstIndentPx: authored, bodyIndentAt: (edge) => Math.ceil((edge + marker) / tab) * tab };
}
/** A numbered first line whose body starts `offset` past the marker's
 * leading edge (a space or nothing suffix). */
function affineNumbering(authored: number, offset: number): NumberingFirstLine {
  return { authoredFirstIndentPx: authored, bodyIndentAt: (edge) => edge + offset };
}

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

  it('fits a line beside a float edge close to the band edge to that edge (kfit9 S rows)', () => {
    // Band [0, 300], text margin 360, mediumKashida, a float edge at 296
    // beside line 1: the reserve runs to the float edge, 0.9 x 296 = 266.4
    // holds the 265 pt prefix (an 8-character word and nine 4-character
    // words), where a reserve to the band edge less the trailing indent (264)
    // would not (blk-0898-kfit9 S001-S003: Word ends line 1 at t007).
    const words = [token('W'.repeat(8) + ' '), ...Array.from({ length: 12 }, () => token('aaaa '))];
    const lines = run(words, { maxWidth: 300, marginRightPx: 360, wrapCtx: wrap([band(296, 300, 0, 10)]), lineFit: fit(1, 0.1) });
    expect(lines[0]!.segments.filter((segment) => 'text' in segment).length).toBe(10);
  });

  it('takes a window ending at a float close to the band edge after a rejection (kfit9)', () => {
    // Band [0, 300], text margin 360, mediumKashida. Line 1 leaves [0, 100];
    // line 2 a float edge at 296, 4 pt inside the band edge, where the
    // reserve runs to the float edge: 0.9 x 296 = 266.4 holds the 265 pt word
    // (blk-0898-kfit9 S rows: Word ends the reserve at such an edge). The
    // clear band below fits only 300 - 0.1 x 360 = 264.
    const lines = run([token('W'.repeat(53))], {
      maxWidth: 300, marginRightPx: 360, wrapCtx: wrap([band(100, 300, 0, 10), band(296, 300, 10, 20)]), lineFit: fit(1, 0.1),
    });
    expect(lines.map((line) => [line.topY, line.xOffset, line.availWidth])).toEqual([[10, 0, 296]]);
  });

  it('keeps the tight-wrap retry steps of the real band (sol round 11)', () => {
    // mediumKashida, band [0, 300], trailing indent 60, a through polygon and
    // a tight 4 pt strip at the far edge: the 90 pt word lands where the
    // unchanged solver puts the interior requirement (100) on the real band.
    const through: FloatRect = { ...obstacle(0, 210), authoredWrap: 'through', imageY: 0, imageH: 60, yTop: 0, yBottom: 60,
      wrapPolygon: [{ xPt: 0, yPt: 0 }, { xPt: 210, yPt: 0 }, { xPt: 150, yPt: 60 }, { xPt: 0, yPt: 60 }] };
    const tight: FloatRect = { ...obstacle(296, 300), authoredWrap: 'tight', imageY: 0, imageH: 60, yTop: 0, yBottom: 60,
      wrapPolygon: [{ xPt: 296, yPt: 0 }, { xPt: 300, yPt: 0 }, { xPt: 300, yPt: 60 }, { xPt: 296, yPt: 60 }] };
    const lines = run([token('W'.repeat(18))], { maxWidth: 300, marginRightPx: 360, wrapCtx: wrap([through, tight]), lineFit: fit(1, 0.1) });
    const expected = computePreparedLineFloatWindow(0, 100, 10, 0, 300, prepareFloatWrap([through, tight]), 0, 360,
      { xLeftPt: 0, xRightPt: 300, readingDirection: 'ltr' }, 100);
    expect(expected.topY).toBe(20);
    expect(lines[0]!.topY).toBe(20);
  });

  it('admits a narrow far-edge window by its own requirement (sol round 11)', () => {
    // highKashida, band [0, 300], trailing indent 1, one 0.65 pt glyph: the
    // 1 pt window at y = 0 fits 0.6, the 1.1 pt window at y = 10 fits 0.68.
    const ctx = { measureText: (text: string) => ({ width: text.length * 0.65,
      fontBoundingBoxAscent: 0.8, fontBoundingBoxDescent: 0.2 }) } as unknown as CanvasRenderingContext2D;
    const lines = layoutLines(ctx, [token('W', { fontSize: 1 })], 300, 0, 1, [],
      wrap([band(0, 299, 0, 10), band(0, 298.9, 10, 20)]), {}, 0,
      undefined, undefined, undefined, 301, false, true, false, undefined, 'bounded', undefined, true, false, fit(1, 0.2));
    expect(lines.map((line) => [line.topY, line.xOffset])).toEqual([[10, 298.9]]);
  });

  it('bounds the far-edge search by line-grid steps up to the wider requirement', () => {
    // highKashida, band [0, 300], trailing 60, a 210 pt word: the first line
    // [0, 250] fits 200; the lines beside a left float leave [30, 300] (far
    // edge: fits 204 < 210, though 270 holds the interior 262.5) for 19
    // lines; the clear band at y = 200 fits 228.
    let calls = 0;
    const floats = [band(250, 300, 0, 10), band(0, 30, 10, 200)];
    const prepared = prepareFloatWrap(floats);
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      // The retry's queries (the interior 262.5 and far-edge 277.5 widths).
      if (input.minimumStartWidthPt > 260) calls += 1;
      const window = computePreparedLineFloatWindow(input.topYPt, input.minimumStartWidthPt, input.probeHeightPt,
        input.paragraphXPt, input.maximumWidthPt, prepared, input.columnXPt, input.columnXPt + input.columnWidthPt,
        { xLeftPt: input.paragraphXPt, xRightPt: input.paragraphXPt + input.maximumWidthPt, readingDirection: 'ltr' },
        input.squareMinimumStartWidthPt ?? input.minimumStartWidthPt);
      return { topYPt: window.topY, xOffsetPt: window.xOffset, maximumWidthPt: window.maxWidth };
    };
    const lines = run([token('W'.repeat(42))], { maxWidth: 300, marginRightPx: 360,
      wrapCtx: { ...wrap(floats), lineWindow }, lineFit: fit(1, 0.2) });
    expect(lines[0]!.topY).toBe(200);
    // Per layout pass (the paragraph is laid out twice): the smaller and the
    // larger width once, then one query per line-grid top up to y = 200.
    expect(calls).toBeLessThanOrEqual(2 * (2 + 200 / 10));
  });

  it('falls back to the wider requirement after the retry step limit (sol round 12)', () => {
    // As the bounded search above, with the left float running to 1e6: every
    // [30, 300] window holds the interior 262.5 and fails the far-edge 277.5.
    let calls = 0;
    const floats = [band(250, 300, 0, 10), band(0, 30, 10, 1e6)];
    const prepared = prepareFloatWrap(floats);
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      calls += 1;
      const window = computePreparedLineFloatWindow(input.topYPt, input.minimumStartWidthPt, input.probeHeightPt,
        input.paragraphXPt, input.maximumWidthPt, prepared, input.columnXPt, input.columnXPt + input.columnWidthPt,
        { xLeftPt: input.paragraphXPt, xRightPt: input.paragraphXPt + input.maximumWidthPt, readingDirection: 'ltr' },
        input.squareMinimumStartWidthPt ?? input.minimumStartWidthPt);
      return { topYPt: window.topY, xOffsetPt: window.xOffset, maximumWidthPt: window.maxWidth };
    };
    const wrapCtx: WrapLayoutCtx = { ...wrap(floats), lineWindow, pageH: 2e6 };
    const lines = run([token('W'.repeat(42))], { maxWidth: 300, marginRightPx: 360, wrapCtx, lineFit: fit(1, 0.2) });
    expect(lines[0]!.topY).toBe(1e6);
    // Per layout pass: the smaller and the larger width, the step limit, and
    // the line's other queries.
    expect(calls).toBeLessThanOrEqual(2 * (2 + KASHIDA_FLOAT_RETRY_STEP_LIMIT + 4));
  });

  it('asks the line-window oracle once per rejection with each kind\'s requirement', () => {
    // A numbered first line with a -50 pt authored hanging indent: every top
    // offers a narrowed 390 pt window, too small for the 320 pt word at 0.8
    // (it needs 400 pt). The retry leaves them in one query.
    let calls = 0;
    const lineWindow: NonNullable<WrapLayoutCtx['lineWindow']> = (input) => {
      calls += 1;
      if (calls > 50) throw new Error('search did not terminate');
      for (let top = input.topYPt; top < 1000; top += 10) {
        if (input.minimumStartWidthPt <= 390) {
          return { topYPt: top, xOffsetPt: 20, maximumWidthPt: 390 };
        }
      }
      return { topYPt: 100000, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
    };
    const wrapCtx: WrapLayoutCtx = { ...wrap([]), floats: [], lineWindow,
      numberingFirstLine: affineNumbering(-50, 0) };
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
  it('places a mode-15 line as the unreserved breaker places its scaled advance (randomized)', () => {
    // WORD_KASHIDA_LINE_FIT from mode 15 is a pure scale: fit = f x width
    // after the first-line offset, so a word of advance a is placed where the
    // unreserved breaker places a word of advance a / f, first-line, hanging
    // and numbered first lines included.
    let seed = 1515;
    const random = (): number => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    for (let trial = 0; trial < 300; trial += 1) {
      const fraction = trial % 2 === 0 ? 0.8 : 0.5;
      const word = fraction === 0.8 ? 5 * 4 * (4 + Math.floor(random() * 10)) : 5 * (8 + Math.floor(random() * 20));
      const scaled = word / fraction;
      const kind = trial % 4;
      const indent = kind === 1 ? 10 + Math.round(random() * 30) : kind === 2 ? -10 - Math.round(random() * 40) : 0;
      const floats = Array.from({ length: 3 }, (_, line) => {
        const left = Math.round(random() * 200);
        return band(left, left + 20 + Math.round(random() * 120), line * 10, line * 10 + 10);
      });
      const numberingFirstLine = kind === 3 ? gridNumbering(-30, 5 + Math.round(random() * 10), 24 + Math.round(random() * 24)) : undefined;
      const wrapCtx: WrapLayoutCtx = { ...wrap(floats, 300), ...(numberingFirstLine ? { numberingFirstLine } : {}) };
      const options = { maxWidth: 300, marginRightPx: 300, firstIndent: kind === 3 ? -30 : indent, wrapCtx };
      const reserved = run([token('W'.repeat(word / 5))], { ...options, lineFit: fit(fraction, 0) });
      const unreserved = run([token('W'.repeat(scaled / 5))], options);
      const label = JSON.stringify({ trial, fraction, word, indent, floats: floats.map((rect) => [rect.xLeft, rect.xRight]) });
      expect(reserved.map((line) => [line.topY, line.xOffset, line.availWidth]), label)
        .toEqual(unreserved.map((line) => [line.topY, line.xOffset, line.availWidth]));
    }
  });

});
