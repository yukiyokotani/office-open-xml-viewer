import { describe, it, expect } from 'vitest';
import {
  prepareFloatWrap,
  resolveLineFloatWindow,
  computePreparedLineFloatWindow,
  computePreparedLineFloatWindowWithDiagnostics,
  skipPastTopAndBottom,
  normalizeWrapSide,
  type FloatRect,
} from './float-layout.js';
import {
  layoutLines,
  type LayoutSeg,
  type LayoutTextSeg,
  type WrapLayoutCtx,
} from './line-layout.js';
import {
  lineGridAfter,
  lineSearchPitch,
  LINE_SEARCH_MAX_PITCH_PT,
  LINE_SEARCH_MIN_PITCH_PT,
  LINE_SEARCH_Y_LIMIT_PT,
} from './layout/float-wrap.js';

// Geometry and numeric-domain regression suite for the float event solver.
// #1670 Word controls supersede #676's inferred one-inch admission policy;
// production admission now measures the next atom and fills ordered gaps.

/** A LEFT-anchored square float band occupying [0, floatRightPx] horizontally on
 *  rows [0, floatBottomPx). paraX is 0, so with a column of `colWpx` the free
 *  RIGHT gap is `colWpx - floatRightPx`. */
function leftBand(floatRightPx: number, floatBottomPx: number): FloatRect {
  return {
    kind: 'shape', mode: 'square', imageKey: 'x',
    imageX: 0, imageY: 0, imageW: floatRightPx, imageH: floatBottomPx,
    xLeft: 0, xRight: floatRightPx, yTop: 0, yBottom: floatBottomPx,
    side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: 1,
  } as FloatRect;
}

function polygonFloat(
  authoredWrap: 'tight' | 'through',
  points: readonly Readonly<{ xPt: number; yPt: number }>[],
  overrides: Partial<Extract<FloatRect, { kind: 'shape' }>> = {},
): FloatRect {
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  const xLeft = Math.min(...xs);
  const xRight = Math.max(...xs);
  const yTop = Math.min(...ys);
  const yBottom = Math.max(...ys);
  return {
    ...leftBand(xRight, yBottom),
    imageKey: authoredWrap,
    authoredWrap,
    wrapPolygon: points,
    xLeft, xRight, yTop, yBottom,
    imageX: xLeft, imageY: yTop,
    imageW: xRight - xLeft, imageH: yBottom - yTop,
    ...overrides,
  };
}

const resolveWithReference = resolveLineFloatWindow as unknown as (
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
  columnXLeftPt?: number,
  columnXRightPt?: number,
  reference?: Readonly<{
    xLeftPt: number;
    xRightPt: number;
    readingDirection: 'ltr' | 'rtl';
  }>,
) => { topY: number; xOffset: number; maxWidth: number };

describe('resolveLineFloatWindow — float geometry and numeric domain', () => {
  const fullBand = (id: string, mode: FloatRect['mode'], yTop: number, yBottom: number): FloatRect => ({
    ...leftBand(100, yBottom),
    imageKey: id, mode, yTop, yBottom,
  });
  it('normalizes unknown legacy wrap sides to bothSides', () => {
    expect(normalizeWrapSide('left')).toBe('left');
    expect(normalizeWrapSide('legacy-unknown')).toBe('bothSides');
    expect(normalizeWrapSide(null)).toBe('bothSides');
  });

  it('projects a triangular tight polygon at the current line Y', () => {
    const triangle = polygonFloat('tight', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    expect(resolveLineFloatWindow(0, 1, 10, 0, 100, [triangle]))
      .toEqual({ topY: 0, xOffset: 0, maxWidth: 20 });
    expect(resolveLineFloatWindow(80, 1, 10, 0, 100, [triangle]))
      .toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
  });

  it('advances a tight line in whole line-height steps (word-tight-wrap-line-step-advance)', () => {
    const triangle = polygonFloat('tight', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    // The free side gap first reaches 40 pt at the contour root 200/3; the
    // line retries one line height (1 pt) lower each time, so it lands at 67.
    const result = resolveLineFloatWindow(0, 40, 1, 0, 100, [triangle]);

    expect(result.topY).toBe(67);
    expect(result.maxWidth).toBeGreaterThanOrEqual(40);
  });

  it('matches a per-line retry on sloped and self-intersecting tight polygons', () => {
    const star = Array.from({ length: 9 }, (_, index) => {
      const angle = (index * 4 * Math.PI) / 9;
      return { xPt: 50 + 45 * Math.cos(angle), yPt: 60 + 55 * Math.sin(angle) };
    });
    const shapes = [
      [{ xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 }],
      [{ xPt: 0, yPt: 10 }, { xPt: 100, yPt: 10 }, { xPt: 100, yPt: 30 },
        { xPt: 45, yPt: 60 }, { xPt: 100, yPt: 90 }, { xPt: 0, yPt: 90 }],
      star,
    ];
    for (const points of shapes) {
      const float = polygonFloat('tight', points);
      const prepared = prepareFloatWrap([float]);
      for (const [required, lineHeight] of [[40, 1], [30, 3.3], [55, 7], [10, 0.7]] as const) {
        const start = 0;
        const actual = computePreparedLineFloatWindow(start, required, lineHeight, 0, 100, prepared);
        // Reference: test every line step while the band still meets the polygon.
        let expected: number | null = null;
        for (let y = start; ; y += lineHeight) {
          const window = computePreparedLineFloatWindow(y, required, lineHeight, 0, 100, prepared);
          if (window.topY === y) { expected = y; break; }
          if (y > float.yBottom) break;
        }
        if (expected !== null) expect(actual.topY).toBeCloseTo(expected, 9);
      }
    }
  });

  /** First grid Y (start + k * lineHeight, k >= 0) at which the line fits,
   * testing every step while the band can still meet a polygon. */
  const perStepReference = (
    floats: FloatRect[],
    start: number,
    required: number,
    lineHeight: number,
    left: number,
    width: number,
  ) => {
    const prepared = prepareFloatWrap(floats);
    const bottom = Math.max(...floats.map((float) => float.yBottom));
    for (let k = 0; ; k += 1) {
      const y = start + k * lineHeight;
      const window = computePreparedLineFloatWindow(y, required, lineHeight, left, width, prepared);
      if (window.topY === y) return window;
      if (y > bottom) return null;
    }
  };

  it('keeps usable steps where a contour crosses the paragraph bounds', () => {
    const floats = [
      polygonFloat('tight', [{ xPt: 120, yPt: 30 }, { xPt: 70, yPt: 100 }, { xPt: 0, yPt: 60 }]),
      polygonFloat('tight', [{ xPt: 60, yPt: 0 }, { xPt: 0, yPt: 100 }, { xPt: 20, yPt: 70 }], { side: 'right' }),
    ];
    const actual = computePreparedLineFloatWindow(20, 58, 1, 0, 100, prepareFloatWrap(floats));
    expect(actual.topY).toBe(32);
    expect(actual.maxWidth).toBeCloseTo(58.2857, 3);
    expect(actual).toEqual(perStepReference(floats, 20, 58, 1, 0, 100));
  });

  /** Compact literal geometry for the representative search cases below. */
  type CaseFloat = Readonly<{
    side: string;
    points?: readonly (readonly [number, number])[];
    square?: readonly [number, number, number, number];
    topAndBottom?: true;
    exempt?: number;
  }>;
  const caseFloat = (spec: CaseFloat, offset = 0): FloatRect => {
    const extra = spec.exempt === undefined ? {} : { exemptLineTopPt: spec.exempt };
    if (spec.square) {
      const [x0, y0, x1, y1] = spec.square;
      const rect = { ...squareRect(x0, offset + y0, x1, offset + y1, spec.side), ...extra };
      return spec.topAndBottom ? { ...rect, mode: 'topAndBottom' } : rect;
    }
    return polygonFloat('tight', spec.points!.map(([xPt, yPt]) => ({ xPt, yPt: offset + yPt })), {
      side: spec.side, ...extra,
    } as Partial<Extract<FloatRect, { kind: 'shape' }>>);
  };

  // Representative cases distilled from a 400-trial seeded search (#1852).
  // Each was chosen because it fails when the walk drops the bottom-edge
  // line-step region, skips grid positions, or steps from the previous
  // candidate instead of the line's own grid; the remaining trials caught no
  // additional local mutant of the search. One mixes sloped polygons, the
  // other axis-aligned rectangles crossing the paragraph bounds.
  it.each([
    { specs: [
      { side: 'right', points: [[45, 104], [23, 65], [-20, 5], [-21, 7]] },
      { side: 'left', points: [[-8, 70], [29, 3], [88, 81]] },
      { side: 'left', points: [[56, 32], [133, 106], [93, 112], [129, 29]] },
    ], start: 22, required: 62, lineHeight: 0.7 },
    { specs: [
      { side: 'left', points: [[11, 15], [42, 15], [42, 44], [11, 44]] },
      { side: 'right', points: [[100, 44], [193, 44], [193, 74], [100, 74]] },
      { side: 'right', points: [[74, 4], [88, 4], [88, 55], [74, 55]] },
    ], start: 20, required: 37, lineHeight: 0.7 },
  ] as const)('matches per-step retries on tight polygons that cross the paragraph bounds (%#)', ({ specs, start, required, lineHeight }) => {
    const floats = specs.map((spec) => caseFloat(spec as CaseFloat));
    const expected = perStepReference(floats, start, required, lineHeight, 0, 100);
    expect(expected).not.toBeNull();
    expect(computePreparedLineFloatWindow(start, required, lineHeight, 0, 100, prepareFloatWrap(floats)))
      .toEqual(expected);
  });

  it('bounds tight line stepping by the polygons it has to pass', () => {
    for (const count of [10, 50, 75]) {
      const floats = Array.from({ length: count }, (_, index) => polygonFloat('tight', [
        { xPt: index - 10, yPt: index * 2 },
        { xPt: index + 50, yPt: index * 2 },
        { xPt: index + 50, yPt: index * 2 + 300 },
        { xPt: index - 10, yPt: index * 2 + 300 },
      ]));
      const prepared = prepareFloatWrap(floats);
      const started = performance.now();
      const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
        0, 60, 1, 0, 100, prepared,
      );
      const elapsedMs = performance.now() - started;
      if (count === 10) expect(window).toEqual(perStepReference(floats, 0, 60, 1, 0, 100));
      // One evaluation per line step until the band clears the lowest polygon
      // (bottom 2 * (count - 1) + 300), plus the structural events.
      const lowestBottom = 2 * (count - 1) + 300;
      expect(diagnostics.evaluatedYCount).toBeLessThanOrEqual(lowestBottom + 2 * count + 8);
      // Generous: the evaluation count is the real bound; this only catches a
      // pathological regression without flaking under load.
      expect(elapsedMs).toBeLessThan(5_000);
    }
  }, 60_000);

  /** Square exclusion band (a rectangle), full or partial column width. */
  const squareRect = (x0: number, y0: number, x1: number, y1: number, side = 'bothSides'): FloatRect => ({
    kind: 'shape', mode: 'square', imageKey: 'square',
    imageX: x0, imageY: y0, imageW: x1 - x0, imageH: y1 - y0,
    xLeft: x0, xRight: x1, yTop: y0, yBottom: y1,
    side, distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 1,
  } as FloatRect);

  /**
   * Independent oracle for a stack of square rectangles and tight polygons:
   * a line tests its own top; while a tight polygon meets its band it retries
   * on the grid start + k * lineHeight; otherwise it moves to the next square
   * bottom or exempt anchor line, or, when the next thing it reaches is a
   * tight polygon, onto the first grid step inside that polygon's band. An
   * exempt anchor line changes only its own window, never the tight region.
   */
  const mixedReference = (
    floats: FloatRect[],
    start: number,
    required: number,
    lineHeight: number,
  ) => {
    const prepared = prepareFloatWrap(floats);
    const fits = (y: number) => {
      const window = computePreparedLineFloatWindow(y, required, lineHeight, 0, 100, prepared);
      return window.topY === y ? window : null;
    };
    const inColumn = floats.filter((float) => float.xRight > 0.01 && float.xLeft < 99.99);
    const tight = inColumn.filter((float) => float.authoredWrap === 'tight');
    const squares = inColumn.filter((float) => float.authoredWrap !== 'tight');
    const meetsTight = (y: number) => tight.some((float) => y + lineHeight > float.yTop && y <= float.yBottom);
    const gridAfter = (y: number) => {
      let k = Math.floor((y - start) / lineHeight) + 1;
      while (start + (k - 1) * lineHeight > y) k -= 1;
      while (start + k * lineHeight <= y) k += 1;
      return start + k * lineHeight;
    };
    let y = start;
    for (let guard = 0; guard < 100_000; guard += 1) {
      const window = fits(y);
      if (window) return window;
      if (meetsTight(y)) {
        y = gridAfter(y);
        continue;
      }
      // Outside tight regions only square rectangles block, so the window can
      // improve only at a square bottom or an exempt anchor line.
      const candidates = [
        ...squares.map((float) => float.yBottom),
        ...inColumn.flatMap((float) => (float.exemptLineTopPt === undefined ? [] : [float.exemptLineTopPt])),
      ].filter((candidate) => candidate > y);
      const next = candidates.length > 0 ? Math.min(...candidates) : Number.POSITIVE_INFINITY;
      const entries = tight.map((float) => float.yTop - lineHeight).filter((entry) => entry >= y);
      const entry = entries.length > 0 ? Math.min(...entries) : Number.POSITIVE_INFINITY;
      if (entry < next) {
        y = gridAfter(entry);
      } else if (Number.isFinite(next)) {
        if (entry === next && !fits(next)) y = gridAfter(entry);
        else y = next;
      } else {
        throw new Error('oracle found no candidate');
      }
    }
    throw new Error('oracle did not terminate');
  };

  it('steps through a tight polygon reached after a square exclusion', () => {
    const floats = [
      squareRect(0, 0, 100, 20),
      polygonFloat('tight', [{ xPt: 0, yPt: 10 }, { xPt: 100, yPt: 10 }, { xPt: 100, yPt: 100 }, { xPt: 0, yPt: 100 }]),
    ];
    const prepared = prepareFloatWrap(floats);
    expect(computePreparedLineFloatWindow(0, 10, 5, 0, 100, prepared).topY).toBe(105);
    expect(computePreparedLineFloatWindow(20, 10, 5, 0, 100, prepared).topY).toBe(105);
    expect(computePreparedLineFloatWindow(0, 10, 5, 0, 100, prepared))
      .toEqual(mixedReference(floats, 0, 10, 5));
  });

  it('keeps a sloped tight polygon reached after a square exclusion on the line grid', () => {
    const floats = [
      squareRect(0, 0, 100, 20),
      polygonFloat('tight', [{ xPt: 0, yPt: 10 }, { xPt: 100, yPt: 10 }, { xPt: 50, yPt: 110 }]),
    ];
    const window = computePreparedLineFloatWindow(0, 40, 5, 0, 100, prepareFloatWrap(floats));
    expect(window.topY % 5).toBe(0);
    expect(window).toEqual(mixedReference(floats, 0, 40, 5));
  });

  it('keeps an exempt anchor line inside the tight line-step region', () => {
    const floats = [
      squareRect(0, 0, 100, 30),
      polygonFloat('tight', [{ xPt: 0, yPt: 30 }, { xPt: 60, yPt: 30 }, { xPt: 60, yPt: 100 }, { xPt: 0, yPt: 100 }],
        { exemptLineTopPt: 25 }),
    ];
    const window = computePreparedLineFloatWindow(3, 30, 10, 0, 100, prepareFloatWrap(floats));
    expect(window.topY).toBe(33);
    expect(window).toEqual(mixedReference(floats, 3, 30, 10));
  });

  // Representative cases distilled from a 400-trial seeded search (#1852):
  // exempt square and tight anchor lines, a sloped polygon entered from the
  // sweep, and `largest` squares above a tight triangle. Together they fail
  // when the sweep stops line-stepping into a tight slab, exempt lines stop
  // being events, the bottom edge leaves the line-step region, steps skip or
  // drift off the line grid, or the step-limit jump passes unmet polygons; the
  // remaining trials caught no additional local mutant of the search.
  it.each([
    { specs: [
      { square: [11, 79, 47, 96], side: 'bothSides', exempt: 12 },
      { square: [82, 118, 189, 124], side: 'left' },
      { side: 'bothSides', points: [[-19, 26], [23, 26], [23, 35], [-19, 35]] },
      { square: [28, 31, 127, 69], side: 'right', exempt: 51 },
    ], start: 28, required: 37, lineHeight: 5 },
    { specs: [
      { side: 'right', points: [[18, 92], [77, 92], [77, 101], [18, 101]] },
      { side: 'left', points: [[-27, 30], [52, 40], [13, 86]], exempt: 33 },
      { square: [74, 2, 151, 65], side: 'right', exempt: 64 },
    ], start: 26, required: 62, lineHeight: 1 },
    { specs: [
      { square: [89, 46, 186, 86], side: 'largest' },
      { square: [19, 56, 65, 62], side: 'largest' },
      { square: [11, 47, 99, 101], side: 'largest' },
      { side: 'right', points: [[16, 98], [40, 109], [28, 121]] },
    ], start: 45, required: 38, lineHeight: 5 },
  ] as const)('matches the oracle on mixed square and tight stacks from varied line tops (%#)', ({ specs, start, required, lineHeight }) => {
    const floats = specs.map((spec) => caseFloat(spec as CaseFloat));
    expect(computePreparedLineFloatWindow(start, required, lineHeight, 0, 100, prepareFloatWrap(floats)))
      .toEqual(mixedReference(floats, start, required, lineHeight));
  });

  it('caps tight line steps and falls back to the sweep aligned to the line grid', () => {
    const tall = polygonFloat('tight', [
      { xPt: 0, yPt: 0 }, { xPt: 100, yPt: 0 }, { xPt: 100, yPt: 10_000 }, { xPt: 0, yPt: 10_000 },
    ]);
    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      0, 10, 0.05, 0, 100, prepareFloatWrap([tall]),
    );
    // 200,000 steps would be needed; the 20,000-step resource limit hands the
    // rest to the sweep, whose answer lands on the next grid step.
    expect(window.topY).toBeGreaterThan(10_000);
    expect(window.topY).toBeLessThanOrEqual(10_000 + 0.05 + 1e-6);
    expect(diagnostics.evaluatedYCount).toBeLessThanOrEqual(20_000 + 16);
  }, 30_000);

  it('tests the first grid step below the tight polygons after the step limit', () => {
    const rectangle = (y0: number, y1: number) => polygonFloat('tight', [
      { xPt: 0, yPt: y0 }, { xPt: 100, yPt: y0 }, { xPt: 100, yPt: y1 }, { xPt: 0, yPt: y1 },
    ]);
    const unrelated = squareRect(200, 2000, 300, 2100);
    const alone = computePreparedLineFloatWindow(0, 10, 0.05, 0, 100, prepareFloatWrap([rectangle(0, 1000)]));
    expect(alone.topY).toBeGreaterThan(1000);
    expect(alone.topY).toBeLessThan(1000.05 + 1e-6);
    // An unrelated object below the region must not pull the line past the
    // first usable step once the step limit is reached.
    expect(computePreparedLineFloatWindow(
      0, 10, 0.05, 0, 100, prepareFloatWrap([rectangle(0, 1000), unrelated]),
    )).toEqual(alone);
    // Several overlapping tight polygons: the fallback passes all of them.
    expect(computePreparedLineFloatWindow(
      0, 10, 0.05, 0, 100,
      prepareFloatWrap([rectangle(0, 400), rectangle(300, 700), rectangle(650, 1000), unrelated]),
    )).toEqual(alone);
  }, 60_000);

  it('terminates after the step limit when a square bridges separated tight regions', () => {
    const rectangle = (y0: number, y1: number) => polygonFloat('tight', [
      { xPt: 0, yPt: y0 }, { xPt: 100, yPt: y0 }, { xPt: 100, yPt: y1 }, { xPt: 0, yPt: y1 },
    ]);
    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      0, 10, 0.5, 0, 100,
      prepareFloatWrap([rectangle(0, 10_000), squareRect(0, 9000, 100, 30_000), rectangle(20_000, 25_000)]),
    );
    expect(window).toEqual({ topY: 30_000, xOffset: 0, maxWidth: 100 });
    expect(diagnostics.evaluatedYCount).toBeLessThan(20_020);
  }, 30_000);

  it('retries a sub-twip line on the one-twip pitch and stays finite outside the grid domain', () => {
    const rectangle = (y0: number, y1: number, extra = {}) => polygonFloat('tight', [
      { xPt: 0, yPt: y0 }, { xPt: 100, yPt: y0 }, { xPt: 100, yPt: y1 }, { xPt: 0, yPt: y1 },
    ], extra);
    // 1e-14 pt: a 1e17-step grid collapses in binary64; the domain pitch is
    // one twip, so the first fit is the first twip step past the bottom edge.
    const tiny = computePreparedLineFloatWindowWithDiagnostics(
      0, 10, 1e-14, 0, 100, prepareFloatWrap([rectangle(0, 1000)]),
    );
    expect(tiny.window.topY).toBe(lineGridAfter(0, LINE_SEARCH_MIN_PITCH_PT, 1000));
    expect(tiny.diagnostics.evaluatedYCount).toBeLessThanOrEqual(20_000 + 16);
    // Number.MIN_VALUE height with an exempt line one ulp below the start:
    // that line top is not on the domain grid, so the earliest in-domain fit
    // is the first twip step past the bottom edge, found within the cap.
    const exempt = 100.00000000000001;
    const minimum = computePreparedLineFloatWindowWithDiagnostics(
      100, 10, Number.MIN_VALUE, 0, 100,
      prepareFloatWrap([rectangle(99, 101, { exemptLineTopPt: exempt })]),
    );
    const expected = lineGridAfter(100, LINE_SEARCH_MIN_PITCH_PT, 101);
    expect(minimum.window.topY).toBe(expected);
    expect(minimum.diagnostics.evaluatedYCount).toBeLessThanOrEqual(Math.ceil(1 / 0.05) + 8);
    // Outside the grid domain (off every page) the exact sweep alone places
    // the line, so coordinates near the binary64 limit stay finite.
    const huge = computePreparedLineFloatWindow(
      1e308, 10, 1e308, 0, 100, prepareFloatWrap([rectangle(0, 1e308)]),
    );
    expect(Number.isFinite(huge.topY)).toBe(true);
    expect(huge.topY).toBeGreaterThanOrEqual(1e308);
    const far = computePreparedLineFloatWindow(
      2 ** 56, 10, 0.05, 0, 100, prepareFloatWrap([rectangle(2 ** 56, 2 ** 56 + 2 ** 14)]),
    );
    expect(far.topY).toBeGreaterThanOrEqual(2 ** 56 + 2 ** 14);
    expect(far.topY).toBeLessThanOrEqual(2 ** 56 + 2 ** 14 + 16);
  }, 30_000);

  it('finds the first line-grid position strictly below any Y of the search domain', () => {
    let seed = 0x1ea5;
    const random = () => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed / 2 ** 32;
    };
    const span = (limit: number) => (random() * 2 - 1) * limit * random() ** 8;
    for (let trial = 0; trial < 5000; trial += 1) {
      const origin = span(LINE_SEARCH_Y_LIMIT_PT);
      const pitch = LINE_SEARCH_MIN_PITCH_PT
        + random() ** 4 * (LINE_SEARCH_MAX_PITCH_PT - LINE_SEARCH_MIN_PITCH_PT);
      const y = Math.min(LINE_SEARCH_Y_LIMIT_PT, Math.max(-LINE_SEARCH_Y_LIMIT_PT,
        origin + random() ** 6 * 3000 * pitch * (random() < 0.1 ? -1 : 1)));
      const next = lineGridAfter(origin, pitch, y);
      const label = `trial ${trial}`;
      expect(next, label).toBeGreaterThan(y);
      // Earliest: the previous index (if any) does not pass y.
      const k = Math.round((next - origin) / pitch);
      expect(origin + k * pitch, label).toBe(next);
      if (k > 1) expect(origin + (k - 1) * pitch, label).toBeLessThanOrEqual(y);
      expect(k, label).toBeGreaterThanOrEqual(1);
    }
    expect(lineSearchPitch(Number.MIN_VALUE)).toBe(LINE_SEARCH_MIN_PITCH_PT);
    expect(lineSearchPitch(1e308)).toBe(LINE_SEARCH_MAX_PITCH_PT);
    expect(lineSearchPitch(12)).toBe(12);
    expect(lineSearchPitch(0)).toBeNull();
    expect(() => lineGridAfter(1e308, 1e308, 1e308)).toThrow(RangeError);
    expect(() => lineGridAfter(100, Number.MIN_VALUE, 100)).toThrow(RangeError);
  });

  // Representative cases distilled from a 300-trial seeded search (#1852):
  // the first two fail when the bottom edge leaves the line-step region, steps
  // skip grid positions, or drift off the line grid; the remaining trials
  // caught no additional local mutant. The third keeps a large in-domain
  // coordinate (2^32 pt) with a one-twip line.
  it.each([
    { offset: 0, specs: [
      { side: 'bothSides', points: [[85, 13], [121, 13], [103, 25.5]] },
      { square: [25, 18, 92, 29.5], side: 'bothSides' },
      { square: [39, 5, 97, 11.5], side: 'right' },
    ], start: 2.75, required: 43, height: 2.5 },
    { offset: 0, specs: [
      { side: 'bothSides', points: [[37, 14.5], [133, 14.5], [133, 29.5], [37, 29.5]] },
      { side: 'right', points: [[76, 0.5], [104, 0.5], [104, 4], [76, 4]] },
      { side: 'largest', points: [[25, 19.5], [64, 19.5], [64, 28.5], [25, 28.5]] },
      { square: [55, 12.5, 89, 26], side: 'right' },
      { side: 'left', points: [[41, 7.5], [88, 16.5], [65, 14.5]] },
    ], start: 2.5, required: 35, height: 0.001 },
    { offset: 2 ** 32, specs: [
      { side: 'bothSides', points: [[6, 5.5], [22, 5.5], [22, 10.5], [6, 10.5]] },
      { square: [12, 17.5, 39, 32.5], side: 'right' },
      { square: [44, 14, 67, 23.5], side: 'right' },
      { side: 'largest', points: [[-19, 11.5], [22, 16.5], [2, 18]] },
      { side: 'right', points: [[77, 13], [101, 13], [101, 26], [77, 26]] },
    ], start: 2 ** 32 + 4.75, required: 22, height: 0.05 },
  ] as const)('keeps the search invariants against a brute-force grid oracle (%#)', ({ offset, specs, start, required, height }) => {
    const floats = specs.map((spec) => caseFloat(spec as CaseFloat, offset));
    const prepared = prepareFloatWrap(floats);
    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      start, required, height, 0, 100, prepared,
    );
    expect(diagnostics.evaluatedYCount).toBeLessThanOrEqual(20_000 + 64 * specs.length + 64);
    expect(Number.isFinite(window.topY)).toBe(true);
    expect(window.topY).toBeGreaterThanOrEqual(start);
    // The answer fits: re-querying at it returns it unchanged.
    expect(computePreparedLineFloatWindow(window.topY, required, height, 0, 100, prepared)).toEqual(window);
    // Nothing above the answer on the domain grid (start + k * pitch) fits.
    const pitch = lineSearchPitch(height)!;
    for (let k = 0; start + k * pitch < window.topY; k += 1) {
      const y = start + k * pitch;
      expect(computePreparedLineFloatWindow(y, required, height, 0, 100, prepared).topY, `k=${k}`).not.toBe(y);
    }
  });

  // Two representative stacks distilled from a 16-trial seeded search
  // (#1852): a tight region longer than the step limit followed by squares that
  // end inside, or cover, the next tight region, on the line grid and with
  // fractional tops between grid steps. They fail when the post-limit jump
  // lands a grid step late, never jumps, or ordinary steps skip positions.
  it.each([
    { start: 0, height: 0.25, regions: [
      { top: 0, bottom: 5108, tight: true },
      { top: 5107.75, bottom: 5114, tight: false },
      { top: 5110.75, bottom: 5115.25, tight: true },
      { top: 5115, bottom: 5124.25, tight: false },
      { top: 5118.75, bottom: 5125.5, tight: true },
    ] },
    { start: 0.125, height: 0.5, regions: [
      { top: 0.125, bottom: 10015.125, tight: true },
      { top: 10014.625, bottom: 10032.125, tight: false },
      { top: 10023.875, bottom: 10029.625, tight: true },
      { top: 10031.625, bottom: 10050.625, tight: false },
      { top: 10037.875, bottom: 10048.125, tight: true },
    ] },
  ] as const)('terminates and returns the first directly usable position after the limit (%#)', ({ start, height, regions }) => {
    const floats = regions.map(({ top, bottom: end, tight }) => tight
      ? polygonFloat('tight', [
        { xPt: 0, yPt: top }, { xPt: 100, yPt: top }, { xPt: 100, yPt: end }, { xPt: 0, yPt: end },
      ])
      : squareRect(0, top, 100, end));
    // Independent direct rectangle test: squares have an open bottom;
    // tight rectangles also block a line resting on their bottom edge.
    const usable = (y: number) => !regions.some((region) =>
      y + height > region.top && (region.tight ? y <= region.bottom : y < region.bottom));
    const bottom = Math.max(...regions.map((region) => region.bottom));
    let earliest = Number.POSITIVE_INFINITY;
    for (const region of regions) if (!region.tight && usable(region.bottom)) earliest = Math.min(earliest, region.bottom);
    for (let step = 0; step <= Math.ceil((bottom - start) / height) + 1; step += 1) {
      const y = start + step * height;
      if (usable(y)) {
        earliest = Math.min(earliest, y);
        break;
      }
    }
    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      start, 10, height, 0, 100, prepareFloatWrap(floats),
    );
    expect(usable(window.topY)).toBe(true);
    // These full-width rectangles cannot open an interior gap, so the
    // resource fallback must also retain the first directly usable position.
    expect(window.topY).toBe(earliest);
    expect(diagnostics.evaluatedYCount).toBeLessThan(20_000 + 12 * regions.length);
  }, 60_000);

  it('steps a through-wrap line down by its line height (word-through-wrap-line-step-advance)', () => {
    const triangle = polygonFloat('through', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    // The left gap [0, 20 + 0.3 y] (the edge read at the band top) first
    // holds 40 pt at y = 200/3; a 1 pt line takes the next line-grid top, 67.
    const result = resolveLineFloatWindow(0, 40, 1, 0, 100, [triangle]);

    expect(result.topY).toBe(67);
    expect(result.maxWidth).toBeCloseTo(40.1, 10);
  });

  it('admits a through opening that crosses the band edge by its clipped width (#2054)', () => {
    // Band [0, 300] in a 360 pt column; the notch open at the top runs from
    // x = 200 - y to x = 350, past the band edge. Word steps the line down by
    // its height and reads each window clipped to the band: a 110 pt unit
    // fits [200 - y, 300] from y = 10, so a 15 pt line takes y = 15.
    const notch = polygonFloat('through', [
      { xPt: 0, yPt: 0 }, { xPt: 200, yPt: 0 }, { xPt: 140, yPt: 60 }, { xPt: 350, yPt: 60 },
      { xPt: 350, yPt: 0 }, { xPt: 360, yPt: 0 }, { xPt: 360, yPt: 80 }, { xPt: 0, yPt: 80 },
    ]);

    expect(resolveLineFloatWindow(0, 110, 15, 0, 300, [notch], 0, 360))
      .toEqual({ topY: 15, xOffset: 185, maxWidth: 115 });
    // The tight copy does not enter the notch open at its top: below it.
    expect(resolveLineFloatWindow(0, 110, 15, 0, 300, [polygonFloat('tight', notch.wrapPolygon!)], 0, 360).topY)
      .toBe(90);
  });

  it('keeps a concave polygon interior gap for through but not tight', () => {
    const concave = [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 80 }, { xPt: 80, yPt: 80 },
      { xPt: 80, yPt: 20 }, { xPt: 20, yPt: 20 },
      { xPt: 20, yPt: 80 }, { xPt: 10, yPt: 80 },
    ];

    expect(resolveLineFloatWindow(30, 1, 10, 0, 100, [polygonFloat('tight', concave)]))
      .toEqual({ topY: 30, xOffset: 0, maxWidth: 10 });
    expect(resolveLineFloatWindow(30, 20, 10, 0, 100, [polygonFloat('through', concave)]))
      .toEqual({ topY: 30, xOffset: 20, maxWidth: 60 });
  });

  it('keeps a through-notch opening whose area begins at the line top', () => {
    const notch = [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ];

    expect(resolveLineFloatWindow(40, 20, 10, 0, 100, [polygonFloat('through', notch)]))
      .toEqual({ topY: 40, xOffset: 30, maxWidth: 40 });
  });

  it('does not apply square compatibility to a polygon gap the square does not constrain', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const containedSquare = {
      ...leftBand(20, 100),
      imageKey: 'contained-square',
      xLeft: 15,
      imageX: 15,
      imageW: 5,
    };

    expect(computePreparedLineFloatWindow(
      50, 20, 10, 0, 100,
      prepareFloatWrap([notch, containedSquare]),
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
      72,
    )).toEqual({ topY: 50, xOffset: 30, maxWidth: 40 });
  });

  it('applies square compatibility when the square constrains the selected polygon gap', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const gapBoundingSquare = {
      ...leftBand(35, 100),
      imageKey: 'gap-bounding-square',
      xLeft: 25,
      imageX: 25,
      imageW: 10,
    };

    expect(computePreparedLineFloatWindow(
      50, 20, 10, 0, 100,
      prepareFloatWrap([notch, gapBoundingSquare]),
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
      72,
    )).toEqual({ topY: 100, xOffset: 0, maxWidth: 100 });
  });

  it('uses an eligible through opening before it becomes the widest gap', () => {
    const opening = polygonFloat('through', [
      { xPt: 130, yPt: 0 }, { xPt: 200, yPt: 0 },
      { xPt: 200, yPt: 100 }, { xPt: 195, yPt: 100 },
      { xPt: 160, yPt: 20 }, { xPt: 140, yPt: 20 },
      { xPt: 140, yPt: 100 }, { xPt: 130, yPt: 100 },
    ]);
    const square = { ...leftBand(80, 100), imageKey: 'square-boundary' };

    const result = computePreparedLineFloatWindow(
      50, 1, 0.5, 0, 200,
      prepareFloatWrap([square, opening]),
      0, 200,
      { xLeftPt: 0, xRightPt: 200, readingDirection: 'ltr' },
      72,
    );

    expect(result.topY).toBe(50);
    expect(result).toMatchObject({ xOffset: 140, maxWidth: 33.125 });
  });

  it('supports an inferred-closure bow-tie whose signed shoelace area is zero', () => {
    const bowTie = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 100 },
      { xPt: 10, yPt: 100 }, { xPt: 90, yPt: 0 },
    ]);

    expect(resolveLineFloatWindow(20, 1, 10, 0, 100, [bowTie]))
      .toEqual({ topY: 20, xOffset: 0, maxWidth: 26 });
  });

  it('applies wrap distances to polygon line-band and horizontal projection', () => {
    const padded = polygonFloat('tight', [
      { xPt: 20, yPt: 20 }, { xPt: 40, yPt: 20 }, { xPt: 30, yPt: 40 },
    ], { xLeft: 15, xRight: 47, yTop: 10, yBottom: 46 });

    expect(resolveLineFloatWindow(45, 1, 1, 0, 100, [padded]))
      .toEqual({ topY: 45, xOffset: 0, maxWidth: 24.5 });
  });

  it('applies authored left, right, and largest sides after polygon projection', () => {
    const triangle = [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ];
    const resolve = (side: FloatRect['side']) => resolveLineFloatWindow(
      80, 1, 10, 0, 100, [polygonFloat('tight', triangle, { side })],
    );

    expect(resolve('left')).toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
    expect(resolve('right')).toEqual({ topY: 80, xOffset: 56, maxWidth: 44 });
    expect(resolve('largest')).toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
  });

  it('selects a polygon largest side from its maximum extent for every line band', () => {
    const asymmetric = prepareFloatWrap([polygonFloat('tight', [
      { xPt: 60, yPt: 0 }, { xPt: 70, yPt: 0 },
      { xPt: 20, yPt: 100 }, { xPt: 10, yPt: 100 },
    ], { side: 'largest' })]);
    const resolve = (topY: number) => computePreparedLineFloatWindow(
      topY, 1, 10, 0, 100, asymmetric,
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
    );

    // §20.4.3.7 selects the object's right side from its full [10, 70]
    // horizontal extent. Its top contour locally lies on the opposite half of
    // the page, but that line-band projection cannot reverse the selected side.
    expect(resolve(0)).toEqual({ topY: 0, xOffset: 70, maxWidth: 30 });
    expect(resolve(90)).toEqual({ topY: 90, xOffset: 25, maxWidth: 75 });
  });

  it('resolves a centered largest object by the first intersecting line direction', () => {
    const centered = { ...leftBand(60, 100), xLeft: 40, imageX: 40, imageW: 20, side: 'largest' };

    expect(resolveWithReference(0, 1, 10, 0, 100, [centered], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr',
    })).toEqual({ topY: 0, xOffset: 0, maxWidth: 40 });
    expect(resolveWithReference(0, 1, 10, 0, 100, [centered], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'rtl',
    })).toEqual({ topY: 0, xOffset: 60, maxWidth: 40 });
  });

  it('intersects the independently selected sides of multiple largest objects', () => {
    const left = { ...leftBand(40, 100), xLeft: 20, imageX: 20, imageW: 20, side: 'largest', imageKey: 'left' };
    const right = { ...leftBand(80, 100), xLeft: 60, imageX: 60, imageW: 20, side: 'largest', imageKey: 'right' };

    expect(resolveWithReference(0, 1, 10, 0, 100, [left, right], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr',
    })).toEqual({ topY: 0, xOffset: 40, maxWidth: 20 });
  });

  it('rejects tight and through floats without a finite nonzero polygon', () => {
    const missing = { ...leftBand(80, 20), authoredWrap: 'tight' as const };
    const nonfinite = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 }, { xPt: 50, yPt: 20 },
    ], { wrapPolygon: [{ xPt: Number.NaN, yPt: 0 }, { xPt: 90, yPt: 0 }, { xPt: 50, yPt: 20 }] });

    expect(() => resolveLineFloatWindow(0, 1, 10, 0, 100, [missing]))
      .toThrow(/invalid tight wrapPolygon/i);
    expect(() => resolveLineFloatWindow(0, 1, 10, 0, 100, [nonfinite]))
      .toThrow(/invalid through wrapPolygon/i);
  });

  it('crosses more than sixteen chained topAndBottom bottoms', () => {
    const floats = Array.from({ length: 20 }, (_, index) =>
      fullBand(`top-${index}`, 'topAndBottom', index, index + 1));

    expect(skipPastTopAndBottom(0, floats, 0, 100)).toBe(20);
    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(20);
  });

  it('crosses more than sixty-four chained square bottoms', () => {
    const floats = Array.from({ length: 70 }, (_, index) =>
      fullBand(`square-${index}`, 'square', index, index + 1));

    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(70);
  });

  it('rechecks topAndBottom after a square push', () => {
    const floats = [
      fullBand('square', 'square', 0, 10),
      fullBand('top', 'topAndBottom', 10, 20),
    ];

    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(20);
  });

  it('ignores square wrap rectangles wholly outside either side of the paragraph column', () => {
    const outsideRanges = [
      { xLeft: 20, xRight: 100 },
      { xLeft: 190, xRight: 270 },
    ];

    for (const side of ['bothSides', 'left', 'right', 'largest']) {
      for (const range of outsideRanges) {
        const outsideColumn = {
          ...leftBand(80, 120),
          ...range,
          imageX: range.xLeft,
          imageW: range.xRight - range.xLeft,
          side,
        };
        const win = resolveLineFloatWindow(
          20,
          10,
          10,
          110,
          70,
          [outsideColumn],
        );

        expect(win).toEqual({ topY: 20, xOffset: 0, maxWidth: 70 });
      }
    }
  });
});

// ── layoutLines integration ──────────────────────────────────────────────────
// Linear mock canvas: glyph advance = perPx · px · chars; ascent/descent 0.8/0.2
// em. Perfectly scale-linear so the wrap ALGORITHM is isolated from font hinting.
function makeLinearCtx(perPx = 0.5): CanvasRenderingContext2D {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      const per = p * perPx;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, fontSize = 10, extra: Partial<LayoutTextSeg> = {}): LayoutSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Times New Roman', vertAlign: null,
    measuredWidth: 0, ...extra,
  } as LayoutSeg;
}

function wrapCtx(floats: FloatRect[]): WrapLayoutCtx {
  return {
    startPageY: 0,
    paraX: 0,
    floats,
    lineBoxH: (asc: number, desc: number) => asc + desc,
    pageH: 100000,
  } as WrapLayoutCtx;
}

/** Was the FIRST line placed beside the band (topY 0, xOffset > 0) or flowed
 *  below it (topY past the band bottom)? */
function firstLinePlacement(lines: ReturnType<typeof layoutLines>): 'beside' | 'below' {
  const l = lines[0];
  return l.topY === 0 && l.xOffset > 0 ? 'beside' : 'below';
}

describe('layoutLines — float admission and probe convergence', () => {
  const scale = 1;
  const colW = 1000;
  const floatBottom = 50;

  // A gap just under 1 inch (70px) and just over (72px) at scale 1.
  const bandFor = (gapPx: number) => [leftBand(colW - gapPx, floatBottom)];


  it('keeps an anchor-host metric-only line on the paragraph-mark threshold', () => {
    const markWrap = {
      ...wrapCtx(bandFor(62)),
      paragraphMarkLineStartWidth: 10,
    };
    const lines = layoutLines(
      makeLinearCtx(),
      [textSeg('', 10, { metricOnly: true })],
      colW,
      0,
      scale,
      [],
      markWrap,
      {},
      0,
    );

    // A zero-advance anchor-character placeholder preserves the run's line
    // metrics, but it does not turn the pilcrow into inline content. The 62pt
    // side gap holds the 10pt mark even though it is below the 1-inch content
    // threshold, so the host line stays beside the float.
    expect(firstLinePlacement(lines)).toBe('beside');
    expect(lines[0].ascent + lines[0].descent).toBe(10);
  });





  it('keeps a fitting word beside the band without splitting', () => {
    // Gap = 200px. "AFTER" = 5 chars × 5px = 25px < 200px → sits beside, no split.
    const lines = layoutLines(makeLinearCtx(), [textSeg('AFTER', 10)], colW, 0, scale, [], wrapCtx(bandFor(200)), {}, 0);
    expect(firstLinePlacement(lines)).toBe('beside');
    expect((lines[0].segments[0] as LayoutTextSeg).text).toBe('AFTER');
  });

  it('remeasures a tall line against its actual polygon band', () => {
    const inverted = polygonFloat('tight', [
      { xPt: 50, yPt: 0 }, { xPt: 100, yPt: 100 }, { xPt: 0, yPt: 100 },
    ]);

    const lines = layoutLines(
      makeLinearCtx(), [textSeg('X', 30), { ...textSeg('', 60), metricOnly: true }], 100, 0, 1, [], wrapCtx([inverted]), {}, 0,
    );

    expect(lines[0].ascent + lines[0].descent).toBe(60);
    expect(lines[0].availWidth).toBe(20);
  });

  it('resolves a height-dependent float reflow monotonically', () => {
    // A 10pt line misses the exclusion at y=15 and admits the 20pt run;
    // that taller union intersects it and narrows to 50pt, excluding the run.
    // Without a monotone probe the short line reopens the band (a 2-cycle).
    // The line keeps its taller observed probe as a floor: the run stays
    // excluded and moves to the next physical line, which owns its band.
    const float: FloatRect = { ...leftBand(100, 100), xLeft: 50,
      imageX: 50, imageW: 50, yTop: 15, imageY: 15, imageH: 85 };
    const wrapping = { ...wrapCtx([float]), columnXPt: 0, columnWidthPt: 100 };
    const lines = layoutLines(makeLinearCtx(), [textSeg('AAAAA '), textSeg('BBB', 20)],
      100, 0, 1, [], wrapping);
    expect(lines.map((line) => [line.topY, line.xOffset, line.availWidth,
      line.segments.map((segment) => 'text' in segment ? segment.text : '')]))
      .toEqual([[0, 0, 50, ['AAAAA ']], [10, 0, 50, ['BBB']]]);
  });

  it('uses the first permitted through opening that admits the next atom', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const context = { ...wrapCtx([notch]), startPageY: 40 };

    const lines = layoutLines(
      makeLinearCtx(), [textSeg('word', 10)], 100, 0, 1, [], context, {}, 0,
    );

    expect(lines[0]).toMatchObject({ topY: 40, xOffset: 30, availWidth: 40 });
  });
});
