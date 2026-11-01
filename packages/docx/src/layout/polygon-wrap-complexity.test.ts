import { describe, expect, it } from 'vitest';
import {
  prepareFloatWrap,
  computePreparedLineFloatWindowWithDiagnostics,
  LINE_SEARCH_Y_LIMIT_PT,
  type FloatRect,
} from './float-wrap.js';
import {
  __test_contourSpanIndexDiagnostics,
  compilePolygonWrap,
  polygonBandExactIntervalFunctions,
  projectPolygonIntervals,
  type CompiledPolygonWrap,
  type PolygonInterval,
} from './polygon-wrap.js';
import {
  compareExactRational,
  decodeBinary64,
  exactRationalFromNumber,
  exactRationalKey,
  exactRationalToNumber,
  midpointExactRational,
  normalizeExactRational,
  scaleExactRationalByPowerOfTwo,
  type ExactRational,
} from './exact-geometry.js';

function starPoints(vertexCount: number, step: number) {
  return Array.from({ length: vertexCount }, (_, index) => {
    const circleIndex = (index * step) % vertexCount;
    const angle = -Math.PI / 2 + (circleIndex * 2 * Math.PI) / vertexCount;
    return {
      xPt: 100 + Math.cos(angle) * 90,
      yPt: 100 + Math.sin(angle) * 90,
    };
  });
}

function denseStar(
  vertexCount: number,
  step: number,
  wrap: 'tight' | 'through' = 'tight',
): FloatRect {
  const points = starPoints(vertexCount, step);
  const xs = points.map(({ xPt }) => xPt);
  const ys = points.map(({ yPt }) => yPt);
  const xLeft = Math.min(...xs);
  const xRight = Math.max(...xs);
  const yTop = Math.min(...ys);
  const yBottom = Math.max(...ys);
  return {
    kind: 'shape',
    mode: 'square',
    authoredWrap: wrap,
    wrapPolygon: points,
    imageKey: 'dense-star',
    imageX: xLeft,
    imageY: yTop,
    imageW: xRight - xLeft,
    imageH: yBottom - yTop,
    xLeft,
    xRight,
    yTop,
    yBottom,
    side: 'bothSides',
    distLeft: 0,
    distRight: 0,
    distTop: 0,
    distBottom: 0,
    paraId: 1,
  };
}

/** Tight and through polygons step a line down by its height inside the
 * line-search domain (WORD_TIGHT/THROUGH_WRAP_LINE_STEP_ADVANCE); the exact
 * affine sweep and its threshold roots place lines beyond it. These sweep
 * tests move their geometry there: T is a power of two past the domain limit,
 * so every coordinate below 2T is a multiple of ULP_AT_T. */
const T = 2 * LINE_SEARCH_Y_LIMIT_PT;
const ULP_AT_T = T * Number.EPSILON;

function translatedY(rect: FloatRect, dy: number): FloatRect {
  return {
    ...rect,
    ...(rect.wrapPolygon ? { wrapPolygon: rect.wrapPolygon.map(({ xPt, yPt }) => ({ xPt, yPt: yPt + dy })) } : {}),
    imageY: rect.imageY + dy,
    yTop: rect.yTop + dy,
    yBottom: rect.yBottom + dy,
  };
}

function frozenDenseStar(vertexCount: number, step: number): Readonly<FloatRect> {
  const star = denseStar(vertexCount, step);
  return Object.freeze({
    ...star,
    wrapPolygon: Object.freeze(star.wrapPolygon!.map((point) => Object.freeze({ ...point }))),
  });
}

function edgeXAt(
  edge: CompiledPolygonWrap['edges'][number],
  yPt: number,
): number {
  const origin = edge.from.yPt <= edge.to.yPt ? edge.from : edge.to;
  return origin.xPt + edge.slopeXPerY * (yPt - origin.yPt);
}

interface OraclePoint {
  readonly x: bigint;
  readonly y: bigint;
}

interface OracleEdge {
  readonly index: number;
  readonly from: OraclePoint;
  readonly to: OraclePoint;
  readonly dx: bigint;
  readonly dy: bigint;
  readonly minY: bigint;
  readonly maxY: bigint;
}

interface ExactProjection {
  l: ExactRational;
  r: ExactRational;
}

interface ExactFullSortOracle {
  readonly scaleExponent: number;
  readonly edges: readonly OracleEdge[];
  readonly eventYs: readonly ExactRational[];
  edgeXAt(edge: OracleEdge, y: ExactRational): ExactRational;
  pairsAt(y: ExactRational): number[][];
  project(top: ExactRational, bottom: ExactRational): PolygonInterval[];
}

function exactInteger(value: bigint): ExactRational {
  return { numerator: value, denominator: 1n };
}

function buildExactFullSortOracle(
  points: readonly Readonly<{ xPt: number; yPt: number }>[],
): ExactFullSortOracle {
  const decoded = points.flatMap(({ xPt, yPt }) => [
    decodeBinary64(xPt),
    decodeBinary64(yPt),
  ]);
  const nonzeroExponents = decoded
    .filter(({ coefficient }) => coefficient !== 0n)
    .map(({ exponent }) => exponent);
  const scaleExponent = nonzeroExponents.length === 0
    ? 0
    : Math.min(...nonzeroExponents);
  let decodedIndex = 0;
  const exactPoints = points.map(() => {
    const x = decoded[decodedIndex++]!;
    const y = decoded[decodedIndex++]!;
    return {
      x: x.coefficient << BigInt(x.exponent - scaleExponent),
      y: y.coefficient << BigInt(y.exponent - scaleExponent),
    };
  });
  const edges = exactPoints.map((from, index): OracleEdge => {
    const to = exactPoints[(index + 1) % exactPoints.length]!;
    return {
      index,
      from,
      to,
      dx: to.x - from.x,
      dy: to.y - from.y,
      minY: from.y < to.y ? from.y : to.y,
      maxY: from.y > to.y ? from.y : to.y,
    };
  });
  const events = new Map<string, ExactRational>();
  for (const point of exactPoints) {
    const eventY = exactInteger(point.y);
    events.set(exactRationalKey(eventY), eventY);
  }
  for (let leftIndex = 0; leftIndex < edges.length; leftIndex += 1) {
    const left = edges[leftIndex]!;
    for (let rightIndex = leftIndex + 1; rightIndex < edges.length; rightIndex += 1) {
      const right = edges[rightIndex]!;
      const denominator = left.dx * right.dy - left.dy * right.dx;
      if (denominator === 0n) continue;
      const betweenX = right.from.x - left.from.x;
      const betweenY = right.from.y - left.from.y;
      const leftNumerator = betweenX * right.dy - betweenY * right.dx;
      const rightNumerator = betweenX * left.dy - betweenY * left.dx;
      const leftParameter = normalizeExactRational(leftNumerator, denominator);
      const rightParameter = normalizeExactRational(rightNumerator, denominator);
      if (compareExactRational(leftParameter, exactInteger(0n)) < 0
        || compareExactRational(leftParameter, exactInteger(1n)) > 0
        || compareExactRational(rightParameter, exactInteger(0n)) < 0
        || compareExactRational(rightParameter, exactInteger(1n)) > 0) continue;
      const y = normalizeExactRational(
        left.from.y * denominator + left.dy * leftNumerator,
        denominator,
      );
      events.set(exactRationalKey(y), y);
    }
  }
  const eventYs = [...events.values()].sort(compareExactRational);
  const edgeXAt = (edge: OracleEdge, y: ExactRational): ExactRational =>
    normalizeExactRational(
      edge.from.x * edge.dy * y.denominator
        + edge.dx * (y.numerator - edge.from.y * y.denominator),
      edge.dy * y.denominator,
    );
  const activeAt = (y: ExactRational): OracleEdge[] => edges
    .filter((edge) => compareExactRational(exactInteger(edge.minY), y) < 0
      && compareExactRational(y, exactInteger(edge.maxY)) < 0)
    .sort((left, right) =>
      compareExactRational(edgeXAt(left, y), edgeXAt(right, y))
        || left.index - right.index);
  const pairsAt = (y: ExactRational): number[][] => {
    const active = activeAt(y);
    expect(active.length % 2).toBe(0);
    const pairs: number[][] = [];
    for (let index = 0; index < active.length; index += 2) {
      pairs.push([active[index]!.index, active[index + 1]!.index]);
    }
    return pairs;
  };
  const coincident = (left: OracleEdge, right: OracleEdge): boolean =>
    left.dx * right.dy === right.dx * left.dy
      && (left.from.x * left.dy - left.dx * left.from.y) * right.dy
        === (right.from.x * right.dy - right.dx * right.from.y) * left.dy;
  const project = (top: ExactRational, bottom: ExactRational): PolygonInterval[] => {
    const slabEvents = [
      top,
      ...eventYs.filter((eventY) =>
        compareExactRational(top, eventY) < 0
          && compareExactRational(eventY, bottom) < 0),
      bottom,
    ];
    const projected: ExactProjection[] = [];
    for (let index = 0; index + 1 < slabEvents.length; index += 1) {
      const slabTop = slabEvents[index]!;
      const slabBottom = slabEvents[index + 1]!;
      const active = activeAt(midpointExactRational(slabTop, slabBottom));
      expect(active.length % 2).toBe(0);
      for (let activeIndex = 0; activeIndex < active.length; activeIndex += 2) {
        const left = active[activeIndex]!;
        const right = active[activeIndex + 1]!;
        if (coincident(left, right)) continue;
        const leftTop = edgeXAt(left, slabTop);
        const leftBottom = edgeXAt(left, slabBottom);
        const rightTop = edgeXAt(right, slabTop);
        const rightBottom = edgeXAt(right, slabBottom);
        projected.push({
          l: compareExactRational(leftTop, leftBottom) <= 0 ? leftTop : leftBottom,
          r: compareExactRational(rightTop, rightBottom) >= 0 ? rightTop : rightBottom,
        });
      }
    }
    projected.sort((left, right) =>
      compareExactRational(left.l, right.l) || compareExactRational(left.r, right.r));
    const merged: ExactProjection[] = [];
    for (const interval of projected) {
      if (compareExactRational(interval.l, interval.r) >= 0) continue;
      const previous = merged.at(-1);
      if (!previous || compareExactRational(previous.r, interval.l) < 0) {
        merged.push({ ...interval });
      } else if (compareExactRational(previous.r, interval.r) < 0) {
        previous.r = interval.r;
      }
    }
    return merged.map(({ l, r }) => ({
      l: exactRationalToNumber(scaleExactRationalByPowerOfTwo(l, scaleExponent)),
      r: exactRationalToNumber(scaleExactRationalByPowerOfTwo(r, scaleExponent)),
    }));
  };
  return { scaleExponent, edges, eventYs, edgeXAt, pairsAt, project };
}

function expectIntervalsClose(
  actual: PolygonInterval[],
  expected: PolygonInterval[],
  label = 'polygon interval projection',
): void {
  expect(actual, label).toHaveLength(expected.length);
  for (let index = 0; index < expected.length; index += 1) {
    expect(actual[index]!.l, `${label}, interval ${index} left`).toBeCloseTo(expected[index]!.l, 9);
    expect(actual[index]!.r, `${label}, interval ${index} right`).toBeCloseTo(expected[index]!.r, 9);
  }
}

describe('polygon line-window active sweep complexity', () => {
  it('retains all pairs when simultaneous starts shift active parity', () => {
    const points = [
      { xPt: 30, yPt: 0 },
      { xPt: 4, yPt: 6 },
      { xPt: 0, yPt: 1 },
      { xPt: 18, yPt: 1 },
      { xPt: 15, yPt: 0 },
      { xPt: 22, yPt: 2 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'simultaneous-start-parity-shift',
      points,
      xLeftPt: 0,
      xRightPt: 30,
      yTopPt: 0,
      yBottomPt: 6,
    });
    const midpoint = 0.5;
    const compiledPairs = polygon.contourSpans
      .filter((span) => span.yTopPt < midpoint && midpoint < span.yBottomPt)
      .sort((left, right) => edgeXAt(left.left, midpoint) - edgeXAt(right.left, midpoint))
      .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);

    expect(compiledPairs).toEqual([[3, 4], [0, 5]]);
    expectIntervalsClose(projectPolygonIntervals(polygon, 0, 1), [
      { l: 15, r: 18.5 },
      { l: 77 / 3, r: 30 },
    ]);
  });

  it('retains the full extent when simultaneous ends shift active parity', () => {
    const points = [
      { xPt: 40, yPt: 0 },
      { xPt: 50, yPt: 60 },
      { xPt: 20, yPt: 20 },
      { xPt: 0, yPt: 40 },
      { xPt: 20, yPt: 0 },
      { xPt: 80, yPt: 100 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'simultaneous-end-parity-shift',
      points,
      xLeftPt: 0,
      xRightPt: 80,
      yTopPt: 0,
      yBottomPt: 100,
    });

    expectIntervalsClose(projectPolygonIntervals(polygon, 20, 40), [
      { l: 0, r: 56 },
    ]);
  });

  it('keeps repeated adjacent-double spans logarithmically indexed', () => {
    const yTopPt = 1;
    const yBottomPt = 1 + Number.EPSILON;
    const spanCount = 1_023;
    expect(yTopPt + (yBottomPt - yTopPt) / 2).toBe(yTopPt);
    const spans = Array.from({ length: spanCount }, () => ({ yTopPt, yBottomPt }));
    const diagnostics = __test_contourSpanIndexDiagnostics(
      spans,
      yBottomPt,
      yBottomPt + Number.EPSILON,
    );
    const logarithmicBound = Math.ceil(Math.log2(spanCount + 1));

    expect(diagnostics.matchingSpanCount).toBe(0);
    expect(diagnostics.queryNodeVisitCount).toBeLessThanOrEqual(logarithmicBound);
    expect(diagnostics.height).toBeLessThanOrEqual(logarithmicBound);
    expect(__test_contourSpanIndexDiagnostics(
      spans,
      yTopPt - Number.EPSILON,
      yTopPt,
    ).matchingSpanCount).toBe(0);
    expect(__test_contourSpanIndexDiagnostics(
      spans,
      yTopPt,
      yBottomPt,
    ).matchingSpanCount).toBe(spanCount);
  });

  it('canonicalizes coordinate-equal endpoints before computing intersection roots', () => {
    const points = [
      { xPt: 1, yPt: 1 },
      { xPt: 12, yPt: 25 },
      { xPt: 1, yPt: 1 },
      { xPt: 14, yPt: 25 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'coordinate-equal-shared-endpoint',
      points,
      xLeftPt: 1,
      xRightPt: 14,
      yTopPt: 1,
      yBottomPt: 25,
    });

    expect(polygon.eventYPts).toEqual([1, 25]);
  });

  it('compiles the exact crossing event of an integer bow-tie', () => {
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'exact-integer-bow-tie',
      points: [
        { xPt: 0, yPt: 0 },
        { xPt: 10, yPt: 10 },
        { xPt: 0, yPt: 10 },
        { xPt: 10, yPt: 0 },
      ],
      xLeftPt: 0,
      xRightPt: 10,
      yTopPt: 0,
      yBottomPt: 10,
    });

    expect(polygon.eventYPts).toEqual([0, 5, 10]);
  });

  it('keeps horizontal contacts structural while diagonal edges reorder', () => {
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'horizontal-contact-with-diagonal-crossing',
      points: [
        { xPt: 0, yPt: 0 },
        { xPt: 10, yPt: 10 },
        { xPt: 0, yPt: 10 },
        { xPt: 10, yPt: 0 },
        { xPt: 10, yPt: 5 },
        { xPt: 0, yPt: 5 },
      ],
      xLeftPt: 0,
      xRightPt: 10,
      yTopPt: 0,
      yBottomPt: 10,
    });
    const midpoint = 7.5;
    const pairs = polygon.contourSpans
      .filter((span) => span.yTopPt < midpoint && midpoint < span.yBottomPt)
      .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);

    expect(polygon.eventYPts).toEqual([0, 5, 10]);
    expect(pairs).toEqual([[2, 0]]);
    expect(projectPolygonIntervals(polygon, 0, 5)).toEqual([{ l: 0, r: 10 }]);
    expect(projectPolygonIntervals(polygon, 5, 10)).toEqual([{ l: 0, r: 10 }]);
  });

  it('keeps shared-endpoint edge order stable above a nearby event', () => {
    const points = [
      { xPt: 60, yPt: 80 },
      { xPt: 80, yPt: 20 },
      { xPt: 50, yPt: 100 },
      { xPt: 10, yPt: 20 },
      { xPt: 40, yPt: 100 },
      { xPt: 100, yPt: 0 },
      { xPt: 20, yPt: 80 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'shared-endpoint-nearby-event-order',
      points,
      xLeftPt: 10,
      xRightPt: 100,
      yTopPt: 0,
      yBottomPt: 100,
    });
    const windowTopPt = 20.00000000000001;
    const midpoint = 35;
    const pairs = polygon.contourSpans
      .filter((span) => span.yTopPt < midpoint && midpoint < span.yBottomPt)
      .sort((left, right) => edgeXAt(left.left, midpoint) - edgeXAt(right.left, midpoint))
      .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);

    expect(pairs).toEqual([[3, 2], [5, 1], [0, 4]]);
    expect(projectPolygonIntervals(polygon, windowTopPt, 50)[0]!.r).toBe(25);
  });

  it('does not manufacture area from retraced coincident edge pairs', () => {
    const points = [
      { xPt: 19, yPt: 11 },
      { xPt: 40, yPt: 0 },
      { xPt: 19, yPt: 11 },
      { xPt: 9, yPt: 7 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'retraced-zero-area-contour',
      points,
      xLeftPt: 9,
      xRightPt: 40,
      yTopPt: 0,
      yBottomPt: 11,
    });

    expect(projectPolygonIntervals(polygon, 8, 10)).toEqual([]);
    expect(polygonBandExactIntervalFunctions(polygon, 2, 8, 10)).toEqual([]);
  });

  it('does not canonicalize a finite segment intersection outside its shared Y range', () => {
    const points = [
      { xPt: 9.968871586024673e56, yPt: 6.033146441914078e56 },
      { xPt: 9.968871586024827e56, yPt: 6.033146441914064e56 },
      { xPt: 9.968871586024692e56, yPt: 6.033146441914223e56 },
      { xPt: 9.968871586024673e56, yPt: 6.033146441914065e56 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'rounded-affine-intersection',
      points,
      xLeftPt: Math.min(...points.map(({ xPt }) => xPt)),
      xRightPt: Math.max(...points.map(({ xPt }) => xPt)),
      yTopPt: Math.min(...points.map(({ yPt }) => yPt)),
      yBottomPt: Math.max(...points.map(({ yPt }) => yPt)),
    });
    const left = polygon.edges[0]!;
    const right = polygon.edges[2]!;
    const affineRootY = (
      right.interceptX - left.interceptX
    ) / (left.slopeXPerY - right.slopeXPerY);
    const sharedMaximumY = Math.min(left.maxYPt, right.maxYPt);

    expect(affineRootY).toBeGreaterThan(sharedMaximumY);
    expect(polygon.eventYPts).not.toContain(affineRootY);
  });

  it('reorders at a valid determinant crossing when the affine root rounds out of range', () => {
    const points = [
      { xPt: 9.999999999996778e55, yPt: 1.0000000000003102e56 },
      { xPt: 1.000000000000298e56, yPt: 9.999999999995848e55 },
      { xPt: 9.99999999999894e55, yPt: 1.0000000000001175e56 },
      { xPt: 9.999999999997248e55, yPt: 1.000000000000255e56 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'determinant-only-ordering-crossing',
      points,
      xLeftPt: Math.min(...points.map(({ xPt }) => xPt)),
      xRightPt: Math.max(...points.map(({ xPt }) => xPt)),
      yTopPt: Math.min(...points.map(({ yPt }) => yPt)),
      yBottomPt: Math.max(...points.map(({ yPt }) => yPt)),
    });
    const determinantYPt = 1.0000000000002542e56;
    const sharedMaximumYPt = 1.000000000000255e56;
    const midpoint = determinantYPt + (sharedMaximumYPt - determinantYPt) / 2;
    const pairs = polygon.contourSpans
      .filter((span) => span.yTopPt < midpoint && midpoint < span.yBottomPt)
      .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);
    const projected = projectPolygonIntervals(polygon, determinantYPt, sharedMaximumYPt);

    expect(pairs).toEqual([[2, 0]]);
    expect(projected[0]!.l).toBe(9.999999999997248e55);
  });

  it('orders a one-probe dyadic slab from exact input geometry', () => {
    const points = [
      { xPt: 4.0740719526631545e90, yPt: 5.070602400921739e30 },
      { xPt: 4.07407195267702e90, yPt: 5.0706024009061487e30 },
      { xPt: 4.0740719526684204e90, yPt: 5.070602400924133e30 },
      { xPt: 4.0740719526751625e90, yPt: 5.070602400905802e30 },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'single-dyadic-probe-order',
      points,
      xLeftPt: Math.min(...points.map(({ xPt }) => xPt)),
      xRightPt: Math.max(...points.map(({ xPt }) => xPt)),
      yTopPt: Math.min(...points.map(({ yPt }) => yPt)),
      yBottomPt: Math.max(...points.map(({ yPt }) => yPt)),
    });
    const slabTopPt = 5.070602400909953e30;
    const slabBottomPt = 5.070602400909954e30;
    const probeYPt = 5.0706024009099537e30;
    const pairs = polygon.contourSpans
      .filter((span) => span.yTopPt < probeYPt && probeYPt < span.yBottomPt)
      .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);

    expect(pairs).toEqual([[3, 2], [0, 1]]);
    expect(projectPolygonIntervals(polygon, slabTopPt, slabBottomPt)).not.toEqual([]);
  });

  it('retains a finite crossing when binary64 determinant products overflow', () => {
    const base = 1e200;
    const delta = 1e190;
    const points = [
      { xPt: base, yPt: base },
      { xPt: base + delta, yPt: base + delta },
      { xPt: base, yPt: base + delta },
      { xPt: base + delta, yPt: base },
    ];
    const polygon = compilePolygonWrap({
      kind: 'through',
      imageKey: 'finite-overflowing-determinant',
      points,
      xLeftPt: base,
      xRightPt: base + delta,
      yTopPt: base,
      yBottomPt: base + delta,
    });

    expect(projectPolygonIntervals(polygon, base, base + delta / 4))
      .toEqual([{ l: base, r: base + delta }]);
  });

  it('finds the earliest line-window root for an overflowing bow-tie', () => {
    const base = 1e308;
    const delta = 1e300;
    const probeHeight = delta * 0.01;
    const requiredWidth = delta * 0.1;
    const prepared = prepareFloatWrap([{
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: base, yPt: base },
        { xPt: base + delta, yPt: base + delta },
        { xPt: base, yPt: base + delta },
        { xPt: base + delta, yPt: base },
      ],
      imageKey: 'overflowing-line-window-bow-tie',
      imageX: base,
      imageY: base,
      imageW: delta,
      imageH: delta,
      xLeft: base,
      xRight: base + delta,
      yTop: base,
      yBottom: base + delta,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }]);

    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      base,
      requiredWidth,
      probeHeight,
      base,
      (base + delta) - base,
      prepared,
      base,
      base + delta,
    );

    expect(window.topY).toBeCloseTo(base + requiredWidth, 12);
    expect(window.topY).toBeLessThan(base + delta * 0.5);
    expect(diagnostics.localRootCandidateCount).toBeGreaterThan(0);
  });

  it('resumes at the first binary64 value above an exact threshold root', () => {
    const prepared = prepareFloatWrap([translatedY({
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: 0, yPt: 0 },
        { xPt: 13, yPt: 5 },
        { xPt: 0, yPt: 5 },
        { xPt: 13, yPt: 0 },
      ],
      imageKey: 'directed-up-threshold-bow-tie',
      imageX: 0,
      imageY: 0,
      imageW: 13,
      imageH: 5,
      xLeft: 0,
      xRight: 13,
      yTop: 0,
      yBottom: 5,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }, T)]);

    const { window } = computePreparedLineFloatWindowWithDiagnostics(
      T, 6, 0.005, 0, 13, prepared, 0, 13,
    );

    // The width threshold root is T + 30/13; the result is the first binary64
    // value above it (offsets from T are exact multiples of ULP_AT_T).
    expect((window.topY - T) * 13).toBeGreaterThan(30);
    expect((window.topY - ULP_AT_T - T) * 13).toBeLessThanOrEqual(30);

    const secondPrepared = prepareFloatWrap([translatedY({
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: 0, yPt: 0 },
        { xPt: 11, yPt: 37 },
        { xPt: 0, yPt: 37 },
        { xPt: 11, yPt: 0 },
      ],
      imageKey: 'second-directed-up-threshold-bow-tie',
      imageX: 0,
      imageY: 0,
      imageW: 11,
      imageH: 37,
      xLeft: 0,
      xRight: 11,
      yTop: 0,
      yBottom: 37,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }, T)]);
    const { window: secondWindow } = computePreparedLineFloatWindowWithDiagnostics(
      T, 5, 0.005, 0, 11, secondPrepared, 0, 11,
    );

    // Root T + 185/11.
    expect((secondWindow.topY - T) * 11).toBeGreaterThan(185);
    expect((secondWindow.topY - ULP_AT_T - T) * 11).toBeLessThanOrEqual(185);
      });

  it('rejects a candidate whose exact final gap is still below the threshold', () => {
    const prepared = prepareFloatWrap([translatedY({
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: 0, yPt: 0 },
        { xPt: 3, yPt: 4 },
        { xPt: 0, yPt: 4 },
        { xPt: 3, yPt: 0 },
      ],
      imageKey: 'exact-final-candidate-bow-tie',
      imageX: 0,
      imageY: 0,
      imageW: 3,
      imageH: 4,
      xLeft: 0,
      xRight: 3,
      yTop: 0,
      yBottom: 4,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }, T)]);

    // The start is the binary64 value just below the root T + 4/3, where the
    // exact final gap is still below the threshold.
    const below = T + Math.floor((4 / 3) / ULP_AT_T) * ULP_AT_T;
    const { window } = computePreparedLineFloatWindowWithDiagnostics(
      below,
      1,
      0.0004,
      0,
      3,
      prepared,
      0,
      3,
    );

    expect(window.topY).toBe(below + ULP_AT_T);
  });

  it('retains a positive probe height that rounds away from the Number bottom', () => {
    const prepared = prepareFloatWrap([{
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'square',
      imageKey: 'rounded-away-probe-height',
      imageX: 0,
      imageY: 1,
      imageW: 3,
      imageH: 1,
      xLeft: 0,
      xRight: 3,
      yTop: 1,
      yBottom: 2,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }]);

    const { window } = computePreparedLineFloatWindowWithDiagnostics(
      1,
      1,
      2 ** -54,
      0,
      3,
      prepared,
      0,
      3,
    );

    expect(window.topY).toBe(2);
  });

  it('returns a representable window contained by exact polygon boundaries', () => {
    const rightGapPrepared = prepareFloatWrap([{
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: 0, yPt: 0 },
        { xPt: 1, yPt: 0 },
        { xPt: 0, yPt: 3 },
      ],
      imageKey: 'directed-window-start',
      imageX: 0,
      imageY: 0,
      imageW: 1,
      imageH: 3,
      xLeft: 0,
      xRight: 1,
      yTop: 0,
      yBottom: 3,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }]);
    const { window: rightGap } = computePreparedLineFloatWindowWithDiagnostics(
      2, 0.1, 0.1, 0, 3, rightGapPrepared, 0, 3,
    );

    expect(rightGap.xOffset).toBe(0.33333333333333337);
    expect(0 + rightGap.xOffset).toBeGreaterThan(1 / 3);

    const leftGapPrepared = prepareFloatWrap([{
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: [
        { xPt: 0, yPt: 0 },
        { xPt: 5, yPt: 0 },
        { xPt: 5, yPt: 25 },
      ],
      imageKey: 'directed-window-width',
      imageX: 0,
      imageY: 0,
      imageW: 5,
      imageH: 25,
      xLeft: 0,
      xRight: 5,
      yTop: 0,
      yBottom: 25,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    }]);
    const { window: leftGap } = computePreparedLineFloatWindowWithDiagnostics(
      11, 0.1, 0.1, 0, 5, leftGapPrepared, 0, 5,
    );

    expect(leftGap.xOffset).toBe(0);
    expect(leftGap.maxWidth).toBe(2.1999999999999997);
    expect(leftGap.xOffset + leftGap.maxWidth).toBeLessThan(2.2);
  });

  it('keeps exact eligibility separate from the safely contained Number width', () => {
    const polygonFloat = (
      imageKey: string,
      points: readonly { xPt: number; yPt: number }[],
      xRight: number,
    ): FloatRect => ({
      kind: 'shape',
      mode: 'square',
      authoredWrap: 'through',
      wrapPolygon: points,
      imageKey,
      imageX: 0,
      imageY: 0,
      imageW: xRight,
      imageH: 9,
      xLeft: 0,
      xRight,
      yTop: 0,
      yBottom: 9,
      side: 'bothSides',
      distLeft: 0,
      distRight: 0,
      distTop: 0,
      distBottom: 0,
      paraId: 1,
    });
    const prepared = prepareFloatWrap([
      translatedY(polygonFloat('left-root-boundary', [
        { xPt: 0, yPt: 0 },
        { xPt: 3, yPt: 9 },
        { xPt: 0, yPt: 9 },
      ], 3), T),
      translatedY(polygonFloat('right-root-boundary', [
        { xPt: 0, yPt: 0 },
        { xPt: 20, yPt: 0 },
        { xPt: 20, yPt: 9 },
        { xPt: 12, yPt: 9 },
      ], 20), T),
    ]);

    const { window } = computePreparedLineFloatWindowWithDiagnostics(
      T, 1, 3, 0, 20, prepared, 0, 20,
    );

    expect(window.topY).toBe(T + 2);
    expect(window.maxWidth).toBeLessThan(1);
  });

  it('compiles a dense self-intersecting star incrementally and queries adjacent boundaries', () => {
    const vertexCount = 31;
    // Inside the line-search domain tight and through lines advance in
    // line-height steps instead of the contour sweep, so the star lies past it.
    const prepared = prepareFloatWrap([translatedY(denseStar(vertexCount, 15, 'through'), T)]);

    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      T + 95, 195, 0.5, 0, 200, prepared,
    );

    expect(window.topY).toBeGreaterThan(T + 95);
    expect(window.maxWidth).toBeGreaterThanOrEqual(195);
    expect(diagnostics.compiledIntersectionCount)
      .toBeGreaterThanOrEqual(vertexCount * (vertexCount - 3) / 4);
    expect(diagnostics.compiledIntersectionCount)
      .toBeLessThanOrEqual(vertexCount * (vertexCount - 1) / 2);
    const binaryInsertionBound = Math.ceil(Math.log2(vertexCount)) + 1;
    expect(diagnostics.compileOrderComparisonCount)
      .toBeLessThanOrEqual(
        diagnostics.compiledIntersectionCount * binaryInsertionBound * 2
          + vertexCount * binaryInsertionBound * 2,
      );
    expect(diagnostics).toHaveProperty('compilePairMembershipVisitCount');
    const pairMembershipVisitCount = (
      diagnostics as typeof diagnostics & { compilePairMembershipVisitCount: number }
    ).compilePairMembershipVisitCount;
    expect(pairMembershipVisitCount).toBeGreaterThan(0);
    const completeVertexScanBound = vertexCount * vertexCount * 2;
    // Each intersection contributes two edges at up to two candidate events;
    // each edge seeds three pair slots before and after, reading two edges.
    const localIntersectionVisitBound = diagnostics.compiledIntersectionCount * 48;
    expect(pairMembershipVisitCount)
      .toBeLessThanOrEqual(completeVertexScanBound + localIntersectionVisitBound);
    expect(diagnostics.localRootCandidateCount).toBeGreaterThan(0);
    expect(diagnostics.localRootEventCount).toBeGreaterThan(0);
    expect(diagnostics.localRootCandidateCount)
      .toBeLessThanOrEqual(vertexCount * diagnostics.structuralEventCount * 8);
    expect(diagnostics.evaluatedYCount)
      .toBeLessThanOrEqual(diagnostics.structuralEventCount + diagnostics.localRootEventCount + 1);

    const polygon = prepared.floats[0]!.polygon!;
    const seen = new Set<object>();
    let maximumIndexDepth = 0;
    const visit = (
      node: NonNullable<typeof polygon.contourSpanIndex> | null,
      depth = 1,
    ): void => {
      if (!node) return;
      maximumIndexDepth = Math.max(maximumIndexDepth, depth);
      expect(node.crossingByBottom).toHaveLength(node.crossingByTop.length);
      expect(new Set(node.crossingByBottom)).toEqual(new Set(node.crossingByTop));
      for (const span of node.crossingByTop) {
        expect(seen.has(span)).toBe(false);
        seen.add(span);
      }
      visit(node.below, depth + 1);
      visit(node.above, depth + 1);
    };
    visit(polygon.contourSpanIndex);
    expect(seen.size).toBe(polygon.contourSpans.length);
    expect(maximumIndexDepth)
      .toBeLessThanOrEqual(Math.ceil(Math.log2(polygon.contourSpans.length + 1)) + 1);
  });

  it('compiles one immutable dense star once across a linear number of acquisitions', () => {
    const vertexCount = 31;
    const acquisitionCount = vertexCount;
    const star = frozenDenseStar(vertexCount, 15);
    const compiled: CompiledPolygonWrap[] = [];
    const totals = {
      polygonCompileCount: 0,
      polygonCacheHitCount: 0,
      polygonSnapshotPointCount: 0,
    };

    for (let index = 0; index < acquisitionCount; index += 1) {
      const diagnostics = {
        polygonCompileCount: 0,
        polygonCacheHitCount: 0,
        polygonSnapshotPointCount: 0,
      };
      compiled.push(prepareFloatWrap([star], diagnostics).floats[0]!.polygon!);
      totals.polygonCompileCount += diagnostics.polygonCompileCount;
      totals.polygonCacheHitCount += diagnostics.polygonCacheHitCount;
      totals.polygonSnapshotPointCount += diagnostics.polygonSnapshotPointCount;
    }

    expect(new Set(compiled).size).toBe(1);
    expect(totals).toEqual({
      polygonCompileCount: 1,
      polygonCacheHitCount: acquisitionCount - 1,
      polygonSnapshotPointCount: acquisitionCount * vertexCount,
    });
    expect(compiled[0]!.intersectionCount)
      .toBeGreaterThanOrEqual(vertexCount * (vertexCount - 3) / 4);
  });

  it('keys an immutable polygon compilation by exact kind and bounds', () => {
    const star = frozenDenseStar(15, 7);
    const first = prepareFloatWrap([star]).floats[0]!.polygon!;
    const changedKind = prepareFloatWrap([{
      ...star,
      authoredWrap: 'through',
    }]).floats[0]!.polygon!;
    expect(changedKind).not.toBe(first);
    expect(prepareFloatWrap([star]).floats[0]!.polygon).toBe(first);

    for (const field of ['xLeft', 'xRight', 'yTop', 'yBottom'] as const) {
      const changedBounds = prepareFloatWrap([{
        ...star,
        [field]: star[field] + 1,
      }]).floats[0]!.polygon!;
      expect(changedBounds).not.toBe(first);
    }
  });

  it('retains alternating immutable kind variants across repeated acquisitions', () => {
    const vertexCount = 31;
    const acquisitionCount = vertexCount;
    const tight = frozenDenseStar(vertexCount, 15);
    const through = Object.freeze({ ...tight, authoredWrap: 'through' as const });
    const diagnostics = {
      polygonCompileCount: 0,
      polygonCacheHitCount: 0,
      polygonSnapshotPointCount: 0,
    };
    const tightCompilations: CompiledPolygonWrap[] = [];
    const throughCompilations: CompiledPolygonWrap[] = [];

    for (let index = 0; index < acquisitionCount; index += 1) {
      const prepared = prepareFloatWrap([tight, through], diagnostics).floats;
      tightCompilations.push(prepared[0]!.polygon!);
      throughCompilations.push(prepared[1]!.polygon!);
    }

    expect(new Set(tightCompilations).size).toBe(1);
    expect(new Set(throughCompilations).size).toBe(1);
    expect(tightCompilations[0]).not.toBe(throughCompilations[0]);
    expect(diagnostics).toEqual({
      polygonCompileCount: 2,
      polygonCacheHitCount: 2 * (acquisitionCount - 1),
      polygonSnapshotPointCount: 2 * acquisitionCount * vertexCount,
    });
  });

  it('does not reuse translated, cloned, or mutable polygon geometry', () => {
    const star = frozenDenseStar(15, 7);
    const first = prepareFloatWrap([star]).floats[0]!.polygon!;
    const clonedPoints = Object.freeze(star.wrapPolygon!.map((point) =>
      Object.freeze({ ...point })));
    const cloned = prepareFloatWrap([{
      ...star,
      wrapPolygon: clonedPoints,
    }]).floats[0]!.polygon!;
    expect(cloned).not.toBe(first);
    expect(cloned.eventYPts).toEqual(first.eventYPts);

    const translatedPoints = Object.freeze(star.wrapPolygon!.map((point) => Object.freeze({
      xPt: point.xPt + 10,
      yPt: point.yPt + 20,
    })));
    const translated = prepareFloatWrap([{
      ...star,
      wrapPolygon: translatedPoints,
      xLeft: star.xLeft + 10,
      xRight: star.xRight + 10,
      yTop: star.yTop + 20,
      yBottom: star.yBottom + 20,
    }]).floats[0]!.polygon!;
    expect(translated).not.toBe(first);
    expect(translated.polygonLeftPt).toBeCloseTo(first.polygonLeftPt + 10, 12);
    expect(translated.polygonTopPt).toBeCloseTo(first.polygonTopPt + 20, 12);

    const scaledPoints = Object.freeze(star.wrapPolygon!.map((point) => Object.freeze({
      xPt: point.xPt * 2,
      yPt: point.yPt * 3,
    })));
    const scaled = prepareFloatWrap([{
      ...star,
      wrapPolygon: scaledPoints,
      xLeft: star.xLeft * 2,
      xRight: star.xRight * 2,
      yTop: star.yTop * 3,
      yBottom: star.yBottom * 3,
    }]).floats[0]!.polygon!;
    expect(scaled).not.toBe(first);
    expect(scaled.polygonRightPt).toBeCloseTo(first.polygonRightPt * 2, 12);
    expect(scaled.polygonBottomPt).toBeCloseTo(first.polygonBottomPt * 3, 12);

    const mutable = denseStar(15, 7);
    const mutablePoints = mutable.wrapPolygon!.map((point) => ({ ...point }));
    const mutableInput = { ...mutable, wrapPolygon: mutablePoints };
    const beforeMutation = prepareFloatWrap([mutableInput]).floats[0]!.polygon!;
    mutablePoints[0]!.xPt += 5;
    const afterMutation = prepareFloatWrap([mutableInput]).floats[0]!.polygon!;
    expect(afterMutation).not.toBe(beforeMutation);
    expect(afterMutation.edges[0]!.from.xPt).toBe(beforeMutation.edges[0]!.from.xPt + 5);
  });

  it('validates current bounds before consulting an immutable polygon memo', () => {
    const star = frozenDenseStar(15, 7);
    prepareFloatWrap([star]);

    expect(() => prepareFloatWrap([{
      ...star,
      imageKey: 'reversed-after-valid-cache-entry',
      xLeft: star.xRight + 1,
    }])).toThrow('Invalid finite wrap bounds for reversed-after-valid-cache-entry');
  });

  it('distinguishes signed-zero bounds in the immutable polygon memo', () => {
    const points = Object.freeze([
      Object.freeze({ xPt: 0, yPt: 0 }),
      Object.freeze({ xPt: 10, yPt: 0 }),
      Object.freeze({ xPt: 0, yPt: 10 }),
    ]);
    const base: FloatRect = {
      kind: 'shape', mode: 'square', authoredWrap: 'through', wrapPolygon: points,
      imageKey: 'signed-zero-bounds', imageX: 0, imageY: 0, imageW: 10, imageH: 10,
      xLeft: 0, xRight: 10, yTop: 0, yBottom: 10,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 1,
    };
    const positiveZero = prepareFloatWrap([base]).floats[0]!.polygon!;
    const negativeZero = prepareFloatWrap([{ ...base, xLeft: -0 }]).floats[0]!.polygon!;

    expect(negativeZero).not.toBe(positiveZero);
  });

  it('matches a full active sort across exact same-Y and endpoint events', () => {
    const fixtures = [
      {
        key: 'dense-odd-star',
        points: starPoints(15, 7),
      },
      {
        key: 'same-y-crossings',
        points: [
          { xPt: 0, yPt: 0 }, { xPt: 40, yPt: 100 },
          { xPt: 0, yPt: 100 }, { xPt: 40, yPt: 0 },
          { xPt: 60, yPt: 0 }, { xPt: 100, yPt: 100 },
          { xPt: 60, yPt: 100 }, { xPt: 100, yPt: 0 },
        ],
      },
      {
        key: 'same-y-multiway',
        points: [
          { xPt: 0, yPt: 0 }, { xPt: 100, yPt: 100 },
          { xPt: 0, yPt: 100 }, { xPt: 100, yPt: 0 },
          { xPt: 50, yPt: 0 }, { xPt: 50, yPt: 100 },
        ],
      },
      {
        key: 'edge-through-vertex',
        points: [
          { xPt: 0, yPt: 0 }, { xPt: 100, yPt: 100 },
          { xPt: 0, yPt: 100 }, { xPt: 50, yPt: 50 },
          { xPt: 100, yPt: 20 }, { xPt: 100, yPt: 0 },
        ],
      },
    ];
    for (const { key, points } of fixtures) {
      const xs = points.map(({ xPt }) => xPt);
      const ys = points.map(({ yPt }) => yPt);
      const polygon = compilePolygonWrap({
        kind: 'through',
        imageKey: key,
        points,
        xLeftPt: Math.min(...xs),
        xRightPt: Math.max(...xs),
        yTopPt: Math.min(...ys),
        yBottomPt: Math.max(...ys),
      });
      const oracle = buildExactFullSortOracle(points);

      for (let index = 0; index + 1 < oracle.eventYs.length; index += 1) {
        const exactTop = oracle.eventYs[index]!;
        const exactBottom = oracle.eventYs[index + 1]!;
        const exactMidpoint = midpointExactRational(exactTop, exactBottom);
        const top = exactRationalToNumber(
          scaleExactRationalByPowerOfTwo(exactTop, oracle.scaleExponent),
        );
        const bottom = exactRationalToNumber(
          scaleExactRationalByPowerOfTwo(exactBottom, oracle.scaleExponent),
        );
        const midpoint = exactRationalToNumber(
          scaleExactRationalByPowerOfTwo(exactMidpoint, oracle.scaleExponent),
        );
        if (midpoint <= top || midpoint >= bottom) continue;
        const expectedPairs = oracle.pairsAt(exactMidpoint);
        const actualPairs = polygon.contourSpans
          .filter((span) => span.yTopPt < midpoint && span.yBottomPt > midpoint)
          .sort((left, right) => {
            const leftIndex = polygon.edges.indexOf(left.left);
            const rightIndex = polygon.edges.indexOf(right.left);
            return compareExactRational(
              oracle.edgeXAt(oracle.edges[leftIndex]!, exactMidpoint),
              oracle.edgeXAt(oracle.edges[rightIndex]!, exactMidpoint),
            );
          })
          .map((span) => [polygon.edges.indexOf(span.left), polygon.edges.indexOf(span.right)]);
        expect(
          actualPairs,
          `${key} compiled pairs in event slab ${index} [${top}, ${bottom})`,
        ).toEqual(expectedPairs);
        const actualIntervals = projectPolygonIntervals(polygon, top, bottom);
        const expectedIntervals = oracle.project(
          scaleExactRationalByPowerOfTwo(
            exactRationalFromNumber(top),
            -oracle.scaleExponent,
          ),
          scaleExactRationalByPowerOfTwo(
            exactRationalFromNumber(bottom),
            -oracle.scaleExponent,
          ),
        );
        expectIntervalsClose(
          actualIntervals,
          expectedIntervals,
          `${key} event slab ${index} [${top}, ${bottom})`,
        );
      }
      expectIntervalsClose(
        projectPolygonIntervals(polygon, polygon.polygonTopPt, polygon.polygonBottomPt),
        oracle.project(oracle.eventYs[0]!, oracle.eventYs.at(-1)!),
        `${key} full-height projection`,
      );
    }
  });
});
