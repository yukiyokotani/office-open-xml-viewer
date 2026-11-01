// Float-wrap geometry for DOCX anchors (ECMA-376 §20.4.2.x).
//
// Pure layout math: given the floats active on a page (as FloatRect exclusion
// records) it answers "where may this line sit?". Object-to-object displacement
// policy lives in floats.ts. No canvas/drawing or document-model deps, so the
// geometry can be unit-reasoned and shared by the renderer and the paginator.
//
// Word gap order is measured by the issue #1670 controls; callers supply
// the next atom width, while this module owns geometry and numeric safety.

import {
  assertValidPolygonCompileInput,
  compilePolygonWrap,
  polygonBandExactIntervalFunctions,
  polygonLineTopEventYPts,
  projectPolygonExactLineIntervals,
  type CompiledPolygonWrap,
} from './polygon-wrap.js';
import {
  compareExactRational,
  decodeBinary64,
  exactRationalToNumberDown,
  exactRationalToNumber,
  exactRationalToNumberUp,
  type ExactRational,
} from './exact-geometry.js';
import {
  axisAlignedRectsOverlap,
} from './axis-aligned-overlap.js';
import { FLOAT_OVERLAP_EPS } from './floats.js';

function unreducedExactFromNumber(value: number): ExactRational {
  const decoded = decodeBinary64(value);
  return decoded.exponent >= 0
    ? {
        numerator: decoded.coefficient << BigInt(decoded.exponent),
        denominator: 1n,
      }
    : {
        numerator: decoded.coefficient,
        denominator: 1n << BigInt(-decoded.exponent),
      };
}

function addUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      + right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function subtractUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      - right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function multiplyUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  return {
    numerator: left.numerator * right.numerator,
    denominator: left.denominator * right.denominator,
  };
}

function divideUnreducedExact(left: ExactRational, right: ExactRational): ExactRational {
  const negative = right.numerator < 0n;
  return {
    numerator: (negative ? -left.numerator : left.numerator) * right.denominator,
    denominator: left.denominator * (negative ? -right.numerator : right.numerator),
  };
}

function exactBinary64Midpoint(left: number, right: number): number {
  const exactLeft = unreducedExactFromNumber(left);
  const exactRight = unreducedExactFromNumber(right);
  return exactRationalToNumber({
    numerator: exactLeft.numerator * exactRight.denominator
      + exactRight.numerator * exactLeft.denominator,
    denominator: 2n * exactLeft.denominator * exactRight.denominator,
  });
}

/**
 * Floating object that affects text wrap on the current page. `FloatRect` is a
 * historical name: tight/through records carry their authoritative polygon;
 * xLeft/xRight/yTop/yBottom are its padded acquisition bounds, not a replacement
 * rectangle for polygon line queries.
 */
interface FloatRectCore {
  /** Transitional registry kind projected into the typed displacement policy
   * in floats.ts. `shape` = DrawingML wp:anchor; `table` = floating table;
   * `frame` = <w:framePr> text frame. */
  mode: 'square' | 'topAndBottom';
  /** Exact retained wrap semantics; `mode` remains the coarse legacy routing key. */
  authoredWrap?: 'square' | 'tight' | 'through' | 'topAndBottom';
  anchorOccurrenceId?: string;
  acquisitionOccurrenceId?: string;
  wrapPolygon?: readonly Readonly<{ xPt: number; yPt: number }>[];
  /** WORD_MODE14_TIGHT_ANCHOR_TOP_TOUCH: a line starting at or below this Y
   * whose bottom lies exactly on this tight polygon's top also wraps around
   * the top edge. */
  topEdgeInclusiveFromYPt?: number;
  /** WORD_MODE14_TIGHT_ANCHOR_LINE_REWRAP: the line starting exactly at this
   * top ignores this float. */
  exemptLineTopPt?: number;
  /** Hex key of the image bitmap (used to defer drawing until final Y is known). */
  imageKey: string;
  /** Absolute X of the object box in the consuming wrap context's linear unit
   * (canonical acquisition uses points; the legacy line-layout bridge may use
   * scaled units). */
  imageX: number;
  imageY: number;
  imageW: number;
  imageH: number;
  /** Padded exclusion rectangle for text wrap. */
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  /** ST_WrapText: "bothSides" | "left" | "right" | "largest". */
  side: string;
  /** dist* padding in the same unit as the box — needed when displacing a float
   * to keep its exclusion padding when re-seating next to a blocking float
   * (ECMA-376 §20.4.2.x). */
  distLeft: number;
  distRight: number;
  distTop: number;
  distBottom: number;
  /** Identifier of the anchoring paragraph. Used only by the named
   *  compatibility rule under allowOverlap=true: floats with the SAME paraId
   *  never displace each other, while different-paragraph floats do. ECMA-376
   *  does not define this scoping; see compatibility.ts and floats.ts. */
  paraId: number;
}

export type FloatRect =
  | (FloatRectCore & {
      kind: 'table';
      /** Required §17.4.56 fact retained for tables placed later. */
      tableOverlap: 'never' | 'overlap';
    })
  | (FloatRectCore & {
      kind: 'shape';
    })
  | (FloatRectCore & {
      kind: 'frame';
    });

export type WrapSide = 'bothSides' | 'left' | 'right' | 'largest';

export function normalizeWrapSide(side: string | null | undefined): WrapSide {
  switch (side) {
    case 'left':
    case 'right':
    case 'largest':
    case 'bothSides': return side;
    default: return 'bothSides';
  }
}

/** A horizontal interval [l, r] in the consuming wrap context's linear unit. */
export interface Gap {
  l: number;
  r: number;
}

/** Minimum width, in the active wrap context's unit, that a polygon free gap
 * must have to hold a line start. It also floors a zero-width direct probe so
 *  a `requiredWidth === 0` call still rejects sub-pixel slivers between
 *  full-width floats. Square queries additionally apply the independently
 *  supplied compatibility width; tight/through deliberately do not. */
export const MIN_LINE_GAP = 1;

export function isWrapFloat(mode?: string | null): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
}

/**
 * Does float `f`'s horizontal extent overlap the paragraph/column text band
 * [paraXLeft, paraXRight]? Touching edges (within FLOAT_OVERLAP_EPS) do not count.
 *
 * ECMA-376 §20.4.2.17 (wrapSquare) and §20.4.2.20 (wrapTopAndBottom) both exclude
 * text only where the object is horizontally placed ("text shall wrap around …
 * THIS OBJECT"). Floats are registered in ABSOLUTE page coordinates and the page
 * float set is shared across a section's newspaper columns (§17.6.4), so a float
 * anchored in one column must be filtered out for a line laid out in another
 * column that it does not horizontally overlap. Both wrap modes route through
 * this one predicate so they share identical column-scoping semantics.
 */
export function floatOverlapsColumnX(
  f: FloatRect,
  paraXLeft: number,
  paraXRight: number,
): boolean {
  return f.xRight > paraXLeft + FLOAT_OVERLAP_EPS && f.xLeft < paraXRight - FLOAT_OVERLAP_EPS;
}

/** Two exclusion rects intersect (strict overlap, touching edges allowed). */
export function rectsOverlap(
  aL: number, aR: number, aT: number, aB: number,
  bL: number, bR: number, bT: number, bB: number,
): boolean {
  return axisAlignedRectsOverlap(
    { left: aL, right: aR, top: aT, bottom: aB },
    { left: bL, right: bR, top: bT, bottom: bB },
    FLOAT_OVERLAP_EPS,
  );
}

interface PreparedFloatRect {
  readonly rect: Readonly<FloatRect>;
  readonly polygon: CompiledPolygonWrap | null;
  readonly wrapMaximumLeftPt: number;
  readonly wrapMaximumRightPt: number;
}

export interface PreparedFloatWrap {
  readonly floats: readonly PreparedFloatRect[];
}

export interface FloatWrapPreparationDiagnostics {
  polygonCompileCount: number;
  polygonCacheHitCount: number;
  polygonSnapshotPointCount: number;
}

const sweepEventCache = new WeakMap<PreparedFloatWrap, Map<number, readonly number[]>>();

interface CompiledPolygonMemo {
  readonly kind: 'tight' | 'through';
  readonly xLeftPt: number;
  readonly xRightPt: number;
  readonly yTopPt: number;
  readonly yBottomPt: number;
  readonly compiled: CompiledPolygonWrap;
}

/**
 * Realm-local derived geometry keyed by the authoritative retained point-array
 * identity. A small variant bound avoids alternating kind/bounds thrash while
 * keeping memory bounded by live immutable polygons; translated/scaled clones
 * intentionally compile under their own identities.
 */
const MAX_COMPILED_POLYGON_VARIANTS = 4;
const compiledPolygonMemo = new WeakMap<
  readonly Readonly<{ xPt: number; yPt: number }>[],
  readonly CompiledPolygonMemo[]
>();

function polygonIsDeepFrozen(
  points: readonly Readonly<{ xPt: number; yPt: number }>[],
): boolean {
  return Object.isFrozen(points) && points.every((point) => Object.isFrozen(point));
}

function memoMatches(
  memo: CompiledPolygonMemo,
  kind: 'tight' | 'through',
  xLeftPt: number,
  xRightPt: number,
  yTopPt: number,
  yBottomPt: number,
): boolean {
  return memo.kind === kind
    && Object.is(memo.xLeftPt, xLeftPt)
    && Object.is(memo.xRightPt, xRightPt)
    && Object.is(memo.yTopPt, yTopPt)
    && Object.is(memo.yBottomPt, yBottomPt);
}

function compiledPolygonFor(
  rect: Readonly<FloatRect>,
  sourcePoints: readonly Readonly<{ xPt: number; yPt: number }>[],
  diagnostics?: FloatWrapPreparationDiagnostics,
): CompiledPolygonWrap {
  const kind = rect.authoredWrap;
  if (kind !== 'tight' && kind !== 'through') {
    throw new Error('Polygon compilation requires tight or through wrap');
  }
  const compileInput = {
    kind,
    imageKey: rect.imageKey,
    points: rect.wrapPolygon,
    xLeftPt: rect.xLeft,
    xRightPt: rect.xRight,
    yTopPt: rect.yTop,
    yBottomPt: rect.yBottom,
  };
  // Validate the current FloatRect before lookup so a prior valid entry cannot
  // mask malformed geometry or replace its current imageKey in the error.
  assertValidPolygonCompileInput(compileInput);
  const cacheable = polygonIsDeepFrozen(sourcePoints);
  if (cacheable) {
    const memo = compiledPolygonMemo.get(sourcePoints)?.find((candidate) => memoMatches(
      candidate,
      kind,
      rect.xLeft,
      rect.xRight,
      rect.yTop,
      rect.yBottom,
    ));
    if (memo) {
      if (diagnostics) diagnostics.polygonCacheHitCount += 1;
      return memo.compiled;
    }
  }
  if (diagnostics) diagnostics.polygonCompileCount += 1;
  const compiled = compilePolygonWrap(compileInput);
  if (cacheable) {
    const memo = Object.freeze({
      kind,
      xLeftPt: rect.xLeft,
      xRightPt: rect.xRight,
      yTopPt: rect.yTop,
      yBottomPt: rect.yBottom,
      compiled,
    });
    compiledPolygonMemo.set(sourcePoints, Object.freeze([
      memo,
      ...(compiledPolygonMemo.get(sourcePoints) ?? [])
        .slice(0, MAX_COMPILED_POLYGON_VARIANTS - 1),
    ]));
  }
  return compiled;
}

interface ExactAttributedGap {
  readonly l: ExactRational;
  readonly r: ExactRational;
  readonly leftSquareBoundary: boolean;
  readonly rightSquareBoundary: boolean;
}

/** Snapshot and compile geometry once at the line-wrap oracle boundary. */
export function prepareFloatWrap(
  floats: readonly FloatRect[],
  diagnostics?: FloatWrapPreparationDiagnostics,
): PreparedFloatWrap {
  const prepared = floats.map((float): PreparedFloatRect => {
    const sourcePoints = float.wrapPolygon;
    if (diagnostics && sourcePoints) {
      diagnostics.polygonSnapshotPointCount += sourcePoints.length;
    }
    const rect = Object.freeze({
      ...float,
      ...(sourcePoints
        ? { wrapPolygon: Object.freeze(sourcePoints.map((point) => Object.freeze({ ...point }))) }
        : {}),
    });
    const polygon = rect.authoredWrap === 'tight' || rect.authoredWrap === 'through'
      ? compiledPolygonFor(rect, sourcePoints ?? [], diagnostics)
      : null;
    return Object.freeze({
      rect,
      polygon,
      // §20.4.3.7 chooses "largest" from the positioned object, not from a
      // contour slice. Retain the full padded wrap extent so every line-band
      // projection observes one stable side even when the polygon is skewed.
      wrapMaximumLeftPt: polygon
        ? Math.min(rect.xLeft, polygon.polygonLeftPt)
        : rect.xLeft,
      wrapMaximumRightPt: polygon
        ? Math.max(rect.xRight, polygon.polygonRightPt)
        : rect.xRight,
    });
  });
  const result = Object.freeze({ floats: Object.freeze(prepared) });
  sweepEventCache.set(result, new Map());
  return result;
}

export interface LineFloatReference {
  readonly xLeftPt: number;
  readonly xRightPt: number;
  readonly readingDirection: 'ltr' | 'rtl';
}

function effectiveWrapSide(
  prepared: PreparedFloatRect,
  reference: LineFloatReference,
): Exclude<WrapSide, 'largest'> {
  const side = normalizeWrapSide(prepared.rect.side);
  if (side !== 'largest') return side;
  const leftWidth = subtractUnreducedExact(
    unreducedExactFromNumber(prepared.wrapMaximumLeftPt),
    unreducedExactFromNumber(reference.xLeftPt),
  );
  const rightWidth = subtractUnreducedExact(
    unreducedExactFromNumber(reference.xRightPt),
    unreducedExactFromNumber(prepared.wrapMaximumRightPt),
  );
  const widthComparison = compareExactRational(leftWidth, rightWidth);
  return widthComparison === 0
    ? (reference.readingDirection === 'ltr' ? 'left' : 'right')
    : widthComparison > 0 ? 'left' : 'right';
}

function floatBlockedIntervals(
  prepared: PreparedFloatRect,
  lineTopPt: number,
  probeHeightPt: number,
  paragraphLeftPt: number,
  paragraphRightPt: number,
  reference: LineFloatReference,
): ExactAttributedGap[] {
  const { rect: float, polygon } = prepared;
  const projected = polygon
    ? projectPolygonExactLineIntervals(polygon, lineTopPt, probeHeightPt)
    : [{
        l: unreducedExactFromNumber(float.xLeft),
        r: unreducedExactFromNumber(float.xRight),
      }];
  if (projected.length === 0) return [];
  const squareBoundary = polygon === null;
  const objectLeft = projected.reduce((minimum, interval) =>
    compareExactRational(interval.l, minimum) < 0 ? interval.l : minimum,
  projected[0]!.l);
  const objectRight = projected.reduce((maximum, interval) =>
    compareExactRational(interval.r, maximum) > 0 ? interval.r : maximum,
  projected[0]!.r);
  switch (effectiveWrapSide(prepared, reference)) {
    case 'left': return [{
      l: objectLeft,
      r: unreducedExactFromNumber(paragraphRightPt),
      leftSquareBoundary: squareBoundary,
      rightSquareBoundary: false,
    }];
    case 'right': return [{
      l: unreducedExactFromNumber(paragraphLeftPt),
      r: objectRight,
      leftSquareBoundary: false,
      rightSquareBoundary: squareBoundary,
    }];
    case 'bothSides': return projected.map((interval) => ({
      ...interval,
      leftSquareBoundary: squareBoundary,
      rightSquareBoundary: squareBoundary,
    }));
  }
}

/**
 * Resolve where a single line box may sit relative to the page's active floats.
 *
 * Given the line's intended top Y and minimum horizontal width, this convenience
 * boundary compiles raw geometry and returns the earliest Y plus horizontal
 * window. Production line layout uses the prepared variant below so compilation
 * occurs once per oracle/acquisition rather than in the per-line hot path.
 *
 * Two ECMA-376 wrap rules are applied, in order:
 *   1. topAndBottom floats (§20.4.2.20): a line intersecting one is pushed below
 *      it — text never sits beside a topAndBottom object.
 *   2. square floats (§20.4.2.17): text wraps around the float's rect + dist
 *      padding; tight/through use their compiled polygons (§20.4.2.18/.19).
 *      Multiple objects are composed only after each `largest` object selects
 *      its own permitted side under §20.4.3.7. The first usable gap in reading order wins.
 *
 * The prepared API can receive distinct square/polygon requirements for
 * direct geometry probes; production supplies the next atom width to both.
 */
export function resolveLineFloatWindow(
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
  // The paragraph's RAW COLUMN band, distinct from the indented text band
  // [paraX, paraX + maxWidth]. Step 1 (topAndBottom) gates by the COLUMN band —
  // §20.4.2.20 blocks the FULL column where the object sits, including the
  // paragraph's indent margins — while step 2 (square side-gap) keeps gating by
  // the narrower indented text band (§20.4.2.17). Defaults to the indented band
  // so a direct unit caller that has no separate column band stays correct.
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
): { topY: number; xOffset: number; maxWidth: number } {
  return computePreparedLineFloatWindow(
    topY,
    requiredWidth,
    probeH,
    paraX,
    maxWidth,
    prepareFloatWrap(floats),
    columnXLeftPt,
    columnXRightPt,
    reference,
    requiredWidth,
  );
}

function finiteStructuralEvents(
  probeH: number,
  prepared: PreparedFloatWrap,
): readonly number[] {
  const cache = sweepEventCache.get(prepared);
  if (!cache) throw new Error('Prepared float geometry omitted its sweep cache');
  const cached = cache.get(probeH);
  if (cached) return cached;
  const candidates = new Set<number>();
  const appendY = (value: number): void => {
    if (Number.isFinite(value)) candidates.add(value);
  };
  for (const { rect, polygon } of prepared.floats) {
    appendY(exactRationalToNumberUp(subtractUnreducedExact(
      unreducedExactFromNumber(rect.yTop),
      unreducedExactFromNumber(probeH),
    )));
    appendY(rect.yBottom);
    // The window changes at an exempt anchor line and at a top-touch threshold.
    if (rect.exemptLineTopPt !== undefined) appendY(rect.exemptLineTopPt);
    if (rect.topEdgeInclusiveFromYPt !== undefined) appendY(rect.topEdgeInclusiveFromYPt);
    if (!polygon) continue;
    for (const eventYPt of polygonLineTopEventYPts(polygon, probeH)) appendY(eventYPt);
  }
  const events = Object.freeze([...candidates].sort((left, right) => left - right));
  cache.set(probeH, events);
  return events;
}

function mergeAttributedIntervals(
  intervals: readonly ExactAttributedGap[],
): ExactAttributedGap[] {
  const sorted = intervals
    .filter((interval) => compareExactRational(interval.r, interval.l) > 0)
    .slice()
    .sort((left, right) =>
      compareExactRational(left.l, right.l)
        || compareExactRational(left.r, right.r));
  const merged: Array<{
    l: ExactRational;
    r: ExactRational;
    leftSquareBoundary: boolean;
    rightSquareBoundary: boolean;
  }> = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || compareExactRational(interval.l, previous.r) > 0) {
      merged.push({ ...interval });
      continue;
    }
    if (compareExactRational(interval.l, previous.l) === 0) {
      previous.leftSquareBoundary = previous.leftSquareBoundary
        && interval.leftSquareBoundary;
    }
    const rightComparison = compareExactRational(interval.r, previous.r);
    if (rightComparison > 0) {
      previous.r = interval.r;
      previous.rightSquareBoundary = interval.rightSquareBoundary;
    } else if (rightComparison === 0) {
      previous.rightSquareBoundary = previous.rightSquareBoundary
        && interval.rightSquareBoundary;
    }
  }
  return merged;
}

function firstUsableFreeGap(
  blocked: readonly ExactAttributedGap[],
  left: number,
  right: number,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
  readingDirection: 'ltr' | 'rtl',
): Readonly<{
  l: ExactRational;
  r: ExactRational;
  squareConstrained: boolean;
}> | null {
  const merged = mergeAttributedIntervals(blocked);
  const exactLeft = unreducedExactFromNumber(left);
  const exactRight = unreducedExactFromNumber(right);
  const gaps: Array<{
    l: ExactRational;
    r: ExactRational;
    squareConstrained: boolean;
  }> = [];
  const consider = (
    l: ExactRational,
    r: ExactRational,
    squareConstrained: boolean,
  ): void => {
    const clippedLeft = compareExactRational(exactLeft, l) >= 0 ? exactLeft : l;
    const clippedRight = compareExactRational(exactRight, r) <= 0 ? exactRight : r;
    if (compareExactRational(clippedRight, clippedLeft) > 0) {
      gaps.push({ l: clippedLeft, r: clippedRight, squareConstrained });
    }
  };
  let cursor = exactLeft;
  let cursorSquareBoundary = false;
  for (const interval of merged) {
    if (compareExactRational(interval.r, exactLeft) <= 0) {
      cursorSquareBoundary = interval.rightSquareBoundary;
      continue;
    }
    if (compareExactRational(interval.l, exactRight) >= 0) {
      consider(cursor, exactRight, cursorSquareBoundary);
      cursor = exactRight;
      break;
    }
    if (compareExactRational(interval.l, cursor) > 0) {
      consider(
        cursor,
        interval.l,
        cursorSquareBoundary || interval.leftSquareBoundary,
      );
    }
    const rightComparison = compareExactRational(interval.r, cursor);
    if (rightComparison > 0) {
      cursor = interval.r;
      cursorSquareBoundary = interval.rightSquareBoundary;
    } else if (rightComparison === 0) {
      cursorSquareBoundary = cursorSquareBoundary && interval.rightSquareBoundary;
    }
    if (compareExactRational(cursor, exactRight) >= 0) break;
  }
  if (compareExactRational(cursor, exactRight) < 0) {
    consider(cursor, exactRight, cursorSquareBoundary);
  }
  // WORD_FLOAT_GAP_FLOW selects physical gaps in reading order;
  // width admission belongs to the next
  // indivisible text atom, not a universal one-inch threshold (#1670).
  for (const gap of readingDirection === 'rtl' ? gaps.reverse() : gaps) {
    const width = subtractUnreducedExact(gap.r, gap.l);
    const requirement = Math.max(
      MIN_LINE_GAP,
      gap.squareConstrained ? squareRequiredWidth : polygonRequiredWidth,
    );
    if (compareExactRational(width, unreducedExactFromNumber(requirement)) >= 0) {
      return {
        l: gap.l,
        r: gap.r,
        squareConstrained: gap.squareConstrained,
      };
    }
  }
  return null;
}

/** Whether the unpadded polygon `points` meets the rectangle
 * [leftPt, rightPt] x [topY, topY + heightPt). */
export function polygonMeetsRect(
  points: readonly Readonly<{ xPt: number; yPt: number }>[],
  topY: number,
  heightPt: number,
  leftPt: number,
  rightPt: number,
): boolean {
  if (points.length < 3 || !(heightPt > 0) || !(rightPt > leftPt)) return false;
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  const compiled = compilePolygonWrap({
    kind: 'tight',
    imageKey: 'anchor-line-rewrap-probe',
    points,
    xLeftPt: Math.min(...xs),
    xRightPt: Math.max(...xs),
    yTopPt: Math.min(...ys),
    yBottomPt: Math.max(...ys),
  });
  const exactLeft = unreducedExactFromNumber(leftPt);
  const exactRight = unreducedExactFromNumber(rightPt);
  return projectPolygonExactLineIntervals(compiled, topY, heightPt).some((interval) =>
    compareExactRational(interval.l, exactRight) < 0
      && compareExactRational(interval.r, exactLeft) > 0);
}

/** Exact binary height of the probe that reads a tight polygon edge. */
const TIGHT_EDGE_PROBE_PT = 2 ** -10;

/** Line-top Y of the probe band that reads a tight polygon's touched edge, or
 * null when the line band does not touch one. WORD_TIGHT_WRAP_BOTTOM_EDGE: a
 * line whose top is exactly the polygon bottom wraps around the bottom edge.
 * WORD_MODE14_TIGHT_ANCHOR_TOP_TOUCH: a line whose bottom is exactly the top
 * of a polygon flagged `topEdgeInclusiveFromYPt`, starting at or below that
 * Y, wraps around the top edge. */
function tightEdgeTouch(
  float: PreparedFloatRect,
  exactTop: ExactRational,
  exactBottom: ExactRational,
): number | null {
  if (float.polygon?.kind !== 'tight') return null;
  const { rect } = float;
  if (compareExactRational(exactTop, unreducedExactFromNumber(rect.yBottom)) === 0) {
    return rect.yBottom - TIGHT_EDGE_PROBE_PT;
  }
  if (rect.topEdgeInclusiveFromYPt !== undefined
    && compareExactRational(exactTop, unreducedExactFromNumber(rect.topEdgeInclusiveFromYPt)) >= 0
    && compareExactRational(exactBottom, unreducedExactFromNumber(rect.yTop)) === 0) {
    return rect.yTop;
  }
  return null;
}

/**
 * The vertical test `lineWindowAtY` uses to decide whether `float` affects a
 * line band [topY, topY + heightPt]: a strict overlap, or a touched tight
 * polygon edge (`tightEdgeTouch`), except on an exempt anchor line.
 * Pagination uses the same predicate to decide whether a carried drawing
 * can change an earlier line.
 */
export function createFloatLineBandPredicate(
  float: FloatRect,
): (topY: number, heightPt: number) => boolean {
  const prepared = prepareFloatWrap([float]).floats[0]!;
  const { rect } = prepared;
  return (topY, heightPt) => {
    const exactTop = unreducedExactFromNumber(topY);
    const exactBottom = addUnreducedExact(exactTop, unreducedExactFromNumber(heightPt));
    const overlaps = compareExactRational(exactBottom, unreducedExactFromNumber(rect.yTop)) > 0
      && compareExactRational(exactTop, unreducedExactFromNumber(rect.yBottom)) < 0;
    if (rect.mode === 'topAndBottom') return overlaps;
    if (rect.exemptLineTopPt === topY) return false;
    return overlaps || tightEdgeTouch(prepared, exactTop, exactBottom) !== null;
  };
}

/** Whether the line band [topY, topY + probeH] lies in a tight polygon's
 * line-step region: it meets the polygon, including the touched edges of
 * `tightEdgeTouch`. An exempt anchor line (`exemptLineTopPt`) is still in the
 * region; the exemption only changes that line's own window. */
function bandMeetsTightPolygon(
  prepared: PreparedFloatWrap,
  topY: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
): boolean {
  return lowestMetTightBottom(prepared, topY, probeH, paraXLeft, paraXRight) !== null;
}

/** The lowest bottom of the tight polygons whose line-step region contains
 * the band [topY, topY + probeH] (see bandMeetsTightPolygon), or null. */
function lowestMetTightBottom(
  prepared: PreparedFloatWrap,
  topY: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
): number | null {
  const exactTop = unreducedExactFromNumber(topY);
  const exactBottom = addUnreducedExact(exactTop, unreducedExactFromNumber(probeH));
  let lowest: number | null = null;
  for (const float of prepared.floats) {
    const { rect } = float;
    if (float.polygon?.kind !== 'tight' || rect.mode !== 'square') continue;
    if (!floatOverlapsColumnX(rect as FloatRect, paraXLeft, paraXRight)) continue;
    const meets = tightEdgeTouch(float, exactTop, exactBottom) !== null
      || (compareExactRational(exactBottom, unreducedExactFromNumber(rect.yTop)) > 0
        && compareExactRational(exactTop, unreducedExactFromNumber(rect.yBottom)) < 0);
    if (meets) lowest = Math.max(lowest ?? Number.NEGATIVE_INFINITY, rect.yBottom);
  }
  return lowest;
}

function lineWindowAtY(
  topY: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number,
  columnXRightPt: number,
  reference: LineFloatReference,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
): { topY: number; xOffset: number; maxWidth: number } | null {
  const exactTop = unreducedExactFromNumber(topY);
  const exactBottom = addUnreducedExact(
    exactTop,
    unreducedExactFromNumber(probeH),
  );
  const exactRectY = (value: number): ExactRational => unreducedExactFromNumber(value);
  if (prepared.floats.some(({ rect }) => rect.mode === 'topAndBottom'
    && floatOverlapsColumnX(rect as FloatRect, columnXLeftPt, columnXRightPt)
    && compareExactRational(exactBottom, exactRectY(rect.yTop)) > 0
    && compareExactRational(exactTop, exactRectY(rect.yBottom)) < 0)) return null;
  const blocked: ExactAttributedGap[] = [];
  for (const float of prepared.floats) {
    const { rect } = float;
    if (rect.mode !== 'square') continue;
    if (rect.exemptLineTopPt === topY) continue;
    const edge = tightEdgeTouch(float, exactTop, exactBottom);
    if (edge === null && (compareExactRational(exactBottom, exactRectY(rect.yTop)) <= 0
      || compareExactRational(exactTop, exactRectY(rect.yBottom)) >= 0)) continue;
    if (!floatOverlapsColumnX(rect as FloatRect, paraXLeft, paraXRight)) continue;
    const intervals = edge === null
      ? floatBlockedIntervals(float, topY, probeH, paraXLeft, paraXRight, reference)
      : floatBlockedIntervals(
          float, edge, TIGHT_EDGE_PROBE_PT, paraXLeft, paraXRight, reference,
        );
    if (intervals.length === 0) continue;
    blocked.push(...intervals);
  }
  if (blocked.length === 0) return { topY, xOffset: 0, maxWidth };
  const best = firstUsableFreeGap(
    blocked,
    paraXLeft,
    paraXRight,
    polygonRequiredWidth,
    squareRequiredWidth,
    reference.readingDirection,
  );
  if (!best) return null;
  const zero: ExactRational = { numerator: 0n, denominator: 1n };
  const xOffsetCandidate = subtractUnreducedExact(
    best.l,
    unreducedExactFromNumber(paraXLeft),
  );
  const xOffset = compareExactRational(xOffsetCandidate, zero) > 0
    ? xOffsetCandidate
    : zero;
  const exactParagraphLeft = unreducedExactFromNumber(paraXLeft);
  let xOffsetNumber = exactRationalToNumberUp(xOffset);
  let returnedStartNumber = paraXLeft + xOffsetNumber;
  let returnedStart = unreducedExactFromNumber(returnedStartNumber);
  if (compareExactRational(returnedStart, best.l) < 0) {
    const targetStartNumber = exactRationalToNumberUp(best.l);
    xOffsetNumber = exactRationalToNumberUp(subtractUnreducedExact(
      unreducedExactFromNumber(targetStartNumber),
      exactParagraphLeft,
    ));
    returnedStartNumber = paraXLeft + xOffsetNumber;
    returnedStart = unreducedExactFromNumber(returnedStartNumber);
  }
  if (compareExactRational(returnedStart, best.l) < 0) {
    throw new Error('Exact float window could not represent a contained start');
  }
  const exactParagraphRight = unreducedExactFromNumber(paraXRight);
  const exactAvailableEnd =
    compareExactRational(best.r, exactParagraphRight) <= 0
      ? best.r
      : exactParagraphRight;
  const returnedEndNumber = exactRationalToNumberDown(exactAvailableEnd);
  const availableWidth = subtractUnreducedExact(
    unreducedExactFromNumber(returnedEndNumber),
    returnedStart,
  );
  const nonnegativeWidth = compareExactRational(availableWidth, zero) > 0
    ? availableWidth
    : zero;
  const maxWidthNumber = exactRationalToNumberDown(nonnegativeWidth);
  const reconstructedEnd = unreducedExactFromNumber(
    returnedStartNumber + maxWidthNumber,
  );
  if (compareExactRational(reconstructedEnd, exactAvailableEnd) > 0) {
    throw new Error('Exact float window could not represent a contained end');
  }
  return {
    topY,
    xOffset: xOffsetNumber,
    maxWidth: maxWidthNumber,
  };
}

/**
 * Left edge of the first free horizontal gap of the band [topY, topY + probeH]
 * between `leftPt` and `rightPt`, around the square/tight/through exclusions of
 * `floats`, or null when the band is fully blocked (including by a
 * topAndBottom object). Gaps narrower than MIN_LINE_GAP are slivers.
 *
 * Used by WORD_MODE14_COLUMN_LINE_START_ORIGIN: in compatibility mode 14 a
 * column-relative horizontal offset is measured from this edge.
 */
export function firstFreeGapLeftPt(
  floats: readonly FloatRect[],
  topY: number,
  probeH: number,
  leftPt: number,
  rightPt: number,
): number | null {
  const prepared = prepareFloatWrap(floats);
  const exactTop = unreducedExactFromNumber(topY);
  const exactBottom = addUnreducedExact(exactTop, unreducedExactFromNumber(probeH));
  const reference: LineFloatReference = { xLeftPt: leftPt, xRightPt: rightPt, readingDirection: 'ltr' };
  const blocked: ExactAttributedGap[] = [];
  for (const float of prepared.floats) {
    const { rect } = float;
    if (compareExactRational(exactBottom, unreducedExactFromNumber(rect.yTop)) <= 0
      || compareExactRational(exactTop, unreducedExactFromNumber(rect.yBottom)) >= 0) continue;
    if (!floatOverlapsColumnX(rect as FloatRect, leftPt, rightPt)) continue;
    if (rect.mode === 'topAndBottom') return null;
    blocked.push(...floatBlockedIntervals(float, topY, probeH, leftPt, rightPt, reference));
  }
  let cursor = unreducedExactFromNumber(leftPt);
  const exactRight = unreducedExactFromNumber(rightPt);
  const minimum = unreducedExactFromNumber(MIN_LINE_GAP);
  for (const interval of mergeAttributedIntervals(blocked)) {
    if (compareExactRational(interval.r, cursor) <= 0) continue;
    const gapRight = compareExactRational(interval.l, exactRight) < 0 ? interval.l : exactRight;
    if (compareExactRational(subtractUnreducedExact(gapRight, cursor), minimum) >= 0) {
      return exactRationalToNumberUp(cursor);
    }
    cursor = interval.r;
    if (compareExactRational(cursor, exactRight) >= 0) return null;
  }
  return compareExactRational(subtractUnreducedExact(exactRight, cursor), minimum) >= 0
    ? exactRationalToNumberUp(cursor)
    : null;
}

interface AffineBoundary {
  readonly exact: Readonly<{
    slope: ExactRational;
    intercept: ExactRational;
  }>;
  readonly square: boolean;
}

interface AffineBlockedInterval {
  readonly left: AffineBoundary;
  readonly right: AffineBoundary;
}

export interface LineFloatSweepDiagnostics {
  readonly compiledIntersectionCount: number;
  readonly compiledContourSpanCount: number;
  readonly compileOrderComparisonCount: number;
  /** Edge-slot reads for complete vertex scans plus local intersection updates: O(V² + K). */
  readonly compilePairMembershipVisitCount: number;
  readonly structuralEventCount: number;
  readonly localRootCandidateCount: number;
  readonly localRootEventCount: number;
  readonly evaluatedYCount: number;
}

interface MutableLineFloatSweepDiagnostics {
  compiledIntersectionCount: number;
  compiledContourSpanCount: number;
  compileOrderComparisonCount: number;
  compilePairMembershipVisitCount: number;
  structuralEventCount: number;
  localRootCandidateCount: number;
  localRootEventCount: number;
  evaluatedYCount: number;
}

function exactAffineValueAt(boundary: AffineBoundary, y: number): ExactRational {
  return addUnreducedExact(
    multiplyUnreducedExact(boundary.exact.slope, unreducedExactFromNumber(y)),
    boundary.exact.intercept,
  );
}

function sameAffine(left: AffineBoundary, right: AffineBoundary): boolean {
  return compareExactRational(left.exact.slope, right.exact.slope) === 0
    && compareExactRational(left.exact.intercept, right.exact.intercept) === 0;
}

function compareAtRight(left: AffineBoundary, right: AffineBoundary, y: number): number {
  return compareExactRational(
    exactAffineValueAt(left, y),
    exactAffineValueAt(right, y),
  ) || compareExactRational(left.exact.slope, right.exact.slope);
}

function exactAffineRoot(
  left: AffineBoundary['exact'],
  right: AffineBoundary['exact'],
  delta: number,
): ExactRational | null {
  const slope = subtractUnreducedExact(left.slope, right.slope);
  if (slope.numerator === 0n) return null;
  const intercept = subtractUnreducedExact(left.intercept, right.intercept);
  return divideUnreducedExact(
    subtractUnreducedExact(unreducedExactFromNumber(delta), intercept),
    slope,
  );
}

function appendLocalRoot(
  roots: number[],
  root: ExactRational | null,
  lower: number,
  upper: number,
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): void {
  if (root === null
    || compareExactRational(root, unreducedExactFromNumber(lower)) <= 0
    || compareExactRational(root, unreducedExactFromNumber(upper)) >= 0) return;
  roots.push(exactRationalToNumberUp(root));
  if (diagnostics) diagnostics.localRootCandidateCount += 1;
}

function selectEnvelope(
  candidates: readonly AffineBoundary[],
  kind: 'min' | 'max',
  lower: number,
  upper: number,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBoundary {
  let winner = candidates[0]!;
  for (const candidate of candidates.slice(1)) {
    const comparison = compareAtRight(candidate, winner, lower);
    if ((kind === 'min' && comparison < 0) || (kind === 'max' && comparison > 0)) {
      winner = candidate;
    }
  }
  let square = winner.square;
  for (const candidate of candidates) {
    if (candidate === winner) continue;
    if (sameAffine(candidate, winner)) {
      square = square && candidate.square;
      continue;
    }
    const canTakeOver = kind === 'min'
      ? compareExactRational(candidate.exact.slope, winner.exact.slope) < 0
      : compareExactRational(candidate.exact.slope, winner.exact.slope) > 0;
    if (canTakeOver) {
      appendLocalRoot(
        roots,
        exactAffineRoot(candidate.exact, winner.exact, 0),
        lower,
        upper,
        diagnostics,
      );
    }
  }
  return { exact: winner.exact, square };
}

function constantBoundary(value: number, square = false): AffineBoundary {
  return {
    exact: {
      slope: { numerator: 0n, denominator: 1n },
      intercept: unreducedExactFromNumber(value),
    },
    square,
  };
}

function objectAffineIntervals(
  prepared: PreparedFloatRect,
  probeH: number,
  lower: number,
  upper: number,
  paragraphLeftPt: number,
  paragraphRightPt: number,
  reference: LineFloatReference,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBlockedInterval[] {
  const { rect, polygon } = prepared;
  const midpoint = exactBinary64Midpoint(lower, upper);
  const exactMidpoint = unreducedExactFromNumber(midpoint);
  const exactMidpointBottom = addUnreducedExact(
    exactMidpoint,
    unreducedExactFromNumber(probeH),
  );
  if (compareExactRational(
    exactMidpointBottom,
    unreducedExactFromNumber(rect.yTop),
  ) <= 0 || compareExactRational(
    exactMidpoint,
    unreducedExactFromNumber(rect.yBottom),
  ) >= 0) return [];
  let intervals: AffineBlockedInterval[] = polygon
    ? polygonBandExactIntervalFunctions(polygon, probeH, lower, upper)
      .map((interval) => ({
        left: { exact: interval.left, square: false },
        right: { exact: interval.right, square: false },
      }))
    : [{
        left: constantBoundary(rect.xLeft, true),
        right: constantBoundary(rect.xRight, true),
      }];
  if (intervals.length === 0) return [];
  const objectLeft = selectEnvelope(
    intervals.map((interval) => interval.left),
    'min', lower, upper, roots, diagnostics,
  );
  const objectRight = selectEnvelope(
    intervals.map((interval) => interval.right),
    'max', lower, upper, roots, diagnostics,
  );
  if (polygon?.kind === 'tight') intervals = [{ left: objectLeft, right: objectRight }];
  switch (effectiveWrapSide(prepared, reference)) {
    case 'left': return [{
      left: objectLeft,
      right: constantBoundary(paragraphRightPt),
    }];
    case 'right': return [{
      left: constantBoundary(paragraphLeftPt),
      right: objectRight,
    }];
    case 'bothSides': return intervals;
  }
}

function mergeAffineIntervalsAtRight(
  intervals: readonly AffineBlockedInterval[],
  lower: number,
  upper: number,
  roots: number[],
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): AffineBlockedInterval[] {
  const sorted = intervals.slice().sort((left, right) =>
    compareAtRight(left.left, right.left, lower)
    || compareAtRight(left.right, right.right, lower));
  for (let index = 0; index + 1 < sorted.length; index += 1) {
    appendLocalRoot(
      roots,
      exactAffineRoot(
        sorted[index]!.left.exact,
        sorted[index + 1]!.left.exact,
        0,
      ),
      lower,
      upper,
      diagnostics,
    );
  }
  const merged: AffineBlockedInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous) {
      merged.push(interval);
      continue;
    }
    appendLocalRoot(
      roots,
      exactAffineRoot(interval.left.exact, previous.right.exact, 0),
      lower,
      upper,
      diagnostics,
    );
    const contact = compareAtRight(interval.left, previous.right, lower);
    if (contact > 0) {
      merged.push(interval);
      continue;
    }
    const right = selectEnvelope(
      [previous.right, interval.right],
      'max', lower, upper, roots, diagnostics,
    );
    const left = sameAffine(previous.left, interval.left)
      ? {
          exact: previous.left.exact,
          square: previous.left.square && interval.left.square,
        }
      : previous.left;
    merged[merged.length - 1] = { left, right };
  }
  return merged;
}

function nextLocalSweepEvent(
  lower: number,
  upper: number,
  probeH: number,
  paraXLeft: number,
  paraXRight: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number,
  columnXRightPt: number,
  reference: LineFloatReference,
  polygonRequiredWidth: number,
  squareRequiredWidth: number,
  diagnostics: MutableLineFloatSweepDiagnostics | null,
): number | null {
  const midpoint = exactBinary64Midpoint(lower, upper);
  const exactMidpoint = unreducedExactFromNumber(midpoint);
  const exactMidpointBottom = addUnreducedExact(
    exactMidpoint,
    unreducedExactFromNumber(probeH),
  );
  if (prepared.floats.some(({ rect }) => rect.mode === 'topAndBottom'
    && floatOverlapsColumnX(rect as FloatRect, columnXLeftPt, columnXRightPt)
    && compareExactRational(
      exactMidpointBottom,
      unreducedExactFromNumber(rect.yTop),
    ) > 0
    && compareExactRational(
      exactMidpoint,
      unreducedExactFromNumber(rect.yBottom),
    ) < 0)) return null;
  const roots: number[] = [];
  const intervals: AffineBlockedInterval[] = [];
  for (const float of prepared.floats) {
    const { rect } = float;
    if (rect.mode !== 'square') continue;
    if (!floatOverlapsColumnX(rect as FloatRect, paraXLeft, paraXRight)) continue;
    intervals.push(...objectAffineIntervals(
      float, probeH, lower, upper, paraXLeft, paraXRight,
      reference, roots, diagnostics,
    ));
  }
  if (intervals.length === 0) return null;
  const merged = mergeAffineIntervalsAtRight(
    intervals, lower, upper, roots, diagnostics,
  );
  const pageLeft = constantBoundary(paraXLeft);
  const pageRight = constantBoundary(paraXRight);
  const gaps: Array<{ exactWidth: AffineBoundary['exact'] }> = [];
  const appendGapRoot = (
    left: AffineBoundary,
    right: AffineBoundary,
    squareConstrained: boolean,
  ): void => {
    const exactWidth = {
      slope: subtractUnreducedExact(right.exact.slope, left.exact.slope),
      intercept: subtractUnreducedExact(right.exact.intercept, left.exact.intercept),
    };
    gaps.push({ exactWidth });
    const requirement = Math.max(
      MIN_LINE_GAP,
      squareConstrained ? squareRequiredWidth : polygonRequiredWidth,
    );
    const widthAtLower = addUnreducedExact(
      multiplyUnreducedExact(exactWidth.slope, unreducedExactFromNumber(lower)),
      exactWidth.intercept,
    );
    if (compareExactRational(
      widthAtLower,
      unreducedExactFromNumber(requirement),
    ) < 0 && exactWidth.slope.numerator > 0n) {
      appendLocalRoot(
        roots,
        exactAffineRoot(right.exact, left.exact, requirement),
        lower,
        upper,
        diagnostics,
      );
    }
  };
  let gapLeft = pageLeft;
  for (const interval of merged) {
    appendGapRoot(
      gapLeft,
      interval.left,
      gapLeft.square || interval.left.square,
    );
    gapLeft = interval.right;
  }
  appendGapRoot(gapLeft, pageRight, gapLeft.square);
  let widestGap = gaps[0];
  for (const gap of gaps.slice(1)) {
    const comparison = compareExactRational(
      addUnreducedExact(
        multiplyUnreducedExact(gap.exactWidth.slope, unreducedExactFromNumber(lower)),
        gap.exactWidth.intercept,
      ),
      addUnreducedExact(
        multiplyUnreducedExact(
          widestGap!.exactWidth.slope,
          unreducedExactFromNumber(lower),
        ),
        widestGap!.exactWidth.intercept,
      ),
    ) || compareExactRational(gap.exactWidth.slope, widestGap!.exactWidth.slope);
    if (comparison > 0) widestGap = gap;
  }
  if (widestGap) {
    for (const gap of gaps) {
      if (gap === widestGap
        || compareExactRational(
          gap.exactWidth.slope,
          widestGap.exactWidth.slope,
        ) <= 0) continue;
      appendLocalRoot(
        roots,
        exactAffineRoot(gap.exactWidth, widestGap.exactWidth, 0),
        lower,
        upper,
        diagnostics,
      );
    }
  }
  if (roots.length === 0) return null;
  return Math.min(...roots);
}

/** Resource limit on ordinary tight line steps per line query (library
 * policy, not an OOXML rule). Within it every grid position is tested; past
 * it, a step jumps to the first grid position below the tight polygons the
 * failed band meets (see `computePreparedLineFloatWindowCore`). */
const TIGHT_LINE_STEP_LIMIT = 20_000;

/** Upper bound on the magnitude of line-top Ys the float line search visits.
 * DrawingML positions and extents are ST_CoordinateUnqualified (ECMA-376
 * §20.1.10.19), bounded by ±27,273,042,316,900 EMU = ±(2^31 − 1) pt, and a
 * page-relative line or object Y is a sum of a few such terms, so 2^33 pt is
 * past every page. Beyond it the search stops (see the domain contract in
 * `computePreparedLineFloatWindowCore`). */
export const LINE_SEARCH_Y_LIMIT_PT = 2 ** 33;
/** Finest retry pitch: one twip, the smallest positive integer
 * ST_SignedTwipsMeasure (§17.18.81) a w:spacing/@w:line exact or atLeast
 * value can express. Finer pitches (universal measures, small auto
 * multiples) retry at this pitch. */
export const LINE_SEARCH_MIN_PITCH_PT = 0.05;
/** Coarsest retry pitch: 1,584 pt (22 in), Word's largest page height and
 * its largest exact/atLeast line spacing. A taller line cannot sit on any
 * page, so a finer retry cannot move it onto one. */
export const LINE_SEARCH_MAX_PITCH_PT = 1584;

/** The line-retry pitch for a line of height `probeH`, clamped to the search
 * domain, or null when the line has no positive finite height (no retries). */
export function lineSearchPitch(probeH: number): number | null {
  if (!(probeH > 0) || !Number.isFinite(probeH)) return null;
  return Math.min(Math.max(probeH, LINE_SEARCH_MIN_PITCH_PT), LINE_SEARCH_MAX_PITCH_PT);
}

/**
 * The first line-grid position strictly below `y`: `originY + k * pitchPt`
 * for the smallest integer k >= 1 whose binary64 value exceeds `y`.
 *
 * Domain (asserted): |originY|, |y| <= LINE_SEARCH_Y_LIMIT_PT and pitchPt in
 * [LINE_SEARCH_MIN_PITCH_PT, LINE_SEARCH_MAX_PITCH_PT]. There every relevant
 * index is below 2^40, so k and k * pitchPt are exact integers times a pitch
 * at least 2^14 ulps of any coordinate: grid values strictly increase with k
 * and are finite. The quotient estimate is then off by at most one, so each
 * correction loop runs at most twice (asserted).
 */
export function lineGridAfter(originY: number, pitchPt: number, y: number): number {
  if (!(Math.abs(originY) <= LINE_SEARCH_Y_LIMIT_PT && Math.abs(y) <= LINE_SEARCH_Y_LIMIT_PT
    && pitchPt >= LINE_SEARCH_MIN_PITCH_PT && pitchPt <= LINE_SEARCH_MAX_PITCH_PT)) {
    throw new RangeError('Line grid query outside the float line-search domain');
  }
  const at = (k: number): number => originY + k * pitchPt;
  let k = Math.max(1, Math.floor((y - originY) / pitchPt) + 1);
  let corrections = 0;
  while (at(k) <= y) {
    k += 1;
    corrections += 1;
  }
  while (k > 1 && at(k - 1) > y) {
    k -= 1;
    corrections += 1;
  }
  if (corrections > 2) throw new Error('Line grid estimate violated its one-step error bound');
  return at(k);
}

/**
 * Hot line query over immutable, acquisition-compiled float geometry.
 *
 * Polygon vertices/intersections and rectangle edges form finite structural
 * top-Y slabs. Inside one slab, the compiled active contour pairs yield affine
 * boundaries. The local queue retains only current extrema, adjacent merged
 * boundaries, per-gap threshold roots, and roots where another
 * adjacent-boundary gap overtakes the current widest gap. A `largest` side is
 * invariant across these slabs because §20.4.3.7 selects it from the object's
 * retained maximum horizontal extent. After the earliest local root, the sweep
 * rebuilds that ordering. Every advance therefore consumes one event in a
 * finite affine arrangement. This proves termination without a global all-pairs
 * candidate set, sampling, bisection, or a pass cap.
 */
function computePreparedLineFloatWindowCore(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
  diagnostics: MutableLineFloatSweepDiagnostics | null = null,
): { topY: number; xOffset: number; maxWidth: number } {
  const paraXLeft = paraX;
  const paraXRight = paraX + maxWidth;
  const structuralEvents = finiteStructuralEvents(probeH, prepared);
  if (diagnostics) {
    diagnostics.structuralEventCount = structuralEvents.length;
    for (const { polygon } of prepared.floats) {
      if (!polygon) continue;
      diagnostics.compiledIntersectionCount += polygon.intersectionCount;
      diagnostics.compiledContourSpanCount += polygon.contourSpans.length;
      diagnostics.compileOrderComparisonCount += polygon.compileOrderComparisonCount;
      diagnostics.compilePairMembershipVisitCount += polygon.compilePairMembershipVisitCount;
    }
  }
  const evaluate = (candidateY: number) => {
    if (diagnostics) diagnostics.evaluatedYCount += 1;
    return lineWindowAtY(
      candidateY, probeH, paraXLeft, paraXRight, maxWidth, prepared,
      columnXLeftPt, columnXRightPt, reference,
      polygonRequiredWidth, squareRequiredWidth,
    );
  };
  const current = evaluate(topY);
  if (current) return current;
  // One candidate walk, WORD_TIGHT_WRAP_LINE_STEP_ADVANCE included. Every
  // candidate is tested with the same predicate (`evaluate`, i.e.
  // `lineWindowAtY`); the first one that fits is the answer, so no returned
  // line collides with a wrap region. Candidates come from two sources:
  //   - line steps: while a failed candidate reached on the line grid
  //     (topY + k * probeH, the Word retry) still meets a tight polygon, or
  //     while the open structural slab ahead is a tight-region slab, the next
  //     candidate is the next grid position (`lineGridAfter`);
  //   - the exact event sweep elsewhere: the earliest local root of the slab
  //     ahead (`nextLocalSweepEvent`), else the slab's upper structural event.
  // Tight-region membership of a band changes only at structural events
  // (polygon top - probeH, bottom, exempt and top-touch Ys), so a slab
  // midpoint decides the whole open slab.
  //
  // Termination and bounds: every candidate is strictly greater than the
  // previous one (checked below). Ordinary line steps are capped by
  // TIGHT_LINE_STEP_LIMIT. Past the cap a line step jumps to the first grid
  // position below the lowest bottom of the tight polygons the failed band
  // meets, which leaves at least one polygon permanently behind (at most B
  // such jumps), or, when the band meets none (a tight slab entered from its
  // top edge), advances one grid position: that lands inside the slab, where
  // the next step is such a jump, or past the slab's upper event (at most E
  // such steps). Sweep candidates consume the finite
  // structural and local events. Each candidate costs one `lineWindowAtY`,
  // one O(B) region test and an O(1) grid search (`lineGridAfter`), so a
  // query is O((LIMIT + E + B + local events) * B log B) for every input in
  // the domain below. Grid positions skipped by a past-the-cap jump
  // are the only positions not tested; below the cap the walk never returns a
  // Y above the first usable candidate.
  //
  // Numeric domain of the line grid: retries use the pitch
  // `lineSearchPitch(probeH)` (1 twip to 1,584 pt; bands keep the line's own
  // height), and a line top takes part in tight retries only within
  // [-LIMIT, LIMIT] (LINE_SEARCH_Y_LIMIT_PT, past every page), where
  // `lineGridAfter` is exact, finite and O(1). "Earliest" is stated on that
  // grid. Beyond the domain (off every page) the exact sweep alone places the
  // line, so extreme binary64 inputs keep origin/main's behaviour.
  const pitch = Math.abs(topY) <= LINE_SEARCH_Y_LIMIT_PT ? lineSearchPitch(probeH) : null;
  const tightAt = (y: number): boolean =>
    pitch !== null && Math.abs(y) <= LINE_SEARCH_Y_LIMIT_PT
    && bandMeetsTightPolygon(prepared, y, probeH, paraXLeft, paraXRight);
  const gridAfter = (y: number): number =>
    lineGridAfter(topY, pitch!, Math.min(y, LINE_SEARCH_Y_LIMIT_PT));
  let stepBudget = TIGHT_LINE_STEP_LIMIT;
  const lineStepAfter = (y: number): number => {
    if (stepBudget > 0) {
      stepBudget -= 1;
      return gridAfter(y);
    }
    const bottom = lowestMetTightBottom(prepared, y, probeH, paraXLeft, paraXRight);
    return gridAfter(bottom === null ? y : Math.max(bottom, y));
  };
  let cursor = topY;
  // topY is grid position 0.
  let cursorOnGrid = true;
  let structuralIndex = 0;
  for (;;) {
    while (structuralIndex < structuralEvents.length
      && structuralEvents[structuralIndex]! <= cursor) structuralIndex += 1;
    let next: number;
    let nextOnGrid: boolean;
    if (cursorOnGrid && tightAt(cursor)) {
      next = lineStepAfter(cursor);
      nextOnGrid = true;
    } else if (structuralIndex < structuralEvents.length) {
      const upper = structuralEvents[structuralIndex]!;
      const midpoint = exactBinary64Midpoint(cursor, upper);
      // An open slab with no binary64 value inside has no line tops.
      if (midpoint > cursor && midpoint < upper && tightAt(midpoint)) {
        next = lineStepAfter(cursor);
        nextOnGrid = true;
      } else {
        const localEvent = nextLocalSweepEvent(
          cursor, upper, probeH, paraXLeft, paraXRight, prepared,
          columnXLeftPt, columnXRightPt, reference,
          polygonRequiredWidth, squareRequiredWidth, diagnostics,
        );
        if (localEvent !== null && diagnostics) diagnostics.localRootEventCount += 1;
        next = localEvent ?? upper;
        nextOnGrid = false;
      }
    } else {
      // Past the last event no object meets a band that starts strictly
      // below every object. Keep layout total rather than throwing when the
      // last event itself is blocked (a line resting on a tight bottom edge).
      const lowest = Math.max(cursor, ...prepared.floats.map(({ rect }) => rect.yBottom));
      let terminalY = pitch !== null && lowest <= LINE_SEARCH_Y_LIMIT_PT
        ? gridAfter(lowest)
        : lowest + Number.EPSILON * Math.max(1, Math.abs(lowest));
      if (!Number.isFinite(terminalY)) terminalY = lowest;
      return evaluate(terminalY) ?? { topY: terminalY, xOffset: 0, maxWidth };
    }
    if (!(next > cursor)) {
      throw new Error('Float line-window search violated strictly increasing candidate progress');
    }
    const window = evaluate(next);
    if (window) return window;
    cursor = next;
    cursorOnGrid = nextOnGrid;
  }
}

export function computePreparedLineFloatWindow(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
): { topY: number; xOffset: number; maxWidth: number } {
  return computePreparedLineFloatWindowCore(
    topY, polygonRequiredWidth, probeH, paraX, maxWidth, prepared,
    columnXLeftPt, columnXRightPt, reference, squareRequiredWidth,
  );
}

/** Internal instrumentation for deterministic solver-complexity regression tests. */
export function computePreparedLineFloatWindowWithDiagnostics(
  topY: number,
  polygonRequiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  prepared: PreparedFloatWrap,
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
  reference: LineFloatReference = {
    xLeftPt: paraX,
    xRightPt: paraX + maxWidth,
    readingDirection: 'ltr',
  },
  squareRequiredWidth: number = polygonRequiredWidth,
): Readonly<{
  window: { topY: number; xOffset: number; maxWidth: number };
  diagnostics: LineFloatSweepDiagnostics;
}> {
  const diagnostics: MutableLineFloatSweepDiagnostics = {
    compiledIntersectionCount: 0,
    compiledContourSpanCount: 0,
    compileOrderComparisonCount: 0,
    compilePairMembershipVisitCount: 0,
    structuralEventCount: 0,
    localRootCandidateCount: 0,
    localRootEventCount: 0,
    evaluatedYCount: 0,
  };
  const window = computePreparedLineFloatWindowCore(
    topY, polygonRequiredWidth, probeH, paraX, maxWidth, prepared,
    columnXLeftPt, columnXRightPt, reference, squareRequiredWidth, diagnostics,
  );
  return Object.freeze({ window, diagnostics: Object.freeze({ ...diagnostics }) });
}

/**
 * If y is inside a topAndBottom float that horizontally overlaps the paragraph's
 * column band [paraXLeft, paraXRight], return that float's bottom; otherwise
 * return y. Mirrors `resolveLineFloatWindow` step 1: §20.4.2.20 excludes text
 * only where THIS OBJECT is placed, so a float anchored in another newspaper
 * column (§17.6.4) — the page float set is shared across columns — is filtered
 * out via the shared `floatOverlapsColumnX` predicate.
 */
export function skipPastTopAndBottom(
  y: number,
  floats: FloatRect[],
  paraXLeft: number,
  paraXRight: number,
): number {
  const consumedBottoms = new Set<number>();
  for (;;) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (!floatOverlapsColumnX(f, paraXLeft, paraXRight)) continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    if (!Number.isFinite(next) || next < y || consumedBottoms.has(next)) {
      throw new Error('Top-and-bottom solver violated strictly increasing finite-bottom progress');
    }
    consumedBottoms.add(next);
    y = next;
  }
}
