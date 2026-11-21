import type {
  LayoutRect,
  Matrix2DData,
  SectionFlow,
  PointPt,
  SectionRegionCoordinateSpace,
  UprightResourceOrientation,
  WritingMode,
} from './types.js';

/** Local counter-turn that keeps a non-text graphic upright in a vertical
 * page of this writing mode: the inverse of the section frame's quarter turn.
 * Horizontal pages need none. */
export function uprightResourceOrientation(
  verticalPageFrame: boolean | undefined,
  pageWritingMode: WritingMode,
): UprightResourceOrientation | undefined {
  if (!verticalPageFrame) return undefined;
  return pageWritingMode === 'sideways-lr'
    ? 'upright-physical-counter-clockwise'
    : 'upright-physical';
}

export type PhysicalPageExtent = Readonly<{
  widthPt: number;
  heightPt: number;
}>;

export type RectEdgeRecord<T> = Readonly<{
  top: T;
  right: T;
  bottom: T;
  left: T;
}>;

export function writingModeFromTextDirection(textDirection: string): WritingMode {
  switch (textDirection) {
    case 'tb':
    case 'tbV':
    case 'lrTb':
    case 'lrTbV':
      return 'horizontal-tb';
    case 'rl':
    case 'rlV':
    case 'tbRl':
    case 'tbRlV':
      return 'vertical-rl';
    case 'btLr':
      // Compatibility rule `word-section-btlr-tbrl-page-frame`; the evidence
      // resolves to the table-driven coordinate-space regression test.
      return 'vertical-rl';
    case 'lr':
    case 'lrV':
      return 'vertical-lr';
    case 'tbLrV':
      // Part 4 §14.11.7 (≡ lrV); its page frame is the compatibility rule
      // `word-section-tblrv-left-page-frame`.
      return 'vertical-lr';
    default:
      throw new RangeError(`Unsupported Transitional text direction ${JSON.stringify(textDirection)}`);
  }
}

/** Writing mode of one section context. The authored token decides it unless
 * a validated section flow fact selects its canonical frame; that fact is only
 * normalized from a producer's private wire, never from a token. */
export function sectionWritingMode(section: Readonly<{
  textDirection: string;
  sectionFlow?: SectionFlow | null;
}>): WritingMode {
  if (section.sectionFlow == null) return writingModeFromTextDirection(section.textDirection);
  if (section.sectionFlow !== 'bottomToTop' || section.textDirection !== 'btLr') {
    throw new RangeError(
      `Section flow ${String(section.sectionFlow)} contradicts text direction ${JSON.stringify(section.textDirection)}`,
    );
  }
  return 'sideways-lr';
}

function requirePage(page: PhysicalPageExtent): void {
  if (!Number.isFinite(page.widthPt) || !Number.isFinite(page.heightPt)
    || page.widthPt <= 0 || page.heightPt <= 0) {
    throw new RangeError('Physical page extents must be positive and finite');
  }
}

function requirePoint(point: PointPt): void {
  if (!Number.isFinite(point.xPt) || !Number.isFinite(point.yPt)) {
    throw new RangeError('Point coordinates must be finite');
  }
}

function requireMatrix(matrix: Matrix2DData): void {
  if (![matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite)) {
    throw new RangeError('Matrix coefficients must be finite');
  }
}

function requireRect(rect: LayoutRect): void {
  requirePoint(rect);
  if (!Number.isFinite(rect.widthPt) || !Number.isFinite(rect.heightPt)
    || rect.widthPt < 0 || rect.heightPt < 0) {
    throw new RangeError('Rectangle extents must be finite and non-negative');
  }
}

export function logicalPageExtent(
  physicalPage: PhysicalPageExtent,
  writingMode: WritingMode,
): PhysicalPageExtent {
  requirePage(physicalPage);
  switch (writingMode) {
    case 'horizontal-tb':
      return { widthPt: physicalPage.widthPt, heightPt: physicalPage.heightPt };
    case 'vertical-rl':
    case 'vertical-lr':
    case 'sideways-lr':
      return { widthPt: physicalPage.heightPt, heightPt: physicalPage.widthPt };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function uprightPhysicalExtent(
  logicalSectionExtent: PhysicalPageExtent,
  writingMode: WritingMode,
): PhysicalPageExtent {
  requirePage(logicalSectionExtent);
  switch (writingMode) {
    case 'horizontal-tb':
      return {
        widthPt: logicalSectionExtent.widthPt,
        heightPt: logicalSectionExtent.heightPt,
      };
    case 'vertical-rl':
    case 'vertical-lr':
    case 'sideways-lr':
      return {
        widthPt: logicalSectionExtent.heightPt,
        heightPt: logicalSectionExtent.widthPt,
      };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function logicalToPhysicalMatrix(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): Matrix2DData {
  requirePage(page);
  switch (writingMode) {
    case 'horizontal-tb':
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    case 'vertical-rl':
      return { a: 0, b: 1, c: -1, d: 0, e: page.widthPt, f: 0 };
    case 'vertical-lr':
      return { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 };
    case 'sideways-lr':
      // Counter-clockwise quarter turn: logical inline +x runs physically
      // upward from the bottom edge, logical block +y runs rightward.
      return { a: 0, b: -1, c: 1, d: 0, e: 0, f: page.heightPt };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function physicalToLogicalMatrix(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): Matrix2DData {
  requirePage(page);
  switch (writingMode) {
    case 'horizontal-tb':
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    case 'vertical-rl':
      return { a: 0, b: -1, c: 1, d: 0, e: 0, f: page.widthPt };
    case 'vertical-lr':
      return { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 };
    case 'sideways-lr':
      return { a: 0, b: 1, c: -1, d: 0, e: page.heightPt, f: 0 };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function transformPoint(matrix: Matrix2DData, point: PointPt): PointPt {
  requireMatrix(matrix);
  requirePoint(point);
  return {
    xPt: matrix.a * point.xPt + matrix.c * point.yPt + matrix.e,
    yPt: matrix.b * point.xPt + matrix.d * point.yPt + matrix.f,
  };
}

export function transformRect(matrix: Matrix2DData, rect: LayoutRect): LayoutRect {
  requireRect(rect);
  const corners = [
    transformPoint(matrix, rect),
    transformPoint(matrix, { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt }),
    transformPoint(matrix, { xPt: rect.xPt, yPt: rect.yPt + rect.heightPt }),
    transformPoint(matrix, {
      xPt: rect.xPt + rect.widthPt,
      yPt: rect.yPt + rect.heightPt,
    }),
  ];
  const xs = corners.map(({ xPt }) => xPt);
  const ys = corners.map(({ yPt }) => yPt);
  const xPt = Math.min(...xs);
  const yPt = Math.min(...ys);
  return {
    xPt,
    yPt,
    widthPt: Math.max(...xs) - xPt,
    heightPt: Math.max(...ys) - yPt,
  };
}

/**
 * Reassign physical edge-owned facts to their transformed logical edges.
 *
 * The matrix supplies the only writing-mode authority: transforming each
 * physical outward unit normal identifies the logical edge that owns the same
 * fact. Values remain opaque, so the same operation projects numeric extents
 * and their provenance labels without duplicating direction-specific formulas.
 */
export function transformRectEdges<T>(
  matrix: Matrix2DData,
  edges: RectEdgeRecord<T>,
): RectEdgeRecord<T> {
  requireMatrix(matrix);
  const origin = transformPoint(matrix, { xPt: 0, yPt: 0 });
  const normals = {
    top: { xPt: 0, yPt: -1 },
    right: { xPt: 1, yPt: 0 },
    bottom: { xPt: 0, yPt: 1 },
    left: { xPt: -1, yPt: 0 },
  } as const;
  const projected: Partial<Record<keyof RectEdgeRecord<T>, T>> = {};
  const assigned = new Set<keyof RectEdgeRecord<T>>();
  for (const source of ['top', 'right', 'bottom', 'left'] as const) {
    const endpoint = transformPoint(matrix, normals[source]);
    const dx = endpoint.xPt - origin.xPt;
    const dy = endpoint.yPt - origin.yPt;
    const target = dy === 0 && dx !== 0
      ? dx > 0 ? 'right' : 'left'
      : dx === 0 && dy !== 0
        ? dy > 0 ? 'bottom' : 'top'
        : null;
    if (target === null || assigned.has(target)) {
      throw new RangeError('Edge transforms require a non-degenerate axis-aligned matrix');
    }
    projected[target] = edges[source];
    assigned.add(target);
  }
  if (assigned.size !== 4) {
    throw new RangeError('Edge transform must map every physical edge exactly once');
  }
  return projected as RectEdgeRecord<T>;
}

export function createSectionRegionCoordinateSpace(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): SectionRegionCoordinateSpace {
  return {
    writingMode,
    logicalToPhysical: logicalToPhysicalMatrix(writingMode, page),
    physicalToLogical: physicalToLogicalMatrix(writingMode, page),
  };
}
