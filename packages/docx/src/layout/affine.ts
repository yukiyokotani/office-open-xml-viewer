import type {
  LayoutRect,
  Matrix2DData,
  PointPt,
  WritingMode,
} from './types.js';

/** Canvas-order composition: the returned transform applies `inner`, then `outer`. */
export function composeAffine(outer: Matrix2DData, inner: Matrix2DData): Matrix2DData {
  return Object.freeze({
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
  });
}

export function scaleAffine(scale: number): Matrix2DData {
  return Object.freeze({ a: scale, b: 0, c: 0, d: scale, e: 0, f: 0 });
}

export function translationAffine(x: number, y: number): Matrix2DData {
  return Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: x, f: y });
}

export function quarterTurnAffine(direction: 1 | -1): Matrix2DData {
  return direction === 1
    ? Object.freeze({ a: 0, b: 1, c: -1, d: 0, e: 0, f: 0 })
    : Object.freeze({ a: 0, b: -1, c: 1, d: 0, e: 0, f: 0 });
}

export function canonicalLogicalToPhysical(
  writingMode: WritingMode,
  pageWidthPt: number,
): Matrix2DData {
  if (!Number.isFinite(pageWidthPt) || pageWidthPt < 0) {
    throw new RangeError('Physical page width must be a finite non-negative point value');
  }
  switch (writingMode) {
    case 'horizontal-tb':
      return Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
    case 'vertical-rl':
      // ECMA-376 §17.18.93 defines inline-down/block-right-to-left axes;
      // the page-width origin is this renderer's explicit logical convention.
      return Object.freeze({ a: 0, b: 1, c: -1, d: 0, e: pageWidthPt, f: 0 });
    case 'vertical-lr':
      return Object.freeze({ a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 });
  }
}

export function mapAffinePoint(matrix: Matrix2DData, point: PointPt): PointPt {
  return {
    xPt: matrix.a * point.xPt + matrix.c * point.yPt + matrix.e,
    yPt: matrix.b * point.xPt + matrix.d * point.yPt + matrix.f,
  };
}

export function mapAffineRect(matrix: Matrix2DData, rect: LayoutRect): LayoutRect {
  const points = [
    mapAffinePoint(matrix, { xPt: rect.xPt, yPt: rect.yPt }),
    mapAffinePoint(matrix, { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt }),
    mapAffinePoint(matrix, { xPt: rect.xPt, yPt: rect.yPt + rect.heightPt }),
    mapAffinePoint(matrix, {
      xPt: rect.xPt + rect.widthPt,
      yPt: rect.yPt + rect.heightPt,
    }),
  ];
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  const xPt = Math.min(...xs);
  const yPt = Math.min(...ys);
  return {
    xPt,
    yPt,
    widthPt: Math.max(...xs) - xPt,
    heightPt: Math.max(...ys) - yPt,
  };
}

export function inverseAffine(matrix: Matrix2DData): Matrix2DData | null {
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(determinant) || determinant === 0) return null;
  return Object.freeze({
    a: matrix.d / determinant,
    b: -matrix.b / determinant,
    c: -matrix.c / determinant,
    d: matrix.a / determinant,
    e: (matrix.c * matrix.f - matrix.d * matrix.e) / determinant,
    f: (matrix.b * matrix.e - matrix.a * matrix.f) / determinant,
  });
}

export function inverseMapAffinePoint(matrix: Matrix2DData, point: PointPt): PointPt | null {
  const inverse = inverseAffine(matrix);
  return inverse ? mapAffinePoint(inverse, point) : null;
}

export function inverseMapAffineVector(matrix: Matrix2DData, vector: PointPt): PointPt | null {
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(determinant) || determinant === 0) return null;
  const result = {
    xPt: (matrix.d * vector.xPt - matrix.c * vector.yPt) / determinant,
    yPt: (-matrix.b * vector.xPt + matrix.a * vector.yPt) / determinant,
  };
  return Number.isFinite(result.xPt) && Number.isFinite(result.yPt) ? result : null;
}

export function sameAffine(left: Matrix2DData, right: Matrix2DData): boolean {
  return left.a === right.a && left.b === right.b
    && left.c === right.c && left.d === right.d
    && left.e === right.e && left.f === right.f;
}
