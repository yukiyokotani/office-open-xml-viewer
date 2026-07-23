import type { ShapeElement, Slide } from './types';

/** A point in the slide's native EMU coordinate space. */
export interface PptxSlidePoint {
  x: number;
  y: number;
}

/** Options for {@link hitTestSlideShape}. */
export interface PptxShapeHitTestOptions {
  /**
   * Extra hit radius in EMU for line-like shapes. Viewers should convert their
   * desired CSS-pixel tolerance using `slideWidth / canvasCssWidth`.
   */
  tolerance?: number;
}

/** Detached result of a successful shape hit test. */
export interface PptxShapeHit {
  slideIndex: number;
  shapeId: string;
  shape: ShapeElement;
  point: PptxSlidePoint;
}

const LINE_GEOMETRIES = new Set(['line', 'straightconnector1']);

function inverseTransformPoint(
  shape: ShapeElement,
  point: PptxSlidePoint,
): PptxSlidePoint {
  const cx = shape.x + shape.width / 2;
  const cy = shape.y + shape.height / 2;
  const radians = (-shape.rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = point.x - cx;
  const dy = point.y - cy;
  let x = cx + cos * dx - sin * dy;
  let y = cy + sin * dx + cos * dy;

  // renderShape applies rotation and then flip around the same frame centre.
  // After undoing rotation, each flip is its own inverse.
  if (shape.flipH) x = 2 * cx - x;
  if (shape.flipV) y = 2 * cy - y;
  return { x, y };
}

function distanceToSegment(
  point: PptxSlidePoint,
  start: PptxSlidePoint,
  end: PptxSlidePoint,
): number {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const lengthSquared = dx * dx + dy * dy;
  if (lengthSquared === 0) return Math.hypot(point.x - start.x, point.y - start.y);
  const projection = Math.max(
    0,
    Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared),
  );
  return Math.hypot(
    point.x - (start.x + projection * dx),
    point.y - (start.y + projection * dy),
  );
}

/** Bounding-frame hit test for one shape, with line tolerance for connectors. */
export function hitTestShapeElement(
  shape: ShapeElement,
  point: PptxSlidePoint,
  tolerance = 0,
): boolean {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  const safeTolerance = Number.isFinite(tolerance) && tolerance > 0 ? tolerance : 0;
  const local = inverseTransformPoint(shape, point);
  const geometry = shape.geometry.toLowerCase();

  if (LINE_GEOMETRIES.has(geometry) || shape.width === 0 || shape.height === 0) {
    return distanceToSegment(
      local,
      { x: shape.x, y: shape.y },
      { x: shape.x + shape.width, y: shape.y + shape.height },
    ) <= safeTolerance;
  }

  const minX = Math.min(shape.x, shape.x + shape.width);
  const maxX = Math.max(shape.x, shape.x + shape.width);
  const minY = Math.min(shape.y, shape.y + shape.height);
  const maxY = Math.max(shape.y, shape.y + shape.height);
  return (
    local.x >= minX &&
    local.x <= maxX &&
    local.y >= minY &&
    local.y <= maxY
  );
}

/**
 * Return the topmost file-authored shape at a slide point.
 *
 * Slide elements are painted in array order, so the hit test walks them in
 * reverse. Parser-synthesized shapes without a stable `cNvPr@id` are skipped.
 */
export function hitTestSlideShape(
  slideIndex: number,
  slide: Slide,
  point: PptxSlidePoint,
  opts: PptxShapeHitTestOptions = {},
): PptxShapeHit | null {
  const tolerance = opts.tolerance ?? 0;
  for (let index = slide.elements.length - 1; index >= 0; index -= 1) {
    const element = slide.elements[index]!;
    if (element.type !== 'shape' || element.id === undefined) continue;
    if (!hitTestShapeElement(element, point, tolerance)) continue;
    return {
      slideIndex,
      shapeId: element.id,
      shape: structuredClone(element),
      point: { ...point },
    };
  }
  return null;
}
