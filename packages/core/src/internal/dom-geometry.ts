export interface RelativeElementRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** Intersect two rectangles in the same coordinate space. */
export function intersectElementRects(
  left: RelativeElementRect,
  right: RelativeElementRect,
): RelativeElementRect | undefined {
  const x = Math.max(left.x, right.x);
  const y = Math.max(left.y, right.y);
  const farX = Math.min(left.x + left.width, right.x + right.width);
  const farY = Math.min(left.y + left.height, right.y + right.height);
  if (farX <= x || farY <= y) return undefined;
  return Object.freeze({ x, y, width: farX - x, height: farY - y });
}

/** Measure an element in a Viewer-owned surface's CSS-pixel coordinate space. */
export function relativeElementRect(
  element: Element,
  surface: Element,
): RelativeElementRect | undefined {
  const rect = element.getBoundingClientRect();
  const origin = surface.getBoundingClientRect();
  if (![rect.left, rect.top, rect.width, rect.height, origin.left, origin.top].every(Number.isFinite)) {
    return undefined;
  }
  if (rect.width <= 0 || rect.height <= 0) return undefined;
  return Object.freeze({
    x: rect.left - origin.left,
    y: rect.top - origin.top,
    width: rect.width,
    height: rect.height,
  });
}
