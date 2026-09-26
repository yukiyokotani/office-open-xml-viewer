export interface ImageTransformRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

function sinCos(rotationDeg: number): readonly [number, number] {
  if (!Number.isFinite(rotationDeg)) return [0, 1];
  const normalized = ((rotationDeg % 360) + 360) % 360;
  if (normalized === 0) return [0, 1];
  if (normalized === 90) return [1, 0];
  if (normalized === 180) return [0, -1];
  if (normalized === 270) return [-1, 0];
  const radians = rotationDeg * Math.PI / 180;
  return [Math.sin(radians), Math.cos(radians)];
}

/** Axis-aligned bounds of a rectangle rotated about its centre. Reflections do
 * not change these bounds. DrawingML rotation is clockwise, but bounds use
 * absolute sine/cosine and are direction-independent. */
export function rotatedImageBounds(
  rect: ImageTransformRect,
  rotationDeg = 0,
): ImageTransformRect {
  const [sin, cos] = sinCos(rotationDeg);
  if (sin === 0 && Math.abs(cos) === 1) return rect;
  const width = Math.abs(rect.width * cos) + Math.abs(rect.height * sin);
  const height = Math.abs(rect.width * sin) + Math.abs(rect.height * cos);
  return {
    x: rect.x + (rect.width - width) / 2,
    y: rect.y + (rect.height - height) / 2,
    width,
    height,
  };
}

/** Map a canvas point back through the authored centre rotation/reflections. */
export function inverseImageTransformPoint(
  point: Readonly<{ x: number; y: number }>,
  rect: ImageTransformRect,
  rotationDeg = 0,
  flipH = false,
  flipV = false,
): Readonly<{ x: number; y: number }> {
  if (rotationDeg === 0 && !flipH && !flipV) return point;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const [sin, cos] = sinCos(-rotationDeg);
  const dx = point.x - cx;
  const dy = point.y - cy;
  let x = cx + dx * cos - dy * sin;
  let y = cy + dx * sin + dy * cos;
  if (flipH) x = 2 * cx - x;
  if (flipV) y = 2 * cy - y;
  return { x, y };
}
