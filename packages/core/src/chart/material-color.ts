import type { ChartThreeD } from '../types/chart';

/** Effective omitted-view camera isolated by the S1-S5 Office boundary set. */
export function isObservedAutomaticSurfaceCamera(view: ChartThreeD): boolean {
  return view.rotationX === 15
    && view.rotationY === 20
    && view.rightAngleAxes === false
    && view.perspective === 30;
}

/**
 * Office's omitted Surface view uses a wider perspective response than the
 * DrawingML angle alone produces in the shared camera. That compatibility
 * gain is observed only for the effective S1-S5 camera; authored cameras keep
 * the specification-derived projection until their own boundaries exist.
 */
export function surfacePerspectiveTangentGain(view: ChartThreeD): number {
  return isObservedAutomaticSurfaceCamera(view) ? 2 : 1;
}

/** Multiply an sRGB chart material by a bounded diffuse-light factor. */
export function scaleHexColor(color: string, factor: number): string {
  const value = color.replace(/^#/, '');
  if (!/^[0-9a-f]{6}$/i.test(value) || !Number.isFinite(factor)) return color;
  const bounded = Math.max(0, factor);
  const channel = (offset: number) => Math.max(0, Math.min(255,
    Math.round(Number.parseInt(value.slice(offset, offset + 2), 16) * bounded),
  )).toString(16).padStart(2, '0');
  return `#${channel(0)}${channel(2)}${channel(4)}`.toUpperCase();
}

/** Surface-only automatic material response.
 *
 * ECMA-376 carries the view and source mesh but leaves the automatic material
 * to the application. The S1-S5/saddle/90° contour boundary set isolates a
 * single camera-space directional response: opposing source-grid triangles
 * may darken or brighten, while authored band paint remains the base colour.
 * Keep that compatibility rule here (and out of general fills/3-D solids).
 */
export function surfaceMaterialFactor(normal: { x: number; y: number; z: number } | null): number {
  if (!normal) return 1;
  const facing = normal.z < 0
    ? { x: -normal.x, y: -normal.y, z: -normal.z }
    : normal;
  // The observed Office surface material is lit from screen upper-right.
  // Shared camera projection maps +x rightward and +y upward on screen.
  const light = { x: 0.24, y: 0.42, z: 0.88 };
  const length = Math.hypot(light.x, light.y, light.z);
  const lambert = Math.max(0,
    (facing.x * light.x + facing.y * light.y + facing.z * light.z) / length,
  );
  return Math.max(0.48, Math.min(1.22, 0.48 + 0.78 * lambert));
}
