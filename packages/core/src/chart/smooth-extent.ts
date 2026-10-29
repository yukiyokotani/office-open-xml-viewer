/** The value range a smoothed line actually draws.
 *
 * A smoothed series (`<c:smooth val="1"/>`) is painted as the Catmull-Rom Bézier
 * `appendCurve` builds, and that curve overshoots its points: a run that leaves a
 * flat stretch of zeros dips below zero. Excel's automatic axis covers the drawn
 * curve, not only the points, so an all-positive smoothed series can get a negative
 * axis minimum. The y control points depend only on neighbouring values, so the
 * extent is the same in value space as on the canvas.
 *
 * Runs break at blank points (the `gap` display of blanks).
 */
export function smoothCurveExtent(
  values: ReadonlyArray<number | null | undefined>,
): [number, number] | null {
  let lo = Infinity;
  let hi = -Infinity;
  let run: number[] = [];
  const flush = () => {
    for (const v of run) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    if (run.length >= 3) {
      for (let i = 0; i < run.length - 1; i++) {
        const p0 = run[i - 1] ?? run[i];
        const p1 = run[i];
        const p2 = run[i + 1];
        const p3 = run[i + 2] ?? p2;
        const c1 = p1 + (p2 - p0) / 6;
        const c2 = p2 - (p3 - p1) / 6;
        // ponytail: sampled rather than solved; 32 steps per segment is far below a pixel.
        for (let k = 1; k < 32; k++) {
          const t = k / 32;
          const u = 1 - t;
          const y = u * u * u * p1 + 3 * u * u * t * c1 + 3 * u * t * t * c2 + t * t * t * p2;
          lo = Math.min(lo, y);
          hi = Math.max(hi, y);
        }
      }
    }
    run = [];
  };
  for (const v of values) {
    if (v == null || !Number.isFinite(v)) flush();
    else run.push(v);
  }
  flush();
  return Number.isFinite(lo) ? [lo, hi] : null;
}
