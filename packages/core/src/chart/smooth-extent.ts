// Value extent of smoothed classic lines (`<c:smooth val="1"/>`, §21.2.2.194).
import type { ChartModel, ChartSeries } from '../types/chart';

import { classicDataPointLineStyle } from './classic-data-point-style.js';

/** Category-index runs a category-axis line series is drawn as. Shared by the
 * line painters and the automatic-axis extent so both see the same curve.
 * A hidden source cell always breaks the run. A blank cell follows
 * `<c:dispBlanksAs>` (§21.2.2.42): `gap` breaks the run, `span` skips the point
 * and keeps the run going, `zero` plots it. `blanksArePlotted` is set for
 * stacked series, whose blank cells already contribute 0 to the stack. */
export function classicLineRunIndices(
  series: ChartSeries,
  n: number,
  dispBlanksAs: string | null | undefined,
  blanksArePlotted = false,
): number[][] {
  const dispBlanks = dispBlanksAs ?? 'gap';
  const runs: number[][] = [];
  let run: number[] = [];
  const flush = (): void => {
    if (run.length > 0) runs.push(run);
    run = [];
  };
  for (let ci = 0; ci < n; ci++) {
    if (series.sourceHidden?.[ci] === true) {
      flush();
      continue;
    }
    if (!blanksArePlotted && series.values[ci] == null) {
      if (dispBlanks === 'gap') { flush(); continue; }
      if (dispBlanks === 'span') continue;
    }
    run.push(ci);
  }
  flush();
  return runs;
}

/** Whether a smoothed series' curve counts toward its axes. Excel 365 decides
 * per series: removing the series line (`a:ln/a:noFill`, markers only) keeps
 * the axis at the points for line and scatter, while hiding only some point
 * segments (`c:dPt` line `noFill`) still widens it to the whole curve. */
export function seriesPaintsSmoothedLine(
  chart: ChartModel,
  series: ChartSeries,
  styleIndex: number,
  smooth = series.smooth === true,
): boolean {
  return smooth
    && classicDataPointLineStyle(chart, 'dataPointLine', series, undefined, styleIndex).paint !== null;
}

/** The value range of the smoothed curves drawn through `runs`, each an
 * ordered list of one coordinate (y for category charts; x or y for scatter).
 *
 * A smoothed line is painted as the Catmull-Rom Bézier `appendCurve` builds
 * (control points p1 + (p2 - p0) / 6 and p2 - (p3 - p1) / 6, ends clamped),
 * and that curve overshoots its points: a run leaving a flat stretch of zeros
 * dips below zero. Excel's automatic value axis covers the drawn curve, not
 * only the points. Observed in Excel 365 (Mac) on line, stacked line, a line
 * over columns, a secondary-axis line, and smooth scatter, for both gap and
 * span blanks (an all-positive series gets a -1 minimum; a peak above the
 * nice maximum raises it). A dip that stays above zero keeps the automatic
 * zero baseline, which the caller's axis planner already applies.
 *
 * Each coordinate of the Bézier depends only on that coordinate of its
 * control points, so the extent is computed per axis. Control points are
 * placed in pixel space, which is linear in the value for a linear axis and in
 * log(value) for a logarithmic one (`logBase`): the curve is evaluated in that
 * space. A log-axis run holding a non-positive value has no finite pixel
 * position and contributes no curve (the painter clamps such a point to the
 * smallest positive number, whose control points are not a meaningful axis
 * extent). Runs of fewer than three points are
 * drawn straight and add nothing beyond their points, which the caller's
 * point extent already covers. Exact cubic extrema are used. */
export function smoothCurveExtent(
  runs: Iterable<ReadonlyArray<number>>,
  logBase?: number | null,
): [number, number] | null {
  const log = logBase != null && Number.isFinite(logBase) && logBase >= 2;
  const toSpace = (value: number): number => log ? Math.log(value) / Math.log(logBase) : value;
  const fromSpace = (value: number): number => log ? logBase ** value : value;
  let lo = Infinity;
  let hi = -Infinity;
  const include = (value: number): void => {
    if (value < lo) lo = value;
    if (value > hi) hi = value;
  };
  for (const run of runs) {
    if (run.length < 3) continue;
    if (run.some(value => !Number.isFinite(value) || (log && value <= 0))) continue;
    const p = run.map(toSpace);
    for (let i = 0; i < p.length - 1; i++) {
      const p0 = p[i - 1] ?? p[i];
      const p1 = p[i];
      const p2 = p[i + 1];
      const p3 = p[i + 2] ?? p2;
      const c1 = p1 + (p2 - p0) / 6;
      const c2 = p2 - (p3 - p1) / 6;
      include(p1);
      include(p2);
      // B'(t) / 3 = a t² + b t + c for the cubic p1, c1, c2, p2. Degeneracy
      // is judged relative to the coefficients so tiny magnitudes keep their
      // extrema.
      const a = -p1 + 3 * c1 - 3 * c2 + p2;
      const b = 2 * (p1 - 2 * c1 + c2);
      const c = c1 - p1;
      const scale = Math.max(Math.abs(a), Math.abs(b), Math.abs(c));
      if (scale === 0) continue;
      const roots: number[] = [];
      if (Math.abs(a) <= scale * 1e-12) {
        if (Math.abs(b) > scale * 1e-12) roots.push(-c / b);
      } else {
        const disc = b * b - 4 * a * c;
        if (disc >= 0) {
          const sq = Math.sqrt(disc);
          roots.push((-b + sq) / (2 * a), (-b - sq) / (2 * a));
        }
      }
      for (const t of roots) {
        if (!(t > 0 && t < 1)) continue;
        const u = 1 - t;
        include(u * u * u * p1 + 3 * u * u * t * c1 + 3 * u * t * t * c2 + t * t * t * p2);
      }
    }
  }
  return Number.isFinite(lo) ? [fromSpace(lo), fromSpace(hi)] : null;
}
