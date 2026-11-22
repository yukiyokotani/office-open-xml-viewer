// Excel-style "nice" value-axis scaling. Pure math (no canvas), extracted so it
// can be unit-tested and reused independently of the chart renderer.

export interface FiniteDataExtent {
  min: number;
  max: number;
}

/**
 * Bounded-stack finite min/max reduction for chart caches and public models.
 * Avoids `Math.min(...values)` / `Math.max(...values)`, whose argument expansion
 * can throw for otherwise valid large OOXML caches.
 */
export function finiteDataExtent(
  values: Iterable<number | null | undefined>,
  fallback: FiniteDataExtent = { min: 0, max: 1 },
): FiniteDataExtent {
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (const value of values) {
    if (value == null || !Number.isFinite(value)) continue;
    if (value < min) min = value;
    if (value > max) max = value;
  }
  return Number.isFinite(min) && Number.isFinite(max) ? { min, max } : fallback;
}

/** A round major-unit step that yields roughly `targetSteps` gridlines across
 *  `range` (1 / 2 / 5 × 10ⁿ — Excel's default ladder). */
export function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
}

/**
 * Automatic percent-stacked major unit in percentage-point data space.
 * OOXML does not define the omitted unit.  The 48-case boundary corpus found
 * two stable density classes:
 *
 * - horizontal value axes: five positive intervals, or four signed intervals;
 * - vertical value axes: the same compact density below a 120pt plot axis,
 *   otherwise ten intervals.
 *
 * The result stays on the shared 1/2/5 ceiling ladder. Authored majorUnit is
 * resolved by the caller and always wins over this compatibility default.
 */
export function automaticPercentMajorUnit(
  min: number,
  max: number,
  orientation: 'vertical' | 'horizontal',
  axisLenPt?: number,
): number {
  const lo = Number.isFinite(min) ? min : 0;
  const hi = Number.isFinite(max) ? max : 100;
  const span = hi > lo ? hi - lo : 100;
  const signed = lo < 0 && hi > 0;
  const compact = orientation === 'horizontal'
    || axisLenPt == null
    || !Number.isFinite(axisLenPt)
    || axisLenPt < 120;
  const intervals = compact ? (signed ? 4 : 5) : 10;
  return ceilingNiceStep(hi > lo ? spanQuotient(lo, hi, intervals) : span / intervals);
}

/**
 * Automatic radar-ring unit in value space. The 36-case Office corpus shows
 * three radial densities: about 4 intervals on a small spoke, 8 on an ordinary
 * spoke and 10 on a large spoke. Keeping this in the shared axis module avoids
 * the former renderer-local fixed five-ring path while preserving authored
 * majorUnit and logarithmic axes at the call site.
 */
export function automaticRadarMajorUnit(
  min: number,
  max: number,
  spokeLenPt?: number,
): number {
  const lo = Number.isFinite(min) ? min : 0;
  const hi = Number.isFinite(max) ? max : 1;
  const span = hi > lo ? null : 1;
  const intervals = spokeLenPt != null && Number.isFinite(spokeLenPt)
    ? spokeLenPt < 45 ? 4 : spokeLenPt < 90 ? 8 : 10
    : 8;
  return ceilingNiceStep(hi > lo ? spanQuotient(lo, hi, intervals) : (span ?? 1) / intervals);
}

/** Automatic major unit for classic Surface value bands.
 *
 * OOXML leaves an omitted major unit application-defined. The Surface
 * boundary corpus (ordinary/tall, compact, scaled and 90-degree contour
 * views) isolates the Office choice to the chart-frame projection of the
 * value direction: one interval per roughly 28pt, with a five-interval
 * compact/edge-on floor. This is intentionally measured before title/legend
 * reserves; using the final plot rect coarsens the verified ordinary S1/S4
 * boundaries. The result remains on the shared ceiling 1/2/5 ladder and an
 * authored majorUnit still wins at the caller.
 */
export function automaticSurfaceMajorUnit(
  min: number,
  max: number,
  projectedValueAxisLenPt?: number,
): number {
  const lo = Number.isFinite(min) ? min : 0;
  const hi = Number.isFinite(max) ? max : 1;
  const span = hi > lo ? hi - lo : 1;
  const intervals = projectedValueAxisLenPt != null
    && Number.isFinite(projectedValueAxisLenPt)
    && projectedValueAxisLenPt > 0
    ? Math.max(5, Math.round(projectedValueAxisLenPt / 28))
    : 5;
  return ceilingNiceStep(span / intervals);
}

/** Hard allocation/paint bound for each numeric-axis tick layer. */
export const MAX_AXIS_TICKS = 512;

export interface LinearValueAxisOptions {
  dataMin: number;
  dataMax: number;
  explicitMin?: number | null;
  explicitMax?: number | null;
  axisLenPt?: number;
  /** Physical direction of the numeric axis. Omitted keeps the legacy policy. */
  axisOrientation?: 'vertical' | 'horizontal';
  majorUnit?: number | null;
  minorUnit?: number | null;
  /** Minor positions are needed by either minor ticks or minor gridlines. */
  needMinor?: boolean;
}

export interface LinearValueAxisPlan {
  min: number;
  max: number;
  majorUnit: number;
  minorUnit: number | null;
  majorTicks: number[];
  minorTicks: number[];
}

/** Smallest 1/2/5 × 10ⁿ step at least as large as `raw`. */
export function ceilingNiceStep(raw: number): number {
  if (!(raw > 0) || !isFinite(raw)) return 1;
  const exponent = Math.floor(Math.log10(raw));
  const magnitude = Math.pow(10, exponent);
  if (!(magnitude > 0) || !isFinite(magnitude)) return raw;
  const normalized = raw / magnitude;
  const factor = normalized <= 1
    ? 1
    : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return Math.min(Number.MAX_VALUE, factor * magnitude);
}

function finiteBound(value: number | null | undefined): number | null {
  return value != null && isFinite(value) ? value : null;
}

function positiveUnit(value: number | null | undefined): number | null {
  return value != null && isFinite(value) && value > 0 ? value : null;
}

/** Compute `(max - min) / divisor` without first materializing an overflowing
 * span. Dividing the endpoints before subtraction keeps opposite-sign finite
 * bounds usable all the way through Number.MAX_VALUE. */
function spanQuotient(min: number, max: number, divisor: number): number {
  const directSpan = max - min;
  const quotient = isFinite(directSpan)
    ? directSpan / divisor
    : max / divisor - min / divisor;
  if (quotient > 0 && isFinite(quotient)) return quotient;
  // A positive subnormal span may underflow when divided. The smallest
  // representable positive step is the only meaningful density input; using a
  // huge fallback would invert the scale (unit much larger than its range).
  if (directSpan > 0 && isFinite(directSpan)) return Number.MIN_VALUE;
  return Number.MAX_VALUE;
}

function tickCount(min: number, max: number, unit: number): number {
  if (!isFinite(min) || !isFinite(max) || !(max >= min) || !(unit > 0) || !isFinite(unit)) return 0;
  const span = max - min;
  const intervals = isFinite(span) ? span / unit : max / unit - min / unit;
  if (!isFinite(intervals) || intervals > Number.MAX_SAFE_INTEGER) return Number.POSITIVE_INFINITY;
  return Math.max(1, Math.floor(intervals + 1e-9) + 1);
}

/**
 * Generate axis positions from integer indices. The count is checked before
 * allocation; explicit hostile units are either truncated (major) or skipped
 * (minor) instead of creating an unbounded render loop.
 */
function indexedTicks(
  min: number,
  max: number,
  unit: number,
  overflow: 'truncate' | 'skip',
  excludeMajorUnit?: number,
): number[] {
  const expected = tickCount(min, max, unit);
  if (expected === 0 || (expected > MAX_AXIS_TICKS && overflow === 'skip')) return [];
  const count = Math.min(expected, MAX_AXIS_TICKS);
  const values: number[] = [];
  const span = max - min;
  const epsilon = Math.max(
    Math.abs(unit),
    isFinite(span) ? Math.abs(span) : Math.max(Math.abs(min), Math.abs(max)),
  ) * 1e-9;
  let previousValue = Number.NEGATIVE_INFINITY;
  for (let index = 0; index < count; index++) {
    const offset = index * unit;
    let value = min + offset;
    // On an opposite-sign near-limit axis, `index * unit` can overflow even
    // though the final translated tick is finite. Retry only that exceptional
    // path in unit coordinates; ordinary values retain their exact old math.
    if (!isFinite(offset) || !isFinite(value) || (index > 0 && !(value > previousValue))) {
      value = (min / unit + index) * unit;
    }
    if (!isFinite(value) || value > max + epsilon) break;
    if (index > 0 && !(value > previousValue)) break;
    previousValue = value;
    if (excludeMajorUnit != null) {
      if (value >= max - epsilon) break;
      const delta = value - min;
      const multiple = isFinite(delta)
        ? delta / excludeMajorUnit
        : value / excludeMajorUnit - min / excludeMajorUnit;
      if (isFinite(multiple) && Math.abs(multiple - Math.round(multiple)) <= 1e-8) continue;
    }
    values.push(value);
  }
  return values;
}

/**
 * Shared, bounded compatibility planner for an automatic linear value axis.
 * OOXML defines authored bounds/units but deliberately leaves omitted values
 * to the consuming application; this compact policy supplies those defaults.
 */
export function planLinearValueAxis(options: LinearValueAxisOptions): LinearValueAxisPlan {
  let dataMin = isFinite(options.dataMin) ? options.dataMin : 0;
  let dataMax = isFinite(options.dataMax) ? options.dataMax : 1;
  if (dataMin > dataMax) [dataMin, dataMax] = [dataMax, dataMin];

  let autoMin: number;
  let autoMax: number;
  let autoMajor: number;
  if (dataMin === 0 && dataMax === 0) {
    autoMin = 0;
    autoMax = 1;
    autoMajor = 0.1;
  } else {
    // A non-zero point range expands toward zero before applying the ordinary
    // symmetric padding policy.
    if (dataMin === dataMax) {
      dataMin = Math.min(0, dataMin);
      dataMax = Math.max(0, dataMax);
    }
    const anchorPositiveAtZero = dataMin >= 0
      && (dataMin === 0 || dataMax > 1.2 * dataMin);
    const anchorNegativeAtZero = dataMax <= 0
      && (dataMax === 0 || Math.abs(dataMin) > 1.2 * Math.abs(dataMax));
    const paddedMin = anchorPositiveAtZero ? 0 : dataMin;
    const paddedMax = anchorNegativeAtZero ? 0 : dataMax;
    const span = paddedMax - paddedMin;
    const padding = isFinite(span)
      ? span * 0.05
      : paddedMax * 0.05 - paddedMin * 0.05;
    let lo = anchorPositiveAtZero ? 0 : dataMin - padding;
    let hi = anchorNegativeAtZero ? 0 : dataMax + padding;
    // Observed Office rule (not in ECMA-376 §21.2.2.107/108/160): the strict
    // 1.2 zero-pin boundary, the 1/2/5 ladder and the 5 % padding come from
    // an Office corpus of 6,354 fully automatic finite linear value axes
    // (#1236 / #1246).
    // The 1.2 boundary is intentionally strict: equality remains offset.
    // Once an endpoint is zero-anchored, padding is computed from that same
    // effective span. This keeps identical automatic axes family-invariant
    // and sign-mirrored instead of retaining the pre-anchor offset range.
    // The broad automatic-axis corpus (6,354 finite line-axis cases) selected
    // the 1/2/5 ladder from roughly ten intervals.  Using the plot length here
    // made ordinary vertical charts collapse to only three or four labelled
    // intervals (for example 0..600 by 200 instead of 0..600 by 100).  Physical
    // size remains relevant only for the separately observed explicit-bounds
    // rule below; the fully automatic compact policy is intentionally one
    // shared ten-interval rule across chart families.
    autoMajor = ceilingNiceStep(hi / 10 - lo / 10);
    autoMin = Math.floor(lo / autoMajor) * autoMajor;
    autoMax = Math.ceil(hi / autoMajor) * autoMajor;
    // Large offsets can erase sub-ULP padding/rounding. Keep a finite range
    // containing the data without inventing special offset heuristics.
    if (!isFinite(autoMin) || !isFinite(autoMax) || !(autoMax > autoMin)) {
      autoMin = Math.min(dataMin, 0);
      autoMax = Math.max(dataMax, autoMin + autoMajor);
    }
  }

  const explicitMin = finiteBound(options.explicitMin);
  const explicitMax = finiteBound(options.explicitMax);
  const authoredMajor = positiveUnit(options.majorUnit);
  // When only a major unit is authored, Excel keeps automatic bounds aligned
  // to that unit.  The previous coarser automatic unit happened to mask this
  // rule for common values; make it explicit now that auto density is finer.
  const min = explicitMin ?? (authoredMajor == null
    ? autoMin
    : Math.floor(autoMin / authoredMajor) * authoredMajor);
  const max = explicitMax ?? (authoredMajor == null
    ? autoMax
    : Math.ceil(autoMax / authoredMajor) * authoredMajor);
  let automaticMajor = autoMajor;
  // OOXML defines authored min/max and majorUnit, but not the unit chosen when
  // both bounds are present and majorUnit is omitted. A bounded corpus across
  // column, stacked-column, line, area, scatter-Y, and combo axes found that
  // axes choose from the authored span, not the interior data range. The
  // vertical fit below matched all 231 measured vertical rows. Horizontal axes
  // form a second density class: a ceiling-ladder step at roughly one interval
  // per 38pt matched all 66 measured bar/scatter-X rows, including the 1600
  // boundary and the compact-chart controls. Keeping both observations here
  // avoids family-local axis paths.
  if (
    authoredMajor == null
    && explicitMin != null
    && explicitMax != null
    && explicitMax > explicitMin
  ) {
    if (options.axisOrientation === 'horizontal') {
      const axisTarget = options.axisLenPt != null
        && isFinite(options.axisLenPt)
        && options.axisLenPt > 0
        ? Math.max(5, Math.round(options.axisLenPt / 38))
        : 8;
      automaticMajor = ceilingNiceStep(
        spanQuotient(explicitMin, explicitMax, axisTarget),
      );
    } else if (options.axisOrientation === 'vertical') {
      const axisTarget = options.axisLenPt != null
        && isFinite(options.axisLenPt)
        && options.axisLenPt > 0
        ? Math.max(5, Math.round(options.axisLenPt / 28))
        : 7;
      automaticMajor = Math.max(
        ceilingNiceStep(spanQuotient(explicitMin, explicitMax, 10)),
        Math.min(
          Number.MAX_VALUE,
          niceStep(spanQuotient(explicitMin, explicitMax, axisTarget), 1),
        ),
      );
    }
  }
  let majorUnit = authoredMajor ?? automaticMajor;
  // An automatic unit may be coarsened to keep explicit wide bounds safe.
  if (authoredMajor == null && tickCount(min, max, majorUnit) > MAX_AXIS_TICKS) {
    majorUnit = ceilingNiceStep(max / (MAX_AXIS_TICKS - 1) - min / (MAX_AXIS_TICKS - 1));
  }
  const authoredMinor = positiveUnit(options.minorUnit);
  // A fully automatic minor plan must cover the entire axis. Since each major
  // interval contains five minor intervals, coarsen the automatic major unit
  // up front instead of returning only the first MAX_AXIS_TICKS positions.
  if (options.needMinor && authoredMajor == null && authoredMinor == null) {
    const maxMajorIntervals = Math.floor((MAX_AXIS_TICKS - 1) / 5);
    if (tickCount(min, max, majorUnit / 5) > MAX_AXIS_TICKS) {
      majorUnit = ceilingNiceStep(max / maxMajorIntervals - min / maxMajorIntervals);
    }
  }
  const automaticMinor = majorUnit / 5;
  const minorUnit = options.needMinor
    ? (authoredMinor ?? (automaticMinor > 0 && isFinite(automaticMinor) ? automaticMinor : majorUnit))
    : null;

  const majorTicks = indexedTicks(min, max, majorUnit, authoredMajor == null ? 'skip' : 'truncate');
  const minorTicks = minorUnit == null
    ? []
    : indexedTicks(min, max, minorUnit, 'skip', majorUnit);
  return { min, max, majorUnit, minorUnit, majorTicks, minorTicks };
}

export function valueAxisScale(
  dataMin: number, dataMax: number,
  explicitMin?: number | null, explicitMax?: number | null,
  axisLenPt?: number,
  majorUnit?: number | null,
): { min: number; max: number; step: number } {
  const plan = planLinearValueAxis({
    dataMin, dataMax, explicitMin, explicitMax, axisLenPt, majorUnit,
  });
  return { min: plan.min, max: plan.max, step: plan.majorUnit };
}

/** Options for {@link axisFraction}: a logarithmic base and/or a reversed
 *  (`maxMin`) orientation. Both default off, in which case the fraction is the
 *  plain linear `(v - min) / (max - min)` — byte-identical to the renderers'
 *  historical inline math. */
export interface AxisFractionOpts {
  /** `<c:scaling><c:logBase val>` (ECMA-376 §21.2.2.98). When set (>= 2) the
   *  value maps in log space. min/max must be positive. */
  logBase?: number | null;
  /** `<c:scaling><c:orientation val="maxMin">` — reverse the axis. */
  reversed?: boolean;
}

/** Map a value to its 0..1 position along an axis spanning `min`..`max`.
 *
 *  This is the single shared primitive the per-chart-type renderers build their
 *  `toY` / `valX` closures on: a caller does `py0 + ph - axisFraction(v, min,
 *  max) * ph`. With no options it returns exactly `(v - min) / (max - min)`, so
 *  a linear, normally-oriented axis is byte-stable. A log base maps in log space
 *  (gridlines fall on powers of the base); `reversed` flips the fraction for a
 *  `maxMin` orientation. A degenerate zero range yields 0 (no NaN/∞). */
export function axisFraction(
  v: number, min: number, max: number, opts?: AxisFractionOpts,
): number {
  let frac: number;
  const logBase = opts?.logBase;
  if (logBase != null && isFinite(logBase) && logBase >= 2 && min > 0 && max > 0) {
    const lo = Math.log(min);
    const hi = Math.log(max);
    const denom = hi - lo;
    frac = denom === 0 ? 0 : (Math.log(Math.max(v, Number.MIN_VALUE)) - lo) / denom;
  } else {
    const denom = max - min;
    if (denom === 0) {
      frac = 0;
    } else if (Number.isFinite(denom) && Number.isFinite(v - min)) {
      // Keep the ordinary path byte-stable for normal chart ranges.
      frac = (v - min) / denom;
    } else {
      // Finite opposite-sign bounds can overflow both subtractions even
      // though their mathematical ratio is representable. Normalize before
      // subtracting so no renderer receives Infinity/Infinity → NaN.
      const scale = Math.max(Math.abs(v), Math.abs(min), Math.abs(max));
      const scaledDenom = max / scale - min / scale;
      frac = scaledDenom === 0 ? 0 : (v / scale - min / scale) / scaledDenom;
    }
  }
  return opts?.reversed ? 1 - frac : frac;
}

/** Logarithmic value-axis bounds + gridline decades. Snaps `dataMin` down and
 *  `dataMax` up to whole powers of `base` (Excel's log-axis behavior) and lists
 *  every power-of-base gridline in `[min, max]`. Explicit `<c:scaling><c:min /
 *  max>` override the snapped bounds. A non-positive `dataMin` (log undefined at
 *  <= 0) floors to the base's lowest decade at or below the smallest positive
 *  datum, defaulting to `base^0 = 1` when there is none. */
export function logAxisScale(
  dataMin: number, dataMax: number, base: number,
  explicitMin?: number | null, explicitMax?: number | null,
): { min: number; max: number; lines: number[] } {
  const b = isFinite(base) && base >= 2 ? base : 10;
  const logB = (x: number): number => Math.log(x) / Math.log(b);
  const posMax = dataMax > 0 ? dataMax : 1;
  // Floor the min to a decade; clamp a non-positive min up to the smallest
  // decade that is still <= the positive max.
  const rawMin = dataMin > 0 ? dataMin : posMax;
  const minExp = Math.floor(logB(rawMin));
  const maxExp = Math.ceil(logB(posMax));
  const finitePow = (exponent: number): number => {
    const value = Math.pow(b, exponent);
    if (value === 0) return Number.MIN_VALUE;
    return isFinite(value) ? value : Number.MAX_VALUE;
  };
  const min = explicitMin != null ? explicitMin : finitePow(minExp);
  const max = explicitMax != null ? explicitMax : finitePow(Math.max(maxExp, minExp + 1));
  const lines: number[] = [];
  const startExp = Math.ceil(logB(min) - 1e-9);
  const endExp = Math.floor(logB(max) + 1e-9);
  if (!isFinite(startExp) || !isFinite(endExp) || endExp < startExp) return { min, max, lines };
  const expected = endExp - startExp + 1;
  const stride = Math.max(1, Math.ceil((expected - 1) / (MAX_AXIS_TICKS - 1)));
  const count = Math.min(MAX_AXIS_TICKS, Math.floor((expected - 1) / stride) + 1);
  for (let index = 0; index < count; index++) {
    const value = finitePow(startExp + index * stride);
    if (value >= min && value <= max && (lines.length === 0 || value > lines[lines.length - 1])) {
      lines.push(value);
    }
  }
  const endValue = finitePow(endExp);
  if (lines.length < MAX_AXIS_TICKS && endValue >= min && endValue <= max
    && endValue > (lines[lines.length - 1] ?? 0)) lines.push(endValue);
  return { min, max, lines };
}

export interface NumericValueAxisOptions extends LinearValueAxisOptions {
  /** `<c:scaling><c:logBase>`; omission keeps the linear planner. */
  logBase?: number | null;
  /** `<c:scaling><c:orientation val="maxMin">`. */
  reversed?: boolean;
}

export interface NumericValueAxisPlan extends LinearValueAxisPlan {
  /** Shared value→axis fraction for linear/logarithmic and normal/reversed axes. */
  fraction: (value: number) => number;
}

/** Shared numeric-axis entry point used by primary, secondary and scatter X/Y
 * axes. It keeps the bounded linear planner and logarithmic decade planner
 * behind one model contract so a chart family cannot silently drop log/reverse. */
export function planNumericValueAxis(options: NumericValueAxisOptions): NumericValueAxisPlan {
  const logBase = options.logBase;
  if (logBase != null && Number.isFinite(logBase) && logBase >= 2) {
    const { min, max, lines } = logAxisScale(
      options.dataMin,
      options.dataMax,
      logBase,
      finiteBound(options.explicitMin),
      finiteBound(options.explicitMax),
    );
    return {
      min,
      max,
      majorUnit: lines.length > 1 ? lines[1] - lines[0] : max - min,
      minorUnit: null,
      majorTicks: lines,
      minorTicks: [],
      fraction: value => axisFraction(value, min, max, {
        logBase,
        reversed: options.reversed,
      }),
    };
  }
  const linear = planLinearValueAxis(options);
  return {
    ...linear,
    fraction: value => axisFraction(value, linear.min, linear.max, {
      reversed: options.reversed,
    }),
  };
}

/** A fitted trendline as a list of `(x, y)` points in DATA space (x = the
 *  point's category index / x-value, y = the fitted value). The renderer maps
 *  these through the chart's category→pixel and value→pixel transforms. Empty
 *  when the type is unsupported or there is too little data to fit. */
export interface TrendlinePoints { xs: number[]; ys: number[] }

/** Least-squares coefficients and coefficient of determination for a linear
 * trendline. This is kept separate from the sampled line points because
 * `<c:dispEq>` / `<c:dispRSqr>` expose the fitted values as chart text. */
export interface LinearTrendlineStats {
  slope: number;
  intercept: number;
  rSquared: number;
}

export function linearTrendlineStats(
  xs: readonly number[],
  ys: readonly number[],
  forcedIntercept?: number | null,
): LinearTrendlineStats | null {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return null;
  let sx = 0, sy = 0, sxx = 0, sxy = 0;
  for (let i = 0; i < n; i++) {
    sx += xs[i]; sy += ys[i]; sxx += xs[i] * xs[i]; sxy += xs[i] * ys[i];
  }
  let slope: number;
  let intercept: number;
  if (forcedIntercept != null && isFinite(forcedIntercept)) {
    slope = sxx === 0 ? 0 : (sxy - forcedIntercept * sx) / sxx;
    intercept = forcedIntercept;
  } else {
    const denom = n * sxx - sx * sx;
    slope = denom === 0 ? 0 : (n * sxy - sx * sy) / denom;
    intercept = (sy - slope * sx) / n;
  }
  const mean = sy / n;
  let residual = 0;
  let total = 0;
  for (let i = 0; i < n; i++) {
    const error = ys[i] - (slope * xs[i] + intercept);
    residual += error * error;
    const centered = ys[i] - mean;
    total += centered * centered;
  }
  const rSquared = total === 0 ? (residual === 0 ? 1 : 0) : 1 - residual / total;
  return { slope, intercept, rSquared };
}

/** Fit a trendline to `(xs, ys)` data points (nulls already filtered by the
 *  caller). Implements every ECMA-376 `ST_TrendlineType` (§21.2.3.50):
 *   - `linear` — ordinary least-squares `y = m·x + b` (honors a forced
 *     `intercept`), sampled at the first and last x (a straight line).
 *   - `exp`, `log`, and `power` — least squares in their standard transformed
 *     domains, sampled as a bounded 65-point curve.
 *   - `poly` — order 2..6 least squares over a centred/scaled Vandermonde
 *     matrix, solved by modified Gram-Schmidt rather than unstable normal
 *     equations, then sampled as a bounded 65-point curve.
 *   - `movingAvg` — the trailing average of the previous `period` points
 *     (default 2), producing a point from index `period-1` onward.
 * Invalid transformed-domain points are omitted. Any non-finite fit or sample
 * fails closed instead of sending invalid geometry to Canvas. */
export function fitTrendline(
  xs: number[], ys: number[], type: string,
  opts?: {
    period?: number | null;
    order?: number | null;
    intercept?: number | null;
    forward?: number | null;
    backward?: number | null;
  },
): TrendlinePoints {
  const n = Math.min(xs.length, ys.length);
  if (n < 2) return { xs: [], ys: [] };
  if (type === 'linear') {
    const stats = linearTrendlineStats(xs, ys, opts?.intercept);
    if (!stats) return { xs: [], ys: [] };
    const x0 = xs[0]; const x1 = xs[n - 1];
    return {
      xs: [x0, x1],
      ys: [stats.slope * x0 + stats.intercept, stats.slope * x1 + stats.intercept],
    };
  }
  if (type === 'movingAvg') {
    const period = Math.max(2, Math.round(opts?.period ?? 2));
    if (n < period) return { xs: [], ys: [] };
    const ox: number[] = []; const oy: number[] = [];
    let sum = 0;
    for (let index = 0; index < period; index++) sum += ys[index];
    for (let i = period - 1; i < n; i++) {
      ox.push(xs[i]); oy.push(sum / period);
      const next = i + 1;
      if (next < n) sum += ys[next] - ys[next - period];
    }
    return { xs: ox, ys: oy };
  }

  const transformed: Array<{ x: number; y: number }> = [];
  for (let index = 0; index < n; index++) {
    const x = xs[index];
    const y = ys[index];
    if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
    if ((type === 'log' && x <= 0) || (type === 'exp' && y <= 0)
      || (type === 'power' && (x <= 0 || y <= 0))) continue;
    transformed.push({ x, y });
  }
  if (transformed.length < 2) return { xs: [], ys: [] };

  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  for (const point of transformed) {
    minX = Math.min(minX, point.x);
    maxX = Math.max(maxX, point.x);
  }
  if (!Number.isFinite(minX) || !Number.isFinite(maxX) || maxX <= minX) {
    return { xs: [], ys: [] };
  }
  const backward = Number.isFinite(opts?.backward) ? Math.max(0, opts?.backward ?? 0) : 0;
  const forward = Number.isFinite(opts?.forward) ? Math.max(0, opts?.forward ?? 0) : 0;
  let sampleMin = minX - backward;
  const sampleMax = maxX + forward;
  if (type === 'log' || type === 'power') sampleMin = Math.max(Number.MIN_VALUE, sampleMin);
  if (!Number.isFinite(sampleMin) || !Number.isFinite(sampleMax) || sampleMax <= sampleMin) {
    return { xs: [], ys: [] };
  }

  const sample = (evaluate: (x: number) => number): TrendlinePoints => {
    const sampledXs: number[] = [];
    const sampledYs: number[] = [];
    for (let index = 0; index <= 64; index++) {
      const fraction = index / 64;
      const x = sampleMin * (1 - fraction) + sampleMax * fraction;
      const y = evaluate(x);
      if (!Number.isFinite(x) || !Number.isFinite(y)) return { xs: [], ys: [] };
      sampledXs.push(x);
      sampledYs.push(y);
    }
    return { xs: sampledXs, ys: sampledYs };
  };

  if (type === 'exp' || type === 'log' || type === 'power') {
    const regressionXs = transformed.map(point =>
      type === 'log' || type === 'power' ? Math.log(point.x) : point.x
    );
    const regressionYs = transformed.map(point =>
      type === 'exp' || type === 'power' ? Math.log(point.y) : point.y
    );
    const stats = linearTrendlineStats(regressionXs, regressionYs);
    if (!stats || ![stats.slope, stats.intercept].every(Number.isFinite)) {
      return { xs: [], ys: [] };
    }
    if (type === 'exp') {
      const coefficient = Math.exp(stats.intercept);
      return sample(x => coefficient * Math.exp(stats.slope * x));
    }
    if (type === 'log') return sample(x => stats.slope * Math.log(x) + stats.intercept);
    const coefficient = Math.exp(stats.intercept);
    return sample(x => coefficient * x ** stats.slope);
  }

  if (type === 'poly') {
    const degree = Math.min(
      6,
      transformed.length - 1,
      Math.max(2, Math.round(opts?.order ?? 2)),
    );
    if (degree < 2) return { xs: [], ys: [] };
    // Halved addition avoids overflow for large opposite-sign endpoints.
    const center = minX / 2 + maxX / 2;
    const scale = Math.max(Math.abs(minX - center), Math.abs(maxX - center));
    if (!Number.isFinite(center) || !Number.isFinite(scale) || scale <= 0) {
      return { xs: [], ys: [] };
    }
    const rowCount = transformed.length;
    const columnCount = degree + 1;
    const columns: number[][] = Array.from({ length: columnCount }, () => Array(rowCount).fill(0));
    for (let row = 0; row < rowCount; row++) {
      const z = (transformed[row].x - center) / scale;
      let power = 1;
      for (let column = 0; column < columnCount; column++) {
        columns[column][row] = power;
        power *= z;
      }
    }
    const orthonormal: number[][] = [];
    const upper: number[][] = Array.from(
      { length: columnCount },
      () => Array(columnCount).fill(0),
    );
    const projectedY = Array(columnCount).fill(0);
    for (let column = 0; column < columnCount; column++) {
      const vector = columns[column].slice();
      for (let prior = 0; prior < column; prior++) {
        let projection = 0;
        for (let row = 0; row < rowCount; row++) projection += orthonormal[prior][row] * vector[row];
        upper[prior][column] = projection;
        for (let row = 0; row < rowCount; row++) vector[row] -= projection * orthonormal[prior][row];
      }
      let magnitudeSquared = 0;
      for (const value of vector) magnitudeSquared += value * value;
      const magnitude = Math.sqrt(magnitudeSquared);
      if (!Number.isFinite(magnitude) || magnitude <= Number.EPSILON * Math.sqrt(rowCount)) {
        return { xs: [], ys: [] };
      }
      upper[column][column] = magnitude;
      const normalized = vector.map(value => value / magnitude);
      orthonormal.push(normalized);
      let projection = 0;
      for (let row = 0; row < rowCount; row++) projection += normalized[row] * transformed[row].y;
      if (!Number.isFinite(projection)) return { xs: [], ys: [] };
      projectedY[column] = projection;
    }
    const coefficients = Array(columnCount).fill(0);
    for (let row = columnCount - 1; row >= 0; row--) {
      let remainder = projectedY[row];
      for (let column = row + 1; column < columnCount; column++) {
        remainder -= upper[row][column] * coefficients[column];
      }
      coefficients[row] = remainder / upper[row][row];
      if (!Number.isFinite(coefficients[row])) return { xs: [], ys: [] };
    }
    return sample(x => {
      const z = (x - center) / scale;
      let value = coefficients[degree];
      for (let index = degree - 1; index >= 0; index--) value = value * z + coefficients[index];
      return value;
    });
  }
  return { xs: [], ys: [] };
}
