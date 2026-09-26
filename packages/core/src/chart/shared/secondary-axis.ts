// Classic chart secondary axis helpers.
import type { ChartModel, ChartRect, ChartSeries, SecondaryValueAxis } from '../../types/chart';
import { planDateCategoryAxis } from '../date-axis.js';
import { isCrossBetween, resolveAxisLine } from '../axis-style.js';
import { planNumericValueAxis } from '../axis-scale.js';
import { axisTitleFontPx, categoryTickLabelGapPx, chartTextFontSizePx } from '../layout.js';
import { rawLinkedChartStyleRole } from '../effective-style.js';
import { categoryLabelAnchorFraction, categoryLabelOffsetPx } from '../category-spacing.js';
import { elideToWidth } from '../text-elide.js';
import { formatCategoryLabel } from '../chart-number-format.js';
import { catAxisReversed, drawAxisTick, drawAxisTitle, formatAxisTickWithUnits, secondaryMajorGridStroke, secondaryMinorGridStroke, strokeAxisSegment, strokeValueGridlineH, valueAxisUnitInRendererSpace } from './axis.js';
import { chartFontCss, chartFontFamily } from './fonts.js';
import { effectiveLinkedLabelBox } from './style-roles.js';


/** Resolved secondary value-axis scale (combo charts). `min`/`max`/`step` are
 *  the "nice" bounds + major unit; `makeToY(py0, ph)` builds the value→pixel
 *  mapping once the final plot rect is known (the scale is computed BEFORE the
 *  pad/gutter math from an estimated plot height, so the mapping factory is
 *  split out). See {@link computeSecondaryAxis}. */
export interface SecondaryAxisScale {
  min: number;
  max: number;
  step: number;
  majorLines: number[];
  minorTicks: number[];
  makeToY: (py0: number, ph: number) => (v: number) => number;
}


/** Build the shared calendar/category mapping for every classic chart family
 * that can be bound to `<c:dateAx>`. Keeping this in one place prevents combo,
 * line, area and stock renderers from assigning different x coordinates to
 * the same authored date cache. */
export function chartDateAxisPlan(
  chart: ChartModel,
  categories: readonly string[],
  reversed = catAxisReversed(chart),
): ReturnType<typeof planDateCategoryAxis> {
  if (chart.catAxisIsDate !== true) return null;
  return planDateCategoryAxis({
    categories,
    date1904: chart.date1904,
    baseTimeUnit: chart.catAxisBaseTimeUnit,
    majorTimeUnit: chart.catAxisMajorTimeUnit,
    majorUnit: chart.catAxisMajorUnit,
    minorTimeUnit: chart.catAxisMinorTimeUnit,
    minorUnit: chart.catAxisMinorUnit,
    explicitMin: chart.catAxisMin,
    explicitMax: chart.catAxisMax,
    crossBetween: isCrossBetween(chart),
    reversed,
  });
}


/** Visit the authored error-bar endpoint values on one numeric axis. The
 * parser has already expanded percentage/fixed/custom forms into per-point
 * positive magnitudes, so scale planning only needs the same plus/minus gates
 * used by paint. */
export function forEachErrorBarEndpoint(
  series: ChartSeries,
  direction: 'x' | 'y',
  baseAt: (index: number) => number | null,
  visit: (value: number) => void,
): void {
  for (const errorBars of series.errBars ?? []) {
    if (errorBars.dir !== direction) continue;
    const drawPlus = errorBars.barType === 'plus' || errorBars.barType === 'both';
    const drawMinus = errorBars.barType === 'minus' || errorBars.barType === 'both';
    const count = Math.max(series.values.length, errorBars.plus.length, errorBars.minus.length);
    for (let index = 0; index < count; index++) {
      const base = baseAt(index);
      if (base == null || !Number.isFinite(base)) continue;
      const plus = errorBars.plus[index];
      const minus = errorBars.minus[index];
      if (drawPlus && plus != null && Number.isFinite(plus)) visit(base + plus);
      if (drawMinus && minus != null && Number.isFinite(minus)) visit(base - minus);
    }
  }
}


/** Compute the INDEPENDENT scale of a secondary value axis from the series that
 *  opt into it (`useSecondaryAxis === true`). Shared by every axis family that
 *  supports a secondary axis (bar-combo line series, and plain line / area
 *  series): the axis has its own bounded automatic plan, with an explicit
 *  `<c:scaling><c:min/max>` (`sec.min`/`sec.max`) overriding. Returns
 *  null when no `SecondaryValueAxis` was parsed OR no series opts into it — the
 *  caller then keeps the single-axis path unchanged.
 *
 *  `plotHeightPt` is the estimated plot height in points (the axis is the
 *  vertical right edge, so its length drives the auto major unit). `getValues`
 *  yields each opted-in series' raw values.
 *
 *  Empty secondary data keeps the neutral 0..1 fallback. */
export function computeSecondaryAxis(
  sec: SecondaryValueAxis | null,
  seriesForSecondary: ChartSeries[],
  plotHeightPt: number,
  errorBarDirection: 'x' | 'y' = 'y',
  percentStacked = false,
  includeZero = false,
  isSecondary: (series: ChartSeries, index: number) => boolean = series =>
    series.useSecondaryAxis === true,
  valueAt: (series: ChartSeries, pointIndex: number, seriesIndex: number) => number | null =
    () => null,
): SecondaryAxisScale | null {
  if (!sec) return null;
  let dMin = Infinity;
  let dMax = -Infinity;
  const include = (value: number): void => {
    if (!Number.isFinite(value)) return;
    dMin = Math.min(dMin, value);
    dMax = Math.max(dMax, value);
  };
  if (includeZero) include(0);
  for (let seriesIndex = 0; seriesIndex < seriesForSecondary.length; seriesIndex++) {
    const s = seriesForSecondary[seriesIndex];
    if (!isSecondary(s, seriesIndex)) continue;
    for (let pointIndex = 0; pointIndex < s.values.length; pointIndex++) {
      const value = valueAt(s, pointIndex, seriesIndex) ?? s.values[pointIndex];
      if (value != null) include(value);
    }
    forEachErrorBarEndpoint(
      s,
      errorBarDirection,
      pointIndex => valueAt(s, pointIndex, seriesIndex) ?? s.values[pointIndex] ?? null,
      include,
    );
  }
  if (!Number.isFinite(dMin) || !Number.isFinite(dMax)) {
    dMin = 0;
    dMax = 1;
  }
  // An explicit `<c:valAx><c:majorUnit>` on the secondary axis (§21.2.2.103)
  // overrides the auto step, mirroring the primary axis. null ⇒ auto.
  const numeric = planNumericValueAxis({
    dataMin: dMin,
    dataMax: dMax,
    explicitMin: valueAxisUnitInRendererSpace(sec.min, percentStacked),
    explicitMax: valueAxisUnitInRendererSpace(sec.max, percentStacked),
    axisLenPt: plotHeightPt,
    axisOrientation: 'vertical',
    majorUnit: valueAxisUnitInRendererSpace(sec.majorUnit, percentStacked),
    minorUnit: valueAxisUnitInRendererSpace(sec.minorUnit, percentStacked),
    needMinor: sec.minorGridlines === true
      || (sec.minorTickMark != null && sec.minorTickMark !== 'none'),
    logBase: sec.logBase,
    reversed: sec.orientation === 'maxMin',
  });
  const { min, max, majorUnit: step } = numeric;
  return {
    min,
    max,
    step,
    majorLines: numeric.majorTicks,
    minorTicks: numeric.minorTicks,
    makeToY: (py0: number, ph: number) => (v: number): number =>
      py0 + ph - numeric.fraction(v) * ph,
  };
}


/** Paint secondary-axis gridlines below chart data. Axis rules, ticks and
 * labels remain in {@link drawSecondaryValueAxis}, which is intentionally
 * called after the series. Keeping these layers separate prevents translucent
 * fills and bars from being incorrectly overpainted by right-axis grids. */
export function drawSecondaryValueGridlines(
  ctx: CanvasRenderingContext2D,
  sec: SecondaryValueAxis,
  secScale: SecondaryAxisScale,
  toYSecondary: (v: number) => number,
  px0: number,
  pw: number,
  ptToPx: number,
): void {
  if (sec.hidden) return;
  ctx.save();
  if (sec.minorGridlines) {
    const grid = secondaryMinorGridStroke(sec, ptToPx);
    for (const value of secScale.minorTicks) {
      strokeValueGridlineH(ctx, px0, pw, toYSecondary(value), false, grid);
    }
  }
  if (sec.majorGridlines) {
    const grid = secondaryMajorGridStroke(sec, ptToPx);
    for (const value of secScale.majorLines) {
      strokeValueGridlineH(ctx, px0, pw, toYSecondary(value), false, grid);
    }
  }
  ctx.restore();
}


/** Draw a secondary value axis on the RIGHT edge of the plot: its rule, mirrored
 *  tick marks + labels, and rotated title. Its scale is INDEPENDENT of the
 *  primary axis (its own "nice" major unit; NOT aligned to the primary
 *  gridlines) — PowerPoint places these marks independently. Shared by the
 *  bar, line and area families so no combo path can retain a divergent rule,
 *  tick, label, or title policy.
 *  Callers pass:
 *  - `secScale`   the resolved scale (from {@link computeSecondaryAxis}),
 *  - `toYSecondary` the value→pixel map (`secScale.makeToY(py0, ph)`),
 *  - `secFontPx` / `secLabelBandW` the tick-label font size + reserved gutter
 *    width (measured up front so the title clears the labels),
 *  - `primaryLabelColor` the fallback tick-label color when the axis specifies
 *    none (the primary value-axis label color). */
export function drawSecondaryValueAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  sec: SecondaryValueAxis,
  secScale: SecondaryAxisScale,
  toYSecondary: (v: number) => number,
  chartRect: ChartRect,
  px0: number, py0: number, pw: number, ph: number,
  ptToPx: number,
  secFontPx: number,
  secLabelBandW: number,
  primaryLabelColor: string,
  date1904: boolean | undefined,
  percentStacked = false,
): void {
  const axX = px0 + pw;
  const { color: secLineColor, width: secLineW } = resolveAxisLine(sec.lineColor, sec.lineWidthEmu, ptToPx);
  if (!sec.lineHidden) {
    strokeAxisSegment(ctx, axX, py0, axX, py0 + ph, secLineColor, secLineW, sec.lineDash);
  }
  if (!sec.hidden) {
    ctx.font = `${sec.fontItalic ? 'italic ' : ''}${sec.fontBold ? 'bold ' : ''}${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    ctx.fillStyle = sec.fontColor ? `#${sec.fontColor}` : primaryLabelColor;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    for (const sval of secScale.majorLines) {
      const gy = toYSecondary(sval);
      // Same tick geometry as the left axis, mirrored to the right edge.
      drawAxisTick(ctx, sec.majorTickMark, 'val', axX, gy, secLineColor, secLineW, true, sec.lineHidden, 'major', ptToPx, sec.lineDash);
      if (sec.tickLabelPos !== 'none') {
        ctx.fillText(
          formatAxisTickWithUnits(
            percentStacked ? sval / 100 : sval,
            sec.formatCode ?? null,
            date1904,
            sec.displayUnits,
          ),
          axX + 14,
          gy,
        );
      }
    }
    if (sec.minorTickMark && sec.minorTickMark !== 'none') {
      for (const value of secScale.minorTicks) {
        drawAxisTick(ctx, sec.minorTickMark, 'val', axX, toYSecondary(value), secLineColor, secLineW, true, sec.lineHidden, 'minor', ptToPx, sec.lineDash);
      }
    }
  }
  if (sec.title) {
    drawSecondaryAxisTitle(
      ctx, chart, sec, chartRect, px0, py0, pw, ph, secLabelBandW, ptToPx,
    );
  }
}


/** Draw a right-side secondary value-axis title. Both the duplicated combo-bar
 *  path and the shared line/area path use this helper, so the fixed 10pt
 *  fallback and +90° reading direction cannot drift. */
export function drawSecondaryAxisTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  sec: SecondaryValueAxis,
  chartRect: ChartRect,
  px0: number, py0: number, pw: number, ph: number,
  secLabelBandW: number,
  ptToPx: number,
): void {
  if (!sec.title) return;
  const fontSizePx = axisTitleFontPx(sec.titleFontSizeHpt, ptToPx);
  const color = sec.titleFontColor
    ? `#${sec.titleFontColor}`
    : (sec.fontColor ? `#${sec.fontColor}` : '#555');
  drawAxisTitle(
    ctx,
    sec.title,
    px0 + pw + secLabelBandW + fontSizePx * 0.6,
    py0 + ph / 2,
    'right',
    fontSizePx,
    sec.titleFontBold ?? true,
    sec.titleFontItalic ?? false,
    color,
    ph,
    chartFontFamily(chart, sec.titleFontFace, 'major'),
    sec.titleRotation,
    sec.titleVerticalMode,
    sec.titleManualLayout,
    chartRect,
    effectiveLinkedLabelBox(
      chart,
      sec.titleStyle ? { style: sec.titleStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
      true,
    ),
    ptToPx,
  );
}


/** Draw the categorical axis paired with a secondary bar/column group. Its
 * top rule, ticks, labels and title are distinct authored objects from the
 * primary bottom category axis; the secondary value axis remains responsible
 * for the paired right-hand numeric scale. */
export function drawSecondaryCategoryAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  axis: SecondaryValueAxis,
  categories: readonly string[],
  chartRect: ChartRect,
  px0: number,
  py0: number,
  pw: number,
  ptToPx: number,
): void {
  if (axis.hidden || categories.length === 0) return;
  const { color, width } = resolveAxisLine(axis.lineColor, axis.lineWidthEmu, ptToPx);
  if (!axis.lineHidden) {
    strokeAxisSegment(ctx, px0, py0, px0 + pw, py0, color, width, axis.lineDash);
  }
  const reversed = axis.orientation === 'maxMin';
  const labelSkip = Math.max(1, Math.floor(axis.tickLabelSkip ?? 1));
  const markSkip = Math.max(1, Math.floor(axis.tickMarkSkip ?? 1));
  const count = categories.length;
  const labelAnchor = (index: number) => categoryLabelAnchorFraction(
    index,
    count,
    isCrossBetween(chart),
    reversed,
    axis.labelAlignment,
  );
  if (!axis.lineHidden && axis.majorTickMark !== 'none') {
    const onBoundary = isCrossBetween(chart);
    const last = onBoundary ? count : count - 1;
    for (let index = 0; index <= last; index += markSkip) {
      const logical = reversed ? last - index : index;
      const fraction = onBoundary
        ? logical / count
        : count === 1 ? 0.5 : logical / (count - 1);
      drawAxisTick(
        ctx, axis.majorTickMark, 'cat', py0, px0 + fraction * pw,
        color, width, true, axis.lineHidden, 'major', ptToPx, axis.lineDash,
      );
    }
  }
  const fontPx = chartTextFontSizePx(axis.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  if (axis.tickLabelPos !== 'none') {
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, axis.fontFace, 'minor'),
      axis.fontBold ?? false,
      axis.fontItalic ?? false,
    );
    ctx.fillStyle = axis.fontColor ? `#${axis.fontColor}` : '#555';
    ctx.textBaseline = 'bottom';
    const budget = Math.max(1, pw / count - 4);
    const labelGap = categoryLabelOffsetPx(
      categoryTickLabelGapPx(fontPx),
      axis.labelOffsetPercent,
    );
    for (let index = 0; index < count; index += labelSkip) {
      const anchor = labelAnchor(index);
      ctx.textAlign = anchor.textAlign;
      ctx.fillText(
        elideToWidth(ctx, formatCategoryLabel(categories[index], axis.formatCode, chart.date1904), budget),
        px0 + anchor.fraction * pw,
        py0 - labelGap,
      );
    }
  }
  if (axis.title) {
    const titleFontPx = axisTitleFontPx(axis.titleFontSizeHpt, ptToPx);
    drawAxisTitle(
      ctx,
      axis.title,
      px0 + pw / 2,
      py0 - (axis.tickLabelPos === 'none'
        ? 0
        : fontPx + categoryLabelOffsetPx(
          categoryTickLabelGapPx(fontPx),
          axis.labelOffsetPercent,
        ))
        - titleFontPx / 2 - 4,
      'horizontal',
      titleFontPx,
      axis.titleFontBold ?? true,
      axis.titleFontItalic ?? false,
      axis.titleFontColor ? `#${axis.titleFontColor}` : '#555',
      pw,
      chartFontFamily(chart, axis.titleFontFace, 'major'),
      axis.titleRotation,
      axis.titleVerticalMode,
      axis.titleManualLayout,
      chartRect,
      effectiveLinkedLabelBox(
        chart,
        axis.titleStyle ? { style: axis.titleStyle } : undefined,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
        true,
      ),
      ptToPx,
    );
  }
}
