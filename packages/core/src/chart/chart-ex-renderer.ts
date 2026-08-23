// Optional Microsoft ChartEx family renderer. Classic DrawingML chart
// families remain in renderer.ts so format entry bundles do not pull this
// module unless the caller imports @silurus/ooxml/chart-ex.

import type {
  ChartDataPointOverride,
  ChartModel,
  ChartRect,
  ChartSeries,
  SecondaryValueAxis,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import {
  AXIS_OUTER_TEXT_MARGIN_PT,
  catAxisLabelBandH,
  categoryTickLabelGapPx,
  chartAxisTitleBands,
  chartLegendBands,
  chartTextFontSizePx,
  computeChartFrame,
  valueTickLabelGapPx,
} from './layout.js';
import { niceStep, planLinearValueAxis } from './axis-scale.js';
import { axisLineWidthPx, resolveAxisLine } from './axis-style.js';
import { formatCategoryLabel } from './chart-number-format.js';
import { resolveCategoryGapWidthPercent } from './category-spacing.js';
import { planHistogramBins } from './histogram-binning.js';
import {
  boxWhiskerGeometry,
  boxWhiskerPointCount,
  computeBoxWhiskerStats,
} from './box-whisker.js';
import { planParetoLayout } from './pareto-layout.js';
import { markerPaintComponents } from './marker-style.js';
import {
  MAX_CANVAS_CHART_POINTS,
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
} from './resource-limits.js';
import {
  buildSunburstTree,
  hierarchyInputTooLarge,
  layoutSunburstAngles,
  sunburstMaxDepth,
  type SunburstNode,
} from './chart-ex-hierarchy.js';
import { paintPlotAreaFrame } from './plot-area-frame.js';
import {
  applyChartExSeriesLineStyle,
  applyResolvedChartExLineStyle,
  axisLabelPx,
  chartColor,
  chartExDataPointFill,
  chartExDataPointPaint,
  chartExFillStyle,
  chartExLegendSeries,
  chartExMarkerPaint,
  chartExSeriesFormatIndex,
  chartExStyleColor,
  chartExValueTickLabelOffsetPx,
  chartFontCss,
  chartFontFamily,
  drawAxisTick,
  drawAxisTitles,
  drawBoundedDataLabelText,
  drawChartTitleForLayout,
  drawLegendForLayout,
  drawMarker,
  drawValMajorGridlines,
  formatPrimaryValueAxisTick,
  indexPointOverrides,
  measuredCartesianTitleBand,
  measuredLegendReserve,
  planValueAxis,
  rejectOversizedCanvasChart,
  renderBarChart,
  renderLineChart,
  resolveChartExLabel,
  resolveChartExSeriesLineStyle,
  richDataLabelOptions,
  strokeAxisSegment,
  strokeValueGridlineH,
  valGridStroke,
  valMinorGridStroke,
  wrapMeasuredText,
  type ChartExStyle,
} from './renderer.js';

function chartExStyleAuthorsFill(style: ChartExStyle | null | undefined): boolean {
  if (!style || style.fillNoStyle === true) return false;
  return style.fillPaintAuthored === true
    || style.fillHidden === true
    || style.fillColors?.some(color => color != null) === true
    || style.fillPaints?.some(paint => paint != null) === true;
}

function waterfallPointPaint(
  chart: ChartModel,
  point: ChartDataPointOverride | undefined,
  series: ChartSeries | undefined,
  semanticIndex: number,
): Fill | null {
  const pointAuthorsFill = point?.fillHidden === true
    || point?.color != null
    || chartExStyleAuthorsFill(point?.chartexStyle);
  if (pointAuthorsFill) {
    const pointStyle = point?.fillHidden === true
      ? { ...point.chartexStyle, fillHidden: true, fillPaintAuthored: true }
      : point?.chartexStyle;
    return chartExMarkerPaint(
      chart, semanticIndex, 3, pointStyle, point?.color,
      // Prevent an authored-but-unresolved point paint from reviving a lower
      // Chart Style or semantic fill.
      { fillHidden: true, fillPaintAuthored: true },
    );
  }
  if (series?.chartexStyle?.fillPaintAuthored === true) {
    return chartExMarkerPaint(
      chart, semanticIndex, 3, series.chartexStyle, series.color,
      // An explicitly authored CT_Series fill choice owns unresolved/noFill;
      // the legacy fillHidden-only public shape lacks that provenance and is
      // intentionally handled by chartExDataPointPaint below.
      { fillHidden: true, fillPaintAuthored: true },
    );
  }
  // CT_Series.spPr formats the series carrier. ChartEx semantic data points
  // keep their dataPoint role when that carrier has noFill, while a positive
  // series paint still wins; chartExDataPointPaint owns that distinction.
  return chartExDataPointPaint(
    chart, semanticIndex, 3, series?.chartexStyle, series?.color,
  );
}

function waterfallPointAuthorsLine(point: ChartDataPointOverride | undefined): boolean {
  const style = point?.chartexStyle;
  return point?.lineHidden != null
    || point?.lineColor != null
    || point?.lineWidthEmu != null
    || point?.lineDash != null
    || style?.linePaintAuthored === true
    || style?.lineHidden != null
    || style?.lineColors?.some(color => color != null) === true
    || style?.linePaints?.some(paint => paint != null) === true
    || style?.lineWidthEmu != null
    || style?.lineDash != null
    || style?.lineCustomDash != null
    || style?.lineCap != null
    || style?.lineJoin != null;
}

/** Bound structured label-shape work owned by ChartEx hierarchy painters. */
/** @internal Exported for resource-boundary regression tests. */
export function chartExHierarchyLabelPaintWorkCount(chart: ChartModel): number | null {
  const hierarchy = chart.chartexSunburst
    ? { rows: chart.chartexSunburst.rows, kind: 'sunburst' as const }
    : chart.chartexTreemap
      ? { rows: chart.chartexTreemap.rows, kind: 'treemap' as const }
      : undefined;
  if (!hierarchy) return null;
  if (hierarchy.rows.length === 0) return 0;
  if (hierarchyInputTooLarge(hierarchy.rows)) return MAX_CHART_PAINT_COMPONENTS + 1;

  const root = buildSunburstTree(hierarchy.rows, hierarchy.kind === 'treemap');
  if (root.layoutWeight <= 0 || root.children.length === 0) return 0;
  if (hierarchy.kind === 'sunburst') {
    root.a0 = -Math.PI / 2;
    root.a1 = root.a0 + Math.PI * 2;
    layoutSunburstAngles(root);
  }
  const series = chart.series[0];
  const overrides = indexPointOverrides(series?.dataLabelOverrides);
  const parentMode = chart.chartexTreemap?.parentLabelLayout ?? 'overlapping';
  let total = 0;
  const pending = [...root.children];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    for (const child of node.children) pending.push(child);
    if (node.layoutWeight <= 0
      || (hierarchy.kind === 'sunburst' && node.a1 - node.a0 <= 1e-4)) continue;
    let label;
    if (hierarchy.kind === 'sunburst') {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: false, showVal: false, showCatName: false },
        overrides,
      );
    } else if (node.children.length > 0) {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: parentMode !== 'none', showVal: false, showCatName: true },
        overrides,
      );
      if (parentMode === 'overlapping' && node.depth !== 0) label = null;
    } else {
      label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: false, showVal: false, showCatName: false },
        overrides,
      );
    }
    for (const paint of [label?.labelBox?.fillPaint, label?.labelBox?.borderFill]) {
      if (!paint) continue;
      const components = markerPaintComponents(paint);
      if ((paint.fillType === 'gradient' && components > MAX_CHART_PAINT_RECIPE_COMPONENTS)
        || components > MAX_CHART_PAINT_COMPONENTS - total) {
        return MAX_CHART_PAINT_COMPONENTS + 1;
      }
      total += components;
    }
  }
  return total;
}

function renderHistogramChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const source = chart.series[0];
  if (!source) return;
  const plan = planHistogramBins(source.values, chart.chartexHistogramBinning ?? {});
  if (plan.kind === 'tooManyInputPoints') {
    rejectOversizedCanvasChart(ctx, rect, MAX_CANVAS_CHART_POINTS + 1);
    return;
  }
  renderBarChart(ctx, {
    ...chart,
    chartType: 'clusteredBar',
    categories: plan.categories,
    series: [{ ...source, categories: undefined, values: plan.counts }],
  }, rect, ptToPx, { gapPolicy: 'chartex' }, shapeRotationDeg);
}

function renderWaterfallChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  const { x, y, w, h } = r;
  const vals = chart.series[0]?.values ?? [];
  const cats = chart.categories;
  // ChartEx numeric and string dimensions are independent. Preserve the union
  // of their point indexes: a missing category removes only its label, while a
  // missing numeric point retains an empty category slot.
  const n = Math.max(vals.length, cats.length);
  if (n === 0) return;
  if (rejectOversizedCanvasChart(ctx, r, n)) return;

  const subSet = new Set(chart.subtotalIndices);
  let cumulativeOverflow = false;
  const safeAdd = (left: number, right: number): number => {
    const sum = left + right;
    if (Number.isFinite(sum)) return sum;
    cumulativeOverflow = true;
    return sum < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
  };
  let running = 0;
  const bars: Array<{
    start: number;
    end: number;
    isSub: boolean;
    isPos: boolean;
    hasValue: boolean;
    paintSlot: boolean;
  }> = [];
  let rawMax = Number.NEGATIVE_INFINITY;
  let rawMin = 0;
  for (let i = 0; i < n; i++) {
    const authoredValue = vals[i];
    const hasValue = authoredValue != null && Number.isFinite(authoredValue);
    // A missing dimension point still owns its authored category slot (the
    // historical zero-height placeholder). A present but non-finite value is
    // invalid numeric input and must not reach axis or Canvas geometry.
    const paintSlot = authoredValue == null || hasValue;
    const v = hasValue ? authoredValue as number : 0;
    const isSub = subSet.has(i);
    if (isSub) {
      const bar = { start: 0, end: v, isSub: true, isPos: true, hasValue, paintSlot };
      bars.push(bar);
      if (paintSlot) {
        rawMax = Math.max(rawMax, bar.start, bar.end);
        rawMin = Math.min(rawMin, bar.start, bar.end);
      }
      if (hasValue) {
        running = v;
      }
    } else {
      const next = safeAdd(running, v);
      const start = v >= 0 ? running : next;
      const end   = v >= 0 ? next : running;
      const bar = { start, end, isSub: false, isPos: v >= 0, hasValue, paintSlot };
      bars.push(bar);
      if (paintSlot) {
        rawMax = Math.max(rawMax, bar.start, bar.end);
        rawMin = Math.min(rawMin, bar.start, bar.end);
      }
      if (hasValue) {
        running = next;
      }
    }
  }

  if (cumulativeOverflow) {
    ctx.fillStyle = '#888';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('(chart values out of range)', x + w / 2, y + h / 2);
    return;
  }

  if (rawMax <= rawMin) return;

  // Office-observed ChartEx compatibility: an all-increase bridge with no
  // subtotal/total points keeps its value grid and rule but omits numeric tick
  // labels. The axis XML is otherwise the same as a subtotal bridge, so keep
  // this narrow family policy out of the shared linear-axis planner.
  const suppressImplicitValueTickLabels = subSet.size === 0
    && bars.every((bar, index) => !bar.hasValue || (vals[index] as number) >= 0);
  const valueTickLabelsVisible = !chart.valAxisHidden
    && chart.valAxisTickLabelPos !== 'none'
    && !suppressImplicitValueTickLabels;

  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const valFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valFont = chartFontFamily(chart, chart.valAxisFontFace, 'minor');
  const catFont = chartFontFamily(chart, chart.catAxisFontFace, 'minor');
  const provisionalPlan = planValueAxis(chart, rawMin, rawMax, h / ptToPx);

  ctx.save();
  let valLabelBandW = 0;
  if (valueTickLabelsVisible) {
    ctx.font = chartFontCss(
      valFontPx,
      valFont,
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    let maxWidth = 0;
    for (const value of provisionalPlan.majorLines) {
      maxWidth = Math.max(
        maxWidth,
        ctx.measureText(formatPrimaryValueAxisTick(chart, value, false)).width,
      );
    }
    valLabelBandW = maxWidth + 8;
  }

  // Category labels participate in layout. Measure wrapped lines with the
  // authored category-axis font and the available category interval rather
  // than placing every word on its own line or reserving a height fraction.
  const estimatedPlotW = Math.max(
    1,
    w - axBands.valBandW - valLabelBandW - w * 0.02,
  );
  const estimatedSlotW = estimatedPlotW / n;
  ctx.font = chartFontCss(
    catFontPx,
    catFont,
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  const wrappedCategories = cats.slice(0, n).map(category =>
    wrapMeasuredText(
      ctx,
      formatCategoryLabel(category, chart.catAxisFormatCode, chart.date1904),
      Math.max(1, estimatedSlotW - 8),
    )
  );
  let maxCategoryLines = 0;
  for (const lines of wrappedCategories) {
    if (lines.some(Boolean)) maxCategoryLines = Math.max(maxCategoryLines, lines.length);
  }
  const categoryLabelBandH = chart.catAxisHidden || maxCategoryLines === 0
    ? 0
    : maxCategoryLines * (catFontPx + 2) + 4;

  const series = chart.series[0];
  const labelOverrides = indexPointOverrides(series?.dataLabelOverrides);
  const pointOverrides = indexPointOverrides(series?.dataPointOverrides);
  const localStyle = series?.chartexStyle;
  const colorPos = `#${series?.color ?? chartExDataPointFill(chart, 0, 3, localStyle)}`;
  const colorNeg = `#${chartExDataPointFill(chart, 1, 3, localStyle)}`;
  const colorSub = `#${chartExDataPointFill(chart, 2, 3, localStyle)}`;
  const legendPaintPos = chartExDataPointPaint(chart, 0, 3, localStyle, series?.color);
  const legendPaintNeg = chartExDataPointPaint(chart, 1, 3, localStyle);
  const legendPaintSub = chartExDataPointPaint(chart, 2, 3, localStyle);
  const legendChart: ChartModel = {
    ...chart,
    chartType: 'clusteredBar',
    series: [
      chartExLegendSeries(
        chart, 'Increase', series, chart.chartexDataPointStyle, 0, 3, colorPos,
      ),
      chartExLegendSeries(
        chart, 'Decrease', series, chart.chartexDataPointStyle, 1, 3, colorNeg,
      ),
      chartExLegendSeries(
        chart, 'Total', series, chart.chartexDataPointStyle, 2, 3, colorSub,
      ),
    ],
  };
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const pad = {
    t: titleBand.bandH + legTopH + valFontPx / 2 + 2,
    r: legRightW + w * 0.02,
    b: legBottomH + axBands.catBandH + categoryLabelBandH,
    l: legLeftW + axBands.valBandW + (chart.valAxisHidden
      ? w * 0.02
      // Office keeps the visible Waterfall value-axis rule roughly 3% inside
      // the chart object even when this layout suppresses its implicit numeric
      // labels. A measured label band may be wider, but never collapses that
      // automatic edge gutter to zero.
      : Math.max(w * 0.03, valLabelBandW)),
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const plan = planValueAxis(chart, rawMin, rawMax, ph / ptToPx);
  const yOf = (value: number): number => py0 + ph - plan.frac(value) * ph;

  const valAxisLine = resolveAxisLine(
    chart.valAxisLineColor,
    chart.valAxisLineWidthEmu,
    ptToPx,
  );
  const catAxisLine = resolveAxisLine(
    chart.catAxisLineColor,
    chart.catAxisLineWidthEmu,
    ptToPx,
  );
  const valGridline = valGridStroke(chart, ptToPx);

  // ECMA-376 / chartEx §axis@hidden: when the value axis is hidden, skip the
  // value-axis gridlines, tick labels and the left segment of the L-frame.
  // This is the canonical PowerPoint look for waterfall analyses where the
  // value scale is implicit in the data labels on each bar.
  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valFontPx,
      valFont,
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#595959';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    const minorGridline = valMinorGridStroke(chart, ptToPx);
    for (const value of plan.minorLines) {
      strokeValueGridlineH(ctx, px0, pw, yOf(value), false, minorGridline);
    }
    for (const value of plan.majorLines) {
      const gy = yOf(value);
      if (drawValMajorGridlines(chart)) {
        ctx.strokeStyle = valGridline.color;
        ctx.lineWidth = valGridline.width;
        const previousDash = valGridline.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (valGridline.dash.length > 0) ctx.setLineDash(valGridline.dash);
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        if (valGridline.dash.length > 0) ctx.setLineDash(previousDash);
      }
      // Locale-independent §18.8.30 formatting (honoring `<c:valAx><c:numFmt>`),
      // matching the other renderers — `toLocaleString()` grouped by the
      // viewer's OS locale, so the same chart read differently across machines.
      if (valueTickLabelsVisible) {
        ctx.fillText(
          formatPrimaryValueAxisTick(chart, value, false),
          px0 - 4,
          gy,
        );
      }
      drawAxisTick(
        ctx,
        chart.valAxisMajorTickMark,
        'val',
        px0,
        gy,
        valAxisLine.color,
        valAxisLine.width,
        false,
        chart.valAxisLineHidden,
        'major',
        ptToPx,
        chart.valAxisLineDash,
      );
    }
    for (const value of plan.minorTicks) {
      drawAxisTick(
        ctx, chart.valAxisMinorTickMark, 'val', px0, yOf(value),
        valAxisLine.color, valAxisLine.width, false, chart.valAxisLineHidden, 'minor', ptToPx,
        chart.valAxisLineDash,
      );
    }
  }

  // L-frame: vertical (value-axis) rule + horizontal (category-axis) baseline.
  // Each segment is independently gated on its axis's `<c:delete>` and on
  // Office's rule/tick suppression for `<c:spPr><a:ln><a:noFill>`.
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden;
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (drawValLine) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      valAxisLine.color, valAxisLine.width, chart.valAxisLineDash,
    );
  }
  if (drawCatLine) {
    strokeAxisSegment(
      ctx, px0, py0 + ph, px0 + pw, py0 + ph,
      catAxisLine.color, catAxisLine.width, chart.catAxisLineDash,
    );
  }

  // ECMA-376 / chartEx §17.18.34 ST_GapAmount: gapWidth is the gap between
  // adjacent categories expressed as a percentage of the bar width
  // (legacy `<c:gapWidth val>`) or as a fraction (chartEx
  // `<cx:catScaling gapWidth>`, normalised to the same percent form by the
  // parser). The bar then occupies `catGap / (1 + gapWidth/100)`. Omitted
  // ChartEx spacing uses the shared ordinal-layout policy; an authored
  // value remains authoritative after parser normalization.
  const gapW = pw / n;
  const gapWidthPct = resolveCategoryGapWidthPercent(chart.barGapWidth, 'chartex');
  const barW = gapW / (1 + gapWidthPct / 100);

  bars.forEach((bar, i) => {
    const bx = px0 + gapW * i + (gapW - barW) / 2;
    const yTop = Math.min(yOf(bar.start), yOf(bar.end));
    const yBot = Math.max(yOf(bar.start), yOf(bar.end));
    const bh = Math.max(1, yBot - yTop);

    const accentIndex = bar.isSub ? 2 : bar.isPos ? 0 : 1;
    const point = pointOverrides.get(i);
    const paint = waterfallPointPaint(chart, point, series, accentIndex);
    const fallback = point?.color
      ? `#${point.color}`
      : bar.isSub ? colorSub : bar.isPos ? colorPos : colorNeg;
    if (bar.paintSlot && paint) {
      ctx.fillStyle = chartExFillStyle(ctx, paint, bx, yTop, barW, bh, fallback, shapeRotationDeg);
      ctx.fillRect(bx, yTop, barW, bh);
    }
    const lineColor = chartExStyleColor(chart, chart.chartexDataPointStyle, 'line', accentIndex, 3);
    const lineCarrier = waterfallPointAuthorsLine(point) ? point : series;
    if (bar.paintSlot && applyChartExSeriesLineStyle(
      ctx,
      chart,
      chart.chartexDataPointStyle,
      lineCarrier,
      accentIndex,
      3,
      lineColor ? `#${lineColor}` : fallback,
      ptToPx,
    )) {
      ctx.strokeRect(bx, yTop, barW, bh);
    }

    if (bar.paintSlot && bars[i + 1]?.paintSlot
      && i < n - 1 && chart.chartexConnectorLines !== false) {
      const nextBx = px0 + gapW * (i + 1) + (gapW - barW) / 2;
      const connY = bar.isPos ? yTop : yBot;
      ctx.save();
      const connectorLine = resolveChartExSeriesLineStyle(
        chart,
        chart.chartexSeriesLineStyle,
        series,
        accentIndex,
        3,
        '#000000',
        { linkedNoStyleFallback: true },
      );
      if (applyResolvedChartExLineStyle(ctx, connectorLine, ptToPx)) {
        // A linked `seriesLine` NoStyle delegates to Waterfall's semantic
        // connector rather than suppressing it. Office vector output from
        // both an unstyled bridge and an explicitly styled bridge establishes
        // the family default as a 0.75pt rule; authored width/color above stay
        // authoritative.
        // Apply the semantic 0.75pt width only when the common direct>linked
        // resolver found no authored width. Looking only at the linked role
        // here used to overwrite a direct CT_Series width.
        if (connectorLine.widthEmu == null) ctx.lineWidth = 0.75 * ptToPx;
        ctx.beginPath();
        ctx.moveTo(bx + barW, connY);
        ctx.lineTo(nextBx, connY);
        ctx.stroke();
      }
      ctx.restore();
    }

    const rawVal = bar.hasValue ? vals[i] as number : 0;
    const label = resolveChartExLabel(
      chart, series, i, cats[i] ?? '', rawVal,
      { visible: chart.showDataLabels, showVal: true, showCatName: false },
      labelOverrides,
      !bar.hasValue,
    );
    if (label) {
      const perPointColor = series?.dataLabelColors?.[i] ?? label.fontColor ?? null;
      const labelColor = perPointColor
        ? `#${perPointColor}`
        : chart.dataLabelFontColor
          ? `#${chart.dataLabelFontColor}`
          : '#595959';
      const dataLabelFontPx = chartTextFontSizePx(label.fontSizeHpt, ptToPx)
        ?? axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
      const dataLabelBold = label.fontBold ?? chart.dataLabelFontBold ?? false;
      const dataLabelFont = chartFontFamily(
        chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
      );
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${dataLabelBold ? 'bold ' : ''}${dataLabelFontPx}px ${dataLabelFont}`;
      drawBoundedDataLabelText(
        ctx,
        label.text,
        {
          kind: 'bar',
          rect: { x: bx, y: yTop, w: barW, h: bh },
          orientation: 'vertical',
          negative: rawVal < 0,
          position: label.position ?? 'outEnd',
        },
        { x: px0, y: py0, w: pw, h: ph },
        dataLabelFontPx,
        labelColor,
        label.manualLayout,
        { x, y, w, h },
        richDataLabelOptions(
          chart, label.richRuns, ptToPx, dataLabelFont, dataLabelBold, label.textStyle,
        ),
        undefined,
        label.textStyle,
        ptToPx,
        label.labelBox,
        shapeRotationDeg,
      );
    }
  });

  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
  // Category (transaction) labels below the bars → category-axis face.
  ctx.font = chartFontCss(
    catFontPx,
    catFont,
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  const labelY = py0 + ph + 4;
  for (let i = 0; i < n && !chart.catAxisHidden; i++) {
    const ccx = px0 + gapW * i + gapW / 2;
    const lines = wrappedCategories[i] ?? [];
    lines.forEach((line, lineIndex) =>
      line && ctx.fillText(line, ccx, labelY + lineIndex * (catFontPx + 2))
    );
  }

  drawAxisTitles(
    ctx,
    chart,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    legLeftW,
    legBottomH,
    axBands.catFontPx,
    axBands.valFontPx,
  );

  drawLegendForLayout(
    ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph,
    titleBand.bandH + 2, ptToPx,
    [legendPaintPos, legendPaintNeg, legendPaintSub], shapeRotationDeg,
  );

  ctx.restore();
}

/** ChartEx `funnel`: one centered horizontal bar per ordinal category. */
function renderFunnelChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  const values = chart.series[0]?.values ?? [];
  // ChartEx dimensions carry independent indexed points. Preserve category-only
  // slots as empty bars and numeric-only slots as unlabeled bars.
  const n = Math.max(values.length, chart.categories.length);
  if (n === 0) return;
  if (rejectOversizedCanvasChart(ctx, r, n)) return;
  let max = 0;
  for (let index = 0; index < n; index++) {
    max = Math.max(max, values[index] ?? 0);
  }
  if (!(max > 0)) return;
  const { x, y, w, h } = r;
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const series = chart.series[0];
  const labelOverrides = indexPointOverrides(series?.dataLabelOverrides);
  const color = `#${series?.color ?? chartExDataPointFill(chart, 0, 1, series?.chartexStyle)}`;
  const paint = chartExDataPointPaint(chart, 0, 1, series?.chartexStyle, series?.color);
  const legendChart: ChartModel = {
    ...chart,
    series: [chartExLegendSeries(
      chart,
      series?.name ?? '',
      series,
      chart.chartexDataPointStyle,
      0,
      1,
      color,
    )],
  };
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  ctx.save();
  ctx.font = chartFontCss(
    catFontPx,
    chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  let labelW = 0;
  if (!chart.catAxisHidden) {
    for (let index = 0; index < Math.min(n, chart.categories.length); index++) {
      labelW = Math.max(labelW, ctx.measureText(chart.categories[index]).width);
    }
    if (chart.categories.length > 0) labelW += 10;
  }
  const pad = {
    t: titleBand.bandH + legTopH + 2,
    r: legRightW + w * 0.02,
    b: legBottomH + h * 0.02,
    l: legLeftW + labelW + w * 0.02,
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const rowH = ph / n;
  const gapWidthPct = resolveCategoryGapWidthPercent(chart.barGapWidth, 'chartex');
  const barH = rowH / (1 + gapWidthPct / 100);
  for (let index = 0; index < n; index++) {
    const value = Math.max(0, values[index] ?? 0);
    const barW = pw * value / max;
    const bx = px0 + (pw - barW) / 2;
    const by = py0 + rowH * index + (rowH - barH) / 2;
    if (paint) {
      ctx.fillStyle = chartExFillStyle(ctx, paint, bx, by, barW, barH, color, shapeRotationDeg);
      ctx.fillRect(bx, by, barW, barH);
    }
    if (applyChartExSeriesLineStyle(
      ctx, chart, chart.chartexDataPointStyle, series, 0, 1, color, ptToPx,
    )) {
      ctx.strokeRect(bx, by, barW, barH);
    }
    const category = chart.categories[index];
    if (!chart.catAxisHidden && category != null) {
      ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
      ctx.textAlign = 'right';
      ctx.textBaseline = 'middle';
      ctx.fillText(category, px0 - 6, by + barH / 2);
    }
    const label = resolveChartExLabel(
      chart, series, index, category ?? '', value,
      { visible: false, showVal: false, showCatName: false },
      labelOverrides,
    );
    if (label) {
      const fontPx = chartTextFontSizePx(label.fontSizeHpt, ptToPx)
        ?? axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
      const labelFont = chartFontFamily(
        chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
      );
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.fontBold ? 'bold ' : ''}${fontPx}px ${labelFont}`;
      drawBoundedDataLabelText(
        ctx,
        label.text,
        {
          kind: 'bar',
          rect: { x: bx, y: by, w: barW, h: barH },
          orientation: 'horizontal',
          negative: false,
          position: label.position ?? 'ctr',
        },
        { x: px0, y: py0, w: pw, h: ph },
        fontPx,
        label.fontColor ? `#${label.fontColor}` : '#ffffff',
        label.manualLayout,
        { x, y, w, h },
        richDataLabelOptions(
          chart, label.richRuns, ptToPx, labelFont, label.fontBold ?? false, label.textStyle,
        ),
        undefined,
        label.textStyle,
        ptToPx,
        label.labelBox,
        shapeRotationDeg,
      );
    }
  }
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    const line = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
    ctx.strokeStyle = line.color;
    ctx.lineWidth = line.width;
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }
  drawLegendForLayout(ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph, titleBand.bandH + 2, ptToPx, [paint], shapeRotationDeg);
  ctx.restore();
}

/** ChartEx Pareto line: cumulative share of the source values. */
function renderParetoLineChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const source = chart.series[0];
  if (!source) return;
  const pointCount = Math.max(
    chart.categories.length,
    source.categories?.length ?? 0,
    source.values.length,
  );
  if (rejectOversizedCanvasChart(ctx, r, pointCount)) return;
  // A standalone `paretoLine` accumulates the authored sequence. Descending
  // frequency sorting belongs to an owner-backed Pareto chart, where columns
  // and their cumulative line are reordered together.
  const layout = planParetoLayout(source, chart.categories, { sortDescending: false });
  if (layout.points.length === 0) return;
  const styleIndex = chartExSeriesFormatIndex(source, 0);
  const paretoLine = resolveChartExSeriesLineStyle(
    chart,
    chart.chartexDataPointLineStyle,
    source,
    styleIndex,
    1,
    chartColor(0, source),
    { linkedNoStyleFallback: true },
  );
  renderLineChart(ctx, {
    ...chart,
    chartType: 'line',
    categories: layout.categories,
    series: [{
      ...layout.series,
      showMarker: false,
      lineHidden: !paretoLine.visible,
      lineColor: paretoLine.color.replace(/^#/, ''),
      lineWidthEmu: paretoLine.widthEmu,
      chartexStyle: {
        lineDash: paretoLine.dash,
        lineCap: paretoLine.cap,
        lineJoin: paretoLine.join,
      },
    }],
    // Pareto's ordinal category labels are suppressed, but its authored
    // category-axis rule remains visible. Keep those two concerns separate.
    catAxisHidden: false,
    catAxisTickLabelPos: 'none',
    showLegend: false,
    // A standalone paretoLine carries cumulative fractions as its actual
    // values, but Office lays them out on an ordinary decimal axis. Two vector
    // references at different chart sizes use the same omitted-axis contract:
    // 0..1.2 in 0.2 steps. Keep authored bounds/major units authoritative and
    // avoid making this semantic cumulative scale depend on pixel height.
    // The 0..100% secondary-axis convention belongs to owner-backed Pareto.
    valMin: chart.valMin ?? 0,
    valMax: chart.valMax ?? 1.2,
    valAxisMajorUnit: chart.valAxisMajorUnit ?? 0.2,
  }, r, ptToPx, shapeRotationDeg);
}

/** Owner-backed ChartEx Pareto: sorted frequency columns plus cumulative line. */
function renderParetoChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const owner = chart.series[0];
  if (!owner) return;
  const pointCount = Math.max(
    chart.categories.length,
    owner.categories?.length ?? 0,
    owner.values.length,
  );
  if (rejectOversizedCanvasChart(ctx, r, pointCount)) return;
  const layout = planParetoLayout(owner, chart.categories);
  if (layout.points.length === 0) return;

  const authoredLine = chart.series.find(series => series.seriesType === 'line');
  const lineStyleIndex = authoredLine?.chartexFormatIdx
    ?? owner.chartexFormatIdx
    ?? 0;
  const cumulativeLine: ChartSeries = {
    ...(authoredLine ?? layout.series),
    name: authoredLine?.name || 'Cumulative %',
    values: layout.series.values,
    categories: layout.categories,
    color: authoredLine?.color
      ?? chartExStyleColor(
        chart, chart.chartexDataPointLineStyle, 'line', lineStyleIndex, 1,
      )
      ?? owner.lineColor
      ?? owner.color,
    seriesType: 'line',
    useSecondaryAxis: true,
    showMarker: false,
  };
  const secondaryAxis: SecondaryValueAxis = {
    min: chart.secondaryValAxis?.min ?? 0,
    max: chart.secondaryValAxis?.max ?? 1,
    title: chart.secondaryValAxis?.title ?? null,
    hidden: chart.secondaryValAxis?.hidden ?? false,
    formatCode: chart.secondaryValAxis?.formatCode ?? '0%',
    fontColor: chart.secondaryValAxis?.fontColor ?? null,
    fontSizeHpt: chart.secondaryValAxis?.fontSizeHpt ?? null,
    fontFace: chart.secondaryValAxis?.fontFace ?? null,
    lineColor: chart.secondaryValAxis?.lineColor ?? null,
    lineWidthEmu: chart.secondaryValAxis?.lineWidthEmu ?? null,
    lineHidden: chart.secondaryValAxis?.lineHidden ?? false,
    majorTickMark: chart.secondaryValAxis?.majorTickMark ?? 'out',
    minorTickMark: chart.secondaryValAxis?.minorTickMark ?? null,
    majorUnit: chart.secondaryValAxis?.majorUnit ?? null,
    minorUnit: chart.secondaryValAxis?.minorUnit ?? null,
    titleFontSizeHpt: chart.secondaryValAxis?.titleFontSizeHpt ?? null,
    titleFontBold: chart.secondaryValAxis?.titleFontBold ?? null,
    titleFontColor: chart.secondaryValAxis?.titleFontColor ?? null,
    titleFontFace: chart.secondaryValAxis?.titleFontFace ?? null,
  };

  renderBarChart(ctx, {
    ...chart,
    chartType: 'clusteredBar',
    categories: layout.categories,
    series: [
      { ...layout.orderedSeries, seriesType: null, useSecondaryAxis: false },
      cumulativeLine,
    ],
    secondaryValAxis: secondaryAxis,
  }, r, ptToPx, {
    gapPolicy: 'chartex',
    semanticLineNoStyleFallback: true,
  }, shapeRotationDeg);
}

// ─── chartEx: box-and-whisker (CH15, MS 2014 chartex ext) ────────────────────

// Application-defined defaults: the ChartEx schema does not encode these.
// Keep them named and local to this family so they cannot silently affect the
// specification-defined classic value-axis planner.
const BOX_WHISKER_AUTO_INTERVAL_TARGET = 7;
const BOX_WHISKER_AUTO_PADDING_RATIO = 0.05;
const BOX_WHISKER_ZERO_ANCHOR_RATIO = 1.2;

/** Compact omitted-axis policy observed in three independent Office vector
 * box/whisker outputs (small, ordinary, and wide numeric ranges). OOXML does
 * not define the automatic major unit. Office consistently chooses the nearest
 * 1/2/5 ladder to about seven raw-range intervals for this family, then applies the same
 * 5% padding / zero-threshold rules as the shared planner. */
function automaticBoxWhiskerAxis(
  dataMin: number,
  dataMax: number,
  explicitMin?: number | null,
  explicitMax?: number | null,
): { min: number; max: number; majorUnit: number } | null {
  // Authored bounds constrain the automatic-unit calculation as well as the
  // final axis. ChartEx may write min/max while omitting majorUnit; Office then
  // chooses the box/whisker family unit from that authored span (for example
  // 1..3 becomes 0.2), rather than deriving a coarse unit from the raw data and
  // merely clipping it to the authored bounds afterwards.
  const boundedMin = explicitMin ?? dataMin;
  const boundedMax = explicitMax ?? dataMax;
  const span = boundedMax - boundedMin;
  if (!(span > 0) || !Number.isFinite(span)) return null;
  const majorUnit = niceStep(span, BOX_WHISKER_AUTO_INTERVAL_TARGET);
  if (!(majorUnit > 0) || !Number.isFinite(majorUnit)) return null;
  let paddedMin = dataMin - span * BOX_WHISKER_AUTO_PADDING_RATIO;
  let paddedMax = dataMax + span * BOX_WHISKER_AUTO_PADDING_RATIO;
  if (dataMin >= 0
    && (dataMin === 0 || dataMax > BOX_WHISKER_ZERO_ANCHOR_RATIO * dataMin)) paddedMin = 0;
  if (dataMax <= 0
    && (dataMax === 0
      || Math.abs(dataMin) > BOX_WHISKER_ZERO_ANCHOR_RATIO * Math.abs(dataMax))) {
    paddedMax = 0;
  }
  const min = explicitMin ?? Math.floor(paddedMin / majorUnit) * majorUnit;
  const max = explicitMax ?? Math.ceil(paddedMax / majorUnit) * majorUnit;
  if (![min, max].every(Number.isFinite) || !(max > min)) return null;
  return { min, max, majorUnit };
}

/**
 * Render a chartEx box-and-whisker chart (MS 2014 chartex extension — there is
 * no ECMA-376 section; the structure is Microsoft's `<cx:chartSpace>` with a
 * `<cx:series layoutId="boxWhisker">` per column, each referencing raw sample
 * points via `<cx:dataId>`). The parser (`parse_chartex_boxwhisker`) groups the
 * raw points by category and threads the `<cx:layoutPr>` visibility/statistics
 * flags into `chart.chartexBox`; this renderer derives the five-number summary
 * per (category, series) and draws, for each box: the IQR rectangle (Q1..Q3),
 * the median line, whiskers to the non-outlier min/max (with end caps), the
 * mean `×` marker, and outlier dots. Colors come from the theme accent palette
 * (`chart.chartexAccents`, cycled by series) — the blue/orange/gray Office
 * default — falling back to `CHART_PALETTE` when a resolver supplies no palette.
 */
function renderBoxWhiskerChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  // ChartEx's linked dataPointMarkerLayout is a generic marker recipe. Office
  // does not use its size for box-and-whisker observations: vector output uses
  // 3pt observation/outlier dots and a 6pt mean `x`, independent of box width.
  // The linked recipe still supplies the marker symbol and paint.
  const observationMarkerSizePt = 3;
  const meanMarkerRadiusPx = 3 * ptToPx;
  const box = chart.chartexBox;
  if (!box || box.categories.length === 0 || box.series.length === 0) return;
  const { x, y, w, h } = r;
  const pointCount = boxWhiskerPointCount(
    box.series.map(series => series.valuesByCategory),
    MAX_CANVAS_CHART_POINTS,
  );
  if (rejectOversizedCanvasChart(ctx, r, pointCount)) return;

  const rejectOutOfRange = (): void => {
    ctx.fillStyle = '#888';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('(chart values out of range)', x + w / 2, y + h / 2);
  };
  const usableScale = (scale: { min: number; max: number; step: number }): boolean => {
    const span = scale.max - scale.min;
    const intervalCount = span / scale.step;
    return Number.isFinite(scale.min)
      && Number.isFinite(scale.max)
      && Number.isFinite(scale.step)
      && Number.isFinite(span)
      && Number.isFinite(intervalCount)
      && scale.max > scale.min
      && scale.step > 0
      // Axis paint is synchronous. Refuse an authored tiny major unit before
      // entering a loop whose finite bounds could still imply vast work.
      && intervalCount <= 1000
      // A finite positive step may still be too small to advance a large
      // floating-point bound, which would wedge the major-gridline loop.
      && scale.min + scale.step > scale.min;
  };

  // Resolve the value range before laying out the plot. The tick labels are a
  // real layout input: reserve their measured width instead of a percentage of
  // the chart width, which made the axis drift right on wide ChartEx charts.
  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of box.series) {
    for (const group of s.valuesByCategory) {
      for (const v of group) {
        if (!Number.isFinite(v)) continue;
        if (v < dataMin) dataMin = v;
        if (v > dataMax) dataMax = v;
      }
    }
  }
  if (!isFinite(dataMin) || !isFinite(dataMax)) return;
  if (!Number.isFinite(dataMax - dataMin)) {
    rejectOutOfRange();
    return;
  }

  // Authored bounds/unit always win. An omitted unit still receives the compact
  // family default inferred from the vector corpus, using authored min/max as
  // the effective span when present.
  const automaticBoxAxis = chart.valAxisMajorUnit == null
    ? automaticBoxWhiskerAxis(dataMin, dataMax, chart.valMin, chart.valMax)
    : null;
  const boxAxisChart: ChartModel = automaticBoxAxis
    ? {
        ...chart,
        valMin: automaticBoxAxis.min,
        valMax: automaticBoxAxis.max,
        valAxisMajorUnit: automaticBoxAxis.majorUnit,
      }
    : chart;

  const font = chartFontFamily(chart, chart.valAxisFontFace, 'minor');
  const valFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  // Keep the existing font-relative layout reserve so the plot geometry stays
  // stable; only the painted ChartEx label uses the observed axis-relative
  // offset below.
  const valTickLabelLayoutGap = valueTickLabelGapPx(valFontPx);
  const valTickLabelPaintGap = chartExValueTickLabelOffsetPx(ptToPx);
  const provisionalScale = planLinearValueAxis({
    dataMin,
    dataMax,
    explicitMin: boxAxisChart.valMin,
    explicitMax: boxAxisChart.valMax,
    axisLenPt: h / ptToPx,
    majorUnit: boxAxisChart.valAxisMajorUnit,
  });
  if (!usableScale({
    min: provisionalScale.min,
    max: provisionalScale.max,
    step: provisionalScale.majorUnit,
  })) {
    rejectOutOfRange();
    return;
  }
  let valLabelBandW = 0;
  if (!chart.valAxisHidden) {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      valFontPx,
      font,
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    let maxLabelW = 0;
    for (const value of provisionalScale.majorTicks) {
      const label = formatPrimaryValueAxisTick(chart, value, false);
      maxLabelW = Math.max(maxLabelW, ctx.measureText(label).width);
    }
    ctx.font = previousFont;
    valLabelBandW = maxLabelW + valTickLabelLayoutGap + AXIS_OUTER_TEXT_MARGIN_PT * ptToPx;
  }

  // Shared title band + cartesian plot rect. Reserve category/value-axis bands
  // and, when present in the chart model, the authored legend band.
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const catAxFontPx0 = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx0 = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const nSer = box.series.length;
  const boxStyleIndices = box.series.map((series, index) =>
    chartExSeriesFormatIndex(series, index)
  );
  const boxLegendSeries = box.series.map((series, index) => {
    const styleIndex = boxStyleIndices[index];
    const fill = series.color ?? chartExDataPointFill(
      chart, styleIndex, nSer, series.chartexStyle,
    );
    return chartExLegendSeries(
      chart,
      series.name,
      series,
      chart.chartexDataPointStyle,
      styleIndex,
      nSer,
      fill,
      true,
    );
  });
  const legendChart: ChartModel = {
    ...chart,
    series: boxLegendSeries,
  };
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const pad = {
    t: titleBand.bandH + legTopH + valAxFontPx0 / 2 + 2,
    r: legRightW + w * 0.02,
    b: legBottomH + axBands.catBandH + (chart.catAxisHidden ? h * 0.02 : catAxisLabelBandH(catAxFontPx0)),
    l: legLeftW + axBands.valBandW + (chart.valAxisHidden ? w * 0.02 : valLabelBandW),
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  const { px0, py0, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  const cats = box.categories;
  const nCat = cats.length;

  // Excel's automatic value axis uses nice-rounded bounds and steps.
  const boxAxisPlan = planValueAxis(boxAxisChart, dataMin, dataMax, ph / ptToPx);
  if (!usableScale(boxAxisPlan)) {
    rejectOutOfRange();
    return;
  }
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, frame.title.fontPx);
  const { min: axisMin, max: axisMax } = boxAxisPlan;
  const span = axisMax - axisMin;
  const yOf = (v: number): number => py0 + ph * (1 - (v - axisMin) / span);

  const valAxisLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const valGridline = valGridStroke(chart, ptToPx);

  // Value-axis gridlines + labels (unless the value axis is hidden).
  ctx.save();
  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valFontPx,
      font,
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    if (chart.valAxisMinorGridlines) {
      const minorGridline = valMinorGridStroke(chart, ptToPx);
      for (const value of boxAxisPlan.minorLines) {
        strokeValueGridlineH(ctx, px0, pw, yOf(value), false, minorGridline);
      }
    }
    for (const v of boxAxisPlan.majorLines) {
      const gy = yOf(v);
      if (chart.valAxisMajorGridlines !== false) {
        ctx.strokeStyle = valGridline.color;
        ctx.lineWidth = valGridline.width;
        const previousDash = valGridline.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (valGridline.dash.length > 0) ctx.setLineDash(valGridline.dash);
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        if (valGridline.dash.length > 0) ctx.setLineDash(previousDash);
      }
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#595959';
      ctx.fillText(
        formatPrimaryValueAxisTick(chart, v, false),
        px0 - valTickLabelPaintGap,
        gy,
      );
      drawAxisTick(
        ctx,
        chart.valAxisMajorTickMark,
        'val',
        px0,
        gy,
        valAxisLine.color,
        valAxisLine.width,
        false,
        chart.valAxisLineHidden,
        'major',
        ptToPx,
        chart.valAxisLineDash,
      );
    }
    for (const value of boxAxisPlan.minorTicks) {
      drawAxisTick(
        ctx, chart.valAxisMinorTickMark, 'val', px0, yOf(value),
        valAxisLine.color, valAxisLine.width, false, chart.valAxisLineHidden, 'minor', ptToPx,
        chart.valAxisLineDash,
      );
    }
    if (!chart.valAxisLineHidden) {
      strokeAxisSegment(
        ctx, px0, py0, px0, py0 + ph,
        valAxisLine.color, valAxisLine.width, chart.valAxisLineDash,
      );
    }
  }
  // Category-axis baseline. ChartEx carries the axis rule in the local
  // `<cx:axis><cx:spPr><a:ln>`; do not replace an authored dark/weighted rule
  // with the old fixed 1px grey fallback.
  const catAxisLine = resolveAxisLine(
    chart.catAxisLineColor,
    chart.catAxisLineWidthEmu,
    ptToPx,
  );
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0 + ph, px0 + pw, py0 + ph,
      catAxisLine.color, catAxisLine.width, chart.catAxisLineDash,
    );
  }

  // ChartEx divides the plot into full category intervals, so the first and
  // last category centres are half an interval from the plot edges. Every
  // series owns one stable equal slot in each category group, including
  // categories where a peer has no retained observations.
  // Formula-only sources are different: Excel stores every visible box as a
  // separate series in one category group. In that form `gapWidth` surrounds
  // the whole group and must not be reapplied between the diagonal entries.
  const slotW = pw / nCat;
  const gapWidthPct = resolveCategoryGapWidthPercent(chart.barGapWidth, 'chartex');
  const paletteOf = (si: number): string => {
    const fill = box.series[si].color
      ?? chartExDataPointFill(
        chart, boxStyleIndices[si], nSer, box.series[si].chartexStyle,
      );
    return `#${fill}`;
  };
  const paintOf = (si: number): Fill | null => chartExDataPointPaint(
    chart,
    boxStyleIndices[si],
    nSer,
    box.series[si].chartexStyle,
    box.series[si].color,
  );
  const statsBySeries = box.series.map(series => series.valuesByCategory.map(values => (
    computeBoxWhiskerStats(values, series.quartileMethod)
  )));
  const boxGeometry = (ci: number, si: number): { bx: number; boxW: number; cx: number } => {
    const geometry = boxWhiskerGeometry(
      px0,
      pw,
      box.oneBoxPerSeries ? 1 : nCat,
      nSer,
      box.oneBoxPerSeries ? 0 : ci,
      si,
      gapWidthPct,
    );
    if (!geometry) return { bx: px0, boxW: 0, cx: px0 };
    return { bx: geometry.boxX, boxW: geometry.boxWidth, cx: geometry.centerX };
  };

  // `<cx:visibility meanLine>` connects the category means for one series.
  // It is a data-point-line role, so it shares the whisker/median style.
  for (let si = 0; si < nSer; si++) {
    const series = box.series[si];
    if (!series.meanLine) continue;
    const lineStyle = chart.chartexDataPointLineStyle ?? chart.chartexDataPointStyle;
    const fallback = series.lineColor ? `#${series.lineColor}` : paletteOf(si);
    ctx.save();
    const styleLineVisible = applyChartExSeriesLineStyle(
      ctx, chart, lineStyle, series, boxStyleIndices[si], nSer, fallback, ptToPx,
    );
    if (styleLineVisible || series.lineColor != null) {
      if (series.lineColor) ctx.strokeStyle = fallback;
      if (series.lineWidthEmu) ctx.lineWidth = axisLineWidthPx(series.lineWidthEmu, ptToPx);
      let open = false;
      ctx.beginPath();
      for (let ci = 0; ci < nCat; ci++) {
        const stats = statsBySeries[si][ci];
        if (!stats) {
          open = false;
          continue;
        }
        const { cx } = boxGeometry(ci, si);
        const meanY = yOf(stats.mean);
        if (open) ctx.lineTo(cx, meanY);
        else ctx.moveTo(cx, meanY);
        open = true;
      }
      ctx.stroke();
    }
    ctx.restore();
  }
  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const catTickLabelGap = categoryTickLabelGapPx(catFontPx);
  for (let ci = 0; ci < nCat; ci++) {
    const categoryCenterX = px0 + slotW * (ci + 0.5);
    if (!chart.catAxisHidden) {
      drawAxisTick(
        ctx,
        chart.catAxisMajorTickMark,
        'cat',
        py0 + ph,
        categoryCenterX,
        catAxisLine.color,
        catAxisLine.width,
        false,
        chart.catAxisLineHidden,
        'major',
        ptToPx,
        chart.catAxisLineDash,
      );
    }
    for (let si = 0; si < nSer; si++) {
      const s = box.series[si];
      const stats = statsBySeries[si][ci];
      if (!stats) continue;
      const { bx, boxW, cx } = boxGeometry(ci, si);
      const fill = paletteOf(si);
      const fillPaint = paintOf(si);
      const pointStyle = chart.chartexDataPointStyle;
      const lineStyle = chart.chartexDataPointLineStyle ?? pointStyle;
      const markerStyle = chart.chartexDataPointMarkerStyle ?? pointStyle;
      const styleIndex = boxStyleIndices[si];
      const styleLine = chartExStyleColor(chart, pointStyle, 'line', styleIndex, nSer);
      const edge = s.lineColor ? `#${s.lineColor}` : styleLine ? `#${styleLine}` : fill;
      const edgeWidth = s.lineWidthEmu
        ? axisLineWidthPx(s.lineWidthEmu, ptToPx)
        : pointStyle?.lineWidthEmu != null
          ? axisLineWidthPx(pointStyle.lineWidthEmu, ptToPx)
          : 1;
      const lineEdge = chartExStyleColor(chart, lineStyle, 'line', styleIndex, nSer);
      const markerFill = chartExStyleColor(chart, markerStyle, 'fill', styleIndex, nSer);
      const markerFillPaint = chartExMarkerPaint(
        chart, styleIndex, nSer, s.chartexStyle, s.color, markerStyle,
      );
      const markerEdge = chartExStyleColor(chart, markerStyle, 'line', styleIndex, nSer);
      const applySeriesLine = (style: ChartExStyle | null | undefined, fallback: string): boolean => {
        const local = s.chartexStyle;
        const hasLocalLine = local?.lineHidden != null
          || local?.lineColors?.some(Boolean)
          || local?.lineWidthEmu != null
          || local?.lineDash != null
          || local?.lineCap != null
          || local?.lineJoin != null;
        const visible = applyChartExSeriesLineStyle(
          ctx, chart, style, s, styleIndex, nSer, fallback, ptToPx,
        );
        // Chart Style `NoStyle` means that this role supplies no decorative
        // override. Box/whisker's semantic outline still exists. This is
        // distinct from an authored `<a:noFill>`, which remains suppressed.
        const semanticNoStyleFallback = style?.lineNoStyle === true
          && !hasLocalLine && s.lineColor == null && s.lineWidthEmu == null;
        if (!visible && semanticNoStyleFallback) {
          ctx.strokeStyle = fallback;
          ctx.lineWidth = 1;
          ctx.setLineDash([]);
        }
        if (s.lineColor) ctx.strokeStyle = edge;
        if (s.lineWidthEmu) ctx.lineWidth = edgeWidth;
        return visible || semanticNoStyleFallback || s.lineColor != null;
      };
      const yQ1 = yOf(stats.q1);
      const yQ3 = yOf(stats.q3);
      const boxTop = Math.min(yQ1, yQ3);
      const boxH = Math.max(1, Math.abs(yQ1 - yQ3));

      // Whiskers: vertical line from box edges to whisker ends, with end caps.
      const capW = boxW * 0.4;
      if (applySeriesLine(lineStyle, lineEdge ?? edge)) {
        ctx.beginPath();
        ctx.moveTo(cx, yOf(stats.whiskerHi)); ctx.lineTo(cx, yQ3);
        ctx.moveTo(cx, yQ1); ctx.lineTo(cx, yOf(stats.whiskerLo));
        ctx.moveTo(cx - capW / 2, yOf(stats.whiskerHi)); ctx.lineTo(cx + capW / 2, yOf(stats.whiskerHi));
        ctx.moveTo(cx - capW / 2, yOf(stats.whiskerLo)); ctx.lineTo(cx + capW / 2, yOf(stats.whiskerLo));
        ctx.stroke();
      }

      // IQR box: solid accent fill + a thin accent×0.8 edge.
      if (fillPaint) {
        ctx.fillStyle = chartExFillStyle(
          ctx,
          fillPaint,
          bx,
          boxTop,
          boxW,
          boxH,
          fill,
          shapeRotationDeg,
        );
        ctx.fillRect(bx, boxTop, boxW, boxH);
      }
      if (applySeriesLine(pointStyle, edge)) {
        ctx.strokeRect(
          bx + edgeWidth / 2,
          boxTop + edgeWidth / 2,
          boxW - edgeWidth,
          boxH - edgeWidth,
        );
      }

      // Median line across the box.
      const yMed = yOf(stats.median);
      if (applySeriesLine(lineStyle, lineEdge ?? edge)) {
        ctx.beginPath(); ctx.moveTo(bx, yMed); ctx.lineTo(bx + boxW, yMed); ctx.stroke();
      }

      // Interior sample points. Excel overlays the raw non-outlier values on
      // the box/whiskers when cx:visibility@nonoutliers is enabled. Their
      // outline follows the owning box series, not the generic linked marker
      // role (which may carry a contrasting line intended for ordinary chart
      // markers).
      if (s.showNonoutliers) {
        const pointSymbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
        for (const point of stats.inner) {
          if (pointSymbol === 'none') continue;
          const markerLineVisible = applySeriesLine(markerStyle, markerEdge ?? edge);
          const pointY = yOf(point);
          drawMarker(
            ctx,
            cx,
            pointY,
            pointSymbol,
            observationMarkerSizePt,
            markerFillPaint ? (markerFill ? `#${markerFill}` : fill) : 'transparent',
            markerLineVisible ? edge : null,
            ptToPx,
            ctx.lineWidth,
            markerFillPaint,
            shapeRotationDeg,
          );
        }
      }

      // Mean `×` marker (same accent×0.8 as the rest of the outline).
      if (s.meanMarker) {
        const mY = yOf(stats.mean);
        const mR = meanMarkerRadiusPx;
        if (applySeriesLine(markerStyle, markerEdge ?? edge)) {
          ctx.beginPath();
          ctx.moveTo(cx - mR, mY - mR); ctx.lineTo(cx + mR, mY + mR);
          ctx.moveTo(cx + mR, mY - mR); ctx.lineTo(cx - mR, mY + mR);
          ctx.stroke();
        }
      }

      // Outlier dots.
      if (s.showOutliers) {
        const pointSymbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
        for (const o of stats.outliers) {
          if (pointSymbol === 'none') continue;
          const markerLineVisible = applySeriesLine(markerStyle, markerEdge ?? edge);
          const outlierY = yOf(o);
          drawMarker(
            ctx,
            cx,
            outlierY,
            pointSymbol,
            observationMarkerSizePt,
            markerFillPaint ? (markerFill ? `#${markerFill}` : fill) : 'transparent',
            markerLineVisible ? edge : null,
            ptToPx,
            ctx.lineWidth,
            markerFillPaint,
            shapeRotationDeg,
          );
        }
      }
    }

    // Category label (centered under the slot), word-wrapped like the other
    // cartesian renderers.
    if (!chart.catAxisHidden) {
      ctx.font = chartFontCss(
        catFontPx,
        chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
        chart.catAxisFontBold ?? false,
        chart.catAxisFontItalic ?? false,
      );
      ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      const label = cats[ci];
      ctx.fillText(label, categoryCenterX, py0 + ph + catTickLabelGap);
    }
  }
  ctx.restore();

  drawAxisTitles(
    ctx,
    chart,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    legLeftW,
    legBottomH,
    axBands.catFontPx,
    axBands.valFontPx,
  );

  drawLegendForLayout(
    ctx,
    legendChart,
    leg,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    titleBand.bandH + 2,
    ptToPx,
    box.series.map((_, index) => paintOf(index)),
    shapeRotationDeg,
  );
}

/** Application-defined automatic ChartEx sunburst center-hole ratio observed
 * across the current desktop-Office vector corpus. It is intentionally local
 * to sunburst and never reused as a generic radial-chart default. */
const SUNBURST_AUTOMATIC_HOLE_RATIO = 0.18;

/**
 * Render a chartEx sunburst (MS 2014 chartex extension — no ECMA-376 section;
 * the structure is a `<cx:series layoutId="sunburst">` over a `<cx:strDim
 * type="cat">` of several `<cx:lvl>` and one `<cx:numDim type="size">`). The
 * parser (`parse_chartex_sunburst`) yields the flat root→leaf `path`/`size`
 * rows in `chart.chartexSunburst`; this renderer folds them into a ring tree,
 * lays out each node's angular span proportional to its aggregated size, and
 * draws concentric rings (inner = root/Branch, outward = Stem, Leaf) from 12
 * o'clock clockwise. Every node in a branch shares that branch's theme accent
 * (`chart.chartexAccents`, cycled by top-level index — the blue/orange/gray
 * Office default). Labels are drawn white and centered in each segment, rotated
 * to follow the arc and elided when the wedge is too small.
 */
function renderSunburstChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  const sb = chart.chartexSunburst;
  if (!sb || sb.rows.length === 0) return;
  const { x, y, w, h } = r;
  if (hierarchyInputTooLarge(sb.rows)) {
    rejectOversizedCanvasChart(ctx, r, MAX_CANVAS_CHART_POINTS + 1);
    return;
  }
  const root = buildSunburstTree(sb.rows);
  if (root.layoutWeight <= 0 || root.children.length === 0) return;
  const series = chart.series[0];
  const labelOverrides = indexPointOverrides(series?.dataLabelOverrides);
  const legendPaints = root.children.map((_, index) =>
    chartExDataPointPaint(chart, index, root.children.length, series?.chartexStyle, series?.color)
  );
  const legendChart: ChartModel = {
    ...chart,
    chartType: 'clusteredBar',
    series: root.children.map(node => {
      const fill = chartExDataPointFill(
        chart, node.branchIndex, root.children.length, series?.chartexStyle,
      );
      return chartExLegendSeries(
        chart,
        node.label,
        series,
        chart.chartexDataPointStyle,
        node.branchIndex,
        root.children.length,
        fill,
        false,
        false,
      );
    }),
  };
  // Reuse the radial frame so an authored top legend reserves space above the
  // rings instead of being painted over the circle.
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0,
    legendReserve: leg,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const cx = px0 + pw / 2;
  const cy = py0 + ph / 2;
  const outerR = Math.min(pw, ph) * 0.46;

  // Full circle from 12 o'clock (−90°), clockwise (canvas angles grow CW), each
  // parent partitioning its range across its children in source (first-seen)
  // order. This is the natural spec-consistent reading of the `<cx:lvl>` point
  // order. The extension does not specify an alternative automatic reordering,
  // so the renderer does not infer one from labels or values.
  root.a0 = -Math.PI / 2;
  root.a1 = -Math.PI / 2 + Math.PI * 2;
  layoutSunburstAngles(root);

  const maxDepth = sunburstMaxDepth(root); // 0-based deepest ring index
  const ringCount = maxDepth + 1;
  // Small center hole (Office draws a modest hole, ~18% of the outer radius);
  // the remaining band is split evenly across the rings.
  const innerR = outerR * SUNBURST_AUTOMATIC_HOLE_RATIO;
  const ringBand = (outerR - innerR) / ringCount;

  const branchColor = (bi: number): string => {
    const hex = chartExDataPointFill(chart, bi, root.children.length, series?.chartexStyle);
    return `#${hex}`;
  };
  const branchPaint = (bi: number): Fill | null => chartExDataPointPaint(
    chart,
    bi,
    root.children.length,
    series?.chartexStyle,
    series?.color,
  );

  const labelDef = series?.seriesDataLabels;
  const labelFont = chartFontFamily(
    chart, labelDef?.fontFace ?? chart.dataLabelFontFace, 'minor',
  );
  const labelPx = chartTextFontSizePx(labelDef?.fontSizeHpt, ptToPx)
    ?? Math.max(7, Math.min(13, outerR * 0.075));
  const labelColor = labelDef?.fontColor ? `#${labelDef.fontColor}` : '#ffffff';

  // Draw every non-root node as a ring segment, deepest-last so borders read on
  // top. Iterate breadth-first by depth.
  const byDepth: SunburstNode[][] = Array.from({ length: ringCount }, () => []);
  const pending = [root];
  while (pending.length > 0) {
    const node = pending.pop() as SunburstNode;
    if (node.depth >= 0) byDepth[node.depth].push(node);
    for (let index = node.children.length - 1; index >= 0; index--) {
      pending.push(node.children[index]);
    }
  }

  ctx.save();
  for (let d = 0; d < ringCount; d++) {
    const rInner = innerR + d * ringBand;
    const rOuter = rInner + ringBand;
    for (const node of byDepth[d]) {
      const sweep = node.a1 - node.a0;
      if (sweep <= 1e-4) continue;
      ctx.beginPath();
      ctx.arc(cx, cy, rOuter, node.a0, node.a1);
      ctx.arc(cx, cy, rInner, node.a1, node.a0, true);
      ctx.closePath();
      const nodePaint = branchPaint(node.branchIndex);
      if (nodePaint) {
        ctx.fillStyle = chartExFillStyle(
          ctx,
          nodePaint,
          cx - rOuter,
          cy - rOuter,
          rOuter * 2,
          rOuter * 2,
          branchColor(node.branchIndex),
          shapeRotationDeg,
        );
        ctx.fill();
      }
      if (applyChartExSeriesLineStyle(
        ctx,
        chart,
        chart.chartexDataPointStyle,
        chart.series[0],
        node.branchIndex,
        root.children.length,
        '#ffffff',
        ptToPx,
      )) {
        ctx.stroke();
      }

      const label = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: false, showVal: false, showCatName: false },
        labelOverrides,
      );
      if (!label) continue;
      const labelText = label.text;
      const nodeLabelPx = chartTextFontSizePx(label.fontSizeHpt, ptToPx) ?? labelPx;
      const nodeLabelColor = label.fontColor ? `#${label.fontColor}` : labelColor;
      const nodeLabelFont = label.fontFace
        ? chartFontFamily(chart, label.fontFace, 'minor')
        : labelFont;

      // Excel's sunburst category labels run along the radius (not around the
      // circumference). Center the text at the wedge mid-radius and wrap it to
      // the available ring-band width; additional lines stack tangentially.
      const midA = (node.a0 + node.a1) / 2;
      const midR = (rInner + rOuter) / 2;
      // Radial room the label may occupy (the ring band, minus padding).
      const radialRoom = ringBand - 4;
      // Tangential arc length at the mid radius.
      const arcLen = sweep * midR;
      // Skip labels that plainly cannot fit even one glyph.
      if (!label.manualLayout &&
          radialRoom < nodeLabelPx * 0.9 && arcLen < nodeLabelPx * 0.9) continue;

      const labelX = cx + Math.cos(midA) * midR;
      const labelY = cy + Math.sin(midA) * midR;
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.fontBold ? 'bold ' : ''}${nodeLabelPx}px ${nodeLabelFont}`;
      if (label.manualLayout) {
        drawBoundedDataLabelText(
          ctx,
          labelText,
          { kind: 'point', x: labelX, y: labelY, position: label.position ?? 'ctr' },
          { x: px0, y: py0, w: pw, h: ph },
          nodeLabelPx,
          nodeLabelColor,
          label.manualLayout,
          { x, y, w, h },
          richDataLabelOptions(
            chart, label.richRuns, ptToPx, nodeLabelFont, label.fontBold ?? false, label.textStyle,
          ),
          undefined,
          label.textStyle,
          ptToPx,
          label.labelBox,
          shapeRotationDeg,
        );
        continue;
      }

      ctx.save();
      ctx.translate(labelX, labelY);
      // Orient the text along the radius and flip on the left half so it stays
      // readable instead of becoming upside-down.
      let rot = midA;
      const deg = ((rot * 180) / Math.PI) % 360;
      if (deg > 90 || deg < -90) rot += Math.PI;
      ctx.rotate(rot);
      ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${label.fontBold ? 'bold ' : ''}${nodeLabelPx}px ${nodeLabelFont}`;
      drawBoundedDataLabelText(
        ctx,
        labelText,
        { kind: 'point', x: 0, y: 0, position: label.position ?? 'ctr' },
        { x: -radialRoom / 2, y: -arcLen / 2, w: radialRoom, h: arcLen },
        nodeLabelPx,
        nodeLabelColor,
        undefined,
        { x: -radialRoom / 2, y: -arcLen / 2, w: radialRoom, h: arcLen },
        richDataLabelOptions(
          chart, label.richRuns, ptToPx, nodeLabelFont, label.fontBold ?? false, label.textStyle,
        ),
        undefined,
        label.textStyle,
        ptToPx,
        label.labelBox,
        shapeRotationDeg,
      );
      ctx.restore();
    }
  }
  ctx.restore();

  drawLegendForLayout(
    ctx,
    legendChart,
    leg,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    frame.title.bandH + 2,
    ptToPx,
    legendPaints,
    shapeRotationDeg,
  );
}

// ─── chartEx: treemap (CH15, MS 2014 chartex ext) ───────────────────────────

interface TreemapRect { x: number; y: number; w: number; h: number }
interface TreemapTile { node: SunburstNode; rect: TreemapRect }

/** Standard squarified-treemap layout. Areas are exactly proportional to node
 * layout weights; descending stable order keeps the aspect ratios useful
 * without any document-specific tuning. */
function layoutTreemapTiles(nodes: SunburstNode[], rect: TreemapRect): TreemapTile[] {
  const positive = nodes
    .map((node, index) => ({
      node,
      index,
      value: node.layoutWeight,
    }))
    .filter(entry => entry.value > 0)
    .sort((a, b) => b.value - a.value || a.index - b.index);
  const total = positive.reduce((sum, entry) => sum + entry.value, 0);
  if (total <= 0 || rect.w <= 0 || rect.h <= 0) return [];

  const scale = (rect.w * rect.h) / total;
  const entries = positive.map(entry => ({ ...entry, area: entry.value * scale }));
  const tiles: TreemapTile[] = [];
  let remaining = { ...rect };
  let row: typeof entries = [];
  let rowArea = 0;
  let rowMin = Number.POSITIVE_INFINITY;
  let rowMax = 0;

  const worstRatio = (sum: number, min: number, max: number, shortSide: number): number => {
    if (sum <= 0 || min <= 0 || shortSide <= 0) return Number.POSITIVE_INFINITY;
    const side2 = shortSide * shortSide;
    return Math.max((side2 * max) / (sum * sum), (sum * sum) / (side2 * min));
  };

  const placeRow = (items: typeof entries, area: number): void => {
    if (items.length === 0) return;
    if (remaining.w >= remaining.h) {
      const colW = remaining.h > 0 ? area / remaining.h : 0;
      let y = remaining.y;
      for (let i = 0; i < items.length; i++) {
        const h = i === items.length - 1 ? remaining.y + remaining.h - y : items[i].area / colW;
        tiles.push({ node: items[i].node, rect: { x: remaining.x, y, w: colW, h } });
        y += h;
      }
      remaining = { x: remaining.x + colW, y: remaining.y, w: Math.max(0, remaining.w - colW), h: remaining.h };
    } else {
      const rowH = remaining.w > 0 ? area / remaining.w : 0;
      let x = remaining.x;
      for (let i = 0; i < items.length; i++) {
        const w = i === items.length - 1 ? remaining.x + remaining.w - x : items[i].area / rowH;
        tiles.push({ node: items[i].node, rect: { x, y: remaining.y, w, h: rowH } });
        x += w;
      }
      remaining = { x: remaining.x, y: remaining.y + rowH, w: remaining.w, h: Math.max(0, remaining.h - rowH) };
    }
  };

  let index = 0;
  while (index < entries.length) {
    const next = entries[index];
    const side = Math.min(remaining.w, remaining.h);
    const nextArea = rowArea + next.area;
    const nextMin = Math.min(rowMin, next.area);
    const nextMax = Math.max(rowMax, next.area);
    if (row.length === 0
      || worstRatio(nextArea, nextMin, nextMax, side)
        <= worstRatio(rowArea, rowMin, rowMax, side)) {
      row.push(next);
      rowArea = nextArea;
      rowMin = nextMin;
      rowMax = nextMax;
      index++;
    } else {
      placeRow(row, rowArea);
      row = [];
      rowArea = 0;
      rowMin = Number.POSITIVE_INFINITY;
      rowMax = 0;
    }
  }
  placeRow(row, rowArea);
  return tiles;
}

/** Render chartEx `layoutId="treemap"` as nested, area-proportional rectangles.
 * The chartEx hierarchy is shared with sunburst; `parentLabelLayout="banner"`
 * reserves a header inside each parent, `none` suppresses parent captions, and
 * the other/absent modes overlay the caption without changing tile area. */
function renderTreemapChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  const treemap = chart.chartexTreemap;
  if (!treemap || treemap.rows.length === 0) return;
  if (hierarchyInputTooLarge(treemap.rows)) {
    rejectOversizedCanvasChart(ctx, r, MAX_CANVAS_CHART_POINTS + 1);
    return;
  }
  // A treemap data point remains a distinct tile even when its terminal label
  // repeats. Only parent path components are grouping keys.
  const root = buildSunburstTree(treemap.rows, true);
  if (root.layoutWeight <= 0 || root.children.length === 0) return;
  const series = chart.series[0];
  const legendPaints = root.children.map(node =>
    chartExDataPointPaint(
      chart, node.branchIndex, root.children.length, series?.chartexStyle, series?.color,
    )
  );
  const legendChart: ChartModel = {
    ...chart,
    chartType: 'clusteredBar',
    series: root.children.map(node => {
      const fill = chartExDataPointFill(
        chart, node.branchIndex, root.children.length, series?.chartexStyle,
      );
      return chartExLegendSeries(
        chart,
        node.label,
        series,
        chart.chartexDataPointStyle,
        node.branchIndex,
        root.children.length,
        fill,
        true,
        false,
      );
    }),
  };
  const leg = measuredLegendReserve(ctx, legendChart, r.w, r.h, 0.22, ptToPx);
  const frame = computeChartFrame(chart, r.x, r.y, r.w, r.h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0,
    legendReserve: leg,
    radialGapFrac: 0.015,
    honorPlotAreaManualLayout: true,
  });
  drawChartTitleForLayout(ctx, chart, r.x, r.y, r.w, r.h, r.y + frame.title.topPad, frame.title.fontPx);
  const { px0, py0, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);
  const plotBounds = { x: px0, y: py0, w: pw, h: ph };

  const parentMode = treemap.parentLabelLayout ?? 'overlapping';
  const labelDef = chart.series[0]?.seriesDataLabels;
  const fontFamily = chartFontFamily(
    chart, labelDef?.fontFace ?? chart.dataLabelFontFace, 'minor',
  );
  const labelFontPx = chartTextFontSizePx(labelDef?.fontSizeHpt, ptToPx)
    ?? Math.max(8, Math.min(13, frame.plotRect.ph * 0.025));
  const labelColor = labelDef?.fontColor ? `#${labelDef.fontColor}` : '#ffffff';
  const labelOverrides = new Map(
    (chart.series[0]?.dataLabelOverrides ?? []).map(override => [override.idx, override]),
  );
  // With no direct or linked line recipe, use a one-CSS-pixel separator derived
  // from the chart background. `applyChartExSeriesLineStyle` still resolves
  // direct series formatting and the linked data-point role before this fallback.
  const automaticSeparator = chart.chartBg
    ? (chart.chartBg.startsWith('#') ? chart.chartBg : `#${chart.chartBg}`)
    : '#ffffff';

  const paint = (node: SunburstNode, tile: TreemapRect): void => {
    if (tile.w < 0.5 || tile.h < 0.5) return;
    const base = chartExDataPointFill(
      chart, node.branchIndex, root.children.length, series?.chartexStyle,
    );
    // Every descendant of a top-level branch uses that branch's exact accent.
    // Hierarchy depth does not tint or whiten ChartEx treemap data points.
    const color = `#${base}`;
    const fillPaint = chartExDataPointPaint(
      chart, node.branchIndex, root.children.length, series?.chartexStyle, series?.color,
    );
    const labelOverride = labelOverrides.get(node.labelIndex);
    const nodeLabelColor = labelOverride?.fontColor ? `#${labelOverride.fontColor}` : labelColor;
    const nodeLabelFontPx = chartTextFontSizePx(labelOverride?.fontSizeHpt, ptToPx)
      ?? labelFontPx;
    const nodeLabelBold = labelOverride?.fontBold ?? labelDef?.fontBold ?? false;

    if (node.children.length > 0) {
      // `overlapping` captions belong to the top-level branch entries exposed
      // by the legend; intermediate nodes still partition their descendants.
      // `banner` remains separate because it reserves a band at each parent.
      const parentLabel = resolveChartExLabel(
        chart, series, node.labelIndex, node.label, node.value,
        { visible: parentMode !== 'none', showVal: false, showCatName: true },
        labelOverrides,
        // Excel treats overlapping/banner entries as hierarchy captions. The
        // series value flag applies to leaf data points, not aggregate parents.
        true,
      );
      const showParent = parentLabel != null
        && (parentMode !== 'overlapping' || node.depth === 0);
      const fontPx = nodeLabelFontPx;
      const parentFontFamily = parentLabel?.fontFace
        ? chartFontFamily(chart, parentLabel.fontFace, 'minor')
        : fontFamily;
      const bannerH = parentMode === 'banner' && showParent
        ? Math.min(tile.h * 0.28, fontPx + 7)
        : 0;
      // `overlapping` (MS-ODRAWXML §2.24.3.69 CT_ParentLabelLayout) places the
      // parent caption over its descendant data points. In Excel it does not
      // create an additional painted parent rectangle; doing so here produced
      // a hairline frame around each branch. Banner mode alone reserves and
      // paints a caption band.
      if (bannerH > 0 && fillPaint) {
        ctx.fillStyle = chartExFillStyle(
          ctx,
          fillPaint,
          tile.x,
          tile.y,
          tile.w,
          bannerH,
          color,
          shapeRotationDeg,
        );
        ctx.fillRect(tile.x, tile.y, tile.w, bannerH);
      }
      const content = {
        x: tile.x,
        y: tile.y + bannerH,
        w: tile.w,
        h: Math.max(0, tile.h - bannerH),
      };
      for (const child of layoutTreemapTiles(node.children, content)) paint(child.node, child.rect);

      if (showParent && (parentLabel.manualLayout || (tile.w > fontPx * 2 && tile.h > fontPx + 4))) {
        ctx.font = `${nodeLabelBold ? 'bold ' : ''}${fontPx}px ${parentFontFamily}`;
        const labelRect = bannerH > 0
          ? { x: tile.x, y: tile.y, w: tile.w, h: bannerH }
          : tile;
        const automaticBounds = labelRect;
        // The series-level data-label position describes leaf values. Office
        // keeps overlapping parent captions at the top-left even when leaves
        // are authored at inEnd; only an indexed parent override changes it.
        const parentPosition = labelOverride?.position ?? 'inBase';
        drawBoundedDataLabelText(
          ctx,
          parentLabel.text,
          parentLabel.manualLayout
            ? { kind: 'point', x: tile.x + tile.w / 2, y: tile.y + tile.h / 2, position: parentPosition }
            : { kind: 'box', rect: automaticBounds, position: parentPosition },
          parentLabel.manualLayout ? plotBounds : automaticBounds,
          fontPx,
          nodeLabelColor,
          parentLabel.manualLayout,
          r,
          richDataLabelOptions(
            chart, parentLabel.richRuns, ptToPx, parentFontFamily, nodeLabelBold,
            parentLabel.textStyle,
          ),
          undefined,
          parentLabel.textStyle,
          ptToPx,
          parentLabel.labelBox,
          shapeRotationDeg,
        );
      }
      return;
    }

    if (fillPaint) {
      ctx.fillStyle = chartExFillStyle(
        ctx,
        fillPaint,
        tile.x,
        tile.y,
        tile.w,
        tile.h,
        color,
        shapeRotationDeg,
      );
      ctx.fillRect(tile.x, tile.y, tile.w, tile.h);
    }
    const hasAuthoredOutline = applyChartExSeriesLineStyle(
      ctx,
      chart,
      chart.chartexDataPointStyle,
      chart.series[0],
      node.branchIndex,
      root.children.length,
      automaticSeparator,
      ptToPx,
      { linkedNoStyleFallback: true },
    );
    if (hasAuthoredOutline) {
      // ChartEx outlines are centered on the tile boundary. An inset stroke
      // creates a second visible outer frame that Excel does not paint.
      ctx.strokeRect(tile.x, tile.y, tile.w, tile.h);
    }

    const leafLabel = resolveChartExLabel(
      chart, series, node.labelIndex, node.label, node.value,
      { visible: false, showVal: false, showCatName: false },
      labelOverrides,
    );
    if (!leafLabel) return;
    const fontPx = chartTextFontSizePx(leafLabel.fontSizeHpt, ptToPx)
      ?? nodeLabelFontPx;
    if (!leafLabel.manualLayout && (tile.w <= fontPx * 1.2 || tile.h <= fontPx * 1.2)) return;
    const leafFontFamily = leafLabel.fontFace
      ? chartFontFamily(chart, leafLabel.fontFace, 'minor')
      : fontFamily;
    ctx.font = `${leafLabel.fontBold ? 'bold ' : ''}${fontPx}px ${leafFontFamily}`;
    const automaticBounds = tile;
    // ChartEx uses `outEnd` for treemap value labels but Excel paints that
    // token inside the tile at its lower-left corner. Keep the family-specific
    // semantic mapping here; the shared box resolver's generic outEnd remains
    // right-centred for other chart families.
    const leafPosition = leafLabel.position === 'outEnd'
      ? 'inEnd'
      : (leafLabel.position ?? 'ctr');
    drawBoundedDataLabelText(
      ctx,
      leafLabel.text,
      leafLabel.manualLayout
        ? { kind: 'point', x: tile.x + tile.w / 2, y: tile.y + tile.h / 2, position: leafPosition }
        : { kind: 'box', rect: automaticBounds, position: leafPosition },
      leafLabel.manualLayout ? plotBounds : automaticBounds,
      fontPx,
      leafLabel.fontColor ? `#${leafLabel.fontColor}` : nodeLabelColor,
      leafLabel.manualLayout,
      r,
      richDataLabelOptions(
        chart, leafLabel.richRuns, ptToPx, leafFontFamily, leafLabel.fontBold ?? false,
        leafLabel.textStyle,
      ),
      undefined,
      leafLabel.textStyle,
      ptToPx,
      leafLabel.labelBox,
      shapeRotationDeg,
    );
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(px0, py0, pw, ph);
  ctx.clip();
  for (const tile of layoutTreemapTiles(root.children, { x: px0, y: py0, w: pw, h: ph })) {
    paint(tile.node, tile.rect);
  }
  ctx.restore();

  drawLegendForLayout(
    ctx,
    legendChart,
    leg,
    r.x,
    r.y,
    r.w,
    r.h,
    px0,
    py0,
    pw,
    ph,
    frame.title.bandH + 2,
    ptToPx,
    legendPaints,
    shapeRotationDeg,
  );
}

/** Render one supported ChartEx family, returning false for classic or future
 * chart identifiers that belong to another optional renderer. */
export function renderChartExChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  const hierarchyLabelWork = chartExHierarchyLabelPaintWorkCount(chart);
  if (hierarchyLabelWork != null && hierarchyLabelWork > MAX_CHART_PAINT_COMPONENTS) {
    rejectOversizedCanvasChart(ctx, rect, MAX_CANVAS_CHART_POINTS + 1);
    return true;
  }
  switch (chart.chartType) {
    case 'waterfall':
      renderWaterfallChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'clusteredColumn':
      renderBarChart(
        ctx,
        { ...chart, chartType: 'clusteredBar' },
        rect,
        ptToPx,
        { gapPolicy: 'chartex' },
        shapeRotationDeg,
      );
      return true;
    case 'histogram':
      renderHistogramChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'funnel':
      renderFunnelChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'paretoLine':
      renderParetoLineChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'pareto':
      renderParetoChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'boxWhisker':
      renderBoxWhiskerChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'sunburst':
      renderSunburstChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    case 'treemap':
      renderTreemapChart(ctx, chart, rect, ptToPx, shapeRotationDeg);
      return true;
    default:
      return false;
  }
}

/** Render text shapes from the chart's related Chart Drawing part.
 *
 * `cdr:relSizeAnchor` coordinates are fractions of the full chart space, not
 * the plot area. Paragraph and run properties are authored DrawingML values.
 * `a:bodyPr@wrap="square"` (and the application default when omitted) wraps
 * text inside the authored rectangle; `wrap="none"` keeps a paragraph on one
 * line. Auto-fit remains deliberately separate because it is a different
 * DrawingML choice with different font-scaling semantics.
 */
