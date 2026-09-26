// Classic bar chart family.
import type { ChartModel, ChartRect, ChartSeries, SecondaryValueAxis } from '../../types/chart';

import { classicDataPointFillDecision } from '../classic-data-point-style.js';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import { mergeChartLabelBoxes } from '../label-box.js';

import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markerFillColorFor,
  markerFillPaintFor,
  pointHasMarkerDetail,
  seriesHasMarkerDetail,
} from '../marker-style.js';
import { chartVariesColorsByPoint } from '../legend-entry-plan.js';
import {
  AXIS_OUTER_TEXT_MARGIN_PT,
  computeChartFrame,
  catAxisLabelBandH,
  chartLegendBands,
  chartAxisTitleBands,
  axisTitleFontPx,
  chartTextFontSizePx,
  chartManualOuterAxisInsets,
  categoryTickLabelGapPx,
  axisTitleMargin,
  valueTickLabelGapPx,
} from '../layout.js';
import { planNumericValueAxis, finiteDataExtent } from '../axis-scale.js';
import { axisLineWidthPx, resolveAxisLine, isCrossBetween } from '../axis-style.js';
import { formatCategoryLabel } from '../chart-number-format.js';
import { elideToWidth } from '../text-elide.js';
import {
  categoryLabelAnchorFraction,
  categoryLabelOffsetPx,
  resolveCategoryGapWidthPercent,
  type CategoryGapPolicy,
} from '../category-spacing.js';

import { resolveChartExLabel } from '../chart-ex-label.js';
import {
  chartDataPointStyleRole,
  chartSeriesVariesByPoint,
  rawLinkedChartStyleRole,
} from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectNoFillDecision,
  chartStyleLineDecision,
} from '../style-paint.js';

import {
  chartColor,
  indexPointOverrides,
  pieSliceColor,
  applyClassicStyleLine,
  IndexedLinePoint,
  paintClassicVaryingLineSegments,
  chartFontFamily,
  chartFontCss,
  drawAxisTitles,
  chartHasDataTable,
  chartDataTableBaseHeight,
  chartDataTableHeaderWidth,
  measureChartDataTable,
  drawChartDataTable,
  createDataLabelLegendKeyResolver,
  measuredLegendReserve,
  drawLegendForLayout,
  manualTopLegendPlotInset,
  drawAxisTick,
  strokeAxisSegment,
  axisTickLengthPx,
  strokeValueGridlineH,
  valGridStroke,
  valMinorGridStroke,
  drawCatMajorGridlines,
  catGridStroke,
  catGridlineFractions,
  catAxisReversed,
  drawValMajorGridlines,
  formatPrimaryValueAxisTick,
  formatAxisTickWithUnits,
  planValueAxis,
  drawSeriesTrendlines,
  trendlineLegendSeries,
  axisLabelPx,
  wrapMeasuredText,
  numericCategoryMetricTolerance,
  catLabelsVisible,
  catLabelRotationRad,
  drawRotatedCatLabel,
  chartDateAxisPlan,
  forEachErrorBarEndpoint,
  computeSecondaryAxis,
  drawSecondaryValueGridlines,
  drawSecondaryValueAxis,
  drawSecondaryCategoryAxis,
  measuredCartesianTitleBand,
  drawChartTitleForLayout,
  chartCategories,
  dataLabelWithinAxisMaximum,
  drawBarDataLabel,
  applyDecorationLineStyle,
  chartStyleRoleLine,
  chartStyleRoleErrorBar,
  drawLineGroupDecorations,
  categoryAxisCrossingValue,
  scatterXValue,
  drawScatterSeriesLayer,
  drawChartMarker,
  seriesHasResolvedMarkerDetail,
  customRichDataLabelOptions,
  clamp,
  appendCurve,
  drawBarErrorBars,
  chartExSeriesFormatIndex,
  chartExDataPointFill,
  chartExDataPointPaint,
  paintClassicDataPointPath,
  paintClassicDataPointRect,
  applyChartExSeriesLineStyle,
  chartExLegendSeries,
} from '../shared/classic.js';

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart — vertical columns + horizontal bars, clustered + stacked +
// percentStacked. Also handles mixed bar+line series (seriesType per series).
// ═══════════════════════════════════════════════════════════════════════════

/**
 * MS ChartEx does not expose a value-axis label-offset property. Office vector
 * output from histogram, box-and-whisker, and Pareto charts (10 pt and 12 pt
 * labels, with and without cross ticks) places the visible label edge about
 * 6.2–7.5 pt from the axis centreline. Keep this compatibility fallback
 * ChartEx-only; classic axes retain their existing font-relative contract.
 */
export function chartExValueTickLabelOffsetPx(ptToPx: number): number {
  return 7 * ptToPx;
}

export function renderBarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  options: {
    gapPolicy?: CategoryGapPolicy;
    semanticLineNoStyleFallback?: boolean;
  } = {},
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const isH = chart.chartType === 'clusteredBarH' || chart.chartType === 'stackedBarH' || chart.chartType === 'stackedBarHPct';
  const stacked = chart.chartType.startsWith('stacked');
  const pct = chart.chartType === 'stackedBarPct' || chart.chartType === 'stackedBarHPct';

  const allBarSeries = chart.series.filter(s => s.seriesType !== 'line' && s.seriesType !== 'scatter' && s.seriesType !== 'area');
  const groupIsHorizontal = (series: ChartSeries): boolean => series.barGroupDirection != null
    ? series.barGroupDirection === 'bar'
    : isH;
  // The first bar group owns the visible axis orientation, but Excel retains
  // each later CT_BarChart group's own barDir. Consequently a schema-valid
  // shared-axis plot may contain vertical columns over horizontal bars. Keep
  // every group in this one data pass; category/value geometry is selected per
  // series below while axes remain owned by the first compatibility family.
  const barSeries = allBarSeries;
  const fallbackBarGroupKey = (series: ChartSeries): string =>
    series.useSecondaryAxis === true ? 'secondary-default' : 'primary-default';
  const barGroupKey = (series: ChartSeries): string => series.barGroupIndex != null
    ? `group-${series.barGroupIndex}`
    : fallbackBarGroupKey(series);
  const groupGrouping = (series: ChartSeries): string => series.barGroupGrouping
    ?? (pct ? 'percentStacked' : stacked ? 'stacked' : 'clustered');
  const groupIsStacked = (series: ChartSeries): boolean => {
    const grouping = groupGrouping(series);
    return grouping === 'stacked' || grouping === 'percentStacked';
  };
  const groupIsPercent = (series: ChartSeries): boolean =>
    groupGrouping(series) === 'percentStacked';
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');
  const areaSeries = chart.series.filter(s => s.seriesType === 'area');
  const scatterSeries = chart.series.filter(s => s.seriesType === 'scatter');
  const sourceSeriesIndices = new Map(chart.series.map((series, index) => [series, index]));
  const plotGroupBySeries = new Map<ChartSeries, NonNullable<ChartModel['plotGroups']>[number]>();
  for (const group of chart.plotGroups ?? []) {
    for (let index = group.seriesStart; index < group.seriesStart + group.seriesCount; index++) {
      const series = chart.series[index];
      if (series) plotGroupBySeries.set(series, group);
    }
  }
  const axisPercentState = new Map<string, { count: number; percentCount: number }>();
  for (const group of chart.plotGroups ?? []) {
    if (group.seriesCount === 0) continue;
    const state = axisPercentState.get(group.valueAxis) ?? { count: 0, percentCount: 0 };
    state.count++;
    if (group.grouping === 'percentStacked') state.percentCount++;
    axisPercentState.set(group.valueAxis, state);
  }
  const axisUsesPercentSpace = (group: NonNullable<ChartModel['plotGroups']>[number]): boolean => {
    const state = axisPercentState.get(group.valueAxis);
    return state != null && state.count === state.percentCount;
  };
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);

  // Combo charts (bar + line) may bind the line series to a SECONDARY value
  // axis drawn on the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with
  // axPos="r" / `<c:crosses val="max">`). `sec` is non-null only when both the
  // axis is declared AND at least one line series opts into it; horizontal bar
  // charts never carry one.
  const hasSecondarySeries = chart.series.some(series => series.useSecondaryAxis === true);
  const sec = !isH && chart.secondaryValAxis && hasSecondarySeries
    ? chart.secondaryValAxis
    : null;
  const secondaryBarSeries = sec
    ? barSeries.filter(series => series.useSecondaryAxis === true)
    : [];
  const primaryBarSeries = sec
    ? barSeries.filter(series => series.useSecondaryAxis !== true)
    : barSeries;
  const secondaryCat = secondaryBarSeries.length > 0 ? chart.secondaryCatAxis : null;
  const secondaryCategories = secondaryBarSeries[0]?.categories?.length
    ? secondaryBarSeries[0].categories
    : chart.categories;

  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;
  const barGroups = new Map<string, ChartSeries[]>();
  for (const barSeriesEntry of barSeries) {
    const key = barGroupKey(barSeriesEntry);
    const group = barGroups.get(key);
    if (group) group.push(barSeriesEntry);
    else barGroups.set(key, [barSeriesEntry]);
  }
  const barGroupFor = (series: ChartSeries): ChartSeries[] =>
    barGroups.get(barGroupKey(series)) ?? [series];
  const percentDenominators = new Map<string, number[]>();
  for (const [key, members] of barGroups) {
    const totals = new Array<number>(n).fill(0);
    for (const member of members) {
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        totals[categoryIndex] += Math.abs(member.values[categoryIndex] ?? 0);
      }
    }
    percentDenominators.set(key, totals);
  }
  const percentDenominator = (series: ChartSeries, categoryIndex: number): number =>
    percentDenominators.get(barGroupKey(series))?.[categoryIndex] || 1;
  const percentGroupMultiplier = (series: ChartSeries): number => {
    const group = plotGroupBySeries.get(series);
    return group == null || axisUsesPercentSpace(group) ? 100 : 1;
  };
  const percentFactor = (series: ChartSeries, categoryIndex: number): number => {
    if (!groupIsPercent(series)) return 1;
    return percentGroupMultiplier(series) / percentDenominator(series, categoryIndex);
  };
  const overlayPlottedValues = new Map<ChartSeries, number[]>();
  const overlayBaseValues = new Map<ChartSeries, number[]>();
  for (const series of [...lineSeries, ...areaSeries]) {
    overlayPlottedValues.set(series, Array.from(
      { length: n }, (_, index) => series.values[index] ?? 0,
    ));
    overlayBaseValues.set(series, new Array<number>(n).fill(0));
  }
  for (const group of chart.plotGroups ?? []) {
    if (group.kind !== 'line' && group.kind !== 'area') continue;
    const members = chart.series.slice(group.seriesStart, group.seriesStart + group.seriesCount);
    const stackedGroup = group.grouping === 'stacked' || group.grouping === 'percentStacked';
    const percentGroup = group.grouping === 'percentStacked';
    const multiplier = percentGroup && axisUsesPercentSpace(group) ? 100 : 1;
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      const denominator = percentGroup
        ? members.reduce((sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0) || 1
        : 1;
      let running = 0;
      for (const series of members) {
        const raw = series.values[categoryIndex] ?? 0;
        const contribution = percentGroup ? raw / denominator * multiplier : raw;
        const baseValues = overlayBaseValues.get(series);
        const plottedValues = overlayPlottedValues.get(series);
        if (baseValues == null || plottedValues == null) continue;
        baseValues[categoryIndex] = stackedGroup ? running : 0;
        running = stackedGroup ? running + contribution : contribution;
        plottedValues[categoryIndex] = running;
      }
    }
  }
  const overlayValue = (series: ChartSeries, categoryIndex: number): number =>
    overlayPlottedValues.get(series)?.[categoryIndex] ?? series.values[categoryIndex] ?? 0;
  const overlayBase = (series: ChartSeries, categoryIndex: number): number =>
    overlayBaseValues.get(series)?.[categoryIndex] ?? 0;
  const effectiveBarErrorBars = (
    series: ChartSeries,
  ): NonNullable<ChartSeries['errBars']> => {
    if (!groupIsPercent(series)) return series.errBars ?? [];
    return (series.errBars ?? []).map(errorBars => ({
      ...errorBars,
      plus: errorBars.plus.map((value, index) => value == null
        ? value
        : value * percentFactor(series, index)),
      minus: errorBars.minus.map((value, index) => value == null
        ? value
        : value * percentFactor(series, index)),
    }));
  };

  // §21.2.2.227 varyColors belongs to each chart-group child. A combo chart
  // can therefore contain a lone varying bar group alongside unrelated line
  // or area groups; total plot-series count is not the ownership domain.
  // Models produced before plot-group metadata existed retain the chart-level
  // fallback for backward compatibility.
  const barVariesByPoint = (series: ChartSeries): boolean => {
    const group = plotGroupBySeries.get(series);
    if (group?.kind === 'bar' || group?.kind === 'bar3D') {
      return group.seriesCount === 1 && group.varyColors === true;
    }
    return chartVariesColorsByPoint(chart);
  };
  const pointOverrides = barSeries.map(series =>
    new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point])),
  );
  const labelOverrides = barSeries.map(series =>
    new Map((series.dataLabelOverrides ?? []).map(label => [label.idx, label])),
  );
  const barStyleIndices = barSeries.map((series, index) =>
    chartExSeriesFormatIndex(series, index)
  );
  // The shared classic-style adapter intentionally exposes effective roles in
  // the historical `chartex*Style` fields. Do not use those aliases as a file-
  // format discriminator: only a true ChartEx model lacks the classic numeric
  // role table. Otherwise a classic bar+line combo would drop every non-bar
  // legend entry by entering the ChartEx synthetic-series path.
  const isChartExColumn = chart.classicChartStyleRoles == null
    && (chart.chartexDataPointStyle != null || chart.chartexColorPalette != null);
  const styledBarLegendSeries = new Map<ChartSeries, ChartSeries>();
  if (isChartExColumn) {
    barSeries.forEach((series, index) => {
      const styleIndex = barStyleIndices[index];
      const fill = series.color
        ?? chartExDataPointFill(chart, styleIndex, barSeries.length, series.chartexStyle);
      styledBarLegendSeries.set(series, chartExLegendSeries(
        chart,
        series.name,
        series,
        chart.chartexDataPointStyle,
        styleIndex,
        barSeries.length,
        fill,
      ));
    });
  }
  const varyingBarSeriesIndex = barSeries.findIndex(barVariesByPoint);
  const varyingBarSeries = varyingBarSeriesIndex >= 0
    ? barSeries[varyingBarSeriesIndex]
    : undefined;
  const legendChart: ChartModel = {
    ...chart,
    series: (isChartExColumn ? barSeries : chart.series).map(series =>
      styledBarLegendSeries.get(series) ?? series
    ),
  };

  // Honor the parser-resolved title font size when present; otherwise use the
  // shared fixed fallback. Reserve the title band from the actual drawn size
  // so the plot shrinks to avoid overlap.
  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  let titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  let titleFontPx = titleBand.fontPx;
  let titleTopPad = titleBand.topPad;
  let titleH = titleBand.bandH;
  // Axis-label font (XML @sz when set) — sizes the bottom tick-label band the
  // same way the line/area families do.
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const categoryLevels = !isH
    && !chartHasDataTable(chart)
    && chart.catAxisNoMultiLevelLabels !== true
    && (chart.categoryLevels?.length ?? 0) > 1
    ? chart.categoryLevels!
    : null;
  const multiLevelCategoryBandH = categoryLevels
    ? (categoryLevels.length - 1) * (catAxFontPx + 4)
    : 0;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);
  const leg = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  // Axis-title bands sized from the *actual* title font (honoring XML @sz)
  // plus a small gap, so big titles get a wide enough gutter
  // and never collide with the tick labels.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  // Horizontal bars swap semantic axes: category title belongs in the left
  // band, while the horizontal value-axis title belongs in the bottom band.
  const catTitleH = isH
    ? (chart.valAxisTitle ? valTitlePx + axisTitleMargin(h) + 4 : 0)
    : axBands.catBandH;
  const valTitleW = isH
    ? (chart.catAxisTitle ? catTitlePx + axisTitleMargin(w) + 4 : 0)
    : axBands.valBandW;
  // Value-axis scales are computed up-front (before `pad`) so the side gutters
  // can be sized to the actual tick-label widths instead of a fixed fraction of
  // the chart width — short numeric labels otherwise leave a big empty gap
  // between the axis title and the labels (PowerPoint sizes the gutter to fit
  // the labels). The scales depend only on the series data, not on `pad`.
  // Vertical pads first (independent of the side gutters) so the plot height —
  // and the value-axis length — are known before the scale + label measuring.
  // The value-axis LENGTH drives the auto major unit (Excel targets a roughly
  // constant gridline spacing, so a longer axis gets finer ticks).
  // Top: title band + a small breathing gap above the topmost gridline.
  // Bottom: PowerPoint's tick-label band (gap + line-height + margin) sized to
  // the label font — the category labels for columns, the value-axis labels for
  // horizontal bars (both a single line of text). A hidden bottom axis keeps a
  // minimal gap. Matches the line/area reserve so the four families agree.
  const secondaryCatFontPx = chartTextFontSizePx(secondaryCat?.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  const secondaryCatLabelBandH = secondaryCat && !secondaryCat.hidden
    && secondaryCat.tickLabelPos !== 'none'
    ? secondaryCatFontPx + categoryLabelOffsetPx(
      categoryTickLabelGapPx(secondaryCatFontPx),
      secondaryCat.labelOffsetPercent,
    ) + 2
    : 0;
  const secondaryCatTitleBandH = secondaryCat?.title
    ? axisTitleFontPx(secondaryCat.titleFontSizeHpt, ptToPx) + 6
    : 0;
  let padT = titleH + legTopH + valAxLabelFontPx / 2 + 2
    + secondaryCatLabelBandH + secondaryCatTitleBandH;
  const padB = isH
    ? (chart.valAxisHidden ? h * 0.02 : catAxisLabelBandH(valAxLabelFontPx))
      + dataTableBaseH + catTitleH + legBottomH
    : (hasDataTable ? 0 : catAxisLabelBandH(
      catAxFontPx,
      chart.catAxisLabelOffsetPercent,
      chart.cartesianAutoLayoutProfile,
    ))
      + multiLevelCategoryBandH + dataTableBaseH + catTitleH + legBottomH;
  const phEst = h - padT - padB;
  let horizontalCategoryLabelBandW = 0;
  if (isH && !chart.catAxisHidden && catLabelsVisible(chart)) {
    // The paint path derives an unauthored tick font from the category slot,
    // not the overall chart height. Use the same resolver here so tall charts
    // do not reserve a gutter for a much larger font than they actually draw.
    const measuredHorizontalCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, (phEst / n) * 0.5));
    ctx.save();
    ctx.font = chartFontCss(
      measuredHorizontalCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    for (const category of cats) {
      horizontalCategoryLabelBandW = Math.max(
        horizontalCategoryLabelBandW,
        ctx.measureText(formatCategoryLabel(category, chart.catAxisFormatCode, chart.date1904)).width,
      );
    }
    ctx.restore();
    horizontalCategoryLabelBandW += categoryLabelOffsetPx(
      chart.catAxisFontSizeHpt != null
        ? valueTickLabelGapPx(measuredHorizontalCatTickFontPx)
        : 4,
      chart.catAxisLabelOffsetPercent,
    ) + AXIS_OUTER_TEXT_MARGIN_PT * ptToPx;
  }
  // Auto layout keeps at least half of the frame available to the data plot;
  // labels beyond that measured budget are elided at paint time. Authored outer
  // layouts still use the full measured band in `manualOuterInsets` below.
  const automaticHorizontalCategoryLabelBandW = Math.min(
    horizontalCategoryLabelBandW,
    Math.max(0, w / 2 - valTitleW - legLeftW),
  );
  // Horizontal bars run the value axis along the (wide) bottom, so its length is
  // the plot WIDTH. Estimate it from the same measured category-label band that
  // the final frame uses so automatic tick density agrees with painted geometry.
  const pwEst = isH
    ? w - ((chart.catAxisHidden ? w * 0.03 : automaticHorizontalCategoryLabelBandW) + valTitleW + legLeftW) - (legRightW + w * 0.03)
    : 0;
  // A deleted value axis has no ticks whose density needs adapting to the
  // available screen length. Office falls back to its default automatic scale
  // target in that case. Feeding the plot length into the visible-tick planner
  // over-refines the major unit and stretches bars relative to slide-authored
  // overlay labels.
  const valAxisLenPt = chart.valAxisHidden ? undefined : (isH ? pwEst : phEst) / ptToPx;

  const plottedBarValue = (seriesIndex: number, categoryIndex: number): number => {
    const owner = barSeries[seriesIndex];
    const raw = owner?.values[categoryIndex] ?? 0;
    if (!owner || !groupIsStacked(owner)) return raw;
    const group = barGroupFor(owner);
    const percent = groupIsPercent(owner);
    let denominator = 1;
    if (percent) {
      denominator = group.reduce(
        (sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0,
      ) || 1;
    }
    const percentMultiplier = percent ? percentFactor(owner, categoryIndex) * denominator : 1;
    const value = percent ? raw / denominator * percentMultiplier : raw;
    let cumulative = 0;
    const ownerIndex = group.indexOf(owner);
    for (let index = 0; index <= ownerIndex; index++) {
      const candidateRaw = group[index]?.values[categoryIndex] ?? 0;
      const candidate = percent ? candidateRaw / denominator * percentMultiplier : candidateRaw;
      if ((value < 0) === (candidate < 0)) cumulative += candidate;
    }
    return cumulative;
  };

  // Value-axis extent. Bars extend from the zero line (the category-axis
  // crossing) toward each value, so the axis must span both the positive and
  // negative reach of the data (ECMA-376 §21.2.2.16 barChart). Negative values
  // pull the axis minimum below 0; positive values push the maximum above it.
  // Clustered charts take the raw extremes; stacked charts accumulate positive
  // and negative contributions on separate sides of the zero line (Excel stacks
  // opposite signs opposite ways), so `dataMax`/`dataMin` come from each
  // category's positive-sum and negative-sum.
  const primaryPlotGroups = (chart.plotGroups ?? []).filter(group =>
    group.seriesCount > 0 && group.valueAxis !== 'secondary'
  );
  const primaryPercentAxis = primaryPlotGroups.length > 0
    ? primaryPlotGroups.every(group => group.grouping === 'percentStacked')
    : primaryBarSeries.some(groupIsPercent);
  let dataMax = 0;
  let dataMin = 0;
  for (let ci = 0; ci < n; ci++) {
    const primaryGroups = new Map<string, ChartSeries[]>();
    for (const series of primaryBarSeries) {
      const key = barGroupKey(series);
      const members = primaryGroups.get(key);
      if (members) members.push(series); else primaryGroups.set(key, [series]);
    }
    for (const members of primaryGroups.values()) {
      const owner = members[0];
      const isStackedGroup = groupIsStacked(owner);
      const isPercentGroup = groupIsPercent(owner);
      const denominator = isPercentGroup
        ? members.reduce((sum, series) => sum + Math.abs(series.values[ci] ?? 0), 0) || 1
        : 1;
      let posSum = 0;
      let negSum = 0;
      for (const series of members) {
        const raw = series.values[ci] ?? 0;
        const value = isPercentGroup
          ? raw / denominator * percentGroupMultiplier(series)
          : raw;
        if (isStackedGroup) {
          if (value >= 0) posSum += value; else negSum += value;
        } else {
          dataMax = Math.max(dataMax, value);
          dataMin = Math.min(dataMin, value);
        }
      }
      if (isStackedGroup) {
        dataMax = Math.max(dataMax, posSum);
        dataMin = Math.min(dataMin, negSum);
      }
    }
  }
  // Combo line series plotted on the PRIMARY value axis (a bar+line chart whose
  // line rides the same `<c:valAx>` as the bars — no secondary axis, or one the
  // line doesn't opt into) must expand the primary axis extent just like the
  // bars do. Excel scales a shared value axis to encompass EVERY series on it,
  // regardless of chart type; a tall line point can exceed the bar stack, so
  // sizing to the bars alone would clip the line. The line is an unstacked overlay, so each raw datum widens the range
  // directly. Secondary-axis line series are excluded (they own an independent
  // scale, mirrored by the `yOf` split below). `sec` matches the draw-time gate.
  for (const s of [...lineSeries, ...areaSeries]) {
    if (sec && s.useSecondaryAxis === true) continue;
    for (let ci = 0; ci < n; ci++) {
      if (s.values[ci] == null) continue;
      const value = overlayValue(s, ci);
      dataMax = Math.max(dataMax, value);
      dataMin = Math.min(dataMin, value);
    }
  }

  // Error bars are part of the plotted value geometry. Their endpoints must
  // participate in automatic scaling; otherwise a valid endpoint can extend
  // beyond an axis planned only from the underlying series values.
  for (const series of primaryBarSeries) {
    const seriesIndex = barSeries.indexOf(series);
    for (const errorBars of effectiveBarErrorBars(series)) {
      forEachErrorBarEndpoint(
        { ...series, errBars: [errorBars] },
        isH ? 'x' : 'y',
        categoryIndex => series.values[categoryIndex] == null
          ? null
          : plottedBarValue(seriesIndex, categoryIndex),
        value => {
          dataMax = Math.max(dataMax, value);
          dataMin = Math.min(dataMin, value);
        },
      );
    }
  }
  for (const series of [...lineSeries, ...areaSeries]) {
    if (sec && series.useSecondaryAxis === true) continue;
    forEachErrorBarEndpoint(
      series,
      'y',
      index => series.values[index] ?? null,
      value => {
        const effective = value;
        dataMax = Math.max(dataMax, effective);
        dataMin = Math.min(dataMin, effective);
      },
    );
  }
  if (primaryPercentAxis) {
    if (primaryBarSeries.some(series => series.values.some(value => value != null && value > 0))) {
      dataMax = Math.max(dataMax, 100);
    }
    if (primaryBarSeries.some(series => series.values.some(value => value != null && value < 0))) {
      dataMin = Math.min(dataMin, -100);
    }
  }
  if (chart.valMax != null) {
    dataMax = primaryPercentAxis ? chart.valMax * 100 : chart.valMax;
  }
  if (chart.valMin != null) {
    dataMin = primaryPercentAxis ? chart.valMin * 100 : chart.valMin;
  }
  if (dataMax === 0 && dataMin === 0) dataMax = 1;
  // `planValueAxis` folds in the CH6 major unit / logBase / orientation; with
  // none set it is byte-identical to `valueAxisScale` + a linear map.
  const plan = planValueAxis(
    chart,
    dataMin,
    dataMax,
    valAxisLenPt,
    primaryPercentAxis,
    isH ? 'horizontal' : 'vertical',
  );
  const { step } = plan;

  // Secondary value-axis scale (combo charts). INDEPENDENT of the primary: its
  // own "nice" major unit / gridline count. Its axis is the vertical right edge,
  // so its length is the plot height. Explicit `<c:scaling>` wins. Computed by
  // the shared `computeSecondaryAxis` helper (same math the line/area families
  // reuse); the fallback keeps the no-secondary path unchanged.
  const renderedBarSeries = new Set(barSeries);
  const allBarSeriesSet = new Set(allBarSeries);
  const secondaryPlotGroups = (chart.plotGroups ?? []).filter(group =>
    group.seriesCount > 0 && group.valueAxis === 'secondary'
  );
  const secondaryPercentAxis = secondaryPlotGroups.length > 0
    ? secondaryPlotGroups.every(group => group.grouping === 'percentStacked')
    : secondaryBarSeries.some(groupIsPercent);
  const secondaryScaleSeries = chart.series
    .filter(series => !allBarSeriesSet.has(series) || renderedBarSeries.has(series))
    .map(series => {
      if (series.useSecondaryAxis !== true) return series;
      if (renderedBarSeries.has(series)) {
        const seriesIndex = barSeries.indexOf(series);
        return {
          ...series,
          values: series.values.map((value, index) => value == null
            ? value
            : plottedBarValue(seriesIndex, index)),
          errBars: effectiveBarErrorBars(series),
        };
      }
      if (!secondaryPercentAxis) return series;
      return {
        ...series,
        values: series.values.map(value => value == null ? value : value * 100),
        errBars: (series.errBars ?? []).map(errorBars => ({
          ...errorBars,
          plus: errorBars.plus.map(value => value == null ? value : value * 100),
          minus: errorBars.minus.map(value => value == null ? value : value * 100),
        })),
      };
    });
  if (secondaryPercentAxis && secondaryBarSeries[0]) {
    const percentBounds: number[] = [];
    if (secondaryBarSeries.some(series =>
      series.values.some(value => value != null && value > 0))) percentBounds.push(100);
    if (secondaryBarSeries.some(series =>
      series.values.some(value => value != null && value < 0))) percentBounds.push(-100);
    secondaryScaleSeries.push({
      ...secondaryBarSeries[0],
      values: percentBounds,
      errBars: [],
    });
  }
  const secScale = computeSecondaryAxis(
    sec,
    secondaryScaleSeries,
    phEst / ptToPx,
    isH ? 'x' : 'y',
    secondaryPercentAxis,
    secondaryBarSeries.length > 0,
  );

  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const measuredValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, phEst / 20));
  const prevFont = ctx.font;
  // Primary value-axis label band (column charts only; horizontal bars keep a
  // wider left band for the category labels).
  let valLabelTextW = 0;
  let valLabelBandW = 0;
  if (!isH && !chart.valAxisHidden) {
    // Measure with the same face the value-axis ticks draw with (below), so the
    // reserved gutter width matches the painted labels when a real face is set.
    ctx.font = chartFontCss(
      measuredValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    let wmax = 0;
    for (const val of plan.majorLines) {
      const label = formatPrimaryValueAxisTick(chart, val, primaryPercentAxis);
      wmax = Math.max(wmax, ctx.measureText(label).width);
    }
    valLabelTextW = wmax;
    valLabelBandW = valLabelTextW + 16; // ~12px tick+gap to the axis + ~4px to the title
  }
  // Secondary value-axis label band (right edge). Measure with the SAME font
  // and number format the axis is drawn with (`secFontPx` / `sec.formatCode`),
  // otherwise a `%`/thousands format or an explicit font size makes the
  // reserved gutter disagree with the painted labels.
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && !sec.hidden) {
    ctx.font = `${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    let wmax = 0;
    for (const value of secScale?.majorLines ?? []) {
      wmax = Math.max(wmax, ctx.measureText(formatAxisTickWithUnits(
        secondaryPercentAxis ? value / 100 : value,
        sec.formatCode ?? null,
        chart.date1904,
        sec.displayUnits,
      )).width);
    }
    secLabelBandW = wmax + 18;
  }
  ctx.font = prevFont;
  const secTitleBandW = sec && sec.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.03 + secLabelBandW + secTitleBandW,
    b: padB,
    // Column charts: title band + measured label band, tight to the axis.
    // Horizontal bars: keep the wider left band for the category labels
    // (`c:catAx/c:delete val="1"` → no category labels, so tighten).
    l: isH
      ? legLeftW + Math.max(
        (chart.catAxisHidden ? w * 0.03 : automaticHorizontalCategoryLabelBandW) + valTitleW,
        dataTableHeaderW,
      )
      : legLeftW + Math.max(valTitleW + valLabelBandW, dataTableHeaderW),
  };
  pad.t = manualTopLegendPlotInset(
    chart, leg, x, y, w, h, titleH, pad.t,
  );

  // `layoutTarget="outer"` includes tick labels and axis titles, but not the
  // chart title or legend. Convert only those measured axis bands to the inner
  // bar/column plot rectangle; an explicit `inner` target ignores the insets in
  // `computeChartFrame`.
  const manualOuterInsets = isH
    ? {
        t: 0,
        r: chart.valAxisHidden ? 0 : measuredValTickFontPx / 2,
        b: chart.valAxisHidden ? 0 : measuredValTickFontPx + catTitleH,
        l: chart.catAxisHidden ? 0 : horizontalCategoryLabelBandW + valTitleW,
      }
    : chartManualOuterAxisInsets({
        valAxisHidden: chart.valAxisHidden,
        catAxisHidden: chart.catAxisHidden,
        valLabelWidth: valLabelTextW,
        valLabelFontPx: measuredValTickFontPx,
        catLabelFontPx: catAxFontPx,
        valLabelGapPx: chart.valAxisFontSizeHpt != null
          ? valueTickLabelGapPx(measuredValTickFontPx)
          : 12,
        catLabelGapPx: chart.catAxisFontSizeHpt != null
          ? categoryLabelOffsetPx(
            categoryTickLabelGapPx(catAxFontPx),
            chart.catAxisLabelOffsetPercent,
          )
          : categoryLabelOffsetPx(3, chart.catAxisLabelOffsetPercent),
        outerTextMarginPx: AXIS_OUTER_TEXT_MARGIN_PT * ptToPx,
        valTitleBandW: valTitleW,
        catTitleBandH: catTitleH,
        secondaryBandW: secLabelBandW + secTitleBandW,
      });

  // Plot-area placement: honor `<c:plotArea><c:layout><c:manualLayout>` when
  // present (ECMA-376 §21.2.2.32). Templates use this to keep bars from
  // overflowing into side annotations; an explicit inner rectangle keeps the
  // data region separate from adjacent authored content.
  // `layoutTarget="inner"` (default) means the rectangle covers the inner
  // data region; "outer" includes axes/labels. We treat both identically
  // because the inner padding stays the same either way. computeChartFrame
  // applies the pad → plot rect and the manual-layout override.
  let frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    // The cartesian title band is already folded into `pad.t`; pass it so
    // `frame.title` (if read) matches the reserved band instead of a stale frac.
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  const plotWidthTitleBand = measuredCartesianTitleBand(
    ctx,
    chart,
    frame.plotRect.pw,
    h,
    ptToPx,
  );
  if (Math.abs(plotWidthTitleBand.bandH - titleBand.bandH) > 0.01) {
    titleBand = plotWidthTitleBand;
    titleFontPx = titleBand.fontPx;
    titleTopPad = titleBand.topPad;
    titleH = titleBand.bandH;
    padT = titleH + legTopH + valAxLabelFontPx / 2 + 2
      + secondaryCatLabelBandH + secondaryCatTitleBandH;
    pad.t = manualTopLegendPlotInset(
      chart, leg, x, y, w, h, titleH, padT,
    );
    frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
      titleBand,
      legendSideReserveFrac: 0.22,
      legendReserve: leg,
      pad,
      honorPlotAreaManualLayout: true,
      manualOuterInsets,
    });
  }
  const { px0, py0, pw } = frame.plotRect;
  let { ph } = frame.plotRect;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleTopPad, titleFontPx,
  );
  if (pw <= 0 || ph <= 0) return;

  // Horizontal bar categories run from bottom to top by default in the
  // existing category-slot contract; preserve that orientation while using
  // the same calendar coordinate plan as every other date-axis family.
  const dateAxisPlan = chartDateAxisPlan(
    chart,
    cats,
    isH ? !catAxisReversed(chart) : catAxisReversed(chart),
  );

  // Horizontal DrawingML category text (`wrap="square"`) wraps within its
  // category slot. Measure the complete strings with the actual tick font and
  // preserve every word instead of replacing most labels with an ellipsis.
  // An authored inner plot rectangle already leaves its own label band; for an
  // automatic/outer layout, move the plot bottom up by the additional wrapped
  // lines so they remain inside the chart frame.
  const catLabelRotation = catLabelRotationRad(chart);
  const wrappedColumnCategories: string[][] = [];
  let wrappedCategoryExtraH = 0;
  if (!hasDataTable && !isH && !dateAxisPlan && !chart.catAxisHidden && catLabelsVisible(chart) && catLabelRotation === 0) {
    const slotW = pw / n;
    const wrapFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, slotW * 0.5));
    ctx.save();
    ctx.font = chartFontCss(
      wrapFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    for (const category of cats) {
      const formattedCategory = formatCategoryLabel(
        category,
        chart.catAxisFormatCode,
        chart.date1904,
      );
      wrappedColumnCategories.push(wrapMeasuredText(
        ctx,
        formattedCategory,
        Math.max(1, slotW),
        numericCategoryMetricTolerance(formattedCategory, wrapFontPx),
      ));
    }
    ctx.restore();
    const maxLines = Math.max(1, ...wrappedColumnCategories.map(lines => lines.length));
    const manualInner = chart.plotAreaManualLayout?.layoutTarget === 'inner' &&
      chart.plotAreaManualLayout.w != null && chart.plotAreaManualLayout.h != null;
    if (!manualInner && maxLines > 1) {
      wrappedCategoryExtraH = (maxLines - 1) * (wrapFontPx + 2);
      ph = Math.max(1, ph - wrappedCategoryExtraH);
    }
  }

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // `axMax`/`step` (primary) and `sMin`/`sMax`/`sStep` (secondary) were computed
  // above the `pad` block so the gutters could be sized to the labels. The
  // line-mapping helpers need the now-final plot rect, so they live here. Line
  // series bound to the secondary axis map through `toYSecondary`; everything
  // else uses the primary `axMax`.
  // Primary value → pixel. `axRange`/`axMin` generalize the old `v / axMax`
  // mapping so the zero line sits wherever the axis crosses it (mid-plot when
  // the data straddles zero); positive-only data keeps `axMin === 0`, so the
  // mapping is unchanged. `valX`/`valY` give the on-axis pixel for a value on
  // the value axis (X for horizontal bars, Y for columns).
  const valY = (v: number): number => py0 + ph - plan.frac(v) * ph;
  const valX = (v: number): number => px0 + plan.frac(v) * pw;
  const zeroY = valY(0); // column zero line
  const zeroX = valX(0); // horizontal-bar zero line
  const toYPrimaryLine = (value: number): number => valY(value);
  // Secondary line series map through the shared scale's factory (identical to
  // the old inline `py0 + ph - ((v - sMin) / sRange) * ph`; `makeToY` uses the
  // same `(max - min) || 1` range). Falls back to the primary map when there is
  // no secondary axis so `toYSecondary` stays callable.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : valY;
  const toYSecondarySeries = (value: number): number => toYSecondary(value);

  // Resolved value-axis gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or
  // the faint `#e0e0e0`/0.5 px default). The vertical (horizontal-bar) path
  // strokes gridlines inline, so it reads `grid.color`/`grid.width` directly.
  const grid = valGridStroke(chart, ptToPx);
  ctx.textBaseline = 'middle';
  const drawnValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxLabelFontPx
    : Math.max(8, Math.min(11, ph / 20));
  ctx.font = chartFontCss(
    drawnValTickFontPx,
    chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
    chart.valAxisFontBold ?? false,
    chart.valAxisFontItalic ?? false,
  );
  // Honor `<c:valAx><c:txPr>…<a:solidFill>` when present (ECMA-376 §21.2.2.*);
  // otherwise keep the neutral gray default.
  const valLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
  ctx.fillStyle = valLabelColor;

  if (!chart.valAxisHidden) {
    // Minor gridlines (under the majors) when the file declares them.
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    for (const val of plan.minorLines) {
      if (!isH) {
        strokeValueGridlineH(ctx, px0, pw, valY(val), false, minorGrid);
      } else {
        const gx = valX(val);
        ctx.strokeStyle = minorGrid.color; ctx.lineWidth = minorGrid.width;
        const previousDash = minorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (minorGrid.dash.length > 0) ctx.setLineDash(minorGrid.dash);
        ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
        if (minorGrid.dash.length > 0) ctx.setLineDash(previousDash);
      }
    }
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const val of plan.majorLines) {
      // The zero line is the emphasized gridline (`si === 0` was that line only
      // while the axis was anchored at 0; with a negative minimum it moves up).
      const isZero = Math.abs(val) < step * 1e-9;
      const label = formatPrimaryValueAxisTick(chart, val, primaryPercentAxis);
      if (!isH) {
        const gy = valY(val);
        if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, isZero, grid);
        if (drawLabels) {
          ctx.textAlign = 'right';
          const gap = options.gapPolicy === 'chartex'
            ? chartExValueTickLabelOffsetPx(ptToPx)
            : chart.valAxisFontSizeHpt != null
              ? valueTickLabelGapPx(drawnValTickFontPx)
              : 12;
          ctx.fillText(label, px0 - gap, gy);
        }
      } else {
        const gx = valX(val);
        if (drawMajorGrid) {
          // Explicit gridline color ⇒ uniform stroke (no zero-line emphasis),
          // matching PowerPoint; otherwise keep the `#aaa`/1 px baseline rule.
          ctx.strokeStyle = grid.explicit ? grid.color : isZero ? '#aaa' : grid.color;
          ctx.lineWidth = grid.explicit ? grid.width : isZero ? 1 : grid.width;
          const previousDash = grid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
          if (grid.dash.length > 0) ctx.setLineDash(grid.dash);
          ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
          if (grid.dash.length > 0) ctx.setLineDash(previousDash);
        }
        if (drawLabels) {
          ctx.textAlign = 'center';
          const gap = chart.valAxisFontSizeHpt != null
            ? categoryTickLabelGapPx(drawnValTickFontPx)
            : 10;
          ctx.fillText(label, gx, py0 + ph + gap);
        }
      }
    }
  }

  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100).
  // Perpendicular to the value gridlines: vertical for a column chart (cat axis
  // runs along x), horizontal for a horizontal-bar chart (cat axis runs along
  // y). Positioned at the same fractions as the category ticks — band
  // boundaries under crossBetween="between" (bar default), category centers
  // under "midCat". Drawn under the bars (like value gridlines). Office omits
  // these by default so the common path is byte-stable.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    const previousDash = cg.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (cg.dash.length > 0) ctx.setLineDash(cg.dash);
    const gridlineFractions = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
      : catGridlineFractions(chart, n);
    for (const frac of gridlineFractions) {
      ctx.beginPath();
      if (!isH) {
        const gx = px0 + frac * pw;
        ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph);
      } else {
        const gy = py0 + frac * ph;
        ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy);
      }
      ctx.stroke();
    }
    if (cg.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // Axis rules. The CATEGORY axis runs along the bars' baseline — bottom
  // (horizontal) for a column chart, left (vertical) for a horizontal bar
  // chart — and the VALUE axis is perpendicular to it. The previous code
  // assumed the left rule was always the value axis, so a horizontal bar
  // chart whose value axis is `<c:delete val="1">` drew
  // no axis line at all even though its category axis carries an explicit
  // `<c:spPr><a:ln>`. `<a:noFill>` on a line suppresses just the rule (labels
  // stay) → `*AxisLineHidden`; an `<a:solidFill>` gives `*AxisLineColor`/Width
  // (ECMA-376 §21.2.2.* line props). Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  // Colour defaults to '#aaa' (Office's faint default rule); the EMU `<a:ln@w>`
  // is scaled to canvas px by `ptToPx`. See `resolveAxisLine`.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const drawCatLine = !chart.catAxisHidden && !chart.catAxisLineHidden;
  const drawValLine = !chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null;
  const primaryCatCrossValue = categoryAxisCrossingValue(chart, plan.min, plan.max);
  const primaryCatAxisY = !isH ? valY(primaryCatCrossValue) : py0 + ph;
  const primaryCatAxisX = isH ? valX(primaryCatCrossValue) : px0;
  const categoryLabelAxisY = !isH && (chart.catAxisTickLabelPos ?? 'nextTo') === 'nextTo'
    ? primaryCatAxisY
    : py0 + ph;
  const catMajorTickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
  const multiLevelBoundariesOwnMajorTicks = !isH
    && categoryLevels != null
    && dateAxisPlan == null
    && !chart.catAxisLineHidden
    && isCrossBetween(chart)
    && Math.abs(categoryLabelAxisY - primaryCatAxisY) < 0.01;
  // Axis rules + tick marks are drawn AFTER the bars/line (see `drawAxesOnTop`
  // below) so the bars don't paint over the category baseline — PowerPoint
  // keeps the axis line crisp on top of the columns.
  const drawAxesOnTop = (): void => {
    if (!isH) {
      if (drawCatLine) strokeAxisSegment(ctx, px0, primaryCatAxisY, px0 + pw, primaryCatAxisY, catLineColor, catLineW, chart.catAxisLineDash);
      if (drawValLine) strokeAxisSegment(ctx, px0, py0, px0, py0 + ph, valLineColor, valLineW, chart.valAxisLineDash);           // left
    } else {
      if (drawCatLine) strokeAxisSegment(ctx, primaryCatAxisX, py0, primaryCatAxisX, py0 + ph, catLineColor, catLineW, chart.catAxisLineDash);
      if (drawValLine) strokeAxisSegment(ctx, px0, py0 + ph, px0 + pw, py0 + ph, valLineColor, valLineW, chart.valAxisLineDash); // bottom
    }

    // Axis major tick marks (`<c:*Ax><c:majorTickMark>` — ECMA-376 §21.2.2.101).
    // PowerPoint draws short ruler ticks even when the axis rule itself is light,
    // so the bar renderer must emit them too (the line renderer already does).
    // `drawAxisTick`'s `axis` arg selects GEOMETRY: 'val' = vertical rule with
    // horizontal ticks, 'cat' = horizontal rule with vertical ticks. For a
    // column chart the value axis is vertical (left) and the category axis
    // horizontal (bottom); a horizontal bar chart swaps the two.
    if (!chart.valAxisHidden && chart.valAxisMajorTickMark && chart.valAxisMajorTickMark !== 'none') {
      for (const val of plan.majorLines) {
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, valY(val), valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.valAxisMajorTickMark, 'cat', py0 + ph, valX(val), valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
        }
      }
    }
    if (!chart.valAxisHidden && chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of plan.minorTicks) {
        if (!isH) {
          drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, valY(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.valAxisMinorTickMark, 'cat', py0 + ph, valX(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
        }
      }
    }
    // Category-axis major ticks share the major-grid positions: with
    // crossBetween="between" they mark the N+1 interval boundaries, while
    // labels and minor ticks sit at the N category centres. This distinction
    // matters when both levels use `cross`: Office's 6pt major boundary marks
    // remain visibly longer than its 4pt centred minor marks.
    if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
      // Multi-level category brackets already occupy every interval boundary.
      // They absorb the coincident major tick into one continuous stroke below,
      // avoiding darker/thicker Canvas seams from painting the same segment
      // twice. Mid-category ticks remain independent of those boundaries.
      const ordinalFractions = multiLevelBoundariesOwnMajorTicks
        ? []
        : catGridlineFractions(chart, n);
      const tickFractions = dateAxisPlan
        ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
        : ordinalFractions.filter((_, index) => index % catMajorTickSkip === 0);
      for (const frac of tickFractions) {
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCatAxisY, px0 + frac * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.catAxisMajorTickMark, 'val', primaryCatAxisX, py0 + frac * ph, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
        }
      }
    }
    if (!chart.catAxisHidden && chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none') {
      const ordinalFractions = isCrossBetween(chart)
        ? Array.from({ length: n }, (_, ci) => (ci + 0.5) / n)
        : Array.from({ length: Math.max(0, n - 1) }, (_, ci) => (ci + 0.5) / (n - 1));
      const tickFractions = dateAxisPlan
        ? dateAxisPlan.minorTicks.map(tick => tick.fraction)
        : ordinalFractions;
      for (const frac of tickFractions) {
        if (!isH) {
          drawAxisTick(ctx, chart.catAxisMinorTickMark, 'cat', primaryCatAxisY, px0 + frac * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
        } else {
          drawAxisTick(ctx, chart.catAxisMinorTickMark, 'val', primaryCatAxisX, py0 + frac * ph, catLineColor, catLineW, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
        }
      }
    }
  };

  // Bar cluster geometry — ECMA-376 §21.2.2.13 (gapWidth = % of bar width
  // between categories, default 150) and §21.2.2.25 (overlap = signed % of
  // bar width within a cluster, default 0). Within a cluster the pitch
  // between consecutive bars is `barW * (1 - overlap/100)`, so with N series:
  //   clusterWidth = barW + (N - 1) * barW * (1 - overlap/100)
  //   catGap       = clusterWidth + barW * gapWidth/100
  //                = barW * (1 + (N-1) * (1 - overlap/100) + gapWidth/100)
  // Solving for barW gives the formula below. Stacked charts render one bar
  // per category so we treat them as N=1 and overlap=0.
  const categoryGap = (horizontal: boolean): number => horizontal ? ph / n : pw / n;
  const catGap = categoryGap(isH);
  const catRev = catAxisReversed(chart);
  const categorySlotIndex = (ci: number, horizontal: boolean): number => horizontal
    ? (catRev ? ci : n - 1 - ci)
    : (catRev ? n - 1 - ci : ci);
  const categoryBandSize = (ci: number, horizontal = isH): number => dateAxisPlan
    ? dateAxisPlan.categoryBandFractions[ci]! * (horizontal ? ph : pw)
    : categoryGap(horizontal);
  const categoryStart = (ci: number, horizontal = isH): number => dateAxisPlan
    ? (horizontal ? py0 : px0)
      + dateAxisPlan.positions[ci]! * (horizontal ? ph : pw)
      - categoryBandSize(ci, horizontal) / 2
    : (horizontal ? py0 : px0)
      + categorySlotIndex(ci, horizontal) * categoryGap(horizontal);
  const categoryCenterX = (ci: number): number => dateAxisPlan
    ? px0 + dateAxisPlan.positions[ci]! * pw
    : px0 + categorySlotIndex(ci, false) * categoryGap(false) + categoryGap(false) / 2;
  const clusterGeometry = (group: readonly ChartSeries[], categorySize: number) => {
    const owner = group[0];
    const isStackedGroup = owner ? groupIsStacked(owner) : stacked;
    const effective = isStackedGroup ? 1 : Math.max(1, group.length);
    const rawOverlap = owner?.barGroupOverlap ?? chart.barOverlap ?? 0;
    const overlapPct = isStackedGroup || !Number.isFinite(rawOverlap)
      ? 0
      : Math.max(-100, Math.min(100, rawOverlap));
    const gapWidthPct = resolveCategoryGapWidthPercent(
      owner?.barGroupGapWidth ?? chart.barGapWidth,
      options.gapPolicy ?? 'legacy',
    );
    const denom = 1 + (effective - 1) * (1 - overlapPct / 100) + gapWidthPct / 100;
    const barW = categorySize / denom;
    const clusterGap = isStackedGroup ? 0 : barW * (1 - overlapPct / 100);
    const clusterWidth = barW + (effective - 1) * clusterGap;
    return { barW, clusterGap, catStart: (categorySize - clusterWidth) / 2 };
  };
  type BarSeriesLinePoint = {
    categoryStart: number;
    categoryEnd: number;
    valueEnd: number;
  };
  const hasVerifiedBarSeriesLines = (chart.barGroupDecorations ?? []).some(
    decoration => decoration.seriesLines?.length === 1,
  );
  const barSeriesLinePoints: Array<Array<BarSeriesLinePoint | null>> | null =
    hasVerifiedBarSeriesLines
      ? barSeries.map(() => new Array<BarSeriesLinePoint | null>(n).fill(null))
      : null;

  // A classic combo chart may place an `<c:areaChart>` group behind a
  // `<c:barChart>` group. Area series are not bars: they share the category
  // coordinate system, map through their bound value axis, and fill from the
  // zero baseline to their authored top edge. Paint them before columns so the
  // later bar group remains visible, matching Excel's foreground columns for
  // this mixed-family layout.
  for (let si = 0; si < areaSeries.length; si++) {
    const series = areaSeries[si];
    const pointOverrides = indexPointOverrides(series.dataPointOverrides);
    const chartIndex = sourceSeriesIndices.get(series) ?? si;
    const styleIndex = chartExSeriesFormatIndex(series, chartIndex);
    const color = chartColor(chartIndex, series);
    const fillDecision = classicDataPointFillDecision(chart, series, undefined, styleIndex);
    const yOf = sec && series.useSecondaryAxis === true
      ? toYSecondarySeries
      : toYPrimaryLine;
    const dispBlanks = chart.dispBlanksAs ?? 'zero';
    let run: Array<{ x: number; y: number; baseY: number }> = [];
    const paintRun = (): void => {
      if (run.length === 0) return;
      if (dateAxisPlan) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(px0, py0, pw, ph);
        ctx.clip();
      }
      ctx.beginPath();
      ctx.moveTo(run[0].x, run[0].baseY);
      ctx.lineTo(run[0].x, run[0].y);
      appendCurve(ctx, run, false);
      for (let index = run.length - 1; index >= 0; index--) {
        ctx.lineTo(run[index].x, run[index].baseY);
      }
      ctx.closePath();
      paintClassicDataPointPath(ctx, fillDecision, {
        x: run[0].x,
        y: py0,
        w: Math.max(1, run[run.length - 1].x - run[0].x),
        h: ph,
      }, color, ptToPx, shapeRotationDeg);
      if (applyClassicStyleLine(
        ctx, chart, 'dataPoint', series, undefined, styleIndex, color,
        1.5, ptToPx, { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      )) {
        ctx.stroke();
      }
      if (dateAxisPlan) ctx.restore();
      run = [];
    };
    for (let ci = 0; ci < n; ci++) {
      if (series.sourceHidden?.[ci] === true) {
        paintRun();
        continue;
      }
      const value = series.values[ci];
      if (value == null) {
        if (dispBlanks === 'gap') paintRun();
        if (dispBlanks !== 'zero') continue;
      }
      run.push({
        x: categoryCenterX(ci),
        y: yOf(overlayValue(series, ci)),
        baseY: yOf(overlayBase(series, ci)),
      });
    }
    paintRun();

    const resolvedMarkerDetail = seriesHasResolvedMarkerDetail(chart, series, chartIndex);
    const seriesMarkersVisible = (series.showMarker === true || seriesHasMarkerDetail(series))
      && series.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(series)) {
      const markerRadius = Math.max(2, 2.5 * ptToPx);
      for (let ci = 0; ci < n; ci++) {
        if (series.sourceHidden?.[ci] === true) continue;
        const value = series.values[ci];
        if (value == null) continue;
        const point = pointOverrides.get(ci);
        const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        const markerX = categoryCenterX(ci);
        const markerY = yOf(overlayValue(series, ci));
        if (resolvedMarkerDetail || pointHasMarkerDetail(point)) {
          const lineWidthEmu = point?.markerLineWidthEmu ?? series.markerLineWidthEmu;
          drawChartMarker(
            ctx, chart, series, point, ci, markerX, markerY, symbol,
            point?.markerSize ?? series.markerSize ?? 5,
            markerFillColorFor(series, point, ci, color),
            point?.markerLine ?? series.markerLine ?? null,
            ptToPx,
            lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
            markerFillPaintFor(series, point, ci),
            shapeRotationDeg,
          );
        } else {
          ctx.fillStyle = color;
          ctx.beginPath();
          ctx.arc(markerX, markerY, markerRadius, 0, Math.PI * 2);
          ctx.fill();
        }
      }
    }
  }

  for (let ci = 0; ci < n; ci++) {
    // Stacked charts accumulate positive and negative contributions on opposite
    // sides of the zero line, so each category tracks two running offsets.
    const positiveOffsets = new Map<string, number>();
    const negativeOffsets = new Map<string, number>();
    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const seriesIsHorizontal = groupIsHorizontal(s);
      const categorySize = categoryBandSize(ci, seriesIsHorizontal);
      const secondary = sec != null && s.useSecondaryAxis === true;
      const group = barGroupFor(s);
      const groupKey = barGroupKey(s);
      const isStackedGroup = groupIsStacked(s);
      const isPercentGroup = groupIsPercent(s);
      const stackSum = isPercentGroup
        ? group.reduce((sum, member) => sum + Math.abs(member.values[ci] ?? 0), 0) || 1
        : 1;
      const posOffset = positiveOffsets.get(groupKey) ?? 0;
      const negOffset = negativeOffsets.get(groupKey) ?? 0;
      const groupIndex = Math.max(0, group.indexOf(s));
      const { barW, clusterGap, catStart } = clusterGeometry(group, categorySize);
      const valueY = secondary && secScale ? secScale.makeToY(py0, ph) : valY;
      // `<c:catAx><c:crosses>` locates the category axis on its paired value
      // axis (ECMA-376 §21.2.2.33/.34). A secondary top axis commonly uses
      // `crosses="max"`; its columns therefore start at the maximum rule and
      // extend downward. Treating every secondary group as zero-based reverses
      // that authored geometry even though the top rule itself is painted.
      const secondaryBase = secScale
        ? secondaryCat?.crossesAt != null && Number.isFinite(secondaryCat.crossesAt)
          ? Math.max(secScale.min, Math.min(secScale.max, secondaryCat.crossesAt))
          : secondaryCat?.crosses === 'max'
            ? secScale.max
            : secondaryCat?.crosses === 'min'
              ? secScale.min
              : Math.max(secScale.min, Math.min(secScale.max, 0))
        : 0;
      const groupZeroY = secondary ? valueY(secondaryBase) : zeroY;
      const raw = s.values[ci] ?? 0;
      // Signed value in axis units (percent keeps its sign — a negative slice of
      // a percentStacked chart reaches below the zero line).
      const sv = isPercentGroup
        ? (raw / stackSum) * percentGroupMultiplier(s)
        : raw;
      const negative = sv < 0;
      const labelAxisMaximum = secondary && secScale ? secScale.max : plan.max;
      const useNegativeStyle = negative
        && (s.invertIfNegative === true || s.automaticNegativeStyle === true);
      // A `<c:dPt>` fill is an explicit point override regardless of the
      // chart-group `varyColors` flag (§21.2.2.52). varyColors only controls
      // the fallback palette for points without an override.
      const pointOverride = pointOverrides[si].get(ci);
      const pointColor = pointOverride?.color ?? s.dataPointColors?.[ci];
      const seriesVariesByPoint = barVariesByPoint(s);
      const pointStyleIndex = seriesVariesByPoint ? ci : barStyleIndices[si];
      const styleFill = classicDataPointFillDecision(
        chart, s, pointOverride, pointStyleIndex, ci,
      );
      const color = styleFill?.fillType === 'solid'
        ? (styleFill.color.startsWith('#') ? styleFill.color : `#${styleFill.color}`)
        : pointColor
          ? `#${pointColor}`
          : seriesVariesByPoint ? pieSliceColor(ci, s) : chartColor(si, s);
      const invertedPaint = useNegativeStyle
        ? s.automaticNegativeStyle === true || s.invertedFillHidden === true
          ? null
          : s.invertedFill
        : undefined;
      const rawDataPointStyle = rawLinkedChartStyleRole(chart, 'dataPoint');
      const pointStructuredFill = chartStyleDirectFillDecision(
        pointOverride?.chartexStyle, rawDataPointStyle, pointStyleIndex,
      );
      const pointNoFill = pointOverride?.fillHidden === true
        ? chartStyleDirectNoFillDecision(rawDataPointStyle)
        : undefined;
      const pointOwnsFill = pointStructuredFill !== undefined
        || pointNoFill !== undefined
        || pointOverride?.color != null;
      const pointPaint = pointOwnsFill
        ? styleFill
        : invertedPaint !== undefined
          ? invertedPaint
          : styleFill;
      const applyPointOutline = (target: CanvasRenderingContext2D): boolean => {
        const hasPointLine = pointOverride?.lineHidden != null
          || pointOverride?.lineColor != null
          || pointOverride?.lineWidthEmu != null
          || pointOverride?.lineDash != null
          || pointOverride?.chartexStyle?.linePaintAuthored === true
          || pointOverride?.chartexStyle?.lineHidden === true;
        if (hasPointLine) {
          return applyClassicStyleLine(
            target, chart, 'dataPoint', s, pointOverride, pointStyleIndex, color,
            1, ptToPx, { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg, false,
          );
        }
        if (useNegativeStyle
          && (s.invertedLineHidden != null
            || s.invertedLineColor != null
            || s.invertedLineWidthEmu != null)) {
          if (s.invertedLineHidden) return false;
          target.strokeStyle = `#${s.invertedLineColor ?? '000000'}`;
          target.lineWidth = axisLineWidthPx(s.invertedLineWidthEmu, ptToPx);
          target.setLineDash([]);
          return true;
        }
        if (s.automaticNegativeStyle === true) {
          target.strokeStyle = '#000000';
          target.lineWidth = 0.75 * ptToPx;
          target.setLineDash([]);
          return true;
        }
        const omittedAlternateLineFallback = useNegativeStyle
          && chart.chartType === 'clusteredBar'
          && chart.legacyChartStyle === 2
          && s.invertedFillAuthored === true
          && s.invertedFill != null
          && s.invertedLineAuthored === false;
        const seriesOwnsOutline = s.lineHidden === true
          || s.lineColor != null
          || s.lineWidthEmu != null;
        if (omittedAlternateLineFallback && seriesOwnsOutline) {
          if (s.lineHidden || !s.lineColor) return false;
          target.strokeStyle = `#${s.lineColor}`;
          target.lineWidth = axisLineWidthPx(s.lineWidthEmu, ptToPx);
          target.setLineDash([]);
          return true;
        }
        // Office 2010's alternate negative-fill extension records no line
        // provenance when `<a:ln>` is omitted. The observed Style 2 clustered
        // column boundary supplies a black 0.75pt outline; keep that
        // application default here, after direct point/series line ownership,
        // instead of forging an authored line in the shared parser model.
        if (omittedAlternateLineFallback) {
          target.strokeStyle = '#000000';
          target.lineWidth = 0.75 * ptToPx;
          target.setLineDash([]);
          return true;
        }
        const linkedDataPoint = chartDataPointStyleRole(
          chart, 'dataPoint', sourceSeriesIndices.get(s) ?? si,
        ) ?? chart.chartexDataPointStyle;
        const pointStructuredLine = chartStyleLineDecision(
          pointOverride?.chartexStyle, pointStyleIndex,
        );
        const directStructuredLine = pointStructuredLine !== undefined
          ? pointStructuredLine : chartStyleLineDecision(s.chartexStyle, pointStyleIndex);
        if (!linkedDataPoint && !isChartExColumn && directStructuredLine === undefined
          && !s.lineColor && s.lineHidden !== true) {
          return false;
        }
        return applyClassicStyleLine(
          target, chart, 'dataPoint', s, pointOverride, pointStyleIndex, color,
          1, ptToPx, { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg, false,
        );
      };
      const pointEffect = chartStyleEffectOwner(
        pointOverride?.chartexStyle,
        s.chartexStyle,
      );
      const paintBarAt = (
        target: CanvasRenderingContext2D,
        bx: number,
        by: number,
        barPaintWidth: number,
        barPaintHeight: number,
      ): void => {
        paintClassicDataPointRect(
          target,
          pointPaint === undefined ? s.fillPattern : pointPaint,
          { x: bx, y: by, w: barPaintWidth, h: barPaintHeight },
          color,
          ptToPx,
          shapeRotationDeg,
        );
        if (barPaintWidth > 0 && barPaintHeight > 0 && applyPointOutline(target)) {
          const outlineW = target.lineWidth;
          target.strokeRect(
            bx + outlineW / 2,
            by + outlineW / 2,
            Math.max(0, barPaintWidth - outlineW),
            Math.max(0, barPaintHeight - outlineW),
          );
        }
      };

      if (!seriesIsHorizontal) {
        const bx = isStackedGroup
          ? categoryStart(ci, false) + catStart
          : categoryStart(ci, false) + catStart + groupIndex * clusterGap;
        // A date axis can explicitly crop categories through min/max. Marks
        // wholly outside that authored plot interval do not bleed into the
        // value-axis/title gutter.
        if (bx + barW <= px0 || bx >= px0 + pw) continue;
        // Column: the bar spans between the zero line and the value. Stacked
        // bars start at the running offset for their sign; clustered bars start
        // at the zero line.
        const y0 = isStackedGroup ? valueY(negative ? negOffset : posOffset) : groupZeroY;
        const y1 = isStackedGroup
          ? valueY((negative ? negOffset : posOffset) + sv)
          : valueY(sv);
        const by = clamp(Math.min(y0, y1), py0, py0 + ph);
        const barBottom = clamp(Math.max(y0, y1), py0, py0 + ph);
        const barH = Math.max(0, barBottom - by);
        if (barSeriesLinePoints && s.values[ci] != null) {
          barSeriesLinePoints[si][ci] = {
            categoryStart: bx,
            categoryEnd: bx + barW,
            valueEnd: clamp(y1, py0, py0 + ph),
          };
        }
        paintChartStyleEffects(
          ctx,
          pointEffect,
          chartDataPointStyleRole(chart, 'dataPoint', sourceSeriesIndices.get(s) ?? si),
          pointStyleIndex,
          { x: bx, y: by, w: barW, h: barH },
          ptToPx,
          target => paintBarAt(target, bx, by, barW, barH),
        );
        const seriesLabels = s.seriesDataLabels;
        const label = resolveChartExLabel(
          chart, s, ci, s.categories?.[ci] ?? cats[ci] ?? '', raw,
          {
            visible: chart.showDataLabels,
            showVal: chart.showDataLabels && !isPercentGroup,
            showPercent: chart.showDataLabels && isPercentGroup,
            showCatName: false,
          },
          labelOverrides[si],
          isPercentGroup ? sv / 100 : undefined,
          s.useSecondaryAxis && sec ? sec.displayUnits : chart.valAxisDisplayUnits,
        );
        if (label && dataLabelWithinAxisMaximum(chart, raw, labelAxisMaximum)) {
          // ECMA-376 §21.2.2.30 / §21.1.2.3.2 — data label font size comes from
          // `<c:dLbls><c:txPr>...<a:defRPr@sz>` (hundredths of a point). When
          // the file specifies one we honor it; otherwise the proportional
          // heuristic keeps small bars readable.
          const sizeHpt = label.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = chartTextFontSizePx(sizeHpt, ptToPx)
            ?? Math.max(7, Math.min(11, barW * 0.6));
          const authoredLabel = labelOverrides[si].get(ci);
          const bold = label.fontBold
            || (seriesLabels?.fontBold == null && authoredLabel?.fontBold == null);
          const labelFont = chartFontFamily(
            chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
          );
          ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${lsz}px ${labelFont}`;
          // drawBarDataLabel takes (bx, by, barL=length, barW=thickness). For
          // a vertical column bar, "length" is the bar's height and
          // "thickness" is its horizontal width — pass them in that order.
          // Previously the args were (barW, barH) which silently swapped the
          // two and made `cx = bx + barW/2` (the horizontal-center formula
          // inside the helper) use the bar's HEIGHT instead of its width,
          // pushing data labels far to the right of the bar.
          drawBarDataLabel(
            ctx, label.text,
            bx, by, barH, barW,
            'vertical',
            label.position ?? chart.dataLabelPosition ?? (isStackedGroup ? 'ctr' : null),
            s.dataLabelColors?.[ci] ?? label.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            lsz,
            { x: px0, y: py0, w: pw, h: ph },
            { x, y, w, h },
            authoredLabel?.manualLayout,
            negative,
            customRichDataLabelOptions(
              chart,
              authoredLabel,
              ptToPx,
              labelFont,
              bold,
              label.textStyle,
            ),
            label.showLegendKey
              ? dataLabelLegendKey(sourceSeriesIndices.get(s) ?? si, ci)
              : undefined,
            label.textStyle,
            ptToPx,
            mergeChartLabelBoxes(authoredLabel?.labelBox, seriesLabels?.labelBox),
            shapeRotationDeg,
          );
        }
      } else {
        // Cluster positions are local to the owning `<c:barChart>` group.
        // A second overlay group restarts at its own first bar; using the
        // flattened plot-area series index pushes it outside the category
        // slot when preceding groups contain more series.
        const siVisual = groupIndex;
        const by = isStackedGroup
          ? categoryStart(ci, true) + catStart
          : categoryStart(ci, true) + catStart + siVisual * clusterGap;
        const x0 = isStackedGroup ? valX(negative ? negOffset : posOffset) : zeroX;
        const x1 = isStackedGroup ? valX((negative ? negOffset : posOffset) + sv) : valX(sv);
        const bx = clamp(Math.min(x0, x1), px0, px0 + pw);
        const barRight = clamp(Math.max(x0, x1), px0, px0 + pw);
        const barL = Math.max(0, barRight - bx);
        if (barSeriesLinePoints && s.values[ci] != null) {
          barSeriesLinePoints[si][ci] = {
            categoryStart: by,
            categoryEnd: by + barW,
            valueEnd: clamp(x1, px0, px0 + pw),
          };
        }
        paintChartStyleEffects(
          ctx,
          pointEffect,
          chartDataPointStyleRole(chart, 'dataPoint', sourceSeriesIndices.get(s) ?? si),
          pointStyleIndex,
          { x: bx, y: by, w: barL, h: barW },
          ptToPx,
          target => paintBarAt(target, bx, by, barL, barW),
        );
        const seriesLabels = s.seriesDataLabels;
        const label = resolveChartExLabel(
          chart, s, ci, s.categories?.[ci] ?? cats[ci] ?? '', raw,
          {
            visible: chart.showDataLabels,
            showVal: chart.showDataLabels && !isPercentGroup,
            showPercent: chart.showDataLabels && isPercentGroup,
            showCatName: false,
          },
          labelOverrides[si],
          isPercentGroup ? sv / 100 : undefined,
          s.useSecondaryAxis && sec ? sec.displayUnits : chart.valAxisDisplayUnits,
        );
        if (label && dataLabelWithinAxisMaximum(chart, raw, labelAxisMaximum)) {
          const sizeHpt = label.fontSizeHpt ?? chart.dataLabelFontSizeHpt;
          const lsz = chartTextFontSizePx(sizeHpt, ptToPx)
            ?? Math.max(7, Math.min(11, barW * 0.6));
          const authoredLabel = labelOverrides[si].get(ci);
          const bold = label.fontBold
            || (seriesLabels?.fontBold == null && authoredLabel?.fontBold == null);
          const labelFont = chartFontFamily(
            chart, label.fontFace ?? chart.dataLabelFontFace, 'minor',
          );
          ctx.font = `${label.textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${lsz}px ${labelFont}`;
          drawBarDataLabel(
            ctx, label.text,
            bx, by, barL, barW,
            'horizontal',
            label.position ?? chart.dataLabelPosition ?? (isStackedGroup ? 'ctr' : null),
            s.dataLabelColors?.[ci] ?? label.fontColor ?? s.labelColor ?? chart.dataLabelFontColor ?? null,
            lsz,
            { x: px0, y: py0, w: pw, h: ph },
            { x, y, w, h },
            authoredLabel?.manualLayout,
            negative,
            customRichDataLabelOptions(
              chart,
              authoredLabel,
              ptToPx,
              labelFont,
              bold,
              label.textStyle,
            ),
            label.showLegendKey
              ? dataLabelLegendKey(sourceSeriesIndices.get(s) ?? si, ci)
              : undefined,
            label.textStyle,
            ptToPx,
            mergeChartLabelBoxes(authoredLabel?.labelBox, seriesLabels?.labelBox),
            shapeRotationDeg,
          );
        }
      }
      if (isStackedGroup) {
        if (negative) negativeOffsets.set(groupKey, negOffset + sv);
        else positiveOffsets.set(groupKey, posOffset + sv);
      }
    }
  }

  // `CT_BarChart/serLines` uses one group-owned line style. MS-OE376
  // 2.1.1578 defines each segment as joining adjacent data points in the same
  // series. Office vector output resolves those points to the value-end edge of
  // each stacked bar and clips the segment to the category gap: columns join
  // right/left facing edges, horizontal bars join bottom/top facing edges.
  // Missing points break the sequence instead of inventing a zero-valued end.
  if (barSeriesLinePoints) {
    const barSeriesByGroup = new Map<number, number[]>();
    for (let seriesIndex = 0; seriesIndex < barSeries.length; seriesIndex++) {
      const groupIndex = barSeries[seriesIndex].barGroupIndex ?? 0;
      const members = barSeriesByGroup.get(groupIndex);
      if (members) members.push(seriesIndex);
      else barSeriesByGroup.set(groupIndex, [seriesIndex]);
    }
    for (const decoration of chart.barGroupDecorations ?? []) {
      // CT_BarChart permits multiple serLines children. The single-child Office
      // geometry is verified; precedence/association for multiple children is
      // application-defined and remains fail-closed until its boundary output is
      // adjudicated rather than guessing first/last/cyclic style semantics.
      if (decoration.seriesLines?.length !== 1) {
        continue;
      }
      ctx.save();
      const seriesLineStyle = chartStyleRoleLine(
        chart, decoration.seriesLines[0], 'seriesLine',
      );
      if (!applyDecorationLineStyle(ctx, seriesLineStyle, ptToPx)) {
        ctx.restore();
        continue;
      }
      ctx.beginPath();
      ctx.rect(px0, py0, pw, ph);
      ctx.clip();
      for (const seriesIndex of barSeriesByGroup.get(decoration.groupIndex) ?? []) {
        const points = barSeriesLinePoints[seriesIndex];
        const seriesIsHorizontal = groupIsHorizontal(barSeries[seriesIndex]);
        for (let categoryIndex = 0; categoryIndex + 1 < n; categoryIndex++) {
          const current = points[categoryIndex];
          const next = points[categoryIndex + 1];
          if (!current || !next) continue;
          const currentCenter = (current.categoryStart + current.categoryEnd) / 2;
          const nextCenter = (next.categoryStart + next.categoryEnd) / 2;
          const forward = nextCenter >= currentCenter;
          ctx.beginPath();
          if (!seriesIsHorizontal) {
            ctx.moveTo(
              forward ? current.categoryEnd : current.categoryStart,
              current.valueEnd,
            );
            ctx.lineTo(
              forward ? next.categoryStart : next.categoryEnd,
              next.valueEnd,
            );
          } else {
            ctx.moveTo(
              current.valueEnd,
              forward ? current.categoryEnd : current.categoryStart,
            );
            ctx.lineTo(
              next.valueEnd,
              forward ? next.categoryStart : next.categoryEnd,
            );
          }
          ctx.stroke();
        }
      }
      ctx.restore();
    }
  }

  // CT_BarSer permits the same trendline and error-bar children as line
  // series. Paint both above the filled rectangles and below axes/labels.
  // Geometry is derived from the same gapWidth/overlap cluster calculation as
  // the bars, so a clustered series' adornments remain centered on its bars.
  const barCategoryCenter = (series: ChartSeries, categoryIndex: number): number => {
    const group = barGroupFor(series);
    const groupIndex = Math.max(0, group.indexOf(series));
    const horizontal = groupIsHorizontal(series);
    const categorySize = categoryBandSize(categoryIndex, horizontal);
    const geometry = clusterGeometry(group, categorySize);
    const start = categoryStart(categoryIndex, horizontal) + geometry.catStart;
    if (groupIsStacked(series)) return start + geometry.barW / 2;
    return start + groupIndex * geometry.clusterGap + geometry.barW / 2;
  };
  const continuousBarCategoryCenter = (series: ChartSeries, index: number): number => {
    if (Number.isInteger(index) && index >= 0 && index < n) {
      return barCategoryCenter(series, index);
    }
    const group = barGroupFor(series);
    const groupIndex = Math.max(0, group.indexOf(series));
    const horizontal = groupIsHorizontal(series);
    const gap = categoryGap(horizontal);
    const geometry = clusterGeometry(group, gap);
    const slot = horizontal
      ? (catRev ? index : n - 1 - index)
      : (catRev ? n - 1 - index : index);
    const start = (horizontal ? py0 : px0) + slot * gap + geometry.catStart;
    const visualIndex = groupIsStacked(series) ? 0 : groupIndex;
    return start + visualIndex * geometry.clusterGap + geometry.barW / 2;
  };
  for (let seriesIndex = 0; seriesIndex < barSeries.length; seriesIndex++) {
    const series = barSeries[seriesIndex];
    const seriesIsHorizontal = groupIsHorizontal(series);
    const secondary = sec != null && series.useSecondaryAxis === true;
    const valueAt = seriesIsHorizontal
      ? valX
      : secondary && secScale ? secScale.makeToY(py0, ph) : valY;
    const color = chartColor(seriesIndex, series);
    const plotted = (categoryIndex: number): number =>
      plottedBarValue(seriesIndex, categoryIndex);
    for (const errorBars of effectiveBarErrorBars(series)) {
      drawBarErrorBars(
        ctx, series, chartStyleRoleErrorBar(chart, errorBars), n, seriesIsHorizontal,
        categoryIndex => barCategoryCenter(series, categoryIndex),
        valueAt,
        plotted, color, ptToPx,
      );
    }
    drawSeriesTrendlines(
      ctx, series, color,
      index => continuousBarCategoryCenter(series, index - 1),
      valueAt,
      ptToPx,
      series.values.map((_value, index) => index + 1),
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
      (index, value) => seriesIsHorizontal
        ? ({
          x: valueAt(value),
          y: continuousBarCategoryCenter(series, index - 1),
        })
        : ({
          x: continuousBarCategoryCenter(series, index - 1),
          y: valueAt(value),
        }),
    );
  }

  if ((!hasDataTable || isH) && !chart.catAxisHidden && catLabelsVisible(chart)) {
    // `<c:catAx><c:txPr>…<a:solidFill>` colors the category tick labels.
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    const drawnCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, catGap * 0.5));
    ctx.font = chartFontCss(
      drawnCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Column: each label is centered in a category slot of width `catGap`, so
    // cap it just under that so neighbours don't collide. Horizontal bars: the
    // label sits right-aligned in the left gutter between the val-title/legend
    // band and the plot edge, so cap it at that band width.
    const catSlotMaxPx = catGap - 4;
    const horizLabelMaxPx = (px0 - 4) - (x + legLeftW + valTitleW);
    // `<c:catAx><c:txPr><a:bodyPr rot>` rotates the column labels (0 = flat).
    const rotRad = catLabelRotation;
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        raw: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        fraction: tick.fraction,
        categoryIndex: -1,
      }))
      : cats.map((category, categoryIndex) => ({
        // §21.2.2.71: a category-axis numFmt formats numeric-serial categories
        // (e.g. dateAx serials → real dates). No-op for string categories.
        raw: formatCategoryLabel(category.toString(), chart.catAxisFormatCode, chart.date1904),
        fraction: null,
        categoryIndex,
      }));
    for (const entry of labelEntries) {
      const { raw } = entry;
      if (!isH) {
        const anchor = entry.fraction != null || entry.categoryIndex < 0
          ? { fraction: entry.fraction ?? 0.5, textAlign: 'center' as CanvasTextAlign }
          : categoryLabelAnchorFraction(
            entry.categoryIndex,
            n,
            isCrossBetween(chart),
            catRev,
            chart.catAxisLabelAlignment,
          );
        const lx = px0 + anchor.fraction * pw;
        ctx.textAlign = anchor.textAlign; ctx.textBaseline = 'top';
        // Rotation elides against a longer diagonal budget. Horizontal labels
        // use the measured word-wrap computed from this category slot.
        const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
        const gap = categoryLabelOffsetPx(
          chart.catAxisFontSizeHpt != null
            ? categoryTickLabelGapPx(drawnCatTickFontPx)
            : 3,
          chart.catAxisLabelOffsetPercent,
        );
        if (rotRad === 0) {
          const lines = entry.categoryIndex >= 0
            ? (wrappedColumnCategories[entry.categoryIndex] ?? [raw])
            : [raw];
          lines.forEach((line, lineIndex) => {
            ctx.fillText(line, lx, categoryLabelAxisY + gap + lineIndex * (drawnCatTickFontPx + 2));
          });
        } else {
          drawRotatedCatLabel(ctx, elideToWidth(ctx, raw, budget), lx, categoryLabelAxisY + gap, rotRad);
        }
      } else {
        const ly = entry.fraction != null
          ? py0 + entry.fraction * ph
          : py0 + categorySlotIndex(entry.categoryIndex, true) * catGap + catGap / 2;
        const gap = categoryLabelOffsetPx(
          chart.catAxisFontSizeHpt != null
            ? valueTickLabelGapPx(drawnCatTickFontPx)
            : 4,
          chart.catAxisLabelOffsetPercent,
        );
        const boxStart = x + legLeftW + valTitleW;
        const boxEnd = px0 - gap;
        const alignment = chart.catAxisLabelAlignment;
        const lx = alignment === 'l'
          ? boxStart
          : alignment === 'ctr' ? (boxStart + boxEnd) / 2 : boxEnd;
        // Horizontal bars historically right-align omitted category labels in
        // the left gutter. `lblAlgn` only replaces that default when authored.
        ctx.textAlign = alignment === 'l' ? 'left' : alignment === 'ctr' ? 'center' : 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(elideToWidth(ctx, raw, horizLabelMaxPx), lx, ly);
      }
    }

    if (!isH && categoryLevels) {
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(drawnCatTickFontPx)
          : 3,
        chart.catAxisLabelOffsetPercent,
      );
      const rowHeight = drawnCatTickFontPx + 4;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      ctx.strokeStyle = catLineColor;
      ctx.lineWidth = catLineW;
      ctx.setLineDash([]);
      const boundaryStartY = (boundaryIndex: number): number => {
        if (!multiLevelBoundariesOwnMajorTicks
          || boundaryIndex % catMajorTickSkip !== 0) {
          return categoryLabelAxisY;
        }
        const tickLength = axisTickLengthPx('major', catLineW, ptToPx);
        if (chart.catAxisMajorTickMark === 'cross') {
          return categoryLabelAxisY - tickLength / 2;
        }
        if (chart.catAxisMajorTickMark === 'in') {
          return categoryLabelAxisY - tickLength;
        }
        return categoryLabelAxisY;
      };

      // Each boundary in the innermost category level separates adjacent
      // labels. Boundaries shared by an outer level are extended below by the
      // outer-level brackets drawn next; extend the remaining boundaries
      // through the first label band here.
      const outerBoundaryIndices = new Set<number>();
      for (let levelIndex = 1; levelIndex < categoryLevels.length; levelIndex++) {
        const level = categoryLevels[levelIndex] ?? [];
        for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
          if ((level[categoryIndex] ?? '') !== '') outerBoundaryIndices.add(categoryIndex);
        }
        outerBoundaryIndices.add(n);
      }
      const firstBandBottom = categoryLabelAxisY + gap + drawnCatTickFontPx + 2;
      for (let boundaryIndex = 0; boundaryIndex <= n; boundaryIndex++) {
        if (outerBoundaryIndices.has(boundaryIndex)) continue;
        const boundary = px0 + boundaryIndex / n * pw;
        ctx.beginPath();
        ctx.moveTo(boundary, boundaryStartY(boundaryIndex));
        ctx.lineTo(boundary, firstBandBottom);
        ctx.stroke();
      }

      const outerBoundaryBottoms = new Map<number, number>();
      for (let levelIndex = 1; levelIndex < categoryLevels.length; levelIndex++) {
        const level = categoryLevels[levelIndex] ?? [];
        const starts: number[] = [];
        for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
          if ((level[categoryIndex] ?? '') !== '') starts.push(categoryIndex);
        }
        for (let groupIndex = 0; groupIndex < starts.length; groupIndex++) {
          const start = starts[groupIndex];
          const end = starts[groupIndex + 1] ?? n;
          const label = level[start] ?? '';
          const left = px0 + start / n * pw;
          const right = px0 + end / n * pw;
          const rowTop = categoryLabelAxisY + gap + levelIndex * rowHeight;
          const alignment = chart.catAxisLabelAlignment;
          const labelX = alignment === 'l'
            ? left
            : alignment === 'r' ? right : (left + right) / 2;
          ctx.textAlign = alignment === 'l' ? 'left' : alignment === 'r' ? 'right' : 'center';
          ctx.fillText(
            elideToWidth(ctx, label, Math.max(0, right - left - 4)),
            labelX,
            rowTop,
          );
          const bracketBottom = rowTop + drawnCatTickFontPx + 2;
          for (const boundaryIndex of [start, end]) {
            outerBoundaryBottoms.set(
              boundaryIndex,
              Math.max(outerBoundaryBottoms.get(boundaryIndex) ?? categoryLabelAxisY, bracketBottom),
            );
          }
        }
      }
      // A boundary can belong to both neighbouring groups, and at deeper
      // levels the same category index can recur. Paint each authored
      // boundary once to its deepest required extent so coincident strokes do
      // not become darker or wider.
      for (const [boundaryIndex, bracketBottom] of outerBoundaryBottoms) {
        const boundary = px0 + boundaryIndex / n * pw;
        ctx.beginPath();
        ctx.moveTo(boundary, boundaryStartY(boundaryIndex));
        ctx.lineTo(boundary, bracketBottom);
        ctx.stroke();
      }
    }
  }

  if (lineSeries.length > 0 && !isH) {
    drawLineGroupDecorations(
      ctx, chart, n, categoryCenterX,
      series => sec && series.useSecondaryAxis === true ? toYSecondarySeries : toYPrimaryLine,
      () => primaryCatAxisY,
      (series, index) => series.values[index] == null ? null : overlayValue(series, index),
      catGap, ptToPx, shapeRotationDeg, 'background',
    );
    if (dateAxisPlan) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(px0, py0, pw, ph);
      ctx.clip();
    }
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      const color = chartColor(barSeries.length + si, s);
      // Series bound to the secondary axis map through its scale; others use
      // the primary (bar) value axis.
      const yOf = sec && s.useSecondaryAxis === true
        ? toYSecondarySeries
        : toYPrimaryLine;
      const sourceIndex = sourceSeriesIndices.get(s) ?? si;
      const varyingLinePaint = chartSeriesVariesByPoint(chart, sourceIndex);
      const styleIndex = chartExSeriesFormatIndex(s, sourceIndex);
      const hasAuthoredLine = s.chartexStyle != null
        || s.lineHidden != null
        || s.lineColor != null
        || s.lineWidthEmu != null
        || chart.chartexDataPointLineStyle != null;
      const smooth = s.smooth === true;
      const dispBlanks = chart.dispBlanksAs ?? 'gap';
      const runs: IndexedLinePoint[][] = [];
      let run: IndexedLinePoint[] = [];
      const flushRun = (): void => {
        if (run.length === 0) return;
        runs.push(run);
        run = [];
      };
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (s.sourceHidden?.[ci] === true) {
          flushRun();
          continue;
        }
        if (v == null) {
          if (dispBlanks === 'gap') flushRun();
          if (dispBlanks !== 'zero') continue;
        }
        const lx = categoryCenterX(ci);
        run.push({ x: lx, y: yOf(overlayValue(s, ci)), index: ci });
      }
      flushRun();
      if (varyingLinePaint) {
        paintClassicVaryingLineSegments(
          ctx, chart, s, runs, smooth, false, color,
          2, ptToPx, { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
          options.semanticLineNoStyleFallback !== false,
        );
      } else {
        const paintOverlayLine = (target: CanvasRenderingContext2D): void => {
          const visible = hasAuthoredLine
            ? applyChartExSeriesLineStyle(
              target,
              chart,
              chart.chartexDataPointLineStyle,
              s,
              styleIndex,
              lineSeries.length,
              color,
              ptToPx,
              { linkedNoStyleFallback: options.semanticLineNoStyleFallback },
            )
            : true;
          if (!visible) return;
          if (!hasAuthoredLine) {
            target.strokeStyle = color;
            target.lineWidth = 2;
            target.setLineDash([]);
          }
          target.beginPath();
          for (const points of runs) {
            target.moveTo(points[0].x, points[0].y);
            appendCurve(target, points, smooth);
          }
          target.stroke();
        };
        paintChartStyleEffects(
          ctx,
          chartStyleEffectOwner(s.chartexStyle),
          chart.chartStyleRoles?.dataPointLine,
          styleIndex,
          { x: px0, y: py0, w: pw, h: ph },
          ptToPx,
          paintOverlayLine,
        );
      }
      const seriesMarkersVisible = s.showMarker !== false && s.markerSymbol !== 'none';
      const drawMarkers = seriesMarkersVisible || hasVisiblePointMarkerOverride(s);
      const hasMarkerDetail = seriesHasResolvedMarkerDetail(chart, s, sourceIndex);
      if (drawMarkers) {
        for (let ci = 0; ci < n; ci++) {
          if (s.sourceHidden?.[ci] === true) continue;
          const v = s.values[ci];
          if (v == null) continue;
          const lx = categoryCenterX(ci);
          const ly = yOf(overlayValue(s, ci));
          const point = pointOverrides.get(ci);
          const symbol = effectiveMarkerSymbol(s, point, 'circle', seriesMarkersVisible);
          if (symbol === 'none') continue;
          if (hasMarkerDetail || pointHasMarkerDetail(point)) {
            const lineWidthEmu = point?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawChartMarker(
              ctx, chart, s, point, ci, lx, ly, symbol,
              point?.markerSize ?? s.markerSize ?? 5,
              markerFillColorFor(s, point, ci, color),
              point?.markerLine ?? s.markerLine ?? null,
              ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, point, ci),
              shapeRotationDeg,
            );
          } else {
            ctx.fillStyle = color;
            ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
          }
        }
      }
      // Trendlines (`<c:trendline>`, §21.2.2.211) for the combo line series.
      drawSeriesTrendlines(
        ctx, s, color,
        (i) => categoryCenterX(i),
        yOf, ptToPx, undefined,
        {
          chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
          shapeRotationDeg,
        },
      );
    }
    drawLineGroupDecorations(
      ctx, chart, n, categoryCenterX,
      series => sec && series.useSecondaryAxis === true ? toYSecondarySeries : toYPrimaryLine,
      () => primaryCatAxisY,
      (series, index) => series.values[index] == null ? null : overlayValue(series, index),
      catGap, ptToPx, shapeRotationDeg, 'foreground',
    );
    if (dateAxisPlan) ctx.restore();
  }

  // A scatter group can be overlaid on a bar chart with its own pair of
  // numeric axes (ECMA-376 CT_ScatterChart `axId`, first X then Y). This is the
  // standard construction for dot/range plots: an invisible horizontal bar
  // series supplies category labels and the visible scatter markers plus
  // custom X error bars supply the dots and connecting ranges.
  if (scatterSeries.length > 0) {
    const allX: number[] = [];
    const allY: number[] = [];
    for (const s of scatterSeries) {
      const sx = s.categories ?? [];
      for (let i = 0; i < s.values.length; i++) {
        const xv = scatterXValue(sx, i, false);
        const yv = s.values[i];
        if (xv == null || yv == null) continue;
        allX.push(xv);
        allY.push(yv);
      }
    }
    if (allX.length && allY.length) {
      const xAxis = chart.secondaryCatAxis;
      const yAxis = chart.secondaryValAxis;
      const xExtent = finiteDataExtent(allX);
      const yExtent = finiteDataExtent(allY);
      const needsMinor = (axis: SecondaryValueAxis | null | undefined): boolean =>
        axis?.minorGridlines === true
        || (axis?.minorTickMark != null && axis.minorTickMark !== 'none');
      const xScale = planNumericValueAxis({
        dataMin: xExtent.min,
        dataMax: xExtent.max,
        explicitMin: xAxis?.min,
        explicitMax: xAxis?.max,
        axisLenPt: pw / ptToPx,
        axisOrientation: 'horizontal',
        majorUnit: xAxis?.majorUnit,
        minorUnit: xAxis?.minorUnit,
        needMinor: needsMinor(xAxis),
        logBase: xAxis?.logBase,
        reversed: xAxis?.orientation === 'maxMin',
      });
      const yScale = planNumericValueAxis({
        dataMin: yExtent.min,
        dataMax: yExtent.max,
        explicitMin: yAxis?.min,
        explicitMax: yAxis?.max,
        axisLenPt: ph / ptToPx,
        axisOrientation: 'vertical',
        majorUnit: yAxis?.majorUnit,
        minorUnit: yAxis?.minorUnit,
        needMinor: needsMinor(yAxis),
        logBase: yAxis?.logBase,
        reversed: yAxis?.orientation === 'maxMin',
      });
      const scatterToX = (value: number): number =>
        px0 + xScale.fraction(value) * pw;
      const scatterToY = (value: number): number =>
        py0 + ph - yScale.fraction(value) * ph;
      drawScatterSeriesLayer(
        ctx,
        chart,
        scatterSeries.map((series, index) => ({
          series,
          index: sourceSeriesIndices.get(series) ?? index,
        })),
        false,
        scatterToX,
        scatterToY,
        r,
        px0,
        py0,
        pw,
        ph,
        ptToPx,
        false,
        chart.scatterStyle ?? 'marker',
        { x, y, w, h },
        yScale.max,
        undefined,
        shapeRotationDeg,
      );
    }
  }

  // Primary axis rules + ticks on top of the bars/line so the category
  // baseline stays visible (the bars would otherwise paint over it).
  drawAxesOnTop();

  if (secondaryCat && !isH) {
    drawSecondaryCategoryAxis(
      ctx, chart, secondaryCat, secondaryCategories, r, px0, py0, pw, ptToPx,
    );
  }

  // Secondary value axis (right edge). Independent scale: its own "nice" major
  // unit drives the tick labels, positioned via `toYSecondary` (NOT aligned to
  // the primary gridlines — PowerPoint places them independently). Draws its
  // rule + ticks on the right; ticks mirror the left axis ("out" points right).
  if (sec && secScale) {
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r, px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, valLabelColor, chart.date1904, secondaryPercentAxis,
    );
  }

  if (dataTableLayout) {
    const tableY = py0 + ph + (isH
      ? (chart.valAxisHidden ? h * 0.02 : catAxisLabelBandH(valAxLabelFontPx))
      : 0);
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, tableY, pw, x + legLeftW, ptToPx,
    );
  }

  const legendPaints = varyingBarSeries && chart.series.length === 1
    ? varyingBarSeries.values.map((_, pointIndex) => classicDataPointFillDecision(
      chart,
      varyingBarSeries,
      pointOverrides[varyingBarSeriesIndex].get(pointIndex),
      pointIndex,
      pointIndex,
    ))
    : isChartExColumn
      ? barSeries.flatMap((series, index) => [
        chartExDataPointPaint(
          chart, barStyleIndices[index], barSeries.length, series.chartexStyle, series.color,
        ),
        ...trendlineLegendSeries([series]).map(() => undefined),
      ])
      : [];
  drawLegendForLayout(
    ctx, legendChart, leg, x, y, w, h, px0, py0, pw, ph,
    titleH + 2, ptToPx, legendPaints,
  );
  drawAxisTitles(
    ctx, chart, x, y, w, h, px0, py0, pw, ph,
    legLeftW, legBottomH, catTitlePx, valTitlePx, isH,
  );
}
