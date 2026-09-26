// Classic area chart family.
import type { ChartModel, ChartRect, ChartSeries } from '../../types/chart';

import { classicDataPointFillDecision } from '../classic-data-point-style.js';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';

import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markerFillColorFor,
  markerFillPaintFor,
  pointHasMarkerDetail,
  seriesHasMarkerDetail,
} from '../marker-style.js';

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
  valueTickLabelGapPx,
} from '../layout.js';

import { axisLineWidthPx, resolveAxisLine, isCrossBetween } from '../axis-style.js';
import { formatCategoryLabel } from '../chart-number-format.js';

import { categoryLabelAnchorFraction, categoryLabelOffsetPx } from '../category-spacing.js';

import { indexChartPlotGroups } from '../plot-groups.js';

import {
  chartDataPointStyleRole,
  chartSeriesSourceIndex,
  chartSeriesVariesByPoint,
} from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';

import { chartColor, indexPointOverrides, applyClassicStyleLine, paintClassicVaryingLineSegments, chartExSeriesFormatIndex } from '../shared/palette.js';
import type { IndexedLinePoint } from '../shared/palette.js';
import { chartFontFamily, chartFontCss } from '../shared/fonts.js';
import { drawAxisTitles } from '../shared/axis.js';
import { chartHasDataTable, chartDataTableBaseHeight, chartDataTableHeaderWidth, measureChartDataTable, drawChartDataTable } from '../shared/data-table.js';
import { createDataLabelLegendKeyResolver, measuredLegendReserve, drawLegendForLayout } from '../shared/legend.js';
import { drawAxisTick, strokeAxisSegment, strokeValueGridlineH, valGridStroke, valMinorGridStroke, drawCatMajorGridlines, catGridStroke, catGridlineFractions, catAxisReversed, drawValMajorGridlines, formatPrimaryValueAxisTick, formatAxisTickWithUnits, planValueAxis, axisLabelPx } from '../shared/axis.js';
import { drawSeriesTrendlines } from '../shared/trendline.js';
import { chartDateAxisPlan, forEachErrorBarEndpoint, computeSecondaryAxis, drawSecondaryValueGridlines, drawSecondaryValueAxis } from '../shared/secondary-axis.js';
import { measuredCartesianTitleBand, drawChartTitleForLayout } from '../shared/title.js';
import { dataLabelWithinAxisMaximum, drawCategoryDataLabels } from '../shared/data-labels.js';
import { chartCategories } from '../category-spacing.js';
import { applyDecorationLineStyle, chartStyleRoleLine, chartStyleRoleErrorBar } from '../shared/style-roles.js';
import { axisCrossingValue, categoryAxisCrossingValue } from '../shared/line-decorations.js';
import { drawChartMarker, seriesHasResolvedMarkerDetail } from '../shared/markers.js';
import { appendCurve } from '../shared/geometry.js';
import { drawCategoryErrorBars } from '../shared/error-bars.js';
import { paintClassicDataPointPath } from '../shared/chartex-style.js';

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════
export function renderAreaChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n === 0) return;
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  // A plot area can contain both `<c:areaChart>` and `<c:lineChart>` groups.
  // Only the ordered area-group series participate in the filled stack; line
  // series share the axes but remain independent overlays (§21.2.2.145).
  const areaSeries = chart.series
    .map((series, chartIndex) => ({ series, chartIndex }))
    .filter(({ series }) => series.seriesType == null || series.seriesType === 'area');
  const lineSeries = chart.series
    .map((series, chartIndex) => ({ series, chartIndex }))
    .filter(({ series }) => series.seriesType === 'line');
  if (areaSeries.length === 0 && lineSeries.length === 0) return;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const legacyGrouping = chart.chartType === 'stackedAreaPct'
    ? 'percentStacked' : chart.chartType === 'stackedArea' ? 'stacked' : 'standard';
  const sourceAreaGroups = chart.plotGroups?.filter(group => group.kind === 'area') ?? [{
    kind: 'area' as const,
    seriesStart: 0,
    seriesCount: areaSeries.length,
    categoryAxis: 'primary' as const,
    valueAxis: 'primary' as const,
    seriesAxis: 'none' as const,
    grouping: legacyGrouping,
  }];
  const areaIndexAtChartIndex = new Array<number>(chart.series.length).fill(-1);
  for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
    areaIndexAtChartIndex[areaSeries[areaIndex].chartIndex] = areaIndex;
  }
  const areaGroupMembers = sourceAreaGroups.map(group => {
    const areaIndices: number[] = [];
    const end = Math.min(chart.series.length, group.seriesStart + group.seriesCount);
    for (let chartIndex = group.seriesStart; chartIndex < end; chartIndex++) {
      const areaIndex = areaIndexAtChartIndex[chartIndex];
      if (areaIndex >= 0) areaIndices.push(areaIndex);
    }
    return { group, areaIndices };
  });
  const stackedByArea = new Array<boolean>(areaSeries.length).fill(false);
  const percentByArea = new Array<boolean>(areaSeries.length).fill(false);
  const percentTotalsByArea = new Array<number[] | null>(areaSeries.length).fill(null);
  const areaBaseValues = areaSeries.map(() => new Array<number>(n).fill(0));
  const areaTopValues = areaSeries.map(() => new Array<number>(n).fill(0));
  const axisPlanningGroups = chart.plotGroups?.filter(group =>
    (group.kind === 'area' || group.kind === 'line') && group.seriesCount > 0
  ) ?? sourceAreaGroups;
  const allPercentByAreaAxis = new Map<string, boolean>();
  for (const group of axisPlanningGroups) {
    const axis = group.valueAxis;
    allPercentByAreaAxis.set(
      axis,
      (allPercentByAreaAxis.get(axis) ?? true) && group.grouping === 'percentStacked',
    );
  }
  for (const { group, areaIndices } of areaGroupMembers) {
    const grouping = group.grouping ?? 'standard';
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const pct = grouping === 'percentStacked';
    const multiplier = pct && allPercentByAreaAxis.get(group.valueAxis) === true ? 100 : 1;
    const totals = pct
      ? cats.map((_, categoryIndex) => areaIndices.reduce(
          (sum, areaIndex) => sum + Math.abs(
            areaSeries[areaIndex].series.values[categoryIndex] ?? 0,
          ), 0,
        ) || 1)
      : null;
    for (const areaIndex of areaIndices) {
      stackedByArea[areaIndex] = stacked;
      percentByArea[areaIndex] = pct;
      percentTotalsByArea[areaIndex] = totals;
    }
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      let positive = 0;
      let negative = 0;
      for (const areaIndex of areaIndices) {
        const raw = areaSeries[areaIndex].series.values[categoryIndex] ?? 0;
        const contribution = pct && totals ? raw / totals[categoryIndex] * multiplier : raw;
        const base = contribution >= 0 ? positive : negative;
        areaBaseValues[areaIndex][categoryIndex] = stacked ? base : 0;
        areaTopValues[areaIndex][categoryIndex] = stacked ? base + contribution : contribution;
        if (stacked) {
          if (contribution >= 0) positive += contribution;
          else negative += contribution;
        }
      }
    }
  }
  const primaryAxisGroups = axisPlanningGroups.filter(group => group.valueAxis !== 'secondary');
  const axisIsPercent = primaryAxisGroups.length > 0
    && primaryAxisGroups.every(group => group.grouping === 'percentStacked');

  // Combo area charts may bind some series to a SECONDARY value axis on the
  // right (ECMA-376 §21.2.2.*). As with line, this applies only to plain
  // (unstacked) area — a stacked/percentStacked secondary combo is not an Office
  // construct. `sec` is null (single-axis, byte-identical to pre-CH7) unless the
  // axis is declared AND a series opts in; secondary series are then excluded
  // from the primary extent and mapped through the secondary scale.
  const areaIndexBySeries = new Map(areaSeries.map((entry, index) => [entry.series, index]));
  const seriesChartIndex = new Map(chart.series.map((series, index) => [series, index]));
  const seriesUsesSecondary = (series: ChartSeries): boolean => {
    const chartIndex = seriesChartIndex.get(series) ?? -1;
    return plotGroupBySeries[chartIndex]?.valueAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  };
  const secondaryAxisGroups = axisPlanningGroups.filter(group => group.valueAxis === 'secondary');
  const secondaryAxisIsPercent = secondaryAxisGroups.length > 0
    && secondaryAxisGroups.every(group => group.grouping === 'percentStacked');
  const sec = chart.secondaryValAxis && chart.series.some(series => seriesUsesSecondary(series))
    ? chart.secondaryValAxis
    : null;
  const isSecondarySeries = (series: ChartSeries): boolean => {
    return sec != null && seriesUsesSecondary(series);
  };

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  // Vertical pads first so the estimated plot height is known before the
  // secondary-axis scale + right-gutter measurement (same ordering as bar/line).
  // Top: title band + half a value-axis label above the top gridline. Bottom:
  // PowerPoint's category-label band (gap + line-height + margin).
  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  const secScale = computeSecondaryAxis(
    sec,
    chart.series,
    phEst / ptToPx,
    'y',
    secondaryAxisIsPercent,
    false,
    series => seriesUsesSecondary(series),
    (series, pointIndex) => {
      const areaIndex = areaIndexBySeries.get(series);
      if (areaIndex == null) return series.values[pointIndex] ?? null;
      return series.values[pointIndex] == null
        ? null : areaTopValues[areaIndex][pointIndex] ?? null;
    },
  );
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = `${secFontPx}px ${chartFontFamily(chart, sec.fontFace, 'minor')}`;
    let wmax = 0;
    for (const value of secScale.majorLines) {
      wmax = Math.max(wmax, ctx.measureText(formatAxisTickWithUnits(value, sec.formatCode ?? null, chart.date1904, sec.displayUnits)).width);
    }
    secLabelBandW = wmax + 18;
    ctx.font = prevFont;
  }
  const secTitleBandW = sec && sec.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  // Resolve the primary extent before frame placement so an authored
  // `layoutTarget="outer"` can be converted to the inner data rectangle using
  // the actual formatted tick-label width. The outer rectangle includes axis
  // labels and ticks (ECMA-376 §21.2.2.89); treating its left edge as `px0`
  // pushes the labels outside chart space.
  const computeAreaDataExtent = (): { min: number; max: number } => {
    let min = Infinity;
    let max = -Infinity;
    for (let ci = 0; ci < n; ci++) {
      for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
        const { series } = areaSeries[areaIndex];
        if (isSecondarySeries(series) || series.values[ci] == null) continue;
        min = Math.min(min, areaBaseValues[areaIndex][ci], areaTopValues[areaIndex][ci]);
        max = Math.max(max, areaBaseValues[areaIndex][ci], areaTopValues[areaIndex][ci]);
      }
      for (const { series } of lineSeries) {
        if (isSecondarySeries(series)) continue;
        const value = series.values[ci];
        if (value == null) continue;
        min = Math.min(min, value);
        max = Math.max(max, value);
      }
    }
    if (!isFinite(min) || !isFinite(max)) return { min: 0, max: 1 };
    if (axisIsPercent) return { min: min < 0 ? -100 : 0, max: max > 0 ? 100 : 0 };
    return { min, max };
  };
  let areaExtent = computeAreaDataExtent();
  if (!axisIsPercent) {
    const includeEndpoint = (value: number): void => {
      areaExtent = {
        min: Math.min(areaExtent.min, value),
        max: Math.max(areaExtent.max, value),
      };
    };
    for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
      const { series } = areaSeries[areaIndex];
      if (isSecondarySeries(series)) continue;
      forEachErrorBarEndpoint(
        series,
        'y',
        index => {
          if (series.values[index] == null) return null;
          return areaTopValues[areaIndex][index];
        },
        includeEndpoint,
      );
    }
    for (const { series } of lineSeries) {
      if (isSecondarySeries(series)) continue;
      forEachErrorBarEndpoint(series, 'y', index => series.values[index] ?? null, includeEndpoint);
    }
  }
  const provisionalScale = planValueAxis(
    chart,
    areaExtent.min,
    areaExtent.max,
    phEst / ptToPx,
    axisIsPercent,
  );
  const manualValTickFontPx = chart.valAxisFontSizeHpt != null
    ? valAxFontPx
    : Math.max(8, Math.min(11, phEst / 20));
  let primaryLabelWidth = 0;
  if (
    !chart.valAxisHidden
    && chart.plotAreaManualLayout != null
    && chart.plotAreaManualLayout.layoutTarget !== 'inner'
  ) {
    const prevFont = ctx.font;
    ctx.font = chartFontCss(
      manualValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    for (const value of provisionalScale.majorLines) {
      primaryLabelWidth = Math.max(
        primaryLabelWidth,
        ctx.measureText(formatPrimaryValueAxisTick(chart, value, axisIsPercent)).width,
      );
    }
    ctx.font = prevFont;
  }
  const manualOuterInsets = chartManualOuterAxisInsets({
    valAxisHidden: chart.valAxisHidden,
    catAxisHidden: chart.catAxisHidden,
    valLabelWidth: primaryLabelWidth,
    valLabelFontPx: manualValTickFontPx,
    catLabelFontPx: catAxFontPx,
    valLabelGapPx: chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(manualValTickFontPx)
      : 6,
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

  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: legLeftW + Math.max(w * 0.12 + valTitleW, dataTableHeaderW),
  };

  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  const areaFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  const { px0, py0, pw } = areaFrame.plotRect;
  let { ph } = areaFrame.plotRect;
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // Primary extent from the PRIMARY series only (secondary series live on their
  // own axis). When `sec` is null every series is primary, byte-identical to
  // the pre-CH7 path.
  // Value axis is vertical → length = plot height (axis-length-aware auto major unit). An
  // explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the auto step.
  const areaPlan = planValueAxis(
    chart, areaExtent.min, areaExtent.max, ph / ptToPx, axisIsPercent,
  );

  // crossBetween="between" (Office's default; ECMA-376 §21.2.2.32 leaves the
  // default application-defined) gives each category a band of width pw/n and
  // plots its point at the band CENTER, leaving a half-band margin before the
  // first and after the last category — matching PowerPoint's Jan…Dec inset.
  // "midCat" anchors points on the category dividers (flush to the axes).
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  const toX = dateAxisPlan
    ? (index: number) => px0 + dateAxisPlan.positions[index]! * pw
    : between
      ? (index: number) => {
        const i = catRev ? n - 1 - index : index;
        return px0 + ((i + 0.5) / n) * pw;
      }
      : (index: number) => {
        const i = catRev ? n - 1 - index : index;
        return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
      };
  const toY = (v: number) => py0 + ph - areaPlan.frac(v) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;
  const primaryCategoryAxisY = toY(
    categoryAxisCrossingValue(chart, areaPlan.min, areaPlan.max),
  );
  const secondaryCategoryAxisY = sec && secScale
    ? toYSecondary(axisCrossingValue(
      chart.secondaryCatAxis?.crossesAt,
      chart.secondaryCatAxis?.crosses,
      secScale.min,
      secScale.max,
    ))
    : primaryCategoryAxisY;
  const categoryAxisYFor = (series: ChartSeries): number =>
    isSecondarySeries(series) ? secondaryCategoryAxisY : primaryCategoryAxisY;

  // Axis line colour/weight from `<c:*Ax><c:spPr><a:ln>` (EMU → px at scale),
  // mirroring the bar/line renderers. Office leaves the value-axis rule off by
  // default (gridlines stand in), so only draw it when the file specifies one.
  const { color: catLineColor, width: catLineW } = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const { color: valLineColor, width: valLineW } = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);

  // Value-axis MAJOR gridlines are drawn UNDER the series (before the fills), so
  // an opaque/translucent area occludes the gridlines inside its region —
  // matching Office vector observations in which opaque area fill occludes
  // gridlines below its top edge. This mirrors the bar/line/stock/
  // scatter/waterfall/box renderers, which already stroke gridlines first. The
  // axis rules, tick marks and value/category labels stay AFTER the series (drawn
  // further below) so they sit atop the plot. `<c:valAx><c:majorGridlines>` is on
  // by default (`drawValMajorGridlines`); `<c:minorGridlines>` only when declared.
  if (!chart.valAxisHidden) {
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    // Minor gridlines (`<c:valAx><c:minorGridlines>`, §21.2.2.129) drawn first,
    // UNDER the majors and the series when the file declares them. An omitted
    // minor unit uses the shared automatic major/5 fallback.
    if (chart.valAxisMinorGridlines) {
      for (const v of areaPlan.minorLines) {
        strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
      }
    }
    if (drawValMajorGridlines(chart)) {
      for (const v of areaPlan.majorLines) {
        strokeValueGridlineH(ctx, px0, pw, toY(v), v === 0, grid);
      }
    }
  }
  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }
  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks, also under the fills. Off by default
  // (byte-stable when the file omits them).
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart)) {
    const cg = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = cg.color;
    ctx.lineWidth = cg.width;
    const previousDash = cg.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (cg.dash.length > 0) ctx.setLineDash(cg.dash);
    const fractions = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => tick.fraction)
      : catGridlineFractions(chart, n);
    for (const frac of fractions) {
      const gx = px0 + frac * pw;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (cg.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // Draw the series area fills ON TOP of the gridlines laid down above.
  // In a stacked area chart, series order is the stacking order: series 0 is
  // adjacent to the category axis, then series 1, and so on (CT_AreaChart's
  // ordered `ser` sequence). Standard areas use the same document paint order:
  // a later series is painted later and therefore overlays earlier series,
  // matching Excel's vector output at their intersections.
  const seriesOrder = areaGroupMembers.flatMap(({ areaIndices }) => areaIndices);
  const plottedAreaValue = (areaIndex: number, categoryIndex: number): number =>
    areaTopValues[areaIndex]?.[categoryIndex] ?? 0;
  for (const areaIndex of seriesOrder) {
    const { series: s, chartIndex } = areaSeries[areaIndex];
    const color = chartColor(chartIndex, s);
    const styleIndex = chartExSeriesFormatIndex(s, chartIndex);
    const fillDecision = classicDataPointFillDecision(chart, s, undefined, styleIndex);
    const baseY = py0 + ph;
    // Unstacked secondary series ride their own vertical scale; the stacked
    // branch is never reached with a secondary axis (`sec` is null when
    // stacked), so its `toY` mapping stays the primary one.
    const yOf = yMapFor(s);

    // Smooth (`<c:ser><c:smooth>`, §21.2.2.194) curves the top edge through the
    // points; the baseline connection stays straight. Non-smooth keeps the exact
    // prior moveTo/lineTo sequence (byte-stable) — appendCurve with smooth=false
    // emits identical lineTo calls.
    //
    // NB: `CT_AreaSer` (§A.5.1) has no `<c:smooth>` child (only `CT_LineSer` /
    // `CT_ScatterSer` do), so `extract_series_smooth` never sets `s.smooth` for
    // a real area series and this branch is dead against actual chart XML —
    // it only fires for a model constructed directly (tests / other producers).
    // Kept for symmetry with the line renderer above rather than dropped.
    const smooth = s.smooth === true;
    const paintAreaSeries = (target: CanvasRenderingContext2D): void => {
      target.beginPath();
      if (stackedByArea[areaIndex]) {
        const topPts = [];
        for (let ci = 0; ci < n; ci++) {
          topPts.push({ x: toX(ci), y: toY(areaTopValues[areaIndex][ci]) });
        }
        target.moveTo(topPts[0].x, topPts[0].y);
        appendCurve(target, topPts, smooth);
        for (let ci = n - 1; ci >= 0; ci--) {
          target.lineTo(toX(ci), toY(areaBaseValues[areaIndex][ci]));
        }
      } else {
        const topPts = [];
        for (let ci = 0; ci < n; ci++) {
          topPts.push({ x: toX(ci), y: yOf(s.values[ci] ?? 0) });
        }
        target.moveTo(toX(0), baseY);
        target.lineTo(topPts[0].x, topPts[0].y);
        appendCurve(target, topPts, smooth);
        target.lineTo(toX(n - 1), baseY);
      }
      target.closePath();
      paintClassicDataPointPath(
        target, fillDecision, { x: px0, y: py0, w: pw, h: ph },
        color, ptToPx, shapeRotationDeg,
      );
      if (applyClassicStyleLine(
        target, chart, 'dataPoint', s, undefined, styleIndex, color,
        1.5, ptToPx, { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      )) target.stroke();
    };
    paintChartStyleEffects(
      ctx,
      chartStyleEffectOwner(s.chartexStyle),
      chartDataPointStyleRole(chart, 'dataPoint', chartSeriesSourceIndex(chart, s)),
      styleIndex,
      { x: px0, y: py0, w: pw, h: ph },
      ptToPx,
      paintAreaSeries,
    );
  }

  // `CT_AreaChart` includes `dropLines` through `EG_AreaChartShared`
  // (ECMA-376 Part 1, dml-chart.xsd). Office vector output establishes one
  // drop line per category, spanning the extrema of the category-axis crossing
  // and every plotted point in the owning group. This matters for a standard
  // multi-series area chart (one envelope line, not one line per series) and
  // for an interior crossing (the line spans points on both sides). Paint after
  // the opaque area fills so the authored geometry remains visible, but before
  // point markers and labels.
  const decorationAreaGroupMembers = new Map<number, Array<{ series: ChartSeries; areaIndex: number }>>();
  for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
    const series = areaSeries[areaIndex].series;
    const groupIndex = series.areaGroupIndex ?? 0;
    const members = decorationAreaGroupMembers.get(groupIndex) ?? [];
    members.push({ series, areaIndex });
    decorationAreaGroupMembers.set(groupIndex, members);
  }
  for (const decoration of chart.areaGroupDecorations ?? []) {
    if (!decoration.dropLines) {
      continue;
    }
    const dropLineStyle = chartStyleRoleLine(chart, decoration.dropLines, 'dropLine');
    if (!applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      continue;
    }
    const members = decorationAreaGroupMembers.get(decoration.groupIndex) ?? [];
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      let minY = Infinity;
      let maxY = -Infinity;
      let hasPoint = false;
      for (const member of members) {
        if (member.series.values[categoryIndex] == null) continue;
        const pointY = yMapFor(member.series)(
          plottedAreaValue(member.areaIndex, categoryIndex),
        );
        const axisY = categoryAxisYFor(member.series);
        if (!Number.isFinite(pointY) || !Number.isFinite(axisY)) continue;
        minY = Math.min(minY, pointY, axisY);
        maxY = Math.max(maxY, pointY, axisY);
        hasPoint = true;
      }
      if (!hasPoint || Math.abs(maxY - minY) < 0.01) continue;
      ctx.beginPath();
      ctx.moveTo(toX(categoryIndex), minY);
      ctx.lineTo(toX(categoryIndex), maxY);
      ctx.stroke();
    }
  }

  // Markers, error bars, and per-point data labels for area series. Drawn in a
  // SEPARATE forward pass (after all fills) so the fill loop above stays
  // byte-identical, and each block fires ONLY for series carrying the relevant
  // fields — an area chart with no marker/errBar/dLbl detail draws exactly as
  // before. The plotted top-of-band value matches where the fill's top edge sat
  // (cumulative for stacked). ECMA-376 §21.2.2.32 / §21.2.2.20 / §21.2.2.45.
  //
  // NB: an area chart's filled region has always read a blank cell as 0
  // (`?? 0`), so `<c:dispBlanksAs>` (§21.2.2.42) is a no-op for the area family
  // here — breaking or spanning a *filled* region is not modeled, and changing
  // the default would break byte-stability. dispBlanksAs steers the line family
  // (where "gap" is the historical default).
  {
    const areaMarkerR = Math.max(2, 2.5 * ptToPx);
    // Top of each series' band per category (stacked); the raw value otherwise.
    // Rebuilt independently of the fill loop's mutated stackBase. The ordered
    // series sequence stacks forward, so band si reaches Σ_{k=0..si}.
    for (let areaIndex = 0; areaIndex < areaSeries.length; areaIndex++) {
      const { series: s, chartIndex } = areaSeries[areaIndex];
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      const color = chartColor(chartIndex, s);
      const yOf = yMapFor(s);
      const plottedOf = (ci: number): number => plottedAreaValue(areaIndex, ci);
      const seriesPercentTotals = percentTotalsByArea[areaIndex];
      // Error bars first (markers overlay their tips).
      for (const eb of s.errBars ?? []) {
        drawCategoryErrorBars(
          ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, color,
        );
      }
      // Markers only when the series opts in (`<c:marker>` symbol/size/… — area
      // charts default to NO markers, so nothing fires without explicit detail).
      const seriesMarkersVisible = (s.showMarker === true || seriesHasMarkerDetail(s))
        && s.markerSymbol !== 'none';
      const resolvedMarkerDetail = seriesHasResolvedMarkerDetail(chart, s, chartIndex);
      if (seriesMarkersVisible || hasVisiblePointMarkerOverride(s)) {
        for (let ci = 0; ci < n; ci++) {
          if (s.sourceHidden?.[ci] === true) continue;
          if (s.values[ci] == null) continue;
          const dpt = pointOverrides.get(ci);
          const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
          if (symbol === 'none') continue;
          const px = toX(ci); const py = yOf(plottedOf(ci));
          if (resolvedMarkerDetail || pointHasMarkerDetail(dpt)) {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = markerFillColorFor(s, dpt, ci, color);
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            const lineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawChartMarker(
              ctx, chart, s, dpt, ci, px, py, symbol, sizePt, fill, line, ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
            );
          } else {
            ctx.fillStyle = color;
            ctx.beginPath(); ctx.arc(px, py, areaMarkerR, 0, Math.PI * 2); ctx.fill();
          }
        }
      }
      // Per-point / series-level data labels. Area's filled region has always
      // read a blank cell as 0 (`?? 0`, see the topValue/plottedOf comment
      // above), so every category index is a "plotted" point here regardless
      // of dispBlanksAs — pass true unconditionally (byte-stable: unchanged
      // from before this parameter existed).
      drawCategoryDataLabels(
        ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false, true,
        chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
        // §21.2.2.48 `<c:dLblPos>` precedence: chart-level position, else the
        // area-chart default `'ctr'` (centered on the point, ECMA-376 default
        // for the areaChart group).
        chart.dataLabelPosition ?? 'ctr',
        { x: px0, y: py0, w: pw, h: ph },
        { x, y, w, h },
        percentByArea[areaIndex] && seriesPercentTotals
          ? ci => (s.values[ci] ?? 0) / seriesPercentTotals[ci]
          : undefined,
        undefined,
        face => chartFontFamily(chart, face, 'minor'),
        isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
        ci => dataLabelLegendKey(chartIndex, ci),
        value => dataLabelWithinAxisMaximum(
          chart, value,
          isSecondarySeries(s) && secScale ? secScale.max : areaPlan.max,
        ),
        shapeRotationDeg,
      );
    }
  }

  // Paint `<c:lineChart>` groups after the area fills, using the same category
  // and value-axis transforms. They do not alter the area stack. This is the
  // OOXML combo-chart z-order: later chart groups overlay earlier ones.
  for (const { series: s, chartIndex } of lineSeries) {
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const color = chartColor(chartIndex, s);
    const stroke = s.lineColor ? `#${s.lineColor}` : color;
    const styleIndex = chartExSeriesFormatIndex(s, chartIndex);
    const yOf = yMapFor(s);
    const varyingLinePaint = chartSeriesVariesByPoint(chart, chartIndex);
    const overlayRuns: IndexedLinePoint[][] = [];
    let overlayRun: IndexedLinePoint[] = [];
    const flushOverlayRun = (): void => {
      if (overlayRun.length > 0) overlayRuns.push(overlayRun);
      overlayRun = [];
    };
    for (let ci = 0; ci < n; ci++) {
      const value = s.values[ci];
      if (s.sourceHidden?.[ci] === true) {
        flushOverlayRun();
        continue;
      }
      if (value == null) {
        if ((chart.dispBlanksAs ?? 'gap') === 'gap') flushOverlayRun();
        if ((chart.dispBlanksAs ?? 'gap') !== 'zero') continue;
      }
      overlayRun.push({ x: toX(ci), y: yOf(value ?? 0), index: ci });
    }
    flushOverlayRun();
    const paintOverlayLine = (target: CanvasRenderingContext2D): void => {
      if (!applyClassicStyleLine(
        target, chart, 'dataPointLine', s, undefined, styleIndex, color,
        Math.max(1, 2.25 * ptToPx), ptToPx,
        { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      )) return;
      target.beginPath();
      for (const run of overlayRuns) {
        target.moveTo(run[0].x, run[0].y);
        appendCurve(target, run, s.smooth === true);
      }
      target.stroke();
    };
    if (varyingLinePaint) {
      paintClassicVaryingLineSegments(
        ctx, chart, s, overlayRuns, s.smooth === true, false, color,
        Math.max(1, 2.25 * ptToPx), ptToPx,
        { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      );
    } else {
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

    const plottedOf = (ci: number): number => s.values[ci] ?? 0;
    for (const eb of s.errBars ?? []) {
      drawCategoryErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, stroke,
      );
    }
    const seriesMarkersVisible = (s.showMarker === true || seriesHasMarkerDetail(s))
      && s.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(s)) {
      for (let ci = 0; ci < n; ci++) {
        if (s.sourceHidden?.[ci] === true) continue;
        const value = s.values[ci];
        if (value == null) continue;
        const dpt = pointOverrides.get(ci);
        const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        drawChartMarker(
          ctx, chart, s, dpt, ci, toX(ci), yOf(value), symbol, dpt?.markerSize ?? s.markerSize ?? 5,
          markerFillColorFor(s, dpt, ci, stroke),
          dpt?.markerLine ?? s.markerLine ?? null, ptToPx,
          (dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu) != null
            ? axisLineWidthPx((dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu) as number, ptToPx)
            : undefined,
          markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
        );
      }
    }
    drawCategoryDataLabels(
      ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false, false,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'), chart.dataLabelPosition ?? 'r',
      { x: px0, y: py0, w: pw, h: ph },
      { x, y, w, h },
      undefined,
      undefined,
      face => chartFontFamily(chart, face, 'minor'),
      isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
      ci => dataLabelLegendKey(chartIndex, ci),
      value => dataLabelWithinAxisMaximum(
        chart, value,
        isSecondarySeries(s) && secScale ? secScale.max : areaPlan.max,
      ),
      shapeRotationDeg,
    );
    drawSeriesTrendlines(
      ctx, s, stroke, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  // Value-axis tick marks + labels. The gridlines themselves were already laid
  // down UNDER the series (above the fill loop); here we only add the tick marks
  // and the value labels, which belong ON TOP of the plot.
  if (!chart.valAxisHidden) {
    const drawnValTickFontPx = chart.valAxisFontSizeHpt != null
      ? valAxFontPx
      : Math.max(8, Math.min(11, ph / 20));
    ctx.font = chartFontCss(
      drawnValTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    for (const v of areaPlan.majorLines) {
      const gy = toY(v);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, valLineColor, valLineW, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
      ctx.textAlign = 'right';
      const gap = chart.valAxisFontSizeHpt != null
        ? valueTickLabelGapPx(drawnValTickFontPx)
        : 6;
      ctx.fillText(formatPrimaryValueAxisTick(chart, v, axisIsPercent), px0 - gap, gy);
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of areaPlan.minorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, toY(value), valLineColor, valLineW, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }
  // Category-axis baseline + value-axis rule. Office treats
  // `<c:*Ax><c:spPr><a:ln><a:noFill>` as suppressing the rule and tick marks
  // while labels/gridlines remain. The value rule is drawn only when the file
  // gives it a colour, matching the bar/line renderers.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, primaryCategoryAxisY, px0 + pw, primaryCategoryAxisY,
      catLineColor, catLineW, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden && chart.valAxisLineColor != null) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      valLineColor, valLineW, chart.valAxisLineDash,
    );
  }
  // Category-axis major tick marks. With crossBetween="between" PowerPoint
  // draws them at the band BOUNDARIES (n+1 dividers); "midCat" ticks centers.
  if (!chart.catAxisHidden && chart.catAxisMajorTickMark && chart.catAxisMajorTickMark !== 'none') {
    const tickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    if (dateAxisPlan) {
      for (const tick of dateAxisPlan.majorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY,
          px0 + tick.fraction * pw, catLineColor, catLineW,
          false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash,
        );
      }
    } else if (between) {
      for (let ci = 0; ci <= n; ci += tickSkip) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, px0 + (ci / n) * pw, catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      }
    } else {
      for (let ci = 0; ci < n; ci += tickSkip) {
        drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, toX(ci), catLineColor, catLineW, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      }
    }
  }
  if (
    !chart.catAxisHidden
    && chart.catAxisMinorTickMark
    && chart.catAxisMinorTickMark !== 'none'
    && dateAxisPlan
  ) {
    for (const tick of dateAxisPlan.minorTicks) {
      drawAxisTick(
        ctx, chart.catAxisMinorTickMark, 'cat', primaryCategoryAxisY,
        px0 + tick.fraction * pw, catLineColor, catLineW,
        false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
      );
    }
  }

  if (!hasDataTable && !chart.catAxisHidden) {
    const drawnCatTickFontPx = chart.catAxisFontSizeHpt != null
      ? catAxFontPx
      : Math.max(8, Math.min(11, pw / n * 0.8));
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      drawnCatTickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Category labels are controlled by the authored `<c:tickLblSkip>` interval.
    // Do not add an automatic collision interval: sparse category caches often
    // deliberately leave alternating entries empty to obtain a two-year label
    // cadence, and a computed interval starting at index 0 can discard every
    // non-empty label. Excel paints the authored sparse labels even when the
    // final pair overlaps.
    // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
    // category-axis numFmt before measuring and drawing; string categories
    // pass through unchanged.
    const authoredSkip = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        label: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        x: px0 + tick.fraction * pw,
        categoryIndex: -1,
      }))
      : Array.from({ length: Math.ceil(n / authoredSkip) }, (_, index) => {
        const ci = index * authoredSkip;
        return {
          label: formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904),
          x: toX(ci),
          categoryIndex: ci,
        };
      });
    for (const entry of labelEntries) {
      const label = entry.label;
      if (!label) continue;
      const anchor = entry.categoryIndex < 0
        ? null
        : categoryLabelAnchorFraction(
          entry.categoryIndex,
          n,
          isCrossBetween(chart),
          catAxisReversed(chart),
          chart.catAxisLabelAlignment,
        );
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(drawnCatTickFontPx)
          : 3,
        chart.catAxisLabelOffsetPercent,
      );
      ctx.textAlign = anchor?.textAlign ?? 'center';
      const labelPosition = chart.catAxisTickLabelPos ?? 'nextTo';
      const labelAxisY = labelPosition === 'nextTo'
        ? primaryCategoryAxisY
        : labelPosition === 'high' ? py0 : py0 + ph;
      ctx.fillText(label, anchor ? px0 + anchor.fraction * pw : entry.x, labelAxisY + gap);
    }
  }

  // Secondary value axis (right edge) — drawn after the fills + category labels
  // so it sits atop the plot, mirroring the bar/line ordering.
  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r, px0, py0, pw, ph, ptToPx,
      secFontPx, secLabelBandW, primaryLabelColor, chart.date1904,
    );
  }

  if (dataTableLayout) {
    drawChartDataTable(
      ctx, chart, dataTableLayout, px0, py0 + ph, pw, x + legLeftW, ptToPx,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}
