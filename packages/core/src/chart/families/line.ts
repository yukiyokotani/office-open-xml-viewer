// Classic line chart family.
import type { ChartModel, ChartRect, ChartSeries } from '../../types/chart';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';

import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markerFillColorFor,
  markerFillPaintFor,
  pointHasMarkerDetail,
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

import { effectiveDataLabelText } from '../data-label-content.js';

import { chartSeriesVariesByPoint } from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';

import {
  chartColor,
  indexPointOverrides,
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
  drawAxisTick,
  strokeAxisSegment,
  strokeValueGridlineH,
  valGridStroke,
  valMinorGridStroke,
  drawCatMajorGridlines,
  catGridStroke,
  catGridlineFractions,
  catAxisReversed,
  drawValMajorGridlines,
  formatPrimaryValueAxisTick,
  displayUnitDivisor,
  formatAxisTickWithUnits,
  planValueAxis,
  drawSeriesTrendlines,
  axisLabelPx,
  catLabelsVisible,
  catLabelRotationRad,
  drawRotatedCatLabel,
  chartDateAxisPlan,
  forEachErrorBarEndpoint,
  computeSecondaryAxis,
  drawSecondaryValueGridlines,
  drawSecondaryValueAxis,
  measuredCartesianTitleBand,
  drawChartTitleForLayout,
  chartCategories,
  dataLabelWithinAxisMaximum,
  chartStyleRoleErrorBar,
  drawLineGroupDecorations,
  axisCrossingValue,
  categoryAxisCrossingValue,
  drawChartMarker,
  seriesHasResolvedMarkerDetail,
  drawDataLabelText,
  appendCurve,
  drawCategoryErrorBars,
  drawCategoryDataLabels,
  chartExSeriesFormatIndex,
} from '../shared/classic.js';

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════
export function renderLineChart(
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

  const legacyGrouping = chart.chartType === 'stackedLinePct'
    ? 'percentStacked' : chart.chartType === 'stackedLine' ? 'stacked' : 'standard';
  const lineGroups = chart.plotGroups?.filter(group => group.kind === 'line') ?? [{
    kind: 'line' as const,
    seriesStart: 0,
    seriesCount: chart.series.length,
    categoryAxis: 'primary' as const,
    valueAxis: 'primary' as const,
    seriesAxis: 'none' as const,
    grouping: legacyGrouping,
  }];
  const stackedBySeries = new Array<boolean>(chart.series.length).fill(false);
  const percentBySeries = new Array<boolean>(chart.series.length).fill(false);
  const secondaryBySeries = new Array<boolean>(chart.series.length).fill(false);
  const percentTotalsBySeries = new Array<number[] | null>(chart.series.length).fill(null);
  const plottedValues = chart.series.map(() => new Array<number>(n).fill(0));
  const allPercentByAxis = new Map<string, boolean>();
  for (const group of lineGroups) {
    const axis = group.valueAxis;
    allPercentByAxis.set(
      axis,
      (allPercentByAxis.get(axis) ?? true) && group.grouping === 'percentStacked',
    );
  }
  for (const group of lineGroups) {
    const grouping = group.grouping ?? 'standard';
    const stacked = grouping === 'stacked' || grouping === 'percentStacked';
    const pct = grouping === 'percentStacked';
    const members = chart.series.slice(group.seriesStart, group.seriesStart + group.seriesCount);
    const percentMultiplier = pct && allPercentByAxis.get(group.valueAxis) === true ? 100 : 1;
    const totals = pct
      ? cats.map((_, categoryIndex) => members.reduce(
          (sum, series) => sum + Math.abs(series.values[categoryIndex] ?? 0), 0,
        ) || 1)
      : null;
    for (let offset = 0; offset < members.length; offset++) {
      const seriesIndex = group.seriesStart + offset;
      stackedBySeries[seriesIndex] = stacked;
      percentBySeries[seriesIndex] = pct;
      secondaryBySeries[seriesIndex] = group.valueAxis === 'secondary';
      percentTotalsBySeries[seriesIndex] = totals;
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        const raw = members[offset].values[categoryIndex] ?? 0;
        if (!stacked) {
          plottedValues[seriesIndex][categoryIndex] = raw;
          continue;
        }
        const prior = offset === 0 ? 0 : plottedValues[seriesIndex - 1][categoryIndex];
        const contribution = pct && totals
          ? raw / totals[categoryIndex] * percentMultiplier
          : raw;
        plottedValues[seriesIndex][categoryIndex] = prior + contribution;
      }
    }
  }
  const plotted = (seriesIndex: number, categoryIndex: number): number =>
    plottedValues[seriesIndex]?.[categoryIndex] ?? 0;
  const primaryGroups = lineGroups.filter(group => group.valueAxis !== 'secondary');
  const axisIsPercent = primaryGroups.length > 0
    && primaryGroups.every(group => group.grouping === 'percentStacked');
  // How null cells are plotted (`<c:dispBlanksAs>`, §21.2.2.42). Default "gap"
  // preserves the historical line break (byte-stable). "zero" treats a null as
  // 0; "span" bridges the neighbours with a straight line (skip the null but
  // keep the run going). Only unstacked charts see nulls — a stacked sum already
  // reads null as 0 — so the value only steers the unstacked path below.
  const dispBlanks = chart.dispBlanksAs ?? 'gap';

  // Combo line charts may bind some series to a SECONDARY value axis drawn on
  // the right (ECMA-376 §21.2.2.* — a second `<c:valAx>` with axPos="r"). `sec`
  // is non-null only when the axis is declared AND at least one series opts in;
  // secondary series are then excluded from the PRIMARY scale and mapped through
  // the secondary one. Stacked line charts stack ALL series onto the primary
  // axis (a percentStacked/stacked secondary combo is not an Office construct),
  // so the split only applies to plain (unstacked) line charts. When `sec` is
  // null every series stays on the primary axis, identical to the pre-CH7 path.
  const secondaryGroups = lineGroups.filter(group => group.valueAxis === 'secondary');
  const secondaryAxisIsPercent = secondaryGroups.length > 0
    && secondaryGroups.every(group => group.grouping === 'percentStacked');
  const sec = chart.secondaryValAxis && chart.series.some(
    (series, index) => secondaryBySeries[index] || (
      chart.plotGroups == null && series.useSecondaryAxis === true
    ),
  )
    ? chart.secondaryValAxis
    : null;
  const seriesIndexByIdentity = new Map(chart.series.map((series, index) => [series, index]));
  const isSecondarySeries = (series: ChartSeries): boolean => {
    const index = seriesIndexByIdentity.get(series) ?? -1;
    return sec != null && (secondaryBySeries[index]
      || (chart.plotGroups == null && series.useSecondaryAxis === true));
  };

  // Resolve the primary extent before frame placement. An authored
  // `layoutTarget="outer"` rectangle includes the value-axis labels, so its
  // conversion to the inner plot rectangle needs the width of the formatted
  // tick labels. This is the same extent used again for the final scale below.
  let dataMin = Infinity; let dataMax = -Infinity;
  for (let ci = 0; ci < n; ci++) {
    for (let si = 0; si < chart.series.length; si++) {
      if (isSecondarySeries(chart.series[si])) continue;
      if (!stackedBySeries[si] && chart.series[si].values[ci] == null) continue;
      const v = plotted(si, ci);
      dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v);
    }
  }
  for (let si = 0; si < chart.series.length; si++) {
    const series = chart.series[si];
    if (isSecondarySeries(series)) continue;
    forEachErrorBarEndpoint(
      series,
      'y',
      index => series.values[index] == null ? null : plotted(si, index),
      value => {
        dataMin = Math.min(dataMin, value);
        dataMax = Math.max(dataMax, value);
      },
    );
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  const isLogAxis = chart.valAxisLogBase != null && chart.valAxisLogBase >= 2;
  if (chart.valMin != null) dataMin = axisIsPercent ? chart.valMin * 100 : chart.valMin;
  else if (axisIsPercent && dataMin > 0 && !isLogAxis) dataMin = 0;
  if (chart.valMax != null) dataMax = axisIsPercent ? chart.valMax * 100 : chart.valMax;
  else if (axisIsPercent && dataMax < 0) dataMax = 0;

  // Shared frame bands. Title + category-label bands follow PowerPoint's chart
  // auto-layout (font-proportional, pinned to the demo slide-5 line-chart PDF);
  // see cartesianTitleBand / catAxisLabelBandH in layout.ts. The default 0.22
  // side-legend reserve is unchanged.
  let titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  let titleFontPx = titleBand.fontPx;
  let titleTopPad = titleBand.topPad;
  let titleH = titleBand.bandH;
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  // Axis-title bands use the real title font (XML @sz when set), independent of
  // the tick-label sizes above, so 18pt titles get a wide enough gutter.
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  // Vertical pads (independent of the right gutter) so an estimated plot height
  // is known before the secondary-axis scale + right-gutter measurement — the
  // same up-front ordering the bar renderer uses. The top adds half a value-axis
  // label so the topmost gridline label rides above the plot; the bottom reserves
  // PowerPoint's full category-label band (gap + line-height + margin).
  let padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;
  const phEst = h - padT - padB;

  // Secondary value-axis scale (shared helper). Its axis is the vertical right
  // edge, so its length is the plot height. Null when there is no secondary axis.
  const secScale = computeSecondaryAxis(
    sec,
    chart.series,
    phEst / ptToPx,
    'y',
    secondaryAxisIsPercent,
    false,
    (_series, index) => secondaryBySeries[index]
      || (chart.plotGroups == null && chart.series[index].useSecondaryAxis === true),
    (series, pointIndex, seriesIndex) => !stackedBySeries[seriesIndex]
      && series.values[pointIndex] == null
      ? null : plotted(seriesIndex, pointIndex),
  );
  // Right-edge gutter for the secondary tick labels + rotated title. Measured
  // with the SAME font/format the axis is drawn with so the reserve matches the
  // painted labels (mirrors the bar renderer). Zero when there is no secondary
  // axis, so `pad.r` is unchanged on the common single-axis path.
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const prevFont = ctx.font;
    ctx.font = chartFontCss(
      secFontPx,
      chartFontFamily(chart, sec.fontFace, 'minor'),
      false,
      sec.fontItalic ?? false,
    );
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

  const titleLeftBandW = legLeftW + Math.max(
    valAxFontPx * 2.2 + 10 + valTitleW,
    dataTableHeaderW,
  );
  const titleRightBandW = legRightW + w * 0.05 + secLabelBandW + secTitleBandW;

  const provisionalPlan = planValueAxis(chart, dataMin, dataMax, phEst / ptToPx, axisIsPercent);
  let primaryLabelWidth = 0;
  if (
    !chart.valAxisHidden
    && chart.valAxisTickLabelPos !== 'none'
    && chart.plotAreaManualLayout != null
    && chart.plotAreaManualLayout.layoutTarget !== 'inner'
  ) {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    for (const value of provisionalPlan.majorLines) {
      primaryLabelWidth = Math.max(
        primaryLabelWidth,
        ctx.measureText(formatPrimaryValueAxisTick(chart, value, axisIsPercent)).width,
      );
    }
    ctx.font = previousFont;
  }
  // Pad based on actual label metrics rather than magic percents so an explicit
  // <c:txPr sz="1000"> (10pt) correctly compresses the plot area.
  const pad = {
    t: padT,
    r: titleRightBandW,
    b: padB,
    l: titleLeftBandW,
  };

  const manualOuterInsets = chartManualOuterAxisInsets({
    valAxisHidden: chart.valAxisHidden,
    catAxisHidden: chart.catAxisHidden,
    valLabelWidth: primaryLabelWidth,
    valLabelFontPx: valAxFontPx,
    catLabelFontPx: catAxFontPx,
    valLabelGapPx: chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(valAxFontPx)
      : 6,
    catLabelGapPx: chart.catAxisFontSizeHpt != null
      ? categoryLabelOffsetPx(
        categoryTickLabelGapPx(catAxFontPx),
        chart.catAxisLabelOffsetPercent,
      )
      : categoryLabelOffsetPx(5, chart.catAxisLabelOffsetPercent),
    outerTextMarginPx: AXIS_OUTER_TEXT_MARGIN_PT * ptToPx,
    valTitleBandW: valTitleW,
    catTitleBandH: catTitleH,
    secondaryBandW: secLabelBandW + secTitleBandW,
  });

  let lineFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
    manualOuterInsets,
  });
  // The automatic title box follows the inner plot width. Resolve that width
  // from the same frame that owns the axis/legend gutters, then run one stable
  // vertical-layout pass if wrapping adds title lines.
  const plotWidthTitleBand = measuredCartesianTitleBand(
    ctx,
    chart,
    lineFrame.plotRect.pw,
    h,
    ptToPx,
  );
  if (Math.abs(plotWidthTitleBand.bandH - titleBand.bandH) > 0.01) {
    titleBand = plotWidthTitleBand;
    titleFontPx = titleBand.fontPx;
    titleTopPad = titleBand.topPad;
    titleH = titleBand.bandH;
    padT = titleH + legTopH + valAxFontPx / 2 + 2;
    pad.t = padT;
    lineFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
      titleBand,
      legendSideReserveFrac: 0.22,
      legendReserve: leg,
      pad,
      honorPlotAreaManualLayout: true,
      manualOuterInsets,
    });
  }
  const { px0, py0, pw } = lineFrame.plotRect;
  let { ph } = lineFrame.plotRect;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleTopPad, titleFontPx,
  );
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // Value axis is vertical → its length is the plot height (axis-length-aware
  // auto major unit, same model as the bar/column renderer). `planValueAxis`
  // folds in the CH6 major unit / logBase / orientation; with none set it is
  // byte-identical to the old `valueAxisScale` + linear `toY`.
  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx, axisIsPercent);
  if (plan.max - plan.min === 0) return;

  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;
  // Secondary series map through their own scale; `secScale` is null on the
  // common single-axis path so `yMapFor` always returns the primary `toY`.
  const toYSecondary = secScale ? secScale.makeToY(py0, ph) : toY;
  const yMapFor = (s: ChartSeries): ((v: number) => number) =>
    isSecondarySeries(s) ? toYSecondary : toY;
  const primaryCategoryAxisY = toY(
    categoryAxisCrossingValue(chart, plan.min, plan.max),
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
  const primaryCatLine = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const primaryValLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  const primaryCatTickColor = chart.catAxisLineColor != null ? primaryCatLine.color : undefined;
  const primaryCatTickWidth = chart.catAxisLineWidthEmu != null ? primaryCatLine.width : undefined;
  const primaryValTickColor = chart.valAxisLineColor != null ? primaryValLine.color : undefined;
  const primaryValTickWidth = chart.valAxisLineWidthEmu != null ? primaryValLine.width : undefined;
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  // crossBetween="between" (default) insets the first/last category by half a
  // step so points aren't flush against the axes. "midCat" anchors them.
  // A `maxMin` category orientation (§21.2.2.130) mirrors the index left↔right.
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const toX = dateAxisPlan
    ? (i0: number) => px0 + dateAxisPlan.positions[i0]! * pw
    : between
      ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
      : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    // Resolved gridline stroke (`<c:majorGridlines><c:spPr><a:ln>` or default).
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    // Minor gridlines first (under the majors), then major gridlines + ticks +
    // labels. Minor lines are only populated when the file declares them.
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, primaryValTickColor, primaryValTickWidth, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        const gap = chart.valAxisFontSizeHpt != null
          ? valueTickLabelGapPx(valAxFontPx)
          : 6;
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, axisIsPercent), px0 - gap, gy);
      }
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of plan.minorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', px0, toY(value), primaryValTickColor, primaryValTickWidth, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }


  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Category-axis MAJOR gridlines (`<c:catAx><c:majorGridlines>`, §21.2.2.100):
  // vertical lines at the category ticks across the plot height. Off by default
  // (byte-stable). Shared placement with the bar renderer via
  // `categoryGridlineFractions`.
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

  // Axis lines: bottom (category) + left (value). Both default to visible
  // unless hidden explicitly. Office treats `<c:spPr><a:ln><a:noFill>` as
  // suppressing the rule and tick marks while retaining labels/gridlines.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, primaryCategoryAxisY, px0 + pw, primaryCategoryAxisY,
      primaryCatLine.color, primaryCatLine.width, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      primaryValLine.color, primaryValLine.width, chart.valAxisLineDash,
    );
  }

  // CT_LineChart owns drop lines, high-low lines, and up/down bars at the
  // group level (ECMA-376 §21.2.2 EG_LineChartShared / CT_LineChart). Paint
  // this background geometry before the series so authored lines and markers
  // remain on top, as in Office. Group provenance prevents a second line group
  // in a combo chart from inheriting the first group's decorations.
  const decorationSlotWidth = dateAxisPlan
    ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw
    : between ? pw / n : n > 1 ? pw / (n - 1) : pw;
  const lineSeriesIndex = new Map(
    chart.series.map((series, seriesIndex) => [series, seriesIndex]),
  );
  drawLineGroupDecorations(
    ctx, chart, n, toX, yMapFor,
    categoryAxisYFor,
    (series, index) => {
      const seriesIndex = lineSeriesIndex.get(series);
      return seriesIndex != null ? plotted(seriesIndex, index) : null;
    },
    decorationSlotWidth, ptToPx, shapeRotationDeg, 'background',
  );

  // Line width and marker size come from OOXML in points (<a:ln w=EMU> /
  // <c:marker><c:size val=pt>). Omitted series strokes keep the PowerPoint
  // defaults (2.25pt line, 5pt marker diameter) scaled to the viewport.
  const lineWidthPx = Math.max(1, 2.25 * ptToPx);
  const markerR = Math.max(2, 2.5 * ptToPx);
  const dataLabelPx = axisLabelPx(chart.dataLabelFontSizeHpt, h, ptToPx);
  // Data labels are a chart-wide foreground layer. Painting them inside the
  // per-series loop lets a later series line (or trendline) cross labels that
  // belong to an earlier series. Excel keeps every series label above all
  // series geometry, so collect the label painters and flush them only after
  // every line, error bar, marker, and trendline has been painted.
  const deferredDataLabels: Array<() => void> = [];
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const seriesStacked = stackedBySeries[si];
    const seriesPercentTotals = percentTotalsBySeries[si];
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const color = chartColor(si, s);
    const styleIndex = chartExSeriesFormatIndex(s, si);
    // Secondary series ride their own vertical scale; primary series (and every
    // series when there is no secondary axis) map through the primary `toY`.
    const yOf = yMapFor(s);
    const smooth = s.smooth === true;
    const runs: IndexedLinePoint[][] = [];
    let run: IndexedLinePoint[] = [];
    const flushRun = (): void => {
      if (run.length > 0) runs.push(run);
      run = [];
    };
    for (let ci = 0; ci < n; ci++) {
      if (s.sourceHidden?.[ci] === true) {
        flushRun();
        continue;
      }
      if (!seriesStacked && s.values[ci] == null) {
        if (dispBlanks === 'gap') { flushRun(); continue; }
        if (dispBlanks === 'span') continue;
      }
      run.push({ x: toX(ci), y: yOf(plotted(si, ci)), index: ci });
    }
    flushRun();
    const lineBounds = { x: px0, y: py0, w: pw, h: ph };
    if (chartSeriesVariesByPoint(chart, si)) {
      paintClassicVaryingLineSegments(
        ctx, chart, s, runs, smooth, false, color, lineWidthPx,
        ptToPx, lineBounds, shapeRotationDeg,
      );
    } else {
      const paintSeriesLine = (target: CanvasRenderingContext2D): void => {
        if (!applyClassicStyleLine(
          target, chart, 'dataPointLine', s, undefined, styleIndex, color,
          lineWidthPx, ptToPx, lineBounds, shapeRotationDeg,
        )) return;
        target.beginPath();
        for (const points of runs) {
          if (points.length === 0) continue;
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
        lineBounds,
        ptToPx,
        paintSeriesLine,
      );
    }

    // Error bars (`<c:errBars>`, §21.2.2.20) — drawn under the markers so the
    // dots overlay the bar tips. Only fires for series that carry them.
    const plottedOf = (ci: number): number => plotted(si, ci);
    for (const eb of s.errBars ?? []) {
      drawCategoryErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), n, toX, yOf, plottedOf, color,
      );
    }

    ctx.fillStyle = color;
    // ECMA-376 §21.2.2.32 — when the series resolves to no marker, skip the
    // data-point dots but keep data labels. Markers / labels pin to the plotted
    // (cumulative) value so they ride the stacked line, not the raw datum.
    const seriesMarkersVisible = s.showMarker !== false && s.markerSymbol !== 'none';
    const drawMarkers = seriesMarkersVisible || hasVisiblePointMarkerOverride(s);
    // Series carrying explicit `<c:marker>` detail route through drawMarker
    // (symbol/size/fill/line + per-point `<c:dPt>` overrides). Series without
    // any detail keep the historical fixed-circle fast path unchanged
    // (byte-stable). `markerSymbol: "none"` is caught by the showMarker gate.
    const hasMarkerDetail = seriesHasResolvedMarkerDetail(chart, s, si);
    // Per-point / series-level data labels (`<c:dLbl idx>` / `<c:dLbls>`) take
    // precedence over the family's simple `showDataLabels` value dump. Merely
    // decide the route here; painting is deferred until all series geometry is
    // complete so labels remain the chart-wide foreground layer.
    const perPointLabels = (s.dataLabelOverrides?.length ?? 0) > 0 || s.seriesDataLabels != null;
    if (perPointLabels) {
      deferredDataLabels.push(() => {
        drawCategoryDataLabels(
          ctx, s, cats, n, toX, yOf, plottedOf, ph, ptToPx, chart.date1904 ?? false,
          // Mirror the marker loop's gate just below: stacked series never see a
          // plotted null (a stacked sum already reads null as 0), and unstacked
          // "zero" mode plots the null at 0 — both cases get a label too.
          seriesStacked || dispBlanks === 'zero',
          chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
          // §21.2.2.48 `<c:dLblPos>` precedence: per-point/series positions win,
          // else the chart-level position, else the line-chart default `'r'`.
          chart.dataLabelPosition ?? 'r',
          // Automatic endpoint labels may occupy the chart gutter outside the
          // plot area (the plot's manual layout often reserves that space).
          // Keep vertical clipping aligned to the plot, but clamp horizontally
          // to the chart rectangle so `l`/`r` remain outside the end markers.
          { x, y: py0, w, h: ph },
          { x, y, w, h },
          percentBySeries[si] && seriesPercentTotals
            ? ci => (s.values[ci] ?? 0) / seriesPercentTotals[ci]
            : undefined,
          ci => {
            if (!drawMarkers) return 0;
            const dpt = pointOverrides.get(ci);
            if (!hasMarkerDetail && !pointHasMarkerDetail(dpt)) return markerR;
            const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
            if (symbol === 'none') return 0;
            return ((dpt?.markerSize ?? s.markerSize ?? 5) / 2) * ptToPx;
          },
          face => chartFontFamily(chart, face, 'minor'),
          isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
          ci => dataLabelLegendKey(si, ci),
          value => dataLabelWithinAxisMaximum(
            chart, value,
            isSecondarySeries(s) && secScale ? secScale.max : plan.max,
          ),
          shapeRotationDeg,
        );
      });
    }
    for (let ci = 0; ci < n; ci++) {
      if (s.sourceHidden?.[ci] === true) continue;
      // A null point gets a marker/label only in "zero" mode (plotted at 0);
      // "gap"/"span" leave the hole empty.
      if (!seriesStacked && s.values[ci] == null && dispBlanks !== 'zero') continue;
      const pv = plotted(si, ci);
      if (drawMarkers) {
        const dpt = pointOverrides.get(ci);
        if (hasMarkerDetail || pointHasMarkerDetail(dpt)) {
          const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkersVisible);
          if (symbol !== 'none') {
            const sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
            const fill = markerFillColorFor(s, dpt, ci, color);
            const line = dpt?.markerLine ?? s.markerLine ?? null;
            const lineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
            drawChartMarker(
              ctx, chart, s, dpt, ci, toX(ci), yOf(pv), symbol, sizePt, fill, line, ptToPx,
              lineWidthEmu != null ? axisLineWidthPx(lineWidthEmu, ptToPx) : undefined,
              markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
            );
          }
        } else {
          ctx.beginPath(); ctx.arc(toX(ci), yOf(pv), markerR, 0, Math.PI * 2); ctx.fill();
        }
      }
    }

    if (chart.showDataLabels && !perPointLabels) {
      deferredDataLabels.push(() => {
        for (let ci = 0; ci < n; ci++) {
          if (s.sourceHidden?.[ci] === true) continue;
          if (!seriesStacked && s.values[ci] == null && dispBlanks !== 'zero') continue;
          const pv = plotted(si, ci);
          // §21.2.2.48 `<c:dLblPos>`: the family-level value dump honors the
          // chart-level position (else the line default `'r'`). The marker gap
          // stays directional while the whole label layer is painted last.
          const labelText = effectiveDataLabelText({
            showValue: true,
            sourceValue: s.values[ci] ?? 0,
            valueDivisor: displayUnitDivisor(
              isSecondarySeries(s) ? sec?.displayUnits : chart.valAxisDisplayUnits,
            ),
            formatCode: chart.dataLabelFormatCode ?? s.valFormatCode ?? null,
            date1904: chart.date1904,
          });
          drawDataLabelText(
            ctx, toX(ci), yOf(pv), labelText,
            chart.dataLabelPosition ?? 'r', dataLabelPx,
            chart.dataLabelFontColor ?? undefined, chart.dataLabelFontBold ?? false,
            chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
            drawMarkers ? markerR + 1 : 2,
            { x: px0, y: py0, w: pw, h: ph },
          );
        }
      });
    }

    // Trendlines (`<c:trendline>`, §21.2.2.211) over this series' points —
    // drawn on top of the line/markers, dashed, in the series color unless the
    // trendline declares its own `<a:ln>`.
    drawSeriesTrendlines(
      ctx, s, color, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  drawLineGroupDecorations(
    ctx, chart, n, toX, yMapFor,
    categoryAxisYFor,
    (series, index) => {
      const seriesIndex = lineSeriesIndex.get(series);
      return seriesIndex != null ? plotted(seriesIndex, index) : null;
    },
    decorationSlotWidth, ptToPx, shapeRotationDeg, 'foreground',
  );

  for (const drawLabels of deferredDataLabels) drawLabels();

  if (!chart.catAxisHidden) {
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      catAxFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    // Tick marks and labels have independent authored skip intervals
    // (§21.2.2.205/§21.2.2.206). When tickLblSkip is absent every non-empty
    // cached category is paintable; sparse caches deliberately use blank
    // indices to author intervals such as every second year.
    const tickInterval = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    const majorTickXs = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => px0 + tick.fraction * pw)
      : Array.from({ length: Math.ceil(n / tickInterval) }, (_, index) => toX(index * tickInterval));
    for (const tx of majorTickXs) {
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', primaryCategoryAxisY, tx, primaryCatTickColor, primaryCatTickWidth, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none' && dateAxisPlan) {
      for (const tick of dateAxisPlan.minorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMinorTickMark, 'cat', primaryCategoryAxisY,
          px0 + tick.fraction * pw, primaryCatTickColor, primaryCatTickWidth,
          false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
        );
      }
    }
    const showLabels = !hasDataTable && catLabelsVisible(chart);
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const rotRad = catLabelRotationRad(chart);
    const labelEntries = dateAxisPlan
      ? dateAxisPlan.majorTicks.map(tick => ({
        label: formatCategoryLabel(String(tick.serial), chart.catAxisFormatCode, chart.date1904),
        x: px0 + tick.fraction * pw,
        categoryIndex: -1,
      }))
      : Array.from({ length: Math.ceil(n / labelInterval) }, (_, index) => {
        const ci = index * labelInterval;
        return {
          label: formatCategoryLabel((cats[ci] ?? '').toString(), chart.catAxisFormatCode, chart.date1904),
          x: toX(ci),
          categoryIndex: ci,
        };
      });
    for (const entry of labelEntries) {
      const anchor = entry.categoryIndex < 0
        ? null
        : categoryLabelAnchorFraction(
          entry.categoryIndex,
          n,
          isCrossBetween(chart),
          catAxisReversed(chart),
          chart.catAxisLabelAlignment,
        );
      const tx = anchor ? px0 + anchor.fraction * pw : entry.x;
      if (!showLabels) continue;
      ctx.textAlign = anchor?.textAlign ?? 'center';
      ctx.fillStyle = catLabelColor;
      // §21.2.2.71: format numeric-serial categories (e.g. dateAx) via the
      // category-axis numFmt; string categories pass through unchanged.
      const label = entry.label;
      if (!label) continue;
      const gap = categoryLabelOffsetPx(
        chart.catAxisFontSizeHpt != null
          ? categoryTickLabelGapPx(catAxFontPx)
          : 5,
        chart.catAxisLabelOffsetPercent,
      );
      const labelPosition = chart.catAxisTickLabelPos ?? 'nextTo';
      const labelAxisY = labelPosition === 'nextTo'
        ? primaryCategoryAxisY
        : labelPosition === 'high' ? py0 : py0 + ph;
      drawRotatedCatLabel(ctx, label, tx, labelAxisY + gap, rotRad);
    }
  }

  // Secondary value axis (right edge) — drawn after the series + category labels
  // so it sits atop the plot, mirroring the bar renderer's ordering.
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
