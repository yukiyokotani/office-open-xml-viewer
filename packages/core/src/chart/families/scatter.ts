// Classic scatter chart family.
import type { ChartDisplayUnits, ChartModel, ChartRect, ChartSeries } from '../../types/chart';

import {
  computeChartFrame,
  catAxisLabelBandH,
  chartLegendBands,
  chartAxisTitleBands,
  chartTextFontSizePx,
  categoryTickLabelGapPx,
  valueTickLabelGapPx,
} from '../layout.js';
import { planNumericValueAxis, finiteDataExtent } from '../axis-scale.js';
import { axisLineWidthPx, resolveAxisLine } from '../axis-style.js';

import { indexChartPlotGroups } from '../plot-groups.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';

import { chartFontFamily, chartFontCss } from '../shared/fonts.js';
import { drawAxisTitles } from '../shared/axis.js';
import { measuredLegendReserve, drawLegendForLayout } from '../shared/legend.js';
import { drawAxisTick, strokeAxisSegment, axisTickOutwardExtentPx, strokeValueGridlineH, valGridStroke, valMinorGridStroke, drawCatMajorGridlines, catGridStroke, catMinorGridStroke, valAxisReversed, catAxisReversed, drawValMajorGridlines, formatPrimaryValueAxisTick, formatAxisTickWithUnits, axisLabelPx } from '../shared/axis.js';
import { forEachErrorBarEndpoint, drawSecondaryValueAxis } from '../shared/secondary-axis.js';
import type { SecondaryAxisScale } from '../shared/secondary-axis.js';
import { measuredCartesianTitleBand, drawChartTitleForLayout } from '../shared/title.js';
import { scatterXValue } from '../shared/scatter-geometry.js';
import type { BubbleGroupSettings } from '../shared/scatter-geometry.js';
import { drawScatterSeriesLayer } from '../shared/scatter-paint.js';
import { clamp } from '../shared/geometry.js';

export function renderScatterChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const entries = chart.series.map((series, index) => ({ series, index }));
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const usesSecondaryX = ({ series, index }: (typeof entries)[number]): boolean =>
    plotGroupBySeries[index]?.categoryAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  const usesSecondaryY = ({ series, index }: (typeof entries)[number]): boolean =>
    plotGroupBySeries[index]?.valueAxis === 'secondary'
      || (chart.plotGroups == null && series.useSecondaryAxis === true);
  const primaryXEntries = entries.filter(entry => !usesSecondaryX(entry));
  const secondaryXEntries = entries.filter(usesSecondaryX);
  const primaryYEntries = entries.filter(entry => !usesSecondaryY(entry));
  const secondaryYEntries = entries.filter(usesSecondaryY);
  const primaryEntries = entries.filter(entry => !usesSecondaryX(entry) && !usesSecondaryY(entry));
  const secondaryEntries = entries.filter(entry => usesSecondaryX(entry) && usesSecondaryY(entry));
  const secondaryX = secondaryXEntries.length > 0 ? chart.secondaryCatAxis : null;
  const secondaryY = secondaryYEntries.length > 0 ? chart.secondaryValAxis : null;

  const numericXValues = (entries: Array<{ series: ChartSeries; index: number }>): number[] => {
    const values: number[] = [];
    for (const { series } of entries) {
      const cats = series.categories ?? chart.categories;
      for (const category of cats) {
        const value = parseFloat(category);
        if (Number.isFinite(value)) values.push(value);
      }
    }
    return values;
  };
  const allNumericX = numericXValues(entries);
  const useIndexX = allNumericX.length === 0;
  const textBubbleOrdinalMax = entries.length === 1
    && entries[0].series.bubbleXSourceIsString === true
    ? entries[0].series.values.length + 1
    : null;
  const pairedExtents = (
    entries: Array<{ series: ChartSeries; index: number }>,
  ): { x: { min: number; max: number }; y: { min: number; max: number } } => {
    const xs: number[] = [];
    const ys: number[] = [];
    for (const { series } of entries) {
      const cats = series.categories ?? chart.categories;
      for (let index = 0; index < series.values.length; index++) {
        const yValue = series.values[index];
        if (yValue == null) continue;
        const xValue = scatterXValue(cats, index, useIndexX);
        if (xValue == null) continue;
        xs.push(xValue);
        ys.push(yValue);
      }
      forEachErrorBarEndpoint(
        series,
        'x',
        index => series.values[index] == null ? null : scatterXValue(cats, index, useIndexX),
        value => xs.push(value),
      );
      forEachErrorBarEndpoint(
        series,
        'y',
        index => {
          const xValue = scatterXValue(cats, index, useIndexX);
          return xValue == null ? null : series.values[index] ?? null;
        },
        value => ys.push(value),
      );
    }
    if (useIndexX && xs.length === 0) {
      let count = 0;
      for (const { series } of entries) count = Math.max(count, series.values.length);
      for (let index = 0; index < count; index++) xs.push(index);
    }
    return { x: finiteDataExtent(xs), y: finiteDataExtent(ys) };
  };
  const primaryExtent = {
    x: pairedExtents(primaryXEntries.length > 0 ? primaryXEntries : secondaryXEntries).x,
    y: pairedExtents(primaryYEntries.length > 0 ? primaryYEntries : secondaryYEntries).y,
  };
  const secondaryExtent = {
    x: pairedExtents(secondaryXEntries).x,
    y: pairedExtents(secondaryYEntries).y,
  };
  // Shared frame bands. Title + bottom axis-label bands follow PowerPoint's
  // chart auto-layout (font-proportional, pinned to the demo slide-5 line-chart
  // PDF); see cartesianTitleBand / catAxisLabelBandH in layout.ts. Scatter's X
  // axis is a numeric value axis, so the bottom band holds its single line of
  // X-value labels (sized like any value-axis label). Default 0.22 side-legend
  // reserve unchanged.
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const xAxLabelFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const yAxLabelFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;

  // Title placement — manual layout overrides the auto position.
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  // Plot area placement: honor `<c:plotArea><c:manualLayout>` when present.
  // ECMA-376: layoutTarget="inner" (default) describes the inner plot rect
  // (no axes / labels); "outer" includes axes. For scatter we treat both
  // identically (the inner padding stays the same). The pad is pure arithmetic
  // and is ignored by computeChartFrame when the manual layout applies.
  const provisionalSecondaryY = secondaryY
    ? planNumericValueAxis({
        dataMin: secondaryExtent.y.min,
        dataMax: secondaryExtent.y.max,
        explicitMin: secondaryY.min,
        explicitMax: secondaryY.max,
        axisLenPt: Math.max(1, h * 0.7 / ptToPx),
        axisOrientation: 'vertical',
        majorUnit: secondaryY.majorUnit,
        minorUnit: secondaryY.minorUnit,
        needMinor: secondaryY.minorGridlines === true
          || (secondaryY.minorTickMark != null && secondaryY.minorTickMark !== 'none'),
        logBase: secondaryY.logBase,
        reversed: secondaryY.orientation === 'maxMin',
      })
    : null;
  let secondaryYLabelWidth = 0;
  if (secondaryY && provisionalSecondaryY && !secondaryY.hidden && secondaryY.tickLabelPos !== 'none') {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      chartTextFontSizePx(secondaryY.fontSizeHpt, ptToPx) ?? yAxLabelFontPx,
      chartFontFamily(chart, secondaryY.fontFace, 'minor'),
      secondaryY.fontBold ?? false,
      secondaryY.fontItalic ?? false,
    );
    for (const value of provisionalSecondaryY.majorTicks) {
      secondaryYLabelWidth = Math.max(
        secondaryYLabelWidth,
        ctx.measureText(formatAxisTickWithUnits(value, secondaryY.formatCode, chart.date1904, secondaryY.displayUnits)).width,
      );
    }
    secondaryYLabelWidth += valueTickLabelGapPx(yAxLabelFontPx) + 4;
    ctx.font = previousFont;
  }
  const secondaryXLabelHeight = secondaryX && !secondaryX.hidden && secondaryX.tickLabelPos !== 'none'
    ? (chartTextFontSizePx(secondaryX.fontSizeHpt, ptToPx) ?? xAxLabelFontPx)
      + categoryTickLabelGapPx(xAxLabelFontPx) + 2
    : 0;
  const pad = {
    t: titleBand.bandH + legTopH + yAxLabelFontPx / 2 + 2 + secondaryXLabelHeight,
    r: legRightW + w * 0.05 + secondaryYLabelWidth,
    b: (chart.catAxisHidden ? h * 0.04 : catAxisLabelBandH(xAxLabelFontPx)) + catTitleH + legBottomH,
    l: (chart.valAxisHidden ? w * 0.04 : w * 0.12) + valTitleW + legLeftW,
  };
  const { plotRect: { px0, py0, pw, ph } } = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  if (pw <= 0 || ph <= 0) return;

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // X / Y data extents. Secondary-group points have their own independent
  // top/right value axes and therefore must not stretch the primary scales.
  let { min: xMin, max: xMax } = primaryExtent.x;
  let { min: yMin, max: yMax } = primaryExtent.y;
  // Apply explicit `<c:valAx><c:scaling><c:min/max>` and `<c:catAx>` scaling.
  // Omitted Y bounds, including point ranges, flow through the shared planner.
  if (chart.valMin != null) yMin = chart.valMin;
  if (chart.valMax != null) yMax = chart.valMax;
  const yNeedsMinor = chart.valAxisMinorGridlines === true
    || (chart.valAxisMinorTickMark != null && chart.valAxisMinorTickMark !== 'none');
  const yAxisPlan = planNumericValueAxis({
    dataMin: yMin,
    dataMax: yMax,
    explicitMin: chart.valMin,
    explicitMax: chart.valMax,
    axisLenPt: ph / ptToPx,
    axisOrientation: 'vertical',
    majorUnit: chart.valAxisMajorUnit,
    minorUnit: chart.valAxisMinorUnit,
    needMinor: yNeedsMinor,
    logBase: chart.valAxisLogBase,
    reversed: valAxisReversed(chart),
  });
  yMin = yAxisPlan.min;
  yMax = yAxisPlan.max;
  const xNeedsMinor = chart.catAxisMinorGridlines === true
    || (chart.catAxisMinorTickMark != null && chart.catAxisMinorTickMark !== 'none');
  const xAxisPlan = planNumericValueAxis({
    dataMin: xMin,
    dataMax: xMax,
    // Office gives a string-backed lone bubble series one empty ordinal slot
    // on each side (four points => 0..5). Authored axis bounds still win.
    explicitMin: chart.catAxisMin ?? (textBubbleOrdinalMax == null ? null : 0),
    explicitMax: chart.catAxisMax ?? textBubbleOrdinalMax,
    axisLenPt: pw / ptToPx,
    axisOrientation: 'horizontal',
    majorUnit: chart.catAxisMajorUnit,
    minorUnit: chart.catAxisMinorUnit,
    needMinor: xNeedsMinor,
    logBase: chart.catAxisLogBase,
    reversed: catAxisReversed(chart),
  });
  xMin = xAxisPlan.min;
  xMax = xAxisPlan.max;

  const secondaryXPlan = secondaryX
    ? planNumericValueAxis({
        dataMin: secondaryExtent.x.min,
        dataMax: secondaryExtent.x.max,
        explicitMin: secondaryX.min,
        explicitMax: secondaryX.max,
        axisLenPt: pw / ptToPx,
        axisOrientation: 'horizontal',
        majorUnit: secondaryX.majorUnit,
        minorUnit: secondaryX.minorUnit,
        needMinor: secondaryX.minorGridlines === true
          || (secondaryX.minorTickMark != null && secondaryX.minorTickMark !== 'none'),
        logBase: secondaryX.logBase,
        reversed: secondaryX.orientation === 'maxMin',
      })
    : null;
  const secondaryYPlan = secondaryY
    ? planNumericValueAxis({
        dataMin: secondaryExtent.y.min,
        dataMax: secondaryExtent.y.max,
        explicitMin: secondaryY.min,
        explicitMax: secondaryY.max,
        axisLenPt: ph / ptToPx,
        axisOrientation: 'vertical',
        majorUnit: secondaryY.majorUnit,
        minorUnit: secondaryY.minorUnit,
        needMinor: secondaryY.minorGridlines === true
          || (secondaryY.minorTickMark != null && secondaryY.minorTickMark !== 'none'),
        logBase: secondaryY.logBase,
        reversed: secondaryY.orientation === 'maxMin',
      })
    : null;

  const toX = (v: number) => px0 + xAxisPlan.fraction(v) * pw;
  const toY = (v: number) => py0 + ph - yAxisPlan.fraction(v) * ph;
  const toSecondaryX = (v: number) => px0 + (secondaryXPlan?.fraction(v) ?? 0) * pw;
  const toSecondaryY = (v: number) => py0 + ph - (secondaryYPlan?.fraction(v) ?? 0) * ph;
  const xStep = xAxisPlan.majorUnit;
  const yMajorTicks = yAxisPlan.majorTicks;
  const yMinorTicks = yAxisPlan.minorTicks;
  const xMajorTicks = xAxisPlan.majorTicks;
  const xMinorTicks = xAxisPlan.minorTicks;

  // Each scatter axis is a numeric value axis. Its crossing coordinate comes
  // from the opposite axis's scale (§21.2.2.31 / §21.2.2.32): autoZero uses
  // zero when the range contains it, while min/max pin the rule to an edge.
  let xAxisY = py0 + ph;
  if (chart.catAxisCrossesAt != null) {
    xAxisY = clamp(toY(chart.catAxisCrossesAt), py0, py0 + ph);
  } else {
    const crosses = chart.catAxisCrosses ?? 'autoZero';
    if (crosses === 'autoZero' && yMin < 0 && yMax > 0) xAxisY = clamp(toY(0), py0, py0 + ph);
    else if (crosses === 'max') xAxisY = py0;
  }

  let yAxisX = px0;
  if (chart.valAxisCrossesAt != null) {
    yAxisX = clamp(toX(chart.valAxisCrossesAt), px0, px0 + pw);
  } else {
    const crosses = chart.valAxisCrosses ?? 'autoZero';
    if (crosses === 'autoZero' && xMin < 0 && xMax > 0) yAxisX = clamp(toX(0), px0, px0 + pw);
    else if (crosses === 'max') yAxisX = px0 + pw;
  }

  // Y-axis gridlines + labels + major tick marks. Scatter has no baseline
  // special-case, so it strokes every gridline in the resolved color/width.
  const grid = valGridStroke(chart, ptToPx);
  if (!chart.valAxisHidden) {
    const yTickFontPx = chart.valAxisFontSizeHpt != null
      ? axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx)
      : Math.max(8, Math.min(11, ph / 20));
    const yTickGap = chart.valAxisFontSizeHpt != null
      ? valueTickLabelGapPx(yTickFontPx)
      : 4;
    ctx.font = chartFontCss(
      yTickFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    const yAxisLineColor = chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : undefined;
    const yAxisLineWidth = axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx);
    const yMajorTickOutset = chart.valAxisLineHidden
      ? 0
      : axisTickOutwardExtentPx(chart.valAxisMajorTickMark, 'major', yAxisLineWidth, ptToPx);
    if (chart.valAxisMinorGridlines) {
      const minorGrid = valMinorGridStroke(chart, ptToPx);
      for (const value of yMinorTicks) {
        strokeValueGridlineH(ctx, px0, pw, toY(value), false, minorGrid);
      }
    }
    for (const v of yMajorTicks) {
      const gy = toY(v);
      ctx.strokeStyle = grid.color; ctx.lineWidth = grid.width;
      if (drawValMajorGridlines(chart)) {
        const previousDash = grid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
        if (grid.dash.length > 0) ctx.setLineDash(grid.dash);
        ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
        if (grid.dash.length > 0) ctx.setLineDash(previousDash);
      }
      if (chart.valAxisTickLabelPos !== 'none') {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        const labelPos = chart.valAxisTickLabelPos ?? 'nextTo';
        let labelX: number;
        if (labelPos === 'high') {
          ctx.textAlign = 'left'; labelX = px0 + pw + yTickGap;
        } else if (labelPos === 'low') {
          ctx.textAlign = 'right'; labelX = px0 - yTickGap;
        } else {
          ctx.textAlign = 'right'; labelX = yAxisX - yMajorTickOutset - yTickGap;
        }
        ctx.textBaseline = 'middle';
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), labelX, gy);
      }
      // Scatter keeps its own undefined colour default (→ drawAxisTick's '#888'),
      // so only the width formula is shared. `axisLineWidthPx`'s 1 px fallback is
      // equivalent to undefined here (drawAxisTick treats both as a hairline).
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', yAxisX, gy, yAxisLineColor, yAxisLineWidth, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
    }
    if (chart.valAxisMinorTickMark && chart.valAxisMinorTickMark !== 'none') {
      for (const value of yMinorTicks) {
        drawAxisTick(ctx, chart.valAxisMinorTickMark, 'val', yAxisX, toY(value), yAxisLineColor, yAxisLineWidth, false, chart.valAxisLineHidden, 'minor', ptToPx, chart.valAxisLineDash);
      }
    }
  }

  // A scatter chart's horizontal axis is represented by the shared category-
  // axis fields in the model even though OOXML stores it as a second valAx.
  // Its major gridlines therefore run vertically through each numeric X tick.
  if (!chart.catAxisHidden && drawCatMajorGridlines(chart) && xStep > 0) {
    const xGrid = catGridStroke(chart, ptToPx);
    ctx.strokeStyle = xGrid.color;
    ctx.lineWidth = xGrid.width;
    const previousDash = xGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (xGrid.dash.length > 0) ctx.setLineDash(xGrid.dash);
    for (const v of xMajorTicks) {
      const gx = toX(v);
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (xGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }
  if (!chart.catAxisHidden && chart.catAxisMinorGridlines && xStep > 0) {
    const xGrid = catMinorGridStroke(chart, ptToPx);
    const previousDash = xGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    ctx.strokeStyle = xGrid.color;
    ctx.lineWidth = xGrid.width;
    if (xGrid.dash.length > 0) ctx.setLineDash(xGrid.dash);
    for (const value of xMinorTicks) {
      const gx = toX(value);
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
    }
    if (xGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }

  // X-axis line (the timeline ruler in Gantt-style scatter charts depends
  // on this line's stroke). Tick labels are skipped when the category axis
  // is hidden via `<c:delete val="1"/>`. Office treats
  // `<c:catAx><c:spPr><a:ln><a:noFill>` as suppressing the rule and tick
  // marks. Color and weight come from
  // `<c:catAx><c:spPr><a:ln>` when present; default otherwise.
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    ctx.save();
    ctx.lineCap = 'butt';
    strokeAxisSegment(
      ctx, px0, xAxisY, px0 + pw, xAxisY,
      chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#888',
      axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx), chart.catAxisLineDash,
    );
    ctx.restore();
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    ctx.save();
    strokeAxisSegment(
      ctx, yAxisX, py0, yAxisX, py0 + ph,
      chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#888',
      axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx), chart.valAxisLineDash,
    );
    ctx.restore();
  }

  // X-axis tick labels (catAxis), formatted via catAxisFormatCode (typically
  // a date code like "m/d/yyyy"). Skipped when catAxisHidden. Drawn just
  // at the authored high/low plot edge or next to the crossing axis. Major
  // tick marks remain attached to the axis rule so `<c:majorTickMark val="cross">` produces
  // the crossing ruler look that templates like the Vertex42 timeline
  // depend on.
  if (!chart.catAxisHidden) {
    const tickFontPx = chart.catAxisFontSizeHpt != null
      ? axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx)
      : Math.max(8, Math.min(11, ph / 20));
    const tickGap = chart.catAxisFontSizeHpt != null
      ? categoryTickLabelGapPx(tickFontPx)
      : 4;
    ctx.font = chartFontCss(
      tickFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.textAlign = 'center';
    const labelPos = chart.catAxisTickLabelPos ?? 'nextTo';
    const lineWidth = axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx);
    const xAxisLineColor = chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : undefined;
    const tickOutset = chart.catAxisLineHidden
      ? 0
      : axisTickOutwardExtentPx(chart.catAxisMajorTickMark, 'major', lineWidth, ptToPx);
    const labelY = labelPos === 'low'
      ? py0 + ph + tickGap
      : labelPos === 'high' ? py0 - tickGap : xAxisY + tickOutset + tickGap;
    ctx.textBaseline = labelPos === 'high' ? 'bottom' : 'top';
    for (const v of xMajorTicks) {
      const gx = toX(v);
      if (labelPos !== 'none') {
        ctx.fillText(formatAxisTickWithUnits(v, chart.catAxisFormatCode, chart.date1904, chart.catAxisDisplayUnits), gx, labelY);
      }
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', xAxisY, gx, xAxisLineColor, lineWidth, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none') {
      for (const value of xMinorTicks) {
        drawAxisTick(ctx, chart.catAxisMinorTickMark, 'cat', xAxisY, toX(value), xAxisLineColor, lineWidth, false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash);
      }
    }
  }

  // Office preserves scatter/bubble group order for overlapping geometry. A
  // group still paints its own line/error/marker/label phases, but the next
  // source group is composited after it. This differs from bar/area/line,
  // whose cross-family layering is application-defined and handled by their
  // dedicated combo path.
  const drawNumericEntries = (
    entries: Array<{ series: ChartSeries; index: number }>,
    isBubble: boolean,
    style: string,
    xMap: (value: number) => number,
    yMap: (value: number) => number,
    maximum: number,
    displayUnits?: ChartDisplayUnits | null,
    bubbleSettings?: BubbleGroupSettings,
  ): void => {
    if (entries.length === 0) return;
    drawScatterSeriesLayer(
      ctx, chart, entries, useIndexX, xMap, yMap, r,
      px0, py0, pw, ph, ptToPx, isBubble, style,
      { x, y, w, h }, maximum, displayUnits, shapeRotationDeg, bubbleSettings,
    );
  };
  if (chart.plotGroups == null) {
    const isBubble = chart.chartType === 'bubble';
    const style = isBubble ? 'marker' : (chart.scatterStyle ?? 'marker');
    drawNumericEntries(
      primaryEntries, isBubble, style, toX, toY,
      yAxisPlan.max, chart.valAxisDisplayUnits,
    );
    if (secondaryEntries.length > 0 && secondaryXPlan && secondaryYPlan) {
      drawNumericEntries(
        secondaryEntries, isBubble, style, toSecondaryX, toSecondaryY,
        secondaryYPlan.max, secondaryY?.displayUnits,
      );
    }
  } else {
    for (const group of chart.plotGroups) {
      if (group.kind !== 'scatter' && group.kind !== 'bubble') continue;
      const entries = chart.series
        .slice(group.seriesStart, group.seriesStart + group.seriesCount)
        .map((series, offset) => ({ series, index: group.seriesStart + offset }));
      if (entries.length === 0) continue;
      const isBubble = group.kind === 'bubble';
      const usesSecondaryX = group.categoryAxis === 'secondary';
      const usesSecondaryY = group.valueAxis === 'secondary';
      drawNumericEntries(
        entries, isBubble,
        isBubble ? 'marker' : (group.scatterStyle ?? chart.scatterStyle ?? 'marker'),
        usesSecondaryX ? toSecondaryX : toX,
        usesSecondaryY ? toSecondaryY : toY,
        usesSecondaryY && secondaryYPlan ? secondaryYPlan.max : yAxisPlan.max,
        usesSecondaryY ? secondaryY?.displayUnits : chart.valAxisDisplayUnits,
        isBubble ? {
          bubbleScale: group.bubbleScale ?? chart.bubbleScale,
          bubbleSizeRepresents: group.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
          showNegativeBubbles: group.showNegativeBubbles ?? chart.showNegativeBubbles,
        } : undefined,
      );
    }
  }

  // The second CT_ScatterChart group owns an independent top X and right Y
  // value-axis pair. Both axes use the same numeric planner and authored unit
  // formatting as the primary pair; only their screen edge differs.
  if (secondaryX && secondaryXPlan && !secondaryX.hidden) {
    const line = resolveAxisLine(secondaryX.lineColor, secondaryX.lineWidthEmu, ptToPx);
    if (!secondaryX.lineHidden) {
      strokeAxisSegment(
        ctx, px0, py0, px0 + pw, py0,
        line.color, line.width, secondaryX.lineDash,
      );
    }
    const fontPx = chartTextFontSizePx(secondaryX.fontSizeHpt, ptToPx) ?? xAxLabelFontPx;
    ctx.font = chartFontCss(
      fontPx,
      chartFontFamily(chart, secondaryX.fontFace, 'minor'),
      secondaryX.fontBold ?? false,
      secondaryX.fontItalic ?? false,
    );
    ctx.fillStyle = secondaryX.fontColor ? `#${secondaryX.fontColor}` : '#555';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    const tickOutset = secondaryX.lineHidden
      ? 0
      : axisTickOutwardExtentPx(secondaryX.majorTickMark, 'major', line.width, ptToPx);
    for (const value of secondaryXPlan.majorTicks) {
      const sx = toSecondaryX(value);
      if (secondaryX.tickLabelPos !== 'none') {
        ctx.fillText(
          formatAxisTickWithUnits(value, secondaryX.formatCode, chart.date1904, secondaryX.displayUnits),
          sx,
          py0 - tickOutset - categoryTickLabelGapPx(fontPx),
        );
      }
      drawAxisTick(
        ctx, secondaryX.majorTickMark, 'cat', py0, sx, line.color, line.width,
        true, secondaryX.lineHidden, 'major', ptToPx, secondaryX.lineDash,
      );
    }
    if (secondaryX.minorTickMark && secondaryX.minorTickMark !== 'none') {
      for (const value of secondaryXPlan.minorTicks) {
        drawAxisTick(
          ctx, secondaryX.minorTickMark, 'cat', py0, toSecondaryX(value), line.color,
          line.width, true, secondaryX.lineHidden, 'minor', ptToPx,
          secondaryX.lineDash,
        );
      }
    }
  }
  if (secondaryY && secondaryYPlan) {
    const scale: SecondaryAxisScale = {
      min: secondaryYPlan.min,
      max: secondaryYPlan.max,
      step: secondaryYPlan.majorUnit,
      majorLines: secondaryYPlan.majorTicks,
      minorTicks: secondaryYPlan.minorTicks,
      makeToY: () => toSecondaryY,
    };
    drawSecondaryValueAxis(
      ctx, chart, secondaryY, scale, toSecondaryY, r,
      px0, py0, pw, ph, ptToPx,
      chartTextFontSizePx(secondaryY.fontSizeHpt, ptToPx) ?? yAxLabelFontPx,
      secondaryYLabelWidth,
      chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555',
      chart.date1904,
    );
  }

  drawLegendForLayout(ctx, chart, leg, x, y, w, h, px0, py0, pw, ph, titleBand.bandH + 2, ptToPx);
  drawAxisTitles(ctx, chart, x, y, w, h, px0, py0, pw, ph, legLeftW, legBottomH, catTitlePx, valTitlePx);
}
