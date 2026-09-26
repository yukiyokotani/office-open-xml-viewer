// Classic stock chart family.
import type { ChartExElementStyle, ChartModel, ChartRect, ChartSeries } from '../../types/chart';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';

import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markerFillColorFor,
  markerFillPaintFor,
  seriesHasMarkerDetail,
} from '../marker-style.js';

import {
  computeChartFrame,
  catAxisLabelBandH,
  chartLegendBands,
  chartAxisTitleBands,
  axisTitleFontPx,
  chartTextFontSizePx,
} from '../layout.js';

import { axisLineWidthPx, resolveAxisLine, isCrossBetween } from '../axis-style.js';
import { formatCategoryLabel } from '../chart-number-format.js';
import { elideToWidth } from '../text-elide.js';
import { categoryLabelAnchorFraction, categoryLabelOffsetPx } from '../category-spacing.js';

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
  catAxisReversed,
  drawValMajorGridlines,
  formatPrimaryValueAxisTick,
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
  applyDecorationLineStyle,
  chartStyleRoleLine,
  chartStyleRoleBarPaint,
  drawDropLineEnvelopes,
  chartStyleRoleErrorBar,
  drawUpDownBars,
  drawChartMarker,
  drawSeriesDataLabels,
  appendCurve,
  drawCategoryErrorBars,
  chartExSeriesFormatIndex,
} from '../shared/classic.js';

// ═══════════════════════════════════════════════════════════════════════════
// Stock chart (ECMA-376 §21.2.2.198)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * High-low-close (and open-high-low-close) stock chart. Series order is fixed
 * by the spec: a 3-series chart is High, Low, Close; a 4-series chart is Open,
 * High, Low, Close. For each category we draw:
 *   - a thin vertical "hi-lo line" from the Low value to the High value
 *     (`<c:hiLowLines>`, §21.2.2.60) — always, when hiLowLines is present;
 *   - the Close series marker at its value (a short tick / dot);
 *   - the Open series marker (4-series only).
 * The value axis, date/category axis, title and legend reuse the shared
 * Cartesian scaffolding (identical to the line renderer). Four-series charts
 * also draw `<c:upDownBars>` (§21.2.2.227) between the first and last
 * authored series. In a four-series stock chart those are Open and Close; the
 * schema also permits the element on a three-series High/Low/Close chart.
 */
export function renderStockChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length;
  if (n === 0) return;
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);

  // Fixed spec series roles by position. With 4 series the first is Open; the
  // last three are always High, Low, Close. Fewer than 3 series can't form a
  // hi-lo-close plot, so fall back to plotting each series' markers only.
  const stockGroup = chart.plotGroups?.find(group => group.kind === 'stock');
  const series = stockGroup
    ? chart.series.slice(stockGroup.seriesStart, stockGroup.seriesStart + stockGroup.seriesCount)
    : chart.series;
  const lineOverlaySeries = chart.plotGroups == null
    ? []
    : chart.plotGroups
        .filter(group => group.kind === 'line')
        .flatMap(group => chart.series.slice(
          group.seriesStart, group.seriesStart + group.seriesCount,
        ));
  const scaleSeries = [...series, ...lineOverlaySeries];
  const sourceSeriesIndices = new Map(chart.series.map((entry, index) => [entry, index]));
  const hasOpen = series.length >= 4;
  const openIdx = hasOpen ? 0 : -1;
  const highIdx = hasOpen ? 1 : 0;
  const lowIdx = hasOpen ? 2 : 1;
  const closeIdx = hasOpen ? 3 : 2;
  const highS = series[highIdx];
  const lowS = series[lowIdx];
  const closeS = series[closeIdx] as ChartSeries | undefined;
  const openS = openIdx >= 0 ? series[openIdx] : undefined;
  const upDownStartS = series[0] as ChartSeries | undefined;
  const upDownEndS = series.at(-1) as ChartSeries | undefined;
  const sec = chart.secondaryValAxis && scaleSeries.some(stockSeries =>
    stockSeries.useSecondaryAxis === true
  ) ? chart.secondaryValAxis : null;
  const isSecondarySeries = (stockSeries: ChartSeries): boolean =>
    sec != null && stockSeries.useSecondaryAxis === true;

  // ── Shared Cartesian frame (mirrors renderLineChart's band computation) ──
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const titleFontPx = titleBand.fontPx;
  const titleTopPad = titleBand.topPad;
  const titleH = titleBand.bandH;
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legBottomH, legTopH } = chartLegendBands(
    leg, chart.legendOverlay === true,
  );
  const catAxFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const valAxFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
  const axBands = chartAxisTitleBands(chart, w, h, ptToPx);
  const catTitlePx = axBands.catFontPx;
  const valTitlePx = axBands.valFontPx;
  const catTitleH = axBands.catBandH;
  const valTitleW = axBands.valBandW;
  const hasDataTable = chartHasDataTable(chart);
  const dataTableBaseH = chartDataTableBaseHeight(chart, ptToPx);
  const dataTableHeaderW = chartDataTableHeaderWidth(ctx, chart, ptToPx);

  const padT = titleH + legTopH + valAxFontPx / 2 + 2;
  const padB = (hasDataTable
    ? dataTableBaseH
    : catAxisLabelBandH(catAxFontPx, chart.catAxisLabelOffsetPercent))
    + catTitleH + legBottomH;

  const phEst = h - padT - padB;
  const secScale = computeSecondaryAxis(sec, scaleSeries, phEst / ptToPx);
  const secTickFontPx = Math.max(8, Math.min(11, h / 20));
  const secFontPx = chartTextFontSizePx(sec?.fontSizeHpt, ptToPx) ?? secTickFontPx;
  let secLabelBandW = 0;
  if (sec && secScale && !sec.hidden) {
    const previousFont = ctx.font;
    ctx.font = chartFontCss(
      secFontPx,
      chartFontFamily(chart, sec.fontFace, 'minor'),
      sec.fontBold ?? false,
      sec.fontItalic ?? false,
    );
    let maxLabelWidth = 0;
    for (const value of secScale.majorLines) {
      maxLabelWidth = Math.max(maxLabelWidth, ctx.measureText(formatAxisTickWithUnits(
        value, sec.formatCode ?? null, chart.date1904, sec.displayUnits,
      )).width);
    }
    secLabelBandW = maxLabelWidth + 18;
    ctx.font = previousFont;
  }
  const secTitleBandW = sec?.title
    ? axisTitleFontPx(sec.titleFontSizeHpt, ptToPx) + 8
    : 0;

  const pad = {
    t: padT,
    r: legRightW + w * 0.05 + secLabelBandW + secTitleBandW,
    b: padB,
    l: legLeftW + Math.max(valAxFontPx * 2.2 + 10 + valTitleW, dataTableHeaderW),
  };

  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + titleTopPad, titleFontPx);

  const stockFrame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    pad,
    honorPlotAreaManualLayout: true,
  });
  const { px0, py0, pw } = stockFrame.plotRect;
  let { ph } = stockFrame.plotRect;
  if (pw <= 0 || ph <= 0) return;

  const dataTableLayout = hasDataTable
    ? measureChartDataTable(ctx, chart, pw / n, ptToPx)
    : null;
  if (dataTableLayout && dataTableLayout.totalHeight > dataTableBaseH) {
    ph = Math.max(1, ph - (dataTableLayout.totalHeight - dataTableBaseH));
  }

  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  // ── Value-axis extent: across every series' plotted values (the hi-lo line
  // needs both the low and high extremes). Authored bounds are retained and
  // omitted bounds flow through the shared automatic planner. ──
  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of scaleSeries) {
    if (isSecondarySeries(s)) continue;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      dataMin = Math.min(dataMin, v);
      dataMax = Math.max(dataMax, v);
    }
  }
  for (const stockSeries of scaleSeries) {
    if (isSecondarySeries(stockSeries)) continue;
    forEachErrorBarEndpoint(
      stockSeries,
      'y',
      index => stockSeries.values[index] ?? null,
      value => {
        dataMin = Math.min(dataMin, value);
        dataMax = Math.max(dataMax, value);
      },
    );
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (chart.valMin != null) dataMin = chart.valMin;
  if (chart.valMax != null) dataMax = chart.valMax;

  const plan = planValueAxis(chart, dataMin, dataMax, ph / ptToPx);
  if (plan.max - plan.min === 0) return;
  const toY = (v: number) => py0 + ph - plan.frac(v) * ph;
  const toYSecondary = secScale?.makeToY(py0, ph) ?? toY;
  const toYFor = (stockSeries: ChartSeries): ((value: number) => number) =>
    isSecondarySeries(stockSeries) ? toYSecondary : toY;

  // Category X mapping — stock charts use crossBetween="between" by default so
  // the first/last hi-lo line isn't flush against the axes (matches Excel).
  const between = isCrossBetween(chart);
  const catRev = catAxisReversed(chart);
  const dateAxisPlan = chartDateAxisPlan(chart, cats);
  const toX = dateAxisPlan
    ? (i0: number) => px0 + dateAxisPlan.positions[i0]! * pw
    : between
      ? (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + ((i + 0.5) / n) * pw; }
      : (i0: number) => { const i = catRev ? n - 1 - i0 : i0; return px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw); };

  // ── Value axis: gridlines + ticks + labels (identical to the line renderer) ──
  if (!chart.valAxisHidden) {
    ctx.font = chartFontCss(
      valAxFontPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.textBaseline = 'middle';
    const grid = valGridStroke(chart, ptToPx);
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    for (const v of plan.minorLines) strokeValueGridlineH(ctx, px0, pw, toY(v), false, minorGrid);
    const drawMajorGrid = drawValMajorGridlines(chart);
    const drawLabels = chart.valAxisTickLabelPos !== 'none';
    for (const v of plan.majorLines) {
      const gy = toY(v);
      if (drawMajorGrid) strokeValueGridlineH(ctx, px0, pw, gy, v === 0, grid);
      drawAxisTick(ctx, chart.valAxisMajorTickMark, 'val', px0, gy, undefined, undefined, false, chart.valAxisLineHidden, 'major', ptToPx, chart.valAxisLineDash);
      if (drawLabels) {
        ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
        ctx.textAlign = 'right';
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), px0 - 6, gy);
      }
    }
    for (const v of plan.minorTicks) {
      drawAxisTick(
        ctx, chart.valAxisMinorTickMark, 'val', px0, toY(v),
        undefined, undefined, false, chart.valAxisLineHidden, 'minor', ptToPx,
        chart.valAxisLineDash,
      );
    }
  }

  if (sec && secScale) {
    drawSecondaryValueGridlines(ctx, sec, secScale, toYSecondary, px0, pw, ptToPx);
  }

  // Axis rules (bottom = category, left = value).
  const stockCatLine = resolveAxisLine(chart.catAxisLineColor, chart.catAxisLineWidthEmu, ptToPx);
  const stockValLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0 + ph, px0 + pw, py0 + ph,
      stockCatLine.color, stockCatLine.width, chart.catAxisLineDash,
    );
  }
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    strokeAxisSegment(
      ctx, px0, py0, px0, py0 + ph,
      stockValLine.color, stockValLine.width, chart.valAxisLineDash,
    );
  }

  // CT_StockChart/dropLines uses the same chart-line paint contract as line
  // and area charts. A stock drop line connects the category axis to the
  // envelope of every finite stock value at that category.
  if (chart.stockDropLines) {
    const compatibility = chart.stockAutomaticStyle ? {
      lineColors: [chart.stockAutomaticStyle.lineColor],
      linePaintAuthored: true,
      lineWidthEmu: chart.stockAutomaticStyle.lineWidthEmu,
    } satisfies ChartExElementStyle : undefined;
    const dropLineStyle = chartStyleRoleLine(
      chart, chart.stockDropLines, 'dropLine', compatibility,
    );
    if ((dropLineStyle.paintAuthored !== true || dropLineStyle.color != null)
      && (dropLineStyle.color != null || dropLineStyle.widthEmu != null
      || dropLineStyle.dash != null) && applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      drawDropLineEnvelopes(
        ctx,
        series,
        n,
        toX,
        stockSeries => toYFor(stockSeries),
        () => py0 + ph,
        (stockSeries, index) => stockSeries.values[index] ?? null,
      );
    }
  }

  // ── Hi-lo lines: vertical Low↔High per category. CT_StockChart makes
  // `<c:hiLowLines>` optional, so absence must remain absence; only a present
  // element receives linked or bounded automatic paint. ──
  const drawHiLo = chart.stockHiLowLines === true && highS != null && lowS != null;
  if (drawHiLo && highS && lowS) {
    const directStyle = chart.stockHiLowLineStyle ?? {
      color: chart.stockHiLowLineColor ?? null,
    };
    const compatibility = chart.stockAutomaticStyle ? {
      lineColors: [chart.stockAutomaticStyle.lineColor],
      linePaintAuthored: true,
      lineWidthEmu: chart.stockAutomaticStyle.lineWidthEmu,
    } satisfies ChartExElementStyle : undefined;
    const lineStyle = chartStyleRoleLine(
      chart, directStyle, 'hiLoLine', compatibility,
    );
    if ((lineStyle.paintAuthored !== true || lineStyle.color != null)
      && (lineStyle.color != null || lineStyle.widthEmu != null || lineStyle.dash != null)
      && applyDecorationLineStyle(ctx, lineStyle, ptToPx)) {
      for (let ci = 0; ci < n; ci++) {
        const hi = highS.values[ci];
        const lo = lowS.values[ci];
        if (hi == null || lo == null) continue;
        const cx = toX(ci);
        const highToY = toYFor(highS);
        const lowToY = toYFor(lowS);
        ctx.beginPath();
        ctx.moveTo(cx, highToY(hi));
        ctx.lineTo(cx, lowToY(lo));
        ctx.stroke();
      }
    }
  }

  // ── Close (and Open) markers. A stock chart's close is drawn as a short tick.
  // If the series carries an explicit `<c:marker>` (symbol/size/fill), honor it;
  // otherwise draw a left/right tick in the series color. ──
  const drawStockTick = (
    s: ChartSeries | undefined,
    seriesIndex: number,
    side: 'left' | 'right' | 'both',
  ): void => {
    if (!s) return;
    const chartIndex = sourceSeriesIndices.get(s) ?? seriesIndex;
    const color = chartColor(chartIndex, s);
    const styleIndex = chartExSeriesFormatIndex(s, chartIndex);
    const pointOverrides = indexPointOverrides(s.dataPointOverrides);
    const seriesMarkerVisible = s.markerSymbol != null && s.markerSymbol !== 'none'
      && seriesHasMarkerDetail(s);
    const tickLen = Math.max(3, (pw / n) * 0.22);
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci];
      if (v == null) continue;
      const cx = toX(ci);
      const cy = toYFor(s)(v);
      const point = pointOverrides.get(ci);
      if (point?.markerSymbol === 'none' || (point?.markerSymbol == null && s.markerSymbol === 'none')) {
        continue;
      }
      const hasExplicitMarker = seriesMarkerVisible
        || (point?.markerSymbol != null && point.markerSymbol !== 'none');
      if (hasExplicitMarker) {
        const symbol = point?.markerSymbol ?? s.markerSymbol ?? 'circle';
        drawChartMarker(
          ctx, chart, s, point, ci, cx, cy, symbol as string,
          point?.markerSize ?? s.markerSize ?? 3,
          markerFillColorFor(s, point, ci, color),
          point?.markerLine ?? s.markerLine ?? null,
          ptToPx,
          (point?.markerLineWidthEmu ?? s.markerLineWidthEmu) != null
            ? axisLineWidthPx(
                (point?.markerLineWidthEmu ?? s.markerLineWidthEmu) as number,
                ptToPx,
              )
            : undefined,
          markerFillPaintFor(s, point, ci), shapeRotationDeg,
        );
        continue;
      }
      // Horizontal tick: close ticks to the RIGHT of the line, open ticks to the
      // LEFT (Excel's open-high-low-close convention). `both` centers it.
      const x0 = side === 'right' ? cx : side === 'left' ? cx - tickLen : cx - tickLen / 2;
      const x1 = side === 'right' ? cx + tickLen : side === 'left' ? cx : cx + tickLen / 2;
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(point?.chartexStyle, s.chartexStyle),
        chart.chartStyleRoles?.dataPointLine,
        styleIndex,
        { x: x0, y: cy, w: x1 - x0, h: 0 },
        ptToPx,
        target => {
          if (!applyClassicStyleLine(
            target, chart, 'dataPointLine', s, point, styleIndex, color,
            Math.max(1, 0.75 * ptToPx), ptToPx,
            { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
          )) return;
          target.beginPath();
          target.moveTo(x0, cy);
          target.lineTo(x1, cy);
          target.stroke();
        },
      );
    }
  };
  // Office accepts a line group after a stock group. Stock decorations retain
  // ownership of the stock slice; the later line group is a normal category
  // line overlay on its resolved value axis rather than becoming a fifth stock
  // role.
  for (const lineSeries of lineOverlaySeries) {
    const chartIndex = sourceSeriesIndices.get(lineSeries) ?? 0;
    const color = chartColor(Math.max(0, chartIndex), lineSeries);
    const yOf = toYFor(lineSeries);
    const pointOverrides = indexPointOverrides(lineSeries.dataPointOverrides);
    const styleIndex = chartExSeriesFormatIndex(lineSeries, chartIndex);
    const varyingLinePaint = chartSeriesVariesByPoint(chart, chartIndex);
    const overlayRuns: IndexedLinePoint[][] = [];
    let overlayRun: IndexedLinePoint[] = [];
    const flushOverlayRun = (): void => {
      if (overlayRun.length > 0) overlayRuns.push(overlayRun);
      overlayRun = [];
    };
    for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
      const value = lineSeries.values[categoryIndex];
      if (value == null) { flushOverlayRun(); continue; }
      overlayRun.push({ x: toX(categoryIndex), y: yOf(value), index: categoryIndex });
    }
    flushOverlayRun();
    const paintOverlayLine = (target: CanvasRenderingContext2D): void => {
      target.save();
      const visible = applyClassicStyleLine(
        target, chart, 'dataPointLine', lineSeries, undefined, styleIndex, color,
        Math.max(1, 2.25 * ptToPx), ptToPx,
        { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      );
      if (visible) {
        target.beginPath();
        for (const run of overlayRuns) {
          target.moveTo(run[0].x, run[0].y);
          appendCurve(target, run, lineSeries.smooth === true);
        }
        target.stroke();
      }
      target.restore();
    };
    if (varyingLinePaint) {
      paintClassicVaryingLineSegments(
        ctx, chart, lineSeries, overlayRuns, lineSeries.smooth === true, false,
        color, Math.max(1, 2.25 * ptToPx), ptToPx,
        { x: px0, y: py0, w: pw, h: ph }, shapeRotationDeg,
      );
    } else {
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(lineSeries.chartexStyle),
        chart.chartStyleRoles?.dataPointLine,
        styleIndex,
        { x: px0, y: py0, w: pw, h: ph },
        ptToPx,
        paintOverlayLine,
      );
    }
    const seriesMarkersVisible = lineSeries.showMarker !== false
      && lineSeries.markerSymbol !== 'none';
    if (seriesMarkersVisible || hasVisiblePointMarkerOverride(lineSeries)) {
      for (let categoryIndex = 0; categoryIndex < n; categoryIndex++) {
        const value = lineSeries.values[categoryIndex];
        if (value == null) continue;
        const point = pointOverrides.get(categoryIndex);
        const symbol = effectiveMarkerSymbol(
          lineSeries, point, 'circle', seriesMarkersVisible,
        );
        if (symbol === 'none') continue;
        drawChartMarker(
          ctx, chart, lineSeries, point, categoryIndex, toX(categoryIndex), yOf(value), symbol,
          point?.markerSize ?? lineSeries.markerSize ?? 5,
          markerFillColorFor(lineSeries, point, categoryIndex, color),
          point?.markerLine ?? lineSeries.markerLine ?? null,
          ptToPx,
          (point?.markerLineWidthEmu ?? lineSeries.markerLineWidthEmu) != null
            ? axisLineWidthPx(
                (point?.markerLineWidthEmu ?? lineSeries.markerLineWidthEmu) as number,
                ptToPx,
              )
            : undefined,
          markerFillPaintFor(lineSeries, point, categoryIndex),
          shapeRotationDeg,
        );
      }
    }
    drawSeriesTrendlines(
      ctx, lineSeries, color, toX, yOf, ptToPx, undefined,
      {
        chart, chartRect: r, plotRect: { x: px0, y: py0, w: pw, h: ph },
        shapeRotationDeg,
      },
    );
  }

  // ── First/last-series up-down bars (§21.2.2.218/227). In DrawingML these
  // decorations follow the line series. Excel paints their opaque bodies over
  // both the high-low rule and the owning series lines, leaving plot geometry
  // visible only outside each body. Explicit stock marker glyphs are replayed
  // after the bodies below, so a marker at a bar endpoint remains fully visible.
  if (chart.stockUpDownBars && upDownStartS && upDownEndS) {
    const directStyle = chart.stockUpDownBarStyle ?? {
      gapWidthPercent: 150,
      up: {},
      down: {},
    };
    const style = {
      ...directStyle,
      up: chartStyleRoleBarPaint(
        chart, directStyle.up, 'upBar', chart.stockAutomaticStyle ?? undefined,
      ),
      down: chartStyleRoleBarPaint(
        chart, directStyle.down, 'downBar', chart.stockAutomaticStyle ?? undefined,
      ),
    };
    const slotWidth = dateAxisPlan
      ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw
      : between ? pw / n : n > 1 ? pw / (n - 1) : pw;
    drawUpDownBars(
      ctx,
      index => upDownStartS.values[index] ?? null,
      index => upDownEndS.values[index] ?? null,
      n, toX, toYFor(upDownStartS), toYFor(upDownEndS),
      slotWidth, style, ptToPx, undefined, shapeRotationDeg,
    );
  }

  // Marker/tick glyphs are the foreground annotation of a stock datum. Desktop
  // Excel paints them after up/down-bar bodies; otherwise a circular High/Low
  // marker intersecting the body is clipped to a semicircle.
  drawStockTick(openS, openIdx, 'left');
  if (highS?.markerSymbol != null || (highS && hasVisiblePointMarkerOverride(highS))) {
    drawStockTick(highS, highIdx, 'both');
  }
  if (lowS?.markerSymbol != null || (lowS && hasVisiblePointMarkerOverride(lowS))) {
    drawStockTick(lowS, lowIdx, 'both');
  }
  drawStockTick(closeS, closeIdx, 'right');

  // CT_LineSer error bars remain attached to their authored stock series.
  // Stock uses a category X axis, so only Y-direction bars have data-unit
  // geometry; the same shared category-series painter is used by line/area.
  for (const stockSeries of scaleSeries) {
    const seriesIndex = sourceSeriesIndices.get(stockSeries) ?? 0;
    const color = chartColor(seriesIndex, stockSeries);
    for (const errorBars of stockSeries.errBars ?? []) {
      drawCategoryErrorBars(
        ctx,
        stockSeries,
        chartStyleRoleErrorBar(chart, errorBars),
        n,
        toX,
        toYFor(stockSeries),
        index => stockSeries.values[index] ?? 0,
        color,
      );
    }
  }

  // If fewer than 3 series (not a real hi-lo-close), still plot each series'
  // markers so nothing is silently dropped.
  if (series.length < 3) {
    for (let si = 0; si < series.length; si++) {
      drawStockTick(series[si], si, 'both');
    }
  }

  // CT_StockChart owns CT_LineSer children, so the same series/default and
  // per-point dLbl contracts used by ordinary category-line charts apply here
  // as well. Paint labels after stock glyphs/error bars so their callout boxes
  // and text remain on top of the plot geometry.
  for (const stockSeries of scaleSeries) {
    const seriesIndex = sourceSeriesIndices.get(stockSeries) ?? 0;
    drawSeriesDataLabels(
      ctx,
      stockSeries,
      cats,
      true,
      toX,
      toYFor(stockSeries),
      ph,
      ptToPx,
      chart.date1904,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      chart.dataLabelPosition ?? 'r',
      { x: px0, y: py0, w: pw, h: ph },
      r,
      face => chartFontFamily(chart, face, 'minor'),
      isSecondarySeries(stockSeries) ? sec?.displayUnits : chart.valAxisDisplayUnits,
      pointIndex => dataLabelLegendKey(seriesIndex, pointIndex),
      value => dataLabelWithinAxisMaximum(
        chart, value, isSecondarySeries(stockSeries) ? secScale?.max ?? plan.max : plan.max,
      ),
      shapeRotationDeg,
    );
  }

  // ── Category (date) axis labels — same path as the line renderer. ──
  if (!chart.catAxisHidden) {
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    const catLabelColor = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#555';
    ctx.fillStyle = catLabelColor; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
    ctx.font = chartFontCss(
      catAxFontPx,
      chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
      chart.catAxisFontBold ?? false,
      chart.catAxisFontItalic ?? false,
    );
    const catSlotMaxPx = dateAxisPlan
      ? (dateAxisPlan.categoryBandFractions[0] ?? 0) * pw - 4
      : (pw / n) * labelInterval - 4;
    const showLabels = !hasDataTable && catLabelsVisible(chart);
    const rotRad = catLabelRotationRad(chart);
    const labelEntries = dateAxisPlan && dateAxisPlan.majorTicks.length > 0
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
      drawAxisTick(ctx, chart.catAxisMajorTickMark, 'cat', py0 + ph, tx, stockCatLine.color, stockCatLine.width, false, chart.catAxisLineHidden, 'major', ptToPx, chart.catAxisLineDash);
      if (!showLabels) continue;
      ctx.textAlign = anchor?.textAlign ?? 'center';
      ctx.fillStyle = catLabelColor;
      const label = entry.label;
      const budget = rotRad === 0 ? catSlotMaxPx : ph * 0.4;
      drawRotatedCatLabel(
        ctx,
        elideToWidth(ctx, label, budget),
        tx,
        py0 + ph + categoryLabelOffsetPx(5, chart.catAxisLabelOffsetPercent),
        rotRad,
      );
    }
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none' && dateAxisPlan) {
      for (const tick of dateAxisPlan.minorTicks) {
        drawAxisTick(
          ctx, chart.catAxisMinorTickMark, 'cat', py0 + ph,
          px0 + tick.fraction * pw, undefined, undefined,
          false, chart.catAxisLineHidden, 'minor', ptToPx, chart.catAxisLineDash,
        );
      }
    }
  }

  if (sec && secScale) {
    const primaryLabelColor = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    drawSecondaryValueAxis(
      ctx, chart, sec, secScale, toYSecondary, r,
      px0, py0, pw, ph, ptToPx,
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
