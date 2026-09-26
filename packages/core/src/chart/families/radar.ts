// Classic radar chart family.
import type { ChartModel, ChartRect } from '../../types/chart';

import { classicDataPointFillDecision } from '../classic-data-point-style.js';

import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';

import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markersSuppressedByChartStyle,
  markerFillColorFor,
  markerFillPaintFor,
} from '../marker-style.js';

import { computeChartFrame } from '../layout.js';
import { automaticRadarMajorUnit, planNumericValueAxis } from '../axis-scale.js';
import { axisLineWidthPx, resolveAxisLine } from '../axis-style.js';
import { formatCategoryLabel } from '../chart-number-format.js';
import { elideToWidth } from '../text-elide.js';
import { categoryLabelOffsetPx } from '../category-spacing.js';

import { chartDataPointStyleRole, chartSeriesVariesByPoint } from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';

import { hexToRgba } from '../../shape/paint.js';

import {
  chartColor,
  indexPointOverrides,
  applyClassicStyleLine,
  IndexedLinePoint,
  paintClassicVaryingLineSegments,
  chartFontFamily,
  chartFontCss,
  measuredLegendReserve,
  drawLegendForLayout,
  drawAxisTick,
  valGridStroke,
  valMinorGridStroke,
  valAxisReversed,
  drawValMajorGridlines,
  formatPrimaryValueAxisTick,
  axisLabelPx,
  catLabelsVisible,
  drawChartTitleForLayout,
  chartCategories,
  drawChartMarker,
  clamp,
  chartExSeriesFormatIndex,
  paintClassicDataPointPath,
} from '../shared/classic.js';

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════
export function renderRadarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const cats = chartCategories(chart);
  const n = cats.length; if (n < 3) return;

  // Shared frame (radial form). Radar uses title pads 0.035 / 0.035 and the
  // default 0.22 side-legend reserve (unlike pie's 0.28). Params keep pixels
  // unchanged.
  const leg = measuredLegendReserve(ctx, chart, w, h, 0.22, ptToPx);
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleTopPadFrac: 0.035,
    titleBottomPadFrac: 0.035,
    legendSideReserveFrac: 0.22,
    legendReserve: leg,
    radialGapFrac: 0.02,
    honorPlotAreaManualLayout: true,
  });
  const titleFontPx = frame.title.fontPx;
  drawChartTitleForLayout(ctx, chart, x, y, w, h, y + frame.title.topPad, titleFontPx);

  const { px0: plotLeft, py0: plotTop, pw, ph } = frame.plotRect;
  paintPlotAreaFrame(
    ctx, chart, plotLeft, plotTop, pw, ph, ptToPx, shapeRotationDeg,
  );
  const cx2 = frame.center.cx;
  const cy2 = frame.center.cy;
  // An explicitly sized `layoutTarget="inner"` rectangle defines the data
  // region itself (ECMA-376 §21.2.2.88), so the outer radar ring is the
  // largest circle inscribed in it. Automatic, position-only, and outer
  // layouts keep the existing label reserve.
  const manualLayout = chart.plotAreaManualLayout;
  const hasExplicitInnerSize = manualLayout?.layoutTarget === 'inner'
    && manualLayout.w != null
    && manualLayout.h != null
    && Number.isFinite(manualLayout.w)
    && Number.isFinite(manualLayout.h)
    && frame.plotAreaManualLayoutApplied;
  const rd = hasExplicitInnerSize
    ? Math.min(pw, ph) / 2
    : Math.min(pw, ph) * 0.38;

  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const s of chart.series) for (const v of s.values) {
    if (v == null) continue;
    dataMin = Math.min(dataMin, v);
    dataMax = Math.max(dataMax, v);
  }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (dataMax === 0) dataMax = 1;
  const needsMinorTicks = chart.valAxisMinorTickMark != null
    && chart.valAxisMinorTickMark !== 'none';
  // An explicit `<c:valAx><c:majorUnit>` (§21.2.2.103) overrides the automatic
  // ring step. Omission uses the radar-specific radial density observed across
  // small/ordinary/large boundary charts, then the shared bounded planner.
  const radarLog = chart.valAxisLogBase != null
    && Number.isFinite(chart.valAxisLogBase)
    && chart.valAxisLogBase >= 2;
  const radarMajorUnit = chart.valAxisMajorUnit ?? (radarLog
    ? null
    : automaticRadarMajorUnit(
        chart.valMin ?? dataMin,
        chart.valMax ?? dataMax,
        rd / ptToPx,
      ));
  const radarAxisPlan = planNumericValueAxis({
    dataMin,
    dataMax,
    explicitMin: chart.valMin,
    explicitMax: chart.valMax,
    axisLenPt: rd / ptToPx,
    axisOrientation: 'vertical',
    majorUnit: radarMajorUnit,
    minorUnit: chart.valAxisMinorUnit,
    needMinor: chart.valAxisMinorGridlines === true || needsMinorTicks,
    logBase: chart.valAxisLogBase,
    reversed: valAxisReversed(chart),
  });
  const radarFrac = (value: number): number => clamp(radarAxisPlan.fraction(value), 0, 1);

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  // Rings sit on the value-axis MAJOR ticks — i.e. at value `ri * step`, whose
  // radius is proportional to the value (`v / axMax`). Deriving the radius from
  // the value (not `ri / rings`) keeps the rings on the major-unit multiples
  // even when `axMax` is not an exact multiple of `step` (e.g. an explicit
  // `<c:majorUnit>` §21.2.2.103 that doesn't divide the auto-rounded max).
  const ringValues = radarAxisPlan.majorTicks.filter(value => radarFrac(value) > 0);
  const strokeRing = (value: number): void => {
    const rr = radarFrac(value) * rd;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  };
  if (chart.valAxisMinorGridlines) {
    const minorGrid = valMinorGridStroke(chart, ptToPx);
    ctx.strokeStyle = minorGrid.color;
    ctx.lineWidth = minorGrid.width;
    const previousDash = minorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (minorGrid.dash.length > 0) ctx.setLineDash(minorGrid.dash);
    for (const value of radarAxisPlan.minorTicks) strokeRing(value);
    if (minorGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }
  if (!chart.valAxisHidden && drawValMajorGridlines(chart)) {
    const majorGrid = valGridStroke(chart, ptToPx);
    ctx.strokeStyle = majorGrid.color;
    ctx.lineWidth = majorGrid.width;
    const previousDash = majorGrid.dash.length > 0 && ctx.getLineDash ? ctx.getLineDash() : [];
    if (majorGrid.dash.length > 0) ctx.setLineDash(majorGrid.dash);
    for (const ringValue of ringValues) strokeRing(ringValue);
    if (majorGrid.dash.length > 0) ctx.setLineDash(previousDash);
  }

  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * rd, cy2 + Math.sin(a) * rd); ctx.stroke();
  }

  // Radial tick labels on the top (12 o'clock) spoke — Excel places the value
  // axis there for radar charts. Respect <c:valAx><c:delete val="1"/> when the
  // caller hides the axis, and skip the 0-label at the center to avoid
  // overlapping the origin point.
  if (!chart.valAxisHidden) {
    const valAxPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
    ctx.font = chartFontCss(
      valAxPx,
      chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
      chart.valAxisFontBold ?? false,
      chart.valAxisFontItalic ?? false,
    );
    ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#555';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    for (const v of ringValues) {
      const rr = radarFrac(v) * rd;
      const y = cy2 - rr;
      const valAxisLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
      drawAxisTick(
        ctx, chart.valAxisMajorTickMark, 'val', cx2, y,
        valAxisLine.color, valAxisLine.width, false, chart.valAxisLineHidden, 'major', ptToPx,
        chart.valAxisLineDash,
      );
      if (chart.valAxisTickLabelPos !== 'none') {
        ctx.fillText(formatPrimaryValueAxisTick(chart, v, false), cx2 - 3, y);
      }
    }
    if (needsMinorTicks) {
      const valAxisLine = resolveAxisLine(chart.valAxisLineColor, chart.valAxisLineWidthEmu, ptToPx);
      for (const value of radarAxisPlan.minorTicks) {
        drawAxisTick(
          ctx,
          chart.valAxisMinorTickMark,
          'val',
          cx2,
          cy2 - radarFrac(value) * rd,
          valAxisLine.color,
          valAxisLine.width,
          false,
          chart.valAxisLineHidden,
          'minor',
          ptToPx,
          chart.valAxisLineDash,
        );
      }
    }
  }

  const radarCatFontPx = chart.catAxisFontSizeHpt != null
    ? axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx)
    : Math.max(8, Math.min(11, rd * 0.2));
  ctx.font = chartFontCss(
    radarCatFontPx,
    chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#444';
  ctx.textBaseline = 'middle';
  // Spoke labels radiate from just outside the ring. Cap each at the room
  // between its anchor and the nearest horizontal plot edge so long category
  // names are elided instead of overrunning the chart frame. Left/right-aligned
  // labels extend toward one edge; centered (top/bottom) labels straddle the
  // anchor, so give them twice the smaller side.
  const plotLeftX = cx2 - pw / 2;
  const plotRightX = cx2 + pw / 2;
  if (!chart.catAxisHidden && catLabelsVisible(chart)) for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const radialLabelOffset = categoryLabelOffsetPx(12, chart.catAxisLabelOffsetPercent);
    const lx = cx2 + Math.cos(a) * (rd + radialLabelOffset);
    const ly = cy2 + Math.sin(a) * (rd + radialLabelOffset);
    const authoredAlignment = chart.catAxisLabelAlignment;
    const align: CanvasTextAlign = authoredAlignment === 'l'
      ? 'left'
      : authoredAlignment === 'r'
        ? 'right'
        : authoredAlignment === 'ctr'
          ? 'center'
          : Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.textAlign = align;
    const maxPx =
      align === 'right' ? lx - plotLeftX
        : align === 'left' ? plotRightX - lx
          : 2 * Math.min(plotRightX - lx, lx - plotLeftX);
    // §21.2.2.71: format numeric-serial categories via the category-axis
    // numFmt; string spoke labels pass through unchanged.
    const label = formatCategoryLabel((cats[i] ?? '').toString(), chart.catAxisFormatCode, chart.date1904);
    ctx.fillText(elideToWidth(ctx, label, maxPx), lx, ly);
  }

  // ECMA-376 §21.2.3.10 c:radarStyle — "filled" closes the polygon with a
  // translucent area fill; "standard" / "marker" (and default) draw the
  // line only. Markers come from per-series `<c:marker>` (which can
  // override the chart-type style by setting `<c:symbol val="none"/>`);
  // A chart may set radarStyle="marker" while every series carries
  // `<c:marker><c:symbol val="none"/>`, in which case Office draws
  // lines only — no dots.
  const filled = markersSuppressedByChartStyle(
    'radar', chart.chartType, chart.scatterStyle, chart.radarStyle,
  );
  const markerRadius = Math.max(2, rd * 0.025);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, s);
    const styleIndex = chartExSeriesFormatIndex(s, si);
    // Build the per-spoke point list, leaving holes where the series has
    // no value (`<c:val>` ptCount > pts implies missing indices), so Office draws an open polyline
    // from idx 1 to idx 10 without bridging back through the top spoke).
    const pts: Array<[number, number] | null> = [];
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v == null) { pts.push(null); continue; }
      const frac = radarFrac(v);
      const a = spoke(i);
      pts.push([cx2 + Math.cos(a) * rd * frac, cy2 + Math.sin(a) * rd * frac]);
    }

    const allPresent = pts.every(p => p != null);
    const indexedRuns: IndexedLinePoint[][] = [];
    let indexedRun: IndexedLinePoint[] = [];
    for (let pointIndex = 0; pointIndex < pts.length; pointIndex++) {
      const point = pts[pointIndex];
      if (point == null) {
        if (indexedRun.length > 0) indexedRuns.push(indexedRun);
        indexedRun = [];
      } else {
        indexedRun.push({ x: point[0], y: point[1], index: pointIndex });
      }
    }
    if (indexedRun.length > 0) indexedRuns.push(indexedRun);
    const radarBounds = { x: plotLeft, y: plotTop, w: pw, h: ph };
    const paintRadarSeries = (target: CanvasRenderingContext2D): void => {
      const previousDash = target.getLineDash ? target.getLineDash() : [];
      const previousCap = target.lineCap;
      const previousJoin = target.lineJoin;
      target.beginPath();
      let pen = false;
      for (const pt of pts) {
        if (pt == null) { pen = false; continue; }
        if (!pen) { target.moveTo(pt[0], pt[1]); pen = true; }
        else { target.lineTo(pt[0], pt[1]); }
      }
      if (allPresent) target.closePath();
      if (filled && allPresent) {
        const fillDecision = classicDataPointFillDecision(chart, s, undefined, styleIndex);
        paintClassicDataPointPath(
          target, fillDecision, radarBounds, hexToRgba(color, 0.25),
          ptToPx, shapeRotationDeg,
        );
      }
      if (applyClassicStyleLine(
        target, chart, 'dataPointLine', s, undefined, styleIndex, color,
        2, ptToPx, radarBounds, shapeRotationDeg,
      )) target.stroke();
      target.setLineDash(previousDash);
      target.lineCap = previousCap;
      target.lineJoin = previousJoin;
    };
    if (chartSeriesVariesByPoint(chart, si)) {
      if (filled && allPresent) {
        const firstPoint = indexPointOverrides(s.dataPointOverrides).get(0);
        const paintRadarFill = (target: CanvasRenderingContext2D): void => {
          const fillDecision = classicDataPointFillDecision(chart, s, firstPoint, 0, 0);
          if (fillDecision === null) return;
          target.beginPath();
          for (let index = 0; index < pts.length; index++) {
            const point = pts[index] as [number, number];
            if (index === 0) target.moveTo(point[0], point[1]);
            else target.lineTo(point[0], point[1]);
          }
          target.closePath();
          paintClassicDataPointPath(
            target, fillDecision, radarBounds, hexToRgba(color, 0.25),
            ptToPx, shapeRotationDeg,
          );
        };
        paintChartStyleEffects(
          ctx,
          chartStyleEffectOwner(firstPoint?.chartexStyle, s.chartexStyle),
          chartDataPointStyleRole(chart, 'dataPoint', si),
          0,
          radarBounds,
          ptToPx,
          paintRadarFill,
          0,
        );
      }
      paintClassicVaryingLineSegments(
        ctx, chart, s, indexedRuns, false, allPresent, color, 2,
        ptToPx, radarBounds, shapeRotationDeg,
      );
    } else {
      paintChartStyleEffects(
        ctx,
        chartStyleEffectOwner(s.chartexStyle),
        chart.chartStyleRoles?.[filled ? 'dataPoint' : 'dataPointLine'],
        styleIndex,
        radarBounds,
        ptToPx,
        paintRadarSeries,
      );
    }

    // Markers: honor the per-series marker_symbol. When the series
    // explicitly carries `<c:marker><c:symbol val="none"/>`, the parser
    // sets showMarker=false — respect that even for radarStyle="marker"
    // charts (the chart-level style is the default; series overrides win).
    const seriesMarkersVisible = !filled && s.showMarker !== false && s.markerSymbol !== 'none';
    if (!filled && (seriesMarkersVisible || hasVisiblePointMarkerOverride(s))) {
      const pointOverrides = indexPointOverrides(s.dataPointOverrides);
      for (let pointIndex = 0; pointIndex < pts.length; pointIndex++) {
        const pt = pts[pointIndex];
        if (pt == null) continue;
        const point = pointOverrides.get(pointIndex);
        const symbol = effectiveMarkerSymbol(s, point, 'circle', seriesMarkersVisible);
        if (symbol === 'none') continue;
        const size = point?.markerSize ?? s.markerSize ?? Math.max(4, markerRadius * 2 / ptToPx);
        const fill = markerFillColorFor(s, point, pointIndex, color);
        const line = point?.markerLine ?? s.markerLine ?? null;
        const lineWidth = point?.markerLineWidthEmu ?? s.markerLineWidthEmu;
        drawChartMarker(
          ctx, chart, s, point, pointIndex, pt[0], pt[1], symbol, size, fill, line, ptToPx,
          lineWidth != null ? axisLineWidthPx(lineWidth, ptToPx) : 1,
          markerFillPaintFor(s, point, pointIndex), shapeRotationDeg,
        );
      }
    }
  }

  drawLegendForLayout(
    ctx, chart, leg,
    x, y, w, h,
    plotLeft, plotTop, pw, ph, frame.title.bandH + 2,
    ptToPx,
  );
}
