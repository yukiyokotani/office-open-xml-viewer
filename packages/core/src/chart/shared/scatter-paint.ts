// Classic chart scatter paint helpers.
import type { ChartDisplayUnits, ChartModel, ChartRect, ChartSeries } from '../../types/chart';
import { bubblePointIsThreeD, effectiveMarkerSymbol, hasVisiblePointMarkerOverride, markerFillPaintFor, markersSuppressedByChartStyle, visibleBubbleSize } from '../marker-style.js';
import type { DataLabelRect } from '../data-label-layout.js';
import { hasFilteredScatterAutomaticPointStyle } from '../source-visibility.js';
import { chartSeriesVariesByPoint } from '../effective-style.js';
import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';
import { axisLineWidthPx } from '../axis-style.js';
import { applyClassicStyleLine, chartColor, chartExSeriesFormatIndex, paintClassicVaryingLineSegments } from './palette.js';
import type { IndexedLinePoint } from './palette.js';
import { bubbleSizeMagnitude, scatterXValue } from './scatter-geometry.js';
import type { BubbleGroupSettings } from './scatter-geometry.js';
import { clamp } from './geometry.js';
import { createDataLabelLegendKeyResolver } from './legend.js';
import { drawSeriesErrorBars } from './error-bars.js';
import { chartStyleRoleErrorBar } from './style-roles.js';
import { bubblePointFill, bubblePointLine, scatterPointFill } from './bubble-paint.js';
import { drawChartMarker } from './markers.js';
import { dataLabelWithinAxisMaximum, drawSeriesDataLabels } from './data-labels.js';
import { chartFontFamily } from './fonts.js';
import { drawSeriesTrendlines } from './trendline.js';


export type ScatterSeriesLayer = {
  series: ChartSeries;
  seriesIndex: number;
  fallbackColor: string;
  cats: string[];
  pointOverrides: Map<number, NonNullable<ChartSeries['dataPointOverrides']>[number]>;
};


export function makeScatterSeriesLayer(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
): ScatterSeriesLayer {
  return {
    series,
    seriesIndex: index,
    fallbackColor: chartColor(index, series),
    cats: series.categories ?? chart.categories,
    pointOverrides: new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point])),
  };
}


/** One `<c:bubbleChart>` group has one size scale: every series must therefore
 * be normalized against the same maximum bubble magnitude. */
export function bubbleSizeToDiameterScale(
  chart: BubbleGroupSettings,
  layers: readonly ScatterSeriesLayer[],
  useIndexX: boolean,
  pw: number,
  ph: number,
): number {
  const bubbleScale = clamp(chart.bubbleScale ?? 100, 0, 300);
  if (bubbleScale <= 0) return 0;
  let maxMagnitude = 0;
  for (const { series, cats, pointOverrides } of layers) {
    if (series.showMarker === false || series.markerSymbol === 'none') continue;
    for (let index = 0; index < series.values.length; index++) {
      if (series.values[index] == null || scatterXValue(cats, index, useIndexX) == null) continue;
      if (pointOverrides.get(index)?.markerSymbol === 'none') continue;
      const value = visibleBubbleSize(chart, series.bubbleSizes?.[index]);
      if (value != null) {
        maxMagnitude = Math.max(maxMagnitude, bubbleSizeMagnitude(chart, value));
      }
    }
  }
  if (maxMagnitude <= 0) return 0;
  // ECMA-376 defines bubbleScale as 0..300% of an application-defined default,
  // but intentionally leaves that default to the consumer. Excel's vector
  // output across the complete 0/25/50/75/100/150/200/300 boundary set follows
  // a bounded scale curve: 0 hides bubbles, 100 uses one quarter of the shorter
  // plot dimension, and 300 approaches one half. The equivalent closed form is
  // `shortSide * scale / (300 + scale)`. Keeping it here (rather than a sample-
  // specific diameter constant) makes the Office compatibility rule depend
  // only on the authored scale and the resolved plot geometry.
  const maximumDiameterPx = Math.min(pw, ph) * bubbleScale / (300 + bubbleScale);
  return maximumDiameterPx / maxMagnitude;
}


/** Paint scatter series into an already-computed plot rectangle. Axis/gridline
 * layout stays with the owning chart renderer, which lets a scatter group be
 * overlaid on a bar chart without duplicating either chart's frame. */
export function drawScatterSeriesLayer(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  entries: Array<{ series: ChartSeries; index: number }>,
  useIndexX: boolean,
  toX: (value: number) => number,
  toY: (value: number) => number,
  chartRect: ChartRect,
  px0: number,
  py0: number,
  pw: number,
  ph: number,
  ptToPx: number,
  isBubble: boolean,
  style: string,
  layoutReferenceRect: DataLabelRect,
  valueAxisMaximum: number,
  valueDisplayUnits?: ChartDisplayUnits | null,
  shapeRotationDeg = 0,
  bubbleSettings?: BubbleGroupSettings,
): void {
  const groupDrawsLines = style === 'line' || style === 'lineMarker' || style === 'lineNoMarker';
  const groupDrawsSmooth = style === 'smooth'
    || style === 'smoothMarker'
    || style === 'smoothNoMarker';
  const hideMarkersByStyle = markersSuppressedByChartStyle(
    'scatter', chart.chartType, style, chart.radarStyle,
  );
  const layers = entries.map(({ series, index }) => makeScatterSeriesLayer(chart, series, index));
  const dataLabelLegendKey = createDataLabelLegendKeyResolver(chart, ptToPx);
  const effectiveBubbleSettings = bubbleSettings ?? chart;
  const bubbleScale = isBubble
    ? bubbleSizeToDiameterScale(effectiveBubbleSettings, layers, useIndexX, pw, ph)
    : 0;

  // Excel paints a scatter group by geometry phase, not one complete series at
  // a time: all series lines/error bars first, then all markers, then all data
  // labels. This is observable in dot/range plots where a final invisible
  // scatter series authors full-width horizontal guides. Painting per series
  // placed those guides on top of earlier series' dots and labels.
  for (const { series: s, fallbackColor, cats } of layers) {
    for (const eb of s.errBars ?? []) {
      drawSeriesErrorBars(
        ctx, s, chartStyleRoleErrorBar(chart, eb), cats, useIndexX, toX, toY,
        fallbackColor,
      );
    }
  }

  for (const { series: s, seriesIndex, fallbackColor, cats } of layers) {
    const automaticPointStyle = style === 'marker'
      && hasFilteredScatterAutomaticPointStyle(s);
    const drawLines = automaticPointStyle || groupDrawsLines;
    const drawSmooth = (!automaticPointStyle && style === 'marker' && !isBubble)
      || groupDrawsSmooth;
    if (drawLines || drawSmooth) {
      const pts: IndexedLinePoint[] = [];
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci];
        if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        pts.push({ x: toX(xv), y: toY(yv), index: ci });
      }
      if (pts.length >= 2) {
        const styleIndex = chartExSeriesFormatIndex(s, seriesIndex);
        const scatterBounds = { x: px0, y: py0, w: pw, h: ph };
        if (chartSeriesVariesByPoint(chart, seriesIndex)) {
          paintClassicVaryingLineSegments(
            ctx, chart, s, [pts], drawSmooth, false, fallbackColor, 1.5,
            ptToPx, scatterBounds, shapeRotationDeg, true,
          );
          continue;
        }
        const paintScatterLine = (target: CanvasRenderingContext2D): void => {
          target.save();
          const paintLine = applyClassicStyleLine(
            target, chart, 'dataPointLine', s, undefined, styleIndex, fallbackColor,
            1.5, ptToPx, scatterBounds, shapeRotationDeg,
            true, false,
          );
          if (paintLine && automaticPointStyle && s.dataPointColors?.some(Boolean)) {
            target.lineWidth = 1.5;
            for (let i = 1; i < pts.length; i++) {
              target.strokeStyle = `#${s.dataPointColors[i] ?? s.color ?? fallbackColor.replace(/^#/, '')}`;
              target.beginPath();
              target.moveTo(pts[i - 1].x, pts[i - 1].y);
              target.lineTo(pts[i].x, pts[i].y);
              target.stroke();
            }
          } else if (paintLine) {
            target.beginPath();
            target.moveTo(pts[0].x, pts[0].y);
            if (drawSmooth && pts.length >= 3) {
              for (let i = 0; i < pts.length - 1; i++) {
                const p0 = pts[i - 1] ?? pts[i];
                const p1 = pts[i];
                const p2 = pts[i + 1];
                const p3 = pts[i + 2] ?? p2;
                target.bezierCurveTo(
                  p1.x + (p2.x - p0.x) / 6,
                  p1.y + (p2.y - p0.y) / 6,
                  p2.x - (p3.x - p1.x) / 6,
                  p2.y - (p3.y - p1.y) / 6,
                  p2.x,
                  p2.y,
                );
              }
            } else {
              for (let i = 1; i < pts.length; i++) target.lineTo(pts[i].x, pts[i].y);
            }
            target.stroke();
          }
          target.restore();
        };
        paintChartStyleEffects(
          ctx,
          chartStyleEffectOwner(s.chartexStyle),
          chart.chartStyleRoles?.dataPointLine,
          styleIndex,
          scatterBounds,
          ptToPx,
          paintScatterLine,
        );
      }
    }
  }

  for (const { series: s, seriesIndex, fallbackColor, cats, pointOverrides } of layers) {
    const seriesMarkersVisible = !hideMarkersByStyle
      && s.showMarker !== false
      && s.markerSymbol !== 'none';
    if (seriesMarkersVisible || (!hideMarkersByStyle && hasVisiblePointMarkerOverride(s))) {
      for (let ci = 0; ci < s.values.length; ci++) {
        const yv = s.values[ci];
        if (yv == null) continue;
        const xv = scatterXValue(cats, ci, useIndexX);
        if (xv == null) continue;
        const dpt = pointOverrides.get(ci);
        const defaultSymbol = isBubble ? 'circle' : (s.automaticMarkerSymbol ?? 'circle');
        const symbol = effectiveMarkerSymbol(s, dpt, defaultSymbol, seriesMarkersVisible);
        if (symbol === 'none') continue;
        let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
        if (isBubble) {
          if (bubbleScale <= 0) continue;
          const bubbleSize = visibleBubbleSize(effectiveBubbleSettings, s.bubbleSizes?.[ci]);
          if (bubbleSize == null) continue;
          sizePt = (bubbleSizeMagnitude(effectiveBubbleSettings, bubbleSize) * bubbleScale) / ptToPx;
        }
        const bubbleFill = isBubble
          ? bubblePointFill(
              chart, s, dpt, ci, chartExSeriesFormatIndex(s, seriesIndex), fallbackColor,
            )
          : null;
        const fill = bubbleFill?.color ?? scatterPointFill(s, dpt, ci, fallbackColor);
        // Bubble geometry is the series shape itself, so its outline comes from
        // `<c:ser><c:spPr><a:ln>` rather than a `<c:marker>` block. Ordinary
        // scatter markers continue to use markerLine only.
        const bubbleLine = isBubble
          ? bubblePointLine(chart, s, dpt, ci, chartExSeriesFormatIndex(s, seriesIndex))
          : null;
        const line = isBubble
          ? bubbleLine!.color
          : dpt?.markerLine ?? s.markerLine ?? null;
        const markerLineWidthEmu = dpt?.markerLineWidthEmu ?? s.markerLineWidthEmu;
        const bubbleLineWidthEmu = bubbleLine?.widthEmu;
        const lineWidthEmu = isBubble ? bubbleLineWidthEmu : markerLineWidthEmu;
        const lineWidthPx = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx)
          : undefined;
        const isThreeD = isBubble ? bubblePointIsThreeD(s, dpt) : false;
        drawChartMarker(
          ctx, chart, s, dpt, ci, toX(xv), toY(yv), symbol, sizePt, fill, line, ptToPx, lineWidthPx,
          isBubble ? bubbleFill!.paint : markerFillPaintFor(s, dpt, ci), shapeRotationDeg,
          isBubble ? bubbleLine!.paint : undefined,
          isBubble ? bubbleLine!.dash : undefined,
          isBubble ? bubbleLine!.customDash : undefined,
          isBubble ? bubbleLine!.cap : undefined,
          isBubble ? bubbleLine!.join : undefined,
          isThreeD,
          isBubble,
        );
      }
    }
  }

  for (const { series: s, seriesIndex, cats, pointOverrides } of layers) {
    const markerGapAt = (pointIndex: number): number => {
      if (hideMarkersByStyle) return 0;
      const seriesMarkerVisible = s.showMarker !== false && s.markerSymbol !== 'none';
      const dpt = pointOverrides.get(pointIndex);
      const symbol = effectiveMarkerSymbol(s, dpt, 'circle', seriesMarkerVisible);
      if (symbol === 'none') return 0;
      let sizePt = dpt?.markerSize ?? s.markerSize ?? 5;
      if (isBubble) {
        if (bubbleScale <= 0) return 0;
        const bubbleSize = visibleBubbleSize(
          effectiveBubbleSettings, s.bubbleSizes?.[pointIndex],
        );
        if (bubbleSize == null) return 0;
        sizePt = bubbleSizeMagnitude(effectiveBubbleSettings, bubbleSize) * bubbleScale / ptToPx;
      }
      return Math.max(0, sizePt * ptToPx / 2);
    };
    drawSeriesDataLabels(
      ctx,
      s,
      cats,
      useIndexX,
      toX,
      toY,
      ph,
      ptToPx,
      chart.date1904,
      chartFontFamily(chart, chart.dataLabelFontFace, 'minor'),
      chart.dataLabelPosition ?? 'r',
      // Office lets automatic left/right endpoint labels occupy the chart-area
      // gutter while keeping their vertical placement constrained to the plot.
      { x: chartRect.x, y: py0, w: chartRect.w, h: ph },
      layoutReferenceRect,
      face => chartFontFamily(chart, face, 'minor'),
      valueDisplayUnits,
      pointIndex => dataLabelLegendKey(seriesIndex, pointIndex),
      value => dataLabelWithinAxisMaximum(chart, value, valueAxisMaximum),
      shapeRotationDeg,
      markerGapAt,
    );
  }

  for (const { series: s, fallbackColor, cats } of layers) {
    const trendlineX = s.values.map((_, index) => scatterXValue(cats, index, useIndexX));
    drawSeriesTrendlines(
      ctx, s, fallbackColor, toX, toY, ptToPx, trendlineX,
      {
        chart,
        chartRect,
        plotRect: { x: px0, y: py0, w: pw, h: ph },
        clipLineToPlot: true,
        shapeRotationDeg,
      },
    );
  }
}
