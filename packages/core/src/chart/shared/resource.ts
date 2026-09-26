// Classic chart resource helpers.
import type { ChartModel, ChartRect, ChartSeries } from '../../types/chart';
import { chartTextFontSizePx } from '../layout.js';
import { chartImageFillPaintWorkUpperBound } from '../image-fill.js';
import type { ChartImageLookup } from '../image-fill.js';
import { PT_TO_PX } from '../../units.js';
import { MAX_CANVAS_CHART_POINTS, classicCanvasPointFamilyIsPainted } from '../resource-limits.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from '../plot-groups.js';
import { deletedLegendEntryIndices, legendEntryIsVisible, legendEntryRanges } from '../legend-entry-plan.js';
import type { Fill } from '../../types/common';
import { bubblePointIsThreeD, classicMarkerPointIsPainted, dataLabelLegendKeyCount, effectiveMarkerSymbol, hasVisiblePointMarkerOverride, markerFillPaintFor, markerPaintComponents, markerSymbolConsumesFill, markersSuppressedByChartStyle, seriesHasMarkerDetail, seriesLegendMarkerIsVisible, seriesMarkerFillPaint, visibleBubbleSize } from '../marker-style.js';
import { dataLabelIsDeleted } from '../data-label-style.js';
import { computeBoxWhiskerStats } from '../box-whisker.js';
import { THREE_D_MAX_SHAPE_FACES_PER_DATUM } from '../three-d-contract.js';
import type { ChartThreeDRenderer } from '../three-d-contract.js';
import { LEGEND_ROW_EXTRA_PX, bubblePointLegendMarker, classicPointLegendMarker } from './legend.js';
import { chartHasDataTable } from './data-table.js';
import { chartCategories } from '../category-spacing.js';
import { bubbleSizeToDiameterScale, makeScatterSeriesLayer } from './scatter-paint.js';
import { CLASSIC_THREE_D_FAMILIES, MAX_CANVAS_MARKER_GRADIENT_STOPS, MAX_CANVAS_MARKER_PAINT_COMPONENTS } from './paint-limits.js';
import { chartColor, chartExSeriesFormatIndex, indexPointOverrides } from './palette.js';
import { BUBBLE_3D_MATERIAL_COMPONENTS } from './markers.js';
import { bubbleSizeMagnitude } from './scatter-geometry.js';
import { bubblePointFill, bubblePointLine } from './bubble-paint.js';
import { chartExMarkerPaint } from './chartex-style.js';


export function markerKeyPaintSizesPx(
  chart: ChartModel,
  series: ChartSeries,
  ptToPx: number,
): { legend: number; table: number; labels: number } {
  const legendFontPx = chartTextFontSizePx(chart.legendFontSizeHpt, ptToPx) ?? 10 * ptToPx;
  const tableFontPx = chartTextFontSizePx(chart.dataTable?.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  let labelFontPx = chartTextFontSizePx(
    series.seriesDataLabels?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
    ptToPx,
  ) ?? 10 * ptToPx;
  for (const label of series.dataLabelOverrides ?? []) {
    labelFontPx = Math.max(
      labelFontPx,
      chartTextFontSizePx(label.fontSizeHpt, ptToPx) ?? labelFontPx,
    );
  }
  // Side legends retain at most two lines. Their marker receives 0.58× the
  // final row height; data-table and data-label keys are no larger than their
  // respective font boxes. Use those actual consumer bounds, not markerSize.
  return {
    legend: Math.max(2, (2 * legendFontPx + LEGEND_ROW_EXTRA_PX) * 0.58),
    table: Math.max(2, tableFontPx),
    labels: Math.max(2, labelFontPx),
  };
}


/** @internal Exported for resource-boundary regression tests; package entry
 * points do not expose the renderer module as public API. */
export function classicMarkerPaintWorkCount(
  chart: ChartModel,
  imageLookup?: ChartImageLookup,
  ptToPx = PT_TO_PX,
  chartRect?: ChartRect,
): number | null {
  const hasClassicMarkers = classicCanvasPointFamilyIsPainted(chart.chartType);
  const hasBoxMarkers = chart.chartexBox != null;
  if (!hasClassicMarkers && !hasBoxMarkers) return null;
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const scatterHasNumericX = chart.series.some((series, seriesIndex) => {
    const group = plotGroupBySeries[seriesIndex];
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    if (family !== 'scatter') return false;
    return (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  const dataTableMarkerKeysVisible = chartHasDataTable(chart)
    && chartCategories(chart).length > 0
    && chart.dataTable?.showKeys === true;
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  const legendRanges = legendEntryRanges(chart, true);
  const legacyBubbleScale = chart.chartType === 'bubble' && chartRect
    ? bubbleSizeToDiameterScale(
        chart,
        chart.series.map((series, index) => makeScatterSeriesLayer(chart, series, index)),
        !scatterHasNumericX,
        chartRect.w,
        chartRect.h,
      )
    : 0;
  const bubbleScaleByGroup = new Map<NonNullable<ChartModel['plotGroups']>[number], number>();
  if (chartRect) for (const group of chart.plotGroups ?? []) {
    if (group.kind !== 'bubble' || group.seriesCount === 0) continue;
    const layers = chart.series
      .slice(group.seriesStart, group.seriesStart + group.seriesCount)
      .map((series, offset) => makeScatterSeriesLayer(chart, series, group.seriesStart + offset));
    bubbleScaleByGroup.set(group, bubbleSizeToDiameterScale({
      bubbleScale: group.bubbleScale ?? chart.bubbleScale,
      bubbleSizeRepresents: group.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
      showNegativeBubbles: group.showNegativeBubbles ?? chart.showNegativeBubbles,
    }, layers, !scatterHasNumericX, chartRect.w, chartRect.h));
  }
  let total = 0;
  const chargeComponents = (components: number, repetitions = 1): boolean => {
    if (repetitions <= 0 || components <= 0) return true;
    if (!Number.isSafeInteger(repetitions)
      || components > Math.floor((MAX_CANVAS_MARKER_PAINT_COMPONENTS - total) / repetitions)) {
      return false;
    }
    total += components * repetitions;
    return true;
  };
  const chargePaint = (
    paint: Fill | null | undefined,
    repetitions = 1,
    sizePx = Math.max(2, 5 * ptToPx),
  ): boolean => {
    if (repetitions <= 0 || paint == null) return true;
    const components = paint.fillType === 'image'
      ? chartImageFillPaintWorkUpperBound(paint, imageLookup, sizePx, sizePx, ptToPx)
      : markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CANVAS_MARKER_GRADIENT_STOPS) {
      return false;
    }
    return chargeComponents(components, repetitions);
  };
  if (hasClassicMarkers) for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const series = chart.series[seriesIndex];
    const group = plotGroupBySeries[seriesIndex];
    const isBubble = group?.kind === 'bubble'
      || (group == null && chart.chartType === 'bubble');
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    const effectiveChartType = markerChartTypeForPlotGroup(chart.chartType, group);
    const effectiveScatterStyle = group?.scatterStyle ?? chart.scatterStyle;
    const effectiveRadarStyle = group?.radarStyle ?? chart.radarStyle;
    const markerContext = {
      chartType: effectiveChartType,
      bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
      showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
    };
    const bubbleSettings = isBubble ? {
      bubbleScale: markerContext.bubbleScale,
      bubbleSizeRepresents: group?.bubbleSizeRepresents ?? chart.bubbleSizeRepresents,
      showNegativeBubbles: markerContext.showNegativeBubbles,
    } : undefined;
    const bubbleScale = group?.kind === 'bubble'
      ? bubbleScaleByGroup.get(group) ?? 0
      : legacyBubbleScale;
    const markerFamily = family === 'line' || family === 'stackedLine'
      || family === 'stackedLinePct' || family === 'area'
      || family === 'stackedArea' || family === 'stackedAreaPct'
      || family === 'scatter' || family === 'radar' || family === 'stock';
    if (!markerFamily) continue;
    // Family-level style choices override every series marker. Keep the
    // availability preflight on the same path as the painters: filled radar
    // never paints markers, and the two no-marker scatter styles suppress even
    // point-local marker overrides. Bubble geometry is unaffected by the
    // scatter style token and therefore remains chargeable.
    if (markersSuppressedByChartStyle(
      family, effectiveChartType, effectiveScatterStyle, effectiveRadarStyle,
    )) continue;
    const areaFamily = family === 'area' || family === 'stackedArea'
      || family === 'stackedAreaPct';
    const seriesVisible = areaFamily
      ? (series.showMarker === true || seriesHasMarkerDetail(series))
        && series.markerSymbol !== 'none'
      : family === 'stock'
        ? series.markerSymbol != null && series.markerSymbol !== 'none'
        : series.showMarker !== false && series.markerSymbol !== 'none';
    if (!seriesVisible && !hasVisiblePointMarkerOverride(series)) continue;
    const pointCount = Math.max(
      series.values.length,
      series.categories?.length ?? 0,
      chart.categories.length,
    );
    const overrides = indexPointOverrides(series.dataPointOverrides);
    const labelOverrides = new Map(
      (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
    );
    const keySizes = markerKeyPaintSizesPx(chart, series, ptToPx);
    const pointDrivenKeys = legendRanges[seriesIndex]?.pointDriven === true;
    for (let index = 0; index < pointCount; index++) {
      const plotPointVisible = classicMarkerPointIsPainted(
        chart, series, family, index, scatterHasNumericX, markerContext,
      ) && (!isBubble
        || (bubbleScale > 0
          && visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]) != null));
      const legendKeyVisible = pointDrivenKeys && chart.showLegend
        && legendEntryIsVisible(legendRanges, deletedLegendEntries, seriesIndex, index);
      const label = labelOverrides.get(index);
      const labelKeyVisible = pointDrivenKeys && plotPointVisible
        && !dataLabelIsDeleted(series.seriesDataLabels, label)
        && (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey ?? false) === true;
      const tableKeyVisible = pointDrivenKeys && dataTableMarkerKeysVisible && index === 0;
      if (!plotPointVisible && !legendKeyVisible && !labelKeyVisible && !tableKeyVisible) continue;
      if (isBubble && plotPointVisible
        && (bubbleScale <= 0
          || visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]) == null)) continue;
      const point = overrides.get(index);
      const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
      if (symbol === 'none') continue;
      const resolvedMarker = isBubble
        ? bubblePointLegendMarker(chart, series, point, index)
        : classicPointLegendMarker(
            chart, series, point, index, seriesIndex, family,
            effectiveScatterStyle, effectiveRadarStyle,
          );
      const fillPaint = resolvedMarker
        ? resolvedMarker.fillPaint
        : (isBubble ? undefined : markerFillPaintFor(series, point, index));
      const linePaint = resolvedMarker?.linePaint;
      const markerHasFill = markerSymbolConsumesFill(resolvedMarker?.symbol ?? symbol);
      const chargeMarker = (sizePx: number): boolean =>
        (!markerHasFill || chargePaint(fillPaint, 1, sizePx))
        && chargePaint(linePaint, 1, sizePx)
        && (!isBubble || resolvedMarker?.bubble3D !== true || fillPaint === null
          || chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS));
      let sizePx = Math.max(2, (point?.markerSize ?? series.markerSize ?? 5) * ptToPx);
      if (family === 'scatter' && isBubble) {
        const size = visibleBubbleSize(bubbleSettings!, series.bubbleSizes?.[index]);
        sizePx = size == null ? 0 : bubbleSizeMagnitude(bubbleSettings!, size) * bubbleScale;
      } else if (family === 'radar' && point?.markerSize == null && series.markerSize == null
        && chartRect) {
        sizePx = Math.max(4 * ptToPx, Math.min(chartRect.w, chartRect.h) * 0.025);
      }
      if (plotPointVisible && !chargeMarker(sizePx)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (legendKeyVisible && !chargeMarker(keySizes.legend)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (tableKeyVisible && !chargeMarker(keySizes.table)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
      if (labelKeyVisible && !chargeMarker(keySizes.labels)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
    }

    const seriesKeySymbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
    if (!pointDrivenKeys && markerSymbolConsumesFill(seriesKeySymbol) && seriesLegendMarkerIsVisible(
      effectiveChartType, effectiveScatterStyle, series, effectiveRadarStyle,
    )) {
      const legendEntryVisible = legendEntryIsVisible(
        legendRanges, deletedLegendEntries, seriesIndex,
      );
      const labelKeys = dataLabelLegendKeyCount(
        chart, series, family, pointCount, scatterHasNumericX, markerContext,
      );
      const bubbleKeyFill = isBubble
        ? bubblePointFill(
            chart, series, undefined, seriesIndex,
            chartExSeriesFormatIndex(series, seriesIndex), chartColor(seriesIndex, series),
          )
        : null;
      const keyPaint = isBubble ? bubbleKeyFill!.paint : seriesMarkerFillPaint(series);
      const bubbleKeyLine = isBubble
        ? bubblePointLine(
            chart, series, undefined, seriesIndex,
            chartExSeriesFormatIndex(series, seriesIndex),
          )
        : null;
      const chargeKey = (repetitions: number, sizePx: number): boolean =>
        chargePaint(keyPaint, repetitions, sizePx)
        && (!isBubble || chargePaint(bubbleKeyLine!.paint, repetitions, sizePx))
        && (!isBubble || !bubblePointIsThreeD(series, undefined)
          || bubbleKeyFill!.paint === null
          || chargeComponents(BUBBLE_3D_MATERIAL_COMPONENTS, repetitions));
      if ((chart.showLegend && legendEntryVisible
          && !chargeKey(1, keySizes.legend))
        || (dataTableMarkerKeysVisible && !chargeKey(1, keySizes.table))
        || !chargeKey(labelKeys, keySizes.labels)) {
        return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
      }
    }
  }
  const box = chart.chartexBox;
  if (box) {
    const seriesCount = box.series.length;
    const markerStyle = chart.chartexDataPointMarkerStyle ?? chart.chartexDataPointStyle;
    const symbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
    if (markerSymbolConsumesFill(symbol)) {
      for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
        const series = box.series[seriesIndex];
        if (!series.showNonoutliers && !series.showOutliers) continue;
        let repetitions = 0;
        for (const values of series.valuesByCategory) {
          const stats = computeBoxWhiskerStats(values, series.quartileMethod);
          if (!stats) continue;
          if (series.showNonoutliers) repetitions += stats.inner.length;
          if (series.showOutliers) repetitions += stats.outliers.length;
        }
        const styleIndex = chartExSeriesFormatIndex(series, seriesIndex);
        const paint = chartExMarkerPaint(
          chart, styleIndex, seriesCount, series.chartexStyle, series.color, markerStyle,
        );
        if (!chargePaint(paint, repetitions, Math.max(2, 3 * ptToPx))) {
          return MAX_CANVAS_MARKER_PAINT_COMPONENTS + 1;
        }
      }
    }
  }
  return total;
}


/** Estimate the expanded synchronous paint work for classic 3-D families.
 * A source point is not one Canvas primitive: a pie slice emits up to 32 wall
 * quads plus its top face, and a round/tapered bar emits a bounded revolved
 * mesh whose cap+facet count is owned by the 3-D renderer.
 * Apply the same 10k availability budget to that derived work before arrays
 * and sort keys are allocated. */
export function classicThreeDWorkCount(
  chart: ChartModel,
  threeD: ChartThreeDRenderer | undefined,
): number | null {
  // The expanded-face budget belongs to the optional mesh renderer. Without
  // the renderer this chart intentionally follows its canonical 2-D family, so
  // rejecting it by a cost that will never be allocated would make the
  // tree-shaken fallback less capable than an ordinary 2-D chart.
  if (!chart.threeD || !threeD) return null;
  if (!CLASSIC_THREE_D_FAMILIES.has(chart.chartType)) return null;
  let total = 0;
  for (const series of chart.series) {
    const points = Math.max(1, series.values.length, series.categories?.length ?? 0);
    const shape = series.threeDShape ?? chart.threeD.shape ?? 'box';
    const weight = chart.chartType === 'pie'
      ? THREE_D_MAX_SHAPE_FACES_PER_DATUM
      : chart.chartType.toLowerCase().includes('bar')
        ? (shape === 'box' ? 4 : THREE_D_MAX_SHAPE_FACES_PER_DATUM)
        // A normal area interval contributes only a small bounded set of
        // visible slab faces. Axis clipping can split it further, but that
        // data-dependent amplification is enforced by the renderer's exact
        // cumulative scene budget. Charging the worst-case split here would
        // reject ordinary charts long before they approach the real limit.
        : chart.chartType.toLowerCase().includes('area') ? 4
          : series.smooth === true ? 25 : 3;
    if (!Number.isSafeInteger(points) || points > Math.floor(MAX_CANVAS_CHART_POINTS / weight)) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    total += points * weight;
    if (total > MAX_CANVAS_CHART_POINTS) return MAX_CANVAS_CHART_POINTS + 1;
  }
  return total;
}


export function rejectOversizedCanvasChart(
  ctx: CanvasRenderingContext2D,
  rect: ChartRect,
  pointCount: number,
): boolean {
  if (pointCount <= MAX_CANVAS_CHART_POINTS) return false;
  ctx.fillStyle = '#888';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('(too many data points)', rect.x + rect.w / 2, rect.y + rect.h / 2);
  return true;
}
