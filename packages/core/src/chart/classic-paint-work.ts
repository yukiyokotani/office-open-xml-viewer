import type { Fill } from '../types/common.js';
import type { ChartDataPointOverride, ChartModel, ChartRect, ChartSeries } from '../types/chart.js';
import {
  classicDataPointFillDecision,
  classicDataPointLineStyle,
} from './classic-data-point-style.js';
import {
  chartSeriesVariesByPoint,
  withChartStyleIndexCache,
} from './effective-style.js';
import { withSparseStyleIndexCache } from './sparse-style-index.js';
import {
  chartDataTableFamilyIsPainted,
  classicMarkerPointIsPainted,
  dataLabelLegendKeyCount,
  markerPaintComponents,
} from './marker-style.js';
import {
  deletedLegendEntryIndices,
  legendEntryIsVisible,
  legendEntryRanges,
  legendSeriesHasVisibleEntry,
} from './legend-entry-plan.js';
import { dataLabelIsDeleted } from './data-label-style.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from './plot-groups.js';
import { planOfPieSecondaryIndices } from './of-pie.js';
import {
  chartImageFillPaintWorkUpperBound,
  type ChartImageLookup,
} from './image-fill.js';
import {
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
} from './resource-limits.js';
import { chartTextFontSizePx } from './layout.js';
import { axisLineWidthPx } from './axis-style.js';
import { pptxPresetDashArray } from '../draw/dash.js';
import { chartStockBarFillDecision } from './style-paint.js';

const FILLED_LEGEND_KEY_PT = 7;

function pointOverrides(series: ChartSeries): ReadonlyMap<number, ChartDataPointOverride> {
  const result = new Map<number, ChartDataPointOverride>();
  for (const point of series.dataPointOverrides ?? []) {
    if (!result.has(point.idx)) result.set(point.idx, point);
  }
  return result;
}

function seriesStyleIndex(series: ChartSeries, sourceIndex: number): number {
  return series.chartexFormatIdx ?? sourceIndex;
}

function finiteNonzero(value: number | null | undefined): boolean {
  return value != null && Number.isFinite(value) && value !== 0;
}

function lineFamily(family: string): boolean {
  return family === 'line' || family === 'stackedLine' || family === 'stackedLinePct'
    || family === 'scatter' || family === 'radar' || family === 'stock';
}

/**
 * Aggregate structured-paint work for classic 2-D marks and their non-marker
 * keys. Marker glyphs have a separate size-aware budget; this counter covers
 * the remaining repeated `resolveFill` consumers (bars, pie-family slices,
 * areas, varying line segments, stock ticks and fill/line legend keys).
 *
 * The count is intentionally an upper bound at ambiguous clipping boundaries,
 * but it follows group ownership, varyColors domains and zero/null visibility.
 * It therefore rejects atomically without making chart layout or style output
 * depend on how much of a maliciously expensive chart happened to paint first.
 */
function classicDataMarkPaintWorkCountImpl(
  chart: ChartModel,
  imageLookup?: ChartImageLookup,
  ptToPx = 4 / 3,
  chartRect?: ChartRect,
  renderedByThreeD = false,
): number | null {
  if (renderedByThreeD || chart.series.length === 0) return null;
  const groupBySeries = indexChartPlotGroups(chart);
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  const legendRanges = legendEntryRanges(chart, true);
  const scatterHasNumericX = chart.series.some((series, index) => {
    const family = markerChartTypeForPlotGroup(chart.chartType, groupBySeries[index]);
    return family === 'scatter' && (series.categories ?? chart.categories)
      .some(value => Number.isFinite(Number.parseFloat(value)));
  });
  const tableKeys = chart.dataTable?.showKeys === true
    && chartDataTableFamilyIsPainted(chart.chartType)
    && (chart.categories.length > 0 || chart.series.some(series =>
      (series.categories?.length ?? series.values.length) > 0
    ));
  let total = 0;
  const chargePaint = (
    paint: Fill | null | undefined,
    width = chartRect?.w ?? 1,
    height = chartRect?.h ?? 1,
  ): boolean => {
    if (paint == null) return true;
    const components = paint.fillType === 'image'
      ? chartImageFillPaintWorkUpperBound(paint, imageLookup, width, height, ptToPx)
      : markerPaintComponents(paint);
    if (paint.fillType === 'gradient' && components > MAX_CHART_PAINT_RECIPE_COMPONENTS) {
      return false;
    }
    if (components > MAX_CHART_PAINT_COMPONENTS - total) return false;
    total += components;
    return true;
  };
  const chargeDatum = (
    series: ChartSeries,
    point: ChartDataPointOverride | undefined,
    styleIndex: number,
    pointIndex: number | undefined,
    role: 'dataPoint' | 'dataPointLine' = 'dataPoint',
    fill = role === 'dataPoint',
    width = chartRect?.w ?? 1,
    height = chartRect?.h ?? 1,
  ): boolean => (!fill || chargePaint(classicDataPointFillDecision(
    chart, series, point, styleIndex, pointIndex,
  ), width, height)) && chargePaint(classicDataPointLineStyle(
    chart, role, series, point, styleIndex,
  ).paint, width, height);
  const keySize = (
    series: ChartSeries,
    point: ChartDataPointOverride | undefined,
    styleIndex: number,
    kind: 'legend' | 'table' | 'label',
    role: 'dataPoint' | 'dataPointLine',
  ): { width: number; height: number } => {
    const tableFont = chartTextFontSizePx(chart.dataTable?.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
    let labelFont = chartTextFontSizePx(
      series.seriesDataLabels?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
      ptToPx,
    ) ?? 10 * ptToPx;
    for (const label of series.dataLabelOverrides ?? []) {
      labelFont = Math.max(
        labelFont,
        chartTextFontSizePx(label.fontSizeHpt, ptToPx) ?? labelFont,
      );
    }
    const font = kind === 'table' ? tableFont : kind === 'label'
      ? labelFont
      : chartTextFontSizePx(chart.legendFontSizeHpt, ptToPx) ?? 10 * ptToPx;
    if (kind === 'table') {
      return { width: Math.max(12 * ptToPx, font * 1.7), height: font };
    }
    if (role !== 'dataPointLine') {
      const side = FILLED_LEGEND_KEY_PT * ptToPx;
      return { width: side, height: side };
    }
    const line = classicDataPointLineStyle(chart, role, series, point, styleIndex);
    const lineWidth = axisLineWidthPx(line.widthEmu, ptToPx);
    const dash = pptxPresetDashArray(line.dash ?? 'solid', lineWidth);
    const completeDash = dash.length > 0
      ? dash.reduce((sum, length) => sum + length, 0) + dash[0]!
      : 0;
    return { width: Math.max(font * 1.6, completeDash), height: font };
  };

  for (let sourceIndex = 0; sourceIndex < chart.series.length; sourceIndex++) {
    const series = chart.series[sourceIndex]!;
    const group = groupBySeries[sourceIndex];
    let family = markerChartTypeForPlotGroup(
      series.seriesType ?? chart.chartType,
      group,
    );
    if (group?.kind === 'bar3D' || group?.kind === 'area3D' || group?.kind === 'line3D') {
      family = chart.chartType;
    } else if (group?.kind === 'pie3D') {
      family = 'pie';
    }
    const directPoints = pointOverrides(series);
    const varies = chartSeriesVariesByPoint(chart, sourceIndex);
    const seriesIndex = seriesStyleIndex(series, sourceIndex);
    const styleIndexAt = (pointIndex: number): number => varies ? pointIndex : seriesIndex;
    const flatBar = group?.kind === 'bar' || group?.kind === 'bar3D' || (group == null && (
      family === 'clusteredBar' || family === 'clusteredBarH'
      || family === 'stackedBar' || family === 'stackedBarH'
      || family === 'stackedBarPct' || family === 'stackedBarHPct'
    ));
    const pieFamily = group?.kind === 'pie' || group?.kind === 'pie3D' || group?.kind === 'doughnut'
      || group?.kind === 'ofPie' || (group == null && (
        family === 'pie' || family === 'doughnut' || family === 'ofPie'
      ));
    const areaFamily = group?.kind === 'area' || group?.kind === 'area3D' || (group == null && (
      family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct'
    ));
    const filledRadar = family === 'radar' && (group?.radarStyle ?? chart.radarStyle) === 'filled';

    if (flatBar || pieFamily) {
      const pointCount = Math.max(1, series.values.length, chart.categories.length);
      const horizontal = group?.barDirection === 'bar' || family.endsWith('BarH')
        || family.endsWith('BarHPct');
      const markWidth = flatBar && !horizontal
        ? (chartRect?.w ?? 1) / pointCount : chartRect?.w ?? 1;
      const markHeight = flatBar && horizontal
        ? (chartRect?.h ?? 1) / pointCount : chartRect?.h ?? 1;
      for (let pointIndex = 0; pointIndex < series.values.length; pointIndex++) {
        if (!finiteNonzero(series.values[pointIndex])) continue;
        if (!chargeDatum(
          series, directPoints.get(pointIndex), styleIndexAt(pointIndex), pointIndex,
          'dataPoint', true, markWidth, markHeight,
        )) return MAX_CHART_PAINT_COMPONENTS + 1;
      }
      if (family === 'ofPie') {
        const secondary = planOfPieSecondaryIndices(chart.ofPie, series.values);
        const aggregateSource = secondary == null ? undefined : [...secondary].find(pointIndex =>
          finiteNonzero(series.values[pointIndex])
        );
        if (aggregateSource != null && !chargeDatum(
          series,
          directPoints.get(aggregateSource),
          styleIndexAt(aggregateSource),
          aggregateSource,
          'dataPoint',
          true,
          chartRect?.w ?? 1,
          chartRect?.h ?? 1,
        )) return MAX_CHART_PAINT_COMPONENTS + 1;
      }
    } else if (areaFamily || filledRadar) {
      const geometry = filledRadar
        ? series.values.length >= 3 && series.values.every(value =>
            value != null && Number.isFinite(value)
          )
        : series.values.length > 0;
      if (geometry && !chargeDatum(series, undefined, seriesIndex, undefined)) {
        return MAX_CHART_PAINT_COMPONENTS + 1;
      }
    }

    const bubbleFamily = group?.kind === 'bubble'
      || (group == null && chart.chartType === 'bubble');
    // Every non-bubble scatter style currently reaches the path painter. In
    // particular Office's/default `marker` style is a smooth connecting line
    // plus markers; it is not a marker-only geometry sentinel.
    const lineGeometryVisible = family !== 'scatter' || !bubbleFamily;
    if (lineFamily(family) && !filledRadar && lineGeometryVisible) {
      const varyingSegments = varies && family !== 'stock';
      if (varyingSegments) {
        const stackedLine = family === 'stackedLine' || family === 'stackedLinePct'
          || (group?.kind === 'line'
            && (group.grouping === 'stacked' || group.grouping === 'percentStacked'));
        const nullAsZero = stackedLine || (family === 'line' && chart.dispBlanksAs === 'zero');
        const span = !nullAsZero && chart.dispBlanksAs === 'span';
        let run: number[] = [];
        const chargeRun = (): boolean => {
          for (let index = 1; index < run.length; index++) {
            const pointIndex = run[index]!;
            if (!chargeDatum(
              series, directPoints.get(pointIndex), pointIndex, pointIndex,
              'dataPointLine', false,
            )) return false;
          }
          if (family === 'radar' && run.length > 1) {
            const pointIndex = run[0]!;
            if (!chargeDatum(
              series, directPoints.get(pointIndex), pointIndex, pointIndex,
              'dataPointLine', false,
            )) return false;
          }
          run = [];
          return true;
        };
        for (let pointIndex = 0; pointIndex < series.values.length; pointIndex++) {
          if (series.sourceHidden?.[pointIndex] === true) {
            if (!chargeRun()) return MAX_CHART_PAINT_COMPONENTS + 1;
            continue;
          }
          const value = series.values[pointIndex];
          if (value != null && Number.isFinite(value) || nullAsZero) {
            run.push(pointIndex);
          } else if (!span && !chargeRun()) {
            return MAX_CHART_PAINT_COMPONENTS + 1;
          }
        }
        if (!chargeRun()) return MAX_CHART_PAINT_COMPONENTS + 1;
      } else if (!chargeDatum(
        series, undefined, seriesIndex, undefined, 'dataPointLine', false,
      )) return MAX_CHART_PAINT_COMPONENTS + 1;
    }

    // Legend/data-table/data-label keys repeat the same non-marker fill/line
    // recipe. Marker glyph paint is deliberately left to the marker budget.
    const pointDrivenKeys = pieFamily || varies;
    const labelKeyCount = dataLabelLegendKeyCount(
      chart, series, family,
      Math.max(series.values.length, series.categories?.length ?? 0, chart.categories.length),
      scatterHasNumericX,
      {
        chartType: markerChartTypeForPlotGroup(chart.chartType, group),
        bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
        showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
      }, sourceIndex,
    );
    const seriesLegendVisible = legendSeriesHasVisibleEntry(
      legendRanges, deletedLegendEntries, sourceIndex,
    );
    const hasAnyKey = family !== 'bubble'
      && ((chart.showLegend && seriesLegendVisible) || tableKeys || labelKeyCount > 0);
    if (hasAnyKey) {
      const role = lineFamily(family) ? 'dataPointLine' : 'dataPoint';
      const fill = role === 'dataPoint';
      const chargeKey = (
        kind: 'legend' | 'table' | 'label',
        point: ChartDataPointOverride | undefined,
        styleIndex: number,
        pointIndex: number | undefined,
      ): boolean => {
        const size = keySize(series, point, styleIndex, kind, role);
        return chargeDatum(
          series, point, styleIndex, pointIndex, role, fill, size.width, size.height,
        );
      };
      if (pointDrivenKeys) {
        const keyPointCount = Math.max(
          series.values.length, series.categories?.length ?? 0, chart.categories.length,
        );
        const labelOverrides = new Map(
          (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
        );
        for (let pointIndex = 0; pointIndex < keyPointCount; pointIndex++) {
          const styleIndex = styleIndexAt(pointIndex);
          const point = directPoints.get(pointIndex);
          const label = labelOverrides.get(pointIndex);
          const labelKey = !dataLabelIsDeleted(series.seriesDataLabels, label)
            && (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey ?? false) === true
            && classicMarkerPointIsPainted(
              chart, series, family, pointIndex, scatterHasNumericX,
              {
                chartType: markerChartTypeForPlotGroup(chart.chartType, group),
                bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
                showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
              },
            );
          if (chart.showLegend && legendEntryIsVisible(
            legendRanges, deletedLegendEntries, sourceIndex, pointIndex,
          ) && !chargeKey('legend', point, styleIndex, pointIndex)) {
            return MAX_CHART_PAINT_COMPONENTS + 1;
          }
          if (labelKey && !chargeKey('label', point, styleIndex, pointIndex)) {
            return MAX_CHART_PAINT_COMPONENTS + 1;
          }
        }
        const tablePoint = directPoints.get(0);
        if (tableKeys && !chargeKey('table', tablePoint, styleIndexAt(0), 0)) {
          return MAX_CHART_PAINT_COMPONENTS + 1;
        }
      } else {
        if (chart.showLegend && seriesLegendVisible
          && !chargeKey('legend', undefined, seriesIndex, undefined)) {
          return MAX_CHART_PAINT_COMPONENTS + 1;
        }
        if (tableKeys && !chargeKey('table', undefined, seriesIndex, undefined)) {
          return MAX_CHART_PAINT_COMPONENTS + 1;
        }
        for (let copy = 0; copy < labelKeyCount; copy++) {
          if (!chargeKey('label', undefined, seriesIndex, undefined)) {
            return MAX_CHART_PAINT_COMPONENTS + 1;
          }
        }
      }
    }
  }

  const chargeUpDownBars = (
    members: readonly ChartSeries[],
    style: NonNullable<ChartModel['stockUpDownBarStyle']>,
  ): boolean => {
    const start = members[0];
    const end = members.at(-1);
    if (!start || !end) return true;
    const count = Math.max(start.values.length, end.values.length);
    for (let index = 0; index < count; index++) {
      const first = start.values[index];
      const last = end.values[index];
      if (first == null || last == null || !Number.isFinite(first) || !Number.isFinite(last)
        || first === last) continue;
      const direction = last > first ? 'upBar' : 'downBar';
      const paint = last > first ? style.up : style.down;
      const horizontal = chart.chartType.endsWith('BarH')
        || chart.chartType.endsWith('BarHPct');
      const width = horizontal ? chartRect?.w ?? 1 : (chartRect?.w ?? 1) / count;
      const height = horizontal ? (chartRect?.h ?? 1) / count : chartRect?.h ?? 1;
      if (!chargePaint(
        chartStockBarFillDecision(chart, paint, direction), width, height,
      )) return false;
    }
    return true;
  };
  for (const decoration of chart.lineGroupDecorations ?? []) {
    if (!decoration.upDownBars) continue;
    let members = chart.series.filter(series => series.lineGroupIndex === decoration.groupIndex);
    if (members.length === 0 && decoration.groupIndex === 0) members = chart.series;
    if (!chargeUpDownBars(members, decoration.upDownBars)) {
      return MAX_CHART_PAINT_COMPONENTS + 1;
    }
  }
  if (chart.stockUpDownBars) {
    const stockGroup = chart.plotGroups?.find(group => group.kind === 'stock');
    const members = stockGroup
      ? chart.series.slice(stockGroup.seriesStart, stockGroup.seriesStart + stockGroup.seriesCount)
      : chart.series;
    if (!chargeUpDownBars(
      members,
      chart.stockUpDownBarStyle ?? { gapWidthPercent: 150, up: {}, down: {} },
    )) return MAX_CHART_PAINT_COMPONENTS + 1;
  }
  return total;
}

/** @internal Exported for resource-boundary regression tests. */
export function classicDataMarkPaintWorkCount(
  chart: ChartModel,
  imageLookup?: ChartImageLookup,
  ptToPx = 4 / 3,
  chartRect?: ChartRect,
  renderedByThreeD = false,
): number | null {
  return withChartStyleIndexCache(() => withSparseStyleIndexCache(() =>
    classicDataMarkPaintWorkCountImpl(
      chart, imageLookup, ptToPx, chartRect, renderedByThreeD,
    )
  ));
}
