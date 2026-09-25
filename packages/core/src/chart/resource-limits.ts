import type {
  ChartDataLabelOverride, ChartExElementStyle, ChartModel, ChartSeries,
} from '../types/chart.js';
import { dataLabelIsDeleted } from './data-label-style.js';
import {
  deletedLegendEntryIndices,
  legendEntryRanges,
} from './legend-entry-plan.js';
import {
  bubblePointIsThreeD,
  classicDataLabelPointIsPainted,
  classicMarkerPointIsPainted,
  chartDataTableFamilyIsPainted,
  dataLabelLegendKeyCount,
  effectiveMarkerSymbol,
  markersSuppressedByChartStyle,
  seriesHasMarkerDetail,
  seriesLegendMarkerIsVisible,
} from './marker-style.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from './plot-groups.js';
import { chartSeriesVariesByPoint } from './effective-style.js';
import { chartDataPointStyleRole } from './effective-style.js';
import { chartStyleEffectOwner, chartStyleEffectRecipe } from './style-effects.js';
import { computeBoxWhiskerStats } from './box-whisker.js';
import { planOfPieSecondaryIndices } from './of-pie.js';
import {
  visitChartExHierarchyBodySites,
  visitChartExHierarchyLabelSites,
} from './chart-ex-hierarchy-labels.js';
import { planWaterfallPaintSites } from './waterfall-plan.js';
import { chartLabelBoxHasVisiblePaint, mergeChartLabelBoxes } from './label-box.js';

/** Shared synchronous Canvas chart point ceiling. Host prefetch and renderer
 * preflight must reject against the same bound before allocating per-point work. */
export const MAX_CANVAS_CHART_POINTS = 10_000;

/** Maximum decoded picture-marker sources retained for one host render pass.
 * This matches the shared decoded-image cache count boundary. */
export const MAX_CHART_MARKER_IMAGE_SOURCES = 256;

/** Shared structured-paint availability ceilings. A single gradient recipe
 * and the aggregate chart work use the same bounds across classic markers,
 * labels, Surface bands, and the optional 3-D renderer. */
export const MAX_CHART_PAINT_RECIPE_COMPONENTS = 4_096;
export const MAX_CHART_PAINT_COMPONENTS = 1_048_576;
/** Maximum drawImage repetitions for one tiled/stacked chart picture fill. */
export const MAX_CHART_IMAGE_FILL_TILES = 4_096;

function ownsResolvedRasterEffect(
  direct: ChartExElementStyle | null | undefined,
  fallback: ChartExElementStyle | null | undefined,
  index: number,
  fallbackIndex = index,
): boolean {
  return chartStyleEffectRecipe(direct, fallback, index, fallbackIndex) != null;
}

function seriesFamily(
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  groups: ReturnType<typeof indexChartPlotGroups>,
): string {
  const group = groups[seriesIndex];
  if (group?.kind === 'bubble' || group?.kind === 'scatter') return 'scatter';
  if (group?.kind === 'line') return chart.chartType === 'stackedLinePct'
    ? 'stackedLinePct' : chart.chartType === 'stackedLine' ? 'stackedLine' : 'line';
  if (group?.kind === 'area') return chart.chartType === 'stackedAreaPct'
    ? 'stackedAreaPct' : chart.chartType === 'stackedArea' ? 'stackedArea' : 'area';
  return group?.kind ?? series.seriesType
    ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
}

function lineDecorationMembers(
  chart: ChartModel,
  groupIndex: number,
): readonly ChartSeries[] {
  const explicit = chart.series.filter(series => series.lineGroupIndex === groupIndex);
  if (explicit.length > 0 || groupIndex !== 0
    || !['line', 'stackedLine', 'stackedLinePct'].includes(chart.chartType)) return explicit;
  return chart.series.filter(series => series.seriesType == null || series.seriesType === 'line');
}

function dataLabelHasContent(
  chart: ChartModel,
  series: ChartSeries,
  label: ChartDataLabelOverride | undefined,
): boolean {
  return Boolean(
    label?.text
    || (label?.showVal ?? series.seriesDataLabels?.showVal ?? chart.showDataLabels)
    || (label?.showCatName ?? series.seriesDataLabels?.showCatName)
    || (label?.showSerName ?? series.seriesDataLabels?.showSerName)
    || (label?.showPercent ?? series.seriesDataLabels?.showPercent)
    || (label?.showBubbleSize ?? series.seriesDataLabels?.showBubbleSize)
    || (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey)
  );
}

function seriesEffectBodyRole(
  chart: ChartModel,
  seriesIndex: number,
  family: string,
  groups: ReturnType<typeof indexChartPlotGroups>,
): 'dataPoint' | 'dataPointLine' | null {
  if (chart.threeD || groups[seriesIndex]?.kind.endsWith('3D') === true
    || chartSeriesVariesByPoint(chart, seriesIndex)) return null;
  if (family === 'line' || family === 'stackedLine' || family === 'stackedLinePct') {
    return 'dataPointLine';
  }
  if (family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct') {
    return 'dataPoint';
  }
  if (family === 'radar') {
    const radarStyle = groups[seriesIndex]?.radarStyle ?? chart.radarStyle;
    return radarStyle === 'filled' ? 'dataPoint' : 'dataPointLine';
  }
  return null;
}

/** Conservative count of effect consumers that can actually own a raster or
 * composed-shadow recipe. Unstyled source points deliberately do not dilute a
 * chart-frame effect's allowance: only an effect-bearing source role/direct
 * style contributes a source-sized domain. The count can overestimate visible
 * styled marks, but cannot grow merely because unrelated data was added. */
export function chartEffectConsumerUpperBound(chart: ChartModel): number {
  const classicPainted = classicCanvasPointFamilyIsPainted(chart.chartType);
  const groups = indexChartPlotGroups(chart);
  const families = chart.series.map((series, index) =>
    seriesFamily(chart, series, index, groups)
  );
  const scatterHasNumericX = chart.series.some((series, index) =>
    families[index] === 'scatter'
      && (series.categories ?? chart.categories).some(category =>
        Number.isFinite(Number.parseFloat(category))
      )
  );
  let consumers = 0;
  const add = (
    direct: ChartExElementStyle | null | undefined,
    fallback: ChartExElementStyle | null | undefined,
    index = 0,
    fallbackIndex = index,
  ): void => {
    if (ownsResolvedRasterEffect(direct, fallback, index, fallbackIndex)) consumers++;
  };

  // Count effective paint sites, not authored carriers. DrawingML effect
  // lists are atomic: an authored empty/unsupported direct list suppresses the
  // linked role and therefore contributes no raster work.
  add(chart.chartAreaStyle, chart.chartStyleRoles?.chartArea);
  add(chart.plotAreaStyle, chart.chartStyleRoles?.[chart.threeD ? 'plotArea3D' : 'plotArea']);
  if (chart.showLegend) add(chart.legendStyle, chart.chartStyleRoles?.legend);
  if (chart.titlePresent === true || chart.title != null
    || (chart.titleRichRuns?.length ?? 0) > 0) add(chart.titleStyle, chart.chartStyleRoles?.title);
  if (chart.catAxisTitle != null) add(chart.catAxisTitleStyle, chart.chartStyleRoles?.axisTitle);
  if (chart.valAxisTitle != null) add(chart.valAxisTitleStyle, chart.chartStyleRoles?.axisTitle);
  for (const axis of [chart.secondaryValAxis, chart.secondaryCatAxis]) {
    if (axis?.title != null) add(axis.titleStyle, chart.chartStyleRoles?.axisTitle);
    if (axis?.displayUnits?.label != null) {
      add(axis.displayUnits.label.boxStyle?.style, chart.chartStyleRoles?.axisTitle);
    }
  }
  if (chart.threeD?.seriesAxis?.title != null) {
    add(chart.threeD.seriesAxis.titleStyle, chart.chartStyleRoles?.axisTitle);
  }
  if (chart.valAxisDisplayUnits?.label != null) {
    add(chart.valAxisDisplayUnits.label.boxStyle?.style, chart.chartStyleRoles?.axisTitle);
  }
  if (chart.catAxisDisplayUnits?.label != null) {
    add(chart.catAxisDisplayUnits.label.boxStyle?.style, chart.chartStyleRoles?.axisTitle);
  }
  visitChartExHierarchyLabelSites(chart, ({ label, linkedStyleIndex }) => {
    const direct = label.labelBox;
    const linked = chartLabelBoxHasVisiblePaint(direct)
      ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
      : chart.chartStyleRoles?.dataLabel;
    add(direct?.style, linked, 0, linkedStyleIndex);
  });
  if (chart.chartType === 'waterfall') {
    const series = chart.series[0];
    const overrides = new Map((series?.dataPointOverrides ?? []).map(point => [point.idx, point]));
    const plan = planWaterfallPaintSites(
      series?.values ?? [], chart.categories.length, chart.subtotalIndices,
    );
    if (!plan.cumulativeOverflow && plan.rawMax > plan.rawMin) {
      for (let index = 0; index < plan.bars.length; index++) {
        const bar = plan.bars[index]!;
        if (!bar.paintSlot) continue;
        add(
          chartStyleEffectOwner(overrides.get(index)?.chartexStyle, series?.chartexStyle),
          chart.chartexDataPointStyle,
          bar.semanticIndex,
          bar.semanticIndex,
        );
      }
    }
  } else if (chart.chartType === 'funnel') {
    const series = chart.series[0];
    const values = series?.values ?? [];
    if (values.some(value => value != null && value > 0)) {
      const count = Math.max(values.length, chart.categories.length);
      for (let index = 0; index < count; index++) {
        if (!((values[index] ?? 0) > 0)) continue;
        add(series?.chartexStyle, chart.chartexDataPointStyle, 0, 0);
      }
    }
  }
  for (let seriesIndex = 0; seriesIndex < (chart.chartexBox?.series.length ?? 0); seriesIndex++) {
    const box = chart.chartexBox!.series[seriesIndex]!;
    const styleIndex = box.chartexFormatIdx ?? seriesIndex;
    for (const values of box.valuesByCategory) {
      if (computeBoxWhiskerStats(values, box.quartileMethod)) {
        add(box.chartexStyle, chart.chartexDataPointStyle, styleIndex, styleIndex);
      }
    }
  }
  const hierarchySeries = chart.series[0];
  visitChartExHierarchyBodySites(chart, ({ node, paintsBody }) => {
    if (paintsBody) add(
      hierarchySeries?.chartexStyle,
      chart.chartexDataPointStyle,
      hierarchySeries?.chartexFormatIdx ?? 0,
      node.branchIndex,
    );
  });

  const legendRanges = legendEntryRanges(chart);
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const series = chart.series[seriesIndex]!;
    const family = families[seriesIndex]!;
    const group = groups[seriesIndex];
    const effectiveChartType = markerChartTypeForPlotGroup(chart.chartType, group);
    const styleIndex = series.chartexFormatIdx ?? seriesIndex;
    const points = new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point]));
    const bodyRole = seriesEffectBodyRole(chart, seriesIndex, family, groups);
    const varies = chartSeriesVariesByPoint(chart, seriesIndex);
    const optionalThreeD = chart.threeD != null || group?.kind.endsWith('3D') === true;
    const bubbleFamily = group?.kind === 'bubble'
      || (group == null && chart.chartType === 'bubble');
    const radarFilled = family === 'radar'
      && (group?.radarStyle ?? chart.radarStyle) === 'filled';
    const scatterDrawsLine = family === 'scatter' && !bubbleFamily;
    const segmentedLine = varies && (
      family === 'line' || family === 'stackedLine' || family === 'stackedLinePct'
      || (family === 'radar' && !radarFilled) || scatterDrawsLine
    );
    const segmentEffectPoints = new Set<number>();
    if (segmentedLine) {
      if (family === 'scatter') {
        const visible: number[] = [];
        for (let pointIndex = 0; pointIndex < series.values.length; pointIndex++) {
          if (classicMarkerPointIsPainted(
            chart, series, family, pointIndex, scatterHasNumericX, {
              chartType: effectiveChartType,
            },
          )) visible.push(pointIndex);
        }
        for (let index = 1; index < visible.length; index++) {
          segmentEffectPoints.add(visible[index]!);
        }
      } else {
        const stackedLine = family === 'stackedLine' || family === 'stackedLinePct'
          || (group?.kind === 'line'
            && (group.grouping === 'stacked' || group.grouping === 'percentStacked'));
        const nullAsZero = stackedLine || (family === 'line' && chart.dispBlanksAs === 'zero');
        const span = family !== 'radar' && !nullAsZero && chart.dispBlanksAs === 'span';
        let run: number[] = [];
        const finishRun = (): void => {
          for (let index = 1; index < run.length; index++) {
            segmentEffectPoints.add(run[index]!);
          }
          if (family === 'radar' && run.length === series.values.length && run.length > 1) {
            segmentEffectPoints.add(run[0]!);
          }
          run = [];
        };
        for (let pointIndex = 0; pointIndex < series.values.length; pointIndex++) {
          if (series.sourceHidden?.[pointIndex] === true) {
            finishRun();
            continue;
          }
          const value = series.values[pointIndex];
          if (value != null && Number.isFinite(value) || nullAsZero) run.push(pointIndex);
          else if (!span) finishRun();
        }
        finishRun();
      }
    }
    const areaBody = bodyRole === 'dataPoint'
      && (family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct');
    if (classicPainted && bodyRole && (!varies || areaBody)) {
      add(
        chartStyleEffectOwner(series.chartexStyle),
        chartDataPointStyleRole(chart, bodyRole, seriesIndex),
        styleIndex,
      );
    }
    if (classicPainted && varies && radarFilled && series.values.length >= 3
      && series.values.every((value, index) => value != null && Number.isFinite(value)
        && series.sourceHidden?.[index] !== true)) {
      const first = points.get(0);
      add(
        chartStyleEffectOwner(first?.chartexStyle, series.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPoint', seriesIndex),
        0,
        0,
      );
    }
    if (classicPainted && family === 'scatter' && !varies && scatterDrawsLine) {
      let visible = 0;
      for (let pointIndex = 0; pointIndex < series.values.length; pointIndex++) {
        if (classicMarkerPointIsPainted(
          chart, series, family, pointIndex, scatterHasNumericX, {
            chartType: effectiveChartType,
          },
        )) visible++;
      }
      if (visible >= 2) add(
        chartStyleEffectOwner(series.chartexStyle),
        chartDataPointStyleRole(chart, 'dataPointLine', seriesIndex),
        styleIndex,
      );
    }
    const pointCount = Math.max(
      series.values.length, series.categories?.length ?? 0,
      series.bubbleSizes?.length ?? 0, chart.categories.length,
    );
    if (classicPainted && !chart.threeD && group?.kind.endsWith('3D') !== true) {
      for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
        const point = points.get(pointIndex);
        const datumReachable = classicMarkerPointIsPainted(
          chart, series, family, pointIndex, scatterHasNumericX, {
            chartType: effectiveChartType,
            bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
            showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
          },
        );
        const markerFamily = bubbleFamily ? 'bubble' : family;
        const areaFamily = markerFamily === 'area' || markerFamily === 'stackedArea'
          || markerFamily === 'stackedAreaPct';
        const markerFamilySupported = markerFamily === 'line'
          || markerFamily === 'stackedLine' || markerFamily === 'stackedLinePct'
          || areaFamily || markerFamily === 'scatter' || markerFamily === 'bubble'
          || markerFamily === 'radar' || markerFamily === 'stock';
        const markerSuppressed = markersSuppressedByChartStyle(
          markerFamily, effectiveChartType,
          group?.scatterStyle ?? chart.scatterStyle,
          group?.radarStyle ?? chart.radarStyle,
        );
        const seriesMarkerVisible = areaFamily
          ? (series.showMarker === true || seriesHasMarkerDetail(series))
            && series.markerSymbol !== 'none'
          : markerFamily === 'stock'
            ? series.markerSymbol != null && series.markerSymbol !== 'none'
            : series.showMarker !== false && series.markerSymbol !== 'none';
        const markerPainted = datumReachable && markerFamilySupported && !markerSuppressed
          && effectiveMarkerSymbol(
            series, point, 'circle', seriesMarkerVisible,
          ) !== 'none';
        const value = series.values[pointIndex];
        const bodyPainted = datumReachable
          && (!(family === 'pie' || family === 'doughnut' || family === 'ofPie')
            || (value != null && Number.isFinite(value) && Math.abs(value) > 0));
        const stockStart = group?.kind === 'stock' ? group.seriesStart : 0;
        const stockCount = group?.kind === 'stock' ? group.seriesCount : chart.series.length;
        const stockTickSeries = family === 'stock'
          && (seriesIndex === stockStart + stockCount - 1
            || (stockCount >= 4 && seriesIndex === stockStart));
        const pointBodySite = segmentedLine
          ? segmentEffectPoints.has(pointIndex)
          : bodyRole == null && (bubbleFamily
            || (family !== 'scatter' && family !== 'radar'));
        if (pointBodySite && bodyPainted
          && (family !== 'stock' || (stockTickSeries && !markerPainted))) {
          const role = bubbleFamily
            ? (bubblePointIsThreeD(series, point) ? 'dataPoint3D' : 'dataPoint')
            : family === 'line' || family === 'stackedLine'
              || family === 'stackedLinePct' || family === 'stock'
              || family === 'scatter' || family === 'radar'
            ? 'dataPointLine' : 'dataPoint';
          add(
            chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle),
            chartDataPointStyleRole(chart, role, seriesIndex),
            family === 'stock' ? styleIndex
              : point?.chartexStyle ? pointIndex : styleIndex,
            family === 'stock' ? styleIndex : varies ? pointIndex : styleIndex,
          );
        }
        if (markerPainted && !bubbleFamily) {
          add(
            chartStyleEffectOwner(
              point?.markerStyle, point?.chartexStyle, series.markerStyle,
            ),
            chartDataPointStyleRole(chart, 'dataPointMarker', seriesIndex),
            point?.markerStyle || point?.chartexStyle ? pointIndex : styleIndex,
            varies ? pointIndex : styleIndex,
          );
        }
      }
    }
    if (classicPainted && family === 'ofPie') {
      const secondary = planOfPieSecondaryIndices(chart.ofPie, series.values);
      const aggregateSource = secondary == null ? undefined : [...secondary].find(pointIndex => {
        const value = series.values[pointIndex];
        return value != null && Number.isFinite(value) && Math.abs(value) > 0;
      });
      if (aggregateSource != null) {
        const point = points.get(aggregateSource);
        add(
          chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle),
          chartDataPointStyleRole(chart, 'dataPoint', seriesIndex),
          varies ? aggregateSource : styleIndex,
        );
      }
    }
    if (classicPainted && optionalThreeD
      && (group?.kind === 'line3D' || group?.kind === 'area3D'
        || family === 'line' || family === 'stackedLine' || family === 'stackedLinePct'
        || family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct')) {
      const areaFamily = group?.kind === 'area3D'
        || family === 'area' || family === 'stackedArea' || family === 'stackedAreaPct';
      const seriesMarkersVisible = (areaFamily
        ? series.showMarker === true || seriesHasMarkerDetail(series)
        : series.showMarker === true) && series.markerSymbol !== 'none';
      for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
        const point = points.get(pointIndex);
        const visible = seriesMarkersVisible
          || (point?.markerSymbol != null && point.markerSymbol !== 'none');
        const value = series.values[pointIndex];
        if (!visible || series.sourceHidden?.[pointIndex] === true
          || value == null || !Number.isFinite(value)) continue;
        add(
          chartStyleEffectOwner(point?.markerStyle, point?.chartexStyle, series.markerStyle),
          chartDataPointStyleRole(chart, 'dataPointMarker', seriesIndex),
          point?.markerStyle || point?.chartexStyle ? pointIndex : styleIndex,
          styleIndex,
        );
      }
    }

    const scatterLineKey = family === 'scatter' && !bubbleFamily
      && series.lineHidden !== true;
    const lineFamily = group?.kind === 'line3D'
      || family === 'line' || family === 'stackedLine'
      || family === 'stackedLinePct' || family === 'stock'
      || scatterLineKey || (family === 'radar' && !radarFilled);
    const markerOnlyKey = bubbleFamily || (family === 'scatter' && !scatterLineKey);
    const keyBodyRole = optionalThreeD ? 'dataPoint3D'
      : lineFamily ? 'dataPointLine' : 'dataPoint';
    const addKey = (pointIndex: number, pointDriven: boolean): void => {
      const point = pointDriven ? points.get(pointIndex) : undefined;
      const keySourceIndex = pointDriven ? pointIndex : seriesIndex;
      if (!markerOnlyKey) {
        const directEffect = chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle);
        add(
          directEffect,
          chartDataPointStyleRole(chart, keyBodyRole, seriesIndex),
          optionalThreeD
            ? directEffect === point?.chartexStyle ? pointIndex : styleIndex
            : point?.chartexStyle ? pointIndex : styleIndex,
          optionalThreeD ? (varies ? keySourceIndex : styleIndex)
            : pointDriven ? pointIndex : styleIndex,
        );
      }
      const markerVisible = (optionalThreeD
        ? lineFamily && series.showMarker === true && series.markerSymbol !== 'none'
        : seriesLegendMarkerIsVisible(
        effectiveChartType,
        group?.scatterStyle ?? chart.scatterStyle,
        series,
        group?.radarStyle ?? chart.radarStyle,
      )) || (point?.markerSymbol != null && point.markerSymbol !== 'none');
      if (!markerVisible) return;
      const markerRole = bubbleFamily && bubblePointIsThreeD(series, point)
        ? 'dataPoint3D' : bubbleFamily ? 'dataPoint' : 'dataPointMarker';
      add(
        chartStyleEffectOwner(
          point?.markerStyle, point?.chartexStyle,
          bubbleFamily ? series.chartexStyle : series.markerStyle,
        ),
        chartDataPointStyleRole(chart, markerRole, seriesIndex),
        point?.markerStyle || point?.chartexStyle ? pointIndex : styleIndex,
        optionalThreeD ? styleIndex : pointDriven ? pointIndex : styleIndex,
      );
    };

    // Legend/data-table/data-label keys are real paint sites too. Point-driven
    // legends resolve the same point index; series-driven keys use the series
    // format index. Deleted legend entries never reach Canvas.
    if (chart.showLegend) {
      const range = legendRanges[seriesIndex];
      if (range) for (let local = 0; local < range.count; local++) {
        if (deletedLegendEntries.has(range.firstIndex + local)) continue;
        const pointIndex = range.pointDriven ? local : 0;
        addKey(pointIndex, range.pointDriven);
      }
    }
    if (chart.dataTable?.showKeys === true && chartDataTableFamilyIsPainted(chart.chartType)) {
      addKey(0, legendRanges[seriesIndex]?.pointDriven === true);
    }
    if (dataLabelLegendKeyCount(
      chart, series, family, pointCount, scatterHasNumericX, {
        chartType: effectiveChartType,
        bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
        showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
      }, seriesIndex,
    ) > 0) {
      const labelOverrides = new Map(
        (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
      );
      for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
        const label = labelOverrides.get(pointIndex);
        if (dataLabelIsDeleted(series.seriesDataLabels, label)
          || (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey) !== true
          || !classicDataLabelPointIsPainted(
            chart, series, family, pointIndex, scatterHasNumericX, seriesIndex,
          )) continue;
        addKey(pointIndex, legendRanges[seriesIndex]?.pointDriven === true);
      }
    }

    for (const label of series.dataLabelOverrides ?? []) {
      if (series.sourceHidden?.[label.idx] === true
        || dataLabelIsDeleted(series.seriesDataLabels, label)
        || !dataLabelHasContent(chart, series, label)
        || !classicDataLabelPointIsPainted(
          chart, series, family, label.idx, scatterHasNumericX, seriesIndex,
        )) continue;
      const labelBox = mergeChartLabelBoxes(
        label.labelBox, series.seriesDataLabels?.labelBox,
      );
      add(
        chartStyleEffectOwner(label.labelBox?.style, series.seriesDataLabels?.labelBox?.style),
        chartLabelBoxHasVisiblePaint(labelBox)
          ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
          : chart.chartStyleRoles?.dataLabel,
        0,
        label.idx,
      );
    }
    const overriddenLabels = new Set((series.dataLabelOverrides ?? []).map(label => label.idx));
    if (series.seriesDataLabels?.labelBox != null || chart.chartStyleRoles?.dataLabel != null) {
      const limit = Math.max(
        series.values.length, series.categories?.length ?? 0, chart.categories.length,
      );
      for (let pointIndex = 0; pointIndex < limit; pointIndex++) {
        if (overriddenLabels.has(pointIndex)) continue;
        if (!classicDataLabelPointIsPainted(
          chart, series, family, pointIndex, scatterHasNumericX, seriesIndex,
        )) continue;
        if (!dataLabelHasContent(chart, series, undefined)) continue;
        add(
          series.seriesDataLabels?.labelBox?.style,
          chartLabelBoxHasVisiblePaint(series.seriesDataLabels?.labelBox)
            ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
            : chart.chartStyleRoles?.dataLabel,
          0,
          series.chartexFormatIdx ?? seriesIndex,
        );
      }
    }
    for (const trendline of classicPainted ? series.trendLines ?? [] : []) {
      const visible = trendline.dispEq === true || trendline.dispRSqr === true
        || Boolean(trendline.labelText)
        || trendline.labelRichRuns?.some(run => run.text.length > 0) === true;
      if (visible) add(
        trendline.labelBox?.style,
        chart.chartStyleRoles?.trendlineLabel,
        0,
        series.chartexFormatIdx ?? seriesIndex,
      );
    }
  }
  for (let seriesIndex = 0; seriesIndex < (chart.chartexBox?.series.length ?? 0); seriesIndex++) {
    const series = chart.chartexBox!.series[seriesIndex]!;
    const symbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
    if (symbol === 'none' || (!series.showNonoutliers && !series.showOutliers)) continue;
    const styleIndex = series.chartexFormatIdx ?? seriesIndex;
    for (const values of series.valuesByCategory) {
      const stats = computeBoxWhiskerStats(values, series.quartileMethod);
      if (!stats) continue;
      const observations = (series.showNonoutliers ? stats.inner.length : 0)
        + (series.showOutliers ? stats.outliers.length : 0);
      for (let index = 0; index < observations; index++) {
        add(
          chartStyleEffectOwner(series.chartexStyle),
          chart.chartStyleRoles?.dataPointMarker,
          styleIndex,
        );
      }
    }
  }
  const addUpDownBarEffects = (
    members: readonly ChartSeries[],
    style: NonNullable<ChartModel['stockUpDownBarStyle']>,
  ): void => {
    const start = members[0];
    const end = members.at(-1);
    if (!start || !end) return;
    const count = Math.max(start.values.length, end.values.length);
    for (let index = 0; index < count; index++) {
      const first = start.values[index];
      const last = end.values[index];
      if (first == null || last == null || !Number.isFinite(first) || !Number.isFinite(last)
        || first === last) continue;
      const role = last > first ? 'upBar' : 'downBar';
      const direct = role === 'upBar' ? style.up : style.down;
      add(direct.style, chart.chartStyleRoles?.[role], index);
    }
  };
  if (chart.stockUpDownBars) {
    const stockGroup = chart.plotGroups?.find(group => group.kind === 'stock');
    const stockMembers = stockGroup
      ? chart.series.slice(stockGroup.seriesStart, stockGroup.seriesStart + stockGroup.seriesCount)
      : chart.series;
    addUpDownBarEffects(
      stockMembers,
      chart.stockUpDownBarStyle ?? { gapWidthPercent: 150, up: {}, down: {} },
    );
  }
  for (const decoration of chart.lineGroupDecorations ?? []) {
    if (!decoration.upDownBars) continue;
    const members = lineDecorationMembers(chart, decoration.groupIndex);
    addUpDownBarEffects(members, decoration.upDownBars);
  }
  return Math.max(1, consumers);
}

const CLASSIC_CANVAS_POINT_FAMILIES = new Set([
  'clusteredBar', 'clusteredBarH', 'stackedBar', 'stackedBarH',
  'stackedBarPct', 'stackedBarHPct', 'clusteredColumn',
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'pie', 'doughnut', 'ofPie', 'radar', 'scatter', 'bubble', 'stock', 'surface', 'surface3D',
]);

export function classicCanvasPointFamilyIsPainted(chartType: string): boolean {
  return CLASSIC_CANVAS_POINT_FAMILIES.has(chartType);
}

/** Bound source/public-model structure before any visibility projection,
 * override maps, image prefetch, or style clones are created. */
export function sourceChartStructureCount(chart: ChartModel): number {
  let total = 0;
  const add = (count: number): boolean => {
    if (!Number.isSafeInteger(count) || count < 0
      || count > MAX_CANVAS_CHART_POINTS - total) return false;
    total += count;
    return true;
  };
  const addProduct = (left: number, right: number): boolean => {
    if (!Number.isSafeInteger(left) || left < 0 || !Number.isSafeInteger(right) || right < 0
      || (left !== 0 && right > Math.floor((MAX_CANVAS_CHART_POINTS - total) / left))) {
      return false;
    }
    total += left * right;
    return true;
  };
  if (!add(chart.legendEntries?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
  if (!add(chart.plotGroups?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
  if (chart.plotGroups != null) {
    let expectedSeriesStart = 0;
    for (const group of chart.plotGroups) {
      if (!Number.isSafeInteger(group.seriesStart) || group.seriesStart < 0
        || !Number.isSafeInteger(group.seriesCount) || group.seriesCount < 0
        || group.seriesStart !== expectedSeriesStart
        || group.seriesCount > chart.series.length - expectedSeriesStart) {
        return MAX_CANVAS_CHART_POINTS + 1;
      }
      expectedSeriesStart += group.seriesCount;
    }
    if (expectedSeriesStart !== chart.series.length) return MAX_CANVAS_CHART_POINTS + 1;
  }
  for (const series of chart.series) {
    const pointSlots = Math.max(
      1,
      chart.categories.length,
      series.values.length,
      series.categories?.length ?? 0,
      series.bubbleSizes?.length ?? 0,
      series.dataPointOverrides?.length ?? 0,
      series.dataLabelOverrides?.length ?? 0,
    );
    if (!add(pointSlots)) return MAX_CANVAS_CHART_POINTS + 1;
    if (!addProduct(series.trendLines?.length ?? 0, Math.max(1, series.values.length))) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    if (!add(series.errBars?.length ?? 0)) return MAX_CANVAS_CHART_POINTS + 1;
    for (const errorBars of series.errBars ?? []) {
      if (!add(Math.max(pointSlots, errorBars.plus.length, errorBars.minus.length))) {
        return MAX_CANVAS_CHART_POINTS + 1;
      }
    }
  }
  if (!add(chart.chartexSunburst?.rows.length ?? 0)
    || !add(chart.chartexTreemap?.rows.length ?? 0)
    || !add(chart.chartexRegionMap?.rows.length ?? 0)
    || !add(chart.chartexBox?.categories.length ?? 0)
    || !add(chart.chartexBox?.series.length ?? 0)) {
    return MAX_CANVAS_CHART_POINTS + 1;
  }
  for (const series of chart.chartexBox?.series ?? []) {
    for (const values of series.valuesByCategory) {
      if (!add(values.length)) return MAX_CANVAS_CHART_POINTS + 1;
    }
  }
  if (!add(chart.ofPie?.customSplitIndices?.length ?? 0)) {
    return MAX_CANVAS_CHART_POINTS + 1;
  }
  return total;
}

/** Count point slots expanded by classic Canvas renderers. */
export function classicCanvasPointCount(chart: ChartModel): number | null {
  if (!classicCanvasPointFamilyIsPainted(chart.chartType)) return null;
  let total = 0;
  for (const series of chart.series) {
    let errorBarPoints = 0;
    for (const errorBars of series.errBars ?? []) {
      errorBarPoints = Math.max(errorBarPoints, errorBars.plus.length, errorBars.minus.length);
    }
    const points = Math.max(
      1,
      chart.categories.length,
      series.categories?.length ?? 0,
      series.values.length,
      series.bubbleSizes?.length ?? 0,
      series.dataPointOverrides?.length ?? 0,
      series.dataLabelOverrides?.length ?? 0,
      series.trendLines?.length ?? 0,
      errorBarPoints,
    );
    if (!Number.isSafeInteger(points) || points > MAX_CANVAS_CHART_POINTS - total) {
      return MAX_CANVAS_CHART_POINTS + 1;
    }
    total += points;
  }
  return total;
}
