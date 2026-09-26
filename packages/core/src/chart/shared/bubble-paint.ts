// Classic chart bubble paint helpers.
import type { ChartModel, ChartSeries } from '../../types/chart';
import { bubblePointIsThreeD, markerFillColorFor, markerFillPaintFor } from '../marker-style.js';
import type { Fill } from '../../types/common';
import { chartDataPointStyleRole, chartSeriesSourceIndex, chartSeriesVariesByPoint, rawLinkedChartStyleRole } from '../effective-style.js';
import { chartStyleDirectFillDecision, chartStyleDirectLineDecision, chartStyleDirectNoFillDecision, chartStyleDirectNoLineDecision } from '../style-paint.js';
import { chartExStyleLinePaintDecision, chartExStylePaintDecision } from './chartex-style.js';


export function scatterPointFill(
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  index: number,
  fallbackColor: string,
): string {
  return markerFillColorFor(series, point, index, fallbackColor);
}


/** Resolve the classic bubble shape fill without collapsing DrawingML
 * provenance into the marker fallback. CT_DPt shape paint wins over CT_Ser
 * shape paint, which wins over the linked dataPoint role. */
export function bubblePointFill(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  seriesStyleIndex: number,
  fallbackColor: string,
  bubble3D = bubblePointIsThreeD(series, point),
): { color: string; paint: Fill | null | undefined } {
  const bubbleSize = series.bubbleSizes?.[pointIndex];
  if (bubbleSize != null && Number.isFinite(bubbleSize) && bubbleSize < 0) {
    // MS-OE376 §2.1.1504(b): Office always inverts a negative bubble,
    // regardless of `<c:invertIfNegative>`. The application-generated default
    // is outline-only for a flat bubble and white material for a 3-D bubble;
    // an authored c14 alternate fill remains authoritative.
    if (series.invertedFillHidden === true) return { color: '00000000', paint: null };
    if (series.invertedFill) {
      return {
        color: series.invertedFill.fillType === 'solid'
          ? series.invertedFill.color : fallbackColor,
        paint: series.invertedFill,
      };
    }
    return bubble3D
      ? { color: 'FFFFFF', paint: undefined }
      : { color: '00000000', paint: null };
  }
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint');
  const directPoint = chartStyleDirectFillDecision(
    point?.chartexStyle, rawLinked, pointIndex,
  );
  if (directPoint !== undefined) {
    return {
      color: directPoint?.fillType === 'solid' ? directPoint.color : fallbackColor,
      paint: directPoint,
    };
  }
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const linkedRole = chartDataPointStyleRole(chart, 'dataPoint', seriesIndex);
  const linkedIndex = chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex : seriesStyleIndex;
  if (point?.fillHidden === true) {
    const noFill = chartStyleDirectNoFillDecision(rawLinked);
    if (noFill !== undefined) return { color: '00000000', paint: noFill };
  }
  if (point?.color != null) return { color: point.color, paint: undefined };
  const pointColor = series.dataPointColors?.[pointIndex];
  if (pointColor != null) return { color: pointColor, paint: undefined };

  const directSeries = chartStyleDirectFillDecision(
    series.chartexStyle, rawLinked, seriesStyleIndex,
  );
  if (directSeries !== undefined) {
    return {
      color: directSeries?.fillType === 'solid' ? directSeries.color : fallbackColor,
      paint: directSeries,
    };
  }
  if (series.color != null) return { color: series.color, paint: undefined };
  const linkedPoint = chartExStylePaintDecision(
    chart,
    linkedRole,
    linkedIndex,
    series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : fallbackColor,
      paint: linkedPoint,
    };
  }
  return {
    color: scatterPointFill(series, point, pointIndex, fallbackColor),
    paint: markerFillPaintFor(series, point, pointIndex),
  };
}


export function bubblePointLine(
  chart: ChartModel,
  series: ChartSeries,
  point: NonNullable<ChartSeries['dataPointOverrides']>[number] | undefined,
  pointIndex: number,
  seriesStyleIndex: number,
): {
  color: string | null;
  paint: ChartModel['plotAreaLineFill'] | null | undefined;
  widthEmu: number | null | undefined;
  dash: string | null | undefined;
  customDash: ChartModel['plotAreaLineCustomDash'];
  cap: string | null | undefined;
  join: string | null | undefined;
} {
  const pointStyle = point?.chartexStyle;
  const seriesStyle = series.chartexStyle;
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const linkedStyle = chartDataPointStyleRole(chart, 'dataPoint', seriesIndex);
  const rawLinked = rawLinkedChartStyleRole(chart, 'dataPoint');
  const linkedIndex = chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex : seriesStyleIndex;
  const linkedGeometry = linkedStyle;
  const dashLayers = [pointStyle, seriesStyle, linkedGeometry];
  let dash: string | null | undefined = point?.lineDash;
  let customDash: ChartModel['plotAreaLineCustomDash'];
  if (dash == null) {
    for (const layer of dashLayers) {
      if (layer?.lineDash != null || layer?.lineCustomDash != null
        || layer?.lineDashAuthored === true) {
        dash = layer.lineDash;
        customDash = layer.lineCustomDash ?? undefined;
        break;
      }
    }
  }
  const geometry = {
    widthEmu: point?.lineWidthEmu
      ?? pointStyle?.lineWidthEmu
      ?? series.lineWidthEmu
      ?? seriesStyle?.lineWidthEmu
      ?? linkedGeometry?.lineWidthEmu
      ?? point?.markerLineWidthEmu
      ?? series.markerLineWidthEmu,
    dash,
    customDash,
    cap: pointStyle?.lineCap ?? seriesStyle?.lineCap ?? linkedGeometry?.lineCap,
    join: pointStyle?.lineJoin ?? seriesStyle?.lineJoin ?? linkedGeometry?.lineJoin,
  };
  const pointPaint = chartStyleDirectLineDecision(pointStyle, rawLinked, pointIndex);
  if (pointPaint !== undefined) {
    return {
      color: pointPaint?.fillType === 'solid' ? pointPaint.color : point?.lineColor ?? null,
      paint: pointPaint,
      ...geometry,
    };
  }
  if (point?.lineHidden === true) {
    const noLine = chartStyleDirectNoLineDecision(rawLinked);
    if (noLine !== undefined) return { color: null, paint: noLine, ...geometry };
  }
  if (point?.lineColor != null) {
    return { color: point.lineColor, paint: undefined, ...geometry };
  }

  const seriesPaint = chartStyleDirectLineDecision(
    seriesStyle, rawLinked, seriesStyleIndex,
  );
  if (seriesPaint !== undefined) {
    return {
      color: seriesPaint?.fillType === 'solid' ? seriesPaint.color : series.lineColor ?? null,
      paint: seriesPaint,
      ...geometry,
    };
  }
  if (series.lineHidden === true) {
    const noLine = chartStyleDirectNoLineDecision(rawLinked);
    if (noLine !== undefined) return { color: null, paint: noLine, ...geometry };
  }
  if (series.lineColor != null) {
    return { color: series.lineColor, paint: undefined, ...geometry };
  }
  const linkedPoint = chartExStyleLinePaintDecision(
    chart, linkedStyle, linkedIndex, series.values.length,
  );
  if (linkedPoint !== undefined) {
    return {
      color: linkedPoint?.fillType === 'solid' ? linkedPoint.color : null,
      paint: linkedPoint,
      ...geometry,
    };
  }
  const bubbleSize = series.bubbleSizes?.[pointIndex];
  const automaticNegativeThreeDLine = bubbleSize != null
    && Number.isFinite(bubbleSize)
    && bubbleSize < 0
    && bubblePointIsThreeD(series, point)
    ? '000000'
    : null;
  return {
    // Current Excel gives its generated white negative 3-D material a black
    // outline. Direct or linked no-line returned above remains authoritative.
    color: point?.markerLine
      ?? series.markerLine
      ?? series.lineColor
      ?? automaticNegativeThreeDLine,
    paint: undefined,
    ...geometry,
  };
}
