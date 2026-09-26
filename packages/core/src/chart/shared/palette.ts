// Classic chart palette helpers.
import type { ChartDataPointOverride, ChartModel, ChartRect, ChartSeries } from '../../types/chart';
import { chartDataPointStyleRole, chartSeriesSourceIndex, chartSeriesVariesByPoint } from '../effective-style.js';
import { classicDataPointLineStyle } from '../classic-data-point-style.js';
import { resolveFill } from '../../shape/paint.js';
import { axisLineWidthPx } from '../axis-style.js';
import { drawingmlLineDashArray } from '../../draw/dash.js';
import { chartStyleEffectOwner, paintChartStyleEffects } from '../style-effects.js';


// ─── Palette + helpers ──────────────────────────────────────────────────────

export const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];


/** Office 2013+ ChartEx fallback accents when no theme/colors sidecar resolves. */
export const CHARTEX_DEFAULT_PALETTE = [
  '5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47',
] as const;


export function chartColor(idx: number, series?: { color?: string | null } | null): string {
  if (series?.color) return `#${series.color}`;
  return `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}


/** Index point-scoped OOXML overrides once while preserving first-in-document
 * precedence for duplicate indexes. */
export function indexPointOverrides<T extends { idx: number }>(
  values: readonly T[] | null | undefined,
): ReadonlyMap<number, T> {
  const indexed = new Map<number, T>();
  for (const value of values ?? []) {
    if (!indexed.has(value.idx)) indexed.set(value.idx, value);
  }
  return indexed;
}


export function pieSliceColor(
  idx: number,
  series: ChartSeries,
  varyColors = true,
  seriesIndex = idx,
): string {
  const override = series.dataPointColors?.[idx];
  if (override) return `#${override}`;
  // When varyColors is off (or the parser deliberately suppresses automatic
  // point colours for a series noFill), every unspecified slice inherits the
  // series fill. Falling straight to the built-in palette would revive a
  // noFill series and recolour a single-colour pie point by point.
  if (series.color === '00000000') return '#00000000';
  return varyColors
    ? `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`
    : chartColor(seriesIndex, series);
}


/** Select the numeric/linked style palette domain used by a pie-family mark.
 * Excel replays a point palette in every ring while `varyColors` is effective,
 * but colors each complete series/ring when it is explicitly disabled. */
export function piePointStyleIndex(
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  pointIndex: number,
): number {
  return chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex
    : chartExSeriesFormatIndex(series, seriesIndex);
}


/** Apply one classic series/point outline with component-wise DrawingML
 * precedence. Series-local paint is direct formatting, while the linked or
 * numeric Chart Style role supplies only missing paint/geometry before the
 * family semantic fallback. */
export function applyClassicStyleLine(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  role: 'dataPoint' | 'dataPointLine',
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  styleIndex: number,
  fallbackColor: string,
  fallbackWidthPx: number,
  ptToPx: number,
  bounds: ChartRect,
  shapeRotationDeg: number,
  semanticVisible = true,
  resetSolidDash = true,
): boolean {
  const line = classicDataPointLineStyle(chart, role, series, point, styleIndex);
  let { paint } = line;
  if (paint === undefined && !semanticVisible) return false;
  if (paint === undefined) {
    paint = { fillType: 'solid', color: fallbackColor.replace(/^#/, '') };
  }
  if (paint === null) return false;

  const stroke = paint.fillType === 'solid'
    ? (paint.color.startsWith('#') ? paint.color : `#${paint.color}`)
    : resolveFill(
        paint, ctx, bounds.x, bounds.y, bounds.w, bounds.h, shapeRotationDeg,
      );
  if (!stroke) return false;
  ctx.strokeStyle = stroke;
  ctx.lineWidth = line.widthEmu != null
    ? axisLineWidthPx(line.widthEmu, ptToPx) : fallbackWidthPx;
  const lineDash = drawingmlLineDashArray(
    line.customDash, line.dash, ctx.lineWidth,
  );
  const currentLineDash = typeof ctx.getLineDash === 'function'
    ? ctx.getLineDash() ?? []
    : [];
  if (resetSolidDash || lineDash.length > 0 || currentLineDash.length > 0) {
    ctx.setLineDash(lineDash);
  }
  ctx.lineCap = line.cap === 'rnd' ? 'round' : line.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = line.join === 'round' || line.join === 'bevel' ? line.join : 'miter';
  return true;
}


export interface IndexedLinePoint {
  x: number;
  y: number;
  index: number;
}


/** Paint a point-varying line as independently styled destination segments.
 * Excel applies §21.2.2.227 to a lone line/scatter/radar series by changing
 * both its marker and the segment that arrives at that point. A multi-series
 * group remains series-coloured. Smooth curves keep the same Catmull-Rom
 * geometry; only the segment paint/effect owner changes at each endpoint. */
export function paintClassicVaryingLineSegments(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  runs: IndexedLinePoint[][],
  smooth: boolean,
  closed: boolean,
  fallbackColor: string,
  fallbackWidthPx: number,
  ptToPx: number,
  bounds: ChartRect,
  shapeRotationDeg: number,
  semanticVisible = true,
): void {
  const seriesIndex = Math.max(0, chartSeriesSourceIndex(chart, series));
  const pointOverrides = indexPointOverrides(series.dataPointOverrides);
  const fallbackRole = chartDataPointStyleRole(chart, 'dataPointLine', seriesIndex);
  const paintSegment = (
    from: IndexedLinePoint,
    to: IndexedLinePoint,
    run: IndexedLinePoint[],
    segmentIndex: number,
    closesRun = false,
  ): void => {
    const point = pointOverrides.get(to.index);
    const paint = (target: CanvasRenderingContext2D): void => {
      target.save();
      if (applyClassicStyleLine(
        target, chart, 'dataPointLine', series, point, to.index,
        fallbackColor, fallbackWidthPx, ptToPx, bounds, shapeRotationDeg,
        semanticVisible,
      )) {
        target.beginPath();
        target.moveTo(from.x, from.y);
        if (smooth && !closesRun) {
          const p0 = run[segmentIndex - 1] ?? from;
          const p3 = run[segmentIndex + 2] ?? to;
          target.bezierCurveTo(
            from.x + (to.x - p0.x) / 6,
            from.y + (to.y - p0.y) / 6,
            to.x - (p3.x - from.x) / 6,
            to.y - (p3.y - from.y) / 6,
            to.x,
            to.y,
          );
        } else {
          target.lineTo(to.x, to.y);
        }
        target.stroke();
      }
      target.restore();
    };
    paintChartStyleEffects(
      ctx,
      chartStyleEffectOwner(point?.chartexStyle, series.chartexStyle),
      fallbackRole,
      to.index,
      bounds,
      ptToPx,
      paint,
      to.index,
    );
  };

  for (const run of runs) {
    for (let index = 0; index + 1 < run.length; index++) {
      paintSegment(run[index], run[index + 1], run, index);
    }
    if (closed && run.length > 1) {
      paintSegment(run[run.length - 1], run[0], run, run.length - 1, true);
    }
  }
}


/** Effective CT_Series formatting index. The shared parser preserves authored
 * `formatIdx` and resolves omission to the original document-order index so a
 * hidden series cannot renumber the visible series' linked Chart Style. */
export function chartExSeriesFormatIndex(
  series: Pick<ChartSeries, 'chartexFormatIdx'> | null | undefined,
  fallbackIndex: number,
): number {
  return series?.chartexFormatIdx ?? fallbackIndex;
}
