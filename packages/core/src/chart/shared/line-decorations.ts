// Classic chart line decorations helpers.
import type { ChartModel, ChartSeries, ChartStockUpDownBarStyle } from '../../types/chart';
import { axisLineWidthPx } from '../axis-style.js';
import { paintChartStyleEffects } from '../style-effects.js';
import { paintClassicDataPointRect } from './chartex-style.js';
import { clamp, dashPatternForPreset } from './geometry.js';
import { applyDecorationLineStyle, chartStyleRoleBarPaint, chartStyleRoleLine, drawDropLineEnvelopes } from './style-roles.js';


export function drawUpDownBars(
  ctx: CanvasRenderingContext2D,
  startValueAt: (index: number) => number | null,
  endValueAt: (index: number) => number | null,
  pointCount: number,
  toX: (index: number) => number,
  toYStart: (value: number) => number,
  toYEnd: (value: number) => number,
  slotWidth: number,
  style: ChartStockUpDownBarStyle,
  ptToPx: number,
  automaticPaint?: {
    lineColor: string;
    lineWidthEmu: number;
    upFillColor: string;
    downFillColor: string;
  },
  shapeRotationDeg = 0,
): void {
  const gapPercent = Number.isFinite(style.gapWidthPercent) && style.gapWidthPercent >= 0
    ? style.gapWidthPercent
    : 150;
  const barWidth = Math.max(0, slotWidth / (1 + gapPercent / 100));
  for (let index = 0; index < pointCount; index++) {
    const start = startValueAt(index);
    const end = endValueAt(index);
    if (start == null || end == null || !Number.isFinite(start) || !Number.isFinite(end)) continue;
    const startY = toYStart(start);
    const endY = toYEnd(end);
    const barHeight = Math.abs(endY - startY);
    if (!(barWidth > 0) || !(barHeight > 0) || !Number.isFinite(barHeight)) continue;
    const paint = end >= start ? style.up : style.down;
    const fillOwned = paint.fillPaintAuthored === true
      || paint.fill != null || paint.fillColor != null || paint.fillHidden === true;
    const automaticFill = fillOwned
      ? undefined
      : end >= start ? automaticPaint?.upFillColor : automaticPaint?.downFillColor;
    const fillColor = paint.fillColor ?? automaticFill;
    const barX = toX(index) - barWidth / 2;
    const barY = Math.min(startY, endY);
    const lineOwned = paint.linePaintAuthored === true
      || paint.lineColor != null || paint.lineHidden === true;
    const lineColor = paint.lineColor ?? (lineOwned ? undefined : automaticPaint?.lineColor);
    const lineWidthEmu = paint.lineWidthEmu
      ?? (lineOwned ? undefined : automaticPaint?.lineWidthEmu);
    const paintBar = (target: CanvasRenderingContext2D): void => {
      if (!paint.fillHidden && (paint.fill != null || fillColor != null)) {
        paintClassicDataPointRect(
          target,
          paint.fill ?? (fillColor ? { fillType: 'solid', color: fillColor } : null),
          { x: barX, y: barY, w: barWidth, h: barHeight },
          fillColor ? `#${fillColor}` : 'rgba(0,0,0,0)',
          ptToPx,
          shapeRotationDeg,
        );
      }
      if (!paint.lineHidden
        && (paint.linePaintAuthored !== true || lineColor != null) && (
        lineColor != null || lineWidthEmu != null
      )) {
        const previousDash = target.getLineDash();
        const previousCap = target.lineCap;
        const previousJoin = target.lineJoin;
        target.strokeStyle = `#${lineColor ?? '000000'}`;
        target.lineWidth = lineWidthEmu != null
          ? axisLineWidthPx(lineWidthEmu, ptToPx)
          : Math.max(1, 0.75 * ptToPx);
        target.setLineDash(dashPatternForPreset(paint.lineDash ?? undefined, target.lineWidth));
        target.lineCap = paint.lineCap === 'rnd'
          ? 'round' : paint.lineCap === 'sq' ? 'square' : 'butt';
        target.lineJoin = paint.lineJoin === 'round' || paint.lineJoin === 'bevel'
          ? paint.lineJoin : 'miter';
        target.strokeRect(barX, barY, barWidth, barHeight);
        target.setLineDash(previousDash);
        target.lineCap = previousCap;
        target.lineJoin = previousJoin;
      }
    };
    paintChartStyleEffects(
      ctx,
      paint.style,
      undefined,
      index,
      { x: barX, y: barY, w: barWidth, h: barHeight },
      ptToPx,
      paintBar,
    );
  }
}


export function drawLineGroupDecorations(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  pointCount: number,
  toX: (index: number) => number,
  yMapFor: (series: ChartSeries) => (value: number) => number,
  categoryAxisYFor: (series: ChartSeries) => number,
  valueFor: (series: ChartSeries, index: number) => number | null,
  slotWidth: number,
  ptToPx: number,
  shapeRotationDeg: number,
  phase: 'background' | 'foreground',
): void {
  for (const decoration of chart.lineGroupDecorations ?? []) {
    let members = chart.series.filter(series => series.lineGroupIndex === decoration.groupIndex);
    // Hand-authored callers predating line-group provenance still represent a
    // single ordinary line group as the whole series list.
    if (members.length === 0 && decoration.groupIndex === 0
      && ['line', 'stackedLine', 'stackedLinePct'].includes(chart.chartType)) {
      members = chart.series.filter(series => series.seriesType == null || series.seriesType === 'line');
    }
    if (members.length === 0) continue;

    if (phase === 'foreground' && decoration.upDownBars && members.length >= 2) {
      const first = members[0];
      const last = members[members.length - 1];
      // Empty upBars/downBars paint is application-defined. The retained
      // Office observation is limited to classic Style 2; it is a family
      // default above the numeric role, while direct and linked paint remain
      // authoritative.
      const automaticPaint = chart.legacyChartStyle === 2 ? {
        lineColor: '000000', lineWidthEmu: 9525,
        upFillColor: 'FFFFFF', downFillColor: '000000',
      } : undefined;
      const upDownBars = {
        ...decoration.upDownBars,
        up: chartStyleRoleBarPaint(chart, decoration.upDownBars.up, 'upBar', automaticPaint),
        down: chartStyleRoleBarPaint(chart, decoration.upDownBars.down, 'downBar', automaticPaint),
      };
      drawUpDownBars(
        ctx, index => valueFor(first, index), index => valueFor(last, index), pointCount, toX,
        yMapFor(first), yMapFor(last), slotWidth, upDownBars, ptToPx,
        undefined, shapeRotationDeg,
      );
    }

    if (phase === 'foreground') continue;

    const dropLineStyle = decoration.dropLines
      ? chartStyleRoleLine(chart, decoration.dropLines, 'dropLine')
      : null;
    if (dropLineStyle && applyDecorationLineStyle(ctx, dropLineStyle, ptToPx)) {
      // Office paints one envelope per category, not one line per series. The
      // envelope includes the effective category-axis crossing and every
      // plotted point in the owning line group. This is observable in vector
      // output for both ordinary and interior crossings; painting per-series
      // segments produces coincident seams and the wrong visible endpoints.
      drawDropLineEnvelopes(
        ctx, members, pointCount, toX, yMapFor, categoryAxisYFor, valueFor,
      );
    }

    const hiLowLineStyle = decoration.hiLowLines
      ? chartStyleRoleLine(chart, decoration.hiLowLines, 'hiLoLine')
      : null;
    if (hiLowLineStyle && members.length >= 2
      && applyDecorationLineStyle(ctx, hiLowLineStyle, ptToPx)) {
      const toY = yMapFor(members[0]);
      for (let index = 0; index < pointCount; index++) {
        let low = Infinity;
        let high = -Infinity;
        for (const series of members) {
          const value = valueFor(series, index);
          if (value == null || !Number.isFinite(value)) continue;
          low = Math.min(low, value);
          high = Math.max(high, value);
        }
        if (!Number.isFinite(low) || !Number.isFinite(high)) continue;
        ctx.beginPath();
        ctx.moveTo(toX(index), toY(low));
        ctx.lineTo(toX(index), toY(high));
        ctx.stroke();
      }
    }
  }
}


export function axisCrossingValue(
  crossesAt: number | null | undefined,
  crosses: string | null | undefined,
  min: number,
  max: number,
): number {
  if (crossesAt != null && Number.isFinite(crossesAt)) {
    return clamp(crossesAt, min, max);
  }
  if (crosses === 'max') return max;
  if (crosses === 'min') return min;
  return clamp(0, min, max);
}


export function categoryAxisCrossingValue(chart: ChartModel, min: number, max: number): number {
  return axisCrossingValue(chart.catAxisCrossesAt, chart.catAxisCrosses, min, max);
}
