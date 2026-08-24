import type { ChartModel } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { resolveFill } from '../shape/paint.js';
import { EMU_PER_PT } from '../units.js';
import { strokeChartFrameRect } from './compound-frame.js';
import { paintChartImageFill } from './image-fill.js';

/** Paint the effective DrawingML plot-area frame behind chart geometry.
 *
 * Linked Chart Style values have already been merged into `ChartModel`, so
 * every 2-D and optional 3-D family consumes the same direct-over-linked
 * precedence result here. */
export function paintPlotAreaFrame(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  if (chart.plotAreaFillHidden !== true) {
    if (chart.plotAreaFill?.fillType === 'image') {
      paintChartImageFill(
        ctx, chart.plotAreaFill, x, y, w, h, ptToPx, shapeRotationDeg,
      );
    } else {
      const fill = chart.plotAreaFill
        ? resolveFill(chart.plotAreaFill, ctx, x, y, w, h, shapeRotationDeg)
        : chart.plotAreaBg ? `#${chart.plotAreaBg}` : null;
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fillRect(x, y, w, h);
      }
    }
  }
  if (chart.plotAreaLineHidden === true
    || (!chart.plotAreaLineFill && !chart.plotAreaLineColor)) return;

  const lineWidth = chart.plotAreaLineWidthEmu
    ? Math.max(0.5, chart.plotAreaLineWidthEmu / EMU_PER_PT) * ptToPx
    : 1;
  ctx.save();
  const stroke = chart.plotAreaLineFill
    ? resolveFill(chart.plotAreaLineFill, ctx, x, y, w, h, shapeRotationDeg)
    : chart.plotAreaLineColor ? `#${chart.plotAreaLineColor}` : null;
  if (!stroke) {
    ctx.restore();
    return;
  }
  ctx.strokeStyle = stroke;
  ctx.setLineDash(drawingmlLineDashArray(
    chart.plotAreaLineCustomDash,
    chart.plotAreaLineDash,
    lineWidth,
  ));
  ctx.lineCap = chart.plotAreaLineCap === 'rnd'
    ? 'round' : chart.plotAreaLineCap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = chart.plotAreaLineJoin === 'round' || chart.plotAreaLineJoin === 'bevel'
    ? chart.plotAreaLineJoin : 'miter';
  strokeChartFrameRect(
    ctx, x, y, w, h, lineWidth, chart.plotAreaLineCompound,
  );
  ctx.restore();
}
