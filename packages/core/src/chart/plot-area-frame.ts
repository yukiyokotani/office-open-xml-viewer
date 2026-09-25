import type { ChartModel } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { resolveFill } from '../shape/paint.js';
import { EMU_PER_PT } from '../units.js';
import { strokeChartFrameRect } from './compound-frame.js';
import { paintChartImageFill } from './image-fill.js';
import { paintChartStyleEffects } from './style-effects.js';

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
  paintChartStyleEffects(
    ctx,
    chart.plotAreaStyle,
    chart.threeD ? chart.chartStyleRoles?.plotArea3D : chart.chartStyleRoles?.plotArea,
    0,
    { x, y, w, h },
    ptToPx,
    target => {
      if (chart.plotAreaFillHidden !== true) {
        if (chart.plotAreaFill?.fillType === 'image') {
          paintChartImageFill(
            target, chart.plotAreaFill, x, y, w, h, ptToPx, shapeRotationDeg,
          );
        } else {
          const fill = chart.plotAreaFill
            ? resolveFill(chart.plotAreaFill, target, x, y, w, h, shapeRotationDeg)
            : chart.plotAreaBg ? `#${chart.plotAreaBg}` : null;
          if (fill) {
            target.fillStyle = fill;
            target.fillRect(x, y, w, h);
          }
        }
      }
      if (chart.plotAreaLineHidden === true
        || (!chart.plotAreaLineFill && !chart.plotAreaLineColor)) return;

      const lineWidth = chart.plotAreaLineWidthEmu
        ? Math.max(0.5, chart.plotAreaLineWidthEmu / EMU_PER_PT) * ptToPx
        : 1;
      target.save();
      const stroke = chart.plotAreaLineFill
        ? resolveFill(chart.plotAreaLineFill, target, x, y, w, h, shapeRotationDeg)
        : chart.plotAreaLineColor ? `#${chart.plotAreaLineColor}` : null;
      if (stroke) {
        target.strokeStyle = stroke;
        target.setLineDash(drawingmlLineDashArray(
          chart.plotAreaLineCustomDash,
          chart.plotAreaLineDash,
          lineWidth,
        ));
        target.lineCap = chart.plotAreaLineCap === 'rnd'
          ? 'round' : chart.plotAreaLineCap === 'sq' ? 'square' : 'butt';
        target.lineJoin = chart.plotAreaLineJoin === 'round' || chart.plotAreaLineJoin === 'bevel'
          ? chart.plotAreaLineJoin : 'miter';
        strokeChartFrameRect(
          target, x, y, w, h, lineWidth, chart.plotAreaLineCompound,
        );
      }
      target.restore();
    },
  );
}
