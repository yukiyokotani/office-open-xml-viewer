import type { ChartModel, ChartRect } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { resolveFill } from '../shape/paint.js';
import { axisLineWidthPx } from './axis-style.js';
import { strokeChartFrameRect } from './compound-frame.js';
import { paintChartStyleEffects } from './style-effects.js';
import { rawLinkedChartStyleRole } from './effective-style.js';
import {
  chartStyleDirectNoFillDecision,
  chartStyleFillCascade,
} from './style-paint.js';
import { paintActiveChartImageFill } from './image-fill-context.js';

/** Effective legend-frame fill before renderer materialization. Kept beside
 * the frame painter so host image preflight and Canvas selection cannot drift. */
export function effectiveLegendFrameFill(
  chart: ChartModel,
): ChartModel['legendFill'] | undefined {
  const rawLinked = rawLinkedChartStyleRole(chart, 'legend');
  const blockedNoFill = chart.legendFillHidden === true
    && chartStyleDirectNoFillDecision(rawLinked) === undefined;
  if (chart.legendFillHidden === true && !blockedNoFill) return undefined;
  if (chart.legendFill != null) return chart.legendFill;
  if (chart.legendFillColor != null
    || chart.legendFillPaintAuthored === true && !blockedNoFill) return undefined;
  return chartStyleFillCascade(
    chart.chartStyleRoles?.legend, rawLinked, 0, chart.legendStyle,
  ) ?? undefined;
}

/** Paint the authored solid `<c:legend><c:spPr>` frame before legend content.
 * Omitted/noFill properties stay transparent; there is no invented default.
 * Direct DrawingML line properties remain authoritative; linked Chart Style
 * dash/cap/join values fill only omitted properties before this helper runs. */
export function paintLegendFrame(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  bounds: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  if ((!chart.legendFill && !chart.legendFillColor || chart.legendFillHidden === true)
    && (!chart.legendLineFill && !chart.legendLineColor || chart.legendLineHidden === true)) return;
  paintChartStyleEffects(
    ctx,
    chart.legendStyle,
    chart.chartStyleRoles?.legend,
    0,
    bounds,
    ptToPx,
    target => {
      target.save();
      if (chart.legendFillHidden !== true && (chart.legendFill || chart.legendFillColor)) {
        const imagePainted = chart.legendFill?.fillType === 'image'
          ? paintActiveChartImageFill(
              target, chart.legendFill,
              bounds.x, bounds.y, bounds.w, bounds.h,
              ptToPx, shapeRotationDeg,
            )
          : false;
        const fill = imagePainted ? null
          : chart.legendFill
            ? resolveFill(
                chart.legendFill, target,
                bounds.x, bounds.y, bounds.w, bounds.h,
                shapeRotationDeg,
              )
            : chart.legendFillColor ? `#${chart.legendFillColor}` : null;
        if (fill) target.fillStyle = fill;
        if (fill) {
          target.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
        }
      }
      if (chart.legendLineHidden !== true
        && (chart.legendLineFill || chart.legendLineColor) && bounds.w > 0 && bounds.h > 0) {
        const width = axisLineWidthPx(chart.legendLineWidthEmu, ptToPx);
        const stroke = chart.legendLineFill
          ? resolveFill(
              chart.legendLineFill, target,
              bounds.x, bounds.y, bounds.w, bounds.h,
              shapeRotationDeg,
            )
          : chart.legendLineColor ? `#${chart.legendLineColor}` : null;
        if (stroke) {
          target.strokeStyle = stroke;
          target.lineCap = chart.legendLineCap === 'rnd'
            ? 'round' : chart.legendLineCap === 'sq' ? 'square' : 'butt';
          target.lineJoin = chart.legendLineJoin === 'round' || chart.legendLineJoin === 'bevel'
            ? chart.legendLineJoin : 'miter';
          target.setLineDash(drawingmlLineDashArray(
            chart.legendLineCustomDash, chart.legendLineDash, width,
          ));
          strokeChartFrameRect(
            target, bounds.x, bounds.y, bounds.w, bounds.h, width, chart.legendLineCompound,
          );
        }
      }
      target.restore();
    }
  );
}
