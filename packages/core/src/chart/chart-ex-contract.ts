import type { ChartModel, ChartRect } from '../types/chart.js';

/**
 * Synchronous optional renderer for Microsoft ChartEx (`cx:*`) chart
 * families. Classic DrawingML (`c:*`) 2-D charts stay in the default renderer;
 * the newer ChartEx dialect is supplied from `@silurus/ooxml/chart-ex`.
 */
export interface ChartExRenderer {
  render(
    ctx: CanvasRenderingContext2D,
    chart: ChartModel,
    rect: ChartRect,
    ptToPx: number,
    shapeRotationDeg?: number,
  ): boolean;
}
