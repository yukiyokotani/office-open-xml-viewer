import type { ChartLabelBox, ChartRect } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { resolveFill } from '../shape/paint.js';
import { EMU_PER_PT } from '../units.js';

/** Whether a label shape supplies an actual Canvas-visible box paint. A bare
 * `<c:spPr>` or explicit `noFill`/no-line still carries authored provenance,
 * but it must not turn an ordinary pie label into a boxed callout. */
export function chartLabelBoxHasVisiblePaint(
  box: ChartLabelBox | null | undefined,
): boolean {
  if (!box) return false;
  const hasFill = box.fillHidden !== true && (box.fill != null || box.fillPaint != null);
  const hasBorder = box.borderHidden !== true
    && (box.borderColor != null || box.borderFill != null);
  return hasFill || hasBorder;
}

/** Merge two directly-authored label shapes property-by-property. The higher
 * precedence shape owns an authored paint/noFill choice even when that choice
 * cannot be resolved to a Canvas paint; omitted geometry continues to inherit
 * from the lower-precedence series/linked shape. */
export function mergeChartLabelBoxes(
  higher: ChartLabelBox | null | undefined,
  lower: ChartLabelBox | null | undefined,
): ChartLabelBox | undefined {
  if (!higher) return lower ?? undefined;
  if (!lower) return higher;
  const higherFillAuthored = higher.fillPaintAuthored === true
    || higher.fill != null || higher.fillPaint != null || higher.fillHidden === true;
  const higherBorderPaintAuthored = higher.borderPaintAuthored === true
    || higher.borderColor != null || higher.borderFill != null || higher.borderHidden === true;
  const higherDashAuthored = higher.borderDashAuthored === true
    || higher.borderDash != null || higher.borderCustomDash != null;
  return {
    ...lower,
    ...higher,
    fill: higherFillAuthored ? higher.fill : lower.fill,
    fillPaint: higherFillAuthored ? higher.fillPaint : lower.fillPaint,
    fillHidden: higherFillAuthored ? higher.fillHidden : lower.fillHidden,
    fillPaintAuthored: higherFillAuthored
      ? higher.fillPaintAuthored : lower.fillPaintAuthored,
    borderColor: higherBorderPaintAuthored ? higher.borderColor : lower.borderColor,
    borderFill: higherBorderPaintAuthored ? higher.borderFill : lower.borderFill,
    borderHidden: higherBorderPaintAuthored ? higher.borderHidden : lower.borderHidden,
    borderPaintAuthored: higherBorderPaintAuthored
      ? higher.borderPaintAuthored : lower.borderPaintAuthored,
    borderWidthEmu: higher.borderWidthEmu ?? lower.borderWidthEmu,
    borderDash: higherDashAuthored ? higher.borderDash : lower.borderDash,
    borderCustomDash: higherDashAuthored
      ? higher.borderCustomDash : lower.borderCustomDash,
    borderDashAuthored: higherDashAuthored
      ? higher.borderDashAuthored : lower.borderDashAuthored,
    borderCap: higher.borderCap ?? lower.borderCap,
    borderJoin: higher.borderJoin ?? lower.borderJoin,
    borderCompound: higher.borderCompound ?? lower.borderCompound,
  };
}

/** Paint a data/trendline label shape from one effective DrawingML recipe. */
export function paintChartLabelBox(
  ctx: CanvasRenderingContext2D,
  box: ChartLabelBox | null | undefined,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  if (!box) return;
  if (box.fillHidden !== true) {
    const fill = box.fillPaint
      ? resolveFill(box.fillPaint, ctx, rect.x, rect.y, rect.w, rect.h, shapeRotationDeg)
      : box.fill ? `#${box.fill}` : null;
    if (fill) {
      ctx.fillStyle = fill;
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    }
  }
  if (box.borderHidden === true) return;
  const stroke = box.borderFill
    ? resolveFill(box.borderFill, ctx, rect.x, rect.y, rect.w, rect.h, shapeRotationDeg)
    : box.borderColor ? `#${box.borderColor}` : null;
  if (!stroke) return;
  ctx.save();
  ctx.strokeStyle = stroke;
  ctx.lineWidth = box.borderWidthEmu != null
    ? Math.max(0.25, box.borderWidthEmu / EMU_PER_PT * ptToPx)
    : Math.max(0.25, 0.75 * ptToPx);
  ctx.setLineDash(drawingmlLineDashArray(
    box.borderCustomDash,
    box.borderDash,
    ctx.lineWidth,
  ));
  if (box.borderCap === 'rnd') ctx.lineCap = 'round';
  else if (box.borderCap === 'sq') ctx.lineCap = 'square';
  else if (box.borderCap === 'flat') ctx.lineCap = 'butt';
  if (box.borderJoin === 'round') ctx.lineJoin = 'round';
  else if (box.borderJoin === 'bevel') ctx.lineJoin = 'bevel';
  else if (box.borderJoin === 'miter') ctx.lineJoin = 'miter';
  // Compound rails remain parsed-only until Office rail geometry is established;
  // painting a single authored stroke is preferable to inventing rail ratios.
  ctx.strokeRect(rect.x, rect.y, rect.w, rect.h);
  ctx.restore();
}
