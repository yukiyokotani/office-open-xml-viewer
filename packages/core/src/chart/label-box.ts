import type { ChartExElementStyle, ChartLabelBox, ChartRect } from '../types/chart.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import { fillCanProduceVisiblePixels, resolveFill } from '../shape/paint.js';
import { EMU_PER_PT } from '../units.js';
import { chartStyleEffectOwner, paintChartStyleEffects } from './style-effects.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectNoFillDecision,
  chartStyleFillDecision,
} from './style-paint.js';
import { paintActiveChartImageFill } from './image-fill-context.js';

/** Resolve only the fill component of a label shape. Hosts use this same pure
 * decision before rendering to warm picture fills; the painter's full label
 * merge adds outline geometry and effects separately. */
export function effectiveChartLabelBoxFill(
  direct: ChartLabelBox | null | undefined,
  linked: ChartExElementStyle | null | undefined,
  rawLinked: ChartExElementStyle | null | undefined,
  createFromLinked: boolean,
  directIndex = 0,
  linkedIndex = directIndex,
): Pick<ChartLabelBox, 'fill' | 'fillPaint' | 'fillHidden' | 'fillPaintAuthored'> | undefined {
  if (!linked || (!direct && !createFromLinked)) return direct ?? undefined;
  const source = direct ?? {};
  const linkedFill = linked.fillNoStyle !== true && (linked.fillPaintAuthored === true
    || linked.fillHidden === true || linked.fillPaints != null || linked.fillColors != null);
  const directNoFill = source.fillHidden === true
    ? chartStyleDirectNoFillDecision(rawLinked) : undefined;
  const directStyleFill = chartStyleDirectFillDecision(source.style, rawLinked, directIndex);
  const directFill = source.fillPaint != null || source.fill != null
    || directStyleFill !== undefined
    || source.fillHidden === true && directNoFill !== undefined
    || source.fillPaintAuthored === true && source.fillHidden !== true;
  const fillDecision = directNoFill !== undefined ? directNoFill
    : source.fillPaint ?? (source.fill ? { fillType: 'solid' as const, color: source.fill }
      : directStyleFill !== undefined
        ? directStyleFill
        : source.fillPaintAuthored === true && source.fillHidden !== true
          ? null
          : chartStyleFillDecision(linked, linkedIndex));
  return {
    fill: fillDecision?.fillType === 'solid' ? fillDecision.color : undefined,
    fillPaint: fillDecision != null && fillDecision.fillType !== 'solid'
      ? fillDecision : undefined,
    fillHidden: fillDecision === null ? true : undefined,
    fillPaintAuthored: directFill ? true : linkedFill ? true : source.fillPaintAuthored,
  };
}

/** Whether a label shape supplies an actual Canvas-visible box paint. A bare
 * `<c:spPr>` or explicit `noFill`/no-line still carries authored provenance,
 * but it must not turn an ordinary pie label into a boxed callout. */
export function chartLabelBoxHasVisiblePaint(
  box: ChartLabelBox | null | undefined,
): boolean {
  if (!box) return false;
  const hasFill = box.fillHidden !== true && (box.fillPaint != null
    ? fillCanProduceVisiblePixels(box.fillPaint)
    : box.fill != null && fillCanProduceVisiblePixels({ fillType: 'solid', color: box.fill }));
  const hasBorder = box.borderHidden !== true && (box.borderFill != null
    ? fillCanProduceVisiblePixels(box.borderFill)
    : box.borderColor != null
      && fillCanProduceVisiblePixels({ fillType: 'solid', color: box.borderColor }));
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
    // DrawingML owns effects as one atomic component. A point shape that only
    // authors fill/line keeps the series/linked effect, while an authored empty
    // or unsupported point effect deliberately suppresses that fallback.
    style: chartStyleEffectOwner(higher.style, lower.style),
    effectFallbackStyle: higher.effectFallbackStyle ?? lower.effectFallbackStyle,
    effectStyleIndex: higher.effectStyleIndex ?? lower.effectStyleIndex,
    effectFallbackIndex: higher.effectFallbackIndex ?? lower.effectFallbackIndex,
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
  const paint = (target: CanvasRenderingContext2D): void => {
    if (box.fillHidden !== true) {
      const imagePainted = box.fillPaint?.fillType === 'image'
        ? paintActiveChartImageFill(
            target, box.fillPaint, rect.x, rect.y, rect.w, rect.h,
            ptToPx, shapeRotationDeg,
          )
        : false;
      const fill = imagePainted ? null
        : box.fillPaint
          ? resolveFill(box.fillPaint, target, rect.x, rect.y, rect.w, rect.h, shapeRotationDeg)
          : box.fill ? `#${box.fill}` : null;
      if (fill) {
        target.fillStyle = fill;
        target.fillRect(rect.x, rect.y, rect.w, rect.h);
      }
    }
    if (box.borderHidden === true) return;
    const stroke = box.borderFill
      ? resolveFill(box.borderFill, target, rect.x, rect.y, rect.w, rect.h, shapeRotationDeg)
      : box.borderColor ? `#${box.borderColor}` : null;
    if (!stroke) return;
    target.save();
    target.strokeStyle = stroke;
    target.lineWidth = box.borderWidthEmu != null
      ? Math.max(0.25, box.borderWidthEmu / EMU_PER_PT * ptToPx)
      : Math.max(0.25, 0.75 * ptToPx);
    target.setLineDash(drawingmlLineDashArray(
      box.borderCustomDash,
      box.borderDash,
      target.lineWidth,
    ));
    if (box.borderCap === 'rnd') target.lineCap = 'round';
    else if (box.borderCap === 'sq') target.lineCap = 'square';
    else if (box.borderCap === 'flat') target.lineCap = 'butt';
    if (box.borderJoin === 'round') target.lineJoin = 'round';
    else if (box.borderJoin === 'bevel') target.lineJoin = 'bevel';
    else if (box.borderJoin === 'miter') target.lineJoin = 'miter';
    // Compound rails remain parsed-only until Office rail geometry is established;
    // painting a single authored stroke is preferable to inventing rail ratios.
    target.strokeRect(rect.x, rect.y, rect.w, rect.h);
    target.restore();
  };
  paintChartStyleEffects(
    ctx,
    box.style,
    box.effectFallbackStyle,
    box.effectStyleIndex ?? 0,
    rect,
    ptToPx,
    paint,
    box.effectFallbackIndex ?? box.effectStyleIndex ?? 0,
  );
}
