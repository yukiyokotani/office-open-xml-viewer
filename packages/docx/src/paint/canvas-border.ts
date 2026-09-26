import {
  crispOffset,
  docxBorderDashArray,
  doubleRailGeometry,
} from '@silurus/ooxml-core';
import type { BorderSegment, TextDecorationLayout } from '../layout/types.js';
import {
  inverseMapAffinePoint,
  inverseMapAffineVector,
  mapAffinePoint,
  scaleAffine,
} from './affine.js';
import type { CanvasPaintContext } from './types.js';

type CompoundBand = Readonly<{ offsetDev: number; widthDev: number }>;

/** ECMA-376 §17.18.2 names the ordered thin/thick rails and gap class. The
 * token supplies relative band weights; device-pixel flooring keeps every
 * authored rail and gap visible at small zoom levels. */
function compoundBorderBands(
  authoredStyle: string,
  widthDev: number,
  dpr: number,
): Readonly<{ bands: readonly CompoundBand[]; spanDev: number }> | null {
  const triple = authoredStyle === 'triple';
  const match = /^(thinThick|thickThin|thinThickThin)(Small|Medium|Large)Gap$/.exec(authoredStyle);
  if (!triple && !match) return null;
  const railWeights = triple ? [1, 1, 1]
    : match?.[1] === 'thinThick' ? [1, 2]
      : match?.[1] === 'thickThin' ? [2, 1]
        : [1, 2, 1];
  const gapWeight = triple || match?.[2] === 'Small' ? 1
    : match?.[2] === 'Medium' ? 2 : 3;
  const totalWeight = railWeights.reduce((sum, weight) => sum + weight, 0)
    + gapWeight * (railWeights.length - 1);
  const unitDev = widthDev / totalWeight;
  const railsDev = railWeights.map((weight) => Math.max(1, Math.round(unitDev * weight * dpr)));
  const gapDev = Math.max(1, Math.round(unitDev * gapWeight * dpr));
  let cursorDev = 0;
  const bands = railsDev.map((railDev, index) => {
    const band = { offsetDev: cursorDev, widthDev: railDev };
    cursorDev += railDev + (index < railsDev.length - 1 ? gapDev : 0);
    return band;
  });
  return { bands, spanDev: cursorDev };
}

function compoundFrameBandsOutsideToInside(
  compound: Readonly<{ bands: readonly CompoundBand[]; spanDev: number }>,
): readonly CompoundBand[] {
  // ST_Border names compound rails from the cell interior toward the exterior
  // (thinThick = thin inside, thick outside). Closed table frames start at the
  // exterior edge, so reverse both order and offsets while retaining gaps.
  return [...compound.bands].reverse().map((band) => ({
    offsetDev: compound.spanDev - band.offsetDev - band.widthDev,
    widthDev: band.widthDev,
  }));
}

/** Paint a rectangular compound border as closed rails. A compound ST_Border
 * token orders its bands from the inside of the box toward the outside, so the
 * frame maps them to outside-in offsets and mirrors the bottom/right edges.
 * Painting the
 * four sides of each rail with overlapping corner rectangles preserves that
 * ordering and gives every rail a continuous join at all four corners. */
export function paintCompoundBorderFrame(
  bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
  border: Pick<BorderSegment, 'authoredStyle' | 'color' | 'widthPt' | 'style'>,
  context: CanvasPaintContext,
): boolean {
  if (border.style !== 'compound') return false;
  const pointToCss = context.pointToCss ?? scaleAffine(context.scale);
  // A skewed or rotated retained frame is no longer axis-aligned in final CSS
  // space. Its individual sides still retain the compound treatment below;
  // the closed-frame optimization is deliberately limited to exact rectangles.
  if (pointToCss.b !== 0 || pointToCss.c !== 0 || pointToCss.a <= 0 || pointToCss.d <= 0) {
    return false;
  }
  const topLeft = mapAffinePoint(pointToCss, { xPt: bounds.xPt, yPt: bounds.yPt });
  const bottomRight = mapAffinePoint(pointToCss, {
    xPt: bounds.xPt + bounds.widthPt,
    yPt: bounds.yPt + bounds.heightPt,
  });
  const horizontal = compoundBorderBands(
    border.authoredStyle,
    border.widthPt * pointToCss.d,
    context.dpr,
  );
  const vertical = compoundBorderBands(
    border.authoredStyle,
    border.widthPt * pointToCss.a,
    context.dpr,
  );
  if (!horizontal || !vertical || horizontal.bands.length !== vertical.bands.length) return false;
  const horizontalFrameBands = compoundFrameBandsOutsideToInside(horizontal);
  const verticalFrameBands = compoundFrameBandsOutsideToInside(vertical);

  const fillFinalRect = (x: number, y: number, width: number, height: number): boolean => {
    const corners = [
      { xPt: x, yPt: y }, { xPt: x + width, yPt: y },
      { xPt: x, yPt: y + height }, { xPt: x + width, yPt: y + height },
    ].map((point) => inverseMapAffinePoint(pointToCss, point));
    if (corners.some((point) => point === null)) return false;
    const local = corners.filter((point): point is { xPt: number; yPt: number } => point !== null);
    const xs = local.map((point) => point.xPt);
    const ys = local.map((point) => point.yPt);
    context.ctx.fillRect(
      Math.min(...xs),
      Math.min(...ys),
      Math.max(...xs) - Math.min(...xs),
      Math.max(...ys) - Math.min(...ys),
    );
    return true;
  };

  const leftStartDev = Math.round(topLeft.xPt * context.dpr - vertical.spanDev / 2);
  const rightEndDev = Math.round(bottomRight.xPt * context.dpr + vertical.spanDev / 2);
  const topStartDev = Math.round(topLeft.yPt * context.dpr - horizontal.spanDev / 2);
  const bottomEndDev = Math.round(bottomRight.yPt * context.dpr + horizontal.spanDev / 2);
  context.ctx.fillStyle = border.color;
  for (let index = 0; index < horizontalFrameBands.length; index += 1) {
    const horizontalBand = horizontalFrameBands[index]!;
    const verticalBand = verticalFrameBands[index]!;
    const leftDev = leftStartDev + verticalBand.offsetDev;
    const rightDev = rightEndDev - verticalBand.offsetDev;
    const topDev = topStartDev + horizontalBand.offsetDev;
    const bottomDev = bottomEndDev - horizontalBand.offsetDev;
    const leftCss = leftDev / context.dpr;
    const rightCss = rightDev / context.dpr;
    const topCss = topDev / context.dpr;
    const bottomCss = bottomDev / context.dpr;
    const verticalWidthCss = verticalBand.widthDev / context.dpr;
    const horizontalWidthCss = horizontalBand.widthDev / context.dpr;
    // Horizontal rails span through the vertical rails, and vice versa. The
    // overlap is intentional: it closes both the inner and outer corner pixels.
    if (!fillFinalRect(leftCss, topCss, rightCss - leftCss, horizontalWidthCss)) return false;
    if (!fillFinalRect(
      leftCss,
      bottomCss - horizontalWidthCss,
      rightCss - leftCss,
      horizontalWidthCss,
    )) return false;
    if (!fillFinalRect(leftCss, topCss, verticalWidthCss, bottomCss - topCss)) return false;
    if (!fillFinalRect(
      rightCss - verticalWidthCss,
      topCss,
      verticalWidthCss,
      bottomCss - topCss,
    )) return false;
  }
  context.ctx.setLineDash([]);
  return true;
}

/** Logical CSS width whose raster footprint is exactly one device pixel. */
export function oneDevicePixelCssWidth(
  context: Pick<CanvasPaintContext, 'dpr'>,
): number {
  return 1 / context.dpr;
}

/** Paint one already-resolved point-space rule; layout owns every conflict and path. */
export function paintStrokeSegment(
  retainedSegment: BorderSegment | TextDecorationLayout,
  context: CanvasPaintContext,
  minimumCssWidthPx = 0,
): void {
  const minimumWidthPt = minimumCssWidthPx / context.scale;
  const segment = minimumWidthPt > retainedSegment.widthPt
    ? {
        ...retainedSegment,
        widthPt: minimumWidthPt,
        ...(typeof retainedSegment.authoredStyle === 'string' ? {
          dashPatternPt: Object.freeze(
            docxBorderDashArray(retainedSegment.authoredStyle, minimumWidthPt),
          ),
        } : {}),
      }
    : retainedSegment;
  const { ctx } = context;
  ctx.strokeStyle = segment.color;
  ctx.lineWidth = segment.widthPt;
  ctx.setLineDash('dashPatternPt' in segment && segment.dashPatternPt
    ? [...segment.dashPatternPt]
    : []);
  ctx.beginPath();
  const path = 'path' in segment && segment.path?.length ? segment.path : [segment.from, segment.to];
  const axisAligned = path.length === 2
    && (path[0]!.xPt === path[1]!.xPt || path[0]!.yPt === path[1]!.yPt);
  const horizontal = axisAligned && path[0]!.yPt === path[1]!.yPt;
  const vertical = axisAligned && path[0]!.xPt === path[1]!.xPt;
  const pointToCss = context.pointToCss ?? scaleAffine(context.scale);
  const finalPath = path.map((point) => mapAffinePoint(pointToCss, point));
  const localDx = axisAligned ? path[1]!.xPt - path[0]!.xPt : 0;
  const localDy = axisAligned ? path[1]!.yPt - path[0]!.yPt : 0;
  const finalDx = pointToCss.a * localDx + pointToCss.c * localDy;
  const finalDy = pointToCss.b * localDx + pointToCss.d * localDy;
  const finalHorizontal = axisAligned && finalDy === 0;
  const finalVertical = axisAligned && finalDx === 0;
  const normalScale = horizontal
    ? Math.hypot(pointToCss.c, pointToCss.d)
    : vertical ? Math.hypot(pointToCss.a, pointToCss.b) : 0;
  const compound = segment.style === 'compound' && axisAligned && normalScale > 0
    ? compoundBorderBands(segment.authoredStyle, segment.widthPt * normalScale, context.dpr)
    : null;
  if (compound) {
    ctx.fillStyle = segment.color;
    const fillFinalRect = (x: number, y: number, width: number, height: number): void => {
      const corners = [
        { xPt: x, yPt: y }, { xPt: x + width, yPt: y },
        { xPt: x, yPt: y + height }, { xPt: x + width, yPt: y + height },
      ].map((point) => inverseMapAffinePoint(pointToCss, point));
      if (corners.some((point) => point === null)) return;
      const local = corners.filter((point): point is { xPt: number; yPt: number } => point !== null);
      const xs = local.map((point) => point.xPt);
      const ys = local.map((point) => point.yPt);
      ctx.fillRect(Math.min(...xs), Math.min(...ys), Math.max(...xs) - Math.min(...xs), Math.max(...ys) - Math.min(...ys));
    };
    const startDev = Math.round(
      (finalHorizontal ? finalPath[0]!.yPt : finalPath[0]!.xPt) * context.dpr
        - compound.spanDev / 2,
    );
    for (const band of compound.bands) {
      const bandStartCss = (startDev + band.offsetDev) / context.dpr;
      const bandWidthCss = band.widthDev / context.dpr;
      if (finalHorizontal) {
        const x = Math.min(finalPath[0]!.xPt, finalPath[1]!.xPt);
        fillFinalRect(x, bandStartCss, Math.abs(finalDx), bandWidthCss);
      } else if (finalVertical) {
        const y = Math.min(finalPath[0]!.yPt, finalPath[1]!.yPt);
        fillFinalRect(bandStartCss, y, bandWidthCss, Math.abs(finalDy));
      } else {
        const offsetPt = (band.offsetDev - compound.spanDev / 2) / context.dpr / normalScale;
        const widthPt = band.widthDev / context.dpr / normalScale;
        if (horizontal) {
          ctx.fillRect(Math.min(path[0]!.xPt, path[1]!.xPt), path[0]!.yPt + offsetPt,
            Math.abs(path[1]!.xPt - path[0]!.xPt), widthPt);
        } else {
          ctx.fillRect(path[0]!.xPt + offsetPt, Math.min(path[0]!.yPt, path[1]!.yPt),
            widthPt, Math.abs(path[1]!.yPt - path[0]!.yPt));
        }
      }
    }
    ctx.setLineDash([]);
    return;
  }
  if (segment.style === 'double' && axisAligned && normalScale > 0) {
    ctx.fillStyle = segment.color;
    if (finalHorizontal || finalVertical) {
      const fillFinalRect = (x: number, y: number, width: number, height: number): void => {
        const corners = [
          { xPt: x, yPt: y },
          { xPt: x + width, yPt: y },
          { xPt: x, yPt: y + height },
          { xPt: x + width, yPt: y + height },
        ].map((point) => inverseMapAffinePoint(pointToCss, point));
        if (corners.some((point) => point === null)) return;
        const local = corners.filter((point): point is { xPt: number; yPt: number } => point !== null);
        const xs = local.map((point) => point.xPt);
        const ys = local.map((point) => point.yPt);
        ctx.fillRect(
          Math.min(...xs), Math.min(...ys),
          Math.max(...xs) - Math.min(...xs),
          Math.max(...ys) - Math.min(...ys),
        );
      };
      const { railDev, gapDev, spanDev } = doubleRailGeometry(
        segment.widthPt * normalScale,
        context.dpr,
      );
      const railCss = railDev / context.dpr;
      // Snapping must happen after the full affine transform. Mapping a snapped
      // device-space rectangle back to local coordinates preserves that result
      // without presenting a partial object as a Canvas context.
      if (finalHorizontal) {
        const startDev = Math.round(finalPath[0]!.yPt * context.dpr - spanDev / 2);
        const x = Math.min(finalPath[0]!.xPt, finalPath[1]!.xPt);
        const width = Math.abs(finalPath[1]!.xPt - finalPath[0]!.xPt);
        fillFinalRect(x, startDev / context.dpr, width, railCss);
        fillFinalRect(x, (startDev + railDev + gapDev) / context.dpr, width, railCss);
      } else {
        const startDev = Math.round(finalPath[0]!.xPt * context.dpr - spanDev / 2);
        const y = Math.min(finalPath[0]!.yPt, finalPath[1]!.yPt);
        const height = Math.abs(finalPath[1]!.yPt - finalPath[0]!.yPt);
        fillFinalRect(startDev / context.dpr, y, railCss, height);
        fillFinalRect((startDev + railDev + gapDev) / context.dpr, y, railCss, height);
      }
    } else {
      // A general affine has no device row/column to snap against, but the two
      // authored rails still retain their point-space separation.
      const { railDev, gapDev, spanDev } = doubleRailGeometry(
        segment.widthPt * normalScale,
        context.dpr,
      );
      const railPt = railDev / context.dpr / normalScale;
      const gapPt = gapDev / context.dpr / normalScale;
      const spanPt = spanDev / context.dpr / normalScale;
      if (horizontal) {
        const x = Math.min(path[0]!.xPt, path[1]!.xPt);
        const width = Math.abs(path[1]!.xPt - path[0]!.xPt);
        ctx.fillRect(x, path[0]!.yPt - spanPt / 2, width, railPt);
        ctx.fillRect(x, path[0]!.yPt - spanPt / 2 + railPt + gapPt, width, railPt);
      } else {
        const y = Math.min(path[0]!.yPt, path[1]!.yPt);
        const height = Math.abs(path[1]!.yPt - path[0]!.yPt);
        ctx.fillRect(path[0]!.xPt - spanPt / 2, y, railPt, height);
        ctx.fillRect(path[0]!.xPt - spanPt / 2 + railPt + gapPt, y, railPt, height);
      }
    }
    ctx.setLineDash([]);
    return;
  }
  if (segment.style === 'double' && !axisAligned && path.length === 2) {
    // A diagonal double border (ECMA-376 §17.4.73/§17.4.79 cell diagonals)
    // keeps the same rail/gap geometry as an axis-aligned double rail, offset
    // along the segment normal; there is no device row/column to snap to.
    const dx = path[1]!.xPt - path[0]!.xPt;
    const dy = path[1]!.yPt - path[0]!.yPt;
    const length = Math.hypot(dx, dy);
    const normal = length > 0 ? { xPt: -dy / length, yPt: dx / length } : null;
    const scale = normal
      ? Math.hypot(
        pointToCss.a * normal.xPt + pointToCss.c * normal.yPt,
        pointToCss.b * normal.xPt + pointToCss.d * normal.yPt,
      )
      : 0;
    if (normal && scale > 0) {
      const { railDev, gapDev } = doubleRailGeometry(segment.widthPt * scale, context.dpr);
      const railPt = railDev / context.dpr / scale;
      const offsetPt = (railDev + gapDev) / context.dpr / scale / 2;
      ctx.lineWidth = railPt;
      for (const side of [-1, 1]) {
        ctx.beginPath();
        ctx.moveTo(path[0]!.xPt + normal.xPt * offsetPt * side, path[0]!.yPt + normal.yPt * offsetPt * side);
        ctx.lineTo(path[1]!.xPt + normal.xPt * offsetPt * side, path[1]!.yPt + normal.yPt * offsetPt * side);
        ctx.stroke();
      }
      ctx.setLineDash([]);
      return;
    }
  }
  const cssOffset = finalVertical && normalScale > 0
    ? { xPt: crispOffset(finalPath[0]!.xPt, segment.widthPt * normalScale, context.dpr), yPt: 0 }
    : finalHorizontal && normalScale > 0
      ? { xPt: 0, yPt: crispOffset(finalPath[0]!.yPt, segment.widthPt * normalScale, context.dpr) }
      : { xPt: 0, yPt: 0 };
  const localOffset = inverseMapAffineVector(pointToCss, cssOffset) ?? { xPt: 0, yPt: 0 };
  const first = path[0]!;
  ctx.moveTo(first.xPt + localOffset.xPt, first.yPt + localOffset.yPt);
  for (const point of path.slice(1)) {
    ctx.lineTo(point.xPt + localOffset.xPt, point.yPt + localOffset.yPt);
  }
  // Layout bounds generated wavy paths by their vertices plus half the stroke
  // width. Canvas' default miter protrudes beyond that envelope at zig-zag
  // peaks, so retain a bevel join for the generated waveform and keep paint
  // exactly inside the acquired point-space bound.
  const boundedWaveJoin = segment.style === 'wavy' && path.length > 2;
  if (boundedWaveJoin) {
    ctx.save();
    (ctx as unknown as { lineJoin: CanvasLineJoin }).lineJoin = 'bevel';
  }
  ctx.stroke();
  if (boundedWaveJoin) ctx.restore();
  ctx.setLineDash([]);
}
