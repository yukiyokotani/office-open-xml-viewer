import type {
  LayoutRect,
  ParagraphLayout,
  ResolvedFloatingTablePlacementLayout,
  TableLayout,
} from '../layout/types.js';
import {
  composeAffine,
  inverseMapAffinePoint,
  mapAffinePoint,
  scaleAffine,
  translationAffine,
} from './affine.js';
import {
  oneDevicePixelCssWidth,
  paintCompoundBorderFrame,
  paintStrokeSegment,
} from './canvas-border.js';
import { paintParagraphLayout } from './canvas-text.js';
import { canvasPaintFrame } from './deferred-paint-frame.js';
import type { CanvasPaintContext } from './types.js';

function outwardRasterClipBounds(
  bounds: LayoutRect,
  context: CanvasPaintContext,
): LayoutRect {
  const pointToCss = context.pointToCss ?? scaleAffine(context.scale);
  // A rotated/skewed clip is a polygon in device space; expanding its
  // axis-aligned bounding box and mapping that box back would change the clip
  // shape. Retain the exact authored rectangle for those uncommon transforms.
  if (pointToCss.b !== 0 || pointToCss.c !== 0) return bounds;
  const corners = [
    { xPt: bounds.xPt, yPt: bounds.yPt },
    { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt },
    { xPt: bounds.xPt, yPt: bounds.yPt + bounds.heightPt },
    { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt + bounds.heightPt },
  ].map((point) => mapAffinePoint(pointToCss, point));
  const cssXs = corners.map((point) => point.xPt);
  const cssYs = corners.map((point) => point.yPt);
  const leftCss = Math.floor(Math.min(...cssXs) * context.dpr) / context.dpr;
  const topCss = Math.floor(Math.min(...cssYs) * context.dpr) / context.dpr;
  const rightCss = Math.ceil(Math.max(...cssXs) * context.dpr) / context.dpr;
  const bottomCss = Math.ceil(Math.max(...cssYs) * context.dpr) / context.dpr;
  const localCorners = [
    { xPt: leftCss, yPt: topCss },
    { xPt: rightCss, yPt: topCss },
    { xPt: leftCss, yPt: bottomCss },
    { xPt: rightCss, yPt: bottomCss },
  ].map((point) => inverseMapAffinePoint(pointToCss, point));
  if (localCorners.some((point) => point === null)) return bounds;
  const local = localCorners.filter(
    (point): point is Readonly<{ xPt: number; yPt: number }> => point !== null,
  );
  const xs = local.map((point) => point.xPt);
  const ys = local.map((point) => point.yPt);
  return {
    xPt: Math.min(...xs),
    yPt: Math.min(...ys),
    widthPt: Math.max(...xs) - Math.min(...xs),
    heightPt: Math.max(...ys) - Math.min(...ys),
  };
}

function paintPlacedChild(
  layout: ParagraphLayout | TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
  paintOwnTableBorders = true,
): void {
  const dxPt = placement.xPt - layout.flowBounds.xPt;
  const dyPt = placement.yPt - layout.flowBounds.yPt;
  const parentLayoutTranslation = context.layoutTranslationPt ?? {
    xPt: 0,
    yPt: 0,
  };
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  const frame = canvasPaintFrame(context.ctx, () => context.ctx.translate(dxPt, dyPt));
  const childContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: {
      xPt: parentLayoutTranslation.xPt + dxPt,
      yPt: parentLayoutTranslation.yPt + dyPt,
    },
  };
  frame(() => {
    if (layout.kind === 'paragraph') paintParagraphLayout(layout, childContext);
    else paintTableLayoutInternal(
      layout,
      childContext,
      layout.resolvedFloatingTables ?? [],
      paintOwnTableBorders,
    );
  })();
}

function paintTableBorders(node: TableLayout, context: CanvasPaintContext): void {
  // Word preserves authored table-border geometry, but rasterizes a subpixel
  // hairline with at least one device pixel of coverage. Keep that distinction
  // in paint so layout/conflict resolution continues to use the OOXML width.
  const minimumCssWidthPx = oneDevicePixelCssWidth(context);
  const framedSegments = new Set<number>();
  for (const frame of node.compoundBorderFrames ?? []) {
    if (paintCompoundBorderFrame(frame.bounds, frame.border, context)) {
      frame.segmentIndexes.forEach((index) => framedSegments.add(index));
    }
  }
  node.borders.forEach((border, index) => {
    if (framedSegments.has(index)) return;
    paintStrokeSegment(border, context, minimumCssWidthPx);
  });
}

function paintPlacedTableBorders(
  node: TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
): void {
  const dxPt = placement.xPt - node.flowBounds.xPt;
  const dyPt = placement.yPt - node.flowBounds.yPt;
  const parentLayoutTranslation = context.layoutTranslationPt ?? { xPt: 0, yPt: 0 };
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  const frame = canvasPaintFrame(context.ctx, () => context.ctx.translate(dxPt, dyPt));
  const borderContext: CanvasPaintContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: {
      xPt: parentLayoutTranslation.xPt + dxPt,
      yPt: parentLayoutTranslation.yPt + dyPt,
    },
  };
  frame(() => {
    const paint = (): void => paintTableBorders(node, borderContext);
    if (!node.clipBounds) {
      paint();
      return;
    }
    const clipFrame = canvasPaintFrame(context.ctx, () => {
      context.ctx.beginPath();
      context.ctx.rect(
        node.clipBounds!.xPt,
        node.clipBounds!.yPt,
        node.clipBounds!.widthPt,
        node.clipBounds!.heightPt,
      );
      context.ctx.clip();
    });
    clipFrame(paint)();
  })();
}

function paintTableContents(
  node: TableLayout,
  context: CanvasPaintContext,
  floatingTables: readonly ResolvedFloatingTablePlacementLayout[],
  paintOwnBorders: boolean,
): void {
  for (const row of node.rows) {
    for (const cell of row.cells) {
      const ownsContinuationPaint = 'visualMergeOwnership' in cell
        && cell.visualMergeOwnership === 'continuation';
      if (cell.verticalMerge === 'continue' && !ownsContinuationPaint) continue;
      if (cell.background) {
        context.ctx.fillStyle = cell.background.color;
        context.ctx.fillRect(
          cell.flowBounds.xPt,
          cell.flowBounds.yPt,
          cell.flowBounds.widthPt,
          cell.flowBounds.heightPt,
        );
      }
      const paintBlocks = (
        blockContext: CanvasPaintContext,
        paintDirectNestedTableBorders = true,
      ): void => {
        for (const block of cell.blocks) {
          paintPlacedChild(block.layout, {
            // Paragraphs are normalized to the cell content origin. A nested
            // table's own flowBounds additionally retain jc/tblInd placement
            // within that band, so translate its coordinate space by the outer
            // content origin without erasing that local offset.
            xPt: cell.contentBounds.xPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.xPt : 0),
            yPt: cell.flowBounds.yPt + block.offsetPt
              + (block.layout.kind === 'table' ? block.layout.flowBounds.yPt : 0),
          }, blockContext, block.layout.kind !== 'table' || paintDirectNestedTableBorders);
        }
      };
      if (cell.verticalText) {
        // ECMA-376 §17.4.72: blocks are in the rotated local frame. Clip to
        // the physical cell, then map the local frame into the table.
        const verticalText = cell.verticalText;
        const clip = cell.clipBounds ?? cell.flowBounds;
        const rotatedFrame = canvasPaintFrame(context.ctx, () => {
          context.ctx.beginPath();
          context.ctx.rect(clip.xPt, clip.yPt, clip.widthPt, clip.heightPt);
          context.ctx.clip();
          const m = verticalText.transform;
          context.ctx.transform(m.a, m.b, m.c, m.d, m.e, m.f);
        });
        const rotatedContext: CanvasPaintContext = {
          ...context,
          pointToCss: composeAffine(
            context.pointToCss ?? scaleAffine(context.scale),
            verticalText.transform,
          ),
          textBoxVerticalMode: verticalText.mode,
        };
        rotatedFrame(() => {
          for (const block of cell.blocks) {
            paintPlacedChild(block.layout, {
              xPt: cell.contentBounds.xPt,
              yPt: block.offsetPt,
            }, rotatedContext, true);
          }
        })();
        continue;
      }
      if (!cell.clipBounds) {
        paintBlocks(context);
        continue;
      }
      // Exact-height cells own an exact content clip. Expanding this clip to a
      // device-pixel boundary lets text, fills, and resources bleed into the
      // next row at fractional viewport scales.
      const contentFrame = canvasPaintFrame(context.ctx, () => {
        context.ctx.beginPath();
        context.ctx.rect(
          cell.clipBounds!.xPt,
          cell.clipBounds!.yPt,
          cell.clipBounds!.widthPt,
          cell.clipBounds!.heightPt,
        );
        context.ctx.clip();
      });
      // Direct nested-table strokes may be centred exactly on the cell edge.
      // Paint their content once under the exact clip, withholding only their
      // own border pass; then paint those retained strokes under an outward
      // raster-coverage clip. No text/fill/resource receives the larger clip.
      contentFrame(() => paintBlocks(context, false))();
      if (cell.blocks.some((block) => block.layout.kind === 'table')) {
        const borderClip = outwardRasterClipBounds(cell.clipBounds, context);
        const borderFrame = canvasPaintFrame(context.ctx, () => {
          context.ctx.beginPath();
          context.ctx.rect(
            borderClip.xPt,
            borderClip.yPt,
            borderClip.widthPt,
            borderClip.heightPt,
          );
          context.ctx.clip();
        });
        borderFrame(() => {
          for (const block of cell.blocks) {
            if (block.layout.kind !== 'table') continue;
            paintPlacedTableBorders(block.layout, {
              xPt: cell.contentBounds.xPt + block.layout.flowBounds.xPt,
              yPt: cell.flowBounds.yPt + block.offsetPt + block.layout.flowBounds.yPt,
            }, context);
          }
        })();
      }
    }
  }
  paintResolvedFloatingTablePlacements(floatingTables, context);
  if (paintOwnBorders) paintTableBorders(node, context);
}

/** Paint stored table geometry. No inheritance, sizing, shaping, or conflict resolution occurs here. */
function paintTableLayoutInternal(
  node: TableLayout,
  context: CanvasPaintContext,
  placements: readonly ResolvedFloatingTablePlacementLayout[],
  paintOwnBorders: boolean,
): void {
  if (!node.clipBounds) {
    paintTableContents(node, context, placements, paintOwnBorders);
    return;
  }
  const clipBounds = node.clipBounds;
  const frame = canvasPaintFrame(context.ctx, () => {
    context.ctx.beginPath();
    context.ctx.rect(
      clipBounds.xPt,
      clipBounds.yPt,
      clipBounds.widthPt,
      clipBounds.heightPt,
    );
    context.ctx.clip();
  });
  frame(() => paintTableContents(node, context, placements, paintOwnBorders))();
}

/** Paint stored table geometry. No inheritance, sizing, shaping, or conflict resolution occurs here. */
export function paintTableLayout(
  node: TableLayout,
  context: CanvasPaintContext,
  floatingTables?: readonly ResolvedFloatingTablePlacementLayout[],
): void {
  paintTableLayoutInternal(
    node,
    context,
    floatingTables ?? node.resolvedFloatingTables ?? [],
    true,
  );
}

/** Paint only point-space placements already resolved by the layout adapter. */
export function paintResolvedFloatingTablePlacements(
  placements: readonly ResolvedFloatingTablePlacementLayout[],
  context: CanvasPaintContext,
): void {
  const parentTranslation = context.layoutTranslationPt ?? { xPt: 0, yPt: 0 };
  for (const placement of placements) {
    paintPlacedChild(placement.child, {
      xPt: placement.xPt - parentTranslation.xPt,
      yPt: placement.yPt - parentTranslation.yPt,
    }, context);
  }
}

/** Bridge an unscaled renderer canvas to the retained point-space table painter. */
export function paintPlacedTableLayout(
  node: TableLayout,
  placement: Readonly<{ xPt: number; yPt: number }>,
  context: CanvasPaintContext,
  floatingTables?: readonly ResolvedFloatingTablePlacementLayout[],
): void {
  const dxPt = placement.xPt - node.flowBounds.xPt;
  const dyPt = placement.yPt - node.flowBounds.yPt;
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    translationAffine(dxPt, dyPt),
  );
  const frame = canvasPaintFrame(context.ctx, () => {
    context.ctx.translate(dxPt * context.scale, dyPt * context.scale);
    context.ctx.scale(context.scale, context.scale);
  });
  const placedContext: CanvasPaintContext = {
    ...context,
    pointToCss,
    layoutTranslationPt: { xPt: dxPt, yPt: dyPt },
  };
  frame(() => paintTableLayout(node, placedContext, floatingTables))();
}
