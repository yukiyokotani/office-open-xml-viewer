import type { DrawingLayout } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';
import {
  canvasFontString,
  clipDrawingMLShape,
  paintDrawingMLShape,
  resolveFill,
  withDrawingMLShapeTransform,
  withPatternCoordinateSpace,
} from '@silurus/ooxml-core';
import { paintRetainedResource } from './canvas-resource.js';

function paintPageAlignedDrawingMLShape(
  plan: Parameters<typeof paintDrawingMLShape>[1],
  context: CanvasPaintContext,
): void {
  const ctx = context.ctx as CanvasRenderingContext2D;
  if (!plan.resolvedGeometry) throw new Error('Missing retained DrawingML geometry');
  const paint = () => paintDrawingMLShape(ctx, plan, 1);
  if (context.patternPageToDevice) {
    withPatternCoordinateSpace(ctx, context.patternPageToDevice, paint);
  } else {
    paint();
  }
}

export function paintDrawingLayout(node: DrawingLayout, context: CanvasPaintContext): void {
  for (const command of node.commands) {
    if (command.kind === 'noop') continue;
    if (command.kind === 'drawingml-shape') {
      // Page setup already maps retained point coordinates to device pixels. A
      // second scale here would multiply stroke widths and arrowhead geometry.
      paintPageAlignedDrawingMLShape(command.plan, context);
      continue;
    }
    if (command.kind === 'drawingml-image-fill') {
      if (!command.plan.resolvedGeometry) throw new Error('Missing retained DrawingML geometry');
      if (!context.resources) throw new Error(`Missing retained resource painter for ${command.resourceKey}`);
      const { x, y, w, h } = command.plan.rect;
      const fillRect = command.fillRect ?? { l: 0, t: 0, r: 0, b: 0 };
      const destination = {
        xPt: x + fillRect.l * w,
        yPt: y + fillRect.t * h,
        widthPt: w * (1 - fillRect.l - fillRect.r),
        heightPt: h * (1 - fillRect.t - fillRect.b),
      };
      if (destination.widthPt > 0 && destination.heightPt > 0) {
        withDrawingMLShapeTransform(
          context.ctx as CanvasRenderingContext2D,
          command.plan,
          () => {
            clipDrawingMLShape(
              context.ctx as CanvasRenderingContext2D,
              command.plan,
            );
            paintRetainedResource(
              command.resourceKey,
              'image',
              destination,
              undefined,
              context,
            );
          },
        );
      }
      // The retained plan has a null base fill; this second pass paints only
      // the authored outline/arrow decorations in the same transform.
      paintPageAlignedDrawingMLShape(command.plan, context);
      continue;
    }
    if (command.kind === 'resource') {
      if (!context.resources) throw new Error(`Missing retained resource painter for ${command.resourceKey}`);
      paintRetainedResource(
        command.resourceKey,
        command.resourceKind,
        command.rect,
        command.orientation,
        context,
        command.nativeImagePlan,
      );
      continue;
    }
    if (command.kind === 'fill-rect') {
      context.ctx.fillStyle = command.fill;
      context.ctx.fillRect(
        command.rect.xPt,
        command.rect.yPt,
        command.rect.widthPt,
        command.rect.heightPt,
      );
      continue;
    }
    if (command.kind === 'stroke-rect') {
      context.ctx.strokeStyle = command.stroke;
      context.ctx.lineWidth = command.lineWidthPt;
      context.ctx.setLineDash([...command.dashPt]);
      context.ctx.strokeRect(
        command.rect.xPt,
        command.rect.yPt,
        command.rect.widthPt,
        command.rect.heightPt,
      );
      context.ctx.setLineDash([]);
      continue;
    }
    if (command.kind === 'watermark-text') {
      // Use the same resolved OOXML fill primitive as DrawingML shapes. `null`
      // is authored no-fill, not a request for a renderer-chosen fallback ink.
      const fill = resolveFill(
        command.fill as Parameters<typeof resolveFill>[0],
        context.ctx as CanvasRenderingContext2D,
        command.rect.xPt,
        command.rect.yPt,
        command.rect.widthPt,
        command.rect.heightPt,
      );
      if (fill === null) continue;
      context.ctx.save();
      const centerXPt = command.rect.xPt + command.rect.widthPt / 2;
      const centerYPt = command.rect.yPt + command.rect.heightPt / 2;
      context.ctx.translate(centerXPt, centerYPt);
      if (command.rotationDeg !== 0) {
        context.ctx.rotate(command.rotationDeg * Math.PI / 180);
      }
      if (command.fitShape) {
        context.ctx.scale(
          command.rect.widthPt / command.sourceBounds.widthPt,
          command.rect.heightPt / command.sourceBounds.heightPt,
        );
        context.ctx.translate(
          -(command.sourceBounds.xPt + command.sourceBounds.widthPt / 2),
          -(command.sourceBounds.yPt + command.sourceBounds.heightPt / 2),
        );
      } else {
        // Without fitshape, preserve authored point size and put the measured
        // source box at the shape box origin; rotation still uses shape centre.
        context.ctx.translate(
          command.rect.xPt - centerXPt - command.sourceBounds.xPt,
          command.rect.yPt - centerYPt - command.sourceBounds.yPt,
        );
      }
      context.ctx.globalAlpha *= command.opacity;
      context.ctx.fillStyle = fill;
      context.ctx.textAlign = 'left';
      context.ctx.textBaseline = 'alphabetic';
      let xPt = 0;
      for (const span of command.spans) {
        context.ctx.font = canvasFontString(
          span.fontRoute,
          command.fontSizePt,
          span.fontWeight,
          span.fontStyle,
        );
        context.ctx.fillText(span.text, xPt, 0);
        xPt += span.advancePt;
      }
      context.ctx.restore();
      continue;
    }
    context.ctx.fillStyle = command.fill;
    context.ctx.font = canvasFontString(
      command.fontRoute,
      command.fontSizePt,
      command.fontWeight,
      command.fontStyle,
    );
    context.ctx.textAlign = command.align === 'start' ? 'left' : command.align === 'end' ? 'right' : 'center';
    context.ctx.textBaseline = command.baseline;
    const xPt = command.align === 'start'
      ? command.rect.xPt
      : command.align === 'end'
        ? command.rect.xPt + command.rect.widthPt
        : command.rect.xPt + command.rect.widthPt / 2;
    const yPt = command.baseline === 'top'
      ? command.rect.yPt
      : command.baseline === 'bottom'
        ? command.rect.yPt + command.rect.heightPt
        : command.rect.yPt + command.rect.heightPt / 2;
    context.ctx.fillText(command.text, xPt, yPt);
  }
}
