import { paintVmlConcentricFill } from './vml-concentric-fill.js';
import { paintVmlImageMask } from './vml-mask.js';
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
import { fillRetainedRouteText } from './canvas-text-fill.js';

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
  paintDrawingCommands(node.commands, context);
}

/** Replay retained drawing commands in the current point space. */
export function paintDrawingCommands(
  commands: DrawingLayout['commands'],
  context: CanvasPaintContext,
): void {
  for (const command of commands) {
    if (command.kind === 'noop') continue;
    if (command.kind === 'drawingml-shape') {
      if (command.concentricFill) {
        paintVmlConcentricFill(command, context);
        continue;
      }
      if (command.insetStroke) {
        withDrawingMLShapeTransform(context.ctx as CanvasRenderingContext2D, command.plan, () => {
          clipDrawingMLShape(context.ctx as CanvasRenderingContext2D, command.plan);
          context.ctx.globalAlpha *= command.opacity ?? 1;
          paintPageAlignedDrawingMLShape(
            { ...command.plan, transform: { rotationDeg: 0, flipH: false, flipV: false } },
            context,
          );
        });
        continue;
      }
      // Page setup already maps retained point coordinates to device pixels. A
      // second scale here would multiply stroke widths and arrowhead geometry.
      if (command.opacity === undefined) {
        paintPageAlignedDrawingMLShape(command.plan, context);
      } else {
        context.ctx.save();
        try {
          context.ctx.globalAlpha *= command.opacity;
          paintPageAlignedDrawingMLShape(command.plan, context);
        } finally {
          context.ctx.restore();
        }
      }
      continue;
    }
    if (command.kind === 'drawingml-image-fill') {
      if (!command.plan.resolvedGeometry) throw new Error('Missing retained DrawingML geometry');
      if (!context.resources)
        throw new Error(`Missing retained resource painter for ${command.resourceKey}`);
      if (command.strokeMask) {
        const { x, y, w, h } = command.plan.rect;
        const pad = command.insetStroke ? 0 : (command.plan.stroke?.width ?? 0) / 2;
        withDrawingMLShapeTransform(context.ctx as CanvasRenderingContext2D, command.plan, () => {
          if (command.insetStroke)
            clipDrawingMLShape(context.ctx as CanvasRenderingContext2D, command.plan);
          const brushBounds = {
            xPt: x - pad,
            yPt: y - pad,
            widthPt: w + pad * 2,
            heightPt: h + pad * 2,
          };
          paintVmlImageMask(
            command.resourceKey,
            command.maskBounds ?? brushBounds,
            context,
            (ctx) =>
              paintDrawingMLShape(
                ctx,
                { ...command.plan, transform: { rotationDeg: 0, flipH: false, flipV: false } },
                1,
              ),
            brushBounds,
            {
              rotationDeg: command.imageRotationDeg ?? 0,
              flipH: command.imageFlipH,
              flipV: command.imageFlipV,
            },
          );
        });
        continue;
      }
      const { x, y, w, h } = command.plan.rect;
      const fillRect = command.fillRect ?? { l: 0, t: 0, r: 0, b: 0 };
      const destination = {
        xPt: x + fillRect.l * w,
        yPt: y + fillRect.t * h,
        widthPt: w * (1 - fillRect.l - fillRect.r),
        heightPt: h * (1 - fillRect.t - fillRect.b),
      };
      if (destination.widthPt > 0 && destination.heightPt > 0) {
        withDrawingMLShapeTransform(context.ctx as CanvasRenderingContext2D, command.plan, () => {
          clipDrawingMLShape(context.ctx as CanvasRenderingContext2D, command.plan);
          context.ctx.save();
          try {
            if (command.imageRotationDeg || command.imageFlipH || command.imageFlipV) {
              context.ctx.translate(x + w / 2, y + h / 2);
              // Replay the inverse retained host transform inside its clip.
              context.ctx.scale(command.imageFlipH ? -1 : 1, command.imageFlipV ? -1 : 1);
              context.ctx.rotate(((command.imageRotationDeg ?? 0) * Math.PI) / 180);
              context.ctx.translate(-x - w / 2, -y - h / 2);
            }
            paintRetainedResource(command.resourceKey, 'image', destination, undefined, context);
          } finally {
            context.ctx.restore();
          }
        });
      }
      // The retained plan has a null base fill; this second pass paints only
      // the authored outline/arrow decorations in the same transform.
      paintPageAlignedDrawingMLShape(command.plan, context);
      continue;
    }
    if (command.kind === 'resource') {
      if (!context.resources)
        throw new Error(`Missing retained resource painter for ${command.resourceKey}`);
      paintRetainedResource(
        command.resourceKey,
        command.resourceKind,
        command.rect,
        command.orientation,
        context,
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
      if (command.fill === null) continue;
      const paintText = (context: CanvasPaintContext) => {
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
        if (fill === null) return;
        context.ctx.save();
        if (command.glyphTransform) {
          const m = command.glyphTransform;
          context.ctx.transform(m.a, m.b, m.c, m.d, m.e, m.f);
        }
        const centerXPt = command.rect.xPt + command.rect.widthPt / 2;
        const centerYPt = command.rect.yPt + command.rect.heightPt / 2;
        context.ctx.translate(centerXPt, centerYPt);
        if (command.pathSlices) {
          context.ctx.translate(-command.rect.widthPt / 2, -command.rect.heightPt / 2);
          context.ctx.globalAlpha *= command.opacity;
          if (fill !== null) context.ctx.fillStyle = fill;
          context.ctx.textAlign = 'left';
          context.ctx.textBaseline = 'alphabetic';
          for (const slice of command.pathSlices) {
            context.ctx.save();
            context.ctx.transform(slice.a, slice.b, slice.c, slice.d, slice.xPt, slice.yPt);
            context.ctx.beginPath();
            context.ctx.rect(
              -slice.sourceWidth / 2,
              command.sourceBounds.yPt,
              slice.sourceWidth,
              command.sourceBounds.heightPt,
            );
            context.ctx.clip();
            let xPt = -slice.sourceX;
            for (const span of command.spans) {
              const font = canvasFontString(
                span.fontRoute,
                command.fontSizePt,
                span.fontWeight,
                span.fontStyle,
              );
              context.ctx.font = font;
              fillRetainedRouteText(context, span.fontRoute, font, span.text, xPt, 0);
              xPt += span.advancePt;
            }
            context.ctx.restore();
          }
          context.ctx.restore();
          return;
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
        if (fill !== null) context.ctx.fillStyle = fill;
        context.ctx.textAlign = 'left';
        context.ctx.textBaseline = 'alphabetic';
        let xPt = 0;
        for (const span of command.spans) {
          const font = canvasFontString(
            span.fontRoute,
            command.fontSizePt,
            span.fontWeight,
            span.fontStyle,
          );
          context.ctx.font = font;
          fillRetainedRouteText(context, span.fontRoute, font, span.text, xPt, 0);
          xPt += span.advancePt;
        }
        context.ctx.restore();
      };
      // §19.1.2.19: rotation after reflection, around the host centre.
      // Both solid glyphs and raster brush masks share this one outer frame;
      // mask callbacks never reapply it. Shadow commands retain the same flags.
      context.ctx.save();
      try {
        const x = command.rect.xPt + command.rect.widthPt / 2,
          y = command.rect.yPt + command.rect.heightPt / 2;
        context.ctx.translate(x, y);
        context.ctx.rotate((command.rotationDeg * Math.PI) / 180);
        context.ctx.scale(command.flipH ? -1 : 1, command.flipV ? -1 : 1);
        context.ctx.translate(-x, -y);
        if (command.imageResourceKey) {
          paintVmlImageMask(
            command.imageResourceKey,
            command.maskBounds ?? command.rect,
            context,
            (ctx) => paintText({ ...context, ctx }),
            command.rect,
            { rotationDeg: (command.imageRotationDeg ?? 0) - command.rotationDeg,
              flipH: command.imageFlipH, flipV: command.imageFlipV },
          );
        } else paintText(context);
      } finally {
        context.ctx.restore();
      }
      continue;
    }
    context.ctx.fillStyle = command.fill;
    const font = canvasFontString(
      command.fontRoute,
      command.fontSizePt,
      command.fontWeight,
      command.fontStyle,
    );
    context.ctx.font = font;
    context.ctx.textAlign =
      command.align === 'start' ? 'left' : command.align === 'end' ? 'right' : 'center';
    context.ctx.textBaseline = command.baseline;
    const xPt =
      command.align === 'start'
        ? command.rect.xPt
        : command.align === 'end'
          ? command.rect.xPt + command.rect.widthPt
          : command.rect.xPt + command.rect.widthPt / 2;
    const yPt =
      command.baseline === 'top'
        ? command.rect.yPt
        : command.baseline === 'bottom'
          ? command.rect.yPt + command.rect.heightPt
          : command.rect.yPt + command.rect.heightPt / 2;
    fillRetainedRouteText(context, command.fontRoute, font, command.text, xPt, yPt);
  }
}
